# 한글 자동화 구현 전 검증 스크립트
# UTF-8로 저장한다.
# 원본 문서를 수정하지 않고, COM 노출 여부와 템플릿 구조를 확인한다.

$ErrorActionPreference = "Continue"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$stylePath = Join-Path $root "보고서 본문 서식.hwpx"
$coverPath = Join-Path $root "표지.hwpx"

function Write-Section($title) {
  ""
  "## $title"
}

function Read-ZipEntryText($zipPath, $entryName) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [IO.Compression.ZipFile]::OpenRead($zipPath)
  try {
    $entry = $zip.Entries | Where-Object FullName -eq $entryName | Select-Object -First 1
    if (-not $entry) { return $null }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
  } finally {
    $zip.Dispose()
  }
}

Write-Section "파일 확인"
"stylePath=$stylePath exists=$(Test-Path $stylePath)"
"coverPath=$coverPath exists=$(Test-Path $coverPath)"

Write-Section "한글 COM 연결"
try {
  $hwp = New-Object -ComObject HWPFrame.HwpObject
  "ProgID=HWPFrame.HwpObject"
  "Version=$($hwp.Version)"
  "Documents=$($hwp.XHwpDocuments.Count)"
  "Windows=$($hwp.XHwpWindows.Count)"
} catch {
  "COM 연결 실패: $($_.Exception.Message)"
  exit 1
}

Write-Section "주요 COM 멤버"
$requiredMembers = @(
  "ImportStyle",
  "ExportStyle",
  "HAction",
  "HParameterSet",
  "RGBColor",
  "HwpLineWidth",
  "HwpLineType",
  "CellApply",
  "Open",
  "SaveAs"
)
foreach ($member in $requiredMembers) {
  $exists = [bool]($hwp | Get-Member -Name $member)
  "$member=$exists"
}

Write-Section "주요 파라미터셋"
$parameterSets = @(
  "HStyle",
  "HStyleTemplate",
  "HCell",
  "HCellBorderFill",
  "HTableBorderFill",
  "HCharShape"
)
foreach ($name in $parameterSets) {
  try {
    $set = $hwp.HParameterSet.$name
    $props = ($set | Get-Member -MemberType Property | Select-Object -ExpandProperty Name) -join ", "
    "$name=$props"
  } catch {
    "$name=확인 실패: $($_.Exception.Message)"
  }
}

Write-Section "색상 팔레트 변환"
$colors = @(
  @{ name = "노랑"; rgb = @(255, 203, 5) },
  @{ name = "진회색"; rgb = @(85, 85, 85) },
  @{ name = "회색"; rgb = @(150, 150, 150) },
  @{ name = "연회색"; rgb = @(204, 204, 204) },
  @{ name = "검정"; rgb = @(0, 0, 0) },
  @{ name = "흰색"; rgb = @(255, 255, 255) }
)
foreach ($color in $colors) {
  $rgb = $color.rgb
  try {
    "$($color.name) RGB($($rgb -join ',')) -> HWP $($hwp.RGBColor($rgb[0], $rgb[1], $rgb[2]))"
  } catch {
    "$($color.name) 변환 실패: $($_.Exception.Message)"
  }
}

Write-Section "선 굵기/선 종류 변환"
foreach ($width in @("0.12mm", "0.3mm", "0.7mm")) {
  try { "$width -> $($hwp.HwpLineWidth($width))" }
  catch { "$width -> 실패: $($_.Exception.Message)" }
}
foreach ($lineType in @("Solid", "Dot", "SlimThick", "ThickSlim")) {
  try { "$lineType -> $($hwp.HwpLineType($lineType))" }
  catch { "$lineType -> 실패: $($_.Exception.Message)" }
}

Write-Section "제목셀/테두리 관련 속성"
try {
  $cellProps = ($hwp.HParameterSet.HCell | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)
  "HCell.Header=$($cellProps -contains 'Header')"
} catch {
  "HCell 확인 실패: $($_.Exception.Message)"
}
try {
  $borderProps = ($hwp.HParameterSet.HCellBorderFill | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)
  foreach ($prop in @("BorderWidthTop", "BorderWidthBottom", "WidthHorz", "WidthVert", "BorderTypeTop", "BorderTypeBottom", "TypeHorz", "TypeVert", "FillAttr")) {
    "$prop=$($borderProps -contains $prop)"
  }
} catch {
  "HCellBorderFill 확인 실패: $($_.Exception.Message)"
}

Write-Section "스타일 라이브러리 구조"
if (Test-Path $stylePath) {
  $header = Read-ZipEntryText $stylePath "Contents/header.xml"
  if ($header) {
    $names = [regex]::Matches($header, '<[^>]*(?:style|Style)[^>]*name="([^"]+)"[^>]*>') |
      ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value) } |
      Sort-Object -Unique
    "styleNameCount=$($names.Count)"
    "sampleStyles=$((($names | Select-Object -First 20) -join ' | '))"
  } else {
    "Contents/header.xml을 찾지 못했다."
  }
}

Write-Section "표지 템플릿 구조"
if (Test-Path $coverPath) {
  $preview = Read-ZipEntryText $coverPath "Preview/PrvText.txt"
  if ($preview) {
    "has문서제목=$($preview.Contains('{{문서제목}}'))"
    "has날짜=$($preview.Contains('{{날짜}}'))"
    "has부서명=$($preview.Contains('{{부서명}}'))"
    "has대외비=$($preview.Contains('대 외 비'))"
  } else {
    "Preview/PrvText.txt를 찾지 못했다."
  }
}

Write-Section "한글 파일 열기 검증"
foreach ($path in @($stylePath, $coverPath)) {
  if (-not (Test-Path $path)) { continue }
  $tempHwp = New-Object -ComObject HWPFrame.HwpObject
  try {
    try { $tempHwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}
    try { $tempHwp.RegisterModule("FilePathCheckerDLL", "FilePathCheckerModule") | Out-Null } catch {}
    $ok = $tempHwp.Open($path, "HWPX", "")
    "$([IO.Path]::GetFileName($path)) open=$ok pageCount=$($tempHwp.PageCount) isModified=$($tempHwp.IsModified)"
  } catch {
    "$([IO.Path]::GetFileName($path)) 열기 실패: $($_.Exception.Message)"
  } finally {
    try { $tempHwp.Quit() | Out-Null } catch {}
  }
}

Write-Section "스타일 가져오기 직접 호출 검증"
try {
  $tempHwp = New-Object -ComObject HWPFrame.HwpObject
  try {
    try { $tempHwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}
    $styleSet = $tempHwp.HParameterSet.HStyleTemplate
    $styleSet.filename = $stylePath
    $importResult = $tempHwp.ImportStyle($styleSet.HSet)
    "ImportStyle direct result=$importResult"
  } finally {
    try { $tempHwp.Quit() | Out-Null } catch {}
  }
} catch {
  "ImportStyle direct call failed=$($_.Exception.Message)"
  "판정=ImportStyle 메서드는 존재하지만 filename만 지정한 직접 호출은 부족하다. NameLocals/NameEngs 또는 HAction 기반 호출법을 추가 조사해야 한다."
}
