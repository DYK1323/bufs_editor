# Build a Windows release folder for BUFS HWP Editor.
# Keep this script ASCII-only so it parses on older Windows PowerShell hosts.

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $root "bufs\.venv\Scripts\python.exe"
$app = Join-Path $root "bufs\hwp_style_mvp.py"
$updaterApp = Join-Path $root "bufs\update_helper.py"
$dist = Join-Path $root "dist"
$build = Join-Path $root "build"

if (-not (Test-Path $venvPython)) {
  Write-Host "Creating virtual environment: bufs\.venv"
  py -3.11 -m venv (Join-Path $root "bufs\.venv")
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $root "requirements-build.txt")

if (Test-Path (Join-Path $dist "BUFS-HWP-Editor")) {
  Remove-Item -LiteralPath (Join-Path $dist "BUFS-HWP-Editor") -Recurse -Force
}
if (Test-Path (Join-Path $build "BUFS-HWP-Editor")) {
  Remove-Item -LiteralPath (Join-Path $build "BUFS-HWP-Editor") -Recurse -Force
}
if (Test-Path (Join-Path $build "BUFS-HWP-Updater")) {
  Remove-Item -LiteralPath (Join-Path $build "BUFS-HWP-Updater") -Recurse -Force
}
if (Test-Path (Join-Path $dist "BUFS-HWP-Updater.exe")) {
  Remove-Item -LiteralPath (Join-Path $dist "BUFS-HWP-Updater.exe") -Force
}

$pyinstallerArgs = @(
  "--noconsole",
  "--onedir",
  "--clean",
  "--name", "BUFS-HWP-Editor",
  "--hidden-import", "pythoncom",
  "--hidden-import", "pywintypes",
  "--hidden-import", "win32timezone",
  "--collect-submodules", "win32com"
)

Get-ChildItem -LiteralPath (Join-Path $root "bufs") -Filter "*.hwpx" | ForEach-Object {
  $pyinstallerArgs += @("--add-data", "$($_.FullName);bufs")
}

$pyinstallerArgs += @(
  "--add-data", "$(Join-Path $root 'bufs\style-sets.json');bufs",
  "--add-data", "$(Join-Path $root 'bufs\table-settings.json');bufs",
  "--add-data", "$(Join-Path $root 'bufs\update-settings.json');bufs",
  "--add-data", "$(Join-Path $root 'bufs\style-order.json');bufs",
  "--add-data", "$(Join-Path $root 'bufs\templates');bufs\templates",
  "--add-data", "$(Join-Path $root 'bufs\icons');bufs\icons",
  "--add-data", "$(Join-Path $root 'bufs\logos');bufs\logos",
  $app
)

& $venvPython -m PyInstaller @pyinstallerArgs

$releaseDir = Join-Path $dist "BUFS-HWP-Editor"

$updaterArgs = @(
  "--noconsole",
  "--onefile",
  "--clean",
  "--name", "BUFS-HWP-Updater",
  $updaterApp
)
& $venvPython -m PyInstaller @updaterArgs

Copy-Item -LiteralPath (Join-Path $dist "BUFS-HWP-Updater.exe") -Destination $releaseDir -Force
Copy-Item -LiteralPath (Join-Path $root "README-ko.md") -Destination $releaseDir -Force
Copy-Item -LiteralPath (Join-Path $root "install-shortcut.ps1") -Destination $releaseDir -Force

$appSource = Get-Content -LiteralPath $app -Raw -Encoding UTF8
$versionMatch = [regex]::Match($appSource, 'APP_VERSION\s*=\s*"([^"]+)"')
if (-not $versionMatch.Success) {
  throw "Cannot find APP_VERSION in $app"
}
$version = $versionMatch.Groups[1].Value
$zipPath = Join-Path $dist "BUFS-HWP-Editor-v$version.zip"
if (Test-Path $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -LiteralPath $releaseDir -DestinationPath $zipPath -Force
$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath "$zipPath.sha256" -Value "$hash  $(Split-Path -Leaf $zipPath)" -Encoding ASCII

Write-Host ""
Write-Host "Release folder created:"
Write-Host "  $releaseDir"
Write-Host "Release zip created:"
Write-Host "  $zipPath"
Write-Host ""
Write-Host "Upload the zip to GitHub Releases as the release asset."
