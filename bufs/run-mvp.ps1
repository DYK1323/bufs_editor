# 한글 스타일 자동화 MVP 실행 스크립트
# UTF-8로 저장한다.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $root ".venv\Scripts\python.exe"
$app = Join-Path $root "hwp_style_mvp.py"

if (-not (Test-Path $python)) {
  throw "가상환경 Python을 찾을 수 없습니다: $python"
}

Set-Location $root
& $python $app
