# BUFS HWP Editor v1.0.6

Release date: 2026-08-13

## 주요 업데이트

- 일반 문단 `일괄처리` 시작 시 대제목 번호박스 검사를 위해 문서 전체 HWPML을 먼저 읽던 동작을 제거해 시작 지연을 줄였습니다.
- 대제목 번호박스 후속 번호 경고는 실제 번호박스 규칙이 처리되는 경우에만 확인하도록 조정했습니다.
- 새 앱 아이콘을 창 아이콘과 PyInstaller 실행 파일 아이콘에 적용했습니다.

## 설치 방법

1. 아래 배포 파일에서 `BUFS-HWP-Editor-v1.0.6.zip`을 다운로드합니다.
2. ZIP 파일의 압축을 원하는 위치에 풉니다.
3. 압축을 푼 `BUFS-HWP-Editor` 폴더 안의 `BUFS-HWP-Editor.exe`를 실행합니다.
4. 바로가기가 필요하면 같은 폴더의 `install-shortcut.ps1`을 PowerShell에서 실행합니다.

PowerShell에서 스크립트 실행이 막히는 경우, 같은 PowerShell 창에서 아래 명령을 먼저 실행한 뒤 다시 시도하세요.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\install-shortcut.ps1
```

## 실행 방법

1. 한컴오피스 한글에서 편집할 문서를 먼저 엽니다.
2. `BUFS-HWP-Editor.exe`를 실행합니다.
3. 앱의 상태 탭에서 한글 연결 및 대상 문서를 확인합니다.
4. 필요한 탭에서 스타일, 표 편집, 표지/로고, 텍스트 정리 기능을 실행합니다.

Python 설치는 필요하지 않습니다. 배포 ZIP 안에 실행에 필요한 파일이 포함되어 있습니다.

## 업데이트 설치

기존 버전 사용자는 앱의 `상태` 탭에서 `업데이트 확인`을 눌러 새 버전을 받을 수 있습니다. 자동 설치가 되지 않는 환경에서는 이 릴리즈의 ZIP 파일을 직접 다운로드해 기존 설치 폴더를 새 폴더로 교체하세요.

## 검증

- 일괄처리 관련 단위 테스트로 일반 문단 처리와 표 셀 처리의 회귀 여부를 확인했습니다.
- 앱 아이콘 `.ico` 파일이 Pillow와 Tk `iconbitmap`에서 정상 로드되는지 확인했습니다.
- v1.0.6 릴리즈 빌드는 `build-release.ps1`로 생성합니다.

## 배포 파일

- `dist/BUFS-HWP-Editor-v1.0.6.zip`
- `dist/BUFS-HWP-Editor-v1.0.6.zip.sha256`
