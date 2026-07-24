# BUFS HWP Editor

Windows용 한글(HWP/HWPX) 문서 편집 보조 앱입니다. 한컴오피스 한글의 COM 자동화를 사용해 스타일 적용, 표 편집, 표지 생성, 로고 삽입, 숫자/날짜/줄바꿈 정리 같은 반복 작업을 빠르게 처리합니다.

## 주요 기능

- 스타일/정리: 스타일 적용, 숫자 쉼표, 날짜/요일, 줄바꿈, 수동 개요번호 정리
- 표 편집: 셀 배경색, 글자색, 테두리, 제목셀, 캡션, 마크다운 표 변환
- 표지/로고: 보고서 표지 생성, BUFS 로고 삽입
- 상태: 한글 연결 확인, 대상 문서 확인, 테스트 문서 열기, 업데이트 확인

## 준비물

- Windows PC
- 한컴오피스 한글 설치
- 개발 실행 시 Python 3.11 권장
- Git, PowerShell

한글 자동화는 `HWPFrame.HwpObject` COM 객체를 사용하므로 한컴오피스 한글이 설치되어 있어야 합니다.

## 배포판 사용

GitHub Releases에서 `BUFS-HWP-Editor-v*.zip`을 내려받아 압축을 풉니다.

권장 ZIP 구조:

```text
BUFS-HWP-Editor-v1.0.2.zip
  BUFS-HWP-Editor/
    BUFS-HWP-Editor.exe
    BUFS-HWP-Updater.exe
    _internal/
    README-ko.md
    install-shortcut.ps1
```

압축을 푼 뒤 `BUFS-HWP-Editor.exe`를 실행합니다. 바탕화면 바로가기가 필요하면 배포 폴더 안의 `install-shortcut.ps1`을 PowerShell에서 실행합니다.

PowerShell 실행 정책 때문에 스크립트가 막히면 현재 창에서만 정책을 완화한 뒤 다시 실행합니다.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\install-shortcut.ps1
```

## 자동 업데이트

앱은 `bufs/update-settings.json` 설정을 번들링하며, 기본값은 GitHub Releases를 확인하도록 되어 있습니다.

```json
{
  "enabled": true,
  "check_on_start": true,
  "provider": "github",
  "github_owner": "DYK1323",
  "github_repo": "bufs_editor",
  "asset_pattern": "BUFS-HWP-Editor-v*.zip",
  "expected_root_dir": "BUFS-HWP-Editor",
  "expected_exe_name": "BUFS-HWP-Editor.exe",
  "timeout_seconds": 5
}
```

새 버전이 있으면 앱이 release ZIP을 다운로드하고 구조를 검증한 뒤, `BUFS-HWP-Updater.exe`를 임시 폴더로 복사해 실행합니다. updater는 실행 중인 앱이 종료되기를 기다린 다음 기존 설치 폴더를 백업하고 새 폴더로 교체한 뒤 앱을 다시 실행합니다.

## 개발 실행

```powershell
git clone https://github.com/DYK1323/bufs_editor.git
cd bufs_editor
py -3.11 -m venv bufs\.venv
bufs\.venv\Scripts\python.exe -m pip install --upgrade pip
bufs\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

한글 COM 연결과 템플릿 파일을 점검합니다.

```powershell
.\bufs\verify-readiness.ps1
```

앱을 실행합니다.

```powershell
.\bufs\run-mvp.ps1
```

직접 Python으로 실행할 수도 있습니다.

```powershell
bufs\.venv\Scripts\python.exe bufs\hwp_style_mvp.py
```

## 테스트

```powershell
python -m unittest bufs.tests.test_cell_matrix_transforms
```

## 릴리즈 빌드

```powershell
.\build-release.ps1
```

빌드가 끝나면 다음 파일이 생성됩니다.

```text
dist/
  BUFS-HWP-Editor/
  BUFS-HWP-Editor-v1.0.2.zip
  BUFS-HWP-Editor-v1.0.2.zip.sha256
```

GitHub Release tag는 `v1.0.2`, asset 이름은 `BUFS-HWP-Editor-v1.0.2.zip`처럼 앱 내부 `APP_VERSION`과 맞춥니다.

## 주요 파일

```text
bufs/
  hwp_style_mvp.py       # 메인 Tkinter 앱과 한글 COM 자동화
  update_helper.py       # 별도 자동 업데이트 helper
  run-mvp.ps1            # 개발 실행 스크립트
  verify-readiness.ps1   # 한글 COM/템플릿 검증 스크립트
  update-settings.json   # 업데이트 provider 설정
  table-settings.json    # 표 색상, 선, 제목셀, 여백, 단축키 설정
  style-sets.json        # 스타일 세트, 문단/글자 구분, 표/캡션 플래그
  style-order.json       # 구버전 스타일 표시 순서 호환용
  templates/             # HWPX 템플릿
  logos/                 # 로고 이미지
```

## 한글 자동화 참고 자료

한컴 자동화 API 이름과 동작은 한컴 개발자 자료, 사용자 제공 API 카탈로그, 실제 한글 COM 동작 검증을 함께 참고했습니다. 한컴DEV PDF와 파생 API 카탈로그 파일은 저장소에 포함하지 않습니다.

## Acknowledgements

- 스레드 `@nursing0705`: 한글 자동화 API 카탈로그와 관련 맥락 제공
- [innae1121-bit/gsghwp](https://github.com/innae1121-bit/gsghwp): 한글 자동화 API 카탈로그 참고

이 저장소에는 위 프로젝트의 코드를 복사해 포함하지 않았습니다.

## 라이선스

이 프로젝트는 [MIT License](LICENSE)로 배포됩니다.
