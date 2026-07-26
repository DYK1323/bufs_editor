# AGENTS.md

## 프로젝트 개요

이 저장소는 BUFS 한글(HWP/HWPX) 문서 서식을 자동화하는 Windows 데스크톱 앱이다. Python Tkinter GUI가 중심이고, 한컴오피스 한글의 `HWPFrame.HwpObject` COM API를 통해 현재 열린 한글 문서에 스타일, 표 서식, 템플릿, 로고, 텍스트 정리 기능을 적용한다.

핵심 구현은 `bufs/hwp_style_mvp.py` 단일 파일에 대부분 들어 있다. 앱이 담당하는 주요 기능은 다음과 같다.

- 실행 중인 한글 COM 객체 연결 및 대상 문서/창 선택
- HWPX 스타일 템플릿에서 BUFS 문단/글자 스타일 읽기 및 적용
- 표 배경색, 글자색, 테두리, 제목행, 캡션, 여백 처리
- 숫자 쉼표, 날짜/요일, 줄바꿈, 수동 개요번호, 단위 변환, 규정명 괄호 등 선택 텍스트 정리
- 표지, 일반보고 양식, 대제목 번호박스, 증빙 양식 HWPX 생성
- 로고/이미지 삽입, PDF 증빙자료 이미지 변환
- 업데이트 메타데이터 확인

현재 구조는 "큰 단일 앱 파일 + JSON 설정 + HWPX/이미지 자산"에 가깝다. 사용자가 명시적으로 리팩터링을 요청하지 않는 한, 광범위한 구조 변경보다 작고 검증 가능한 수정이 우선이다.

## 주요 경로

- `bufs/hwp_style_mvp.py` - 메인 Tkinter 앱, 순수 유틸 함수, HWP COM 래퍼, `MvpApp`.
- `bufs/tests/test_cell_matrix_transforms.py` - 순수 함수와 fake-COM 기반 동작 테스트.
- `bufs/run-mvp.ps1` - 로컬 실행 스크립트. `bufs/.venv/Scripts/python.exe`를 사용한다.
- `bufs/verify-readiness.ps1` - Windows/Hancom COM 및 템플릿 준비 상태 검증.
- `build-release.ps1` - PyInstaller 배포 폴더 생성.
- `BUFS-HWP-Editor.spec` - PyInstaller spec. `.gitignore`에는 `*.spec`가 있지만 현재 파일은 저장소에 존재한다.
- `requirements.txt` - 런타임 의존성: `pywin32`, `PyMuPDF`, `Pillow`.
- `requirements-build.txt` - 빌드 의존성: 런타임 의존성 + `pyinstaller`.
- `bufs/style-sets.json` - 스타일 세트, 개요/인라인 규칙 설정.
- `bufs/table-settings.json` - 팔레트, 선 프리셋, 캡션 파서, 표 기본값.
- `bufs/style-order.json` - 스타일 표시 순서 기본값.
- `bufs/update-settings.json` - 업데이트 확인 기본 설정.
- `bufs/templates/` - 앱이 생성/삽입하는 HWPX 템플릿.
- `bufs/icons/` - 표 스타일/테두리 버튼 아이콘.
- `bufs/logos/` - BUFS 로고 자산.
- `HancomDev/` - 한글 자동화 공식 PDF. 새로운 COM 액션명, 파라미터셋, 속성명을 다룰 때 먼저 확인한다.
- `공문양식/` - 공문 HWPX 양식 자산. 일반 코드 fixture처럼 취급하지 말고, 사용자/문서 자산으로 조심해서 다룬다.

## 런타임 구조

`hwp_style_mvp.py`는 `app_root()`로 `ROOT`를 계산한다. 소스 실행 시 `ROOT`는 `bufs/`이고, PyInstaller 빌드 실행 시에는 번들 내부 또는 실행 파일 주변에서 `style-sets.json`이나 `보고서 본문 서식.hwpx`를 찾아 앱 루트를 결정한다.

번들 JSON은 최초 실행 시 사용자 설정 폴더로 복사될 수 있다.

- Windows: `%APPDATA%/BUFS-HWP-Editor`
- fallback: `~/.bufs-hwp-editor`

따라서 번들 JSON 기본값을 바꿔도 기존 사용자의 설정 파일에는 자동 반영되지 않을 수 있다. 기존 사용자에게 반드시 적용되어야 하는 변경은 마이그레이션, reset, 안내 흐름까지 함께 고려한다.

## 개발 명령

PowerShell을 저장소 루트에서 실행한다.

가상환경 생성 및 의존성 설치:

```powershell
py -3.11 -m venv bufs\.venv
bufs\.venv\Scripts\python.exe -m pip install --upgrade pip
bufs\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

앱 실행:

```powershell
.\bufs\run-mvp.ps1
```

한글 COM 없이 검증 가능한 단위 테스트:

```powershell
bufs\.venv\Scripts\python.exe -m unittest bufs.tests.test_cell_matrix_transforms
```

앱 스모크 체크:

```powershell
bufs\.venv\Scripts\python.exe bufs\hwp_style_mvp.py --smoke
```

`--smoke`는 한글 COM 연결을 시도하므로, 한컴오피스 한글이 설치되지 않았거나 COM 등록이 깨진 환경에서는 실패할 수 있다.

한글/HWPX 준비 상태 확인:

```powershell
.\bufs\verify-readiness.ps1
```

배포 폴더 빌드:

```powershell
.\build-release.ps1
```

## 작업 지침

- `hwp_style_mvp.py`는 UTF-8로 유지한다. 한글 UI 문자열, HWPX placeholder, 스타일명은 의미 있는 데이터이므로 임의로 영문화하거나 치환하지 않는다.
- GUI 이벤트 핸들러나 COM 호출부를 직접 바꾸기 전에, 가능하면 순수 유틸 함수와 테스트를 먼저 추가/수정한다.
- HWP COM 호출은 기존 래퍼를 우선 재사용한다: `run_hwp_command`, `execute_first_hwp_action`, `set_com_attr`, `set_hset_item`, `set_parameter_item`.
- 검증되지 않은 HWP 액션명, 파라미터셋, 속성명은 추측하지 말고 `HancomDev/` 공식 PDF에서 먼저 확인한다.
- 표, 문단, 셀, 도형처럼 한글 내부 구조를 "생성"하거나 "변형"하는 작업은 HWPML/XML 직접 삽입보다 한글 카탈로그 액션(`HAction.Execute`, `Run`)과 공식 파라미터셋을 우선 사용한다. `SetTextFile(..., "HWPML2X", ...)` 같은 직접 삽입은 공식 액션으로 불가능하거나 검증된 fallback인 경우에만 사용한다.
- HWP 자동화는 현재 커서 위치, 선택 모드, 포커스된 창, 표 선택 상태, Windows 클립보드 상태에 민감하다. 실패 로그와 복원 동작을 함께 고려한다.
- 테스트 중 원본 HWP/HWPX 템플릿을 직접 덮어쓰지 않는다. 생성물은 가능하면 `bufs/test-output/` 아래에 둔다.
- `build/`, `dist/`, `bufs/.venv/`, `__pycache__/`, `bufs/test-output/` 생성물은 커밋하지 않는다.
- 사용자 설정 JSON의 순서와 값은 UI 흐름에 직접 영향을 준다. 불필요한 재정렬이나 대량 포맷 변경을 피한다.
- 새 의존성은 꼭 필요할 때만 추가한다. 현재 외부 런타임 의존성은 Windows/HWP 자동화와 이미지/PDF 처리에 집중되어 있다.

## 테스트 전략

텍스트 변환, 날짜/숫자 처리, 클립보드 셀 행렬, 스타일 규칙, fake-COM 동작은 `bufs/tests/test_cell_matrix_transforms.py`에 테스트를 추가하거나 갱신한다.

한글 설치가 필요한 변경은 가능한 경우 다음을 실행한다.

- `.\bufs\verify-readiness.ps1`
- `bufs\.venv\Scripts\python.exe bufs\hwp_style_mvp.py --smoke`

현재 환경에 한글이 없어서 COM 검증을 실행할 수 없다면, 그 사실을 명시하고 다음으로 강한 검증을 수행한다. 보통 단위 테스트, import/syntax 확인, 관련 HWPX XML 구조 확인이 대안이다.

## 현재 작업트리 주의사항

이 파일을 추가할 당시 `공문양식/` 아래 HWPX 파일 3개가 이미 수정된 상태였다. 사용자가 명시적으로 요청하지 않는 한 되돌리거나 덮어쓰지 않는다.

일부 PowerShell 세션에서는 README나 한글 경로가 mojibake처럼 깨져 보일 수 있다. 새 문서는 UTF-8로 유지하고, 한글 내용은 UTF-8을 제대로 해석하는 편집기에서 확인하는 것이 안전하다.
