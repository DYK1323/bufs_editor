# BUFS 한글 스타일 자동화 MVP

한글(HWP/HWPX) 문서 편집을 빠르게 처리하기 위한 Windows용 Tkinter 앱입니다. 스타일 적용, 표 편집, 표지 생성, 로고 삽입, 숫자/날짜/줄바꿈 정리 기능을 제공합니다.

## 준비물

- Windows PC
- 한컴오피스 한글 설치
- Python 3.11 권장
- Git
- PowerShell

한글 자동화는 `HWPFrame.HwpObject` COM 객체를 사용하므로, 한컴오피스 한글이 설치되어 있어야 합니다.

## 처음 설치하기

PowerShell에서 원하는 작업 폴더로 이동한 뒤 실행합니다.

```powershell
git clone https://github.com/DYK1323/bufs_editor.git
cd bufs_editor
```

가상환경을 만들고 의존성을 설치합니다.

```powershell
py -3.11 -m venv bufs\.venv
bufs\.venv\Scripts\python.exe -m pip install --upgrade pip
bufs\.venv\Scripts\python.exe -m pip install pywin32
```

Python 3.11이 없다면 설치된 Python으로도 시도할 수 있습니다.

```powershell
python -m venv bufs\.venv
bufs\.venv\Scripts\python.exe -m pip install --upgrade pip
bufs\.venv\Scripts\python.exe -m pip install pywin32
```

## 실행 전 확인

한글 COM 연결과 템플릿 파일을 점검합니다.

```powershell
.\bufs\verify-readiness.ps1
```

PowerShell 실행 정책 때문에 스크립트가 막히면 현재 프로세스에서만 정책을 완화한 뒤 다시 실행합니다.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\bufs\verify-readiness.ps1
```

## 프로그램 실행

```powershell
.\bufs\run-mvp.ps1
```

직접 Python으로 실행할 수도 있습니다.

```powershell
bufs\.venv\Scripts\python.exe bufs\hwp_style_mvp.py
```

## 기본 사용 흐름

1. 한글에서 편집할 문서를 엽니다.
2. `.\bufs\run-mvp.ps1`로 앱을 실행합니다.
3. 앱의 `상태` 탭에서 `한글 연결 확인` 또는 `대상 확인`을 누릅니다.
4. 필요한 탭에서 기능을 실행합니다.

주요 탭:

- `스타일/정리`: 스타일 적용, 숫자 쉼표, 날짜/요일, 줄바꿈/개요번호 정리
- `표 편집`: 셀 배경색, 글자색, 표 테두리, 제목셀, 캡션, 마크다운 표 변환
- `표지/로고`: 표지 생성, BUFS 로고 삽입
- `상태`: 한글 연결 확인, 테스트 문서 열기, 로그 확인

## 주요 파일

```text
bufs/
  hwp_style_mvp.py       # 메인 프로그램
  run-mvp.ps1            # 실행 스크립트
  verify-readiness.ps1   # 한글 COM/템플릿 검증 스크립트
  table-settings.json    # 표 색상, 선, 제목셀, 여백, 단축키 설정
  style-order.json       # 스타일 표시 순서
  보고서 본문 서식.hwpx   # 스타일 라이브러리
  표지.hwpx              # 표지 템플릿
  logos/                 # 로고 이미지
  APP_INDEX.md           # 개발자를 위한 앱 구조 인덱스
HancomDev/
  *.pdf                  # 한글 자동화 공식 참고 문서
```

## 개발 참고

앱 구조, 의존성, 함수 구성은 `bufs/APP_INDEX.md`에 정리되어 있습니다.

한글 자동화 기능을 새로 추가할 때는 `HancomDev/` 폴더의 PDF를 먼저 확인하세요.

- `HwpAutomation_2504.pdf`: HWP 자동화 객체 모델
- `ActionTable_2504.pdf`: HAction 액션 이름
- `ParameterSetTable_2504.pdf`: HParameterSet과 속성
- `한글오토메이션EventHandler추가_2504.pdf`: 이벤트 처리

## Git에 올리지 않는 파일

다음 파일은 로컬 실행 부산물이므로 커밋하지 않습니다.

- `bufs/.venv/`
- `__pycache__/`
- `.omx/`
- `bufs/test-output/*.hwpx`

## 문제 해결

`가상환경 Python을 찾을 수 없습니다`

- `bufs\.venv`가 없거나 다른 위치에 만들어진 상태입니다.
- 위의 `처음 설치하기` 단계대로 `py -3.11 -m venv bufs\.venv`를 다시 실행하세요.

`COM 연결 실패` 또는 `HWPFrame.HwpObject` 관련 오류

- 한컴오피스 한글이 설치되어 있는지 확인하세요.
- 한글을 한 번 직접 실행한 뒤 다시 시도하세요.
- 회사/기관 PC 보안 정책에 따라 COM 자동화가 제한될 수 있습니다.

PowerShell에서 스크립트 실행이 막힘

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

위 명령은 현재 PowerShell 창에서만 적용됩니다.
