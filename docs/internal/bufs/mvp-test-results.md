# MVP 테스트 결과

검증 일시: 2026-07-15

## 생성한 파일

- `bufs/.venv/`
  - 개발용 가상환경
  - `pywin32` 설치됨
- `bufs/hwp_style_mvp.py`
  - Tkinter 기반 MVP 앱
- `bufs/run-mvp.ps1`
  - MVP 실행 스크립트

## 현재 MVP 범위

현재 MVP는 자산 확인과 일부 실제 적용을 지원한다.

지원 기능:

- 한글 COM 연결 확인
- `bufs/templates/보고서 본문 서식.hwpx` 스타일 이름 읽기
- 목록에서 선택한 스타일을 현재 한글 문서의 현재 문단/선택 영역에 적용
- `bufs/templates/표지.hwpx` 자리표시자 확인
- `bufs/logos/` 아래 로고 이미지 목록 읽기
- 표 선 프리셋/색상 팔레트 표시
- 숫자 쉼표 넣기/빼기 유틸 테스트

아직 연결하지 않은 기능:

- 스타일 세트 불러오기
- 표 테두리 실제 적용
- 제목셀 속성 실제 적용
- 표지 삽입/치환

## Smoke Test

실행 명령:

```powershell
.\bufs\.venv\Scripts\python.exe .\bufs\hwp_style_mvp.py --smoke
```

결과:

```text
styles=16
cover={'{{문서제목}}': True, '{{날짜}}': True, '{{부서명}}': True, '<대 외 비>': True}
logos=3
commas=1,234,567 2026-07-15 051-123-4567
remove=1234567
hwp_version=13, 0, 0, 3622
hwp_docs=1
```

판정:

- 통과

## 2026-07-15 스타일 적용 버튼 추가

사용자가 MVP에서 실제로 해볼 수 있도록 `스타일` 탭에 다음 동작을 추가했다.

- 스타일 목록 클릭 즉시 적용
- `선택 스타일 다시 적용`
- `서식 테스트 문서 열기`

동작:

1. `서식 테스트 문서 열기`를 누르면 `bufs/templates/보고서 본문 서식.hwpx`의 복사본을 `bufs/test-output/`에 만들고 연다.
2. 한글 문서에서 스타일을 적용할 문단 또는 영역을 선택한다.
3. MVP의 `스타일` 탭에서 스타일을 클릭한다.
4. 한글 COM의 `HAction.Execute("Style", HStyle.HSet)` 경로로 즉시 스타일을 적용한다.

디버그:

- 하단 `디버그 로그` 영역을 항상 보이게 했다.
- 스타일 적용 요청, 한글 연결, `GetDefault`, `Execute` 결과를 로그에 남긴다.

주의:

- 현재 문서에 같은 스타일 세트가 들어 있어야 정확하게 동작한다.
- `bufs/templates/보고서 본문 서식.hwpx`에서 시작한 문서나 스타일 세트를 이미 가져온 문서에서 먼저 테스트한다.
- 스타일 세트가 없는 문서에서는 같은 스타일 ID가 다른 스타일을 가리킬 수 있다.

재검증:

```powershell
.\bufs\.venv\Scripts\python.exe .\bufs\hwp_style_mvp.py --smoke
.\bufs\.venv\Scripts\python.exe -m py_compile .\bufs\hwp_style_mvp.py
```

판정:

- 통과

참고:

- 현재 적용용으로 읽는 스타일은 `hh:style`의 명시적 `id/type/name` 기준이다.
- 현재 `bufs/templates/보고서 본문 서식.hwpx`에서 이 기준으로 읽히는 스타일 수는 `16`개다.

## 문법 검사

실행 명령:

```powershell
.\bufs\.venv\Scripts\python.exe -m py_compile .\bufs\hwp_style_mvp.py
```

판정:

- 통과

## 2026-07-15 실행 오류 수정

사용자 실행 중 다음 오류가 발생했다.

```text
AttributeError: '_tkinter.tkapp' object has no attribute 'style_list'
```

원인:

- `상태` 탭 생성 중 `refresh_all()`이 너무 일찍 호출됐다.
- 아직 `스타일` 탭의 `style_list` 위젯이 생성되기 전이라 초기화 순서가 맞지 않았다.

수정:

- 모든 탭을 생성한 뒤 `refresh_all()`을 한 번 호출하도록 초기화 순서를 변경했다.

재검증:

```powershell
.\bufs\.venv\Scripts\python.exe .\bufs\hwp_style_mvp.py --smoke
.\bufs\.venv\Scripts\python.exe -m py_compile .\bufs\hwp_style_mvp.py
.\bufs\.venv\Scripts\python.exe -c "import sys; sys.path.insert(0, r'.\bufs'); import hwp_style_mvp as m; app=m.MvpApp(); print('app-created'); app.destroy(); print('app-destroyed')"
```

판정:

- 통과

## 로고 인식

현재 인식된 로고 파일 수:

- `3`

현재 로고 폴더명:

- `가로형`
- `블록형CMYK`
- `블록형흑백`

MVP는 고정 폴더명에 의존하지 않고 `bufs/logos/` 아래 이미지 파일을 재귀적으로 읽는다.

## 실행 방법

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\bufs\run-mvp.ps1
```

또는:

```powershell
.\bufs\.venv\Scripts\python.exe .\bufs\hwp_style_mvp.py
```

## 다음 단계

1. 테스트 문서 복사본을 만든다.
2. 실제 스타일 적용 버튼을 연결한다.
3. 제목셀/표 테두리 적용을 복사본에서 검증한다.
4. 검증된 동작만 UI 버튼으로 승격한다.

