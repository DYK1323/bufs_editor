# 구현 전 검증 결과

검증 일시: 2026-07-15

검증 스크립트:

- `bufs/verify-readiness.ps1`

## 요약

구현을 시작해도 되는 상태다.

다만 `스타일 세트 불러오기`는 COM 메서드가 존재하지만, 단순 호출 방식은 실패했으므로 추가 조사가 필요하다.

## 검증 완료

### Python / pywin32 상태

- 시스템 Python: `Python 3.11.9`
- 현재 시스템 Python에는 `win32com`이 설치되어 있지 않다.
- Tkinter는 사용 가능하다.
- PowerShell COM으로는 한글 자동화 객체 연결이 성공했다.

판정:

- 실제 Python 구현에는 `pywin32` 설치가 필요하다.
- 배포본은 PyInstaller로 빌드해 `pywin32`를 포함하는 방식이 적합하다.
- 최종 사용자에게 Python 또는 pywin32 수동 설치를 요구하지 않는 방향이 좋다.

### 한글 COM 연결

- `HWPFrame.HwpObject` 생성 성공
- 한글 버전: `13, 0, 0, 3622`
- `HAction`, `HParameterSet`, `Open`, `SaveAs` 사용 가능

### 스타일 관련 COM 멤버

- `ImportStyle` 존재
- `ExportStyle` 존재
- `HParameterSet.HStyle` 존재
- `HParameterSet.HStyleTemplate` 존재
  - `filename`
  - `NameLocals`
  - `NameEngs`

### 스타일 라이브러리 파일

- `bufs/보고서 본문 서식.hwpx` 존재
- COM으로 열기 성공
- 열기 직후 수정 상태: `False`
- 스타일 이름 추출 가능
- 추출된 스타일 이름 수: `102`

### 표지 템플릿 파일

- `bufs/표지.hwpx` 존재
- COM으로 열기 성공
- 열기 직후 수정 상태: `False`
- 자리표시자 확인 완료
  - `{{문서제목}}`
  - `{{날짜}}`
  - `{{부서명}}`
- 대외비 표시 확인 완료
  - `<대 외 비>`

### 색상 팔레트

한글 COM의 `RGBColor(byte, byte, byte)`로 변환 가능하다.

| 이름 | RGB | HWP 색상값 |
| --- | --- | ---: |
| 노랑 | `255, 203, 5` | `379903` |
| 진회색 | `85, 85, 85` | `5592405` |
| 회색 | `150, 150, 150` | `9868950` |
| 연회색 | `204, 204, 204` | `13421772` |
| 검정 | `0, 0, 0` | `0` |
| 흰색 | `255, 255, 255` | `16777215` |

### 선 굵기

한글 COM의 `HwpLineWidth(string)`로 변환 가능하다.

| 계획상 굵기 | 호출 문자열 | HWP 선 굵기값 |
| --- | --- | ---: |
| `0.12` | `0.12mm` | `1` |
| `0.3` | `0.3mm` | `5` |
| `0.7` | `0.7mm` | `9` |

### 선 종류

한글 COM의 `HwpLineType(string)`로 변환 가능하다.

| 선 종류 | 호출 문자열 | HWP 선 종류값 |
| --- | --- | ---: |
| 실선 | `Solid` | `1` |
| 점선 | `Dot` | `3` |
| 이중선 후보 1 | `SlimThick` | `9` |
| 이중선 후보 2 | `ThickSlim` | `10` |

참고:

- `Double` 문자열은 현재 COM 호출에서 실패했다.
- 이중선은 `SlimThick` 또는 `ThickSlim` 중 실제 시각 결과가 더 맞는 쪽을 선택해야 한다.

### 제목셀 속성

`HParameterSet.HCell`에 `Header` 속성이 존재한다.

따라서 한글 표 속성의 `제목셀` 설정은 COM 구현 가능성이 높다.

### 표 테두리 관련 속성

`HParameterSet.HCellBorderFill`에서 다음 속성을 확인했다.

- `BorderWidthTop`
- `BorderWidthBottom`
- `WidthHorz`
- `WidthVert`
- `BorderTypeTop`
- `BorderTypeBottom`
- `TypeHorz`
- `TypeVert`
- `FillAttr`

위/아래/내부 가로/내부 세로를 구분한 테두리 프리셋 구현 가능성이 높다.

## 남은 검증

### 스타일 세트 불러오기

`ImportStyle` 메서드는 존재하지만, 다음 단순 호출은 실패했다.

```powershell
$set = $hwp.HParameterSet.HStyleTemplate
$set.filename = "bufs/보고서 본문 서식.hwpx"
$hwp.ImportStyle($set.HSet)
```

실패 내용:

- `RPC_E_SERVERFAULT`
- `The server threw an exception.`

판정:

- `filename`만 지정하는 방식으로는 부족하다.
- `NameLocals` / `NameEngs` 배열을 채우는 방식 또는 `HAction` 기반 호출법을 추가 조사해야 한다.
- 대안으로 샘플 문단 복사 후 정리 방식도 유지한다.

### 스타일 이름 적용

아직 실제 문단에 스타일 이름을 적용하는 동작은 검증하지 않았다.

다음 단계에서 테스트 문서 복사본을 만들어 검증한다.

### 표 제목셀/테두리 실제 적용

속성 존재는 확인했지만, 실제 선택 셀에 적용하는 테스트는 아직 하지 않았다.

다음 단계에서 테스트 표가 들어 있는 문서 복사본을 만들어 검증한다.

### 표지 삽입/치환

자리표시자 존재와 파일 열기는 확인했다.

다음 단계에서 복사본 대상으로 다음을 검증한다.

- 자리표시자 치환
- 대외비 표시 제거
- 현재 문서 맨 앞 삽입

## 다음 권장 작업

1. `bufs` 안에 개발용 가상환경을 만들고 `pywin32`를 설치한다.
2. `bufs` 안에 테스트용 문서 복사본을 만든다.
3. 스타일 적용, 제목셀 설정, 테두리 적용을 실제 문서 복사본에서 검증한다.
4. `ImportStyle` 호출법을 추가 조사한다.
5. `ImportStyle`가 계속 불안정하면 스타일 적용 기능은 먼저 제공하고, 스타일 세트 불러오기는 2차 방식으로 우회한다.
