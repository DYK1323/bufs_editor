# 한글 표 셀 선택 위치 탐색 메모

작성일: 2026-07-15

## 배경

여러 표 셀을 선택한 뒤 숫자 쉼표 넣기나 날짜 변환을 셀마다 적용하려면, 선택 범위를 한 번에 붙여넣는 방식이 아니라 선택된 셀을 하나씩 순회해야 한다.

실험 결과, 한글은 표 셀 블록 선택 상태에서 일반 텍스트 블록처럼 안정적인 시작/끝 위치를 바로 주지 않는다. 따라서 현재 커서 위치와 선택 셀 복사값을 조합해 선택 범위를 역산하는 방식이 필요하다.

## 확인한 사실

### 1. `GetSelectedPos()`는 표 셀 블록에서 유효하지 않았다

표 셀 블록 선택 상태에서 다음 값을 확인했다.

```text
GetSelectedPos() -> (False, 0, 0, 0, 0, 0, 0)
```

## 표/셀 여백 적용 메모

- `ActionTable_2504.pdf` 기준 `TablePropertyDialog`의 ParameterSet ID는 `ShapeObject`다.
- `ParameterSetTable_2504.pdf` 기준 셀 정보는 `ShapeObject` 안의 `ShapeTableCell`(`Cell` 타입)에 들어 있다.
- 따라서 셀 자체 여백은 `HParameterSet.HCell` 단독 실행보다 `HParameterSet.HShapeObject.ShapeTableCell`에 `HasMargin`, `MarginLeft/Right/Top/Bottom`을 넣고 `TablePropertyDialog`를 실행하는 경로가 문서 구조상 맞다.
- 매크로 기록상 셀 여백 변경에는 `HSet.SetItem("ShapeType", 3)` 및 `HSet.SetItem("ShapeCellSize", 0)`을 같이 넣어야 한다.
- 표의 기본 셀 안쪽 여백은 같은 `HShapeObject`의 `CellMarginLeft/Right/Top/Bottom`에도 존재하므로 함께 설정한다.
- 표 바깥 여백은 `HShapeObject.OutsideMarginLeft/Right/Top/Bottom`이다.

즉, 일반 텍스트 선택 범위용 위치 API로는 표 셀 블록의 시작/끝을 얻을 수 없었다.

### 2. `KeyIndicator()`와 `GetPos()`는 현재 셀, 즉 선택 끝점 쪽을 보여준다

예를 들어 A1에서 시작해 B4까지 선택하면:

```text
KeyIndicator -> (B4): 문자 입력
GetPos       -> 끝점에 해당하는 위치 튜플
```

B4에서 시작해 A1까지 반대로 선택하면:

```text
KeyIndicator -> (A1): 문자 입력
GetPos       -> 끝점에 해당하는 위치 튜플
```

따라서 `GetPos()`는 선택 범위의 anchor가 아니라 현재 끝점으로 봐야 한다.

### 3. 셀 블록 `Copy` 결과는 탭 구분 행렬이 아니라 선형 목록이다

임시 3x3 표에서 A1:B3을 선택하고 `Copy`했을 때, 기대한 값은 다음처럼 행/열 구조가 있는 문자열이었다.

```text
1\t2
4\t5
7\t8
```


하지만 실제 클립보드는 다음처럼 줄바꿈 목록이었다.

```text
\r\n1\r\n2\r\n4\r\n5\r\n7\r\n8\r\n
```

즉, 복사값만으로는 원래 선택이 `6 x 1`인지 `3 x 2`인지 확정할 수 없다.

### 4. 선택 범위에 변환된 전체 목록을 다시 붙여넣으면 각 셀에 전체 목록이 들어갈 수 있다

선택 범위 전체를 변환한 뒤 그대로 `Paste`하면 한글이 셀별 분배를 하지 않고, 각 선택 셀 안에 목록 전체를 넣는 문제가 발생했다.

따라서 여러 셀 변환은 반드시 다음 방식이어야 한다.

```text
선택 범위 위치 탐색
-> 시작 셀로 이동
-> 현재 셀 Copy
-> 현재 셀 텍스트만 변환
-> 현재 셀 Paste
-> 다음 셀로 이동
-> 반복
```

## 표 셀 이동 명령 확인

임시 표에서 확인한 이동 규칙:

```text
TableRightCell: 오른쪽 셀로 이동, 행 끝에서는 다음 행 첫 셀로 이동
TableLeftCell : 왼쪽 셀로 이동, 행 처음에서는 이전 행 마지막 셀로 이동
TableLowerCell: 아래 셀로 이동
TableUpperCell: 위 셀로 이동
TableCellBlock: 현재 셀 선택
```

예: 3x3 표에서 C3에서 `TableLeftCell` 반복:

```text
C3 -> B3 -> A3 -> C2 -> B2 -> A2 -> C1 -> B1 -> A1
```

## 선택 범위 역산 방법

목표는 현재 선택 끝점과 복사된 선택값 목록을 이용해 실제 직사각형 범위를 찾는 것이다.

### 입력

```text
endpoint = hwp.GetPos()
selected_values = 선택 셀 블록 Copy 결과에서 빈 줄을 제거한 값 목록
```

예:

```text
endpoint = (170, 0, 0)
KeyIndicator = (D7): 문자 입력
selected_values = [
  "269.0", "3571.0",
  "236.4", "3501.1",
  "505.4", "7072.1",
  "209.0", "3666.5",
  "208.9", "3606.5",
  "417.9", "7273.0"
]
```

### 후보 생성

선택값 개수 `N`의 모든 약수 조합을 후보 행/열로 본다.

```text
N = 12
후보:
1 x 12
2 x 6
3 x 4
4 x 3
6 x 2
12 x 1
```

그리고 현재 끝점이 직사각형의 어느 모서리인지 네 가지 경우를 모두 시도한다.

```text
bottom_right
bottom_left
top_right
top_left
```

### 후보 검증

각 후보에 대해:

1. `SetPos(endpoint)`로 끝점으로 이동한다.
2. 후보 모서리 정보에 따라 직사각형의 좌상단 셀로 이동한다.
3. 좌상단부터 행 우선 순서로 셀을 순회하며 각 셀 값을 읽는다.
4. 읽은 값 목록이 `selected_values`와 정확히 일치하면 그 후보를 실제 선택 범위로 확정한다.

실제 백업 문서에서 다음 후보가 12/12로 일치했다.

```text
rows = 6
cols = 2
endpoint_corner = bottom_right
start = (C2): 문자 입력
endpoint = (D7): 문자 입력
```

즉, 선택 범위는 C2:D7이었다.

## 구현 시 주의점

### 1. 선택 방향에 의존하지 말 것

앞에서 뒤로 긁든, 뒤에서 앞으로 긁든 결과는 같아야 한다. 따라서 “현재 위치에서 왼쪽으로 몇 칸” 같은 단순 추측으로 처리하면 안 된다.

반드시 가능한 행/열/끝점 모서리 후보를 검증해서 실제 범위를 찾아야 한다.

### 2. 붙여넣기 대화상자 방지

셀 선택 상태에서 `Paste`를 실행하면 한글이 덮어쓰기/내용만 덮어쓰기/셀 안에 표로 넣기 선택 대화상자를 띄울 수 있다.

대화상자 방지를 위해 다음 방식이 필요할 수 있다.

```python
option = hwp.HParameterSet.HSelectionOpt
hwp.HAction.GetDefault("Paste", option.HSet)
option.Option = 5
hwp.HAction.Execute("Paste", option.HSet)
```

`Option = 5`는 내용만 덮어쓰기 의도로 사용했다.

### 3. 셀 값 읽기

현재 셀을 읽는 임시 방식:

```text
TableCellBlock
Copy
클립보드 텍스트 읽기
Cancel
```

더 안정적인 방식으로는 `InitScan/GetText/ReleaseScan` 또는 현재 셀에 임시 필드명을 주고 `GetFieldText`로 읽는 방법이 후보지만, 아직 이 프로젝트에서는 검증하지 않았다.

### 4. 병합 셀

병합 셀이 포함된 선택 범위는 위 직사각형 후보 검증이 실패하거나 잘못된 후보를 만들 수 있다. 우선은 병합 셀이 없는 직사각형 범위를 대상으로 구현하는 것이 안전하다.

## 현재 결론

한글 표 셀 범위 변환은 다음 구조로 구현해야 한다.

```text
1. 선택 상태에서 Copy
2. selected_values 추출
3. endpoint = GetPos()
4. 가능한 rows/cols/corner 후보를 실제 셀 값과 대조
5. 완전 일치 후보를 실제 선택 범위로 확정
6. 시작 셀로 이동
7. 셀 하나씩 Copy -> 변환 -> Paste
8. 행 우선 순서로 다음 셀 이동
```

이 방식은 선택 방향에 덜 의존하고, 선택 범위 전체 붙여넣기에서 생긴 “각 셀에 전체 목록이 들어가는 문제”를 피할 수 있다.

## HancomDev PDF 확인 내용

로컬 `HancomDev` 폴더의 PDF 4개를 확인했다.

```text
ActionTable_2504.pdf
HwpAutomation_2504.pdf
ParameterSetTable_2504.pdf
한글오토메이션EventHandler추가_2504.pdf
```

### ActionTable_2504.pdf

표 셀 이동/선택 관련 액션이 명시되어 있다.

```text
ShapeObjTableSelCell - 테이블 선택상태에서 첫 번째 셀 선택하기
TableCellBlock - 셀 블록
TableCellBlockCol - 셀 블록(칸)
TableCellBlockExtend - 셀 블록 연장(F5 + F5)
TableCellBlockExtendAbs - 셀 블록 연장(SHIFT + F5)
TableCellBlockRow - 셀 블록(줄)
TableColBegin - 셀 이동: 열 시작
TableColEnd - 셀 이동: 열 끝
TableLeftCell - 셀 이동: 셀 왼쪽
TableLowerCell - 셀 이동: 셀 아래
TableRightCell - 셀 이동: 셀 오른쪽
TableRightCellAppend - 셀 이동: 셀 오른쪽에 이어서
TableUpperCell - 셀 이동: 셀 위
```

특히 `ShapeObjTableSelCell` 설명이 중요하다.

```text
테이블 선택상태에서 첫 번째 셀 선택하기
```

한컴 포럼 예제의 흐름과 일치한다.

```python
hwp.SetPosBySet(table_ctrl.GetAnchorPos(0))
hwp.Run("SelectCtrlReverse")
hwp.Run("ShapeObjTableSelCell")
```

즉, 표 컨트롤을 선택한 상태에서 `ShapeObjTableSelCell`을 실행하면 표의 첫 번째 셀로 들어갈 수 있다.

### HwpAutomation_2504.pdf

`CellShape` 설명에 따르면 현재 선택된 표/셀 모양 정보를 `ParameterSet/Table`로 얻을 수 있고, 그 안의 `"Cell"` 아이템이 셀 속성이다.

다만 이 정보는 셀의 모양/속성 중심이며, 선택 범위의 시작/끝 행열을 직접 주는 API로 확인되지는 않았다.

또 `CurFieldState`는 현재 캐럿 위치가 셀 필드인지 확인하는 데 사용할 수 있다.

```text
bit 0-3 = 필드 종류
1 = 셀
```

### ParameterSetTable_2504.pdf

`ListParaPos`는 다음 API에서 쓰는 커서 위치 파라미터다.

```text
HwpCtrl.GetPosBySet
HwpCtrl.SetPosBySet
HwpCtrlCode.GetAnchorPos
```

`TableCreation`에는 표 생성 시의 행/열 수가 있다.

```text
Rows
Cols
```

`TableSplitCell`에도 나누기용 행/열 수가 있다.

```text
Rows
Cols
```

하지만 기존 표에서 현재 선택 셀 범위의 행/열을 직접 반환하는 파라미터셋은 확인하지 못했다.

### 실용적 의미

PDF와 포럼을 합치면, 한컴이 제공하는 안정적인 접근 방식은 다음에 가깝다.

```text
표 컨트롤 anchor로 이동
-> SelectCtrlReverse
-> ShapeObjTableSelCell로 표 첫 셀 진입
-> 필요한 row/col만큼 TableLowerCell/TableRightCell 이동
```

따라서 “특정 표의 특정 row/col” 접근은 가능하다.

하지만 “사용자가 임의로 선택한 셀 블록의 시작 row/col과 크기”를 직접 읽는 API는 여전히 확인되지 않았다.

앞으로 구현 방향은 둘 중 하나다.

```text
1. 선택 범위를 직접 읽는 대신 사용자에게 행/열 수 또는 시작 방향 제약을 둔다.
2. 선택값 목록과 실제 표 값 대조로 범위를 역산하되, 탐색 과정이 사용자 문서를 어지럽히지 않도록 별도 복사본/숨김 문서 기반으로 검증한다.
```
