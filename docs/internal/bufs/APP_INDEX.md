# BUFS 한글 스타일 자동화 앱 인덱스

이 문서는 `bufs/` 폴더의 한글 스타일 자동화 MVP를 수정하거나 확장할 때 먼저 읽는 구조 안내서다. 루트에 흩어져 있던 내부 메모 문서는 정리 과정에서 `docs/internal/bufs/`로 옮겼다.

## 새 기능 개발 원칙

한글 자동화 기능을 새로 추가하거나 기존 COM/HAction 호출을 바꿀 때는 한컴 공식 개발자 문서 또는 신뢰 가능한 API 카탈로그를 먼저 확인한다. 한컴DEV PDF와 파생 카탈로그 파일은 공개 저장소에 포함하지 않는다.

- `HwpAutomation`: HWP 자동화 객체 모델, COM 진입점, 주요 객체/메서드 확인
- `ActionTable`: `HAction.Execute`, `Run` 등에 사용할 액션 이름 확인
- `ParameterSetTable`: 액션별 `HParameterSet` 이름과 속성 확인
- `한글오토메이션 EventHandler`: 이벤트 처리나 대화상자/상호작용 자동화가 필요한 경우 확인

구현 전에 관련 PDF에서 액션명, 파라미터셋, 속성명을 확인하고, 코드에는 추측한 이름보다 검증된 이름을 사용한다. 새 기능 검증 결과나 한글 COM 특이사항은 이 문서 또는 관련 설계/검증 문서에 남긴다.

## 실행 환경과 의존성

- OS: Windows
- 한글: Hancom HWP/HWPX를 열 수 있고 `HWPFrame.HwpObject` COM 객체가 등록된 환경
- Python: `bufs/.venv/Scripts/python.exe`를 사용하는 로컬 가상환경
- GUI: Python 표준 라이브러리 `tkinter`
- 한글 자동화: `pywin32`
  - `pythoncom`
  - `win32com.client`
  - `win32clipboard`
  - `win32con`
  - `win32gui`
- 표준 라이브러리 주요 사용처:
  - `zipfile`: HWPX 내부 XML/텍스트 읽기와 표지 템플릿 생성
  - `json`: 스타일 순서와 표 설정 저장
  - `re`: 텍스트 정리, 날짜/숫자/마크다운 표/캡션 파싱
  - `shutil`, `Path`, `datetime`, `time`: 파일 생성, 테스트 출력, COM 대기 처리
- 이미지/썸네일:
  - PNG/BMP/JPEG 크기 파싱은 자체 구현
  - 로고 썸네일 생성은 Windows WIA COM이 가능할 때만 동작

현재 별도 `requirements.txt`는 없다. 배포/재현성을 강화할 때는 `pywin32` 의존성을 명시하는 파일을 추가하는 것이 좋다.

## 실행 방법

- 일반 실행: `bufs/run-mvp.ps1`
  - `bufs/.venv/Scripts/python.exe`로 `bufs/hwp_style_mvp.py`를 실행한다.
- 환경 검증: `bufs/verify-readiness.ps1`
  - 한글 COM 연결, 주요 COM 멤버, 파라미터셋, 색상/선 변환, 템플릿 구조를 확인한다.
- 앱 내 스모크 테스트: `python bufs/hwp_style_mvp.py --smoke-test`
  - 코드 하단의 `smoke_test()`가 지원하는 경우 사용한다.

## 파일 구조

```text
bufs/
  hwp_style_mvp.py              # Tkinter GUI와 한글 COM 자동화 메인 앱
  run-mvp.ps1                   # 로컬 가상환경 기반 실행 스크립트
  verify-readiness.ps1          # 한글 COM/템플릿 사전 검증 스크립트
  table-settings.json           # 팔레트, 선 프리셋, 표 제목셀, 여백, 단축키 설정
  style-sets.json               # 스타일 세트, 문단/글자 구분, 표/캡션 플래그
  style-order.json              # 구버전 스타일 목록 표시 순서 호환용
  templates/
    보고서 본문 서식.hwpx        # 스타일 기준 문서
    표지.hwpx                   # 표지 생성 템플릿
    표지.hwp                    # 표지 호환/원본 파일
  logos/                        # 삽입 가능한 BUFS 로고 이미지
  test-output/                  # 테스트/생성 산출물
docs/
  internal/
    bufs/
      APP_INDEX.md
      implementation-plan.md
      verification-results.md
      mvp-test-results.md
      hwp-table-cell-selection-notes.md
      demo-cases.md
      리팩터링-기능명세서.md
```

## 데이터와 설정 흐름

- `STYLE_FILE = bufs/templates/보고서 본문 서식.hwpx`
  - 스타일 목록과 스타일 ID를 읽는 기준 파일이다.
  - 앱은 현재 문서에 같은 스타일 이름이 있는지 찾아 적용한다.
- `COVER_FILE = bufs/templates/표지.hwpx`
  - `{{문서제목}}`, `{{날짜}}`, `{{부서명}}` 같은 표지 플레이스홀더를 채워 새 표지 파일을 만든다.
- `LOGO_DIR = bufs/logos`
  - GUI 로고 탭에서 이미지 후보를 표시하고 현재 커서 위치에 삽입한다.
- `STYLE_ORDER_FILE = bufs/style-order.json`
  - 사용자가 스타일 버튼/목록 순서를 바꾸면 저장한다.
- `TABLE_SETTINGS_FILE = bufs/table-settings.json`
  - 표 편집 팔레트, 선, 제목셀, 여백, 캡션 파서, 단축키 설정을 저장한다.
- `TEST_OUTPUT_DIR = bufs/test-output`
  - 표지 생성 파일, 스타일 테스트 복사본, 로고 썸네일 같은 산출물을 저장한다.

## 메인 코드 구조

### 데이터 클래스

- `PaletteColor`: UI와 한글 COM 색상 변환에 쓰는 팔레트 항목
- `StyleRecord`: 스타일 ID, 스타일 타입, 스타일 이름
- `HwpCandidate`: 연결 가능한 한글 인스턴스 후보 정보
- `CaptionParts`: 표 캡션 조립을 위한 prefix/suffix 분리 결과

### 설정/템플릿 읽기

- `read_zip_text`: HWPX(zip) 내부 엔트리 텍스트 읽기
- `read_style_records`, `read_style_names`: 스타일 템플릿에서 스타일 목록 추출
- `load_style_order`, `save_style_order`, `apply_saved_style_order`: 스타일 표시 순서 관리
- `merge_dict`, `normalize_table_settings`, `load_table_settings`, `save_table_settings`: 표 설정 기본값 병합과 저장
- `read_cover_placeholders`: 표지 템플릿의 플레이스홀더 존재 여부 확인

### 표지/로고 처리

- `parse_cover_input`: 선택한 3줄을 제목/날짜/부서명으로 파싱
- `replace_marker_with_paragraphs`, `collapse_cover_paragraph_after`, `remove_confidential_cover_table`: 표지 XML 수정
- `create_filled_cover_file`: 채워진 표지 HWPX 생성
- `list_logos`, `read_image_pixel_size`, `create_logo_thumbnail`: 로고 후보와 미리보기 처리
- `insert_file_at_cursor`, `insert_logo_at_cursor`, `insert_picture_at_cursor`: 현재 한글 문서 커서 위치에 파일/이미지 삽입

### 텍스트 정리 유틸

- `add_thousand_commas`, `remove_number_commas`: 숫자 쉼표 추가/제거
- `normalize_dates`, `add_weekdays_to_dates`, `remove_weekdays_from_dates`: 날짜 표기와 요일 처리
- `clean_manual_line_breaks`: 수동 줄바꿈 정리
- `remove_manual_outline_prefix`, `remove_manual_outline_prefixes`: 수동 개요번호 제거
- `parse_markdown_table`, `markdown_rows_to_html_table`, `rows_to_tsv`: 마크다운 표를 한글에 붙여넣을 수 있는 표 형태로 변환

### 한글 COM 연결과 대상 선택

- `connect_hwp`: `HWPFrame.HwpObject`에 연결하고 보안 모듈 등록을 시도한다.
- `get_foreground_title`, `list_visible_window_titles`: 현재 Windows 창 제목 확인
- `list_hwp_candidates`: 실행 중인 한글 문서 후보를 수집한다.
- `describe_hwp`: 연결된 한글 인스턴스 상태를 로그용으로 요약한다.
- `MvpApp.ensure_hwp`: 앱 동작 전에 한글 연결을 보장한다.

### GUI 앱 `MvpApp`

`MvpApp(tk.Tk)`가 전체 GUI와 사용자 액션을 소유한다.

- 탭 구성:
  - `_build_styles_tab`: 스타일 적용과 텍스트 정리 기능
  - `_build_table_tab`, `build_table_tab_content`: 표 편집 기능
  - `_build_cover_logo_tab`: 표지 생성과 로고 삽입 기능
  - `_build_status_tab`: 연결 확인, 대상 확인, 테스트 문서 열기, 로그
- 로그/상태:
  - `log`, `debug`, `refresh_all`, `check_hwp`, `show_current_target`
- 스타일:
  - `populate_style_list`, `move_selected_style`, `reset_style_order`
  - `refresh_current_doc_style_map`, `apply_selected_style`, `apply_style_by_name`, `apply_first_style_containing`
- 색상/셀:
  - `apply_cell_fill_color`, `clear_cell_fill_color`, `set_cell_fill_color`, `set_cell_fill_none`
  - `apply_text_color`, `set_text_color`
- 표 테두리/제목셀:
  - `line_type_value`, `line_width_value`, `mm_to_hwpunit`
  - `set_border_edge`, `set_inside_border_from_preset`, `set_table_border_preset`
  - `apply_table_border_preset`, `apply_title_cell_preset`, `apply_table_outside_margins`
- 캡션:
  - `split_table_caption_parts`
  - `cut_selected_table_caption_title`
  - `apply_pending_table_caption_to_next_table`
  - `cut_title_and_apply_to_next_table_caption`
- 클립보드/붙여넣기:
  - `get_clipboard_text`, `set_clipboard_text`, `set_clipboard_html_table`
  - `paste_text_into_selected_cell`, `paste_with_selection_option`, `paste_html_original_format`
- 선택 셀 반복 처리:
  - `detect_selected_cell_rect`
  - `transform_selected_cells_by_iteration`
  - `transform_selected_text`
- 사용자 버튼 액션:
  - `convert_selected_markdown_table`
  - `clean_selected_line_breaks`
  - `clean_selected_outline_prefixes`
  - `add_commas_to_selection`, `remove_commas_from_selection`
  - `normalize_dates_to_korean`, `normalize_dates_to_dot`, `normalize_dates_to_dot_padded`
  - `add_weekdays_to_selection`, `remove_weekdays_from_selection`
  - `clean_excel_table`

## 기능 추가 체크리스트

1. 한컴 공식 개발자 문서 또는 검증된 API 카탈로그에서 필요한 액션명, 파라미터셋, 속성명을 확인한다.
2. 기존 `MvpApp`의 한글 COM 헬퍼를 먼저 재사용한다.
   - `run_hwp_command`
   - `execute_first_hwp_action`
   - `set_com_attr`
   - `set_hset_item`
   - `set_parameter_item`
3. 설정이 필요한 기능은 하드코딩보다 `table-settings.json` 또는 별도 JSON 설정으로 분리할지 검토한다.
4. HWP 원본 파일을 직접 덮어쓰지 말고 `test-output/`에 복사본이나 생성물을 만든 뒤 검증한다.
5. COM 동작은 실패 가능성이 높으므로 로그에 액션명, 파라미터셋, 결과값을 남긴다.
6. UI 버튼을 추가할 때는 관련 동작을 작은 메서드로 분리하고, 실패 시 사용자가 확인할 수 있는 메시지를 제공한다.
7. 기능 검증 후 `verification-results.md` 또는 관련 문서에 테스트 조건과 결과를 기록한다.

## 주의할 점

- 한글 COM은 설치 버전, 보안 모듈 등록, 현재 선택 상태에 따라 같은 액션도 실패할 수 있다.
- 표/셀 관련 기능은 현재 커서 위치, 셀 선택 범위, 표 개체 선택 여부에 민감하다.
- 클립보드 기반 기능은 한글 포커스와 Windows 클립보드 상태에 영향을 받는다.
- `bufs/.venv`, `__pycache__`, 생성된 `test-output/*.hwpx`는 커밋하지 않는다.
- `test-output/logo-thumbs/*.png`는 현재 레포에서 추적 중인 파일이므로 일괄 삭제/ignore 정책을 바꿀 때 주의한다.

