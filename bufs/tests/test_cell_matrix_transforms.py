import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import hwp_style_mvp  # noqa: E402
from hwp_style_mvp import (  # noqa: E402
    add_thousand_commas,
    add_weekdays_to_dates,
    cell_matrix_size,
    cell_matrix_to_tsv,
    clean_manual_line_breaks,
    flatten_cell_matrix,
    is_rectangular_cell_matrix,
    looks_like_cell_clipboard_matrix,
    manual_outline_prefix_length,
    normalize_dates,
    normalize_clipboard_newlines,
    parse_cell_clipboard_matrix,
    remove_manual_outline_prefixes,
    remove_number_commas,
    remove_weekdays_from_dates,
    parse_cell_addresses_from_formula_command,
    parse_cell_address_from_keyindicator,
    looks_like_single_copied_cell,
    TABLE_BORDER_ICON_BUTTON_SIZE,
    TABLE_BORDER_ICON_PRESETS,
    TABLE_ICON_BUTTON_BG,
    TABLE_ICON_BUTTON_GAP,
    TABLE_STYLE_ICON_BUTTON_SIZE,
    TABLE_STYLE_ICON_PRESETS,
    StyleRecord,
    read_style_records,
    read_style_records_from_hwpml,
    transform_cell_matrix,
    MvpApp,
)


class CellMatrixTransformTests(unittest.TestCase):
    def test_table_icon_presets_match_expected_grid_sizes_and_files(self) -> None:
        self.assertEqual(len(TABLE_STYLE_ICON_PRESETS), 4)
        self.assertEqual(len(TABLE_BORDER_ICON_PRESETS), 6)
        self.assertEqual(TABLE_STYLE_ICON_BUTTON_SIZE, 48)
        self.assertEqual(TABLE_BORDER_ICON_BUTTON_SIZE, 48)
        self.assertEqual(TABLE_ICON_BUTTON_GAP, 6)
        self.assertEqual(TABLE_ICON_BUTTON_BG, "#FFFFFF")
        for icon_name, _preset_name, _tooltip in TABLE_STYLE_ICON_PRESETS + TABLE_BORDER_ICON_PRESETS:
            self.assertTrue((ROOT / "icons" / "2x" / f"{icon_name}@2x.png").exists())

    def test_style_records_keep_document_order_index_separate_from_xml_id(self) -> None:
        header = """<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:styles>
    <hh:style id="0" type="PARA" name="바탕글"/>
    <hh:style id="99" type="PARA" name="보고서본문"/>
    <hh:style id="3" type="CHAR" name="강조"/>
  </hh:styles>
</hh:head>
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "styles.hwpx"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("Contents/header.xml", header)

            records = {record.name: record for record in read_style_records(path)}

        self.assertEqual(records["보고서본문"].style_id, 99)
        self.assertEqual(records["보고서본문"].style_index, 1)

    def test_hwpml_style_records_use_runtime_style_ids(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<HWPML>
  <HEAD>
    <STYLE Id="95" Name="xl65" EngName="xl65"/>
    <STYLE Id="97" Name="표내용-중간"/>
  </HEAD>
</HWPML>
"""
        records = {record.name: record for record in read_style_records_from_hwpml(xml)}

        self.assertEqual(records["표내용-중간"].style_id, 97)
        self.assertEqual(records["표내용-중간"].style_index, 1)

    def test_execute_style_record_applies_current_document_style_id_after_name_match(self) -> None:
        class FakeStyleSet:
            HSet = object()

            def __init__(self) -> None:
                self.Apply = -1

        class FakeAction:
            def __init__(self, pset: FakeStyleSet) -> None:
                self.pset = pset

            def GetDefault(self, name, hset):
                return name == "Style"

            def Execute(self, name, hset):
                return name == "Style" and self.pset.Apply == 99

        pset = FakeStyleSet()
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"HParameterSet": type("Sets", (), {"HStyle": pset})(), "HAction": FakeAction(pset)})()
        app.debug = lambda _message: None

        ok = app.execute_style_record(StyleRecord(style_id=99, style_index=1, style_type="PARA", name="보고서본문"))

        self.assertTrue(ok)
        self.assertEqual(pset.Apply, 99)

    def test_apply_style_by_name_stops_when_current_document_name_match_is_missing(self) -> None:
        app = object.__new__(MvpApp)
        app.refresh_current_doc_style_map = lambda *, force=False: None
        app.find_current_doc_style_record = lambda _name: None
        app.debug = lambda _message: None
        app.execute_style_record = lambda _record: self.fail("이름 매칭 실패 시 스타일 번호 fallback을 쓰면 안 됩니다")
        warnings = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))

        try:
            self.assertFalse(app.apply_style_by_name("보고서본문"))
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertEqual(len(warnings), 1)
        self.assertIn("보고서본문", warnings[0][1])

    def test_apply_first_style_containing_does_not_fallback_to_base_style_records(self) -> None:
        app = object.__new__(MvpApp)
        app.refresh_current_doc_style_map = lambda *, force=False: None
        app.current_doc_style_map = {}
        app.style_records = [StyleRecord(style_id=99, style_index=1, style_type="PARA", name="[표/그림] 캡션")]
        app.debug = lambda _message: None
        app.execute_style_record = lambda _record: self.fail("현재 문서 스타일 맵 없이 기준 번호 fallback을 쓰면 안 됩니다")

        self.assertFalse(app.apply_first_style_containing("캡션"))

    def test_refresh_current_doc_style_map_prefers_hwp_memory_styles(self) -> None:
        class FakeHwp:
            Path = "C:/temp/current.hwpx"

            def GetTextFile(self, format_name, option):
                self.last_request = (format_name, option)
                return """<HWPML><HEAD>
                    <STYLE Id="95" Name="xl65"/>
                    <STYLE Id="97" Name="표내용-중간"/>
                </HEAD></HWPML>"""

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.current_doc_style_path = None
        app.current_doc_style_map = {}
        app.current_doc_style_norm_map = {}
        app.style_records = [StyleRecord(style_id=33, style_index=31, style_type="PARA", name="표내용-중간")]
        app.debug = lambda _message: None

        app.refresh_current_doc_style_map(force=True)

        self.assertEqual(app.find_current_doc_style_record("표내용-중간").style_id, 97)

    def test_set_selected_cells_as_header_commits_cell_shape_table_header(self) -> None:
        class FakeSet:
            def __init__(self) -> None:
                self.items = {}

            def SetItem(self, name, value):
                self.items[name] = value

        class FakeCell:
            def __init__(self) -> None:
                self.HSet = FakeSet()

        class FakeCellShape:
            def __init__(self) -> None:
                self.Cell = FakeCell()

        class FakeShapeObject:
            def __init__(self) -> None:
                self.HSet = FakeSet()
                self.ShapeTableCell = FakeCell()

        class FakeAction:
            def __init__(self) -> None:
                self.defaults = []
                self.executed = []

            def GetDefault(self, action, hset):
                self.defaults.append(action)
                return action == "TablePropertyDialog"

            def Execute(self, action, hset):
                self.executed.append(action)
                return action == "TablePropertyDialog"

        shape_object = FakeShapeObject()
        action = FakeAction()
        app = object.__new__(MvpApp)
        app.hwp = type(
            "FakeHwp",
            (),
            {
                "CellShape": FakeCellShape(),
                "HParameterSet": type("Sets", (), {"HShapeObject": shape_object})(),
                "HAction": action,
            },
        )()
        app.debug = lambda _message: None

        action_name, set_ok, action_ok = app.set_selected_cells_as_header()

        self.assertEqual(action_name, "CellShape")
        self.assertTrue(set_ok)
        self.assertTrue(action_ok)
        self.assertEqual(app.hwp.CellShape.RepeatHeader, 1)
        self.assertEqual(app.hwp.CellShape.Cell.HSet.items["Header"], 1)
        self.assertEqual(action.executed, [])

    def test_parse_tsv_matrix_preserves_empty_cells(self) -> None:
        matrix = parse_cell_clipboard_matrix("A\t\tC\r\n1\t2\t")

        self.assertEqual(matrix, [["A", "", "C"], ["1", "2", ""]])
        self.assertTrue(is_rectangular_cell_matrix(matrix))
        self.assertEqual(cell_matrix_size(matrix), (2, 3))

    def test_parse_single_column_matrix(self) -> None:
        matrix = parse_cell_clipboard_matrix("1234\r\n5678\r\n")

        self.assertEqual(matrix, [["1234"], ["5678"]])
        self.assertTrue(is_rectangular_cell_matrix(matrix))
        self.assertTrue(looks_like_cell_clipboard_matrix("1234\r\n5678\r\n"))

    def test_parse_formula_command_addresses(self) -> None:
        addresses = parse_cell_addresses_from_formula_command("=(C2,D2,C3,D3,C4,D4,C5,D5)")

        self.assertEqual(addresses, [(3, 2), (4, 2), (3, 3), (4, 3), (3, 4), (4, 4), (3, 5), (4, 5)])

    def test_parse_keyindicator_cell_address(self) -> None:
        self.assertEqual(parse_cell_address_from_keyindicator("(D5): 표 안 입력"), (4, 5))

    def test_multiline_dates_are_cell_like_for_single_column_selection(self) -> None:
        self.assertTrue(looks_like_cell_clipboard_matrix("2026.7.22\r\n2026-07-23\r\n"))

    def test_rejects_non_rectangular_matrix(self) -> None:
        matrix = parse_cell_clipboard_matrix("A\tB\r\nC")

        self.assertFalse(is_rectangular_cell_matrix(matrix))

    def test_matrix_to_tsv_round_trip_shape(self) -> None:
        matrix = [["A", "", "C"], ["1", "2", ""]]

        self.assertEqual(cell_matrix_to_tsv(matrix), "A\t\tC\r\n1\t2\t")

    def test_flatten_cell_matrix_preserves_rectangular_order(self) -> None:
        matrix = [[" A ", ""], ["1", "2"]]

        self.assertEqual(flatten_cell_matrix(matrix), ["A", "", "1", "2"])

    def test_clipboard_newlines_are_crlf_normalized(self) -> None:
        self.assertEqual(normalize_clipboard_newlines("첫 문단\n둘째 문단"), "첫 문단\r\n둘째 문단")
        self.assertEqual(normalize_clipboard_newlines("첫 문단\r\n둘째 문단"), "첫 문단\r\n둘째 문단")

    def test_add_commas_cell_by_cell(self) -> None:
        matrix = [["1234", "2026-07-22"], ["1234567", "051-123-4567"]]

        transformed, changed = transform_cell_matrix(matrix, add_thousand_commas)

        self.assertEqual(transformed, [["1,234", "2026-07-22"], ["1,234,567", "051-123-4567"]])
        self.assertEqual(changed, 2)

    def test_remove_commas_cell_by_cell(self) -> None:
        matrix = [["1,234", "1,234,567"]]

        transformed, changed = transform_cell_matrix(matrix, remove_number_commas)

        self.assertEqual(transformed, [["1234", "1234567"]])
        self.assertEqual(changed, 2)

    def test_normalize_dates_cell_by_cell(self) -> None:
        matrix = [["2026.7.22", "2026-07-22"]]

        transformed, changed = transform_cell_matrix(matrix, lambda text: normalize_dates(text, "dot_padded"))

        self.assertEqual(transformed, [["2026. 07. 22.", "2026. 07. 22."]])
        self.assertEqual(changed, 2)

    def test_weekday_add_remove_cell_by_cell(self) -> None:
        matrix = [["2026. 7. 22.", "2026. 7. 22.(수)"]]

        added, add_changed = transform_cell_matrix(matrix, add_weekdays_to_dates)
        removed, remove_changed = transform_cell_matrix(added, remove_weekdays_from_dates)

        self.assertEqual(added, [["2026. 7. 22.(수)", "2026. 7. 22.(수)"]])
        self.assertEqual(add_changed, 1)
        self.assertEqual(removed, [["2026. 7. 22.", "2026. 7. 22."]])
        self.assertEqual(remove_changed, 2)

    def test_outline_prefix_removed_cell_by_cell(self) -> None:
        matrix = [["  ◦ 본문", "- 항목"], ["본문", "∙내용"]]

        transformed, changed = transform_cell_matrix(
            matrix,
            remove_manual_outline_prefixes,
            strip_wrapping_lines=True,
        )

        self.assertEqual(transformed, [["본문", "항목"], ["본문", "내용"]])
        self.assertEqual(changed, 3)

    def test_manual_outline_prefix_length_includes_leading_spaces_and_marker(self) -> None:
        text = "  ○ 본문"

        self.assertEqual(manual_outline_prefix_length(text), len("  ○ "))
        self.assertEqual(text[manual_outline_prefix_length(text):], "본문")

    def test_line_break_cleanup_cell_by_cell(self) -> None:
        matrix = [["문장 중간\n줄바꿈", "항목:\n- 유지"]]

        transformed, changed = transform_cell_matrix(matrix, clean_manual_line_breaks)

        self.assertEqual(transformed, [["문장 중간 줄바꿈", "항목:\n- 유지"]])
        self.assertEqual(changed, 1)

    def test_unchanged_matrix_reports_zero_changes(self) -> None:
        matrix = [["본문", "2026년"]]

        transformed, changed = transform_cell_matrix(matrix, add_thousand_commas)

        self.assertEqual(transformed, matrix)
        self.assertEqual(changed, 0)

    def test_selection_mode_text_is_not_cell_block(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.debug = lambda _message: None

        self.assertFalse(app.is_selected_cell_block())

    def test_selection_mode_cells_is_cell_block(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None

        self.assertTrue(app.is_selected_cell_block())

    def test_single_copied_cell_without_formula_addresses_uses_text_path(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_selected_cell_range_by_formula = lambda: None

        self.assertTrue(looks_like_single_copied_cell("2021. 6. 8.\r\n"))
        self.assertFalse(
            app.transform_selected_cell_matrix(
                "요일 추가",
                "2021. 6. 8.\r\n",
                add_weekdays_to_dates,
            )
        )

    def test_cell_matrix_transform_stops_when_cell_address_snapshot_is_missing(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 99, "cols": 99}
        app.detect_selected_cell_rect = lambda *args, **kwargs: self.fail("주소 없을 때 셀 감지 스캔을 실행하면 안 됩니다")
        app.transform_selected_cells_by_iteration = lambda *args, **kwargs: self.fail("주소 없을 때 한 열 순회를 실행하면 안 됩니다")
        app.set_clipboard_text = lambda _text: self.fail("주소 없을 때 TSV 붙여넣기를 실행하면 안 됩니다")
        warnings: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))

        try:
            handled = app.transform_selected_cell_matrix(
                "날짜 정규화",
                "2021. 6. 8.\t2019. 4. 2.",
                lambda text: normalize_dates(text, "dot_padded"),
            )
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertTrue(handled)
        self.assertTrue(warnings)
        self.assertIn("테이블 전체를 추측해서 훑지 않습니다", warnings[0][1])

    def test_formula_address_iteration_runs_before_detection_scan(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.get_selected_cell_range_by_formula = lambda: {
            "rows": 4,
            "cols": 2,
            "addresses": [(3, 2), (4, 2), (3, 3), (4, 3), (3, 4), (4, 4), (3, 5), (4, 5)],
        }

        called: dict[str, object] = {}

        def fake_formula_iteration(label, addresses, transform, *, strip_wrapping_lines=False):
            called["label"] = label
            called["addresses"] = addresses
            return 8, 8

        app.transform_formula_address_cells = fake_formula_iteration
        app.detect_selected_cell_rect = lambda *args, **kwargs: self.fail("address path should run before detection scan")

        handled = app.transform_selected_cell_matrix(
            "요일 제거",
            "\r\n".join(str(index) for index in range(8)),
            remove_weekdays_from_dates,
        )

        self.assertTrue(handled)
        self.assertEqual(called["addresses"][0], (3, 2))
        self.assertEqual(called["addresses"][-1], (4, 5))

    def test_formula_address_failure_does_not_fall_back_to_detection_scan(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_selected_cell_range_by_formula = lambda: {
            "rows": 4,
            "cols": 2,
            "addresses": [(3, 2), (4, 2), (3, 3), (4, 3), (3, 4), (4, 4), (3, 5), (4, 5)],
        }
        app.transform_formula_address_cells = lambda *args, **kwargs: None
        app.detect_selected_cell_rect = lambda *args, **kwargs: self.fail("formula address failure must not scan")

        warnings: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))
        try:
            handled = app.transform_selected_cell_matrix(
                "요일 제거",
                "\r\n".join(str(index) for index in range(8)),
                remove_weekdays_from_dates,
            )
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertTrue(handled)
        self.assertTrue(warnings)
        self.assertIn("다른 셀을 추측해서 훑지 않습니다", warnings[0][1])

    def test_transform_text_captures_cell_range_before_copy(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: None
        app.get_clipboard_text = lambda: "\r\n".join(str(index) for index in range(8))
        app.debug = lambda _message: None
        app.log = lambda _message: None

        ranges = [
            {
                "rows": 4,
                "cols": 2,
                "addresses": [(3, 2), (4, 2), (3, 3), (4, 3), (3, 4), (4, 4), (3, 5), (4, 5)],
            }
        ]
        app.get_selected_cell_range_by_formula = lambda: ranges.pop(0)
        app.run_hwp_command = lambda command: command == "Copy"

        received: dict[str, object] = {}

        def fake_matrix(label, text, transform, *, strip_wrapping_lines=False, range_info=None):
            received["range_info"] = range_info
            return True

        app.transform_selected_cell_matrix = fake_matrix

        app.transform_selected_text(
            "요일 제거",
            remove_weekdays_from_dates,
            allow_cell_iteration=True,
        )

        self.assertEqual(received["range_info"]["rows"], 4)
        self.assertEqual(received["range_info"]["cols"], 2)
        self.assertEqual(received["range_info"]["addresses"][0], (3, 2))

    def test_transform_text_does_not_treat_post_copy_cell_state_as_selected_block(self) -> None:
        class FakeHwp:
            SelectionMode = 0

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: None
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_clipboard_text = lambda: "2022. 9. 1."
        app.run_hwp_command = lambda command: command in {"Copy", "Paste", "TableCellBlock"}
        app.transform_selected_cell_matrix = lambda *args, **kwargs: self.fail("셀 블록이 아닌 상태에서는 셀 행렬 경로를 타면 안 됩니다")
        written: dict[str, str] = {}
        app.set_clipboard_text = lambda text: written.setdefault("text", text)
        app.activate_hwp_window = lambda: None

        app.transform_selected_text(
            "요일 추가",
            add_weekdays_to_dates,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

        self.assertEqual(written["text"], "2022. 9. 1.(목)")

    def test_formula_address_iteration_moves_from_current_cell_to_each_formula_address(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        addresses = iter([(4, 5), (3, 2), (4, 2)])
        app.get_current_cell_address = lambda: next(addresses)
        moves: list[tuple[tuple[int, int], tuple[int, int]]] = []

        def fake_move(current, target):
            moves.append((current, target))
            return True

        app.move_between_table_addresses = fake_move
        values = iter(["2022년 9월 1일(목)", "2022. 9. 1."])
        app.read_current_cell_text = lambda: next(values)
        app.select_current_table_cell_for_replace = lambda: True
        app.set_clipboard_text = lambda _text: None
        app.paste_text_into_selected_cell = lambda: True
        app.clear_hwp_selection = lambda: True
        positions = iter(["pos-c2", "pos-d2"])
        restored: list[str] = []
        app.get_hwp_pos_by_set = lambda: next(positions)
        app.set_hwp_pos_by_set = lambda pos: restored.append(pos) or True
        app.log = lambda _message: None

        result = app.transform_formula_address_cells(
            "요일 제거",
            [(3, 2), (4, 2)],
            remove_weekdays_from_dates,
        )

        self.assertEqual(result, (2, 1))
        self.assertEqual(moves, [((4, 5), (3, 2)), ((3, 2), (4, 2))])
        self.assertEqual(restored, ["pos-c2", "pos-c2", "pos-d2", "pos-d2"])

    def test_formula_address_iteration_uses_snapshotted_positions_after_unchanged_cell(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.clear_hwp_selection = lambda: True

        snapshot_addresses = iter([(3, 2), (4, 2), (3, 3)])
        app.get_current_cell_address = lambda: next(snapshot_addresses)
        moves: list[tuple[tuple[int, int], tuple[int, int]]] = []
        transforming = False

        def fake_move(current, target):
            self.assertFalse(transforming, "변환 중에는 상대 셀 이동을 다시 쓰면 안 됩니다")
            moves.append((current, target))
            return True

        app.move_between_table_addresses = fake_move
        positions = iter(["pos-c2", "pos-d2", "pos-c3"])
        app.get_hwp_pos_by_set = lambda: next(positions)
        restored: list[str] = []

        def fake_set_pos(pos):
            restored.append(pos)
            return True

        app.set_hwp_pos_by_set = fake_set_pos
        values = iter(["2022. 9. 1.", "이미 본문", "2021. 6. 8."])

        def fake_read():
            nonlocal transforming
            transforming = True
            return next(values)

        app.read_current_cell_text = fake_read
        written: list[str] = []
        app.select_current_table_cell_for_replace = lambda: True
        app.set_clipboard_text = written.append
        app.paste_text_into_selected_cell = lambda: True

        result = app.transform_formula_address_cells(
            "요일 추가",
            [(3, 2), (4, 2), (3, 3)],
            add_weekdays_to_dates,
        )

        self.assertEqual(result, (3, 2))
        self.assertEqual(moves, [((3, 2), (4, 2)), ((4, 2), (3, 3))])
        self.assertEqual(restored, ["pos-c2", "pos-c2", "pos-d2", "pos-d2", "pos-c3", "pos-c3"])
        self.assertEqual(written, ["2022. 9. 1.(목)", "2021. 6. 8.(화)"])

    def test_addressless_multi_column_selection_does_not_try_detection_endpoints(self) -> None:
        class FakeHwp:
            SelectionMode = 19

            def GetPos(self):
                return (99, 0, 0)

            def GetSelectedPos(self):
                return (True, 10, 0, 0, 20, 0, 0)

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 4, "cols": 2}

        tried: list[tuple[int, int, int]] = []

        def fake_detect(endpoint, values, preferred_shape=None, only_preferred_shape=False):
            tried.append(endpoint)
            if endpoint == (20, 0, 0):
                return {"rows": 4, "cols": 2, "corner": "bottom_right", "start_label": "C2"}
            return None

        app.detect_selected_cell_rect = fake_detect
        app.transform_detected_cell_rect = lambda *args, **kwargs: (8, 8)
        app.set_hwp_pos = lambda _pos: True
        app.set_clipboard_text = lambda _text: self.fail("TSV paste path should not be used")
        app.run_hwp_command = lambda _command: self.fail("TSV paste path should not be used")
        warnings: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))

        try:
            handled = app.transform_selected_cell_matrix(
                "요일 제거",
                "\r\n".join(str(index) for index in range(8)),
                remove_weekdays_from_dates,
            )
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertTrue(handled)
        self.assertEqual(tried, [])
        self.assertTrue(warnings)

    def test_detection_accepts_preferred_shape_with_column_major_clipboard_values(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.set_hwp_pos = lambda _pos: True
        app.clear_hwp_selection = lambda: True
        app.move_table_cell = lambda _command, count=1: True

        reads = iter(
            [
                "2022년 9월 1일(목)",
                "2022. 9. 1.",
                "2021년 6월 8일(화)",
                "2021. 6. 8.",
                "2019년 4월 2일(화)",
                "2019. 4. 2.",
                "2026년 3월 18일(수)",
                "2026. 3. 18.",
            ]
        )
        app.read_current_cell_text = lambda: next(reads)

        column_major_clipboard = [
            "2022년 9월 1일(목)",
            "2021년 6월 8일(화)",
            "2019년 4월 2일(화)",
            "2026년 3월 18일(수)",
            "2022. 9. 1.",
            "2021. 6. 8.",
            "2019. 4. 2.",
            "2026. 3. 18.",
        ]

        rect = app.detect_selected_cell_rect((10, 0, 0), column_major_clipboard, (4, 2))

        self.assertIsNotNone(rect)
        self.assertEqual(rect["rows"], 4)
        self.assertEqual(rect["cols"], 2)
        self.assertEqual(rect["match"], "unordered")

    def test_detection_prefers_multi_column_shape_without_formula(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.set_hwp_pos = lambda _pos: True
        app.clear_hwp_selection = lambda: True
        app.move_table_cell = lambda _command, count=1: True

        reads = iter(["A", "B", "C", "D", "E", "F", "G", "H"])
        app.read_current_cell_text = lambda: next(reads)

        rect = app.detect_selected_cell_rect((10, 0, 0), ["A", "B", "C", "D", "E", "F", "G", "H"])

        self.assertIsNotNone(rect)
        self.assertGreater(rect["cols"], 1)

    def test_multiline_number_text_selection_is_not_blocked_as_cells(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: None
        app.run_hwp_command = lambda command: command in {"Copy", "Paste"}
        app.get_clipboard_text = lambda: "1234567\r\n-1234567\r\n1234567.89"
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None

        written: dict[str, str] = {}
        app.set_clipboard_text = lambda text: written.setdefault("text", text)

        app.transform_selected_text(
            "천원단위 쉼표 넣기",
            add_thousand_commas,
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
        )

        self.assertEqual(written["text"], "1,234,567\r\n-1,234,567\r\n1,234,567.89")

    def test_paragraph_outline_cleanup_excludes_zero_width_end_paragraph_and_avoids_paste(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 3), (0, 5, 0))
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.clear_current_paragraph_heading = lambda: False
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        messages: list[tuple[str, str]] = []
        original_showinfo = hwp_style_mvp.messagebox.showinfo
        hwp_style_mvp.messagebox.showinfo = lambda title, message: messages.append((title, message))
        try:
            visited_positions: list[tuple[int, int, int]] = []
            app.set_hwp_pos = lambda pos: visited_positions.append(pos) or True
            paragraphs = {2: "○ 첫 문단", 3: "본문", 4: "  - 셋째 문단", 5: "○ 제외 문단"}
            app.read_current_paragraph_text = lambda _list_id, para: paragraphs[para]
            deleted: list[tuple[int, int, int, int]] = []
            app.delete_hwp_text_range = lambda list_id, para, start, end: deleted.append((list_id, para, start, end)) or True
            app.run_hwp_command = lambda command: self.fail(f"문단 개요 제거는 {command} 명령을 직접 호출하지 않아야 합니다")
            app.clean_selected_paragraph_outline_prefixes()
        finally:
            hwp_style_mvp.messagebox.showinfo = original_showinfo

        self.assertEqual(deleted, [(0, 2, 0, len("○ ")), (0, 4, 0, len("  - "))])
        self.assertNotIn((0, 5, 0), visited_positions)
        self.assertEqual(restored, ["original"])
        self.assertEqual(messages, [])


if __name__ == "__main__":
    unittest.main()
