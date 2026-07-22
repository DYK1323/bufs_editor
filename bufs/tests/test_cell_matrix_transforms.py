import sys
import unittest
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
    normalize_dates,
    normalize_clipboard_newlines,
    parse_cell_clipboard_matrix,
    remove_manual_outline_prefixes,
    remove_number_commas,
    remove_weekdays_from_dates,
    parse_cell_addresses_from_formula_command,
    parse_cell_address_from_keyindicator,
    transform_cell_matrix,
    MvpApp,
)


class CellMatrixTransformTests(unittest.TestCase):
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

    def test_cell_matrix_transform_continues_when_formula_size_is_stale(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        messages: list[str] = []
        app.debug = messages.append
        app.log = messages.append
        app.activate_hwp_window = lambda: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 99, "cols": 99}
        app.run_hwp_command = lambda command: command == "Paste"

        written: dict[str, str] = {}
        app.set_clipboard_text = lambda text: written.setdefault("text", text)

        handled = app.transform_selected_cell_matrix(
            "날짜 정규화",
            "2021. 6. 8.\t2019. 4. 2.",
            lambda text: normalize_dates(text, "dot_padded"),
        )

        self.assertTrue(handled)
        self.assertEqual(written["text"], "2021. 06. 08.\t2019. 04. 02.")
        self.assertTrue(any("클립보드 행렬 기준으로 진행" in message for message in messages))

    def test_single_column_cell_matrix_uses_cell_iteration(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 99, "cols": 99}

        called: dict[str, object] = {}

        def fake_iteration(label, text, transform, *, strip_wrapping_lines=False, preserve_clipboard_cells=False):
            called["label"] = label
            called["text"] = text
            called["strip_wrapping_lines"] = strip_wrapping_lines
            called["preserve_clipboard_cells"] = preserve_clipboard_cells
            return True

        app.transform_selected_cells_by_iteration = fake_iteration

        handled = app.transform_selected_cell_matrix(
            "날짜 정규화",
            "2021. 6. 8.\r\n2019. 4. 2.",
            lambda text: normalize_dates(text, "dot_padded"),
        )

        self.assertTrue(handled)
        self.assertEqual(called["label"], "날짜 정규화")
        self.assertEqual(called["text"], "2021. 6. 8.\r\n2019. 4. 2.")
        self.assertFalse(called["preserve_clipboard_cells"])

    def test_multi_column_cell_matrix_prefers_detected_cell_iteration(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19, "GetPos": lambda self: (10, 0, 0)})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 2, "cols": 2}
        app.detect_selected_cell_rect = lambda endpoint, values, preferred_shape=None, only_preferred_shape=False: {
            "rows": 2,
            "cols": 2,
            "corner": "bottom_right",
            "start_label": "A1",
        }

        called: dict[str, object] = {}

        def fake_transform_rect(label, endpoint, rect, transform, *, strip_wrapping_lines=False):
            called["label"] = label
            called["endpoint"] = endpoint
            called["rect"] = rect
            called["strip_wrapping_lines"] = strip_wrapping_lines
            return 4, 4

        app.transform_detected_cell_rect = fake_transform_rect
        app.set_clipboard_text = lambda _text: self.fail("TSV paste path should not be used")
        app.run_hwp_command = lambda _command: self.fail("TSV paste path should not be used")

        handled = app.transform_selected_cell_matrix(
            "날짜 정규화",
            "2021. 6. 8.\t2019. 4. 2.\r\n2026. 3. 18.\t2022. 9. 1.",
            lambda text: normalize_dates(text, "dot_padded"),
        )

        self.assertTrue(handled)
        self.assertEqual(called["endpoint"], (10, 0, 0))
        self.assertEqual(called["rect"]["rows"], 2)
        self.assertEqual(called["rect"]["cols"], 2)

    def test_line_copied_multi_column_selection_uses_formula_shape_for_iteration(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19, "GetPos": lambda self: (10, 0, 0)})()
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.get_selected_cell_range_by_formula = lambda: {"rows": 4, "cols": 2}

        detected: dict[str, object] = {}

        def fake_detect(endpoint, values, preferred_shape=None, only_preferred_shape=False):
            detected["endpoint"] = endpoint
            detected["values"] = values
            detected["preferred_shape"] = preferred_shape
            detected["only_preferred_shape"] = only_preferred_shape
            return {
                "rows": 4,
                "cols": 2,
                "corner": "bottom_right",
                "start_label": "A1",
            }

        app.detect_selected_cell_rect = fake_detect
        app.transform_detected_cell_rect = lambda *args, **kwargs: (8, 8)
        app.set_clipboard_text = lambda _text: self.fail("line-copied multi-column cells must not use TSV paste")
        app.run_hwp_command = lambda _command: self.fail("line-copied multi-column cells must not use TSV paste")

        handled = app.transform_selected_cell_matrix(
            "날짜 정규화",
            "\r\n".join(
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
            ),
            lambda text: normalize_dates(text, "dot_padded"),
        )

        self.assertTrue(handled)
        self.assertEqual(detected["endpoint"], (10, 0, 0))
        self.assertEqual(len(detected["values"]), 8)
        self.assertEqual(detected["preferred_shape"], (4, 2))
        self.assertTrue(detected["only_preferred_shape"])

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

    def test_formula_address_iteration_moves_from_current_cell_to_each_formula_address(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.get_current_cell_address = lambda: (4, 5)
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
        app.log = lambda _message: None

        result = app.transform_formula_address_cells(
            "요일 제거",
            [(3, 2), (4, 2)],
            remove_weekdays_from_dates,
        )

        self.assertEqual(result, (2, 1))
        self.assertEqual(moves, [((4, 5), (3, 2)), ((3, 2), (4, 2))])

    def test_multi_column_detection_tries_selection_endpoints_after_current_pos(self) -> None:
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
        app.activate_hwp_window = lambda: None
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

        handled = app.transform_selected_cell_matrix(
            "요일 제거",
            "\r\n".join(str(index) for index in range(8)),
            remove_weekdays_from_dates,
        )

        self.assertTrue(handled)
        self.assertEqual(tried, [(99, 0, 0), (10, 0, 0), (20, 0, 0)])

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


if __name__ == "__main__":
    unittest.main()
