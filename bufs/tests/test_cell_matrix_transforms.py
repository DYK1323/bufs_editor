import json
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
    consume_numbering_entry,
    convert_decimal_numbers_to_currency_unit,
    flatten_cell_matrix,
    find_suspicious_decimal_numbers,
    google_drive_direct_download_url,
    google_drive_file_id,
    github_api_latest_release_url,
    initial_app_geometry,
    is_rectangular_cell_matrix,
    is_newer_version,
    looks_like_cell_clipboard_matrix,
    manual_outline_prefix_length,
    normalize_decimal_numbers,
    normalize_dates,
    normalize_years,
    normalize_clipboard_newlines,
    normalize_table_settings,
    parse_cell_clipboard_matrix,
    parse_cover_input,
    parse_github_release_info,
    parse_update_info,
    create_filled_title_number_box_file,
    create_title_number_box_hwpml,
    resolve_config_path,
    migrate_title_number_box_roles,
    normalize_title_box_text,
    sync_builtin_templates,
    style_entry_has_role,
    title_number_box_numbers_from_hwpml,
    remove_manual_outline_prefixes,
    remove_number_commas,
    remove_weekdays_from_dates,
    scale_decimal_numbers,
    scale_decimal_numbers_for_unit_conversion,
    split_table_caption_parts,
    merge_dict,
    markdown_rows_to_html_table,
    find_markdown_table_block,
    DEFAULT_TABLE_SETTINGS,
    SuspiciousDecimalNumberError,
    UnsafeUnitConversionNumberError,
    ensure_no_suspicious_decimal_numbers,
    ensure_safe_decimal_unit_conversion_text,
    find_outline_style_rule,
    is_standard_regulation_wrapper,
    parse_cell_addresses_from_formula_command,
    parse_cell_address_from_keyindicator,
    parse_style_marker_text,
    looks_like_single_copied_cell,
    TABLE_BORDER_ICON_BUTTON_SIZE,
    TABLE_BORDER_ICON_PRESETS,
    TABLE_ICON_BUTTON_BG,
    TABLE_ICON_BUTTON_GAP,
    TABLE_STYLE_ICON_BUTTON_SIZE,
    TABLE_STYLE_ICON_PRESETS,
    InlineRule,
    NumberingRunState,
    StyleEntry,
    StyleRecord,
    StyleSet,
    find_inline_rule_ranges,
    find_inline_rule_ranges_for_style_set,
    format_page_settings,
    parse_report_template_input,
    page_settings_match,
    PageSettings,
    read_hwpx_page_settings,
    load_style_sets,
    read_style_records,
    read_style_records_from_hwp_file,
    read_style_records_from_hwpml,
    needs_regulation_wrapper_conversion,
    save_style_sets,
    style_set_from_records,
    today_ymd_text,
    today_ym_text,
    transform_cell_matrix,
    validate_update_zip,
    MvpApp,
)


class CellMatrixTransformTests(unittest.TestCase):
    def test_initial_app_geometry_fills_work_area_height_except_margins(self) -> None:
        self.assertEqual(initial_app_geometry((0, 0, 1920, 1040)), "420x1024+1492+8")

    def test_sync_builtin_templates_overwrites_builtin_and_preserves_custom(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            bundled = root / "bundled"
            builtin = root / "config" / "templates" / "builtin"
            custom = root / "config" / "templates" / "custom"
            bundled.mkdir()
            builtin.mkdir(parents=True)
            custom.mkdir(parents=True)
            (bundled / "base.hwpx").write_text("new builtin", encoding="utf-8")
            (builtin / "base.hwpx").write_text("old builtin", encoding="utf-8")
            (custom / "base.hwpx").write_text("custom file", encoding="utf-8")

            sync_builtin_templates(bundled, builtin, custom)

            self.assertEqual((builtin / "base.hwpx").read_text(encoding="utf-8"), "new builtin")
            self.assertEqual((custom / "base.hwpx").read_text(encoding="utf-8"), "custom file")

    def test_sync_builtin_templates_adds_new_nested_builtin_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            bundled = root / "bundled"
            builtin = root / "config" / "templates" / "builtin"
            custom = root / "config" / "templates" / "custom"
            (bundled / "nested").mkdir(parents=True)
            (bundled / "nested" / "new.hwpx").write_text("new", encoding="utf-8")

            sync_builtin_templates(bundled, builtin, custom)

            self.assertEqual((builtin / "nested" / "new.hwpx").read_text(encoding="utf-8"), "new")
            self.assertTrue(custom.exists())

    def test_resolve_config_path_maps_legacy_templates_path_to_builtin(self) -> None:
        self.assertEqual(
            resolve_config_path("templates/legacy.hwpx"),
            hwp_style_mvp.BUILTIN_TEMPLATE_DIR / "legacy.hwpx",
        )

    def test_resolve_config_path_keeps_custom_templates_path_in_user_config(self) -> None:
        self.assertEqual(
            resolve_config_path("templates/custom/my.hwpx"),
            hwp_style_mvp.CONFIG_ROOT / "templates" / "custom" / "my.hwpx",
        )

    def test_today_ymd_text_uses_unpadded_year_month_day(self) -> None:
        self.assertEqual(today_ymd_text(hwp_style_mvp.datetime(2026, 7, 2)), "2026. 7. 2.")

    def test_today_ym_text_uses_unpadded_year_month(self) -> None:
        self.assertEqual(today_ym_text(hwp_style_mvp.datetime(2026, 7, 2)), "2026. 7.")

    def test_parse_cover_input_uses_today_month_when_date_line_is_empty(self) -> None:
        title, date_text, department = parse_cover_input(
            "\ubb38\uc11c\uc81c\ubaa9\n\n\ubd80\uc11c\uba85",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(title, "\ubb38\uc11c\uc81c\ubaa9")
        self.assertEqual(date_text, "2026. 7.")
        self.assertEqual(department, "\ubd80\uc11c\uba85")

    def test_parse_cover_input_uses_today_month_when_date_line_is_missing(self) -> None:
        title, date_text, department = parse_cover_input(
            "\ubb38\uc11c\uc81c\ubaa9\n\ubd80\uc11c\uba85",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(title, "\ubb38\uc11c\uc81c\ubaa9")
        self.assertEqual(date_text, "2026. 7.")
        self.assertEqual(department, "\ubd80\uc11c\uba85")

    def test_parse_cover_input_removes_labels(self) -> None:
        title, date_text, department = parse_cover_input(
            "\uc81c\ubaa9: \ubb38\uc11c\uc81c\ubaa9\n\ub0a0\uc9dc: 2026. 8.\n\ubd80\uc11c: \ubd80\uc11c\uba85",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(title, "\ubb38\uc11c\uc81c\ubaa9")
        self.assertEqual(date_text, "2026. 8.")
        self.assertEqual(department, "\ubd80\uc11c\uba85")

    def test_parse_cover_input_uses_today_month_when_labeled_date_is_missing(self) -> None:
        title, date_text, department = parse_cover_input(
            "\uc81c\ubaa9: \ubb38\uc11c\uc81c\ubaa9\n\ubd80\uc11c: \ubd80\uc11c\uba85",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(title, "\ubb38\uc11c\uc81c\ubaa9")
        self.assertEqual(date_text, "2026. 7.")
        self.assertEqual(department, "\ubd80\uc11c\uba85")

    def test_parse_cover_input_accepts_today_markers(self) -> None:
        for marker in ("{{\uc624\ub298}}", "\uc624\ub298"):
            with self.subTest(marker=marker):
                title, date_text, department = parse_cover_input(
                    f"\uc81c\ubaa9: \ubb38\uc11c\uc81c\ubaa9\n\ub0a0\uc9dc: {marker}\n\ubd80\uc11c: \ubd80\uc11c\uba85",
                    hwp_style_mvp.datetime(2026, 7, 22),
                )

                self.assertEqual(title, "\ubb38\uc11c\uc81c\ubaa9")
                self.assertEqual(date_text, "2026. 7.")
                self.assertEqual(department, "\ubd80\uc11c\uba85")

    def test_parse_report_template_input_uses_today_when_date_value_is_empty(self) -> None:
        fields = parse_report_template_input(
            "\uc81c\ubaa9: \ubcf4\uace0\n\ubd80\uc11c: \uc13c\ud130\n\ub0a0\uc9dc:",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(fields.date_text, "2026. 7. 22.")

    def test_parse_report_template_input_uses_today_when_date_field_is_missing(self) -> None:
        fields = parse_report_template_input(
            "\uc81c\ubaa9: \ubcf4\uace0\n\ubd80\uc11c: \uc13c\ud130",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(fields.date_text, "2026. 7. 22.")

    def test_parse_report_template_input_accepts_today_markers(self) -> None:
        for marker in ("{{\uc624\ub298}}", "\uc624\ub298"):
            with self.subTest(marker=marker):
                fields = parse_report_template_input(
                    f"\uc81c\ubaa9: \ubcf4\uace0\n\ubd80\uc11c: \uc13c\ud130\n\ub0a0\uc9dc: {marker}",
                    hwp_style_mvp.datetime(2026, 7, 22),
                )

                self.assertEqual(fields.date_text, "2026. 7. 22.")

    def test_parse_report_template_input_replaces_today_marker(self) -> None:
        fields = parse_report_template_input(
            "제목 : 오마이갓 이런일이\n부서 : 대학성과분석센터\n날짜 : {{오늘}}\n",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(fields.title, "오마이갓 이런일이")
        self.assertEqual(fields.department, "대학성과분석센터")
        self.assertEqual(fields.date_text, "2026. 7. 22.")

    def test_parse_report_template_input_keeps_explicit_date_text(self) -> None:
        fields = parse_report_template_input(
            "제목: 보고\n부서: 센터\n날짜: 2026. 8. 3.",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(fields.date_text, "2026. 8. 3.")

    def test_parse_report_template_input_does_not_reuse_title_for_table_placeholders(self) -> None:
        fields = parse_report_template_input(
            "\uc81c\ubaa9: \ubcf4\uace0\n\ubd80\uc11c: \uc13c\ud130\n\ub0a0\uc9dc: 2026. 8. 3.",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertIsNone(fields.table_title)
        self.assertIsNone(fields.table_title_line)

    def test_parse_report_template_input_accepts_explicit_table_placeholders(self) -> None:
        fields = parse_report_template_input(
            "\uc81c\ubaa9: \ubcf4\uace0\n\ubd80\uc11c: \uc13c\ud130\n\ub0a0\uc9dc: 2026. 8. 3.\n"
            "\ud45c\uc81c\ubaa9: \ud45c \uc81c\ubaa9\n\ud45c\uc81c\ubaa9\uc904: \ud45c \uc81c\ubaa9\uc904",
            hwp_style_mvp.datetime(2026, 7, 22),
        )

        self.assertEqual(fields.table_title, "\ud45c \uc81c\ubaa9")
        self.assertEqual(fields.table_title_line, "\ud45c \uc81c\ubaa9\uc904")

    def test_read_hwpx_page_settings_reads_general_report_template_margins(self) -> None:
        settings = read_hwpx_page_settings(hwp_style_mvp.GENERAL_REPORT_TEMPLATE_FILE)

        self.assertIsNotNone(settings)
        assert settings is not None
        self.assertEqual(settings.paper_width, 59528)
        self.assertEqual(settings.paper_height, 84188)
        self.assertEqual(settings.left_margin, 5669)
        self.assertEqual(settings.right_margin, 5669)
        self.assertEqual(settings.top_margin, 4252)
        self.assertEqual(settings.bottom_margin, 4252)
        self.assertEqual(settings.header_len, 2835)
        self.assertEqual(settings.footer_len, 2835)

    def test_read_hwpx_page_settings_reads_cover_template_margins(self) -> None:
        settings = read_hwpx_page_settings(hwp_style_mvp.COVER_FILE)

        self.assertIsNotNone(settings)
        assert settings is not None
        self.assertGreater(settings.paper_width, 0)
        self.assertGreater(settings.paper_height, 0)
        self.assertGreaterEqual(settings.left_margin, 0)
        self.assertGreaterEqual(settings.right_margin, 0)

    def test_page_settings_match_compares_all_page_values(self) -> None:
        expected = PageSettings(59528, 84188, 0, 5669, 5669, 4252, 4252, 2835, 2835, 0, 0)
        actual = PageSettings(59528, 84188, 0, 5669, 5669, 4252, 4252, 2835, 2835, 0, 0)
        changed = PageSettings(59528, 84188, 0, 5669, 5669, 4252, 4252, 2835, 2835, 1, 0)

        self.assertTrue(page_settings_match(expected, actual))
        self.assertFalse(page_settings_match(expected, changed))
        self.assertFalse(page_settings_match(expected, None))

    def test_format_page_settings_includes_margins_for_diagnostics(self) -> None:
        settings = PageSettings(59528, 84188, 0, 5669, 5669, 4252, 4252, 2835, 2835, 0, 0)

        self.assertIn("margin=L5669 R5669 T4252 B4252", format_page_settings(settings))
        self.assertEqual(format_page_settings(None), "none")

    def test_inline_rule_ranges_match_parenthesized_keyword_with_spaces(self) -> None:
        rule = InlineRule("keyword", "leading_parenthesized_after_identifier", include_wrapper=True)

        self.assertEqual(find_inline_rule_ranges("  (project plan keyword) body", rule), [(2, 24, None)])

    def test_inline_rule_ranges_match_markdown_bold_inner_text_when_removing_markers(self) -> None:
        rule = InlineRule("markdown", "markdown_bold", remove_markers=True)

        self.assertEqual(find_inline_rule_ranges("pre **bold text** post", rule), [(6, 15, (4, 17))])

    def test_inline_rule_ranges_match_leading_text_before_colon(self) -> None:
        rule = InlineRule("colon", "leading_text_before_colon_after_identifier", include_wrapper=False)

        self.assertEqual(find_inline_rule_ranges("project plan: body", rule), [(0, 12, None)])

    def test_inline_rule_ranges_match_after_outline_marker_removed_examples(self) -> None:
        keyword_rule = InlineRule("keyword", "leading_parenthesized_after_identifier", include_wrapper=True)
        colon_rule = InlineRule("colon", "leading_text_before_colon_after_identifier", include_wrapper=False)

        self.assertEqual(find_inline_rule_ranges("(background) text", keyword_rule), [(0, 12, None)])
        self.assertEqual(find_inline_rule_ranges("purpose : text", colon_rule), [(0, 7, None)])

    def test_inline_rule_ranges_skip_configured_outline_marker_without_deleting_it(self) -> None:
        style_set = StyleSet("set", [StyleEntry("body", outline_markers=("w",))], [])
        keyword_rule = InlineRule("keyword", "leading_parenthesized_after_identifier", include_wrapper=True)
        colon_rule = InlineRule("colon", "leading_text_before_colon_after_identifier", include_wrapper=False)

        self.assertEqual(find_inline_rule_ranges_for_style_set("w (background) text", keyword_rule, style_set), [(2, 14, None)])
        self.assertEqual(find_inline_rule_ranges_for_style_set("w purpose : text", colon_rule, style_set), [(2, 9, None)])

    def test_inline_rule_ranges_only_match_first_leading_parentheses(self) -> None:
        style_set = StyleSet("set", [StyleEntry("body", outline_markers=("o",))], [])
        rule = InlineRule("keyword", "leading_parenthesized_after_identifier", include_wrapper=True)

        text = "o (airfare) text with later (exchange professor) parentheses"

        self.assertEqual(find_inline_rule_ranges_for_style_set(text, rule, style_set), [(2, 11, None)])

    def test_numbering_restart_state_restarts_lower_level_after_higher_level(self) -> None:
        state = NumberingRunState()
        level_1 = StyleEntry("body-outline-1", numbering_group="body", numbering_level=1, restart_after_higher_level=True)
        level_2 = StyleEntry("body-outline-2", numbering_group="body", numbering_level=2, restart_after_higher_level=True)

        self.assertFalse(consume_numbering_entry(state, level_2))
        self.assertFalse(consume_numbering_entry(state, level_2))
        self.assertFalse(consume_numbering_entry(state, level_1))
        self.assertTrue(consume_numbering_entry(state, level_2))
        self.assertFalse(consume_numbering_entry(state, level_2))

    def test_numbering_restart_state_restarts_first_lower_level_after_higher_level(self) -> None:
        state = NumberingRunState()
        level_1 = StyleEntry("body-outline-1", numbering_group="body", numbering_level=1, restart_after_higher_level=True)
        level_2 = StyleEntry("body-outline-2", numbering_group="body", numbering_level=2, restart_after_higher_level=True)

        self.assertFalse(consume_numbering_entry(state, level_1))
        self.assertTrue(consume_numbering_entry(state, level_2))
        self.assertFalse(consume_numbering_entry(state, level_2))

    def test_outline_rule_filters_duplicate_markers_by_table_context(self) -> None:
        style_set = StyleSet(
            "set",
            [
                StyleEntry("본문-개요2", outline_markers=("ㅇ",)),
                StyleEntry("표내용-개요2", table_style=True, outline_markers=("ㅇ",)),
            ],
            [],
        )

        body_entry, body_prefix = find_outline_style_rule("ㅇ 본문", style_set, in_table=False)
        table_entry, table_prefix = find_outline_style_rule("ㅇ 표내용", style_set, in_table=True)

        self.assertEqual(body_entry.name, "본문-개요2")
        self.assertEqual(table_entry.name, "표내용-개요2")
        self.assertEqual(body_prefix, len("ㅇ "))
        self.assertEqual(table_prefix, len("ㅇ "))

    def test_table_icon_presets_match_expected_grid_sizes_and_files(self) -> None:
        self.assertEqual(len(TABLE_STYLE_ICON_PRESETS), 6)
        self.assertEqual(len(TABLE_BORDER_ICON_PRESETS), 6)
        self.assertEqual(TABLE_STYLE_ICON_BUTTON_SIZE, 48)
        self.assertEqual(TABLE_BORDER_ICON_BUTTON_SIZE, 48)
        self.assertEqual(TABLE_ICON_BUTTON_GAP, 6)
        self.assertEqual(TABLE_ICON_BUTTON_BG, "#FFFFFF")
        for icon_name, _preset_name, _tooltip in TABLE_STYLE_ICON_PRESETS + TABLE_BORDER_ICON_PRESETS:
            self.assertTrue((ROOT / "icons" / "2x" / f"{icon_name}@2x.png").exists())

    def test_clear_selected_character_style_runs_char_shape_normal_action(self) -> None:
        app = object.__new__(MvpApp)
        commands: list[str] = []
        app.ensure_hwp = lambda: True
        app.run_hwp_command = lambda command: commands.append(command) or True
        app.log = lambda _message: None

        app.clear_selected_character_style()

        self.assertEqual(commands, ["StyleClearCharStyle"])

    def test_select_current_table_first_row_selects_table_object_then_row(self) -> None:
        app = object.__new__(MvpApp)
        commands: list[str] = []
        app.run_hwp_command = lambda command: commands.append(command) or True
        app.select_current_table_object = lambda: commands.append("SelectCtrlReverse") or True
        app.select_selected_table_first_cell = lambda: commands.append("ShapeObjTableSelCell") or True
        app.move_to_nearby_table_cell = lambda: self.fail("첫 셀 진입 성공 시 nearby fallback을 쓰지 않아야 합니다")

        ok, steps = app.select_current_table_first_row()

        self.assertTrue(ok)
        self.assertEqual(commands, ["Cancel", "SelectCtrlReverse", "ShapeObjTableSelCell", "TableCellBlock", "TableCellBlockRow"])
        self.assertIn("TableCellBlockRow=True", steps)

    def test_select_current_table_first_row_accepts_existing_cell_block_after_false_block_command(self) -> None:
        app = object.__new__(MvpApp)
        commands: list[str] = []
        app.select_current_table_object = lambda: commands.append("SelectCtrlReverse") or True
        app.select_selected_table_first_cell = lambda: commands.append("ShapeObjTableSelCell") or True
        app.move_to_nearby_table_cell = lambda: self.fail("첫 셀 진입 성공 시 nearby fallback을 쓰지 않아야 합니다")
        app.is_selected_cell_block = lambda: True

        def run_command(command: str) -> bool:
            commands.append(command)
            return command != "TableCellBlock"

        app.run_hwp_command = run_command

        ok, steps = app.select_current_table_first_row()

        self.assertTrue(ok)
        self.assertEqual(commands, ["Cancel", "SelectCtrlReverse", "ShapeObjTableSelCell", "TableCellBlock", "TableCellBlockRow"])
        self.assertIn("TableCellBlock=False, ready=True", steps)

    def test_table_style_preset_applies_header_background_after_border(self) -> None:
        app = object.__new__(MvpApp)
        calls: list[tuple[str, object | None]] = []
        app.ensure_hwp = lambda: True
        app.select_current_table_all_cells = lambda: calls.append(("all_cells", None)) or True
        app.set_table_border_preset = lambda preset: calls.append(("border", preset)) or ("CellZoneBorderFill", True)
        app.select_current_table_first_row = lambda: calls.append(("header", None)) or (True, ["TableCellBlockRow=True"])
        app.table_settings = {
            "title_cell": {
                "background_color": "셀배경",
                "paragraph_style": "표내용-중간",
                "character_style": "표내용-굵게",
            }
        }
        app.set_selected_cells_as_header = lambda: calls.append(("header_attr", None)) or ("CellShape", True, True)
        app.apply_style_by_name = lambda name: calls.append(("style", name)) or True
        app.palette_color = lambda name: calls.append(("palette", name)) or hwp_style_mvp.PaletteColor(
            name,
            (220, 221, 221),
            "#DCDDDE",
            "",
        )
        app.set_cell_fill_color = lambda color: calls.append(("fill", color.name)) or ("CellBorderFill", True, True)
        app.log = lambda _message: None

        app.apply_table_style_preset("thin_top_bottom")

        self.assertEqual(
            calls,
            [
                ("all_cells", None),
                ("border", "thin_top_bottom"),
                ("header", None),
                ("header_attr", None),
                ("style", "표내용-중간"),
                ("style", "표내용-굵게"),
                ("palette", "셀배경"),
                ("fill", "셀배경"),
            ],
        )

    def test_table_style_preset_can_add_header_double_bottom_border(self) -> None:
        app = object.__new__(MvpApp)
        calls: list[tuple[str, object | None]] = []
        app.ensure_hwp = lambda: True
        app.select_current_table_all_cells = lambda: calls.append(("all_cells", None)) or True
        app.set_table_border_preset = lambda preset: calls.append(("border", preset)) or ("CellZoneBorderFill", True)
        app.select_current_table_first_row = lambda: calls.append(("header", None)) or (True, ["TableCellBlockRow=True"])
        app.table_settings = {
            "title_cell": {
                "background_color": "셀배경",
                "paragraph_style": "표내용-중간",
                "character_style": "표내용-굵게",
            }
        }
        app.set_selected_cells_as_header = lambda: calls.append(("header_attr", None)) or ("CellShape", True, True)
        app.apply_style_by_name = lambda name: calls.append(("style", name)) or True
        app.palette_color = lambda name: calls.append(("palette", name)) or hwp_style_mvp.PaletteColor(
            name,
            (220, 221, 221),
            "#DCDDDE",
            "",
        )
        app.set_cell_fill_color = lambda color: calls.append(("fill", color.name)) or ("CellBorderFill", True, True)
        app.log = lambda _message: None

        app.apply_table_style_preset("thick_top_bottom_header_double")

        self.assertEqual(
            calls,
            [
                ("all_cells", None),
                ("border", "thick_top_bottom"),
                ("header", None),
                ("header_attr", None),
                ("style", "표내용-중간"),
                ("style", "표내용-굵게"),
                ("border", "double_bottom"),
                ("palette", "셀배경"),
                ("fill", "셀배경"),
            ],
        )

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

    def test_style_sets_round_trip_table_and_caption_flags(self) -> None:
        old_path = hwp_style_mvp.STYLE_SETS_FILE
        old_bundled_path = hwp_style_mvp.BUNDLED_STYLE_SETS_FILE
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                hwp_style_mvp.STYLE_SETS_FILE = Path(temp_dir) / "style-sets.json"
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE = Path(temp_dir) / "style-sets.default.json"
                style_sets = [
                    StyleSet(
                        "보고서용",
                        [
                            StyleEntry("표내용-중간", table_style=True, template_file="templates/box-a.hwpx"),
                            StyleEntry("[표/그림] 캡션", caption_style=True),
                            StyleEntry(
                                "본문-개요2",
                                outline_markers=("ㅇ", "○"),
                                numbering_group="body",
                                numbering_level=2,
                                restart_after_higher_level=True,
                            ),
                        ],
                        [StyleEntry("표내용-굵게", table_style=True)],
                    )
                ]

                save_style_sets("보고서용", style_sets)
                active_name, loaded_sets = load_style_sets()
            finally:
                hwp_style_mvp.STYLE_SETS_FILE = old_path
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE = old_bundled_path

        self.assertEqual(active_name, "보고서용")
        self.assertTrue(loaded_sets[0].paragraph_styles[0].table_style)
        self.assertEqual(loaded_sets[0].paragraph_styles[0].template_file, "templates/box-a.hwpx")
        self.assertTrue(loaded_sets[0].paragraph_styles[1].caption_style)
        self.assertEqual(loaded_sets[0].paragraph_styles[2].outline_markers, ("ㅇ", "○"))
        self.assertEqual(loaded_sets[0].paragraph_styles[2].numbering_group, "body")
        self.assertEqual(loaded_sets[0].paragraph_styles[2].numbering_level, 2)
        self.assertTrue(loaded_sets[0].paragraph_styles[2].restart_after_higher_level)
        self.assertTrue(loaded_sets[0].character_styles[0].table_style)

    def test_style_sets_copy_bundled_default_to_user_config_on_first_load(self) -> None:
        old_path = hwp_style_mvp.STYLE_SETS_FILE
        old_bundled_path = hwp_style_mvp.BUNDLED_STYLE_SETS_FILE
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                temp_root = Path(temp_dir)
                hwp_style_mvp.STYLE_SETS_FILE = temp_root / "user" / "style-sets.json"
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE = temp_root / "bundle" / "style-sets.json"
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE.parent.mkdir()
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE.write_text(
                    json.dumps(
                        {
                            "active_set": "배포기본",
                            "style_sets": [
                                {
                                    "name": "배포기본",
                                    "paragraph_styles": [{"name": "본문", "outline_markers": ["ㅇ"]}],
                                    "character_styles": [],
                                    "inline_rules": [],
                                }
                            ],
                        },
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )

                active_name, loaded_sets = load_style_sets()
                user_file_exists = hwp_style_mvp.STYLE_SETS_FILE.exists()
            finally:
                hwp_style_mvp.STYLE_SETS_FILE = old_path
                hwp_style_mvp.BUNDLED_STYLE_SETS_FILE = old_bundled_path

        self.assertEqual(active_name, "배포기본")
        self.assertEqual(loaded_sets[0].paragraph_styles[0].outline_markers, ("ㅇ",))
        self.assertTrue(user_file_exists)

    def test_imported_hwpx_style_records_can_build_style_set(self) -> None:
        header = """<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:styles>
    <hh:style id="1" type="PARA" name="표내용-중간"/>
    <hh:style id="2" type="PARA" name="[표/그림] 캡션"/>
    <hh:style id="3" type="CHAR" name="표내용-굵게"/>
  </hh:styles>
</hh:head>
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("Contents/header.xml", header)

            records = read_style_records_from_hwp_file(path)

        style_set = style_set_from_records("sample", records)

        self.assertEqual(style_set.name, "sample")
        self.assertTrue(style_set.paragraph_styles[0].table_style)
        self.assertTrue(style_set.paragraph_styles[1].caption_style)
        self.assertTrue(style_set.character_styles[0].table_style)

    def test_style_set_from_records_assigns_template_to_title_number_box_entries(self) -> None:
        records = [StyleRecord(1, 1, "PARA", "대제목")]

        style_set = style_set_from_records("sample", records, Path("templates/sample.hwpx"))

        self.assertEqual(style_set.role_configs.get("title_number_box", {}).get("template_file"), "templates/sample.hwpx")
        self.assertEqual(style_set.paragraph_styles[0].template_file, "templates/sample.hwpx")

    def test_parse_style_marker_text_splits_commas_and_removes_duplicates(self) -> None:
        self.assertEqual(parse_style_marker_text("ㅇ, ○, 원기호, 원, ㅇ"), ["ㅇ", "○", "원기호", "원"])

    def test_regulation_wrapper_conversion_detects_nonstandard_corner_brackets(self) -> None:
        self.assertTrue(needs_regulation_wrapper_conversion("「규정명」"))
        self.assertFalse(needs_regulation_wrapper_conversion("｢규정명｣"))

    def test_version_comparison_detects_newer_release(self) -> None:
        self.assertTrue(is_newer_version("0.1.1", "0.1.0"))
        self.assertTrue(is_newer_version("v0.2.0", "0.1.9"))
        self.assertFalse(is_newer_version("0.1.0", "0.1.0"))
        self.assertFalse(is_newer_version("0.1", "0.1.0"))

    def test_google_drive_file_url_can_be_converted_to_download_url(self) -> None:
        url = "https://drive.google.com/file/d/abc123XYZ/view?usp=sharing"
        self.assertEqual(google_drive_file_id(url), "abc123XYZ")
        self.assertEqual(
            google_drive_direct_download_url(url),
            "https://drive.google.com/uc?export=download&id=abc123XYZ",
        )

    def test_parse_update_info_accepts_release_notes_alias(self) -> None:
        info = parse_update_info(
            {
                "version": "0.1.2",
                "download_url": "https://drive.google.com/file/d/app/view",
                "release_notes": "업데이트 확인 추가",
            }
        )
        self.assertEqual(info.latest_version, "0.1.2")
        self.assertEqual(info.notes, "업데이트 확인 추가")

    def test_github_latest_release_url_uses_owner_and_repo(self) -> None:
        self.assertEqual(
            github_api_latest_release_url("DYK1323", "bufs_editor"),
            "https://api.github.com/repos/DYK1323/bufs_editor/releases/latest",
        )

    def test_parse_github_release_info_selects_matching_zip_asset(self) -> None:
        info = parse_github_release_info(
            {
                "tag_name": "v1.0.2",
                "html_url": "https://github.com/DYK1323/bufs_editor/releases/tag/v1.0.2",
                "body": "release notes",
                "assets": [
                    {
                        "name": "source.zip",
                        "browser_download_url": "https://example.test/source.zip",
                    },
                    {
                        "name": "BUFS-HWP-Editor-v1.0.2.zip",
                        "browser_download_url": "https://example.test/app.zip",
                        "digest": "sha256:abcdef",
                        "size": 123,
                    },
                ],
            }
        )
        self.assertEqual(info.latest_version, "v1.0.2")
        self.assertEqual(info.download_url, "https://example.test/app.zip")
        self.assertEqual(info.asset_name, "BUFS-HWP-Editor-v1.0.2.zip")
        self.assertEqual(info.sha256, "abcdef")
        self.assertEqual(info.source, "github")

    def test_validate_update_zip_accepts_expected_release_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = Path(temp_dir) / "release.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("BUFS-HWP-Editor/BUFS-HWP-Editor.exe", "")
                archive.writestr("BUFS-HWP-Editor/_internal/runtime.dat", "")
            validate_update_zip(zip_path, "BUFS-HWP-Editor", "BUFS-HWP-Editor.exe")

    def test_validate_update_zip_rejects_files_outside_release_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = Path(temp_dir) / "release.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("BUFS-HWP-Editor/BUFS-HWP-Editor.exe", "")
                archive.writestr("BUFS-HWP-Editor/_internal/runtime.dat", "")
                archive.writestr("unexpected.txt", "")
            with self.assertRaises(ValueError):
                validate_update_zip(zip_path, "BUFS-HWP-Editor", "BUFS-HWP-Editor.exe")

    def test_validate_update_zip_rejects_path_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = Path(temp_dir) / "release.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("../escape.txt", "")
            with self.assertRaises(ValueError):
                validate_update_zip(zip_path, "BUFS-HWP-Editor", "BUFS-HWP-Editor.exe")

    def test_title_number_box_text_is_normalized_to_single_line(self) -> None:
        self.assertEqual(normalize_title_box_text("  첫 줄\r\n둘째\t줄  "), "첫 줄 둘째 줄")

    def test_create_filled_title_number_box_replaces_title_text(self) -> None:
        target = create_filled_title_number_box_file("사업 <성과> & 계획", number=7)
        with zipfile.ZipFile(target) as zf:
            section = zf.read("Contents/section0.xml").decode("utf-8")
            preview = zf.read("Preview/PrvText.txt").decode("utf-8")

        self.assertIn("사업 &lt;성과&gt; &amp; 계획", section)
        self.assertIn(">7<", section)
        self.assertNotIn("{{header}}", section)
        self.assertNotIn("{{no}}", section)
        self.assertIn("사업 <성과> & 계획", preview)
        self.assertNotIn("{{header}}", preview)
        self.assertNotIn("{{no}}", preview)

    def test_create_filled_title_number_box_replaces_hidden_marker_when_series_marker_is_given(self) -> None:
        target = create_filled_title_number_box_file("사업 계획", number=2, marker="{{bufs_title}}::series-a")
        with zipfile.ZipFile(target) as zf:
            section = zf.read("Contents/section0.xml").decode("utf-8")

        self.assertIn("{{bufs_title}}::series-a", section)
        self.assertNotIn("<hp:t>{{bufs_title}}</hp:t>", section)

    def test_create_title_number_box_hwpml_replaces_number_and_title(self) -> None:
        hwpml = create_title_number_box_hwpml("사업 <성과> & 계획", number=7)

        self.assertIn("사업 &lt;성과&gt; &amp; 계획", hwpml)
        self.assertIn(">7<", hwpml)
        self.assertIn("{{bufs_title}}", hwpml)
        self.assertNotIn("{{header}}", hwpml)
        self.assertNotIn("{{no}}", hwpml)

    def test_create_cover_from_selected_lines_inserts_empty_cover_when_selection_is_missing(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: True
        commands: list[str] = []

        def fake_run(command: str) -> bool:
            commands.append(command)
            if command == "Copy":
                return False
            self.fail(f"빈 표지 삽입에서는 {command} 명령을 호출하지 않아야 합니다")

        app.run_hwp_command = fake_run
        app.get_clipboard_text = lambda: ""
        app.cover_confidential = type("Flag", (), {"get": lambda self: False})()
        inserted: list[Path] = []
        app.insert_file_at_cursor = lambda path: inserted.append(path) or True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.apply_page_settings = lambda settings, apply_to=3: ("apply", True, "detail")
        cover_path = Path("empty-cover.hwpx")
        created: list[tuple[str, str, str, bool]] = []
        original_create_cover = hwp_style_mvp.create_filled_cover_file
        original_read_page = hwp_style_mvp.read_hwpx_page_settings
        hwp_style_mvp.create_filled_cover_file = lambda title, date_text, department, confidential: (
            created.append((title, date_text, department, confidential)) or cover_path
        )
        hwp_style_mvp.read_hwpx_page_settings = lambda _path: None
        try:
            app.create_cover_from_selected_lines()
        finally:
            hwp_style_mvp.create_filled_cover_file = original_create_cover
            hwp_style_mvp.read_hwpx_page_settings = original_read_page

        self.assertEqual(commands, ["Copy"])
        self.assertEqual(created, [("", "", "", False)])
        self.assertEqual(inserted, [cover_path])

    def test_insert_general_report_template_inserts_plain_template_when_selection_is_missing(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: True
        app.ensure_open_hwp_document = lambda _label: True
        app.get_selected_text_positions = lambda: None
        commands: list[str] = []

        def fake_run(command: str) -> bool:
            commands.append(command)
            self.fail(f"빈 보고양식 삽입에서는 {command} 명령을 호출하지 않아야 합니다")

        app.run_hwp_command = fake_run
        app.get_clipboard_text = lambda: ""
        inserted: list[tuple[Path, int]] = []
        app.insert_file_at_cursor = lambda path, keep_section=0: inserted.append((path, keep_section)) or True
        app.apply_page_settings = lambda settings, apply_to=3: ("apply", True, "detail")
        app.refresh_current_doc_style_map = lambda force=False: None
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        original_template = hwp_style_mvp.GENERAL_REPORT_TEMPLATE_FILE
        original_read_styles = hwp_style_mvp.read_style_records
        original_read_page = hwp_style_mvp.read_hwpx_page_settings
        with tempfile.TemporaryDirectory() as temp_name:
            template_path = Path(temp_name) / "plain-report.hwpx"
            template_path.write_bytes(b"template")
            try:
                hwp_style_mvp.GENERAL_REPORT_TEMPLATE_FILE = template_path
                hwp_style_mvp.read_style_records = lambda _path: []
                hwp_style_mvp.read_hwpx_page_settings = lambda _path: None
                app.insert_general_report_template()
            finally:
                hwp_style_mvp.GENERAL_REPORT_TEMPLATE_FILE = original_template
                hwp_style_mvp.read_style_records = original_read_styles
                hwp_style_mvp.read_hwpx_page_settings = original_read_page

            self.assertEqual(commands, [])
            self.assertEqual(inserted, [(template_path, 0)])

    def test_insert_general_report_template_stops_before_parsing_when_hwp_is_unavailable(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: False
        app.run_hwp_command = lambda _command: self.fail("한글 연결 실패 시 Copy를 호출하지 않아야 합니다")
        app.get_clipboard_text = lambda: self.fail("한글 연결 실패 시 클립보드를 읽지 않아야 합니다")

        app.insert_general_report_template()

    def test_insert_general_report_template_stops_before_parsing_when_document_is_missing(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: True
        app.ensure_open_hwp_document = lambda _label: False
        app.run_hwp_command = lambda _command: self.fail("문서 없음 시 Copy를 호출하지 않아야 합니다")
        app.get_clipboard_text = lambda: self.fail("문서 없음 시 클립보드를 읽지 않아야 합니다")

        app.insert_general_report_template()

    def test_style_entry_has_title_number_box_role(self) -> None:
        entry = StyleEntry("대제목", roles=("title_number_box",))

        self.assertTrue(style_entry_has_role(entry, "title_number_box"))
        self.assertFalse(style_entry_has_role(entry, "table_bold"))

    def test_migrate_title_number_box_roles_adds_role_to_existing_user_style_set(self) -> None:
        style_sets = [
            StyleSet(
                "일반보고용",
                [StyleEntry("대제목", outline_markers=("대제목",), roles=())],
                [],
            )
        ]

        migrated = migrate_title_number_box_roles(style_sets)

        self.assertEqual(migrated[0].paragraph_styles[0].roles, ("title_number_box",))

    def test_title_number_box_numbers_from_hwpml_reads_marker_tables(self) -> None:
        hwpml = """<?xml version="1.0" encoding="UTF-16"?>
<HWPML xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl rowCnt="1" colCnt="3">
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>첫 제목</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>{{bufs_title}}</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:tbl rowCnt="1" colCnt="3">
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>3</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>셋째 제목</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>{{bufs_title}}</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:tbl rowCnt="1" colCnt="2">
    <hp:tr>
      <hp:tc><hp:t>99</hp:t></hp:tc>
      <hp:tc><hp:t>마커 없는 표는 무시</hp:t></hp:tc>
    </hp:tr>
  </hp:tbl>
</HWPML>
"""
        self.assertEqual(title_number_box_numbers_from_hwpml(hwpml), [1, 3])

    def test_title_number_box_numbers_from_hwpml_filters_by_series_marker(self) -> None:
        hwpml = """<?xml version="1.0" encoding="UTF-16"?>
<HWPML xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl rowCnt="1" colCnt="3">
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>첫 제목</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>{{bufs_title}}::alpha</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:tbl rowCnt="1" colCnt="3">
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>2</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>둘째 제목</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>{{bufs_title}}::beta</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>
</HWPML>
"""
        self.assertEqual(title_number_box_numbers_from_hwpml(hwpml, "{{bufs_title}}::alpha"), [1])
        self.assertEqual(title_number_box_numbers_from_hwpml(hwpml, "{{bufs_title}}::beta"), [2])
        self.assertEqual(title_number_box_numbers_from_hwpml(hwpml), [1, 2])

    def test_title_number_box_numbers_from_hwpml_accepts_multi_row_number_box_template(self) -> None:
        hwpml = """<?xml version="1.0" encoding="UTF-16"?>
<HWPML xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl rowCnt="2" colCnt="3">
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>소제목 A</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>{{bufs_title}}::sub</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t></hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t>본문 설명</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc><hp:subList><hp:p><hp:run><hp:t></hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>
</HWPML>
"""
        self.assertEqual(title_number_box_numbers_from_hwpml(hwpml, "{{bufs_title}}::sub"), [1])

    def test_next_title_number_box_number_uses_current_document_max_plus_one(self) -> None:
        class FakeHwp:
            def GetTextFile(self, _format_name, _option):
                return """<HWPML>
  <TABLE RowCount="1" ColCount="3"><ROW>
    <CELL><P><TEXT><CHAR>2</CHAR></TEXT></P></CELL><CELL><P><TEXT><CHAR>제목</CHAR></TEXT></P></CELL>
    <CELL><P><TEXT><CHAR>{{bufs_title}}</CHAR></TEXT></P></CELL>
  </ROW></TABLE>
  <TABLE RowCount="1" ColCount="3"><ROW>
    <CELL><P><TEXT><CHAR>5</CHAR></TEXT></P></CELL><CELL><P><TEXT><CHAR>제목</CHAR></TEXT></P></CELL>
    <CELL><P><TEXT><CHAR>{{bufs_title}}</CHAR></TEXT></P></CELL>
  </ROW></TABLE>
</HWPML>"""

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.debug = lambda _message: None

        self.assertEqual(app.next_title_number_box_number(), (6, [2, 5]))
        self.assertTrue(is_standard_regulation_wrapper("｢규정명｣"))

    def test_next_title_number_box_number_uses_previous_title_when_context_is_given(self) -> None:
        app = object.__new__(MvpApp)
        app.current_title_number_box_numbers = lambda: [1, 2, 3]
        app.title_number_box_numbers_before_paragraph = lambda list_id, para: [1]

        self.assertEqual(app.next_title_number_box_number(list_id=0, para=5), (2, [1, 2, 3]))

    def test_next_title_number_box_number_uses_entry_specific_series(self) -> None:
        app = object.__new__(MvpApp)
        entry = StyleEntry("대제목A", roles=("title_number_box",), template_file="templates/title-a.hwpx")
        app.current_title_number_box_numbers = lambda current_entry=None: [3, 4] if current_entry is entry else [1, 2, 3, 4, 9]
        app.title_number_box_numbers_before_paragraph = lambda list_id, para, current_entry=None: [3] if current_entry is entry else [1, 2, 3]

        self.assertEqual(app.next_title_number_box_number(list_id=0, para=5, entry=entry), (4, [3, 4]))

    def test_has_title_number_boxes_after_paragraph_only_checks_following_boxes(self) -> None:
        app = object.__new__(MvpApp)
        app.current_title_number_box_numbers = lambda: [1, 2, 3]
        app.title_number_box_numbers_before_paragraph = lambda list_id, para: [1, 2, 3] if para == 6 else [1]

        self.assertFalse(app.has_title_number_boxes_after_paragraph(0, 5))
        self.assertTrue(app.has_title_number_boxes_after_paragraph(0, 0))

    def test_has_title_number_boxes_after_paragraph_checks_same_series_when_entry_is_given(self) -> None:
        app = object.__new__(MvpApp)
        entry = StyleEntry("대제목B", roles=("title_number_box",), template_file="templates/title-b.hwpx")
        app.current_title_number_box_numbers = lambda current_entry=None: [1, 2] if current_entry is entry else [1, 2, 8]
        app.title_number_box_numbers_before_paragraph = lambda list_id, para, current_entry=None: [1, 2] if current_entry is entry else [1]

        self.assertFalse(app.has_title_number_boxes_after_paragraph(0, 5, entry))

    def test_ai_prompt_for_style_set_uses_active_outline_markers(self) -> None:
        app = object.__new__(MvpApp)
        style_set = StyleSet(
            "일반보고용",
            [
                StyleEntry("대제목", outline_markers=("대제목",), roles=("title_number_box",)),
                StyleEntry("소제목", outline_markers=("소제목",), numbering_group="body", numbering_level=2),
                StyleEntry("본문1", outline_markers=("원",)),
                StyleEntry("표내용", table_style=True, outline_markers=("행",)),
            ],
            [],
        )

        prompt = app.ai_prompt_for_style_set(style_set)

        self.assertIn("* 대제목 문단은 반드시 `대제목 `으로 시작한다.", prompt)
        self.assertIn("* 소제목 문단은 반드시 `소제목 `으로 시작한다.", prompt)
        self.assertIn("* 본문1 문단은 반드시 `원 `으로 시작한다.", prompt)
        self.assertIn("`대제목`, `소제목`, `원`", prompt)
        self.assertIn("[최우선 규칙]", prompt)
        self.assertIn("보고서 전체 제목, 문서 제목, 보고서명은 쓰지 말고 본문부터 시작한다.", prompt)
        self.assertIn("`대제목`은 한 보고서 안에서 여러 번 반복될 수 있는 본문 섹션 제목이다.", prompt)
        self.assertIn("`대제목`은 배경, 현황, 운영 방안, 기대 효과처럼 본문을 여러 장으로 나누는 데 사용한다.", prompt)
        self.assertIn("[구조 예시]", prompt)
        self.assertIn("* `대제목 배경 및 목적`", prompt)
        self.assertIn("* `소제목 추진 배경`", prompt)
        self.assertIn("* `원 제도 도입 필요`", prompt)
        self.assertIn("[계층 규칙]", prompt)
        self.assertIn("`대제목 → 소제목 → 본문1`", prompt)
        self.assertNotIn("표내용 문단", prompt)
        self.assertIn("행정 보고서용 개조식 문체", prompt)
        self.assertIn("`~함`, `~필요`, `~예정`, `~추진`", prompt)
        self.assertIn("표만 `markdown` 코드블록 안에 작성한다.", prompt)
        self.assertIn("[출력 전 확인]", prompt)
        self.assertIn("모든 문단이 `대제목`, `소제목`, `원` 중 하나로 시작하는지 확인한다.", prompt)

    def test_draw_title_number_box_uses_table_actions_and_cell_moves(self) -> None:
        calls: list[tuple[str, object]] = []
        app = object.__new__(MvpApp)
        app.table_settings = {}
        app.create_table_at_cursor = lambda rows, cols: calls.append(("create", (rows, cols))) or ("TableCreate", True)
        app.title_number_box_settings = lambda: {"marker": "{{bufs_title}}", "table_style_preset": "thin_all"}
        app.title_number_box_cell_specs = lambda: [{}, {}, {}]
        app.is_in_table_cell = lambda: calls.append(("in_cell", True)) or True
        app.select_current_table_object = lambda: calls.append(("select_object", True)) or True
        app.select_selected_table_first_cell = lambda: calls.append(("select_first", True)) or True
        app.apply_title_number_box_cell = lambda spec, text: calls.append(("cell", text)) or True
        app.move_table_cell = lambda command: calls.append(("move", command)) or True
        app.select_current_table_all_cells = lambda rows, cols: calls.append(("select_all", (rows, cols))) or True
        app.set_table_border_preset = lambda preset: calls.append(("border", preset)) or ("CellBorderFill", True)
        app.run_hwp_command_best_effort = lambda command: calls.append(("command", command)) or True
        app.clear_hwp_selection = lambda: calls.append(("clear", None)) or True
        app.apply_current_table_outside_margins = lambda values: calls.append(("outside", values)) or True
        app.title_number_box_outside_margin_values = lambda: {"OutsideMarginBottom": 850}
        app.debug = lambda _message: None

        self.assertTrue(app.draw_title_number_box_at_cursor("  Title\r\nHere  ", 7))

        self.assertEqual(
            calls,
            [
                ("create", (1, 3)),
                ("in_cell", True),
                ("cell", "7"),
                ("move", "TableRightCell"),
                ("cell", "Title Here"),
                ("move", "TableRightCell"),
                ("cell", "{{bufs_title}}"),
                ("select_all", (1, 3)),
                ("border", "thin_all"),
                ("command", "TableVAlignCenter"),
                ("clear", None),
                ("select_object", True),
                ("outside", {"OutsideMarginBottom": 850}),
                ("clear", None),
            ],
        )

    def test_apply_title_number_box_cell_applies_dimensions_fill_style_color_then_text(self) -> None:
        calls: list[tuple[str, object]] = []
        app = object.__new__(MvpApp)
        app.apply_current_cell_dimensions = lambda width, height, margins: calls.append(("dimensions", (width, height, margins))) or True
        app.set_cell_fill_color = lambda color: calls.append(("fill", color.rgb)) or ("CellBorderFill", True, True)
        app.palette_color_or_rgb = lambda name, rgb: calls.append(("palette", (name, rgb))) or hwp_style_mvp.PaletteColor(name, rgb, "#000000", "")
        app.clear_hwp_selection = lambda: calls.append(("clear", None)) or True
        app.apply_style_by_name = lambda name: calls.append(("style", name)) or True
        app.set_text_color = lambda color: calls.append(("text_color", color.rgb)) or (True, True)
        app.insert_hwp_text = lambda text: calls.append(("text", text)) or True

        cell_spec = {
            "width": 2487,
            "height": 2414,
            "paragraph_style": "바탕글 사본1",
            "fill_rgb": [217, 217, 217],
            "text_color_rgb": [255, 255, 255],
            "margins": {"left": 510, "right": 510, "top": 141, "bottom": 141},
        }

        self.assertTrue(app.apply_title_number_box_cell(cell_spec, "7"))
        self.assertEqual(
            calls,
            [
                ("dimensions", (2487, 2414, {"left": 510, "right": 510, "top": 141, "bottom": 141})),
                ("palette", ("대제목번호박스배경", (217, 217, 217))),
                ("fill", (217, 217, 217)),
                ("clear", None),
                ("style", "바탕글 사본1"),
                ("palette", ("대제목번호박스글자", (255, 255, 255))),
                ("text_color", (255, 255, 255)),
                ("text", "7"),
            ],
        )

    def test_insert_title_number_box_file_at_cursor_uses_filled_template_insert_action(self) -> None:
        calls: list[tuple[str, object]] = []
        original = hwp_style_mvp.create_filled_title_number_box_file

        def fake_create(title: str, number: int, template, **kwargs):
            calls.append(("create_file", (title, number, template, kwargs)))
            return Path("C:/temp/title-box.hwpx")

        try:
            hwp_style_mvp.create_filled_title_number_box_file = fake_create
            app = object.__new__(MvpApp)
            app.title_number_box_settings = lambda entry=None: {"template_path": Path("C:/templates/title.hwpx")}
            app.insert_file_at_cursor = lambda path, keep_section=0: calls.append(("insert_file", (path, keep_section))) or True

            self.assertTrue(app.insert_title_number_box_file_at_cursor("대제목 테스트", 4))
        finally:
            hwp_style_mvp.create_filled_title_number_box_file = original

        self.assertEqual(
            calls,
            [
                ("create_file", ("대제목 테스트", 4, Path("C:/templates/title.hwpx"), {"marker": hwp_style_mvp.TITLE_NUMBER_BOX_MARKER})),
                ("insert_file", (Path("C:/temp/title-box.hwpx"), 0)),
            ],
        )

    def test_replace_paragraph_with_title_number_box_inserts_template_file(self) -> None:
        calls: list[tuple[str, object]] = []
        app = object.__new__(MvpApp)
        entry = StyleEntry("대제목", roles=("title_number_box",), template_file="templates/box-a.hwpx")
        app.next_title_number_box_number = lambda **kwargs: calls.append(("number", kwargs)) or (4, [1, 3])
        app.read_current_paragraph_text = lambda list_id, para: "대제목 테스트"
        app.delete_hwp_text_range = lambda list_id, para, start, end: calls.append(("delete", (list_id, para, start, end))) or True
        app.set_hwp_pos = lambda pos: calls.append(("pos", pos)) or True
        app.insert_title_number_box_file_at_cursor = (
            lambda title, number, style_entry=None: calls.append(("insert_file", (title, number, style_entry.template_file if style_entry else ""))) or True
        )

        result = app.replace_paragraph_with_title_number_box(2, 5, "  대제목\r\n테스트  ", entry)

        self.assertEqual(result, (True, 4, [1, 3]))
        self.assertEqual(
            calls,
            [
                ("number", {"list_id": 2, "para": 5, "entry": entry}),
                ("delete", (2, 5, 0, len("대제목 테스트"))),
                ("pos", (2, 5, 0)),
                ("insert_file", ("대제목 테스트", 4, "templates/box-a.hwpx")),
            ],
        )

    def test_title_number_box_settings_prefers_style_template_file_over_set_default(self) -> None:
        app = object.__new__(MvpApp)
        app.table_settings = {"title_number_box": {"template_file": "templates/global.hwpx"}}
        app.active_style_set = lambda: StyleSet(
            "set",
            [StyleEntry("대제목", roles=("title_number_box",))],
            [],
            role_configs={"title_number_box": {"template_file": "templates/default.hwpx"}},
        )

        settings = app.title_number_box_settings(
            StyleEntry("대제목A", roles=("title_number_box",), template_file="templates/custom.hwpx")
        )

        self.assertEqual(settings["template_file"], "templates/custom.hwpx")
        self.assertEqual(settings["template_path"], hwp_style_mvp.BUILTIN_TEMPLATE_DIR / "custom.hwpx")

    def test_find_outline_style_rule_prefers_longer_marker(self) -> None:
        style_set = StyleSet(
            "보고서용",
            [
                StyleEntry("본문-개요2", outline_markers=("원",)),
                StyleEntry("본문-개요3", outline_markers=("원기호",)),
            ],
            [],
        )

        entry, prefix_length = find_outline_style_rule("  원기호 세부 내용", style_set)

        self.assertEqual(entry.name, "본문-개요3")
        self.assertEqual(prefix_length, len("  원기호 "))

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

    def test_apply_style_by_name_uses_existing_style_cache_without_refreshing(self) -> None:
        app = object.__new__(MvpApp)
        record = StyleRecord(style_id=7, style_index=1, style_type="PARA", name="표내용-중간")
        app.current_doc_style_map = {"표내용-중간": record}
        app.current_doc_style_norm_map = {}
        app.refresh_current_doc_style_map = lambda *, force=False: self.fail("cached style apply should not refresh")
        app.debug = lambda _message: None
        applied: list[StyleRecord] = []
        app.execute_style_record = lambda item: applied.append(item) or True

        self.assertTrue(app.apply_style_by_name("표내용-중간"))
        self.assertEqual(applied, [record])

    def test_apply_first_style_containing_does_not_fallback_to_base_style_records(self) -> None:
        app = object.__new__(MvpApp)
        app.refresh_current_doc_style_map = lambda *, force=False: None
        app.current_doc_style_map = {}
        app.style_records = [StyleRecord(style_id=99, style_index=1, style_type="PARA", name="[표/그림] 캡션")]
        app.debug = lambda _message: None
        app.execute_style_record = lambda _record: self.fail("현재 문서 스타일 맵 없이 기준 번호 fallback을 쓰면 안 됩니다")

        self.assertFalse(app.apply_first_style_containing("캡션"))

    def test_table_content_style_records_include_all_para_and_char_table_content_styles(self) -> None:
        app = object.__new__(MvpApp)
        app.style_records = [
            StyleRecord(1, 1, "PARA", "본문"),
            StyleRecord(8, 8, "PARA", "\ufeff표내용-숨은문자"),
            StyleRecord(2, 2, "PARA", "표내용-중간"),
            StyleRecord(3, 3, "CHAR", "표내용-굵게"),
            StyleRecord(4, 4, "PARA", "표내용-오른쪽"),
            StyleRecord(5, 5, "PARA", "표내용-왼쪽"),
            StyleRecord(6, 6, "PARA", "표내용-가운데"),
            StyleRecord(7, 7, "PARA", "표내용-여분"),
        ]

        self.assertEqual(
            [record.name for record in app.table_content_style_records()],
            [
                "\ufeff표내용-숨은문자",
                "표내용-중간",
                "표내용-굵게",
                "표내용-오른쪽",
                "표내용-왼쪽",
                "표내용-가운데",
                "표내용-여분",
            ],
        )
        self.assertEqual(app.table_content_style_button_label(app.style_records[1]), "숨은문자")
        self.assertEqual(app.table_content_style_button_label(app.style_records[3]), "굵게")

    def test_table_content_style_records_include_checked_table_styles(self) -> None:
        app = object.__new__(MvpApp)
        app.style_records = [
            StyleRecord(1, 1, "PARA", "표 셀 기본"),
            StyleRecord(2, 2, "CHAR", "강조 문자"),
            StyleRecord(3, 3, "PARA", "본문"),
        ]
        app.active_style_set_name = "보고서용"
        app.style_sets = [
            StyleSet(
                "보고서용",
                [StyleEntry("표 셀 기본", table_style=True), StyleEntry("본문")],
                [StyleEntry("강조 문자", table_style=True)],
            )
        ]

        self.assertEqual(
            [record.name for record in app.table_content_style_records()],
            ["표 셀 기본", "강조 문자"],
        )

    def test_caption_style_record_uses_checked_caption_style(self) -> None:
        app = object.__new__(MvpApp)
        app.style_records = [
            StyleRecord(1, 1, "PARA", "본문"),
            StyleRecord(2, 2, "PARA", "캡션 스타일"),
        ]
        app.active_style_set_name = "보고서용"
        app.style_sets = [
            StyleSet(
                "보고서용",
                [StyleEntry("본문"), StyleEntry("캡션 스타일", caption_style=True)],
                [],
            )
        ]

        self.assertEqual(app.caption_style_record().name, "캡션 스타일")

    def test_caption_parser_templates_can_use_custom_named_groups(self) -> None:
        settings = {
            "pattern": r"^\[\s*(?P<kind>표|그림)(?P<gap>\s*)(?P<prefix>(?:\d+(?:\.\d+)*\s*)?-)\s*(?P<number>\d+)?\s*\]\s*(?P<title>.*)$",
            "prefix_template": "[{kind}{gap}{prefix}",
            "suffix_template": "] {title}",
        }

        spaced = split_table_caption_parts("[그림 2.4-1] 학사관리 체계 개선 실적", settings)
        compact = split_table_caption_parts("[표-17] 제목", settings)

        self.assertEqual(spaced.prefix, "[그림 2.4-")
        self.assertEqual(spaced.suffix, "] 학사관리 체계 개선 실적")
        self.assertEqual(compact.prefix, "[표-")
        self.assertEqual(compact.suffix, "] 제목")

    def test_table_settings_upgrade_legacy_default_caption_parser(self) -> None:
        settings = {
            "caption_parser": {
                "pattern": r"^\[\s*(?P<kind>표|그림)\s+(?P<prefix>[^\]]*-)\s*(?:\d+)?\s*\]\s*(?P<title>.*)$",
                "prefix_template": "[{kind} {prefix}",
                "suffix_template": "] {title}",
            }
        }

        normalized = normalize_table_settings(merge_dict(DEFAULT_TABLE_SETTINGS, settings))
        caption_parser = normalized["caption_parser"]

        self.assertIn("(?P<gap>", caption_parser["pattern"])
        self.assertIn("(?P<number>", caption_parser["pattern"])
        self.assertEqual(caption_parser["prefix_template"], "[{kind}{gap}{prefix}")

    def test_table_settings_upgrade_numbered_default_caption_prefix_template(self) -> None:
        settings = {
            "caption_parser": {
                "pattern": r"^\[\s*(?P<kind>표|그림)(?P<gap>\s*)(?P<prefix>(?:\d+(?:\.\d+)*\s*)?-)\s*(?P<number>\d+)?\s*\]\s*(?P<title>.*)$",
                "prefix_template": "[{kind}{gap}{prefix}{number}",
                "suffix_template": "] {title}",
            }
        }

        normalized = normalize_table_settings(merge_dict(DEFAULT_TABLE_SETTINGS, settings))

        self.assertEqual(normalized["caption_parser"]["prefix_template"], "[{kind}{gap}{prefix}")

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

    def test_refresh_current_doc_style_map_reads_memory_styles_for_unsaved_document(self) -> None:
        class FakeHwp:
            Path = ""

            def GetTextFile(self, _format_name, _option):
                return """<HWPML><HEAD>
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

    def test_set_table_cell_margins_uses_table_property_dialog_shape_table_cell(self) -> None:
        class FakeSet:
            def __init__(self) -> None:
                self.items = {}
                self.shape_table_cell = FakeSubset()

            def SetItem(self, name, value):
                self.items[name] = value

            def Item(self, name):
                if name == "ShapeTableCell":
                    return self.shape_table_cell
                raise KeyError(name)

        class FakeSubset:
            def __init__(self) -> None:
                self.items = {}

            def SetItem(self, name, value):
                self.items[name] = value

        class FakeAction:
            def __init__(self) -> None:
                self.set = FakeSet()
                self.default_sets = []
                self.executed_sets = []

            def CreateSet(self):
                return self.set

            def GetDefault(self, hset):
                self.default_sets.append(hset)
                return True

            def Execute(self, hset):
                self.executed_sets.append(hset)
                return True

        action = FakeAction()

        class FakeHwp:
            def __init__(self) -> None:
                self.action_names = []

            def CreateAction(self, name):
                self.action_names.append(name)
                return action

            def MiliToHwpUnit(self, value):
                return int(value * 283465 / 1000)

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.table_settings = {
            "table_margins": {
                "cell_left": "2",
                "cell_right": "3",
                "cell_top": "4",
                "cell_bottom": "5",
            }
        }
        app.debug = lambda _message: None

        action_name, ok = app.set_table_cell_margins()

        self.assertEqual(action_name, "CellShape+CreateAction(TablePropertyDialog)")
        self.assertTrue(ok)
        self.assertEqual(app.hwp.action_names, ["TablePropertyDialog"])
        self.assertEqual(action.set.items["ShapeType"], 3)
        self.assertEqual(action.set.items["ShapeCellSize"], 0)
        self.assertEqual(action.set.items["CellMarginLeft"], 566)
        self.assertEqual(action.set.items["CellMarginRight"], 850)
        self.assertEqual(action.set.items["CellMarginTop"], 1133)
        self.assertEqual(action.set.items["CellMarginBottom"], 1417)
        self.assertEqual(action.set.shape_table_cell.items["HasMargin"], 1)
        self.assertEqual(action.set.shape_table_cell.items["MarginLeft"], 566)
        self.assertEqual(action.set.shape_table_cell.items["MarginRight"], 850)
        self.assertEqual(action.set.shape_table_cell.items["MarginTop"], 1133)
        self.assertEqual(action.set.shape_table_cell.items["MarginBottom"], 1417)
        self.assertEqual(action.executed_sets, [action.set])

    def test_set_table_cell_margins_commits_cell_shape_cell_margin_items(self) -> None:
        class FakeSubset:
            def __init__(self) -> None:
                self.items = {}

            def SetItem(self, name, value):
                self.items[name] = value

            def Item(self, name):
                return self.items.get(name)

        class FakeCellShape:
            def __init__(self) -> None:
                self.items = {"Cell": FakeSubset()}

            def SetItem(self, name, value):
                self.items[name] = value

            def Item(self, name):
                return self.items.get(name)

        cell_shape = FakeCellShape()
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"CellShape": cell_shape})()
        app.table_settings = {
            "table_margins": {
                "cell_left": "2",
                "cell_right": "2",
                "cell_top": "2",
                "cell_bottom": "2",
            }
        }
        app.mm_to_hwpunit = lambda value: int(float(value) * 100)
        app.debug = lambda _message: None

        action_name, ok = app.set_table_cell_margins_by_cell_shape()

        self.assertEqual(action_name, "CellShape")
        self.assertTrue(ok)
        self.assertIs(app.hwp.CellShape, cell_shape)
        self.assertEqual(cell_shape.items["CellMarginLeft"], 200)
        self.assertEqual(cell_shape.items["CellMarginRight"], 200)
        self.assertEqual(cell_shape.items["CellMarginTop"], 200)
        self.assertEqual(cell_shape.items["CellMarginBottom"], 200)
        self.assertEqual(cell_shape.items["Cell"].items["HasMargin"], 1)
        self.assertEqual(cell_shape.items["Cell"].items["MarginLeft"], 200)
        self.assertEqual(cell_shape.items["Cell"].items["MarginRight"], 200)
        self.assertEqual(cell_shape.items["Cell"].items["MarginTop"], 200)
        self.assertEqual(cell_shape.items["Cell"].items["MarginBottom"], 200)

    def test_markdown_table_post_formatting_runs_cell_steps_before_object_steps(self) -> None:
        app = object.__new__(MvpApp)
        calls: list[str] = []
        app.set_table_page_width = lambda: calls.append("width") or ("HWPML2X", True, 1000)
        app.select_current_table_object = lambda: calls.append("select_object") or True
        app.set_table_outside_margins = lambda: calls.append("outside") or ("TablePropertyDialog", True)
        app.select_current_table_all_cells = lambda rows=None, cols=None: calls.append(f"select_cells:{rows}x{cols}") or True
        app.set_table_cell_margins = lambda: calls.append("cell_margin") or ("CellShape", True)
        app.apply_table_style_preset_steps = lambda preset: calls.append(f"table_style:{preset}") or (["ok"], True, None)
        app.clear_hwp_selection = lambda: calls.append("clear") or True
        app.table_settings = {"markdown_table": {"paragraph_style": "(적용 안 함)", "table_style_preset": "thin_top_bottom"}}
        app.debug = lambda _message: None

        results = app.apply_markdown_table_post_formatting("Markdown 표 → 한글 표", 3, 2)

        self.assertEqual(
            calls,
            [
                "select_cells:3x2",
                "cell_margin",
                "table_style:thin_top_bottom",
                "clear",
                "select_object",
                "outside",
                "width",
            ],
        )
        self.assertEqual(results, [("셀여백", True), ("표스타일:얇은상하", True), ("표여백", True), ("너비맞추기", True)])

    def test_markdown_table_post_formatting_applies_configured_paragraph_style(self) -> None:
        app = object.__new__(MvpApp)
        calls: list[str] = []
        app.table_settings = {
            "markdown_table": {
                "paragraph_style": "표내용-중간",
                "table_style_preset": "thick_top_bottom_header_double",
            }
        }
        app.set_table_page_width = lambda: calls.append("width") or ("HWPML2X", True, 1000)
        app.select_current_table_object = lambda: calls.append("select_object") or True
        app.set_table_outside_margins = lambda: calls.append("outside") or ("TablePropertyDialog", True)
        app.select_current_table_all_cells = lambda rows=None, cols=None: calls.append(f"select_cells:{rows}x{cols}") or True
        app.apply_style_by_name = lambda style_name: calls.append(f"style:{style_name}") or True
        app.set_table_cell_margins = lambda: calls.append("cell_margin") or ("CellShape", True)
        app.apply_table_style_preset_steps = lambda preset: calls.append(f"table_style:{preset}") or (["ok"], True, None)
        app.clear_hwp_selection = lambda: calls.append("clear") or True
        app.debug = lambda _message: None

        results = app.apply_markdown_table_post_formatting("Markdown 표 → 한글 표", 3, 2)

        self.assertEqual(
            calls,
            [
                "select_cells:3x2",
                "style:표내용-중간",
                "cell_margin",
                "table_style:thick_top_bottom_header_double",
                "clear",
                "select_object",
                "outside",
                "width",
            ],
        )
        self.assertEqual(
            results,
            [
                ("문단스타일:표내용-중간", True),
                ("셀여백", True),
                ("표스타일:굵은상하이중선", True),
                ("표여백", True),
                ("너비맞추기", True),
            ],
        )

    def test_selected_table_object_can_switch_to_all_cell_selection(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = object()
        app.has_selected_object = lambda: True
        app.is_in_table_cell = lambda: True
        app.is_selected_cell_block = lambda: True
        app.activate_hwp_window = lambda: True
        app.clear_hwp_selection = lambda: True
        app.debug = lambda _message: None
        calls: list[str] = []

        def fake_run(command):
            calls.append(command)
            return command in {
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableCellBlockCol",
                "TableCellBlockRow",
            }

        app.run_hwp_command = fake_run

        self.assertTrue(app.select_current_table_all_cells())
        self.assertEqual(
            calls,
            ["ShapeObjTableSelCell", "TableCellBlock", "TableCellBlockRow", "TableCellBlockCol", "TableCellBlockExtend"],
        )

    def test_selected_table_object_can_select_markdown_dimensions(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = object()
        app.has_selected_object = lambda: True
        app.is_in_table_cell = lambda: True
        app.is_selected_cell_block = lambda: True
        app.activate_hwp_window = lambda: True
        app.debug = lambda _message: None
        calls: list[str] = []

        def fake_run(command):
            calls.append(command)
            return command in {
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableRightCell",
                "TableLowerCell",
            }

        app.run_hwp_command = fake_run

        self.assertTrue(app.select_current_table_all_cells(2, 3))
        self.assertEqual(
            calls,
            [
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableLowerCell",
                "TableRightCell",
                "TableRightCell",
            ],
        )

    def test_markdown_dimension_selection_continues_when_cell_block_command_returns_false_but_state_is_cell_block(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = object()
        app.has_selected_object = lambda: True
        app.is_in_table_cell = lambda: True
        app.is_selected_cell_block = lambda: True
        app.activate_hwp_window = lambda: True
        app.debug = lambda _message: None
        calls: list[str] = []

        def fake_run(command):
            calls.append(command)
            if command == "TableCellBlock":
                return False
            return command in {
                "ShapeObjTableSelCell",
                "TableCellBlockExtend",
                "TableRightCell",
                "TableLowerCell",
            }

        app.run_hwp_command = fake_run

        self.assertTrue(app.select_current_table_all_cells(2, 2))
        self.assertEqual(
            calls,
            [
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableLowerCell",
                "TableRightCell",
            ],
        )

    def test_current_selected_table_dimensions_counts_colspan_cells(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None

        class FakeHwp:
            def GetTextFile(self, _format, _option):
                return """
<HWPML>
  <BODY>
    <TABLE>
      <ROW>
        <CELL RowSpan="2"/>
        <CELL ColSpan="4"/>
      </ROW>
      <ROW>
        <CELL/><CELL/><CELL/><CELL/>
      </ROW>
      <ROW>
        <CELL/><CELL/><CELL/><CELL/><CELL/>
      </ROW>
    </TABLE>
  </BODY>
</HWPML>
"""

        app.hwp = FakeHwp()

        self.assertEqual(app.current_selected_table_dimensions(), (3, 5))

    def test_select_current_table_all_cells_variant_uses_hwpml_dimensions_when_available(self) -> None:
        app = object.__new__(MvpApp)
        app.has_selected_object = lambda: True
        app.is_in_table_cell = lambda: True
        app.is_selected_cell_block = lambda: True
        app.activate_hwp_window = lambda: True
        app.debug = lambda _message: None
        calls: list[str] = []

        class FakeHwp:
            def GetTextFile(self, _format, _option):
                return """
<HWPML><BODY><TABLE>
  <ROW><CELL RowSpan="2"/><CELL ColSpan="4"/></ROW>
  <ROW><CELL/><CELL/><CELL/><CELL/></ROW>
  <ROW><CELL/><CELL/><CELL/><CELL/><CELL/></ROW>
</TABLE></BODY></HWPML>
"""

        app.hwp = FakeHwp()

        def fake_run(command):
            calls.append(command)
            return command in {
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableLowerCell",
                "TableRightCell",
            }

        app.run_hwp_command = fake_run

        ok, steps = app.select_current_table_all_cells_by_variant(horizontal_first=False)

        self.assertTrue(ok)
        self.assertIn("dimensions=(3, 5)", steps)
        self.assertEqual(
            calls,
            [
                "ShapeObjTableSelCell",
                "TableCellBlock",
                "TableCellBlockExtend",
                "TableLowerCell",
                "TableLowerCell",
                "TableRightCell",
                "TableRightCell",
                "TableRightCell",
                "TableRightCell",
            ],
        )

    def test_fill_current_table_with_rows_moves_row_by_row(self) -> None:
        app = object.__new__(MvpApp)
        moves: list[tuple[str, int]] = []
        inserted: list[str] = []
        app.is_in_table_cell = lambda: True
        app.select_current_table_object = lambda: self.fail("table object selection should not run")
        app.select_selected_table_first_cell = lambda: self.fail("first-cell selection should not run")
        app.clear_hwp_selection = lambda: True
        app.move_table_cell = lambda command, count=1: moves.append((command, count)) or True
        app.insert_hwp_text = lambda text: inserted.append(text) or True

        ok = app.fill_current_table_with_rows([["A", "B"], ["1", "2"]])

        self.assertTrue(ok)
        self.assertEqual(inserted, ["A", "B", "1", "2"])
        self.assertEqual(
            moves,
            [
                ("TableRightCell", 1),
                ("TableLowerCell", 1),
                ("TableLeftCell", 1),
                ("TableRightCell", 1),
            ],
        )

    def test_fill_current_table_with_rows_converts_newline_tokens(self) -> None:
        app = object.__new__(MvpApp)
        inserted: list[str] = []
        app.is_in_table_cell = lambda: True
        app.select_current_table_object = lambda: True
        app.select_selected_table_first_cell = lambda: True
        app.clear_hwp_selection = lambda: True
        app.move_table_cell = lambda _command, count=1: True
        app.insert_hwp_text = lambda text: inserted.append(text) or True

        ok = app.fill_current_table_with_rows([["첫줄<br>둘째줄", r"가\n나"]])

        self.assertTrue(ok)
        self.assertEqual(inserted, ["첫줄\r\n둘째줄", "가\r\n나"])

    def test_markdown_rows_to_html_table_converts_newline_tokens_to_br(self) -> None:
        html = markdown_rows_to_html_table([["항목", "내용"], ["A", "첫줄<br>둘째줄"], ["B", r"가\n나"]])

        self.assertIn("<td>첫줄<br/>둘째줄</td>", html)
        self.assertIn("<td>가<br/>나</td>", html)

    def test_find_markdown_table_block_returns_longest_contiguous_block(self) -> None:
        block = find_markdown_table_block(
            [
                "| 번호 | 내용 |",
                "|---|---|",
                "| 1 | 첫 줄 |",
                "| 2 | 둘째 줄 |",
                "다음 문단",
            ]
        )

        self.assertEqual(block, (0, 3, [["번호", "내용"], ["1", "첫 줄"], ["2", "둘째 줄"]]))

    def test_convert_markdown_table_delegates_to_draw_path(self) -> None:
        app = object.__new__(MvpApp)
        calls: list[tuple[str, bool]] = []
        app.draw_selected_markdown_table = (
            lambda *, label="Markdown 표 → 표 그리기", warn_if_missing=True: calls.append((label, warn_if_missing)) or "converted"
        )

        app.convert_selected_markdown_table()

        self.assertEqual(calls, [("Markdown 표 → 한글 표", True)])

    def test_draw_selected_markdown_table_creates_table_after_delete(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: True
        commands: list[str] = []
        app.get_selected_text_positions = lambda: ((0, 2, 0), (0, 4, 0))
        app.run_hwp_command = lambda command: commands.append(command) or command in {"Copy", "Delete"}
        app.get_clipboard_text = lambda: "| A | B |\n|---|---|\n| 1 | 2 |"
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.set_clipboard_text = lambda _text: self.fail("restore should not run")
        app.create_table_at_cursor = lambda rows, cols: ("TableCreate", rows == 2 and cols == 2)
        app.fill_current_table_with_rows = lambda rows: rows == [["A", "B"], ["1", "2"]]
        formatted: list[str] = []
        app.apply_markdown_table_post_formatting = lambda label, rows, cols: formatted.append((label, rows, cols)) or [
            ("셀여백", True),
            ("얇은상하", True),
            ("표여백", True),
            ("너비맞추기", True),
        ]

        app.draw_selected_markdown_table()

        self.assertEqual(commands, ["Copy", "Delete"])
        self.assertEqual(formatted, [("Markdown 표 → 표 그리기", 2, 2)])

    def test_bulk_processing_converts_inline_markdown_block_and_continues(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = None
        app.ensure_hwp = lambda: True
        app.is_selected_cell_block = lambda: False
        app.selected_paragraph_range = lambda: (0, 0, 5)
        app.current_paragraph_range = lambda: None
        app.ensure_current_doc_style_cache = lambda: None
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.active_style_set_name = "set"
        app.has_title_number_boxes_after_paragraph = lambda _list_id, _para: False
        app.style_sets = [StyleSet("set", [StyleEntry("본문-개요2", outline_markers=("ㅇ",))], [])]
        paragraphs = {
            0: "대제목 테스트",
            1: "소제목 테스트",
            2: "| 번호 | 내용 |",
            3: "|---|---|",
            4: "| 1 | 값 |",
            5: "ㅇ 마지막 문단",
        }
        app.read_current_paragraph_text = lambda _list_id, para: paragraphs.get(para)
        replaced: list[tuple[int, int, str, list[list[str]]]] = []

        def replace_markdown_table_block(label, list_id, start_para, end_para, source_text, rows):
            replaced.append((start_para, end_para, source_text, rows))
            paragraphs[2] = "[TABLE]"
            paragraphs[3] = "ㅇ 마지막 문단"
            paragraphs.pop(4, None)
            paragraphs.pop(5, None)
            return "converted"

        app.replace_markdown_table_block = replace_markdown_table_block
        processed: list[int] = []

        def fake_process(_label, _list_id, para, _active_set, _missing_styles, _missing_roles, **_kwargs):
            processed.append(para)
            return True, para == 3, para == 3, 0

        app.process_configured_style_paragraph = fake_process

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(processed, [0, 1, 3])
        self.assertEqual(
            replaced,
            [(2, 4, "| 번호 | 내용 |\n|---|---|\n| 1 | 값 |", [["번호", "내용"], ["1", "값"]])],
        )
        self.assertEqual(restored, ["original"])

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

    def test_current_cell_address_prefers_keyindicator_when_com_addr_disagrees(self) -> None:
        class FakeHwp:
            def KeyIndicator(self):
                return ("", "(A15): 표 안 입력")

            def GetTableCellAddr(self, index):
                return (15, 1)[index]

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        messages: list[str] = []
        app.debug = messages.append

        self.assertEqual(app.get_current_cell_address(), (1, 15))
        self.assertTrue(any("불일치" in message for message in messages))

    def test_multiline_dates_are_cell_like_for_single_column_selection(self) -> None:
        self.assertTrue(looks_like_cell_clipboard_matrix("2026.7.22\r\n2026-07-23\r\n"))
        self.assertTrue(looks_like_cell_clipboard_matrix("19910114\r\n910114\r\n"))

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

    def test_remove_commas_handles_korean_unit_suffixes(self) -> None:
        self.assertEqual(remove_number_commas("1,250,000원"), "1250000원")
        self.assertEqual(remove_number_commas("2,500천원"), "2500천원")
        self.assertEqual(remove_number_commas("-1,234.56원"), "-1234.56원")

    def test_normalize_decimal_numbers_rounding_modes(self) -> None:
        self.assertEqual(
            normalize_decimal_numbers("1234.567 -1234.561", 2, "반올림"),
            "1,234.57 -1,234.56",
        )
        self.assertEqual(
            normalize_decimal_numbers("1234.561 -1234.561", 2, "올림"),
            "1,234.57 -1,234.57",
        )
        self.assertEqual(
            normalize_decimal_numbers("1234.569 -1234.569", 2, "버림"),
            "1,234.56 -1,234.56",
        )

    def test_normalize_decimal_numbers_can_skip_commas_and_preserve_dates(self) -> None:
        self.assertEqual(normalize_decimal_numbers("1,234.567", 1, "반올림", use_commas=False), "1234.6")
        self.assertEqual(
            normalize_decimal_numbers("2026. 7. 22. 051-123-4567 12.345", 1, "반올림"),
            "2026. 7. 22. 051-123-4567 12.3",
        )

    def test_find_suspicious_decimal_numbers_detects_bad_commas(self) -> None:
        text = "정상 1,234 1,234.56 오류 2,00,000 1,234,56 123,45 날짜 2026. 7. 22. 전화 051-123-4567"

        self.assertEqual(find_suspicious_decimal_numbers(text), ["2,00,000", "1,234,56", "123,45"])

    def test_decimal_preflight_blocks_suspicious_numbers(self) -> None:
        with self.assertRaises(SuspiciousDecimalNumberError) as ctx:
            ensure_no_suspicious_decimal_numbers("매출 2,00,000 비용 1,234,56")

        self.assertEqual(ctx.exception.values, ["2,00,000", "1,234,56"])

    def test_decimal_preflight_allows_valid_commas_dates_and_phones(self) -> None:
        ensure_no_suspicious_decimal_numbers("1,234 1,234.56 2026. 7. 22. 051-123-4567")

    def test_normalize_decimal_numbers_cell_by_cell(self) -> None:
        matrix = [["1234.567", "2026. 7. 22."], ["-1234.561", "051-123-4567"]]

        transformed, changed = transform_cell_matrix(
            matrix,
            lambda text: normalize_decimal_numbers(text, 1, "반올림"),
        )

        self.assertEqual(transformed, [["1,234.6", "2026. 7. 22."], ["-1,234.6", "051-123-4567"]])
        self.assertEqual(changed, 2)

    def test_scale_decimal_numbers_unit_conversion(self) -> None:
        self.assertEqual(
            scale_decimal_numbers("1,234,567 2,345", hwp_style_mvp.Decimal("0.001"), 1, "반올림"),
            "1,234.6 2.3",
        )
        self.assertEqual(
            scale_decimal_numbers("1,234.6 -2.3", hwp_style_mvp.Decimal("1000"), 0, "반올림"),
            "1,234,600 -2,300",
        )

    def test_scale_decimal_numbers_preserves_dates_and_phones(self) -> None:
        self.assertEqual(
            scale_decimal_numbers("2026. 7. 22. 051-123-4567 1200", hwp_style_mvp.Decimal("0.001"), 2, "반올림"),
            "2026. 7. 22. 051-123-4567 1.20",
        )

    def test_scale_decimal_numbers_cell_by_cell(self) -> None:
        matrix = [["1234567", "2026. 7. 22."], ["2000000", "051-123-4567"]]

        transformed, changed = transform_cell_matrix(
            matrix,
            lambda text: scale_decimal_numbers(text, hwp_style_mvp.Decimal("0.000001"), 2, "반올림"),
        )

        self.assertEqual(transformed, [["1.23", "2026. 7. 22."], ["2.00", "051-123-4567"]])
        self.assertEqual(changed, 2)

    def test_unit_conversion_updates_matching_currency_suffix(self) -> None:
        self.assertEqual(
            scale_decimal_numbers_for_unit_conversion(
                "1,234천원 12백만원 500",
                hwp_style_mvp.Decimal("1000"),
                {"천원": "원", "백만원": "천원"},
                0,
                "반올림",
            ),
            "1,234,000원 12,000천원 500,000",
        )
        self.assertEqual(
            scale_decimal_numbers_for_unit_conversion(
                "1,234,000원 2,000천원",
                hwp_style_mvp.Decimal("0.001"),
                {"원": "천원", "천원": "백만원"},
                0,
                "반올림",
            ),
            "1,234천원 2백만원",
        )

    def test_currency_unit_conversion_uses_target_unit_buttons(self) -> None:
        self.assertEqual(
            convert_decimal_numbers_to_currency_unit(
                "1\ubc31\ub9cc 1\ucc9c 500",
                "\uc6d0",
                0,
                "\ubc18\uc62c\ub9bc",
            ),
            "1,000,000\uc6d0 1,000\uc6d0 500\uc6d0",
        )
        self.assertEqual(
            convert_decimal_numbers_to_currency_unit(
                "1,000\uc6d0 1\ubc31\ub9cc",
                "\ucc9c",
                3,
                "\ubc18\uc62c\ub9bc",
            ),
            "1.000\ucc9c\uc6d0 1,000.000\ucc9c\uc6d0",
        )
        self.assertEqual(
            convert_decimal_numbers_to_currency_unit(
                "1,000,000\uc6d0 1,000\ucc9c\uc6d0",
                "\ubc31\ub9cc",
                0,
                "\ubc18\uc62c\ub9bc",
            ),
            "1\ubc31\ub9cc\uc6d0 1\ubc31\ub9cc\uc6d0",
        )

    def test_unit_conversion_preflight_blocks_other_currency_units(self) -> None:
        with self.assertRaises(UnsafeUnitConversionNumberError) as ctx:
            ensure_safe_decimal_unit_conversion_text("1,234원 12억원", {"천원", "백만원"})

        self.assertEqual(ctx.exception.values, ["1,234원", "12억원"])

    def test_unit_conversion_preflight_blocks_non_currency_units(self) -> None:
        with self.assertRaises(UnsafeUnitConversionNumberError) as ctx:
            ensure_safe_decimal_unit_conversion_text("12.5% 120m^2 30㎡", {"원"})

        self.assertEqual(ctx.exception.values, ["12.5%", "120m^2", "30㎡"])

    def test_normalize_dates_cell_by_cell(self) -> None:
        matrix = [["2026.7.22", "2026-07-22"]]

        transformed, changed = transform_cell_matrix(matrix, lambda text: normalize_dates(text, "dot_padded"))

        self.assertEqual(transformed, [["2026. 07. 22.", "2026. 07. 22."]])
        self.assertEqual(changed, 2)

    def test_normalize_dates_accepts_compact_yyyymmdd_and_yymmdd(self) -> None:
        self.assertEqual(
            normalize_dates("19910114 910114", "dot_padded"),
            "1991. 01. 14. 1991. 01. 14.",
        )
        self.assertEqual(
            normalize_dates("19910114 260723", "korean"),
            "1991\ub144 1\uc6d4 14\uc77c 2026\ub144 7\uc6d4 23\uc77c",
        )
        self.assertEqual(normalize_dates("19910231 991399", "dot"), "19910231 991399")

    def test_normalize_years_handles_year_and_school_year_without_dates(self) -> None:
        source = (
            "2026\ub144 26\ub144 '26\ub144 "
            "2026\ud559\ub144\ub3c4 26\ud559\ub144\ub3c4 '26\ud559\ub144\ub3c4 "
            "2026\ub144 1\ud559\uae30 26\ud559\ub144\ub3c4 2\ud559\uae30"
        )
        self.assertEqual(
            normalize_years(source, "full"),
            (
                "2026\ub144 2026\ub144 2026\ub144 "
                "2026\ud559\ub144\ub3c4 2026\ud559\ub144\ub3c4 2026\ud559\ub144\ub3c4 "
                "2026\ub144 1\ud559\uae30 2026\ud559\ub144\ub3c4 2\ud559\uae30"
            ),
        )
        self.assertEqual(
            normalize_years(source, "short"),
            (
                "26\ub144 26\ub144 26\ub144 "
                "26\ud559\ub144\ub3c4 26\ud559\ub144\ub3c4 26\ud559\ub144\ub3c4 "
                "26\ub144 1\ud559\uae30 26\ud559\ub144\ub3c4 2\ud559\uae30"
            ),
        )
        self.assertEqual(
            normalize_years(source, "apostrophe"),
            (
                "'26\ub144 '26\ub144 '26\ub144 "
                "'26\ud559\ub144\ub3c4 '26\ud559\ub144\ub3c4 '26\ud559\ub144\ub3c4 "
                "'26\ub144 1\ud559\uae30 '26\ud559\ub144\ub3c4 2\ud559\uae30"
            ),
        )

    def test_normalize_years_skips_year_month_or_day_date_context(self) -> None:
        self.assertEqual(
            normalize_years(
                "2026\ub144 7\uc6d4 26\ub144 7\uc77c '26\ub144 12\uc6d4 2026\ud559\ub144\ub3c4 1\ud559\uae30",
                "apostrophe",
            ),
            "2026\ub144 7\uc6d4 26\ub144 7\uc77c '26\ub144 12\uc6d4 '26\ud559\ub144\ub3c4 1\ud559\uae30",
        )

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

    def test_transform_text_replaces_single_paragraph_selection_by_position_without_paste(self) -> None:
        class FakeHwp:
            SelectionMode = 0

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((7, 3, 25), (7, 3, 36))
        app.is_selected_cell_block = lambda: False
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.get_clipboard_text = lambda: "Article 41-2(2026. 2. 27. amended)"
        app.read_single_paragraph_selection_text = lambda _positions: ("2026. 2. 27.", 25, 36)
        commands: list[str] = []

        def fake_run_hwp_command(command: str) -> bool:
            commands.append(command)
            return command in {"Copy", "TableCellBlock"}

        app.run_hwp_command = fake_run_hwp_command
        app.transform_selected_cell_matrix = lambda *args, **kwargs: self.fail(
            "non-cell selection should not use cell matrix iteration"
        )
        app.set_clipboard_text = lambda _text: self.fail("position replacement should not write clipboard")
        deleted: list[tuple[int, int, int, int]] = []
        inserted: list[str] = []
        app.delete_hwp_text_range = lambda list_id, para, start, end: deleted.append(
            (list_id, para, start, end)
        ) or True
        app.insert_hwp_text = lambda text: inserted.append(text) or True
        app.activate_hwp_window = lambda: None

        app.transform_selected_text(
            "date normalize",
            lambda text: normalize_dates(text, "dot_padded"),
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
            replace_single_paragraph_by_position=True,
        )

        self.assertEqual(deleted, [(7, 3, 25, 36)])
        self.assertEqual(inserted, ["2026. 02. 27."])
        self.assertNotIn("Paste", commands)

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
        app.log = lambda _message: None

        result = app.transform_formula_address_cells(
            "요일 제거",
            [(3, 2), (4, 2)],
            remove_weekdays_from_dates,
        )

        self.assertEqual(result, (2, 1))
        self.assertEqual(moves, [((4, 5), (3, 2)), ((3, 2), (4, 2))])

    def test_formula_address_iteration_moves_by_address_after_changed_cell(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.clear_hwp_selection = lambda: True

        current_addresses = iter([(3, 2), (4, 2), (3, 3)])
        app.get_current_cell_address = lambda: next(current_addresses)
        moves: list[tuple[tuple[int, int], tuple[int, int]]] = []

        def fake_move(current, target):
            moves.append((current, target))
            return True

        app.move_between_table_addresses = fake_move
        values = iter(["2022. 9. 1.", "이미 본문", "2021. 6. 8."])
        app.read_current_cell_text = lambda: next(values)
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

    def test_configured_outline_style_bulk_applies_matching_paragraph_styles(self) -> None:
        app = object.__new__(MvpApp)
        selected_ranges: list[tuple[int, int, int, int]] = []

        class FakeHwp:
            SelectionMode = 1

            def SelectText(self, spara, spos, epara, epos):
                selected_ranges.append((spara, spos, epara, epos))
                return True

        app.hwp = FakeHwp()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 0), (0, 4, 3))
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.ensure_current_doc_style_cache = lambda: None
        app.is_hwp_paragraph_in_table = lambda _list_id, _para: False
        app.active_style_set_name = "보고서용"
        app.style_sets = [
            StyleSet(
                "보고서용",
                [StyleEntry("본문-개요2", outline_markers=("ㅇ", "○"))],
                [],
            )
        ]
        current_record = StyleRecord(97, 1, "PARA", "본문-개요2")
        app.find_current_doc_style_record = lambda name: current_record if name == "본문-개요2" else None
        paragraphs = {2: "ㅇ 첫째", 3: "본문", 4: "○ 둘째"}
        app.read_current_paragraph_text = lambda _list_id, para: paragraphs[para]
        deleted: list[tuple[int, int, int, int]] = []
        def delete_range(list_id, para, start, end):
            deleted.append((list_id, para, start, end))
            paragraphs[para] = paragraphs[para][:start] + paragraphs[para][end:]
            return True

        app.delete_hwp_text_range = delete_range
        positions: list[tuple[int, int, int]] = []
        app.set_hwp_pos = lambda pos: positions.append(pos) or True
        applied: list[StyleRecord] = []
        app.execute_style_record = lambda record: applied.append(record) or True

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(deleted, [(0, 2, 0, len("ㅇ ")), (0, 4, 0, len("○ "))])
        self.assertEqual(selected_ranges, [(2, 0, 2, len("첫째")), (4, 0, 4, len("둘째"))])
        self.assertEqual(applied, [current_record, current_record])
        self.assertEqual(restored, ["original"])

    def test_select_hwp_text_range_offsets_document_start_visible_text(self) -> None:
        app = object.__new__(MvpApp)
        selected_ranges: list[tuple[int, int, int, int]] = []
        positions: list[tuple[int, int, int]] = []

        class FakeHwp:
            def SetPos(self, list_id, para, pos):
                positions.append((list_id, para, pos))
                return True

            def GetPos(self):
                return (0, 0, 16)

            def SelectText(self, spara, spos, epara, epos):
                selected_ranges.append((spara, spos, epara, epos))
                return True

            def GetSelectedPos(self):
                return (True, 0, 0, 16, 0, 0, 21)

        app.hwp = FakeHwp()
        app.clear_hwp_selection = lambda: True
        app.debug = lambda _message: None

        self.assertTrue(app.select_hwp_text_range(0, 0, 0, 5))

        self.assertEqual(selected_ranges, [(0, 16, 0, 21)])
        self.assertEqual(positions[-1], (0, 0, 16))

    def test_configured_outline_style_bulk_normal_selection_does_not_force_table_context(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 0), (0, 3, 3))
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        app.style_sets = [StyleSet("set", [], [], [InlineRule("본문 괄호", "leading_parenthesized_after_identifier")])]
        calls: list[dict] = []

        def fake_process(*args, **kwargs):
            calls.append(kwargs)
            return True, False, False, 1

        app.process_configured_style_paragraph = fake_process

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual([call.get("force_in_table") for call in calls], [None, None])
        self.assertTrue(all(isinstance(call.get("numbering_state"), NumberingRunState) for call in calls))
        self.assertTrue(restored)
        self.assertEqual(restored[-1], "original")

    def test_configured_outline_style_bulk_normal_selection_uses_body_marker_style_when_duplicate(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 0), (0, 2, 4))
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        app.style_sets = [
            StyleSet(
                "set",
                [
                    StyleEntry("본문-개요2", outline_markers=("ㅇ",)),
                    StyleEntry("표내용-개요2", table_style=True, outline_markers=("ㅇ",)),
                ],
                [],
            )
        ]
        app.is_hwp_paragraph_in_table = lambda _list_id, _para: False
        app.read_current_paragraph_text = lambda _list_id, _para: "ㅇ 본문"
        app.delete_hwp_text_range = lambda _list_id, _para, _start, _end: True
        app.apply_style_to_paragraph_text = lambda _list_id, _para, record: applied.append(record.name) or True
        app.apply_inline_rules_to_paragraph = lambda *_args, **_kwargs: 0
        app.find_current_doc_style_record = lambda name: StyleRecord(1, 1, "PARA", name)
        applied: list[str] = []

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(applied, ["본문-개요2"])

    def test_configured_outline_style_bulk_title_box_replaces_paragraph(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 1})()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 0), (0, 2, 6))
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None
        app.ensure_current_doc_style_cache = lambda: None
        app.is_hwp_paragraph_in_table = lambda _list_id, _para: False
        app.active_style_set_name = "set"
        app.style_sets = [
            StyleSet(
                "set",
                [StyleEntry("대제목", outline_markers=("대",), roles=("title_number_box",), template_file="templates/title-a.hwpx")],
                [],
            )
        ]
        app.read_current_paragraph_text = lambda _list_id, _para: "대 사업계획"
        calls: list[tuple[int, int, str, str]] = []
        app.replace_paragraph_with_title_number_box = (
            lambda list_id, para, title, entry=None, forced_number=None: calls.append((list_id, para, title, entry.template_file if entry else "")) or (True, 3, [1, 2])
        )
        app.find_current_doc_style_record = lambda name: self.fail("title_number_box role should not use paragraph style lookup")
        app.apply_inline_rules_to_paragraph = lambda *_args, **_kwargs: self.fail("title_number_box replacement should stop before inline rules")
        app.current_doc_style_reconnect_hint = lambda: ""
        warnings: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))
        try:
            app.apply_configured_outline_styles_to_selection()
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertEqual(calls, [(0, 2, "사업계획", "templates/title-a.hwpx")])
        self.assertEqual(restored, ["original"])
        self.assertEqual(warnings, [])

    def test_process_configured_style_paragraph_restarts_title_box_after_higher_level(self) -> None:
        app = object.__new__(MvpApp)
        texts = {
            0: "대제목 배경 및 목적",
            1: "소제목 일정 및 기간",
            2: "소제목 이동 수단",
        }
        app.read_current_paragraph_text = lambda _list_id, para: texts[para]
        app.is_hwp_paragraph_in_table = lambda _list_id, _para: False
        app.is_title_number_box_entry = lambda entry: True
        app.has_title_number_boxes_after_paragraph = lambda _list_id, _para, _entry=None: False
        app.debug = lambda _message: None
        calls: list[tuple[int, int, str, str, int | None]] = []
        app.replace_paragraph_with_title_number_box = (
            lambda list_id, para, title, entry=None, forced_number=None: (
                calls.append((list_id, para, title, entry.name if entry else "", forced_number)) or (True, forced_number or 2, [])
            )
        )

        active_set = StyleSet(
            "set",
            [
                StyleEntry("대제목", outline_markers=("대제목",), roles=("title_number_box",), numbering_group="body", numbering_level=1),
                StyleEntry(
                    "소제목",
                    outline_markers=("소제목",),
                    roles=("title_number_box",),
                    numbering_group="body",
                    numbering_level=2,
                    restart_after_higher_level=True,
                ),
            ],
            [],
        )
        numbering_state = NumberingRunState()
        missing_styles: set[str] = set()
        missing_roles: set[str] = set()

        self.assertEqual(
            app.process_configured_style_paragraph("일괄처리", 0, 0, active_set, missing_styles, missing_roles, numbering_state=numbering_state),
            (True, True, True, 0),
        )
        self.assertEqual(
            app.process_configured_style_paragraph("일괄처리", 0, 1, active_set, missing_styles, missing_roles, numbering_state=numbering_state),
            (True, True, True, 0),
        )
        self.assertEqual(
            app.process_configured_style_paragraph("일괄처리", 0, 2, active_set, missing_styles, missing_roles, numbering_state=numbering_state),
            (True, True, True, 0),
        )

        self.assertEqual(
            calls,
            [
                (0, 0, "배경 및 목적", "대제목", None),
                (0, 1, "일정 및 기간", "소제목", 1),
                (0, 2, "이동 수단", "소제목", None),
            ],
        )

    def test_configured_outline_style_bulk_falls_back_to_current_paragraph(self) -> None:
        class FakeHwp:
            SelectionMode = 1

            def GetPos(self):
                return (0, 5, 0)

        app = object.__new__(MvpApp)
        app.hwp = FakeHwp()
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: None
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = restored.append
        app.clear_hwp_selection = lambda: True
        app.activate_hwp_window = lambda: None
        logs: list[str] = []
        app.log = logs.append
        app.debug = lambda _message: None
        app.ensure_current_doc_style_cache = lambda: None
        app.is_hwp_paragraph_in_table = lambda _list_id, _para: False
        app.active_style_set_name = "set"
        app.style_sets = [
            StyleSet(
                "set",
                [StyleEntry("대제목", outline_markers=("대제목",), roles=("title_number_box",), template_file="templates/title-b.hwpx")],
                [],
            )
        ]
        app.read_current_paragraph_text = lambda _list_id, _para: "대제목 사업계획"
        calls: list[tuple[int, int, str, str]] = []
        app.replace_paragraph_with_title_number_box = (
            lambda list_id, para, title, entry=None, forced_number=None: calls.append((list_id, para, title, entry.template_file if entry else "")) or (True, 1, [])
        )

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(calls, [(0, 5, "사업계획", "templates/title-b.hwpx")])
        self.assertTrue(any("현재 커서 문단" in message for message in logs))
        self.assertEqual(restored, ["original"])

    def test_configured_outline_style_bulk_routes_cell_block_through_cell_addresses(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        active_set = StyleSet("set", [StyleEntry("body", outline_markers=("o",))], [])
        app.style_sets = [active_set]
        app.get_selected_cell_range_by_formula = lambda: {"addresses": [(1, 1), (2, 1)]}
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = lambda pos: restored.append(pos) or True
        app.snapshot_formula_address_positions = lambda addresses: [((1, 1), "cell-a1"), ((2, 1), "cell-b1")]
        app.scan_current_cell_paragraph_positions = lambda: []
        ranges = iter([(7, 2, 3), (8, 5, 5)])
        app.selected_current_cell_paragraph_range = lambda: next(ranges)
        processed: list[tuple[int, int]] = []

        force_values: list[bool | None] = []

        def fake_process(_label, list_id, para, _active_set, _missing_styles, _missing_roles, **kwargs):
            processed.append((list_id, para))
            force_values.append(kwargs.get("force_in_table"))
            return True, para % 2 == 0, True, 0

        app.process_configured_style_paragraph = fake_process

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(processed, [(7, 2), (7, 3), (8, 5)])
        self.assertEqual(force_values, [True, True, True])
        self.assertEqual(restored, ["cell-a1", "cell-b1", "original"])

    def test_configured_outline_style_bulk_prefers_cell_scan_paragraph_positions(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        active_set = StyleSet("set", [StyleEntry("body", outline_markers=("o",))], [])
        app.style_sets = [active_set]
        app.get_selected_cell_range_by_formula = lambda: {"addresses": [(1, 1), (2, 1)]}
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = lambda pos: restored.append(pos) or True
        app.snapshot_formula_address_positions = lambda addresses: [((1, 1), "cell-a1"), ((2, 1), "cell-b1")]
        scans = iter([[(1799, 0)], [(1802, 0), (1802, 1)]])
        app.scan_current_cell_paragraph_positions = lambda: next(scans)
        app.selected_current_cell_paragraph_range = lambda: self.fail("scan 성공 시 GetSelectedPos fallback을 쓰지 않아야 합니다")
        processed: list[tuple[int, int]] = []

        force_values: list[bool | None] = []

        def fake_process(_label, list_id, para, _active_set, _missing_styles, _missing_roles, **kwargs):
            processed.append((list_id, para))
            force_values.append(kwargs.get("force_in_table"))
            return True, True, True, 0

        app.process_configured_style_paragraph = fake_process

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(processed, [(1799, 0), (1802, 0), (1802, 1)])
        self.assertEqual(force_values, [True, True, True])
        self.assertEqual(restored, ["cell-a1", "cell-b1", "original"])

    def test_init_hwp_scan_retries_with_full_argument_list_for_com_wrappers(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        calls: list[tuple[int, ...]] = []

        class FakeHwp:
            def InitScan(self, *args):
                calls.append(args)
                if len(args) == 2:
                    raise RuntimeError("매개 변수의 개수가 잘못되었습니다.")
                return True

        app.hwp = FakeHwp()

        self.assertTrue(app.init_hwp_scan(0, 0x0055))
        self.assertEqual(calls, [(0, 0x0055), (0, 0x0055, 0, 0, 0, 0)])

    def test_move_hwp_pos_retries_with_full_argument_list_for_com_wrappers(self) -> None:
        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        calls: list[tuple[int, ...]] = []

        class FakeHwp:
            def MovePos(self, *args):
                calls.append(args)
                if len(args) == 1:
                    raise RuntimeError("매개 변수의 개수가 잘못되었습니다.")
                return True

        app.hwp = FakeHwp()

        self.assertTrue(app.move_hwp_pos(201))
        self.assertEqual(calls, [(201,), (201, 0, 0)])

    def test_configured_outline_style_bulk_cell_selection_uses_table_marker_style_when_duplicate(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.activate_hwp_window = lambda: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        app.style_sets = [
            StyleSet(
                "set",
                [
                    StyleEntry("본문-개요2", outline_markers=("ㅇ",)),
                    StyleEntry("표내용-개요2", table_style=True, outline_markers=("ㅇ",)),
                ],
                [],
            )
        ]
        app.get_selected_cell_range_by_formula = lambda: {"addresses": [(1, 1)]}
        app.get_hwp_pos_by_set = lambda: "original"
        app.set_hwp_pos_by_set = lambda _pos: True
        app.snapshot_formula_address_positions = lambda _addresses: [((1, 1), "cell-a1")]
        app.scan_current_cell_paragraph_positions = lambda: [(1802, 0)]
        app.read_current_paragraph_text = lambda _list_id, _para: "ㅇ 표내용"
        app.delete_hwp_text_range = lambda _list_id, _para, _start, _end: True
        app.apply_style_to_paragraph_text = lambda _list_id, _para, record: applied.append(record.name) or True
        app.apply_inline_rules_to_paragraph = lambda *_args, **_kwargs: 0
        app.find_current_doc_style_record = lambda name: StyleRecord(1, 1, "PARA", name)
        applied: list[str] = []

        app.apply_configured_outline_styles_to_selection()

        self.assertEqual(applied, ["표내용-개요2"])

    def test_configured_outline_style_bulk_cell_block_warns_without_addresses(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.debug = lambda _message: None
        app.log = lambda _message: None
        app.active_style_set_name = "set"
        app.style_sets = [StyleSet("set", [StyleEntry("body", outline_markers=("o",))], [])]
        app.get_selected_cell_range_by_formula = lambda: None
        app.ensure_current_doc_style_cache = lambda: self.fail("addressless cell block should stop before cache refresh")
        warnings: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))
        try:
            app.apply_configured_outline_styles_to_selection()
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning

        self.assertTrue(warnings)

    def test_configured_outline_style_bulk_warns_when_cell_paragraph_ranges_fail(self) -> None:
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"SelectionMode": 19})()
        app.ensure_hwp = lambda: True
        app.debug = lambda _message: None
        logs: list[str] = []
        app.log = logs.append
        app.activate_hwp_window = lambda: None
        app.ensure_current_doc_style_cache = lambda: None
        app.active_style_set_name = "set"
        active_set = StyleSet("set", [StyleEntry("body", outline_markers=("o",))], [])
        app.style_sets = [active_set]
        app.get_selected_cell_range_by_formula = lambda: {"addresses": [(1, 1), (2, 1)]}
        app.get_hwp_pos_by_set = lambda: "original"
        restored: list[object] = []
        app.set_hwp_pos_by_set = lambda pos: restored.append(pos) or True
        app.snapshot_formula_address_positions = lambda addresses: [((1, 1), "cell-a1"), ((2, 1), "cell-b1")]
        app.scan_current_cell_paragraph_positions = lambda: []
        app.selected_current_cell_paragraph_range = lambda: None
        processed: list[tuple[int, int]] = []
        app.process_configured_style_paragraph = lambda *args: processed.append((args[1], args[2])) or (True, True, True, 0)
        warnings: list[tuple[str, str]] = []
        infos: list[tuple[str, str]] = []
        original_showwarning = hwp_style_mvp.messagebox.showwarning
        original_showinfo = hwp_style_mvp.messagebox.showinfo
        hwp_style_mvp.messagebox.showwarning = lambda title, message: warnings.append((title, message))
        hwp_style_mvp.messagebox.showinfo = lambda title, message: infos.append((title, message))
        try:
            app.apply_configured_outline_styles_to_selection()
        finally:
            hwp_style_mvp.messagebox.showwarning = original_showwarning
            hwp_style_mvp.messagebox.showinfo = original_showinfo

        self.assertEqual(processed, [])
        self.assertEqual(restored, ["cell-a1", "cell-b1", "original"])
        self.assertTrue(warnings)
        self.assertIn("문단 위치", warnings[0][1])
        self.assertEqual(infos, [])
        self.assertTrue(any("range_failed=2" in message for message in logs))

    def test_wrap_regulation_button_converts_existing_corner_brackets(self) -> None:
        app = object.__new__(MvpApp)
        app.ensure_hwp = lambda: True
        app.get_selected_text_positions = lambda: ((0, 2, 8), (0, 2, 13))
        app.run_hwp_command = lambda command: command == "Copy"
        app.get_clipboard_text = lambda: "「규정명」"
        app.read_single_paragraph_selection_text = lambda _positions: ("「규정명」", 5, 10)
        app.activate_hwp_window = lambda: None
        app.log = lambda _message: None
        app.debug = lambda _message: None

        deleted: list[tuple[int, int, int, int]] = []
        positions: list[tuple[int, int, int]] = []
        inserted: list[str] = []
        app.delete_hwp_text_range = lambda list_id, para, start, end: deleted.append((list_id, para, start, end)) or True
        app.set_hwp_pos = lambda pos: positions.append(pos) or True
        app.insert_hwp_text = lambda text: inserted.append(text) or True

        app.wrap_regulation_names_in_selection()

        self.assertEqual(deleted, [(0, 2, 9, 10), (0, 2, 5, 6)])
        self.assertEqual(positions, [(0, 2, 9), (0, 2, 5)])
        self.assertEqual(inserted, ["｣", "｢"])

    def test_paragraph_auto_right_tab_sets_auto_tab_right(self) -> None:
        class FakeTabDef:
            AutoTabRight = 0

        class FakeParaShape:
            HSet = None

            def __init__(self) -> None:
                self.HSet = self
                self.TabDef = FakeTabDef()

        class FakeAction:
            def __init__(self) -> None:
                self.actions: list[str] = []

            def GetDefault(self, action, _hset):
                self.actions.append(f"default:{action}")
                return True

            def Execute(self, action, _hset):
                self.actions.append(f"execute:{action}")
                return True

        para = FakeParaShape()
        action = FakeAction()
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"HParameterSet": type("Sets", (), {"HParaShape": para})(), "HAction": action})()
        app.debug = lambda _message: None

        action_name, ok = app.set_paragraph_auto_right_tab()

        self.assertEqual(action_name, "ParagraphShape.AutoTabRight")
        self.assertTrue(ok)
        self.assertEqual(para.TabDef.AutoTabRight, 1)
        self.assertEqual(action.actions, ["default:ParagraphShape", "execute:ParagraphShape"])

    def test_paragraph_dotted_right_tab_uses_page_text_width(self) -> None:
        class FakeTabItem:
            def __init__(self) -> None:
                self.values: dict[int, int] = {}

            def SetItem(self, index, value):
                self.values[index] = value

        class FakeTabDef:
            def __init__(self) -> None:
                self.TabItem = FakeTabItem()
                self.array_spec: tuple[str, int] | None = None

            def CreateItemArray(self, name, size):
                self.array_spec = (name, size)

        class FakeParaShape:
            HSet = None
            LeftMargin = 100
            RightMargin = 200

            def __init__(self) -> None:
                self.HSet = self
                self.TabDef = FakeTabDef()

        class FakeAction:
            def GetDefault(self, _action, _hset):
                return True

            def Execute(self, _action, _hset):
                return True

        para = FakeParaShape()
        app = object.__new__(MvpApp)
        app.hwp = type("FakeHwp", (), {"HParameterSet": type("Sets", (), {"HParaShape": para})(), "HAction": FakeAction()})()
        app.debug = lambda _message: None
        app.is_in_table_cell = lambda: False
        app.read_current_table_cell_size = lambda _prefix: (_ for _ in ()).throw(AssertionError("page paragraph should not read cell width"))
        app.current_page_text_width = lambda: 5000

        action_name, ok, tab_position, source = app.set_paragraph_dotted_right_tab()

        self.assertEqual(action_name, "ParagraphShape.TabDef.TabItem")
        self.assertTrue(ok)
        self.assertEqual(tab_position, 10000)
        self.assertIn("쪽", source)
        self.assertEqual(para.TabDef.array_spec, ("TabItem", 3))
        self.assertEqual(para.TabDef.TabItem.values, {0: 10000, 1: 7, 2: 1})

    def test_paragraph_dotted_right_tab_uses_page_width_even_if_cell_shape_is_stale(self) -> None:
        class FakeParaShape:
            HSet = None
            LeftMargin = 100
            RightMargin = 200

            def __init__(self) -> None:
                self.HSet = self

        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.is_in_table_cell = lambda: False
        app.read_current_table_cell_size = lambda _prefix: (_ for _ in ()).throw(AssertionError("page paragraph should not read stale cell width"))
        app.current_page_text_width = lambda: 5000

        width, source = app.current_paragraph_tab_width(FakeParaShape())

        self.assertEqual(width, 10000)
        self.assertIn("쪽", source)

    def test_paragraph_dotted_right_tab_uses_cell_inner_width_inside_table_cell(self) -> None:
        class FakeParaShape:
            HSet = None
            LeftMargin = 100
            RightMargin = 200

            def __init__(self) -> None:
                self.HSet = self

        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.is_in_table_cell = lambda: True
        app.read_current_table_cell_size = lambda _prefix: (5000, 1000, "fake-cell")
        app.current_cell_horizontal_margins = lambda: (600, "fake-margin")

        width, source = app.current_paragraph_tab_width(FakeParaShape())

        self.assertEqual(width, 8800)
        self.assertIn("셀", source)

    def test_paragraph_dotted_right_tab_floors_to_tenth_point_precision(self) -> None:
        app = object.__new__(MvpApp)

        self.assertEqual(app.paragraph_tab_effective_width(26056), 26050)
        self.assertEqual(app.paragraph_tab_stop_value(26056), 52100)
        self.assertEqual(app.paragraph_tab_stop_value(48190), 96380)

    def test_paragraph_dotted_right_tab_falls_back_to_page_width_when_cell_width_fails(self) -> None:
        class FakeParaShape:
            HSet = None
            LeftMargin = 100
            RightMargin = 200

            def __init__(self) -> None:
                self.HSet = self

        app = object.__new__(MvpApp)
        app.debug = lambda _message: None
        app.is_in_table_cell = lambda: False
        app.read_current_table_cell_size = lambda _prefix: (_ for _ in ()).throw(AssertionError("page paragraph should not read cell width"))
        app.current_page_text_width = lambda: 5000

        width, source = app.current_paragraph_tab_width(FakeParaShape())

        self.assertEqual(width, 10000)
        self.assertIn("쪽", source)


if __name__ == "__main__":
    unittest.main()
