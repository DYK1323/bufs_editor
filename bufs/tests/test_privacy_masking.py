import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import hwp_style_mvp as app_module
from privacy_masking import mask_personal_text


class MaskingRuleTests(unittest.TestCase):
    def test_names_and_permitted_general_words(self):
        examples = {
            '홍길동': '홍*동', '기획팀': '기*팀', '재학생': '재*생',
            '김수': '김*', '김': '*', 'Alice': 'A***e', 'Li': 'L*',
            'John Smith': 'J*** ****h', 'Anne-Marie': 'A***-****e',
            '  홍길동  ': '  홍*동  ', '홍길동\n김수': '홍*동\n김*',
            'Jose\u0301 Luis': 'J*** ***s',
        }
        for source, expected in examples.items():
            with self.subTest(source=source):
                result = mask_personal_text(source)
                self.assertEqual(result.text, expected)
                self.assertEqual(mask_personal_text(result.text).text, expected)

    def test_structured_values_and_boundaries(self):
        examples = {
            '00123456': '****3456', '20260907': '****0907',
            '1234567': '1234567', '123456789': '123456789',
            '010-1234-5678': '010-****-5678', '01012345678': '010****5678',
            '02-123-4567': '02-***-4567', '051 123 4567': '051 *** 4567',
            '1234-5678-9012-3456': '1234-****-****-3456',
            '1234567890123456': '1234********3456',
            '1234 5678 9012 3456': '1234 **** **** 3456',
            '900101-1234567': '******-*******',
            '900101-5234567': '******-*******',
            'hong@example.com': 'h***@example.com', 'a@example.com': '*@example.com',
            '1234567890123456789': '1234567890123456789',
            '12.5%': '12.5%', '1,234,567': '1,234,567',
        }
        for source, expected in examples.items():
            with self.subTest(source=source):
                result = mask_personal_text(source)
                self.assertEqual(result.text, expected)
                self.assertEqual(mask_personal_text(expected).text, expected)

    def test_mixed_names_and_email_do_not_remask_domain(self):
        source = '홍길동: hong@example.com / 010-1234-5678'
        result = mask_personal_text(source)
        self.assertEqual(result.text, '홍*동: h***@example.com / 010-****-5678')
        self.assertFalse(mask_personal_text(result.text).edits)

    def test_edits_preserve_unchanged_runs(self):
        source = 'John Smith'
        result = mask_personal_text(source)
        self.assertEqual([(e.start, e.end) for e in result.edits], [(1, 4), (5, 9)])
        for edit in reversed(result.edits):
            source = source[:edit.start] + edit.replacement + source[edit.end:]
        self.assertEqual(source, result.text)

    def test_unrecognized_long_number_is_reported(self):
        result = mask_personal_text('1234567890123456789')
        self.assertEqual(result.skipped, 1)
        self.assertFalse(result.edits)

    def test_adjacent_label_and_number(self):
        masked = mask_personal_text('학번20261234').text
        self.assertEqual(masked, '학*****1234')
        self.assertEqual(mask_personal_text(masked).text, masked)


class MaskingCellTests(unittest.TestCase):
    def setUp(self):
        self.app = object.__new__(app_module.MvpApp)
        self.values = {(1, 0): '홍길동', (2, 0): '010-1234-5678', (3, 0): '선택밖'}
        self.pos = (1, 0, 0)
        self.selected = None
        self.commands = []
        self.logs = []
        self.app.ensure_hwp = lambda: True
        self.app.is_selected_cell_block = lambda: True
        self.app.get_current_hwp_path = lambda: 'test.hwpx'
        self.app.get_hwp_pos_by_set = lambda: self.pos
        self.app.set_hwp_pos = self.set_pos
        self.app.set_hwp_pos_by_set = self.set_pos
        self.app.get_current_cell_address = lambda: (1, self.pos[0])
        self.app.get_selected_cell_range_by_formula = lambda: {'addresses': [(1, 1), (1, 2)]}
        self.app.clear_hwp_selection = lambda: True
        self.app.move_between_table_addresses = lambda current, target: self.set_pos((target[1], 0, 0))
        self.app.scan_current_cell_paragraph_positions = lambda: [(self.pos[0], 0)]
        self.app.read_current_paragraph_text = lambda list_id, para: self.values[(list_id, para)]
        self.app.actual_hwp_text_range = lambda _list, _para, start, end: (start, end)
        self.app.select_hwp_text_range = self.select
        self.app.selected_hwp_text_range_matches = lambda *args: self.selected == args
        self.app.run_hwp_command = self.command
        self.app.insert_hwp_text = self.insert
        self.app.select_current_table_cell_for_replace = lambda: True
        self.app.get_clipboard_text = lambda: self.values[(self.pos[0], 0)]
        self.app.set_clipboard_text = self.set_clipboard_text
        self.app.paste_text_into_selected_cell = self.paste_clipboard_into_cell
        self.clipboard = None
        self.app.log = self.logs.append
        self.warning = patch.object(app_module.messagebox, 'showwarning').start()
        self.addCleanup(patch.stopall)

    def set_pos(self, pos):
        self.pos = pos
        return True

    def select(self, list_id, para, start, end):
        self.selected = (list_id, para, start, end)
        self.pos = (list_id, para, start)
        return True

    def command(self, command):
        self.commands.append(command)
        if command == 'Copy':
            return True
        self.assertEqual(command, 'Delete')
        list_id, para, start, end = self.selected
        source = self.values[(list_id, para)]
        self.values[(list_id, para)] = source[:start] + source[end:]
        return True

    def insert(self, text):
        self.commands.append('InsertText')
        list_id, para, start = self.pos
        source = self.values[(list_id, para)]
        self.values[(list_id, para)] = source[:start] + text + source[start:]
        return True

    def set_clipboard_text(self, text):
        self.clipboard = text

    def paste_clipboard_into_cell(self):
        self.commands.append('ClipboardPaste')
        self.values[(self.pos[0], 0)] = self.clipboard
        return True

    def test_multiple_cells_and_repeat_preserve_outside(self):
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values, {(1, 0): '홍*동', (2, 0): '010-****-5678', (3, 0): '선택밖'})
        self.assertFalse(self.warning.called)
        calls = list(self.commands)
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.commands, calls)
        self.assertNotIn('홍길동', ''.join(self.logs))
        self.assertNotIn('010-1234-5678', ''.join(self.logs))

    def test_single_cell(self):
        self.app.get_selected_cell_range_by_formula = lambda: {'addresses': [(1, 1)]}
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')

    def test_missing_range_never_guesses(self):
        self.app.get_selected_cell_range_by_formula = lambda: None
        self.app.get_selected_text_positions = lambda: ((1, 0, 0), (2, 0, 5))
        self.app.mask_selected_cell_personal_data()
        self.assertFalse(self.commands)
        self.assertTrue(self.warning.called)

    def test_selection_mismatch_falls_back_to_clipboard(self):
        # GetSelectedPos가 무효를 반환하는 하이퍼링크 필드 텍스트 등에서는 문자 단위 편집이
        # 구조적으로 검증 불가능하다. 이 경우 셀 전체를 클립보드로 치환해 그래도 마스킹을 끝낸다.
        self.app.selected_hwp_text_range_matches = lambda *args: False
        self.app.mask_selected_cell_personal_data()
        self.assertFalse(self.warning.called)
        self.assertEqual(self.values, {(1, 0): '홍*동', (2, 0): '010-****-5678', (3, 0): '선택밖'})
        self.assertIn('ClipboardPaste', self.commands)

    def test_delete_lies_about_success_falls_back_to_clipboard(self):
        # HWP의 Delete가 True를 반환했는데 실제로는 아무것도 안 지운 경우(하이퍼링크 필드 등)도
        # 같은 클립보드 폴백으로 복구되어야 한다.
        self.app.run_hwp_command = lambda command: self.commands.append(command) or True
        self.app.mask_selected_cell_personal_data()
        self.assertFalse(self.warning.called)
        self.assertEqual(self.values, {(1, 0): '홍*동', (2, 0): '010-****-5678', (3, 0): '선택밖'})

    def test_insert_failure_after_real_delete_stops_without_guessing(self):
        # Delete는 실제로 문자를 지웠는데 Insert가 실패하면, 문단이 이미 source와 달라져 있다.
        # 이 상태에서 클립보드로 다시 마스킹하면 잘못된 결과를 만들 수 있으므로 자동 복구하지
        # 않고 안전하게 중단해야 한다(클립보드 폴백은 문단이 손대지지 않았을 때만 시도한다).
        self.app.insert_hwp_text = lambda text: False
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')
        self.assertTrue(self.warning.called)
        self.assertFalse(self.app.privacy_masking_busy)

    def test_fallback_failure_still_aborts_safely(self):
        # 문자 단위 편집도, 클립보드 폴백도 모두 실패하면(아직 아무것도 지우기 전이라면) 문서를
        # 손대지 않고 안전하게 중단해야 한다.
        self.app.select_hwp_text_range = lambda *args: False
        self.app.select_current_table_cell_for_replace = lambda: False
        self.app.mask_selected_cell_personal_data()
        self.assertTrue(self.warning.called)
        self.assertEqual(self.values[(1, 0)], '홍길동')
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')

    def test_reentrant_call_does_nothing(self):
        self.app.privacy_masking_busy = True
        self.app.mask_selected_cell_personal_data()
        self.assertFalse(self.commands)

    def test_raw_exception_is_not_logged(self):
        def fail(*args):
            raise RuntimeError('홍길동 secret@example.com')
        self.app.read_current_paragraph_text = fail
        self.app.mask_selected_cell_personal_data()
        self.assertNotIn('secret@example.com', ''.join(self.logs))
        self.assertNotIn('홍길동', str(self.warning.call_args))

    def test_document_change_prevents_edit_and_cursor_restore(self):
        paths = iter(['original.hwpx', 'other.hwpx', 'other.hwpx'])
        self.app.get_current_hwp_path = lambda: next(paths)
        self.app.mask_selected_cell_personal_data()
        self.assertFalse(self.commands)
        self.assertTrue(self.warning.called)
        self.assertFalse(self.app.privacy_masking_busy)

    def test_read_moves_cursor_but_insertion_position_is_restored(self):
        def read(list_id, para):
            self.pos = (3, 0, 0)
            return self.values[(list_id, para)]
        self.app.read_current_paragraph_text = read
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(3, 0)], '선택밖')
        self.assertFalse(self.warning.called)

    def test_multiple_paragraphs_and_duplicate_cell_are_not_repeated(self):
        self.values[(1, 1)] = 'Alice'
        self.app.get_selected_cell_range_by_formula = lambda: {'addresses': [(1, 1), (1, 1)]}
        self.app.scan_current_cell_paragraph_positions = lambda: [(1, 0), (1, 1)]
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(1, 1)], 'A***e')
        self.assertEqual(self.commands.count('Delete'), 2)

    def test_moves_between_cells_without_pre_snapshotting_all_positions(self):
        moves = []
        real_move = self.app.move_between_table_addresses

        def tracking_move(current, target):
            moves.append((current, target))
            return real_move(current, target)

        self.app.move_between_table_addresses = tracking_move
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(moves, [((1, 1), (1, 2))])
        self.assertEqual(self.values, {(1, 0): '홍*동', (2, 0): '010-****-5678', (3, 0): '선택밖'})
        self.assertFalse(self.warning.called)

    def test_cell_move_failure_aborts_without_editing(self):
        self.app.move_between_table_addresses = lambda current, target: False
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.commands.count('Delete'), 1)
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')
        self.assertTrue(self.warning.called)
        self.assertFalse(self.app.privacy_masking_busy)

    def test_cell_move_verification_mismatch_aborts(self):
        self.app.move_between_table_addresses = lambda current, target: self.set_pos((99, 0, 0))
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.commands.count('Delete'), 1)
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')
        self.assertTrue(self.warning.called)

    def test_verified_single_cell_without_formula(self):
        self.app.get_selected_cell_range_by_formula = lambda: None
        self.app.get_selected_text_positions = lambda: ((1, 0, 0), (1, 0, 3))
        self.app.current_field_is_table_cell = lambda: True
        self.app.mask_selected_cell_personal_data()
        self.assertEqual(self.values[(1, 0)], '홍*동')
        self.assertEqual(self.values[(2, 0)], '010-1234-5678')

    def test_busy_debug_drops_private_content(self):
        self.app.privacy_masking_busy = True
        self.app.debug('홍길동 secret@example.com')
        self.assertFalse(self.logs)


if __name__ == '__main__':
    unittest.main()
