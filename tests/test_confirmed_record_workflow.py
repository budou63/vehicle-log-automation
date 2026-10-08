from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "Module1").read_text(encoding="utf-8")
SHEET_CODE_PATH = ROOT / "記録シートコード.txt"
CHECK_MARK = chr(0x2713)


def procedure(source, name, kind="Sub"):
    match = re.search(
        rf"(?ims)^\s*(?:(?:Public|Private)\s+)?{kind}\s+{re.escape(name)}\b.*?^\s*End\s+{kind}\b",
        source,
    )
    if not match:
        raise AssertionError(f"{kind} {name} was not found")
    return match.group(0)


def unresolved_auto_warning(cell):
    return bool(cell.get("managed_warning")) and cell.get("font_color") == "red"


def can_confirm(record):
    if not str(record.get("answer", "")).strip():
        return False, "回答番号がありません。"
    unresolved = [name for name, cell in record.get("cells", {}).items() if unresolved_auto_warning(cell)]
    if unresolved:
        return False, "未確認の警告: " + "、".join(unresolved)
    return True, ""


def click_check(record):
    if record.get("check") == CHECK_MARK:
        record["check"] = ""
        return True, ""
    if record.get("check", "") != "":
        return False, "チェック欄の値を確認してください。"
    allowed, message = can_confirm(record)
    if allowed:
        record["check"] = CHECK_MARK
    return allowed, message


def atomic_edit(records, changes):
    """Apply a whole user paste or undo the whole operation if it touches a confirmed row."""
    if any(records[index].get("check") == CHECK_MARK for index, _field, _value in changes):
        return False
    for index, field, value in changes:
        records[index][field] = value
    return True


class ConfirmedRowModelTests(unittest.TestCase):
    def test_clean_row_is_confirmable(self):
        record = {"answer": "601", "check": "", "cells": {}}
        self.assertEqual(click_check(record), (True, ""))
        self.assertEqual(record["check"], CHECK_MARK)

    def test_red_managed_warning_blocks_confirmation_and_names_item(self):
        record = {"answer": "601", "check": "", "cells": {"運転者氏名": {"managed_warning": True, "font_color": "red"}}}
        allowed, message = click_check(record)
        self.assertFalse(allowed)
        self.assertIn("運転者氏名", message)
        self.assertEqual(record["check"], "")

    def test_black_font_resolves_warning_even_when_managed_comment_remains(self):
        record = {"answer": "601", "check": "", "cells": {"運転者氏名": {"managed_warning": True, "font_color": "automatic"}}}
        self.assertEqual(click_check(record), (True, ""))

    def test_ordinary_user_comment_alone_does_not_block(self):
        record = {"answer": "601", "check": "", "cells": {"行先": {"user_comment": "手書きメモ", "font_color": "red"}}}
        self.assertEqual(click_check(record), (True, ""))

    def test_missing_answer_is_not_confirmed(self):
        record = {"answer": "", "check": "", "cells": {}}
        self.assertFalse(click_check(record)[0])

    def test_confirmed_click_turns_off(self):
        record = {"answer": "601", "check": CHECK_MARK, "cells": {"項目": {"managed_warning": True, "font_color": "red"}}}
        self.assertEqual(click_check(record), (True, ""))
        self.assertEqual(record["check"], "")

    def test_multirow_paste_is_atomic_when_one_row_is_confirmed(self):
        records = [{"answer": "601", "check": CHECK_MARK, "vehicle": "1889"}, {"answer": "602", "check": "", "vehicle": "7391"}]
        before = [record.copy() for record in records]
        self.assertFalse(atomic_edit(records, [(0, "vehicle", "7391"), (1, "vehicle", "1889")]))
        self.assertEqual(records, before)

    def test_edit_after_uncheck_is_allowed(self):
        records = [{"answer": "601", "check": "", "vehicle": "1889"}]
        self.assertTrue(atomic_edit(records, [(0, "vehicle", "7391")]))
        self.assertEqual(records[0]["vehicle"], "7391")


class ConfirmedRowWiringTests(unittest.TestCase):
    def test_check_mark_is_generated_with_chr_w_not_embedded(self):
        self.assertIn("ChrW(&H2713)", MODULE)
        self.assertNotIn(CHECK_MARK, MODULE)

    def test_shared_header_and_last_row_helpers_exist(self):
        self.assertIn("Public Function 記録列番号", MODULE)
        self.assertIn("Public Function 記録最終行", MODULE)
        body = procedure(MODULE, "記録最終行", "Function")
        self.assertNotIn('Cells(ws.Rows.Count, "A")', body)
        self.assertIn('HeaderColumn(ws, "回答番号")', body)

    def test_csv_maps_answer_by_header_and_new_row_check_is_blank(self):
        body = procedure(MODULE, "CSV取込")
        self.assertIn("CSV行を記録へ追加", body)
        helper = procedure(MODULE, "CSV行を記録へ追加", "Function")
        self.assertIn("回答番号", helper)
        self.assertIn("チェック", helper)
        self.assertIn("ClearContents", helper)
        self.assertIn("HasFormula", helper)
        self.assertIn("targetColumn = checkColumn", helper)
        self.assertNotIn("Cells(targetRow, checkColumn).ClearContents", helper)
        self.assertIn("targetRow", helper)

    def test_csv_history_uses_answer_number_header_not_check_column(self):
        for name in ("取込履歴を準備", "回答番号行を検索"):
            self.assertIn(name, MODULE)
        self.assertIn("記録列番号(recordWS, \"回答番号\")", procedure(MODULE, "取込履歴を準備", "Function"))

    def test_manual_flow_autoconfirms_only_after_all_checks(self):
        body = procedure(MODULE, "入力内容チェック")
        calls = ["入力内容チェック実行", "運転者氏名をチェック", "給油記録を照合実行", "車両候補確認画面を表示", "未確認行を自動確定"]
        positions = [body.index(token) for token in calls]
        self.assertEqual(positions, sorted(positions))

    def test_internal_false_check_never_auto_confirms(self):
        body = procedure(MODULE, "入力内容チェック実行", "Function")
        self.assertNotIn("未確認行を自動確定", body)
        for name in ("CSV取込", "記録シート整理を実行"):
            self.assertNotIn("未確認行を自動確定", procedure(MODULE, name, "Function" if name == "記録シート整理を実行" else "Sub"))

    def test_auto_confirm_requires_valid_answer_and_no_red_managed_warnings(self):
        body = procedure(MODULE, "未確認行を自動確定", "Sub")
        self.assertIn("回答番号", body)
        self.assertIn("記録行は確認可能", body)
        self.assertIn("確認チェックマーク", body)
        can_confirm = procedure(MODULE, "記録行は確認可能", "Function")
        self.assertIn("未解決自動警告", can_confirm)

    def test_warning_detector_uses_managed_markers_and_red_color(self):
        body = procedure(MODULE, "未解決自動警告", "Function")
        self.assertIn("CHECK_WARNING_START", body)
        self.assertIn("CHECK_WARNING_END", body)
        self.assertIn("vbRed", body)
        self.assertIn("Font.Color", body)

    def test_sheet_click_event_toggles_single_check_cell_and_moves_selection(self):
        source = SHEET_CODE_PATH.read_text(encoding="utf-8")
        self.assertIn("Worksheet_SelectionChange", source)
        self.assertIn("Target.CountLarge <> 1", source)
        self.assertIn("確認チェックマーク", source)
        self.assertIn("Me.Cells(Target.Row, answerColumn).Select", source)

    def test_sheet_events_restore_the_captured_application_event_state_on_errors(self):
        source = SHEET_CODE_PATH.read_text(encoding="utf-8")
        for name in ("Worksheet_SelectionChange", "Worksheet_Change"):
            body = procedure(source, name, "Sub")
            self.assertIn("eventsCaptured", body)
            self.assertIn("If eventsCaptured Then Application.EnableEvents = previousEvents", body)
            self.assertNotIn("Application.EnableEvents = True", body)
            self.assertLess(body.index("previousEvents = Application.EnableEvents"), body.index("eventsCaptured = True"))

    def test_sheet_change_event_rejects_confirmed_rows_atomically(self):
        source = SHEET_CODE_PATH.read_text(encoding="utf-8")
        self.assertIn("Worksheet_Change", source)
        self.assertIn("Application.Undo", source)
        self.assertIn("記録行は確認済み", source)
        self.assertIn("Application.EnableEvents", source)

    def test_checked_rows_are_skipped_by_automated_content_and_name_checks(self):
        standard = procedure(MODULE, "入力内容チェック実行", "Function")
        drivers = procedure(MODULE, "運転者氏名をチェック", "Function")
        self.assertIn("記録行は確認済み", standard)
        self.assertIn("記録行は確認済み", drivers)

    def test_checked_rows_are_skipped_by_reformat_and_meter_fill(self):
        organizer = procedure(MODULE, "記録シート整理を実行", "Function")
        self.assertIn("記録行は確認済み", organizer)
        self.assertIn("終業時メーター(km)", organizer)
        self.assertIn("始業時メーター(km)", organizer)

    def test_sort_includes_check_column_and_full_used_table(self):
        organizer = procedure(MODULE, "記録シート整理を実行", "Function")
        self.assertIn("Cells(2, 1)", organizer)
        self.assertIn("記録最終列", organizer)
        self.assertNotIn('Range("A2:AO"', organizer)

    def test_vehicle_candidate_and_approval_use_headers_and_guard_confirmed_rows(self):
        detector = procedure(MODULE, "車両選択候補をチェック")
        approval = procedure(MODULE, "車両候補を承認して修正", "Function")
        self.assertIn("記録列番号(ws, \"回答番号\")", detector)
        self.assertIn("記録列番号(ws, \"車両番号\")", detector)
        self.assertIn("記録行は確認済み", detector)
        self.assertIn("記録行は確認済み", approval)

    def test_fuel_matching_reads_dynamic_headers_and_does_not_warn_checked_rows(self):
        log_reader = procedure(MODULE, "ReadLogs", "Function")
        warning_writer = procedure(MODULE, "給油警告を記録シートへ反映", "Function")
        self.assertIn("記録列番号(ws, \"回答番号\")", log_reader)
        self.assertIn("記録列番号(ws, \"車両番号\")", log_reader)
        self.assertIn("記録列番号(ws, \"運転した日\")", log_reader)
        self.assertIn("記録行は確認済み", warning_writer)

    def test_print_transfer_reads_record_fields_by_header(self):
        body = procedure(MODULE, "転記to印刷様式")
        self.assertIn("記録列番号(wsSrc, \"回答番号\")", body)
        self.assertIn("記録列番号(wsSrc, \"車両番号\")", body)
        self.assertNotRegex(body, r'wsSrc\.Cells\(i,\s*"[A-Z]+"\)')

    def test_auto_warning_refresh_ignores_confirmed_rows(self):
        cleanup = procedure(MODULE, "前回の自動チェック警告を除去")
        self.assertIn("記録行は確認済み", cleanup)

    def test_check_state_is_defined_only_by_exact_generated_mark(self):
        helper = procedure(MODULE, "記録行は確認済み", "Function")
        self.assertIn("確認チェックマーク", helper)
        self.assertIn('記録列番号(ws, "チェック")', helper)

    def test_csv_append_guard_preserves_existing_data_and_check_cell(self):
        body = procedure(MODULE, "CSV行を記録へ追加", "Function")
        self.assertIn("existingLastRow = 記録最終行(recordWS)", body)
        self.assertIn("If targetRow <= existingLastRow Then Exit Function", body)
        self.assertIn("recordWS.Cells(targetRow, checkColumn).HasFormula", body)
        self.assertNotIn("Cells(targetRow, checkColumn).ClearContents", body)
        self.assertIn("If targetColumn = 0 Or targetColumn = checkColumn Then Exit Function", body)

    def test_blank_answer_row_is_ignored_by_manual_click_event(self):
        source = SHEET_CODE_PATH.read_text(encoding="utf-8")
        self.assertIn("記録行に回答番号あり(Me, Target.Row)", source)
        self.assertIn("If Not 記録行に回答番号あり", source)

    def test_header_lookup_normalizes_whitespace_and_rejects_duplicate_names(self):
        body = procedure(MODULE, "HeaderColumn", "Function")
        self.assertIn("HeaderNameKey(CStr(headerValue))", body)
        self.assertIn("If foundColumn > 0 Then", body)
        self.assertIn("HeaderColumn = 0", body)

    def test_manual_check_requires_unique_answer_and_check_headers(self):
        body = procedure(MODULE, "入力内容チェック", "Sub")
        self.assertIn('記録列番号(ws, "回答番号") = 0', body)
        self.assertIn('記録列番号(ws, "チェック") = 0', body)
        self.assertLess(body.index('記録列番号(ws, "チェック") = 0'), body.index("入力内容チェック実行(False)"))

    def test_gasoline_column_uses_normalized_shared_header_lookup(self):
        body = procedure(MODULE, "ガソリン給油列を取得", "Function")
        self.assertIn("HeaderColumn(ws, GASOLINE_FUEL_HEADER)", body)

    def test_missing_check_header_fails_closed_for_macro_writes(self):
        body = procedure(MODULE, "記録行は確認済み", "Function")
        self.assertIn("If checkColumn = 0 Then", body)
        self.assertIn("記録行は確認済み = True", body)
        self.assertIn("Exit Function", body)

    def test_checked_rows_are_not_rewarned_by_fuel_reconciliation(self):
        body = procedure(MODULE, "給油警告を記録シートへ反映", "Function")
        self.assertIn("If Not 記録行は確認済み(logWS, rowNo)", body)
        self.assertIn("If Not 記録行は確認済み(logWS, CLng(ticketItem))", body)
        self.assertIn("If Not 記録行は確認済み(logWS, CLng(candidateLog(4)))", body)

    def test_cp932_roundtrip_for_all_deployable_vba_sources(self):
        paths = [ROOT / "Module1", SHEET_CODE_PATH]
        for path in paths:
            text = path.read_text(encoding="utf-8")
            encoded = text.encode("cp932", errors="strict")
            self.assertEqual(encoded.decode("cp932"), text)
            self.assertNotIn(CHECK_MARK, text)


if __name__ == "__main__":
    unittest.main()
