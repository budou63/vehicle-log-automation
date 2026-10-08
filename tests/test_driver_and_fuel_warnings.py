from datetime import date
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "Module1"
MODULE_SOURCE = MODULE.read_text(encoding="utf-8")


DRIVER_SPACE = " \u3000"
BAD_NAME_CHARS = set("0123456789０１２３４５６７８９?？!！/／#＃@＠")


def trim_edges(value):
    return value.strip(DRIVER_SPACE)


def remove_spaces(value):
    return value.replace(" ", "").replace("\u3000", "")


def levenshtein_one(left, right):
    if left == right:
        return False
    if abs(len(left) - len(right)) > 1:
        return False
    previous = list(range(len(right) + 1))
    for i, a in enumerate(left, 1):
        current = [i]
        for j, b in enumerate(right, 1):
            current.append(min(current[-1] + 1, previous[j] + 1, previous[j - 1] + (a != b)))
        previous = current
    return previous[-1] == 1


def classify_driver(value, master_names=None, exact_counts=None, normalized_counts=None, variants=None, representatives=None):
    raw = str(value)
    trimmed = trim_edges(raw)
    compact = remove_spaces(trimmed)
    issues = []
    if raw != trimmed:
        issues.append("leading_trailing_space")
    if (" " in raw and "　" in raw) or re.search(r"[ \u3000]{2,}", raw):
        issues.append("space_format")
    if any(c in BAD_NAME_CHARS for c in raw):
        issues.append("bad_character")
    if compact and len(compact) == 1:
        issues.append("too_short")
    if not compact:
        return issues

    exact_counts = exact_counts or {raw: 1}
    normalized_counts = normalized_counts or {compact: 1}
    variants = variants or {compact: {raw: exact_counts.get(raw, 1)}}
    representatives = representatives or {compact: raw}
    trimmed_master = [trim_edges(name) for name in (master_names or []) if remove_spaces(trim_edges(name))]
    if trimmed in trimmed_master:
        return issues
    current_count = exact_counts.get(raw, 1)
    if current_count == 1:
        frequent_spacing = [name for name, count in variants.get(compact, {}).items() if name != raw and count >= 2]
        if frequent_spacing:
            issues.append("history_spacing")
        near = sorted(
            representatives[other]
            for other, count in normalized_counts.items()
            if other != compact and count >= 2 and normalized_counts.get(compact, current_count) < count and levenshtein_one(compact, other)
        )
        if len(near) == 1:
            issues.append("history_near:" + near[0])
        elif len(near) > 1:
            issues.append("history_near_multiple")

    if master_names:
        trimmed_master = [trim_edges(name) for name in master_names if remove_spaces(trim_edges(name))]
        exact = [name for name in trimmed_master if trimmed == name]
        spacing = [name for name in trimmed_master if remove_spaces(name) == compact]
        if not exact:
            if spacing:
                issues.append("master_spacing_difference")
            else:
                near_master = sorted({name for name in trimmed_master if levenshtein_one(compact, remove_spaces(name))})
                if len(near_master) == 1:
                    issues.append("master_near:" + near_master[0])
                elif len(near_master) > 1:
                    issues.append("master_near_multiple")
                else:
                    issues.append("new_name")
    return issues


def classify_driver_history(values, master_names=None):
    exact_counts = {}
    normalized_counts = {}
    variants = {}
    order = {}
    for value in values:
        raw = str(value)
        if not remove_spaces(trim_edges(raw)):
            continue
        compact = remove_spaces(trim_edges(raw))
        exact_counts[raw] = exact_counts.get(raw, 0) + 1
        normalized_counts[compact] = normalized_counts.get(compact, 0) + 1
        variants.setdefault(compact, {})[raw] = variants.setdefault(compact, {}).get(raw, 0) + 1
        order.setdefault(compact, [])
        if raw not in order[compact]:
            order[compact].append(raw)
    representatives = {}
    for compact, forms in variants.items():
        representatives[compact] = max(order[compact], key=lambda name: forms[name])
    return {
        str(value): classify_driver(value, master_names, exact_counts, normalized_counts, variants, representatives)
        for value in values if remove_spaces(trim_edges(str(value)))
    }


def same_vehicle_date_warnings(tickets, records):
    per_record = {}
    unlocated = 0
    for ticket in tickets:
        status = ticket["status"]
        if status == "入力値要確認" or status in ("同日一致", "後日一致"):
            continue
        if status == "未照合":
            candidates = [r for r in records if r["car"] == ticket["car"] and r["day"] == ticket["day"]]
            if not candidates:
                unlocated += 1
            else:
                for record in candidates:
                    per_record.setdefault(record["row"], []).append("unmatched")
        elif status == "候補複数・要確認":
            candidates = [r for r in records if r["car"] == ticket["car"] and r["amount"] == ticket["amount"] and r["day"] >= ticket["day"] and r.get("answer_unique", True)]
            for record in candidates:
                per_record.setdefault(record["row"], []).append("ambiguous")
    return per_record, unlocated


def merge_category(categories, category, warning):
    result = {key: list(values) for key, values in categories.items()}
    result.setdefault(category, [])
    if warning not in result[category]:
        result[category].append(warning)
    return result


def remove_category(categories, category):
    return {key: list(values) for key, values in categories.items() if key != category and values}


def initial_master_names(sheet_names):
    seen = set()
    result = []
    for name in sheet_names:
        cleaned = trim_edges(name)
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)
    return result


def procedure(source, name, kind="Sub"):
    match = re.search(rf"(?ims)^\s*(?:(?:Public|Private)\s+)?{kind}\s+{re.escape(name)}\b.*?^\s*End\s+{kind}\b", source)
    if not match:
        raise AssertionError(f"missing {kind} {name}")
    return match.group(0)


class FuelWarningLogicTests(unittest.TestCase):
    def test_unmatched_ticket_warns_only_same_vehicle_and_date(self):
        tickets = [{"status": "未照合", "car": "7391", "day": date(2026, 4, 8), "amount": 30}]
        records = [
            {"row": 2, "car": "7391", "day": date(2026, 4, 8)},
            {"row": 3, "car": "1889", "day": date(2026, 4, 8)},
            {"row": 4, "car": "7391", "day": date(2026, 4, 9)},
        ]
        flagged, missing = same_vehicle_date_warnings(tickets, records)
        self.assertEqual(set(flagged), {2})
        self.assertEqual(missing, 0)

    def test_multiple_same_day_records_all_warn_without_selecting_one(self):
        tickets = [{"status": "未照合", "car": "7391", "day": date(2026, 4, 8), "amount": 30}]
        records = [{"row": n, "car": "7391", "day": date(2026, 4, 8)} for n in (2, 3)]
        flagged, _ = same_vehicle_date_warnings(tickets, records)
        self.assertEqual(set(flagged), {2, 3})

    def test_no_same_day_record_does_not_choose_nearest_day_and_counts_ticket(self):
        tickets = [{"status": "未照合", "car": "7391", "day": date(2026, 4, 8), "amount": 30}]
        records = [{"row": 2, "car": "7391", "day": date(2026, 4, 7)}]
        flagged, missing = same_vehicle_date_warnings(tickets, records)
        self.assertEqual(flagged, {})
        self.assertEqual(missing, 1)

    def test_unlocated_header_count_is_exact(self):
        tickets = [{"status": "未照合", "car": str(i), "day": date(2026, 4, 8)} for i in range(3)]
        flagged, missing = same_vehicle_date_warnings(tickets, [])
        self.assertFalse(flagged)
        self.assertEqual(missing, 3)

    def test_ambiguous_status_warns_all_valid_candidates_and_never_matches(self):
        ticket = {"status": "候補複数・要確認", "car": "7391", "day": date(2026, 4, 8), "amount": 25}
        records = [
            {"row": 2, "car": "7391", "day": date(2026, 4, 8), "amount": 25, "answer_unique": True},
            {"row": 3, "car": "7391", "day": date(2026, 4, 9), "amount": 25, "answer_unique": True},
            {"row": 4, "car": "1889", "day": date(2026, 4, 8), "amount": 25, "answer_unique": True},
            {"row": 5, "car": "7391", "day": date(2026, 4, 8), "amount": 25, "answer_unique": False},
        ]
        flagged, _ = same_vehicle_date_warnings([ticket], records)
        self.assertEqual(set(flagged), {2, 3})
        self.assertEqual(flagged[2], ["ambiguous"])

    def test_resolved_and_invalid_ticket_statuses_do_not_warn_record_sheet(self):
        tickets = [{"status": status, "car": "7391", "day": date(2026, 4, 8), "amount": 25} for status in ("同日一致", "後日一致", "入力値要確認")]
        records = [{"row": 2, "car": "7391", "day": date(2026, 4, 8), "amount": 25}]
        self.assertEqual(same_vehicle_date_warnings(tickets, records), ({}, 0))

    def test_category_updates_are_idempotent_and_independent(self):
        categories = merge_category({}, "通常チェック", "invalid fuel quantity")
        categories = merge_category(categories, "給油照合警告", "fuel warning")
        categories = merge_category(categories, "運転者名警告", "driver warning")
        categories = merge_category(categories, "給油照合警告", "fuel warning")
        self.assertEqual(categories["給油照合警告"], ["fuel warning"])
        remaining = remove_category(categories, "給油照合警告")
        self.assertEqual(remaining, {"通常チェック": ["invalid fuel quantity"], "運転者名警告": ["driver warning"]})
        remaining = remove_category(remaining, "通常チェック")
        self.assertEqual(remaining, {"運転者名警告": ["driver warning"]})

    def test_unmatched_ticket_warning_copy_mentions_details_and_header_fallback(self):
        self.assertIn("給油記録に", MODULE_SOURCE)
        self.assertIn("対応する給油入力が確認できません", MODULE_SOURCE)
        self.assertIn("対応先を特定できない未照合チケットが", MODULE_SOURCE)
        self.assertIn('CStr(unlocatedCount) & "件あります。詳細は給油記録シートD:Gを確認してください。"', MODULE_SOURCE)

    def test_header_warning_only_changes_comment_not_header_value(self):
        body = procedure(MODULE_SOURCE, "給油警告を記録シートへ反映", "Function")
        self.assertIn("logWS.Cells(1, gasCol)", body)
        self.assertNotRegex(body, r'logWS\.Cells\(1,\s*gasCol\)\.Value\s*=')

    def test_fuel_warnings_are_category_cleared_and_apply_to_gas_column(self):
        self.assertIn('"給油照合警告"', MODULE_SOURCE)
        self.assertIn("給油警告を記録シートへ反映", MODULE_SOURCE)
        self.assertIn("セルの自動チェック警告カテゴリを除去", MODULE_SOURCE)
        self.assertIn("gasCol", procedure(MODULE_SOURCE, "給油警告を記録シートへ反映", "Function"))

    def test_standalone_and_content_check_both_refresh_fuel_warning(self):
        self.assertIn("給油警告を記録シートへ反映", procedure(MODULE_SOURCE, "給油記録を照合実行", "Function"))


class DriverNameLogicTests(unittest.TestCase):
    def test_history_without_master_does_not_warn_on_a_normal_first_name(self):
        issues = classify_driver_history(["田中次郎"])["田中次郎"]
        self.assertEqual(issues, [])

    def test_exact_master_name_is_not_overruled_by_rare_history(self):
        issues = classify_driver_history(["山田 太郎"] * 5 + ["山田太郎"], ["山田太郎"])
        self.assertEqual(issues["山田太郎"], [])

    def test_exact_master_match_precedes_history_frequency_warnings(self):
        checker = procedure(MODULE_SOURCE, "運転者名警告一覧", "Function")
        self.assertIn("exactMasterMatch", checker)
        self.assertLess(checker.index("If exactMasterMatch Then"), checker.index("If currentCount = 1"))

    def test_history_spacing_warns_only_singleton_variant(self):
        issues = classify_driver_history(["山田 太郎"] * 5 + ["山田太郎"])
        self.assertNotIn("history_spacing", issues["山田 太郎"])
        self.assertIn("history_spacing", issues["山田太郎"])

    def test_history_one_character_near_warns_only_singleton_variant(self):
        issues = classify_driver_history(["山田太郎"] * 5 + ["山田太朗"])
        self.assertFalse(any(item.startswith("history_near") for item in issues["山田太郎"]))
        self.assertIn("history_near:山田太郎", issues["山田太朗"])

    def test_history_near_comparison_uses_compact_group_frequency_too(self):
        values = ["山田太郎"] * 2 + ["山田太朗"] + ["山田 太朗"] * 2
        issues = classify_driver_history(values)
        self.assertFalse(any(item.startswith("history_near") for item in issues["山田太朗"]))

    def test_equal_frequency_near_names_do_not_warn(self):
        issues = classify_driver_history(["山田太郎", "山田太朗"])
        self.assertFalse(any(item.startswith("history_near") for row in issues.values() for item in row))

    def test_history_threshold_requires_singleton_target_and_two_common_entries(self):
        two_target = classify_driver_history(["山田太郎"] * 3 + ["山田太朗"] * 2)
        self.assertFalse(any(item.startswith("history_near") for row in two_target.values() for item in row))
        one_common = classify_driver_history(["山田太郎", "山田太朗"])
        self.assertFalse(any(item.startswith("history_near") for row in one_common.values() for item in row))

    def test_exact_match_has_no_warning(self):
        self.assertEqual(classify_driver("山田太郎", ["山田太郎"]), [])

    def test_new_name_is_flagged(self):
        self.assertIn("new_name", classify_driver("佐藤花子", ["山田太郎"]))

    def test_half_and_full_width_space_spelling_variants_are_flagged(self):
        self.assertIn("master_spacing_difference", classify_driver("山田太郎", ["山田 太郎"]))
        self.assertIn("master_spacing_difference", classify_driver("山田　太郎", ["山田太郎"]))

    def test_one_character_substitution_is_near_match(self):
        self.assertIn("master_near:山田太郎", classify_driver("山田太朗", ["山田太郎"]))
        self.assertIn("master_near:高橋一郎", classify_driver("髙橋一郎", ["高橋一郎"]))

    def test_one_character_insertion_or_deletion_is_near_match(self):
        self.assertTrue(levenshtein_one("山田三郎", "山田三郎一"))
        self.assertTrue(levenshtein_one("山田太郎", "山田太"))
        self.assertFalse(levenshtein_one("山田太郎", "鈴木花子"))

    def test_multiple_near_names_are_not_resolved_to_one(self):
        result = classify_driver("山田太郎", ["山田太朗", "山田大郎"])
        self.assertIn("master_near_multiple", result)
        self.assertFalse(any(item.startswith("master_near:") for item in result))

    def test_half_width_edge_space_is_flagged(self):
        self.assertIn("leading_trailing_space", classify_driver(" 山田太郎 ", ["山田太郎"]))

    def test_full_width_edge_space_is_flagged(self):
        self.assertIn("leading_trailing_space", classify_driver("　山田太郎　", ["山田太郎"]))

    def test_mixed_and_repeated_spaces_are_flagged(self):
        self.assertIn("space_format", classify_driver("山田 太郎　次郎", []))
        self.assertIn("space_format", classify_driver("山田  太郎", []))

    def test_digits_and_forbidden_symbols_are_flagged(self):
        for name in ("山田1", "山田１", "山田?", "山田？", "山田!", "山田！", "山田/", "山田／", "山田#", "山田＃", "山田@", "山田＠"):
            with self.subTest(name=name):
                self.assertIn("bad_character", classify_driver(name, []))

    def test_hyphen_like_name_characters_are_not_forbidden(self):
        for name in ("アンナ・リー", "ジョン-スミス", "ジョンースミス"):
            self.assertNotIn("bad_character", classify_driver(name, [name]))

    def test_single_character_name_is_flagged_but_two_character_name_is_not_short(self):
        self.assertIn("too_short", classify_driver("李", ["李"]))
        self.assertNotIn("too_short", classify_driver("李明", ["李明"]))

    def test_initial_master_list_trims_edges_keeps_internal_spaces_and_deduplicates(self):
        self.assertEqual(initial_master_names([" 山田 太郎　", "山田 太郎", "", "　", "李明"]), ["山田 太郎", "李明"])

    def test_name_checker_is_manual_only_and_does_not_write_w_or_add_names(self):
        public = procedure(MODULE_SOURCE, "入力内容チェック")
        self.assertLess(public.index("運転者氏名をチェック"), public.index("給油記録を照合実行(False)"))
        self.assertNotIn("運転者氏名をチェック", procedure(MODULE_SOURCE, "入力内容チェック実行", "Function"))
        for name in ("CSV取込", "記録シート整理を実行"):
            self.assertNotIn("運転者氏名をチェック", procedure(MODULE_SOURCE, name, "Function" if name == "記録シート整理を実行" else "Sub"))
        self.assertIn("Public Sub 運転者名マスタを初期作成()", MODULE_SOURCE)
        self.assertNotRegex(MODULE_SOURCE, r'(?im)^\s*ws\.Cells\([^\n]*,\s*"W"\)\.Value\s*=')

    def test_missing_master_has_no_popup_and_history_check_keeps_running(self):
        public = procedure(MODULE_SOURCE, "入力内容チェック")
        checker = procedure(MODULE_SOURCE, "運転者氏名をチェック", "Function")
        self.assertIn("運転者氏名をチェック", public)
        self.assertNotIn("運転者名マスタが未作成です", MODULE_SOURCE)
        self.assertNotIn("MasterMissing:", checker)
        self.assertIn("記録", checker)
        self.assertIn("運転者名警告一覧", checker)
        self.assertIn("exactCounts", checker)
        self.assertIn("normalizedCounts", checker)
        self.assertIn("運転者名の編集距離", procedure(MODULE_SOURCE, "運転者名警告一覧", "Function"))
        self.assertLess(public.index("運転者氏名をチェック"), public.index("給油記録を照合実行(False)"))
        self.assertLess(public.index("給油記録を照合実行(False)"), public.index("車両候補確認画面を表示"))

    def test_name_comment_category_is_refreshed_not_hand_comments(self):
        body = procedure(MODULE_SOURCE, "運転者氏名をチェック", "Function")
        self.assertIn('"運転者名警告"', body)
        self.assertIn("セルの自動チェック警告カテゴリを除去", body)

    def test_initializer_refuses_to_overwrite_existing_master_entries(self):
        body = procedure(MODULE_SOURCE, "運転者名マスタを初期作成")
        self.assertIn("A列に既存データ", body)
        self.assertIn("Exit Sub", body)
        self.assertIn('Worksheets("運転者名マスタ")', MODULE_SOURCE)

    def test_master_values_are_written_only_by_explicit_initializer(self):
        body = procedure(MODULE_SOURCE, "運転者名マスタを初期作成")
        self.assertIn('Cells(outputRow, "A").Value', body)
        self.assertNotIn('Cells(outputRow, "W").Value', body)
        self.assertNotIn('Cells(outputRow, "A").Value =', procedure(MODULE_SOURCE, "運転者氏名をチェック", "Function"))


class VBAWarningWiringTests(unittest.TestCase):
    def test_categories_share_one_managed_comment_and_preserve_other_categories(self):
        self.assertIn("【通常チェック警告開始】", MODULE_SOURCE)
        self.assertIn("【運転者名警告開始】", MODULE_SOURCE)
        self.assertIn("【給油照合警告開始】", MODULE_SOURCE)
        self.assertIn("Private Sub セルの自動チェック警告カテゴリを除去", MODULE_SOURCE)
        self.assertIn("originalColor", procedure(MODULE_SOURCE, "セルの自動チェック警告カテゴリを除去", "Sub"))

    def test_hand_comment_is_not_deleted_when_category_removed(self):
        remover = procedure(MODULE_SOURCE, "セルの自動チェック警告カテゴリを除去", "Sub")
        self.assertIn("note.Delete", remover)
        self.assertIn("If Len(Trim$(rebuiltText)) = 0 Then", remover)
        self.assertIn("prefixText", remover)
        self.assertIn("suffixText", remover)

    def test_malformed_or_duplicate_managed_markers_fail_closed(self):
        add = procedure(MODULE_SOURCE, "自動チェック警告を追加", "Sub")
        remove = procedure(MODULE_SOURCE, "セルの自動チェック警告カテゴリを除去", "Sub")
        self.assertIn("If markerEnd = 0 Then GoTo AddFailed", add)
        self.assertIn("Then GoTo AddFailed", add)
        self.assertIn("CHECK_WARNING_START, vbBinaryCompare) > 0 Then", add)
        self.assertIn("CHECK_WARNING_START, vbBinaryCompare) > 0 Then", remove)

    def test_standard_check_clears_only_standard_warning_category(self):
        self.assertIn('Call セルの自動チェック警告カテゴリを除去(target, errorCount, "通常チェック")', procedure(MODULE_SOURCE, "セルの自動チェック警告を除去", "Sub"))

    def test_comment_merge_deduplicates_identical_warning(self):
        self.assertIn("InStr(1, currentCategoryText, warningText", procedure(MODULE_SOURCE, "自動チェック警告を追加", "Sub"))

    def test_other_categories_keep_red_bold_and_restore_original_when_empty(self):
        source = procedure(MODULE_SOURCE, "セルの自動チェック警告カテゴリを除去", "Sub")
        self.assertIn("If hasManagedWarnings Then", source)
        self.assertIn("target.Font.Color = originalColor", source)
        self.assertIn("target.Font.Bold = originalBold", source)


if __name__ == "__main__":
    unittest.main()
