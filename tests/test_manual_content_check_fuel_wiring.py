from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "Module1").read_text(encoding="utf-8")
FORM = (ROOT / "UserForm1").read_text(encoding="utf-8")


def procedure(source, name, kind="Sub"):
    match = re.search(
        rf"(?ims)^\s*(?:(?:Public|Private)\s+)?{kind}\s+{re.escape(name)}\b.*?^\s*End\s+{kind}\b",
        source,
    )
    if match is None:
        raise AssertionError(f"{kind} {name} was not found")
    return match.group(0)


class ManualContentCheckFuelWiringTests(unittest.TestCase):
    def test_manual_content_check_runs_in_required_order(self):
        body = procedure(MODULE, "入力内容チェック")
        calls = [
            "Call 入力内容チェック実行(True)",
            "Call 給油記録を照合実行(False)",
            "Call 車両候補確認画面を表示",
        ]
        self.assertEqual(body.count("Call 給油記録を照合実行(False)"), 1)
        positions = [body.index(call) for call in calls]
        self.assertEqual(positions, sorted(positions))

    def test_internal_content_check_does_not_run_fuel_reconciliation(self):
        body = procedure(MODULE, "入力内容チェック実行", "Function")
        self.assertNotIn("給油記録を照合", body)
        self.assertEqual(MODULE.count("Call 入力内容チェック実行(False)"), 2)
        self.assertNotIn("給油記録を照合実行", procedure(MODULE, "CSV取込"))
        self.assertNotIn("給油記録を照合実行", procedure(MODULE, "記録シート整理を実行"))

    def test_standalone_fuel_macro_remains_available(self):
        body = procedure(MODULE, "給油記録を照合")
        self.assertIn("Call 給油記録を照合実行(True)", body)

    def test_success_message_is_suppressed_only_for_internal_silent_call(self):
        body = procedure(MODULE, "給油記録を照合実行", "Function")
        self.assertIn("If 結果表示 Then MsgBox \"給油記録の照合が完了しました。\"", body)
        self.assertIn("MsgBox \"給油照合を中止しました。\"", body)
        self.assertIn("MsgBox \"記録シートのガソリン給油列が見つかりません。\"", body)

    def test_form_button_still_calls_only_manual_content_check(self):
        body = procedure(FORM, "CommandButton5_Click")
        self.assertIn("Call 入力内容チェック", body)
        self.assertNotIn("給油記録を照合", body)


if __name__ == "__main__":
    unittest.main()
