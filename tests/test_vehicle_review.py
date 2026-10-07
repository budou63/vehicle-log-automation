"""Static safety tests for the modeless vehicle-correction review feature."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "Module1"
FORM = ROOT / "frmVehicleReview.frm"


def approval_allowed(records, answer, expected_vehicle, expected_candidate):
    matches = [r for r in records if r["answer"] == answer]
    if len(matches) != 1:
        return False
    record = matches[0]
    return record["vehicle"] == expected_vehicle and record["candidate"] == expected_candidate and record["still_valid"]


class VehicleReviewTests(unittest.TestCase):
    def test_approval_requires_unique_answer_unchanged_vehicle_and_current_candidate(self):
        base = {"answer": "601", "vehicle": "1889", "candidate": "7391", "still_valid": True}
        self.assertTrue(approval_allowed([base], "601", "1889", "7391"))
        self.assertFalse(approval_allowed([], "601", "1889", "7391"))
        self.assertFalse(approval_allowed([base, base], "601", "1889", "7391"))
        self.assertFalse(approval_allowed([{**base, "vehicle": "5555"}], "601", "1889", "7391"))
        self.assertFalse(approval_allowed([{**base, "candidate": "5555"}], "601", "1889", "7391"))
        self.assertFalse(approval_allowed([{**base, "still_valid": False}], "601", "1889", "7391"))

    def test_module_only_opens_review_from_manual_content_check(self):
        source = MODULE.read_text(encoding="utf-8")
        public_check = re.search(r"Public Sub 入力内容チェック\(\).*?^End Sub", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(public_check)
        self.assertIn("車両候補確認画面を表示", public_check.group(0))
        self.assertEqual(source.count("車両候補確認画面を表示"), 2)
        self.assertIn("VBA.UserForms.Add(\"frmVehicleReview\")", source)
        self.assertIn("vbModeless", source)
        self.assertIn("車両候補を承認して修正", source)
        self.assertIn("Private Sub 車両候補確認画面を表示", source)
        approval = re.search(r"Public Function 車両候補を承認して修正.*?^End Function", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(approval)
        self.assertIn("最新の車両候補を再判定", approval.group(0))
        self.assertNotIn("入力内容チェック実行", approval.group(0))
        self.assertIn('Cells(foundRow, "J").Value', approval.group(0))
        self.assertIn("記録シート整理を実行", source)

    def test_modeless_form_has_required_buttons_and_safe_close(self):
        source = FORM.read_text(encoding="utf-8")
        self.assertIn("ShowModal", source)
        self.assertIn("=   0", source)
        for name in ("cmdGoToRow", "cmdApply", "cmdReject", "cmdDefer", "cmdFinish"):
            self.assertIn(name, source)
        self.assertIn("QueryClose", source)
        self.assertIn("車両候補確認を終了", source)
        self.assertIn("車両候補を承認して修正", source)


if __name__ == "__main__":
    unittest.main()
