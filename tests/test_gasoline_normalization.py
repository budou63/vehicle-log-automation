"""Static and decision-table tests for gasoline entry normalization in Module1."""
from __future__ import annotations

from pathlib import Path
import re
import unittest

MODULE = Path(__file__).resolve().parents[1] / "Module1"


def normalize_gasoline(value):
    """Model the required strict acceptance rule, not a permissive digit extractor."""
    if value is None or value == "":
        return "none", None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        amount = float(value)
    else:
        text = str(value).strip().translate(str.maketrans("０１２３４５６７８９Ｌ", "0123456789L"))
        if text == "なし":
            return "none", None
        for unit in ("リットル", "ℓ", "L", "l"):
            if text.endswith(unit):
                text = text[: -len(unit)].strip()
                break
        negative = text.startswith("-")
        digits = text[1:] if negative else text
        if not digits or digits.count(".") > 1:
            return "invalid", None
        integer, dot, fraction = digits.partition(".")
        if not integer or (dot and not fraction) or not integer.isascii() or not integer.isdigit():
            return "invalid", None
        if fraction and (not fraction.isascii() or not fraction.isdigit()):
            return "invalid", None
        amount = float(text)
    return ("valid", amount) if amount > 0 else ("nonpositive", amount)


class GasolineNormalizationTests(unittest.TestCase):
    def test_numeric_entries_are_accepted_as_positive_numbers(self):
        for value, expected in ((20, 20.0), ("25", 25.0), ("30", 30.0), ("25.5", 25.5)):
            with self.subTest(value=value):
                self.assertEqual(normalize_gasoline(value), ("valid", expected))

    def test_clear_liter_units_are_normalized(self):
        for value, expected in (
            ("25L", 25.0), ("30L", 30.0), ("25l", 25.0), ("25 L", 25.0),
            ("25Ｌ", 25.0), ("２５Ｌ", 25.0), ("25ℓ", 25.0), ("25リットル", 25.0),
        ):
            with self.subTest(value=value):
                self.assertEqual(normalize_gasoline(value), ("valid", expected))

    def test_none_and_blank_remain_non_numeric_without_a_warning(self):
        self.assertEqual(normalize_gasoline("なし"), ("none", None))
        self.assertEqual(normalize_gasoline(""), ("none", None))
        self.assertEqual(normalize_gasoline(None), ("none", None))

    def test_ambiguous_text_is_never_reduced_to_digits(self):
        for value in (
            "25?", "25？", "25くらい", "約25", "不明", "たぶん25", "25L?",
            "1e3", "1,000", "$25", "25.", ".5", "+25",
        ):
            with self.subTest(value=value):
                self.assertEqual(normalize_gasoline(value), ("invalid", None))

    def test_zero_and_negative_values_are_warned_as_nonpositive(self):
        self.assertEqual(normalize_gasoline(0), ("nonpositive", 0.0))
        self.assertEqual(normalize_gasoline("-5"), ("nonpositive", -5.0))

    def test_module_uses_a_header_resolved_column_and_managed_warning_path(self):
        source = MODULE.read_text(encoding="utf-8")
        check = re.search(r"Private Function 入力内容チェック実行.*?^End Function", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(check)
        body = check.group(0)
        for required in (
            "ガソリン給油列を取得", "ガソリン給油量をチェック", "gasolineColumn",
            "前回の自動チェック警告を除去", "自動チェック警告を追加",
        ):
            self.assertIn(required, body)
        self.assertIn("ガソリン給油の有無 [給油の場合はその他に数値入力（?）]", source)
        self.assertIn("Private Function 十進数を取得", source)
        self.assertNotIn("Not IsNumeric(numericText)", source)
        self.assertIn("ガソリン給油量を数値として認識できません", source)
        self.assertIn("ガソリン給油量が0以下です", source)


if __name__ == "__main__":
    unittest.main()
