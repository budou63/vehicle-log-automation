"""Regression tests for one-star fuel indicator placement in repeated forms."""
from __future__ import annotations

from pathlib import Path
import re
import unittest

MODULE = Path(__file__).resolve().parents[1] / "Module1"


def indicator_left(area_left, area_width, glyph_size, alignment, indent=0):
    """Expected first-glyph position using only the specific destination cell area."""
    if alignment == "center":
        return area_left + (area_width - glyph_size) / 2
    if alignment == "right":
        return area_left + area_width - glyph_size - indent * glyph_size
    return area_left + indent * glyph_size


def fuel_indicator(level):
    """The half-width shape is level 1 only; higher levels render full glyphs."""
    if level == 1:
        return "half-shape", 0
    if 2 <= level <= 5:
        return "text", level - 1
    return "empty", 0


class FuelIndicatorPositionTests(unittest.TestCase):
    def test_each_repeated_form_uses_its_own_gauge_area_not_a_previous_form(self):
        # Distinct report positions and merged widths model the repeated 22-row blocks.
        report_areas = ((82.0, 56.0), (614.5, 56.0), (1147.0, 73.0), (1679.5, 56.0))
        for area_left, area_width in report_areas:
            with self.subTest(area_left=area_left, area_width=area_width):
                left = indicator_left(area_left, area_width, 12.0, "center")
                expected_first_full_glyph = area_left + (area_width - 12.0) / 2
                self.assertAlmostEqual(left, expected_first_full_glyph)
                # Level 1 has no preceding full glyphs; the black shape is its left half.
                self.assertEqual(fuel_indicator(1), ("half-shape", 0))
                self.assertAlmostEqual(left + 12.0 / 2, expected_first_full_glyph + 6.0)

    def test_level_one_shape_is_separate_from_levels_two_through_five_text(self):
        self.assertEqual(fuel_indicator(1), ("half-shape", 0))
        for level in range(2, 6):
            with self.subTest(level=level):
                self.assertEqual(fuel_indicator(level), ("text", level - 1))

        source = MODULE.read_text(encoding="utf-8")
        transfer = re.search(r"Sub 転記to印刷様式\(\).*?^End Sub", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(transfer)
        body = transfer.group(0)
        level_one = body.split("If fuelLevel = 1 Then", 1)[1].split("ElseIf fuelLevel >= 2", 1)[0]
        self.assertIn("燃料残量表示開始位置(fuelCell)", level_one)
        self.assertNotIn("String(fuelLevel - 1,", level_one)
        self.assertIn('String(fuelLevel - 1, "■")', body)

    def test_center_left_and_right_alignment_follow_each_target_area(self):
        cases = (
            ("center", 0, 100.0 + (60.0 - 12.0) / 2),
            ("left", 1, 100.0 + 12.0),
            ("right", 1, 100.0 + 60.0 - 12.0 - 12.0),
        )
        for alignment, indent, expected in cases:
            with self.subTest(alignment=alignment):
                self.assertAlmostEqual(indicator_left(100.0, 60.0, 12.0, alignment, indent), expected)

    def test_vba_placement_is_cell_based_and_does_not_branch_on_arbitrary_group_shapes(self):
        source = MODULE.read_text(encoding="utf-8")
        match = re.search(
            r"Private Function 燃料残量表示開始位置\b.*?^End Function",
            source,
            re.MULTILINE | re.DOTALL,
        )
        self.assertIsNotNone(match)
        body = match.group(0)
        self.assertIn("targetCell.MergeArea", body)
        self.assertIn("xlHAlignCenter", body)
        self.assertIn("xlHAlignRight", body)
        self.assertIn("IsNull(alignment)", body)
        self.assertNotIn("TopLeftCell", body)
        self.assertNotIn("msoGroup", body)

    def test_transfer_keeps_half_width_shape_and_existing_multi_star_path(self):
        source = MODULE.read_text(encoding="utf-8")
        transfer = re.search(r"Sub 転記to印刷様式\(\).*?^End Sub", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(transfer)
        body = transfer.group(0)
        self.assertIn('Left(wsDest.Shapes(fuelShapeIndex).Name, Len("FuelRemainHalf_"))', body)
        self.assertIn("燃料残量表示開始位置(fuelCell)", body)
        self.assertIn("fuelCell.Font.Size / 2", body)
        self.assertIn('fuelCell.MergeArea.Top + (fuelCell.MergeArea.Height - fuelCell.Font.Size) / 2', body)
        self.assertIn('String(fuelLevel - 1, "■")', body)


if __name__ == "__main__":
    unittest.main()
