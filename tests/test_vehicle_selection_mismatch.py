"""Static and logic tests for Module1's vehicle-selection mismatch detection.

Excel/VBA cannot run on this PC.  The focused model mirrors the required Q-and-time
rules while the source checks ensure Module1's real detection path uses those rules.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import unittest

TOLERANCE_KM = 2.0
MODULE = Path(__file__).resolve().parents[1] / "Module1"


@dataclass(frozen=True)
class Record:
    answer: str
    vehicle: str
    start: datetime
    end: datetime
    end_meter: float


def overlaps(a: Record, b: Record) -> bool:
    return a.start < b.end and a.end > b.start


def nearest_neighbors(records: list[Record], target: Record, vehicle: str):
    before = [r for r in records if r != target and r.vehicle == vehicle and r.end <= target.start]
    after = [r for r in records if r != target and r.vehicle == vehicle and r.start >= target.end]
    return (
        max(before, key=lambda r: r.end, default=None),
        min(after, key=lambda r: r.start, default=None),
    )


def candidate_vehicle(records: list[Record], target: Record) -> str | None:
    assigned_previous, assigned_next = nearest_neighbors(records, target, target.vehicle)
    if assigned_previous is None or assigned_next is None:
        return None
    assigned_is_unnatural = (
        target.end_meter < assigned_previous.end_meter - TOLERANCE_KM
        or target.end_meter > assigned_next.end_meter + TOLERANCE_KM
    )
    if not assigned_is_unnatural:
        return None

    candidates: list[str] = []
    for vehicle in sorted({r.vehicle for r in records if r.vehicle != target.vehicle}):
        if any(r.vehicle == vehicle and r != target and overlaps(r, target) for r in records):
            continue
        previous, following = nearest_neighbors(records, target, vehicle)
        if previous is None or following is None:
            continue
        if (
            previous.end_meter <= target.end_meter + TOLERANCE_KM
            and target.end_meter <= following.end_meter + TOLERANCE_KM
        ):
            candidates.append(vehicle)
    return candidates[0] if len(candidates) == 1 else None


def rec(answer: str, vehicle: str, start: str, end: str, meter: float) -> Record:
    return Record(answer, vehicle, datetime.fromisoformat(start), datetime.fromisoformat(end), meter)


class VehicleSelectionMismatchTests(unittest.TestCase):
    def test_case_a_detects_7391_without_using_start_meter(self):
        target = rec("601", "1889", "2026-09-17 09:30", "2026-09-17 12:00", 90707)
        records = [
            rec("600", "1889", "2026-09-16 09:00", "2026-09-16 12:00", 60392),
            target,
            rec("602", "1889", "2026-09-18 09:00", "2026-09-18 12:00", 60409),
            rec("700", "7391", "2026-09-16 09:00", "2026-09-16 12:00", 90684),
            rec("701", "7391", "2026-09-18 09:00", "2026-09-18 12:00", 90729),
        ]
        self.assertEqual(candidate_vehicle(records, target), "7391")

    def test_case_b_detects_7391(self):
        target = rec("608", "1889", "2026-09-20 09:30", "2026-09-20 12:00", 90930)
        records = [
            rec("607", "1889", "2026-09-19 09:00", "2026-09-19 12:00", 60409),
            target,
            rec("609", "1889", "2026-09-21 09:00", "2026-09-21 12:00", 60484),
            rec("702", "7391", "2026-09-19 09:00", "2026-09-19 12:00", 90878),
            rec("703", "7391", "2026-09-21 09:00", "2026-09-21 12:00", 90945),
        ]
        self.assertEqual(candidate_vehicle(records, target), "7391")

    def test_normal_series_has_no_warning(self):
        target = rec("610", "1889", "2026-09-23 09:30", "2026-09-23 12:00", 60500)
        records = [
            rec("609", "1889", "2026-09-22 09:00", "2026-09-22 12:00", 60480),
            target,
            rec("611", "1889", "2026-09-24 09:00", "2026-09-24 12:00", 60520),
            rec("704", "7391", "2026-09-22 09:00", "2026-09-22 12:00", 60490),
            rec("705", "7391", "2026-09-24 09:00", "2026-09-24 12:00", 60530),
        ]
        self.assertIsNone(candidate_vehicle(records, target))

    def test_two_equally_valid_candidates_are_not_named(self):
        target = rec("612", "1889", "2026-09-26 09:30", "2026-09-26 12:00", 90707)
        records = [
            rec("611", "1889", "2026-09-25 09:00", "2026-09-25 12:00", 60392),
            target,
            rec("613", "1889", "2026-09-27 09:00", "2026-09-27 12:00", 60409),
            rec("710", "7391", "2026-09-25 09:00", "2026-09-25 12:00", 90684),
            rec("711", "7391", "2026-09-27 09:00", "2026-09-27 12:00", 90729),
            rec("720", "5555", "2026-09-25 09:00", "2026-09-25 12:00", 90684),
            rec("721", "5555", "2026-09-27 09:00", "2026-09-27 12:00", 90729),
        ]
        self.assertIsNone(candidate_vehicle(records, target))

    def test_overlapping_candidate_vehicle_is_excluded(self):
        target = rec("614", "1889", "2026-09-29 09:30", "2026-09-29 12:00", 90707)
        records = [
            rec("613", "1889", "2026-09-28 09:00", "2026-09-28 12:00", 60392),
            target,
            rec("615", "1889", "2026-09-30 09:00", "2026-09-30 12:00", 60409),
            rec("730", "7391", "2026-09-28 09:00", "2026-09-28 12:00", 90684),
            rec("731", "7391", "2026-09-29 10:00", "2026-09-29 11:00", 90700),
            rec("732", "7391", "2026-09-30 09:00", "2026-09-30 12:00", 90729),
        ]
        self.assertIsNone(candidate_vehicle(records, target))

    def test_module_source_uses_q_and_time_not_p_for_vehicle_detection(self):
        source = MODULE.read_text(encoding="utf-8")
        section = re.search(
            r"Private Sub 車両選択候補をチェック.*?^End Sub",
            source,
            re.MULTILINE | re.DOTALL,
        )
        self.assertIsNotNone(section)
        body = section.group(0)
        self.assertNotIn('"P"', body)
        for required in (
            '"Q"', '"J"', '"K"', '"N"', '"O"', '"W"',
            "vehicleOverlaps", "回答番号：", "運転日：", "使用時間：", "運転者：",
            "登録車両", "候補車両", "登録車両 ", "候補車両 ",
        ):
            self.assertIn(required, body)

    def test_record_sheet_organize_checks_candidates_before_p_is_filled(self):
        source = MODULE.read_text(encoding="utf-8")
        organize = re.search(r"Private Sub 記録シート整理を実行.*?^End Sub", source, re.MULTILINE | re.DOTALL)
        self.assertIsNotNone(organize)
        body = organize.group(0)
        self.assertLess(body.index("車両選択候補の事前チェック"), body.index('Cells(i, "P").Value'))


if __name__ == "__main__":
    unittest.main()
