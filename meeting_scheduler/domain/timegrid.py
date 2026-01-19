# meeting_scheduler/domain/timegrid.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import List, Tuple


@dataclass(frozen=True)
class TimeGrid:
    """日付×開始スロット（30分刻み）を扱う"""
    day_start_hour: int = 9
    slots_per_day: int = 26
    slot_minutes: int = 30

    def slot_to_time(self, slot: int) -> time:
        minutes = self.day_start_hour * 60 + slot * self.slot_minutes
        hh = minutes // 60
        mm = minutes % 60
        return time(hh, mm)

    def meeting_end_time(self, start_slot: int, meeting_slots: int) -> time:
        minutes = self.day_start_hour * 60 + (start_slot + meeting_slots) * self.slot_minutes
        return time(minutes // 60, minutes % 60)

    def meeting_slots_covered(self, start_slot: int, meeting_slots: int) -> List[int]:
        return list(range(start_slot, start_slot + meeting_slots))

    def dt_index(self, d: date, start_slot: int) -> int:
        """順序制約用：日付と開始スロットから単調な整数インデックス"""
        # date.toordinalは日単位連番
        return d.toordinal() * 100 + start_slot
