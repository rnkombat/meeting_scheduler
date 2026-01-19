# meeting_scheduler/preprocessing/preprocess.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, timedelta
from typing import Dict, List, Set, Tuple

from meeting_scheduler.domain.models import InputData, CandidateSlot, FixedMeeting
from meeting_scheduler.domain.timegrid import TimeGrid
from meeting_scheduler.config import AppConfig


@dataclass(frozen=True)
class Preprocessed:
    candidates_by_team: Dict[str, List[CandidateSlot]]  # tid -> 候補
    can_attend: Dict[str, Dict[date, Dict[int, bool]]]  # pid -> day -> start_slot -> bool
    occupied: Dict[str, Dict[date, Dict[int, bool]]]    # pid -> day -> slot -> bool（固定会議で埋まっている）
    fixed_by_team_sorted: Dict[str, List[FixedMeeting]]


def build_can_attend(data: InputData, cfg: AppConfig, grid: TimeGrid) -> Dict[str, Dict[date, Dict[int, bool]]]:
    out: Dict[str, Dict[date, Dict[int, bool]]] = {}
    for pid, dmap in data.avail.items():
        out[pid] = {}
        for d, slots in dmap.items():
            out[pid][d] = {}
            for s in range(0, cfg.latest_start_slot + 1):
                ok = True
                for ss in grid.meeting_slots_covered(s, cfg.meeting_slots):
                    v = slots.get(ss, 4)
                    if v in (0, 4):
                        ok = False
                        break
                out[pid][d][s] = ok
    return out


def build_occupied_from_fixed(data: InputData, cfg: AppConfig, grid: TimeGrid) -> Dict[str, Dict[date, Dict[int, bool]]]:
    occ: Dict[str, Dict[date, Dict[int, bool]]] = {pid: {} for pid in data.persons.keys()}
    for fm in data.fixed_meetings:
        involved = [fm.leader_pid] + list(fm.commissioner_pids)
        for pid in involved:
            occ.setdefault(pid, {})
            occ[pid].setdefault(fm.day, {i: False for i in range(cfg.slots_per_day)})
            for sl in grid.meeting_slots_covered(fm.start_slot, cfg.meeting_slots):
                if 0 <= sl < cfg.slots_per_day:
                    occ[pid][fm.day][sl] = True
    return occ


def fixed_by_team_sorted(data: InputData, grid: TimeGrid) -> Dict[str, List[FixedMeeting]]:
    by: Dict[str, List[FixedMeeting]] = {}
    for fm in data.fixed_meetings:
        by.setdefault(fm.team_tid, []).append(fm)
    for tid in by:
        by[tid].sort(key=lambda x: (x.day, x.start_slot))
    return by


def generate_candidates(
    data: InputData,
    cfg: AppConfig,
    grid: TimeGrid,
    can_attend: Dict[str, Dict[date, Dict[int, bool]]],
    occupied: Dict[str, Dict[date, Dict[int, bool]]],
    generation_start: datetime,
) -> Dict[str, List[CandidateSlot]]:
    """
    候補 (day, start_slot) を列挙する。
    条件：
    - generation_start以降
    - 締切日以前
    - 当日は開始時刻が generation_start より前は禁止
    - リーダーが参加可能（can_attend）かつ occupiedで埋まっていない
    """
    start_day = generation_start.date()
    start_minutes = generation_start.hour * 60 + generation_start.minute

    out: Dict[str, List[CandidateSlot]] = {}

    for tid, team in data.teams.items():
        leader = team.leader_pid
        out[tid] = []

        d = start_day
        while d <= team.deadline:
            # その日がavailに存在しない場合は候補無し
            if d not in data.avail.get(leader, {}):
                d += timedelta(days=1)
                continue

            for s in range(0, cfg.latest_start_slot + 1):
                st_time = grid.slot_to_time(s)
                st_min = st_time.hour * 60 + st_time.minute

                # generation_start当日: 開始が生成開始時刻より前は禁止
                if d == start_day and st_min < start_minutes:
                    continue

                # リーダーの可否
                if not can_attend.get(leader, {}).get(d, {}).get(s, False):
                    continue

                # 固定会議で占有されていないか（リーダーの4セルにoccupiedがあれば不可）
                ok_occ = True
                for sl in grid.meeting_slots_covered(s, cfg.meeting_slots):
                    if occupied.get(leader, {}).get(d, {}).get(sl, False):
                        ok_occ = False
                        break
                if not ok_occ:
                    continue

                out[tid].append(CandidateSlot(
                    team_tid=tid,
                    day=d,
                    start_slot=s,
                    dt_idx=grid.dt_index(d, s)
                ))

            d += timedelta(days=1)

        # 候補を時系列順に
        out[tid].sort(key=lambda c: c.dt_idx)

    return out


def preprocess_all(data: InputData, cfg: AppConfig, grid: TimeGrid, generation_start: datetime) -> Preprocessed:
    can = build_can_attend(data, cfg, grid)
    occ = build_occupied_from_fixed(data, cfg, grid)
    fixed_sorted = fixed_by_team_sorted(data, grid)
    cand = generate_candidates(data, cfg, grid, can, occ, generation_start)
    return Preprocessed(
        candidates_by_team=cand,
        can_attend=can,
        occupied=occ,
        fixed_by_team_sorted=fixed_sorted
    )
