# scheduler_core/reporting/report.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Tuple

import pandas as pd

from scheduler_core.domain.models import InputData, SolveResult, SolutionMeeting, FixedMeeting
from scheduler_core.domain.timegrid import TimeGrid
from scheduler_core.config import AppConfig


def _pid_to_name(data: InputData, pid: str) -> str:
    return data.persons[pid].name


def build_meeting_table(
    data: InputData,
    result: SolveResult,
    cfg: AppConfig,
    grid: TimeGrid
) -> pd.DataFrame:
    rows = []

    # 既存固定（統合表示）
    for fm in data.fixed_meetings:
        team = data.teams[fm.team_tid]
        st = grid.slot_to_time(fm.start_slot).strftime("%H:%M")
        et = grid.meeting_end_time(fm.start_slot, cfg.meeting_slots).strftime("%H:%M")
        comm_names = [_pid_to_name(data, p) for p in fm.commissioner_pids]
        senior_cnt = sum(1 for p in fm.commissioner_pids if data.persons[p].is_senior_commissioner)

        rows.append(dict(
            source="fixed",
            team_name=team.name,
            meeting_date=fm.day.isoformat(),
            start_time=st,
            end_time=et,
            leader_name=_pid_to_name(data, fm.leader_pid),
            comm1=comm_names[0], comm2=comm_names[1], comm3=comm_names[2], comm4=comm_names[3],
            senior_count=senior_cnt,
            meeting_no=fm.meeting_no if fm.meeting_no is not None else "",
            handover_person=""
        ))

    # 新規
    for sm in result.meetings:
        team = data.teams[sm.team_tid]
        st = grid.slot_to_time(sm.start_slot).strftime("%H:%M")
        et = grid.meeting_end_time(sm.start_slot, cfg.meeting_slots).strftime("%H:%M")
        comm_names = [_pid_to_name(data, p) for p in sm.commissioner_pids]
        senior_cnt = sum(1 for p in sm.commissioner_pids if data.persons[p].is_senior_commissioner)

        rows.append(dict(
            source="new",
            team_name=team.name,
            meeting_date=sm.day.isoformat(),
            start_time=st,
            end_time=et,
            leader_name=_pid_to_name(data, sm.leader_pid),
            comm1=comm_names[0], comm2=comm_names[1], comm3=comm_names[2], comm4=comm_names[3],
            senior_count=senior_cnt,
            meeting_no=sm.meeting_no,
            handover_person=_pid_to_name(data, sm.handover_person_pid) if sm.handover_person_pid else ""
        ))

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["team_name", "meeting_date", "start_time", "source"]).reset_index(drop=True)
    return df


def build_team_summary(data: InputData, meeting_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for t in data.teams.values():
        need = t.base_required + t.add_required
        done = int((meeting_df["team_name"] == t.name).sum()) if not meeting_df.empty else 0
        rows.append(dict(
            team_name=t.name,
            required_total=need,
            done_total=done,
            normal_plus_one_ok=(done >= t.base_required + 1) if t.base_required > 0 else False,
            # 1日バッファは近似なのでここでは未評価（厳密化するなら最終会議日を見て判定）
        ))
    return pd.DataFrame(rows)


def build_person_summary(data: InputData, meeting_df: pd.DataFrame) -> pd.DataFrame:
    # 既存+新規の統合 df を使う
    counts = {p.name: dict(total=0, leader=0, commissioner=0) for p in data.persons.values()}

    for _, r in meeting_df.iterrows():
        leader = r["leader_name"]
        counts[leader]["total"] += 1
        counts[leader]["leader"] += 1
        for c in ["comm1", "comm2", "comm3", "comm4"]:
            nm = r[c]
            counts[nm]["total"] += 1
            counts[nm]["commissioner"] += 1

    rows = []
    for name, d in counts.items():
        rows.append(dict(
            person_name=name,
            total_attend=d["total"],
            leader_count=d["leader"],
            commissioner_count=d["commissioner"]
        ))
    df = pd.DataFrame(rows).sort_values(["total_attend", "person_name"], ascending=[False, True]).reset_index(drop=True)
    return df
