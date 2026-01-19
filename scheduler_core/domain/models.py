# scheduler_core/domain/models.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Set, Tuple, Optional


@dataclass(frozen=True)
class Person:
    pid: str
    name: str
    is_commissioner: bool
    is_senior_commissioner: bool


@dataclass(frozen=True)
class Team:
    tid: str
    name: str
    leader_pid: str
    member_pids: Set[str]
    deadline: date
    base_required: int
    add_required: int


@dataclass(frozen=True)
class FixedMeeting:
    """既存固定会議（または過去結果リザルトから取り込む固定扱い）"""
    team_tid: str
    day: date
    start_slot: int
    leader_pid: str
    commissioner_pids: Tuple[str, str, str, str]
    meeting_no: Optional[int] = None  # 任意


@dataclass(frozen=True)
class CandidateSlot:
    """新規会議の候補（チーム×日付×開始スロット）"""
    team_tid: str
    day: date
    start_slot: int
    dt_idx: int  # 順序制約用


@dataclass
class InputData:
    persons: Dict[str, Person]                 # pid -> Person
    teams: Dict[str, Team]                     # tid -> Team
    name_to_pid: Dict[str, str]                # 人名 -> pid
    team_name_to_tid: Dict[str, str]           # 登山隊名 -> tid

    # avail[pid][day][slot] = 0..4（複数月マージ済）
    avail: Dict[str, Dict[date, Dict[int, int]]]

    fixed_meetings: List[FixedMeeting]         # 既存＋過去結果リザルト取込（固定扱い）


@dataclass
class SolutionMeeting:
    team_tid: str
    day: date
    start_slot: int
    leader_pid: str
    commissioner_pids: Tuple[str, str, str, str]
    meeting_no: int
    handover_person_pid: Optional[str]  # 引き継ぎ担当（前回と共通の許可委員から1名選ぶ）


@dataclass
class SolveResult:
    feasible: bool
    status: str
    meetings: List[SolutionMeeting]     # 新規のみ
    iis_summary: Optional[str] = None
