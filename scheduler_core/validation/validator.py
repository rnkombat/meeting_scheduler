# scheduler_core/validation/validator.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple, Dict

from scheduler_core.domain.models import InputData


@dataclass(frozen=True)
class ValidationError(Exception):
    message: str


@dataclass(frozen=True)
class ValidationWarning:
    message: str


def validate_generation_start(now: datetime, generation_start: datetime) -> None:
    # 仕様：過去なら即エラー（最適化は実行しない）
    if generation_start < now:
        raise ValidationError("生成開始日時が現在より過去です。設定を確認してください。")


def validate_integrity(data: InputData) -> Tuple[List[ValidationWarning], None]:
    warnings: List[ValidationWarning] = []

    # 許可委員の人数チェック（全体で4名未満なら不可能）
    commissioners = [p for p in data.persons.values() if p.is_commissioner]
    if len(commissioners) < 4:
        raise ValidationError("許可委員フラグが立っている人物が4名未満です。")

    seniors = [p for p in commissioners if p.is_senior_commissioner]
    if len(seniors) < 2:
        raise ValidationError("上級生許可委員が全体で2名未満です（どの会議も上級生2名要件を満たせません）。")

    # 登山隊リーダーが人物マスタに存在するか
    for t in data.teams.values():
        if t.leader_pid not in data.persons:
            raise ValidationError(f"登山隊 {t.name} のリーダーIDが人物マスタにありません: {t.leader_pid}")
        # メンバーIDの存在チェック
        for mp in t.member_pids:
            if mp not in data.persons:
                raise ValidationError(f"登山隊 {t.name} のメンバーIDが人物マスタにありません: {mp}")

    # 固定会議の整合性チェック（固定は変更できないため厳格に）
    for fm in data.fixed_meetings:
        team = data.teams.get(fm.team_tid)
        if not team:
            raise ValidationError(f"固定会議の登山隊IDがマスタにありません: {fm.team_tid}")
        if fm.leader_pid != team.leader_pid:
            raise ValidationError(f"固定会議のリーダーが登山隊のリーダーと一致しません: {team.name}")
        # 利益相反（固定でも禁止）
        forbidden = set(team.member_pids) | {team.leader_pid}
        if len(set(fm.commissioner_pids)) != 4:
            raise ValidationError(f"固定会議で許可委員が重複しています: {team.name}")
        for pid in fm.commissioner_pids:
            if pid in forbidden:
                raise ValidationError(f"固定会議で利益相反が発生しています: {team.name}")
            if not data.persons[pid].is_commissioner:
                raise ValidationError(f"固定会議で許可委員フラグのない人物が割当されています: {team.name}")
        # 上級生2名以上（固定でも満たすべき）
        senior_cnt = sum(1 for pid in fm.commissioner_pids if data.persons[pid].is_senior_commissioner)
        if senior_cnt < 2:
            raise ValidationError(f"固定会議で上級生許可委員が2名未満です: {team.name}")

    # 固定会議の二重ブッキング警告
    occupied_map: Dict[Tuple[str, str, int], List[str]] = {}
    for fm in data.fixed_meetings:
        for pid in [fm.leader_pid] + list(fm.commissioner_pids):
            for sl in range(fm.start_slot, fm.start_slot + 4):
                key = (pid, fm.day.isoformat(), sl)
                occupied_map.setdefault(key, []).append(fm.team_tid)
    for (pid, day, sl), tids in occupied_map.items():
        if len(tids) >= 2:
            warnings.append(ValidationWarning(
                f"固定会議で人物の二重ブッキングが検出されました: pid={pid} {day} slot={sl}"
            ))

    # 締切日/必要回数の軽い警告（過密可能性）
    for t in data.teams.values():
        need = t.base_required + t.add_required
        if need <= 0:
            continue
        # 同日複数禁止なので、最短でも必要日数=need日
        # ※ここは実際には候補数・中1日等が絡むので警告止まり
        # 呼び出し側で「生成開始日〜締切」の日数と比較して警告できる。
        pass

    return warnings, None
