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
