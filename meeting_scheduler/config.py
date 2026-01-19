# meeting_scheduler/config.py
from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class PenaltyConfig:
    """セル値のペナルティ（ソフト制約用）"""
    value2: int = 1   # 2:できれば避けたい
    value3: int = 2   # 3:未定


@dataclass(frozen=True)
class ObjectiveWeights:
    """目的関数の重み（必要ならGUIで調整）"""
    w_availability: float = 1.0     # 2/3セル利用ペナルティ
    w_gap_rule: float = 0.5         # 中1日以上（違反ペナルティ）
    w_finish_buffer: float = 0.8    # 最終会議が締切前日以前の加点
    w_normal_plus_one: float = 0.4  # 通常+1回バッファの加点
    w_load_balance: float = 1.5     # 最大負担最小化（Wmax）


@dataclass(frozen=True)
class SolverConfig:
    time_limit_sec: int = 60
    mip_gap: float = 0.01
    threads: int = 0  # 0=自動


@dataclass(frozen=True)
class AppConfig:
    # 1日スロット（固定）
    day_start_hour: int = 9
    slots_per_day: int = 26  # 09:00〜21:30 start slot
    meeting_slots: int = 4   # 2時間固定（30分×4）
    latest_start_slot: int = 22  # 20:00開始が最遅（09:00=slot0 → 20:00=slot22）

    # xlsxスケジュールの列仕様
    schedule_col_start: str = "C"
    schedule_col_end: str = "AB"
    schedule_row_start: int = 2  # 行2が日1

    # 生成開始日時チェック（Asia/Tokyo想定）
    timezone_name: str = "Asia/Tokyo"

    penalty: PenaltyConfig = PenaltyConfig()
    weights: ObjectiveWeights = ObjectiveWeights()
    solver: SolverConfig = SolverConfig()


DEFAULT_CONFIG = AppConfig()
