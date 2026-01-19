# meeting_scheduler/gui/app.py
from __future__ import annotations

import streamlit as st
from datetime import datetime
from dateutil import tz

from meeting_scheduler.config import DEFAULT_CONFIG
from meeting_scheduler.domain.timegrid import TimeGrid
from meeting_scheduler.io_layer.xlsx_reader import XlsxReader
from meeting_scheduler.validation.validator import validate_generation_start, validate_integrity, ValidationError
from meeting_scheduler.preprocessing.preprocess import preprocess_all
from meeting_scheduler.optimization.milp import solve_milp
from meeting_scheduler.reporting.report import build_meeting_table, build_team_summary, build_person_summary
from meeting_scheduler.reporting.export_xlsx import export_result_xlsx


def main():
    cfg = DEFAULT_CONFIG
    grid = TimeGrid(day_start_hour=cfg.day_start_hour, slots_per_day=cfg.slots_per_day)
    reader = XlsxReader(cfg=cfg, grid=grid)

    st.title("登山サークル審議会議スケジューラ（MILP + Gurobi）")

    jst = tz.gettz(cfg.timezone_name)
    now = datetime.now(tz=jst)

    st.header("入力")
    schedule_files = st.text_area("人員スケジュール（月ごとxlsx）複数可（改行区切り）").strip().splitlines()
    teams_file = st.text_input("登山隊情報 xlsx").strip()
    fixed_files = st.text_area("既存スケジュール xlsx 複数可（改行区切り）").strip().splitlines()
    prev_files = st.text_area("過去結果（result.xlsx）複数可（改行区切り）").strip().splitlines()
    add_file = st.text_input("追加審議要求 xlsx").strip()

    st.subheader("生成開始日時（必須）")
    gen_date = st.date_input("日付", value=now.date())
    gen_time = st.time_input("時刻", value=now.time().replace(second=0, microsecond=0))
    generation_start = datetime.combine(gen_date, gen_time).replace(tzinfo=jst)

    st.header("Solver設定")
    time_limit = st.number_input("TimeLimit(sec)", min_value=1, max_value=600, value=cfg.solver.time_limit_sec)
    mip_gap = st.number_input("MIPGap", min_value=0.0, max_value=0.5, value=float(cfg.solver.mip_gap), step=0.001)

    # 反映（簡易）
    cfg2 = cfg.__class__(
        day_start_hour=cfg.day_start_hour,
        slots_per_day=cfg.slots_per_day,
        meeting_slots=cfg.meeting_slots,
        latest_start_slot=cfg.latest_start_slot,
        schedule_col_start=cfg.schedule_col_start,
        schedule_col_end=cfg.schedule_col_end,
        schedule_row_start=cfg.schedule_row_start,
        timezone_name=cfg.timezone_name,
        penalty=cfg.penalty,
        weights=cfg.weights,
        solver=cfg.solver.__class__(time_limit_sec=int(time_limit), mip_gap=float(mip_gap), threads=cfg.solver.threads)
    )

    run =
