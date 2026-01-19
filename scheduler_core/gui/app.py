# scheduler_core/gui/app.py
from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
import sys

import pandas as pd
import streamlit as st
from dateutil import tz

# 日本語コメント: Streamlitは実行ディレクトリが変わるため、リポジトリルートをパスに追加する。
ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from scheduler_core.config import DEFAULT_CONFIG
from scheduler_core.domain.timegrid import TimeGrid
from scheduler_core.io_layer.paths import InputPaths
from scheduler_core.io_layer.xlsx_reader import XlsxReader
from scheduler_core.validation.validator import validate_generation_start, validate_integrity, ValidationError
from scheduler_core.preprocessing.preprocess import preprocess_all
from scheduler_core.optimization.milp import solve_milp
from scheduler_core.reporting.report import build_meeting_table, build_team_summary, build_person_summary


def _export_result_bytes(meeting_df: pd.DataFrame, team_df: pd.DataFrame, person_df: pd.DataFrame) -> bytes:
    """日本語コメント: Streamlitダウンロード用にxlsxをメモリに書き出す。"""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        meeting_df.to_excel(w, sheet_name="result", index=False)
        team_df.to_excel(w, sheet_name="team_summary", index=False)
        person_df.to_excel(w, sheet_name="person_summary", index=False)
    return buf.getvalue()


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

    st.caption("※複数月は複数ファイルを指定してください。過去結果は追加審議のため固定扱いで読み込みます。")

    run = st.button("最適化を実行")

    if not run:
        st.stop()

    # 日本語コメント: 入力の最低限チェック
    if not schedule_files or not schedule_files[0]:
        st.error("人員スケジュール（月ごとxlsx）のパスが未入力です。")
        st.stop()
    if not teams_file:
        st.error("登山隊情報 xlsx のパスが未入力です。")
        st.stop()
    if not add_file:
        st.error("追加審議要求 xlsx のパスが未入力です。")
        st.stop()

    # 日本語コメント: 生成開始日時の検証（過去なら即エラー）
    try:
        validate_generation_start(now, generation_start)
    except ValidationError as e:
        st.error(e.message)
        st.stop()

    # 日本語コメント: 入力読み込み
    paths = InputPaths(
        schedule_month_files=[p for p in schedule_files if p.strip()],
        teams_file=teams_file,
        fixed_files=[p for p in fixed_files if p.strip()],
        previous_result_files=[p for p in prev_files if p.strip()],
        add_request_file=add_file
    )
    fixed_pairs = [(f, paths.fixed_sheet_name) for f in paths.fixed_files]
    prev_pairs = [(f, paths.result_sheet_name) for f in paths.previous_result_files]

    try:
        data = reader.build_input_data(
            schedule_month_files=paths.schedule_month_files,
            teams_file=paths.teams_file,
            teams_sheet=paths.team_sheet_name,
            fixed_files=fixed_pairs,
            prev_result_files=prev_pairs,
            add_file=paths.add_request_file,
            add_sheet=paths.add_sheet_name,
        )
    except Exception as e:
        st.error(f"入力読み込みでエラーが発生しました: {e}")
        st.stop()

    # 日本語コメント: データ整合性チェック
    try:
        warnings, _ = validate_integrity(data)
        for w in warnings:
            st.warning(w.message)
    except ValidationError as e:
        st.error(e.message)
        st.stop()

    # 日本語コメント: 前処理 → 最適化
    pre = preprocess_all(data, cfg2, grid, generation_start)
    result = solve_milp(
        data=data,
        pre_cand=pre.candidates_by_team,
        fixed_sorted=pre.fixed_by_team_sorted,
        cfg=cfg2,
        grid=grid
    )

    if not result.feasible:
        st.error("生成不可能")
        if result.iis_summary:
            st.code(result.iis_summary)
        st.stop()

    # 日本語コメント: 出力整形
    meeting_df = build_meeting_table(data, result, cfg2, grid)
    team_df = build_team_summary(data, meeting_df)
    person_df = build_person_summary(data, meeting_df)

    st.success("最適化が完了しました。")

    tab1, tab2, tab3 = st.tabs(["会議一覧", "登山隊別サマリ", "人物別負担"])
    with tab1:
        st.dataframe(meeting_df, use_container_width=True)
    with tab2:
        st.dataframe(team_df, use_container_width=True)
    with tab3:
        st.dataframe(person_df, use_container_width=True)

    # 日本語コメント: ダウンロード（xlsx）
    xlsx_bytes = _export_result_bytes(meeting_df, team_df, person_df)
    st.download_button(
        label="結果xlsxをダウンロード",
        data=xlsx_bytes,
        file_name="result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    main()
