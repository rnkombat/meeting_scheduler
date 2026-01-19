# meeting_scheduler/main_cli.py
from __future__ import annotations

import argparse
from datetime import datetime
from dateutil import tz

from meeting_scheduler.config import DEFAULT_CONFIG
from meeting_scheduler.domain.timegrid import TimeGrid
from meeting_scheduler.io_layer.paths import InputPaths
from meeting_scheduler.io_layer.xlsx_reader import XlsxReader
from meeting_scheduler.validation.validator import validate_generation_start, validate_integrity, ValidationError
from meeting_scheduler.preprocessing.preprocess import preprocess_all
from meeting_scheduler.optimization.milp import solve_milp
from meeting_scheduler.reporting.report import build_meeting_table, build_team_summary, build_person_summary
from meeting_scheduler.reporting.export_xlsx import export_result_xlsx


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--schedule", nargs="+", required=True, help="人員スケジュール月ファイル群（xlsx）")
    p.add_argument("--teams", required=True, help="登山隊情報 xlsx")
    p.add_argument("--fixed", nargs="*", default=[], help="既存スケジュール xlsx（複数可）")
    p.add_argument("--prev", nargs="*", default=[], help="過去結果 xlsx（複数可。追加審議で固定として読む）")
    p.add_argument("--add", required=True, help="追加審議要求 xlsx")
    p.add_argument("--generation_start", required=True, help="生成開始日時（例: 2026-01-20 10:30）")
    p.add_argument("--out", default="assets/output/result.xlsx", help="出力xlsx")
    return p.parse_args()


def main():
    args = parse_args()
    cfg = DEFAULT_CONFIG
    grid = TimeGrid(day_start_hour=cfg.day_start_hour, slots_per_day=cfg.slots_per_day)

    # timezone
    jst = tz.gettz(cfg.timezone_name)
    now = datetime.now(tz=jst)
    generation_start = datetime.strptime(args.generation_start, "%Y-%m-%d %H:%M").replace(tzinfo=jst)

    try:
        validate_generation_start(now, generation_start)
    except ValidationError as e:
        print(f"[ERROR] {e.message}")
        return 1

    paths = InputPaths(
        schedule_month_files=args.schedule,
        teams_file=args.teams,
        fixed_files=args.fixed,
        previous_result_files=args.prev,
        add_request_file=args.add
    )

    reader = XlsxReader(cfg=cfg, grid=grid)

    # fixed_files / prev_files は「ファイル,シート名」の組で渡す
    fixed_pairs = [(f, paths.fixed_sheet_name) for f in paths.fixed_files]
    prev_pairs = [(f, paths.result_sheet_name) for f in paths.previous_result_files]

    data = reader.build_input_data(
        schedule_month_files=paths.schedule_month_files,
        teams_file=paths.teams_file,
        teams_sheet=paths.team_sheet_name,
        fixed_files=fixed_pairs,
        prev_result_files=prev_pairs,
        add_file=paths.add_request_file,
        add_sheet=paths.add_sheet_name,
    )

    try:
        warnings, _ = validate_integrity(data)
        for w in warnings:
            print(f"[WARN] {w.message}")
    except ValidationError as e:
        print(f"[ERROR] {e.message}")
        return 1

    pre = preprocess_all(data, cfg, grid, generation_start)

    result = solve_milp(
        data=data,
        pre_cand=pre.candidates_by_team,
        fixed_sorted=pre.fixed_by_team_sorted,
        cfg=cfg,
        grid=grid
    )

    if not result.feasible:
        print("[RESULT] 生成不可能")
        if result.iis_summary:
            print(result.iis_summary)
        return 2

    meeting_df = build_meeting_table(data, result, cfg, grid)
    team_df = build_team_summary(data, meeting_df)
    person_df = build_person_summary(data, meeting_df)

    out_path = export_result_xlsx(args.out, meeting_df, team_df, person_df)
    print(f"[RESULT] OK: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
