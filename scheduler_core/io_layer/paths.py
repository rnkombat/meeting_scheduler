# scheduler_core/io_layer/paths.py
from dataclasses import dataclass
from typing import List, Optional


@dataclass(frozen=True)
class InputPaths:
    """
    複数月=複数ファイルを読み込む
    schedule_month_files: 人員スケジュール xlsx（月ごと）
    teams_file: 登山隊情報 xlsx
    fixed_files: 既存スケジュール xlsx（複数可）
    previous_result_files: 過去に出力した結果xlsx（追加審議のため固定として取り込む）
    add_request_file: 追加審議要求 xlsx
    """
    schedule_month_files: List[str]
    teams_file: str
    fixed_files: List[str]
    previous_result_files: List[str]
    add_request_file: str

    # シート名（運用で変えるならここだけ）
    team_sheet_name: str = "teams"
    fixed_sheet_name: str = "fixed"
    add_sheet_name: str = "add"
    result_sheet_name: str = "result"  # 過去結果xlsxの会議一覧シート名
