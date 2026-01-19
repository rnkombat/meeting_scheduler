# meeting_scheduler/reporting/export_xlsx.py
from __future__ import annotations

from pathlib import Path
from typing import Tuple

import pandas as pd


def export_result_xlsx(
    out_path: str,
    meeting_df: pd.DataFrame,
    team_df: pd.DataFrame,
    person_df: pd.DataFrame,
) -> str:
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        # 重要：過去結果として取り込むのは "result" シート名を想定
        meeting_df.to_excel(w, sheet_name="result", index=False)
        team_df.to_excel(w, sheet_name="team_summary", index=False)
        person_df.to_excel(w, sheet_name="person_summary", index=False)
    return out_path
