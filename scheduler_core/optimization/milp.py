# scheduler_core/optimization/milp.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Tuple, Set, Optional

import gurobipy as gp
from gurobipy import GRB

from scheduler_core.domain.models import InputData, CandidateSlot, SolveResult, SolutionMeeting, FixedMeeting
from scheduler_core.domain.timegrid import TimeGrid
from scheduler_core.config import AppConfig


@dataclass(frozen=True)
class ModelArtifacts:
    model: gp.Model


def _availability_penalty(value: int, cfg: AppConfig) -> int:
    # 1:0, 2:小, 3:中, 4:不可（候補生成で除外される想定）
    if value == 2:
        return cfg.penalty.value2
    if value == 3:
        return cfg.penalty.value3
    return 0


def solve_milp(
    data: InputData,
    pre_cand: Dict[str, List[CandidateSlot]],
    fixed_sorted: Dict[str, List[FixedMeeting]],
    cfg: AppConfig,
    grid: TimeGrid,
) -> SolveResult:
    """
    実装方針（重要）
    - 既存固定会議は「固定としてカウント＆占有」されている前提（preprocessでoccupiedへ反映済）
    - 引き継ぎ：チームごとに固定会議を時系列で k=1..Kfixed とみなし、その後の新規を接続
    - 新規会議数は「必要回数 - fixed_count」を最低限満たす（不足なら infeasible）
    """
    m = gp.Model("scheduler_core")
    m.Params.OutputFlag = 0
    m.Params.TimeLimit = cfg.solver.time_limit_sec
    m.Params.MIPGap = cfg.solver.mip_gap
    if cfg.solver.threads and cfg.solver.threads > 0:
        m.Params.Threads = cfg.solver.threads

    persons = list(data.persons.keys())

    seniors = {pid for pid, p in data.persons.items() if p.is_senior_commissioner and p.is_commissioner}
    commissioners = {pid for pid, p in data.persons.items() if p.is_commissioner}

    # --- kの最大数をチームごとに決める（必要分＋通常+1バッファ用の余地を作る） ---
    # 追加回数には+1を適用しないが、実装簡略のため「新規上限」は base+add+1 まで許す
    # （目的関数で余計な会議は基本選ばれないようにする）
    Kmax_by_team: Dict[str, int] = {}
    need_new_by_team: Dict[str, int] = {}
    fixed_count_by_team: Dict[str, int] = {}

    for tid, t in data.teams.items():
        fixed_count = len(fixed_sorted.get(tid, []))
        fixed_count_by_team[tid] = fixed_count

        need_total = t.base_required + t.add_required
        need_new = max(0, need_total - fixed_count)
        need_new_by_team[tid] = need_new

        # 通常+1バッファ（追加がある場合でも base+1 の余地だけ許す）
        kmax = need_new + (1 if t.base_required > 0 else 0)
        Kmax_by_team[tid] = kmax

        # 候補がゼロで必要があるなら即 infeasible に近いが、ここでは後で解かせる
        # （ただし候補ゼロの場合、Gurobiは即 infeasibleになる）

    # --- 変数 ---
    # y[tid,k,c] = 会議k(新規のk=1..Kmax) を候補cに置く
    # x[tid,k,c,p] = その会議にpを許可委員で割当
    y: Dict[Tuple[str, int, int], gp.Var] = {}
    x: Dict[Tuple[str, int, int, str], gp.Var] = {}

    # 負担カウント w[p] と Wmax
    w = {pid: m.addVar(vtype=GRB.INTEGER, lb=0, name=f"w[{pid}]") for pid in persons}
    Wmax = m.addVar(vtype=GRB.INTEGER, lb=0, name="Wmax")

    # 目的用：最終会議日が締切前日以前か（簡易：締切当日の会議にペナルティ）
    # → 会議が締切当日に置かれたら penalty を与える（本当に「最終日」判定までやると重い）
    # ここは「1日バッファ達成」を厳密化したければ追加実装（後述の拡張ポイント参照）。
    # 今回は実用優先で「締切当日の会議を嫌う」近似にしてある。
    deadline_day_penalty_terms = []

    availability_penalty_terms = []
    gap_penalty_terms = []  # 中1日違反（近似：連日配置を嫌う）
    normal_plus_one_reward_terms = []

    # 既存固定の負担を初期値として加算するため、固定出席回数を数える
    fixed_attend_count = {pid: 0 for pid in persons}
    for fm in data.fixed_meetings:
        fixed_attend_count[fm.leader_pid] += 1
        for pid in fm.commissioner_pids:
            fixed_attend_count[pid] += 1

    # --- 変数作成 ---
    for tid, cand_list in pre_cand.items():
        for k in range(1, Kmax_by_team[tid] + 1):
            for ci, c in enumerate(cand_list):
                y[(tid, k, ci)] = m.addVar(vtype=GRB.BINARY, name=f"y[{tid},{k},{ci}]")
                for pid in persons:
                    x[(tid, k, ci, pid)] = m.addVar(vtype=GRB.BINARY, name=f"x[{tid},{k},{ci},{pid}]")

    m.update()

    # --- 制約 ---
    # (1) 各kは高々1候補に配置
    for tid, cand_list in pre_cand.items():
        for k in range(1, Kmax_by_team[tid] + 1):
            m.addConstr(
                gp.quicksum(y[(tid, k, ci)] for ci in range(len(cand_list))) <= 1,
                name=f"at_most_one_slot[{tid},{k}]"
            )

    # (2) 必要回数達成（既存＋新規 >= base+add）
    for tid, t in data.teams.items():
        cand_list = pre_cand.get(tid, [])
        lhs_new = gp.quicksum(
            y[(tid, k, ci)]
            for k in range(1, Kmax_by_team[tid] + 1)
            for ci in range(len(cand_list))
        )
        m.addConstr(
            fixed_count_by_team[tid] + lhs_new >= (t.base_required + t.add_required),
            name=f"required_count[{tid}]"
        )

    # (3) 同一登山隊の同日複数禁止（新規同士）
    for tid, cand_list in pre_cand.items():
        # 日ごとに集計
        days = sorted({c.day for c in cand_list})
        for d in days:
            idxs = [ci for ci, c in enumerate(cand_list) if c.day == d]
            for k in range(1, Kmax_by_team[tid] + 1):
                # kごとではなく「全k合計でその日1回まで」
                pass
            m.addConstr(
                gp.quicksum(y[(tid, k, ci)]
                            for k in range(1, Kmax_by_team[tid] + 1)
                            for ci in idxs) <= 1,
                name=f"no_multi_same_day[{tid},{d}]"
            )

    # (4) 許可委員はちょうど4名（会議が置かれたときのみ）
    for tid, cand_list in pre_cand.items():
        for k in range(1, Kmax_by_team[tid] + 1):
            for ci in range(len(cand_list)):
                m.addConstr(
                    gp.quicksum(x[(tid, k, ci, pid)] for pid in persons) == 4 * y[(tid, k, ci)],
                    name=f"exact4[{tid},{k},{ci}]"
                )

    # (5) 上級生2名以上
    for tid, cand_list in pre_cand.items():
        for k in range(1, Kmax_by_team[tid] + 1):
            for ci in range(len(cand_list)):
                m.addConstr(
                    gp.quicksum(x[(tid, k, ci, pid)] for pid in seniors) >= 2 * y[(tid, k, ci)],
                    name=f"senior2[{tid},{k},{ci}]"
                )

    # (6) 利益相反・役割排他・許可委員属性
    for tid, t in data.teams.items():
        forbidden = set(t.member_pids) | {t.leader_pid}
        cand_list = pre_cand.get(tid, [])
        for k in range(1, Kmax_by_team[tid] + 1):
            for ci in range(len(cand_list)):
                for pid in persons:
                    if pid not in commissioners:
                        m.addConstr(x[(tid, k, ci, pid)] == 0, name=f"not_comm[{tid},{k},{ci},{pid}]")
                        continue
                    if pid in forbidden:
                        m.addConstr(x[(tid, k, ci, pid)] == 0, name=f"conflict[{tid},{k},{ci},{pid}]")

    # (7) 二重ブッキング禁止（新規同士＋固定によるoccupiedは候補生成で排除済）
    # 人物p・日d・スロットhについて、そのスロットにかかる会議へ同時参加しない
    # 参加= leader（そのチームのリーダー） + commissioner（x）
    # ※ occupied（固定）は preprocessで候補削減しているが、保険でここにも入れるなら上限を(1-occ)にできる
    # 今回は「新規同士」だけを厳密化する。
    for pid in persons:
        # 関係する候補を走査
        # dayごとにまとめる
        # 小規模前提の素直実装（重くなったら候補を事前にインデックス化）
        day_to_terms: Dict[date, Dict[int, List[gp.LinExpr]]] = {}
        for tid, cand_list in pre_cand.items():
            leader_pid = data.teams[tid].leader_pid
            for k in range(1, Kmax_by_team[tid] + 1):
                for ci, c in enumerate(cand_list):
                    covered = grid.meeting_slots_covered(c.start_slot, cfg.meeting_slots)
                    for h in covered:
                        day_to_terms.setdefault(c.day, {}).setdefault(h, [])
                        # commissioner参加
                        day_to_terms[c.day][h].append(x[(tid, k, ci, pid)])
                        # leader参加（pidがリーダーならyが参加）
                        if pid == leader_pid:
                            day_to_terms[c.day][h].append(y[(tid, k, ci)])

        for d, hmap in day_to_terms.items():
            for h, terms in hmap.items():
                if not terms:
                    continue
                m.addConstr(
                    gp.quicksum(terms) <= 1,
                    name=f"no_double_booking[{pid},{d},{h}]"
                )

    # (8) 引き継ぎ要件（前回と許可委員が1名以上共通）
    # A案：固定会議を系列に取り込む
    # 新規k=1 は「直前が固定の最後」または「固定無しなら引継ぎ不要」
    for tid, cand_list in pre_cand.items():
        fixed_list = fixed_sorted.get(tid, [])
        # 新規kは固定の後に続く、という位置づけ
        for k in range(1, Kmax_by_team[tid] + 1):
            # kに置かれた会議が存在するなら、引き継ぎ成立が必要
            # ただし「固定が0件かつk=1」は初回なので免除
            if len(fixed_list) == 0 and k == 1:
                continue

            # 前回の許可委員集合（固定 or 新規）
            # ここは線形式で「共通>=1」を作る必要がある。
            # 実装簡略：k-1が新規なら共通、k=1なら固定最後との共通。
            # ただし固定→新規の共通は「固定最後のメンバーに含まれる誰かが今回選ばれる」なので簡単。
            if k == 1:
                prev_comm = set(fixed_list[-1].commissioner_pids)
                # Σ_{ci} Σ_{p∈prev} x[tid,k,ci,p] >= Σ_{ci} y[tid,k,ci]
                m.addConstr(
                    gp.quicksum(x[(tid, k, ci, pid)] for ci in range(len(cand_list)) for pid in prev_comm)
                    >= gp.quicksum(y[(tid, k, ci)] for ci in range(len(cand_list))),
                    name=f"handover_fixed_to_new[{tid},{k}]"
                )
            else:
                # 新規同士の引き継ぎは後段で z 変数を使って実装する。
                continue

    # 上の「新規同士の引き継ぎ」を、k単位の共通変数z[tid,k,p]で実装する（爆発回避）
    z: Dict[Tuple[str, int, str], gp.Var] = {}
    for tid, cand_list in pre_cand.items():
        for k in range(2, Kmax_by_team[tid] + 1):
            for pid in persons:
                z[(tid, k, pid)] = m.addVar(vtype=GRB.BINARY, name=f"z[{tid},{k},{pid}]")
    m.update()

    for tid, cand_list in pre_cand.items():
        for k in range(2, Kmax_by_team[tid] + 1):
            # z <= commissioner_in_k, z <= commissioner_in_(k-1)
            # commissioner_in_k = Σ_ci x[tid,k,ci,p]
            for pid in persons:
                m.addConstr(
                    z[(tid, k, pid)] <= gp.quicksum(x[(tid, k, ci, pid)] for ci in range(len(cand_list))),
                    name=f"z_le_cur[{tid},{k},{pid}]"
                )
                m.addConstr(
                    z[(tid, k, pid)] <= gp.quicksum(x[(tid, k-1, ci, pid)] for ci in range(len(cand_list))),
                    name=f"z_le_prev[{tid},{k},{pid}]"
                )

            # 共通>=1（会議kが置かれたときだけ）
            placed_k = gp.quicksum(y[(tid, k, ci)] for ci in range(len(cand_list)))
            m.addConstr(
                gp.quicksum(z[(tid, k, pid)] for pid in persons) >= placed_k,
                name=f"handover_new_to_new[{tid},{k}]"
            )

    # (9) kの時系列整合（第k回が第k-1回より後）
    # 候補が時系列順に並んでいる前提で、idxの逆転を禁止する：
    # y[k,ci_cur] + y[k-1,ci_prev] <= 1 for ci_prev >= ci_cur
    for tid, cand_list in pre_cand.items():
        n = len(cand_list)
        for k in range(2, Kmax_by_team[tid] + 1):
            for ci_cur in range(n):
                for ci_prev in range(ci_cur, n):
                    m.addConstr(
                        y[(tid, k, ci_cur)] + y[(tid, k-1, ci_prev)] <= 1,
                        name=f"order[{tid},{k},{ci_cur},{ci_prev}]"
                    )

    # --- ソフト制約（目的関数） ---
    # 1) availabilityペナルティ（参加者×4セルを足し込む）
    for tid, cand_list in pre_cand.items():
        leader_pid = data.teams[tid].leader_pid
        for k in range(1, Kmax_by_team[tid] + 1):
            for ci, c in enumerate(cand_list):
                # リーダー分
                pen_leader = 0
                slots = data.avail.get(leader_pid, {}).get(c.day, {})
                for sl in grid.meeting_slots_covered(c.start_slot, cfg.meeting_slots):
                    pen_leader += _availability_penalty(slots.get(sl, 4), cfg)
                availability_penalty_terms.append(pen_leader * y[(tid, k, ci)])

                # 許可委員分（誰が入るかで変わる）
                for pid in commissioners:
                    pen_p = 0
                    slots_p = data.avail.get(pid, {}).get(c.day, {})
                    for sl in grid.meeting_slots_covered(c.start_slot, cfg.meeting_slots):
                        pen_p += _availability_penalty(slots_p.get(sl, 4), cfg)
                    if pen_p != 0:
                        availability_penalty_terms.append(pen_p * x[(tid, k, ci, pid)])

                # 締切当日ペナルティ（1日バッファ近似）
                if c.day == data.teams[tid].deadline:
                    deadline_day_penalty_terms.append(y[(tid, k, ci)])

    # 2) 中1日以上（近似：連日配置を嫌う）
    # チーム内で、日付が連続する候補を両方選んだらペナルティ
    # （厳密な「中1日以上」を全ペアでやると重いので、連日だけ抑える）
    # 日本語コメント: 連日配置のペナルティは線形化が必要なので、
    # placed_day を導入して「連日ならペナルティ」を表現する。

    placed_day: Dict[Tuple[str, date], gp.Var] = {}
    for tid, cand_list in pre_cand.items():
        days = sorted({c.day for c in cand_list})
        for d in days:
            placed_day[(tid, d)] = m.addVar(vtype=GRB.BINARY, name=f"placed_day[{tid},{d}]")
    m.update()

    for tid, cand_list in pre_cand.items():
        # placed_day >= 各候補y
        day_to_cis: Dict[date, List[int]] = {}
        for ci, c in enumerate(cand_list):
            day_to_cis.setdefault(c.day, []).append(ci)
        for d, cis in day_to_cis.items():
            max_per_day = max(1, Kmax_by_team[tid])
            m.addConstr(
                placed_day[(tid, d)] >= gp.quicksum(y[(tid, k, ci)]
                                                    for k in range(1, Kmax_by_team[tid] + 1)
                                                    for ci in cis) / float(max_per_day),
                name=f"placed_day_lb[{tid},{d}]"
            )
            m.addConstr(
                placed_day[(tid, d)] <= gp.quicksum(y[(tid, k, ci)]
                                                    for k in range(1, Kmax_by_team[tid] + 1)
                                                    for ci in cis),
                name=f"placed_day_ub[{tid},{d}]"
            )

        # 連日ペナルティ：placed_day(d) + placed_day(d+1) - 1 <= v, v>=0 を最小化
        days = sorted(day_to_cis.keys())
        day_set = set(days)
        for d in days:
            d2 = d.fromordinal(d.toordinal() + 1)
            if d2 not in day_set:
                continue
            v = m.addVar(vtype=GRB.CONTINUOUS, lb=0.0, name=f"v_consecutive[{tid},{d}]")
            m.addConstr(v >= placed_day[(tid, d)] + placed_day[(tid, d2)] - 1, name=f"v_consecutive_c[{tid},{d}]")
            gap_penalty_terms.append(v)

    # 3) 通常+1回バッファ（加点）：base_required+1 を満たしたら reward
    # バイナリ buf_ok を導入
    buf_ok: Dict[str, gp.Var] = {}
    for tid, t in data.teams.items():
        buf_ok[tid] = m.addVar(vtype=GRB.BINARY, name=f"buf_ok[{tid}]")
    m.update()

    for tid, t in data.teams.items():
        cand_list = pre_cand.get(tid, [])
        total_meetings = fixed_count_by_team[tid] + gp.quicksum(
            y[(tid, k, ci)]
            for k in range(1, Kmax_by_team[tid] + 1)
            for ci in range(len(cand_list))
        )
        # total >= base+1 なら buf_ok=1（Mで線形化）
        # Mは最大会議数の上限で十分
        M = fixed_count_by_team[tid] + Kmax_by_team[tid]
        m.addConstr(total_meetings >= (t.base_required + 1) - M * (1 - buf_ok[tid]), name=f"buf_ok1[{tid}]")
        m.addConstr(total_meetings <= (t.base_required) + M * buf_ok[tid], name=f"buf_ok2[{tid}]")
        # 追加がある場合でも「通常+1」のみ評価（仕様に沿う）
        if t.base_required > 0:
            normal_plus_one_reward_terms.append(buf_ok[tid])

    # 4) 負担均一化：Wmax最小化
    # w[p] = 固定出席 + 新規出席（leader + commissioner）
    for pid in persons:
        expr = fixed_attend_count[pid]

        for tid, cand_list in pre_cand.items():
            leader_pid = data.teams[tid].leader_pid
            for k in range(1, Kmax_by_team[tid] + 1):
                for ci in range(len(cand_list)):
                    if pid == leader_pid:
                        expr += y[(tid, k, ci)]
                    expr += x[(tid, k, ci, pid)]

        m.addConstr(w[pid] == expr, name=f"w_def[{pid}]")
        m.addConstr(Wmax >= w[pid], name=f"Wmax_ge[{pid}]")

    # --- 目的関数 ---
    obj = gp.LinExpr()

    # availability penalty
    obj += cfg.weights.w_availability * gp.quicksum(availability_penalty_terms)

    # deadline day penalty（締切当日を嫌う）＝「1日バッファ」近似
    obj += cfg.weights.w_finish_buffer * gp.quicksum(deadline_day_penalty_terms)

    # consecutive day penalty
    obj += cfg.weights.w_gap_rule * gp.quicksum(gap_penalty_terms)

    # rewardは「最小化」なのでマイナスで入れる
    obj += -cfg.weights.w_normal_plus_one * gp.quicksum(normal_plus_one_reward_terms)

    # load balance
    obj += cfg.weights.w_load_balance * Wmax

    m.setObjective(obj, GRB.MINIMIZE)

    # --- solve ---
    m.optimize()

    status = m.Status
    if status in (GRB.INFEASIBLE, GRB.INF_OR_UNBD):
        # IIS診断（任意）
        iis_text = None
        try:
            m.computeIIS()
            bad = []
            for c in m.getConstrs():
                if c.IISConstr:
                    bad.append(c.ConstrName)
            iis_text = "IIS（矛盾に関与する制約）:\n" + "\n".join(bad[:200])
        except Exception:
            iis_text = None
        return SolveResult(feasible=False, status="INFEASIBLE", meetings=[], iis_summary=iis_text)

    if status not in (GRB.OPTIMAL, GRB.TIME_LIMIT, GRB.SUBOPTIMAL):
        return SolveResult(feasible=False, status=f"STATUS_{status}", meetings=[], iis_summary=None)

    # --- 解の復元（新規会議のみ） ---
    meetings: List[SolutionMeeting] = []
    for tid, cand_list in pre_cand.items():
        fixed_list = fixed_sorted.get(tid, [])
        base_no = len(fixed_list)

        # 新規kごとに配置候補を探す
        for k in range(1, Kmax_by_team[tid] + 1):
            chosen_ci = None
            for ci in range(len(cand_list)):
                if y[(tid, k, ci)].X > 0.5:
                    chosen_ci = ci
                    break
            if chosen_ci is None:
                continue

            c = cand_list[chosen_ci]
            leader_pid = data.teams[tid].leader_pid

            comms = []
            for pid in persons:
                if x[(tid, k, chosen_ci, pid)].X > 0.5:
                    comms.append(pid)
            comms = comms[:4]
            if len(comms) != 4:
                # 数値誤差・復元の保険
                continue

            meeting_no = base_no + k

            # 引き継ぎ担当者（前回と共通の誰か1名）
            handover_pid = None
            if meeting_no >= 2:
                # 前回が固定か新規かで共通集合を取り、最初の1名を採用
                prev_comm_set: Set[str] = set()
                if meeting_no - 1 <= base_no and base_no > 0:
                    prev_comm_set = set(fixed_list[meeting_no - 2].commissioner_pids)
                else:
                    # 前回が新規（meeting_no-1 = base_no + (k-1)）
                    prev_k = k - 1
                    prev_ci = None
                    for ci2 in range(len(cand_list)):
                        if y[(tid, prev_k, ci2)].X > 0.5:
                            prev_ci = ci2
                            break
                    if prev_ci is not None:
                        prev_comm_set = {pid for pid in persons if x[(tid, prev_k, prev_ci, pid)].X > 0.5}
                common = [pid for pid in comms if pid in prev_comm_set]
                if common:
                    handover_pid = common[0]

            meetings.append(SolutionMeeting(
                team_tid=tid,
                day=c.day,
                start_slot=c.start_slot,
                leader_pid=leader_pid,
                commissioner_pids=(comms[0], comms[1], comms[2], comms[3]),
                meeting_no=meeting_no,
                handover_person_pid=handover_pid
            ))

    return SolveResult(feasible=True, status=("OPTIMAL" if status == GRB.OPTIMAL else "FEASIBLE"), meetings=meetings)
