# 登山サークル審議会議スケジューラ（MILP + Gurobi）

## まず最初に：ディレクトリ＆空ファイル作成コマンド
以下は「新規環境でコピー用に空ファイルとディレクトリを一括作成する」ためのコマンド例です。
（このリポジトリでは既に作成済みですが、再利用時のテンプレとして残しています。）

```bash
mkdir -p meeting_scheduler/{domain,io_layer,validation,preprocessing,optimization,reporting,gui} assets/{input,output}
touch meeting_scheduler/__init__.py \
  meeting_scheduler/config.py \
  meeting_scheduler/domain/{__init__.py,models.py,timegrid.py} \
  meeting_scheduler/io_layer/{__init__.py,paths.py,xlsx_reader.py} \
  meeting_scheduler/validation/{__init__.py,validator.py} \
  meeting_scheduler/preprocessing/{__init__.py,preprocess.py} \
  meeting_scheduler/optimization/{__init__.py,milp.py} \
  meeting_scheduler/reporting/{__init__.py,report.py,export_xlsx.py} \
  meeting_scheduler/gui/{__init__.py,app.py} \
  main_cli.py requirements.txt
```

## 目的と特徴（要約）
* 複数月のスケジュール（複数ファイル）を読み込み、締切までに必要回数の会議を自動生成します。
* 既存スケジュールは固定扱いで変更しません。
* 追加審議は「過去出力の result シート」を固定会議として読み込みます。
* 生成開始日時が現在より過去なら即エラーで停止します。
* 出力は全て名称ベース（IDは内部のみ）です。

## ディレクトリ構成（概要）
```
meeting_scheduler/
  config.py           # 設定（時間スロット・重みなど）
  domain/             # エンティティ定義
  io_layer/           # スプレッドシート読み込み
  preprocessing/      # 前処理（候補生成など）
  optimization/       # MILPモデル（Gurobi）
  reporting/          # 出力整形・xlsx出力
  validation/         # バリデーション
  gui/                # Streamlit GUI
main_cli.py           # CLI実行用
```

## スプレッドシート仕様（重要）
### 1. 人員マスタ＋スケジュール（Sheets①）
* 人員マスタは `master` シートに置きます。
  * 列: `A:人物ID`, `B:人物名`, `C:許可委員フラグ(0/1)`, `D:上級生フラグ(0/1)`
* 各人物のスケジュールは「人物名シート」で管理します。
  * 列: `C(09:00)` 〜 `AB(21:30)` が時間スロット
  * 行: `2 = 日1` 〜 `最終日+1 = 月末`（閏年もカレンダーに従う）
  * 値: `1/2/3 = 参加可`、`0/4 = 参加不可`（0は4として扱います）
  * 会議は2時間固定（連続4セル）なので、4連続が 1/2/3 であることが参加条件です。

### 2. 登山隊情報（Sheets②）
* シート名: `teams`
* 必須列: `tid, name, leader_pid, deadline, base_required`
* 任意列: `member_pids`（カンマ区切り）

### 3. 既存スケジュール（Sheets③）
* シート名: `fixed`
* 必須列: `team_name, meeting_date, start_time, leader_name, comm1, comm2, comm3, comm4`

### 4. 追加審議要求（Sheets④）
* シート名: `add`
* 必須列: `team_name, add_required`

### 5. 追加審議のための過去結果
* 過去出力ファイルの `result` シートを読み込み、固定会議として扱います。

## 実行方法（CLI）
```bash
python main_cli.py \
  --schedule data/schedule_2026-01.xlsx data/schedule_2026-02.xlsx \
  --teams data/teams.xlsx \
  --fixed data/fixed.xlsx \
  --prev data/result_prev.xlsx \
  --add data/add_request.xlsx \
  --generation_start "2026-01-20 10:30" \
  --out assets/output/result.xlsx
```

## 実行方法（GUI）
```bash
streamlit run meeting_scheduler/gui/app.py
```
* 入力パスは改行区切りで複数指定可能です。
* 生成開始日時が過去の場合は即エラーで停止します。

## 出力
* `result` シート: 会議一覧（名称ベース）
* `team_summary` シート: 登山隊別サマリ
* `person_summary` シート: 人物別負担サマリ

---

### 注意点（運用）
* ファイル名に `YYYY-MM` が含まれることを前提に、月を推定しています。
* 追加審議は「過去結果の result シート」を固定会議として読み込みます。
* 生成開始日時は現在より過去であれば最適化を実行しません。
