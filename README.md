# 時間割自動生成

## 目的
- 教員の担当・不可コマ・技能科目入力をもとに、時間割を自動生成する。
- 出力Excelはテンプレートコピーではなく、構造（シート/セル）をコードで生成する。

## 構成
- `timetable_gui.py` : 入力〜自動割り当て〜Excel出力までを行うGUI。
- `scheduler_engine.py` : ルールベースの探索エンジン（TT割当・診断/提案付き）。
- `scheduler_engine_cp.py` : CP-SAT（OR-Tools）エンジン。GUIの自動割り当てで使用。
- `timetable_tool.py` : CLIツール（JSON/Excel生成、PDF抽出、簡易CSP解法）。
- `requirements.txt` : CLIで必要な依存（`openpyxl`, `pypdf`, `orjson`）。
- `../R6.json`, `../R7.json` : 実データ例（リポジトリ直下）。
- `POLICY.md` : 仕様メモ。

## セットアップ
```bash
cd program
pip install -r requirements.txt
```

GUIの自動割り当て（CP-SAT）を使う場合は追加で以下が必要です。
```bash
pip install ttkbootstrap ortools
```

## GUIでの運用フロー（推奨）
1. `python timetable_gui.py`
2. 基本設定タブで以下を入力し、「基本設定反映」を実行。
   - 出力Excel / 教科一覧 / クラス一覧 / 曜日時限
3. 教員タブ
   - 教員名・担当教科を入力
   - 「教員名を確定して不可コマ欄を生成」を押す
   - 不可コマを✕で入力、担当クラスの時間数を入力
   - TTはセルをダブルクリックで薄青に変更（TT希望）
4. 技能科目入力タブ
   - 「1年音美」「1年総合」を手入力
5. その他科目割り振りタブ
   - 「技能科目入力タブをコピー」→「教員設定を基に他教科割り当て」
6. 確認タブでサマリを確認し、「Excelに出力」

## CLIでの運用フロー
### 設定テンプレート生成
```bash
python timetable_tool.py init --config config.json
```

### PDF抽出（入力補助）
```bash
python timetable_tool.py extract-pdf --pdf ../YOURPDF.pdf --out ./output/pdf_text.txt
```

### 自動生成
```bash
python timetable_tool.py solve --config ../R6.json
```

注意:
- `timetable_tool.py` は内蔵のCSPソルバを使用し、`solver.engine` は参照しません。
- GUIの自動割り当て結果を使う場合は、`fixed_assignments` が全コマ埋まるため `weekly_requirements` が空でも生成できます。

## 設定JSONの要点
- `subjects`, `classes`, `day_periods`
- `teachers` : `name`, `subjects`, `class_assignments`(hours/tt), `unavailable_slots`
- `class_subject_teacher` : CLIソルバ用の教科担当（GUI保存時に自動生成）
- `scenarios` : `id`, `target_block`(upper/lower), `weekly_requirements`, `fixed_assignments`, `manual_skill_assignments`
- `solver` : GUIの自動割り当てで使用（`engine: cp_sat` 推奨）

## 出力Excelの構造
- `時間割（略図）` : 空シート（現状未使用）
- `完成` : 各シナリオ×クラス×曜日時限の一覧
- `技能科目` : 1つ目のシナリオから技能科目のみ抽出
- `予備` : `upper`=行3-13 / `lower`=行17-27 に書き込み

## 制約ルール（実装済み）
- 同日同教科の重複禁止
- 学年内同時刻の同教科重複禁止（除外教科あり）
- 教員の同時刻重複禁止
- 教員不可コマの考慮
- 固定入力の厳守
- TT割当（CP-SAT/ルールエンジン）

## 既知のポイント
- GUIの自動割り当ては `scheduler_engine_cp.py` を使用し、`ortools` が必須です。
- CLIのCSPソルバは `weekly_requirements` を前提とするため、未設定の場合は固定入力で全コマを埋めてください。
