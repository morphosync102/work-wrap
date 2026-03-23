# work-wrap

業務終了後の勤怠・工数記録を Markdown ファイルで管理し、各種申請システムへ自動転記するツールです。

## 対応システム

| システム | 用途 | ステータス |
|----------|------|----------|
| **CATW** | プロジェクト案件の工数入力 | ✅ MD管理・Excel転記実装済み / Selenium: 統合待ち |
| **PSA** | 案件外工数・祝日勤怠提出 | MD管理済み（フォーマット確認後に詳細化）/ 転記: TODO |
| **Workday** | 日次の出退勤時刻提出 | MD管理済み / 転記: TODO |
| **Concur** | 交通費申請 | ✅ MD管理・Excel転記実装済み / Selenium: 統合待ち |

## ディレクトリ構成

```
work-wrap/
├── config.example.yaml           # 設定テンプレート（Git管理対象）
├── config.yaml                   # 実際の設定（.gitignore で除外・各自作成）
├── pyproject.toml                # Python依存関係
├── .gitignore
├── scripts/
│   ├── generate_month.py         # ✅ 月次MDテンプレート生成（祝日自動判定）
│   ├── transfer_catw.py          # ✅ MD → CATW Excel 転記
│   ├── transfer_psa.py           # 🔲 MD → PSA 転記（フォーマット確認後）
│   ├── transfer_concur.py        # ✅ MD → Concur Excel 転記
│   ├── catw_selenium.py          # 🔲 CATW Excel → ブラウザ Selenium（既存統合待ち）
│   └── concur_selenium.py        # 🔲 Concur Excel → ブラウザ Selenium（既存統合待ち）
└── 勤怠/
    └── YYYY/
        └── MM月.md               # 月次勤怠データ（実データ）
```

## セットアップ

```bash
# 1. 仮想環境の作成と依存ライブラリのインストール
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e .

# 2. 設定ファイルの作成
cp config.example.yaml config.yaml

# 3. config.yaml を編集して自分の環境に合わせる
#    - プロジェクト名・WBSコード
#    - 各ExcelファイルのパスとExcel転記先列
```

## 使い方

### 月次テンプレートの生成

```bash
python scripts/generate_month.py 2026 4    # 特定の年月
python scripts/generate_month.py           # 引数なしで今月
```

`勤怠/2026/4月.md` が生成されます。あとはそのファイルを直接編集して記録します。

### 生成されるMDファイルの構造

```
# 2026年4月 勤怠

## サマリー          ← 月末に手動/集計スクリプトで更新

## CATW（案件工数）  ← プロジェクト別工数（config.yamlのprojectsから列が生成される）
## PSA（案件外工数） ← 案件外工数・祝日など
## Workday（出退勤） ← 出勤09:00固定、退勤・総時間・残業を手入力
## Concur（交通費）  ← 交通費の行を都度追加
```

#### 各テーブルの役割と記入ルール

**CATW**
- プロジェクト別の工数（時間）を記入
- `合計` 列に当日のCATW合計を記入
- 残業分もCATWに入れてよい

**PSA**
- CATW以外の工数（内勤・教育・祝日対応など）を記入
- CATW合計 + PSA合計 = 総勤務時間
- 詳細フォーマットは連携後に更新

**Workday**
- 出勤: `09:00`（テンプレート時点で自動入力済み）
- 退勤: `09:00 + コアタイム(7.5h) + 残業 + 休憩(1h)` で計算
  - 残業0h → 退勤 `17:30`
  - 残業3h → 退勤 `20:30`
- 総時間・残業を手入力（または集計スクリプトで自動化予定）

**Concur**
- 交通費が発生した日の行を手動で追加
- 列: 日付 / 交通手段 / 金額(円) / コメント / Business Purpose

### CATW Excel への転記（transfer_catw.py）

MDファイルの `## CATW（案件工数）` テーブルを読み込み、CATWの提出用 Excel ファイルに転記します。

```bash
python scripts/transfer_catw.py 2026 3   # 2026年3月分
python scripts/transfer_catw.py          # 引数なしで今月
```

**前提条件:**
- `config.yaml` の `catw.excel.path` に Excel ファイルのパスが設定されていること
- `config.yaml` の `projects` に対象案件の WBS・description・aa_type が定義されていること

**動作:**
1. `勤怠/YYYY/MM月.md` の CATW テーブルをパース
2. 書き込み内容をプレビュー表示（案件名・日数・合計工数）
3. `転記を実行しますか？ [y/N]` の確認後に Excel へ書き込み

**Excel の書き込み先（シート構造）:**
- `C4`=年, `C5`=月
- 週ブロック（Week1〜6）の各行に WBS / Description / AA Type / Memo / 日別工数を書き込み
- G列（週合計）の数式は上書きしない

---

### Concur Excel への転記（transfer_concur.py）

MDファイルの `## Concur（交通費）` テーブルを読み込み、Concur 提出用 Excel ファイルに転記します。

```bash
python scripts/transfer_concur.py 2026 3   # 2026年3月分
python scripts/transfer_concur.py          # 引数なしで今月
```

**前提条件:**
- `config.yaml` の `concur.excel.path` に Excel ファイルのパスが設定されていること

**動作:**
1. `勤怠/YYYY/MM月.md` の Concur テーブルをパース
2. 書き込み内容をプレビュー表示（日付・交通手段・金額・件数合計）
3. `転記を実行しますか？ [y/N]` の確認後に Excel へ書き込み

**MDテーブルの列と Excel 列の対応:**

| MD列 | Excel列 | 内容 |
|------|---------|------|
| 日付 | A | 日付（`YYYY-MM-DD` / `YYYY/MM/DD` / `M/D` 形式を自動解析）|
| 交通手段 | B | 交通手段 |
| 金額 | C | 金額（円、カンマ・¥記号は自動除去）|
| コメント | D | コメント |
| Business Purpose | E | Business Purpose |

---

### PSA 転記（transfer_psa.py）

**現在未実装。** PSA のフォーマット確認後に実装予定。

```bash
python scripts/transfer_psa.py   # [TODO] 現時点では即座に終了します
```

---

### CATW 一括実行（run_catw.py）

Excel 転記 → OpenCATW マクロ → Web 自動入力 を一本のスクリプトで順次実行できます。

```bash
python scripts/run_catw.py 2026 3   # 2026年3月分
python scripts/run_catw.py          # 引数なしで今月
```

**実行ステップ:**
1. `transfer_catw.py` を呼び出し（確認プロンプットあり）
2. CATW Web 自動入力を続けるか確認
3. win32com で Excel を開き `OpenCATW` マクロを実行（Edge がデバッグモードで起動）
4. Edge の CDP ポート（デフォルト 9222）が応答するまで待機
5. `catw_selenium.py` を呼び出し

> **注意**: Step 3〜5 は `pywin32` が必要なため **Windows 環境のみ** 動作します。

**必要な config.yaml 設定:**
```yaml
catw:
  excel:
    path: "/path/to/CATW.xlsm"
    macro_name: "OpenCATW"   # VBA マクロ名（デフォルト: OpenCATW）
    cdp_port: 9222           # Edge デバッグポート（デフォルト: 9222）
```

---

### 運用フロー（完成後）

```
1. 月初: python scripts/generate_month.py   → 月次MDを生成
2. 毎日: MDファイルに当日分を記入
3. 月末: python scripts/run_catw.py         → CATW Excel転記 + Web自動入力（一括）
          python scripts/transfer_concur.py  → Concur Excel に転記
          python scripts/concur_selenium.py  → Concur Web に自動入力
          python scripts/transfer_psa.py     → PSA に転記（フォーマット確認後）
```

## 設定ファイル（config.yaml）の主要項目

```yaml
projects:               # CATWに登録している案件一覧
  - name: "ProjectA"
    wbs: "A-001"        # WBSコード（案件終了時はここから削除）
    catw_column: "B"

workday:
  start_time: "09:00"
  core_hours: 7.5       # コアタイム
  break_hours: 1.0      # 休憩時間

catw:
  excel:
    path: "/path/to/catw.xlsx"

concur:
  excel:
    path: "/path/to/concur.xlsx"
```

## 依存ライブラリ

| ライブラリ | 用途 |
|---|---|
| `PyYAML` | `config.yaml` の読み込み |
| `holidays` | 日本の祝日自動判定 |
| `openpyxl` | Excel転記（転記スクリプト実装後に使用） |
| `selenium` | ブラウザ自動操作 |
| `webdriver-manager` | ChromeDriverの自動管理 |

## 実装状況

### スクリプト一覧

| スクリプト | 状態 | 実行環境 | 内容 |
|---|---|---|---|
| `generate_month.py` | ✅ 実装済み | Mac/Win | 月次 MD テンプレート生成 |
| `transfer_catw.py` | ✅ 実装済み | Mac/Win | MD → CATW Excel 転記 |
| `transfer_concur.py` | ✅ 実装済み | Mac/Win | MD → Concur Excel 転記 |
| `run_catw.py` | ✅ 実装済み・**未通しテスト** | **Windows のみ** | CATW 一括実行オーケストレーター |
| `catw_selenium.py` | ✅ 実装済み・**未通しテスト** | **Windows のみ** | Playwright で CATW Web 自動入力 |
| `concur_selenium.py` | ✅ 実装済み・**未通しテスト** | **Windows のみ** | Selenium で Concur Web 自動入力 |
| `transfer_psa.py` | 🔲 未実装 | - | PSA フォーマット確認後に実装 |

> `catw_selenium.py` / `concur_selenium.py` は `.gitignore` 対象（社内システム用）。

---

## Windows 環境での動作確認チェックリスト

`run_catw.py` を通しで動かす前に以下を確認すること。

### 1. config.yaml の設定確認

```yaml
catw:
  excel:
    path: "/path/to/FY26_03月_CATW入力マクロ.xlsm"  # 実際のパスに変更
    macro_name: "OpenCATW"   # Excel VBA マクロ名（実際のマクロ名と一致しているか）
    cdp_port: 9222           # Edge デバッグポート（変更している場合は修正）
```

### 2. Step 別確認ポイント

| Step | コマンド / 操作 | 確認すること |
|---|---|---|
| Step 1 | `python scripts/transfer_catw.py 2026 3` | 確認プロンプトが出る / Excel に正しく書き込まれる / N で終了する |
| Step 2 | `python scripts/transfer_concur.py 2026 3` | Concur シートへの書き込みが正しい |
| Step 3 | `python scripts/run_catw.py 2026 3` | Step1 の確認後 → Step2 の Web 入力確認が出る |
| Step 4 | 上記 y 入力後 | win32com で Excel が開く / OpenCATW マクロが動く / Edge がデバッグモードで起動する |
| Step 5 | 上記継続 | CDP ポート 9222 に接続できる / catw_selenium.py が年月を正しく読む |

### 3. 残タスク

| 優先度 | タスク |
|---|---|
| 高 | Windows で `run_catw.py` を通しテスト（特に Step3 の `excel.Run()` マクロ名フォーマット） |
| 高 | `config.yaml` の WBS コードを実際の案件に合わせる |
| 中 | `run_concur.py` を作成（run_catw.py と同パターンで Concur を一括実行） |
| 低 | PSA フォーマット確認 → `transfer_psa.py` 実装 |
| 低 | Workday 退勤時刻の自動計算スクリプト |

## AI向け補足

このリポジトリに対して作業する際の注意点:

- **`config.yaml` は `.gitignore` 対象**。存在しない場合は `config.example.yaml` をコピーして作成させること
- **プロジェクト一覧は `config.yaml` の `projects` セクション**で管理。スクリプト内にハードコードしない
- **コアタイムは7.5時間**（8時間ではない）。退勤計算: `09:00 + 7.5h + 残業h + 1h(休憩)`
- **MDテーブルのパース規則**:
  - 土日・祝日行: 日付が `**太字**`、工数セルが ` - `
  - 祝日の日付フォーマット: `M/D (祝)祝日名`
  - 平日の日付フォーマット: `M/D (曜日)`
- **transfer_catw.py・transfer_concur.py は実装済み**。transfer_psa.py のみ未実装（PSA フォーマット確認後に実装）
- **Seleniumスクリプトは既存コードの統合待ち**。プレースホルダーのみ存在する
