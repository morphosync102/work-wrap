# work-wrap

業務終了後の勤怠・工数記録を Markdown で管理し、各種申請システムへ自動転記するツールです。

## ディレクトリ構成

```
work-wrap/
├── config.example.yaml     # 設定テンプレート（コピーして config.yaml を作成）
├── config.yaml             # 実際の設定（.gitignore 対象・各自作成）
├── pyproject.toml
├── scripts/
│   ├── generate_month.py   # 月次 MD テンプレート生成
│   ├── transfer_catw.py    # MD → CATW Excel 転記
│   ├── transfer_concur.py  # MD → Concur Excel 転記
│   ├── transfer_psa.py     # MD → PSA 転記（未実装）
│   ├── run_catw.py         # CATW 一括実行（転記 + Web 自動入力）
│   ├── catw_selenium.py    # CATW Web 自動入力 ※.gitignore 対象
│   └── concur_selenium.py  # Concur Web 自動入力 ※.gitignore 対象
├── tests/
│   ├── conftest.py
│   └── test_catw_logic.py
└── 勤怠/
    └── YYYY/
        └── MM月.md
```

---

## セットアップ

```bash
# 1. 依存ライブラリのインストール
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e .

# 2. 設定ファイルを作成
cp config.example.yaml config.yaml
# → config.yaml を編集してパス・WBS コード・案件名を設定する
```

---

## 日常の使い方

### 月初: テンプレート生成

```bash
python scripts/generate_month.py 2026 4   # 特定の年月
python scripts/generate_month.py          # 今月
```

`勤怠/2026/4月.md` が生成されます。あとはそのファイルを日々記入します。

### 毎日: MD ファイルを記入

生成された MD ファイルには以下のテーブルがあります：

| テーブル | 内容 | 記入ルール |
|---|---|---|
| **CATW（案件工数）** | プロジェクト別工数（時間） | `合計` 列も記入。残業分も CATW に入れてよい |
| **PSA（案件外工数）** | 内勤・教育・祝日など | CATW + PSA = 総勤務時間 |
| **Workday（出退勤）** | 出勤 / 退勤 / 総時間 / 残業 | 出勤は `09:00` 固定。退勤 = `09:00 + 7.5h + 残業h + 1h（休憩）` |
| **Concur（交通費）** | 日付 / 交通手段 / 金額 / コメント / Business Purpose | 交通費が発生した日だけ行を追加 |

### 月末: 各システムへ転記・入力

**CATW（推奨: 一括実行）**
```bash
python scripts/run_catw.py 2026 3   # 転記 → OpenCATW マクロ → Web 自動入力を一括実行
```

**Concur**
```bash
python scripts/transfer_concur.py 2026 3   # Excel 転記
# → concur_selenium.py を別途実行（run_concur.py は未作成）
```

---

## スクリプトリファレンス

### generate_month.py

月次 MD テンプレートを生成します。日本の祝日は自動判定されます。

```bash
python scripts/generate_month.py [年] [月]
```

### transfer_catw.py

MD の `## CATW（案件工数）` テーブルを CATW 提出用 Excel に転記します。

```bash
python scripts/transfer_catw.py [年] [月]
```

- プレビューを表示 → `転記を実行しますか？ [y/N]` で確認後に書き込み
- `config.yaml` の `projects` から WBS / Description / AA Type を解決
- Excel シート構造: `C4`=年, `C5`=月, Week1〜6 の各ブロックに日別工数を書き込み

### transfer_concur.py

MD の `## Concur（交通費）` テーブルを Concur 提出用 Excel に転記します。

```bash
python scripts/transfer_concur.py [年] [月]
```

| MD 列 | Excel 列 | 備考 |
|---|---|---|
| 日付 | A | `YYYY-MM-DD` / `YYYY/MM/DD` / `M/D` を自動解析 |
| 交通手段 | B | |
| 金額 | C | カンマ・¥記号を自動除去 |
| コメント | D | |
| Business Purpose | E | |

### run_catw.py（Windows のみ）

Excel 転記 → OpenCATW マクロ実行 → Web 自動入力 を一括で実行します。

```bash
python scripts/run_catw.py [年] [月]
```

| Step | 処理 |
|---|---|
| 1 | `transfer_catw.py` を実行（確認プロンプトあり） |
| 2 | 「CATW Web への自動入力を実行しますか？」確認 |
| 3 | win32com で Excel を開き `OpenCATW` マクロを実行（Edge がデバッグモードで起動） |
| 4 | Edge の CDP ポート（デフォルト: 9222）が応答するまで待機 |
| 5 | `catw_selenium.py` を実行 |

> Step 3〜5 は `pywin32` が必要なため **Windows 環境のみ** 動作します。

---

## 設定（config.yaml）

```yaml
projects:                          # CATW に登録している案件一覧
  - name: "ProjectA"               # MD ファイルの列名
    wbs: "JP3-XXXXX.XX.XX.XX"     # WBS コード
    description: "412 Implementation"
    aa_type: "412 Implementation"  # Attend/Absence Type（config.example.yaml の一覧から選択）
    memo: "プロジェクトA"

workday:
  start_time: "09:00"
  core_hours: 7.5                  # コアタイム（7.5h = 定時退勤 17:30）
  break_hours: 1.0

catw:
  excel:
    path: "/path/to/FY26_03月_CATW入力マクロ.xlsm"
    sheet_name: "CATW"
    macro_name: "OpenCATW"         # VBA マクロ名（run_catw.py が使用）
    cdp_port: 9222                 # Edge デバッグポート（run_catw.py が使用）

concur:
  excel:
    path: "/path/to/FY26_03月_CATW入力マクロ.xlsm"
    sheet_name: "Concur"
    header_row: 1
    data_start_row: 2
```

Attend/Absence Type の全一覧は `config.example.yaml` を参照してください。

---

## 依存ライブラリ

<!-- AUTO-GENERATED from pyproject.toml -->
| ライブラリ | 用途 | 環境 |
|---|---|---|
| `PyYAML` | config.yaml の読み込み | 全環境 |
| `holidays` | 日本の祝日自動判定 | 全環境 |
| `openpyxl` | Excel 転記 | 全環境 |
| `selenium` | Concur Web 自動入力 | 全環境 |
| `webdriver-manager` | ChromeDriver 自動管理 | 全環境 |
| `pywin32` | Excel VBA マクロ呼び出し | **Windows のみ** |

開発用（`pip install -e ".[dev]"`）:

| ライブラリ | 用途 |
|---|---|
| `pytest` | ユニットテスト |
| `pytest-mock` | モック |
<!-- END AUTO-GENERATED -->

---

## 実装状況と残タスク

| スクリプト | 状態 | 環境 | 内容 |
|---|---|---|---|
| `generate_month.py` | ✅ | Mac/Win | 月次 MD テンプレート生成 |
| `transfer_catw.py` | ✅ | Mac/Win | MD → CATW Excel 転記 |
| `transfer_concur.py` | ✅ | Mac/Win | MD → Concur Excel 転記 |
| `run_catw.py` | ✅ 未通しテスト | **Win** | CATW 一括実行オーケストレーター |
| `catw_selenium.py` | ✅ 未通しテスト | **Win** | Playwright で CATW Web 自動入力 |
| `concur_selenium.py` | ✅ 未通しテスト | **Win** | Selenium で Concur Web 自動入力 |
| `transfer_psa.py` | 🔲 未実装 | - | PSA フォーマット確認後に実装 |
| `run_concur.py` | 🔲 未作成 | **Win** | Concur 一括実行（run_catw.py と同パターン） |

### Windows 通しテストのチェックリスト

`run_catw.py` を初めて動かす前に確認すること：

- [ ] `config.yaml` の `catw.excel.path` を実際のファイルパスに設定した
- [ ] `macro_name` が実際の VBA マクロ名と一致している
- [ ] `cdp_port` が Edge のデバッグポートと一致している（デフォルト: 9222）
- [ ] `transfer_catw.py 2026 3` 単体で Excel 書き込みが正しく動く
- [ ] `run_catw.py 2026 3` でStep1〜2 が通る（転記 → 確認プロンプト）
- [ ] Step 3 で OpenCATW マクロが動き、Edge がデバッグモードで起動する
- [ ] Step 4〜5 で catw_selenium.py が年月を正しく読み込んで Web 入力できる

---

## AI 向け補足

- **`config.yaml` は `.gitignore` 対象**。存在しない場合は `config.example.yaml` をコピーして作成させること
- **プロジェクト一覧は `config.yaml` の `projects` セクション**で管理。スクリプト内にハードコードしない
- **コアタイムは 7.5 時間**（8 時間ではない）。退勤 = `09:00 + 7.5h + 残業h + 1h（休憩）`
- **MD テーブルのパース規則**:
  - 土日・祝日行: 日付が `**太字**`、工数セルが ` - `
  - 祝日フォーマット: `M/D (祝)祝日名`
  - 平日フォーマット: `M/D (曜日)`
- **`catw_selenium.py` / `concur_selenium.py` は `.gitignore` 対象**（社内システム用）。プレースホルダーではなく本実装済みのスクリプトが別途存在する
- **`transfer_psa.py` のみ未実装**。PSA フォーマット確認後に実装すること
