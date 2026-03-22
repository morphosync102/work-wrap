# work-wrap

業務終了後の勤怠・工数記録を Markdown ファイルで管理し、各種申請システムへ自動転記するツールです。

## 対応システム

| システム | 用途 | ステータス |
|----------|------|----------|
| **CATW** | プロジェクト案件の工数入力 | MD管理済み / Excel転記・Selenium: TODO |
| **PSA** | 案件外工数・祝日勤怠提出 | MD管理済み（フォーマット確認後に詳細化）/ 転記: TODO |
| **Workday** | 日次の出退勤時刻提出 | MD管理済み / 転記: TODO |
| **Concur** | 交通費申請 | MD管理済み / Excel転記・Selenium: TODO |

## ディレクトリ構成

```
work-wrap/
├── config.example.yaml           # 設定テンプレート（Git管理対象）
├── config.yaml                   # 実際の設定（.gitignore で除外・各自作成）
├── pyproject.toml                # Python依存関係
├── .gitignore
├── scripts/
│   ├── generate_month.py         # ✅ 月次MDテンプレート生成（祝日自動判定）
│   ├── transfer_catw.py          # 🔲 MD → CATW Excel 転記（TODO）
│   ├── transfer_psa.py           # 🔲 MD → PSA 転記（フォーマット確認後）
│   ├── transfer_concur.py        # 🔲 MD → Concur Excel 転記（TODO）
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

### 運用フロー（完成後）

```
1. 月初: python scripts/generate_month.py   → 月次MDを生成
2. 毎日: MDファイルに当日分を記入
3. 月末: python scripts/transfer_catw.py    → CATW Excel に転記
          python scripts/catw_selenium.py    → CATW Web に自動入力
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

## TODO

- [ ] PSAのフォーマット確認 → `transfer_psa.py` 実装、PSAテーブル詳細化
- [ ] CATWのExcelフォーマット確認 → `transfer_catw.py` 実装
- [ ] ConcurのExcelフォーマット確認 → `transfer_concur.py` 実装
- [ ] 既存の CATW Selenium スクリプトを `catw_selenium.py` に統合
- [ ] 既存の Concur Selenium スクリプトを `concur_selenium.py` に統合
- [ ] Attend/Absence Type の選択肢を確定して `config.yaml` に追加
- [ ] WBSコードを実際の案件に合わせて更新
- [ ] Workday の退勤・総時間を自動計算するヘルパースクリプト

## AI向け補足

このリポジトリに対して作業する際の注意点:

- **`config.yaml` は `.gitignore` 対象**。存在しない場合は `config.example.yaml` をコピーして作成させること
- **プロジェクト一覧は `config.yaml` の `projects` セクション**で管理。スクリプト内にハードコードしない
- **コアタイムは7.5時間**（8時間ではない）。退勤計算: `09:00 + 7.5h + 残業h + 1h(休憩)`
- **MDテーブルのパース規則**:
  - 土日・祝日行: 日付が `**太字**`、工数セルが ` - `
  - 祝日の日付フォーマット: `M/D (祝)祝日名`
  - 平日の日付フォーマット: `M/D (曜日)`
- **転記スクリプトはすべてTODO**。実装時は各システムのExcelフォーマットを確認してから着手すること
- **Seleniumスクリプトは既存コードの統合待ち**。プレースホルダーのみ存在する
