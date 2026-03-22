# work-wrap

業務終了後の勤怠・工数記録を Markdown ファイルで管理し、既存の Excel ファイルへ転記するための自動化ツールです。

## 概要

- 月次の勤怠データを `勤怠/YYYY/MM月.md` に記録する
- 日本の祝日を自動判定してテンプレートを生成する
- プロジェクト別工数・残業時間・WBSコードなどを一覧表形式で管理する
- 記録完了後、Python スクリプトで既存の Excel ファイルへ転記する（実装予定）

## ディレクトリ構成

```
work-wrap/
├── config.example.yaml        # 設定テンプレート（Git管理対象）
├── config.yaml                # 実際の設定（.gitignore で除外・各自作成）
├── pyproject.toml             # Python依存関係定義
├── .gitignore
├── scripts/
│   ├── generate_month.py      # 月次MDテンプレート生成スクリプト
│   └── transfer_to_excel.py   # Excel転記スクリプト（TODO）
└── 勤怠/
    └── YYYY/
        └── MM月.md            # 月次勤怠データ（実データ）
```

## セットアップ

```bash
# 1. 仮想環境の作成と依存ライブラリのインストール
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install PyYAML holidays

# 2. 設定ファイルの作成
cp config.example.yaml config.yaml

# 3. config.yaml を編集して自分の環境に合わせる
#    - プロジェクト名・WBSコード
#    - Excel転記先ファイルパス（後で設定）
```

## 使い方

### 月次テンプレートの生成

```bash
# 特定の年月を指定
python scripts/generate_month.py 2026 4

# 引数なしで今月を生成
python scripts/generate_month.py
```

`勤怠/2026/4月.md` が生成されます。あとはそのファイルを直接編集して勤怠を記録します。

### 生成されるMDファイルの構造

```markdown
# 2026年4月 勤怠

## サマリー
| 項目 | 値 |
...（月末に手動または集計スクリプトで更新）

## 日次記録
| 日付         | ProjectA | ProjectB | 残業 | WBS element | Description | Attend/Absence Type | Memo |
|--------------|----------|----------|------|-------------|-------------|---------------------|------|
| 4/1 (水)     |          |          |      |             |             |                     |      |
| **4/4 (土)** |  -       |  -       |  -   |             |             |                     |      |
| **4/29 (祝)昭和の日** | - | -     |  -   |             |             |                     |      |
```

- **太字・ハイフン（ -）**: 土日・祝日（入力不要）
- **祝日**: 日本の祝日名を自動付与（`holidays` ライブラリ使用）
- **WBS element**: プロジェクトの一意識別コード（`config.yaml` で管理）
- **Attend/Absence Type**: 業務種別（選択肢は `config.yaml` で定義）

## 設定ファイル（config.yaml）

```yaml
projects:
  - name: "ProjectA"
    wbs: "A-001"         # WBSコード（案件終了時はここから削除する）
    excel_column: "B"    # Excel転記先列
  - name: "ProjectB"
    wbs: "B-002"
    excel_column: "C"

attend_absence_types:
  - "開発"
  - "設計"
  - "レビュー"
  - "打ち合わせ"
  - "社内業務"

regular_hours: 8.0       # 所定労働時間

excel:
  path: "/path/to/勤怠表2026.xlsx"
  sheet_name: "4月"
  date_column: "A"
  data_start_row: 2
  overtime_column: "E"
```

### WBSコードの管理

- 案件ごとに `wbs` フィールドを設定する
- 案件が終了したら `config.yaml` から該当プロジェクトを削除する
- 選択肢は今後随時追加予定

## 依存ライブラリ

| ライブラリ | 用途 |
|---|---|
| `PyYAML` | `config.yaml` の読み込み |
| `holidays` | 日本の祝日自動判定 |
| `openpyxl` | Excel転記（TODO実装時に追加予定） |

## TODO（今後の実装予定）

- [ ] `scripts/transfer_to_excel.py` — MDファイルから既存Excelへの転記スクリプト
- [ ] `Attend/Absence Type` の選択肢を確定して `config.yaml` に追加
- [ ] WBSコードの選択肢を実際の案件に合わせて更新
- [ ] サマリーセクションの自動集計スクリプト

## AI向け補足

このリポジトリに対して作業する際の注意点：

- **`config.yaml` は `.gitignore` 対象**。存在しない場合は `config.example.yaml` をコピーして作成させること
- **プロジェクト一覧は `config.yaml` の `projects` セクションで管理**。スクリプト内にハードコードしない
- **MDファイルのパース**は `|` 区切りのテーブル形式。日付列のフォーマットは `M/D (曜日)` または `M/D (祝)祝日名`
- **土日・祝日の行**は工数セルが ` - ` で埋まっている。転記スクリプトではスキップする
- **Excel転記スクリプト**はExcelのフォーマット確認後に実装予定。`config.yaml` の `excel` セクションを参照すること
