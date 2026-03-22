#!/usr/bin/env python3
"""
MD → Concur Excel 転記スクリプト

MDファイルの「## Concur（交通費）」テーブルを読み込み、
config.yaml で指定した Concur 提出用 Excel ファイルに転記します。

使い方:
  python scripts/transfer_concur.py 2026 3

TODO: ConcurのExcelフォーマット確認後に実装
"""

# TODO: 実装待ち（ConcurのExcelフォーマット確認後）
# 実装方針:
#   1. load_config() で config.yaml を読み込む
#   2. 勤怠/YYYY/MM月.md を開いて「## Concur」セクションのテーブルをパース
#   3. openpyxl で concur.excel.path を開く
#   4. 各行（日付/交通手段/金額/コメント/Business Purpose）を書き込む
#   5. 保存

import sys
print("[TODO] transfer_concur.py は未実装です。ConcurのExcelフォーマット確認後に実装します。")
sys.exit(1)
