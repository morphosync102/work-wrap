#!/usr/bin/env python3
"""
MD → CATW Excel 転記スクリプト

MDファイルの「## CATW（案件工数）」テーブルを読み込み、
config.yaml で指定した CATW 提出用 Excel ファイルに転記します。

使い方:
  python scripts/transfer_catw.py 2026 3

TODO: CATWのExcelフォーマット確認後に実装
"""

# TODO: 実装待ち（CATWのExcelフォーマット確認後）
# 実装方針:
#   1. load_config() で config.yaml を読み込む
#   2. 勤怠/YYYY/MM月.md を開いて「## CATW」セクションのテーブルをパース
#   3. openpyxl で catw.excel.path を開く
#   4. 日付列を走査して今月の行を特定し、各プロジェクト列に工数を書き込む
#   5. 保存

import sys
print("[TODO] transfer_catw.py は未実装です。CATWのExcelフォーマット確認後に実装します。")
sys.exit(1)
