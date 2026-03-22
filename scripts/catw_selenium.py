#!/usr/bin/env python3
"""
CATW Excel → ブラウザ 自動転記スクリプト（Selenium）

config.yaml で指定した CATW 提出用 Excel を読み込み、
Selenium で CATW の Web サイトに自動入力します。

使い方:
  python scripts/catw_selenium.py 2026 3

TODO: 既存の Selenium スクリプトをここに統合してください。
"""

# TODO: 既存スクリプトをここに移植・統合する
# 統合手順:
#   1. 既存のSeleniumスクリプトのコードをここにコピー
#   2. Excelパスを config.yaml の catw.excel.path から読み込むよう変更
#   3. 年月を sys.argv から受け取るよう変更
#
# 依存ライブラリ（pyproject.toml に追加が必要）:
#   selenium>=4.0
#   webdriver-manager>=4.0  # ChromeDriverの自動管理

import sys
print("[TODO] catw_selenium.py は未実装です。既存のSeleniumスクリプトをここに統合してください。")
sys.exit(1)
