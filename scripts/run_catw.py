#!/usr/bin/env python3
"""
CATW 一括実行オーケストレーター

MD → Excel 転記 → OpenCATW マクロ実行 → catw_selenium.py を順次実行します。

使い方:
  python scripts/run_catw.py 2026 3   # 特定の年月
  python scripts/run_catw.py          # 引数なしで今月

実行フロー:
  Step 1: transfer_catw.py  — MD → Excel 転記（確認プロンプトあり）
  Step 2: 確認               — CATW Web 自動入力を続けるか確認
  Step 3: OpenCATW マクロ   — win32com で Excel を開き VBA マクロを実行
                               → Edge がデバッグモードで起動し CATW URL を開く
  Step 4: CDP 接続待機      — Edge の CDP ポートが応答するまで待機
  Step 5: catw_selenium.py  — Playwright で CATW Web UI に自動入力

実行環境: Windows（win32com / pywin32 が必要）
"""

import sys
import subprocess
import time
import urllib.request
import urllib.error
from datetime import date
from pathlib import Path

import yaml

# win32com は Windows のみ利用可能
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

SCRIPTS_DIR = Path(__file__).parent


def load_config(path: Path = Path("config.yaml")) -> dict:
    if not path.exists():
        print(f"[エラー] 設定ファイルが見つかりません: {path}")
        print("  cp config.example.yaml config.yaml  を実行して設定してください。")
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_transfer(year: int, month: int) -> int:
    """transfer_catw.py をサブプロセスで実行する。

    戻り値:
        0 = 転記成功
        1 = エラー
        2 = ユーザーキャンセル
    """
    script = SCRIPTS_DIR / "transfer_catw.py"
    result = subprocess.run(
        [sys.executable, str(script), str(year), str(month)],
        # stdin/stdout/stderr をそのまま端末に繋ぐ（確認プロンプトをパススルー）
    )
    return result.returncode


def open_excel_and_run_macro(excel_path: str, macro_name: str) -> None:
    """win32com で Excel ファイルを開き、指定の VBA マクロを実行する。

    Excel が既に開いている場合は既存インスタンスを再利用する。
    """
    if not WIN32COM_AVAILABLE:
        print("[エラー] win32com が利用できません。Windows 環境で実行してください。")
        sys.exit(1)

    # Excel インスタンスを取得（既存 or 新規）
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("  既存の Excel インスタンスを使用します。")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
        print("  新しい Excel インスタンスを起動します。")

    excel.Visible = True

    # 対象ブックを検索（既に開いている場合はそれを使う）
    target_wb = None
    excel_name = Path(excel_path).name
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks.Item(i)
        if wb.Name == excel_name:
            target_wb = wb
            print(f"  既に開いているブックを使用: {excel_name}")
            break

    if target_wb is None:
        print(f"  Excel ファイルを開きます: {excel_path}")
        target_wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))

    print(f"  マクロを実行します: {macro_name}")
    excel.Run(f"'{target_wb.Name}'!{macro_name}")


def wait_for_edge_cdp(port: int, timeout: int = 30) -> bool:
    """Edge の CDP ポートが応答するまでポーリングする。

    Args:
        port: Edge デバッグポート（デフォルト 9222）
        timeout: 最大待機秒数

    Returns:
        True = 接続確立 / False = タイムアウト
    """
    url = f"http://127.0.0.1:{port}/json/version"
    print(f"  Edge CDP ポート ({port}) の応答を待機中...", end="", flush=True)

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2):
                print(" 接続しました。")
                return True
        except (urllib.error.URLError, OSError):
            print(".", end="", flush=True)
            time.sleep(1)

    print(f"\n[警告] {timeout}秒待機しましたが CDP に接続できませんでした。")
    return False


def run_selenium() -> int:
    """catw_selenium.py をサブプロセスで実行する。

    戻り値:
        0 = 成功
        それ以外 = エラー
    """
    script = SCRIPTS_DIR / "catw_selenium.py"
    if not script.exists():
        print(f"[エラー] スクリプトが見つかりません: {script}")
        print("  catw_selenium.py は .gitignore 対象です。")
        print("  既存の catw_selenium.py を scripts/ に配置してください。")
        return 1

    result = subprocess.run([sys.executable, str(script)])
    return result.returncode


def main() -> None:
    today = date.today()
    if len(sys.argv) == 3:
        year, month = int(sys.argv[1]), int(sys.argv[2])
    elif len(sys.argv) == 1:
        year, month = today.year, today.month
    else:
        print("使い方: python scripts/run_catw.py [年] [月]")
        sys.exit(1)

    print(f"\n{'='*50}")
    print(f"  CATW 一括実行: {year}年{month}月")
    print(f"{'='*50}\n")

    config = load_config()
    catw_cfg = config.get("catw", {}).get("excel", {})
    excel_path = catw_cfg.get("path", "")
    macro_name = catw_cfg.get("macro_name", "OpenCATW")
    cdp_port = int(catw_cfg.get("cdp_port", 9222))

    if not excel_path or "/path/to/" in excel_path:
        print("[エラー] config.yaml の catw.excel.path を設定してください。")
        sys.exit(1)

    # ── Step 1: MD → Excel 転記 ──────────────────────────────────────────
    print("【Step 1】 MD → Excel 転記")
    rc = run_transfer(year, month)
    if rc == 2:
        print("\n転記をキャンセルしました。終了します。")
        sys.exit(0)
    if rc != 0:
        print(f"\n[エラー] transfer_catw.py が失敗しました（終了コード: {rc}）")
        sys.exit(1)

    # ── Step 2: Web 自動入力を続けるか確認 ───────────────────────────────
    print("\n【Step 2】 CATW Web への自動入力")
    ans = input("CATW Web への自動入力を実行しますか？ [y/N]: ").strip().lower()
    if ans != "y":
        print("Web 自動入力をスキップしました。")
        sys.exit(0)

    # ── Step 3: Excel を開いて OpenCATW マクロを実行 ─────────────────────
    print("\n【Step 3】 OpenCATW マクロを実行（Edge をデバッグモードで起動）")
    open_excel_and_run_macro(excel_path, macro_name)

    # ── Step 4: Edge CDP 接続待機 ────────────────────────────────────────
    print("\n【Step 4】 Edge CDP 接続待機")
    if not wait_for_edge_cdp(cdp_port):
        ans2 = input("CDP に接続できませんでした。それでも続けますか？ [y/N]: ").strip().lower()
        if ans2 != "y":
            print("終了します。")
            sys.exit(1)

    # ── Step 5: catw_selenium.py 実行 ────────────────────────────────────
    print("\n【Step 5】 catw_selenium.py を実行")
    rc = run_selenium()
    if rc != 0:
        print(f"\n[エラー] catw_selenium.py が失敗しました（終了コード: {rc}）")
        sys.exit(1)

    print(f"\n{'='*50}")
    print(f"  完了: {year}年{month}月 CATW 一括実行")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
