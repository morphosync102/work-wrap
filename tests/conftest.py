"""pytest 共通設定

catw_selenium.py は playwright / win32com / winreg（Windows専用）を import するため、
テスト実行前にスタブを差し込んでおく。
"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

# ── Windows / Playwright 依存モジュールをスタブ化 ──────────────────────────────
# catw_selenium.py が import される前に sys.modules に差し込む必要がある

_pw_stub = MagicMock()
_pw_sync_stub = MagicMock()
# TimeoutError だけは本物の Exception を継承させる（isinstance チェックがあるため）
_pw_sync_stub.TimeoutError = type("PlaywrightTimeout", (Exception,), {})

sys.modules.setdefault("playwright", _pw_stub)
sys.modules.setdefault("playwright.sync_api", _pw_sync_stub)
sys.modules.setdefault("win32com", MagicMock())
sys.modules.setdefault("win32com.client", MagicMock())
# winreg は mock しない: catw_selenium.py が try/except ImportError で保護済み。
# mock すると macOS の mimetypes.py が winreg を使おうとしてエラーになる。

# ── scripts/ ディレクトリを sys.path に追加 ────────────────────────────────────
_scripts_dir = Path(__file__).parent.parent / "scripts"
if str(_scripts_dir) not in sys.path:
    sys.path.insert(0, str(_scripts_dir))
