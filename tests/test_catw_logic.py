"""catw_selenium.py の純粋ロジック関数ユニットテスト

ブラウザ・win32com 不要の関数を対象とする。
ブラウザを使う統合テスト（CATWAutomation.*）は Windows + Edge 環境でのみ実行可能なため
pytest.mark.skip でマークしてある。

実行方法:
    pip install -e ".[dev]"
    pytest tests/test_catw_logic.py -v
"""
import sys
from datetime import datetime
from unittest.mock import MagicMock, patch

import pytest

# conftest.py でスタブ化済みのため、ここで import できる
import catw_selenium as cs


# ──────────────────────────────────────────────────────────────────────────────
# extract_month_from_filename
# ──────────────────────────────────────────────────────────────────────────────
class TestExtractMonthFromFilename:
    def test_space_separator(self):
        """スペース区切り: 'FY26 3月_test.xlsm' → 3"""
        assert cs.extract_month_from_filename("FY26 3月_test.xlsm") == 3

    def test_no_separator_two_digit(self):
        """区切りなし2桁: 'FY2612月' → 12"""
        assert cs.extract_month_from_filename("FY2612月") == 12

    def test_double_space(self):
        """ダブルスペース: 'FY26  12月' → 12"""
        assert cs.extract_month_from_filename("FY26  12月") == 12

    def test_underscore_returns_none(self):
        """アンダースコア区切りは regex にマッチしない → None
        （config.yaml の標準ファイル名 'FY26_03月_...' は後方互換ではない）
        """
        assert cs.extract_month_from_filename("FY26_03月_CATW入力マクロ.xlsm") is None

    def test_empty_string(self):
        assert cs.extract_month_from_filename("") is None

    def test_none_input(self):
        assert cs.extract_month_from_filename(None) is None

    def test_no_month_marker(self):
        assert cs.extract_month_from_filename("report_2026.xlsx") is None


# ──────────────────────────────────────────────────────────────────────────────
# get_target_month_days  (2026年3月 で検証)
#
# 2026年3月1日 = 日曜 (weekday=6)
#   Week1_monday = 2/23(月)
#   Week 1: 2/23(月)～3/1(日)  → 3月に属するのは 3/1(日) のみ  → {"sun"}
#   Week 2: 3/2(月)～3/8(日)   → 全日3月                        → all 7 days
#   Week 6: 3/30(月)～4/5(日)  → 3月に属するのは 3/30(月),3/31(火) → {"mon","tue"}
# ──────────────────────────────────────────────────────────────────────────────
_ALL_DAYS = {"mon", "tue", "wed", "thu", "fri", "sat", "sun"}
_YEAR, _MONTH = 2026, 3


class TestGetTargetMonthDays:
    def test_week1_boundary_start(self):
        """Week1: 月が始まる前の月曜から始まる週 → 最終日のみ当月"""
        result = cs.get_target_month_days(1, _YEAR, _MONTH)
        assert result == {"sun"}

    def test_week2_full_in_month(self):
        """Week2: 完全に当月内 → 全曜日"""
        result = cs.get_target_month_days(2, _YEAR, _MONTH)
        assert result == _ALL_DAYS

    def test_week3_full_in_month(self):
        result = cs.get_target_month_days(3, _YEAR, _MONTH)
        assert result == _ALL_DAYS

    def test_week6_boundary_end(self):
        """Week6: 月末で翌月にまたがる週 → 月初2日だけが当月"""
        result = cs.get_target_month_days(6, _YEAR, _MONTH)
        assert result == {"mon", "tue"}

    def test_january_2026_week1_all_days(self):
        """2026年1月1日 = 木曜 → Week1_monday = 12/29(月)
        Week1: 12/29～1/4 → 1月に属するのは木・金・土・日 → {thu,fri,sat,sun}
        """
        result = cs.get_target_month_days(1, 2026, 1)
        assert result == {"thu", "fri", "sat", "sun"}

    def test_result_is_subset_of_all_days(self):
        """結果は常に有効な曜日キーのみを含む"""
        for week in range(1, 7):
            result = cs.get_target_month_days(week, _YEAR, _MONTH)
            assert result.issubset(_ALL_DAYS)


# ──────────────────────────────────────────────────────────────────────────────
# get_week_date
#
# 2026年3月:
#   week1_monday = 2026-02-23
#   Week1 の水曜 = 2/23 + 2 = 2/25
#   Week2 の水曜 = 3/2 + 2 = 3/4
#   Week3 の水曜 = 3/9 + 2 = 3/11
# ──────────────────────────────────────────────────────────────────────────────
class TestGetWeekDate:
    def test_week1_returns_wednesday(self):
        d = cs.get_week_date(1, month=3, year=2026)
        assert d == datetime(2026, 2, 25)

    def test_week2_returns_wednesday(self):
        d = cs.get_week_date(2, month=3, year=2026)
        assert d == datetime(2026, 3, 4)

    def test_week3_returns_wednesday(self):
        d = cs.get_week_date(3, month=3, year=2026)
        assert d == datetime(2026, 3, 11)

    def test_week6_returns_wednesday(self):
        """Week6 の水曜 = 3/30 + 2 = 4/1"""
        d = cs.get_week_date(6, month=3, year=2026)
        assert d == datetime(2026, 4, 1)

    def test_week_num_none_returns_today(self):
        """week_num=None のとき month/year 月の1日を返す（内部で today として使う）"""
        d = cs.get_week_date(None, month=3, year=2026)
        assert d == datetime(2026, 3, 1)

    def test_result_is_always_wednesday(self):
        """戻り値は常に水曜日（weekday=2）"""
        for week in range(1, 7):
            d = cs.get_week_date(week, month=3, year=2026)
            assert d.weekday() == 2, f"Week{week}: {d} は水曜ではない"


# ──────────────────────────────────────────────────────────────────────────────
# get_current_week_num
# ──────────────────────────────────────────────────────────────────────────────
class TestGetCurrentWeekNum:
    @patch("catw_selenium.datetime")
    def test_first_day_of_month_is_week1(self, mock_dt):
        """月初日は Week1"""
        mock_dt.now.return_value = datetime(2026, 3, 1)  # 日曜
        result = cs.get_current_week_num()
        assert result == 1

    @patch("catw_selenium.datetime")
    def test_mid_month(self, mock_dt):
        """3/16 (月) は Week4"""
        mock_dt.now.return_value = datetime(2026, 3, 16)
        result = cs.get_current_week_num()
        assert result == 4

    @patch("catw_selenium.datetime")
    def test_last_days_of_month(self, mock_dt):
        """3/30 (月) は Week6"""
        mock_dt.now.return_value = datetime(2026, 3, 30)
        result = cs.get_current_week_num()
        assert result == 6

    @patch("catw_selenium.datetime")
    def test_result_clamped_to_1_6(self, mock_dt):
        """結果は必ず 1〜6 の範囲内"""
        for day in range(1, 32):
            try:
                mock_dt.now.return_value = datetime(2026, 3, day)
                result = cs.get_current_week_num()
                assert 1 <= result <= 6, f"day={day}: week={result} が範囲外"
            except ValueError:
                pass  # 3/32 など無効日付はスキップ


# ──────────────────────────────────────────────────────────────────────────────
# get_year_month_from_excel  (win32com をモック)
# ──────────────────────────────────────────────────────────────────────────────
class TestGetYearMonthFromExcel:
    def test_returns_none_if_win32com_unavailable(self, monkeypatch):
        """win32com が使えない場合は (None, None)"""
        monkeypatch.setattr(cs, "WIN32COM_AVAILABLE", False)
        assert cs.get_year_month_from_excel() == (None, None)

    def test_returns_none_if_excel_not_open(self, monkeypatch):
        """Excel が開いていない（GetActiveObject が例外）→ (None, None)"""
        monkeypatch.setattr(cs, "WIN32COM_AVAILABLE", True)
        mock_win32 = MagicMock()
        mock_win32.client.GetActiveObject.side_effect = Exception("Excel not running")
        monkeypatch.setattr(cs, "win32com", mock_win32)
        assert cs.get_year_month_from_excel() == (None, None)

    def test_returns_year_month_from_cells(self, monkeypatch):
        """Excel の C4=2026 / C5=3 から (2026, 3) を返す"""
        monkeypatch.setattr(cs, "WIN32COM_AVAILABLE", True)

        mock_cells = MagicMock()
        mock_cells.return_value.Value = None

        def cells_side_effect(row, col):
            m = MagicMock()
            if row == 4 and col == 3:
                m.Value = 2026
            elif row == 5 and col == 3:
                m.Value = 3
            else:
                m.Value = None
            return m

        mock_ws = MagicMock()
        mock_ws.Cells.side_effect = cells_side_effect

        mock_wb = MagicMock()
        mock_wb.Name = "FY26_03月_CATW入力マクロ.xlsm"
        mock_wb.Sheets.return_value = mock_ws

        mock_excel = MagicMock()
        mock_excel.Workbooks.Count = 1
        mock_excel.Workbooks.Item.return_value = mock_wb

        mock_win32 = MagicMock()
        mock_win32.client.GetActiveObject.return_value = mock_excel

        monkeypatch.setattr(cs, "win32com", mock_win32)

        year, month = cs.get_year_month_from_excel()
        assert year == 2026
        assert month == 3

    def test_returns_none_if_no_catw_workbook(self, monkeypatch):
        """CATW という名前のブックがない → (None, None)"""
        monkeypatch.setattr(cs, "WIN32COM_AVAILABLE", True)

        mock_wb = MagicMock()
        mock_wb.Name = "other_file.xlsx"

        mock_excel = MagicMock()
        mock_excel.Workbooks.Count = 1
        mock_excel.Workbooks.Item.return_value = mock_wb

        mock_win32 = MagicMock()
        mock_win32.client.GetActiveObject.return_value = mock_excel

        monkeypatch.setattr(cs, "win32com", mock_win32)

        assert cs.get_year_month_from_excel() == (None, None)


# ──────────────────────────────────────────────────────────────────────────────
# CATWAutomation ブラウザ統合テスト（Windows + Edge 環境が必要）
# ──────────────────────────────────────────────────────────────────────────────
@pytest.mark.skip(reason="Windows + Edge デバッグモード環境が必要")
class TestCATWAutomationIntegration:
    """ブラウザ統合テスト。Windows + OpenCATW マクロで Edge が起動している状態で実行。

    実行方法:
        1. Excel で OpenCATW マクロを実行して Edge をデバッグモードで起動
        2. pytest tests/test_catw_logic.py::TestCATWAutomationIntegration -v -s
    """

    def test_connect_to_existing_browser(self):
        automation = cs.CATWAutomation()
        assert automation.connect(reuse_browser=True) is True

    def test_navigate_to_week1(self):
        automation = cs.CATWAutomation()
        automation.connect(reuse_browser=True)
        assert automation.navigate_to_week(1, month=_MONTH, year=_YEAR) is True

    def test_read_week_data_returns_list(self):
        automation = cs.CATWAutomation()
        automation.connect(reuse_browser=True)
        automation.navigate_to_week(1, month=_MONTH, year=_YEAR)
        data = automation.read_week_data()
        assert isinstance(data, list)
