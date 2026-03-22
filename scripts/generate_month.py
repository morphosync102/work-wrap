#!/usr/bin/env python3
"""
月次勤怠MDファイル生成スクリプト

使い方:
  python scripts/generate_month.py 2026 3
  python scripts/generate_month.py          # 引数なしで今月を生成
"""

import sys
import calendar
from datetime import date
from pathlib import Path

import yaml
import holidays


WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


def load_config(path: Path = Path("config.yaml")) -> dict:
    if not path.exists():
        print(f"[エラー] 設定ファイルが見つかりません: {path}")
        print("  cp config.example.yaml config.yaml  を実行して設定してください。")
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_jp_holidays(year: int) -> dict[date, str]:
    return holidays.Japan(years=year, language="ja")


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5  # 土=5, 日=6


def format_date_label(d: date, jp_holidays: dict) -> str:
    """日付ラベルを生成。例: 3/2 (月) / 3/7 (土) / 3/20 (祝)春分の日"""
    weekday = WEEKDAY_JP[d.weekday()]
    if d in jp_holidays:
        holiday_name = jp_holidays[d]
        return f"{d.month}/{d.day} (祝){holiday_name}"
    return f"{d.month}/{d.day} ({weekday})"


def is_non_working(d: date, jp_holidays: dict) -> bool:
    return is_weekend(d) or d in jp_holidays


def generate_md(year: int, month: int, config: dict) -> str:
    projects = config.get("projects", [])
    attend_types = config.get("attend_absence_types", [])
    jp_holidays = get_jp_holidays(year)

    _, num_days = calendar.monthrange(year, month)
    days = [date(year, month, d) for d in range(1, num_days + 1)]

    lines = []

    # ヘッダー
    lines.append(f"# {year}年{month}月 勤怠\n")

    # サマリーセクション（スクリプト実行時は空欄、手動または集計スクリプトで更新）
    lines.append("## サマリー\n")
    lines.append("| 項目 | 値 |")
    lines.append("|------|-----|")
    lines.append("| 稼働日数 |  |")
    lines.append("| 総稼働時間 |  |")
    for p in projects:
        lines.append(f"| {p['name']} 合計 |  |")
    lines.append("| 残業合計 |  |")
    lines.append("| 有給取得日数 |  |")
    lines.append("")

    # 日次記録テーブルヘッダー
    lines.append("## 日次記録\n")

    project_names = [p["name"] for p in projects]
    header_cols = ["日付"] + project_names + ["残業", "WBS element", "Description", "Attend/Absence Type", "Memo"]
    sep_cols = ["------"] + ["----------"] * len(project_names) + ["------", "-------------", "-------------", "--------------------", "------"]

    lines.append("| " + " | ".join(header_cols) + " |")
    lines.append("| " + " | ".join(sep_cols) + " |")

    # 日次行
    for d in days:
        label = format_date_label(d, jp_holidays)
        non_working = is_non_working(d, jp_holidays)

        if non_working:
            # 土日・祝日はハイフンで埋めて視覚的に区別
            proj_cells = [" - " for _ in projects]
            row = [f"**{label}**"] + proj_cells + [" - ", "", "", "", ""]
        else:
            proj_cells = ["" for _ in projects]
            row = [label] + proj_cells + ["", "", "", "", ""]

        lines.append("| " + " | ".join(row) + " |")

    lines.append("")

    # フッター注記
    lines.append("---")
    lines.append("")
    lines.append("### 備考")
    lines.append("")
    lines.append(f"- **所定労働時間**: {config.get('regular_hours', 8.0)}h/日")
    if attend_types:
        lines.append(f"- **Attend/Absence Type 選択肢**: {', '.join(attend_types)}")
    if projects:
        lines.append("- **WBSコード一覧**:")
        for p in projects:
            lines.append(f"  - {p['name']}: `{p.get('wbs', 'TBD')}`")
    lines.append("")

    return "\n".join(lines)


def main():
    today = date.today()

    if len(sys.argv) == 3:
        year, month = int(sys.argv[1]), int(sys.argv[2])
    elif len(sys.argv) == 1:
        year, month = today.year, today.month
    else:
        print("使い方: python scripts/generate_month.py [年] [月]")
        print("例:     python scripts/generate_month.py 2026 3")
        sys.exit(1)

    config = load_config()

    # 出力先ディレクトリ: 勤怠/YYYY/
    out_dir = Path("勤怠") / str(year)
    out_dir.mkdir(parents=True, exist_ok=True)

    out_path = out_dir / f"{month}月.md"

    if out_path.exists():
        print(f"[警告] すでに存在します: {out_path}")
        answer = input("上書きしますか？ [y/N]: ").strip().lower()
        if answer != "y":
            print("キャンセルしました。")
            sys.exit(0)

    content = generate_md(year, month, config)
    out_path.write_text(content, encoding="utf-8")
    print(f"生成しました: {out_path}")


if __name__ == "__main__":
    main()
