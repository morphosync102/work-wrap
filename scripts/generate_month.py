#!/usr/bin/env python3
"""
月次勤怠MDファイル生成スクリプト

使い方:
  python scripts/generate_month.py 2026 4
  python scripts/generate_month.py          # 引数なしで今月を生成
"""

import sys
import calendar
from datetime import date, time, timedelta
from pathlib import Path

import yaml
import holidays

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

# 会社設定（config.yaml で上書き可能）
DEFAULT_START_TIME = "09:00"
DEFAULT_CORE_HOURS = 7.5
DEFAULT_BREAK_HOURS = 1.0  # 休憩時間


def load_config(path: Path = Path("config.yaml")) -> dict:
    if not path.exists():
        print(f"[エラー] 設定ファイルが見つかりません: {path}")
        print("  cp config.example.yaml config.yaml  を実行して設定してください。")
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_jp_holidays(year: int) -> dict[date, str]:
    return holidays.Japan(years=year, language="ja")


def is_non_working(d: date, jp_holidays: dict) -> bool:
    return d.weekday() >= 5 or d in jp_holidays


def format_date_label(d: date, jp_holidays: dict) -> str:
    weekday = WEEKDAY_JP[d.weekday()]
    if d in jp_holidays:
        return f"{d.month}/{d.day} (祝){jp_holidays[d]}"
    return f"{d.month}/{d.day} ({weekday})"


def calc_end_time(start: str, core_hours: float, overtime_hours: float, break_hours: float) -> str:
    """出勤時刻・コアタイム・残業・休憩から退勤時刻を計算する。"""
    h, m = map(int, start.split(":"))
    total_minutes = int((core_hours + overtime_hours + break_hours) * 60)
    end_minutes = h * 60 + m + total_minutes
    return f"{end_minutes // 60:02d}:{end_minutes % 60:02d}"


def generate_md(year: int, month: int, config: dict) -> str:
    projects = config.get("projects", [])
    jp_hol = get_jp_holidays(year)
    _, num_days = calendar.monthrange(year, month)
    days = [date(year, month, d) for d in range(1, num_days + 1)]

    work_cfg = config.get("workday", {})
    start_time = work_cfg.get("start_time", DEFAULT_START_TIME)
    core_hours = float(work_cfg.get("core_hours", DEFAULT_CORE_HOURS))
    break_hours = float(work_cfg.get("break_hours", DEFAULT_BREAK_HOURS))
    # 残業0の定時退勤時刻
    default_end = calc_end_time(start_time, core_hours, 0, break_hours)

    lines = []

    # ===== ヘッダー =====
    lines.append(f"# {year}年{month}月 勤怠\n")

    # ===== サマリー =====
    lines.append("## サマリー\n")
    lines.append("| 項目 | 値 |")
    lines.append("|------|-----|")
    lines.append("| 稼働日数 |  |")
    lines.append("| 残業合計 |  |")
    for p in projects:
        lines.append(f"| CATW {p['name']} 合計 |  |")
    lines.append("| PSA 合計 |  |")
    lines.append("| 有給取得日数 |  |")
    lines.append("")
    lines.append("---\n")

    # ===== CATW テーブル =====
    lines.append("## CATW（案件工数）\n")

    catw_header = ["日付"] + [p["name"] for p in projects] + ["合計"]
    catw_sep    = ["------"] + ["----------"] * len(projects) + ["------"]
    lines.append("| " + " | ".join(catw_header) + " |")
    lines.append("| " + " | ".join(catw_sep) + " |")

    for d in days:
        label = format_date_label(d, jp_hol)
        if is_non_working(d, jp_hol):
            cells = [f"**{label}**"] + [" - "] * len(projects) + [" - "]
        else:
            cells = [label] + [""] * len(projects) + [""]
        lines.append("| " + " | ".join(cells) + " |")

    lines.append("")
    lines.append("---\n")

    # ===== PSA テーブル =====
    lines.append("## PSA（案件外工数）\n")
    lines.append("> PSAの詳細フォーマットは連携待ち。現在は工数合計と種別のみ記録。\n")

    psa_types = config.get("psa", {}).get("types", [])
    if psa_types:
        psa_cols = ["日付"] + psa_types + ["合計"]
        psa_sep  = ["------"] + ["----------"] * len(psa_types) + ["------"]
    else:
        psa_cols = ["日付", "工数合計", "種別・備考"]
        psa_sep  = ["------", "----------", "----------------"]

    lines.append("| " + " | ".join(psa_cols) + " |")
    lines.append("| " + " | ".join(psa_sep) + " |")

    for d in days:
        label = format_date_label(d, jp_hol)
        if is_non_working(d, jp_hol):
            cells = [f"**{label}**"] + [" - "] * (len(psa_cols) - 1)
        else:
            cells = [label] + [""] * (len(psa_cols) - 1)
        lines.append("| " + " | ".join(cells) + " |")

    lines.append("")
    lines.append("---\n")

    # ===== Workday テーブル =====
    lines.append("## Workday（出退勤）\n")
    lines.append(f"> 出勤: {start_time} 固定 / コアタイム: {core_hours}h / 休憩: {break_hours}h\n")
    lines.append(f"> 定時退勤: {default_end} / 退勤 = 出勤 + コアタイム + 残業 + 休憩\n")

    wd_header = ["日付", "出勤", "退勤", "総時間(h)", "残業(h)"]
    wd_sep    = ["------", "------", "------", "----------", "--------"]
    lines.append("| " + " | ".join(wd_header) + " |")
    lines.append("| " + " | ".join(wd_sep) + " |")

    for d in days:
        label = format_date_label(d, jp_hol)
        if is_non_working(d, jp_hol):
            cells = [f"**{label}**", " - ", " - ", " - ", " - "]
        else:
            cells = [label, start_time, "", "", ""]
        lines.append("| " + " | ".join(cells) + " |")

    lines.append("")
    lines.append("---\n")

    # ===== Concur テーブル =====
    lines.append("## Concur（交通費）\n")

    concur_header = ["日付", "交通手段", "金額(円)", "コメント", "Business Purpose"]
    concur_sep    = ["------", "----------", "----------", "----------", "----------------"]
    lines.append("| " + " | ".join(concur_header) + " |")
    lines.append("| " + " | ".join(concur_sep) + " |")
    lines.append("| | | | | |")  # 空行（交通費は毎日あるわけではないため行は空）

    lines.append("")
    lines.append("---\n")

    # ===== 備考 =====
    lines.append("### 備考\n")
    lines.append(f"- **コアタイム**: {core_hours}h / **休憩**: {break_hours}h / **出勤**: {start_time}")
    lines.append(f"- **定時退勤**: {default_end}")
    if projects:
        lines.append("- **WBSコード一覧** (CATW):")
        for p in projects:
            lines.append(f"  - {p['name']}: `{p.get('wbs', 'TBD')}`")
    attend_types = config.get("attend_absence_types", [])
    if attend_types:
        lines.append(f"- **Attend/Absence Type 選択肢**: {', '.join(attend_types)}")
    if psa_types:
        lines.append(f"- **PSA 種別**: {', '.join(psa_types)}")
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
        sys.exit(1)

    config = load_config()

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
