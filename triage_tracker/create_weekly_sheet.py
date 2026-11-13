#!/usr/bin/env python3
"""Create or update the weekly tab on the triage effort workbook."""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.report_week import (  # noqa: E402
    ReportWeekConfig,
    format_report_week,
    report_week_start,
)

ROOT = Path(__file__).resolve().parent
LEGACY_SHEETS = ("Summary", "IOTREQ")

COLUMNS = [
    "Engineer",
    "Stack",
    "Task",
    "Pipeline",
    "Started",
    "Target complete",
    "SDK build",
    "Test suite(s)",
    "Failures assigned",
    "Effort (days)",
    "Status",
    "Comments",
    "Done?",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(size=14, bold=True, color="1F4E79")
SECTION_FONT = Font(size=12, bold=True, color="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
PIPELINE_FILL = PatternFill("solid", fgColor="E8EDF5")
SUMMARY_HEADER_FILL = PatternFill("solid", fgColor="D6E6F5")
_WRAP_TITLES = frozenset({"Test suite(s)", "Comments"})

_THIN = Border(
    left=Side(style="thin", color="D0D7E2"),
    right=Side(style="thin", color="D0D7E2"),
    top=Side(style="thin", color="D0D7E2"),
    bottom=Side(style="thin", color="D0D7E2"),
)

COL_WIDTHS = {
    "Engineer": 20,
    "Stack": 10,
    "Task": 18,
    "Pipeline": 14,
    "Started": 14,
    "Target complete": 16,
    "SDK build": 12,
    "Test suite(s)": 44,
    "Failures assigned": 16,
    "Effort (days)": 12,
    "Status": 16,
    "Comments": 48,
    "Done?": 8,
}

PIPELINE_SUMMARY = (
    ("smoke", "Smoke"),
    ("regression", "Regression"),
    ("ext_smoke", "Ext smoke"),
    ("ext_regression", "Ext regression"),
)

DATA_ROW_HEIGHT = 42
HEADER_ROW_HEIGHT = 32
WRAP_ROW_HEIGHT = 54

WEEK_HEADER_ROW = 4
DATA_FIRST_ROW = WEEK_HEADER_ROW + 1
DATA_LAST_ROW = 220
SUMMARY_MARKER = "effort by pipeline"
WEEK_TAB_RE = re.compile(r"^week\s*(\d+)$", re.IGNORECASE)


def load_config() -> dict:
    path = ROOT / "config.yaml"
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def report_week_cfg(cfg: dict) -> ReportWeekConfig:
    return ReportWeekConfig.from_mapping(cfg.get("report_week"))


def week_tabs_cfg(cfg: dict) -> dict:
    return cfg.get("week_tabs") or {}


def parse_week_number(tab_name: str) -> int | None:
    m = WEEK_TAB_RE.match(str(tab_name or "").strip())
    return int(m.group(1)) if m else None


def tab_name_for_number(number: int) -> str:
    return f"Week {number}"


def week_1_report_start(cfg: dict) -> datetime:
    raw = week_tabs_cfg(cfg).get("week_1_start") or "2026-01-01"
    anchor = datetime.strptime(str(raw).strip()[:10], "%Y-%m-%d")
    return report_week_start(anchor, report_week_cfg(cfg))


def report_start_for_week_number(cfg: dict, number: int) -> datetime:
    return week_1_report_start(cfg) + timedelta(days=7 * (number - 1))


def current_week_number(cfg: dict | None = None, at: datetime | None = None) -> int:
    cfg = cfg or load_config()
    rw = report_week_cfg(cfg)
    at = at or datetime.now().replace(microsecond=0)
    cycle = report_week_start(at, rw).date()
    anchor = week_1_report_start(cfg).date()
    if cycle < anchor:
        return 1
    return (cycle - anchor).days // 7 + 1


def week_tab_name(cfg: dict | None = None, at: datetime | None = None) -> str:
    cfg = cfg or load_config()
    return tab_name_for_number(current_week_number(cfg, at))


def previous_week_tab_name(week_name: str) -> str:
    number = parse_week_number(week_name)
    if number is None or number <= 1:
        return ""
    return tab_name_for_number(number - 1)


def week_tab_label(cfg: dict, week_name: str) -> str:
    number = parse_week_number(week_name)
    if number is None:
        return week_name
    start = report_start_for_week_number(cfg, number)
    return format_report_week(start, report_week_cfg(cfg))


def minimum_tab_count(cfg: dict) -> int:
    return max(1, int(week_tabs_cfg(cfg).get("min_tabs") or 10))


def sanitize_tab(name: str) -> str:
    name = re.sub(r'[\\/*?:\[\]]', "-", name)[:31]
    return name or "Week"


def _col_letter(title: str, columns: list[str]) -> str:
    return get_column_letter(columns.index(title) + 1)


def _sheet_ref(tab_name: str) -> str:
    safe = tab_name.replace("'", "''")
    return f"'{safe}'"


def _apply_border(ws: Worksheet, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row, col).border = _THIN


def _cell_alignment(title: str) -> Alignment:
    if title in _WRAP_TITLES:
        return Alignment(
            horizontal="left",
            vertical="top",
            wrap_text=True,
            indent=1,
        )
    if title in ("Done?", "Pipeline", "Stack"):
        return Alignment(horizontal="center", vertical="center", wrap_text=True)
    if title == "Effort (days)":
        return Alignment(horizontal="center", vertical="center")
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def _style_header(ws: Worksheet, row: int, columns: list[str]) -> None:
    ws.row_dimensions[row].height = HEADER_ROW_HEIGHT
    for col, title in enumerate(columns, start=1):
        cell = ws.cell(row, col, title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(title, 16)
    _apply_border(ws, row, row, 1, len(columns))


def _yellow_fields(columns: list[str]) -> set[str]:
    return {
        "Started",
        "Target complete",
        "SDK build",
        "Test suite(s)",
        "Failures assigned",
        "Effort (days)",
        "Status",
        "Comments",
    }


def _format_data_row(ws: Worksheet, row: int, columns: list[str], *, tall: bool = False) -> None:
    ws.row_dimensions[row].height = WRAP_ROW_HEIGHT if tall else DATA_ROW_HEIGHT
    for col, title in enumerate(columns, start=1):
        cell = ws.cell(row, col)
        cell.alignment = _cell_alignment(title)
        if title == "Effort (days)":
            cell.number_format = "0.##"


def _format_table(ws: Worksheet, header_row: int, last_row: int, columns: list[str]) -> None:
    if last_row < header_row:
        return
    _apply_border(ws, header_row, last_row, 1, len(columns))
    for row in range(header_row + 1, last_row + 1):
        if not any(ws.cell(row, c).value for c in range(1, len(columns) + 1)):
            ws.row_dimensions[row].height = 14
            continue
        _format_data_row(ws, row, columns)


def _sumif_on_sheet(
    tab_name: str | None,
    pipeline_key: str,
    columns: list[str],
    *,
    first_row: int = DATA_FIRST_ROW,
    last_row: int = DATA_LAST_ROW,
) -> str:
    pipe_col = _col_letter("Pipeline", columns)
    days_col = _col_letter("Effort (days)", columns)
    if tab_name:
        ref = _sheet_ref(tab_name)
        target_pipe = f"{ref}!{pipe_col}{first_row}:{pipe_col}{last_row}"
        target_days = f"{ref}!{days_col}{first_row}:{days_col}{last_row}"
    else:
        target_pipe = f"{pipe_col}{first_row}:{pipe_col}{last_row}"
        target_days = f"{days_col}{first_row}:{days_col}{last_row}"
    return (
        f"SUMIF({target_pipe},"
        f'"{pipeline_key}",{target_days})'
    )


def _last_engineer_data_row(ws: Worksheet) -> int:
    last = WEEK_HEADER_ROW
    for row in range(WEEK_HEADER_ROW + 1, DATA_LAST_ROW + 1):
        if ws.cell(row, 1).value or ws.cell(row, 3).value:
            last = row
    return last


def _find_summary_start(ws: Worksheet) -> int | None:
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if val and SUMMARY_MARKER in str(val).lower():
            return row
    return None


def _clear_summary_block(ws: Worksheet, start_row: int, height: int = 18) -> None:
    ncol = 4
    for row in range(start_row, start_row + height):
        for col in range(1, ncol + 1):
            cell = ws.cell(row, col)
            cell.value = None
            cell.font = Font()
            cell.fill = PatternFill()
            cell.number_format = "General"


def _write_pipeline_summary_block(
    ws: Worksheet,
    start_row: int,
    week_name: str,
    columns: list[str],
) -> int:
    prev_tab = previous_week_tab_name(week_name)

    ws.cell(start_row, 1, "This week — effort by pipeline (days)").font = SECTION_FONT
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=4)
    prev_note = (
        f"Compared to previous week tab: {sanitize_tab(prev_tab)}"
        if prev_tab
        else "No previous week tab (Week 1)"
    )
    ws.cell(start_row + 1, 1, prev_note)
    ws.cell(start_row + 1, 1).font = Font(color="667085", italic=True)
    ws.merge_cells(start_row=start_row + 1, start_column=1, end_row=start_row + 1, end_column=4)

    header = start_row + 3
    headers = ("Pipeline", "This week", "Previous week", "Change")
    for col, label in enumerate(headers, start=1):
        cell = ws.cell(header, col, label)
        cell.fill = SUMMARY_HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _apply_border(ws, header, header, 1, 4)

    row = header + 1
    this_cells: list[str] = []
    prev_cells: list[str] = []
    for key, label in PIPELINE_SUMMARY:
        ws.cell(row, 1, label)
        this_formula = f"={_sumif_on_sheet(None, key, columns)}"
        if prev_tab:
            prev_inner = _sumif_on_sheet(sanitize_tab(prev_tab), key, columns)
            prev_formula = f'=IFERROR({prev_inner},"")'
        else:
            prev_formula = '=""'
        ws.cell(row, 2, this_formula)
        ws.cell(row, 3, prev_formula)
        ws.cell(row, 4, f'=IF(C{row}="","",B{row}-C{row})')
        for col in (2, 3, 4):
            ws.cell(row, col).number_format = "0.##"
        this_cells.append(f"B{row}")
        prev_cells.append(f"C{row}")
        row += 1

    ws.cell(row, 1, "All pipelines")
    ws.cell(row, 1).font = Font(bold=True)
    ws.cell(row, 2, f"=SUM({','.join(this_cells)})")
    ws.cell(row, 3, f"=SUM({','.join(prev_cells)})")
    ws.cell(row, 4, f'=IF(C{row}="","",B{row}-C{row})')
    for col in (2, 3, 4):
        ws.cell(row, col).number_format = "0.##"
        ws.cell(row, col).font = Font(bold=True)
    _apply_border(ws, header, row, 1, 4)

    note = row + 2
    ws.cell(
        note,
        1,
        "Fill Effort (days) with numbers (e.g. 0.5, 2). Change = this week minus previous week (same pipeline).",
    )
    ws.cell(note, 1).font = Font(color="667085", italic=True)
    ws.merge_cells(start_row=note, start_column=1, end_row=note, end_column=len(columns))
    ws.cell(note, 1).alignment = Alignment(wrap_text=True)
    return row


def _refresh_pipeline_summary(ws: Worksheet, week_name: str) -> None:
    existing = _find_summary_start(ws)
    if existing:
        _clear_summary_block(ws, existing)
        start = existing
    else:
        start = _last_engineer_data_row(ws) + 3
    _write_pipeline_summary_block(ws, start, week_name, COLUMNS)


def task_labels(cfg: dict, engineer: str = "") -> list[tuple[str, str]]:
    """Return (pipeline, task label) for each row to seed per engineer."""
    skip_smoke = engineer.strip() in {
        str(name).strip() for name in (cfg.get("no_smoke_engineers") or [])
    }
    labels: list[tuple[str, str]] = []
    for item in cfg.get("weekly_tasks_per_engineer") or []:
        pipeline = str(item.get("pipeline") or "").strip()
        if skip_smoke and pipeline == "smoke":
            continue
        count = int(item.get("count") or 1)
        for index in range(1, count + 1):
            if pipeline == "smoke" and count > 1:
                labels.append((pipeline, f"smoke {index}/{count}"))
            else:
                labels.append((pipeline, pipeline if count == 1 else f"{pipeline} {index}"))
    return labels


def _write_row(
    ws: Worksheet,
    row: int,
    columns: list[str],
    *,
    preset: dict[str, str],
) -> None:
    yellow = _yellow_fields(columns)
    for col, title in enumerate(columns, start=1):
        cell = ws.cell(row, col)
        if title in preset:
            cell.value = preset[title]
        elif title == "Done?":
            cell.value = "N"
        if title in yellow:
            cell.fill = INPUT_FILL
        elif title == "Pipeline":
            cell.fill = PIPELINE_FILL
    _format_data_row(ws, row, columns)


def _seed_engineer_week(
    ws: Worksheet, start_row: int, stack: str, engineer: str, tasks: list[tuple[str, str]]
) -> int:
    row = start_row
    for pipeline, task in tasks:
        _write_row(
            ws,
            row,
            COLUMNS,
            preset={
                "Engineer": engineer,
                "Stack": stack,
                "Task": task,
                "Pipeline": pipeline,
            },
        )
        row += 1
    return row


def _create_week_tab(wb: Workbook, week_name: str, cfg: dict) -> Worksheet:
    tab = sanitize_tab(week_name)
    ws = wb.create_sheet(tab)
    span = week_tab_label(cfg, week_name)
    ws["A1"] = f"{week_name} — report week {span}"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    ws["A2"] = (
        "One row per task for the week (5× smoke + regression + ext_smoke + ext_regression per person). "
        "Effort (days) drives totals and week-over-week comparison at the bottom."
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLUMNS))
    ws["A2"].font = Font(color="667085")
    ws["A2"].alignment = Alignment(wrap_text=True)
    header_row = WEEK_HEADER_ROW
    _style_header(ws, header_row, COLUMNS)
    engineers = cfg.get("engineers") or {}
    row = header_row + 1
    last_data = header_row
    for stack in ("rail", "rail_sdk"):
        for engineer in engineers.get(stack) or []:
            tasks = task_labels(cfg, engineer)
            row = _seed_engineer_week(ws, row, stack, engineer, tasks)
            last_data = row - 1
            row += 1
    _format_table(ws, header_row, last_data, COLUMNS)
    ws.freeze_panes = f"A{header_row + 1}"
    return ws


def ensure_week_tab(wb: Workbook, week_name: str, cfg: dict) -> Worksheet:
    tab = sanitize_tab(week_name)
    if tab not in wb.sheetnames:
        ws = _create_week_tab(wb, week_name, cfg)
    else:
        ws = wb[tab]
    _refresh_pipeline_summary(ws, week_name)
    return ws


def remove_legacy_sheets(wb: Workbook) -> None:
    for name in LEGACY_SHEETS:
        if name in wb.sheetnames:
            del wb[name]


def ensure_numbered_week_tabs(wb: Workbook, cfg: dict) -> None:
    """Create Week 1 … Week N (at least min_tabs; N includes current report week)."""
    floor = minimum_tab_count(cfg)
    current = current_week_number(cfg)
    last = max(floor, current)
    for number in range(1, last + 1):
        ensure_week_tab(wb, tab_name_for_number(number), cfg)
    reorder_week_tabs(wb)


def reorder_week_tabs(wb: Workbook) -> None:
    numbered: list[tuple[int, Worksheet]] = []
    rest: list[Worksheet] = []
    for ws in wb.worksheets:
        n = parse_week_number(ws.title)
        if n is not None:
            numbered.append((n, ws))
        else:
            rest.append(ws)
    if not numbered:
        return
    numbered.sort(key=lambda item: item[0])
    wb._sheets = [ws for _, ws in numbered] + rest  # type: ignore[attr-defined]


def main() -> int:
    parser = argparse.ArgumentParser(description="Add this week's triage effort tab.")
    parser.add_argument(
        "--week",
        help='Tab override (e.g. "Week 3"; default = current report week number)',
    )
    parser.add_argument(
        "--recreate-week",
        action="store_true",
        help="Replace this week's tab if it already exists",
    )
    args = parser.parse_args()
    cfg = load_config()
    path = Path(cfg.get("workbook") or ROOT / "RAIL-triage-effort.xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)
    week_name = args.week or week_tab_name(cfg)

    if path.is_file():
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    tab = sanitize_tab(week_name)
    if args.recreate_week and tab in wb.sheetnames:
        del wb[tab]

    remove_legacy_sheets(wb)
    if args.week:
        ensure_week_tab(wb, week_name, cfg)
    else:
        ensure_numbered_week_tabs(wb, cfg)
    wb.save(path)
    print(f"Workbook: {path}")
    print(f"Active report week tab: {week_tab_name(cfg)}")
    print(f"Tabs ensured: Week 1 … Week {max(minimum_tab_count(cfg), current_week_number(cfg))}")
    prev = previous_week_tab_name(week_tab_name(cfg))
    if prev:
        print(f"Compares to: {prev}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
