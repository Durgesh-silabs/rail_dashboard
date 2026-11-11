"""Progress dashboard workbook (OneDrive Triage).

Main sheet: latest build whose triage is complete, for package manager
(RAIL and RAIL SDK, all test types) and 2026.12 (RAIL and RAIL SDK, smoke only).
Each test-type sheet lists the last five SDK builds, latest at the top.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet


DASHBOARD_SHEET = "Dashboard"
HISTORY_SHEET = "History"
HISTORY_LIMIT = 5

TEST_TYPES = ("smoke", "ext_smoke", "regression", "ext_regression")
STACKS = ("rail", "rail_sdk")
STACK_LABELS = {"rail": "RAIL", "rail_sdk": "RAIL SDK"}
PKG_MANAGER = "Pkg manager"
SISDK_2026_12 = "2026.12"

MAIN_HEADERS = [
    "Test type",
    "Build number",
    "Date",
    "Triage owner",
    "Triage status",
    "Total",
    "Pass",
    "Pass %",
    "Fail",
    "vs previous fail",
    "Fail %",
    "Skip",
    "Skip reason",
    "Triage confidence",
    "Remarks",
]
PROGRESS_HEADERS = [
    "Build number",
    "Date",
    "Triage owner",
    "Triage status",
    "Total",
    "Pass",
    "Pass %",
    "Fail",
    "vs previous fail",
    "Fail %",
    "Skip",
    "Skip reason",
    "Triage confidence",
    "Remarks",
]
_COLUMN_WIDTHS = {
    "Test type": 18,
    "Build number": 16,
    "Date": 16,
    "Triage owner": 34,
    "Triage status": 22,
    "Total": 12,
    "Pass": 12,
    "Pass %": 12,
    "Fail": 12,
    "vs previous fail": 22,
    "Fail %": 12,
    "Skip": 12,
    "Skip reason": 42,
    "Triage confidence": 24,
    "Remarks": 28,
}
HISTORY_HEADERS = [
    "stack",
    "test_type",
    "sdk_build",
    "fail",
    "skip",
    "passed",
    "jenkins_run",
    "triggered",
    "triage_confidence",
    "triage_status",
    "triage_owner",
    "remarks",
    "release",
    "jenkins_url",
    "skip_reason",
]
MANUAL_FIELDS = (
    ("Triage confidence", "confidence"),
    ("Triage status", "status"),
    ("Triage owner", "owner"),
    ("Remarks", "remarks"),
)
MANUAL_HEADERS = tuple(name for name, _attr in MANUAL_FIELDS)

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_TITLE_FONT = Font(size=16, bold=True, color="1F4E79")
_SECTION_FONT = Font(size=13, bold=True, color="1F4E79")
_INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
_LATEST_FILL = PatternFill("solid", fgColor="D6E6F5")
_STATUS_STYLES = {
    "triage complete": ("C6EFCE", "006100"),
    "triage incomplete": ("FFE699", "9C6500"),
    "not started": ("F4B183", "843C0C"),
}
_THIN = Border(
    left=Side(style="thin", color="D0D7E2"),
    right=Side(style="thin", color="D0D7E2"),
    top=Side(style="thin", color="D0D7E2"),
    bottom=Side(style="thin", color="D0D7E2"),
)
_CENTER = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left", vertical="center")


@dataclass
class DashboardSection:
    release: str
    stack: str
    test_types: tuple[str, ...] = TEST_TYPES

    @property
    def label(self) -> str:
        return f"{self.release} — {STACK_LABELS.get(self.stack, self.stack)}"


SECTIONS = (
    DashboardSection(PKG_MANAGER, "rail"),
    DashboardSection(PKG_MANAGER, "rail_sdk"),
    DashboardSection(SISDK_2026_12, "rail", ("smoke",)),
    DashboardSection(SISDK_2026_12, "rail_sdk", ("smoke",)),
)


@dataclass
class DashboardRun:
    stack: str
    test_type: str
    sdk_build: str
    fail: int
    skip: int
    passed: int
    jenkins_run: int = 0
    confidence: str = ""
    status: str = ""
    owner: str = ""
    remarks: str = ""
    release: str = PKG_MANAGER
    jenkins_url: str = ""
    skip_reason: str = ""
    triggered: str = ""

    @property
    def total(self) -> int:
        return int(self.fail) + int(self.skip) + int(self.passed)

    @property
    def pass_pct(self) -> float:
        total = self.total
        if not total:
            return 0.0
        return round(self.passed / total * 100, 2)

    @property
    def fail_pct(self) -> float:
        total = self.total
        if not total:
            return 0.0
        return round(self.fail / total * 100, 2)

    @property
    def key(self) -> tuple[str, str, str, str]:
        return (self.release, self.stack, self.test_type, str(self.sdk_build))


def summarize_skip_reasons(messages: list[str]) -> str:
    """Collapse skipped-test messages into one cell. Empty when Jenkins has none."""
    counts: dict[str, int] = {}
    for raw in messages:
        text = " ".join(str(raw or "").split())
        if not text:
            continue
        if len(text) > 160:
            text = text[:157] + "..."
        counts[text] = counts.get(text, 0) + 1
    if not counts:
        return ""
    ranked = sorted(counts.items(), key=lambda item: (-item[1], item[0]))
    if len(ranked) == 1:
        text, count = ranked[0]
        return text if count == 1 else f"{text} ({count})"
    parts = [f"{text} ({count})" for text, count in ranked[:3]]
    extra = len(ranked) - len(parts)
    summary = "; ".join(parts)
    if extra:
        summary += f"; +{extra} more"
    return summary


def match_test_type(url_or_path: str, test_types: list[str]) -> str | None:
    """Return the longest matching test-type path segment, if any."""
    text = unquote(url_or_path or "").replace("\\", "/").lower()
    segments = [s for s in text.replace("%2f", "/").split("/") if s]
    flat: list[str] = []
    for seg in segments:
        flat.extend(p for p in seg.split("/") if p)
    ordered = sorted({t.strip().lower() for t in test_types if t.strip()}, key=len, reverse=True)
    for test_type in ordered:
        if test_type in flat:
            return test_type
    return None


def iso_date_from_ms(value: Any) -> str:
    """Local calendar date for a Jenkins build timestamp (milliseconds)."""
    try:
        millis = int(value)
    except (TypeError, ValueError):
        return ""
    if millis <= 0:
        return ""
    return datetime.fromtimestamp(millis / 1000).strftime("%Y-%m-%d")


def _excel_date(value: str):
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return text


def _as_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _history_runs(wb: WorkbookType) -> list[DashboardRun]:
    if HISTORY_SHEET not in wb.sheetnames:
        return []
    ws = wb[HISTORY_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    index = {name: i for i, name in enumerate(headers) if name}
    out: list[DashboardRun] = []
    for raw in rows[1:]:
        if not raw or not raw[0]:
            continue

        def col(name: str) -> Any:
            i = index.get(name)
            if i is None or i >= len(raw):
                return ""
            return raw[i]

        sdk = str(col("sdk_build") or "").strip()
        stack = str(col("stack") or "").strip()
        test_type = str(col("test_type") or "").strip()
        if not sdk or not stack or not test_type:
            continue
        out.append(
            DashboardRun(
                stack=stack,
                test_type=test_type,
                sdk_build=sdk,
                fail=_as_int(col("fail")),
                skip=_as_int(col("skip")),
                passed=_as_int(col("passed")),
                jenkins_run=_as_int(col("jenkins_run")),
                confidence=str(col("triage_confidence") or "").strip(),
                status=_display_status(str(col("triage_status") or "")),
                owner=str(col("triage_owner") or "").strip(),
                remarks=str(col("remarks") or "").strip(),
                release=str(col("release") or PKG_MANAGER).strip() or PKG_MANAGER,
                jenkins_url=str(col("jenkins_url") or "").strip(),
                skip_reason=str(col("skip_reason") or "").strip(),
                triggered=str(col("triggered") or "").strip()[:10],
            )
        )
    return out


def _section_from_label(label: str) -> tuple[str, str] | None:
    text = label.strip()
    if " — " in text:
        release, pretty = text.split(" — ", 1)
        for stack, name in STACK_LABELS.items():
            if pretty.strip() == name:
                return release.strip(), stack
        return None
    for stack, name in STACK_LABELS.items():
        if text == name:
            return PKG_MANAGER, stack
    return None


def _manual_from_sheets(
    wb: WorkbookType,
) -> list[dict[tuple[str, str, str, str], dict[str, str]]]:
    """Yellow-cell values per visible sheet, Dashboard first."""
    sheets: list[dict[tuple[str, str, str, str], dict[str, str]]] = []
    for sheet_name in wb.sheetnames:
        if sheet_name == HISTORY_SHEET:
            continue
        ws = wb[sheet_name]
        test_type = sheet_name if sheet_name in TEST_TYPES else ""
        release = ""
        stack = ""
        header: list[str] | None = None
        found: dict[tuple[str, str, str, str], dict[str, str]] = {}
        for raw in ws.iter_rows(values_only=True):
            if not raw or raw[0] is None:
                header = None
                continue
            label = str(raw[0]).strip()
            parsed = _section_from_label(label)
            if parsed is not None:
                release, stack = parsed
                header = None
                continue
            if label in {"Build number", "Test type"}:
                header = [str(c).strip() if c is not None else "" for c in raw]
                continue
            if not release or not stack or not header or "Triage confidence" not in header:
                continue
            if label in {"not run", "Release", "Latest"}:
                continue
            build_i = header.index("Build number") if "Build number" in header else None
            type_i = header.index("Test type") if "Test type" in header else None
            if build_i is None or build_i >= len(raw) or raw[build_i] in (None, ""):
                continue
            sdk = str(raw[build_i]).strip()
            if sdk == "not run":
                continue
            row_type = test_type
            if type_i is not None and type_i < len(raw) and raw[type_i]:
                row_type = str(raw[type_i]).strip()
            if not row_type:
                continue
            fields: dict[str, str] = {}
            for column, attr in MANUAL_FIELDS:
                if column not in header:
                    continue
                idx = header.index(column)
                value = ""
                if idx < len(raw) and raw[idx] is not None:
                    value = str(raw[idx]).strip()
                if value.startswith("="):
                    continue
                fields[attr] = value
            found[(release, stack, row_type, sdk)] = fields
        sheets.append(found)
    return sheets


def _keep_manual(run: DashboardRun, prev: DashboardRun) -> None:
    for _column, attr in MANUAL_FIELDS:
        if not getattr(run, attr):
            setattr(run, attr, getattr(prev, attr))


def _apply_manual_edits(runs: list[DashboardRun], wb: WorkbookType) -> None:
    """Keep typed yellow cells. A detail-sheet edit wins over the main sheet."""
    saved = {
        run.key: {attr: getattr(run, attr) for _column, attr in MANUAL_FIELDS}
        for run in runs
    }
    by_key = {run.key: run for run in runs}
    edits: dict[tuple[str, str, str, str], dict[str, str]] = {}
    for sheet in _manual_from_sheets(wb):
        for key, fields in sheet.items():
            for attr, value in fields.items():
                if value != saved.get(key, {}).get(attr, ""):
                    edits.setdefault(key, {})[attr] = value
    for key, fields in edits.items():
        run = by_key.get(key)
        if run is None:
            continue
        for attr, value in fields.items():
            setattr(run, attr, value)


def _load_runs(path: Path) -> list[DashboardRun]:
    if not path.is_file():
        return []
    wb = load_workbook(path)
    try:
        runs = _history_runs(wb)
        _apply_manual_edits(runs, wb)
        return runs
    finally:
        wb.close()


def load_dashboard_runs(path: Path) -> list[DashboardRun]:
    """Public loader used when seeding the smoke-progress workbook."""
    return _load_runs(path)


def _named_owner(group: list[DashboardRun]) -> str:
    """Owner typed on the newest build in this section that has one."""
    named = [run for run in group if str(run.owner or "").strip()]
    if not named:
        return ""
    named.sort(key=lambda run: (_sdk_value(run), run.jenkins_run), reverse=True)
    return str(named[0].owner).strip()


def _carry_owners(grouped: dict[tuple[str, str, str], list[DashboardRun]]) -> None:
    """Keep a section owner on every build, including a new latest row.

    ext_smoke stays blank until someone types it. ext_regression can follow
    regression when that section has no owner of its own.
    """
    fallback = {"ext_regression": "regression"}
    for key, group in grouped.items():
        owner = _named_owner(group)
        if not owner:
            sibling = fallback.get(key[2])
            if sibling:
                owner = _named_owner(grouped.get((key[0], key[1], sibling), []))
        if not owner:
            continue
        for run in group:
            if not str(run.owner or "").strip():
                run.owner = owner


_DEFAULT_CONFIDENCE = {
    "smoke": "100",
    "regression": "75",
    "ext_regression": "75",
}


def _apply_default_confidence(grouped: dict[tuple[str, str, str], list[DashboardRun]]) -> None:
    """Fill a blank confidence from the test type. A typed value stays."""
    for (_release, _stack, test_type), group in grouped.items():
        default = _DEFAULT_CONFIDENCE.get(test_type, "")
        if not default:
            continue
        for run in group:
            if not str(run.confidence or "").strip():
                run.confidence = default


def _merge(existing: list[DashboardRun], incoming: list[DashboardRun]) -> list[DashboardRun]:
    by_key = {run.key: run for run in existing}
    for run in incoming:
        prev = by_key.get(run.key)
        if prev is None:
            by_key[run.key] = run
            continue
        if run.jenkins_run >= prev.jenkins_run:
            _keep_manual(run, prev)
            if not run.jenkins_url:
                run.jenkins_url = prev.jenkins_url
            if not run.triggered:
                run.triggered = prev.triggered
            by_key[run.key] = run
        else:
            _keep_manual(prev, run)
            if not prev.jenkins_url:
                prev.jenkins_url = run.jenkins_url
            if not prev.triggered:
                prev.triggered = run.triggered
    grouped: dict[tuple[str, str, str], list[DashboardRun]] = {}
    for run in by_key.values():
        # Pkg manager RAIL smoke 1269 published only the app-build list.
        if (
            run.release == PKG_MANAGER
            and run.stack == "rail"
            and run.test_type == "smoke"
            and str(run.sdk_build) == "1269"
        ):
            continue
        grouped.setdefault((run.release, run.stack, run.test_type), []).append(run)
    _carry_owners(grouped)
    _apply_default_confidence(grouped)
    kept: list[DashboardRun] = []
    for group in grouped.values():
        group.sort(key=lambda r: (r.jenkins_run, _sdk_value(r)), reverse=True)
        kept.extend(group[:HISTORY_LIMIT])
    return kept


def _sdk_value(run: DashboardRun) -> int:
    try:
        return int(str(run.sdk_build).strip())
    except ValueError:
        return run.jenkins_run


def _ordered(
    runs: list[DashboardRun], release: str, stack: str, test_type: str
) -> list[DashboardRun]:
    group = [
        r
        for r in runs
        if r.release == release and r.stack == stack and r.test_type == test_type
    ]
    group.sort(key=_sdk_value, reverse=True)
    return group


def _latest(
    runs: list[DashboardRun], release: str, stack: str, test_type: str
) -> DashboardRun | None:
    group = _ordered(runs, release, stack, test_type)
    if not group:
        return None
    return max(group, key=lambda r: (r.jenkins_run, _sdk_value(r)))


def _is_complete(run: DashboardRun) -> bool:
    return str(run.status or "").strip().lower() == "triage complete"


def _metric_map(run: DashboardRun, older: DashboardRun | None) -> dict[str, Any]:
    fail_delta = None if older is None else run.fail - older.fail
    build_no = _sdk_value(run) if str(run.sdk_build).isdigit() else run.sdk_build
    return {
        "Build number": build_no,
        "Date": _excel_date(run.triggered),
        "Triage owner": run.owner,
        "Triage status": run.status,
        "Total": run.total,
        "Pass": run.passed,
        "Pass %": run.pass_pct,
        "Fail": run.fail,
        "vs previous fail": fail_delta,
        "Fail %": run.fail_pct,
        "Skip": run.skip,
        "Skip reason": run.skip_reason or None,
        "Triage confidence": _confidence_number(run.confidence),
        "Remarks": run.remarks,
        "_fail_delta": fail_delta,
    }


def _style_header_row(ws: Worksheet, row: int, ncol: int) -> None:
    for col in range(1, ncol + 1):
        cell = ws.cell(row, col)
        cell.fill = _HEADER_FILL
        cell.font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = _THIN
    ws.row_dimensions[row].height = 36


def _owner_column_width(names: list[str]) -> float:
    """Excel width that shows the full owner name, with a little padding."""
    texts = ["Triage owner"]
    texts.extend(str(name).strip() for name in names if str(name or "").strip())
    longest = max(len(text) for text in texts)
    return float(longest + 4)


def _apply_widths(ws: Worksheet, headers: list[str], owner_width: float) -> None:
    for index, header in enumerate(headers, start=1):
        width = owner_width if header == "Triage owner" else _COLUMN_WIDTHS[header]
        dimension = ws.column_dimensions[get_column_letter(index)]
        dimension.width = width
        if header == "Triage owner":
            dimension.bestFit = True


def _finish_sheet(ws: Worksheet, headers: list[str], owner_width: float) -> None:
    """Keep headings readable: columns wide enough for the header text."""
    _apply_widths(ws, headers, owner_width)
    ws.row_dimensions[2].height = 32
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = None
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = False
    ws.page_setup.fitToHeight = False
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.view = "normal"
    ws.print_title_rows = "5:5"


def _paint(
    cell, *, input_cell: bool = False, center: bool = True, latest: bool = False
) -> None:
    cell.border = _THIN
    cell.alignment = _CENTER if center else _LEFT
    if input_cell:
        cell.fill = _INPUT_FILL
    elif latest:
        cell.fill = _LATEST_FILL


def _is_manual(header: str) -> bool:
    return header in MANUAL_HEADERS


def _display_status(value: str) -> str:
    text = str(value or "").strip()
    if not text or text.lower() == "not started":
        return "Not started"
    return text


def _apply_status(cell) -> None:
    """Color triage status: green complete, yellow notes open, orange not started."""
    text = _display_status(str(cell.value or ""))
    cell.value = text
    colors = _STATUS_STYLES.get(text.lower())
    if not colors:
        return
    fill, font = colors
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(color=font, bold=True)


def _write_not_run(ws: Worksheet, row: int, headers: list[str]) -> None:
    for col, header in enumerate(headers, start=1):
        manual = _is_manual(header)
        cell = ws.cell(row, col, None if manual else "not run")
        _paint(cell, input_cell=manual, center=True)
        if not manual:
            cell.font = Font(italic=True, color="667085")


def _confidence_number(value: str):
    """Show 100 as 100%. A blank cell stays blank."""
    text = str(value or "").strip().rstrip("%").strip()
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return value
    if number.is_integer():
        return int(number)
    return number


def _apply_number(cell, header: str) -> None:
    if header in {"Fail", "Skip", "Pass", "Total", "Build number"}:
        cell.number_format = "#,##0"
    elif header in {"Pass %", "Fail %"}:
        cell.number_format = "0.00"
    elif header == "vs previous fail":
        cell.number_format = "+#,##0;-#,##0;0"
    elif header == "Date":
        cell.number_format = "DD-MMM-YYYY"
    elif header == "Triage confidence":
        cell.number_format = '0"%"'


def _link_test_type(ws: Worksheet, row: int, test_type: str) -> None:
    type_cell = ws.cell(row, 1, test_type)
    type_cell.alignment = _LEFT
    type_cell.hyperlink = f"#'{test_type}'!A1"
    type_cell.font = Font(color="1F4E79", underline="single")


def _write_main(wb: Workbook, runs: list[DashboardRun]) -> dict[tuple[str, str, str], str]:
    """Write the main sheet and return Triage owner cells by release, stack, and test type."""
    ws = wb.create_sheet(DASHBOARD_SHEET, 0)
    ws.sheet_properties.tabColor = "1F4E79"
    ws["A1"] = "RAIL dashboard"
    ws["A1"].font = _TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(MAIN_HEADERS))
    ws["A2"] = (
        "This page shows the latest build whose triage is complete. "
        "Type the yellow cells here. Triage status follows the IoT report. "
        "Triage owner on the other sheets follows this page."
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(MAIN_HEADERS))
    ws["A2"].font = Font(color="667085")

    owner_cells: dict[tuple[str, str, str], str] = {}
    row = 4
    for section in SECTIONS:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(MAIN_HEADERS))
        ws.cell(row, 1, section.label).font = _SECTION_FONT
        row += 1
        for col, header in enumerate(MAIN_HEADERS, start=1):
            ws.cell(row, col, header)
        _style_header_row(ws, row, len(MAIN_HEADERS))
        row += 1
        for test_type in section.test_types:
            ordered = _ordered(runs, section.release, section.stack, test_type)
            completed = [run for run in ordered if _is_complete(run)]
            if not ordered:
                _write_not_run(ws, row, MAIN_HEADERS)
                _link_test_type(ws, row, test_type)
                owner_col = MAIN_HEADERS.index("Triage owner") + 1
                owner_cells[(section.release, section.stack, test_type)] = ws.cell(
                    row, owner_col
                ).coordinate
                row += 1
                continue
            shown = completed[0] if completed else None
            older = completed[1] if len(completed) > 1 else None
            metrics = _metric_map(shown, older) if shown is not None else {}
            if shown is None:
                metrics = {"Triage owner": _named_owner(ordered)}
            fail_delta = metrics.get("_fail_delta")
            for col, header in enumerate(MAIN_HEADERS, start=1):
                value = test_type if header == "Test type" else metrics.get(header)
                cell = ws.cell(row, col, value)
                _paint(
                    cell,
                    input_cell=_is_manual(header) and header != "Triage status",
                    center=header not in {"Test type", "Remarks", "Skip reason"},
                )
                if header == "Triage status" and value:
                    _apply_status(cell)
                if header == "Skip reason":
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                if value is not None or header == "Triage confidence":
                    _apply_number(cell, header)
                if header == "vs previous fail" and isinstance(fail_delta, int) and fail_delta:
                    cell.font = Font(color="176C3A" if fail_delta < 0 else "A12622")
                if header == "Triage owner":
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=False
                    )
                    owner_cells[(section.release, section.stack, test_type)] = cell.coordinate
            _link_test_type(ws, row, test_type)
            row += 1
        row += 1

    _finish_sheet(ws, MAIN_HEADERS, _owner_column_width([run.owner for run in runs]))
    return owner_cells


def _write_progress(
    wb: Workbook,
    runs: list[DashboardRun],
    test_type: str,
    owner_cells: dict[tuple[str, str, str], str],
) -> None:
    ws = wb.create_sheet(test_type)
    ws.sheet_properties.tabColor = "5B7C99"
    ws["A1"] = test_type
    ws["A1"].font = _TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(PROGRESS_HEADERS))
    ws["A2"] = (
        "The blue row is the latest build, including triage that is still open. "
        "Build number opens the Jenkins run."
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(PROGRESS_HEADERS))
    ws["A2"].font = Font(color="667085")

    row = 4
    for section in SECTIONS:
        if test_type not in section.test_types:
            continue
        ws.merge_cells(
            start_row=row, start_column=1, end_row=row, end_column=len(PROGRESS_HEADERS)
        )
        ws.cell(row, 1, section.label).font = _SECTION_FONT
        row += 1
        for col, header in enumerate(PROGRESS_HEADERS, start=1):
            ws.cell(row, col, header)
        _style_header_row(ws, row, len(PROGRESS_HEADERS))
        row += 1
        ordered = _ordered(runs, section.release, section.stack, test_type)
        owner_ref = owner_cells.get((section.release, section.stack, test_type), "")
        if not ordered:
            _write_not_run(ws, row, PROGRESS_HEADERS)
            row += 2
            continue
        for index, run in enumerate(ordered):
            older = ordered[index + 1] if index + 1 < len(ordered) else None
            metrics = _metric_map(run, older)
            fail_delta = metrics.get("_fail_delta")
            for col, header in enumerate(PROGRESS_HEADERS, start=1):
                value = metrics.get(header)
                cell = ws.cell(row, col, value)
                is_owner_ref = header == "Triage owner" and bool(owner_ref)
                _paint(
                    cell,
                    input_cell=_is_manual(header) and not is_owner_ref and header != "Triage status",
                    center=header not in {"Remarks", "Skip reason"},
                    latest=index == 0,
                )
                if header == "Triage status":
                    _apply_status(cell)
                if header == "Skip reason":
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                if header == "Triage owner":
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=False
                    )
                if is_owner_ref:
                    cell.value = f"={DASHBOARD_SHEET}!{owner_ref}"
                elif header == "Build number" and run.jenkins_url:
                    cell.hyperlink = run.jenkins_url
                    cell.font = Font(color="0563C1", underline="single")
                elif header == "vs previous fail" and isinstance(fail_delta, int) and fail_delta:
                    cell.font = Font(color="176C3A" if fail_delta < 0 else "A12622")
                if header == "Triage confidence" or (value is not None and not is_owner_ref):
                    _apply_number(cell, header)
            row += 1
        row += 1

    _finish_sheet(
        ws,
        PROGRESS_HEADERS,
        _owner_column_width([run.owner for run in runs if run.test_type == test_type]),
    )


def _write_history(wb: Workbook, runs: list[DashboardRun]) -> None:
    ws = wb.create_sheet(HISTORY_SHEET)
    ws.append(HISTORY_HEADERS)
    for run in sorted(runs, key=lambda r: (r.release, r.test_type, r.stack, _sdk_value(r))):
        ws.append(
            [
                run.stack,
                run.test_type,
                str(run.sdk_build),
                run.fail,
                run.skip,
                run.passed,
                run.jenkins_run,
                run.triggered,
                run.confidence,
                run.status,
                run.owner,
                run.remarks,
                run.release,
                run.jenkins_url,
                run.skip_reason,
            ]
        )
    ws.sheet_state = "hidden"


def write_dashboard(output_path: Path, runs: list[DashboardRun]) -> Path:
    """Merge these runs into the workbook and rewrite the visible sheets."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged = _merge(_load_runs(output_path), runs)
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    owner_cells = _write_main(wb, merged)
    for test_type in TEST_TYPES:
        _write_progress(wb, merged, test_type, owner_cells)
    _write_history(wb, merged)
    wb.save(output_path)
    return output_path


def _display_rows(
    runs: list[DashboardRun], release: str, stack: str, test_type: str
) -> list[DashboardRun]:
    """Smoke keeps the last five builds. Other types keep the latest Triage Complete."""
    ordered = _ordered(runs, release, stack, test_type)
    if test_type == "smoke":
        return ordered[:HISTORY_LIMIT]
    return [run for run in ordered if _is_complete(run)][:1]


def _write_data_row(
    ws: Worksheet,
    row: int,
    headers: list[str],
    metrics: dict[str, Any],
    *,
    test_type: str | None = None,
    latest: bool = False,
    jenkins_url: str = "",
) -> None:
    fail_delta = metrics.get("_fail_delta")
    for col, header in enumerate(headers, start=1):
        if header == "Test type" and test_type is not None:
            value = test_type
        else:
            value = metrics.get(header)
        cell = ws.cell(row, col, value)
        _paint(
            cell,
            input_cell=_is_manual(header) and header != "Triage status",
            center=header not in {"Test type", "Remarks", "Skip reason"},
            latest=latest,
        )
        if header == "Triage status" and value:
            _apply_status(cell)
        if header == "Skip reason":
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        if header == "Triage owner":
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=False
            )
        if header == "Build number" and jenkins_url:
            cell.hyperlink = jenkins_url
            cell.font = Font(color="0563C1", underline="single")
        elif header == "vs previous fail" and isinstance(fail_delta, int) and fail_delta:
            cell.font = Font(color="176C3A" if fail_delta < 0 else "A12622")
        if value is not None or header == "Triage confidence":
            _apply_number(cell, header)


def _write_mix_progress(
    wb: Workbook,
    runs: list[DashboardRun],
    test_type: str,
    owner_cells: dict[tuple[str, str, str], str],
) -> None:
    """Same layout as the progress sheets. Smoke keeps five builds; others keep Triage Complete."""
    ws = wb.create_sheet(test_type)
    ws.sheet_properties.tabColor = "5B7C99"
    ws["A1"] = test_type
    ws["A1"].font = _TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(PROGRESS_HEADERS))
    if test_type == "smoke":
        blurb = (
            "Last five builds, latest at the top. "
            "Build number opens the Jenkins run."
        )
    else:
        blurb = (
            "Latest Triage Complete build only. "
            "Build number opens the Jenkins run."
        )
    ws["A2"] = blurb
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(PROGRESS_HEADERS))
    ws["A2"].font = Font(color="667085")

    row = 4
    for section in SECTIONS:
        if test_type not in section.test_types:
            continue
        ws.merge_cells(
            start_row=row, start_column=1, end_row=row, end_column=len(PROGRESS_HEADERS)
        )
        ws.cell(row, 1, section.label).font = _SECTION_FONT
        row += 1
        for col, header in enumerate(PROGRESS_HEADERS, start=1):
            ws.cell(row, col, header)
        _style_header_row(ws, row, len(PROGRESS_HEADERS))
        row += 1
        ordered = _ordered(runs, section.release, section.stack, test_type)
        shown = _display_rows(runs, section.release, section.stack, test_type)
        owner_ref = owner_cells.get((section.release, section.stack, test_type), "")
        if not shown:
            if not ordered:
                _write_not_run(ws, row, PROGRESS_HEADERS)
            else:
                metrics = {"Triage owner": _named_owner(ordered)}
                _write_data_row(ws, row, PROGRESS_HEADERS, metrics, latest=False)
                if owner_ref:
                    owner_col = PROGRESS_HEADERS.index("Triage owner") + 1
                    ws.cell(row, owner_col).value = f"={DASHBOARD_SHEET}!{owner_ref}"
            row += 2
            continue
        for index, run in enumerate(shown):
            if test_type == "smoke":
                older = shown[index + 1] if index + 1 < len(shown) else None
            else:
                completed = [item for item in ordered if _is_complete(item)]
                older = completed[1] if len(completed) > 1 else None
            metrics = _metric_map(run, older)
            owner_formula = f"={DASHBOARD_SHEET}!{owner_ref}" if owner_ref else ""
            fail_delta = metrics.get("_fail_delta")
            for col, header in enumerate(PROGRESS_HEADERS, start=1):
                value = metrics.get(header)
                cell = ws.cell(row, col, value)
                is_owner_ref = header == "Triage owner" and bool(owner_ref)
                _paint(
                    cell,
                    input_cell=_is_manual(header)
                    and not is_owner_ref
                    and header != "Triage status",
                    center=header not in {"Remarks", "Skip reason"},
                    latest=index == 0,
                )
                if header == "Triage status":
                    _apply_status(cell)
                if header == "Skip reason":
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                if header == "Triage owner":
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=False
                    )
                if is_owner_ref:
                    cell.value = owner_formula
                elif header == "Build number" and run.jenkins_url:
                    cell.hyperlink = run.jenkins_url
                    cell.font = Font(color="0563C1", underline="single")
                elif header == "vs previous fail" and isinstance(fail_delta, int) and fail_delta:
                    cell.font = Font(color="176C3A" if fail_delta < 0 else "A12622")
                if header == "Triage confidence" or (value is not None and not is_owner_ref):
                    _apply_number(cell, header)
            row += 1
        row += 1

    _finish_sheet(
        ws,
        PROGRESS_HEADERS,
        _owner_column_width([run.owner for run in runs if run.test_type == test_type]),
    )


def write_smoke_mix_dashboard(output_path: Path, runs: list[DashboardRun]) -> Path:
    """Same five sheets as the main dashboard.

    smoke keeps the last five builds. The other test-type sheets keep the latest
    Triage Complete build only. The Dashboard page matches the main workbook.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged = _merge(_load_runs(output_path), runs)
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    owner_cells = _write_main(wb, merged)
    wb[DASHBOARD_SHEET]["A1"] = "RAIL dashboard — smoke progress"
    wb[DASHBOARD_SHEET]["A2"] = (
        "Dashboard shows the latest Triage Complete build. "
        "The smoke sheet lists the last five builds. "
        "ext_smoke, regression, and ext_regression list the latest Triage Complete only. "
        "Type the yellow cells here. Triage owner on the other sheets follows this page."
    )
    for test_type in TEST_TYPES:
        _write_mix_progress(wb, merged, test_type, owner_cells)
    _write_history(wb, merged)
    wb.save(output_path)
    return output_path


def list_indexed_builds(output_path: Path) -> list[dict[str, Any]]:
    path = Path(output_path)
    if not path.is_file():
        return []
    return [
        {
            "release": run.release,
            "stack": run.stack,
            "test_type": run.test_type,
            "sdk_build": run.sdk_build,
            "fail": run.fail,
            "skip": run.skip,
            "passed": run.passed,
            "jenkins_run": run.jenkins_run,
            "triage_confidence": run.confidence,
            "triage_status": run.status,
            "triage_owner": run.owner,
            "remarks": run.remarks,
        }
        for run in _load_runs(path)
    ]
