"""Fill dashboard triage status from the IoT report viewer.

The viewer stores pass/fail on the report, and the Complete / Incomplete chip
on the triage notes service that the report uses. A build is complete when
every note group that exists (test results and app builds) is marked complete.
"""

from __future__ import annotations

from collections import defaultdict

import requests

from .jenkins_client import _resolve_verify
from .dashboard_excel import DashboardRun

_REPORTS = "https://iot-reports.silabs.net"
_TRIAGE = "https://triage.silabs.net"
_SQA_ORG = "60143082100035c608bfd3fd"
_REPORT_TYPES = ("testResult", "appBuildResult")
COMPLETE = "Triage Complete"
INCOMPLETE = "Triage Incomplete"
NOT_STARTED = "Not started"


def apply_iot_triage(runs: list[DashboardRun], report_id: str) -> None:
    """Set Triage status from the IoT report for this dashboard source."""
    report_id = (report_id or "").strip()
    if not report_id or not runs:
        return
    branch = _branch_name(report_id)
    if not branch:
        print(f"  IoT report {report_id}: no branch, status left as typed", flush=True)
        return
    grouped: dict[tuple[str, str], list[DashboardRun]] = defaultdict(list)
    for run in runs:
        grouped[(run.stack, run.test_type)].append(run)
    complete = incomplete = not_started = 0
    for (stack, test_type), group in grouped.items():
        builds = {_build_number(run.sdk_build) for run in group}
        builds.discard(None)
        if not builds:
            continue
        flags = _completion_flags(branch, stack, test_type, builds)
        for run in group:
            number = _build_number(run.sdk_build)
            status = _status_label(flags.get(number, {})) if number is not None else ""
            if not status:
                status = NOT_STARTED
            run.status = status
            if status == COMPLETE:
                complete += 1
            elif status == INCOMPLETE:
                incomplete += 1
            else:
                not_started += 1
    print(
        f"  IoT triage ({branch}): {complete} complete, "
        f"{incomplete} in progress, {not_started} not started",
        flush=True,
    )


def _branch_name(report_id: str) -> str:
    url = f"{_REPORTS}/api/v1/{_SQA_ORG}/reportAndType/{report_id}"
    payload = _get_json(url)
    if not isinstance(payload, dict):
        return ""
    report = payload.get("report") or {}
    data = report.get("reportData") or {}
    for section in data.get("sections") or []:
        branch = str(section.get("branchName") or "").strip()
        if branch:
            return branch
    return ""


def _completion_flags(
    branch: str, stack: str, test_type: str, builds: set[int]
) -> dict[int, dict[str, list[bool]]]:
    """build number -> report type -> completed flags on that type's notes."""
    flags: dict[int, dict[str, list[bool]]] = defaultdict(lambda: defaultdict(list))
    wanted = ",".join(str(number) for number in sorted(builds))
    for report_type in _REPORT_TYPES:
        page = 1
        seen = 0
        while page <= 10:
            payload = _get_json(
                f"{_TRIAGE}/api/v3/notes/",
                params={
                    "branchName": branch,
                    "stack": stack.upper(),
                    "reportType": report_type,
                    "resultType": test_type.lower(),
                    "buildNumberIn": f"[{wanted}]",
                    "includeJiraData": "False",
                    "page": page,
                    "pageSize": 100,
                },
            )
            if not isinstance(payload, dict):
                break
            rows = payload.get("data") or []
            for row in rows:
                number = _build_number(row.get("buildNumber"))
                if number not in builds:
                    continue
                flags[number][report_type].append(bool(row.get("completed")))
            seen += len(rows)
            count = int(payload.get("count") or 0)
            if not rows or seen >= count:
                break
            page += 1
    return flags


def _status_label(by_type: dict[str, list[bool]]) -> str:
    """Match the IoT chip: a note group is complete when any of its notes is."""
    present = [any(done) for done in by_type.values() if done]
    if not present:
        return ""
    if all(present):
        return COMPLETE
    return INCOMPLETE


def _build_number(value) -> int | None:
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def _get_json(url: str, params: dict | None = None) -> dict | None:
    try:
        response = requests.get(
            url, params=params, timeout=45, verify=_resolve_verify()
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        print(f"  IoT request failed: {exc}", flush=True)
        return None
    try:
        payload = response.json()
    except ValueError:
        return None
    return payload if isinstance(payload, dict) else None
