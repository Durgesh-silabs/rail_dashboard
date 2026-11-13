"""Discover Jenkins jobs under the dashboard root and sync latest builds."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, unquote, urljoin

from .config import DashboardSource
from .jenkins_client import JenkinsClient, JenkinsError, api_json_url
from .dashboard_excel import (
    HISTORY_LIMIT,
    WEEKLY_SCHEDULED_TYPES,
    WEEKLY_SCHEDULED_WINDOW,
    DashboardRun,
    include_run_for_dashboard,
    iso_date_from_ms,
    match_test_type,
    summarize_skip_reasons,
    write_dashboard,
)
from .iot_triage import COMPLETE, apply_iot_triage


@dataclass
class DashboardJobTarget:
    test_type: str
    job_url: str
    job_name: str
    full_name: str
    leaf_job: str
    last_completed_build_url: str | None


def _job_api_url(job_url: str) -> str:
    base = job_url if job_url.endswith("/") else job_url + "/"
    return api_json_url(base)


def _child_job_url(parent_url: str, child_name: str) -> str:
    base = parent_url if parent_url.endswith("/") else parent_url + "/"
    return urljoin(base, f"job/{quote(child_name, safe='')}/")


def _normalize_build_url(url: str | None) -> str | None:
    if not url:
        return None
    return url if url.endswith("/") else url + "/"


def _walk_jobs(
    client: JenkinsClient,
    job_url: str,
    *,
    depth: int = 0,
    max_depth: int = 14,
    seen: set[str] | None = None,
) -> list[dict]:
    """Recursively collect job/folder JSON nodes under job_url."""
    if seen is None:
        seen = set()
    key = job_url.rstrip("/")
    if key in seen or depth > max_depth:
        return []
    seen.add(key)

    data = client.get_json(
        _job_api_url(job_url),
        params={
            "tree": (
                "name,fullName,url,_class,"
                "lastCompletedBuild[url,number,result],"
                "jobs[name,fullName,url,_class]"
            )
        },
    )
    if not data.get("url"):
        data["url"] = job_url if job_url.endswith("/") else job_url + "/"

    nodes = [data]
    for child in data.get("jobs") or []:
        child_name = child.get("name") or ""
        child_url = child.get("url") or _child_job_url(job_url, child_name)
        nodes.extend(
            _walk_jobs(
                client,
                child_url,
                depth=depth + 1,
                max_depth=max_depth,
                seen=seen,
            )
        )
    return nodes


def _path_segments(url: str, full_name: str) -> list[str]:
    text = unquote(f"{url}|{full_name}").replace("\\", "/").lower()
    parts: list[str] = []
    for chunk in text.replace("%2f", "/").split("/"):
        parts.extend(p for p in chunk.split("|") if p)
    return parts


def _matches_stack(segments: list[str], stacks: list[str]) -> bool:
    if not stacks:
        return True
    wanted = {s.strip().lower() for s in stacks if s.strip()}
    return any(seg in wanted for seg in segments)


def discover_dashboard_targets(
    client: JenkinsClient,
    root_url: str,
    test_types: Iterable[str],
    stacks: Iterable[str] | None = None,
) -> tuple[list[DashboardJobTarget], list[str]]:
    """
    Find leaf jobs under root_url whose path matches a configured test type.

    When stacks is set (e.g. rail, rail_sdk), only jobs under those folders match.

    Returns (targets, discovered_job_names).
    """
    types = [t.strip() for t in test_types if str(t).strip()]
    if not types:
        raise JenkinsError("dashboard_test_types is empty")
    stack_list = [s.strip() for s in (stacks or []) if str(s).strip()]

    root = root_url if root_url.endswith("/") else root_url + "/"
    # Walk only configured stack folders when set (rail / rail_sdk), not every product.
    walk_roots = [_child_job_url(root, s) for s in stack_list] if stack_list else [root]
    nodes: list[dict] = []
    for walk_root in walk_roots:
        try:
            nodes.extend(_walk_jobs(client, walk_root))
        except JenkinsError as exc:
            if "404" in str(exc):
                continue
            raise

    discovered: list[str] = []
    targets: list[DashboardJobTarget] = []
    seen_urls: set[str] = set()

    for node in nodes:
        url = (node.get("url") or "").strip()
        name = (node.get("name") or "").strip()
        full_name = (node.get("fullName") or name).strip()
        if full_name:
            discovered.append(full_name)
        if not url:
            continue

        node_class = (node.get("_class") or "").lower()
        # Skip folders / multibranch containers
        if "folder" in node_class or "organizationfolder" in node_class:
            continue
        # Prefer jobs that can have builds
        if "workflowjob" not in node_class and "freestyle" not in node_class:
            # Still allow unknown classes if they expose lastCompletedBuild
            if not node.get("lastCompletedBuild") and (node.get("jobs") is not None):
                continue

        segments = _path_segments(url, full_name)
        if not _matches_stack(segments, stack_list):
            continue

        test_type = match_test_type(f"{url}|{full_name}", types)
        if not test_type:
            continue

        # Leaf job name should be the product folder (rail / rail_sdk), not the
        # test-type job name (smoke / regression).
        leaf = next(
            (s for s in reversed(segments) if s in {x.lower() for x in stack_list}),
            name or full_name.split("/")[-1],
        )
        lcb = node.get("lastCompletedBuild") or {}
        build_url = _normalize_build_url(lcb.get("url"))
        if not build_url:
            try:
                detail = client.get_json(
                    _job_api_url(url),
                    params={
                        "tree": "name,fullName,url,_class,lastCompletedBuild[url,number,result]"
                    },
                )
                build_url = _normalize_build_url(
                    (detail.get("lastCompletedBuild") or {}).get("url")
                )
                name = detail.get("name") or name
                full_name = detail.get("fullName") or full_name
                node_class = (detail.get("_class") or node_class).lower()
                if "folder" in node_class:
                    continue
            except JenkinsError:
                build_url = None

        key = url.rstrip("/")
        if key in seen_urls:
            continue
        seen_urls.add(key)
        targets.append(
            DashboardJobTarget(
                test_type=test_type,
                job_url=url if url.endswith("/") else url + "/",
                job_name=name or leaf,
                full_name=full_name,
                leaf_job=leaf,
                last_completed_build_url=build_url,
            )
        )

    targets.sort(key=lambda t: (t.test_type, t.leaf_job, t.full_name))
    discovered = sorted(set(discovered))
    return targets, discovered


def build_matches_dashboard(
    build_url: str,
    *,
    root_url: str,
    test_types: Iterable[str],
) -> str | None:
    """Return matched test_type if build_url is under root and matches a type."""
    root = unquote(root_url.rstrip("/") + "/").lower()
    url = unquote(build_url or "").lower()
    if root.rstrip("/") not in url:
        root_path = root.split(".net/", 1)[-1] if ".net/" in root else root
        url_path = url.split(".net/", 1)[-1] if ".net/" in url else url
        if root_path.rstrip("/") not in url_path:
            return None
    return match_test_type(url, list(test_types))


def _sdk_number(client: JenkinsClient, build_url: str) -> str:
    data = client.get_json(
        api_json_url(build_url),
        params={"tree": "actions[parameters[name,value]]"},
    )
    for action in data.get("actions") or []:
        for param in action.get("parameters") or []:
            if (param.get("name") or "").strip() == "TEST_BUILD_NUMBER":
                return str(param.get("value") or "").strip()
    page = client.get_text(build_url)
    match = re.search(r"TEST_BUILD_NUMBER\s*</\w+>\s*<\w+[^>]*>\s*(\d+)", page)
    if match:
        return match.group(1)
    match = re.search(r"TEST_BUILD_NUMBER.{0,80}?(\d{3,6})", page)
    return match.group(1) if match else ""


def _skip_messages(client: JenkinsClient, build_url: str) -> list[str]:
    report = client.get_json_or_none(
        build_url + "testReport/api/json",
        params={
            "tree": (
                "suites[cases[status,skipped,skippedMessage,errorDetails,"
                "errorStackTrace,stdout,stderr]]"
            )
        },
    )
    if not report:
        return []
    messages: list[str] = []
    for suite in report.get("suites") or []:
        for case in suite.get("cases") or []:
            status = (case.get("status") or "").upper()
            if status != "SKIPPED" and not case.get("skipped"):
                continue
            for key in (
                "skippedMessage",
                "errorDetails",
                "stdout",
                "stderr",
                "errorStackTrace",
            ):
                text = str(case.get(key) or "").strip()
                if text:
                    messages.append(text)
                    break
    return messages


def _triggered_on(client: JenkinsClient, build_url: str, timestamp: object) -> str:
    """Calendar date Jenkins started the run."""
    found = iso_date_from_ms(timestamp)
    if found:
        return found
    info = client.get_json_or_none(api_json_url(build_url), params={"tree": "timestamp"})
    if not info:
        return ""
    return iso_date_from_ms(info.get("timestamp"))


def _tests_were_run(report: dict) -> bool:
    """False when the report has no pytest suite.

    rail_smoke and rail_sdk_smoke are the app-build list. A report that only
    contains that list never executed the tests.
    """
    suites = report.get("suites") or []
    if not suites:
        return True
    return any(str(suite.get("name") or "").strip().lower() == "pytest" for suite in suites)


def collect_recent_runs(
    client: JenkinsClient,
    job_url: str,
    *,
    release: str,
    stack: str,
    test_type: str,
    limit: int = HISTORY_LIMIT,
    window: int = 80,
    weekly_schedule: bool = False,
) -> list[DashboardRun]:
    """Last distinct SDK builds that published a test report, newest first."""
    job = job_url if job_url.endswith("/") else job_url + "/"
    listing = client.get_json(
        api_json_url(job),
        params={"tree": f"builds[number,result,timestamp]{{0,{window}}}"},
    )
    seen: set[str] = set()
    runs: list[DashboardRun] = []
    for build in listing.get("builds") or []:
        if build.get("result") in (None, "ABORTED"):
            continue
        number = int(build["number"])
        build_url = f"{job}{number}/"
        triggered = _triggered_on(client, build_url, build.get("timestamp"))
        if not include_run_for_dashboard(
            test_type,
            triggered,
            build.get("timestamp"),
            weekly_schedule=weekly_schedule,
        ):
            continue
        sdk = _sdk_number(client, build_url)
        if not sdk or sdk in seen:
            continue
        report = client.get_json_or_none(
            build_url + "testReport/api/json",
            params={"tree": "passCount,failCount,skipCount,suites[name]"},
        )
        if not report:
            continue
        if not _tests_were_run(report):
            seen.add(sdk)
            continue
        skip = int(report.get("skipCount") or 0)
        seen.add(sdk)
        runs.append(
            DashboardRun(
                stack=stack,
                test_type=test_type,
                sdk_build=sdk,
                fail=int(report.get("failCount") or 0),
                skip=skip,
                passed=int(report.get("passCount") or 0),
                jenkins_run=number,
                release=release,
                jenkins_url=build_url,
                triggered=triggered,
                skip_reason=summarize_skip_reasons(_skip_messages(client, build_url))
                if skip
                else "",
            )
        )
        if len(runs) >= limit:
            break
    return runs


def _source_jobs(
    client: JenkinsClient, source: DashboardSource
) -> list[tuple[str, str, str]]:
    if source.jobs:
        return [
            (job.url, job.stack, job.test_type)
            for job in source.jobs
            if job.url and job.stack and job.test_type
        ]
    if not source.root:
        return []
    targets, _discovered = discover_dashboard_targets(
        client, source.root, source.test_types, source.stacks
    )
    return [(target.job_url, target.leaf_job, target.test_type) for target in targets]


def match_source(
    build_url: str, sources: Iterable[DashboardSource]
) -> tuple[str, str, str] | None:
    """Return (label, stack, test_type) when a build belongs on the dashboard."""
    url = unquote(build_url or "").lower()
    source_list = list(sources)
    for source in source_list:
        for job in source.jobs:
            job_url = unquote(job.url or "").rstrip("/").lower()
            if job_url and job_url in url:
                return source.label, job.stack, job.test_type
    for source in source_list:
        if not source.root:
            continue
        test_type = build_matches_dashboard(
            build_url, root_url=source.root, test_types=source.test_types
        )
        if not test_type:
            continue
        lowered = url.replace("%2f", "/")
        for stack in source.stacks:
            if f"/job/{stack.lower()}/" in lowered:
                return source.label, stack, test_type
    return None


def _collect_runs(
    config, *, limit: int = HISTORY_LIMIT, weekly_schedule: bool = False
) -> list[DashboardRun]:
    if not config.resolved_dashboard_sources():
        raise JenkinsError("No dashboard sources are configured")
    client = JenkinsClient()
    runs: list[DashboardRun] = []
    for source in config.resolved_dashboard_sources():
        source_runs: list[DashboardRun] = []
        for job_url, stack, test_type in _source_jobs(client, source):
            print(f"Fetching {source.label} / {stack} / {test_type}", flush=True)
            reg_window = (
                WEEKLY_SCHEDULED_WINDOW
                if weekly_schedule and test_type in WEEKLY_SCHEDULED_TYPES
                else 80
            )
            found = collect_recent_runs(
                client,
                job_url,
                release=source.label,
                stack=stack,
                test_type=test_type,
                limit=limit,
                window=reg_window,
                weekly_schedule=weekly_schedule,
            )
            print(f"  {len(found)} build(s)", flush=True)
            source_runs.extend(found)
        if source.iot_report:
            apply_iot_triage(source_runs, source.iot_report)
        runs.extend(source_runs)
    return runs


def refresh_dashboard(config) -> Path:
    """Pull recent Jenkins results into the daily dashboard workbook."""
    if not config.dashboard_excel:
        raise JenkinsError("dashboard_excel is not set")
    runs = _collect_runs(config)
    return write_dashboard(Path(config.dashboard_excel), runs)


def refresh_dashboard_weekly(config) -> Path:
    """Write the weekly workbook: Triage Complete builds only, five per section."""
    weekly = getattr(config, "dashboard_weekly_excel", None)
    if not weekly:
        raise JenkinsError("dashboard_weekly_excel is not set")
    # Fetch like daily; pipeline + backfill filtering happens in write_dashboard.
    runs = _collect_runs(config, limit=50, weekly_schedule=False)
    triage_complete = sum(1 for run in runs if str(run.status or "").strip() == COMPLETE)
    print(f"  Triage Complete rows: {triage_complete}", flush=True)
    return write_dashboard(
        Path(weekly), runs, completed_only=True, report_week=config.report_week
    )
