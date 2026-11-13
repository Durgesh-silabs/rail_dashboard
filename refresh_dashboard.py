#!/usr/bin/env python3
"""Refresh RAIL dashboard Excel workbooks from Jenkins and IoT triage."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def _load_dotenv(path: Path) -> None:
    try:
        from dotenv import load_dotenv

        load_dotenv(path)
        return
    except ImportError:
        pass
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh the RAIL Excel dashboard.")
    parser.add_argument(
        "--weekly",
        action="store_true",
        help="Write RAIL-dashboard-weekly.xlsx with Triage Complete builds only.",
    )
    args = parser.parse_args()
    _load_dotenv(ROOT / ".env")
    from src.config import load_config
    from src.dashboard_sync import refresh_dashboard, refresh_dashboard_weekly
    from src.jenkins_client import JenkinsError

    config = load_config()
    print("=== Configuration ===")
    for line in config.summary_lines():
        print(line)
    target = config.dashboard_weekly_excel if args.weekly else config.dashboard_excel
    if not target:
        name = "dashboard_weekly_excel" if args.weekly else "dashboard_excel"
        print(f"Config error: {name} is not set", file=sys.stderr)
        return 1
    if args.weekly:
        from src.report_week import format_report_week, report_week_start

        start = report_week_start(cfg=config.report_week)
        print(
            "=== Refreshing weekly dashboard ===",
            f"(report week {format_report_week(start, config.report_week)})",
            flush=True,
        )
    else:
        print("=== Refreshing dashboard ===", flush=True)
    try:
        path = (
            refresh_dashboard_weekly(config) if args.weekly else refresh_dashboard(config)
        )
    except PermissionError:
        print(
            "The Excel file is open. Close it, then run this again:\n"
            f"  {target}",
            file=sys.stderr,
        )
        return 1
    except JenkinsError as exc:
        print(f"Jenkins error: {exc}", file=sys.stderr)
        return 1
    print(f"Workbook: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
