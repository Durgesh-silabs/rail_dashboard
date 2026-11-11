#!/usr/bin/env python3
"""Refresh RAIL dashboard Excel workbooks from Jenkins and IoT triage."""

from __future__ import annotations

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
    _load_dotenv(ROOT / ".env")
    from src.config import load_config
    from src.dashboard_sync import refresh_dashboard
    from src.jenkins_client import JenkinsError

    config = load_config()
    print("=== Configuration ===")
    for line in config.summary_lines():
        print(line)
    if not config.dashboard_excel:
        print("Config error: dashboard_excel is not set", file=sys.stderr)
        return 1
    print("=== Refreshing dashboard ===")
    try:
        path = refresh_dashboard(config)
    except PermissionError:
        targets = [str(config.dashboard_excel)]
        if config.dashboard_smoke_excel:
            targets.append(str(config.dashboard_smoke_excel))
        print(
            "An Excel file is open. Close it, then run this again:\n  "
            + "\n  ".join(targets),
            file=sys.stderr,
        )
        return 1
    except JenkinsError as exc:
        print(f"Jenkins error: {exc}", file=sys.stderr)
        return 1
    print(f"Workbook: {path}")
    if config.dashboard_smoke_excel:
        print(f"Smoke workbook: {config.dashboard_smoke_excel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
