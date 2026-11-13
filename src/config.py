"""Load dashboard settings from config.yaml."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .report_week import ReportWeekConfig


@dataclass(frozen=True)
class DashboardJob:
    stack: str
    test_type: str
    url: str


@dataclass(frozen=True)
class DashboardSource:
    label: str
    root: str = ""
    stacks: tuple[str, ...] = ()
    test_types: tuple[str, ...] = ()
    jobs: tuple[DashboardJob, ...] = ()
    iot_report: str = ""


@dataclass
class AppConfig:
    dashboard_excel: Path | None = None
    dashboard_smoke_excel: Path | None = None
    dashboard_weekly_excel: Path | None = None
    dashboard_root: str = ""
    dashboard_stacks: tuple[str, ...] = ()
    dashboard_test_types: tuple[str, ...] = ()
    dashboard_sources: tuple[DashboardSource, ...] = ()
    source_path: Path | None = None
    report_week: ReportWeekConfig = ReportWeekConfig()

    def resolved_dashboard_sources(self) -> tuple[DashboardSource, ...]:
        if self.dashboard_sources:
            return self.dashboard_sources
        if not self.dashboard_root:
            return ()
        return (
            DashboardSource(
                label="Pkg manager",
                root=self.dashboard_root,
                stacks=self.dashboard_stacks or ("rail", "rail_sdk"),
                test_types=self.dashboard_test_types,
            ),
        )

    def summary_lines(self) -> list[str]:
        return [
            f"config source:     {self.source_path or '-'}",
            f"dashboard_excel:   {self.dashboard_excel or '-'}",
            f"dashboard_weekly:  {self.dashboard_weekly_excel or '-'}",
            f"dashboard_root:    {self.dashboard_root or '-'}",
            f"dashboard_stacks:  {', '.join(self.dashboard_stacks) or '-'}",
            f"dashboard_types:   {', '.join(self.dashboard_test_types) or '-'}",
            "dashboard_sources: "
            + (
                ", ".join(source.label for source in self.resolved_dashboard_sources())
                or "-"
            ),
            f"report_week:       {self.report_week.label()} (Wed–Wed cycles)",
        ]


def _as_path(value: Any) -> Path | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    return Path(text)


def _as_str_tuple(value: Any) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        parts = [p.strip() for p in value.split(",")]
        return tuple(p for p in parts if p)
    if isinstance(value, (list, tuple)):
        return tuple(str(p).strip() for p in value if str(p).strip())
    return ()


def _as_sources(value: Any) -> tuple[DashboardSource, ...]:
    if not value:
        return ()
    if not isinstance(value, (list, tuple)):
        raise ValueError("dashboard_sources must be a list")
    sources: list[DashboardSource] = []
    for item in value:
        if not isinstance(item, dict):
            raise ValueError("each dashboard source must be a mapping")
        jobs: list[DashboardJob] = []
        for job in item.get("jobs") or []:
            if not isinstance(job, dict):
                raise ValueError("each dashboard job must be a mapping")
            jobs.append(
                DashboardJob(
                    stack=str(job.get("stack") or "").strip(),
                    test_type=str(job.get("test_type") or "").strip(),
                    url=str(job.get("url") or "").strip(),
                )
            )
        sources.append(
            DashboardSource(
                label=str(item.get("label") or "").strip(),
                root=str(item.get("root") or "").strip(),
                stacks=_as_str_tuple(item.get("stacks")),
                test_types=_as_str_tuple(item.get("test_types")),
                jobs=tuple(jobs),
                iot_report=str(item.get("iot_report") or "").strip(),
            )
        )
    return tuple(sources)


def load_config(path: Path | None = None) -> AppConfig:
    """Load config.yaml from the project root when path is omitted."""
    import yaml

    root = Path(__file__).resolve().parents[1]
    source = Path(path) if path else root / "config.yaml"
    if not source.is_file():
        raise FileNotFoundError(f"config.yaml not found: {source}")
    raw = yaml.safe_load(source.read_text(encoding="utf-8")) or {}
    if not isinstance(raw, dict):
        raise ValueError("config.yaml must be a mapping")
    return AppConfig(
        dashboard_excel=_as_path(raw.get("dashboard_excel")),
        dashboard_smoke_excel=_as_path(raw.get("dashboard_smoke_excel")),
        dashboard_weekly_excel=_as_path(raw.get("dashboard_weekly_excel")),
        dashboard_root=str(raw.get("dashboard_root") or "").strip(),
        dashboard_stacks=_as_str_tuple(raw.get("dashboard_stacks")),
        dashboard_test_types=_as_str_tuple(raw.get("dashboard_test_types")),
        dashboard_sources=_as_sources(raw.get("dashboard_sources")),
        source_path=source,
        report_week=ReportWeekConfig.from_mapping(raw.get("report_week")),
    )
