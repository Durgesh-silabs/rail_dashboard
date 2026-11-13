"""Wednesday–Wednesday reporting weeks for the weekly workbook."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta

_WEEKDAY_NAMES = {
    "monday": 0,
    "tuesday": 1,
    "wednesday": 2,
    "thursday": 3,
    "friday": 4,
    "saturday": 5,
    "sunday": 6,
}


@dataclass(frozen=True)
class ReportWeekConfig:
    weekday: int = 2
    hour: int = 11
    minute: int = 0

    @classmethod
    def from_mapping(cls, raw: object) -> ReportWeekConfig:
        if not isinstance(raw, dict):
            return cls()
        name = str(raw.get("weekday") or "wednesday").strip().lower()
        weekday = _WEEKDAY_NAMES.get(name, 2)
        return cls(
            weekday=weekday,
            hour=int(raw.get("hour") or 11),
            minute=int(raw.get("minute") or 0),
        )

    def label(self) -> str:
        names = {v: k.title() for k, v in _WEEKDAY_NAMES.items()}
        day = names.get(self.weekday, "Wednesday")
        return f"{day} {self.hour:02d}:{self.minute:02d} local"


def report_week_start(
    at: datetime | None = None, cfg: ReportWeekConfig | None = None
) -> datetime:
    """Start (inclusive) of the reporting week that contains `at`."""
    cfg = cfg or ReportWeekConfig()
    at = at or datetime.now().replace(microsecond=0)
    anchor = at.replace(hour=cfg.hour, minute=cfg.minute, second=0, microsecond=0)
    days_since = (at.weekday() - cfg.weekday) % 7
    start = anchor - timedelta(days=days_since)
    if at < start:
        start -= timedelta(days=7)
    return start


def report_week_end(start: datetime, cfg: ReportWeekConfig | None = None) -> datetime:
    """End (exclusive) of the reporting week that begins at `start`."""
    _ = cfg
    return start + timedelta(days=7)


def report_week_start_for_triggered(
    triggered: str, cfg: ReportWeekConfig | None = None
) -> datetime | None:
    """Map a Jenkins calendar date (YYYY-MM-DD) to its reporting week start."""
    text = str(triggered or "").strip()[:10]
    if not text:
        return None
    try:
        day = datetime.strptime(text, "%Y-%m-%d")
    except ValueError:
        return None
    return report_week_start(day + timedelta(hours=12), cfg)


def format_report_week(start: datetime, cfg: ReportWeekConfig | None = None) -> str:
    end = report_week_end(start, cfg) - timedelta(seconds=1)
    return f"{start.date().isoformat()} – {end.date().isoformat()}"
