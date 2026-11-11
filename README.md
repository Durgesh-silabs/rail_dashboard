# RAIL dashboard

Refresh the RAIL Excel triage dashboards from Jenkins and IoT triage notes.

## Workbooks

- `RAIL-dashboard.xlsx` — main dashboard (latest Triage Complete on the Dashboard sheet; last five builds on each test-type sheet)
- `RAIL-dashboard-smoke.xlsx` — same five sheets; smoke keeps the last five builds, other test types keep the latest Triage Complete only

Both paths are set in `config.yaml` (OneDrive Triage folder by default).

## Setup

```powershell
cd C:\Projects\rail_dashboard
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env
```

Edit `.env` with your Jenkins user and API token.

## Refresh

Close the Excel files first, then:

```powershell
.\.venv\Scripts\python.exe refresh_dashboard.py
```

Or use the scheduled-task wrapper:

```powershell
scripts\refresh_dashboard.cmd
```

## What gets updated

**From Jenkins:** build number (`TEST_BUILD_NUMBER`), date, pass / fail / skip, skip reason, Jenkins link.

**From IoT triage notes:** Triage status (`Triage Complete`, `Triage Incomplete`, `Not started`).

**Kept as typed:** Triage owner, remarks, and a confidence you changed. Blank confidence defaults to 100% for smoke and 75% for regression / ext_regression.
