# RAIL dashboard

Refreshes the RAIL and RAIL SDK triage workbooks from Jenkins and IoT triage notes.

Two Excel files are written into the OneDrive Triage folder (paths are in `config.yaml`):

| File | Command | What it shows |
| --- | --- | --- |
| `RAIL-dashboard.xlsx` | `refresh_dashboard.py` | Latest Triage Complete build on the Dashboard sheet. Last five builds on each test-type sheet, including triage that is still open. |
| `RAIL-dashboard-weekly.xlsx` | `refresh_dashboard.py --weekly` | Same five sheets. Triage Complete builds only, up to five per section. |

## Sheets

Both workbooks have the same visible sheets:

- **Dashboard** — one row per section: the latest Triage Complete build. If none is complete, the result cells stay blank.
- **smoke**, **ext_smoke**, **regression**, **ext_regression** — history for that test type. Newest build is the top row.

Sections, in order:

1. Pkg manager — RAIL (smoke, ext_smoke, regression, ext_regression)
2. Pkg manager — RAIL SDK (same four types)
3. 2026.12 — RAIL (smoke only)
4. 2026.12 — RAIL SDK (smoke only)

A hidden **History** sheet stores the rows so the next refresh can keep typed cells.

## Columns

Dashboard sheet: Test type, Build number, Date, Triage owner, Triage status, Total, Pass, Pass %, Fail, vs previous fail, Fail %, Skip, Skip reason, Triage confidence, Remarks.

Test-type sheets use the same columns without Test type. On those sheets the build number is a link to the Jenkins run.

| Column | Source |
| --- | --- |
| Build number | Jenkins `TEST_BUILD_NUMBER` (the SDK build, not the Jenkins run number) |
| Date | Local date Jenkins started the run |
| Total, Pass, Pass %, Fail, Fail %, Skip | Jenkins test report. Total is pass + fail + skip. |
| vs previous fail | This build's fail count minus the next older build in the same section |
| Skip reason | Jenkins `skippedMessage`, summarized when tests were skipped |
| Triage status | IoT triage notes |
| Triage owner, Triage confidence, Remarks | Typed in Excel and kept on the next refresh |

Triage owner is typed on the Dashboard sheet. The other sheets follow that cell. A new build keeps the owner already typed for that section. ext_smoke starts blank.

Blank confidence is filled in as 100 for smoke and 75 for regression and ext_regression. ext_smoke stays blank until someone types it. A value you type is kept.

## Triage status

Status comes from the IoT report notes, not from a typed cell.

- **Triage Complete** — every note group that exists (test result and app-build result) has at least one completed note.
- **Triage Incomplete** — notes exist, and at least one group is still open.
- **Not started** — IoT has no note for that build.

The weekly workbook keeps Triage Complete rows only. A section with no completed build in the recent window shows `not run`.

## What is collected

Product stacks are `rail` and `rail_sdk`. Test types are `smoke`, `ext_smoke`, `regression`, and `ext_regression`.

Jenkins jobs:

- Package manager: `2026.12-package-manager-sqa-pipelines` for both stacks and all four test types.
- 2026.12: smoke only, for both stacks.

A build is skipped when the test report has suites and none of them is `pytest` (the app-build list ran, the tests did not). Package-manager RAIL smoke SDK 1269 is always left out.

Pkg-manager pipeline schedule (local Jenkins start):

| Test type | Trigger | Weekly workbook keeps |
| --- | --- | --- |
| smoke | Daily | All runs (no day filter) |
| regression | Sunday and Tuesday | **Sunday** (+ early **Monday** spill); **not Tuesday** |
| ext_smoke | Thursday | **Thursday** through **Monday** (same as Fri ext_regression triage window) |
| ext_regression | Friday | **Friday** through **Monday** (you triage when back after Sat/Sun) |

The **daily** workbook lists the latest Jenkins builds with no weekday filter. The **weekly** workbook is the meeting view: **Triage Complete only** (no incomplete, no `not run` placeholders). **Smoke:** one complete row per Wed–Wed week (five weeks). **Regression, ext_smoke, ext_regression:** the **last three** completed triages per section (newest first).

The daily refresh looks at the last five SDK builds per job. The weekly refresh scans more Jenkins history for scheduled test types so it can still find five Triage Complete builds.

## Setup

Python 3.11 or newer.

```powershell
cd C:\Projects\rail_dashboard
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env
```

Edit `.env`:

```
JENKINS_USER=your.user
JENKINS_TOKEN=your_api_token
JENKINS_VERIFY_SSL=false
```

`JENKINS_TOKEN` is a Jenkins API token. IoT report and triage-note requests do not use a key.

Change workbook paths or job URLs in `config.yaml`.

## Refresh

Close the workbook in Excel first. If it is open, the save fails and the file name is left unchanged.

Daily:

```powershell
.\.venv\Scripts\python.exe refresh_dashboard.py
```

Weekly:

```powershell
.\.venv\Scripts\python.exe refresh_dashboard.py --weekly
```

Scheduled-task wrappers (they append a log under `reports\`):

```powershell
scripts\refresh_dashboard.cmd
scripts\refresh_dashboard_weekly.cmd
scripts\refresh_all.cmd
```

**Wednesday 11:00 AM (recommended):** refresh **both** workbooks for the **Wed–Wed report week** (`report_week` in `config.yaml`). Register a weekly task (run once from an elevated prompt if needed):

```powershell
scripts\register_daily_auto.cmd
```

Registers **every day at 11:00 AM** (`refresh_all_scheduled.cmd` → both OneDrive workbooks). Optional (runs when locked): `register_daily_task.cmd` (asks for Windows password).

The script asks for your **Windows password** once so tasks can run **while the screen is locked** (you must still be signed in). If the PC is off at the scheduled time, the job runs after the next sign-in when the PC is on. Close Excel before the run; connect **VPN** so Jenkins is reachable. Logs: `reports\`.

## Triage effort tracker (team Excel)

Separate shared workbook for **effort days** by engineer (`Week 1` … `Week 10+`), week-over-week compare, and manager view. Not auto-filled from Jenkins — the team edits in Excel (OneDrive **AutoSave**).

```powershell
copy triage_tracker\config.yaml   # edit workbook path and engineer names
.\.venv\Scripts\python.exe triage_tracker\create_weekly_sheet.py
scripts\open_triage_effort.cmd
scripts\register_triage_weekly_task.cmd   # new week tab every Wed 00:00
```

Full guide: [`triage_tracker/README.md`](triage_tracker/README.md).

## Layout of this repo

| Path | Role |
| --- | --- |
| `refresh_dashboard.py` | Command-line entry. `--weekly` selects the completed-only workbook. |
| `config.yaml` | Workbook paths, Jenkins jobs, IoT report ids. |
| `triage_tracker/` | Effort workbook generator (`create_weekly_sheet.py`, team config). |
| `src/dashboard_sync.py` | Jenkins fetch and IoT status. |
| `src/dashboard_excel.py` | Reads and writes the workbooks. |
| `src/jenkins_client.py` | Jenkins HTTP client. |
| `src/iot_triage.py` | IoT report and triage-note status. |
| `src/report_week.py` | Wed–Wed report week helpers. |
| `scripts/` | Wrappers for Task Scheduler. |
