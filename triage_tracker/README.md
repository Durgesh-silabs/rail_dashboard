# Triage effort tracker (separate from dashboard refresh)

Manual **editable** workbook for managers: who owns which suites, effort days, blockers.

## How it differs from `RAIL-dashboard.xlsx`

| Dashboard refresh | Triage effort tracker |
| ------------------- | ------------------- |
| Auto from Jenkins + IoT | **You edit** during the week |
| Pass/fail totals | **Suite splits**, effort days, comments |
| Meeting status | **Effort & week-over-week compare** |

## Workbook

- Path: `config.yaml` → `workbook` (same OneDrive folder as dashboards).
- **Tabs `Week 1` … `Week 10`** (minimum 10) plus new weeks as the calendar advances.
- Row 1 on each tab shows the **Wed–Wed dates** for that week number.
- **Report week:** Wednesday **00:00** through the following Tuesday (same Wed–Wed cycle as `RAIL-dashboard-weekly.xlsx`).
- No Summary or IOTREQ sheets — everything lives on the week tabs.

### Weekly tab (per person, per week)

Each engineer gets **8 task rows** by default: **5× smoke**, **1× regression**, **1× ext_smoke**, **1× ext_regression**.  
Names in `no_smoke_engineers` skip smoke (e.g. Ankita, SVN).

| Column | Use |
|--------|-----|
| **Started** / **Target complete** | Multi-day work (not one day only) |
| **Task** | smoke 1/5 … smoke 5/5, regression, etc. |
| **Pipeline** | smoke, regression, ext_smoke, ext_regression (seeded; used for totals) |
| **Failures assigned** | Count at handoff or when you start the row |
| **Effort (days)** | **Number** (e.g. `0.5`, `2`) — drives totals and compare |
| **Status / Comments** | PR review, validation pending, stuck, … (wrapped, taller rows) |

Grid **borders** match the RAIL dashboard. **Test suite(s)** and **Comments** use extra width, wrap, and row height.

At the **bottom of each week tab**:

| Pipeline | This week | Previous week | Change |
|----------|-----------|---------------|--------|
| Smoke, Regression, … | SUMIF on this tab | SUMIF on **`Week N-1`** | This − previous |

If the previous week tab is missing, **Previous week** is blank.

Edit engineer names and task counts in `config.yaml`.

**Week numbering** (`week_tabs` in `config.yaml`):

| Setting | Meaning |
|---------|---------|
| `min_tabs: 10` | Always keep at least **Week 1** … **Week 10** sheets |
| `week_1_start` | Calendar date in the first report week (adjust so “this” cycle is the week you want) |

After week 10, **Week 11**, … appear automatically when the calendar moves forward (Wednesday job).

## Shared editing — auto-save for everyone

The workbook lives on **OneDrive** (`config.yaml` → `workbook`). There is no extra server: **Excel + OneDrive** sync changes to the cloud so others see updates.

### One-time setup (you / admin)

1. Put the file in a folder the team can access (e.g. `RAIL-dashboard` on OneDrive).
2. **Share** that folder (or `RAIL-triage-effort.xlsx`) with Ankita, Adithya, SVN, Durgesh, Vijay, Sai Krishna, and your manager (edit access).
3. Tell everyone to open **that OneDrive file**, not a copy in `Downloads` or email attachment.

### Each person, every time they edit

1. Open the workbook:
   ```powershell
   scripts\open_triage_effort.cmd
   ```
   Or open from OneDrive in the browser → **Open in Desktop App** / **Excel for the web**.
2. Turn **AutoSave** **ON** (top-left in Excel when the file is on OneDrive).  
   - Changes save to the cloud automatically (every few seconds).  
   - Others with the file open see updates without closing and reopening (co-authoring).
3. **Excel for the web** always auto-saves — good option if desktop AutoSave is off.

### What you will see

| Situation | What happens |
|-----------|----------------|
| You type in a yellow cell, AutoSave ON | Saved to OneDrive; teammates see it shortly |
| Manager has the file open | They refresh or see live updates (co-authoring) |
| Wednesday midnight task adds a new week tab | Close Excel before that run **or** use co-authoring; if save fails, check `reports\triage_effort_weekly.log` |

The Python scripts only **add week tabs** and refresh the **compare** block at the bottom. They do **not** replace your row data unless you use `--recreate-week`.

## Commands

From repo root (venv already has `openpyxl` + `PyYAML`):

```powershell
cd C:\Users\dugupta\Downloads\rail_dashboard
scripts\open_triage_effort.cmd
.\.venv\Scripts\python.exe triage_tracker\create_weekly_sheet.py
```

If the tab for the current report week already exists, the script **refreshes the compare block** at the bottom and **removes** legacy Summary/IOTREQ sheets on save.

**Automatic new tab — every Wednesday at 12:00 AM:**

```powershell
scripts\register_triage_weekly_task.cmd
```

Runs `scripts\triage_effort_weekly.cmd` (log: `reports\triage_effort_weekly.log`).

Recreate the current report week tab (new engineer rows + layout):

```powershell
.\.venv\Scripts\python.exe triage_tracker\create_weekly_sheet.py --recreate-week
```

## Presentation tips for your manager

1. Open the **current date tab** and scroll to the **compare** table.
2. Sort by **SDK build** + **pipeline** to show regression burn-down.
3. **Failures assigned** + **Status** for progress; **Change** column shows more/less effort vs last week.
4. **Status** values: In progress | Blocked | PR review | Validation pending | Done.
