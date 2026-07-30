================================================================================
  README — generate_dashboard.py
  Weekly Output Dashboard — Auto Data Injection Script
================================================================================

LOCATION
--------
  Script : .agents/skills/Ops_weekly_output_report/scripts/generate_dashboard.py
  Output : .agents/skills/Ops_weekly_output_report/reports/weekly_output_dashboard.html

WHAT THIS SCRIPT DOES
----------------------
generate_dashboard.py is the data pipeline that brings the Weekly Output
Dashboard to life. Every time you run it, the script:

  1. Reads the latest Workload Excel file (Jobs + Tasks sheets)
  2. Reads the Attendance Excel file (FACT_Attendance_Daily, Req_Leave,
     Req_OT, DIM_Employee sheets)
  3. Computes all metrics the dashboard requires:
       • customer_weekly      — completed jobs / sqm / units / layouts per
                                week × team × customer (5-week trend)
       • employee_weekly      — per-employee task counts and volumes per week
       • rework_team_weekly   — design vs. review-fix tasks per team per week
       • passthrough_weekly   — jobs sent / intra-week / pass-week / backlog
       • jobtype_weekly       — job type breakdown (Amd./Est./Det.) per week
       • job_scatter          — per-job bubble data: sqm, units, hours invested
       • leadtime_weekly      — average lead-time per customer per week
       • sla_active_jobs      — current active job counts per customer (SLA widget)
       • sla_distribution     — SLA bucket breakdown (0d – 5d)
       • backlog_forecast     — capacity vs. backlog clearance forecast table
       • headcount_weekly     — unique designers per customer per week
       • CUSTOMER_TARGETS     — weekly KPI targets (jobs or layouts) per customer
       • RAW (Gantt)          — per-employee task bars and leave blocks
  4. Embeds all computed data into the HTML dashboard template
  5. Writes the fully updated dashboard to the reports/ folder

WHICH DASHBOARD DOES THIS FEED?
---------------------------------
  Template (source HTML):
    .agents/skills/Ops_weekly_output_report/template/weekly_output_dashboard.html

  Output (open this file in your browser):
    .agents/skills/Ops_weekly_output_report/reports/weekly_output_dashboard.html

  The output file is a fully self-contained HTML file — open it directly in
  any web browser. No internet connection, login, or web server required.

SOURCE DATA FILES
-----------------
  Workload Excel:
    G:\My Drive\Dữ liệu nhân sự\Workload\Construction Team\
    Report_Past Month.xlsx
    (script auto-finds the latest "Report_Past Month*.xlsx" in that folder)

  Attendance Excel:
    G:\My Drive\Dữ liệu nhân sự\Data\Timesheet\HR_Fact_Attendance.xlsx

  KPI Targets JSON (optional, for CUSTOMER_TARGETS):
    G:\My Drive\Dữ liệu nhân sự\.agents\skills\workload_daily_report\
    references\Customer KPI Mapping.json

HOW TO RUN
----------
  Basic (auto-detects previous complete week):
    python generate_dashboard.py

  Specify a custom reporting week end date:
    python generate_dashboard.py --week-end 2026-07-20

  Override the workload or attendance file:
    python generate_dashboard.py --workload "path\to\file.xlsx"
    python generate_dashboard.py --attendance "path\to\attendance.xlsx"

REQUIREMENTS
------------
  Python 3.10+
  pip install pandas openpyxl python-dateutil numpy

REPORTING WEEK LOGIC
--------------------
  • W0 = the last completed Mon–Sun week (ends on the most recent Sunday)
  • The dashboard shows 15+ historical weeks automatically (from Apr 2026
    cutoff date through W0) — all weeks where data exists
  • Gantt chart default view covers the 2 most recent weeks

DATA REFRESH WORKFLOW
---------------------
  1. Export the latest Workload Excel from the Construction Team tracker
     and save it in the Workload\Construction Team\ folder.
  2. Ensure HR_Fact_Attendance.xlsx is up to date.
  3. Run:  python generate_dashboard.py
  4. Open:  reports\weekly_output_dashboard.html  in your browser.
  5. The dashboard will reflect all data changes automatically.

NOTES
-----
  • Users listed in EXCLUDE_USERS (management, IT, HR) are excluded from
    all workload metrics automatically.
  • The script overrides completion status for RES Engineering & Prime
    Design jobs: a job is considered complete when its latest Check task
    is marked Complete (matching the ops team's counting convention).
  • CUSTOMER_TARGETS (weekly KPI targets per customer) are loaded from
    Customer KPI Mapping.json. If that file is unavailable, the targets
    fall back to the hardcoded values in the HTML template.
  • All times are treated as local time (no timezone conversion).

================================================================================
