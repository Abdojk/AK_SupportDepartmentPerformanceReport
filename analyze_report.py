import pandas as pd
import json
import numpy as np
from datetime import datetime

# ── Configuration ──────────────────────────────────────────────────────────
REPORT_DATE = "2026-02-13"
PLANNING_FENCE = "2026-03-15"

AGENTS = [
    'Rebecca Estephan', 'Mennatullah El Bahr', 'Abdo Khoury',
    'Jana Sweid', 'Fadi Hanna', 'Georges Mouaikel', 'Raji Aoun'
]

VALID_STATUSES = [
    'Estimation', 'In Process with Microsoft', 'In Progress',
    'Not Started', 'Requirement Gathering',
    'Requirement Gathering/Pending Briefing',
    'Researching', 'Solution Design'
]

OPEN_FILE  = "/home/user/AK_SupportDepartmentPerformanceReport/data/Support Unit Open Cases 2-13-2026 8-35-49 AM.xlsx"
RESOLVED_FILE = "/home/user/AK_SupportDepartmentPerformanceReport/data/Support Unit Resolved Cases 2-13-2026 8-36-33 AM.xlsx"

# ── Helper: make values JSON-serialisable ──────────────────────────────────
def sanitize(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        return float(val)
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d %H:%M:%S") if not pd.isna(val) else None
    return val

def sanitize_list(series):
    return [sanitize(v) for v in series]

# ══════════════════════════════════════════════════════════════════════════
# 1. READ BOTH FILES
# ══════════════════════════════════════════════════════════════════════════
print("=" * 80)
print("READING EXCEL FILES")
print("=" * 80)

df_open = pd.read_excel(OPEN_FILE, engine="openpyxl")
df_resolved = pd.read_excel(RESOLVED_FILE, engine="openpyxl")

print(f"\nOpen cases file  : {len(df_open)} rows")
print(f"Resolved cases file: {len(df_resolved)} rows")

# ══════════════════════════════════════════════════════════════════════════
# 2. PRINT ALL COLUMN NAMES
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("COLUMN NAMES — OPEN CASES")
print("=" * 80)
for i, col in enumerate(df_open.columns, 1):
    print(f"  {i:>3}. {col}")

print("\n" + "=" * 80)
print("COLUMN NAMES — RESOLVED CASES")
print("=" * 80)
for i, col in enumerate(df_resolved.columns, 1):
    print(f"  {i:>3}. {col}")

# ══════════════════════════════════════════════════════════════════════════
# 3. FILTER OPEN CASES BY OWNER
# ══════════════════════════════════════════════════════════════════════════
total_before_owner_filter = len(df_open)
df_open = df_open[df_open['Owner'].isin(AGENTS)].copy()
print(f"\n--- Owner filter: {total_before_owner_filter} -> {len(df_open)} rows (excluded {total_before_owner_filter - len(df_open)})")

# ══════════════════════════════════════════════════════════════════════════
# 4. RENAME Raji Aoun -> Unassigned Cases
# ══════════════════════════════════════════════════════════════════════════
df_open['Owner'] = df_open['Owner'].replace('Raji Aoun', 'Unassigned Cases')
print("--- Renamed 'Raji Aoun' -> 'Unassigned Cases'")

# ══════════════════════════════════════════════════════════════════════════
# 5. STATUS REASON — UNIQUE VALUES & COUNTS (before status filter)
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("STATUS REASON — ALL UNIQUE VALUES & COUNTS (after owner filter, before status filter)")
print("=" * 80)
status_counts = df_open['Status Reason'].value_counts(dropna=False)
for status, count in status_counts.items():
    print(f"  {str(status):50s} : {count}")
print(f"  {'TOTAL':50s} : {status_counts.sum()}")

# ══════════════════════════════════════════════════════════════════════════
# 6. APPLY STATUS FILTER
# ══════════════════════════════════════════════════════════════════════════
total_before_status = len(df_open)
df_open = df_open[df_open['Status Reason'].isin(VALID_STATUSES)].copy()
total_after_status = len(df_open)

print(f"\n--- Status filter: {total_before_status} -> {total_after_status} rows (excluded {total_before_status - total_after_status})")

# ══════════════════════════════════════════════════════════════════════════
# 7. PRINT TOTALS SUMMARY
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("FILTER SUMMARY")
print("=" * 80)
print(f"  Total before owner filter  : {total_before_owner_filter}")
print(f"  After owner filter         : {total_before_status}")
print(f"  After status filter (final): {total_after_status}")
print(f"  Excluded by status         : {total_before_status - total_after_status}")

# ══════════════════════════════════════════════════════════════════════════
# 8. PER-AGENT DETAIL — OPEN CASES
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("PER-AGENT DETAIL — OPEN CASES")
print("=" * 80)

agent_details = {}
for agent in sorted(df_open['Owner'].unique()):
    agent_df = df_open[df_open['Owner'] == agent]
    detail = {
        "name": agent,
        "case_count": int(len(agent_df)),
        "priorities": sanitize_list(agent_df['Priority'].tolist()),
        "ages": sanitize_list(agent_df['Age'].tolist()) if 'Age' in agent_df.columns else [],
        "due_dates": sanitize_list(agent_df['Due Date'].tolist()) if 'Due Date' in agent_df.columns else [],
        "initial_estimation_durations": sanitize_list(agent_df['Initial Estimation Duration'].tolist()) if 'Initial Estimation Duration' in agent_df.columns else [],
        "status_reasons": sanitize_list(agent_df['Status Reason'].tolist()),
        "customers": sanitize_list(agent_df['Customer'].tolist()) if 'Customer' in agent_df.columns else [],
    }
    agent_details[agent] = detail

    print(f"\n  AGENT: {agent}")
    print(f"    Case count : {detail['case_count']}")
    print(f"    Priorities : {detail['priorities']}")
    print(f"    Ages       : {detail['ages']}")
    print(f"    Due Dates  : {detail['due_dates']}")
    print(f"    Init Est.  : {detail['initial_estimation_durations']}")
    print(f"    Statuses   : {detail['status_reasons']}")
    print(f"    Customers  : {detail['customers']}")

# ══════════════════════════════════════════════════════════════════════════
# 9. OVERALL STATS — OPEN CASES
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("OVERALL STATS — OPEN CASES")
print("=" * 80)

total_cases = len(df_open)

# Average age
if 'Age' in df_open.columns:
    avg_age = df_open['Age'].mean()
    avg_age = round(float(avg_age), 2) if not pd.isna(avg_age) else 0
else:
    avg_age = 0

# Priority distribution
priority_dist = df_open['Priority'].value_counts(dropna=False).to_dict()
priority_dist = {str(k): int(v) for k, v in priority_dist.items()}

# Customer distribution (all customers with counts)
if 'Customer' in df_open.columns:
    customer_dist = df_open['Customer'].value_counts(dropna=False).to_dict()
    customer_dist = {str(k): int(v) for k, v in customer_dist.items()}
else:
    customer_dist = {}

print(f"  Total cases     : {total_cases}")
print(f"  Average age     : {avg_age}")
print(f"  Priority dist.  : {json.dumps(priority_dist, indent=4)}")
print(f"  Customer dist.  :")
for cust, cnt in sorted(customer_dist.items(), key=lambda x: -x[1]):
    print(f"    {cust:50s} : {cnt}")

# ══════════════════════════════════════════════════════════════════════════
# 10. RESOLVED CASES — PER AGENT DURATIONS
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("RESOLVED CASES — PER-AGENT DURATIONS")
print("=" * 80)

df_resolved_filtered = df_resolved[df_resolved['Owner'].isin(AGENTS)].copy()
df_resolved_filtered['Owner'] = df_resolved_filtered['Owner'].replace('Raji Aoun', 'Unassigned Cases')

print(f"  Resolved cases (filtered by agents): {len(df_resolved_filtered)}")

resolved_agent_details = {}
for agent in sorted(df_resolved_filtered['Owner'].unique()):
    agent_df = df_resolved_filtered[df_resolved_filtered['Owner'] == agent]

    init_est_col = 'Initial Estimation Duration' if 'Initial Estimation Duration' in agent_df.columns else None
    actual_dur_col = 'Actual Duration' if 'Actual Duration' in agent_df.columns else None

    init_sum = float(agent_df[init_est_col].fillna(0).sum()) if init_est_col else 0
    actual_sum = float(agent_df[actual_dur_col].fillna(0).sum()) if actual_dur_col else 0

    resolved_agent_details[agent] = {
        "name": agent,
        "resolved_count": int(len(agent_df)),
        "initial_estimation_duration_sum": round(init_sum, 2),
        "actual_duration_sum": round(actual_sum, 2),
    }

    print(f"\n  AGENT: {agent}")
    print(f"    Resolved count              : {len(agent_df)}")
    print(f"    Initial Estimation Dur. SUM : {round(init_sum, 2)}")
    print(f"    Actual Duration SUM         : {round(actual_sum, 2)}")

# ══════════════════════════════════════════════════════════════════════════
# 11. STRUCTURED JSON OUTPUT
# ══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("STRUCTURED JSON OUTPUT")
print("=" * 80)

output = {
    "report_date": REPORT_DATE,
    "planning_fence": PLANNING_FENCE,
    "open_cases": {
        "total_before_owner_filter": total_before_owner_filter,
        "total_after_owner_filter": total_before_status,
        "total_after_status_filter": total_after_status,
        "excluded_by_status_filter": total_before_status - total_after_status,
        "average_age": avg_age,
        "priority_distribution": priority_dist,
        "customer_distribution": customer_dist,
        "agents": agent_details,
    },
    "resolved_cases": {
        "total_filtered": len(df_resolved_filtered),
        "agents": resolved_agent_details,
    },
}

print(json.dumps(output, indent=2, default=str))

print("\n" + "=" * 80)
print("DONE")
print("=" * 80)
