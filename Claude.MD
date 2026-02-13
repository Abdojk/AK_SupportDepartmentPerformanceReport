# PRMT01-E v2.3.0 - STATUS REASON FILTER UPDATE

## CRITICAL CHANGE: Mandatory Status Reason Filtering

### CHANGE SUMMARY
Add mandatory filtering to ONLY include cases with specific active Status Reasons.
This ensures capacity calculations reflect only cases actively being worked on or ready to start.

---

## NEW FILTERING REQUIREMENT (STEP 1 - BEFORE ALL CALCULATIONS)

### LOCATION IN PROMPT
Add this filtering step IMMEDIATELY AFTER the department agent filtering and 
BEFORE any data quality audit or KPI calculations.

### FILTER SPECIFICATION

**Column:** Status Reason (Column J in Excel)

**ONLY Include Cases With These Status Reasons:**
1. Estimation
2. In Process with Microsoft
3. In Progress
4. Not Started
5. Requirement Gathering
6. Requirement Gathering/Pending Briefing
7. Researching
8. Solution Design

**Exclude ALL Other Status Reasons**, including but not limited to:
- Pending Client Confirmation
- Pending Dev
- Pending Customer
- Pending Delivery Date
- Waiting on Customer
- Waiting on Third Party
- On Hold
- Any other status not in the approved list

---

## UPDATED CODE IMPLEMENTATION

### Step 1: Filter for Department Agents (EXISTING)

```python
# Read Excel file
open_df = pd.read_excel(file_path)

# Filter for department agents FIRST
DEPARTMENT_AGENTS = [
    'Rebecca Estephan',
    'Mennatullah El Bahr', 
    'Abdo Khoury',
    'Jana Sweid',
    'Fadi Hanna',
    'Georges Mouaikel',
    'Raji Aoun'
]

open_df = open_df[open_df['Owner'].isin(DEPARTMENT_AGENTS)].copy()

# Rename Raji Aoun to Unassigned Cases
open_df['Owner'] = open_df['Owner'].replace('Raji Aoun', 'Unassigned Cases')
```

### Step 2: Filter for Active Status Reasons (NEW - MANDATORY)

```python
# ═══════════════════════════════════════════════════════════════
# STEP 2: FILTER FOR ACTIVE STATUS REASONS ONLY
# ═══════════════════════════════════════════════════════════════

# Define approved active status reasons (cases actively being worked)
ACTIVE_STATUS_REASONS = [
    'Estimation',
    'In Process with Microsoft',
    'In Progress',
    'Not Started',
    'Requirement Gathering',
    'Requirement Gathering/Pending Briefing',
    'Researching',
    'Solution Design'
]

# Store total before filter for reporting
total_before_status_filter = len(open_df)

# Filter for active status reasons only
open_df = open_df[open_df['Status Reason'].isin(ACTIVE_STATUS_REASONS)].copy()

total_after_status_filter = len(open_df)
excluded_by_status = total_before_status_filter - total_after_status_filter

print(f"\n╔═══════════════════════════════════════════════════════════════╗")
print(f"║ STATUS REASON FILTERING                                      ║")
print(f"╚═══════════════════════════════════════════════════════════════╝\n")
print(f"Total Cases Before Status Filter:  {total_before_status_filter}")
print(f"Total Cases After Status Filter:   {total_after_status_filter}")
print(f"Cases Excluded (Non-Active):       {excluded_by_status}")
print(f"Exclusion Rate:                    {(excluded_by_status/total_before_status_filter*100):.1f}%")
print(f"\nActive Status Reasons Included:")
for status in ACTIVE_STATUS_REASONS:
    count = len(open_df[open_df['Status Reason'] == status])
    if count > 0:
        print(f"  • {status:<40}: {count:>3} cases")
```

### Step 3: Continue with Data Quality Audit (EXISTING)

```python
# ═══════════════════════════════════════════════════════════════
# STEP 3: DATA QUALITY AUDIT (on filtered dataset)
# ═══════════════════════════════════════════════════════════════

total_cases = len(open_df)  # Now reflects filtered count
print(f"\n╔═══════════════════════════════════════════════════════════════╗")
print(f"║ DATA QUALITY AUDIT (ACTIVE CASES ONLY)                       ║")
print(f"╚═══════════════════════════════════════════════════════════════╝\n")
print(f"Total Active Cases Analyzed: {total_cases}")

# Continue with existing data quality checks...
```

---

## UPDATED REPORT SECTIONS

### Executive Summary Update

**OLD:**
```
As of 11 February 2026, the Support Department holds 180 total open cases 
with 171 cases (95.0%) due within the 30-day planning fence...
```

**NEW:**
```
As of 11 February 2026, the Support Department holds [X] active open cases 
(filtered for active work status) with [Y] cases ([Z]%) due within the 
30-day planning fence. [N] cases excluded from analysis due to pending/waiting 
status (not actively being worked). Data quality is excellent...
```

### Header Tags Update

**Add to header tags:**
```html
<span class="tag">ACTIVE: [X] CASES</span>
<span class="tag">EXCLUDED: [N] PENDING/WAITING</span>
```

### Data Quality Section Update

**Add Status Filter Summary:**
```
╔═══════════════════════════════════════════════════════════════╗
║ STATUS REASON FILTERING                                      ║
╚═══════════════════════════════════════════════════════════════╝

Total Cases in System:         [total before filter]
Active Cases (included):       [total after filter]
Excluded (Pending/Waiting):    [excluded count]

Active Status Reasons:
  • Estimation:                [count]
  • In Process with Microsoft: [count]
  • In Progress:               [count]
  • Not Started:               [count]
  • Requirement Gathering:     [count]
  • Req Gathering/Pend Brief:  [count]
  • Researching:               [count]
  • Solution Design:           [count]

Excluded Status Reasons:
  • Pending Client Confirm:    [count]
  • Pending Dev:               [count]
  • Pending Customer:          [count]
  • Other Pending/Waiting:     [count]

Rationale: Capacity calculations reflect only cases actively being 
worked or ready to start. Cases in pending/waiting states are excluded 
as they do not consume agent capacity until reactivated.
```

### Footer Update

**Add to footer:**
```
STATUS FILTERING: Only active work states included (8 approved statuses)
EXCLUDED: Pending/Waiting states that do not consume active capacity
```

---

## GOVERNANCE IMPACT

### Compliance Statement Update

```
Status Reason Filtering (NEW - v2.3.0):
  ✓ Filter Applied: 8 active status reasons only
  ✓ Excluded: Pending/Waiting states (non-active work)
  ✓ Rationale: Capacity reflects actively worked cases only
  ✓ Exclusion Count: [N] cases ([X]% of total)
```

### Report Limitations Update

**Add to report limitations:**
```
- Status Filter: Only cases with active work status included
- Excluded cases: [N] cases in Pending/Waiting states
- Assumption: Pending cases do not consume capacity until reactivated
- Impact: Capacity % reflects active workload only
```

---

## VERIFICATION CHECKLIST UPDATE

**Add these verification steps:**

- [ ] Status Reason filter applied BEFORE data quality audit
- [ ] Only 8 approved status reasons included
- [ ] Excluded cases count reported
- [ ] Status filter summary in Section 0 or header
- [ ] Capacity calculations based on filtered dataset only
- [ ] Footer states status filtering applied
- [ ] Governance statement includes status filter compliance

---

## EXAMPLE IMPACT ANALYSIS

### Before Status Filter:
- Total Open Cases: 180
- Department Capacity: 232.9%
- Rebecca Estephan: 408.7%

### After Status Filter (Example):
- Total Active Cases: 120 (60 excluded as Pending/Waiting)
- Department Capacity: 155.3% (33% reduction)
- Rebecca Estephan: 272.5% (33% reduction)

**Key Insight:** Filtering for active status dramatically improves capacity 
accuracy by excluding cases waiting on external parties.

---

## PROMPT INTEGRATION POINTS

### In "📊 INPUT SOURCES" Section:
Add after confirmed available fields:
```
CRITICAL STATUS FILTERING (v2.3.0):
- Column J: Status Reason
- ONLY include these 8 active status reasons:
  * Estimation
  * In Process with Microsoft
  * In Progress
  * Not Started
  * Requirement Gathering
  * Requirement Gathering/Pending Briefing
  * Researching
  * Solution Design
- EXCLUDE all Pending/Waiting statuses
```

### In "📐 MANDATORY KPIs" Section:
Update capacity calculation step 2:
```
# STEP 2: Filter Cases Within Fence (WITH REQUIRED FIELDS AND ACTIVE STATUS)
capacity_cases = open_df[
    (open_df['Due Date'].notna()) &
    (open_df['Initial Estimation Duration'].notna()) &
    (open_df['Initial Estimation Duration'] > 0) &
    (pd.to_datetime(open_df['Due Date']) <= planning_fence_date) &
    (open_df['Status Reason'].isin(ACTIVE_STATUS_REASONS))  # NEW
].copy()
```

### In "✅ DO's" Section:
Add:
```
14. ALWAYS filter for 8 active Status Reasons BEFORE any calculations
15. ALWAYS report excluded case count and exclusion rate
16. ALWAYS explain status filtering rationale in report
```

### In "❌ DON'Ts" Section:
Add:
```
13. DON'T include Pending/Waiting status cases in capacity
14. DON'T skip status reason filtering step
15. DON'T calculate capacity before applying status filter
```

---

## RATIONALE FOR CHANGE

**Why This Matters:**
1. **Accuracy:** Cases in "Pending Client Confirmation" or "Pending Dev" 
   don't consume agent capacity - they're waiting on external parties
2. **Actionability:** Executives need to know active workload, not total 
   case count including blocked items
3. **Resource Planning:** Capacity should reflect work agents can actively 
   progress, not cases on hold
4. **Consistency:** Aligns capacity metrics with actual work-in-progress

**Business Impact:**
- More accurate capacity planning
- Better resource allocation decisions
- Clearer visibility into actual vs. blocked workload
- Enables targeted unblocking of pending cases

---

## VERSION INFORMATION

**Version:** PRMT01-E v2.3.0
**Effective Date:** 12 February 2026
**Previous Version:** PRMT01-E v2.2.1
**Change Type:** Major - Calculation Methodology
**Backward Compatible:** No (results will differ from v2.2.1)

**Key Changes from v2.2.1:**
1. Added mandatory Status Reason filtering (8 active statuses only)
2. Updated capacity calculation to use filtered dataset
3. Added status filter summary reporting
4. Updated governance compliance statement
5. Enhanced exclusion tracking and reporting

---

## IMPLEMENTATION EXAMPLE

```python
# Complete filtering sequence (in order):

# 1. Filter for department agents
open_df = open_df[open_df['Owner'].isin(DEPARTMENT_AGENTS)].copy()
open_df['Owner'] = open_df['Owner'].replace('Raji Aoun', 'Unassigned Cases')

# 2. Filter for active status reasons (NEW v2.3.0)
ACTIVE_STATUS_REASONS = [
    'Estimation', 'In Process with Microsoft', 'In Progress',
    'Not Started', 'Requirement Gathering', 
    'Requirement Gathering/Pending Briefing',
    'Researching', 'Solution Design'
]
open_df = open_df[open_df['Status Reason'].isin(ACTIVE_STATUS_REASONS)].copy()

# 3. Proceed with data quality audit
# 4. Calculate 30-day capacity on filtered dataset
# 5. Generate report with status filter disclosure
```

---

**END OF UPDATE SPECIFICATION - PRMT01-E v2.3.0**

