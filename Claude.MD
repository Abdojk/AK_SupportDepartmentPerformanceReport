# PRMT01-E v2.3.0 — Support Department Performance Report Specification

## COMPLETE REFERENCE FOR REPORT GENERATION

This document is the **single source of truth** for generating Support Department Performance Reports.
Follow it exactly to avoid repeating past mistakes. When in doubt, refer to `best_report_format/reference_design.html`.

---

## 1. OVERVIEW & FILE NAMING

- **Report Title:** `D365 F&O Support Report – {DD} {Mon} {YYYY} (v2.3.0)`
- **File Name:** `SupportDepartmentPerformance{YYYYMMDD}-{N}.html` where N is the version number for that day
- **Reference Design:** `best_report_format/reference_design.html` — the HTML template to match exactly
- **Technology:** Single-file HTML5 with embedded CSS and Chart.js 4.4.1 via CDN
- **CDN:** `https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js`

---

## 2. DATA SOURCES & INPUT

### Excel Files (in `data/` folder)
- **Open Cases:** `Support Unit Open Cases {M}-{DD}-{YYYY} {time}.xlsx`
- **Resolved Cases:** `Support Unit Resolved Cases {M}-{DD}-{YYYY} {time}.xlsx`

### Key Columns Used
| Column | Purpose |
|--------|---------|
| Owner | Agent assignment (filter + grouping) |
| Status Reason | Active vs pending filtering |
| Due Date | Planning fence + SLA calculation |
| Initial Estimation Duration | Capacity calculation (in hours) |
| Case Type | Classification |
| Priority | Normal / High / Low — determines SLA threshold |
| Customer | Customer impact analysis |
| Age | Case aging distribution |
| Actual Duration | Efficiency calculation (resolved cases) |

---

## 3. DATA PROCESSING PIPELINE (STRICT ORDER)

Execute these steps IN THIS EXACT ORDER. Do not skip or reorder.

### Step 1: Read Excel File
```python
open_df = pd.read_excel(file_path)
```

### Step 2: Filter for Department Agents
```python
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
```

### Step 3: Rename Raji Aoun
```python
open_df['Owner'] = open_df['Owner'].replace('Raji Aoun', 'Unassigned Cases')
```

### Step 4: Filter for Active Status Reasons (v2.3.0 MANDATORY)
```python
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

total_before_status_filter = len(open_df)
open_df = open_df[open_df['Status Reason'].isin(ACTIVE_STATUS_REASONS)].copy()
total_after_status_filter = len(open_df)
excluded_by_status = total_before_status_filter - total_after_status_filter
```

**EXCLUDE all other statuses** including: Pending Client Confirmation, Pending Dev, Pending Customer, Pending Delivery Date, Waiting on Customer, Waiting on Third Party, On Hold, and any others not in the approved list.

### Step 5: Data Quality Audit (on filtered dataset)
- Check mandatory fields: Due Date, Initial Estimation Duration
- Report coverage percentages
- Note any missing data

### Step 6: Calculate KPIs (on filtered dataset)
- See Section 5 for formulas

---

## 4. DEPARTMENT AGENTS

| Agent | Display Name | Notes |
|-------|-------------|-------|
| Rebecca Estephan | Rebecca Estephan | |
| Mennatullah El Bahr | Mennatullah El Bahr | |
| Abdo Khoury | Abdo Khoury | |
| Jana Sweid | Jana Sweid | |
| Fadi Hanna | Fadi Hanna | |
| Georges Mouaikel | Georges Mouaikel | |
| Raji Aoun | **Unassigned Cases** | ALWAYS rename — never show "Raji Aoun" |

**Sort order in tables and charts:** By capacity % DESCENDING (highest first).

---

## 5. KPI FORMULAS

### Capacity (30-Day Time-Fenced)
```
Planning Fence = Report Date + 30 days

For each agent, sum estimated hours by priority for cases due within the fence:
  Normal_Est_Hours = sum of Initial Estimation Duration for Normal priority cases
  High_Est_Hours   = sum of Initial Estimation Duration for High priority cases
  Low_Est_Hours    = sum of Initial Estimation Duration for Low priority cases

Overall_Capacity_% = (Normal_Est_Hours / 128 × 100)
                   + (High_Est_Hours / 16 × 100)
                   + (Low_Est_Hours / 240 × 100)

Baselines:
  Normal = 128 hours (16 working days × 8 hours)
  High   = 16 hours  (2 working days × 8 hours)
  Low    = 240 hours (30 working days × 8 hours)
```

### Efficiency
```
Efficiency_% = (Total_Estimated_Duration / Total_Actual_Duration) × 100

Uses RESOLVED cases data. Show "—" if agent has no resolved cases in period.
```

### SLA Breach
```
SLA Thresholds:
  Normal Priority = 16 days
  High Priority   = 2 days
  Low Priority    = 30 days

A case is in SLA breach if its Age exceeds the threshold for its priority.
Breach_% = (breached_cases / total_cases) × 100
```

### Agent Status Badge
```
If Capacity > 100% → OVERLOAD (badge-overload, row-overload)
If Capacity ≤ 100% → OK (badge-ok, row-ok)
```

### Customer Risk Level
```
CRITICAL: ≥20% of total cases OR avg age >50d with high case count
HIGH:     ≥10% of total cases OR avg age >40d
MEDIUM:   All others with ≥3 cases
```

---

## 6. HTML REPORT STRUCTURE (EXACT SECTION ORDER)

**DO NOT change this order. It must match the reference design.**

### 6.1 Header
```html
<div class="header">
  <h1>📊 Dynamics 365 Finance & Operations Support Report</h1>
  <div class="sub">Support Department — Executive Performance Analysis</div>
  <div class="meta">SNAPSHOT: {DD} {MON} {YYYY} • ACTIVE: {N} CASES • EXCLUDED: {N} PENDING/WAITING • VERSION: PRMT01-E v2.3.0</div>
  <div class="tags">
    <span class="tag">✅ DATA QUALITY: PASS</span>
    <span class="tag">🎯 STATUS FILTERED</span>
    <span class="tag">📆 30-DAY FENCE</span>
  </div>
</div>
```

### 6.2 Status Filter Disclosure Box
- Blue gradient background (`linear-gradient(135deg,#1e3a8a,#1e40af)`)
- Title: "🎯 STATUS REASON FILTERING (v2.3.0)"
- Explanation paragraph
- 4 stat cards: Total In System, Active (Included), Excluded, Exclusion Rate
- List of active status reasons with counts

### 6.3 Section 1: Executive Summary
- 6 metric cards in a responsive grid:
  1. **Active Cases** (card.dark) — total filtered count
  2. **Avg Agent Capacity** (card.green) — average of all agents
  3. **Agents >100% Capacity** (card.purple) — count fraction like "5/6"
  4. **Peak Capacity** (card.orange) — highest agent name + %
  5. **Avg Case Age** (card.blue) — average days
  6. **Top Customer** (card.dark) — name + % of workload

### 6.4 Section 2: Data Quality & Capacity Overview
- Green alert box: data quality summary (coverage %)
- Full-width capacity chart (`chart-box full`) — horizontal bar
- Blue info alert: capacity formula, planning fence date, status filtering note

### 6.5 Section 3: Agent Performance Dashboard
- 8-column table:

| Column | Content |
|--------|---------|
| Agent | Name (bold) |
| Status | OVERLOAD or OK badge |
| Cases | Active case count |
| Capacity % | With v-overload/v-ok class |
| Efficiency % | With v-good/v-poor class, or "—" |
| Avg Age | Days with 1 decimal |
| SLA Breach | Format: "X / Y" |
| Breach % | With color class |

- Orange warning alert below: capacity alert text

### 6.6 Section 4: Operational Analytics
- 2×2 chart grid (`chart-grid`):
  1. **Case Aging Distribution** — vertical bar, 6 buckets
  2. **Customer Concentration (Top 5 + Others)** — doughnut
  3. **Priority Distribution** — pie
  4. **Status Reason Distribution (Active Only)** — horizontal bar

### 6.7 Section 5: Customer Impact Analysis — Top 10
- 7-column table:

| Column | Content |
|--------|---------|
| # | Rank number |
| Customer | Name (bold for top 3) |
| Active Cases | Count |
| % of Total | Percentage |
| High Priority | Count of high priority cases |
| Avg Age | Days with 1 decimal |
| Risk Level | CRITICAL / HIGH / MEDIUM badge |

- Orange warning alert below: customer concentration risk text

### 6.8 Footer
```html
<div class="footer">
  <strong>Support Unit Executive Performance Report</strong> • Snapshot: {date} • Generated by Senior Operations Manager & D365 F&O Solution Architect<br>
  <strong>Data Sources:</strong> Active Cases ({N} filtered from {total} total) • Planning Fence: ≤{fence date} • Agents: 6 active + 1 Unassigned<br>
  <strong>Version:</strong> PRMT01-E v2.3.0 (Status Reason Filtering + 30-Day Time-Fenced Capacity + Data Quality Audit)<br>
  <div class="formula">
    CAPACITY FORMULA: Overall_Capacity_% = (Normal_Est_Hours / 128 × 100) + (High_Est_Hours / 16 × 100) + (Low_Est_Hours / 240 × 100)
    Normal Baseline = 128h (16 days × 8h) | High Baseline = 16h (2 days × 8h) | Low Baseline = 240h (30 days × 8h)
    STATUS FILTERING (v2.3.0): Only active work states included (8 approved status reasons)
    EXCLUDED: Pending/Waiting states ({N} cases, {X}%) — not actively consuming capacity
    EFFICIENCY FORMULA: (Total_Estimated_Duration / Total_Actual_Duration) × 100
    SLA THRESHOLDS: Normal Priority = 16 days | High Priority = 2 days | Low Priority = 30 days
  </div>
</div>
```

---

## 7. DESIGN SYSTEM & CSS

### Colors
| Element | Color |
|---------|-------|
| Body background | `#0f0f1a` |
| Body text | `#e2e8f0` |
| Primary accent | `#7c3aed` |
| Light accent | `#a78bfa` |
| Dark accent | `#4c1d95` |
| Section title text | `#c4b5fd` |
| Table background | `#1a1a2e` |
| Chart box gradient | `#1e1b4b` → `#0f0f1a` |

### Card Classes
| Class | Gradient |
|-------|----------|
| `.card.green` | `#10b981` → `#34d399` |
| `.card.orange` | `#ea580c` → `#facc15` |
| `.card.purple` | `#7c3aed` → `#c084fc` |
| `.card.dark` | `#1e1b4b` → `#312e81` (+ purple border) |
| `.card.blue` | `#0369a1` → `#22d3ee` |

### Badge Classes
| Class | Background | Text Color |
|-------|------------|------------|
| `.badge-overload` | `rgba(250,204,21,.15)` | `#fde047` |
| `.badge-ok` | `rgba(34,197,94,.15)` | `#86efac` |
| `.badge-critical` | `rgba(220,38,38,.15)` | `#fca5a5` |
| `.badge-high` | `rgba(234,88,12,.15)` | `#fdba74` |
| `.badge-medium` | `rgba(250,204,21,.15)` | `#fde047` |

### Value Color Classes
| Class | Color | Use |
|-------|-------|-----|
| `.v-overload` | `#fbbf24` (yellow) | Capacity >100% |
| `.v-ok` | `#34d399` (green) | Capacity ≤100% |
| `.v-good` | `#34d399` (green) | Efficiency ≥100% |
| `.v-med` | `#fbbf24` (yellow) | Breach 30-60% |
| `.v-poor` | `#f87171` (red) | Efficiency <50% or Breach >60% |

### Alert Classes
| Class | Border Color | Use |
|-------|-------------|-----|
| `.alert-success` | `#22c55e` | Data quality pass |
| `.alert-info` | `#3b82f6` | Capacity formula / info |
| `.alert-warning` | `#ea580c` | Capacity alerts, risk warnings |

### Typography
- Primary font: `'Aptos','Segoe UI',system-ui,-apple-system,sans-serif`
- Monospace (meta, formula): `'Consolas','Courier New',monospace`
- Max width: `1320px`

---

## 8. CHART CONFIGURATIONS

### Chart 1: Capacity by Agent (Horizontal Bar)
```javascript
type: 'bar'
indexAxis: 'y'
// Sorted by capacity descending
// Yellow (#fbbf24) for >100%, Green (#34d399) for ≤100%
// X-axis: ticks with '%' suffix, title "Capacity %"
// No legend
borderRadius: 6
```

### Chart 2: Case Aging Distribution (Vertical Bar)
```javascript
type: 'bar'
labels: ['0-10d','11-20d','21-30d','31-60d','61-90d','91+d']
// Color gradient: green → yellow → orange → red
// rgba(52,211,153,0.8), rgba(52,211,153,0.7), rgba(251,191,36,0.7),
// rgba(251,146,60,0.7), rgba(239,68,68,0.7), rgba(220,38,38,0.8)
borderRadius: 6
```

### Chart 3: Customer Concentration (Doughnut)
```javascript
type: 'doughnut'
// Top 5 customers + "Others" bucket
cutout: '55%'
legend: position 'right'
// Colors: #dc2626, #ea580c, #f59e0b, #8b5cf6, #3b82f6, #475569
// Label format: "CustomerName (count)"
// Tooltip shows percentage of total active cases
borderColor: '#0f0f1a'
```

### Chart 4: Priority Distribution (Pie)
```javascript
type: 'pie'
// Labels: "Normal (N)", "Low (N)", "High (N)"
// Colors: #8b5cf6 (Normal), #3b82f6 (Low), #dc2626 (High)
legend: position 'bottom'
borderColor: '#0f0f1a'
```

### Chart 5: Status Reason Distribution (Horizontal Bar)
```javascript
type: 'bar'
indexAxis: 'y'
// All active status reasons, sorted by count descending
// Abbreviated labels: "Req Gathering", "In Process MS", "Req Gather/Brief"
// Mixed colors for visual distinction
borderRadius: 4
```

### Global Chart Defaults
```javascript
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = "'Aptos','Segoe UI',system-ui,sans-serif";
Chart.defaults.font.size = 12;
```

---

## 9. STATUS FILTERING SPECIFICATION (v2.3.0)

### ONLY Include These 8 Status Reasons:
1. Estimation
2. In Process with Microsoft
3. In Progress
4. Not Started
5. Requirement Gathering
6. Requirement Gathering/Pending Briefing
7. Researching
8. Solution Design

### EXCLUDE All Others, Including:
- Pending Client Confirmation
- Pending Dev
- Pending Customer
- Pending Delivery Date
- Waiting on Customer
- Waiting on Third Party
- On Hold
- Any status not in the approved list above

### Rationale
Capacity calculations must reflect only cases actively being worked. Cases in pending/waiting states are blocked on external parties and do not consume agent capacity until reactivated.

### Reporting Requirements
- Show total before filter, total after filter, and exclusion count
- Display exclusion rate as percentage
- List active status reasons with per-status case counts
- Include status filter disclosure in header meta line AND dedicated filter box

---

## 10. COMMON MISTAKES TO AVOID

1. **DO NOT change the section order** — Header → Filter Box → Exec Summary → Data Quality → Agent Table → Analytics Charts → Customer Table → Footer
2. **DO NOT omit the status filter disclosure box** — it must appear between header and executive summary
3. **DO NOT use a different color scheme or font** — match the design system exactly
4. **DO NOT calculate capacity BEFORE applying status filter** — filter first, then calculate
5. **DO NOT forget exclusion counts** in header meta line (format: `ACTIVE: N CASES • EXCLUDED: N PENDING/WAITING`)
6. **DO NOT use "Top 10" for customer doughnut chart** — use "Top 5 + Others" (top 10 is for the TABLE only)
7. **DO NOT change chart types** — capacity is horizontal bar, aging is vertical bar, customer is doughnut, priority is pie, status is horizontal bar
8. **DO NOT show "Raji Aoun"** — always rename to "Unassigned Cases"
9. **DO NOT forget planning fence** = report date + 30 days
10. **DO NOT omit the formula box** in the footer
11. **DO NOT make the capacity chart vertical** — it must be horizontal (indexAxis: 'y')
12. **DO NOT forget to sort agents by capacity descending** in both table and chart
13. **DO NOT use different badge logic** — >100% = OVERLOAD, ≤100% = OK
14. **DO NOT skip the "Generated by Senior Operations Manager & D365 F&O Solution Architect"** attribution in footer
15. **DO NOT include pending/waiting cases in ANY KPI calculation** — all metrics use filtered dataset only

---

## 11. VERIFICATION CHECKLIST

Before finalizing any report, verify:

- [ ] File named correctly: `SupportDepartmentPerformance{YYYYMMDD}-{N}.html`
- [ ] Status filter applied BEFORE all calculations
- [ ] Only 8 approved status reasons included
- [ ] Exclusion count shown in header meta and filter box
- [ ] All 6 executive summary cards present with correct data
- [ ] Data quality alert shows coverage %
- [ ] Capacity chart is horizontal bar, sorted descending
- [ ] Agent table has 8 columns, sorted by capacity descending
- [ ] All 4 analytics charts render (aging, customer, priority, status)
- [ ] Customer doughnut shows Top 5 + Others
- [ ] Customer table shows Top 10 with risk badges
- [ ] Footer includes formula box with all formulas
- [ ] Footer includes status filtering disclosure
- [ ] Planning fence = report date + 30 days
- [ ] Chart.js CDN loaded (v4.4.1)
- [ ] Design matches reference: dark theme, purple accents, correct fonts

---

## 12. VERSION HISTORY

| Version | Date | Change |
|---------|------|--------|
| v2.3.0 | 12 Feb 2026 | Added mandatory status reason filtering (8 active statuses only). Major change — not backward compatible. |
| v2.2.1 | Prior | Previous version without status filtering |

**Key Changes in v2.3.0:**
1. Added mandatory Status Reason filtering (8 active statuses only)
2. Updated capacity calculation to use filtered dataset
3. Added status filter summary reporting (disclosure box)
4. Updated governance compliance statement
5. Enhanced exclusion tracking and reporting
6. Adopted reference design template (dark theme, Chart.js, responsive layout)
