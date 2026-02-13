# PRMT01-E v2.3.0 — Support Department Performance Report Specification

## COMPLETE REFERENCE FOR REPORT GENERATION

This document is the **single source of truth** for generating Support Department Performance Reports.

> **CRITICAL:** The authoritative design template is **`best_report_format/Reference_Design.html`**.
> All components, sections, CSS classes, chart types, color schemes, and layout MUST match that file exactly.
> When in doubt, open and inspect `Reference_Design.html` — it overrides anything described here.

---

## 1. OVERVIEW & FILE NAMING

- **Report Title:** `D365 F&O Support Report — {DD} {Mon} {YYYY}`
- **File Name:** `SupportDepartmentPerformance{YYYYMMDD}-{N}.html` where N is the version number for that day
- **Design Template:** `best_report_format/Reference_Design.html` — copy CSS, structure, and class names from this file
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
| Case Type | Classification + Case Type Breakdown chart |
| Priority | Normal / High / Low — determines SLA threshold |
| Customer | Customer impact analysis |
| Age | Case aging distribution |
| Actual Duration | Efficiency calculation (resolved cases) |
| Created On / Deactivated On | Monthly throughput calculation (resolved cases) |

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
- Check mandatory fields: Due Date, Initial Estimation Duration, Case Type
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

Uses RESOLVED cases data. Show "0.0%" if agent has no resolved cases or zero actual duration.
```

### SLA Breach
```
SLA Thresholds:
  Normal Priority = 16 days
  High Priority   = 2 days
  Low Priority    = 30 days

A case is in SLA breach if its Age exceeds the threshold for its priority.
Breach_% = (breached_cases / total_cases) × 100
SLA_Compliance_% = 100 - overall_breach_%
```

### Agent Status Badge
```
If Capacity > 300% → EXTREME  (badge .b-ex, row .rx)
If Capacity > 100% → CRITICAL (badge .b-cr, row .rc)
If Capacity ≤ 100% → OK       (badge .b-lo)
Unassigned Cases   → N/A      (badge .b-md)
```

### Customer Risk Level
```
CRITICAL: ≥20% of total cases OR avg age >50d with high case count  → badge .b-ex, row .rx
HIGH:     ≥10% of total cases OR avg age >40d                       → badge .b-hi, row .rc
MEDIUM:   All others with ≥3 cases                                  → badge .b-md
```

---

## 6. HTML REPORT STRUCTURE

> **Source of truth:** `best_report_format/Reference_Design.html`
> Copy the exact HTML structure, CSS classes, and section order from that file.
> Replace only the data values (numbers, names, percentages, chart data arrays).

### Section Order (MUST match Reference_Design.html exactly)

| # | Section | CSS Class | Description |
|---|---------|-----------|-------------|
| — | Header | `.hdr` | Title, subtitle, tag pills |
| — | Data Quality Alert | `.al.al-s` | Green success alert |
| 1 | Executive Summary | `.sec` `.sn`=1 | 6 metric cards |
| — | Dept Capacity Gauges | `.dept-hero` | 3 half-doughnut gauges + breakdown |
| 2 | Agent Performance Dashboard | `.sec` `.sn`=2 | Info alert + 9-column table + critical alert |
| 3 | 30-Day Capacity Analysis | `.sec` `.sn`=3 | Full-width stacked bar chart + warning alert |
| 4 | Operational Analytics | `.sec` `.sn`=4 | 2×2 chart grid |
| 5 | Workload Composition | `.sec` `.sn`=5 | 2+1 chart grid |
| 6 | Customer Impact — Top 10 | `.sec` `.sn`=6 | 7-column table + warning alert |
| — | Footer | `.ft` | Attribution + formula box |

### 6.1 Header (`.hdr`)
- H1: `Dynamics 365 Finance & Operations Support Report`
- Sub: `Support Department — Executive Performance Analysis`
- Tags (`.tag` pills): SNAPSHOT date, OPEN case count, AGENTS count, DATA QUALITY status, COMPLIANCE version
- See Reference_Design.html lines 116–126

### 6.2 Data Quality Alert (`.al.al-s`)
- Green success alert with field coverage percentages
- Includes planning fence computable case count
- See Reference_Design.html lines 129–131

### 6.3 Section 1: Executive Summary — 6 Cards
| # | Card Class | Icon | Value | Label |
|---|-----------|------|-------|-------|
| 1 | `.cd.dk` | 📋 | Total case count | Total Open Cases |
| 2 | `.cd.red` | ⚠️ | Avg capacity % | Avg Agent Capacity (30d) |
| 3 | `.cd.ros` | 🚨 | Breach rate % | SLA Breach Rate |
| 4 | `.cd.org` | 🎯 | Peak % (agent name) | Peak Capacity (Agent) |
| 5 | `.cd.pur` | 📅 | Avg age days | Avg Case Age |
| 6 | `.cd.blu` | 🏢 | Customer name (%) | Top Risk Customer |
- See Reference_Design.html lines 134–171

### 6.4 Department Capacity Gauge Hero (`.dept-hero`)
- 3 half-doughnut gauge charts using the `makeGauge()` helper function:
  1. **Overall Capacity** (`.gc` #deptGauge) — large value (`.gv.big`)
  2. **Normal Priority** (`.gc` #normGauge) — medium value (`.gv.med`)
  3. **High Priority** (`.gc` #highGauge) — medium value (`.gv.med`)
- Breakdown text: agents >300%, agents >200%, highest agent name + capacity
- Status tag: `.status-tag.overloaded` with overall status text
- See Reference_Design.html lines 174–205

### 6.5 Section 2: Agent Performance Dashboard
- Blue info alert (`.al.al-i`) with capacity formula and planning fence
- **9-column table:**

| Column | Content | Notes |
|--------|---------|-------|
| Agent | Name (bold) | |
| Total Cases | Count | |
| Status | Badge | `.b-ex` EXTREME / `.b-cr` CRITICAL / `.b-lo` OK / `.b-md` N/A |
| Capacity Visual | Mini bar (`.mb`) | Gradient fill with percentage |
| Overall % | Value with color class | `.ve` extreme / `.vc` critical / `.vb` blue / `.vm` medium |
| Efficiency % | Value with color class | `.vg` good / `.vm` medium / `.vp` poor |
| Avg Age | Days with 1 decimal | |
| SLA Breach | Format: "X / Y" | |
| Breach % | With `.vp` color class | |

- **Mini bar visual** in Capacity Visual column: `<div class="mb"><div class="mb-bg"><div class="mb-f" style="width:{pct}%;background:{gradient}"></div></div><div class="mb-v {colorclass}">{value}%</div></div>`
- Row classes: `.rx` for EXTREME, `.rc` for CRITICAL, none for OK
- Red critical alert (`.al.al-c`) below table with capacity crisis narrative
- See Reference_Design.html lines 207–355

### 6.6 Section 3: 30-Day Time-Fenced Capacity Analysis
- Full-width tall chart box (`.cb.full.tall`)
- **Stacked horizontal bar chart** (#capChart) with 3 datasets:
  - Normal Cap % (orange `rgba(251,146,60,.7)`)
  - High Cap % (red `rgba(239,68,68,.7)`)
  - Low Cap % (yellow `rgba(251,191,36,.7)`)
- Agents sorted by total capacity descending (excluding Unassigned)
- Orange warning alert (`.al.al-w`) with capacity breakdown narrative
- See Reference_Design.html lines 358–372

### 6.7 Section 4: Operational Analytics — 2×2 Grid (`.cg`)
| # | Chart Title | Type | Canvas ID | Details |
|---|------------|------|-----------|---------|
| 1 | Case Aging Distribution | Vertical bar | #ageChart | 7 buckets: 0-5d, 6-10d, 11-16d, 17-30d, 31-60d, 61-90d, 91+d |
| 2 | Customer Concentration Risk | Doughnut | #custPie | Top 5 + Others, cutout 55%, legend right |
| 3 | Efficiency % by Agent | Horizontal bar | #effChart | Sorted by efficiency desc, color-coded |
| 4 | SLA Compliance Rate | Gauge | #slaGauge | Half-doughnut with compliance %, text below |
- See Reference_Design.html lines 374–413

### 6.8 Section 5: Workload Composition — 2+1 Grid (`.cg`)
| # | Chart Title | Type | Canvas ID | Details |
|---|------------|------|-----------|---------|
| 1 | Case Status Distribution | Horizontal bar | #statChart | All status reasons sorted by count desc |
| 2 | Case Type Breakdown | Doughnut | #typeChart | All case types, cutout 50%, legend right |
| 3 | Monthly Throughput by Agent | Grouped bar (full-width) | #grpTP | Last 6 months, per-agent resolved cases |
- See Reference_Design.html lines 416–442

### 6.9 Section 6: Customer Impact Analysis — Top 10
- **7-column table:**

| Column | Content |
|--------|---------|
| # | Rank number |
| Customer | Name (bold for top rows) |
| Open Cases | Count |
| % of Total | Percentage |
| High Priority | Count of high priority cases |
| Avg Age | Days with 1 decimal |
| Risk Level | CRITICAL / HIGH / MEDIUM badge |

- Row classes: `.rx` for CRITICAL, `.rc` for HIGH, none for MEDIUM
- Orange warning alert (`.al.al-w`) with customer concentration risk narrative
- See Reference_Design.html lines 444–554

### 6.10 Footer (`.ft`)
- Attribution line: report title, snapshot date, "Generated by Senior Operations Manager & D365 F&O Solution Architect"
- Data Sources line: open case count, resolved case count, agent count
- Data Quality line: pass/fail status, capacity computable count
- Formula box (`.fm`): capacity formula, baselines, efficiency formula, SLA thresholds, planning fence, data quality status + compliance version
- See Reference_Design.html lines 557–570

---

## 7. DESIGN SYSTEM & CSS

> **Source of truth:** `best_report_format/Reference_Design.html` `<style>` block (lines 8–110).
> Copy the entire `<style>` block from Reference_Design.html into each generated report.
> The class names below are abbreviated — this is intentional; match the Reference_Design.html exactly.

### Key CSS Classes (abbreviated names from Reference_Design.html)

| Reference Class | Purpose |
|----------------|---------|
| `.w` | Wrapper container (max-width: 1360px) |
| `.hdr` | Header block (purple gradient + SVG pattern) |
| `.cd` | Metric card base |
| `.cd.dk` | Dark card (indigo gradient + purple border) |
| `.cd.red` | Red card (red → orange gradient) |
| `.cd.org` | Orange card (orange → yellow gradient) |
| `.cd.pur` | Purple card |
| `.cd.blu` | Blue card (sky → cyan gradient) |
| `.cd.ros` | Rose card (crimson → pink gradient) |
| `.cd.grn` | Green card (emerald gradient) |
| `.v` | Card value (2.4rem, 900 weight) |
| `.l` | Card label (uppercase, 0.72rem) |
| `.ic` | Card icon (emoji) |
| `.sec` | Section container |
| `.sh` | Section header (flex row with number + title) |
| `.sn` | Section number circle (purple gradient) |
| `.st` | Section title text (light purple) |
| `.dept-hero` | Department gauge hero section |
| `.gc` | Gauge chart container (250×145px) |
| `.gv` | Gauge value overlay |
| `.cg` | Chart grid (2-column) |
| `.cb` | Chart box (dark gradient, purple border) |
| `.cb.full` | Full-width chart box |
| `.cb.tall` | Tall chart box (max-height: 440px) |
| `.tw` | Table wrapper (rounded, shadowed) |
| `.b` | Badge base (inline-block, pill shape) |
| `.b-ex` | Badge EXTREME (solid red bg, white text) |
| `.b-cr` | Badge CRITICAL (red bg 15%, red text) |
| `.b-hi` | Badge HIGH (orange bg 15%, orange text) |
| `.b-md` | Badge MEDIUM/N/A (yellow bg 15%, yellow text) |
| `.b-lo` | Badge LOW/OK (green bg 15%, green text) |
| `.ve` | Value extreme (red #f87171, 800 weight) |
| `.vc` | Value critical (orange #fb923c, 700 weight) |
| `.vg` | Value good (green #34d399, 700 weight) |
| `.vm` | Value medium (yellow #fbbf24, 600 weight) |
| `.vp` | Value poor (red #f87171, 600 weight) |
| `.vb` | Value blue (#60a5fa, 600 weight) |
| `.mb` | Mini bar container |
| `.mb-bg` | Mini bar background track |
| `.mb-f` | Mini bar fill (colored, variable width) |
| `.mb-v` | Mini bar value label |
| `.al` | Alert base |
| `.al-s` | Alert success (green border #10b981) |
| `.al-i` | Alert info (blue border #3b82f6) |
| `.al-w` | Alert warning (orange border #ea580c) |
| `.al-c` | Alert critical (red border #dc2626) |
| `.rx` | Table row extreme/critical (red bg 10%) |
| `.rc` | Table row caution (orange bg 6%) |
| `.ft` | Footer container |
| `.fm` | Footer formula box (monospace, green text) |

### Core Colors
| Element | Color |
|---------|-------|
| Body background | `#0a0a14` |
| Body text | `#e2e8f0` |
| Wrapper max-width | `1360px` |
| Header gradient | `#4c1d95` → `#6d28d9` → `#7c3aed` → `#a78bfa` |
| Section title text | `#c4b5fd` |
| Chart box gradient | `#1a1740` → `#0e0e1a` |
| Table background | `#141428` |
| Footer gradient | `#1e1b4b` → `#0a0a14` |
| Chart label color | `#a78bfa` |
| Default chart text | `#94a3b8` |

### Typography
- Primary font: `'Aptos','Segoe UI',system-ui,-apple-system,sans-serif`
- Monospace (tags, formula): `Consolas,'Courier New',monospace`
- Header tags: `Consolas,monospace` at 0.78rem

---

## 8. CHART CONFIGURATIONS

> **Source of truth:** `best_report_format/Reference_Design.html` `<script>` block (lines 574–773).
> Copy chart initialization code structure from Reference_Design.html. Replace only data arrays.

### Global Chart Defaults
```javascript
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = "'Aptos','Segoe UI',system-ui,sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.pointStyle = 'circle';
```

### Gauge Helper Function
```javascript
function makeGauge(id, val, max, ok, warn) {
  const pct = Math.min(val/max, 1);
  const col = val > warn ? 'rgba(239,68,68,.75)' : val > ok ? 'rgba(251,146,60,.75)' : 'rgba(52,211,153,.75)';
  new Chart(document.getElementById(id), {
    type: 'doughnut',
    data: { datasets: [{ data: [pct*100, (1-pct)*100], backgroundColor: [col, 'rgba(255,255,255,.06)'], borderWidth: 0, circumference: 180, rotation: 270 }] },
    options: { responsive: false, cutout: '75%', plugins: { legend: { display: false }, tooltip: { enabled: false } } }
  });
}
```

### Chart Inventory

| # | Chart ID | Type | Location | Key Config |
|---|----------|------|----------|------------|
| 1 | #deptGauge | Half-doughnut (gauge) | Dept Hero | `makeGauge('deptGauge', avgCapacity, 400, 100, 200)` |
| 2 | #normGauge | Half-doughnut (gauge) | Dept Hero | `makeGauge('normGauge', normalCap, 300, 100, 150)` |
| 3 | #highGauge | Half-doughnut (gauge) | Dept Hero | `makeGauge('highGauge', highCap, 200, 80, 100)` |
| 4 | #slaGauge | Half-doughnut (gauge) | Section 4 | `makeGauge('slaGauge', compliance%, 100, 60, 40)` |
| 5 | #capChart | Stacked horizontal bar | Section 3 | 3 datasets: Normal/High/Low capacity %, `indexAxis:'y'`, stacked |
| 6 | #ageChart | Vertical bar | Section 4 | 7 buckets, color gradient green→red |
| 7 | #custPie | Doughnut | Section 4 | Top 5 + Others, cutout 55%, legend right |
| 8 | #effChart | Horizontal bar | Section 4 | Per-agent efficiency, sorted desc, color-coded by threshold |
| 9 | #statChart | Horizontal bar | Section 5 | All status reasons, sorted by count desc |
| 10 | #typeChart | Doughnut | Section 5 | All case types, cutout 50%, legend right |
| 11 | #grpTP | Grouped vertical bar | Section 5 | Last 6 months, one dataset per agent, legend bottom |

### Chart 5: Capacity Stacked Bar (#capChart)
```javascript
type: 'bar', indexAxis: 'y'
datasets: [
  { label: 'Normal Cap %', backgroundColor: 'rgba(251,146,60,.7)', borderRadius: 4 },
  { label: 'High Cap %',   backgroundColor: 'rgba(239,68,68,.7)',  borderRadius: 4 },
  { label: 'Low Cap %',    backgroundColor: 'rgba(251,191,36,.7)', borderRadius: 4 }
]
scales: { x: { stacked: true }, y: { stacked: true } }
legend: { position: 'top' }
```

### Chart 6: Case Aging (#ageChart)
```javascript
type: 'bar'
labels: ['0-5d','6-10d','11-16d','17-30d','31-60d','61-90d','91+d']
// 7 buckets with color gradient: green → yellow → orange → red → dark red
backgroundColor: [
  'rgba(52,211,153,.7)', 'rgba(52,211,153,.7)',
  'rgba(251,191,36,.7)', 'rgba(251,146,60,.7)',
  'rgba(239,68,68,.6)', 'rgba(220,38,38,.7)', 'rgba(153,27,27,.8)'
]
borderRadius: 6
```

### Chart 7: Customer Doughnut (#custPie)
```javascript
type: 'doughnut'
// Top 5 customers + "Others" bucket
cutout: '55%', legend: { position: 'right' }
backgroundColor: ['#dc2626','#ea580c','#f59e0b','#8b5cf6','#3b82f6','#475569']
borderColor: '#0a0a14'
// Tooltip: show percentage of total
```

### Chart 8: Efficiency Bar (#effChart)
```javascript
type: 'bar', indexAxis: 'y'
// Sorted by efficiency descending
// Green for ≥100%, Yellow for 50-100%, Red for <50%, Gray for 0%
borderRadius: 6
```

### Chart 9: Status Distribution (#statChart)
```javascript
type: 'bar', indexAxis: 'y'
// All status reasons, sorted by count descending
// Mixed colors for visual distinction
borderRadius: 4
```

### Chart 10: Case Type Doughnut (#typeChart)
```javascript
type: 'doughnut'
cutout: '50%', legend: { position: 'right' }
backgroundColor: ['#7c3aed','#ec4899','#f59e0b','#10b981','#3b82f6','#6366f1','#f43f5e','#14b8a6']
borderColor: '#0a0a14'
```

### Chart 11: Monthly Throughput (#grpTP)
```javascript
type: 'bar' // vertical grouped
// One dataset per agent, last 6 months
// Calculate from resolved cases: group by month(Deactivated On) and Owner
backgroundColor per agent:
  Georges Mouaikel:    'rgba(139,92,246,.75)'
  Rebecca Estephan:    'rgba(239,68,68,.7)'
  Fadi Hanna:          'rgba(251,146,60,.7)'
  Jana Sweid:          'rgba(56,189,248,.7)'
  Mennatullah El Bahr: 'rgba(52,211,153,.7)'
  Abdo Khoury:         'rgba(251,191,36,.7)'
legend: { position: 'bottom' }
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
- Pending Dev - Require Delivery Date
- Pending Customer
- Pending Delivery Date
- Waiting on Customer
- Waiting on Third Party
- On Hold
- Pending CR
- Pending Functional
- Pending Internal Approval
- Any status not in the approved list above

### Rationale
Capacity calculations must reflect only cases actively being worked. Cases in pending/waiting states are blocked on external parties and do not consume agent capacity until reactivated.

### Reporting Requirements
- Show total before filter, total after filter, and exclusion count
- Display exclusion rate as percentage
- List active status reasons with per-status case counts
- Report exclusion in header tags and data quality alert

---

## 10. COMMON MISTAKES TO AVOID

1. **DO NOT deviate from Reference_Design.html structure** — always copy CSS classes, section order, and component layout from `best_report_format/Reference_Design.html`
2. **DO NOT use long CSS class names** — the reference uses abbreviated names (`.w`, `.hdr`, `.cd`, `.sec`, `.sh`, `.sn`, `.st`, `.cg`, `.cb`, `.tw`, `.b`, `.al`, `.ft`, `.fm`, etc.)
3. **DO NOT omit the Department Capacity Gauge Hero section** — it must appear between Executive Summary and Agent Dashboard
4. **DO NOT use a simple bar for capacity** — Section 3 uses a **stacked** horizontal bar (Normal + High + Low)
5. **DO NOT omit the mini bar visuals** in the Agent Performance table — each row needs a `.mb` capacity bar
6. **DO NOT use only 2 badge levels** — use EXTREME (`.b-ex`), CRITICAL (`.b-cr`), OK (`.b-lo`), and N/A (`.b-md`)
7. **DO NOT use 6 aging buckets** — use 7 buckets: 0-5d, 6-10d, 11-16d, 17-30d, 31-60d, 61-90d, 91+d
8. **DO NOT omit the Efficiency chart** — Section 4 includes a horizontal bar chart for efficiency by agent
9. **DO NOT omit the SLA Compliance gauge** — Section 4 includes a half-doughnut gauge for SLA compliance
10. **DO NOT omit Section 5: Workload Composition** — includes Status Distribution, Case Type Breakdown, and Monthly Throughput
11. **DO NOT calculate capacity BEFORE applying status filter** — filter first, then calculate
12. **DO NOT show "Raji Aoun"** — always rename to "Unassigned Cases"
13. **DO NOT forget planning fence** = report date + 30 days
14. **DO NOT omit the formula box** (`.fm`) in the footer
15. **DO NOT include pending/waiting cases in ANY KPI calculation** — all metrics use filtered dataset only
16. **DO NOT use `#0f0f1a` as body background** — the reference uses `#0a0a14`
17. **DO NOT use `1320px` max-width** — the reference uses `1360px`
18. **DO NOT skip the "Generated by Senior Operations Manager & D365 F&O Solution Architect"** attribution in footer

---

## 11. VERIFICATION CHECKLIST

Before finalizing any report, verify against `best_report_format/Reference_Design.html`:

- [ ] File named correctly: `SupportDepartmentPerformance{YYYYMMDD}-{N}.html`
- [ ] CSS classes match Reference_Design.html (abbreviated names: `.w`, `.hdr`, `.cd`, `.sec`, etc.)
- [ ] Status filter applied BEFORE all calculations
- [ ] Only 8 approved status reasons included
- [ ] Header (`.hdr`) has 5 tag pills matching reference format
- [ ] Data quality alert (`.al.al-s`) present
- [ ] All 6 executive summary cards present with correct card classes
- [ ] Department Capacity Gauge Hero (`.dept-hero`) with 3 gauges
- [ ] Agent table has 9 columns including mini bar visual column
- [ ] Agent table uses 4-level badge system (EXTREME/CRITICAL/OK/N/A)
- [ ] Section 3: Stacked horizontal bar capacity chart
- [ ] Section 4: 4 charts (aging, customer doughnut, efficiency bar, SLA gauge)
- [ ] Section 5: 3 charts (status bar, case type doughnut, monthly throughput)
- [ ] Customer table has 7 columns with risk badges and row highlighting
- [ ] Footer (`.ft`) includes formula box (`.fm`) with all formulas
- [ ] Planning fence = report date + 30 days
- [ ] Chart.js CDN loaded (v4.4.1)
- [ ] Body background is `#0a0a14`, max-width is `1360px`

---

## 12. VERSION HISTORY

| Version | Date | Change |
|---------|------|--------|
| v2.3.0 | 12 Feb 2026 | Added mandatory status reason filtering (8 active statuses only). Major change — not backward compatible. |
| v2.2.1 | Prior | Previous version without status filtering |

**Key Changes in v2.3.0:**
1. Added mandatory Status Reason filtering (8 active statuses only)
2. Updated capacity calculation to use filtered dataset
3. Added status filter summary reporting
4. Updated governance compliance statement
5. Enhanced exclusion tracking and reporting
6. Adopted Reference_Design.html template (dark theme, gauges, stacked charts, mini bars)

---

## 13. QUICK REFERENCE — RUNTIME DETAILS

These are practical details discovered during actual report generation that prevent common rework.

### Resolved Cases File — Column Differences
The resolved cases Excel has **different columns** than open cases:
```
Columns: (Do Not Modify) Case, (Do Not Modify) Row Checksum, (Do Not Modify) Modified On,
         Case Number, Owner, Status, Customer, Case Title, Created On, Modified On,
         Age, Initial Estimation Duration, Actual Duration, Actual Duration (days), Deactivated On
```
- **No** `Priority`, `Status Reason`, `Case Type`, `Due Date` columns
- **Has** `Deactivated On` (datetime) — use this for monthly throughput grouping
- **Has** `Case Number` (string) — not present in open cases
- Apply same agent filter + Raji rename. Do NOT apply status filter (these are already resolved).

### Efficiency Calculation — Scope
- Uses **ALL** resolved cases for each agent (no date range filter)
- `Efficiency = sum(Initial Estimation Duration) / sum(Actual Duration) × 100`
- Many resolved cases have null `Initial Estimation Duration` (≈42% nulls) — these are excluded from numerator naturally via `.sum()` treating NaN as 0
- If `Actual Duration` sum is 0 or agent has no resolved cases → show `0.0%`

### Monthly Throughput — Calculation
```python
# Group resolved cases by Deactivated On month + Owner
# Last 6 months relative to report date (including current partial month)
# Current month label gets asterisk: "Feb 2026*"
# Only include department agents (excluding Unassigned/Raji)
```

### Case Type Doughnut — "Other" Grouping
- Show top 7 case types individually in the doughnut chart
- Group all remaining types into a single **"Other (N)"** bucket
- This prevents the legend from being too crowded

### Mini Bar Width Calculation
```
Mini bar fill width % = min(agent_capacity / peak_capacity × 100, 100)
```
- Peak capacity = highest agent's overall capacity %
- Green gradient (`#34d399` → `#6ee7b7`) for OK agents (≤100%)
- Orange gradient (`#ea580c` → `#fbbf24`) for CRITICAL agents (>100%)
- Red gradient (`#dc2626` → `#f97316`) for EXTREME agents (>300%)
- Gray (`#64748b`) for Unassigned Cases

### Efficiency Chart Color Thresholds
```
≥100%  → Green   rgba(52,211,153,.75)
50-99% → Yellow  rgba(251,191,36,.7)
<50%   → Red     rgba(239,68,68,.7)
0%     → Gray    rgba(100,116,139,.5)
```

### Header Tag #2 Format (v2.3.0)
The second tag pill must show BOTH active and excluded counts:
```
ACTIVE: {N} CASES • EXCLUDED: {N} PENDING/WAITING
```

### Footer Excluded Status Breakdown
List each excluded status with its count in the formula box:
```
EXCLUDED: Pending Client Confirmation (N), Pending Dev (N), Pending Customer (N), ...
```

### Data Quality Coverage Fields
Report coverage for three mandatory fields:
- Due Date coverage %
- Initial Estimation Duration coverage %
- Case Type coverage %
