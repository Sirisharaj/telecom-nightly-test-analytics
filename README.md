# 📡 Telecom Daily Test Pass/Fail Analytics & Regression Tracker

An end-to-end operational analytics built to automate nightly regression triage across **5G NR**, **LTE**, and **NB-IoT** test simulator exports, cutting daily morning analysis time from **45 minutes to under 5 minutes** while establishing a single source of truth for KPI reporting.

---

## 🎯 Business Problem & Objective
* **The Problem**: In a 3GPP cellular simulator lab, hundreds of automated regression tests run nightly across multiple radio access technologies. Each morning, QA engineers manually parsed raw CSV exports to find failures, while management received conflicting pass rates because `SKIPPED` tests and hardware-related `BLOCKED` test cases were inconsistently calculated across teams.
* **The Objective**: 
  1. Standardize cross-team KPI definitions: $\text{Pass Rate} = \frac{\text{PASS}}{\text{PASS} + \text{FAIL}}$ (excluding `SKIPPED` and `BLOCKED` tests from the denominator while maintaining visibility into bench downtime).
  2. Automate regression classification into **NEW** vs. **REPEAT** failures to accelerate root-cause triage.
  3. Deliver a robust, 1-page morning executive summary report in Excel that refreshes in under 5 minutes before the 09:00 AM daily standup.

---

## 🏗️ Repository Structure

```text
telecom-nightly-test-analytics/
├── README.md                                  <- Comprehensive project documentation & case study
├── data/
│   ├── raw/
│   │   ├── nightly_test_results_2026-10-30.csv <- Day 1 test run export (42 rows with duplicates/whitespace)
│   │   ├── nightly_test_results_2026-10-31.csv <- Day 2 test run export (Yesterday's baseline)
│   │   └── nightly_test_results_2026-11-01.csv <- Day 3 test run export (Today's active dataset)
├── reports/
│   └── daily_test_pass_fail_report.xlsx        <- Automated multi-tab workbook with dynamic models & charts
└── docs/
    └── SOP_Refresh_Guide.md                    <- 5-Minute standard operating procedure for daily triage
```

---

## 📊 Executive KPI Summary (Today vs. Yesterday)

| Radio Technology | PASS | FAIL | SKIPPED / BLOCKED | Total Executed | Pass Rate (Today) | Pass Rate (Yesterday) | Daily Shift ($\Delta$) | Primary Driver / Triage Insight |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **5G NR** | 10 | 4 | 2 | 14 | **71.4%** | 78.6% | **-7.1%** 🔻 | 3 new attach/PDU regressions isolated to `SIM-BENCH-01` |
| **LTE** | 11 | 2 | 1 | 13 | **84.6%** | 84.6% | **0.0%** ➖ | Steady performance; 1 new RRC re-establishment bug |
| **NB-IoT** | 7 | 1 | 2 | 8 | **87.5%** | 77.8% | **+9.7%** 🔺 | NPRACH CE1 recovered; repeat failure on CE2 |
| **Total / Overall** | **28** | **7** | **5** | **35** | **80.0%** | **81.1%** | **-1.1%** | 35 executed; 5 excluded from denominator |

---

## 🔍 Data Quality Audit & Cleaning Workflow

During initial exploratory analysis of the raw CSV feeds, 4 critical data hygiene issues were detected and resolved:
1. **Trailing & Leading Whitespace**: Test cases and verdicts had uneven whitespace padding (e.g., `" fail "`, `"  NR_PDU_session_est "`), breaking exact string lookups. Cleaned via `=TRIM()` / Power Query Text Trim.
2. **Inconsistent Case Formatting**: Mixed casing (`"Pass"`, `"PASS"`) resolved via `=UPPER()` normalization.
3. **Duplicate Simulator Lines**: Simulator re-transmissions produced duplicate `run_date` + `test_case` entries (42 raw rows $\rightarrow$ 40 unique tests). Cleaned using composite-key deduplication.
4. **Schema Variation**: Uncovered a 7th undocumented column (`bench_id`) identifying hardware testbeds (`SIM-BENCH-01` to `04`). Retained in raw sheets to assist infrastructure vs. software failure attribution.

---

## ⚙️ Analytical Methodology & Formula Implementation

### 1. Dynamic Technology Extraction
Instead of hardcoding row labels, technology categories populate dynamically:
```excel
=SORT(UNIQUE(Raw_Data_Today!$C$2:$C$41))
```

### 2. Standardized KPI Aggregation
```excel
=COUNTIFS(Raw_Data_Today!$C$2:$C$41, A7, Raw_Data_Today!$D$2:$D$41, "PASS")
```
$$\text{Pass Rate} = \frac{\text{PASS}}{\text{Executed Tests}} = \frac{B7}{E7}$$

### 3. Automated Regression Classification (`XLOOKUP`)
To immediately isolate new overnight software regressions from known issues, yesterday's verdict is looked up dynamically:
```excel
=XLOOKUP(B5, Raw_Data_Yesterday!$B$2:$B$41, Raw_Data_Yesterday!$D$2:$D$41, "NOT_FOUND")
```
The failure status flag is assigned via:
```excel
=IF(E5="FAIL", "REPEAT", "NEW")
```

#### Triage Result for Today's Failures (`2026-11-01`):
* **NEW Failures (4)**: `NR_attach_basic` (RRC timeout), `NR_PDU_session_est` (SMF N4 timeout), `NR_MIMO_4x4_downlink_tput` (Throughput SLA breach), `LTE_RRC_reestablishment` (T311 expiry).
* **REPEAT Failures (3)**: `NR_gNB_beam_failure_rec`, `LTE_VoLTE_SIP_invite`, `NBIoT_NPRACH_coverage_CE2`.

---

## ⏱️ 5-Minute Morning Refresh SOP

1. **Roll Yesterday's Baseline**: Copy today's cleaned data and paste as values into `Raw_Data_Yesterday`.
2. **Ingest Nightly CSV**: Paste new export into `Raw_Data_Today` (or click **Data > Refresh All** if connected to Power Query).
3. **Data Quality Gate**: Confirm 40 unique test cases after automated deduplication (`run_date` + `test_case`).
4. **Triage Review**: Filter `Failures_Today` for `Status = "NEW"` and post root-cause ticket links.
5. **Standup Handoff**: Review the 3-line summary and Technology Clustered Bar Chart before 08:50 AM.

---

## 💼 Transitioning from QA to Data Analyst (Interview Positioning)

### The Strategic Value of QA Background in Analytics
Having a QA engineering background is a distinct competitive advantage for data roles:
* **Deep Data Provenance Knowledge**: You understand the upstream systems and edge conditions that create data noise.
* **Rigorous Edge-Case Thinking**: Instinctive verification of boundary conditions, null values, and dirty schemas before building metrics.
* **Root-Cause Orientation**: Moving beyond passive reporting to actively isolating drivers of change (e.g., tying a 7.1% drop to a specific simulator hardware bench).

### Resume Bullet Points
* **Telecom Regression Analytics**: Designed and automated an Excel/Power Query daily test reporting pipeline across **5G NR, LTE, and NB-IoT**, reducing daily manual triage time by **88% (45 min $\rightarrow$ 5 min)**.
* **KPI Governance & Standardization**: Implemented strict data hygiene rules excluding `SKIPPED`/`BLOCKED` hardware states from pass-rate calculations, eliminating cross-team reporting discrepancies between QA and engineering managers.
* **Regression Triage Automation**: Engineered dynamic lookup models (`XLOOKUP`/`INDEX-MATCH`) to classify overnight failures into **NEW** vs. **REPEAT**, directly localizing a **7.1% drop in 5G NR pass rate** to attach timeouts on testbed `SIM-BENCH-01`.

---

## 🛠️ Tools & Technologies Used
* **Microsoft Excel**: Dynamic Array Formulas (`SORT`, `UNIQUE`, `FILTER`), `XLOOKUP`, `COUNTIFS`, PivotTables, Clustered Bar Visualizations.
* **Power Query (M) / ETL**: Text cleaning, casing normalization, multi-column deduplication.
* **Git & GitHub**: Version control, analytical documentation, markdown reporting.
