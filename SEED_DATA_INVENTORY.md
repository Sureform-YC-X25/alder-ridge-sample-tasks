# Alder Ridge Seed Data Inventory

This catalog links directly to every file included in the complete Alder Ridge sample seed. It covers the full 132-file shared company world, the accounting database, and the seed control/support files.
Click any filename below to open that file directly in GitHub.

## Summary

| Measure | Volume |
|---|---:|
| Company-world source files | 132 |
| Complete seed files, including controls/database | 139 |
| Complete seed size | 78.37 MiB |
| Source-artifact size | 3.62 MiB |
| Accounting database size | 74.35 MiB |
| Accounting database rows | 575,008 |
| Excel workbooks / worksheets | 54 / 260 |
| Detected Excel list records | 20,121 |
| PDF files / pages | 19 / 87 |
| Word files / rendered pages | 16 / 75 |
| PowerPoint files / slides | 9 / 62 |
| CSV files / records | 7 / 1,052 |

### File types

| Type | Files |
|---|---:|
| `xlsx` | 54 |
| `eml` | 26 |
| `pdf` | 19 |
| `docx` | 16 |
| `pptx` | 9 |
| `csv` | 7 |
| `txt` | 1 |

### Source authority/version status

| Status | Files |
|---|---:|
| current-support | 71 |
| current-working | 49 |
| historical-final | 6 |
| stale-working | 3 |
| approved-plan | 2 |
| superseded | 1 |

### Native-format volume

- PDF pages: total 87 / average 4.58 / median 4 / range 2-12
- Word rendered pages: total 75 / average 4.69 / median 3.0 / range 2-19
- PowerPoint slides: total 62 / average 6.89 / median 7 / range 5-9
- Excel worksheets: total 260 / average 4.81 / median 4.0 / range 1-9
- Excel populated rows: total 23,123 / average 428.20 / median 113.0 / range 12-5,598
- Excel detected list records per list-containing workbook: total 20,121 / average 428.11 / median 112 / range 13-5,544
- CSV records: total 1,052 / average 150.29 / median 170 / range 78-240
- Email body words: total 5,344 / average 205.54 / median 122.0 / range 68-914

Excel list records use a structural estimate: a mainly textual header spanning at least three columns followed by at least eight contiguous populated data rows. Model output rows, repeated versions, and reconciled copies can represent the same underlying business fact, so these are workload-volume measures rather than unique-event counts.

## Accounting database

[`accounting.db`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/accounting.db) contains **575,008 rows across 27 business tables**.

| Table | Rows |
|---|---:|
| `accounts` | 108 |
| `audit_events` | 0 |
| `bank_transactions` | 59,406 |
| `company` | 1 |
| `cost_codes` | 24 |
| `customer_invoice_lines` | 31,479 |
| `customer_invoices` | 31,479 |
| `customers` | 160 |
| `departments` | 6 |
| `employees` | 176 |
| `fixed_assets` | 81 |
| `job_cost_entries` | 20,627 |
| `journal_headers` | 98,518 |
| `journal_lines` | 212,805 |
| `payroll_run_lines` | 15,114 |
| `payroll_runs` | 215 |
| `prepaid_items` | 57 |
| `project_change_orders` | 8 |
| `projects` | 68 |
| `purchase_order_lines` | 2,057 |
| `purchase_orders` | 2,057 |
| `service_agreements` | 90 |
| `service_work_orders` | 30,576 |
| `timecard_entries` | 45,452 |
| `vendor_bill_lines` | 12,042 |
| `vendor_bills` | 12,042 |
| `vendors` | 360 |
| **Total** | **575,008** |

## Seed documentation and controls

These files document or verify the seed package. Control files are not mounted into the agent-visible workspace.

- [`ACCOUNTING_MCP.md`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/ACCOUNTING_MCP.md) — 6.17 KiB
- [`accounting.db`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/accounting.db) — 74.35 MiB
- [`sample_manifest.json`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sample_manifest.json) — 12.56 KiB
- [`controls/corporate_finance_source_map.json`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/controls/corporate_finance_source_map.json) — 15.41 KiB
- [`controls/corporate_finance_template_inputs.json`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/controls/corporate_finance_template_inputs.json) — 293.34 KiB
- [`controls/file_registry.json`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/controls/file_registry.json) — 56.67 KiB
- [`controls/task_source_dependencies.json`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/controls/task_source_dependencies.json) — 19.57 KiB

## Company-world source files

### `Deliverables`

- [`ARM-2417 June close and PCO-017 review.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Deliverables/ARM-2417%20June%20close%20and%20PCO-017%20review.docx) — DOCX · 39.71 KiB · current-support · 2 rendered pages · 6 tables · 33 table rows

### `Requests`

- [`7.1.26_822am - Fwd_ wip close need today.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.1.26_822am%20-%20Fwd_%20wip%20close%20need%20today.eml) — EML · 1.98 KiB · current-support · 313 body words
- [`7.1.26_904am - cash refresh + Fri deck.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.1.26_904am%20-%20cash%20refresh%20%2B%20Fri%20deck.eml) — EML · 895 B · current-support · 112 body words
- [`7.2.26_1027am - Q2 bank compliance package.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_1027am%20-%20Q2%20bank%20compliance%20package.eml) — EML · 946 B · current-support · 110 body words
- [`7.2.26_1103am - CFO close decision note - internal only.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_1103am%20-%20CFO%20close%20decision%20note%20-%20internal%20only.eml) — EML · 789 B · current-support · 87 body words
- [`7.2.26_611am - June ops review deck + notes.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_611am%20-%20June%20ops%20review%20deck%20%2B%20notes.eml) — EML · 697 B · current-support · 68 body words
- [`7.2.26_736am - FW_ wc cleanup + backlog coverage.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_736am%20-%20FW_%20wc%20cleanup%20%2B%20backlog%20coverage.eml) — EML · 963 B · current-support · 109 body words
- [`7.2.26_805am - payroll bridge + capex before noon.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_805am%20-%20payroll%20bridge%20%2B%20capex%20before%20noon.eml) — EML · 1.06 KiB · current-support · 128 body words
- [`7.2.26_842am - service vans capacity check - LT fwd.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.2.26_842am%20-%20service%20vans%20capacity%20check%20-%20LT%20fwd.eml) — EML · 952 B · current-support · 102 body words
- [`7.4.26_0711am - FY27 plan first pass + branch submissions.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_0711am%20-%20FY27%20plan%20first%20pass%20%2B%20branch%20submissions.eml) — EML · 1.53 KiB · current-support · 186 body words
- [`7.4.26_0838am - capital committee + five year plan.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_0838am%20-%20capital%20committee%20%2B%20five%20year%20plan.eml) — EML · 1.13 KiB · current-support · 115 body words
- [`7.4.26_0956am - July liquidity bank work + covenant outlook.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_0956am%20-%20July%20liquidity%20bank%20work%20%2B%20covenant%20outlook.eml) — EML · 1.29 KiB · current-support · 140 body words
- [`7.4.26_1106am - corp dev IC workstream status.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_1106am%20-%20corp%20dev%20IC%20workstream%20status.eml) — EML · 1.26 KiB · current-support · 139 body words
- [`7.4.26_1216pm - tax provision + EBITDA definitions.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_1216pm%20-%20tax%20provision%20%2B%20EBITDA%20definitions.eml) — EML · 1.16 KiB · current-support · 125 body words
- [`7.4.26_1344pm - Q2 board lender + IR refresh.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.4.26_1344pm%20-%20Q2%20board%20lender%20%2B%20IR%20refresh.eml) — EML · 1.15 KiB · current-support · 124 body words
- [`7.5.26_0618am - FY27 plan review comments - no stretch case.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0618am%20-%20FY27%20plan%20review%20comments%20-%20no%20stretch%20case.eml) — EML · 1.12 KiB · current-support · 120 body words
- [`7.5.26_0642am - capital cases after committee review.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0642am%20-%20capital%20cases%20after%20committee%20review.eml) — EML · 900 B · current-support · 90 body words
- [`7.5.26_0709am - bank follow-up + treasury version check.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0709am%20-%20bank%20follow-up%20%2B%20treasury%20version%20check.eml) — EML · 1020 B · current-support · 110 body words
- [`7.5.26_0727am - deal IC follow-up + banker case warning.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0727am%20-%20deal%20IC%20follow-up%20%2B%20banker%20case%20warning.eml) — EML · 1.04 KiB · current-support · 108 body words
- [`7.5.26_0751am - tax review comments + close version.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0751am%20-%20tax%20review%20comments%20%2B%20close%20version.eml) — EML · 930 B · current-support · 84 body words
- [`7.5.26_0812am - board review changes + source tie.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Requests/7.5.26_0812am%20-%20board%20review%20changes%20%2B%20source%20tie.eml) — EML · 941 B · current-support · 95 body words

### `Shared/Finance/Accounting Policies + old memos`

- [`Revenue recognition - WIP policy_rev11-24 SIGNED scan.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Accounting%20Policies%20%2B%20old%20memos/Revenue%20recognition%20-%20WIP%20policy_rev11-24%20SIGNED%20scan.pdf) — PDF · 52.11 KiB · current-support · 4 pages

### `Shared/Finance/Close/2026/05 May/4_WIP`

- [`WIP 5.31.26_FINAL_v7_revised NB.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/05%20May/4_WIP/WIP%205.31.26_FINAL_v7_revised%20NB.xlsx) — XLSX · 372.68 KiB · historical-final · 9 sheets · 5,598 populated rows · 207 formulas · 5,544 detected list records

### `Shared/Finance/Close/2026/06 June/1 close mgmt`

- [`June close tracker - working - upd 7.1 810am.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/1%20close%20mgmt/June%20close%20tracker%20-%20working%20-%20upd%207.1%20810am.xlsx) — XLSX · 14.26 KiB · current-working · 6 sheets · 82 populated rows · 0 formulas · 42 detected list records
- [`READ ME - where we are 7.1 810am.txt`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/1%20close%20mgmt/READ%20ME%20-%20where%20we%20are%207.1%20810am.txt) — TXT · 530 B · current-support · 6 lines

### `Shared/Finance/Close/2026/06 June/2 reconciliations`

- [`AR roll + retainage 0630 KS.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/2%20reconciliations/AR%20roll%20%2B%20retainage%200630%20KS.xlsx) — XLSX · 113.20 KiB · current-support · 4 sheets · 1,835 populated rows · 1,833 formulas · 1,811 detected list records
- [`TB rec_6.30 prelim - NB v4.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/2%20reconciliations/TB%20rec_6.30%20prelim%20-%20NB%20v4.xlsx) — XLSX · 16.81 KiB · current-working · 4 sheets · 138 populated rows · 154 formulas · 116 detected list records
- [`bank recs 6.30 - MP working.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/2%20reconciliations/bank%20recs%206.30%20-%20MP%20working.xlsx) — XLSX · 41.44 KiB · current-working · 4 sheets · 497 populated rows · 501 formulas · 471 detected list records

### `Shared/Finance/Close/2026/06 June/3 accruals`

- [`AP cutoff + accrual list_7.1 OL.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/3%20accruals/AP%20cutoff%20%2B%20accrual%20list_7.1%20OL.xlsx) — XLSX · 78.65 KiB · current-support · 4 sheets · 1,261 populated rows · 1,259 formulas · 1,237 detected list records
- [`payroll accrual 6.30 - FINAL? v3.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/3%20accruals/payroll%20accrual%206.30%20-%20FINAL%3F%20v3.xlsx) — XLSX · 16.65 KiB · current-working · 4 sheets · 140 populated rows · 150 formulas · 112 detected list records
- [`prepaids_amort sched FY26 - copy.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/3%20accruals/prepaids_amort%20sched%20FY26%20-%20copy.xlsx) — XLSX · 14.14 KiB · current-working · 4 sheets · 81 populated rows · 79 formulas · 57 detected list records

### `Shared/Finance/Close/2026/06 June/4 WIP`

- [`ARM-2409 June WIP controller sign-off - WORKING.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/4%20WIP/ARM-2409%20June%20WIP%20controller%20sign-off%20-%20WORKING.docx) — DOCX · 39.63 KiB · current-working · 4 rendered pages · 7 tables · 47 table rows
- [`WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/4%20WIP/WIP%20risk%20cases_7.1%20847am%20-%20NB%20REVIEW%20COPY.xlsx) — XLSX · 9.91 KiB · current-working · 3 sheets · 43 populated rows · 4 formulas

### `Shared/Finance/Close/2026/06 June/5 fixed assets`

- [`FA + CIP rollforward_June26 rev2.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Close/2026/06%20June/5%20fixed%20assets/FA%20%2B%20CIP%20rollforward_June26%20rev2.xlsx) — XLSX · 15.23 KiB · current-support · 4 sheets · 108 populated rows · 115 formulas · 81 detected list records

### `Shared/Finance/Controllership/Q2 working`

- [`FY26 tax + non-GAAP reporting policy - signed.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Controllership/Q2%20working/FY26%20tax%20%2B%20non-GAAP%20reporting%20policy%20-%20signed.pdf) — PDF · 10.32 KiB · current-working · 4 pages

### `Shared/Finance/Corporate Development`

- [`Orion integration value capture - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Orion%20integration%20value%20capture%20-%20WORKING.xlsx) — XLSX · 13.53 KiB · current-working · 7 sheets · 78 populated rows · 0 formulas · 20 detected list records

### `Shared/Finance/Corporate Development/Active analyses`

- [`IC diligence review notes - 7.5 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Active%20analyses/IC%20diligence%20review%20notes%20-%207.5%20DC.docx) — DOCX · 43.91 KiB · current-support · 8 rendered pages · 1 tables · 16 table rows
- [`IC model inputs - 6.29 banker case.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Active%20analyses/IC%20model%20inputs%20-%206.29%20banker%20case.xlsx) — XLSX · 27.42 KiB · current-support · 6 sheets · 253 populated rows · 12 formulas · 227 detected list records
- [`IC model inputs - 7.5 controller tie.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Active%20analyses/IC%20model%20inputs%20-%207.5%20controller%20tie.xlsx) — XLSX · 43.15 KiB · current-support · 6 sheets · 511 populated rows · 12 formulas · 485 detected list records
- [`M&A underwriting policy + IC return definitions.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Active%20analyses/M%26A%20underwriting%20policy%20%2B%20IC%20return%20definitions.pdf) — PDF · 13.48 KiB · current-support · 5 pages

### `Shared/Finance/Corporate Development/Data room exports`

- [`target monthly financials + diligence index_7.3.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Corporate%20Development/Data%20room%20exports/target%20monthly%20financials%20%2B%20diligence%20index_7.3.csv) — CSV · 30.19 KiB · current-support · 170 records · 13 columns

### `Shared/Finance/FP&A/Backlog`

- [`FY27 backlog burn and capacity - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Backlog/FY27%20backlog%20burn%20and%20capacity%20-%20WORKING.xlsx) — XLSX · 32.50 KiB · current-working · 8 sheets · 928 populated rows · 0 formulas · 197 detected list records
- [`backlog burn + award pipeline_6.30 v6.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Backlog/backlog%20burn%20%2B%20award%20pipeline_6.30%20v6.xlsx) — XLSX · 13.15 KiB · current-support · 4 sheets · 64 populated rows · 62 formulas · 40 detected list records

### `Shared/Finance/FP&A/Capex/2026 equipment requests`

- [`CapEx asks + ROI_2026 midyear - MW.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Capex/2026%20equipment%20requests/CapEx%20asks%20%2B%20ROI_2026%20midyear%20-%20MW.xlsx) — XLSX · 10.15 KiB · current-support · 4 sheets · 32 populated rows · 39 formulas
- [`midyear capex committee memo - draft 7.1.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Capex/2026%20equipment%20requests/midyear%20capex%20committee%20memo%20-%20draft%207.1.docx) — DOCX · 39.20 KiB · current-working · 2 rendered pages · 1 tables · 6 table rows

### `Shared/Finance/FP&A/FY26 plan + reforecast`

- [`FY26 AOP approval deck_12.18.25 FINAL.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY26%20plan%20%2B%20reforecast/FY26%20AOP%20approval%20deck_12.18.25%20FINAL.pptx) — PPTX · 32.25 KiB · approved-plan · 8 slides
- [`FY26 Op Plan_BoardApproved_12.18.25_FINAL2.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY26%20plan%20%2B%20reforecast/FY26%20Op%20Plan_BoardApproved_12.18.25_FINAL2.xlsx) — XLSX · 19.54 KiB · approved-plan · 9 sheets · 86 populated rows · 39 formulas · 36 detected list records
- [`FY26 rolling forecast_v12 - pre WIP.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY26%20plan%20%2B%20reforecast/FY26%20rolling%20forecast_v12%20-%20pre%20WIP.xlsx) — XLSX · 13.01 KiB · stale-working · 4 sheets · 87 populated rows · 94 formulas · 60 detected list records
- [`department reforecast inputs - June pull.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY26%20plan%20%2B%20reforecast/department%20reforecast%20inputs%20-%20June%20pull.xlsx) — XLSX · 12.96 KiB · current-support · 4 sheets · 87 populated rows · 94 formulas · 60 detected list records

### `Shared/Finance/FP&A/FY27 plan`

- [`FY27 EBITDA scenarios - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20EBITDA%20scenarios%20-%20WORKING.xlsx) — XLSX · 27.62 KiB · current-working · 7 sheets · 835 populated rows · 0 formulas · 95 detected list records
- [`FY27 field workforce plan - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20field%20workforce%20plan%20-%20WORKING.xlsx) — XLSX · 12.81 KiB · current-working · 5 sheets · 84 populated rows · 0 formulas · 54 detected list records
- [`FY27 planning assumptions - v4 branch draft.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20planning%20assumptions%20-%20v4%20branch%20draft.xlsx) — XLSX · 34.61 KiB · current-working · 7 sheets · 359 populated rows · 15 formulas · 329 detected list records
- [`FY27 planning assumptions - v6 controller tie.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20planning%20assumptions%20-%20v6%20controller%20tie.xlsx) — XLSX · 89.08 KiB · current-support · 7 sheets · 1,299 populated rows · 15 formulas · 1,269 detected list records
- [`FY27 planning definitions + scenario guardrails - APPROVED.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20planning%20definitions%20%2B%20scenario%20guardrails%20-%20APPROVED.pdf) — PDF · 12.75 KiB · current-support · 5 pages
- [`FY27 steering committee notes - 7.3 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/FY27%20plan/FY27%20steering%20committee%20notes%20-%207.3%20DC.docx) — DOCX · 43.63 KiB · current-support · 19 rendered pages · 1 tables · 18 table rows

### `Shared/Finance/FP&A/Labor`

- [`2026 field labor loaded rate build - est copy.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Labor/2026%20field%20labor%20loaded%20rate%20build%20-%20est%20copy.xlsx) — XLSX · 5.99 KiB · current-working · 1 sheets · 12 populated rows · 4 formulas
- [`HC + field labor productivity Q2 working.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Labor/HC%20%2B%20field%20labor%20productivity%20Q2%20working.xlsx) — XLSX · 66.45 KiB · current-working · 4 sheets · 960 populated rows · 958 formulas · 936 detected list records
- [`field burden rate approval - 2026 estimator use.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/FP%26A/Labor/field%20burden%20rate%20approval%20-%202026%20estimator%20use.pdf) — PDF · 49.77 KiB · current-support · 4 pages

### `Shared/Finance/Investor Relations/Peer analysis`

- [`Q2 peer market + competitor support - 6.27 banker pull.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Peer%20analysis/Q2%20peer%20market%20%2B%20competitor%20support%20-%206.27%20banker%20pull.xlsx) — XLSX · 11.60 KiB · current-support · 4 sheets · 68 populated rows · 6 formulas · 50 detected list records
- [`Q2 peer market + competitor support - 7.3 close.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Peer%20analysis/Q2%20peer%20market%20%2B%20competitor%20support%20-%207.3%20close.xlsx) — XLSX · 11.58 KiB · current-support · 4 sheets · 70 populated rows · 6 formulas · 52 detected list records
- [`market study + valuation methodology - approved.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Peer%20analysis/market%20study%20%2B%20valuation%20methodology%20-%20approved.pdf) — PDF · 4.88 KiB · current-support · 2 pages
- [`peer market data pull_7.3.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Peer%20analysis/peer%20market%20data%20pull_7.3.csv) — CSV · 14.99 KiB · current-support · 85 records · 13 columns
- [`peer review notes + comparability decisions - 7.3 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Peer%20analysis/peer%20review%20notes%20%2B%20comparability%20decisions%20-%207.3%20DC.docx) — DOCX · 39.32 KiB · current-support · 2 rendered pages · 1 tables · 6 table rows

### `Shared/Finance/Investor Relations/Q2 working`

- [`KPI definitions + lender presentation policy.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Investor%20Relations/Q2%20working/KPI%20definitions%20%2B%20lender%20presentation%20policy.pdf) — PDF · 8.17 KiB · current-working · 4 pages

### `Shared/Finance/Reporting/2026/05 May`

- [`May Ops Review - 6.5 mtg - FINAL_v4.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/05%20May/May%20Ops%20Review%20-%206.5%20mtg%20-%20FINAL_v4.pptx) — PPTX · 54.07 KiB · historical-final · 8 slides
- [`notes from 6.5 ops mtg - dc.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/05%20May/notes%20from%206.5%20ops%20mtg%20-%20dc.docx) — DOCX · 38.66 KiB · historical-final · 2 rendered pages · 1 tables · 5 table rows

### `Shared/Finance/Reporting/2026/06 June`

- [`June executive performance review - WORKING.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/June%20executive%20performance%20review%20-%20WORKING.pptx) — PPTX · 52.17 KiB · current-working · 9 slides
- [`June flash - CFO scratch v2 7.2.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/June%20flash%20-%20CFO%20scratch%20v2%207.2.pptx) — PPTX · 23.84 KiB · current-support · 5 slides
- [`June flash bridge - review copy 7.2.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/June%20flash%20bridge%20-%20review%20copy%207.2.xlsx) — XLSX · 7.97 KiB · current-working · 3 sheets · 27 populated rows · 3 formulas
- [`June flash review notes - DC 7.2 643am.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/June%20flash%20review%20notes%20-%20DC%207.2%20643am.docx) — DOCX · 40.00 KiB · current-support · 3 rendered pages · 1 tables · 17 table rows
- [`Q2 board + lender narrative review notes - 7.4 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/Q2%20board%20%2B%20lender%20narrative%20review%20notes%20-%207.4%20DC.docx) — DOCX · 40.75 KiB · current-support · 4 rendered pages · 1 tables · 8 table rows
- [`Q2 management reporting cube extract_7.2.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/Q2%20management%20reporting%20cube%20extract_7.2.csv) — CSV · 17.31 KiB · current-support · 78 records · 13 columns
- [`Q2 management reporting data book - v5 CFO scratch.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/Q2%20management%20reporting%20data%20book%20-%20v5%20CFO%20scratch.xlsx) — XLSX · 17.42 KiB · current-support · 4 sheets · 150 populated rows · 6 formulas · 132 detected list records
- [`Q2 management reporting data book - v7 controller tie.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/2026/06%20June/Q2%20management%20reporting%20data%20book%20-%20v7%20controller%20tie.xlsx) — XLSX · 33.72 KiB · current-support · 4 sheets · 407 populated rows · 6 formulas · 389 detected list records

### `Shared/Finance/Reporting/Board/Q1 2026`

- [`Q1 board package - FINAL signed off.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/Board/Q1%202026/Q1%20board%20package%20-%20FINAL%20signed%20off.pptx) — PPTX · 29.44 KiB · historical-final · 7 slides

### `Shared/Finance/Reporting/Board/Q2 2026`

- [`Board performance dashboard - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/Board/Q2%202026/Board%20performance%20dashboard%20-%20WORKING.xlsx) — XLSX · 14.03 KiB · current-working · 6 sheets · 103 populated rows · 0 formulas · 51 detected list records
- [`CFO narrative Q2 board pre-read - v3 comments.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Reporting/Board/Q2%202026/CFO%20narrative%20Q2%20board%20pre-read%20-%20v3%20comments.docx) — DOCX · 39.99 KiB · current-support · 2 rendered pages · 1 tables · 17 table rows

### `Shared/Finance/Risk + Insurance`

- [`2026 renewal exposure notes - broker call followup.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Risk%20%2B%20Insurance/2026%20renewal%20exposure%20notes%20-%20broker%20call%20followup.docx) — DOCX · 39.12 KiB · current-support · 2 rendered pages · 1 tables · 6 table rows
- [`2026-27 binder + schedule of coverage - broker draft.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Risk%20%2B%20Insurance/2026-27%20binder%20%2B%20schedule%20of%20coverage%20-%20broker%20draft.pdf) — PDF · 49.61 KiB · current-working · 4 pages
- [`bonding insurance schedule 2026 - renewal working.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Risk%20%2B%20Insurance/bonding%20insurance%20schedule%202026%20-%20renewal%20working.xlsx) — XLSX · 11.60 KiB · current-working · 4 sheets · 49 populated rows · 59 formulas · 21 detected list records

### `Shared/Finance/Strategic Finance/FY27 working`

- [`capital committee cases + LRP assumptions - v3 pre-review.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Strategic%20Finance/FY27%20working/capital%20committee%20cases%20%2B%20LRP%20assumptions%20-%20v3%20pre-review.xlsx) — XLSX · 17.25 KiB · current-working · 5 sheets · 131 populated rows · 9 formulas · 109 detected list records
- [`capital committee cases + LRP assumptions - v5.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Strategic%20Finance/FY27%20working/capital%20committee%20cases%20%2B%20LRP%20assumptions%20-%20v5.xlsx) — XLSX · 39.65 KiB · current-working · 5 sheets · 517 populated rows · 9 formulas · 495 detected list records
- [`capital committee review notes - 7.4 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Strategic%20Finance/FY27%20working/capital%20committee%20review%20notes%20-%207.4%20DC.docx) — DOCX · 40.28 KiB · current-working · 3 rendered pages · 1 tables · 8 table rows
- [`investment hurdle + capital guardrails - board approved.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Strategic%20Finance/FY27%20working/investment%20hurdle%20%2B%20capital%20guardrails%20-%20board%20approved.pdf) — PDF · 6.87 KiB · current-working · 3 pages

### `Shared/Finance/Tax/2026`

- [`1099 sales use tax tracker - 6.30.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026/1099%20sales%20use%20tax%20tracker%20-%206.30.xlsx) — XLSX · 18.83 KiB · current-support · 4 sheets · 177 populated rows · 184 formulas · 150 detected list records
- [`FY26 tax provision - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026/FY26%20tax%20provision%20-%20WORKING.xlsx) — XLSX · 12.91 KiB · current-working · 6 sheets · 98 populated rows · 0 formulas · 19 detected list records
- [`OR DOR desk review notice + response checklist.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026/OR%20DOR%20desk%20review%20notice%20%2B%20response%20checklist.pdf) — PDF · 49.52 KiB · current-support · 4 pages

### `Shared/Finance/Tax/2026 working`

- [`FY26 tax + non-GAAP workpapers - v2 pre-close.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026%20working/FY26%20tax%20%2B%20non-GAAP%20workpapers%20-%20v2%20pre-close.xlsx) — XLSX · 20.24 KiB · current-working · 5 sheets · 168 populated rows · 9 formulas · 146 detected list records
- [`FY26 tax + non-GAAP workpapers - v4 controller review.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026%20working/FY26%20tax%20%2B%20non-GAAP%20workpapers%20-%20v4%20controller%20review.xlsx) — XLSX · 29.04 KiB · current-working · 5 sheets · 316 populated rows · 9 formulas · 294 detected list records
- [`tax jurisdiction + book trial balance detail_6.30.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026%20working/tax%20jurisdiction%20%2B%20book%20trial%20balance%20detail_6.30.csv) — CSV · 17.43 KiB · current-working · 88 records · 13 columns
- [`tax provision review notes - 7.4 NB.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Tax/2026%20working/tax%20provision%20review%20notes%20-%207.4%20NB.docx) — DOCX · 42.19 KiB · current-working · 5 rendered pages · 1 tables · 10 table rows

### `Shared/Finance/Treasury/13 week cash`

- [`13wk cash_v9 - 6.26 roll - DCHO edits.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/13%20week%20cash/13wk%20cash_v9%20-%206.26%20roll%20-%20DCHO%20edits.xlsx) — XLSX · 130.81 KiB · stale-working · 9 sheets · 2,010 populated rows · 1,576 formulas · 1,956 detected list records
- [`downside assumptions_7.2 618am - DC marks.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/13%20week%20cash/downside%20assumptions_7.2%20618am%20-%20DC%20marks.xlsx) — XLSX · 10.08 KiB · current-support · 3 sheets · 48 populated rows · 3 formulas · 13 detected list records
- [`downside liquidity sensitivity_7.2 618am - current base treatment.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/13%20week%20cash/downside%20liquidity%20sensitivity_7.2%20618am%20-%20current%20base%20treatment.xlsx) — XLSX · 8.52 KiB · current-support · 3 sheets · 37 populated rows · 0 formulas

### `Shared/Finance/Treasury/AR calls - notes`

- [`AR notes_6.30 - NB working copy.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/AR%20calls%20-%20notes/AR%20notes_6.30%20-%20NB%20working%20copy.docx) — DOCX · 44.51 KiB · current-working · 8 rendered pages · 6 tables · 77 table rows

### `Shared/Finance/Treasury/Bank - covenants`

- [`FY27 covenant forecast - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/FY27%20covenant%20forecast%20-%20WORKING.xlsx) — XLSX · 14.00 KiB · current-working · 5 sheets · 118 populated rows · 0 formulas · 72 detected list records

### `Shared/Finance/Treasury/Bank - covenants/2026 Q1 submitted`

- [`Q1 2026 compliance pkg - submitted 4.28.26.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q1%20submitted/Q1%202026%20compliance%20pkg%20-%20submitted%204.28.26.pdf) — PDF · 59.25 KiB · historical-final · 12 pages
- [`covenant summary - old 2024 (use agreement).pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q1%20submitted/covenant%20summary%20-%20old%202024%20%28use%20agreement%29.pdf) — PDF · 43.84 KiB · superseded · 2 pages
- [`lender + surety update Q1 - submitted copy.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q1%20submitted/lender%20%2B%20surety%20update%20Q1%20-%20submitted%20copy.pptx) — PPTX · 26.52 KiB · historical-final · 6 slides

### `Shared/Finance/Treasury/Bank - covenants/2026 Q2 working`

- [`Q2 covenant headroom - lender review working.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q2%20working/Q2%20covenant%20headroom%20-%20lender%20review%20working.xlsx) — XLSX · 7.05 KiB · current-working · 2 sheets · 17 populated rows · 3 formulas
- [`Q2 lender update - review working v3.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q2%20working/Q2%20lender%20update%20-%20review%20working%20v3.pptx) — PPTX · 23.83 KiB · current-working · 5 slides
- [`public-owner assignment status - lender reply 7.2.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/2026%20Q2%20working/public-owner%20assignment%20status%20-%20lender%20reply%207.2.eml) — EML · 1.91 KiB · current-working · 250 body words

### `Shared/Finance/Treasury/Bank - covenants/Agreement + amendments`

- [`USBank AR eligibility exhibit - closing set copy.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/Agreement%20%2B%20amendments/USBank%20AR%20eligibility%20exhibit%20-%20closing%20set%20copy.pdf) — PDF · 69.79 KiB · current-working · 4 pages
- [`USBank_Amdt2_9.30.25_EXECUTED scan.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Bank%20-%20covenants/Agreement%20%2B%20amendments/USBank_Amdt2_9.30.25_EXECUTED%20scan.pdf) — PDF · 55.62 KiB · current-support · 9 pages

### `Shared/Finance/Treasury/Cash positioning`

- [`July daily cash position - WORKING.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Cash%20positioning/July%20daily%20cash%20position%20-%20WORKING.xlsx) — XLSX · 14.29 KiB · current-working · 6 sheets · 104 populated rows · 0 formulas · 69 detected list records
- [`bank portal settlements + expected items_7.3.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Cash%20positioning/bank%20portal%20settlements%20%2B%20expected%20items_7.3.csv) — CSV · 37.83 KiB · current-support · 208 records · 13 columns

### `Shared/Finance/Treasury/Debt`

- [`debt sched - 6.30 before bank pkg.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Debt/debt%20sched%20-%206.30%20before%20bank%20pkg.xlsx) — XLSX · 9.74 KiB · current-support · 4 sheets · 28 populated rows · 23 formulas
- [`equipment line proposal - bank copy 6.18.26.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/Debt/equipment%20line%20proposal%20-%20bank%20copy%206.18.26.pdf) — PDF · 49.53 KiB · current-working · 4 pages

### `Shared/Finance/Treasury/FY27 working`

- [`treasury and bank review notes - 7.4 DC.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/FY27%20working/treasury%20and%20bank%20review%20notes%20-%207.4%20DC.docx) — DOCX · 43.38 KiB · current-working · 6 rendered pages · 1 tables · 16 table rows
- [`treasury outlook support - July v6 pre-bank.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/FY27%20working/treasury%20outlook%20support%20-%20July%20v6%20pre-bank.xlsx) — XLSX · 26.93 KiB · current-working · 6 sheets · 251 populated rows · 12 formulas · 225 detected list records
- [`treasury outlook support - July v8 controller tie.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/FY27%20working/treasury%20outlook%20support%20-%20July%20v8%20controller%20tie.xlsx) — XLSX · 47.85 KiB · current-working · 6 sheets · 601 populated rows · 12 formulas · 575 detected list records
- [`treasury risk limits + funding policy - board approved.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Finance/Treasury/FY27%20working/treasury%20risk%20limits%20%2B%20funding%20policy%20-%20board%20approved.pdf) — PDF · 12.77 KiB · current-working · 5 pages

### `Shared/Operations/Commercial/CO Log`

- [`CO log MASTER (do not sort)_6.30.26 - PR copy.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Commercial/CO%20Log/CO%20log%20MASTER%20%28do%20not%20sort%29_6.30.26%20-%20PR%20copy.xlsx) — XLSX · 13.66 KiB · current-working · 5 sheets · 55 populated rows · 23 formulas · 24 detected list records

### `Shared/Operations/Commercial/CO Log/6.30 support - not all final`

- [`2409_PCO11_backcharge backup - draft2.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Commercial/CO%20Log/6.30%20support%20-%20not%20all%20final/2409_PCO11_backcharge%20backup%20-%20draft2.pdf) — PDF · 49.66 KiB · current-working · 4 pages
- [`2417_PCO17_owner emails + FD summary_6.30.pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Commercial/CO%20Log/6.30%20support%20-%20not%20all%20final/2417_PCO17_owner%20emails%20%2B%20FD%20summary_6.30.pdf) — PDF · 69.49 KiB · current-support · 4 pages
- [`2506_PCO6_DB27 backup + cost est (working).pdf`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Commercial/CO%20Log/6.30%20support%20-%20not%20all%20final/2506_PCO6_DB27%20backup%20%2B%20cost%20est%20%28working%29.pdf) — PDF · 49.65 KiB · current-working · 4 pages

### `Shared/Operations/Fleet`

- [`fleet + equipment utilization_June working.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Fleet/fleet%20%2B%20equipment%20utilization_June%20working.xlsx) — XLSX · 11.48 KiB · current-working · 4 sheets · 58 populated rows · 56 formulas · 34 detected list records

### `Shared/Operations/Planning exports`

- [`CRM backlog service renewals + forecast history_6.30.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Planning%20exports/CRM%20backlog%20service%20renewals%20%2B%20forecast%20history_6.30.csv) — CSV · 39.93 KiB · current-support · 183 records · 13 columns

### `Shared/Operations/Project Controls/cash curves`

- [`major jobs cash curves_6.30 scenario B.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Project%20Controls/cash%20curves/major%20jobs%20cash%20curves_6.30%20scenario%20B.xlsx) — XLSX · 20.45 KiB · current-support · 4 sheets · 231 populated rows · 226 formulas · 208 detected list records

### `Shared/Operations/Project Forecasts/June26 - PM updates`

- [`PM ETC updates_6.29 530pm_COMBINED_v3.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Project%20Forecasts/June26%20-%20PM%20updates/PM%20ETC%20updates_6.29%20530pm_COMBINED_v3.xlsx) — XLSX · 16.47 KiB · current-support · 6 sheets · 95 populated rows · 32 formulas · 72 detected list records

### `Shared/Operations/Project Reviews`

- [`weekly top jobs review - 6.29 PM copy.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Project%20Reviews/weekly%20top%20jobs%20review%20-%206.29%20PM%20copy.pptx) — PPTX · 29.38 KiB · current-working · 7 slides

### `Shared/Operations/Quarterly Reviews`

- [`Q2 ops + safety review - 6.26 draft.pptx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Quarterly%20Reviews/Q2%20ops%20%2B%20safety%20review%20-%206.26%20draft.pptx) — PPTX · 29.28 KiB · stale-working · 7 slides
- [`Q2 project execution status and rework review protocol - 6.30 approved.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Quarterly%20Reviews/Q2%20project%20execution%20status%20and%20rework%20review%20protocol%20-%206.30%20approved.xlsx) — XLSX · 6.97 KiB · current-support · 2 sheets · 36 populated rows · 4 formulas · 30 detected list records

### `Shared/Operations/Service/monthly KPI`

- [`Q2 service pricing + callback cut - 7.1 SA.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/Q2%20service%20pricing%20%2B%20callback%20cut%20-%207.1%20SA.xlsx) — XLSX · 66.03 KiB · current-support · 3 sheets · 841 populated rows · 6 formulas · 824 detected list records
- [`service KPI pack source_6.30 - LT edits.xlsx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20KPI%20pack%20source_6.30%20-%20LT%20edits.xlsx) — XLSX · 65.25 KiB · current-working · 4 sheets · 854 populated rows · 870 formulas · 832 detected list records
- [`service access signal out-of-time observations - 7.10 DC.csv`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20access%20signal%20out-of-time%20observations%20-%207.10%20DC.csv) — CSV · 9.70 KiB · current-support · 240 records · 7 columns
- [`service access signal production-readiness review - 7.10 DC.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20access%20signal%20production-readiness%20review%20-%207.10%20DC.eml) — EML · 2.16 KiB · current-support · 290 body words
- [`service callback action window - 7.6 DC.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20callback%20action%20window%20-%207.6%20DC.eml) — EML · 2.42 KiB · current-support · 336 body words
- [`service callback evidence follow-up - 7.3 SA.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20callback%20evidence%20follow-up%20-%207.3%20SA.eml) — EML · 2.94 KiB · current-support · 396 body words
- [`service commitment timing decision - 7.8 DC.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20commitment%20timing%20decision%20-%207.8%20DC.eml) — EML · 4.40 KiB · current-support · 693 body words
- [`service operating-review allocation economics - 7.7 DC.eml`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20operating-review%20allocation%20economics%20-%207.7%20DC.eml) — EML · 6.02 KiB · current-support · 914 body words
- [`service pricing + callback review notes - SA 7.1.docx`](https://github.com/Sureform-YC-X25/alder-ridge-sample-tasks/blob/main/environment/seed/sources/Shared/Operations/Service/monthly%20KPI/service%20pricing%20%2B%20callback%20review%20notes%20-%20SA%207.1.docx) — DOCX · 39.99 KiB · current-support · 3 rendered pages · 1 tables · 17 table rows

## Measurement notes

- PDF pages are read from the PDF page tree.
- Word pages are measured after headless LibreOffice rendering and may differ slightly from Microsoft Word pagination.
- Excel populated rows include model rows, summaries, assumptions, checks, and source data—not only independent records.
- Accounting counts are physical table rows. A single economic event can appear in a header table, one or more detail tables, and the general ledger.
- File sizes reflect compressed source files and do not measure analytical difficulty.

Generated by `environment/tools/generate_seed_inventory.py`.
