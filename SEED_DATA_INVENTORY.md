# Alder Ridge Seed Data Inventory

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

## Company-world source files

### `Deliverables`

- `ARM-2417 June close and PCO-017 review.docx` — DOCX · 39.71 KiB · current-support · 2 rendered pages · 6 tables · 33 table rows

### `Requests`

- `7.1.26_822am - Fwd_ wip close need today.eml` — EML · 1.98 KiB · current-support · 313 body words
- `7.1.26_904am - cash refresh + Fri deck.eml` — EML · 895 B · current-support · 112 body words
- `7.2.26_1027am - Q2 bank compliance package.eml` — EML · 946 B · current-support · 110 body words
- `7.2.26_1103am - CFO close decision note - internal only.eml` — EML · 789 B · current-support · 87 body words
- `7.2.26_611am - June ops review deck + notes.eml` — EML · 697 B · current-support · 68 body words
- `7.2.26_736am - FW_ wc cleanup + backlog coverage.eml` — EML · 963 B · current-support · 109 body words
- `7.2.26_805am - payroll bridge + capex before noon.eml` — EML · 1.06 KiB · current-support · 128 body words
- `7.2.26_842am - service vans capacity check - LT fwd.eml` — EML · 952 B · current-support · 102 body words
- `7.4.26_0711am - FY27 plan first pass + branch submissions.eml` — EML · 1.53 KiB · current-support · 186 body words
- `7.4.26_0838am - capital committee + five year plan.eml` — EML · 1.13 KiB · current-support · 115 body words
- `7.4.26_0956am - July liquidity bank work + covenant outlook.eml` — EML · 1.29 KiB · current-support · 140 body words
- `7.4.26_1106am - corp dev IC workstream status.eml` — EML · 1.26 KiB · current-support · 139 body words
- `7.4.26_1216pm - tax provision + EBITDA definitions.eml` — EML · 1.16 KiB · current-support · 125 body words
- `7.4.26_1344pm - Q2 board lender + IR refresh.eml` — EML · 1.15 KiB · current-support · 124 body words
- `7.5.26_0618am - FY27 plan review comments - no stretch case.eml` — EML · 1.12 KiB · current-support · 120 body words
- `7.5.26_0642am - capital cases after committee review.eml` — EML · 900 B · current-support · 90 body words
- `7.5.26_0709am - bank follow-up + treasury version check.eml` — EML · 1020 B · current-support · 110 body words
- `7.5.26_0727am - deal IC follow-up + banker case warning.eml` — EML · 1.04 KiB · current-support · 108 body words
- `7.5.26_0751am - tax review comments + close version.eml` — EML · 930 B · current-support · 84 body words
- `7.5.26_0812am - board review changes + source tie.eml` — EML · 941 B · current-support · 95 body words

### `Shared/Finance/Accounting Policies + old memos`

- `Revenue recognition - WIP policy_rev11-24 SIGNED scan.pdf` — PDF · 52.11 KiB · current-support · 4 pages

### `Shared/Finance/Close/2026/05 May/4_WIP`

- `WIP 5.31.26_FINAL_v7_revised NB.xlsx` — XLSX · 372.68 KiB · historical-final · 9 sheets · 5,598 populated rows · 207 formulas · 5,544 detected list records

### `Shared/Finance/Close/2026/06 June/1 close mgmt`

- `June close tracker - working - upd 7.1 810am.xlsx` — XLSX · 14.26 KiB · current-working · 6 sheets · 82 populated rows · 0 formulas · 42 detected list records
- `READ ME - where we are 7.1 810am.txt` — TXT · 530 B · current-support · 6 lines

### `Shared/Finance/Close/2026/06 June/2 reconciliations`

- `AR roll + retainage 0630 KS.xlsx` — XLSX · 113.20 KiB · current-support · 4 sheets · 1,835 populated rows · 1,833 formulas · 1,811 detected list records
- `TB rec_6.30 prelim - NB v4.xlsx` — XLSX · 16.81 KiB · current-working · 4 sheets · 138 populated rows · 154 formulas · 116 detected list records
- `bank recs 6.30 - MP working.xlsx` — XLSX · 41.44 KiB · current-working · 4 sheets · 497 populated rows · 501 formulas · 471 detected list records

### `Shared/Finance/Close/2026/06 June/3 accruals`

- `AP cutoff + accrual list_7.1 OL.xlsx` — XLSX · 78.65 KiB · current-support · 4 sheets · 1,261 populated rows · 1,259 formulas · 1,237 detected list records
- `payroll accrual 6.30 - FINAL? v3.xlsx` — XLSX · 16.65 KiB · current-working · 4 sheets · 140 populated rows · 150 formulas · 112 detected list records
- `prepaids_amort sched FY26 - copy.xlsx` — XLSX · 14.14 KiB · current-working · 4 sheets · 81 populated rows · 79 formulas · 57 detected list records

### `Shared/Finance/Close/2026/06 June/4 WIP`

- `ARM-2409 June WIP controller sign-off - WORKING.docx` — DOCX · 39.63 KiB · current-working · 4 rendered pages · 7 tables · 47 table rows
- `WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx` — XLSX · 9.91 KiB · current-working · 3 sheets · 43 populated rows · 4 formulas

### `Shared/Finance/Close/2026/06 June/5 fixed assets`

- `FA + CIP rollforward_June26 rev2.xlsx` — XLSX · 15.23 KiB · current-support · 4 sheets · 108 populated rows · 115 formulas · 81 detected list records

### `Shared/Finance/Controllership/Q2 working`

- `FY26 tax + non-GAAP reporting policy - signed.pdf` — PDF · 10.32 KiB · current-working · 4 pages

### `Shared/Finance/Corporate Development`

- `Orion integration value capture - WORKING.xlsx` — XLSX · 13.53 KiB · current-working · 7 sheets · 78 populated rows · 0 formulas · 20 detected list records

### `Shared/Finance/Corporate Development/Active analyses`

- `IC diligence review notes - 7.5 DC.docx` — DOCX · 43.91 KiB · current-support · 8 rendered pages · 1 tables · 16 table rows
- `IC model inputs - 6.29 banker case.xlsx` — XLSX · 27.42 KiB · current-support · 6 sheets · 253 populated rows · 12 formulas · 227 detected list records
- `IC model inputs - 7.5 controller tie.xlsx` — XLSX · 43.15 KiB · current-support · 6 sheets · 511 populated rows · 12 formulas · 485 detected list records
- `M&A underwriting policy + IC return definitions.pdf` — PDF · 13.48 KiB · current-support · 5 pages

### `Shared/Finance/Corporate Development/Data room exports`

- `target monthly financials + diligence index_7.3.csv` — CSV · 30.19 KiB · current-support · 170 records · 13 columns

### `Shared/Finance/FP&A/Backlog`

- `FY27 backlog burn and capacity - WORKING.xlsx` — XLSX · 32.50 KiB · current-working · 8 sheets · 928 populated rows · 0 formulas · 197 detected list records
- `backlog burn + award pipeline_6.30 v6.xlsx` — XLSX · 13.15 KiB · current-support · 4 sheets · 64 populated rows · 62 formulas · 40 detected list records

### `Shared/Finance/FP&A/Capex/2026 equipment requests`

- `CapEx asks + ROI_2026 midyear - MW.xlsx` — XLSX · 10.15 KiB · current-support · 4 sheets · 32 populated rows · 39 formulas
- `midyear capex committee memo - draft 7.1.docx` — DOCX · 39.20 KiB · current-working · 2 rendered pages · 1 tables · 6 table rows

### `Shared/Finance/FP&A/FY26 plan + reforecast`

- `FY26 AOP approval deck_12.18.25 FINAL.pptx` — PPTX · 32.25 KiB · approved-plan · 8 slides
- `FY26 Op Plan_BoardApproved_12.18.25_FINAL2.xlsx` — XLSX · 19.54 KiB · approved-plan · 9 sheets · 86 populated rows · 39 formulas · 36 detected list records
- `FY26 rolling forecast_v12 - pre WIP.xlsx` — XLSX · 13.01 KiB · stale-working · 4 sheets · 87 populated rows · 94 formulas · 60 detected list records
- `department reforecast inputs - June pull.xlsx` — XLSX · 12.96 KiB · current-support · 4 sheets · 87 populated rows · 94 formulas · 60 detected list records

### `Shared/Finance/FP&A/FY27 plan`

- `FY27 EBITDA scenarios - WORKING.xlsx` — XLSX · 27.62 KiB · current-working · 7 sheets · 835 populated rows · 0 formulas · 95 detected list records
- `FY27 field workforce plan - WORKING.xlsx` — XLSX · 12.81 KiB · current-working · 5 sheets · 84 populated rows · 0 formulas · 54 detected list records
- `FY27 planning assumptions - v4 branch draft.xlsx` — XLSX · 34.61 KiB · current-working · 7 sheets · 359 populated rows · 15 formulas · 329 detected list records
- `FY27 planning assumptions - v6 controller tie.xlsx` — XLSX · 89.08 KiB · current-support · 7 sheets · 1,299 populated rows · 15 formulas · 1,269 detected list records
- `FY27 planning definitions + scenario guardrails - APPROVED.pdf` — PDF · 12.75 KiB · current-support · 5 pages
- `FY27 steering committee notes - 7.3 DC.docx` — DOCX · 43.63 KiB · current-support · 19 rendered pages · 1 tables · 18 table rows

### `Shared/Finance/FP&A/Labor`

- `2026 field labor loaded rate build - est copy.xlsx` — XLSX · 5.99 KiB · current-working · 1 sheets · 12 populated rows · 4 formulas
- `HC + field labor productivity Q2 working.xlsx` — XLSX · 66.45 KiB · current-working · 4 sheets · 960 populated rows · 958 formulas · 936 detected list records
- `field burden rate approval - 2026 estimator use.pdf` — PDF · 49.77 KiB · current-support · 4 pages

### `Shared/Finance/Investor Relations/Peer analysis`

- `Q2 peer market + competitor support - 6.27 banker pull.xlsx` — XLSX · 11.60 KiB · current-support · 4 sheets · 68 populated rows · 6 formulas · 50 detected list records
- `Q2 peer market + competitor support - 7.3 close.xlsx` — XLSX · 11.58 KiB · current-support · 4 sheets · 70 populated rows · 6 formulas · 52 detected list records
- `market study + valuation methodology - approved.pdf` — PDF · 4.88 KiB · current-support · 2 pages
- `peer market data pull_7.3.csv` — CSV · 14.99 KiB · current-support · 85 records · 13 columns
- `peer review notes + comparability decisions - 7.3 DC.docx` — DOCX · 39.32 KiB · current-support · 2 rendered pages · 1 tables · 6 table rows

### `Shared/Finance/Investor Relations/Q2 working`

- `KPI definitions + lender presentation policy.pdf` — PDF · 8.17 KiB · current-working · 4 pages

### `Shared/Finance/Reporting/2026/05 May`

- `May Ops Review - 6.5 mtg - FINAL_v4.pptx` — PPTX · 54.07 KiB · historical-final · 8 slides
- `notes from 6.5 ops mtg - dc.docx` — DOCX · 38.66 KiB · historical-final · 2 rendered pages · 1 tables · 5 table rows

### `Shared/Finance/Reporting/2026/06 June`

- `June executive performance review - WORKING.pptx` — PPTX · 52.17 KiB · current-working · 9 slides
- `June flash - CFO scratch v2 7.2.pptx` — PPTX · 23.84 KiB · current-support · 5 slides
- `June flash bridge - review copy 7.2.xlsx` — XLSX · 7.97 KiB · current-working · 3 sheets · 27 populated rows · 3 formulas
- `June flash review notes - DC 7.2 643am.docx` — DOCX · 40.00 KiB · current-support · 3 rendered pages · 1 tables · 17 table rows
- `Q2 board + lender narrative review notes - 7.4 DC.docx` — DOCX · 40.75 KiB · current-support · 4 rendered pages · 1 tables · 8 table rows
- `Q2 management reporting cube extract_7.2.csv` — CSV · 17.31 KiB · current-support · 78 records · 13 columns
- `Q2 management reporting data book - v5 CFO scratch.xlsx` — XLSX · 17.42 KiB · current-support · 4 sheets · 150 populated rows · 6 formulas · 132 detected list records
- `Q2 management reporting data book - v7 controller tie.xlsx` — XLSX · 33.72 KiB · current-support · 4 sheets · 407 populated rows · 6 formulas · 389 detected list records

### `Shared/Finance/Reporting/Board/Q1 2026`

- `Q1 board package - FINAL signed off.pptx` — PPTX · 29.44 KiB · historical-final · 7 slides

### `Shared/Finance/Reporting/Board/Q2 2026`

- `Board performance dashboard - WORKING.xlsx` — XLSX · 14.03 KiB · current-working · 6 sheets · 103 populated rows · 0 formulas · 51 detected list records
- `CFO narrative Q2 board pre-read - v3 comments.docx` — DOCX · 39.99 KiB · current-support · 2 rendered pages · 1 tables · 17 table rows

### `Shared/Finance/Risk + Insurance`

- `2026 renewal exposure notes - broker call followup.docx` — DOCX · 39.12 KiB · current-support · 2 rendered pages · 1 tables · 6 table rows
- `2026-27 binder + schedule of coverage - broker draft.pdf` — PDF · 49.61 KiB · current-working · 4 pages
- `bonding insurance schedule 2026 - renewal working.xlsx` — XLSX · 11.60 KiB · current-working · 4 sheets · 49 populated rows · 59 formulas · 21 detected list records

### `Shared/Finance/Strategic Finance/FY27 working`

- `capital committee cases + LRP assumptions - v3 pre-review.xlsx` — XLSX · 17.25 KiB · current-working · 5 sheets · 131 populated rows · 9 formulas · 109 detected list records
- `capital committee cases + LRP assumptions - v5.xlsx` — XLSX · 39.65 KiB · current-working · 5 sheets · 517 populated rows · 9 formulas · 495 detected list records
- `capital committee review notes - 7.4 DC.docx` — DOCX · 40.28 KiB · current-working · 3 rendered pages · 1 tables · 8 table rows
- `investment hurdle + capital guardrails - board approved.pdf` — PDF · 6.87 KiB · current-working · 3 pages

### `Shared/Finance/Tax/2026`

- `1099 sales use tax tracker - 6.30.xlsx` — XLSX · 18.83 KiB · current-support · 4 sheets · 177 populated rows · 184 formulas · 150 detected list records
- `FY26 tax provision - WORKING.xlsx` — XLSX · 12.91 KiB · current-working · 6 sheets · 98 populated rows · 0 formulas · 19 detected list records
- `OR DOR desk review notice + response checklist.pdf` — PDF · 49.52 KiB · current-support · 4 pages

### `Shared/Finance/Tax/2026 working`

- `FY26 tax + non-GAAP workpapers - v2 pre-close.xlsx` — XLSX · 20.24 KiB · current-working · 5 sheets · 168 populated rows · 9 formulas · 146 detected list records
- `FY26 tax + non-GAAP workpapers - v4 controller review.xlsx` — XLSX · 29.04 KiB · current-working · 5 sheets · 316 populated rows · 9 formulas · 294 detected list records
- `tax jurisdiction + book trial balance detail_6.30.csv` — CSV · 17.43 KiB · current-working · 88 records · 13 columns
- `tax provision review notes - 7.4 NB.docx` — DOCX · 42.19 KiB · current-working · 5 rendered pages · 1 tables · 10 table rows

### `Shared/Finance/Treasury/13 week cash`

- `13wk cash_v9 - 6.26 roll - DCHO edits.xlsx` — XLSX · 130.81 KiB · stale-working · 9 sheets · 2,010 populated rows · 1,576 formulas · 1,956 detected list records
- `downside assumptions_7.2 618am - DC marks.xlsx` — XLSX · 10.08 KiB · current-support · 3 sheets · 48 populated rows · 3 formulas · 13 detected list records
- `downside liquidity sensitivity_7.2 618am - current base treatment.xlsx` — XLSX · 8.52 KiB · current-support · 3 sheets · 37 populated rows · 0 formulas

### `Shared/Finance/Treasury/AR calls - notes`

- `AR notes_6.30 - NB working copy.docx` — DOCX · 44.51 KiB · current-working · 8 rendered pages · 6 tables · 77 table rows

### `Shared/Finance/Treasury/Bank - covenants`

- `FY27 covenant forecast - WORKING.xlsx` — XLSX · 14.00 KiB · current-working · 5 sheets · 118 populated rows · 0 formulas · 72 detected list records

### `Shared/Finance/Treasury/Bank - covenants/2026 Q1 submitted`

- `Q1 2026 compliance pkg - submitted 4.28.26.pdf` — PDF · 59.25 KiB · historical-final · 12 pages
- `covenant summary - old 2024 (use agreement).pdf` — PDF · 43.84 KiB · superseded · 2 pages
- `lender + surety update Q1 - submitted copy.pptx` — PPTX · 26.52 KiB · historical-final · 6 slides

### `Shared/Finance/Treasury/Bank - covenants/2026 Q2 working`

- `Q2 covenant headroom - lender review working.xlsx` — XLSX · 7.05 KiB · current-working · 2 sheets · 17 populated rows · 3 formulas
- `Q2 lender update - review working v3.pptx` — PPTX · 23.83 KiB · current-working · 5 slides
- `public-owner assignment status - lender reply 7.2.eml` — EML · 1.91 KiB · current-working · 250 body words

### `Shared/Finance/Treasury/Bank - covenants/Agreement + amendments`

- `USBank AR eligibility exhibit - closing set copy.pdf` — PDF · 69.79 KiB · current-working · 4 pages
- `USBank_Amdt2_9.30.25_EXECUTED scan.pdf` — PDF · 55.62 KiB · current-support · 9 pages

### `Shared/Finance/Treasury/Cash positioning`

- `July daily cash position - WORKING.xlsx` — XLSX · 14.29 KiB · current-working · 6 sheets · 104 populated rows · 0 formulas · 69 detected list records
- `bank portal settlements + expected items_7.3.csv` — CSV · 37.83 KiB · current-support · 208 records · 13 columns

### `Shared/Finance/Treasury/Debt`

- `debt sched - 6.30 before bank pkg.xlsx` — XLSX · 9.74 KiB · current-support · 4 sheets · 28 populated rows · 23 formulas
- `equipment line proposal - bank copy 6.18.26.pdf` — PDF · 49.53 KiB · current-working · 4 pages

### `Shared/Finance/Treasury/FY27 working`

- `treasury and bank review notes - 7.4 DC.docx` — DOCX · 43.38 KiB · current-working · 6 rendered pages · 1 tables · 16 table rows
- `treasury outlook support - July v6 pre-bank.xlsx` — XLSX · 26.93 KiB · current-working · 6 sheets · 251 populated rows · 12 formulas · 225 detected list records
- `treasury outlook support - July v8 controller tie.xlsx` — XLSX · 47.85 KiB · current-working · 6 sheets · 601 populated rows · 12 formulas · 575 detected list records
- `treasury risk limits + funding policy - board approved.pdf` — PDF · 12.77 KiB · current-working · 5 pages

### `Shared/Operations/Commercial/CO Log`

- `CO log MASTER (do not sort)_6.30.26 - PR copy.xlsx` — XLSX · 13.66 KiB · current-working · 5 sheets · 55 populated rows · 23 formulas · 24 detected list records

### `Shared/Operations/Commercial/CO Log/6.30 support - not all final`

- `2409_PCO11_backcharge backup - draft2.pdf` — PDF · 49.66 KiB · current-working · 4 pages
- `2417_PCO17_owner emails + FD summary_6.30.pdf` — PDF · 69.49 KiB · current-support · 4 pages
- `2506_PCO6_DB27 backup + cost est (working).pdf` — PDF · 49.65 KiB · current-working · 4 pages

### `Shared/Operations/Fleet`

- `fleet + equipment utilization_June working.xlsx` — XLSX · 11.48 KiB · current-working · 4 sheets · 58 populated rows · 56 formulas · 34 detected list records

### `Shared/Operations/Planning exports`

- `CRM backlog service renewals + forecast history_6.30.csv` — CSV · 39.93 KiB · current-support · 183 records · 13 columns

### `Shared/Operations/Project Controls/cash curves`

- `major jobs cash curves_6.30 scenario B.xlsx` — XLSX · 20.45 KiB · current-support · 4 sheets · 231 populated rows · 226 formulas · 208 detected list records

### `Shared/Operations/Project Forecasts/June26 - PM updates`

- `PM ETC updates_6.29 530pm_COMBINED_v3.xlsx` — XLSX · 16.47 KiB · current-support · 6 sheets · 95 populated rows · 32 formulas · 72 detected list records

### `Shared/Operations/Project Reviews`

- `weekly top jobs review - 6.29 PM copy.pptx` — PPTX · 29.38 KiB · current-working · 7 slides

### `Shared/Operations/Quarterly Reviews`

- `Q2 ops + safety review - 6.26 draft.pptx` — PPTX · 29.28 KiB · stale-working · 7 slides
- `Q2 project execution status and rework review protocol - 6.30 approved.xlsx` — XLSX · 6.97 KiB · current-support · 2 sheets · 36 populated rows · 4 formulas · 30 detected list records

### `Shared/Operations/Service/monthly KPI`

- `Q2 service pricing + callback cut - 7.1 SA.xlsx` — XLSX · 66.03 KiB · current-support · 3 sheets · 841 populated rows · 6 formulas · 824 detected list records
- `service KPI pack source_6.30 - LT edits.xlsx` — XLSX · 65.25 KiB · current-working · 4 sheets · 854 populated rows · 870 formulas · 832 detected list records
- `service access signal out-of-time observations - 7.10 DC.csv` — CSV · 9.70 KiB · current-support · 240 records · 7 columns
- `service access signal production-readiness review - 7.10 DC.eml` — EML · 2.16 KiB · current-support · 290 body words
- `service callback action window - 7.6 DC.eml` — EML · 2.42 KiB · current-support · 336 body words
- `service callback evidence follow-up - 7.3 SA.eml` — EML · 2.94 KiB · current-support · 396 body words
- `service commitment timing decision - 7.8 DC.eml` — EML · 4.40 KiB · current-support · 693 body words
- `service operating-review allocation economics - 7.7 DC.eml` — EML · 6.02 KiB · current-support · 914 body words
- `service pricing + callback review notes - SA 7.1.docx` — DOCX · 39.99 KiB · current-support · 3 rendered pages · 1 tables · 17 table rows

## Measurement notes

- PDF pages are read from the PDF page tree.
- Word pages are measured after headless LibreOffice rendering and may differ slightly from Microsoft Word pagination.
- Excel populated rows include model rows, summaries, assumptions, checks, and source data—not only independent records.
- Accounting counts are physical table rows. A single economic event can appear in a header table, one or more detail tables, and the general ledger.
- File sizes reflect compressed source files and do not measure analytical difficulty.

Generated by `environment/tools/generate_seed_inventory.py`.
