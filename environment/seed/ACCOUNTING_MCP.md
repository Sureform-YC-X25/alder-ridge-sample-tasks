# Contractor Accounting MCP

The `contractor_accounting` MCP provides the accounting and operating records for Alder Ridge Mechanical. The records cover the period from the company's formation on April 16, 2018 through the June 30, 2026 reporting cutoff. May 31, 2026 is the most recent closed month; June remains the current close period.

Alder Ridge reports on an accrual basis and recognizes revenue on long-term contracts using cost-to-cost percentage of completion. Records carry the dimensions normally used in contractor accounting, including account, department, project, cost code, customer, vendor, employee and service work order.

## Company structure

The database includes six operating and corporate departments:

| Code | Department | Branch |
| --- | --- | --- |
| 10 | Preconstruction | Portland |
| 20 | Construction - Portland | Portland |
| 21 | Construction - Central Oregon | Bend |
| 30 | Service | Portland |
| 40 | Building Controls | Portland |
| 90 | General & Administrative | Corporate |

Master data is available for customers, vendors, employees, projects, departments, cost codes, service agreements, fixed assets and prepaid items. The chart of accounts identifies account type, normal balance and cash accounts.

## Financial reporting and ledger

| Tool | Returns |
| --- | --- |
| `get_trial_balance` | Posted account activity for an inclusive date range |
| `get_profit_and_loss` | Posted P&L by month, department or account |
| `get_balance_sheet` | Posted account balances as of a selected date |
| `get_general_ledger` | Posted journal-line detail, including stable journal and line IDs, filterable by account or project |
| `list_accounts` | Chart of accounts, with account type and normal balance |
| `get_audit_log` | MCP write history by document or transaction type |

The general-ledger endpoint is paginated at 100 lines per request. Summary reports can also be exported to CSV through `export_accounting_report`.

## Project accounting

| Tool | Returns |
| --- | --- |
| `get_project_summary` | Contract value, cost budget, cost to date and billings by job |
| `get_job_cost_detail` | Project cost transactions with stable entry IDs through a selected as-of date, optionally bounded by a start date |
| `get_project_change_orders` | Change-order status and related revenue and cost-budget amounts |
| `get_open_purchase_orders` | Remaining commitments on open and partially received POs |

Project and job-cost records retain the project and cost-code detail needed for WIP, margin, change-order and forecast analysis.

## Receivables, payables and cash

| Tool | Returns |
| --- | --- |
| `get_ar_aging` | Open customer invoices, retainage and aging buckets |
| `get_ap_aging` | Open vendor bills, due dates and payment holds |
| `get_cash_balances` | Book cash by bank account as of a selected date |
| `get_bank_activity` | Bank-feed transactions with book links and cleared status |

The underlying invoice, bill, purchase-order, journal and bank records can be searched with `query_transactions` and retrieved individually with `get_transaction`. Invoice, bill, purchase-order, journal and payroll-run lookups include their line detail.

## Service, labor and payroll

| Tool | Returns |
| --- | --- |
| `get_service_work_orders` | Service dispatch records by customer, technician and status |
| `get_service_agreements` | Customer/site agreements, billing frequency and annual value |
| `get_timecard_detail` | Employee time by date, project or service work order |
| `get_payroll_register` | Payroll detail by run, employee and department |

## Other accounting schedules

| Tool | Returns |
| --- | --- |
| `get_fixed_asset_register` | In-service assets, cost, depreciation, net book value, location and vendor |
| `get_prepaid_schedule` | Contract term, original cost, amortization and remaining prepaid balance |

## Finding records

`get_company_profile` returns the company profile, reporting dates and department list. `list_available_reports` lists the report names accepted by the export tool. `list_master_data` searches master records, while `query_transactions` searches invoices, bills, purchase orders, journals, work orders, timecards, payroll runs, bank transactions and change orders. `list_workspace_attachments` lists the company files available alongside the accounting records.

Most detail tools accept a date range or as-of date plus optional project, account, employee, department, customer, technician or status filters. Paginated endpoints return a `limit` and `offset` so larger populations can be reviewed in sections.

## Accounting entries and workflow actions

The MCP also supports the transaction work required by tasks that call for entries or document updates:

| Tool | Action |
| --- | --- |
| `create_journal` | Creates a dimensional draft journal and can post it when balanced |
| `post_or_void_journal` | Posts a balanced draft or voids an existing journal |
| `create_invoice` | Creates and posts a customer invoice, including optional line detail |
| `create_bill` | Creates and posts a vendor bill, including project and cost-code detail |
| `create_purchase_order` | Creates an open PO with commitment lines |
| `update_transaction` | Updates supported workflow fields on invoices, bills, POs or draft journals |
| `apply_payment` | Applies a customer receipt or vendor payment and records the cash journal |
| `update_document_status` | Updates invoice, bill or PO workflow status and supporting notes |

Posted accounting fields cannot be rewritten through `update_transaction`. Every MCP write is recorded in the audit log. Individual task instructions determine whether the expected work is analysis only, preparation of a draft entry, or an authorized accounting action.

## Report exports

`export_accounting_report` writes a CSV report to the task workspace. Available exports are:

- profit and loss
- balance sheet
- trial balance
- general ledger
- project summary and job-cost detail
- AR and AP aging
- open purchase orders
- cash balances and bank activity
- service work orders and service agreements
- timecard and payroll detail
- fixed assets and prepaid items
- project change orders

Export parameters follow the corresponding MCP report: date range or as-of date, reporting grain and any applicable account, project, customer, technician, employee, department, status or cleared-state filter.
