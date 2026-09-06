# Alder Ridge Seed Data Inventory

This inventory is generated from the packaged five-task manifest, source registry, and read-only accounting snapshot. Run `python scripts/build_seed_inventory.py --check` to verify it.

## Summary

| Measure | Value |
|---|---:|
| Selected tasks | 5 |
| Company-world source files | 144 |
| Company-world source size | 4.00 MiB |
| Accounting database size | 74.35 MiB |
| Accounting business tables | 27 |
| Accounting rows | 575,008 |
| Accounting SHA256 | `00ddfe914af7abb109a05f703623f43c241ac5f90d73e0f4f5e46d9459b76e7d` |
| Source world | `alder-ridge-mechanical-v2` |
| Snapshot | `2026-06-30-pre-close` |

### File types

| Type | Files |
|---|---:|
| `csv` | 16 |
| `docx` | 16 |
| `eml` | 27 |
| `pdf` | 21 |
| `pptx` | 9 |
| `txt` | 1 |
| `xlsx` | 54 |

### Source authority/version status

| Status | Files |
|---|---:|
| `approved-plan` | 2 |
| `approved-policy` | 1 |
| `approved-record` | 1 |
| `current-support` | 81 |
| `current-working` | 49 |
| `historical-final` | 6 |
| `stale-working` | 3 |
| `superseded` | 1 |

## Selected task source contracts

| Task | Minimum source artifacts | Accounting MCP |
|---|---:|---|
| `task_001` | 10 | required |
| `task_004` | 13 | required |
| `task_015` | 19 | required |
| `task_035` | 10 | not required |
| `task_068` | 20 | required |

## Accounting database

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

| Path | Type | Size | Status | SHA256 |
|---|---|---:|---|---|
| `Deliverables/ARM-2417 June close and PCO-017 review.docx` | `docx` | 39.71 KiB | `current-support` | `fa1ca83927b09eea221ba8ed29c0ff60b9e0289284295ca85a5093b6bdd27259` |
| `Requests/7.1.26_822am - Fwd_ wip close need today.eml` | `eml` | 1.04 KiB | `current-support` | `e5a239130fffa8397c8dfefa35fd33042e7752d476c87442f9248973258e73cf` |
| `Requests/7.1.26_904am - cash refresh + Fri deck.eml` | `eml` | 895 B | `current-support` | `82e17582f1ff16e67a7e3110e02b1e51ee39d44110af168716ef188da9a0eeb4` |
| `Requests/7.2.26_1027am - Q2 bank compliance package.eml` | `eml` | 1.17 KiB | `current-support` | `39fadb63e526e013663a8b9f50765ecdd96044c03b1a8c4f35f6548693cff30c` |
| `Requests/7.2.26_1103am - CFO close decision note - internal only.eml` | `eml` | 789 B | `current-support` | `125ccd309f5832012818529df727e99962dd8a88c8dc45f6f73cd5ab78f0a58e` |
| `Requests/7.2.26_611am - June ops review deck + notes.eml` | `eml` | 697 B | `current-support` | `577de977992c61b903403626f6569470d1443c6d16b14870aaf92c3683cd5d6b` |
| `Requests/7.2.26_736am - FW_ wc cleanup + backlog coverage.eml` | `eml` | 963 B | `current-support` | `22b257c942553bb0863c67090b93ec2a13a20597ab1493996017044b4b01ef5d` |
| `Requests/7.2.26_805am - payroll bridge + capex before noon.eml` | `eml` | 1.06 KiB | `current-support` | `c09d1da6d54dc3eaf87260560b78210984fa940771474c598361cd4d51d3558c` |
| `Requests/7.2.26_842am - service vans capacity check - LT fwd.eml` | `eml` | 952 B | `current-support` | `09ef2b12e8b1c3398ae8bf8397a5112c1bfd6971cc7211434b9cb7b06f2a5186` |
| `Requests/7.4.26_0711am - FY27 plan first pass + branch submissions.eml` | `eml` | 1.42 KiB | `current-support` | `3245fdbf484a84710e90b0d311a35484d8b55649e9643a392a3ed968804e94b2` |
| `Requests/7.4.26_0838am - capital committee + five year plan.eml` | `eml` | 1.13 KiB | `current-support` | `7e0e78a286b91c6e775c9f10ffcfc756a38e26510599f02f42d574699c18a896` |
| `Requests/7.4.26_0956am - July liquidity bank work + covenant outlook.eml` | `eml` | 1.29 KiB | `current-support` | `cd4b283387c318fe0b7e0897737a43c4c85d3f6e09b12709b900fa869d0067c0` |
| `Requests/7.4.26_1106am - corp dev IC workstream status.eml` | `eml` | 1.26 KiB | `current-support` | `4cf64bc351cd7a6a68e8c23d3e43daf3330cf9d8bdda16871c598c930c1937f3` |
| `Requests/7.4.26_1216pm - tax provision + EBITDA definitions.eml` | `eml` | 1.16 KiB | `current-support` | `5f0cf4373eeda6051bf1d9da2c91974ad1b35c0f14f0f0d432be4dbdc3578644` |
| `Requests/7.4.26_1344pm - Q2 board lender + IR refresh.eml` | `eml` | 1.15 KiB | `current-support` | `1d5917863fbb30ddd4b70e33bcc8f829a7194968529ca172a67c455582b4c498` |
| `Requests/7.5.26_0618am - FY27 plan review follow-up.eml` | `eml` | 1.31 KiB | `current-support` | `27ee557525c77a24efc5780480317c8fee21c4b13f82f2cc7c83b28d802bbe0b` |
| `Requests/7.5.26_0642am - capital cases after committee review.eml` | `eml` | 900 B | `current-support` | `2ef2ed5223517786bbcc82886000c1b08eaddcf5e739c4e2c7a686ba6035454d` |
| `Requests/7.5.26_0709am - bank follow-up + treasury version check.eml` | `eml` | 1020 B | `current-support` | `7ee07e198ebb2dbe243b5bd0dbf4bef27e766375a103fceec882985bb0637236` |
| `Requests/7.5.26_0727am - deal IC follow-up + banker case warning.eml` | `eml` | 1.04 KiB | `current-support` | `91391f800c937b2cc6fe69228053a83c48abd6fe441d3e627601c7b7d39b586c` |
| `Requests/7.5.26_0751am - tax review comments + close version.eml` | `eml` | 930 B | `current-support` | `316288e2c4821b48af43d4221892ebf3dca689b6d8995a3cc49d1ed41defc0e3` |
| `Requests/7.5.26_0812am - board review changes + source tie.eml` | `eml` | 963 B | `current-support` | `9ebdf491b7d6747051b34ffd4114b995f3a0c2fba5c1b4a4932d0f2850d4d8da` |
| `Shared/Finance/AP/Invoice Batches/2026/07/01/batch 20260701-02.pdf` | `pdf` | 81.70 KiB | `current-support` | `18bf0f58a5a02b35d38ff210784274ec23d3d44ff085d6c6dced346ea38a2f4b` |
| `Shared/Finance/Accounting Policies + old memos/Revenue recognition - WIP policy_rev11-24 SIGNED scan.pdf` | `pdf` | 71.52 KiB | `current-support` | `714c589af37ce0f56821402e0842fefd342bceffda34022589dd8b8f31d272c7` |
| `Shared/Finance/Close/2026/05 May/4_WIP/WIP 5.31.26_FINAL_v7_revised NB.xlsx` | `xlsx` | 372.68 KiB | `historical-final` | `e4687e5de974150d1ee9331746d454eae751e30b8bc8d9cf43ae09e740bb3426` |
| `Shared/Finance/Close/2026/06 June/1 close mgmt/June close tracker - working - upd 7.1 810am.xlsx` | `xlsx` | 14.34 KiB | `current-working` | `22e77cb10e58cabf009475cef8ee18dc3ad6d653b04a64eeec5b39a838352dcd` |
| `Shared/Finance/Close/2026/06 June/1 close mgmt/READ ME - where we are 7.1 810am.txt` | `txt` | 518 B | `current-support` | `dfa6964c5cfc8323c7090244e1c899aa0865b152b997de19bc38de31d7a77f44` |
| `Shared/Finance/Close/2026/06 June/2 reconciliations/AR roll + retainage 0630 KS.xlsx` | `xlsx` | 113.20 KiB | `current-support` | `e4be682828eea46c837f88eb6174a612b2cdcccbbe2bd1a3787230363a064dcc` |
| `Shared/Finance/Close/2026/06 June/2 reconciliations/TB rec_6.30 prelim - NB v4.xlsx` | `xlsx` | 16.81 KiB | `current-working` | `bf74cf03714d36bb0a3e1b17ce5ff0b9212494ca5b143e01b146443b4d07958d` |
| `Shared/Finance/Close/2026/06 June/2 reconciliations/Vista billing batch history through 7.5.csv` | `csv` | 6.53 KiB | `current-support` | `91e80f1d46a28223903170023119e0e1cadd78644aa45f44fae1e5523af57668` |
| `Shared/Finance/Close/2026/06 June/2 reconciliations/bank recs 6.30 - MP working.xlsx` | `xlsx` | 41.44 KiB | `current-working` | `6155d9165bda54f740dd2002ee24ce2aeabcd5fa1377d4098a6bcc35ea42fd18` |
| `Shared/Finance/Close/2026/06 June/3 accruals/AP cutoff + accrual list_7.1 OL.xlsx` | `xlsx` | 87.43 KiB | `current-support` | `15cebc513bb80e0bf7de09078bd761f2a78d6e8a7271b7e4ef9a6401c2a0bbb5` |
| `Shared/Finance/Close/2026/06 June/3 accruals/payroll accrual 6.30 - FINAL? v3.xlsx` | `xlsx` | 16.65 KiB | `current-working` | `958b1a34b3928ec4a84007083c27d38f4dc7cf5571cba15f65b248bf3b4f4bd9` |
| `Shared/Finance/Close/2026/06 June/3 accruals/prepaids_amort sched FY26 - copy.xlsx` | `xlsx` | 14.14 KiB | `current-working` | `aff9acf6b5fe75405f797b203ac0f53f175c16ae55479a3966560be56b9f24e8` |
| `Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx` | `docx` | 42.59 KiB | `current-working` | `ca7a510681f801b9a9697b1fb503f6a05c0d52b5d88aa9c5052a6241ad0110af` |
| `Shared/Finance/Close/2026/06 June/4 WIP/WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx` | `xlsx` | 6.66 KiB | `current-working` | `6e82251b717b922d6067e467a99419ed134d4385fd11b5878608c53813ef9120` |
| `Shared/Finance/Close/2026/06 June/5 fixed assets/FA + CIP rollforward_June26 rev2.xlsx` | `xlsx` | 15.23 KiB | `current-support` | `bd7c1dc8cbaea27cc29156a7e2d43c870e972d24c1e540285c8274c21c0b41f9` |
| `Shared/Finance/Controllership/Q2 working/FY26 tax + non-GAAP reporting policy - signed.pdf` | `pdf` | 10.32 KiB | `current-working` | `2343e866a3f8ecb3b5f91b8e569d6087e056098a7e7b7ceddabc292eb17a5d15` |
| `Shared/Finance/Corporate Development/Active analyses/IC diligence review notes - 7.5 DC.docx` | `docx` | 43.88 KiB | `current-support` | `5eaf1c2f035fe3c23d920532421b0b67c20cb0609b4b8e469592df16012eb913` |
| `Shared/Finance/Corporate Development/Active analyses/IC model inputs - 6.29 banker case.xlsx` | `xlsx` | 27.42 KiB | `current-support` | `0fc2a22868d072927e3ffc786331ca98a128a9c1ad2f18f834ec0e99f7ec497b` |
| `Shared/Finance/Corporate Development/Active analyses/IC model inputs - 7.5 controller tie.xlsx` | `xlsx` | 43.15 KiB | `current-support` | `5949c806d30a4f8e6b9d01b8509f5b46caecd2929ccaec64ec55bd6ac4c60d74` |
| `Shared/Finance/Corporate Development/Active analyses/M&A underwriting policy + IC return definitions.pdf` | `pdf` | 13.48 KiB | `current-support` | `cba81a0ceb2851dc6509eee46a7c4bab3f8ef95ce623ea014d08f0cc5d2ccc2e` |
| `Shared/Finance/Corporate Development/Data room exports/target monthly financials + diligence index_7.3.csv` | `csv` | 30.19 KiB | `current-support` | `86df27d68018d562fd563267d3bfc03d6802d8793df025af180180d041b98ed2` |
| `Shared/Finance/Corporate Development/Orion integration value capture - WORKING.xlsx` | `xlsx` | 13.53 KiB | `current-working` | `a38c949d85be8626e3597fcfc7edfa104d0839aa8311b6591c21e771a3e0a2a2` |
| `Shared/Finance/FP&A/Backlog/FY27 backlog burn and capacity - WORKING.xlsx` | `xlsx` | 24.65 KiB | `current-working` | `d2aa283d16f9d875371388994f317b49420fb39bfbcb15b342cd3b31dfb8d37f` |
| `Shared/Finance/FP&A/Backlog/backlog burn + award pipeline_6.30 v6.xlsx` | `xlsx` | 13.15 KiB | `current-support` | `7420b603202cf0cb5c1e870500817e40c6c0ac15184160a3dc01545dc2a5af38` |
| `Shared/Finance/FP&A/Capex/2026 equipment requests/CapEx asks + ROI_2026 midyear - MW.xlsx` | `xlsx` | 10.15 KiB | `current-support` | `43a4d7600057ee1f4a8cb53ce1cdb43fc5ae2dee5d45befbc991cc96ec0f3bf0` |
| `Shared/Finance/FP&A/Capex/2026 equipment requests/midyear capex committee memo - draft 7.1.docx` | `docx` | 39.20 KiB | `current-working` | `e3a8e430a11d0a1fe726620881dceac6098f523aeec82d9a6b7df5f80bfd0f01` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/FY26 AOP approval deck_12.18.25 FINAL.pptx` | `pptx` | 38.07 KiB | `approved-plan` | `f35e67b15359bb8a8d506e10a58e3fa1c28b5b7a5ada39ee5cefe04b6c55428c` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/FY26 Op Plan_BoardApproved_12.18.25_FINAL2.xlsx` | `xlsx` | 19.54 KiB | `approved-plan` | `a343602eec8aacf00861914257bd0cb5854e0e671adca830e295cf8592e5d713` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/FY26 guidance cases_7.5.csv` | `csv` | 1.31 KiB | `current-support` | `92fa7c0cabd403420c606304d9aa32b9590b9befa67040b9f491e5821c48bf73` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/FY26 rolling forecast_v12 - pre WIP.xlsx` | `xlsx` | 13.01 KiB | `stale-working` | `af52aadb93a221ef583d9aa00caf938fe956184aa324e2c7248c3bc591f85039` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/June case sensitivity inputs_7.5.csv` | `csv` | 1.69 KiB | `current-support` | `7974c4519ee33d2b1ccd59b65230302cbd1e8ae587b7a8614447a27e8000a40f` |
| `Shared/Finance/FP&A/FY26 plan + reforecast/department reforecast inputs - June pull.xlsx` | `xlsx` | 12.96 KiB | `current-support` | `637dddc8e52bbc182d4da7be3d40f816c2ce6cc527923260742a78dc7bd23847` |
| `Shared/Finance/FP&A/FY27 plan/FY27 EBITDA scenarios - WORKING.xlsx` | `xlsx` | 27.62 KiB | `current-working` | `8aa8641ccf8268e2fd097fa524051cc9a8c8f8ad8edd20315cdd667e02ba0aec` |
| `Shared/Finance/FP&A/FY27 plan/FY27 field workforce plan - WORKING.xlsx` | `xlsx` | 12.81 KiB | `current-working` | `9ad4aa9d1169868b3ea235349241759747395e225e0c6aa166e766af8399ff53` |
| `Shared/Finance/FP&A/FY27 plan/FY27 planning assumptions - v4 branch draft.xlsx` | `xlsx` | 35.10 KiB | `current-working` | `c0091953a125394a8c647b292ab5a8c550401b5bc98950a99c41578ac1ce8839` |
| `Shared/Finance/FP&A/FY27 plan/FY27 planning assumptions - v6 controller tie.xlsx` | `xlsx` | 89.46 KiB | `current-support` | `f365b0a68552c88cb6fe2b45383ce90025894cd8c959ed03fa9b13263c08ca7e` |
| `Shared/Finance/FP&A/FY27 plan/FY27 planning definitions + scenario guardrails - APPROVED.pdf` | `pdf` | 19.86 KiB | `current-support` | `69a1fb011d8d4e2c0a59d949f2f110bd8b7be88d218dea12b2a4924de843b144` |
| `Shared/Finance/FP&A/FY27 plan/FY27 steering committee notes - 7.3 DC.docx` | `docx` | 46.38 KiB | `current-support` | `d28b9f8d355c9dbcb332ff727b746ea06813b9912fa8138d7b69e40e6b3e6933` |
| `Shared/Finance/FP&A/Labor/2026 field labor loaded rate build - est copy.xlsx` | `xlsx` | 5.99 KiB | `current-working` | `dcafebced0d934738606a38dba6695d410af0ff71beb5dcf5058a5fcc64e3e2c` |
| `Shared/Finance/FP&A/Labor/HC + field labor productivity Q2 working.xlsx` | `xlsx` | 66.45 KiB | `current-working` | `40f788059f543e57143d6e03d2811b24e6f3b187fe30850e9bd2960f16cb783e` |
| `Shared/Finance/FP&A/Labor/field burden rate approval - 2026 estimator use.pdf` | `pdf` | 69.22 KiB | `current-support` | `52e7f76461754c7e64be6041a94cbd77e0991a02a86f01ca43a25c4f1ce816d3` |
| `Shared/Finance/Investor Relations/Peer analysis/Q2 peer market + competitor support - 6.27 banker pull.xlsx` | `xlsx` | 11.60 KiB | `current-support` | `805c3698b6d19541ef3ff9cfc48a73dc1335d0d5e7b60bdb9b38923d96b8b5c5` |
| `Shared/Finance/Investor Relations/Peer analysis/Q2 peer market + competitor support - 7.3 close.xlsx` | `xlsx` | 11.58 KiB | `current-support` | `d45a10d16317ba42de48e83f284d18570fe1bc37d8f7020abb234a4d562cb849` |
| `Shared/Finance/Investor Relations/Peer analysis/market study + valuation methodology - approved.pdf` | `pdf` | 4.88 KiB | `current-support` | `47ad024313a8bc8e435fdf3b526bda066dcbd30afe775e8232f511919c92b677` |
| `Shared/Finance/Investor Relations/Peer analysis/peer market data pull_7.3.csv` | `csv` | 14.99 KiB | `current-support` | `abeb50bfc1875a0305fce3e579027a3400cc7363c8673073144697bebbebd5a5` |
| `Shared/Finance/Investor Relations/Peer analysis/peer review notes + comparability decisions - 7.3 DC.docx` | `docx` | 39.28 KiB | `current-support` | `b126cccc8a1443b971ac02b4b9404db91f7656cebc979a92412eb750395d0620` |
| `Shared/Finance/Investor Relations/Q2 working/FY26 published range record_6.18.csv` | `csv` | 632 B | `approved-record` | `bd8188890a513e7a0becf4851f15e332b051f50bd24c9b06072f56e71779585c` |
| `Shared/Finance/Investor Relations/Q2 working/KPI definitions + lender presentation policy.pdf` | `pdf` | 5.43 KiB | `current-working` | `e5293a2b6d7b71fa80d605fa455e9040232714e2c65083e20adc68a54458a2ea` |
| `Shared/Finance/Investor Relations/Q2 working/guidance interval settings_6.18.csv` | `csv` | 507 B | `approved-policy` | `2a649cacd6936252d95dee26a303144970e1d02b3067107ff5eafe6e611e6dd4` |
| `Shared/Finance/Reporting/2026/05 May/May Ops Review - 6.5 mtg - FINAL_v4.pptx` | `pptx` | 54.07 KiB | `historical-final` | `c49153203328e857d098661bf057a859a17761f932877aa7694e70f430e0605e` |
| `Shared/Finance/Reporting/2026/05 May/notes from 6.5 ops mtg - dc.docx` | `docx` | 38.67 KiB | `historical-final` | `72b12bb540ea00ffc06eb1b58663591e3ab29c28242dd176aaa11fe3f69840a6` |
| `Shared/Finance/Reporting/2026/06 June/FY26 outlook risk register_7.5.csv` | `csv` | 1.50 KiB | `current-support` | `429d9332b6fe881bd9f60d0e29255e8625778563968d3536da60faed8884c628` |
| `Shared/Finance/Reporting/2026/06 June/Finance action review notes_7.5.csv` | `csv` | 6.09 KiB | `current-support` | `385ee6303c4d42670c439a2af7f398cdb1557356e9b748246be288b836087955` |
| `Shared/Finance/Reporting/2026/06 June/June executive performance review - WORKING.pptx` | `pptx` | 49.58 KiB | `current-working` | `377e6c3f7cfb8b28625ac4c674ab01139d45de6082e6c99bb7025add47319d33` |
| `Shared/Finance/Reporting/2026/06 June/June flash - CFO scratch v2 7.2.pptx` | `pptx` | 26.63 KiB | `current-support` | `7f72f3fe6a88bc93ec3477b487d1a9605f93120bae9f3dc9b46aa1fb585e4a3c` |
| `Shared/Finance/Reporting/2026/06 June/June flash bridge - review copy 7.2.xlsx` | `xlsx` | 7.97 KiB | `current-working` | `2bd1efaa5af554077b1fa946261ec4d52b599f6c13ed0f0ee8800130bb7dc999` |
| `Shared/Finance/Reporting/2026/06 June/June flash review notes - DC 7.2 643am.docx` | `docx` | 40.00 KiB | `current-support` | `c5300ad9d1b927989545441cfdeaf829462c4d93b8db25d6844a8a3dac4b5dd0` |
| `Shared/Finance/Reporting/2026/06 June/Q2 board + lender narrative review notes - 7.4 DC.docx` | `docx` | 39.17 KiB | `current-support` | `e3e85a9634ad4b50af59bad0d5b911d0d50432b43f4cfc4defca455e288ab40c` |
| `Shared/Finance/Reporting/2026/06 June/Q2 management reporting cube extract_7.2.csv` | `csv` | 7.96 KiB | `current-support` | `53a79ff1afa7cbeb5b3f047ac14f4a0397339f4eaeaa90a524bca2b1e01fbdbc` |
| `Shared/Finance/Reporting/2026/06 June/Q2 management reporting data book - v5 CFO scratch.xlsx` | `xlsx` | 20.45 KiB | `current-support` | `74910243fb70ae7b9e33f60b9263c484f93ed24f0bc856ce0228e5848b05c5e2` |
| `Shared/Finance/Reporting/2026/06 June/Q2 management reporting data book - v7 controller tie.xlsx` | `xlsx` | 23.37 KiB | `current-support` | `d41bc01197966303831d3490bfdc2ae78d19a5d1d32063d7aaac06522766daa6` |
| `Shared/Finance/Reporting/Board/Q1 2026/Q1 board package - FINAL signed off.pptx` | `pptx` | 29.44 KiB | `historical-final` | `d5833bba0c2daead046a3de58f3fba38394f104836247e3778c5707ab85bc222` |
| `Shared/Finance/Reporting/Board/Q2 2026/Board performance dashboard - WORKING.xlsx` | `xlsx` | 13.78 KiB | `current-working` | `7860c50d3d475b2bf29087bbef4a4edc15407fecb66834b3b7af1c4ac2341b3f` |
| `Shared/Finance/Reporting/Board/Q2 2026/CFO narrative Q2 board pre-read - v3 comments.docx` | `docx` | 39.99 KiB | `current-support` | `45c0f82f70f3702484983796b6d07196afc6430c6db304220eb0894b05f300ed` |
| `Shared/Finance/Risk + Insurance/2026 renewal exposure notes - broker call followup.docx` | `docx` | 39.12 KiB | `current-support` | `ecb189ea84374ac06abdbc1901a886333aeed44111ff22d67ea21be0c42bdeb7` |
| `Shared/Finance/Risk + Insurance/2026-27 binder + schedule of coverage - broker draft.pdf` | `pdf` | 69.13 KiB | `current-working` | `fcea8ff5f050212498993b95ef8684595483da2f284666b79f8218b2caca5685` |
| `Shared/Finance/Risk + Insurance/bonding insurance schedule 2026 - renewal working.xlsx` | `xlsx` | 11.60 KiB | `current-working` | `26472578f3be58e43e3993014c549865e70a2b2a9c6ce6196f8284efd6784356` |
| `Shared/Finance/Strategic Finance/FY27 working/capital committee cases + LRP assumptions - v3 pre-review.xlsx` | `xlsx` | 17.25 KiB | `current-working` | `2fa9c56fe4c932bc23f7adbeb878670cbd8147eb2272ec3d170e68858ba28ad3` |
| `Shared/Finance/Strategic Finance/FY27 working/capital committee cases + LRP assumptions - v5.xlsx` | `xlsx` | 39.65 KiB | `current-working` | `4126a92b3e467bf59dcf06469d1888566c5dbf10b887a9594ac353e2a54229b4` |
| `Shared/Finance/Strategic Finance/FY27 working/capital committee review notes - 7.4 DC.docx` | `docx` | 40.25 KiB | `current-working` | `325ae309b3086ccc2cc68839ebcc1e945620db9b7a9c5b3856d3dd5b69fc5d46` |
| `Shared/Finance/Strategic Finance/FY27 working/investment hurdle + capital guardrails - board approved.pdf` | `pdf` | 6.87 KiB | `current-working` | `637ef3268b51a5e69b03380cf7ff7c69e05395c87dcaabf07958d11b6fc57607` |
| `Shared/Finance/Tax/2026 working/FY26 tax + non-GAAP workpapers - v2 pre-close.xlsx` | `xlsx` | 20.24 KiB | `current-working` | `d78f83006b6bbe164e73c543df2a3af3678ef87ca0206c6b3c75c07036c0a979` |
| `Shared/Finance/Tax/2026 working/FY26 tax + non-GAAP workpapers - v4 controller review.xlsx` | `xlsx` | 29.04 KiB | `current-working` | `4eeef4b23c4cd7307d4fc85506c788be06e8b84b9037be616825d2d7ee0b0afc` |
| `Shared/Finance/Tax/2026 working/tax jurisdiction + book trial balance detail_6.30.csv` | `csv` | 17.43 KiB | `current-working` | `6aa8c43da94c4a686e2ae3e852100c067729e905a1bb4433de1d67518f80611a` |
| `Shared/Finance/Tax/2026 working/tax provision review notes - 7.4 NB.docx` | `docx` | 42.15 KiB | `current-working` | `2ca3609ea5f73994517bb6516113d0af92f61366a10f995f6111f695979bb168` |
| `Shared/Finance/Tax/2026/1099 sales use tax tracker - 6.30.xlsx` | `xlsx` | 18.83 KiB | `current-support` | `9f4ebf1899dfec525779240f8d28b328fc8bed22771f246cf00e7c826db123fe` |
| `Shared/Finance/Tax/2026/FY26 tax provision - WORKING.xlsx` | `xlsx` | 12.91 KiB | `current-working` | `cf5ab06d8849f8cbdff75c579e9c566462ab12f375553f609afb29a7771ffc75` |
| `Shared/Finance/Tax/2026/OR DOR desk review notice + response checklist.pdf` | `pdf` | 68.98 KiB | `current-support` | `71fa586b609379bbe7e98b77d7b34250dd7753bc7a50de115246bdeaa861d157` |
| `Shared/Finance/Treasury/13 week cash/13wk cash_v9 - 6.26 roll - DCHO edits.xlsx` | `xlsx` | 130.81 KiB | `stale-working` | `89fb71f97a37fba166f26bdc739fde795df72c91a233124377bcc8d6719742b6` |
| `Shared/Finance/Treasury/13 week cash/downside assumptions_7.2 618am - DC marks.xlsx` | `xlsx` | 10.08 KiB | `current-support` | `a4511056691b83915b0e386e70081f946052c376a89189fbe7e59c1e376c3fbe` |
| `Shared/Finance/Treasury/13 week cash/downside liquidity sensitivity_7.2 618am - current base treatment.xlsx` | `xlsx` | 8.52 KiB | `current-support` | `e63dc689b9477549b3ec8aaf33ef2afa2099cdd94cccbe40f76228a4fcd2c830` |
| `Shared/Finance/Treasury/AR calls - notes/AR notes_6.30 - NB working copy.docx` | `docx` | 44.51 KiB | `current-working` | `060f3847ebaed62a089370338d6bdc33f9c3a7512f5a0fbf3079556bb6a0d816` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q1 submitted/Q1 2026 compliance pkg - submitted 4.28.26.pdf` | `pdf` | 78.46 KiB | `historical-final` | `069de958e11d7b4425b96243221e53079831ac86670f4b338b3eb97f64a2d1d2` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q1 submitted/covenant summary - old 2024 (use agreement).pdf` | `pdf` | 63.42 KiB | `superseded` | `8692981ffbe3bb97e9e83e47680c2557238b76f3f67367d7ca66d354b37855d5` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q1 submitted/lender + surety update Q1 - submitted copy.pptx` | `pptx` | 26.52 KiB | `historical-final` | `9dcc1f24ab0e153b21d2613f44db2a0eefa67b1ae1fd91f4923b1ada7e2eda0d` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 covenant headroom - lender review working.xlsx` | `xlsx` | 7.05 KiB | `current-working` | `54dd0ed1913e3009687f3c22abef130caa78ee6f20673f15e5bb594f55abbcb8` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 lender update - review working v3.pptx` | `pptx` | 27.01 KiB | `current-working` | `77e6107f163af332d357bb0d676be4dd03d47b723e8faa932659f1bb45992c0d` |
| `Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/public-owner assignment status - lender reply 7.2.eml` | `eml` | 1.91 KiB | `current-working` | `b511ebbd38d1acb9c8880a96cff63474cc5b2fd85933214858a7d879c234568c` |
| `Shared/Finance/Treasury/Bank - covenants/Agreement + amendments/USBank AR eligibility exhibit - closing set copy.pdf` | `pdf` | 69.50 KiB | `current-working` | `a5864ccdaa8a9f4569bf272d4b6c503808d14f8e9ab65a393859107aa5b243a5` |
| `Shared/Finance/Treasury/Bank - covenants/Agreement + amendments/USBank_Amdt2_9.30.25_EXECUTED scan.pdf` | `pdf` | 75.08 KiB | `current-support` | `f421b58250ffda87755630d94e80aacdd8f66e0745d163655f469b9fe30092dc` |
| `Shared/Finance/Treasury/Bank - covenants/FY27 covenant forecast - WORKING.xlsx` | `xlsx` | 14.00 KiB | `current-working` | `1a56fc2031b46e30d87dcf43934ae7792621eb918d4003e314b874130163e9b3` |
| `Shared/Finance/Treasury/Cash positioning/July daily cash position - WORKING.xlsx` | `xlsx` | 14.29 KiB | `current-working` | `5fdab1830a88eb1e4c5001f68f52608097daffc20753a66110e6b6a5c734c930` |
| `Shared/Finance/Treasury/Cash positioning/bank portal settlements + expected items_7.3.csv` | `csv` | 37.83 KiB | `current-support` | `783eb84798a727f8b173cf14584d542ca34c3eccab142c7ecd71822f78306c7d` |
| `Shared/Finance/Treasury/Debt/debt sched - 6.30 before bank pkg.xlsx` | `xlsx` | 9.74 KiB | `current-support` | `f3ee4a488e34799d7980b5606c07aef84df79156046506853ef8aab984179be0` |
| `Shared/Finance/Treasury/Debt/equipment line proposal - bank copy 6.18.26.pdf` | `pdf` | 69.00 KiB | `current-working` | `dc772ef5ae2053d92627031fd16dca81eedbfa1cef0b6c09ff5de2937f6ed7f5` |
| `Shared/Finance/Treasury/FY27 working/treasury and bank review notes - 7.4 DC.docx` | `docx` | 43.34 KiB | `current-working` | `c11c054fe0f019fbe55b592fb3470397fecaededfc4be7c245505dec4733b7d9` |
| `Shared/Finance/Treasury/FY27 working/treasury outlook support - July v6 pre-bank.xlsx` | `xlsx` | 26.93 KiB | `current-working` | `fcd3d7895194a4be276960df27cb5eb1641784dc95fd56bb032bc6f65f97893d` |
| `Shared/Finance/Treasury/FY27 working/treasury outlook support - July v8 controller tie.xlsx` | `xlsx` | 47.85 KiB | `current-working` | `25068b2464be5bbe2d690c1b4f04f98e68b5dc4e58068ef30a28a4f67346f2a5` |
| `Shared/Finance/Treasury/FY27 working/treasury risk limits + funding policy - board approved.pdf` | `pdf` | 12.77 KiB | `current-working` | `cb34608ea293f5daf2f7cd4526a961da6ec9bf74a3a7ac8bcf41d13e2acd1a56` |
| `Shared/Operations/Commercial/CO Log/6.30 support - not all final/2409_commercial working_6.24.pdf` | `pdf` | 66.53 KiB | `current-working` | `62f50db0199c82610dd15d8969e04d7e1a83bb0c69edcb9f6ff8eab9674e0016` |
| `Shared/Operations/Commercial/CO Log/6.30 support - not all final/2417_PCO17_owner emails + FD summary_6.30.pdf` | `pdf` | 66.40 KiB | `current-support` | `4e1fa1ce8d6d8e959179410a21806beb910c8fd5043b969517984a2f86b4fb56` |
| `Shared/Operations/Commercial/CO Log/6.30 support - not all final/2506_PCO6_DB27 backup + cost est (working).pdf` | `pdf` | 66.31 KiB | `current-working` | `e2cb22ed385b217b25a0c268172b11d4cae1cbae79385ffc76f0b0f9d43b63a0` |
| `Shared/Operations/Commercial/CO Log/CO log MASTER (do not sort)_6.30.26 - PR copy.xlsx` | `xlsx` | 13.49 KiB | `current-working` | `202486070d2b00632526e7bcfee5a55033982d487845a68a5092334341756b41` |
| `Shared/Operations/Commercial/Contract Records/2026/07/02/scan batch 20260702-01.pdf` | `pdf` | 71.48 KiB | `current-support` | `80f9f1a98c5c5b80ae6506d29ccf0df1d800945d79ee9b87caf2a28b2c60bc13` |
| `Shared/Operations/Commercial/Correspondence Archive/2026/06/30/sent 1706.eml` | `eml` | 690 B | `current-support` | `0848b5bb100d485d63eb9f9dc68eb6319e22ffc30ef09df4a813d4b6d3f74e81` |
| `Shared/Operations/Fleet/fleet + equipment utilization_June working.xlsx` | `xlsx` | 11.48 KiB | `current-working` | `c43381173a7c1814cc34648a25be94f9bfa6e913fc233b529db3c87ee28882ba` |
| `Shared/Operations/Planning exports/CRM backlog service renewals + forecast history_6.30.csv` | `csv` | 40.79 KiB | `current-support` | `31fd81d80484059dd8b6825860ca6d45d3bedc1fe793333b0d298bc3480d05bc` |
| `Shared/Operations/Planning exports/July owner support log_7.5.csv` | `csv` | 4.39 KiB | `current-support` | `7e50e98730b4b42457d47e304b01c234ff219b51612e0961b588fac724ee3bf7` |
| `Shared/Operations/Planning exports/Q2 operating evidence extract_7.5.csv` | `csv` | 7.79 KiB | `current-support` | `1399ed49be4a631fbd379d8ffad071281eac5549197337e7e567f43fcab93dae` |
| `Shared/Operations/Project Controls/cash curves/major jobs cash curves_6.30 scenario B.xlsx` | `xlsx` | 20.45 KiB | `current-support` | `d3cd0948eafc562b1ead187c86e9d8f1bbe27e2d44a31897efb77443d57ff525` |
| `Shared/Operations/Project Forecasts/June26 - PM updates/PM ETC updates_6.29 530pm_COMBINED_v3.xlsx` | `xlsx` | 18.54 KiB | `current-support` | `8ca9b478dd7613a0a9289b4b44bf04fe6de0542161c9bf9713dc3d6cf2866c9a` |
| `Shared/Operations/Project Reviews/weekly top jobs review - 6.29 PM copy.pptx` | `pptx` | 34.37 KiB | `current-working` | `4d567c06e2f8c926d4e29b9c60092aeeec37b92c2439f4050e653fd489c2b133` |
| `Shared/Operations/Quarterly Reviews/Q2 ops + safety review - 6.26 draft.pptx` | `pptx` | 34.18 KiB | `stale-working` | `bd278aced71658521eddc42590f93e709ddac67d215e5257a313b4c28c287a18` |
| `Shared/Operations/Quarterly Reviews/Q2 project execution status and rework review protocol - 6.30 approved.xlsx` | `xlsx` | 6.97 KiB | `current-support` | `ed98688e73f9a8960fbbb8c707955e06e3f4df499a19118d872369228b3f6513` |
| `Shared/Operations/Service/monthly KPI/Q2 service pricing + callback cut - 7.1 SA.xlsx` | `xlsx` | 66.03 KiB | `current-support` | `88cb2372dc70369429383110440b23df69084d2411092911e8f387b060c0b6ef` |
| `Shared/Operations/Service/monthly KPI/service KPI pack source_6.30 - LT edits.xlsx` | `xlsx` | 65.25 KiB | `current-working` | `3c4c11129e0c78cb1c10cee1e0985aa855fe920e6bcaeefc16f70a86433665cf` |
| `Shared/Operations/Service/monthly KPI/service access signal out-of-time observations - 7.10 DC.csv` | `csv` | 9.70 KiB | `current-support` | `4ed1c9cb2c22a8fce0db43068fc4dc32f07a43962408f4c582fa4d0e88596437` |
| `Shared/Operations/Service/monthly KPI/service access signal production-readiness review - 7.10 DC.eml` | `eml` | 2.16 KiB | `current-support` | `df59cba2c2bde7c075d70e1559faee6d4ecc694eb43a53392fa27f60033a3c6f` |
| `Shared/Operations/Service/monthly KPI/service callback action window - 7.6 DC.eml` | `eml` | 2.42 KiB | `current-support` | `6a1332edc7a9a75a9e13333d0e37437752435120522a656542006393b1c126e9` |
| `Shared/Operations/Service/monthly KPI/service callback evidence follow-up - 7.3 SA.eml` | `eml` | 2.94 KiB | `current-support` | `6f2cf3b87a624117a7d96a92760d96f74b54033c1ea0f1bc0ea5f27f28bdc7d0` |
| `Shared/Operations/Service/monthly KPI/service commitment timing decision - 7.8 DC.eml` | `eml` | 4.40 KiB | `current-support` | `c1c3dfae2252a4ec211d2a6e223a0d2d080421db4af0826a1c5dc162bf072f9f` |
| `Shared/Operations/Service/monthly KPI/service operating-review allocation economics - 7.7 DC.eml` | `eml` | 6.02 KiB | `current-support` | `b8cf9309dba6f73b1b11142247b9bef94c2b85907e4757887bfcf7855f16ab74` |
| `Shared/Operations/Service/monthly KPI/service pricing + callback review notes - SA 7.1.docx` | `docx` | 39.99 KiB | `current-support` | `608c73e84f0a5bb4a67e9c913536bfc4522710a2903586c85c8080677ba2f518` |
