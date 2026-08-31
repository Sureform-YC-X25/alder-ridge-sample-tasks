from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    slug: str
    title: str
    workflow: str
    output_mode: str
    prompt: str
    difficulty: str = 'advanced'


TASKS: tuple[TaskSpec, ...] = (
    TaskSpec(
        task_id='task_001',
        slug='arm-2409-wip-backcharge',
        title='ARM-2409 June WIP controller sign-off memorandum',
        workflow='project-accounting',
        output_mode='document_edit',
        prompt='Finalize the ARM-2409 June WIP sign-off memorandum at `Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx`.\n\nGive the Controller a review-ready June 30 close recommendation. Reconcile the final May WIP through June accounting activity and the current PM and commercial evidence, resolve PCO-011 under the signed WIP policy, and quantify the resulting WIP and proposed posting. Confirm the complete June job-cost and billing populations tie to the posted ledger or clearly identify any exceptions. Document the controlling sources, open items, owners, and approval and posting boundary.\n\nComplete only the working memorandum and save it in place. This is preparer support: do not post a journal, imply Controller approval, alter another close file, or create separate support exports. A brief completion note is sufficient in the final response.\n',
        difficulty='advanced',
    ),
    TaskSpec(
        task_id='task_004',
        slug='complete-june-wip-risk-template',
        title='Complete the four-project WIP risk template',
        workflow='controllership',
        output_mode='spreadsheet_edit',
        prompt="Complete the Controller's four-project June WIP risk review at\n`Shared/Finance/Close/2026/06 June/4 WIP/WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx`.\n\nReconcile the June positions to the current accounting records and May close, resolve conflicts in the PM, commercial, and policy evidence, and quantify both the booked close case and material unbooked commercial exposure. Leave a formula-driven accounting tie and an unambiguous release or escalation decision for each project and for the portfolio.\n\nPreserve the source/reference content and workbook style, recalculate, and save in place. Do not post or imply approval. A brief completion note is sufficient in the final response.\n",
        difficulty='advanced',
    ),
    TaskSpec(
        task_id='task_015',
        slug='append-q2-covenant-slide',
        title='Append one Q2 covenant-headroom slide',
        workflow='lender-reporting',
        output_mode='presentation_edit',
        prompt="Append one decision-ready June 30 covenant-headroom slide to\n`Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 lender update - review working v3.pptx`.\n\nReperform the covenants from the executed agreement and current support. Show the posted and proposed-WIP bases, the meaningful downside capacity, the calculation/source basis, and the resulting compliance and circulation conclusion. Keep the proposal visibly unposted.\n\nPreserve the five existing slides and their order, match the deck's visual system, append exactly one slide, and save in place. This remains a working deck; do not sign, submit, or describe it as submitted.\n",
        difficulty='advanced',
    ),
    TaskSpec(
        task_id='task_027',
        slug='complete-fy27-scenario-model',
        title='Complete the FY27 EBITDA scenario model',
        workflow='fpa',
        output_mode='spreadsheet_edit',
        prompt="Please finish the FY27 EBITDA scenario workbook for the next planning review.\n\nPlease work in `Shared/Finance/FP&A/FY27 plan/FY27 EBITDA scenarios - WORKING.xlsx` and save the finished file back there.\n\nThe July FY27 planning correspondence in `Requests/`, including Daniel's follow-up after the first steering review, sets the request. Use the current controller-tied planning book, operating extracts, and dated review notes. Refer to earlier branch drafts only to understand changes, and settle any conflicts using dates, approval status, document status, and the management correspondence. Tie the posted Alder Ridge baseline to the company accounting records, and keep any external evidence tied to its own controlling finance source.\n\nPreserve `Driver Inputs` and `Contingency Inputs`. Complete formula-driven `Base`, `Downside`, `Upside`, `Contingency Plan`, and `Checks` sheets. Each case must include branch revenue, gross margin, fixed and variable opex, EBITDA, EBITDA margin, cash taxes, capex, working-capital investment, and free cash flow. Use the scenario-specific driver table rather than applying one blanket percentage. Add the approved quarterly liquidity bridge through ending cash and revolver for each scenario. Bridge Downside and Upside to Base by branch revenue and margin, consolidated EBITDA and free-cash-flow driver, and liquidity quarter; reconcile each bridge to zero and identify the primary EBITDA driver and liquidity quarter. From the approved downside action register, select the smallest executable action set that restores the board EBITDA-margin and free-cash-flow guardrails, minimize implementation cash cost, and carry the selected actions through a formula-driven 13-week cash and revolver release bridge. Apply the executed lender definitions to Base, Downside, and the selected-action case and release the plan only if the funded-debt, interest, leverage, fixed-charge-coverage, and liquidity conditions agree. Include consolidation, scenario-selection, cash-flow, liquidity, covenant, and source controls.\n\nDocument the controlling sources inside the deliverable. Recalculate the finished file and save it to the requested path.\n\nLeave only the finished file in the shared company workspace, with no draft scripts or temporary exports. Open the saved workbook once more before you finish and make sure it contains no spreadsheet errors.",
        difficulty='expert-long-horizon',
    ),
    TaskSpec(
        task_id='task_035',
        slug='complete-backlog-capacity-model',
        title='Complete the backlog burn and capacity model',
        workflow='fpa-operations',
        output_mode='spreadsheet_edit',
        prompt='Finish the FY27 backlog burn and labor-capacity model at\n`Shared/Finance/FP&A/Backlog/FY27 backlog burn and capacity - WORKING.xlsx` for the next planning review.\n\nReconcile the current planning sources, distinguish the approved probability plan from the full signed-commitment exposure, and determine what can actually be delivered under the approved labor and response constraints. Translate any constraint into the earnings, contractual, sequencing, and recovery decision management must make. Keep the model formula-driven, auditable, and internally consistent with the approved definitions and authority evidence.\n\nPreserve the three input schedules, recalculate the workbook, and save it in place. Leave no spreadsheet errors or temporary files in the shared workspace.',
        difficulty='expert-long-horizon',
    ),
    TaskSpec(
        task_id='task_037',
        slug='create-capital-portfolio-model',
        title='Create the FY27 capital-project portfolio model',
        workflow='capital-planning',
        output_mode='spreadsheet_create',
        prompt='Get the FY27 capital-project portfolio ready for the next capital-committee review.\n\nUse the July capital-committee and long-range-plan correspondence, including the post-committee clarification, as the starting point. Work from the current controller-tied workpapers and operating extracts, and use earlier versions only to trace changes. Resolve any disagreement by checking dates, approval status, document status, and management correspondence. Tie posted Alder Ridge balances to the company accounting records, while keeping external project evidence tied to its controlling finance source.\n\nBuild `Projects`, `Cash Flow Models`, `Portfolio Selection`, and `Checks`. For each project, calculate after-tax cash flows, depreciation tax shield, working-capital investment and release, terminal value, NPV, IRR, payback, and downside NPV. Select the highest-NPV feasible portfolio within the annual cash, debt, technician-capacity, and mandatory-safety constraints. Formulas, rather than hardcoded rankings, must drive the selected portfolio. Include formula-driven `Portfolio NPV` and `Downside portfolio NPV` headlines that aggregate only the selected projects, an independent selection-tie control proving both headlines use that exact set, and the necessary cash, constraint, valuation, and source controls.\n\nAdd a compact formula-driven resilience schedule that re-optimizes the portfolio for each selected nonmandatory-project outage, tighter cash/debt/technician capacity, and a downside-NPV objective. For the downside-NPV objective, retain the same annual cash, debt, technician-capacity, and mandatory-safety constraints, maximize portfolio downside NPV, and do not add a separate minimum base-NPV constraint. Show each replacement portfolio, funding, headroom, NPV sacrifice, binding constraint, and committee release status.\n\nSave the completed workbook as `Deliverables/FY27 capital project portfolio - 7.4.xlsx`. Document the controlling sources, recalculate the saved file, and confirm that the checks clear with no spreadsheet errors. Leave no draft scripts or temporary exports in the shared company workspace.',
        difficulty='expert-long-horizon',
    ),
    TaskSpec(
        task_id='task_055',
        slug='create-horizon-accretion-model',
        title='Create the Horizon acquisition accretion model',
        workflow='corporate-development',
        output_mode='spreadsheet_create',
        prompt="The Horizon acquisition accretion and dilution workbook needs to be ready for the next investment-committee review.\n\nUse the July corporate-development correspondence with the 7/5 controller-tied IC book, current data-room export, diligence review notes, and approved underwriting policy. The 6/29 banker case is change history only. Follow the controller's warning on seller adjustments, revenue synergies, and timing assumptions. Horizon is an external target, so keep the analysis tied to the controlling transaction sources rather than Alder Ridge's posted ledger.\n\nBuild `Buyer Standalone`, `Target Standalone`, `Sources & Uses`, `Purchase Accounting`, `Combined`, `Accretion Dilution`, and `Checks`. Calculate the financing mix, incremental interest and foregone cash yield, seller share issuance, purchase-accounting amortization, phased cost synergies, tax effects, and one-time integration costs. Show year-one and year-two GAAP and adjusted EPS accretion or dilution, with the financing, balance, EPS, and source controls needed for committee review.\n\nAdd a compact committee sensitivity matrix for 50% synergy realization, a 200-basis-point increase in debt cost, and the combined case. Show the year-one and year-two GAAP and adjusted EPS bridges and release status for each case. A stress case requires committee mitigation if either GAAP or adjusted EPS is dilutive in either year; otherwise it is ready for committee release. Normal status labels such as mitigation required, conditional, hold, or ready for release are acceptable when they communicate that decision correctly.\n\nDocument the controlling sources, recalculate the finished workbook, and save it as `Deliverables/Horizon acquisition accretion dilution model - 7.4.xlsx`. Leave no spreadsheet errors, draft scripts, or temporary exports in the shared company workspace.\n",
        difficulty='expert-long-horizon',
    ),
    TaskSpec(
        task_id='task_061',
        slug='complete-tax-provision-model',
        title='Complete the annual income-tax provision model',
        workflow='controllership-tax',
        output_mode='spreadsheet_edit',
        prompt="Complete the FY26 annual income-tax provision for the controller's close review.\n\nStart with the tax-provision and EBITDA-definition correspondence in `Requests/`, including the close-review follow-up. Use the v4 controller-review workpapers, jurisdiction detail, tax provision review notes, and signed reporting policy as the current support. Keep the v2 pre-close workbook as change history only, and resolve differences using approval status, dates, document status, and management correspondence.\n\nUpdate `Shared/Finance/Tax/2026/FY26 tax provision - WORKING.xlsx` in place. Preserve `Trial Balance` and `Tax Adjustments`. Complete `Current Tax`, `Deferred Tax`, `Rate Reconciliation`, and `Checks`. Reconcile pretax book income through current taxable income, keeping permanent, temporary, and current-only adjustments distinct. Apply the federal and jurisdiction-specific apportionment, rate, credit, and NOL rules in the controlling tax support.\n\nShow the temporary-difference DTA, state-credit DTA, valuation allowance, net DTA, DTL, and net deferred-tax liability as separate formula-driven balances at the enacted rates. Present deferred tax expense as positive and a deferred tax benefit as negative. Reconcile current tax and deferred tax to total provision and effective tax rate, with book-tax, deferred roll-forward, rate, and source controls.\n\nComplete the federal and state current-tax payable rollforwards and a balanced close-entry bridge from current and deferred tax through estimated payments, cash, ending payables, and the controller posting-release decision.\n\nFor that close-entry bridge, use the controller-approved opening federal payable of $410,000, opening state payable of $165,000, FY26 federal estimated payments of $1,125,000, and FY26 state estimated payments of $505,000. These four close inputs supplement the v4 workbook and are part of this assignment's controlling materials.\n\nDocument the controlling sources and workpaper status. Recalculate the workbook, confirm the checks clear and no spreadsheet errors remain, and save it in place.\n",
        difficulty='expert-long-horizon',
    ),
    TaskSpec(
        task_id='task_068',
        slug='complete-executive-performance-deck',
        title='Complete the June executive performance review',
        workflow='fpa-investor-relations',
        output_mode='presentation_edit',
        prompt='Finish the June executive performance review at\n`Shared/Finance/Reporting/2026/06 June/June executive performance review - WORKING.pptx` for the leadership meeting.\n\nReconcile the controller-tied actuals to the approved plan, current forecast, operating support, accounting records, and the applicable KPI and reporting policies. Explain the material performance and cash drivers, assess backlog and outlook, and turn the documented risk and action evidence into a quantified guidance recommendation with accountable next steps. Make the executive story and final control tie auditable.\n\nPreserve the nine-slide template, master, section order, and title slide; keep the deck readable and source-footnoted; render-check it; and save it in place. Leave no draft files or temporary exports in the shared workspace.',
        difficulty='advanced-long-horizon',
    ),
    TaskSpec(
        task_id='task_072',
        slug='create-quarterly-financial-narrative',
        title='Create the Q2 earnings and lender narrative',
        workflow='investor-relations',
        output_mode='document_create',
        prompt='Can you finish the Q2 earnings and lender narrative for the next board and bank review?\n\nStart with the Q2 board, lender, and investor-relations correspondence, including the source-tie follow-up, and use the current controller-tied workpapers and operating extracts. Earlier versions are only there for change history. Resolve any conflicts from the approval status, dates, document status, and follow-up correspondence. Tie posted Alder Ridge balances to the company accounting records, and keep market evidence and bank terms tied to their own controlling sources.\n\nSave the finished Word narrative at `Deliverables/Q2 earnings and lender narrative - 7.4.docx`. Use `Executive Summary`, `Quarterly Performance`, `Segment Performance`, `Cash Flow and Liquidity`, `Credit Metrics`, `Outlook`, and `Appendix: KPI Definitions` as the main sections. Explain why Q2 revenue and adjusted EBITDA moved differently from plan, rank the supported operating drivers by financial impact, and reconcile free cash flow. Show lender leverage and fixed-charge coverage using posted funded debt, approved LTM EBITDA, and the executed definitions.\n\nPlease make a clear recommendation on whether to hold or revise published guidance. Ground that call in the signed downside stress, headroom to the published EBITDA low end, headroom to the formal update trigger, and a measurable monitoring threshold. Include the latest approved outlook, the main risks and opportunities, and actions with owners and timing. Use the review notes and signed policy for the controller-tie headline basis. Keep Board Performance, GAAP-posted, IR-comparable, and lender covenant measures separately labeled and reconciled.\n\nUse a real Word table for the quarterly and credit-metric comparison so the actual results, plan or comparative figures, variances, and applicable definitions can be reviewed together.\n\nCite the controlling sources and leave no draft scripts or temporary exports in the shared company workspace.\n',
        difficulty='advanced-long-horizon',
    ),
    TaskSpec(
        task_id='task_073',
        slug='credit-metrics-debt-capacity',
        title='Credit metrics and incremental debt capacity',
        workflow='treasury-investor-relations',
        output_mode='console',
        prompt='Please refresh our credit-capacity readout for the next bank discussion.\n\nStart with the July liquidity and covenant correspondence, the current controller-tied treasury workpapers, and the executed credit agreement. Keep earlier working versions as change history only, and tie posted funded debt to the company accounting records.\n\nFor this capacity readout, the v8 `Credit Outlook` schedule is the controlling management case. Its current pro-forma funded-debt basis equals June 30 posted funded debt plus the documented July 15 $24,000,000 committed acquisition-facility draw. Use the executed agreement for legal definitions and covenant thresholds; use the v8 schedule for the management assumptions and internal ratings limit. The v6 workbook is change history only.\n\nPlease show LTM covenant EBITDA, funded debt, cash, gross and net leverage, interest and fixed charges, coverage, and tangible-net-worth headroom. Then quantify incremental debt capacity under the gross-leverage covenant, fixed-charge-coverage covenant, and internal ratings limit. Include interest on new debt at the documented borrowing rate, identify the binding constraint, and show the pro-forma debt, interest, leverage, coverage, and remaining nonbinding capacity at that limit.\n\nDo not post transactions or change company files. Return one JSON object with these keys: `ltm_covenant_ebitda`, `funded_debt`, `unrestricted_cash`, `gross_leverage`, `net_leverage`, `cash_interest`, `fixed_charges`, `interest_coverage`, `fixed_charge_coverage`, `tangible_net_worth_headroom`, `maximum_debt_at_leverage_limit`, `leverage_debt_capacity`, `fixed_charge_debt_capacity`, `ratings_guardrail_debt_capacity`, `maximum_incremental_debt`, `binding_constraint`, `pro_forma_funded_debt`, `pro_forma_cash_interest`, `pro_forma_gross_leverage`, `pro_forma_net_leverage`, `pro_forma_interest_coverage`, `pro_forma_fixed_charge_coverage`, `remaining_fixed_charge_capacity`, `remaining_ratings_capacity`. Use dollars to two decimals and ratios as decimals to at least four places.',
        difficulty='advanced-long-horizon',
    ),
    TaskSpec(
        task_id='task_100',
        slug='create-cfo-board-strategic-finance-deck',
        title='Create the CFO strategic-finance board deck',
        workflow='cfo-synthesis',
        output_mode='presentation_create',
        prompt='Please finish the FY26 outlook and FY27 priorities deck for the next board meeting.\n\nUse the Q2 board, lender, and investor-relations request and the source-tie follow-up, together with the v7 controller-tied reporting book, current reporting-cube extract, review notes, signed KPI and presentation policy, and Board-approved operating plan. Keep the v5 CFO scratch book for change history only. Resolve any differences using the dates, approval status, document status, and follow-up correspondence.\n\nBuild a board-ready 10-slide deck covering the title, executive decision summary, consolidated forecast versus plan, branch performance, EBITDA bridge, cash and liquidity, leverage and covenant risk, FY27 strategic priorities, decisions and gates, and an appendix with definitions and source reconciliation. Rebuild the operating headlines from branch economics, compare the forecast with the Board-approved plan, identify the largest negative branch EBITDA variance, and show Base, Downside, and Severe cash, leverage, revolver headroom, and guardrail results.\n\nDistinguish Approved, Gated, Conditional, and Deferred priorities and identify the executable board decision portfolio under the documented cash, count, dependency, status, and downside constraints. Show total probability-weighted benefit, executable benefit, incremental financing, remaining revolver headroom, any unfunded shortfall, ending cash and leverage, and the post-action guardrail results.\n\nUse editable charts and tables, concise board narrative, page numbers, source footnotes, consistent units, and a quantitative source reconciliation for revenue, EBITDA, cash, and covenant measures. Save the finished deck as `Deliverables/FY26 outlook and FY27 priorities board deck - 7.6.pptx`. Render and review every slide, and leave no draft scripts or temporary exports in the shared company workspace.',
        difficulty='expert-adversarial-long-horizon',
    ),
)

TASK_BY_ID = {task.task_id: task for task in TASKS}
TASK_BY_SLUG = {task.slug: task for task in TASKS}
