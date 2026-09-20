from __future__ import annotations
import json
import math
import re
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation
from runtime.grading.apex import Criterion, SEED_WORKSPACE, _answer_mapping, _close, _get, _normalize, _number, _recalculated_data_workbook, _result
from runtime.grading.semantic import contains_concept, date_matches, iter_numeric_candidates, ordered_semantic_list_matches, semantic_equal, semantic_value_matches, unordered_semantic_list_matches
from runtime.grading.hybrid_semantic import semantic_requirement
GOLD_PATH = Path(__file__).resolve().parent / 'gold' / 'tasks_026_100.json'
SAMPLE_CORPORATE_TASK_IDS = frozenset({'task_035', 'task_068'})
_TASK_035_LABEL_ALIASES = {'signed_backlog': ['Signed backlog', 'Portfolio signed backlog'], 'fy27_backlog_revenue_burn': ['Approved probability plan', 'Probability-plan revenue', 'Probability-weighted revenue', 'Risk-adjusted backlog plan', 'Weighted delivery plan', 'Planning-case revenue', 'Expected backlog revenue', 'FY27 backlog revenue'], 'fy27_required_hours': ['Approved probability plan', 'Probability-plan required hours', 'Probability-weighted required hours', 'Risk-adjusted required hours', 'Planning-case hours'], 'fy27_available_hours': ['FY27 available hours', 'Base available hours'], 'constrained_month_count': ['Constrained months', 'Months constrained'], 'first_constrained_month': ['First constrained month', 'Probability-plan first constrained month'], 'maximum_capacity_shortfall_hours': ['Maximum capacity shortfall', 'Probability-plan maximum shortfall'], 'probability_plan_unresolved_hours': ['Approved probability plan', 'Probability-plan unresolved hours', 'Risk-adjusted unresolved hours', 'Weighted-plan residual hours'], 'probability_plan_remediation_cost': ['Approved probability plan', 'Probability-plan remediation cost', 'Risk-adjusted response cost', 'Weighted-plan capacity cost'], 'gross_commitment_revenue_burn': ['Full signed commitment', 'Gross-commitment revenue', '100% signed backlog', 'Signed-backlog stress revenue', 'Contracted exposure revenue', 'Full-obligation revenue'], 'gross_commitment_required_hours': ['Full signed commitment', 'Gross-commitment required hours', '100% signed-backlog hours', 'Signed-backlog stress hours', 'Full-obligation hours'], 'gross_commitment_first_constrained_month': ['Full signed commitment', 'Gross-commitment first constrained month', 'Signed-backlog stress first constrained month'], 'gross_commitment_maximum_capacity_shortfall_hours': ['Gross-commitment maximum shortfall', 'Full-commitment maximum shortfall', 'Signed-backlog stress maximum shortfall'], 'gross_commitment_unresolved_hours': ['Full signed commitment', 'Gross-commitment unresolved hours', 'Signed-backlog stress residual hours', 'Full-obligation unresolved hours'], 'gross_commitment_remediation_cost': ['Full signed commitment', 'Gross-commitment remediation cost', 'Signed-backlog stress response cost', 'Full-obligation capacity cost'], 'gross_commitment_revenue_variance_to_probability_plan': ['Gross-commitment revenue variance', 'Revenue variance to probability plan'], 'gross_commitment_required_hours_variance_to_probability_plan': ['Gross-commitment required-hours variance', 'Required-hours variance to probability plan'], 'gross_commitment_unresolved_hours_variance_to_probability_plan': ['Gross-commitment unresolved-hours variance', 'Unresolved-hours variance to probability plan'], 'gross_commitment_remediation_cost_variance_to_probability_plan': ['Gross-commitment remediation-cost variance', 'Remediation-cost variance to probability plan'], 'gross_commitment_release_status': ['Full signed commitment', 'Gross-commitment release conclusion', 'Signed-backlog stress status', 'Full-obligation release status'], 'controlling_capacity_case': ['Controlling case', 'Capacity case controlling', 'Binding capacity case', 'Governing delivery case', 'Capacity planning basis'], 'execution_portfolio_completed_revenue': ['Portfolio completed revenue', 'Execution completed revenue', 'Deliverable portfolio revenue', 'Capacity-constrained delivered revenue'], 'execution_portfolio_ending_deferred_revenue': ['Portfolio deferred revenue', 'Ending deferred revenue', 'Delivery carryforward revenue', 'Capacity-constrained carryforward'], 'execution_portfolio_revenue_check_delta': ['Portfolio revenue check', 'Revenue bridge check'], 'execution_portfolio_ending_deferred_hours': ['Portfolio deferred hours', 'Ending deferred hours'], 'execution_portfolio_completed_gross_profit': ['Portfolio completed gross profit', 'Completed gross profit'], 'execution_portfolio_liquidated_damages': ['Portfolio liquidated damages', 'Liquidated damages'], 'execution_portfolio_net_gross_profit_after_damages': ['Portfolio net gross profit', 'Executable gross profit after damages', 'Deliverable portfolio net margin', 'Capacity-constrained net gross profit'], 'execution_portfolio_highest_priority_deferred_project': ['Highest-priority deferred project', 'Priority deferred item'], 'execution_portfolio_release_status': ['Portfolio release decision', 'Execution release status', 'Delivery portfolio disposition', 'Capacity-constrained release status'], 'release_bridge_probability_plan_gross_profit_after_remediation': ['Probability-plan GP after response', 'Probability-plan GP after remediation', 'Weighted-plan GP after capacity cost', 'Risk-adjusted GP after response'], 'release_bridge_gross_commitment_gross_profit_after_remediation': ['Gross-commitment GP after response', 'Gross-commitment GP after remediation', 'Signed-backlog GP after response', 'Full-obligation GP after capacity cost'], 'release_bridge_execution_net_gross_profit_after_damages': ['Executable GP after damages', 'Execution net GP after damages', 'Deliverable portfolio GP after contract cost', 'Capacity-constrained net GP'], 'release_bridge_deferred_revenue_at_risk': ['Deferred revenue at risk', 'Deferred revenue'], 'release_bridge_management_decision': ['Management decision', 'Release decision', 'Portfolio disposition', 'Management recommendation'], 'executive_recovery_probability_plan_gp_after_remediation': ['Probability-plan GP', 'Probability-plan GP after response'], 'executive_recovery_gross_commitment_gp_after_remediation': ['Gross-commitment GP', 'Gross-commitment GP after response'], 'executive_recovery_execution_net_gp_after_damages': ['Executable GP after damages', 'Execution net GP after damages'], 'executive_recovery_gross_commitment_shortfall': ['Gross-commitment shortfall', 'Earnings shortfall'], 'executive_recovery_deferred_revenue_at_risk': ['Deferred revenue', 'Deferred revenue at risk'], 'executive_recovery_deferred_gross_profit_at_risk': ['Deferred gross profit', 'Deferred gross profit at risk'], 'executive_recovery_liquidated_damages': ['Damages', 'Liquidated damages'], 'executive_recovery_required_recovery': ['Recovery required', 'Required recovery'], 'executive_recovery_highest_priority_deferred_project': ['Priority item', 'Highest-priority deferred project'], 'executive_recovery_decision': ['Recommendation', 'Executive decision', 'Management recommendation', 'Release recommendation', 'Portfolio disposition']}
_TASK_035_INPUT_SHEETS = ('Contract Backlog', 'Capacity Inputs', 'Execution Authority')

def _task_035_output_sheet_names(workbook) -> tuple[str, ...]:
    """Return every authored-output surface while excluding seeded sources."""
    return tuple((name for name in workbook.sheetnames if name not in _TASK_035_INPUT_SHEETS))
_TASK_035_MODEL_CONTROL_MEANING = {'model_content__burn_curve': 'a formula-driven monthly revenue or backlog burn schedule', 'model_content__required_hours': 'formula-driven productive labor hours required by the delivery schedule', 'model_content__available_hours': 'formula-driven labor capacity or available productive hours', 'model_content__overtime': 'formula-driven overtime capacity or usage', 'model_content__subcontract': 'formula-driven subcontract or contingent-labor capacity or usage', 'model_content__capacity_gap': 'a formula-driven capacity variance, shortfall, surplus, or residual', 'controls__source': 'a completed formula-driven source-population tie or completeness control', 'controls__period': 'a completed formula-driven period, date-coverage, or cutoff control', 'controls__scenario': 'a completed formula-driven scenario, case, or planning-basis control', 'controls__unit': 'a completed formula-driven unit-consistency or unit-conversion control', 'controls__version': 'a completed formula-driven current-versus-prior or source-version control', 'controls__check': 'a completed formula-driven roll-forward, bridge, reconciliation, or cross-foot control', 'controls__model_status': 'a completed formula-driven overall model or review status'}
_TASK_035_SHARED_ROW_ALIASES = frozenset(['approved probability plan', 'deferred revenue', 'deferred revenue at risk', 'executable gp after damages', 'execution net gp after damages', 'full signed commitment', 'gross commitment gp after response', 'highest priority deferred project', 'liquidated damages', 'management recommendation', 'portfolio disposition', 'probability plan gp after response'])
_TASK_068_COMPATIBILITY_LABEL_ALIASES = {'q2_revenue': ['Q2 revenue', 'Board Performance Q2 revenue'], 'q2_approved_plan_revenue': ['Q2 revenue plan', 'Approved Q2 plan revenue'], 'q2_revenue_variance_to_plan': ['Q2 revenue variance to plan', 'Revenue variance vs plan', 'Revenue variance to plan', 'Q2 revenue variance'], 'q2_organic_growth': ['Organic growth'], 'q2_adjusted_ebitda': ['Adjusted EBITDA'], 'q2_approved_plan_adjusted_ebitda': ['Q2 adjusted EBITDA plan', 'Approved Q2 plan adjusted EBITDA', 'Adjusted EBITDA plan'], 'q2_adjusted_ebitda_variance_to_plan': ['Q2 adjusted EBITDA variance to plan', 'Adjusted EBITDA variance vs plan', 'EBITDA variance to plan', 'Adjusted EBITDA variance'], 'construction_gross_profit_impact': ['Construction gross profit impact', 'Construction gross profit downside'], 'service_labor_productivity_impact': ['Service labor productivity impact', 'Service labor downside'], 'controls_mix_impact': ['Controls mix impact', 'Controls mix downside'], 'q2_free_cash_flow': ['Free cash flow'], 'maximum_revolver': ['Maximum downside revolver', 'FY27 downside maximum revolver', 'Downside revolver'], 'latest_full_year_revenue_outlook': ['Latest approved outlook', 'Revenue outlook'], 'ltm_adjusted_ebitda': ['LTM adjusted EBITDA', 'Approved LTM adjusted EBITDA', 'Lender adjusted EBITDA'], 'funded_debt': ['Posted funded debt', 'Lender funded debt'], 'lender_leverage': ['Lender leverage', 'Covenant leverage', 'Gross leverage'], 'fixed_charge_coverage': ['Fixed-charge coverage', 'Fixed charge coverage', 'FCCR', 'Lender FCCR'], 'revenue_guidance_low': ['Revenue guidance low', 'Published revenue guidance low'], 'revenue_guidance_high': ['Revenue guidance high', 'Published revenue guidance high'], 'ebitda_guidance_low': ['EBITDA guidance low', 'Published EBITDA guidance low'], 'ebitda_guidance_high': ['EBITDA guidance high', 'Published EBITDA guidance high'], 'largest_downside_driver': ['Largest downside driver', 'Principal downside driver'], 'combined_stress_ebitda': ['Combined stress EBITDA', 'Signed downside stress EBITDA'], 'ebitda_shortfall_to_guidance_low': ['EBITDA shortfall to guidance low', 'Shortfall to published EBITDA low end'], 'additional_revenue_decline_to_update_trigger': ['Additional revenue decline to update trigger', 'Headroom to formal update trigger'], 'guidance_update_required': ['Guidance update required', 'Revise published guidance']}
_ARTIFACT_TOKEN_ALIASES = {'management correspondence': ['management correspondence', 'management email', 'request thread', 'Requests/', 'controller follow-up', 'source-tie thread', 'board review', 'review notes'], 'executive summary': ['executive summary', 'executive performance summary', 'performance summary'], 'cash roll forward': ['cash roll forward', 'cash roll-forward', 'cash liquidity scenario', 'ending cash before financing', 'cash & liquidity'], 'risk': ['risk', 'downside', 'guardrail', 'failure', 'shortfall'], 'version': ['version', 'source version', 'document version', 'file version', 'v4', 'v6', 'v7.4', 'v7.6', 'v8', 'current approved', 'current source', 'controller-tied', 'controlling case', 'current - controller tie complete'], 'unit': ['unit', 'units', 'USD in thousands', 'USD thousands', '$000', '$mm', 'USD millions'], 'model status': ['model status', 'overall status', 'control status', 'feasibility', 'model integrity']}

def load_corporate_finance_gold(task_id: str | None=None) -> dict[str, Any]:
    payload = json.loads(GOLD_PATH.read_text(encoding='utf-8'))
    return payload[task_id] if task_id else payload

def _boolean_matches(actual: Any, expected: bool) -> bool:
    if isinstance(actual, bool):
        return actual is expected
    normalized = _normalize(actual)
    positive = {'true', 'yes', 'y', 'pass', 'compliant', 'met'}
    negative = {'false', 'no', 'n', 'fail', 'noncompliant'}
    if normalized in positive | negative | {'1', '0', 'not met'}:
        return normalized in positive | {'1'} if expected else normalized in negative | {'0', 'not met'}
    words = set(normalized.split())
    positive_hit = bool(words & positive)
    negative_hit = bool(words & negative) or 'not met' in normalized
    return positive_hit and (not negative_hit) if expected else negative_hit and (not positive_hit)

def _xlsx_label_cell_index(workbook):
    """Index normalized workbook labels once for repeated atomic criteria.

    The production workbooks deliberately expose hundreds of controller-tied
    outputs under exact row labels.  Re-normalizing every cell for every atomic
    criterion is quadratic and can consume the residual rollout wall clock.
    Exact-label lookups cover the normal template path; callers retain their
    existing professional-alias fallback when an exact label is absent or does
    not carry the expected value.
    """
    cached = getattr(workbook, '_alder_label_cell_index', None)
    if cached is not None:
        return cached
    cells: list[tuple[str, int, int, Any]] = []
    exact: dict[str, list[tuple[str, int, int, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                record = (sheet_name, cell.row, cell.column, cell.value)
                cells.append(record)
                normalized = _normalize(cell.value)
                if normalized:
                    exact.setdefault(normalized, []).append(record)
    cached = (cells, exact)
    setattr(workbook, '_alder_label_cell_index', cached)
    return cached

def _xlsx_label_value(workbook, values, label: str, expected: Any, *, directional_strings: bool=False) -> bool:
    wanted = _normalize(label)
    (indexed_cells, exact_labels) = _xlsx_label_cell_index(workbook)

    def record_matches(record: tuple[str, int, int, Any]) -> bool:
        (sheet_name, row_number, column_number, _value) = record
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else workbook[sheet_name]
        candidates = [value_sheet.cell(row_number, column_number + offset).value for offset in (1, 2, 3)]
        candidates.append(value_sheet.cell(row_number + 1, column_number).value)
        return _value_candidates_match(candidates, expected, directional_strings=directional_strings)
    exact_records = exact_labels.get(wanted, [])
    if any((record_matches(record) for record in exact_records)):
        return True
    exact_record_ids = {(record[0], record[1], record[2]) for record in exact_records}
    for record in indexed_cells:
        (sheet_name, row_number, column_number, value) = record
        if (sheet_name, row_number, column_number) in exact_record_ids:
            continue
        if not (semantic_equal(value, label) or contains_concept(value, label)):
            continue
        if record_matches(record):
            return True
    return False

def _xlsx_label_formula(workbook, label: str) -> tuple[bool, str]:
    wanted = _normalize(label)
    (indexed_cells, exact_labels) = _xlsx_label_cell_index(workbook)
    label_locations: list[str] = []
    seen: set[tuple[str, int, int]] = set()

    def check_record(record: tuple[str, int, int, Any]) -> tuple[bool, str | None]:
        (sheet_name, row_number, column_number, _value) = record
        sheet = workbook[sheet_name]
        coordinate = sheet.cell(row_number, column_number).coordinate
        label_locations.append(f'{sheet_name}!{coordinate}')
        candidates = [sheet.cell(row_number, column_number + offset) for offset in (1, 2, 3) if column_number + offset <= sheet.max_column]
        if row_number + 1 <= sheet.max_row:
            candidates.append(sheet.cell(row_number + 1, column_number))
        for candidate in candidates:
            formula = candidate.value
            if not isinstance(formula, str) or not formula.startswith('='):
                continue
            if re.search("(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\\$?[A-Z]{1,3}\\$?\\d+", formula):
                return (True, f'{sheet_name}!{candidate.coordinate}={formula}')
        return (False, None)
    for record in exact_labels.get(wanted, []):
        seen.add((record[0], record[1], record[2]))
        (matched, evidence) = check_record(record)
        if matched:
            return (True, str(evidence))
    for record in indexed_cells:
        identity = (record[0], record[1], record[2])
        if identity in seen:
            continue
        if not (semantic_equal(record[3], label) or contains_concept(record[3], label)):
            continue
        (matched, evidence) = check_record(record)
        if matched:
            return (True, str(evidence))
    if label_locations:
        return (False, f'label found at {label_locations!r}, but no adjacent source-linked formula')
    return (False, 'headline label not found')

def _is_source_linked_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith('=') and (re.search("(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\\$?[A-Z]{1,3}\\$?\\d+", value) is not None)

def _workbook_sheet_lineage_graph(workbook) -> tuple[dict[str, set[str]], dict[tuple[str, str], list[str]]]:
    """Build exact direct and named-range sheet dependencies once.

    Excel models routinely stage source records through workbook-scoped named
    ranges before a downstream schedule consumes them. Those references are no
    less deterministic or auditable than a literal ``'Source'!A1`` token. The
    graph also permits a material downstream sheet to reach a source through an
    intermediate calculation schedule.
    """
    cached = getattr(workbook, '_alder_sheet_lineage_graph', None)
    if cached is not None:
        return cached
    name_sources: dict[str, set[str]] = {}
    for defined_name in workbook.defined_names.values():
        if getattr(defined_name, 'type', None) != 'RANGE':
            continue
        try:
            destinations = {str(sheet_name) for (sheet_name, _range) in defined_name.destinations}
        except (AttributeError, TypeError, ValueError):
            continue
        if destinations:
            name_sources.setdefault(str(defined_name.name).casefold(), set()).update(destinations)
    graph = {sheet_name: set() for sheet_name in workbook.sheetnames}
    evidence: dict[tuple[str, str], list[str]] = {}
    sheet_names = tuple(workbook.sheetnames)
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                formula = cell.value
                if not (isinstance(formula, str) and formula.startswith('=')):
                    continue
                formula_folded = formula.casefold()
                referenced: set[tuple[str, str]] = set()
                for source_name in sheet_names:
                    escaped = source_name.replace("'", "''").casefold()
                    if f"'{escaped}'!" in formula_folded or f'{source_name.casefold()}!' in formula_folded:
                        referenced.add((source_name, 'direct'))
                for (defined_name, source_names) in name_sources.items():
                    if re.search(f'(?<![A-Za-z0-9_.]){re.escape(defined_name)}(?![A-Za-z0-9_.])', formula_folded, flags=re.I):
                        referenced.update(((source_name, f'name:{defined_name}') for source_name in source_names))
                for (source_name, method) in referenced:
                    if source_name == sheet.title:
                        continue
                    graph[sheet.title].add(source_name)
                    evidence.setdefault((sheet.title, source_name), []).append(f'{sheet.title}!{cell.coordinate}={formula} [{method}]')
    cached = (graph, evidence)
    setattr(workbook, '_alder_sheet_lineage_graph', cached)
    return cached

def _xlsx_sheet_lineage(workbook, target_name: str, source_name: str) -> tuple[bool, str]:
    """Prove direct, named-range, or transitive worksheet formula lineage."""
    if target_name not in workbook.sheetnames or source_name not in workbook.sheetnames:
        return (False, f'target={target_name!r} or source={source_name!r} sheet is missing')
    (graph, edge_evidence) = _workbook_sheet_lineage_graph(workbook)
    queue: list[tuple[str, tuple[str, ...]]] = [(target_name, (target_name,))]
    visited = {target_name}
    while queue:
        (current, path) = queue.pop(0)
        for dependency in sorted(graph.get(current, ())):
            candidate_path = (*path, dependency)
            if dependency == source_name:
                examples: list[str] = []
                for (upstream, downstream) in zip(candidate_path, candidate_path[1:]):
                    examples.extend(edge_evidence.get((upstream, downstream), ())[:1])
                return (True, f"lineage_path={' -> '.join(candidate_path)}; examples={examples!r}")
            if dependency not in visited:
                visited.add(dependency)
                queue.append((dependency, candidate_path))
    return (False, f'no direct, named-range, or transitive formula path from {target_name!r} to {source_name!r}')

def _display_tolerance(expected: float) -> float:
    if abs(expected) <= 10:
        return 0.0005
    return max(0.02, abs(expected) * 0.0005)

def _sample_directional_semantic_value_matches(actual: Any, expected: str) -> bool:
    """Accept professional sample-task wording without accepting its opposite.

    Several sample outputs contain decision states and classifications where
    keyword overlap alone is unsafe (``ready ... no``, ``hold is not
    required``, or ``leverage is not the binding constraint``). This guard is
    intentionally called only from the sample-task paths below so it does not
    change grading semantics for the other Alder Ridge tasks.
    """
    if not semantic_value_matches(actual, expected):
        return False
    normalized = _normalize(actual)
    expected_normalized = _normalize(expected)
    if not normalized or not expected_normalized:
        return False
    if 'within commitment and cash floor' in expected_normalized:
        clean = normalized.replace('no breach', '').replace('no shortfall', '')
        if any((phrase in clean for phrase in ('breach', 'shortfall', 'over commitment', 'below cash floor', 'outside commitment', 'insufficient liquidity'))):
            return False
    if 'capacity cleared' in expected_normalized and any((phrase in normalized for phrase in ('not cleared', 'sequencing required', 'decision required', 'recovery required', 'rephase required', 'shortfall remains', 'constraint remains', 'insufficient capacity'))):
        return False
    if expected_normalized == 'selected' and any((phrase in normalized for phrase in ('not selected', 'unselected', 'excluded'))):
        return False
    if expected_normalized == 'not selected' and (not any((phrase in normalized for phrase in ('not selected', 'unselected', 'excluded')))):
        return False
    if 'feasible alternative' in expected_normalized:
        return not any((phrase in normalized for phrase in ('not feasible', 'no feasible', 'infeasible', 'no alternative', 'feasible alternative no', 'feasible alternative is false')))
    if 'ready for controller posting' in expected_normalized and any((phrase in normalized for phrase in ('not ready', 'do not post', 'posting prohibited', 'hold for'))):
        return False
    guidance_text = re.sub('guidance update required\\s*(?::|=|-)?\\s*(?:false|no|0)\\b', '', normalized)
    if 'maintain with heightened monitoring' in expected_normalized and any((phrase in guidance_text for phrase in ('do not maintain', 'guidance update required', 'update guidance now', 'revise guidance now', 'withdraw guidance now', 'reduce guidance now', 'cut guidance now'))):
        return False
    if expected_normalized == 'release' or expected_normalized.startswith('release only'):
        if any((phrase in normalized for phrase in ('do not release', 'not released', 'release not approved', 'hold', 'withhold', 'not ready'))):
            return False
    if any((phrase in expected_normalized for phrase in ('hold', 'sequencing required', 'decision required', 'mitigation required', 'recovery plan'))) and any((phrase in normalized for phrase in ('hold not required', 'no hold', 'sequencing not required', 'decision not required', 'mitigation not required', 'no mitigation', 'unconditional release', 'release approved', 'approved to proceed', 'clear to proceed'))):
        return False
    escaped = re.escape(expected_normalized).replace('\\ ', '\\s+')
    if re.search(f'\\b{escaped}\\b\\s*(?:(?:is|are|was|were|remains?)\\s+|[:\\-]\\s*)?(?:not|no)\\b', normalized):
        return False
    return True

def _value_candidates_match(candidates: Iterable[Any], expected: Any, *, directional_strings: bool=False) -> bool:
    if isinstance(expected, list):
        return ordered_semantic_list_matches(list(candidates), expected)
    if isinstance(expected, str):
        matcher = _sample_directional_semantic_value_matches if directional_strings else semantic_value_matches
        return any((date_matches(candidate, expected) or matcher(candidate, expected) for candidate in candidates))
    if isinstance(expected, bool):
        return any((_boolean_matches(candidate, expected) for candidate in candidates))
    return any((_close(candidate, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0) for candidate in candidates))

def _task_035_numeric_targets(sheet, label: str, expected: Any, *, row_number: int, column_number: int) -> list[float]:
    """Return unit-aware targets only when the workbook discloses the scale.

    Task035 outputs are often presented in dollars, $000s, or $mm.  Hours and
    counts must never inherit a currency scale merely because another schedule
    uses one.  Keep the exact target first and add a scaled target only for a
    monetary output with a local row/header/sheet disclosure.
    """
    if not isinstance(expected, (int, float)) or isinstance(expected, bool):
        return []
    normalized_label = _normalize(label)
    monetary = any((token in normalized_label for token in ('backlog', 'revenue', 'gross profit', 'gp ', 'cost', 'damages', 'required recovery'))) and 'hours' not in normalized_label
    targets = [float(expected)]
    if not monetary:
        return targets
    context_values: list[Any] = [cell.value for cell in sheet[row_number] if cell.value not in (None, '')]
    for candidate_row in range(max(1, row_number - 6), row_number):
        value = sheet.cell(candidate_row, column_number).value
        if value not in (None, ''):
            context_values.append(value)
    for candidate_row in range(1, min(sheet.max_row, 6) + 1):
        for candidate_column in range(1, min(sheet.max_column, 12) + 1):
            value = sheet.cell(candidate_row, candidate_column).value
            if value not in (None, ''):
                context_values.append(value)
    context = ' '.join((str(value) for value in context_values)).casefold()
    millions = any((marker in context for marker in ('$mm', '$ mm', 'usd mm', 'usd millions', '$ in millions', 'dollars in millions', 'amounts in millions')))
    thousands = any((marker in context for marker in ('$000', '$ 000', '$000s', 'usd thousands', '$ in thousands', 'dollars in thousands', 'amounts in thousands')))
    if millions:
        targets.append(float(expected) / 1000000.0)
    if thousands:
        targets.append(float(expected) / 1000.0)
    return targets

def _task_035_metric_header_aliases(label: str) -> tuple[str, ...]:
    """Return concise metric headers for a case-row layout."""
    tail = label
    for prefix in ('release_bridge_probability_plan_', 'release_bridge_gross_commitment_', 'release_bridge_execution_', 'executive_recovery_probability_plan_', 'executive_recovery_gross_commitment_', 'executive_recovery_execution_', 'gross_commitment_', 'execution_portfolio_', 'executive_recovery_', 'release_bridge_', 'probability_plan_', 'fy27_'):
        if tail.startswith(prefix):
            tail = tail[len(prefix):]
            break
    human = tail.replace('_', ' ')
    aliases = {'backlog revenue burn': ('revenue', 'revenue burn', 'backlog revenue'), 'required hours': ('required hours', 'productive hours'), 'available hours': ('available hours', 'capacity hours'), 'constrained month count': ('constrained months', 'month count'), 'maximum capacity shortfall hours': ('maximum shortfall', 'capacity shortfall'), 'unresolved hours': ('unresolved hours', 'residual hours'), 'remediation cost': ('remediation cost', 'response cost', 'capacity cost'), 'revenue variance to probability plan': ('revenue variance',), 'required hours variance to probability plan': ('required hours variance',), 'unresolved hours variance to probability plan': ('unresolved hours variance', 'residual hours variance'), 'remediation cost variance to probability plan': ('remediation cost variance', 'response cost variance'), 'gross profit after remediation': ('gross profit after remediation', 'gp after response', 'gp after capacity cost'), 'net gross profit after damages': ('net gross profit', 'gp after damages', 'gross profit after damages'), 'highest priority deferred project': ('highest priority deferred project', 'priority item'), 'management decision': ('management decision', 'recommendation', 'portfolio disposition'), 'decision': ('decision', 'recommendation', 'release recommendation')}
    return (human, *aliases.get(human, ()))

def _task_035_metric_header_present(sheet, label: str, cell) -> bool:
    aliases = _task_035_metric_header_aliases(label)
    for row_number in range(max(1, cell.row - 8), cell.row):
        for column_number in range(max(1, cell.column - 1), min(sheet.max_column, cell.column + 1) + 1):
            value = sheet.cell(row_number, column_number).value
            if value in (None, ''):
                continue
            if any((semantic_equal(value, alias) or contains_concept(value, alias) for alias in aliases)):
                return True
    return False

def _task_035_row_match(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[bool, str]:
    """Accept ordinary planning labels while preserving value association.

    Task035 is intentionally laid out as normal case, portfolio, and executive
    tables rather than an answer-key registry.  A required output may therefore
    be identified by its case/portfolio row instead of a private snake_case
    label.  Match the exact expected value only on a row carrying an approved
    business alias, and require the matching cell itself to be source-linked
    for formula-lineage criteria.
    """
    canonical_label = label.replace('_', ' ')
    aliases = [canonical_label, *_TASK_035_LABEL_ALIASES.get(label, [])]
    specific_aliases = [alias for alias in aliases if alias == canonical_label or _normalize(alias) not in _TASK_035_SHARED_ROW_ALIASES]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            if not any((semantic_equal(cell.value, alias) or contains_concept(cell.value, alias) for cell in row for alias in aliases)):
                continue
            row_has_specific_metric = any((semantic_equal(cell.value, alias) or contains_concept(cell.value, alias) for cell in row for alias in specific_aliases))
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                if require_formula and (not _is_source_linked_formula(cell.value)):
                    continue
                if not row_has_specific_metric and (not _task_035_metric_header_present(sheet, label, cell)):
                    continue
                if isinstance(expected, (int, float)) and (not isinstance(expected, bool)):
                    targets = _task_035_numeric_targets(sheet, label, expected, row_number=cell.row, column_number=cell.column)
                    matched = isinstance(cached, (int, float)) and (not isinstance(cached, bool)) and any((_close(float(cached), target, abs_tol=_display_tolerance(target), rel_tol=0.0) for target in targets))
                else:
                    matched = _value_candidates_match([cached], expected, directional_strings=True)
                if matched:
                    return (True, f'{sheet_name}!{cell.coordinate}={cell.value!r}; cached={cached!r}; associated aliases={aliases!r}')
    return (False, f'no associated Task035 row value matched {label!r}')

def _artifact_text_entries(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    entries: list[str] = []
    if suffix == '.xlsx':
        workbook = load_workbook(path, data_only=False, read_only=False)
        for sheet in workbook.worksheets:
            entries.append(sheet.title)
            entries.extend((str(cell.value) for row in sheet.iter_rows() for cell in row if cell.value is not None))
    elif suffix == '.pptx':
        presentation = Presentation(path)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text:
                    entries.append(shape.text)
                if getattr(shape, 'has_table', False):
                    entries.extend((cell.text for row in shape.table.rows for cell in row.cells))
                    entries.extend((' | '.join((cell.text for cell in row.cells)) for row in shape.table.rows))
                if getattr(shape, 'has_chart', False):
                    for plot in shape.chart.plots:
                        try:
                            categories = [category.label for category in plot.categories]
                        except (AttributeError, TypeError, ValueError):
                            categories = []
                        for series in plot.series:
                            entries.append(str(series.name))
                            try:
                                values = list(series.values)
                            except (AttributeError, TypeError, ValueError):
                                continue
                            for (category, value) in zip(categories, values):
                                if value is not None:
                                    entries.append(f'{category} | {series.name}: {value}')
    elif suffix == '.docx':
        document = Document(path)
        entries.extend((paragraph.text for paragraph in document.paragraphs if paragraph.text))
        entries.extend((cell.text for table in document.tables for row in table.rows for cell in row.cells if cell.text))
    return entries

def _artifact_text(path: Path) -> str:
    return '\n'.join(_artifact_text_entries(path))

def _semantic_evidence_pack(path: Path, workbook=None, values=None, *, max_chars: int=60000) -> str:
    """Render the artifact as compact, grader-side evidence for a bounded judge."""
    chunks: list[str] = [f'ARTIFACT: {path.name}']
    if path.suffix.lower() == '.docx':
        document = Document(path)
        for (index, paragraph) in enumerate(document.paragraphs, start=1):
            if paragraph.text.strip():
                chunks.append(f'PARAGRAPH {index}: {paragraph.text.strip()}')
        for (table_index, table) in enumerate(document.tables, start=1):
            chunks.append(f'TABLE {table_index}:')
            for (row_index, row) in enumerate(table.rows, start=1):
                chunks.append(f'  ROW {row_index}: ' + ' | '.join((cell.text.strip() for cell in row.cells)))
    elif path.suffix.lower() == '.xlsx' and workbook is not None and (values is not None):
        sheet_chunks: list[str] = []
        per_sheet_budget = max(1000, (max_chars - 1000) // max(1, len(workbook.sheetnames)))
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            rendered_rows: list[tuple[int, str, str, str, tuple[Any, ...]]] = []
            for (row_number, row) in enumerate(sheet.iter_rows(), start=1):
                rendered: list[str] = []
                compact: list[str] = []
                searchable: list[str] = []
                semantic_values: list[Any] = []
                for cell in row:
                    if cell.value is None:
                        continue
                    number_format = str(cell.number_format or 'General')
                    format_suffix = f'[FORMAT={number_format!r}]' if number_format != 'General' and any((marker in number_format.casefold() for marker in ('$', '%', 'x'))) else ''
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        cached = value_sheet[cell.coordinate].value
                        rendered.append(f'{cell.coordinate}=FORMULA({cell.value})=>{cached!r}{format_suffix}')
                        compact.append(f'{cell.coordinate}=FORMULA=>{cached!r}{format_suffix}')
                        if cached is not None:
                            searchable.append(str(cached))
                            semantic_values.append(cached)
                    else:
                        rendered.append(f'{cell.coordinate}={cell.value!r}{format_suffix}')
                        compact.append(f'{cell.coordinate}={cell.value!r}{format_suffix}')
                        searchable.append(str(cell.value))
                        semantic_values.append(cell.value)
                if rendered:
                    rendered_rows.append((row_number, '  ' + ' | '.join(rendered), ' '.join(searchable).casefold(), '  ' + ' | '.join(compact), tuple(semantic_values)))
            rendered_sheet = '\n'.join([f'SHEET: {sheet_name}', *(row[1] for row in rendered_rows)])
            if len(rendered_sheet) > per_sheet_budget:
                landmark_terms = ('selected project', 'selected portfolio', 'winning combo', 'recommendation', 'decision', 'owner', 'timing', 'deadline', 'model status', 'overall status', 'control status', 'model integrity', 'source', 'policy', 'correspondence', 'mcp', 'check', 'control', 'maximum', 'variance', 'headroom', 'unfunded', 'conclusion')
                selected_rows: set[int] = set()
                follow_through_terms = ('selected project', 'selected portfolio', 'winning combo')
                for (index, (_, _, searchable, _, _)) in enumerate(rendered_rows):
                    if any((term in searchable for term in landmark_terms)):
                        selected_rows.update(range(max(0, index - 1), min(len(rendered_rows), index + 2)))
                    if any((term in searchable for term in follow_through_terms)):
                        selected_rows.update(range(index, min(len(rendered_rows), index + 11)))
                selector_terms = ('selected portfolio id', 'selected combination id', 'selected row id', 'winning combo', 'winning combination')

                def selector_key(value: Any) -> str | None:
                    if isinstance(value, bool) or value is None:
                        return None
                    if isinstance(value, (int, float)):
                        numeric = float(value)
                        return str(int(numeric)) if numeric.is_integer() else format(numeric, '.15g')
                    text = str(value).strip().casefold()
                    if not text or any((term in text for term in selector_terms)):
                        return None
                    return text if re.fullmatch('[a-z]{0,8}-?\\d{1,8}', text) else None
                selector_values: set[str] = set()
                for (_, _, searchable, _, row_values) in rendered_rows:
                    if any((term in searchable for term in selector_terms)):
                        selector_values.update((key for value in row_values if (key := selector_key(value)) is not None))
                if selector_values:
                    for (index, (_, _, _, _, row_values)) in enumerate(rendered_rows):
                        row_keys = {key for value in row_values if (key := selector_key(value)) is not None}
                        if row_keys & selector_values:
                            selected_rows.update(range(max(0, index - 1), min(len(rendered_rows), index + 2)))
                salient = '\n'.join((rendered_rows[index][3] for index in sorted(selected_rows)))
                salient_budget = int(per_sheet_budget * 0.42)
                if len(salient) > salient_budget:
                    salient = salient[:salient_budget] + '\n[SALIENT ROWS TRUNCATED]'
                head_budget = int(per_sheet_budget * 0.4)
                tail_budget = per_sheet_budget - head_budget - len(salient) - 120
                tail_budget = max(int(per_sheet_budget * 0.12), tail_budget)
                rendered_sheet = rendered_sheet[:head_budget] + f'\n[SALIENT LABELED ROWS FROM {sheet_name}]\n' + salient + f'\n[TRUNCATED MIDDLE OF {sheet_name}; TAIL PRESERVED]\n' + rendered_sheet[-tail_budget:]
                if len(rendered_sheet) > per_sheet_budget:
                    rendered_sheet = rendered_sheet[:per_sheet_budget]
            sheet_chunks.append(rendered_sheet)
        chunks.extend(sheet_chunks)
    else:
        chunks.extend(_artifact_text_entries(path))
    rendered = '\n'.join(chunks)
    if len(rendered) <= max_chars:
        return rendered
    return rendered[:max_chars - 120] + '\n[TRUNCATED AFTER BALANCED DETERMINISTIC EXTRACTION]'

def _scoped_xlsx_label_evidence(workbook, values, label: str, *, aliases: Iterable[str]=()) -> tuple[str, bool, str]:
    """Return only the submitted workbook row needed for one semantic check."""
    (indexed_cells, exact_labels) = _xlsx_label_cell_index(workbook)
    accepted_labels = (label, *tuple(aliases))
    matches = [record for accepted in accepted_labels for record in exact_labels.get(_normalize(accepted), [])]
    if not matches:
        matches = [record for record in indexed_cells if any((semantic_equal(record[3], accepted) or contains_concept(record[3], accepted) for accepted in accepted_labels))]
    if not matches:
        return ('required labeled output row is missing', False, 'labeled output row is missing')
    rendered: list[str] = []
    has_submitted_value = False
    for (sheet_name, row_number, column_number, _) in matches[:3]:
        formula_sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else formula_sheet
        parts: list[str] = []
        for column in range(1, formula_sheet.max_column + 1):
            formula_cell = formula_sheet.cell(row_number, column)
            cached = value_sheet.cell(row_number, column).value
            if formula_cell.value is None and cached is None:
                continue
            if isinstance(formula_cell.value, str) and formula_cell.value.startswith('='):
                parts.append(f'{formula_cell.coordinate}=FORMULA({formula_cell.value})=>{cached!r}')
            else:
                parts.append(f'{formula_cell.coordinate}={formula_cell.value!r}')
            if column > column_number and cached not in (None, '') and (_normalize(cached) not in {'-', '—', 'tbd', 'to be completed'}):
                has_submitted_value = True
        if parts:
            rendered.append(f'SHEET {sheet_name} ROW {row_number}: ' + ' | '.join(parts))
    return ('\n'.join(rendered)[:12000], has_submitted_value, f'labeled row found and a submitted result cell is populated={has_submitted_value}')

def _scoped_xlsx_token_evidence(workbook, values, token: str, *, aliases: Iterable[str]=(), sheet_names: Iterable[str] | None=None) -> tuple[str, bool, str]:
    """Return bounded workbook rows containing one provenance concept."""
    accepted_concepts = (token, *tuple(aliases))
    rows: list[str] = []
    completed = False
    selected_sheets = tuple(sheet_names) if sheet_names is not None else tuple(workbook.sheetnames)
    for sheet_name in selected_sheets:
        if sheet_name not in workbook.sheetnames:
            continue
        formula_sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else formula_sheet
        for row_number in range(1, formula_sheet.max_row + 1):
            row_values = [formula_sheet.cell(row_number, column).value for column in range(1, formula_sheet.max_column + 1)]
            if not any((contains_concept(value, concept) for value in row_values if value is not None for concept in accepted_concepts)):
                continue
            parts: list[str] = []
            for column in range(1, formula_sheet.max_column + 1):
                formula_cell = formula_sheet.cell(row_number, column)
                cached = value_sheet.cell(row_number, column).value
                if formula_cell.value is None and cached is None:
                    continue
                if isinstance(formula_cell.value, str) and formula_cell.value.startswith('='):
                    parts.append(f'{formula_cell.coordinate}=FORMULA({formula_cell.value})=>{cached!r}')
                else:
                    parts.append(f'{formula_cell.coordinate}={formula_cell.value!r}')
                if column > 1 and cached not in (None, '') and (_normalize(cached) not in {'-', '—', 'tbd', 'to be completed'}):
                    completed = True
            if parts:
                rows.append(f'SHEET {sheet_name} ROW {row_number}: ' + ' | '.join(parts))
    return ('\n'.join(rows[:8])[:12000] or 'required supporting row is missing', completed, f'supporting row found with populated submitted content={completed}')
_TASK_035_SOURCE_REFERENCES = {'controller-tied': {'aliases': ('FY27 planning assumptions - v6 controller tie.xlsx', 'FY27 planning assumptions v6 controller tie', 'controller tie', 'controller-tied'), 'answer_key': 'Identify the current controller-tied FY27 planning-assumptions workbook or the current controller-tied locked planning-input tabs, and distinguish them from prior/superseded drafts. The exact filename or an unambiguous shortened internal reference is sufficient because the task does not disclose a required filename.'}, 'management correspondence': {'aliases': ('7.4.26_0711am - FY27 plan first pass + branch submissions.eml', 'FY27 plan first pass + branch submissions', '7.5.26_0618am - FY27 plan review follow-up.eml', 'FY27 plan review follow-up', 'steering review comments', 'controller follow-up', 'management correspondence'), 'answer_key': 'Identify at least one of the two governing FY27 management threads: the July 4 first-pass/branch-submission thread or the July 5 plan-review follow-up thread. Exact filename, subject, or an unambiguous ordinary business reference is sufficient.'}, 'policy': {'aliases': ('FY27 planning definitions + scenario guardrails - APPROVED.pdf', 'FY27 planning definitions + scenario guardrails', 'planning definitions', 'scenario guardrails', 'approved planning policy', 'approved capacity-response policy', 'capacity response policy', 'executed contingent labor framework', 'contingent labor framework', 'policy'), 'answer_key': 'Identify an applicable approved authority used by the model: the FY27 planning-definitions/scenario-guardrails policy, the approved capacity-response policy, or the executed contingent-labor framework. The exact filename or an unambiguous shortened reference is sufficient.'}}

def _pptx_slide_evidence(path: Path, slide_numbers: Iterable[int]) -> tuple[str, bool, str]:
    """Render only the slides relevant to one presentation criterion."""
    stat = path.stat()
    return _pptx_slide_evidence_cached(str(path.resolve()), stat.st_mtime_ns, stat.st_size, tuple(slide_numbers))

@lru_cache(maxsize=256)
def _pptx_slide_evidence_cached(path_text: str, _mtime_ns: int, _size: int, slide_numbers: tuple[int, ...]) -> tuple[str, bool, str]:
    path = Path(path_text)
    presentation = Presentation(path)
    chunks: list[str] = []
    authored_values: list[str] = []
    placeholders = {'', '-', '—', 'tbd', 'to be completed', 'complete with supported conclusion', 'refresh coverage', 'draft data not refreshed'}
    for slide_number in slide_numbers:
        if slide_number < 1 or slide_number > len(presentation.slides):
            continue
        slide = presentation.slides[slide_number - 1]
        entries: list[str] = []
        for shape in _task_068_leaf_shapes(slide.shapes):
            if hasattr(shape, 'text') and str(shape.text or '').strip():
                entries.append(str(shape.text).strip())
            if getattr(shape, 'has_table', False):
                entries.extend((' | '.join((cell.text.strip() for cell in row.cells)) for row in shape.table.rows))
        authored_values.extend((entry for entry in entries if _normalize(entry) not in placeholders and (not _normalize(entry).isdigit())))
        chunks.append(f'SLIDE {slide_number}:\n' + '\n'.join(entries))
    substantive = bool(chunks) and any((len(_normalize(entry).split()) >= 3 for entry in authored_values))
    return ('\n\n'.join(chunks)[:12000] or 'required slide is missing', substantive, f'required slide exists with non-placeholder content={substantive}')
_TASK_035_DETERMINISTIC_TEXT_IDS = {'headline_values__first_constrained_month', 'headline_values__gross_commitment_first_constrained_month', 'headline_values__execution_portfolio_highest_priority_deferred_project', 'headline_values__executive_recovery_highest_priority_deferred_project'}
_TASK_035_EXACT_TEXT_LABELS = frozenset((criterion_id.split('__', 1)[1] for criterion_id in _TASK_035_DETERMINISTIC_TEXT_IDS))

def _task_035_status_reference(label: str, expected: Any) -> dict[str, Any]:
    meaning_by_status = {'capacity cleared': 'available capacity and permitted remedies fully cover required hours; no unresolved shortfall remains', 'executive sequencing required': 'a residual shortfall remains after permitted remedies and leadership must sequence or rephase work', 'portfolio resequencing required': 'the execution schedule leaves deferred work and requires portfolio-level resequencing', 'executive decision required': 'the project remains deferred or constrained and requires an explicit management decision', 'executive portfolio decision required': 'the portfolio retains deferred work and requires an executive portfolio decision before release', 'hold for executive portfolio sequencing': 'hold release until the deferred portfolio is sequenced or rephased', 'recover or rephase deferred priority backlog before release': 'recover capacity for, or rephase, the highest-priority deferred backlog before release', 'hold for executive recovery plan': 'hold release until management approves and demonstrates the required recovery plan'}
    normalized_expected = _normalize(expected)
    decision_scope = 'full decision'
    grading_boundary = 'Judge the complete operational decision in this field. The separately scored amounts, project identifier, and formula outputs do not need to be repeated.'
    if label == 'controlling_capacity_case':
        decision_scope = 'binding capacity basis'
        required_meaning = 'the full signed-backlog or 100% customer-commitment exposure is the binding capacity case, rather than only the probability-weighted planning view. Accept Gross commitment, Signed-backlog stress, Full obligation, Contracted exposure, or an equally clear professional equivalent'
        grading_boundary = 'Judge only which capacity basis controls. The related revenue, hours, cost, and release conclusion are independently scored.'
    elif re.fullmatch('(?:gross_commitment_)?20\\d\\d_\\d\\d_capacity_release_status', label):
        decision_scope = 'monthly capacity status'
        if normalized_expected == 'capacity cleared':
            required_meaning = 'the month has no unresolved capacity constraint. Accept Capacity cleared, Cleared, Within capacity, Capacity available, No constraint, Not constrained, or an equally clear professional equivalent'
        else:
            required_meaning = 'the month retains a capacity shortfall and therefore requires executive sequencing, rephasing, or an equivalent constrained-capacity decision'
        grading_boundary = 'Judge only whether this monthly capacity row communicates clear/available versus constrained/short. Detailed release and recovery actions are scored separately.'
    elif re.fullmatch('execution_20\\d\\d_\\d\\d_release_status', label):
        decision_scope = 'monthly execution status'
        if normalized_expected == 'capacity cleared':
            required_meaning = 'the month is released or cleared because no unresolved work remains. Accept Release/Released/Cleared or an equally clear release-status equivalent; a timing-only label such as On schedule is insufficient'
        else:
            required_meaning = 'the month is held, deferred, or not fully released because work remains. Accept Deferred backlog, Deferred work carried, Hold, or an equivalent. Do not require this monthly status cell to repeat the separately scored portfolio-level resequencing action'
        grading_boundary = 'Judge only whether this monthly row communicates released versus deferred/held. Portfolio escalation and recovery actions are scored in separate overall-decision fields.'
    elif re.fullmatch('release_bridge_20\\d\\d_\\d\\d_management_release_status', label):
        decision_scope = 'monthly earnings-release status'
        if normalized_expected == 'capacity cleared':
            required_meaning = 'the month is released or cleared. Accept Release, Release supported, Release cleared, or an equivalent'
        else:
            required_meaning = 'the month is held or conditional because deferred work or revenue risk remains. Accept Hold, Hold - deferred revenue/backlog risk, Conditional - remediation + deferral, or an equivalent. Do not require the monthly cell to restate the separately scored portfolio action'
        grading_boundary = 'Judge only the monthly release/hold conclusion. The portfolio sequencing action is scored in release_bridge_management_decision.'
    elif re.fullmatch('execution_p_\\d+_release_status', label):
        decision_scope = 'project execution status'
        if normalized_expected == 'capacity cleared':
            required_meaning = 'the project is complete, released, or otherwise cleared with no deferred work. Accept Complete, Complete - release, Release/Released, or an equivalent'
        else:
            required_meaning = 'the project is deferred or held and is not released. Accept Deferred, Deferred backlog, Hold - deferred backlog, or an equivalent. Do not require each project row to repeat the separately scored executive portfolio action'
        grading_boundary = 'Judge only whether this project is released/complete versus deferred/held. Executive portfolio action is scored separately.'
    elif label == 'execution_portfolio_release_status':
        decision_scope = 'portfolio execution status'
        required_meaning = 'the portfolio is not a full unconditional release because deferred work remains. Accept a hold, conditional/partial release with deferred backlog, or an equivalent. The detailed sequencing action is scored separately'
        grading_boundary = 'Judge the overall released-versus-deferred portfolio status, not whether the same cell repeats the management action or recovery condition.'
    elif label == 'executive_recovery_required_action':
        decision_scope = 'executive recovery action'
        required_meaning = 'management must recover capacity for, sequence, or rephase the remaining deferred priority/unresolved backlog. Accept an action that stages the approved remedies and then escalates residual unresolved hours for added capacity, customer schedule relief, or executive sequencing. The action field need not repeat the separately scored release timing'
        grading_boundary = 'Judge the recovery action itself. The highest-priority project, recovery amount, release condition, and hold/release decision are independently scored.'
    elif label == 'executive_recovery_release_condition':
        decision_scope = 'full-portfolio release condition'
        required_meaning = 'full release requires both (1) eliminating or resolving the remaining capacity shortfall, unresolved hours, deferred priority backlog, or the associated deferred revenue/revenue-at-risk through recovery/sequencing/rephasing and (2) curing the required recovery amount or gross-commitment gross-profit shortfall. Accept equivalent causal wording. Merely recovering gross profit while still allowing backlog to remain deferred is insufficient'
        grading_boundary = 'Judge the two-part condition for full portfolio release. Exact amounts and the separate executive decision are independently scored.'
    elif label == 'release_bridge_management_decision':
        decision_scope = 'portfolio management release decision'
        required_meaning = 'hold the gross-commitment/full-portfolio release until the deferred portfolio is sequenced or rephased. It may separately support release of the probability-weighted plan. Reject a conditional or partial full-portfolio release that expressly leaves backlog deferred instead of holding that release for executive sequencing'
        grading_boundary = 'Judge the portfolio-level management release decision. Monthly and project statuses, recovery amounts, action mechanics, and the final recovery condition are independently scored.'
    elif label == 'executive_recovery_decision':
        decision_scope = 'current full-portfolio recovery decision'
        required_meaning = "the gross-commitment/full-portfolio release is currently held or not authorized until executive recovery, sequencing, or rephasing resolves the deferred/constrained work. Accept HOLD plus executive sequencing before release, or conditional wording that unambiguously identifies the unresolved constraint, requires executive sequencing/recovery, and says that constraint persists until the work is resequenced or recovered. Because the submitted field is already labeled as the executive recovery decision, wording such as 'Conditional — executive sequencing required; shortfall remains until deferred work is resequenced' is a valid hold-pending-action equivalent and need not repeat the word release. Reject an unconditional release and reject a present 'conditional release' that merely lists future remediation or sequencing without an explicit until/before dependency showing that the unresolved condition must be cured first"
        grading_boundary = 'Judge only the current full-portfolio hold/release decision and its dependency on executive recovery or sequencing. The recovery action, release-condition mechanics, project, and exact amounts are independently scored and need not be repeated here.'
    else:
        required_meaning = meaning_by_status.get(normalized_expected, str(expected))
    return {'output_label': label, 'decision_scope': decision_scope, 'canonical_status': expected, 'required_meaning': required_meaning, 'grading_boundary': grading_boundary, 'equivalence_rule': 'Accept any ordinary professional wording with the same operational decision and direction. Do not require the canonical phrase. Reject a status that reverses cleared versus held, or omits an action that this field itself is specifically responsible for communicating.'}

def _task_035_preferred_fact_sheets(label: str) -> tuple[str, ...]:
    """Prioritize the ordinary output surfaces most likely to hold a fact.

    This ordering is used only to keep the semantic evidence packet small.  It
    never decides whether a label is acceptable or whether a criterion earns
    credit.
    """
    if label.startswith('execution_portfolio_'):
        return ('Project Execution', 'Checks', 'Gap Analysis', 'Revenue Burn')
    if label.startswith(('release_bridge_', 'executive_recovery_')):
        return ('Checks', 'Project Execution', 'Gap Analysis', 'Revenue Burn')
    if label.startswith('gross_commitment_') or label == 'controlling_capacity_case':
        return ('Gap Analysis', 'Revenue Burn', 'Checks', 'Project Execution')
    if label.startswith('probability_plan_'):
        return ('Gap Analysis', 'Revenue Burn', 'Checks')
    if label in {'fy27_available_hours'}:
        return ('Labor Capacity', 'Gap Analysis', 'Checks')
    return ('Revenue Burn', 'Gap Analysis', 'Labor Capacity', 'Checks')

def _task_035_numeric_cell_matches(sheet, label: str, expected: float, actual: Any, *, row_number: int, column_number: int) -> bool:
    """Prove a cached numeric magnitude under a locally disclosed unit.

    Cached formulas retain their underlying precision, so the tolerance is
    deliberately much tighter than a rendered-text tolerance.  Either sign is
    admitted at this objective gate; the semantic reviewer decides whether an
    accounting negative, adverse magnitude, or positive balance expresses the
    requested business meaning.
    """
    if not isinstance(actual, (int, float)) or isinstance(actual, bool):
        return False
    return any((_close(abs(float(actual)), abs(float(target)), abs_tol=max(0.02, abs(float(target)) * 1e-06), rel_tol=0.0) for target in _task_035_numeric_targets(sheet, label, expected, row_number=row_number, column_number=column_number)))

def _task_035_numeric_literal_matches(literal: str, label: str, expected: float, *, context: str) -> bool:
    """Match a displayed Task035 magnitude without inventing a unit scale."""
    text = literal.replace('−', '-').strip()
    accounting_negative = text.startswith('(') and text.endswith(')')
    cleaned = text.replace('$', '').replace(',', '').strip('() ')
    scale_match = re.search('(?:\\s*)(thousand|k|million|mm|m|billion|bn|b)\\s*$', cleaned, flags=re.I)
    scale_token = scale_match.group(1).casefold() if scale_match else ''
    if scale_match:
        cleaned = cleaned[:scale_match.start()].strip()
    cleaned = cleaned.rstrip('%x×').strip()
    try:
        displayed = float(cleaned)
    except ValueError:
        return False
    if accounting_negative or text.lstrip().startswith('-'):
        displayed = -abs(displayed)
    multiplier = {'thousand': 1000.0, 'k': 1000.0, 'million': 1000000.0, 'mm': 1000000.0, 'm': 1000000.0, 'billion': 1000000000.0, 'bn': 1000000000.0, 'b': 1000000000.0}.get(scale_token, 1.0)
    normalized_label = _normalize(label)
    monetary = any((token in normalized_label for token in ('backlog', 'revenue', 'gross profit', 'gp ', 'cost', 'damages', 'required recovery', 'shortfall'))) and 'hours' not in normalized_label
    normalized_context = context.casefold()
    if monetary and (not scale_token):
        if any((marker in normalized_context for marker in ('$mm', '$ mm', 'usd mm', 'usd millions', '$ in millions', 'dollars in millions', 'amounts in millions'))):
            multiplier = 1000000.0
        elif any((marker in normalized_context for marker in ('$000', '$ 000', '$000s', 'usd thousands', '$ in thousands', 'dollars in thousands', 'amounts in thousands'))):
            multiplier = 1000.0
    displayed *= multiplier
    decimals = len(cleaned.rsplit('.', 1)[1]) if '.' in cleaned else 0
    rounding_tolerance = 0.5 * 10 ** (-decimals) * multiplier
    return _close(abs(displayed), abs(float(expected)), abs_tol=max(0.02, abs(float(expected)) * 1e-06, rounding_tolerance + 1e-12), rel_tol=0.0)

def _task_035_render_header_context(sheet, row_number: int) -> str:
    """Render the nearest visible table headers above a Task035 output row."""
    header_rows: list[tuple[int, list[str]]] = []
    for candidate_row in range(1, row_number):
        header_cells: list[str] = []
        populated_cells = 0
        for cell in sheet[candidate_row]:
            if cell.value not in (None, ''):
                populated_cells += 1
            if isinstance(cell.value, str) and (not cell.value.startswith('=')):
                value = cell.value.strip()
                if value:
                    header_cells.append(f'{cell.coordinate}={value!r}')
        if len(header_cells) >= 2 and len(header_cells) * 2 >= populated_cells:
            header_rows.append((candidate_row, header_cells))
    header_context = '\n'.join((f'SHEET {sheet.title} HEADER ROW {candidate_row}: ' + ' | '.join(header_cells) for (candidate_row, header_cells) in header_rows[-3:]))
    return header_context

def _task_035_render_fact_row(sheet, value_sheet, row_number: int) -> str:
    header_context = _task_035_render_header_context(sheet, row_number)
    parts: list[str] = []
    for cell in sheet[row_number]:
        cached = value_sheet[cell.coordinate].value
        if cell.value is None and cached is None:
            continue
        if isinstance(cell.value, str) and cell.value.startswith('='):
            formula = cell.value
            if len(formula) > 360:
                formula = formula[:357] + '...'
            parts.append(f'{cell.coordinate}=FORMULA({formula})=>{cached!r}')
        else:
            parts.append(f'{cell.coordinate}={cell.value!r}')
    fact_row = f'SHEET {sheet.title} ROW {row_number}: ' + ' | '.join(parts)
    return f'{header_context}\n{fact_row}' if header_context else fact_row

def _task_035_compact_formula_row(sheet, value_sheet, row_number: int) -> str:
    """Keep a formula row interpretable without letting long formulas crowd out peers."""
    row = sheet[row_number]
    cached_cells: list[str] = []
    formula_coordinates: list[str] = []
    source_linked_count = 0
    formula_examples: list[str] = []
    for cell in row:
        cached = value_sheet[cell.coordinate].value
        if cached not in (None, ''):
            rendered = repr(cached)
            if len(rendered) > 120:
                rendered = rendered[:117] + '...'
            cached_cells.append(f'{cell.coordinate}={rendered}')
        if not (isinstance(cell.value, str) and cell.value.startswith('=')):
            continue
        if cached is None:
            continue
        formula_coordinates.append(cell.coordinate)
        if _is_source_linked_formula(cell.value):
            source_linked_count += 1
        if len(formula_examples) < 3:
            formula = cell.value
            if len(formula) > 180:
                formula = formula[:177] + '...'
            formula_examples.append(f'{cell.coordinate}=FORMULA({formula})=>{cached!r}')
    cached_text = ' | '.join(cached_cells[:16])
    return f'SHEET {sheet.title} ROW {row_number}: {cached_text}; CALCULATED_FORMULA_CELLS={formula_coordinates!r}; SOURCE_LINKED_FORMULA_COUNT={source_linked_count}; FORMULA_EXAMPLES={formula_examples!r}'[:850]

def _task_035_scenario_context_evidence(workbook, values, label: str) -> str:
    """Supply ordinary workbook rows that establish a scenario basis."""
    normalized_label = _normalize(label)
    controlling_case = normalized_label == 'controlling capacity case'
    needs_probability = 'probability plan' in normalized_label or controlling_case
    needs_gross = 'gross commitment' in normalized_label or controlling_case
    if not (needs_probability or needs_gross):
        return ''
    terms: list[str] = []
    preferred: list[str] = []
    if needs_probability:
        terms.extend(('risk adjusted', 'execution probability', 'probability weighted', 'probability plan'))
        preferred.extend(('Revenue Burn', 'Gap Analysis', 'Checks'))
    if needs_gross:
        terms.extend(('gross commitment', 'full signed backlog', '100 signed backlog', 'without probability'))
        preferred.extend(('Gap Analysis', 'Revenue Burn', 'Checks'))
    sheet_order = tuple(dict.fromkeys((*preferred, *_task_035_output_sheet_names(workbook))))
    sheet_rank = {name: index for (index, name) in enumerate(sheet_order)}
    candidates: list[tuple[int, int, int, str, int]] = []
    for sheet_name in sheet_order:
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row_number in range(1, sheet.max_row + 1):
            formula_values = [cell.value for cell in sheet[row_number] if isinstance(cell.value, str) and cell.value.startswith('=')]
            if not formula_values:
                continue
            source_text = ' | '.join((str(cell.value) for cell in sheet[row_number] if cell.value not in (None, '')))
            cached_text = ' | '.join((str(value_sheet[cell.coordinate].value) for cell in sheet[row_number] if value_sheet[cell.coordinate].value not in (None, '')))
            normalized_row = _normalize(f'{source_text} {cached_text}')
            overlap = sum((term in normalized_row for term in terms))
            if not overlap:
                continue
            total_bonus = int(any((token in normalized_row for token in ('total', 'tie', 'check'))))
            candidates.append((sheet_rank[sheet_name], -total_bonus, -overlap, sheet_name, row_number))
    candidates.sort()
    selected: list[str] = []
    selected_by_sheet: dict[str, int] = {}
    emitted_headers: set[tuple[str, str]] = set()
    used = 0
    for (_rank, _total, _overlap, sheet_name, row_number) in candidates:
        if selected_by_sheet.get(sheet_name, 0) >= 2:
            continue
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        header = _task_035_render_header_context(sheet, row_number)
        header_key = (sheet_name, header)
        parts: list[str] = []
        if header and header_key not in emitted_headers:
            parts.append(header[:900])
        parts.append(_task_035_compact_formula_row(sheet, value_sheet, row_number))
        rendered = '\n'.join(parts)
        if selected and used + len(rendered) + 1 > 3800:
            continue
        selected.append(rendered)
        used += len(rendered) + 1
        selected_by_sheet[sheet_name] = selected_by_sheet.get(sheet_name, 0) + 1
        if header:
            emitted_headers.add(header_key)
        if len(selected) >= 6:
            break
    return '\n'.join(selected)

def _task_035_exact_text_candidate_matches(actual: Any, expected: Any) -> bool:
    """Recognize objective project IDs and ordinary year-month displays."""
    if actual in (None, ''):
        return False
    expected_text = str(expected).strip()
    actual_text = str(actual).strip()
    if date_matches(actual, expected) or _normalize(actual) == _normalize(expected) or expected_text.casefold() in actual_text.casefold():
        return True
    month_match = re.fullmatch('(\\d{4})-(\\d{2})', expected_text)
    if not month_match:
        return False
    (year, month_number) = month_match.groups()
    month_index = int(month_number)
    if not 1 <= month_index <= 12:
        return False
    month_names = ('january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december')
    month_name = month_names[month_index - 1]
    normalized_actual = _normalize(actual_text)
    return year in normalized_actual and any((token in normalized_actual for token in (month_name, month_name[:3], month_number)))

def _task_035_semantic_fact_evidence(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[str, bool, str]:
    """Separate objective Task035 facts from editable business association.

    Exact magnitudes and exact date/project identifiers are deterministic.  A
    source-linked formula is additionally mandatory for formula-lineage rows.
    The metric, period, scenario, sign, and business-label association is left
    to the bounded semantic reviewer.  This prevents both failure modes: a
    finite alias list cannot reject correct professional wording, and a number
    copied beside the wrong metric cannot earn credit merely because it occurs
    somewhere in the workbook.
    """
    cache = getattr(workbook, '_alder_task_035_semantic_fact_cache', None)
    if cache is None or cache.get('values_id') != id(values):
        cache = {'values_id': id(values), 'results': {}}
        setattr(workbook, '_alder_task_035_semantic_fact_cache', cache)
    cache_key = (label, repr(expected), require_formula)
    cached_result = cache['results'].get(cache_key)
    if cached_result is not None:
        return cached_result
    exact_text = label in _TASK_035_EXACT_TEXT_LABELS
    objective_fact = isinstance(expected, (int, float)) and (not isinstance(expected, bool)) or exact_text
    preferred = _task_035_preferred_fact_sheets(label)
    preferred_rank = {sheet_name: index for (index, sheet_name) in enumerate(preferred)}
    concept_text = ' '.join((label.replace('_', ' '), *_TASK_035_LABEL_ALIASES.get(label, ()), *_task_035_metric_header_aliases(label), str(expected)))
    concept_tokens = {token for token in _normalize(concept_text).split() if token not in {'a', 'all', 'and', 'at', 'by', 'case', 'for', 'from', 'fy27', 'in', 'of', 'on', 'plan', 'portfolio', 'the', 'to', 'total'}}
    candidates: list[tuple[int, int, str, int, str]] = []
    objective_match_count = 0
    decision_candidate_count = 0
    sequence = 0
    for sheet_name in _task_035_output_sheet_names(workbook):
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row_number in range(1, sheet.max_row + 1):
            row = sheet[row_number]
            cached_values = [value_sheet[cell.coordinate].value for cell in row]
            nonblank = [value for value in cached_values if value not in (None, '')]
            row_context = ' | '.join((str(value) for value in nonblank))
            normalized_row = _normalize(row_context)
            row_tokens = set(normalized_row.split())
            matched_coordinates: list[str] = []
            has_text_result = False
            for (cell, cached) in zip(row, cached_values):
                formula_ok = _is_source_linked_formula(cell.value)
                if require_formula and (not formula_ok):
                    continue
                if isinstance(cached, str) and cached.strip():
                    if formula_ok or (row_number > 4 and len(nonblank) >= 2):
                        has_text_result = True
                matched = False
                if isinstance(expected, (int, float)) and (not isinstance(expected, bool)):
                    matched = _task_035_numeric_cell_matches(sheet, label, float(expected), cached, row_number=row_number, column_number=cell.column)
                    if not matched and isinstance(cached, str):
                        matched = any((_task_035_numeric_literal_matches(literal, label, float(expected), context=row_context) for literal in _NUMERIC_LITERAL_PATTERN.findall(cached)))
                elif exact_text and cached not in (None, ''):
                    matched = _task_035_exact_text_candidate_matches(cached, expected)
                if matched:
                    matched_coordinates.append(cell.coordinate)
            if matched_coordinates:
                objective_match_count += 1
            if has_text_result:
                decision_candidate_count += 1
            if objective_fact and (not matched_coordinates):
                continue
            if not objective_fact and (not has_text_result):
                continue
            overlap = len(concept_tokens & row_tokens)
            sheet_bonus = max(0, len(preferred) - preferred_rank.get(sheet_name, len(preferred)))
            exact_bonus = 100 if matched_coordinates else 0
            formula_bonus = 10 if any((_is_source_linked_formula(cell.value) for cell in row)) else 0
            score = exact_bonus + 5 * overlap + sheet_bonus + formula_bonus
            rendered = _task_035_render_fact_row(sheet, value_sheet, row_number)
            if matched_coordinates:
                rendered += f' | EXACT_CANDIDATE_CELLS={matched_coordinates!r}'
            candidates.append((-score, sequence, sheet_name, row_number, rendered))
            sequence += 1
    candidates.sort()
    selected: list[str] = []
    used = 0
    for (_score, _sequence, _sheet, _row, rendered) in candidates:
        if selected and used + len(rendered) + 1 > 12000:
            continue
        selected.append(rendered)
        used += len(rendered) + 1
        if len(selected) >= 24:
            break
    (lexical_evidence, _lexical_populated, _lexical_gate) = _scoped_xlsx_label_evidence(workbook, values, label.replace('_', ' '), aliases=_TASK_035_LABEL_ALIASES.get(label, ()))
    evidence_parts = []
    scenario_context = _task_035_scenario_context_evidence(workbook, values, label)
    if scenario_context:
        evidence_parts.append('SCENARIO-BASIS ROWS (context only; association remains semantic):\n' + scenario_context)
    if lexical_evidence != 'required labeled output row is missing':
        evidence_parts.append('VETTED-LABEL ROWS (evidence only; not a wording gate):\n' + lexical_evidence)
    evidence_parts.append('EXACT-FACT / DECISION CANDIDATE ROWS:\n' + ('\n'.join(selected) if selected else 'No qualifying output row was found.'))
    if objective_fact:
        hard_gate_met = objective_match_count > 0
        gate_evidence = f'deterministic exact fact candidate rows={objective_match_count}; require_source_linked_formula={require_formula}; metric association remains semantic'
    else:
        hard_gate_met = decision_candidate_count > 0
        gate_evidence = f'populated decision candidate rows={decision_candidate_count}; require_source_linked_formula={require_formula}; decision meaning remains semantic'
    result = ('\n\n'.join(evidence_parts)[:12000], hard_gate_met, gate_evidence)
    cache['results'][cache_key] = result
    return result

def _task_035_model_or_control_check(workbook, values, criterion_id: str) -> tuple[bool, str]:
    model_requirements = {'model_content__burn_curve': (('Revenue Burn',), (('revenue',), ('burn',))), 'model_content__required_hours': (('Revenue Burn',), (('required', 'hours'), ('productive', 'hours'))), 'model_content__available_hours': (('Labor Capacity',), (('available', 'hours'), ('capacity', 'hours'))), 'model_content__overtime': (('Labor Capacity', 'Gap Analysis'), (('overtime',), ('extra', 'hours'))), 'model_content__subcontract': (('Labor Capacity', 'Gap Analysis'), (('subcontract',), ('contingent', 'labor'))), 'model_content__capacity_gap': (('Gap Analysis',), (('shortfall',), ('capacity', 'gap'), ('variance',), ('residual',)))}
    if criterion_id in model_requirements:
        (sheet_names, term_sets) = model_requirements[criterion_id]
        matches: list[str] = []
        candidate_sheets = dict.fromkeys((*sheet_names, *_task_035_output_sheet_names(workbook)))
        for sheet_name in candidate_sheets:
            if sheet_name not in workbook.sheetnames:
                continue
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            for label_cell in (cell for row in sheet.iter_rows() for cell in row if cell.value not in (None, '')):
                label_text = _normalize(label_cell.value)
                if not any((all((term in label_text for term in terms)) for terms in term_sets)):
                    continue
                candidates = list(sheet[label_cell.row]) + [sheet.cell(row, label_cell.column) for row in range(1, sheet.max_row + 1)]
                for cell in candidates:
                    cached = value_sheet[cell.coordinate].value
                    if isinstance(cell.value, str) and cell.value.startswith('=') and (cached is not None):
                        matches.append(f'{sheet_name}!{cell.coordinate}={cell.value}=>{cached!r}')
        return (bool(matches), f'formula-driven completed schedule examples={matches[:5]!r}' if matches else 'the relevant schedule exists only as an empty starter row or lacks formula-driven results')
    checks = {'controls__source': ('source population tie', 'source reconciliation', 'input population', 'source completeness'), 'controls__version': ('current versus prior', 'version control', 'source version', 'current source selection'), 'controls__period': ('period completeness', 'date coverage', 'month coverage', 'cutoff control'), 'controls__scenario': ('scenario validity', 'basis control', 'case control', 'planning basis'), 'controls__unit': ('unit conversion', 'unit consistency', 'units check', 'unit control'), 'controls__check': ('roll forward', 'bridge', 'reconciliation', 'cross foot'), 'controls__model_status': ('model status', 'overall status', 'review status')}
    aliases = checks.get(criterion_id)
    if aliases is None:
        return (False, 'no task-specific substantive control is defined')
    attempted: list[str] = []
    candidate_sheets = dict.fromkeys(('Checks', *_task_035_output_sheet_names(workbook)))
    for sheet_name in candidate_sheets:
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            if not any((semantic_equal(cell.value, alias) or contains_concept(cell.value, alias) for cell in row if cell.value is not None for alias in aliases)):
                continue
            populated: list[str] = []
            formula_results: list[str] = []
            for cell in row[1:]:
                cached = value_sheet[cell.coordinate].value
                if cached not in (None, ''):
                    populated.append(f'{cell.coordinate}={cached!r}')
                if isinstance(cell.value, str) and cell.value.startswith('=') and (cached is not None):
                    formula_results.append(f'{cell.coordinate}={cell.value}=>{cached!r}')
            status_values = [value_sheet.cell(row[0].row, column).value for column in range(1, sheet.max_column + 1)]
            status_ok = any((contains_concept(status, token) for status in status_values if status not in (None, '') for token in ('ok', 'pass', 'cleared', 'complete', 'no exception', 'zero')))
            met = len(populated) >= 2 and bool(formula_results) and status_ok
            evidence = f'{sheet_name}: populated={populated!r}; formula_results={formula_results!r}; row_values={status_values!r}'
            if met:
                return (True, evidence)
            attempted.append(evidence)
    return (False, f'required completed control concept {aliases!r} is missing; candidates={attempted[:4]!r}')

def _task_035_formula_row_evidence(workbook, values, criterion_id: str) -> tuple[str, bool, str]:
    """Collect formula-bearing output rows without making a wording decision.

    The deterministic boundary is only that the submitted workbook has a
    calculated formula row on an authored-output surface.  Whether that row is
    the requested burn, capacity, or control concept is a semantic decision,
    so normal professional relabeling cannot be rejected by an alias parser.
    """
    preferred = {'model_content__burn_curve': ('Revenue Burn',), 'model_content__required_hours': ('Revenue Burn',), 'model_content__available_hours': ('Labor Capacity',), 'model_content__overtime': ('Labor Capacity', 'Gap Analysis'), 'model_content__subcontract': ('Labor Capacity', 'Gap Analysis'), 'model_content__capacity_gap': ('Gap Analysis',)}.get(criterion_id, ('Checks',))
    sheet_names = tuple(dict.fromkeys((*preferred, *_task_035_output_sheet_names(workbook))))
    concept = _TASK_035_MODEL_CONTROL_MEANING.get(criterion_id, criterion_id)
    concept_terms = tuple((token for token in _normalize(concept).split() if token not in {'a', 'or', 'the', 'and', 'formula', 'driven', 'completed'}))
    preferred_set = set(preferred)
    rows: list[tuple[int, int, int, str, int]] = []
    formula_result_count = 0
    sequence = 0
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            formula_cells: list[str] = []
            for cell in row:
                if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                    continue
                cached = value_sheet[cell.coordinate].value
                if cached is None:
                    continue
                formula_result_count += 1
                formula_cells.append(cell.coordinate)
            if not formula_cells:
                continue
            populated = [str(value_sheet[cell.coordinate].value) for cell in row if value_sheet[cell.coordinate].value not in (None, '')]
            row_text = ' | '.join(populated[:16])
            normalized_row = _normalize(row_text)
            relevance = sum((term in normalized_row for term in concept_terms))
            sheet_priority = 0 if sheet_name in preferred_set else 1 if sheet.sheet_state == 'visible' else 2
            rows.append((sheet_priority, -relevance, sequence, sheet_name, row[0].row))
            sequence += 1
    rows.sort()
    selected: list[str] = []
    emitted_headers: set[tuple[str, str]] = set()
    used = 0
    preferred_rows_available = sum((row[3] in preferred_set for row in rows))
    preferred_rows_selected = 0
    for (_sheet_priority, _relevance, _sequence, sheet_name, row_number) in rows:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        header = _task_035_render_header_context(sheet, row_number)
        header_key = (sheet_name, header)
        parts: list[str] = []
        if header and header_key not in emitted_headers:
            parts.append(header[:900])
        parts.append(_task_035_compact_formula_row(sheet, value_sheet, row_number))
        row_text = '\n'.join(parts)
        if selected and used + len(row_text) + 1 > 12000:
            continue
        selected.append(row_text)
        used += len(row_text) + 1
        if header:
            emitted_headers.add(header_key)
        if sheet_name in preferred_set:
            preferred_rows_selected += 1
        if len(selected) >= 36:
            break
    submitted = '\n'.join(selected) or 'No calculated formula rows were found on an output sheet.'
    completed = formula_result_count > 0
    gate_evidence = f'calculated output formula cells={formula_result_count}; submitted formula rows={len(selected)}; preferred formula rows included={preferred_rows_selected}/{preferred_rows_available}'
    return (submitted, completed, gate_evidence)
_TASK_068_SLIDES_BY_CRITERION = {'preservation__title': (1,), 'headline_values__guidance_release_status': (2, 7, 8), 'headline_values__guidance_protection_portfolio_treatment': (2, 4, 7, 8, 9), 'headline_values__largest_downside_driver': (8,), 'headline_values__guidance_update_required': (2, 7, 8), 'narrative__executive_summary': (2,), 'narrative__revenue': (3, 5), 'narrative__ebitda': (4, 5), 'narrative__cash': (6,), 'narrative__backlog': (7,), 'narrative__outlook': (7,), 'narrative__risk': (8,), 'narrative__owner': (8,), 'narrative__action': (7, 8), 'controls__numerical_tie_out': (3, 4, 5, 6, 7, 9), 'controls__revenue_bridge': (3, 5, 9), 'controls__ebitda_bridge': (4, 5, 9), 'controls__cash_roll_forward': (6, 9), 'controls__latest_forecast': (7, 8, 9)}
_TASK_068_NUMERIC_ASSOCIATION_RULES = {'q2_revenue': 'This is close-adjusted Q2 actual revenue, not approved-plan, budget, target, AOP, or another scenario. Reject the expected magnitude when it appears only in a Plan/AOP cell while the deck displays a different Q2 Actual value.', 'q2_revenue_variance_to_plan': 'This is close-adjusted Q2 actual revenue less approved-plan Q2 revenue. If the deck displays the actual and plan bases, they must support the variance. Reject a coincidentally similar rounded variance when either displayed basis is materially wrong or belongs to another scenario.', 'q2_adjusted_ebitda': 'This is close-adjusted Q2 actual adjusted EBITDA, not approved-plan, budget, target, AOP, or another scenario. Reject the expected magnitude when it appears only in a Plan/AOP cell while the deck displays a different Q2 Actual value.', 'q2_approved_plan_adjusted_ebitda': 'Judge the local Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when the surrounding schedule and source context establish the approved basis and no competing GAAP or other EBITDA basis is shown. Do not require every modifier to be repeated in the value cell or column heading.', 'construction_q2_approved_plan_adjusted_ebitda': 'For the Construction row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'service_q2_approved_plan_adjusted_ebitda': 'For the Service row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'controls_q2_approved_plan_adjusted_ebitda': 'For the Controls row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'posted_ytd_revenue': 'At a June 30 reporting date, accept ordinary cumulative-period wording such as YTD, H1, first half, or six months for revenue when the slide or cited source/control context identifies the actuals as posted accounting and no conflicting reporting basis is shown. Do not require `posted` and the period label to be repeated beside the number.', 'signed_backlog_coverage_of_remaining_outlook': 'This is signed backlog divided by the close-adjusted remaining FY26 revenue outlook. Accept normal ratio rounding when that basis is either stated correctly or not contradicted. If the deck explicitly displays a signed-backlog numerator or remaining-outlook denominator, reject a coincidentally matching rounded ratio when either displayed basis value is wrong.', 'service_labor_productivity_impact': 'This is the gross adverse Q2 labor-productivity impact derived from excess paid hours and the applicable loaded hourly cost. Evaluate it independently from the smaller supported or probability-adjusted recovery action, which may appear elsewhere in the same deck. Accept normal board rounding such as $0.29M for the verified $285,000 gross impact when the excess-hours productivity context and adverse role are clear. Reject a value presented only as mitigation or recovery.', 'latest_full_year_revenue_outlook': 'This is the current close-adjusted FY26 revenue outlook, after the controller-approved June close entry. It is distinct from the earlier June reforecast base. Reject a deck that presents the unadjusted June base as the current outlook, even if normal display rounding makes the two values pass the broad numeric-presence gate.', 'q2_posted_ytd_revenue_reconciliation_check': 'This zero check is supported only when the raw reporting-cube mapping and the controller-approved June revenue entries bridge to the current Q2 reporting basis. Accept an explicitly tied/no-plug bridge showing the raw Q2 cube allocation plus the approved close entries equals the current Q2 reporting amount; the separately scored current Q2 share need not be repeated. Do not confuse a disclosed GAAP-posted Q2 amount on a pre-June-WIP posting basis with an unresolved difference in this management-reporting bridge. Reject a zero or TIED statement based only on the earlier cube allocation, or any bridge that omits the close entries.', 'q2_operating_cash_flow_plan': 'This is the approved Board-plan Q2 operating-cash-flow component, not actual Q2 operating cash flow. The expected amount may be shown monthly or as a Q2 total, but it must be unmistakably associated with the plan scenario. Reject the expected amount when the deck presents or uses it as the actual cash-flow component.', 'q2_capital_expenditure_plan': 'This is the approved Board-plan Q2 capital-expenditure component, not actual Q2 capital expenditure. The expected amount may be shown as a cash outflow or positive spend magnitude, but it must be unmistakably associated with the plan scenario. Reject the expected amount when the deck presents or uses it as the actual cash-flow component.', 'q2_free_cash_flow_plan_derived': 'This is the mathematically derived Board-plan Q2 free cash flow: plan operating cash flow less plan capital expenditure. It is not actual Q2 free cash flow and it is distinct from the separately stated plan free-cash-flow line. A separately and correctly labeled $2.55M stated-plan line is not a contradiction to the $1.65M derived value. Accept the derived value when the plan components or an explicit derived row bind it to $3.47M less $1.82M. Reject the expected amount when it is presented or used as actual free cash flow, or merely appears as an unlabeled bridge result.', 'q2_free_cash_flow_plan_stated': 'This is the incompatible free-cash-flow line stated in the approved plan, $2.55M, distinct from the $1.65M free cash flow derived from plan operating cash flow less plan capital expenditure. Accept ordinary board rounding, including $2.6M when the deck displays one decimal place. A separately and correctly labeled $1.65M derived-plan value is not a contradiction. Reject a value presented only as actual free cash flow or as the derived plan.', 'q2_free_cash_flow_plan_discrepancy': 'This is the internal inconsistency within the approved plan records: stated plan free cash flow less the free cash flow derived from the plan operating-cash-flow and capex components. It is not the actual-versus-plan performance variance. Reject the expected magnitude when it is presented as actual free cash flow versus plan, even if the arithmetic uses the same two displayed amounts.', 'guidance_update_trigger': 'This is the formal-update threshold: the active published adjusted-EBITDA low end plus the signed minimum release buffer. The expected amount must be labeled or used as that trigger, threshold, or required headroom boundary. Reject the same number when it appears only as the upper endpoint of a recommended EBITDA range or in another unrelated role.'}

def _task_068_slide_numbers(criterion_id: str) -> tuple[int, ...]:
    if criterion_id in _TASK_068_SLIDES_BY_CRITERION:
        return _TASK_068_SLIDES_BY_CRITERION[criterion_id]
    if criterion_id.startswith('headline_values__'):
        return _task_068_value_slides(criterion_id.removeprefix('headline_values__'))
    if criterion_id.startswith('sources__'):
        return (9,)
    return tuple(range(1, 10))

def _task_068_leaf_shapes(shapes: Iterable[Any]) -> Iterable[Any]:
    """Yield authored leaf shapes, including descendants of grouped shapes."""
    for shape in shapes:
        child_shapes = getattr(shape, 'shapes', None)
        if child_shapes is not None:
            yield from _task_068_leaf_shapes(child_shapes)
        else:
            yield shape

def _task_068_shape_text_fragments(shapes: Iterable[Any]) -> tuple[str, ...]:
    """Return visible text from text frames, native tables, and groups."""
    fragments: list[str] = []
    for shape in _task_068_leaf_shapes(shapes):
        if getattr(shape, 'has_table', False):
            fragments.extend((str(cell.text).strip() for row in shape.table.rows for cell in row.cells if str(cell.text or '').strip()))
        elif hasattr(shape, 'text') and str(shape.text or '').strip():
            fragments.append(str(shape.text).strip())
    return tuple(fragments)

def _task_068_slides_changed(path: Path, artifact_relative: str, slide_numbers: Iterable[int]) -> bool:
    seed_path = SEED_WORKSPACE / artifact_relative
    if not seed_path.is_file():
        return True
    current_stat = path.stat()
    seed_stat = seed_path.stat()
    return _task_068_slides_changed_cached(str(path.resolve()), current_stat.st_mtime_ns, current_stat.st_size, str(seed_path.resolve()), seed_stat.st_mtime_ns, seed_stat.st_size, tuple(slide_numbers))

@lru_cache(maxsize=256)
def _task_068_slides_changed_cached(path_text: str, _path_mtime_ns: int, _path_size: int, seed_path_text: str, _seed_mtime_ns: int, _seed_size: int, slide_numbers: tuple[int, ...]) -> bool:
    path = Path(path_text)
    seed_path = Path(seed_path_text)
    current = Presentation(path)
    seed = Presentation(seed_path)
    for slide_number in slide_numbers:
        if slide_number > len(current.slides) or slide_number > len(seed.slides):
            return False
        current_text = _normalize('\n'.join(_task_068_shape_text_fragments(current.slides[slide_number - 1].shapes)))
        seed_text = _normalize('\n'.join(_task_068_shape_text_fragments(seed.slides[slide_number - 1].shapes)))
        if current_text != seed_text:
            return True
    return False

def _task_068_reference_context(criterion_id: str, expected_facts: Any, answer: dict[str, Any]) -> dict[str, Any]:
    answer_keys = {'headline_values__guidance_release_status': 'The current release decision is to prepare a formal guidance update because probability-adjusted operating recovery does not restore the required EBITDA headroom. Accept ordinary concise equivalents such as `prepare guidance update`, `revise guidance`, `update required`, or `do not maintain the current range`. Reject a maintain, hold, or unchanged-guidance conclusion.', 'headline_values__guidance_protection_portfolio_treatment': 'The guidance-protection portfolio is a separate contingency set and is excluded from the operating-recovery release bridge unless formally activated. Accept ordinary equivalents such as `not included`, `held in reserve`, `exclude pending activation`, or stating that the protection actions stay separately disclosed until activated with dated support. Reject treating it as current executable operating recovery.', 'headline_values__largest_downside_driver': "Identify the largest adverse probability-weighted FY26 adjusted-EBITDA effect in the signed guidance risk register. Copper escalation is the correct driver because its probability-weighted EBITDA downside is larger than Helix's. A professional risk table may separately describe Helix as the largest gross revenue exposure; that distinction is not a contradiction. Accept ordinary labels, an explicit note that identifies Copper as the largest probability-weighted EBITDA downside, a displayed probability-weighted EBITDA column with Copper highest, or a table whose adverse-risk rows are explicitly ranked by that measure. Do not award this criterion merely because an unsorted table contains gross impacts and probabilities from which a reader could calculate the answer; a table that lists Helix before Copper without a stated ranking basis does not identify the required driver. Reject Copper when it is associated only with gross revenue exposure, when another driver has a larger displayed probability-weighted EBITDA downside, or when the deck negates Copper's ranking.", 'narrative__executive_summary': 'Summarize the direction and scale of Q2 revenue and EBITDA performance, cash/liquidity, latest outlook, principal risk, and management response. The summary should communicate that the current published range requires a formal update under the probability-adjusted operating-recovery threshold; owner names, detailed actions, and dates are independently scored on slide 8 and need not be repeated on the summary slide. Do not require a scripted sentence. If the summary displays a Q2 free-cash-flow, combined-stress EBITDA, post-mitigation EBITDA, or other verified headline amount, a materially wrong displayed amount makes this criterion unmet; correct directional prose cannot rescue that contradiction.', 'narrative__revenue': 'Explain that the current Q2 revenue basis carries the controller-approved June close entries against the earlier cube extract, then compare the resulting Construction, Service, and Controls amounts with approved plan. The three branch variances must reconcile to the consolidated variance. Accept normal business-unit and favorable/unfavorable variance wording; do not require the authored labels.', 'narrative__ebitda': 'Explain that the current Q2 adjusted-EBITDA basis carries the controller-approved June close entries against the earlier cube extract, then compare the resulting Construction, Service, and Controls amounts with approved plan. The three branch-level actual-versus-plan variances must reconcile to the consolidated shortfall. Accept normal EBITDA and unfavorable-variance wording. Do not treat a consolidated operational-driver bridge or the larger diagnostic impacts for construction gross profit, service labor, and controls mix as the required business-unit EBITDA reconciliation.', 'narrative__cash': 'Explain actual Q2 operating cash flow less actual capital expenditure to free cash flow, distinguish those actuals from the approved-plan components, disclose the incompatible stated plan line, quantify the variance to derived plan, and cover unrestricted cash/revolver posture plus the collections-timing drag. Reject an earnings explanation for the collections-only cash item.', 'narrative__backlog': 'Connect signed backlog to remaining FY26 revenue outlook and its resulting coverage, without treating probability-weighted pipeline as signed backlog.', 'narrative__outlook': 'State the close-adjusted current FY26 outlook separately from both its June reforecast base and the active published revenue and EBITDA ranges, and explain the minimum-headroom threshold and formal-update posture.', 'narrative__risk': 'Identify the probability-weighted principal downside and correctly connect combined-stress EBITDA, full supported operating recovery, probability-adjusted executable recovery, and post-mitigation guidance headroom.', 'narrative__owner': 'Assign each material risk or mitigation action shown on the risk-and-action slide to an identifiable management owner; job titles or unambiguous role abbreviations are acceptable.', 'narrative__action': 'State concrete mitigation actions and their next evidence gate or decision date, distinguish full supported recovery from probability-adjusted executable recovery, keep the unactivated guidance-protection contingency separate, and make the release conclusion consistent with the verified bridge. A materially contradictory displayed action pool, recovery, post-mitigation EBITDA, headroom, or release decision makes this criterion unmet.', 'controls__numerical_tie_out': 'Show substantive controls across the deck for the raw Q2 revenue mapping plus the controller-approved close entries to the current Q2 reporting basis, adjusted EBITDA, cash/free cash flow, signed backlog, and the close-adjusted FY26 outlook. The final source/control slide may use concise subject-specific TIED/OK/OPEN rows when the corresponding calculation and value are clear on the relevant analytical slide; do not require duplicate figures on slide 9. A generic list with no identifiable subject or supporting bridge is insufficient.', 'controls__revenue_bridge': 'Show the raw cube revenue basis plus the controller-approved June revenue entries to current Q2 actual, then show that the three business-unit actual and plan amounts and variances reconcile with no unexplained difference.', 'controls__ebitda_bridge': 'Show the controller-approved June EBITDA entries in the current Q2 actual basis and show that the three business-unit actual and plan EBITDA amounts and variances reconcile with no unexplained difference.', 'controls__cash_roll_forward': 'Show actual operating cash flow less actual capital expenditure equals actual free cash flow, distinguish the approved-plan components and stated plan line, and relate the result to liquidity. Equivalent professional labels are acceptable.', 'controls__latest_forecast': 'Identify the June reforecast as the approved starting outlook and carry the later controller-approved close entry into the current FY26 outlook. The resulting current outlook need not be duplicated on slide 9 when it is clear on slide 7 and the final control row says the latest forecast ties. The unadjusted June base, AOP, or April working forecast alone is not sufficient.', 'sources__posted_accounting_records': 'Identify Vista ERP, the general ledger, or ordinary company accounting records as the authority for posted YTD revenue. Accept normal system and ledger wording; no tool-specific acronym is required.', 'sources__controller_tied': 'Identify `Q2 management reporting data book - v7 controller tie.xlsx` or an unambiguous v7 controller-tied reporting-book reference; reject the v5 CFO scratch book as controlling.', 'sources__management_correspondence': 'Identify the July 4 Q2 board/lender/IR refresh thread, the July 5 board-review/source-tie follow-up, or the Q2 board/lender narrative review notes as management authority for the outlook, risks, or actions. The structured `Finance action review notes_7.5.csv`, a generic `Finance action review notes` reference, or a `Finance operating-recovery review` table alone is operating support, not management correspondence, and does not satisfy this criterion unless the deck also identifies one of the correspondence authorities above or an unambiguous professional equivalent.', 'sources__policy': 'Identify the signed KPI definitions/lender-presentation policy or signed FY26 non-GAAP policy for the definitions or reporting treatment it governs.'}
    branch_revenue_keys = tuple((f'{branch}_q2_{suffix}' for branch in ('construction', 'service', 'controls') for suffix in ('revenue', 'approved_plan_revenue', 'revenue_variance_to_plan')))
    branch_ebitda_keys = tuple((f'{branch}_q2_{suffix}' for branch in ('construction', 'service', 'controls') for suffix in ('adjusted_ebitda', 'approved_plan_adjusted_ebitda', 'adjusted_ebitda_variance_to_plan')))
    context_keys_by_criterion = {'headline_values__largest_downside_driver': ('largest_downside_driver',), 'narrative__executive_summary': ('q2_revenue', 'q2_revenue_variance_to_plan', 'q2_adjusted_ebitda', 'q2_adjusted_ebitda_variance_to_plan', 'q2_free_cash_flow', 'q2_unrestricted_cash', 'latest_full_year_revenue_outlook', 'combined_stress_ebitda', 'post_mitigation_combined_stress_ebitda', 'guidance_release_status'), 'narrative__revenue': ('posted_ytd_revenue', 'q2_posted_ytd_revenue_share', 'q2_posted_ytd_revenue_reconciliation_check', 'q2_revenue', 'q2_approved_plan_revenue', 'q2_revenue_variance_to_plan', 'q2_raw_cube_revenue', 'q2_close_revenue_adjustment', *branch_revenue_keys), 'narrative__ebitda': ('q2_adjusted_ebitda', 'q2_approved_plan_adjusted_ebitda', 'q2_adjusted_ebitda_variance_to_plan', 'q2_close_ebitda_adjustment', *branch_ebitda_keys), 'narrative__cash': ('q2_operating_cash_flow', 'q2_capital_expenditure', 'q2_free_cash_flow', 'q2_operating_cash_flow_plan', 'q2_capital_expenditure_plan', 'q2_free_cash_flow_plan_derived', 'q2_free_cash_flow_plan_stated', 'q2_free_cash_flow_plan_discrepancy', 'q2_free_cash_flow_variance_to_derived_plan', 'q2_unrestricted_cash', 'collections_timing_cash_impact', 'maximum_revolver'), 'narrative__backlog': ('signed_backlog', 'remaining_fy26_revenue_outlook', 'signed_backlog_coverage_of_remaining_outlook', 'latest_full_year_revenue_outlook'), 'narrative__outlook': ('latest_full_year_revenue_outlook', 'revenue_guidance_low', 'revenue_guidance_high', 'ebitda_guidance_low', 'ebitda_guidance_high', 'minimum_guidance_update_headroom', 'guidance_update_trigger', 'guidance_action_gap_to_required_headroom', 'full_year_outlook_pre_close_adjustment', 'full_year_outlook_close_adjustment', 'guidance_update_required', 'guidance_release_status'), 'narrative__risk': ('largest_downside_driver', 'combined_stress_ebitda', 'ebitda_shortfall_to_guidance_low', 'identified_ebitda_action_pool', 'full_supported_mitigation_conversion', 'full_supported_ebitda_mitigation', 'probability_adjusted_mitigation_conversion', 'probability_adjusted_executable_ebitda_mitigation', 'post_mitigation_combined_stress_ebitda', 'post_mitigation_headroom_to_guidance_low'), 'narrative__owner': (), 'narrative__action': ('identified_ebitda_action_pool', 'full_supported_mitigation_conversion', 'full_supported_ebitda_mitigation', 'probability_adjusted_mitigation_conversion', 'probability_adjusted_executable_ebitda_mitigation', 'combined_stress_ebitda', 'post_mitigation_combined_stress_ebitda', 'post_mitigation_headroom_to_guidance_low', 'minimum_guidance_update_headroom', 'guidance_action_gap_to_required_headroom', 'guidance_protection_portfolio_treatment', 'guidance_release_status'), 'controls__numerical_tie_out': ('posted_ytd_revenue', 'q2_posted_ytd_revenue_share', 'q2_posted_ytd_revenue_reconciliation_check', 'q2_revenue', 'q2_raw_cube_revenue', 'q2_close_revenue_adjustment', 'q2_adjusted_ebitda', 'q2_close_ebitda_adjustment', 'q2_free_cash_flow', 'signed_backlog', 'latest_full_year_revenue_outlook'), 'controls__revenue_bridge': ('q2_revenue', 'q2_approved_plan_revenue', 'q2_revenue_variance_to_plan', 'q2_raw_cube_revenue', 'q2_close_revenue_adjustment', *branch_revenue_keys), 'controls__ebitda_bridge': ('q2_adjusted_ebitda', 'q2_approved_plan_adjusted_ebitda', 'q2_adjusted_ebitda_variance_to_plan', 'q2_close_ebitda_adjustment', *branch_ebitda_keys), 'controls__cash_roll_forward': ('q2_operating_cash_flow', 'q2_capital_expenditure', 'q2_free_cash_flow', 'q2_operating_cash_flow_plan', 'q2_capital_expenditure_plan', 'q2_free_cash_flow_plan_derived', 'q2_free_cash_flow_plan_stated', 'q2_free_cash_flow_plan_discrepancy', 'q2_free_cash_flow_variance_to_derived_plan', 'q2_unrestricted_cash'), 'controls__latest_forecast': ('latest_full_year_revenue_outlook', 'revenue_guidance_low', 'revenue_guidance_high', 'ebitda_guidance_low', 'ebitda_guidance_high', 'full_year_outlook_pre_close_adjustment', 'full_year_outlook_close_adjustment')}
    context_keys = context_keys_by_criterion.get(criterion_id, ())
    return {'criterion_answer_key': answer_keys.get(criterion_id, expected_facts), 'expected_facts': expected_facts, 'equivalence_rule': 'Accept ordinary professional wording and equivalent labels; do not require the authored phrase.', 'verified_finance_context': {key: answer[key] for key in context_keys if key in answer and answer[key] is not None}}
_TASK_068_NUMERIC_ASSOCIATION_RULES.update({'q2_approved_plan_adjusted_ebitda': 'Judge the local Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when the surrounding schedule and source context establish the approved basis and no competing GAAP or other EBITDA basis is shown. Do not require every modifier to be repeated in the value cell or column heading.', 'construction_q2_approved_plan_adjusted_ebitda': 'For the Construction row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'service_q2_approved_plan_adjusted_ebitda': 'For the Service row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'controls_q2_approved_plan_adjusted_ebitda': 'For the Controls row, judge the Q2 actual-versus-plan schedule together with its cited planning authority. Accept any concise professional plan, budget, target, or AOP heading for adjusted EBITDA when context establishes the approved basis and no competing EBITDA basis is shown.', 'posted_ytd_revenue': 'At a June 30 reporting date, accept ordinary cumulative-period wording such as YTD, H1, first half, or six months for revenue when the slide or cited source/control context identifies the actuals as posted accounting and no conflicting reporting basis is shown. Do not require `posted` and the period label to be repeated beside the number.'})

def _workbook_value_present(values, expected: Any, *, formula_workbook=None) -> bool:
    candidates: list[Any]
    cached_candidates = getattr(values, '_alder_all_value_candidates', None) if formula_workbook is None else None
    if cached_candidates is not None:
        candidates = cached_candidates
    else:
        candidates = []
        for sheet_name in values.sheetnames:
            value_sheet = values[sheet_name]
            formula_sheet = formula_workbook[sheet_name] if formula_workbook is not None else None
            for row in value_sheet.iter_rows():
                for cell in row:
                    if formula_sheet is not None:
                        formula = formula_sheet[cell.coordinate].value
                        if not (isinstance(formula, str) and formula.startswith('=')):
                            continue
                    candidates.append(cell.value)
        if formula_workbook is None:
            setattr(values, '_alder_all_value_candidates', candidates)
    if isinstance(expected, bool):
        return any((_boolean_matches(candidate, expected) for candidate in candidates))
    if isinstance(expected, str):
        return any((date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates))
    if isinstance(expected, list):
        return ordered_semantic_list_matches(candidates, expected) or unordered_semantic_list_matches(candidates, expected)
    target = float(expected)
    targets = [target]
    if abs(target) > 10:
        targets.append(target / 1000.0)
    if abs(target) >= 100000:
        unit_text = getattr(values, '_alder_unit_text_cache', None)
        if unit_text is None:
            unit_text = '\n'.join((str(candidate) for candidate in candidates if candidate is not None)).casefold()
            setattr(values, '_alder_unit_text_cache', unit_text)
        if any((marker in unit_text for marker in ('usd millions', '$ in millions', '$ millions', '$mm', '$ mm'))):
            targets.append(target / 1000000.0)
    return any((isinstance(candidate, (int, float)) and (not isinstance(candidate, bool)) and any((_close(candidate, candidate_target, abs_tol=max(2e-05, abs(candidate_target) * 1e-06), rel_tol=0.0) for candidate_target in targets)) for candidate in candidates))

def _workbook_numeric_targets(values, expected: float) -> list[float]:
    """Return only unit transformations explicitly disclosed by the workbook."""
    target = float(expected)
    targets = [target]
    if abs(target) > 10:
        targets.append(target / 1000.0)
    unit_text = getattr(values, '_alder_unit_text_cache', None)
    if unit_text is None:
        unit_text = '\n'.join((str(cell.value) for sheet in values.worksheets for row in sheet.iter_rows() for cell in row if cell.value is not None)).casefold()
        setattr(values, '_alder_unit_text_cache', unit_text)
    if abs(target) >= 100000 and any((marker in unit_text for marker in ('usd millions', '$ in millions', '$ millions', '$mm', '$ mm', 'all dollars in millions', 'dollars in millions'))):
        targets.append(target / 1000000.0)
    return list(dict.fromkeys(targets))

def _workbook_semantic_row_index(workbook, values):
    """Build the reusable row/label index once for all semantic hard gates."""
    cache = getattr(workbook, '_alder_semantic_row_index', None)
    if cache is not None and cache[0] == id(values):
        return (cache[1], cache[2])
    rows: list[dict[str, Any]] = []
    label_index: dict[str, list[dict[str, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            numbers: list[float] = []
            labels: list[str] = []
            formulas: list[tuple[str, str, Any]] = []
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                if isinstance(cached, (int, float)) and (not isinstance(cached, bool)):
                    numbers.append(float(cached))
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    formulas.append((cell.coordinate, cell.value, cached))
                elif cell.value is not None:
                    normalized = _normalize(cell.value)
                    if normalized:
                        labels.append(normalized)
            if not numbers and (not labels) and (not formulas):
                continue
            indexed_row = {'sheet': sheet_name, 'row': row[0].row, 'numbers': tuple(numbers), 'labels': tuple(labels), 'formulas': tuple(formulas)}
            rows.append(indexed_row)
            for normalized in set(labels):
                label_index.setdefault(normalized, []).append(indexed_row)
    setattr(workbook, '_alder_semantic_row_index', (id(values), rows, label_index))
    return (rows, label_index)

def _explicit_label_value_conflict(workbook, values, label: str, expected: Any) -> str | None:
    """Reject a directly labeled numeric contradiction without enforcing aliases.

    Semantic review remains responsible for professional-equivalent labels.
    This gate is deliberately narrower: when the artifact itself uses the exact
    requested metric label, that row may not display a different number while
    the expected number happens to occur under another metric elsewhere.
    """
    if isinstance(expected, (bool, str, list)):
        return None
    wanted = _normalize(label.replace('_', ' '))
    targets = _workbook_numeric_targets(values, float(expected))
    (rows, label_index) = _workbook_semantic_row_index(workbook, values)
    labeled_rows: list[tuple[str, int, tuple[float, ...]]] = []
    wanted_tokens = set(wanted.split())

    def row_matches_expected(numbers) -> bool:
        return any((_close(number, target, abs_tol=max(2e-05, abs(target) * 1e-06), rel_tol=0.0) for number in numbers for target in targets))
    for (row_label, indexed_rows) in label_index.items():
        if not wanted_tokens <= set(row_label.split()):
            continue
        if any((row_matches_expected(row['numbers']) for row in indexed_rows)):
            return None
    for row in label_index.get(wanted, []):
        if row['numbers']:
            labeled_rows.append((row['sheet'], row['row'], row['numbers']))
    if not labeled_rows:
        return None
    for (_sheet_name, _row_number, numbers) in labeled_rows:
        if row_matches_expected(numbers):
            return None
    rendered = [f'{sheet_name}!{row_number}={numbers!r}' for (sheet_name, row_number, numbers) in labeled_rows]
    return f'exact metric label {label!r} displays conflicting numeric row(s) {rendered!r}; expected one of {targets!r}'

def _extreme_headline_association_present(workbook, values, label: str, expected: Any) -> tuple[bool, str] | None:
    """Require an explicit max/min association when the requested fact is an extreme."""
    if isinstance(expected, (bool, str, list)):
        return None
    normalized_label = _normalize(label.replace('_', ' '))
    if any((word in normalized_label.split() for word in ('maximum', 'max', 'peak', 'highest'))):
        direction = 'maximum'
        terms = ('maximum', 'max', 'peak', 'highest')
        formula_function = 'MAX'
    elif any((word in normalized_label.split() for word in ('minimum', 'min', 'lowest'))):
        direction = 'minimum'
        terms = ('minimum', 'min', 'lowest')
        formula_function = 'MIN'
    else:
        return None
    targets = _workbook_numeric_targets(values, float(expected))
    (rows, _label_index) = _workbook_semantic_row_index(workbook, values)

    def matches(value: Any) -> bool:
        return isinstance(value, (int, float)) and (not isinstance(value, bool)) and any((_close(value, target, abs_tol=max(2e-05, abs(target) * 1e-06), rel_tol=0.0) for target in targets))
    for row in rows:
        if not any((matches(number) for number in row['numbers'])):
            continue
        row_text = ' '.join(row['labels'])
        if any((re.search(f'\\b{re.escape(term)}\\b', row_text) for term in terms)):
            return (True, f"{direction} is explicitly labeled on {row['sheet']}!{row['row']}")
        for (coordinate, formula, cached) in row['formulas']:
            if matches(cached) and re.search(f'\\b{formula_function}\\s*\\(', formula, flags=re.I):
                return (True, f"{direction} is formula-identified at {row['sheet']}!{coordinate}={formula}")
    return (False, f'expected fact is present but is not explicitly identified as a {direction} by a professional label, selected-extreme row, or MAX/MIN formula')
_FORMULA_REFERENCE_PATTERN = re.compile("(?:(?:'(?P<quoted>[^']+)'|(?P<plain>[A-Za-z_][A-Za-z0-9_. -]*))!)?\\$?[A-Z]{1,3}\\$?(?P<row_start>\\d+)(?::\\$?[A-Z]{1,3}\\$?(?P<row_end>\\d+))?")

def _scenario_marker(text: str) -> str | None:
    normalized = _normalize(text)
    markers: set[str] = set()
    if any((token in normalized for token in ('downside', 'severe downside', 'stress'))):
        markers.add('downside')
    if 'upside' in normalized:
        markers.add('upside')
    if re.search('\\bbase(?: case)?\\b', normalized):
        markers.add('base')
    return next(iter(markers)) if len(markers) == 1 else None

def _scenario_aggregate_formula_miswires(workbook) -> list[str]:
    """Find scenario summaries aggregating a visibly different scenario row.

    This intentionally does not reject ordinary downside formulas that use a
    base-case input.  It targets the much narrower, auditable defect where a
    scenario-labeled summary applies MAX/MIN/SUM/AVERAGE to a row explicitly
    labeled as another scenario, such as a downside maximum pointing at an
    upside free-cash-flow row.
    """
    cached = getattr(workbook, '_alder_scenario_miswire_cache', None)
    if cached is not None:
        return list(cached)
    comparison_terms = ('variance', 'difference', 'delta', ' versus ', ' vs ', 'comparison', 'bridge', 'reconciliation', 'sensitivity')
    findings: list[str] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            target_text = ' | '.join((str(cell.value) for cell in row if cell.value is not None and (not (isinstance(cell.value, str) and cell.value.startswith('=')))))
            target_scenario = _scenario_marker(target_text)
            normalized_target = f' {_normalize(target_text)} '
            if target_scenario is None or any((term in normalized_target for term in comparison_terms)):
                continue
            for cell in row:
                formula = cell.value
                if not (isinstance(formula, str) and formula.startswith('=') and re.search('\\b(?:MAX|MIN|SUM|AVERAGE)\\s*\\(', formula, flags=re.I)):
                    continue
                for match in _FORMULA_REFERENCE_PATTERN.finditer(formula):
                    referenced_sheet_name = match.group('quoted') or match.group('plain') or sheet_name
                    if referenced_sheet_name not in workbook.sheetnames:
                        continue
                    referenced_sheet = workbook[referenced_sheet_name]
                    row_start = int(match.group('row_start'))
                    row_end = int(match.group('row_end') or row_start)
                    if row_end - row_start > 250:
                        continue
                    for referenced_row_number in range(row_start, row_end + 1):
                        source_text = ' | '.join((str(source_cell.value) for source_cell in referenced_sheet[referenced_row_number] if source_cell.value is not None and (not (isinstance(source_cell.value, str) and source_cell.value.startswith('=')))))
                        source_scenario = _scenario_marker(source_text)
                        if source_scenario is None or source_scenario == target_scenario:
                            continue
                        findings.append(f'{sheet_name}!{cell.coordinate} {target_scenario!r} summary uses {referenced_sheet_name}!{referenced_row_number} labeled {source_scenario!r}: {formula}')
    setattr(workbook, '_alder_scenario_miswire_cache', tuple(findings))
    return findings

def _numeric_literal_is_compatible_with_label(label: str | None, literal: str) -> bool:
    """Keep a displayed unit from satisfying an unrelated metric type.

    In particular, a rounded monetary literal such as ``$0.00m`` has a wide
    dollar display tolerance.  It must never satisfy a percentage criterion
    merely because the expected decimal happens to fall inside that dollar
    rounding band.
    """
    if not label:
        return True
    label_tokens = set(_normalize(label).split())
    if label_tokens & {'pct', 'percent', 'percentage'}:
        return '%' in literal
    return True

def _artifact_expected_fact_present(path: Path, expected: Any, *, label: str | None=None) -> bool:
    """Require objective displayed facts without requiring a brittle label.

    Numeric facts are deterministic hard gates.  Categorical conclusions and
    list membership stay with the bounded semantic judge because their visible
    expression necessarily depends on wording, negation, and table context.
    """
    if isinstance(expected, (str, bool, list)):
        return True
    text = _artifact_text(path)
    return any((_numeric_literal_is_compatible_with_label(label, literal) and _numeric_literal_matches(literal, float(expected)) for literal in _NUMERIC_LITERAL_PATTERN.findall(text)))

def _workbook_error_values(values) -> list[str]:
    """Return recalculated spreadsheet errors with stable cell evidence."""
    cached = getattr(values, '_alder_error_values_cache', None)
    if cached is not None:
        return list(cached)
    errors = [f'{sheet.title}!{cell.coordinate}={cell.value}' for sheet in values.worksheets for row in sheet.iter_rows() for cell in row if cell.data_type == 'e' or (isinstance(cell.value, str) and re.fullmatch('#(?:REF!|DIV/0!|VALUE!|NAME\\?|N/A|NUM!|NULL!)', cell.value))]
    setattr(values, '_alder_error_values_cache', tuple(errors))
    return errors

def _hybrid_review_hard_gate(task_id: str, spec: dict[str, Any], gold: dict[str, Any], path: Path, workbook, values) -> tuple[bool, str]:
    kind = spec['kind']
    suffix = path.suffix.lower()
    if suffix == '.xlsx':
        if spec['id'].endswith('__model_status'):
            errors = _workbook_error_values(values)
            if errors:
                return (False, 'recalculated spreadsheet errors=' + repr(errors[:8]))
            miswires = _scenario_aggregate_formula_miswires(workbook)
            if miswires:
                return (False, 'visible cross-scenario aggregate formula miswire(s): ' + '; '.join(miswires[:8]))
        if kind == 'xlsx_label_values':
            if task_id == 'task_035':
                (label, expected) = next(iter(spec['label_values'].items()))
                (_evidence, populated, gate_evidence) = _task_035_semantic_fact_evidence(workbook, values, label, expected, require_formula=False)
                return (populated, gate_evidence)
            checkable = {label: expected for (label, expected) in spec['label_values'].items() if isinstance(expected, (int, float)) and (not isinstance(expected, bool))}
            missing = [label for (label, expected) in checkable.items() if not _workbook_value_present(values, expected)]
            if missing:
                return (False, f'required numeric facts missing anywhere in workbook={missing!r}')
            extreme_associations = [outcome for (label, expected) in spec['label_values'].items() if (outcome := _extreme_headline_association_present(workbook, values, label, expected)) is not None]
            failed_extremes = [evidence for (met, evidence) in extreme_associations if not met]
            if failed_extremes:
                return (False, '; '.join(failed_extremes))
            conflicts = [conflict for (label, expected) in spec['label_values'].items() if (conflict := _explicit_label_value_conflict(workbook, values, label, expected)) is not None]
            return (not conflicts, 'required numeric facts are present and no exact-label contradiction exists' if not conflicts else '; '.join(conflicts))
        if kind == 'xlsx_headline_formula' and task_id == 'task_035':
            label = str(spec['headline_label'])
            expected = gold['answer'][label]
            (_evidence, formula_fact_present, gate_evidence) = _task_035_semantic_fact_evidence(workbook, values, label, expected, require_formula=True)
            return (formula_fact_present, gate_evidence)
        if kind == 'artifact_tokens':
            if task_id == 'task_035':
                if spec['id'].startswith(('model_content__', 'controls__')):
                    (_submitted, completed, gate_evidence) = _task_035_formula_row_evidence(workbook, values, spec['id'])
                    return (completed, gate_evidence)
                token = spec.get('tokens', [''])[0]
                source_reference = _TASK_035_SOURCE_REFERENCES.get(token, {})
                (_evidence, completed, gate_evidence) = _scoped_xlsx_token_evidence(workbook, values, token, aliases=source_reference.get('aliases', ()), sheet_names=_task_035_output_sheet_names(workbook))
                return (completed, gate_evidence)
            return (True, 'workbook parsed; criterion is intentionally semantic')
        if kind == 'xlsx_formula_lineage':
            by_sheet = {sheet.title: sum((1 for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('='))) for sheet in workbook.worksheets}
            formulas = [cell.value for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')]
            cross_sheet = sum((1 for formula in formulas if re.search("(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\\$?[A-Z]{1,3}\\$?\\d+", formula)))
            missing_sheets = [name for name in spec['formula_sheets'] if by_sheet.get(name, 0) == 0]
            met = not missing_sheets and cross_sheet >= int(spec['min_cross_sheet_formulas'])
            return (met, f"cross_sheet={cross_sheet}/{spec['min_cross_sheet_formulas']}; missing_formula_sheets={missing_sheets!r}. Headline facts are enforced by their separate exact-value gate; the semantic judge verifies professional association to the formula-driven schedules.")
    if suffix == '.docx':
        if kind == 'docx_heading':
            return (True, 'document parsed; substantive heading equivalence is semantic')
        if kind == 'docx_structure':
            tables = len(Document(path).tables)
            return (tables >= int(spec['min_tables']), f"tables={tables}; required={spec['min_tables']}")
        if kind == 'artifact_label_values':
            missing = [label for (label, expected) in spec['label_values'].items() if not _artifact_expected_fact_present(path, expected, label=label)]
            return (not missing, f'exact numeric facts missing anywhere in artifact={missing!r}')
        if kind == 'artifact_tokens':
            return (True, 'artifact parsed; criterion is intentionally semantic')
    if suffix == '.pptx':
        if task_id == 'task_068':
            slide_numbers = _task_068_slide_numbers(spec['id'])
            (_evidence, substantive, evidence) = _pptx_slide_evidence(path, slide_numbers)
            changed = _task_068_slides_changed(path, gold['artifact']['path'], slide_numbers)
            if spec['id'] == 'preservation__title':
                return (substantive, evidence)
            if spec['id'] == 'narrative__backlog':
                required_labels = ('signed_backlog', 'remaining_fy26_revenue_outlook', 'signed_backlog_coverage_of_remaining_outlook')
                missing: list[str] = []
                fact_evidence: list[str] = []
                for label in required_labels:
                    (present, detail) = _task_068_numeric_fact_present(path, label, float(gold['answer'][label]), slide_numbers)
                    if not present:
                        missing.append(label)
                    fact_evidence.append(f'{label}: {detail}')
                return (substantive and changed and (not missing), f'{evidence}; relevant slide content changed from starter={changed}; missing required backlog/outlook magnitudes={missing!r}; ' + '; '.join(fact_evidence))
            if kind == 'artifact_label_values':
                (label, expected) = next(iter(spec['label_values'].items()))
                if isinstance(expected, (int, float)) and (not isinstance(expected, bool)):
                    if label == 'q2_posted_ytd_revenue_reconciliation_check':
                        (components_present, components_evidence) = _task_068_current_q2_reconciliation_components_present(path, gold['answer'], slide_numbers)
                        return (substantive and changed and components_present, f'{evidence}; relevant slide content changed from starter={changed}; {components_evidence}')
                    (primary_status, primary_evidence) = _task_068_unchanged_template_numeric_status(path, gold['artifact']['path'], label, float(expected))
                    if primary_status == 'conflict':
                        return (False, f'an unchanged canonical primary field contains a conflicting value: {primary_evidence}')
                    (branch_status, branch_evidence) = _task_068_branch_primary_value_status(path, label, float(expected))
                    if branch_status == 'conflict':
                        return (False, branch_evidence)
                    variance_bases = {'q2_revenue_variance_to_plan': ('q2_revenue', 'q2_approved_plan_revenue'), 'q2_adjusted_ebitda_variance_to_plan': ('q2_adjusted_ebitda', 'q2_approved_plan_adjusted_ebitda')}.get(label, ())
                    conflicting_bases: list[str] = []
                    basis_evidence: list[str] = []
                    for basis_label in variance_bases:
                        (status, detail) = _task_068_unchanged_template_numeric_status(path, gold['artifact']['path'], basis_label, float(gold['answer'][basis_label]))
                        basis_evidence.append(f'{basis_label}={status}: {detail}')
                        if status == 'conflict':
                            conflicting_bases.append(basis_label)
                    if conflicting_bases:
                        return (False, f'the displayed actual/plan basis conflicts with the required variance: conflicting bases={conflicting_bases!r}; ' + '; '.join(basis_evidence))
                    if label == 'signed_backlog_coverage_of_remaining_outlook':
                        (fact_present, fact_evidence) = _task_068_numeric_fact_present(path, label, float(expected), slide_numbers)
                        basis_conflicts: list[str] = []
                        basis_evidence: list[str] = []
                        for basis_label in ('signed_backlog', 'remaining_fy26_revenue_outlook'):
                            (status, detail) = _task_068_unchanged_template_numeric_status(path, gold['artifact']['path'], basis_label, float(gold['answer'][basis_label]))
                            basis_evidence.append(f'{basis_label}={status}: {detail}')
                            if status == 'conflict':
                                basis_conflicts.append(basis_label)
                        return (fact_present and changed and (not basis_conflicts), f'{fact_evidence}; relevant slide content changed from starter={changed}; conflicting displayed coverage basis={basis_conflicts!r}; ' + '; '.join(basis_evidence))
                    (fact_present, fact_evidence) = _task_068_numeric_fact_present(path, label, float(expected), slide_numbers)
                    return (fact_present and changed, f'{fact_evidence}; relevant slide content changed from starter={changed}')
            return (substantive and changed, f'{evidence}; relevant slide content changed from starter={changed}')
        if kind == 'pptx_title':
            return (True, 'presentation parsed; substantive title equivalence is semantic')
        if kind == 'pptx_structure':
            presentation = Presentation(path)
            expected_slides = spec.get('exact_slides')
            met = len(presentation.slides) == int(expected_slides) if expected_slides is not None else len(presentation.slides) >= int(spec['min_slides'])
            return (met, f"slides={len(presentation.slides)}; required={expected_slides or spec['min_slides']}")
        if kind == 'artifact_label_values':
            missing = [label for (label, expected) in spec['label_values'].items() if not _artifact_expected_fact_present(path, expected, label=label)]
            return (not missing, f'exact displayed numeric facts missing anywhere in deck={missing!r}')
        if kind == 'artifact_tokens':
            return (True, 'deck parsed; criterion is intentionally semantic')
    return (False, 'criterion has no benchmark-wide hybrid hard gate')

def _hybrid_semantic_review(task_id: str, gold: dict[str, Any], path: Path, workbook, values, criteria: list[Criterion]) -> dict[str, Any] | None:
    by_id = {criterion.id: criterion for criterion in criteria}
    reviews: list[dict[str, Any]] = []
    for spec in gold['criteria']:
        if not spec.get('semantic'):
            continue
        canonical_numeric_fast_path = False
        (hard_gate_met, hard_gate_evidence) = _hybrid_review_hard_gate(task_id, spec, gold, path, workbook, values)
        expected_facts = spec.get('label_values')
        if expected_facts is None and spec.get('tokens'):
            expected_facts = {'required_concept': spec['tokens'][0]}
        if expected_facts is None and spec.get('heading'):
            expected_facts = {'required_section': spec['heading']}
        if expected_facts is None and spec.get('title_token'):
            expected_facts = {'required_title': spec['title_token']}
        requirement = semantic_requirement(criterion_id=spec['id'], description=spec['description'], expected_facts=expected_facts, artifact_type={'.xlsx': 'workbook', '.docx': 'memorandum', '.pptx': 'board presentation'}.get(path.suffix.lower(), 'artifact'))
        if task_id == 'task_035':
            if spec['kind'] in {'xlsx_label_values', 'xlsx_headline_formula'}:
                if spec['kind'] == 'xlsx_label_values':
                    (label, expected) = next(iter(spec['label_values'].items()))
                    require_formula = False
                else:
                    label = str(spec['headline_label'])
                    expected = gold['answer'][label]
                    require_formula = True
                (submitted_evidence, _populated, _gate) = _task_035_semantic_fact_evidence(workbook, values, label, expected, require_formula=require_formula)
                exact_text = label in _TASK_035_EXACT_TEXT_LABELS
                numeric = isinstance(expected, (int, float)) and (not isinstance(expected, bool))
                if numeric or exact_text:
                    reference_context = {'metric': label.replace('_', ' '), 'expected_value': expected, 'formula_required': require_formula, 'equivalence_rule': 'Accept any unambiguous professional label, abbreviation, table layout, or locally disclosed unit that associates the deterministically verified fact with this metric. Do not require an authored label or finite alias match.', 'association_rule': 'The verified fact must belong to this metric, period, and scenario. Reject an unlabeled value, a value belonging to a different metric, or a conflicting primary value even if the expected fact appears elsewhere in the output sheets.', 'sign_rule': 'For numeric facts, accept an accounting negative, adverse magnitude, or positive-balance convention only when the surrounding label preserves the same business meaning.'}
                    normalized_label = _normalize(label)
                    probability_scenario = 'probability plan' in normalized_label
                    gross_scenario = 'gross commitment' in normalized_label
                    if probability_scenario:
                        reference_context['probability_plan_definition'] = 'The probability plan is the execution-probability-adjusted, risk-adjusted signed-backlog view. A visibly source-linked risk-adjusted or execution-probability schedule is sufficient; the literal phrase probability plan is not required.'
                    if gross_scenario:
                        reference_context['gross_commitment_definition'] = 'The gross-commitment case is 100% of signed customer backlog without an execution-probability reduction. It is distinct from the risk-adjusted probability plan.'
                    execution_revenue_check = label == 'execution_portfolio_revenue_check_delta'
                    if execution_revenue_check:
                        reference_context['execution_portfolio_revenue_check_definition'] = 'This control reconciles execution-portfolio completed revenue plus execution-portfolio ending deferred revenue to gross signed backlog or contract value. A generic tie between probability-weighted or risk-adjusted revenue and its planning source is a different control and is insufficient even when that different check also equals zero.'
                    requirement = f"Evaluate only whether the submitted workbook clearly associates the deterministically verified fact with {label.replace('_', ' ')} for the correct period, scenario, and business meaning. Accept normal professional wording, abbreviations, layout, locally disclosed units, and valid sign conventions. Reject a right value attached to the wrong metric, an unlabeled number or identifier, an ambiguous binding, or a contradictory displayed primary value."
                    if require_formula:
                        requirement += ' The exact fact must be the calculated result of the source-linked formula identified by the deterministic hard gate.'
                    if probability_scenario:
                        requirement += ' Treat a visibly linked risk-adjusted signed-backlog or execution-probability schedule as an ordinary professional expression of the probability plan; do not require those literal words in the output label.'
                    if gross_scenario:
                        requirement += ' Treat only a full signed-backlog view without probability reduction as the gross-commitment case; do not confuse it with the risk-adjusted plan.'
                    if execution_revenue_check:
                        requirement += ' This exact zero must be the execution-portfolio revenue bridge: completed revenue plus ending deferred revenue reconciled to gross signed backlog or contract value. Reject a zero from a probability-plan or risk-adjusted revenue-source tie because that is a different control.'
                else:
                    reference_context = _task_035_status_reference(label, expected)
                    reference_context['formula_required'] = require_formula
                    if label == 'controlling_capacity_case':
                        reference_context['probability_plan_definition'] = 'The probability plan is the execution-probability-adjusted, risk-adjusted signed-backlog planning view.'
                        reference_context['gross_commitment_definition'] = 'The gross-commitment case is 100% of signed customer backlog without execution-probability reduction.'
                    requirement = f"Evaluate only whether the submitted value for {label!r} expresses the correct operational decision under the supplied answer key. Accept an unambiguous professional equivalent. Apply the answer key's explicit grading boundary: do not demand that a monthly or project status repeat an executive action scored in another criterion, but reject an opposite or ambiguous released-versus-held decision."
                    if require_formula:
                        requirement += ' The decision must be the calculated result of a source-linked workbook formula; a static label or unrelated formula is insufficient.'
                    if label == 'controlling_capacity_case':
                        requirement += ' Select gross commitment only when the output establishes the full 100% signed-backlog exposure as binding. A risk-adjusted or execution-probability plan alone does not establish that controlling case.'
                evidence_scope = f'authored output rows containing exact fact or decision candidates for {label!r}; vetted aliases are evidence hints only'
            elif spec['id'].startswith(('model_content__', 'controls__')):
                (submitted_evidence, _completed, _gate) = _task_035_formula_row_evidence(workbook, values, spec['id'])
                required_meaning = _TASK_035_MODEL_CONTROL_MEANING[spec['id']]
                reference_context = {'required_model_or_control_capability': required_meaning, 'equivalence_rule': 'Accept any ordinary professional label or layout that clearly performs this capability. Do not require the authored label.', 'completion_rule': 'The capability must be populated and formula-driven, with a calculated result. A heading, blank starter row, or uncalculated formula is insufficient.'}
                evidence_scope = 'calculated formula-bearing rows on authored output sheets'
                requirement = f'Evaluate only whether the submitted calculated rows perform this model/control capability: {required_meaning}. Accept equivalent business wording and a different professional layout. Reject a merely named heading, an unrelated formula, or an incomplete control.'
            else:
                token = spec.get('tokens', [''])[0]
                source_reference = _TASK_035_SOURCE_REFERENCES.get(token, {})
                (submitted_evidence, _populated, _gate) = _scoped_xlsx_token_evidence(workbook, values, token, aliases=source_reference.get('aliases', ()), sheet_names=_task_035_output_sheet_names(workbook))
                reference_context = {'required_source_or_control': token, 'criterion_answer_key': source_reference.get('answer_key', spec['description']), 'accepted_authority_references': list(source_reference.get('aliases', ())), 'equivalence_rule': 'Accept a clearly identified equivalent source/control; do not require the authored phrase.'}
                evidence_scope = f'workbook rows containing the {token!r} source/control concept'
                requirement = f'Evaluate only whether this submitted source row identifies an authoritative source satisfying the supplied answer key for {token!r}. Accept a clear filename, subject, abbreviation, or ordinary professional equivalent; reject an unidentified generic claim or the wrong/superseded authority.'
            task_context = {'assignment': 'Complete the FY27 signed-backlog burn, capacity, remediation, execution, earnings-release, and executive-recovery model.', 'sign_convention': 'capacity gap = required hours minus available hours; positive means shortage, negative means spare capacity; unresolved shortfall is floored at zero', 'status_conventions': {'cleared': 'no unresolved shortfall remains after permitted remedies', 'sequencing_or_resequencing_required': 'a residual shortfall or deferred portfolio remains', 'executive_decision_required': 'the project or portfolio cannot be released without management action'}, 'grading_boundary': 'Exact magnitudes and exact date/project occurrences remain deterministic hard gates, as does formula-result presence where required. This judge decides their metric/period/scenario association and all editable status, source, or formula-row business meaning under the individual criterion boundary.'}
        elif task_id == 'task_068':
            slide_numbers = _task_068_slide_numbers(spec['id'])
            (submitted_evidence, _substantive, _gate) = _pptx_slide_evidence(path, slide_numbers)
            numeric_association = False
            numeric_label = ''
            numeric_expected: float | None = None
            if spec['kind'] == 'artifact_label_values':
                (numeric_label, candidate) = next(iter(spec['label_values'].items()))
                if isinstance(candidate, (int, float)) and (not isinstance(candidate, bool)):
                    numeric_association = True
                    numeric_expected = float(candidate)
            if numeric_association:
                canonical_numeric_fast_path = _task_068_safe_canonical_numeric_fast_path(path, gold['artifact']['path'], numeric_label, float(numeric_expected))
                exact_candidate_contexts = _task_068_exact_value_candidate_contexts(path, numeric_label, float(numeric_expected), slide_numbers) if float(numeric_expected) != 0.0 else ()
                if exact_candidate_contexts:
                    submitted_evidence = '\n\n'.join((f'EXACT-VALUE CANDIDATE CONTEXT {index}:\n{context}' for (index, context) in enumerate(exact_candidate_contexts, start=1)))
                reference_context = {'metric': numeric_label.replace('_', ' '), 'expected_value': numeric_expected, 'equivalence_rule': 'Accept any unambiguous professional label, abbreviation, table layout, or locally disclosed unit that associates the expected value with this metric. Do not require the authored label or a finite alias-list match.', 'association_rule': 'The expected value must belong to this metric, period, and scenario. Reject an unlabeled number, a value belonging to a different metric, or a conflicting primary value even if the expected number appears elsewhere on the supplied slides.', 'candidate_context_rule': 'Each supplied EXACT-VALUE CANDIDATE CONTEXT contains a literal that passed the deterministic magnitude gate, including ordinary displayed rounding. Award credit only when one of those same compact contexts binds that literal to the requested metric, period, scenario, and role. Do not combine the requested label from one row with the expected magnitude from another row.', 'sign_rule': 'Accept a correctly described adverse magnitude, accounting negative, or cash outflow convention; reject a direction that changes the business meaning.'}
                criterion_specific_rule = _TASK_068_NUMERIC_ASSOCIATION_RULES.get(numeric_label)
                if criterion_specific_rule:
                    reference_context['criterion_specific_association_rule'] = criterion_specific_rule
                if numeric_label == 'signed_backlog_coverage_of_remaining_outlook':
                    reference_context['required_ratio_basis'] = {'signed_backlog': gold['answer']['signed_backlog'], 'remaining_fy26_revenue_outlook': gold['answer']['remaining_fy26_revenue_outlook']}
            else:
                reference_context = _task_068_reference_context(spec['id'], expected_facts, gold['answer'])
            evidence_scope = 'slide(s) ' + ', '.join((str(number) for number in slide_numbers))
            if numeric_association:
                requirement = f"Evaluate only whether the supplied slide content clearly associates the deterministically verified value with the {numeric_label.replace('_', ' ')} metric for the correct period and scenario. Accept normal professional wording, abbreviations, layout, rounding, and sign conventions. Reject an unlabeled pile of numbers, the right number attached to the wrong metric, an ambiguous binding, or a contradictory displayed primary value. When exact-value candidate contexts are supplied, the requested meaning and the verified magnitude must be bound in the same compact context; never borrow a label from a different row or shape."
            else:
                requirement = f"Evaluate only this criterion in the supplied slide content: {spec['description']} Use the criterion answer key and verified finance context below. Accept normal professional wording; reject a missing, contradictory, or unsupported conclusion."
            task_context = {'assignment': 'Complete the existing nine-slide June/Q2 executive performance review for leadership.', 'period_scope': "The requested performance, plan, and cash-flow comparison is Q2. A compact row on a Q2-scoped slide need not repeat the token 'Q2' in every cell unless it expressly identifies another period.", 'central_work': 'Tie the accounting and reporting figures and support the resulting guidance recommendation.', 'grading_boundary': 'Checkable magnitudes are deterministic hard gates. This judge grades only their professional association, or the named narrative, source, control, or decision meaning, within the supplied slides.'}
        else:
            submitted_evidence = ''
            reference_context = expected_facts or {}
            evidence_scope = 'legacy whole artifact'
            task_context = {}
        reviews.append({'criterion_id': spec['id'], 'requirement': requirement, 'hard_gate_met': hard_gate_met, 'hard_gate_evidence': hard_gate_evidence, 'legacy_lexical_match': bool(by_id[spec['id']].met), **({'submitted_evidence': submitted_evidence, 'reference_context': reference_context, 'task_context': task_context, 'evidence_scope': evidence_scope, 'always_judge': not canonical_numeric_fast_path} if task_id in {'task_035', 'task_068'} else {})})
    if not reviews:
        return None
    return {'version': 2, 'mode': 'deterministic_hard_gates_plus_bounded_semantic_judge', 'task_id': task_id, 'artifact': gold['artifact']['path'], 'evidence': _semantic_evidence_pack(path, workbook, values), 'execution_mode': 'scoped_per_criterion' if task_id in {'task_035', 'task_068'} else 'legacy_batched', 'criteria': reviews, 'policy': 'A criterion passes only when its deterministic hard gate passes and the semantic judge finds the professional-language association or narrative substance MET.'}

def _artifact_label_value(entries: list[str], label: str, expected: Any, *, directional_strings: bool=False) -> bool:
    for (index, entry) in enumerate(entries):
        normalized = _normalize(entry)
        if not contains_concept(entry, label):
            continue
        candidates: list[Any] = [entry]
        candidates.extend(entries[index + 1:index + 4])
        if isinstance(expected, list):
            if ordered_semantic_list_matches(candidates, expected):
                return True
            continue
        if isinstance(expected, str):
            matcher = _sample_directional_semantic_value_matches if directional_strings else semantic_value_matches
            if any((date_matches(candidate, expected) or matcher(candidate, expected) for candidate in candidates)):
                return True
            continue
        if isinstance(expected, bool):
            if any((_boolean_matches(candidate, expected) for candidate in candidates)):
                return True
            continue
        numbers = []
        for candidate in candidates:
            numbers.extend(re.findall('\\(?-?\\$?\\d[\\d,]*(?:\\.\\d+)?[ \\t]*(?:%|[kmbx×])?\\)?', str(candidate), flags=re.I))

        def literal_tolerance(number: str) -> float:
            cleaned = number.strip().replace('$', '').replace(',', '').strip('() ')
            suffix = cleaned[-1:].casefold()
            multiplier = {'k': 1000.0, 'm': 1000000.0, 'b': 1000000000.0}.get(suffix, 1.0)
            if suffix in {'k', 'm', 'b', 'x', '×'}:
                cleaned = cleaned[:-1].strip()
            percent = cleaned.endswith('%')
            cleaned = cleaned.rstrip('%').strip()
            decimals = len(cleaned.rsplit('.', 1)[1]) if '.' in cleaned else 0
            tolerance = 0.5 * 10 ** (-decimals) * multiplier
            return tolerance / 100 if percent else tolerance
        if any((_close(number, float(expected), abs_tol=max(_display_tolerance(float(expected)), literal_tolerance(number) + 1e-12), rel_tol=0.0) for number in numbers)):
            return True
    return False
_NUMERIC_LITERAL_PATTERN = re.compile('\n    (?<![A-Za-z0-9])\n    (?:[-−]?\\$?[ \\t]*\\(?|\\([ \\t]*\\$?[-−]?[ \\t]*)\n    \\d[\\d,]*(?:\\.\\d+)?[ \\t]*\\)?\n    (?:[ \\t]*(?:\n        %\n        |[x×]\n        |k(?![A-Za-z])\n        |thousand(?![A-Za-z])\n        |m(?:m|illion)?(?![A-Za-z])\n        |b(?:n|illion)?(?![A-Za-z])\n    ))?\n    [ \\t]*\\)?\n    ', flags=re.I | re.X)

def _numeric_literal_matches(literal: str, expected: float) -> bool:
    normalized_literal = literal.replace('−', '-')
    candidates = list(iter_numeric_candidates(normalized_literal))
    accounting_negative = '(' in literal and ')' in literal
    explicit_negative = '-' in normalized_literal.partition(next((c for c in normalized_literal if c.isdigit()), ''))[0]
    percent = literal.rstrip().rstrip(')').rstrip().endswith('%')
    cleaned_for_precision = normalized_literal.strip().replace('$', '').replace(',', '').replace('(', '').replace(')', '').strip()
    scale_match = re.search('(?:\\s*)(thousand|k|million|mm|m|billion|bn|b)\\s*$', cleaned_for_precision, flags=re.I)
    scale_token = scale_match.group(1).casefold() if scale_match else ''
    display_multiplier = {'thousand': 1000.0, 'k': 1000.0, 'million': 1000000.0, 'mm': 1000000.0, 'm': 1000000.0, 'billion': 1000000000.0, 'bn': 1000000000.0, 'b': 1000000000.0}.get(scale_token, 1.0)
    if scale_match:
        cleaned_for_precision = cleaned_for_precision[:scale_match.start()].strip()
    cleaned_for_precision = cleaned_for_precision.rstrip('%x×').strip()
    try:
        displayed_value = float(cleaned_for_precision) * display_multiplier
        if accounting_negative or explicit_negative:
            displayed_value = -abs(displayed_value)
        if percent:
            displayed_value /= 100.0
        candidates.append(displayed_value)
    except ValueError:
        pass
    decimals = len(cleaned_for_precision.rsplit('.', 1)[1]) if '.' in cleaned_for_precision else 0
    literal_tolerance = 0.5 * 10 ** (-decimals) * display_multiplier
    if percent:
        literal_tolerance /= 100.0
    if any((_close(candidate, expected, abs_tol=max(_display_tolerance(expected), literal_tolerance + 1e-12), rel_tol=0.0) for candidate in candidates)):
        return True
    cleaned = normalized_literal.strip().replace('$', '').replace(',', '').strip('() ')
    if abs(expected) <= 10000 or cleaned[-1:].casefold() in {'k', 'm', 'b', '%', 'x', '×'}:
        return False
    try:
        displayed = float(cleaned)
    except ValueError:
        return False
    if accounting_negative or explicit_negative:
        displayed = -abs(displayed)
    scaled = displayed * 1000000.0
    return _close(scaled, expected, abs_tol=max(5000.0, _display_tolerance(expected)), rel_tol=0.0)

def _context_has_alias(context: str, alias: str) -> bool:
    if contains_concept(context, alias):
        return True
    context_tokens = set(_normalize(context).split())
    alias_tokens = [token for token in _normalize(alias).split() if len(token) > 1]
    return bool(alias_tokens) and all((token in context_tokens for token in alias_tokens))
_TASK_068_NUMERIC_LITERAL_PATTERN = re.compile('\n    (?<![A-Za-z0-9])\n    (?:[-−]?\\$?[ \\t]*\\(?|\\([ \\t]*\\$?[-−]?[ \\t]*)\n    \\d[\\d,]*(?:\\.\\d+)?[ \\t]*\\)?\n    (?:[ \\t]*(?:\n        %\n        |[x×]\n        |k(?![A-Za-z])\n        |thousand(?![A-Za-z])\n        |m(?:m|illion)?(?![A-Za-z])\n        |b(?:n|illion)?(?![A-Za-z])\n    ))?\n    [ \\t]*\\)?\n    ', flags=re.I | re.X)

def _task_068_value_slides(label: str) -> tuple[int, ...]:
    if label.startswith(('construction_q2_', 'service_q2_', 'controls_q2_')):
        return (5,)
    if label in {'construction_gross_profit_impact', 'service_labor_productivity_impact', 'controls_mix_impact'}:
        return (2, 4, 5, 8, 9)
    if label in {'identified_ebitda_action_pool', 'full_supported_mitigation_conversion', 'full_supported_ebitda_mitigation', 'probability_adjusted_mitigation_conversion', 'probability_adjusted_executable_ebitda_mitigation', 'post_mitigation_combined_stress_ebitda', 'post_mitigation_headroom_to_guidance_low', 'guidance_action_gap_to_required_headroom'}:
        return (2, 4, 7, 8, 9)
    if label == 'june_revenue':
        return (3,)
    if label == 'june_ebitda':
        return (4,)
    if label in {'posted_ytd_revenue', 'q2_posted_ytd_revenue_share'}:
        return (2, 3, 5, 7, 9)
    if label in {'q2_posted_ytd_revenue_reconciliation_check', 'q2_revenue', 'q2_approved_plan_revenue', 'q2_revenue_variance_to_plan', 'q2_organic_growth'}:
        return (2, 3, 5, 9)
    if label in {'q2_adjusted_ebitda', 'q2_approved_plan_adjusted_ebitda', 'q2_adjusted_ebitda_variance_to_plan'}:
        return (2, 4, 5, 9)
    if label in {'q2_operating_cash_flow', 'q2_capital_expenditure', 'q2_free_cash_flow', 'q2_operating_cash_flow_plan', 'q2_capital_expenditure_plan', 'q2_free_cash_flow_plan_derived', 'q2_free_cash_flow_plan_stated', 'q2_free_cash_flow_plan_discrepancy', 'q2_free_cash_flow_variance_to_derived_plan', 'collections_timing_cash_impact', 'q2_unrestricted_cash', 'maximum_revolver'}:
        return (2, 6, 9)
    if label in {'signed_backlog', 'remaining_fy26_revenue_outlook', 'signed_backlog_coverage_of_remaining_outlook', 'latest_full_year_revenue_outlook'}:
        return (2, 7, 9)
    if label in {'revenue_guidance_low', 'revenue_guidance_high', 'ebitda_guidance_low', 'ebitda_guidance_high', 'minimum_guidance_update_headroom', 'guidance_update_trigger', 'combined_stress_ebitda', 'ebitda_shortfall_to_guidance_low'}:
        return (2, 7, 8, 9)
    return (2, 4, 8, 9)

def _task_068_template_numeric_bindings(label: str) -> tuple[tuple[int, tuple[str, ...], str], ...]:
    """Return only pre-authored Task068 label/value surfaces.

    Each tuple is ``(slide number, unchanged label/heading shape names,
    editable value shape name)``. A recognized alias added elsewhere is not a
    canonical surface and therefore still requires semantic review.
    """
    direct: dict[str, tuple[tuple[int, tuple[str, ...], str], ...]] = {'june_revenue': ((3, ('metric-label-3-0',), 'metric-value-3-0'),), 'q2_revenue': ((2, ('metric-label-2-0',), 'metric-value-2-0'), (3, ('metric-label-3-2',), 'metric-value-3-2'), (3, ('revenue-bridge-heading', 'rev-cat-5'), 'rev-val-5'), (5, ('score-text-4-0', 'score-head-text-1'), 'score-text-4-1')), 'q2_approved_plan_revenue': ((3, ('revenue-bridge-heading', 'rev-cat-0'), 'rev-val-0'), (5, ('score-text-4-0', 'score-head-text-2'), 'score-text-4-2')), 'q2_revenue_variance_to_plan': ((5, ('score-text-4-0', 'score-head-text-3'), 'score-text-4-3'),), 'q2_adjusted_ebitda': ((2, ('metric-label-2-1',), 'metric-value-2-1'), (4, ('metric-label-4-0',), 'metric-value-4-0'), (4, ('ebitda-bridge-heading', 'ebitda-driver-5'), 'ebitda-val-5'), (5, ('score-text-4-0', 'score-head-text-4'), 'score-text-4-4')), 'q2_approved_plan_adjusted_ebitda': ((4, ('ebitda-bridge-heading', 'ebitda-driver-0'), 'ebitda-val-0'), (5, ('score-text-4-0', 'score-head-text-5'), 'score-text-4-5')), 'q2_adjusted_ebitda_variance_to_plan': ((4, ('metric-label-4-2',), 'metric-value-4-2'), (5, ('score-text-4-0', 'score-head-text-6'), 'score-text-4-6')), 'q2_free_cash_flow': ((2, ('metric-label-2-2',), 'metric-value-2-2'), (6, ('metric-label-6-0',), 'metric-value-6-0'), (6, ('cash-heading', 'cash-label-2'), 'cash-actual-2')), 'latest_full_year_revenue_outlook': ((2, ('metric-label-2-3',), 'metric-value-2-3'), (3, ('metric-label-3-3',), 'metric-value-3-3'), (7, ('metric-label-7-1',), 'metric-value-7-1'), (7, ('outlook-heading', 'outlook-name-0'), 'outlook-val-0')), 'q2_operating_cash_flow': ((6, ('cash-heading', 'cash-label-0'), 'cash-actual-0'),), 'q2_operating_cash_flow_plan': ((6, ('cash-heading', 'cash-label-0'), 'cash-plan-0'),), 'q2_capital_expenditure': ((6, ('cash-heading', 'cash-label-1'), 'cash-actual-1'),), 'q2_capital_expenditure_plan': ((6, ('cash-heading', 'cash-label-1'), 'cash-plan-1'),), 'q2_free_cash_flow_plan_derived': ((6, ('cash-heading', 'cash-label-2'), 'cash-plan-2'),), 'q2_free_cash_flow_plan_stated': ((6, ('cash-heading', 'cash-label-3'), 'cash-plan-3'), (6, ('cash-heading', 'cash-label-3'), 'cash-actual-3')), 'q2_free_cash_flow_variance_to_derived_plan': ((6, ('metric-label-6-3',), 'metric-value-6-3'), (6, ('cash-heading', 'cash-label-4'), 'cash-actual-4')), 'collections_timing_cash_impact': ((6, ('cash-heading', 'cash-label-5'), 'cash-actual-5'),), 'q2_unrestricted_cash': ((6, ('metric-label-6-1',), 'metric-value-6-1'),), 'maximum_revolver': ((6, ('metric-label-6-2',), 'metric-value-6-2'),), 'signed_backlog': ((7, ('metric-label-7-0',), 'metric-value-7-0'),), 'remaining_fy26_revenue_outlook': ((7, ('metric-label-7-2',), 'metric-value-7-2'),), 'signed_backlog_coverage_of_remaining_outlook': ((7, ('metric-label-7-3',), 'metric-value-7-3'),), 'revenue_guidance_low': ((7, ('outlook-heading', 'outlook-name-1'), 'outlook-val-1'),), 'revenue_guidance_high': ((7, ('outlook-heading', 'outlook-name-2'), 'outlook-val-2'),)}
    branch_match = re.fullmatch('(construction|service|controls)_q2_(revenue|approved_plan_revenue|revenue_variance_to_plan|adjusted_ebitda|approved_plan_adjusted_ebitda|adjusted_ebitda_variance_to_plan)', label)
    if branch_match:
        row = {'construction': 0, 'service': 1, 'controls': 2}[branch_match.group(1)]
        column = {'revenue': 1, 'approved_plan_revenue': 2, 'revenue_variance_to_plan': 3, 'adjusted_ebitda': 4, 'approved_plan_adjusted_ebitda': 5, 'adjusted_ebitda_variance_to_plan': 6}[branch_match.group(2)]
        return ((5, (f'score-text-{row}-0', f'score-head-text-{column}'), f'score-text-{row}-{column}'),)
    return direct.get(label, ())

def _task_068_shape_geometry_unchanged(current: Any, seed: Any) -> bool:
    """Return whether a named template shape still occupies its fixed surface."""
    return all((float(getattr(current, attribute, 0) or 0) == float(getattr(seed, attribute, 0) or 0) for attribute in ('left', 'top', 'width', 'height', 'rotation')))

def _task_068_unchanged_template_numeric_match(path: Path, artifact_relative: str, label: str, expected: float, *, current: Presentation | None=None, seed: Presentation | None=None) -> bool:
    """Allow a deterministic fast path only on an unchanged template label.

    Any relabeling, moved value, added alias, or alternative layout returns
    ``False`` and is sent to semantic review even when the broader deterministic
    parser recognizes it.
    """
    bindings = _task_068_template_numeric_bindings(label)
    seed_path = SEED_WORKSPACE / artifact_relative
    if not bindings or not seed_path.is_file():
        return False
    if current is None:
        current = Presentation(path)
    if seed is None:
        seed = Presentation(seed_path)
    for (slide_number, label_shape_names, value_shape_name) in bindings:
        if slide_number > len(current.slides) or slide_number > len(seed.slides):
            continue
        current_shapes = {shape.name: shape for shape in current.slides[slide_number - 1].shapes}
        seed_shapes = {shape.name: shape for shape in seed.slides[slide_number - 1].shapes}
        if value_shape_name not in current_shapes:
            continue
        if any((name not in current_shapes or name not in seed_shapes or (not hasattr(current_shapes[name], 'text')) or (not hasattr(seed_shapes[name], 'text')) or (_normalize(current_shapes[name].text) != _normalize(seed_shapes[name].text)) or (not _task_068_shape_geometry_unchanged(current_shapes[name], seed_shapes[name])) for name in label_shape_names)):
            continue
        value_shape = current_shapes[value_shape_name]
        if not hasattr(value_shape, 'text') or value_shape_name not in seed_shapes or (not _task_068_shape_geometry_unchanged(value_shape, seed_shapes[value_shape_name])):
            continue
        context = ' | '.join((*(str(current_shapes[name].text).strip() for name in label_shape_names), str(value_shape.text).strip()))
        targets = (expected,) if expected == 0 else (expected, -expected)
        if any((_task_068_numeric_literal_matches(literal, target, context=context) for literal in _TASK_068_NUMERIC_LITERAL_PATTERN.findall(context) for target in targets)):
            return True
    return False

def _task_068_unchanged_template_numeric_status(path: Path, artifact_relative: str, label: str, expected: float) -> tuple[str, str]:
    """Classify an unchanged canonical value field as match/conflict/absent.

    This is intentionally narrower than semantic association. It is used to
    reject an explicitly contradictory value in a fixed starter field while
    leaving relabeled, moved, or alternate-layout work to the semantic judge.
    """
    bindings = _task_068_template_numeric_bindings(label)
    seed_path = SEED_WORKSPACE / artifact_relative
    if not bindings or not seed_path.is_file():
        return ('absent', 'no canonical template binding')
    current = Presentation(path)
    seed = Presentation(seed_path)
    saw_numeric_value = False
    conflict_evidence = ''
    for (slide_number, label_shape_names, value_shape_name) in bindings:
        if slide_number > len(current.slides) or slide_number > len(seed.slides):
            continue
        current_shapes = {shape.name: shape for shape in current.slides[slide_number - 1].shapes}
        seed_shapes = {shape.name: shape for shape in seed.slides[slide_number - 1].shapes}
        if value_shape_name not in current_shapes:
            continue
        if any((name not in current_shapes or name not in seed_shapes or (not hasattr(current_shapes[name], 'text')) or (not hasattr(seed_shapes[name], 'text')) or (_normalize(current_shapes[name].text) != _normalize(seed_shapes[name].text)) or (not _task_068_shape_geometry_unchanged(current_shapes[name], seed_shapes[name])) for name in label_shape_names)):
            continue
        value_shape = current_shapes[value_shape_name]
        if not hasattr(value_shape, 'text') or value_shape_name not in seed_shapes or (not _task_068_shape_geometry_unchanged(value_shape, seed_shapes[value_shape_name])):
            continue
        labels = tuple((str(current_shapes[name].text).strip() for name in label_shape_names))
        value_text = str(value_shape.text).strip()
        context = ' | '.join((*labels, value_text))
        literals = _TASK_068_NUMERIC_LITERAL_PATTERN.findall(value_text)
        if not literals:
            continue
        saw_numeric_value = True
        targets = (expected,) if expected == 0 else (expected, -expected)
        matched = any((_task_068_numeric_literal_matches(literal, target, context=context) for literal in literals for target in targets))
        if matched and len(literals) == 1:
            return ('match', f'unchanged canonical field matched: {context}')
        if matched:
            return ('absent', f'canonical value field contains multiple numeric roles: {context}')
        conflict_evidence = context
    if saw_numeric_value:
        return ('conflict', f'unchanged canonical field contains a conflicting value: {conflict_evidence}')
    return ('absent', 'unchanged canonical field is blank or unavailable')
_TASK_068_CANONICAL_NUMERIC_FAST_PATH_LABELS = frozenset((f'{branch}_q2_{metric}' for branch in ('construction', 'service', 'controls') for metric in ('approved_plan_revenue', 'approved_plan_adjusted_ebitda')))
_TASK_068_SCORECARD_VALUE_SHAPE_PATTERN = re.compile('score-text-[0-4]-[1-6]')

def _task_068_canonical_scorecard_structure_unchanged(path: Path, artifact_relative: str, label: str) -> bool:
    """Keep the fixed-cell shortcut limited to the unchanged scorecard schema.

    A professional source footnote outside the fixed scorecard must not turn six
    objective plan cells into stochastic semantic judgments.  Conversely, an
    added shape that overlaps the scorecard or introduces a competing
    branch/plan statement disables the shortcut and returns the criterion to
    semantic review.
    """
    bindings = _task_068_template_numeric_bindings(label)
    seed_path = SEED_WORKSPACE / artifact_relative
    if not bindings or not seed_path.is_file():
        return False
    current = Presentation(path)
    seed = Presentation(seed_path)
    for (slide_number, _label_shape_names, _value_shape_name) in bindings:
        if slide_number > len(current.slides) or slide_number > len(seed.slides):
            return False
        current_shapes = {shape.name: shape for shape in current.slides[slide_number - 1].shapes}
        seed_shapes = {shape.name: shape for shape in seed.slides[slide_number - 1].shapes}
        if not set(seed_shapes).issubset(current_shapes):
            return False
        for (shape_name, seed_shape) in seed_shapes.items():
            current_shape = current_shapes[shape_name]
            if not _task_068_shape_geometry_unchanged(current_shape, seed_shape):
                return False
            if _TASK_068_SCORECARD_VALUE_SHAPE_PATTERN.fullmatch(shape_name):
                continue
            if hasattr(current_shape, 'text') != hasattr(seed_shape, 'text'):
                return False
            if hasattr(current_shape, 'text') and _normalize(current_shape.text) != _normalize(seed_shape.text):
                return False
        scorecard_shapes = [shape for (name, shape) in seed_shapes.items() if name.startswith(('score-head-', 'score-text-', 'score-row-', 'score-grid-'))]
        if not scorecard_shapes:
            return False
        scorecard_left = min((int(shape.left) for shape in scorecard_shapes))
        scorecard_top = min((int(shape.top) for shape in scorecard_shapes))
        scorecard_right = max((int(shape.left + shape.width) for shape in scorecard_shapes))
        scorecard_bottom = max((int(shape.top + shape.height) for shape in scorecard_shapes))
        branch = label.split('_q2_', 1)[0]
        branch_terms = {'construction': ('construction',), 'service': ('service',), 'controls': ('controls', 'building controls')}[branch]
        metric_terms = ('revenue', 'sales') if '_revenue' in label else ('ebitda', 'operating earnings', 'adjusted earnings')
        plan_terms = ('plan', 'aop', 'budget', 'target')
        for shape_name in set(current_shapes) - set(seed_shapes):
            shape = current_shapes[shape_name]
            shape_left = int(shape.left)
            shape_top = int(shape.top)
            shape_right = int(shape.left + shape.width)
            shape_bottom = int(shape.top + shape.height)
            overlaps_scorecard = not (shape_right <= scorecard_left or shape_left >= scorecard_right or shape_bottom <= scorecard_top or (shape_top >= scorecard_bottom))
            if overlaps_scorecard:
                return False
            fragments: list[str] = []
            if getattr(shape, 'has_table', False):
                fragments.extend((' | '.join((str(cell.text or '') for cell in row.cells)) for row in shape.table.rows))
            elif hasattr(shape, 'text'):
                fragments.append(str(shape.text or ''))
            for fragment in fragments:
                for segment in re.split('[.;\\n]+', fragment):
                    normalized = _normalize(segment)
                    has_branch = any((term in normalized for term in branch_terms))
                    has_plan = any((term in normalized for term in plan_terms))
                    has_metric = any((term in normalized for term in metric_terms))
                    if has_branch and has_plan and has_metric:
                        return False
    return True

def _task_068_safe_canonical_numeric_fast_path(path: Path, artifact_relative: str, label: str, expected: float) -> bool:
    """Use a deterministic final association only for unambiguous fixed cells.

    The six business-unit plan cells have one stable row/column meaning in the
    neutral Q2 scorecard. Their unchanged row and column labels plus the exact
    value fully establish the association. All actuals, variances, aggregate
    metrics, edited labels, moved values, and alternate layouts remain semantic.
    """
    if label not in _TASK_068_CANONICAL_NUMERIC_FAST_PATH_LABELS:
        return False
    if not _task_068_canonical_scorecard_structure_unchanged(path, artifact_relative, label):
        return False
    (status, _evidence) = _task_068_unchanged_template_numeric_status(path, artifact_relative, label, expected)
    return status == 'match'

def _task_068_local_contexts(path: Path, slide_numbers: Iterable[int]) -> list[str]:
    """Return local shape/row contexts, never an indiscriminate whole-deck dump."""
    stat = path.stat()
    return list(_task_068_local_contexts_cached(str(path.resolve()), stat.st_mtime_ns, stat.st_size, tuple(slide_numbers)))

@lru_cache(maxsize=256)
def _task_068_local_contexts_cached(path_text: str, _mtime_ns: int, _size: int, slide_numbers: tuple[int, ...]) -> tuple[str, ...]:
    path = Path(path_text)
    presentation = Presentation(path)
    contexts: list[str] = []
    for slide_number in slide_numbers:
        if slide_number < 1 or slide_number > len(presentation.slides):
            continue
        slide = presentation.slides[slide_number - 1]
        leaf_shapes = list(_task_068_leaf_shapes(slide.shapes))
        text_shapes = [shape for shape in leaf_shapes if hasattr(shape, 'text') and str(shape.text or '').strip()]
        ordered = sorted(text_shapes, key=lambda shape: (int(shape.top), int(shape.left)))
        texts = [str(shape.text).strip() for shape in ordered]
        contexts.extend(texts)
        for start in range(len(texts)):
            for width in range(2, 7):
                window = texts[start:start + width]
                if len(window) == width:
                    contexts.append(' | '.join(window))
        row_tolerance = 180000
        column_tolerance = 300000
        for shape in text_shapes:
            shape_text = str(shape.text).strip()
            same_row = sorted((candidate for candidate in text_shapes if abs(int(candidate.top) - int(shape.top)) <= row_tolerance), key=lambda candidate: int(candidate.left))
            if len(same_row) >= 2:
                row_context = ' | '.join((str(candidate.text).strip() for candidate in same_row))
                contexts.append(row_context)
                row_left = min((int(candidate.left) for candidate in same_row))
                row_top = min((int(candidate.top) for candidate in same_row))
                row_headings = [candidate for candidate in text_shapes if int(candidate.top) < row_top and int(candidate.left) <= row_left <= int(candidate.left + candidate.width) and re.search('(?:^|[-_])(?:title|heading)(?:[-_]|$)', str(getattr(candidate, 'name', '')), flags=re.I) and (not _TASK_068_NUMERIC_LITERAL_PATTERN.search(str(candidate.text))) and (len(str(candidate.text).strip()) <= 160)]
                if row_headings:
                    heading = max(row_headings, key=lambda candidate: int(candidate.top))
                    contexts.append(' | '.join((str(heading.text).strip(), row_context)))
            left_candidates = [candidate for candidate in same_row if int(candidate.left) < int(shape.left)]
            if left_candidates:
                row_label = min(left_candidates, key=lambda candidate: int(candidate.left))
                contexts.append(' | '.join((str(row_label.text).strip(), shape_text)))
            above_candidates = [candidate for candidate in text_shapes if int(candidate.top) < int(shape.top) and abs(int(candidate.left) - int(shape.left)) <= column_tolerance and re.search('(?:^|[-_])(?:head|header)(?:[-_]|$)', str(getattr(candidate, 'name', '')), flags=re.I) and (not _TASK_068_NUMERIC_LITERAL_PATTERN.search(str(candidate.text)))]
            if left_candidates and above_candidates:
                for header in sorted(above_candidates, key=lambda candidate: int(shape.top) - int(candidate.top))[:8]:
                    contexts.append(' | '.join((str(row_label.text).strip(), str(header.text).strip(), shape_text)))
            shape_center = int(shape.left) + int(shape.width) // 2
            shape_slot = re.search('(?P<slot>\\d+(?:-\\d+)*)$', str(getattr(shape, 'name', '')))
            aligned = sorted((candidate for candidate in text_shapes if candidate is not shape and abs(int(candidate.left) + int(candidate.width) // 2 - shape_center) <= column_tolerance and (abs(int(candidate.top) - int(shape.top)) <= 350000 or (shape_slot is not None and (candidate_slot := re.search('(?P<slot>\\d+(?:-\\d+)*)$', str(getattr(candidate, 'name', '')))) is not None and (candidate_slot.group('slot') == shape_slot.group('slot'))))), key=lambda candidate: abs(int(candidate.top) - int(shape.top)))
            for candidate in aligned[:8]:
                candidate_text = str(candidate.text).strip()
                if _TASK_068_NUMERIC_LITERAL_PATTERN.search(shape_text) and _TASK_068_NUMERIC_LITERAL_PATTERN.search(candidate_text):
                    continue
                pair = (candidate_text, shape_text)
                contexts.append(' | '.join(pair))
                pair_top = min(int(candidate.top), int(shape.top))
                headings = [heading for heading in text_shapes if heading is not shape and heading is not candidate and (int(heading.top) < pair_top) and (int(heading.left) <= shape_center <= int(heading.left + heading.width)) and re.search('(?:^|[-_])(?:title|heading)(?:[-_]|$)', str(getattr(heading, 'name', '')), flags=re.I) and (not _TASK_068_NUMERIC_LITERAL_PATTERN.search(str(heading.text))) and (len(str(heading.text).strip()) <= 160)]
                if headings:
                    heading = max(headings, key=lambda item: int(item.top))
                    contexts.append(' | '.join((str(heading.text).strip(), *pair)))
        for shape in leaf_shapes:
            if not getattr(shape, 'has_table', False):
                continue
            rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
            if not rows:
                continue
            headers = rows[0]
            contexts.extend((' | '.join(row) for row in rows))
            for row in rows[1:]:
                for (column, cell) in enumerate(row):
                    header = headers[column] if column < len(headers) else ''
                    contexts.append(' | '.join((row[0] if row else '', header, cell)))
    return tuple(dict.fromkeys((context for context in contexts if context.strip())))
_TASK_068_LABEL_ALIASES = {'june_revenue': ('June revenue', 'June actual revenue', 'June actual'), 'june_ebitda': ('June EBITDA', 'June adjusted EBITDA', 'June actual'), 'posted_ytd_revenue': ('posted YTD revenue', 'Vista YTD revenue'), 'q2_posted_ytd_revenue_share': ('Q2 share', 'Q2 revenue share', '53% Q2 share'), 'q2_posted_ytd_revenue_reconciliation_check': ('Q2 revenue reconciliation check', 'posted YTD revenue tie', 'no plug'), 'q2_revenue': ('Q2 revenue', 'Q2 actual revenue'), 'q2_approved_plan_revenue': ('Q2 revenue plan', 'approved plan revenue', 'revenue plan', 'plan'), 'q2_revenue_variance_to_plan': ('Q2 revenue vs plan', 'revenue variance to plan'), 'q2_organic_growth': ('organic growth', 'organic YoY', 'organic year over year', 'IR organic'), 'q2_adjusted_ebitda': ('Q2 adjusted EBITDA', 'Q2 EBITDA actual'), 'q2_approved_plan_adjusted_ebitda': ('Q2 adjusted EBITDA plan', 'approved plan EBITDA', 'EBITDA plan', 'plan'), 'q2_adjusted_ebitda_variance_to_plan': ('Q2 EBITDA vs plan', 'adjusted EBITDA variance to plan', 'EBITDA variance'), 'q2_operating_cash_flow': ('Q2 operating cash flow', 'operating cash flow', 'OCF'), 'q2_capital_expenditure': ('Q2 capital expenditure', 'capital expenditure', 'capex'), 'q2_free_cash_flow': ('Q2 free cash flow', 'free cash flow', 'FCF'), 'q2_operating_cash_flow_plan': ('Q2 operating cash flow plan', 'plan OCF', 'OCF plan'), 'q2_capital_expenditure_plan': ('Q2 capital expenditure plan', 'plan capex', 'capex plan'), 'q2_free_cash_flow_plan_derived': ('derived Q2 free cash flow plan', 'derived plan FCF', 'FCF plan derived'), 'q2_free_cash_flow_plan_stated': ('stated Q2 free cash flow plan', 'stated plan FCF', 'FCF plan stated'), 'q2_free_cash_flow_plan_discrepancy': ('FCF plan discrepancy', 'plan line difference', 'plan inconsistency'), 'q2_free_cash_flow_variance_to_derived_plan': ('FCF vs derived plan', 'free cash flow variance to derived plan'), 'collections_timing_cash_impact': ('collections timing', 'collections cash impact'), 'q2_unrestricted_cash': ('Q2 unrestricted cash', 'unrestricted cash'), 'maximum_revolver': ('maximum revolver', 'max revolver', 'downside revolver'), 'signed_backlog': ('signed backlog',), 'remaining_fy26_revenue_outlook': ('remaining FY26 revenue', 'remaining revenue outlook', 'H2 remaining revenue', 'H2 remaining revenue need'), 'signed_backlog_coverage_of_remaining_outlook': ('backlog coverage of remaining outlook', 'FY26 remaining coverage', 'backlog coverage'), 'latest_full_year_revenue_outlook': ('FY26 revenue outlook', 'June reforecast'), 'ltm_adjusted_ebitda': ('LTM adjusted EBITDA', 'lender EBITDA'), 'funded_debt': ('funded debt',), 'lender_leverage': ('lender leverage', 'gross leverage', 'leverage'), 'fixed_charge_coverage': ('fixed charge coverage', 'FCCR'), 'revenue_guidance_low': ('revenue guidance low', 'revenue guidance'), 'revenue_guidance_high': ('revenue guidance high', 'revenue guidance'), 'ebitda_guidance_low': ('EBITDA guidance low', 'guidance low'), 'ebitda_guidance_high': ('EBITDA guidance high', 'EBITDA guidance'), 'construction_gross_profit_impact': ('construction gross profit', 'construction GP'), 'service_labor_productivity_impact': ('service labor productivity', 'service labor'), 'controls_mix_impact': ('controls mix',), 'identified_ebitda_action_pool': ('identified EBITDA action pool', 'action pool'), 'full_supported_mitigation_conversion': ('full supported mitigation conversion', 'supported recovery conversion'), 'full_supported_ebitda_mitigation': ('full supported mitigation', 'supported EBITDA recovery', 'action pool supported recovery'), 'probability_adjusted_mitigation_conversion': ('probability-adjusted mitigation conversion', 'expected recovery conversion'), 'probability_adjusted_executable_ebitda_mitigation': ('probability-adjusted executable mitigation', 'expected executable recovery', 'probability weighted operating recovery'), 'combined_stress_ebitda': ('combined stress EBITDA', 'pre-mitigation stress EBITDA'), 'ebitda_shortfall_to_guidance_low': ('shortfall to guidance low', 'guidance shortfall'), 'post_mitigation_combined_stress_ebitda': ('post-mitigation stress EBITDA',), 'post_mitigation_headroom_to_guidance_low': ('post-mitigation headroom', 'headroom to guidance low'), 'minimum_guidance_update_headroom': ('minimum update buffer', 'minimum headroom'), 'guidance_update_trigger': ('guidance update trigger', 'update trigger'), 'guidance_action_gap_to_required_headroom': ('action gap to required headroom', 'remaining EBITDA recovery gap', 'shortfall to update threshold'), 'guidance_protection_portfolio_treatment': ('guidance protection portfolio', 'guidance protection register', 'contingency portfolio'), 'largest_downside_driver': ('largest downside driver', 'copper escalation'), 'guidance_release_status': ('guidance release status', 'guidance'), 'guidance_update_required': ('guidance update', 'formal update')}

def _task_068_numeric_literal_matches(literal: str, expected: float, *, context: str) -> bool:
    """Match a board-deck number without inventing an undisclosed unit.

    The generic investment-committee matcher accepts an unqualified decimal as
    millions because those tables use a single ``$mm`` heading.  Task068 also
    contains ratios such as the 0.62 mitigation-conversion factor.  Treating
    that ratio as $0.62 million can falsely satisfy an unrelated EBITDA value.
    Here an unqualified compact number receives million scaling only when its
    *local* context explicitly declares a dollar-in-millions display unit.
    Explicit ``k``/``M`` suffixes and full-dollar figures retain the normal
    professional board-rounding tolerance.
    """
    normalized_literal = literal.replace('−', '-').strip()
    cleaned = normalized_literal.replace('$', '').replace(',', '').replace('(', '').replace(')', '').strip()
    unit_match = re.search('(%|[x×]|thousand|k|million|mm|m|billion|bn|b)\\s*$', cleaned, flags=re.I)
    unit_token = unit_match.group(1).casefold() if unit_match else ''
    numeric_text = cleaned[:unit_match.start()].strip() if unit_match else cleaned
    try:
        displayed = float(numeric_text)
    except ValueError:
        return False
    accounting_negative = '(' in literal and ')' in literal
    explicit_negative = normalized_literal.lstrip().startswith('-')
    if accounting_negative or explicit_negative:
        displayed = -abs(displayed)
    decimals = len(numeric_text.rsplit('.', 1)[1]) if '.' in numeric_text else 0
    display_step = 0.5 * 10 ** (-decimals)
    scale = {'thousand': 1000.0, 'k': 1000.0, 'million': 1000000.0, 'mm': 1000000.0, 'm': 1000000.0, 'billion': 1000000000.0, 'bn': 1000000000.0, 'b': 1000000000.0}.get(unit_token)
    if scale is not None:
        return _close(displayed * scale, expected, abs_tol=max(_display_tolerance(expected), display_step * scale + 1e-12), rel_tol=0.0)
    if unit_token == '%':
        return _close(displayed / 100.0, expected, abs_tol=max(_display_tolerance(expected), display_step / 100.0 + 1e-12), rel_tol=0.0)
    if unit_token in {'x', '×'}:
        if decimals == 0 and (not _close(displayed, expected, abs_tol=_display_tolerance(expected), rel_tol=0.0)):
            return False
        return _close(displayed, expected, abs_tol=max(_display_tolerance(expected), display_step + 1e-12), rel_tol=0.0)
    if '$' in normalized_literal:
        if abs(expected) <= 10:
            return False
        return _close(displayed, expected, abs_tol=max(_display_tolerance(expected), display_step + 1e-12), rel_tol=0.0)
    if abs(displayed) >= 10000:
        return _close(displayed, expected, abs_tol=max(_display_tolerance(expected), display_step + 1e-12), rel_tol=0.0)
    if abs(expected) <= 10000:
        if decimals == 0 and (not _close(displayed, expected, abs_tol=_display_tolerance(expected), rel_tol=0.0)):
            return False
        return _close(displayed, expected, abs_tol=max(_display_tolerance(expected), display_step + 1e-12), rel_tol=0.0)
    dollars_in_millions = bool(re.search('(?:\\$\\s*(?:in\\s+)?(?:millions?|mm|m)\\b|\\busd\\s*(?:in\\s+)?(?:millions?|mm|m)\\b|\\bdollars?\\s+in\\s+millions?\\b)', context, flags=re.I))
    if not dollars_in_millions:
        return False
    scaled = displayed * 1000000.0
    rounding_tolerance = display_step * 1000000.0
    return _close(scaled, expected, abs_tol=max(_display_tolerance(expected), rounding_tolerance + 1e-12), rel_tol=0.0)
_TASK_068_SHARED_UNIT_RANGE_PATTERN = re.compile('(?<![A-Za-z0-9])(?P<currency>\\$?)\\s*(?P<low>\\d[\\d,]*(?:\\.\\d+)?)\\s*(?:[-–—]|\\bto\\b)\\s*\\$?\\s*(?P<high>\\d[\\d,]*(?:\\.\\d+)?)\\s*(?P<unit>k|m|mm|million|b|bn|billion)\\b', flags=re.I)

def _task_068_numeric_literals(context: str) -> tuple[str, ...]:
    """Return ordinary literals plus both endpoints of a shared-unit range.

    Board decks conventionally write ranges such as ``$3.0–4.25M`` with the
    unit only once.  Treating the low endpoint as three dollars rejects normal
    business notation, while copying the unit to both endpoints preserves the
    exact objective magnitude and leaves its EBITDA/revenue meaning to the
    semantic reviewer.
    """
    literals = list(_TASK_068_NUMERIC_LITERAL_PATTERN.findall(context))
    for match in _TASK_068_SHARED_UNIT_RANGE_PATTERN.finditer(context):
        currency = match.group('currency')
        unit = match.group('unit')
        literals.extend((f"{currency}{match.group('low')}{unit}", f"{currency}{match.group('high')}{unit}"))
    return tuple(dict.fromkeys(literals))

def _task_068_exact_value_candidate_contexts(path: Path, label: str, expected: float, slide_numbers: Iterable[int] | None=None) -> tuple[str, ...]:
    """Return compact local contexts containing the hard-gated magnitude.

    The bounded judge must bind the *verified value* to editable business
    language.  Supplying every number on several slides lets a judge borrow an
    expected magnitude from one row and a target label from another.  These
    contexts are selected without a label alias: deterministic code filters
    only for the objective magnitude, then the semantic judge decides whether
    the same compact context establishes the metric, period, scenario, sign,
    and role.  This preserves alternate labels while preventing cross-row
    number/label leakage.
    """
    scopes = tuple(slide_numbers or _task_068_value_slides(label))
    targets = (expected,) if expected == 0 else (expected, -expected)
    candidates: list[tuple[int, int, int, str]] = []
    fallback: list[tuple[int, int, int, str]] = []
    for context in _task_068_local_contexts(path, scopes):
        literals = _task_068_numeric_literals(context)
        if not any((_task_068_numeric_literal_matches(literal, target, context=context) for literal in literals for target in targets)):
            continue
        words = [word.casefold() for word in re.findall('[A-Za-z][A-Za-z-]+', context) if word.casefold() not in {'k', 'm', 'mm', 'b', 'bn', 'x'}]
        row = (len(literals), -min(len(words), 12), len(context), context)
        fallback.append(row)
        if words:
            candidates.append(row)
    ordered = sorted(candidates or fallback, key=lambda row: row[:3])
    selected: list[str] = []
    normalized_selected: set[str] = set()
    for (_literal_count, _word_rank, _length, context) in ordered:
        normalized = _normalize(context)
        if normalized in normalized_selected:
            continue
        selected.append(context)
        normalized_selected.add(normalized)
        if len(selected) >= 16:
            break
    return tuple(selected)

def _task_068_numeric_fact_present(path: Path, label: str, expected: float, slide_numbers: Iterable[int] | None=None) -> tuple[bool, str]:
    """Prove a Task068 magnitude before semantic metric association.

    The existing task-specific parser remains a safe fast path for canonical
    or already-vetted professional labels. If the author relabels or rearranges
    the deck, this broader gate asks only whether the checkable magnitude
    occurs on a relevant slide under a disclosed unit. The semantic reviewer
    is solely responsible for binding it to the right metric, period, scenario,
    and direction, so a correct number beside the wrong label cannot earn
    credit.
    """
    if _task_068_artifact_value(path, label, expected):
        return (True, 'exact value matched a vetted local metric association')
    scopes = tuple(slide_numbers or _task_068_value_slides(label))
    contexts = _task_068_local_contexts(path, scopes)
    targets = (expected,) if expected == 0 else (expected, -expected)
    for context in contexts:
        for literal in _task_068_numeric_literals(context):
            if any((_task_068_numeric_literal_matches(literal, target, context=context) for target in targets)):
                return (True, 'exact magnitude appears on relevant slide(s); metric, period, scenario, sign, and unit association requires semantic review')
    return (False, 'required exact magnitude is absent from the relevant slide(s) under any supported displayed unit')

def _task_068_current_q2_reconciliation_components_present(path: Path, answer: dict[str, Any], slide_numbers: Iterable[int]) -> tuple[bool, str]:
    """Require the objective close-basis components before judging a zero tie.

    A generic ``TIED`` or ``no plug`` statement can truthfully describe an
    earlier raw-cube bridge while omitting the later controller-approved close
    entry.  The zero itself has no useful numeric literal, so prove the three
    checkable components deterministically and leave their business association
    and the tie conclusion to the semantic reviewer.
    """
    required = ('q2_raw_cube_revenue', 'q2_close_revenue_adjustment', 'q2_revenue')
    missing: list[str] = []
    evidence: list[str] = []
    scopes = tuple(slide_numbers)
    for label in required:
        (present, detail) = _task_068_numeric_fact_present(path, label, float(answer[label]), scopes)
        if not present:
            missing.append(label)
        evidence.append(f'{label}: {detail}')
    return (not missing, f'missing current close-basis components={missing!r}; ' + '; '.join(evidence))

def _task_068_branch_primary_value_status(path: Path, label: str, expected: float) -> tuple[str, str]:
    """Return ``match``, ``conflict``, or ``absent`` for a BU scorecard cell.

    This is deliberately limited to an unambiguous row/column intersection.
    A recognized alternative prose or table layout remains eligible for the
    broad exact-magnitude gate plus semantic review.  When a primary scorecard
    cell is present, however, an unrelated number elsewhere on slide 5 cannot
    rescue a contradictory actual, plan, or variance.
    """
    match = re.fullmatch('(construction|service|controls)_q2_(.+)', label)
    if not match:
        return ('absent', 'not a business-unit scorecard criterion')
    (branch, metric) = match.groups()
    branch_aliases = {'construction': ('Construction',), 'service': ('Service',), 'controls': ('Controls', 'Building Controls')}[branch]

    def is_header(text: str) -> bool:
        normalized = _normalize(text)
        has_revenue = 'revenue' in normalized and 'ebitda' not in normalized
        has_ebitda = 'ebitda' in normalized
        has_plan = 'plan' in normalized
        has_variance = 'variance' in normalized or 'vs plan' in normalized
        if metric == 'revenue':
            return has_revenue and (not has_plan) and (not has_variance)
        if metric == 'approved_plan_revenue':
            return has_revenue and has_plan and (not has_variance)
        if metric == 'revenue_variance_to_plan':
            return has_revenue and has_variance or normalized == 'vs plan'
        if metric == 'adjusted_ebitda':
            return has_ebitda and (not has_plan) and (not has_variance) and ('%' not in text)
        if metric == 'approved_plan_adjusted_ebitda':
            return has_ebitda and has_plan and (not has_variance)
        return has_ebitda and has_variance
    presentation = Presentation(path)
    if len(presentation.slides) < 5:
        return ('absent', 'business-unit scorecard slide is absent')
    shapes = [shape for shape in presentation.slides[4].shapes if hasattr(shape, 'text') and str(shape.text or '').strip()]
    branch_shapes = []
    for shape in shapes:
        text = str(shape.text).strip()
        if len(text) > 80 or _TASK_068_NUMERIC_LITERAL_PATTERN.search(text):
            continue
        if any((_context_has_alias(text, alias) for alias in branch_aliases)):
            branch_shapes.append(shape)
    header_shapes = [shape for shape in shapes if is_header(str(shape.text))]
    row_tolerance = 180000
    saw_conflict = False
    conflict_context = ''
    for branch_shape in branch_shapes:
        row_shapes = [shape for shape in shapes if abs(int(shape.top) - int(branch_shape.top)) <= row_tolerance]
        for header in header_shapes:
            header_center = int(header.left) + int(header.width) // 2
            inherited_unit_headers: list[str] = []
            if metric == 'revenue_variance_to_plan':
                same_header_row = [shape for shape in shapes if abs(int(shape.top) - int(header.top)) <= row_tolerance and int(shape.left) < int(header.left) and ('revenue' in _normalize(str(shape.text))) and ('ebitda' not in _normalize(str(shape.text)))]
                if same_header_row:
                    inherited_unit_headers.append(str(max(same_header_row, key=lambda shape: int(shape.left)).text).strip())
            candidates = sorted((shape for shape in row_shapes if shape is not branch_shape and abs(int(shape.left) + int(shape.width) // 2 - header_center) <= 300000), key=lambda shape: abs(int(shape.left) + int(shape.width) // 2 - header_center))
            if not candidates:
                continue
            value_shape = candidates[0]
            context = ' | '.join((str(branch_shape.text).strip(), *inherited_unit_headers, str(header.text).strip(), str(value_shape.text).strip()))
            literals = _TASK_068_NUMERIC_LITERAL_PATTERN.findall(str(value_shape.text))
            if not literals:
                continue
            if any((_task_068_numeric_literal_matches(literal, expected, context=context) for literal in literals)):
                return ('match', f'primary business-unit scorecard cell matched: {context}')
            saw_conflict = True
            conflict_context = context
    if saw_conflict:
        return ('conflict', f'contradictory primary business-unit scorecard value cannot be rescued by an unrelated magnitude elsewhere: {conflict_context}')
    return ('absent', 'no unambiguous primary business-unit scorecard cell found')

def _task_068_branch_artifact_value(path: Path, label: str, expected: float) -> bool:
    """Read a business-unit scorecard without confusing actual, plan, and variance.

    Board decks commonly implement scorecards as aligned text boxes rather
    than native PowerPoint tables.  A branch name alone is not enough to bind
    every number in that row to every branch criterion.  This parser first
    accepts an explicit branch/metric/value statement, then uses the column
    heading and row alignment for grid-style scorecards.
    """
    match = re.fullmatch('(construction|service|controls)_q2_(.+)', label)
    if not match:
        return False
    (branch, metric) = match.groups()
    branch_aliases = {'construction': ('Construction',), 'service': ('Service',), 'controls': ('Controls', 'Building Controls')}[branch]
    metric_aliases = {'revenue': ('Q2 revenue', f'{branch} revenue'), 'approved_plan_revenue': ('approved plan revenue', 'plan revenue', f'{branch} plan revenue'), 'revenue_variance_to_plan': ('revenue vs plan', 'revenue variance to plan', f'{branch} revenue vs plan'), 'adjusted_ebitda': ('Q2 adjusted EBITDA', f'{branch} adjusted EBITDA', f'{branch} EBITDA'), 'approved_plan_adjusted_ebitda': ('approved plan adjusted EBITDA', 'plan adjusted EBITDA', 'plan EBITDA', f'{branch} plan EBITDA'), 'adjusted_ebitda_variance_to_plan': ('adjusted EBITDA vs plan', 'EBITDA variance to plan', f'{branch} EBITDA vs plan')}[metric]
    (primary_status, _primary_evidence) = _task_068_branch_primary_value_status(path, label, expected)
    if primary_status == 'match':
        return True
    if primary_status == 'conflict':
        return False
    contexts = _task_068_local_contexts(path, (5,))
    for context in contexts:
        if not any((_context_has_alias(context, alias) for alias in branch_aliases)):
            continue
        if not any((_context_has_alias(context, alias) for alias in metric_aliases)):
            continue
        if any((_task_068_numeric_literal_matches(literal, expected, context=context) for literal in _TASK_068_NUMERIC_LITERAL_PATTERN.findall(context))):
            return True
    return False

def _task_068_bridge_artifact_value(path: Path, label: str, expected: float) -> bool:
    """Bind a Plan/Actual waterfall value to its local Q2 bridge and unit."""
    mapping = {'q2_approved_plan_revenue': (3, 'revenue', 'plan'), 'q2_revenue': (3, 'revenue', 'actual'), 'q2_approved_plan_adjusted_ebitda': (4, 'ebitda', 'plan'), 'q2_adjusted_ebitda': (4, 'ebitda', 'actual')}
    if label not in mapping:
        return False
    (slide_number, metric, state) = mapping[label]
    presentation = Presentation(path)
    if len(presentation.slides) < slide_number:
        return False
    shapes = [shape for shape in presentation.slides[slide_number - 1].shapes if hasattr(shape, 'text') and str(shape.text or '').strip()]
    headings = [shape for shape in shapes if metric in _normalize(str(shape.text)) and 'bridge' in _normalize(str(shape.text))]
    labels = [shape for shape in shapes if _normalize(str(shape.text)) == state]
    for state_shape in labels:
        state_center = int(state_shape.left) + int(state_shape.width) // 2
        aligned = sorted((shape for shape in shapes if shape is not state_shape and abs(int(shape.left) + int(shape.width) // 2 - state_center) <= 300000 and (int(shape.top) < int(state_shape.top))), key=lambda shape: int(state_shape.top) - int(shape.top))
        for value_shape in aligned[:2]:
            for heading in headings:
                context = ' | '.join((str(heading.text).strip(), str(state_shape.text).strip(), str(value_shape.text).strip()))
                if any((_task_068_numeric_literal_matches(literal, expected, context=context) for literal in _TASK_068_NUMERIC_LITERAL_PATTERN.findall(context))):
                    return True
    return False

def _task_068_artifact_value(path: Path, label: str, expected: Any) -> bool:
    if re.fullmatch('(?:construction|service|controls)_q2_.+', label) and (not isinstance(expected, (bool, str))):
        return _task_068_branch_artifact_value(path, label, float(expected))
    if not isinstance(expected, (bool, str)) and _task_068_bridge_artifact_value(path, label, float(expected)):
        return True
    aliases = [label.replace('_', ' '), *_TASK_068_LABEL_ALIASES.get(label, ())]
    contexts = _task_068_local_contexts(path, _task_068_value_slides(label))
    if isinstance(expected, bool):
        if label == 'guidance_update_required' and expected is False:
            text = _normalize('\n'.join(contexts))
            maintained = any((phrase in text for phrase in ('no guidance update', 'maintain guidance', 'hold guidance', 'guidance unchanged', 'do not revise guidance')))
            contrary = any((phrase in text for phrase in ('guidance update required', 'update guidance now', 'revise guidance now', 'withdraw guidance now', 'reduce guidance now', 'cut guidance now')))
            return maintained and (not contrary)
        return any((any((_context_has_alias(context, alias) for alias in aliases)) and _boolean_matches(context, expected) for context in contexts))
    if isinstance(expected, str):
        return any((any((_context_has_alias(context, alias) for alias in aliases)) and _sample_directional_semantic_value_matches(context, expected) for context in contexts))
    if label == 'q2_posted_ytd_revenue_reconciliation_check' and float(expected) == 0.0:
        for context in contexts:
            if not any((_context_has_alias(context, alias) for alias in aliases)):
                continue
            normalized = _normalize(context)
            if any((phrase in normalized for phrase in ('no plug', 'zero difference', 'reconciles', 'ties', 'tied'))):
                return True
    if label in {'revenue_guidance_low', 'revenue_guidance_high', 'ebitda_guidance_low', 'ebitda_guidance_high'}:
        range_pattern = re.compile('\\$?\\s*(\\d[\\d,]*(?:\\.\\d+)?)\\s*[-–—]\\s*\\$?\\s*(\\d[\\d,]*(?:\\.\\d+)?)\\s*(k|m|mm|million|b|bn|billion)\\b', flags=re.I)
        scale = {'k': 1000.0, 'm': 1000000.0, 'mm': 1000000.0, 'million': 1000000.0, 'b': 1000000000.0, 'bn': 1000000000.0, 'billion': 1000000000.0}
        for context in contexts:
            if not any((_context_has_alias(context, alias) for alias in aliases)):
                continue
            for match in range_pattern.finditer(context):
                multiplier = scale[match.group(3).casefold()]
                values = (float(match.group(1).replace(',', '')) * multiplier, float(match.group(2).replace(',', '')) * multiplier)
                if any((_close(value, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0) for value in values)):
                    return True
    if label == 'minimum_guidance_update_headroom':
        for context in contexts:
            normalized = _normalize(context)
            if 'guidance' not in normalized or 'trigger' not in normalized:
                continue
            literals = _TASK_068_NUMERIC_LITERAL_PATTERN.findall(context)
            values: list[float] = []
            for literal in literals:
                cleaned = literal.replace('−', '-').strip()
                negative = cleaned.startswith('-') or (cleaned.startswith('(') and cleaned.endswith(')'))
                cleaned = cleaned.replace('$', '').replace(',', '').strip('() ')
                suffix = cleaned[-1:].casefold()
                multiplier = {'k': 1000.0, 'm': 1000000.0, 'b': 1000000000.0}.get(suffix, 1.0)
                if suffix in {'k', 'm', 'b'}:
                    cleaned = cleaned[:-1].strip()
                cleaned = cleaned.rstrip('%x×').strip()
                try:
                    value = float(cleaned) * multiplier
                except ValueError:
                    continue
                if negative:
                    value = -abs(value)
                if abs(value) >= 100000:
                    values.append(value)
            if any((_close(high - low, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0) for high in values for low in values if high > low)):
                return True
    adverse_magnitude_labels = {'construction_gross_profit_impact', 'service_labor_productivity_impact', 'controls_mix_impact', 'collections_timing_cash_impact'}
    for context in contexts:
        if not any((_context_has_alias(context, alias) for alias in aliases)):
            continue
        literals = _TASK_068_NUMERIC_LITERAL_PATTERN.findall(context)
        if label == 'remaining_fy26_revenue_outlook':
            normalized = _normalize(context)
            explanatory_need = 'remaining revenue need' in normalized or 'remaining fy26 revenue need' in normalized
            if explanatory_need and any(('(' in literal and ')' in literal and _task_068_numeric_literal_matches(literal, -abs(float(expected)), context=context) for literal in literals)):
                return True
        if label == 'q2_capital_expenditure' and any((_task_068_numeric_literal_matches(literal, abs(float(expected)), context=context) or _task_068_numeric_literal_matches(literal, -abs(float(expected)), context=context) for literal in literals)):
            return True
        if label in adverse_magnitude_labels:
            normalized = _normalize(context)
            adverse = any((term in normalized for term in ('unfavorable', 'unfav', 'adverse', 'drag', 'miss', 'risk', 'erosion', 'pressure', 'timing')))
            recovery = any((term in normalized for term in ('recovery', 'action', 'mitigation', 'opportunity', 'action pool')))
            if adverse and any((_task_068_numeric_literal_matches(literal, -abs(float(expected)), context=context) or (_task_068_numeric_literal_matches(literal, abs(float(expected)), context=context) and (not any((sign in literal for sign in ('+', '-', '−', '('))))) for literal in literals)):
                return True
            if recovery and any((_task_068_numeric_literal_matches(literal, abs(float(expected)), context=context) for literal in literals)):
                return True
            continue
        if any((_task_068_numeric_literal_matches(literal, float(expected), context=context) for literal in literals)):
            return True
    return False

def _preserved_sheets(output, seed, names: list[str]) -> tuple[bool, str]:
    for name in names:
        if name not in output.sheetnames or name not in seed.sheetnames:
            return (False, f'missing protected sheet {name!r}')
        (left, right) = (output[name], seed[name])
        max_row = max(left.max_row, right.max_row)
        max_col = max(left.max_column, right.max_column)
        for row in range(1, max_row + 1):
            for column in range(1, max_col + 1):
                left_value = left.cell(row, column).value
                right_value = right.cell(row, column).value
                if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
                    same = math.isclose(float(left_value), float(right_value), rel_tol=1e-12, abs_tol=1e-09)
                else:
                    same = left_value == right_value
                if not same:
                    return (False, f'protected source changed at {name}!{left.cell(row, column).coordinate}')
    return (True, 'protected source sheets match the seeded template')

def _task_035_has_authored_work(output, seed) -> bool:
    """Distinguish substantive workbook edits from a copied/resaved starter.

    The three reference schedules are intentionally pre-populated.  They and
    the static output labels must not earn reward when an agent returns the
    untouched starter (including a metadata-only resave).  A changed cell on
    any working/output sheet, or a genuinely added sheet, is authored work.
    """
    output_sheets = list(_task_035_output_sheet_names(output))
    if any((name not in seed.sheetnames for name in output_sheets)):
        return True
    for name in output_sheets:
        left = output[name]
        right = seed[name]
        max_row = max(left.max_row, right.max_row)
        max_col = max(left.max_column, right.max_column)
        for row in range(1, max_row + 1):
            for column in range(1, max_col + 1):
                if left.cell(row, column).value != right.cell(row, column).value:
                    return True
    return False

def _file_failure(gold: dict[str, Any], reason: str) -> dict[str, Any]:
    result = _result([Criterion(spec['id'], spec['description'], False, reason) for spec in gold['criteria']])
    return _attach_gold_policy(result, gold)

def _attach_gold_policy(result: dict[str, Any], gold: dict[str, Any]) -> dict[str, Any]:
    by_spec = {str(spec['id']): spec for spec in gold['criteria']}
    for row in result.get('criteria', []):
        spec = by_spec[str(row['id'])]
        for key in ('category', 'weight', 'failure_cap', 'semantic', 'kind'):
            if key in spec:
                row[key] = spec[key]
    return result

def _headline_formula_match(task_id: str, workbook, values, label: str, gold: dict[str, Any]) -> tuple[bool, str]:
    candidates = [label.replace('_', ' ')]
    if task_id == 'task_035':
        candidates.extend(_TASK_035_LABEL_ALIASES.get(label, []))
    attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
    (matched, detail) = next((attempt for attempt in attempts if attempt[0]), attempts[0])
    if task_id == 'task_035' and (not matched):
        (matched, detail) = _task_035_row_match(workbook, values, label, gold['answer'][label], require_formula=True)
    return (matched, detail)

def _grade_artifact(task_id: str, workspace_root: Path) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    relative = Path(gold['artifact']['path'])
    path = workspace_root / relative
    if not path.is_file():
        return _file_failure(gold, f'missing required artifact: {relative}')
    try:
        entries = _artifact_text_entries(path)
        normalized_text = _normalize('\n'.join(entries))
        workbook = values = None
        if path.suffix.lower() == '.xlsx':
            workbook = load_workbook(path, data_only=False, read_only=False)
            values = _recalculated_data_workbook(path)
            if task_id == 'task_035':
                seed_path = SEED_WORKSPACE / relative
                if seed_path.is_file():
                    seed = load_workbook(seed_path, data_only=False, read_only=False)
                    if not _task_035_has_authored_work(workbook, seed):
                        return _file_failure(gold, 'untouched or metadata-only Task035 starter; no authored working-model cells')
        elif task_id == 'task_068':
            seed_path = SEED_WORKSPACE / relative
            if seed_path.is_file() and (not _task_068_slides_changed(path, gold['artifact']['path'], range(2, 10))):
                return _file_failure(gold, 'untouched or metadata-only Task068 starter; no authored executive-slide content')
    except Exception as exc:
        return _file_failure(gold, f'unreadable artifact: {exc}')
    criteria: list[Criterion] = []
    for spec in gold['criteria']:
        kind = spec['kind']
        evidence = ''
        if kind == 'xlsx_sheet_present':
            met = workbook is not None and spec['sheet'] in workbook.sheetnames
            evidence = f"required_sheet={spec['sheet']!r}; sheets={getattr(workbook, 'sheetnames', [])!r}"
        elif kind == 'xlsx_sheet_preserved':
            seed_path = SEED_WORKSPACE / relative
            if workbook is None or not seed_path.is_file():
                met = False
                evidence = 'submitted workbook or seeded edit template is missing'
            else:
                seed = load_workbook(seed_path, data_only=False, read_only=False)
                (met, evidence) = _preserved_sheets(workbook, seed, [spec['sheet']])
        elif kind == 'xlsx_formula_count':
            formulas = sum((1 for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')))
            met = formulas >= int(spec['min_formulas'])
            evidence = f"formulas={formulas}; required={spec['min_formulas']}"
        elif kind == 'xlsx_no_errors':
            errors = [f'{sheet.title}!{cell.coordinate}={cell.value}' for sheet in values.worksheets for row in sheet.iter_rows() for cell in row if cell.data_type == 'e']
            met = not errors
            evidence = f'errors={errors[:10]!r}'
        elif kind == 'xlsx_headline_formula':
            (met, evidence) = _headline_formula_match(task_id, workbook, values, str(spec['headline_label']), gold)
        elif kind == 'xlsx_formula_sheet':
            sheet = workbook[spec['sheet']] if spec['sheet'] in workbook.sheetnames else None
            formulas = 0 if sheet is None else sum((1 for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')))
            met = formulas > 0
            evidence = f"sheet={spec['sheet']!r}; formulas={formulas}"
        elif kind == 'xlsx_cross_sheet_formulas':
            formulas = [cell.value for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')]
            cross_sheet = sum((1 for formula in formulas if re.search("(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\\$?[A-Z]{1,3}\\$?\\d+", formula)))
            met = cross_sheet >= int(spec['min_cross_sheet_formulas'])
            evidence = f"cross_sheet={cross_sheet}; required={spec['min_cross_sheet_formulas']}"
        elif kind == 'xlsx_sheet_lineage':
            target_name = str(spec['target_sheet'])
            source_name = str(spec['source_sheet'])
            (met, evidence) = _xlsx_sheet_lineage(workbook, target_name, source_name)
        elif kind == 'pptx_slide_count':
            presentation = Presentation(path)
            expected = spec.get('exact_slides')
            met = len(presentation.slides) == int(expected) if expected is not None else len(presentation.slides) >= int(spec['min_slides'])
            evidence = f"slides={len(presentation.slides)}; required={expected or spec.get('min_slides')}"
        elif kind == 'pptx_title':
            met = contains_concept('\n'.join(entries), spec['title_token'])
            evidence = f"title_concept={spec['title_token']!r}; present={met}"
        elif kind == 'docx_heading':
            body = '\n'.join((paragraph.text for paragraph in Document(path).paragraphs))
            met = contains_concept(body, spec['heading'])
            evidence = f"heading_concept={spec['heading']!r}; present={met}"
        elif kind == 'docx_tables':
            tables = len(Document(path).tables)
            met = tables >= int(spec['min_tables'])
            evidence = f"tables={tables}; required={spec['min_tables']}"
        elif kind == 'xlsx_structure':
            required = spec['sheets']
            sheets_ok = workbook is not None and all((name in workbook.sheetnames for name in required))
            (preserve_ok, preserve_evidence) = (True, 'not an edit task')
            if sheets_ok and spec.get('preserve_source_sheets'):
                seed_path = SEED_WORKSPACE / relative
                if not seed_path.is_file():
                    (preserve_ok, preserve_evidence) = (False, 'seeded edit template is missing')
                else:
                    seed = load_workbook(seed_path, data_only=False, read_only=False)
                    (preserve_ok, preserve_evidence) = _preserved_sheets(workbook, seed, spec['preserve_source_sheets'])
            met = sheets_ok and preserve_ok
            evidence = f"sheets={getattr(workbook, 'sheetnames', [])!r}; {preserve_evidence}"
        elif kind == 'xlsx_model_integrity':
            formulas = sum((1 for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')))
            errors = [f'{sheet.title}!{cell.coordinate}={cell.value}' for sheet in values.worksheets for row in sheet.iter_rows() for cell in row if cell.data_type == 'e']
            met = formulas >= spec['min_formulas'] and (not errors)
            evidence = f'formulas={formulas}; errors={errors[:5]!r}'
        elif kind == 'xlsx_formula_lineage':
            by_sheet = {sheet.title: sum((1 for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('='))) for sheet in workbook.worksheets}
            all_formulas = [cell.value for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=')]
            cross_sheet = sum((1 for formula in all_formulas if re.search("(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\\$?[A-Z]{1,3}\\$?\\d+", formula)))
            missing_formula_sheets = [name for name in spec['formula_sheets'] if by_sheet.get(name, 0) == 0]
            headline_failures = []
            headline_evidence = []
            for label in spec['headline_labels']:
                candidates = [label.replace('_', ' ')]
                if task_id == 'task_035':
                    candidates.extend(_TASK_035_LABEL_ALIASES.get(label, []))
                attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
                (matched, detail) = next((attempt for attempt in attempts if attempt[0]), attempts[0])
                if task_id == 'task_035' and (not matched):
                    expected = gold['answer'][label]
                    (matched, detail) = _task_035_row_match(workbook, values, label, expected, require_formula=True)
                if not matched:
                    headline_failures.append(label)
                else:
                    headline_evidence.append(detail)
            met = not missing_formula_sheets and (not headline_failures) and (cross_sheet >= spec['min_cross_sheet_formulas'])
            evidence = f'formula_sheets={by_sheet!r}; cross_sheet={cross_sheet}; missing_formula_sheets={missing_formula_sheets!r}; headline_failures={headline_failures!r}; headline_examples={headline_evidence[:3]!r}'
        elif kind == 'xlsx_label_values':
            failures = []
            for (label, expected) in spec['label_values'].items():
                labels = [label.replace('_', ' ')]
                if task_id == 'task_035':
                    labels.extend(_TASK_035_LABEL_ALIASES.get(label, []))
                matched = any((_xlsx_label_value(workbook, values, candidate, expected, directional_strings=task_id in {'task_035'}) for candidate in labels))
                if task_id == 'task_035' and (not matched):
                    (matched, _) = _task_035_row_match(workbook, values, label, expected, require_formula=False)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = 'all labeled values matched' if met else f'missing or incorrect labels={failures!r}'
        elif kind == 'artifact_label_values':
            failures = []
            for (label, expected) in spec['label_values'].items():
                labels = [label.replace('_', ' ')]
                if task_id in {'task_068'}:
                    labels.extend(_TASK_068_COMPATIBILITY_LABEL_ALIASES.get(label, []))
                matched = _task_068_artifact_value(path, label, expected) if task_id == 'task_068' else any((_artifact_label_value(entries, candidate, expected, directional_strings=False) for candidate in labels))
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = 'all labeled values matched' if met else f'missing or incorrect labels={failures!r}'
        elif kind == 'artifact_tokens':
            if task_id == 'task_035' and spec['id'].startswith(('model_content__', 'controls__')):
                (met, evidence) = _task_035_model_or_control_check(workbook, values, spec['id'])
            elif task_id == 'task_035' and spec['id'].startswith('sources__'):
                token = spec.get('tokens', [''])[0]
                source_reference = _TASK_035_SOURCE_REFERENCES.get(token, {})
                (_submitted, met, evidence) = _scoped_xlsx_token_evidence(workbook, values, token, aliases=source_reference.get('aliases', ()), sheet_names=_task_035_output_sheet_names(workbook))
            else:
                missing = [token for token in spec['tokens'] if not any((contains_concept(normalized_text, candidate) for candidate in _ARTIFACT_TOKEN_ALIASES.get(token, [token])))]
                met = not missing
                evidence = 'all required content present' if met else f'missing={missing!r}'
        elif kind == 'pptx_structure':
            presentation = Presentation(path)
            expected_slides = spec.get('exact_slides')
            count_ok = len(presentation.slides) == expected_slides if expected_slides is not None else len(presentation.slides) >= spec['min_slides']
            joined_entries = '\n'.join(entries)
            title_present = contains_concept(joined_entries, spec['title_token'])
            met = count_ok and title_present
            evidence = f'slides={len(presentation.slides)}; title_present={title_present}'
        elif kind == 'docx_structure':
            document = Document(path)
            body = _normalize('\n'.join((paragraph.text for paragraph in document.paragraphs)))
            missing = [heading for heading in spec['headings'] if not contains_concept(body, heading)]
            met = not missing and len(document.tables) >= spec['min_tables']
            evidence = f'missing_headings={missing!r}; tables={len(document.tables)}'
        else:
            raise ValueError(f'Unsupported artifact criterion: {kind}')
        criteria.append(Criterion(spec['id'], spec['description'], met, evidence))
    result = _attach_gold_policy(_result(criteria), gold)
    semantic_review = _hybrid_semantic_review(task_id, gold, path, workbook, values, criteria)
    if semantic_review is not None:
        result['semantic_review'] = semantic_review
    return result

def grade_corporate_finance_task(task_id: str, answer: Any, workspace_root: str | Path) -> dict[str, Any]:
    if task_id not in SAMPLE_CORPORATE_TASK_IDS:
        raise KeyError(f'Task is not part of this sample: {task_id}')
    return _grade_artifact(task_id, Path(workspace_root))
