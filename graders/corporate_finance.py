from __future__ import annotations
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation
from graders.apex import Criterion, SEED_WORKSPACE, _answer_mapping, _close, _get, _normalize, _number, _recalculated_data_workbook, _result
from graders.semantic import contains_concept, date_matches, iter_numeric_candidates, ordered_semantic_list_matches, semantic_equal, semantic_value_matches, unordered_semantic_list_matches
from graders.hybrid_semantic import semantic_requirement
GOLD_PATH = Path(__file__).resolve().parent / 'gold' / 'tasks_026_100.json'
_TASK_037_LABEL_ALIASES = {'selected_portfolio': ['Selected ID'], 'selected_capex': ['Cash capex'], 'portfolio_npv': ['Selected NPV'], 'downside_portfolio_npv': ['Downside portfolio NPV']}
_TASK_055_LABEL_ALIASES = {'buyer_standalone_eps': ['Standalone diluted EPS'], 'seller_shares_issued': ['Seller shares issued (equity ÷ price)'], 'pro_forma_diluted_shares': ['Pro forma diluted shares'], 'target_ebit': ['Target EBIT'], 'incremental_debt_interest': ['Incremental new-debt interest'], 'foregone_cash_yield': ['Foregone cash yield (interest)'], 'total_incremental_financing_cost': ['Total incremental financing cost'], 'year_one_gaap_eps_accretion': ['GAAP EPS accretion / (dilution) %'], 'year_one_adjusted_eps_accretion': ['Adjusted EPS accretion / (dilution) %'], 'year_two_gaap_eps_accretion': ['GAAP EPS accretion / (dilution) %'], 'year_two_adjusted_eps_accretion': ['Adjusted EPS accretion / (dilution) %']}
_TASK_055_YEAR_ROW_ALIASES = {'run_rate_synergy': ['Cost synergies', 'Run-rate synergy'], 'integration_expense': ['One-time integration expense'], 'gaap_incremental_pre_tax_income': ['Target-side pre-tax income (GAAP)'], 'gaap_incremental_after_tax_income': ['Target-side after-tax (GAAP)'], 'adjusted_incremental_pre_tax_income': ['Target-side pre-tax income (Adjusted)'], 'adjusted_incremental_after_tax_income': ['Target-side after-tax (Adjusted)'], 'gaap_pro_forma_net_income': ['Combined net income (GAAP)'], 'adjusted_pro_forma_net_income': ['Combined net income (Adjusted)'], 'gaap_pro_forma_eps': ['Combined GAAP EPS'], 'adjusted_pro_forma_eps': ['Combined adjusted EPS']}
_TASK_061_LABEL_ALIASES = {'ending_temporary_difference_dta': ['temporary_difference_dta'], 'ending_state_credit_dta': ['state_credit_dta'], 'ending_valuation_allowance': ['valuation_allowance'], 'ending_net_dta': ['net_dta'], 'ending_net_deferred_tax_liability': ['net_deferred_tax_liability']}
_TASK_072_LABEL_ALIASES = {'q2_revenue': ['Q2 revenue', 'Board Performance Q2 revenue'], 'q2_approved_plan_revenue': ['Q2 revenue plan', 'Approved Q2 plan revenue'], 'q2_revenue_variance_to_plan': ['Q2 revenue variance to plan', 'Revenue variance vs plan', 'Revenue variance to plan'], 'q2_organic_growth': ['Organic growth'], 'q2_adjusted_ebitda': ['Adjusted EBITDA'], 'q2_approved_plan_adjusted_ebitda': ['Q2 adjusted EBITDA plan', 'Approved Q2 plan adjusted EBITDA'], 'q2_adjusted_ebitda_variance_to_plan': ['Q2 adjusted EBITDA variance to plan', 'Adjusted EBITDA variance vs plan', 'EBITDA variance to plan'], 'construction_gross_profit_impact': ['Construction gross profit impact', 'Construction gross profit downside'], 'service_labor_productivity_impact': ['Service labor productivity impact', 'Service labor downside'], 'controls_mix_impact': ['Controls mix impact', 'Controls mix downside'], 'q2_free_cash_flow': ['Free cash flow'], 'maximum_revolver': ['Maximum downside revolver', 'FY27 downside maximum revolver', 'Downside revolver'], 'latest_full_year_revenue_outlook': ['Latest approved outlook', 'Revenue outlook'], 'ltm_adjusted_ebitda': ['LTM adjusted EBITDA', 'Approved LTM adjusted EBITDA', 'Lender adjusted EBITDA'], 'funded_debt': ['Posted funded debt', 'Lender funded debt'], 'lender_leverage': ['Lender leverage', 'Covenant leverage', 'Gross leverage'], 'fixed_charge_coverage': ['Fixed-charge coverage', 'Fixed charge coverage', 'FCCR', 'Lender FCCR'], 'revenue_guidance_low': ['Revenue guidance low', 'Published revenue guidance low'], 'revenue_guidance_high': ['Revenue guidance high', 'Published revenue guidance high'], 'ebitda_guidance_low': ['EBITDA guidance low', 'Published EBITDA guidance low'], 'ebitda_guidance_high': ['EBITDA guidance high', 'Published EBITDA guidance high'], 'largest_downside_driver': ['Largest downside driver', 'Principal downside driver'], 'combined_stress_ebitda': ['Combined stress EBITDA', 'Signed downside stress EBITDA'], 'ebitda_shortfall_to_guidance_low': ['EBITDA shortfall to guidance low', 'Shortfall to published EBITDA low end'], 'additional_revenue_decline_to_update_trigger': ['Additional revenue decline to update trigger', 'Headroom to formal update trigger'], 'guidance_update_required': ['Guidance update required', 'Revise published guidance']}
_TASK_100_LABEL_ALIASES = {'forecast_revenue': ['FY26 outlook', 'FY26 forecast revenue', 'revenue'], 'revenue_variance_to_board_aop': ['vs Board AOP', 'versus Board AOP', 'Board AOP'], 'revenue_variance_pct_to_board_aop': ['vs Board AOP', 'versus Board AOP', 'Board AOP'], 'forecast_branch_ebitda': ['FY26 forecast branch EBITDA', 'forecast branch EBITDA', 'branch EBITDA'], 'branch_ebitda_variance_to_current_comparator': ['branch EBITDA variance', 'branch EBITDA gain', 'branch economics', 'vs current comparator'], 'largest_branch_ebitda_miss': ['largest branch EBITDA miss', 'largest negative branch variance'], 'largest_branch_ebitda_miss_amount': ['largest branch EBITDA miss', 'largest negative branch variance', 'Controls'], 'downside_leverage': ['Downside funded-debt leverage', 'Downside leverage', 'Downside'], 'severe_leverage': ['Severe funded-debt leverage', 'Severe leverage', 'Severe'], 'downside_ending_cash': ['Downside ending cash', 'Downside cash', 'Downside'], 'severe_ending_cash': ['Severe ending cash', 'Severe cash', 'Severe'], 'severe_revolver_headroom': ['Severe revolver headroom', 'Severe headroom', 'Severe'], 'total_probability_weighted_priority_ebitda': ['total probability-weighted benefit', 'total priority benefit', 'PW benefit', 'unconstrained PW benefit', 'all priorities', 'full priority list', 'priority value'], 'executable_probability_weighted_priority_ebitda': ['executable', 'Approved or Gated', 'Approved plus Gated', 'Approved + Gated PW'], 'severe_covenant_breach': ['Severe covenant breach', 'breach flags'], 'optimized_board_priority_portfolio': ['Optimized priority portfolio', 'Selected board priorities', 'Selected portfolio'], 'optimized_priority_cash_spend': ['Selected priority spend', 'Optimized cash spend', 'Selected cash', 'Budget'], 'optimized_probability_weighted_base_benefit': ['Optimized Base benefit', 'Selected probability-weighted benefit', 'PW Base benefit', 'Approved + Gated', 'Executable FY27 portfolio'], 'optimized_probability_weighted_severe_benefit': ['Optimized Severe benefit', 'Selected Severe benefit', 'Severe benefit'], 'post_action_severe_incremental_revolver_draw': ['Post-action incremental revolver draw', 'Severe incremental draw', 'Incremental revolver funding', 'Incremental draw'], 'post_action_severe_remaining_revolver_headroom': ['Post-action revolver headroom', 'Remaining Severe revolver headroom', 'Remaining revolver capacity', 'Remaining headroom'], 'post_action_severe_unfunded_liquidity_shortfall': ['Post-action unfunded shortfall', 'Severe unfunded liquidity', 'Residual liquidity shortfall'], 'post_action_severe_ending_cash': ['Post-action Severe cash', 'Severe cash after priorities'], 'post_action_severe_leverage': ['Post-action Severe leverage', 'Severe leverage after priorities'], 'post_action_severe_cash_breach': ['Post-action Severe cash breach', 'Severe cash gate'], 'post_action_severe_leverage_breach': ['Post-action Severe leverage breach', 'Severe leverage gate']}
_ARTIFACT_TOKEN_ALIASES = {'management correspondence': ['management correspondence', 'management email', 'request thread', 'Requests/', 'controller follow-up', 'source-tie thread', 'board review', 'review notes'], 'contractor accounting mcp': ['contractor accounting mcp', 'vista erp', 'posted vista', 'company accounting records'], 'executive summary': ['executive summary', 'executive performance summary', 'performance summary'], 'cash roll forward': ['cash roll forward', 'cash roll-forward', 'cash liquidity scenario', 'ending cash before financing', 'cash & liquidity'], 'risk': ['risk', 'downside', 'guardrail', 'failure', 'shortfall'], 'funding gap': ['funding gap', 'unfunded gap', 'funds flow check', 'excess shortfall', 'funds exactly', 'sources less uses', 'source use difference'], 'version': ['version', 'source version', 'document version', 'file version', 'v4', 'v6', 'v7.4', 'v7.6', 'v8', 'current approved', 'current source', 'controller-tied', 'controlling case', 'current - controller tie complete'], 'working capital peg': ['working capital peg', 'NWC peg', 'median NWC peg', 'monthly NWC'], 'unit': ['unit', 'units', 'USD in thousands', 'USD thousands', '$000', '$mm', 'USD millions'], 'cash credit': ['cash credit', 'usable cash', 'usable target cash', 'target cash credit'], 'locked box': ['locked box', 'locked-box', 'unauthorized leakage', 'permitted leakage', 'leakage'], 'investor leverage': ['investor leverage', 'investor net leverage'], 'covenant leverage': ['covenant leverage', 'covenant net leverage', 'covenant-basis net leverage', 'lender leverage'], 'model status': ['model status', 'overall status', 'control status', 'feasibility', 'model integrity']}

def load_corporate_finance_gold(task_id: str | None=None) -> dict[str, Any]:
    payload = json.loads(GOLD_PATH.read_text(encoding='utf-8'))
    return payload[task_id] if task_id else payload

def _numeric_matches(actual: Any, expected: float, abs_tol: float) -> bool:
    return any((_close(candidate, expected, abs_tol=abs_tol, rel_tol=0.0) for candidate in iter_numeric_candidates(actual)))

def _numeric_matches_spec(actual: Any, expected: float, spec: dict[str, Any]) -> bool:
    abs_tol = float(spec.get('abs_tol', 0.02))
    aggregate_field = spec.get('aggregate_field')
    if aggregate_field and isinstance(actual, (list, tuple)):
        aggregate_aliases = [aggregate_field, *spec.get('aggregate_field_aliases', [])]
        amounts = []
        for item in actual:
            if not isinstance(item, dict):
                continue
            raw = next((value for candidate in aggregate_aliases for (key, value) in item.items() if semantic_equal(key, candidate)), None)
            if isinstance(raw, (int, float)) and (not isinstance(raw, bool)):
                amounts.append(float(raw))
        concepts = list(spec.get('expected_concepts', []))
        concepts_match = not concepts or (unordered_semantic_list_matches(actual, concepts) if spec.get('unordered_aggregate_concepts') else ordered_semantic_list_matches(actual, concepts))
        if len(amounts) == len(actual) and amounts and concepts_match and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0):
            return True
    if spec.get('sum_numeric_sequence') and isinstance(actual, (list, tuple)):
        amounts = [float(item) for item in actual if isinstance(item, (int, float)) and (not isinstance(item, bool))]
        if len(amounts) == len(actual) and amounts and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0):
            return True
    if spec.get('aggregate_structured_amounts') and isinstance(actual, (list, tuple)):
        amounts: list[float] = []
        for item in actual:
            if not isinstance(item, dict):
                return False
            raw = next((item[key] for key in ('amount', 'value') if key in item), None)
            if not isinstance(raw, (int, float)) or isinstance(raw, bool):
                return False
            amounts.append(float(raw))
        concepts = list(spec.get('expected_concepts', []))
        return len(amounts) == len(concepts) and ordered_semantic_list_matches(actual, concepts) and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0)
    if _numeric_matches(actual, expected, abs_tol):
        return True
    for alternative in spec.get('numeric_alternatives', []):
        if _numeric_matches(actual, float(alternative), abs_tol):
            return True
    denominator = spec.get('ratio_alternative_denominator')
    if denominator:
        ratio = expected / float(denominator)
        ratio_tol = float(spec.get('ratio_alternative_abs_tol', 0.0001))
        if any((_close(candidate, ratio, abs_tol=ratio_tol, rel_tol=0.0) for candidate in iter_numeric_candidates(actual))):
            return True
    decimals = spec.get('display_decimals')
    if decimals is not None:
        tolerance = 0.5 * 10 ** (-int(decimals)) + 1e-12
        return any((_close(candidate, expected, abs_tol=max(abs_tol, tolerance), rel_tol=0.0) for candidate in iter_numeric_candidates(actual)))
    return False

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

def _structured_value_matches(actual: Any, expected: Any) -> bool:
    """Match nested disclosed JSON values without stringifying containers."""
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            return False
        return all(((found := _semantic_find(actual, str(key)))[0] and _structured_value_matches(found[1], value) for (key, value) in expected.items()))
    if isinstance(expected, list):
        return isinstance(actual, list) and len(actual) == len(expected) and all((_structured_value_matches(actual_value, expected_value) for (actual_value, expected_value) in zip(actual, expected, strict=True)))
    if isinstance(expected, bool):
        return _boolean_matches(actual, expected)
    if isinstance(expected, (int, float)) and (not isinstance(expected, bool)):
        return _numeric_matches(actual, float(expected), 0.02)
    if expected is None:
        return actual is None or _normalize(actual) in {'none', 'null', 'n a'}
    return semantic_value_matches(actual, str(expected))

def _semantic_get(mapping: dict[str, Any], key: str) -> Any:
    """Prefer the requested key, then accept harmless label morphology."""
    actual = _get(mapping, key)
    if actual is not None:
        return actual
    for (candidate, value) in mapping.items():
        if semantic_equal(candidate, key):
            return value
    return None

def _semantic_find(mapping: dict[str, Any], key: str) -> tuple[bool, Any]:
    if key in mapping:
        return (True, mapping[key])
    for (candidate, value) in mapping.items():
        if semantic_equal(candidate, key):
            return (True, value)
    return (False, None)

def _structured_row(mapping: dict[str, Any], *, key: str, row_key: str, row_value: str) -> dict[str, Any] | None:
    rows = _semantic_get(mapping, key)
    if not isinstance(rows, list):
        return None
    for item in rows:
        if not isinstance(item, dict):
            continue
        actual_row_value = _semantic_get(item, row_key)
        if semantic_value_matches(actual_row_value, row_value):
            return item
    return None

def _structured_field(row: dict[str, Any] | None, spec: dict[str, Any]) -> Any:
    """Read a disclosed structured field while honoring narrow rubric aliases."""
    if row is None:
        return None
    for field in (str(spec['field']), *[str(value) for value in spec.get('field_aliases', [])]):
        (present, value) = _semantic_find(row, field)
        if present:
            return value
    return None

def _period_label_matches(actual: Any, expected: Any) -> bool:
    """Accept a bare ordinal only when the output key already supplies the period type."""

    def parse(value: Any) -> tuple[str | None, int] | None:
        if isinstance(value, (int, float)) and (not isinstance(value, bool)) and float(value).is_integer():
            return (None, int(value))
        normalized = _normalize(value)
        if re.fullmatch('\\d+', normalized):
            return (None, int(normalized))
        match = re.fullmatch('(?:fiscal )?(year|yr|y|month|mo|week|wk|quarter|q)\\s*-?\\s*(\\d+)', normalized)
        if not match:
            return None
        period = {'yr': 'year', 'y': 'year', 'mo': 'month', 'wk': 'week', 'q': 'quarter'}.get(match.group(1), match.group(1))
        return (period, int(match.group(2)))
    (left, right) = (parse(actual), parse(expected))
    return bool(left and right and (left[1] == right[1]) and (left[0] is None or right[0] is None or left[0] == right[0]))

def _console_criterion(mapping: dict[str, Any], spec: dict[str, Any]) -> Criterion:
    kind = spec['kind']
    if kind == 'numeric':
        failures = []
        for (key, expected) in spec['expected'].items():
            actual = _semantic_get(mapping, key)
            if not _numeric_matches_spec(actual, float(expected), spec):
                failures.append(f'{key}={actual!r}; expected {expected}')
        met = not failures
        evidence = 'reported value matched' if met else '; '.join(failures)
    elif kind == 'string':
        failures = []
        for (key, expected) in spec['expected'].items():
            (present, actual) = _semantic_find(mapping, key)
            expected_none = expected is None or _normalize(expected) in {'none', 'null'}
            actual_none = actual is None or _normalize(actual) in {'none', 'no draw', 'no breach', 'not applicable', 'n a'}
            typed_period = key.endswith(('_year', '_month', '_week', '_quarter')) and _period_label_matches(actual, expected)
            numeric_alternative = False
            if present and 'numeric_alternative' in spec:
                numeric_alternative = _numeric_matches_spec(actual, float(spec['numeric_alternative']), {'abs_tol': float(spec.get('numeric_alternative_abs_tol', 0.02))})
            if not present or not (expected_none and actual_none or typed_period or numeric_alternative or semantic_value_matches(actual, str(expected))):
                failures.append(f'{key}={actual!r}; expected {expected!r}')
        met = not failures
        evidence = 'classification matched' if met else '; '.join(failures)
    elif kind == 'boolean':
        failures = []
        for (key, expected) in spec['expected'].items():
            actual = _semantic_get(mapping, key)
            if not _boolean_matches(actual, bool(expected)):
                failures.append(f'{key}={actual!r}; expected {expected!r}')
        met = not failures
        evidence = 'boolean conclusion matched' if met else '; '.join(failures)
    elif kind == 'list_item':
        actual = _semantic_get(mapping, spec['key'])
        expected = spec['expected_item']
        if isinstance(actual, dict):
            flattened = [item for (key, value) in actual.items() for item in (key, value)]
        elif isinstance(actual, (list, tuple, set)):
            flattened = list(actual)
        elif actual is not None:
            flattened = [actual]
        else:
            flattened = []
        present_in_required_bucket = any((_structured_value_matches(item, expected) if isinstance(expected, (dict, list)) else semantic_value_matches(item, str(expected)) for item in flattened))
        conflicting_matches = {}
        for conflicting_key in spec.get('exclusive_with_keys', []):
            conflicting_actual = _semantic_get(mapping, conflicting_key)
            if isinstance(conflicting_actual, dict):
                conflicting_items = [item for (key, value) in conflicting_actual.items() for item in (key, value)]
            elif isinstance(conflicting_actual, (list, tuple, set)):
                conflicting_items = list(conflicting_actual)
            elif conflicting_actual is not None:
                conflicting_items = [conflicting_actual]
            else:
                conflicting_items = []
            if any((semantic_value_matches(item, str(expected)) for item in conflicting_items)):
                conflicting_matches[str(conflicting_key)] = conflicting_actual
        met = present_in_required_bucket and (not conflicting_matches)
        evidence = f'actual={actual!r}; required_item={expected!r}; conflicting_matches={conflicting_matches!r}'
    elif kind in {'list', 'list_exact'}:
        actual = _semantic_get(mapping, spec['key'])
        expected_values = list(spec['expected'])
        if any((isinstance(value, (dict, list)) for value in expected_values)):
            if spec.get('unordered'):
                remaining = list(actual) if isinstance(actual, list) else []
                met = len(remaining) == len(expected_values)
                for expected_value in expected_values:
                    matched_index = next((index for (index, actual_value) in enumerate(remaining) if _structured_value_matches(actual_value, expected_value)), None)
                    if matched_index is None:
                        met = False
                        break
                    remaining.pop(matched_index)
                met = met and (not remaining)
                expectation = 'unordered structured concepts'
            else:
                met = isinstance(actual, list) and len(actual) == len(expected_values) and all((_structured_value_matches(actual_value, expected_value) for (actual_value, expected_value) in zip(actual, expected_values, strict=True)))
                expectation = 'ordered structured concepts'
        elif spec.get('unordered'):
            met = unordered_semantic_list_matches(actual, expected_values)
            expectation = 'unordered concepts'
        else:
            met = isinstance(actual, list) and (not actual) if not expected_values else ordered_semantic_list_matches(actual, expected_values)
            expectation = 'ordered concepts'
        evidence = f'actual={actual!r}; expected {expectation}={expected_values!r}'
    elif kind == 'structured_row_order':
        actual = _semantic_get(mapping, spec['key'])
        actual_rows = []
        if isinstance(actual, list):
            actual_rows = [_semantic_get(item, spec['row_key']) for item in actual if isinstance(item, dict)]
        expected_rows = list(spec['expected_rows'])
        met = len(actual_rows) == len(expected_rows) and all((semantic_value_matches(actual_value, expected_value) for (actual_value, expected_value) in zip(actual_rows, expected_rows)))
        evidence = f'actual_rows={actual_rows!r}; expected_rows={expected_rows!r}'
    elif kind == 'structured_row_numeric':
        row = _structured_row(mapping, key=str(spec['key']), row_key=str(spec['row_key']), row_value=str(spec['row_value']))
        actual = _structured_field(row, spec)
        expected = float(spec['expected_value'])
        met = row is not None and _numeric_matches_spec(actual, expected, spec)
        evidence = f"row={spec['row_value']!r}; field={spec['field']!r}; actual={actual!r}; expected={expected!r}"
    elif kind == 'structured_row_boolean':
        row = _structured_row(mapping, key=str(spec['key']), row_key=str(spec['row_key']), row_value=str(spec['row_value']))
        actual = _structured_field(row, spec)
        expected = bool(spec['expected_value'])
        if isinstance(actual, bool):
            normalized = actual
        elif isinstance(actual, str):
            normalized = {'true': True, 'false': False}.get(actual.strip().casefold())
        else:
            normalized = None
        met = row is not None and normalized is expected
        evidence = f"row={spec['row_value']!r}; field={spec['field']!r}; actual={actual!r}; expected={expected!r}"
    elif kind == 'structured_row_list':
        row = _structured_row(mapping, key=str(spec['key']), row_key=str(spec['row_key']), row_value=str(spec['row_value']))
        actual = _structured_field(row, spec)
        expected = list(spec['expected_value'])
        met = row is not None and unordered_semantic_list_matches(actual, expected)
        evidence = f"row={spec['row_value']!r}; field={spec['field']!r}; actual={actual!r}; expected unordered values={expected!r}"
    elif kind == 'structured_row_string':
        row = _structured_row(mapping, key=str(spec['key']), row_key=str(spec['row_key']), row_value=str(spec['row_value']))
        actual = _structured_field(row, spec)
        expected = str(spec['expected_value'])
        met = row is not None and semantic_value_matches(actual, expected)
        evidence = f"row={spec['row_value']!r}; field={spec['field']!r}; actual={actual!r}; expected={expected!r}"
    elif kind == 'structured_row_nullable_numeric':
        row = _structured_row(mapping, key=str(spec['key']), row_key=str(spec['row_key']), row_value=str(spec['row_value']))
        actual = None if row is None else _semantic_get(row, str(spec['field']))
        met = row is not None and actual is None
        evidence = f"row={spec['row_value']!r}; field={spec['field']!r}; actual={actual!r}; expected=None"
    elif kind == 'structured_row_order_composite':
        actual = _semantic_get(mapping, spec['key'])
        row_keys = [str(value) for value in spec['row_keys']]
        actual_rows = []
        if isinstance(actual, list):
            actual_rows = [[_semantic_get(item, row_key) for row_key in row_keys] for item in actual if isinstance(item, dict)]
        expected_rows = [list(row) for row in spec['expected_rows']]
        met = len(actual_rows) == len(expected_rows) and all((len(actual_row) == len(expected_row) and all((semantic_value_matches(actual_value, expected_value) for (actual_value, expected_value) in zip(actual_row, expected_row))) for (actual_row, expected_row) in zip(actual_rows, expected_rows)))
        evidence = f'actual_rows={actual_rows!r}; expected_rows={expected_rows!r}'
    elif kind == 'structured_row_numeric_composite':
        actual_rows = _semantic_get(mapping, spec['key'])
        row = None
        if isinstance(actual_rows, list):
            for candidate in actual_rows:
                if not isinstance(candidate, dict):
                    continue
                if all((semantic_value_matches(_semantic_get(candidate, str(row_key)), row_value) for (row_key, row_value) in zip(spec['row_keys'], spec['row_values']))):
                    row = candidate
                    break
        actual = None if row is None else _semantic_get(row, str(spec['field']))
        expected = float(spec['expected_value'])
        met = row is not None and _numeric_matches_spec(actual, expected, spec)
        evidence = f"row={spec['row_values']!r}; field={spec['field']!r}; actual={actual!r}; expected={expected!r}"
    else:
        raise ValueError(f'Unsupported console criterion: {kind}')
    return Criterion(spec['id'], spec['description'], met, evidence)

def _normalize_console_mapping(task_id: str, mapping: dict[str, Any]) -> dict[str, Any]:
    """Accept financially equivalent field names without weakening the math.

    Console tasks ask for professional schedules, not a private serialization
    dialect.  Normalize only exact, source-equivalent aliases or identities;
    every value still has to match the deterministic gold calculation.
    """
    normalized = dict(mapping)

    def copied_rows(key: str) -> list[dict[str, Any]]:
        raw = mapping.get(key)
        if not isinstance(raw, list):
            return []
        rows = [dict(row) for row in raw if isinstance(row, dict)]
        normalized[key] = rows
        return rows

    def canonical_label(value: Any, rules: dict[str, tuple[str, ...]]) -> str | None:
        text = _normalize(value)
        candidates = [(target, alias) for (target, aliases) in rules.items() for alias in (target, *aliases)]
        candidates.sort(key=lambda pair: len(_normalize(pair[1])), reverse=True)
        for (target, alias) in candidates:
            if semantic_value_matches(value, alias) or _normalize(alias) in text:
                return target
        return None
    return normalized

def _grade_console(task_id: str, answer: Any) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    mapping = _normalize_console_mapping(task_id, _answer_mapping(answer))
    criteria = [_console_criterion(mapping, spec) for spec in gold['criteria']]
    result = _result(criteria)
    reviews: list[dict[str, Any]] = []
    by_id = {criterion.id: criterion for criterion in criteria}
    for spec in gold['criteria']:
        if not spec.get('semantic'):
            continue
        if spec['kind'] in {'string', 'boolean'}:
            expected_facts = dict(spec['expected'])
        elif spec['kind'] == 'list_item':
            if spec.get('exclusive_with_keys'):
                expected_facts = {spec['key']: {'must_include': spec['expected_item']}}
                expected_facts.update({conflicting_key: {'must_exclude': spec['expected_item']} for conflicting_key in spec['exclusive_with_keys']})
            else:
                expected_facts = {spec['key']: spec['expected_item']}
        else:
            expected_facts = {spec['key']: list(spec['expected'])}
        reviews.append({'criterion_id': spec['id'], 'requirement': semantic_requirement(criterion_id=spec['id'], description=spec['description'], expected_facts=expected_facts, artifact_type='submitted answer') + (' For a list criterion, require the complete requested set: reject missing, extra, duplicated, or wrongly classified items.' if spec['kind'] in {'list', 'list_exact'} else '') + (' For this membership criterion, require the item in the named bucket and absent from every conflicting bucket.' if spec.get('exclusive_with_keys') else ''), 'hard_gate_met': bool(by_id[spec['id']].met) if spec['kind'] == 'boolean' or (spec['kind'] == 'list_item' and spec.get('exclusive_with_keys')) else bool(mapping), 'hard_gate_evidence': by_id[spec['id']].evidence if spec['kind'] == 'boolean' or (spec['kind'] == 'list_item' and spec.get('exclusive_with_keys')) else 'answer parsed into a non-empty field mapping' if mapping else 'answer did not parse into a non-empty field mapping', 'legacy_lexical_match': bool(by_id[spec['id']].met)})
    if reviews:
        result['semantic_review'] = {'version': 2, 'mode': 'deterministic_hard_gates_plus_bounded_semantic_judge', 'task_id': task_id, 'artifact': None, 'evidence': str(answer or '')[:60000], 'criteria': reviews, 'policy': 'A criterion passes only when its parsing gate passes and the semantic judge finds the exact professional conclusion or complete list MET.'}
    by_spec = {str(spec['id']): spec for spec in gold['criteria']}
    for row in result['criteria']:
        spec = by_spec[str(row['id'])]
        for key in ('category', 'weight', 'failure_cap', 'semantic', 'kind'):
            if key in spec:
                row[key] = spec[key]
    return result

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

def _xlsx_label_value(workbook, values, label: str, expected: Any) -> bool:
    wanted = _normalize(label)
    (indexed_cells, exact_labels) = _xlsx_label_cell_index(workbook)

    def record_matches(record: tuple[str, int, int, Any]) -> bool:
        (sheet_name, row_number, column_number, _value) = record
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else workbook[sheet_name]
        candidates = [value_sheet.cell(row_number, column_number + offset).value for offset in (1, 2, 3)]
        candidates.append(value_sheet.cell(row_number + 1, column_number).value)
        return _value_candidates_match(candidates, expected)
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

def _xlsx_exact_label_value(workbook, values, label: str, expected: Any) -> bool:
    """Require the value beside the exact normalized workbook label.

    Task 027 publishes a complete controller output registry, so an incorrect
    value on that named row must not be rescued by a similar label elsewhere
    in the model.  Other tasks retain the broader professional-alias fallback.
    """
    (_indexed_cells, exact_labels) = _xlsx_label_cell_index(workbook)
    for (sheet_name, row_number, column_number, _value) in exact_labels.get(_normalize(label), []):
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else workbook[sheet_name]
        candidates = [value_sheet.cell(row_number, column_number + offset).value for offset in (1, 2, 3)]
        candidates.append(value_sheet.cell(row_number + 1, column_number).value)
        if _value_candidates_match(candidates, expected):
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

def _xlsx_row_column_formula(workbook, row_label: str, column_label: str) -> tuple[bool, str]:
    """Find a source-linked formula at a semantic row/column intersection."""
    label_locations: list[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            row_labels = [cell for cell in row if semantic_equal(cell.value, row_label) or contains_concept(cell.value, row_label)]
            if not row_labels:
                continue
            label_locations.extend((f'{sheet.title}!{cell.coordinate}' for cell in row_labels))
            for candidate in row:
                if not _is_source_linked_formula(candidate.value):
                    continue
                headers = [sheet.cell(header_row, candidate.column).value for header_row in range(max(1, candidate.row - 12), candidate.row)]
                if any((semantic_equal(header, column_label) or contains_concept(header, column_label) for header in headers)):
                    return (True, f'{sheet.title}!{candidate.coordinate}={candidate.value}; row={row_label!r}; column={column_label!r}')
    if label_locations:
        return (False, f'row label found at {label_locations!r}, but no source-linked formula was under column {column_label!r}')
    return (False, f'row label {row_label!r} not found')

def _task_037_selection_tie_control(workbook, values) -> tuple[bool, str]:
    """Require a real independent selected-set calculation that ties to zero."""
    if 'Checks' not in workbook.sheetnames or 'Checks' not in values.sheetnames:
        return (False, 'Checks sheet is missing')
    sheet = workbook['Checks']
    value_sheet = values['Checks']
    reference_pattern = re.compile("(?:'Portfolio Selection'|Portfolio Selection)!\\$?([A-Z]{1,3})\\$?(\\d+)(?::\\$?([A-Z]{1,3})\\$?(\\d+))?", flags=re.I)
    for row in sheet.iter_rows():
        for cell in row:
            if not (contains_concept(cell.value, 'selection tie') or contains_concept(cell.value, 'selected portfolio tie')):
                continue
            candidates = [sheet.cell(cell.row, cell.column + offset) for offset in (1, 2, 3) if cell.column + offset <= sheet.max_column]
            if cell.row + 1 <= sheet.max_row:
                candidates.append(sheet.cell(cell.row + 1, cell.column))
            for candidate in candidates:
                formula = candidate.value
                if not isinstance(formula, str) or not formula.startswith('='):
                    continue
                references = reference_pattern.findall(formula)
                normalized_references = {(start_column.casefold(), int(start_row), end_column.casefold() if end_column else '', int(end_row) if end_row else 0) for (start_column, start_row, end_column, end_row) in references}
                single_references = {reference for reference in normalized_references if not reference[2]}
                range_references = {reference for reference in normalized_references if reference[2]}
                recalculated = value_sheet[candidate.coordinate].value
                if 'sumproduct(' in formula.casefold().replace(' ', '') and len(single_references) >= 1 and (len(range_references) >= 2) and (len(normalized_references) >= 3) and isinstance(recalculated, (int, float)) and (not isinstance(recalculated, bool)) and _close(float(recalculated), 0.0, abs_tol=0.01, rel_tol=0.0):
                    return (True, f'Checks!{candidate.coordinate}={formula}; value={recalculated}')
            return (False, f'selection-tie label found at Checks!{cell.coordinate}, but no independent selected-set zero formula')
    return (False, 'no selection-tie control on Checks')

def _display_tolerance(expected: float) -> float:
    if abs(expected) <= 10:
        return 0.0005
    return max(0.02, abs(expected) * 0.0005)

def _value_candidates_match(candidates: Iterable[Any], expected: Any) -> bool:
    if isinstance(expected, list):
        return ordered_semantic_list_matches(list(candidates), expected)
    if isinstance(expected, str):
        return any((date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates))
    if isinstance(expected, bool):
        return any((_boolean_matches(candidate, expected) for candidate in candidates))
    return any((_close(candidate, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0) for candidate in candidates))

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
    return any((_numeric_literal_is_compatible_with_label(label, literal) and _shared_helper_dc68be01_numeric_literal_matches(literal, float(expected)) for literal in _SHARED_HELPER_DC68BE01_NUMERIC_LITERAL_PATTERN.findall(text)))

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
            missing = [label for (label, expected) in spec['label_values'].items() if not _workbook_value_present(values, expected)]
            if missing:
                return (False, f'exact cached facts missing anywhere in workbook={missing!r}')
            shared_helper_089e2eb4_equivalent_labels = set()
            extreme_associations = [outcome for (label, expected) in spec['label_values'].items() if label not in shared_helper_089e2eb4_equivalent_labels if (outcome := _extreme_headline_association_present(workbook, values, label, expected)) is not None]
            failed_extremes = [evidence for (met, evidence) in extreme_associations if not met]
            if failed_extremes:
                return (False, '; '.join(failed_extremes))
            conflicts = [conflict for (label, expected) in spec['label_values'].items() if (conflict := _explicit_label_value_conflict(workbook, values, label, expected)) is not None]
            return (not conflicts, 'exact cached facts are present and no exact-label contradiction exists' if not conflicts else '; '.join(conflicts))
        if kind == 'artifact_tokens':
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
        (hard_gate_met, hard_gate_evidence) = _hybrid_review_hard_gate(task_id, spec, gold, path, workbook, values)
        expected_facts = spec.get('label_values')
        if expected_facts is None and spec.get('tokens'):
            expected_facts = {'required_concept': spec['tokens'][0]}
        if expected_facts is None and spec.get('heading'):
            expected_facts = {'required_section': spec['heading']}
        if expected_facts is None and spec.get('title_token'):
            expected_facts = {'required_title': spec['title_token']}
        requirement = semantic_requirement(criterion_id=spec['id'], description=spec['description'], expected_facts=expected_facts, artifact_type={'.xlsx': 'workbook', '.docx': 'memorandum', '.pptx': 'board presentation'}.get(path.suffix.lower(), 'artifact'))
        reviews.append({'criterion_id': spec['id'], 'requirement': requirement, 'hard_gate_met': hard_gate_met, 'hard_gate_evidence': hard_gate_evidence, 'legacy_lexical_match': bool(by_id[spec['id']].met)})
    if not reviews:
        return None
    return {'version': 2, 'mode': 'deterministic_hard_gates_plus_bounded_semantic_judge', 'task_id': task_id, 'artifact': gold['artifact']['path'], 'evidence': _semantic_evidence_pack(path, workbook, values), 'criteria': reviews, 'policy': 'A criterion passes only when its deterministic hard gate passes and the semantic judge finds the professional-language association or narrative substance MET.'}

def _artifact_label_value(entries: list[str], label: str, expected: Any) -> bool:
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
            if any((date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates)):
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
_SHARED_HELPER_DC68BE01_NUMERIC_LITERAL_PATTERN = re.compile('\n    (?<![A-Za-z0-9])\n    (?:-?\\$?[ \\t]*\\(?|\\([ \\t]*\\$?-?[ \\t]*)\n    \\d[\\d,]*(?:\\.\\d+)?[ \\t]*\\)?\n    (?:[ \\t]*(?:\n        %\n        |[x×]\n        |k(?![A-Za-z])\n        |thousand(?![A-Za-z])\n        |m(?:m|illion)?(?![A-Za-z])\n        |b(?:n|illion)?(?![A-Za-z])\n    ))?\n    [ \\t]*\\)?\n    ', flags=re.I | re.X)

def _shared_helper_dc68be01_numeric_literal_matches(literal: str, expected: float) -> bool:
    candidates = list(iter_numeric_candidates(literal))
    accounting_negative = '(' in literal and ')' in literal
    explicit_negative = '-' in literal.partition(next((c for c in literal if c.isdigit()), ''))[0]
    percent = literal.rstrip().rstrip(')').rstrip().endswith('%')
    cleaned_for_precision = literal.strip().replace('$', '').replace(',', '').replace('(', '').replace(')', '').strip()
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
    cleaned = literal.strip().replace('$', '').replace(',', '').strip('() ')
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

def _presentation_contexts(path: Path) -> list[str]:
    """Return slide, table-row, and table-intersection contexts."""
    presentation = Presentation(path)
    contexts: list[str] = []
    for slide in presentation.slides:
        slide_parts = [shape.text.strip() for shape in slide.shapes if hasattr(shape, 'text') and shape.text.strip()]
        slide_text = ' | '.join(slide_parts)
        contexts.append(slide_text)
        for shape in slide.shapes:
            if not getattr(shape, 'has_table', False):
                continue
            rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
            if not rows:
                continue
            headers = rows[0]
            for row in rows[1:]:
                contexts.append(' | '.join(row))
                for (column, cell) in enumerate(row):
                    header = headers[column] if column < len(headers) else ''
                    contexts.append(' | '.join((row[0] if row else '', header, cell)))
    return contexts

def _context_has_alias(context: str, alias: str) -> bool:
    if contains_concept(context, alias):
        return True
    context_tokens = set(_normalize(context).split())
    alias_tokens = [token for token in _normalize(alias).split() if len(token) > 1]
    return bool(alias_tokens) and all((token in context_tokens for token in alias_tokens))

def _task_068_artifact_value(path: Path, label: str, expected: Any) -> bool:
    aliases = [label.replace('_', ' '), *_TASK_072_LABEL_ALIASES.get(label, [])]
    aliases.extend({'june_revenue': ['June actual', 'June revenue'], 'june_ebitda': ['June EBITDA'], 'q2_approved_plan_revenue': ['Revenue plan', 'Plan revenue'], 'q2_revenue_variance_to_plan': ['Revenue vs plan', 'Revenue variance'], 'q2_approved_plan_adjusted_ebitda': ['EBITDA plan', 'Plan EBITDA'], 'q2_adjusted_ebitda_variance_to_plan': ['EBITDA vs plan', 'EBITDA variance'], 'q2_organic_growth': ['IR-organic', 'organic'], 'construction_gross_profit_impact': ['Construction GP', 'Construction gross profit'], 'service_labor_productivity_impact': ['Service labor', 'Labor productivity'], 'controls_mix_impact': ['Controls mix'], 'revenue_guidance_low': ['Guidance'], 'revenue_guidance_high': ['Guidance'], 'ebitda_guidance_low': ['Guidance'], 'ebitda_guidance_high': ['Guidance'], 'largest_downside_driver': ['largest downside', 'copper escalation'], 'guidance_update_required': ['guidance']}.get(label, []))
    contexts = _presentation_contexts(path)
    if isinstance(expected, bool):
        if label == 'guidance_update_required' and expected is False:
            text = _normalize('\n'.join(contexts))
            return any((phrase in text for phrase in ('no guidance update', 'maintain guidance', 'hold guidance', 'guidance unchanged', 'do not revise guidance')))
        return any((any((_context_has_alias(context, alias) for alias in aliases)) and _boolean_matches(context, expected) for context in contexts))
    if isinstance(expected, str):
        return any((any((_context_has_alias(context, alias) for alias in aliases)) and semantic_value_matches(context, expected) for context in contexts))
    for context in contexts:
        if not any((_context_has_alias(context, alias) for alias in aliases)):
            continue
        if any((_shared_helper_dc68be01_numeric_literal_matches(literal, float(expected)) for literal in _SHARED_HELPER_DC68BE01_NUMERIC_LITERAL_PATTERN.findall(context))):
            return True
    return False

def _task_100_artifact_value(path: Path, label: str, expected: Any) -> bool:
    """Grade a board deck by meaning, including editable chart data and $m units.

    Board materials normally state units once at the slide or deck level and
    then show compact values such as ``10.77`` or ``3.17x``.  The generic
    artifact reader cannot safely infer those units across every task, so this
    reader is intentionally scoped to task 100 and its controlled labels.
    """
    aliases = [label.replace('_', ' '), *_TASK_100_LABEL_ALIASES.get(label, [])]
    entries = _artifact_text_entries(path)
    entries.extend(_presentation_contexts(path))
    text = _normalize('\n'.join(entries))
    if isinstance(expected, list):
        if label == 'optimized_board_priority_portfolio':
            start = next((i for (i, entry) in enumerate(entries) if contains_concept(entry, 'optimized decision portfolio')), None)
            if start is None:
                return False
            stop = next((i for i in range(start + 1, len(entries)) if contains_concept(entries[i], 'constraints')), min(len(entries), start + 24))
            selected_entries = entries[start:stop]
            return all((any((semantic_value_matches(entry, item) for entry in selected_entries)) for item in expected))
        for (index, entry) in enumerate(entries):
            if not any((contains_concept(entry, alias) for alias in aliases)):
                continue
            candidates = entries[index:index + max(4, len(expected) + 2)]
            if unordered_semantic_list_matches(candidates[1:1 + len(expected)], expected):
                return True
            if ordered_semantic_list_matches(candidates, expected):
                return True
        return False
    if isinstance(expected, bool):
        if label == 'severe_covenant_breach':
            severe_is_flagged = bool(re.search('severe.{0,240}breach', text) or re.search('breach.{0,240}severe', text) or ('breach flags' in text and 'severe' in text and ('executed leverage' in text)))
            return severe_is_flagged if expected else not severe_is_flagged
        if label == 'post_action_severe_cash_breach':
            flagged = bool(re.search('breach.{0,120}cash\\s*<', text) or re.search('cash.{0,120}breach', text))
            return flagged if expected else not flagged
        if label == 'post_action_severe_leverage_breach':
            flagged = bool(re.search('breach.{0,180}leverage\\s*>', text) or re.search('leverage.{0,120}breach', text))
            return flagged if expected else not flagged
        return any((any((contains_concept(segment, alias) for alias in aliases)) and _boolean_matches(segment, expected) for segment in entries))
    if isinstance(expected, str):
        if label == 'largest_branch_ebitda_miss':
            named_callout = bool(re.search('controls.{0,180}(?:miss|negative|below)', text) or re.search('(?:miss|negative|below).{0,180}controls', text))
            negative_row = any((contains_concept(segment, expected) and bool(re.search('(?:\\(\\s*\\$?\\d|[-−]\\s*\\$?\\d)', segment)) for segment in entries))
            return named_callout or negative_row
        return any((any((contains_concept(segment, alias) for alias in aliases)) and semantic_value_matches(segment, expected) for segment in entries))
    pattern = re.compile('\\(?[-−]?\\$?\\d[\\d,]*(?:\\.\\d+)?[ \\t]*(?:%|[kmbx×])?\\)?', flags=re.I)
    for (index, entry) in enumerate(entries):
        if not any((_context_has_alias(entry, alias) for alias in aliases)):
            continue
        segment = ' | '.join(entries[index:index + 4])
        for literal in pattern.findall(segment):
            normalized_literal = literal.replace('−', '-')
            candidates = list(iter_numeric_candidates(normalized_literal))
            precision_text = normalized_literal.strip().replace('$', '').replace(',', '').strip('() ')
            precision_suffix = precision_text[-1:].casefold()
            precision_multiplier = {'k': 1000.0, 'm': 1000000.0, 'b': 1000000000.0}.get(precision_suffix, 1.0)
            if precision_suffix in {'k', 'm', 'b', 'x', '×'}:
                precision_text = precision_text[:-1].strip()
            precision_text = precision_text.rstrip('%').strip()
            precision_decimals = len(precision_text.rsplit('.', 1)[1]) if '.' in precision_text else 0
            precision_tolerance = 0.5 * 10 ** (-precision_decimals) * precision_multiplier
            if normalized_literal.strip().endswith('%'):
                precision_tolerance /= 100.0
            if any((_close(candidate, float(expected), abs_tol=max(_display_tolerance(float(expected)), precision_tolerance + 1e-12), rel_tol=0.0) for candidate in candidates)):
                return True
            cleaned = normalized_literal.strip().replace('$', '').replace(',', '').strip('() ')
            suffix = cleaned[-1:].casefold()
            if abs(float(expected)) <= 10000 or suffix in {'k', 'm', 'b', '%', 'x', '×'}:
                continue
            try:
                displayed = float(cleaned)
            except ValueError:
                continue
            if normalized_literal.strip().startswith('('):
                displayed = -abs(displayed)
            scaled = displayed * 1000000.0
            decimals = len(cleaned.rsplit('.', 1)[1]) if '.' in cleaned else 0
            display_tolerance = 0.5 * 10 ** (-decimals) * 1000000.0 + 1e-09
            if _close(scaled, float(expected), abs_tol=max(display_tolerance, _display_tolerance(float(expected))), rel_tol=0.0):
                return True
    return False

def _task_037_selected_projects(workbook, values, expected: list[str]) -> bool:
    selected: list[str] = []
    index = _task_037_workbook_index(workbook, values)
    for (project, records) in index['cached_projects'].items():
        for record in records:
            row_values = record['cached'][:12]
            chosen = any((semantic_equal(value, 'select') or semantic_equal(value, 'selected') or value is True or (isinstance(value, (int, float)) and (not isinstance(value, bool)) and (float(value) == 1.0)) for value in row_values))
            if chosen and project not in selected:
                selected.append(project)
    return unordered_semantic_list_matches(selected, expected)

def _task_037_workbook_index(workbook, values) -> dict[str, Any]:
    """Index Task037 rows once instead of rescanning 8,000+ cells per fact.

    Task037 deliberately has hundreds of atomic finance checks.  The prior
    implementation performed a full workbook scan for nearly every one, which
    was exact but quadratic on a realistic portfolio model and could exhaust
    HUD's rollout clock after the agent had already finished.  This index holds
    the same formula/cached values and does not change any matching rule.
    """
    cached = getattr(workbook, '_alder_task_037_index', None)
    if cached is not None:
        return cached
    rows: list[dict[str, Any]] = []
    by_sheet_row: dict[tuple[str, int], dict[str, Any]] = {}
    formula_projects: dict[str, list[dict[str, Any]]] = {}
    cached_projects: dict[str, list[dict[str, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        (max_row, max_column) = (sheet.max_row, sheet.max_column)
        for row_index in range(1, max_row + 1):
            formula_values = tuple((sheet.cell(row_index, column).value for column in range(1, max_column + 1)))
            cached_values = tuple((value_sheet.cell(row_index, column).value for column in range(1, max_column + 1)))
            record = {'sheet': sheet_name, 'row': row_index, 'max_column': max_column, 'formula': formula_values, 'cached': cached_values}
            rows.append(record)
            by_sheet_row[sheet_name, row_index] = record
            for value in formula_values:
                for project in re.findall('\\bCP-\\d{2}\\b', str(value or ''), flags=re.I):
                    bucket = formula_projects.setdefault(project.upper(), [])
                    if not bucket or bucket[-1] is not record:
                        bucket.append(record)
            for value in cached_values:
                if re.fullmatch('CP-\\d{2}', str(value or ''), flags=re.I):
                    cached_projects.setdefault(str(value).upper(), []).append(record)
                    break
    cached = {'rows': rows, 'by_sheet_row': by_sheet_row, 'formula_projects': formula_projects, 'cached_projects': cached_projects}
    setattr(workbook, '_alder_task_037_index', cached)
    return cached

def _professional_row_value_match(workbook, values, aliases: list[str], expected: Any, require_formula: bool=False, year: int | None=None) -> tuple[bool, str]:
    """Match a normal finance schedule by row label and, when needed, year column."""
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        year_columns: set[int] = set()
        if year is not None:
            for row in sheet.iter_rows():
                for cell in row:
                    if semantic_equal(cell.value, f'Year {year}'):
                        year_columns.add(cell.column)
        for row_index in range(1, sheet.max_row + 1):
            row_labels = [sheet.cell(row_index, column).value for column in range(1, sheet.max_column + 1)]
            if not any((any((contains_concept(value, alias) for alias in aliases)) for value in row_labels)):
                continue
            columns = sorted(year_columns) if year_columns else list(range(1, sheet.max_column + 1))
            for column in columns:
                formula_value = sheet.cell(row_index, column).value
                cached_value = value_sheet.cell(row_index, column).value
                if require_formula and (not (isinstance(formula_value, str) and formula_value.startswith('='))):
                    continue
                if _value_candidates_match([cached_value], expected):
                    return (True, f'professional schedule match at {sheet_name}!{sheet.cell(row_index, column).coordinate}')
    return (False, 'no professional schedule row match')

def _task_055_row_match(workbook, values, label: str, expected: Any, require_formula: bool=False) -> tuple[bool, str]:
    match = re.fullmatch('year_(one|two)_(.+)', label)
    if match:
        year = 1 if match.group(1) == 'one' else 2
        suffix = match.group(2)
        aliases = _TASK_055_YEAR_ROW_ALIASES.get(suffix)
        if aliases:
            return _professional_row_value_match(workbook, values, aliases, expected, require_formula, year)
    return _professional_row_value_match(workbook, values, [label.replace('_', ' '), *_TASK_055_LABEL_ALIASES.get(label, [])], expected, require_formula)

def _task_037_row_match(workbook, values, label: str, expected: Any, require_formula: bool=False) -> tuple[bool, str]:
    index = _task_037_workbook_index(workbook, values)
    project_match = re.fullmatch('cp_(\\d{2})_(.+)', label)
    if project_match:
        project = f'CP-{project_match.group(1)}'
        suffix = project_match.group(2)
        aliases = [suffix.replace('_', ' '), *_TASK_037_LABEL_ALIASES.get(label, [])]
        year_flow = re.fullmatch('year_(\\d)_(base|downside)_after_tax_cash_flow', suffix)
        if year_flow:
            (year, case) = year_flow.groups()
            aliases.extend([f"Yr{year} {('down' if case == 'downside' else 'CF')}", f'Year {year} {case} cash flow'])
        aliases.extend({'technician_capacity': ['Technician req.', 'Technician capacity required'], 'mandatory_safety_flag': ['Mandatory safety'], 'npv': ['NPV base', 'NPV (base)'], 'irr': ['IRR base', 'IRR (base)'], 'payback_years': ['Disc payback', 'Payback'], 'selected': ['Selected', 'Select', 'Portfolio decision']}.get(suffix, []))
        for record in index['formula_projects'].get(project, []):
            sheet_name = record['sheet']
            row_index = record['row']
            for column in range(1, record['max_column'] + 1):
                headers = [index['by_sheet_row'][sheet_name, header_row]['formula'][column - 1] for header_row in range(max(1, row_index - 4), row_index)]
                if not any((any((contains_concept(header, alias) for alias in aliases)) for header in headers)):
                    continue
                formula_value = record['formula'][column - 1]
                cached_value = record['cached'][column - 1]
                if require_formula and (not (isinstance(formula_value, str) and formula_value.startswith('='))):
                    continue
                if _value_candidates_match([cached_value], expected):
                    coordinate = f'{get_column_letter(column)}{row_index}'
                    return (True, f'project schedule match at {sheet_name}!{coordinate}')
    aliases = [label.replace('_', ' '), *_TASK_037_LABEL_ALIASES.get(label, [])]
    for record in index['rows']:
        if not any((contains_concept(value, alias) for value in record['formula'] for alias in aliases)):
            continue
        for (column, (formula_value, cached_value)) in enumerate(zip(record['formula'], record['cached']), start=1):
            if require_formula and (not (isinstance(formula_value, str) and formula_value.startswith('='))):
                continue
            if _value_candidates_match([cached_value], expected):
                coordinate = f"{get_column_letter(column)}{record['row']}"
                return (True, f"professional schedule match at {record['sheet']}!{coordinate}")
    return (False, 'no professional schedule row match')

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
    if task_id == 'task_037':
        candidates.extend(_TASK_037_LABEL_ALIASES.get(label, []))
    if task_id == 'task_061':
        candidates.extend(_TASK_061_LABEL_ALIASES.get(label, []))
    if task_id == 'task_037':
        (matched, detail) = _task_037_row_match(workbook, values, label, gold['answer'][label], require_formula=True)
        if matched:
            return (matched, detail)
    attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
    (matched, detail) = next((attempt for attempt in attempts if attempt[0]), attempts[0])
    if task_id == 'task_055' and (not matched):
        row_column = {'year_one_gaap_eps_accretion': ('GAAP EPS accretion / (dilution) %', 'Year 1'), 'year_one_adjusted_eps_accretion': ('Adjusted EPS accretion / (dilution) %', 'Year 1'), 'year_two_gaap_eps_accretion': ('GAAP EPS accretion / (dilution) %', 'Year 2'), 'year_two_adjusted_eps_accretion': ('Adjusted EPS accretion / (dilution) %', 'Year 2')}
        if label in row_column:
            (matched, detail) = _xlsx_row_column_formula(workbook, *row_column[label])
        if not matched:
            (matched, detail) = _task_055_row_match(workbook, values, label, gold['answer'][label], require_formula=True)
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
            target_sheet = workbook[target_name] if target_name in workbook.sheetnames else None
            quoted_reference = f"'{source_name.replace(chr(39), chr(39) * 2)}'!".casefold()
            unquoted_reference = f'{source_name}!'.casefold()
            matches = [] if target_sheet is None else [f'{target_name}!{cell.coordinate}={cell.value}' for row in target_sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith('=') and (quoted_reference in cell.value.casefold() or unquoted_reference in cell.value.casefold())]
            met = bool(matches)
            evidence = f'target_sheet={target_name!r}; source_sheet={source_name!r}; linked_formula_examples={matches[:3]!r}'
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
                if task_id == 'task_037':
                    candidates.extend(_TASK_037_LABEL_ALIASES.get(label, []))
                attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
                (matched, detail) = next((attempt for attempt in attempts if attempt[0]), attempts[0])
                if task_id == 'task_055' and (not matched):
                    expected = gold['answer'][label]
                    (matched, detail) = _task_055_row_match(workbook, values, label, expected, require_formula=True)
                if task_id == 'task_037' and (not matched):
                    expected = gold['answer'][label]
                    (matched, detail) = _task_037_row_match(workbook, values, label, expected, require_formula=True)
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
                if task_id == 'task_037':
                    if label == 'selected_portfolio' and isinstance(expected, list):
                        matched = _task_037_selected_projects(workbook, values, expected)
                    else:
                        (matched, _) = _task_037_row_match(workbook, values, label, expected, require_formula=False)
                    if not matched:
                        matched = any((_xlsx_label_value(workbook, values, candidate, expected) for candidate in [*labels, *_TASK_037_LABEL_ALIASES.get(label, [])]))
                else:
                    matched = any((_xlsx_exact_label_value(workbook, values, candidate, expected) if task_id == 'task_027' else _xlsx_label_value(workbook, values, candidate, expected) for candidate in labels))
                if task_id == 'task_055' and (not matched):
                    (matched, _) = _task_055_row_match(workbook, values, label, expected, require_formula=False)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = 'all labeled values matched' if met else f'missing or incorrect labels={failures!r}'
        elif kind == 'task_037_selection_tie':
            (met, evidence) = _task_037_selection_tie_control(workbook, values)
        elif kind == 'artifact_label_values':
            failures = []
            for (label, expected) in spec['label_values'].items():
                labels = [label.replace('_', ' ')]
                if task_id in {'task_068', 'task_072'}:
                    labels.extend(_TASK_072_LABEL_ALIASES.get(label, []))
                if task_id == 'task_100':
                    labels.extend(_TASK_100_LABEL_ALIASES.get(label, []))
                matched = any((_artifact_label_value(entries, candidate, expected) for candidate in labels))
                if task_id == 'task_068' and (not matched):
                    matched = _task_068_artifact_value(path, label, expected)
                if task_id == 'task_100' and (not matched):
                    matched = _task_100_artifact_value(path, label, expected)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = 'all labeled values matched' if met else f'missing or incorrect labels={failures!r}'
        elif kind == 'artifact_tokens':
            missing = [token for token in spec['tokens'] if not any((contains_concept(normalized_text, candidate) for candidate in _ARTIFACT_TOKEN_ALIASES.get(token, [token])))]
            met = not missing
            evidence = 'all required content present' if met else f'missing={missing!r}'
        elif kind == 'pptx_structure':
            presentation = Presentation(path)
            expected_slides = spec.get('exact_slides')
            count_ok = len(presentation.slides) == expected_slides if expected_slides is not None else len(presentation.slides) >= spec['min_slides']
            joined_entries = '\n'.join(entries)
            title_present = contains_concept(joined_entries, spec['title_token'])
            if task_id == 'task_100' and (not title_present):
                normalized_title = _normalize(joined_entries)
                title_present = all((token in normalized_title for token in ('fy26', 'outlook', 'fy27', 'priorities')))
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
    gold = load_corporate_finance_gold(task_id)
    if 'artifact' in gold:
        return _grade_artifact(task_id, Path(workspace_root))
    return _grade_console(task_id, answer)
