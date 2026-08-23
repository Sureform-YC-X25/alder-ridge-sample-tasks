from __future__ import annotations
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader
from pptx import Presentation
from runtime.grading.semantic import contains_concept, ordered_semantic_list_matches, semantic_equal, semantic_value_matches
from runtime.grading.hybrid_semantic import semantic_requirement
from runtime.grading.rubric import apply_reward_policy, attach_default_policy
from runtime.mcp.paths import resolve_project_root, resolve_seed_root
CORPORATE_TASK_IDS = frozenset(['task_027', 'task_035', 'task_037', 'task_055', 'task_061', 'task_068', 'task_072', 'task_073', 'task_100'])

GOLD_PATH = Path(__file__).resolve().parent / 'gold' / 'tasks_001_025.json'
SEED_WORKSPACE = resolve_seed_root(__file__) / 'sources'
TASK_GRADING_REVISIONS = {'task_015': {'id': 'task-015-covenant-slide-presentation-quality-v2', 'effective_date': '2026-08-09', 'basis': 'source-relative presentation quality, readable table structure, and round-trip-tolerant preservation of the existing lender deck'}, 'task_027': {'id': 'task-027-covenant-release-decision-v3', 'effective_date': '2026-08-13', 'basis': 'controller-visible FY27 source conventions plus formula-driven branch, EBITDA, free-cash-flow, weekly liquidity, lender leverage and fixed-charge coverage release decisions with proportional section-normalized finance and formula-lineage scoring'}, 'task_035': {'id': 'task-035-executive-recovery-decision-v1', 'effective_date': '2026-08-13', 'basis': 'probability-plan, gross-commitment, executable-priority, contractual-damages, deferred-gross-profit, required-recovery, and release-condition decisions with proportional section-normalized spreadsheet scoring'}}

@dataclass(frozen=True)
class Criterion:
    id: str
    description: str
    met: bool
    evidence: str
    category: str | None = None
    weight: int | None = None
    semantic: bool | None = None
    failure_cap: float | None = None

def load_apex_gold(task_id: str | None=None) -> dict[str, Any]:
    payload = json.loads(GOLD_PATH.read_text(encoding='utf-8'))
    return payload[task_id] if task_id else payload

def _normalize(value: Any) -> str:
    return re.sub('[^a-z0-9]+', ' ', str(value or '').lower()).strip()

def _number(value: Any) -> float | None:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)) and math.isfinite(float(value)):
        return float(value)
    if not isinstance(value, str):
        return None
    text = value.strip()
    percent = text.endswith('%')
    negative = text.startswith('(') and text.endswith(')')
    text = text.strip('() ')
    suffix_match = re.search('([kmb])\\s*$', text, flags=re.I)
    multiplier = {'k': 1000.0, 'm': 1000000.0, 'b': 1000000000.0}.get(suffix_match.group(1).lower() if suffix_match else '', 1.0)
    if suffix_match:
        text = text[:suffix_match.start()]
    text = text.replace('$', '').replace(',', '').replace('×', '').replace('x', '')
    text = text.strip('% ')
    try:
        result = float(text)
    except ValueError:
        return None
    if negative:
        result = -result
    result *= multiplier
    return result / 100 if percent else result

def _answer_mapping(answer: Any) -> dict[str, Any]:
    if isinstance(answer, dict):
        return answer
    text = str(answer or '').strip()
    candidates = [text]
    candidates.extend(re.findall('```(?:json)?\\s*(\\{.*?\\})\\s*```', text, flags=re.I | re.S))
    brace = re.search('\\{.*\\}', text, flags=re.S)
    if brace:
        candidates.append(brace.group(0))
    for candidate in reversed(candidates):
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except (json.JSONDecodeError, TypeError):
            pass
    mapping: dict[str, Any] = {}
    for line in text.splitlines():
        match = re.match('\\s*[-*]?\\s*([A-Za-z][A-Za-z0-9 _/-]{1,60})\\s*[:=]\\s*(.+?)\\s*$', line)
        if match:
            mapping[re.sub('\\s+', '_', match.group(1).strip().lower())] = match.group(2).strip()
    return mapping

def _is_single_json_object(answer: Any) -> bool:
    if isinstance(answer, dict):
        return True
    if not isinstance(answer, str):
        return False
    text = answer.strip()
    if not (text.startswith('{') and text.endswith('}')):
        return False
    try:
        return isinstance(json.loads(text), dict)
    except json.JSONDecodeError:
        return False

def _get(mapping: dict[str, Any], key: str) -> Any:
    if key in mapping:
        return mapping[key]
    wanted = _normalize(key)
    for (candidate, value) in mapping.items():
        if _normalize(candidate) == wanted:
            return value
    return None

def _close(actual: Any, expected: float, *, abs_tol: float=0.02, rel_tol: float=1e-06) -> bool:
    value = _number(actual)
    if value is None:
        return False
    return abs(value - expected) <= max(abs_tol, abs(expected) * rel_tol)

def _numeric_criteria(mapping: dict[str, Any], criterion_id: str, description: str, expected: dict[str, float], *, abs_tol: float=0.02, rel_tol: float=0.0) -> list[Criterion]:
    criteria: list[Criterion] = []
    for (key, target) in expected.items():
        actual = _get(mapping, key)
        met = _close(actual, target, abs_tol=abs_tol, rel_tol=rel_tol)
        criteria.append(Criterion(f"{criterion_id}__{_normalize(key).replace(' ', '_')}", f'{description}: `{key}` is correct', met, 'reported value matched' if met else f'{key}={actual!r}; expected {target}', category='core_finance', weight=10, semantic=False))
    return criteria

def _field_numeric_criteria(mapping: dict[str, Any], criterion_id: str, description: str, expected: dict[str, tuple[float, float]], *, weight: int=10) -> list[Criterion]:
    criteria: list[Criterion] = []
    for (key, (target, abs_tol)) in expected.items():
        actual = _get(mapping, key)
        met = _close(actual, target, abs_tol=abs_tol + 1e-09, rel_tol=0.0)
        criteria.append(Criterion(f"{criterion_id}__{_normalize(key).replace(' ', '_')}", f'{description}: `{key}` is correct', met, 'reported value matched at required precision' if met else f'{key}={actual!r}; expected {target} ± {abs_tol}', category='core_finance', weight=weight, semantic=False))
    return criteria

def _percentage_criterion(mapping: dict[str, Any], criterion_id: str, description: str, key: str, expected_ratio: float, *, abs_tol: float=5e-05) -> Criterion:
    actual = _get(mapping, key)
    matched = _close(actual, expected_ratio, abs_tol=abs_tol, rel_tol=0.0) or _close(actual, expected_ratio * 100, abs_tol=abs_tol * 100, rel_tol=0.0)
    return Criterion(criterion_id, description, matched, 'percentage matched as a decimal ratio or percentage points' if matched else f'{key}={actual!r}; expected {expected_ratio} or {expected_ratio * 100}', category='core_finance', weight=10, semantic=False)

def _result(criteria: list[Criterion]) -> dict[str, Any]:
    met_count = sum((criterion.met for criterion in criteria))
    reward = met_count / len(criteria) if criteria else 0.0
    return {'reward': round(reward, 6), 'strict_pass': bool(criteria) and met_count == len(criteria), 'criteria_met': met_count, 'criteria_total': len(criteria), 'criteria': [{'id': criterion.id, 'description': criterion.description, 'value': int(criterion.met), 'evidence': criterion.evidence, **({'category': criterion.category, 'weight': criterion.weight, 'semantic': criterion.semantic, **({'failure_cap': criterion.failure_cap} if criterion.failure_cap is not None else {})} if criterion.category is not None else {})} for criterion in criteria]}

def _legacy_file_atomic_specs(task_id: str) -> list[tuple[str, str]]:
    """Stable atomic denominator for missing or unreadable legacy artifacts."""
    gold = load_apex_gold(task_id)
    specs: list[tuple[str, str]] = []
    if task_id == 'task_004':
        specs.extend(((f"preservation__{_normalize(sheet).replace(' ', '_')}", f'Preserve {sheet}') for sheet in ('READ ME first', 'Risk Review', 'Evidence Map')))
        for expected in gold['source_inputs']:
            specs.extend(((f"source__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} source {label}") for label in ('current_contract', 'cost_to_date', 'pm_etc', 'documented_etc_overlay', 'billings', 'prior_close_margin_percent')))
        for expected in gold['projects']:
            specs.append((f"formula_lineage__{expected['project_id'].casefold()}", f"{expected['project_id']} formula lineage"))
        for expected in gold['projects']:
            specs.extend(((f"result__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} result {label}") for label in ('close_etc', 'eac', 'earned', 'asset', 'liability', 'margin', 'percent_complete', 'margin_percent', 'prior_margin_percent', 'margin_movement_bps')))
        for expected in gold['review_rows']:
            specs.extend(((f"review__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} review {label}") for label in ('flag', 'commercial_treatment', 'source')))
        specs.append(('total_formula_lineage', 'Total/control row formula lineage'))
        specs.extend(((f'total_result__{coordinate.casefold()}', f'Total result {coordinate}') for coordinate in [f'{column}11' for column in 'CDEFGHIJKLMNOPQ']))
        specs.append(('decision_bridge__structure', 'Controller decision-bridge structure'))
        for expected in gold['decision_rows']:
            project_id = expected['project_id'].casefold()
            specs.extend(((f'decision_bridge__{project_id}__formula_lineage', f"{expected['project_id']} decision-bridge formula lineage"), (f'decision_bridge__{project_id}__pm_case_margin', f"{expected['project_id']} PM-case margin"), (f'decision_bridge__{project_id}__margin_impact', f"{expected['project_id']} overlay margin impact"), (f'decision_bridge__{project_id}__close_disposition', f"{expected['project_id']} close disposition"), (f'decision_bridge__{project_id}__required_follow_up', f"{expected['project_id']} required follow-up")))
        specs.extend((('decision_control__formula_lineage', 'Decision-control formula lineage'), ('decision_control__pm_case_margin', 'Aggregate PM-case margin'), ('decision_control__margin_impact', 'Aggregate overlay margin impact'), ('decision_control__pending_revenue', 'Pending commercial revenue exclusion'), ('decision_control__review_count', 'Review-required project count'), ('decision_control__largest_overlay', 'Largest finance overlay'), ('decision_control__review_queue', 'Controller review queue'), ('decision_control__close_release', 'Aggregate close-release decision'), ('decision_control__policy_basis', 'Decision-control policy basis')))
    else:
        if task_id == 'task_015':
            specs.append(('preservation__slide_count', 'Exactly one slide appended'))
            specs.extend(((f'preservation__slide_{index:02d}', f'Preserve slide {index}') for index in range(1, 6)))
            specs.extend((('structure__title', 'Required slide title'), ('structure__table', 'Required PowerPoint table')))
            specs.extend((('style__visual_system', 'Match the existing deck visual system'), ('style__title_hierarchy', 'Use a clear title hierarchy consistent with the deck'), ('style__table_readability', 'Present the covenant comparison in a reviewable table')))
            for metric in ('leverage', 'fccr', 'tangible_net_worth'):
                specs.extend(((f'metric__{metric}__{field}', f'{metric} {field}') for field in ('posted', 'pro_forma', 'threshold', 'status')))
            specs.extend((('status__proposed', 'Proposed WIP status'), ('status__unposted', 'Unposted WIP status'), ('status__compliant', 'Compliance status')))
        else:
            raise KeyError(task_id)
    ids = [criterion_id for (criterion_id, _) in specs]
    if len(ids) != len(set(ids)):
        raise ValueError(f'duplicate legacy atomic criterion ids for {task_id}')
    return specs

def _file_failure(task_id: str, evidence: str) -> dict[str, Any]:
    return _result([Criterion(criterion_id, description, False, evidence) for (criterion_id, description) in _legacy_file_atomic_specs(task_id)])

def _legacy_file_semantic_ids(task_id: str) -> set[str]:
    ids = {criterion_id for (criterion_id, _) in _legacy_file_atomic_specs(task_id)}
    if task_id == 'task_004':
        return {criterion_id for criterion_id in ids if criterion_id.endswith('__flag') or criterion_id.endswith('__commercial_treatment') or criterion_id.endswith('__close_disposition') or criterion_id.endswith('__required_follow_up') or (criterion_id in {'decision_control__pending_revenue', 'decision_control__review_queue', 'decision_control__close_release'})}
    if task_id == 'task_015':
        return {'structure__title', 'status__compliant', 'status__proposed', 'status__unposted', 'metric__leverage__status', 'metric__fccr__status', 'metric__tangible_net_worth__status'}
    return set()

def _canonicalize_legacy_file_policy(task_id: str, result: dict[str, Any]) -> dict[str, Any]:
    """Make missing, corrupt, and complete artifact rubrics policy-identical.

    Artifact evidence descriptions necessarily differ between a missing file
    and a parsed file, so importance and semantic routing must never be inferred
    from whichever description happens to be present on that execution path.
    """
    canonical = attach_default_policy(_file_failure(task_id, 'canonical policy'))
    semantic_ids = _legacy_file_semantic_ids(task_id)
    policy_by_id: dict[str, dict[str, Any]] = {}
    for row in canonical['criteria']:
        criterion_id = str(row['id'])
        row['semantic'] = criterion_id in semantic_ids
        if criterion_id in semantic_ids and criterion_id.startswith('source__'):
            row.update({'category': 'provenance', 'weight': 3})
            row.pop('failure_cap', None)
        elif task_id == 'task_004':
            if criterion_id.startswith('preservation__'):
                row.update({'category': 'integrity', 'weight': 10, 'failure_cap': 0.0})
            elif criterion_id.startswith('source__') or criterion_id.endswith('__source') or criterion_id == 'decision_control__policy_basis':
                row.update({'category': 'provenance', 'weight': 3})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('formula_lineage__') or criterion_id == 'total_formula_lineage' or criterion_id.endswith('__formula_lineage'):
                row.update({'category': 'auditability', 'weight': 5})
                row.pop('failure_cap', None)
            elif criterion_id == 'decision_bridge__structure':
                row.update({'category': 'structure', 'weight': 1, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.endswith('__flag') or criterion_id.endswith('__commercial_treatment') or criterion_id.endswith('__close_disposition') or criterion_id.endswith('__required_follow_up') or (criterion_id in {'decision_control__pending_revenue', 'decision_control__review_count', 'decision_control__largest_overlay', 'decision_control__review_queue', 'decision_control__close_release'}):
                row.update({'category': 'decision', 'weight': 10, 'semantic': criterion_id in semantic_ids})
                row.pop('failure_cap', None)
            else:
                row.update({'category': 'core_finance', 'weight': 10, 'semantic': False})
                row.pop('failure_cap', None)
        else:
            if task_id == 'task_015' and criterion_id.startswith('style__'):
                row.update({'category': 'presentation_quality', 'weight': 3, 'semantic': False})
                row.pop('failure_cap', None)
            else:
                if row['semantic'] and row.get('category') == 'decision':
                    row.update({'weight': 10, 'failure_cap': 0.49})
        policy_by_id[criterion_id] = {key: row[key] for key in ('category', 'weight', 'semantic', 'failure_cap') if key in row}
    actual_ids = {str(row['id']) for row in result.get('criteria', [])}
    if actual_ids != set(policy_by_id):
        raise ValueError(f'legacy artifact rubric ids drifted for {task_id}: missing={sorted(set(policy_by_id) - actual_ids)}, extra={sorted(actual_ids - set(policy_by_id))}')
    review = result.get('semantic_review')
    if isinstance(review, dict):
        reviewed = {str(row.get('criterion_id')) for row in review.get('criteria', []) if isinstance(row, dict)}
        if reviewed != semantic_ids:
            raise ValueError(f'legacy semantic route drifted for {task_id}: missing={sorted(semantic_ids - reviewed)}, extra={sorted(reviewed - semantic_ids)}')
    for row in result['criteria']:
        row.pop('failure_cap', None)
        row.update(policy_by_id[str(row['id'])])
    return result

def _grade_numeric(task_id: str, answer: Any) -> dict[str, Any]:
    gold = load_apex_gold(task_id)
    mapping = _answer_mapping(answer)
    criteria: list[Criterion]
    if task_id == 'task_001':
        project = gold['project']
        prior = gold['prior_month']
        expected_keys = {'current_contract', 'posted_cost', 'billings', 'pm_etc', 'required_etc_adjustment', 'close_etc', 'eac', 'percent_complete', 'earned_revenue', 'contract_asset', 'contract_liability', 'estimated_margin', 'margin_percent', 'may_earned_revenue', 'may_current_contract', 'may_posted_cost', 'may_close_etc', 'may_eac', 'may_percent_complete', 'may_billings', 'may_contract_asset', 'may_contract_liability', 'may_estimated_margin', 'may_margin_percent', 'current_contract_change', 'posted_cost_change', 'close_etc_change', 'eac_change', 'percent_complete_change_bps', 'earned_revenue_change', 'billings_change', 'contract_asset_change', 'contract_liability_change', 'estimated_margin_change', 'margin_rate_change_bps', 'draft_change_contract_treatment', 'disputed_recovery_etc_action', 'controller_review_required'}
        criteria = [*_numeric_criteria(mapping, 'wip_chain', 'ARM-2409 WIP calculation chain is correct', {'current_contract': project['current_contract'], 'posted_cost': project['cost_to_date'], 'billings': project['billings'], 'pm_etc': gold['pm_etc'], 'required_etc_adjustment': gold['required_etc_adjustment'], 'close_etc': project['estimated_cost_to_complete'], 'eac': project['estimated_cost_at_completion'], 'earned_revenue': project['earned_revenue'], 'contract_asset': project['underbilling'], 'contract_liability': project['overbilling'], 'estimated_margin': project['estimated_total_margin']}, abs_tol=0.02), _percentage_criterion(mapping, 'wip_chain__percent_complete', 'ARM-2409 percent complete is correct', 'percent_complete', project['percent_complete']), _percentage_criterion(mapping, 'wip_chain__margin_percent', 'ARM-2409 total estimated margin rate is correct', 'margin_percent', project['estimated_margin_percent']), *_numeric_criteria(mapping, 'may_position', 'Final May ARM-2409 WIP position is correct', {'may_current_contract': prior['current_contract'], 'may_posted_cost': prior['posted_cost'], 'may_close_etc': prior['close_etc'], 'may_eac': prior['eac'], 'may_earned_revenue': prior['earned_revenue'], 'may_billings': prior['billings'], 'may_contract_asset': prior['contract_asset'], 'may_contract_liability': prior['contract_liability'], 'may_estimated_margin': prior['estimated_margin']}, abs_tol=0.02), _percentage_criterion(mapping, 'may_position__percent_complete', 'Final May ARM-2409 percent complete is correct', 'may_percent_complete', prior['percent_complete']), _percentage_criterion(mapping, 'may_position__margin_percent', 'Final May ARM-2409 total estimated margin rate is correct', 'may_margin_percent', prior['margin_percent']), *_field_numeric_criteria(mapping, 'month_bridge', 'May-to-June ARM-2409 bridge is correct', {'current_contract_change': (gold['bridge']['current_contract_change'], 0.02), 'posted_cost_change': (gold['bridge']['posted_cost_change'], 0.02), 'close_etc_change': (gold['bridge']['close_etc_change'], 0.02), 'eac_change': (gold['bridge']['eac_change'], 0.02), 'percent_complete_change_bps': (gold['bridge']['percent_complete_change_bps'], 0.5), 'earned_revenue_change': (gold['bridge']['earned_revenue_change'], 0.02), 'billings_change': (gold['bridge']['billings_change'], 0.02), 'contract_asset_change': (gold['bridge']['contract_asset_change'], 0.02), 'contract_liability_change': (gold['bridge']['contract_liability_change'], 0.02), 'estimated_margin_change': (gold['bridge']['estimated_margin_change'], 0.02), 'margin_rate_change_bps': (gold['bridge']['margin_rate_change_bps'], 0.5)}), Criterion('decision__draft_change_contract_treatment', 'The disputed draft change is correctly excluded from contract value', _normalize(_get(mapping, 'draft_change_contract_treatment')) == 'exclude', f"actual={_get(mapping, 'draft_change_contract_treatment')!r}; expected='exclude'", category='decision', weight=10, semantic=False), Criterion('decision__disputed_recovery_etc_action', 'The disputed recovery is correctly added back to PM ETC', _normalize(_get(mapping, 'disputed_recovery_etc_action')) == _normalize('add_back_to_pm_etc'), f"actual={_get(mapping, 'disputed_recovery_etc_action')!r}; expected='add_back_to_pm_etc'", category='decision', weight=10, semantic=False), Criterion('decision__controller_review_required', 'The controller-review conclusion is correct', _get(mapping, 'controller_review_required') is True, f"actual={_get(mapping, 'controller_review_required')!r}; expected=True", category='decision', weight=10, semantic=False), Criterion('format__single_json_object', 'The final response contains only the requested JSON object', _is_single_json_object(answer), f'single_json_object={_is_single_json_object(answer)}', category='formatting', weight=1, semantic=False), Criterion('format__exact_keys', 'The JSON object contains exactly the requested keys', set(mapping) == expected_keys, f'actual_keys={sorted(mapping)}; expected_keys={sorted(expected_keys)}', category='formatting', weight=1, semantic=False)]
    else:
        raise KeyError(task_id)
    result = _result(criteria)
    semantic_specs: list[tuple[str, dict[str, Any]]] = []
    if semantic_specs:
        result = _attach_semantic_review(result, task_id=task_id, evidence=str(answer or '')[:60000], artifact_type='submitted answer', specs=[{'criterion_id': criterion_id, 'expected_facts': expected, 'hard_gate_met': bool(mapping), 'hard_gate_evidence': 'answer parsed into a non-empty field mapping' if mapping else 'answer did not parse'} for (criterion_id, expected) in semantic_specs])
    if task_id in TASK_GRADING_REVISIONS:
        result['task_grading_revision'] = dict(TASK_GRADING_REVISIONS[task_id])
    return result

def _attach_semantic_review(result: dict[str, Any], *, task_id: str, evidence: str, artifact_type: str, specs: list[dict[str, Any]], decision_failure_cap: float | None=0.49) -> dict[str, Any]:
    result = attach_default_policy(result)
    by_id = {str(row['id']): row for row in result.get('criteria', [])}
    reviews = []
    for spec in specs:
        criterion_id = str(spec['criterion_id'])
        criterion = by_id[criterion_id]
        criterion['semantic'] = True
        if criterion.get('category') == 'decision':
            criterion['weight'] = 10
            if decision_failure_cap is None:
                criterion.pop('failure_cap', None)
            else:
                criterion['failure_cap'] = decision_failure_cap
        reviews.append({'criterion_id': criterion_id, 'requirement': semantic_requirement(criterion_id=criterion_id, description=str(criterion['description']), expected_facts=spec.get('expected_facts'), artifact_type=artifact_type), 'hard_gate_met': bool(spec['hard_gate_met']), 'hard_gate_evidence': str(spec['hard_gate_evidence']), 'legacy_lexical_match': bool(criterion['value'])})
    result['semantic_review'] = {'version': 2, 'mode': 'deterministic_hard_gates_plus_bounded_semantic_judge', 'task_id': task_id, 'artifact': None, 'evidence': evidence[:60000], 'criteria': reviews, 'policy': 'Semantic wording cannot rescue a failed objective numeric, formula, or structure gate.'}
    return result

def _legacy_artifact_evidence(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == '.pptx':
        presentation = Presentation(path)
        chunks = []
        for (index, slide) in enumerate(presentation.slides, start=1):
            chunks.append(f'[Slide {index}]')
            for shape in slide.shapes:
                if getattr(shape, 'text', ''):
                    chunks.append(shape.text)
                if getattr(shape, 'has_table', False):
                    chunks.append('TABLE:')
                    chunks.extend((' | '.join((cell.text for cell in row.cells)) for row in shape.table.rows))
        return '\n'.join(chunks)[:60000]
    if suffix == '.docx':
        return _document_text(Document(path))[:60000]
    if suffix == '.xlsx':
        workbook = load_workbook(path, data_only=False, read_only=False)
        chunks = []
        for sheet in workbook.worksheets:
            chunks.append(f'[Sheet: {sheet.title}]')
            for row in sheet.iter_rows():
                values = [f'{cell.coordinate}={cell.value}' for cell in row if cell.value is not None]
                if values:
                    chunks.append(' | '.join(values))
        return '\n'.join(chunks)[:60000]
    return ''

def _formula_text(cell: Any) -> str:
    return str(getattr(cell, 'value', '') or '').lower().replace('$', '').replace("'", '').replace(' ', '')

def _formula_has_refs(cell: Any, *references: str) -> bool:
    formula = _formula_text(cell)
    return formula.startswith('=') and all((reference.lower().replace('$', '').replace("'", '').replace(' ', '') in formula for reference in references))

def _formula_has_large_literal(cell: Any) -> bool:
    formula = _formula_text(cell)
    if not formula.startswith('='):
        return False
    without_references = re.sub('(?:[a-z0-9_ ]+!)?[a-z]{1,3}[0-9]+', '', formula)
    return any((abs(float(token)) >= 10000 for token in re.findall('(?<![a-z])\\d+(?:\\.\\d+)?', without_references)))

def _recalculated_data_workbook(path: Path):
    """Return a data-only workbook after a safe LibreOffice recalc copy."""
    formula_workbook = load_workbook(path, data_only=False, read_only=False)
    original_values = load_workbook(path, data_only=True, read_only=False)

    def cached_formula_values(candidate) -> int:
        count = 0
        for sheet in formula_workbook.worksheets:
            if sheet.title not in candidate.sheetnames:
                continue
            value_sheet = candidate[sheet.title]
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        if value_sheet[cell.coordinate].value is not None:
                            count += 1
        return count

    def preserves_submitted_cache(candidate) -> bool:
        for sheet in formula_workbook.worksheets:
            if sheet.title not in original_values.sheetnames or sheet.title not in candidate.sheetnames:
                continue
            original_sheet = original_values[sheet.title]
            candidate_sheet = candidate[sheet.title]
            for row in sheet.iter_rows():
                for cell in row:
                    if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                        continue
                    if original_sheet[cell.coordinate].value is not None and candidate_sheet[cell.coordinate].value is None:
                        return False
        return True
    executable = shutil.which('libreoffice') or shutil.which('soffice')
    if not executable:
        return original_values
    with tempfile.TemporaryDirectory(prefix='arm-grade-xlsx-') as directory:
        root = Path(directory)
        source_dir = root / 'source'
        output_dir = root / 'output'
        source_dir.mkdir()
        output_dir.mkdir()
        source = source_dir / path.name
        shutil.copy2(path, source)
        try:
            subprocess.run([executable, f'-env:UserInstallation={root.as_uri()}/profile', '--headless', '--convert-to', 'xlsx', '--outdir', str(output_dir), str(source)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30, check=False, env={**os.environ, 'HOME': str(root)})
        except (OSError, subprocess.TimeoutExpired):
            return original_values
        recalculated = output_dir / path.name
        if recalculated.exists():
            recalculated_values = load_workbook(recalculated, data_only=True, read_only=False)
            if cached_formula_values(recalculated_values) >= cached_formula_values(original_values) and preserves_submitted_cache(recalculated_values):
                return recalculated_values
    return original_values

def _same_cells(left, right, sheet: str, ranges: list[str]) -> bool:
    (lws, rws) = (left[sheet], right[sheet])
    for cell_range in ranges:
        for (lrow, rrow) in zip(lws[cell_range], rws[cell_range]):
            for (lcell, rcell) in zip(lrow, rrow):
                if lcell.value != rcell.value:
                    return False
    return True

def _task_004_expected_row(ws: Any, row: int) -> dict[str, float] | None:
    inputs = {column: _number(ws[f'{column}{row}'].value) for column in 'CDEFKP'}
    if any((value is None for value in inputs.values())):
        return None
    contract = float(inputs['C'])
    posted_cost = float(inputs['D'])
    close_etc = float(inputs['E'] + inputs['F'])
    eac = posted_cost + close_etc
    if abs(eac) < 1e-12 or abs(contract) < 1e-12:
        return None
    percent_complete = min(posted_cost / eac, 1.0)
    earned = contract * percent_complete
    billings = float(inputs['K'])
    margin = contract - eac
    prior_margin = float(inputs['P'])
    return {'G': close_etc, 'H': eac, 'I': percent_complete, 'J': earned, 'L': max(earned - billings, 0.0), 'M': max(billings - earned, 0.0), 'N': margin, 'O': margin / contract, 'Q': (margin / contract - prior_margin) * 10000}

def _task_004_perturbation_checks(path: Path) -> tuple[dict[int, bool], bool, dict[int, bool], bool] | None:
    """Confirm submitted formulas still calculate correctly after source inputs move."""
    if not (shutil.which('libreoffice') or shutil.which('soffice')):
        return None
    try:
        with tempfile.TemporaryDirectory(prefix='task004-lineage-') as temporary:
            candidate = Path(temporary) / path.name
            shutil.copy2(path, candidate)
            workbook = load_workbook(candidate, data_only=False, read_only=False)
            sheet = workbook['Risk Review']
            for (offset, row) in enumerate(range(6, 10), start=1):
                for (column, delta) in {'C': 10000 * offset, 'D': 700 * offset, 'E': 300 * offset, 'F': 100 * offset, 'K': 500 * offset, 'P': 0.001 * offset}.items():
                    sheet[f'{column}{row}'] = float(sheet[f'{column}{row}'].value) + delta
            workbook.save(candidate)
            recalculated = _recalculated_data_workbook(candidate)
            value_sheet = recalculated['Risk Review']
            row_checks: dict[int, bool] = {}
            expected_rows: list[dict[str, float]] = []
            for row in range(6, 10):
                expected = _task_004_expected_row(value_sheet, row)
                if expected is None:
                    return None
                expected_rows.append(expected)
                row_checks[row] = all((_close(value_sheet[f'{column}{row}'].value, target, abs_tol=0.5 if column == 'Q' else 5e-05 if column in 'IO' else 0.05, rel_tol=0.0) for (column, target) in expected.items()))
            currency_columns = 'CDEFGHJKLMN'
            total_expected = {column: sum((float(value_sheet[f'{column}{row}'].value) for row in range(6, 10))) for column in currency_columns}
            total_expected.update({'I': total_expected['D'] / total_expected['H'], 'O': total_expected['N'] / total_expected['C'], 'P': sum((float(value_sheet[f'C{row}'].value) * float(value_sheet[f'P{row}'].value) for row in range(6, 10))) / total_expected['C']})
            total_expected['Q'] = (total_expected['O'] - total_expected['P']) * 10000
            total_ok = all((_close(value_sheet[f'{column}11'].value, target, abs_tol=0.5 if column == 'Q' else 5e-05 if column in 'IOP' else 0.05, rel_tol=0.0) for (column, target) in total_expected.items()))
            decision_row_checks: dict[int, bool] = {}
            for (source_row, decision_row) in zip(range(6, 10), range(15, 19), strict=True):
                pm_eac = float(value_sheet[f'D{source_row}'].value) + float(value_sheet[f'E{source_row}'].value)
                pm_margin = float(value_sheet[f'C{source_row}'].value) - pm_eac
                close_eac = float(value_sheet[f'H{source_row}'].value)
                close_margin = float(value_sheet[f'N{source_row}'].value)
                impact = close_margin - pm_margin
                expected_decision = {'B': pm_eac, 'C': pm_margin, 'D': close_eac, 'E': close_margin, 'F': impact}
                decision_row_checks[decision_row] = all((_close(value_sheet[f'{column}{decision_row}'].value, target, abs_tol=0.05, rel_tol=0.0) for (column, target) in expected_decision.items()))
            decision_control_expected = {'B21': sum((float(value_sheet[f'C{row}'].value) for row in range(15, 19))), 'D21': sum((float(value_sheet[f'E{row}'].value) for row in range(15, 19))), 'F21': sum((float(value_sheet[f'F{row}'].value) for row in range(15, 19)))}
            decision_control_ok = all((_close(value_sheet[coordinate].value, target, abs_tol=0.05, rel_tol=0.0) for (coordinate, target) in decision_control_expected.items()))
            return (row_checks, total_ok, decision_row_checks, decision_control_ok)
    except Exception:
        return None

def _task_004_review_flag_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    if expected['controller_review_required']:
        met = text in {'yes', 'required'} or any((token in text for token in ('review required', 'controller review', 'flag', 'escalat', 'hold', 'controller signoff', 'controller sign off')))
    else:
        met = text in {'no', 'monitor'} or any((token in text for token in ('below threshold', 'within threshold', 'no review', 'not required', 'monitor', 'release', 'clear')))
    if expected['netted_claim_or_backcharge']:
        met = met and any((token in text for token in ('backcharge', 'recovery', 'netted')))
    return met

def _task_004_treatment_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = str(value or '')
    normalized = _normalize(text)
    if expected['netted_claim_or_backcharge']:
        treatment_ok = any((token in normalized for token in ('backcharge', 'recovery'))) and any((token in normalized for token in ('do not net', 'not netted', 'de net', 'denet', 'add back', 'added back', 'restore', 'restored', 'remove from pm etc', 'removed from pm etc', 'cost added')))
        return treatment_ok and _text_contains_number(text, expected['cost_amount'], abs_tol=1.0)
    revenue_ok = any((token in normalized for token in ('exclude', 'excluded', 'omit', 'omitted', 'no revenue', 'not recognize', 'not recognised', 'unapproved', 'pending revenue'))) and any((token in normalized for token in ('revenue', 'recovery', 'change order', 'commercial value', 'pco'))) and _text_contains_number(text, expected['revenue_amount'], abs_tol=1.0)
    cost_ok = any((token in normalized for token in ('include', 'included', 'add', 'added', 'carry', 'carried', 'record', 'recorded', 'overlay', 'probable cost'))) and any((token in normalized for token in ('cost', 'exposure', 'etc', 'overlay')))
    return revenue_ok and cost_ok and _text_contains_number(text, expected['cost_amount'], abs_tol=1.0)

def _task_004_source_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    reference = _normalize(expected['reference'])
    return reference in text and ('fin rev 04' in text or 'signed wip policy' in text or 'wip policy' in text)

def _task_004_close_disposition_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    if expected['controller_review_required']:
        return any((token in text for token in ('hold', 'do not release', 'not release'))) and any((token in text for token in ('review', 'controller', 'signoff', 'sign off', 'approval', 'pending')))
    return any((token in text for token in ('release', 'clear'))) and any((token in text for token in ('monitor', 'follow up', 'follow-up', 'closeout')))

def _task_004_follow_up_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    return all((any((_normalize(token) in text for token in token_group)) for token_group in expected['required_follow_up_tokens']))

def _task_004_pending_revenue_excluded(value: Any) -> bool:
    if value is False or value == 0:
        return True
    text = _normalize(value)
    return text in {'no', 'none', 'zero'} or any((token in text for token in ('excluded', 'not included', 'no pending revenue')))

def _task_004_review_queue_ok(value: Any, expected: list[str]) -> bool:
    text = _normalize(value)
    expected_present = all((_normalize(project_id) in text for project_id in expected))
    if not expected_present or 'arm 2318' not in text:
        return expected_present
    return any((token in text for token in ('arm 2318 released', 'release arm 2318', 'arm 2318 cleared', 'clear arm 2318', 'except arm 2318', 'excluding arm 2318')))

def _task_004_close_release_ok(value: Any) -> bool:
    text = _normalize(value)
    return 'hold' in text and any((token in text for token in ('review', 'controller', 'signoff', 'sign off', 'approval', 'pending')))

def _task_004_decision_support(value_sheet: Any, gold: dict[str, Any]) -> dict[str, Any]:
    """Credit useful decision support without treating a wrong conclusion as correct."""
    support: dict[str, dict[str, Any]] = {}
    review_direction: dict[str, bool] = {}
    disposition_direction: dict[str, bool] = {}
    for (source_row, decision_row, review, decision) in zip(range(6, 10), range(15, 19), gold['review_rows'], gold['decision_rows'], strict=True):
        project_id = review['project_id']
        slug = project_id.casefold()
        treatment = str(value_sheet[f'S{source_row}'].value or '')
        controlling_support = str(value_sheet[f'T{source_row}'].value or '')
        treatment_support = f'{treatment}\n{controlling_support}'
        amounts_present = all((_text_contains_number(treatment, amount, abs_tol=1.0) for amount in {review['revenue_amount'], review['cost_amount']}))
        evidence_present = _task_004_source_ok(treatment_support, review)
        treatment_score = 0.5 if amounts_present and evidence_present else 0.25 if amounts_present or evidence_present else 0.0
        support[f'review__{slug}__commercial_treatment'] = {'score': treatment_score, 'evidence': f'expected amounts present={amounts_present}; controlling commercial and policy evidence present={evidence_present}'}
        flag_text = value_sheet[f'R{source_row}'].value
        direction_expected = bool(review['controller_review_required'])
        normalized_flag = _normalize(flag_text)
        direction_met = any((token in normalized_flag for token in ('review', 'hold', 'escalat', 'signoff', 'sign off'))) if direction_expected else any((token in normalized_flag for token in ('no review', 'not required', 'below threshold', 'within threshold', 'release', 'clear', 'monitor')))
        review_direction[project_id] = direction_met
        decision_support_ok = _normalize(value_sheet[f'A{decision_row}'].value) == _normalize(project_id) and _normalize(decision['policy_reference']) in _normalize(value_sheet[f'K{decision_row}'].value) and (_normalize(decision['commercial_reference']) in _normalize(value_sheet[f'L{decision_row}'].value))
        disposition_text = f"{value_sheet[f'I{decision_row}'].value or ''} {value_sheet[f'J{decision_row}'].value or ''}"
        disposition_met = _task_004_close_disposition_ok(disposition_text, decision)
        disposition_direction[project_id] = disposition_met
        support[f'decision_bridge__{slug}__close_disposition'] = {'score': 0.5 if direction_met and decision_support_ok else 0.25 if direction_met or decision_support_ok else 0.0, 'evidence': f'project-level review direction supported={direction_met}; policy and commercial references associated to decision row={decision_support_ok}'}
    excluded_count = sum((_task_004_pending_revenue_excluded(value_sheet[f'G{row}'].value) for row in range(15, 19)))
    support['decision_control__pending_revenue'] = {'score': round(0.5 * excluded_count / 4, 6), 'evidence': f'{excluded_count}/4 project bridge rows exclude pending revenue'}
    correct_review_directions = sum(review_direction.values())
    support['decision_control__review_queue'] = {'score': round(0.5 * correct_review_directions / 4, 6), 'evidence': f'{correct_review_directions}/4 project review directions support the aggregate queue'}
    correct_dispositions = sum(disposition_direction.values())
    support['decision_control__close_release'] = {'score': round(0.5 * correct_dispositions / 4, 6), 'evidence': f'{correct_dispositions}/4 project dispositions support the aggregate release decision'}
    return {'version': 1, 'criteria': support}

def _grade_task_004(workspace_root: Path) -> dict[str, Any]:
    relative = Path('Shared/Finance/Close/2026/06 June/4 WIP/WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx')
    path = workspace_root / relative
    gold = load_apex_gold('task_004')
    if not path.exists():
        return _file_failure('task_004', 'missing workbook')
    try:
        wb = load_workbook(path, data_only=False, read_only=False)
        original = load_workbook(SEED_WORKSPACE / relative, data_only=False, read_only=False)
        values = _recalculated_data_workbook(path)
    except Exception as exc:
        return _file_failure('task_004', str(exc))
    criteria: list[Criterion] = []
    for (sheet, ranges) in (('READ ME first', ['A5:B17']), ('Risk Review', ['A5:T5', 'A6:B9', 'A11:B11']), ('Evidence Map', ['A5:G9'])):
        met = _same_cells(wb, original, sheet, ranges)
        criteria.append(Criterion(f"preservation__{_normalize(sheet).replace(' ', '_')}", f'Protected {sheet!r} content remains unchanged', met, f'ranges={ranges!r}; preserved={met}'))
    value_sheet = values['Risk Review']
    for (row_number, expected) in enumerate(gold['source_inputs'], start=6):
        checks = {'current_contract': (value_sheet[f'C{row_number}'].value, expected['current_contract'], 0.05), 'cost_to_date': (value_sheet[f'D{row_number}'].value, expected['cost_to_date'], 0.05), 'pm_etc': (value_sheet[f'E{row_number}'].value, expected['pm_etc'], 0.05), 'documented_etc_overlay': (value_sheet[f'F{row_number}'].value, expected['documented_etc_overlay'], 0.05), 'billings': (value_sheet[f'K{row_number}'].value, expected['billings'], 0.05), 'prior_close_margin_percent': (value_sheet[f'P{row_number}'].value, expected['prior_close_margin_percent'], 5e-05)}
        for (label, (actual, target, tolerance)) in checks.items():
            met = _close(actual, target, abs_tol=tolerance, rel_tol=0.0)
            criteria.append(Criterion(f"source__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} source input `{label}` is correct", met, f'actual={actual!r}; expected={target}'))
    formula_ws = wb['Risk Review']
    perturbation = _task_004_perturbation_checks(path)
    required_references = {'G': ('E{row}', 'F{row}'), 'H': ('D{row}', 'G{row}'), 'I': ('D{row}', 'H{row}'), 'J': ('C{row}', 'I{row}'), 'L': ('J{row}', 'K{row}'), 'M': ('K{row}', 'J{row}'), 'N': ('C{row}', 'H{row}'), 'O': ('N{row}', 'C{row}'), 'Q': ('O{row}', 'P{row}')}
    for (row_number, expected) in enumerate(gold['projects'], start=6):
        cell_checks = []
        evidence = []
        for (column, references) in required_references.items():
            cell = formula_ws[f'{column}{row_number}']
            has_refs = _formula_has_refs(cell, *(reference.format(row=row_number) for reference in references))
            no_hardcoded_output = column == 'Q' or not _formula_has_large_literal(cell)
            cell_checks.append(has_refs and no_hardcoded_output)
            evidence.append(f'{cell.coordinate}={cell.value!r}')
        perturbation_met = perturbation is None or perturbation[0].get(row_number, False)
        criteria.append(Criterion(f"formula_lineage__{expected['project_id'].casefold()}", f"{expected['project_id']} calculations respond correctly to their source inputs", all(cell_checks) and perturbation_met, '; '.join(evidence) + f'; perturbation_check={perturbation_met}'))
    for (row_number, expected) in enumerate(gold['projects'], start=6):
        checks = {'close_etc': (value_sheet[f'G{row_number}'].value, expected['estimated_cost_to_complete'], 0.05), 'eac': (value_sheet[f'H{row_number}'].value, expected['estimated_cost_at_completion'], 0.05), 'earned': (value_sheet[f'J{row_number}'].value, expected['earned_revenue'], 0.05), 'asset': (value_sheet[f'L{row_number}'].value, expected['underbilling'], 0.05), 'liability': (value_sheet[f'M{row_number}'].value, expected['overbilling'], 0.05), 'margin': (value_sheet[f'N{row_number}'].value, expected['estimated_total_margin'], 0.05), 'percent_complete': (value_sheet[f'I{row_number}'].value, expected['percent_complete'], 5e-05), 'margin_percent': (value_sheet[f'O{row_number}'].value, expected['estimated_margin_percent'], 5e-05), 'prior_margin_percent': (value_sheet[f'P{row_number}'].value, expected['prior_margin_percent'], 5e-05), 'margin_movement_bps': (value_sheet[f'Q{row_number}'].value, expected['margin_movement_bps'], 0.5)}
        for (label, (actual, target, tolerance)) in checks.items():
            met = _close(actual, target, abs_tol=tolerance, rel_tol=0.0)
            criteria.append(Criterion(f"result__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} `{label}` is correct", met, f'actual={actual!r}; expected={target}'))
    for (row_number, expected) in enumerate(gold['review_rows'], start=6):
        flag = value_sheet[f'R{row_number}'].value
        treatment = value_sheet[f'S{row_number}'].value
        source = value_sheet[f'T{row_number}'].value
        for (label, met, evidence) in (('flag', _task_004_review_flag_ok(flag, expected), f'actual={flag!r}'), ('commercial_treatment', _task_004_treatment_ok(treatment, expected), f'actual={treatment!r}'), ('source', _task_004_source_ok(f"{treatment or ''}\n{source or ''}", expected), f'treatment={treatment!r}; controlling_support={source!r}')):
            criteria.append(Criterion(f"review__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} review `{label}` is professionally supported", met, evidence))
    total_lineage_checks = [_formula_has_refs(formula_ws[f'{column}11'], f'{column}6:{column}9') and (not _formula_has_large_literal(formula_ws[f'{column}11'])) for column in 'CDEFGHJKLMN']
    total_lineage_checks.extend([_formula_has_refs(formula_ws['I11'], 'D11', 'H11') and (not _formula_has_large_literal(formula_ws['I11'])), _formula_has_refs(formula_ws['O11'], 'N11', 'C11') and (not _formula_has_large_literal(formula_ws['O11'])), _formula_has_refs(formula_ws['P11'], 'C6:C9', 'P6:P9', 'C11') and (not _formula_has_large_literal(formula_ws['P11'])), _formula_has_refs(formula_ws['Q11'], 'O11', 'P11')])
    perturbation_total_met = perturbation is None or perturbation[1]
    criteria.append(Criterion('total_formula_lineage', 'Total/control row responds correctly to the four project rows', all(total_lineage_checks) and perturbation_total_met, f"formulas={[formula_ws[f'{column}11'].value for column in 'CDEFGHIJKLMNOPQ']!r}; perturbation_check={perturbation_total_met}"))
    for coordinate in [f'{column}11' for column in 'CDEFGHIJKLMNOPQ']:
        target = gold['totals'][coordinate]
        actual = value_sheet[coordinate].value
        column = coordinate[0]
        tolerance = 0.5 if column == 'Q' else 5e-05 if column in 'IOP' else 0.05
        criteria.append(Criterion(f'total_result__{coordinate.casefold()}', f'Total/control cell {coordinate} is correct', _close(actual, target, abs_tol=tolerance, rel_tol=0.0), f'actual={actual!r}; expected={target}'))
    decision_headers = ['Project ID', 'PM case EAC', 'PM case margin', 'Close case EAC', 'Close case margin', 'Margin impact', 'Pending revenue included?', 'Review required?', 'Close disposition', 'Required follow-up', 'Policy', 'Commercial support']
    decision_structure_met = [formula_ws.cell(14, column).value for column in range(1, 13)] == decision_headers and [formula_ws[f'A{row}'].value for row in range(15, 19)] == [row['project_id'] for row in gold['decision_rows']] and ([formula_ws[coordinate].value for coordinate in ('A21', 'C21', 'E21', 'A22', 'C22', 'E22', 'A23', 'A24', 'A25')] == ['Aggregate PM case margin', 'Aggregate close case margin', 'Aggregate margin impact', 'Pending commercial revenue included?', 'Review-required count', 'Largest overlay', 'Review queue', 'Close release decision', 'Policy basis'])
    criteria.append(Criterion('decision_bridge__structure', 'The workbook retains the controller decision-bridge and decision-control structure', decision_structure_met, f'headers={[formula_ws.cell(14, column).value for column in range(1, 13)]!r}'))
    for (source_row, decision_row, expected) in zip(range(6, 10), range(15, 19), gold['decision_rows'], strict=True):
        lineage_checks = (_formula_has_refs(formula_ws[f'B{decision_row}'], f'D{source_row}', f'E{source_row}'), _formula_has_refs(formula_ws[f'C{decision_row}'], f'C{source_row}', f'B{decision_row}'), _formula_has_refs(formula_ws[f'D{decision_row}'], f'H{source_row}'), _formula_has_refs(formula_ws[f'E{decision_row}'], f'N{source_row}'), _formula_has_refs(formula_ws[f'F{decision_row}'], f'E{decision_row}', f'C{decision_row}'))
        no_literals = all((not _formula_has_large_literal(formula_ws[f'{column}{decision_row}']) for column in 'BCDEF'))
        perturbation_met = perturbation is None or perturbation[2].get(decision_row, False)
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__formula_lineage", f"{expected['project_id']} PM-to-close decision bridge responds to source changes", all(lineage_checks) and no_literals and perturbation_met, f"formulas={[formula_ws[f'{column}{decision_row}'].value for column in 'BCDEF']!r}; perturbation_check={perturbation_met}"))
        for (label, actual, target) in (('pm_case_margin', value_sheet[f'C{decision_row}'].value, expected['pm_case_margin']), ('margin_impact', value_sheet[f'F{decision_row}'].value, expected['margin_impact'])):
            criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} decision bridge `{label}` is correct", _close(actual, target, abs_tol=0.05, rel_tol=0.0), f'actual={actual!r}; expected={target}'))
        disposition = value_sheet[f'I{decision_row}'].value
        follow_up = value_sheet[f'J{decision_row}'].value
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__close_disposition", f"{expected['project_id']} close disposition distinguishes release from controller hold", _task_004_close_disposition_ok(f"{disposition or ''} {follow_up or ''}", expected), f"disposition={disposition!r}; follow_up={follow_up!r}; review_required={expected['controller_review_required']}"))
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__required_follow_up", f"{expected['project_id']} required follow-up addresses the unresolved evidence", _task_004_follow_up_ok(follow_up, expected), f"actual={follow_up!r}; required_groups={expected['required_follow_up_tokens']!r}"))
    decision_control_lineage = all((_formula_has_refs(formula_ws['B21'], 'C15:C18'), _formula_has_refs(formula_ws['D21'], 'E15:E18'), _formula_has_refs(formula_ws['F21'], 'D21', 'B21'), _formula_has_refs(formula_ws['D22'], 'H15:H18'))) and all((not _formula_has_large_literal(formula_ws[coordinate]) for coordinate in ('B21', 'D21', 'F21', 'D22')))
    perturbation_decision_control_met = perturbation is None or perturbation[3]
    criteria.append(Criterion('decision_control__formula_lineage', 'The controller decision control responds to the project bridge', decision_control_lineage and perturbation_decision_control_met, f"formulas={[formula_ws[coordinate].value for coordinate in ('B21', 'D21', 'F21', 'D22')]!r}; perturbation_check={perturbation_decision_control_met}"))
    control = gold['decision_control']
    for (label, coordinate, target) in (('pm_case_margin', 'B21', control['pm_case_margin']), ('margin_impact', 'F21', control['margin_impact'])):
        actual = value_sheet[coordinate].value
        criteria.append(Criterion(f'decision_control__{label}', f'Decision control `{label}` is correct', _close(actual, target, abs_tol=0.05, rel_tol=0.0), f'actual={actual!r}; expected={target}'))
    pending_revenue = value_sheet['B22'].value
    criteria.append(Criterion('decision_control__pending_revenue', 'The control confirms that pending commercial revenue is excluded from the close case', _task_004_pending_revenue_excluded(pending_revenue), f'actual={pending_revenue!r}'))
    review_count = _number(value_sheet['D22'].value)
    criteria.append(Criterion('decision_control__review_count', 'The control identifies the number of projects requiring controller review', review_count is not None and _close(review_count, control['review_required_count'], abs_tol=0.0, rel_tol=0.0), f"actual={review_count!r}; expected={control['review_required_count']}"))
    largest_overlay = _normalize(value_sheet['F22'].value)
    criteria.append(Criterion('decision_control__largest_overlay', 'The control identifies the project with the largest finance overlay', _normalize(control['largest_overlay_project']) in largest_overlay, f"actual={value_sheet['F22'].value!r}; expected={control['largest_overlay_project']}"))
    review_queue = value_sheet['B23'].value
    criteria.append(Criterion('decision_control__review_queue', 'The control identifies the exact controller review queue', _task_004_review_queue_ok(review_queue, control['review_queue']), f"actual={review_queue!r}; expected={control['review_queue']!r}"))
    close_release = _normalize(value_sheet['B24'].value)
    criteria.append(Criterion('decision_control__close_release', 'The aggregate close-release decision holds the flagged jobs for controller review', _task_004_close_release_ok(close_release), f"actual={value_sheet['B24'].value!r}"))
    policy_basis = _normalize(value_sheet['B25'].value)
    criteria.append(Criterion('decision_control__policy_basis', 'The decision control cites the signed WIP policy', 'fin rev 04' in policy_basis or 'signed wip policy' in policy_basis, f"actual={value_sheet['B25'].value!r}"))
    result = _result(criteria)
    result['decision_support'] = _task_004_decision_support(value_sheet, gold)
    semantic_specs: list[dict[str, Any]] = []
    for (source_row, expected) in zip(range(6, 10), gold['review_rows'], strict=True):
        project_id = expected['project_id']
        project_ok = _normalize(value_sheet[f'A{source_row}'].value) == _normalize(project_id)
        flag = value_sheet[f'R{source_row}'].value
        treatment = str(value_sheet[f'S{source_row}'].value or '')
        source = str(value_sheet[f'T{source_row}'].value or '')
        amounts_present = all((_text_contains_number(treatment, amount, abs_tol=1.0) for amount in {expected['revenue_amount'], expected['cost_amount']}))
        evidence_present = _task_004_source_ok(f'{treatment}\n{source}', expected)
        slug = project_id.casefold()
        semantic_specs.extend([{'criterion_id': f'review__{slug}__flag', 'expected_facts': {'project_id': project_id, 'controller_review_required': expected['controller_review_required'], 'netted_claim_or_backcharge': expected['netted_claim_or_backcharge']}, 'hard_gate_met': project_ok and bool(_normalize(flag)), 'hard_gate_evidence': f'project_associated={project_ok}; review conclusion nonblank={bool(_normalize(flag))}'}, {'criterion_id': f'review__{slug}__commercial_treatment', 'expected_facts': {'project_id': project_id, 'pending_commercial_amount_excluded': expected['revenue_amount'], 'probable_cost_or_recovery_overlay': expected['cost_amount'], 'disputed_recovery_de_netted': expected['netted_claim_or_backcharge']}, 'hard_gate_met': project_ok and amounts_present and evidence_present, 'hard_gate_evidence': f'project_associated={project_ok}; expected_amounts_present={amounts_present}; commercial_and_policy_evidence_present={evidence_present}'}])
    for (decision_row, expected) in zip(range(15, 19), gold['decision_rows'], strict=True):
        project_id = expected['project_id']
        project_ok = _normalize(value_sheet[f'A{decision_row}'].value) == _normalize(project_id)
        policy_ok = _normalize(expected['policy_reference']) in _normalize(value_sheet[f'K{decision_row}'].value)
        commercial_ok = _normalize(expected['commercial_reference']) in _normalize(value_sheet[f'L{decision_row}'].value)
        evidence_ok = policy_ok and commercial_ok
        disposition = value_sheet[f'I{decision_row}'].value
        follow_up = value_sheet[f'J{decision_row}'].value
        slug = project_id.casefold()
        semantic_specs.extend([{'criterion_id': f'decision_bridge__{slug}__close_disposition', 'expected_facts': {'project_id': project_id, 'close_disposition': expected['close_disposition'], 'controller_review_required': expected['controller_review_required']}, 'hard_gate_met': project_ok and evidence_ok and bool(_normalize(disposition)), 'hard_gate_evidence': f'project_associated={project_ok}; policy_and_commercial_support={evidence_ok}; disposition_nonblank={bool(_normalize(disposition))}'}, {'criterion_id': f'decision_bridge__{slug}__required_follow_up', 'expected_facts': {'project_id': project_id, 'required_follow_up': expected['required_follow_up']}, 'hard_gate_met': project_ok and evidence_ok and bool(_normalize(follow_up)), 'hard_gate_evidence': f'project_associated={project_ok}; policy_and_commercial_support={evidence_ok}; follow_up_nonblank={bool(_normalize(follow_up))}'}])
    control = gold['decision_control']
    project_rows_ok = all((_normalize(value_sheet[f'A{row}'].value) == _normalize(expected['project_id']) for (row, expected) in zip(range(15, 19), gold['decision_rows'], strict=True)))
    pending_rows_present = all((bool(_normalize(value_sheet[f'G{row}'].value)) for row in range(15, 19)))
    review_count = _number(value_sheet['D22'].value)
    review_count_ok = review_count is not None and _close(review_count, control['review_required_count'], abs_tol=0.0, rel_tol=0.0)
    queue_has_expected = all((_normalize(project_id) in _normalize(value_sheet['B23'].value) for project_id in control['review_queue']))
    semantic_specs.extend([{'criterion_id': 'decision_control__pending_revenue', 'expected_facts': {'pending_commercial_revenue_included': False}, 'hard_gate_met': project_rows_ok and pending_rows_present, 'hard_gate_evidence': f'project_rows_associated={project_rows_ok}; per-project pending-revenue conclusions present={pending_rows_present}'}, {'criterion_id': 'decision_control__review_queue', 'expected_facts': {'held_for_controller_review': control['review_queue'], 'released_with_monitoring': 'ARM-2318'}, 'hard_gate_met': project_rows_ok and review_count_ok and queue_has_expected, 'hard_gate_evidence': f'project_rows_associated={project_rows_ok}; review_count_correct={review_count_ok}; expected_review_projects_present={queue_has_expected}'}, {'criterion_id': 'decision_control__close_release', 'expected_facts': {'close_release_decision': control['close_release_decision']}, 'hard_gate_met': project_rows_ok and review_count_ok and queue_has_expected and bool(_normalize(value_sheet['B24'].value)), 'hard_gate_evidence': f"project_rows_associated={project_rows_ok}; review_count_correct={review_count_ok}; review_queue_supported={queue_has_expected}; decision_nonblank={bool(_normalize(value_sheet['B24'].value))}"}])
    return _attach_semantic_review(result, task_id='task_004', evidence=_legacy_artifact_evidence(path), artifact_type='controller WIP workbook', specs=semantic_specs, decision_failure_cap=None)

def _slide_text(slide) -> str:
    return '\n'.join((getattr(shape, 'text', '') for shape in slide.shapes if getattr(shape, 'text', '')))

def _pptx_color_token(color: Any) -> str | None:
    try:
        rgb = color.rgb
    except (AttributeError, TypeError, ValueError):
        rgb = None
    if rgb is not None:
        return f'rgb:{str(rgb).upper()}'
    try:
        theme = color.theme_color
    except (AttributeError, TypeError, ValueError):
        theme = None
    if theme is None:
        return None
    try:
        brightness = round(float(color.brightness), 3)
    except (AttributeError, TypeError, ValueError):
        brightness = 0.0
    return f'theme:{theme}:{brightness}'

def _pptx_fill_signature(fill: Any) -> tuple[str | None, str | None]:
    try:
        fill_type = str(fill.type) if fill.type is not None else None
    except (AttributeError, TypeError, ValueError):
        fill_type = None
    try:
        color = _pptx_color_token(fill.fore_color) if fill_type is not None else None
    except (AttributeError, TypeError, ValueError):
        color = None
    if color in {None, 'rgb:FFFFFF'}:
        return (None, None)
    return ('color', color)

def _pptx_font_signature(font: Any) -> tuple[float | None, bool | None, bool | None, str, str | None]:
    try:
        size = float(round(float(font.size.pt))) if font.size is not None else None
    except (AttributeError, TypeError, ValueError):
        size = None
    underline_value = getattr(font, 'underline', None)
    underline = 'underlined' if underline_value not in {None, False} and 'NONE' not in str(underline_value).upper() else ''
    return (size, True if getattr(font, 'bold', None) is True else None, True if getattr(font, 'italic', None) is True else None, underline, _pptx_color_token(getattr(font, 'color', None)))

def _pptx_text_frame_signature(text_frame: Any) -> tuple[Any, ...]:
    paragraphs = []
    for paragraph in text_frame.paragraphs:
        runs = tuple(((_normalize(run.text), _pptx_font_signature(run.font)) for run in paragraph.runs))
        if not runs:
            runs = (('', _pptx_font_signature(paragraph.font)),)
        paragraphs.append((_normalize(paragraph.text), int(getattr(paragraph, 'level', 0) or 0), None if getattr(paragraph, 'alignment', None) is None or str(getattr(paragraph, 'alignment', None)).startswith('LEFT') else str(getattr(paragraph, 'alignment', None)), runs))
    return tuple(paragraphs)

def _pptx_shape_visual_signature(shape: Any, slide_width: int, slide_height: int) -> tuple[Any, ...]:
    geometry = (round(float(getattr(shape, 'left', 0) or 0) / max(slide_width, 1), 1), round(float(getattr(shape, 'top', 0) or 0) / max(slide_height, 1), 1), round(float(getattr(shape, 'width', 0) or 0) / max(slide_width, 1), 1), round(float(getattr(shape, 'height', 0) or 0) / max(slide_height, 1), 1))
    try:
        line_color = _pptx_color_token(shape.line.color)
        line_signature = (None, None, None if line_color in {None, 'rgb:FFFFFF'} else line_color)
    except (AttributeError, TypeError, ValueError):
        line_signature = (None, None, None)
    text_signature = _pptx_text_frame_signature(shape.text_frame) if getattr(shape, 'has_text_frame', False) else ()
    table_signature: tuple[Any, ...] = ()
    if getattr(shape, 'has_table', False):
        table_signature = tuple((tuple(((_normalize(cell.text), _pptx_fill_signature(cell.fill), _pptx_text_frame_signature(cell.text_frame)) for cell in row.cells)) for row in shape.table.rows))
    return (str(getattr(shape, 'shape_type', None)), geometry, _pptx_fill_signature(getattr(shape, 'fill', None)), line_signature, _normalize(getattr(shape, 'text', '')), text_signature, table_signature)

def _pptx_slide_visual_signature(slide: Any, slide_width: int, slide_height: int) -> tuple[Any, ...]:
    try:
        background = _pptx_fill_signature(slide.background.fill)
    except (AttributeError, TypeError, ValueError):
        background = (None, None)
    shapes = [_pptx_shape_visual_signature(shape, slide_width, slide_height) for shape in slide.shapes]
    return (background, tuple(sorted(shapes, key=repr)))

def _pptx_font_observations(text_frame: Any) -> list[tuple[float | None, bool | None, str | None]]:
    observations: list[tuple[float | None, bool | None, str | None]] = []
    for paragraph in text_frame.paragraphs:
        fonts = [run.font for run in paragraph.runs] or [paragraph.font]
        for font in fonts:
            (size, bold, _, _, color) = _pptx_font_signature(font)
            observations.append((size, bold, color))
    return observations

def _pptx_slide_colors(slide: Any) -> set[str]:
    colors: set[str] = set()

    def add(color: str | None) -> None:
        if color:
            colors.add(color)
    try:
        add(_pptx_color_token(slide.background.fill.fore_color))
    except (AttributeError, TypeError, ValueError):
        pass
    for shape in slide.shapes:
        try:
            add(_pptx_color_token(shape.fill.fore_color))
        except (AttributeError, TypeError, ValueError):
            pass
        try:
            add(_pptx_color_token(shape.line.color))
        except (AttributeError, TypeError, ValueError):
            pass
        if getattr(shape, 'has_text_frame', False):
            for (_, _, color) in _pptx_font_observations(shape.text_frame):
                add(color)
        if getattr(shape, 'has_table', False):
            for row in shape.table.rows:
                for cell in row.cells:
                    try:
                        add(_pptx_color_token(cell.fill.fore_color))
                    except (AttributeError, TypeError, ValueError):
                        pass
                    for (_, _, color) in _pptx_font_observations(cell.text_frame):
                        add(color)
    return colors

def _pptx_colors_match(first: str | None, second: str | None) -> bool:
    if first is None or second is None:
        return False
    if first == second:
        return True
    if not first.startswith('rgb:') or not second.startswith('rgb:'):
        return False
    try:
        first_rgb = tuple((int(first[index:index + 2], 16) for index in (4, 6, 8)))
        second_rgb = tuple((int(second[index:index + 2], 16) for index in (4, 6, 8)))
    except (TypeError, ValueError):
        return False
    return math.sqrt(sum(((left - right) ** 2 for (left, right) in zip(first_rgb, second_rgb)))) <= 36

def _pptx_background_color(slide: Any) -> str | None:
    try:
        return _pptx_color_token(slide.background.fill.fore_color)
    except (AttributeError, TypeError, ValueError):
        return None

def _pptx_task_015_style_checks(presentation: Any, original: Any, added: Any, table_shape: Any | None) -> tuple[bool, str, bool, str, bool, str]:
    if added is None:
        return (False, 'appended slide missing', False, 'appended slide missing', False, 'appended slide missing')
    reference_slides = list(original.slides)[1:] or list(original.slides)
    reference_backgrounds = {color for color in (_pptx_background_color(slide) for slide in reference_slides) if color}
    reference_colors = set().union(*(_pptx_slide_colors(slide) for slide in reference_slides))
    added_colors = _pptx_slide_colors(added)
    added_background = _pptx_background_color(added)
    distinctive_reference = {color for color in reference_colors if color not in reference_backgrounds and color not in {'rgb:FFFFFF', 'rgb:000000'}}
    matched_colors = {reference for reference in distinctive_reference if any((_pptx_colors_match(reference, candidate) for candidate in added_colors))}
    background_matches = any((_pptx_colors_match(added_background, reference) for reference in reference_backgrounds))
    visual_system_ok = background_matches and len(matched_colors) >= 2
    visual_evidence = f'background={added_background}; reference_backgrounds={sorted(reference_backgrounds)}; matched_reference_colors={sorted(matched_colors)}'
    title_shape = next((shape for shape in added.shapes if all((token in _normalize(getattr(shape, 'text', '')) for token in ('q2', 'covenant', 'headroom', 'posted', 'pro forma')))), None)
    reference_title_sizes: list[float] = []
    reference_title_tops: list[float] = []
    for slide in reference_slides:
        candidates: list[tuple[float, Any]] = []
        for shape in slide.shapes:
            if not _normalize(getattr(shape, 'text', '')):
                continue
            sizes = [size for (size, _, _) in _pptx_font_observations(shape.text_frame) if size is not None] if getattr(shape, 'has_text_frame', False) else []
            if sizes and float(getattr(shape, 'top', 0) or 0) / max(presentation.slide_height, 1) < 0.25:
                candidates.append((max(sizes), shape))
        if candidates:
            (size, shape) = max(candidates, key=lambda item: item[0])
            reference_title_sizes.append(size)
            reference_title_tops.append(float(shape.top) / max(presentation.slide_height, 1))
    sorted_sizes = sorted(reference_title_sizes)
    reference_title_size = sorted_sizes[len(sorted_sizes) // 2] if sorted_sizes else 20.0
    title_observations = _pptx_font_observations(title_shape.text_frame) if title_shape is not None and getattr(title_shape, 'has_text_frame', False) else []
    title_sizes = [size for (size, _, _) in title_observations if size is not None]
    title_bold = any((bold is True for (_, bold, _) in title_observations))
    title_top = float(title_shape.top) / max(presentation.slide_height, 1) if title_shape is not None else 1.0
    title_top_limit = min(0.25, max(reference_title_tops, default=0.17) + 0.08)
    title_hierarchy_ok = bool(title_shape is not None and title_sizes and (max(title_sizes) >= reference_title_size * 0.8) and title_bold and (title_top <= title_top_limit))
    title_evidence = f'title_present={title_shape is not None}; max_size={(max(title_sizes) if title_sizes else None)}; reference_size={reference_title_size}; bold={title_bold}; top_ratio={round(title_top, 3)}; top_limit={round(title_top_limit, 3)}'
    table_readability_ok = False
    table_evidence = 'real table missing'
    if table_shape is not None:
        table = table_shape.table
        header_cells = list(table.rows[0].cells) if len(table.rows) else []
        body_cells = [cell for row in list(table.rows)[1:] for cell in row.cells]
        header_fills = {color for cell in header_cells for color in [_pptx_fill_signature(cell.fill)[1]] if color}
        body_fills = {color for cell in body_cells for color in [_pptx_fill_signature(cell.fill)[1]] if color}
        header_bold = bool(header_cells) and all((any((bold is True for (_, bold, _) in _pptx_font_observations(cell.text_frame))) for cell in header_cells))
        header_fill_distinct = bool(header_fills) and (not all((any((_pptx_colors_match(header, body) for body in body_fills)) for header in header_fills)))
        table_properties = table._tbl.tblPr
        style_nodes = table_properties.xpath('./a:tableStyleId')
        uses_header_table_style = bool(table_properties.get('firstRow') == '1' and style_nodes and str(style_nodes[0].text or '').strip())
        within_slide = table_shape.left >= 0 and table_shape.top >= 0 and (table_shape.left + table_shape.width <= presentation.slide_width) and (table_shape.top + table_shape.height <= presentation.slide_height)
        below_title = bool(title_shape is not None and table_shape.top >= title_shape.top + title_shape.height)
        table_readability_ok = bool(len(table.rows) >= 4 and len(table.columns) >= 5 and within_slide and below_title and (header_bold or header_fill_distinct or uses_header_table_style))
        table_evidence = f'rows={len(table.rows)}; columns={len(table.columns)}; within_slide={within_slide}; below_title={below_title}; header_bold={header_bold}; header_fill_distinct={header_fill_distinct}; uses_header_table_style={uses_header_table_style}'
    return (visual_system_ok, visual_evidence, title_hierarchy_ok, title_evidence, table_readability_ok, table_evidence)

def _grade_task_015(workspace_root: Path) -> dict[str, Any]:
    relative = Path('Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 lender update - review working v3.pptx')
    path = workspace_root / relative
    gold = load_apex_gold('task_015')
    if not path.exists():
        return _file_failure('task_015', 'missing deck')
    try:
        presentation = Presentation(str(path))
        original = Presentation(str(SEED_WORKSPACE / relative))
    except Exception as exc:
        return _file_failure('task_015', str(exc))
    original_texts = [_normalize(_slide_text(slide)) for slide in original.slides]
    current_texts = [_normalize(_slide_text(slide)) for slide in presentation.slides]
    original_signatures = [_pptx_slide_visual_signature(slide, original.slide_width, original.slide_height) for slide in original.slides]
    current_signatures = [_pptx_slide_visual_signature(slide, presentation.slide_width, presentation.slide_height) for slide in presentation.slides]
    added = presentation.slides[-1] if len(presentation.slides) > len(original.slides) else None
    added_text = _normalize(_slide_text(added)) if added else ''
    title_ok = 'q2 covenant headroom' in added_text and 'posted' in added_text and ('pro forma' in added_text)
    table_shapes = [shape for shape in added.shapes if getattr(shape, 'has_table', False)] if added else []
    tables = [shape.table for shape in table_shapes]
    has_table = len(tables) == 1
    (visual_system_ok, visual_system_evidence, title_hierarchy_ok, title_hierarchy_evidence, table_readability_ok, table_readability_evidence) = _pptx_task_015_style_checks(presentation, original, added, table_shapes[0] if has_table else None)
    metric_failures: list[str] = []
    metric_results: list[Criterion] = []
    joined = ''
    if has_table:
        table = tables[0]
        table_rows = [[cell.text for cell in row.cells] for row in table.rows]
        joined = ' | '.join((value for row in table_rows for value in row))
        headers = [_normalize(value) for value in table_rows[0]] if table_rows else []

        def column_index(*aliases: str) -> int | None:
            return next((index for (index, header) in enumerate(headers) if any((_normalize(alias) == header for alias in aliases))), None)
        posted_column = column_index('posted')
        pro_forma_column = column_index('pro forma', 'pro-forma')
        threshold_column = column_index('threshold')
        status_column = column_index('status')
        metric_targets = {'leverage': (gold['leverage_posted'], gold['leverage_pro_forma'], gold['max_leverage']), 'fccr': (gold['fccr_posted'], gold['fccr_pro_forma'], gold['min_fccr']), 'tangible net worth': (gold['tangible_net_worth_posted'], gold['tangible_net_worth_pro_forma'], gold['min_tangible_net_worth'])}
        for (metric, expected) in metric_targets.items():
            row = next((values for values in table_rows[1:] if metric in _normalize(values[0])), None)
            if row is None or None in (posted_column, pro_forma_column, threshold_column, status_column):
                metric_failures.append(f'{metric}: missing row/columns')
                for field in ('posted', 'pro_forma', 'threshold', 'status'):
                    metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__{field}", f'{metric.title()} `{field}` is correctly presented', False, 'missing row or required column'))
                continue
            for (field, column, target) in zip(('posted', 'pro_forma', 'threshold'), (posted_column, pro_forma_column, threshold_column), expected):
                met = _text_contains_number(row[column], target, abs_tol=0.0001)
                if not met:
                    metric_failures.append(f'{metric} {field}: value={row[column]!r}')
                metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__{field}", f'{metric.title()} `{field}` value is correct', met, f'actual={row[column]!r}; expected={target}'))
            status_met = 'compliant' in _normalize(row[status_column])
            if not status_met:
                metric_failures.append(f'{metric}: status={row[status_column]!r}')
            metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__status", f'{metric.title()} status is correctly reported as compliant', status_met, f'actual={row[status_column]!r}'))
    else:
        for metric in ('leverage', 'fccr', 'tangible net worth'):
            metric_id = _normalize(metric).replace(' ', '_')
            for field in ('posted', 'pro_forma', 'threshold', 'status'):
                metric_results.append(Criterion(f'metric__{metric_id}__{field}', f'{metric.title()} `{field}` is correctly presented', False, 'required appended covenant table is missing'))
    numbers_ok = has_table and (not metric_failures)
    status_ok = 'proposed' in added_text and ('unposted' in added_text or 'not posted' in added_text) and ('compliant' in added_text)
    result = _result([Criterion('preservation__slide_count', 'Exactly one slide is appended to the original deck', len(presentation.slides) == len(original.slides) + 1, f'original={len(original.slides)}; current={len(presentation.slides)}'), *[Criterion(f'preservation__slide_{index:02d}', f'Original slide {index} retains its content, layout, and styling', index <= len(current_signatures) and current_texts[index - 1] == text and (current_signatures[index - 1] == original_signatures[index - 1]), f'text_match={index <= len(current_texts) and current_texts[index - 1] == text}; visual_signature_match={index <= len(current_signatures) and current_signatures[index - 1] == original_signatures[index - 1]}') for (index, text) in enumerate(original_texts, start=1)], Criterion('structure__title', 'The appended slide has the requested title', title_ok, f'title_ok={title_ok}'), Criterion('structure__table', 'The appended slide contains exactly one real PowerPoint table', has_table, f'tables={len(tables)}'), Criterion('style__visual_system', 'The appended slide uses the existing deck background and visual palette', visual_system_ok, visual_system_evidence), Criterion('style__title_hierarchy', "The appended slide title follows the deck's established hierarchy", title_hierarchy_ok, title_hierarchy_evidence), Criterion('style__table_readability', 'The covenant comparison is presented as a bounded, clearly headed review table', table_readability_ok, table_readability_evidence), *metric_results, Criterion('status__proposed', 'The WIP entry is identified as proposed', 'proposed' in added_text, added_text[:800]), Criterion('status__unposted', 'The WIP entry is identified as unposted', 'unposted' in added_text or 'not posted' in added_text, added_text[:800]), Criterion('status__compliant', 'Both covenant bases are reported as compliant', 'compliant' in added_text, added_text[:800])])
    numeric_gate = has_table and all((any((_text_contains_number(joined, value, abs_tol=0.0001) for value in targets)) for targets in ((gold['leverage_posted'],), (gold['leverage_pro_forma'],), (gold['max_leverage'],), (gold['fccr_posted'],), (gold['fccr_pro_forma'],), (gold['min_fccr'],), (gold['tangible_net_worth_posted'],), (gold['tangible_net_worth_pro_forma'],), (gold['min_tangible_net_worth'],))))
    return _attach_semantic_review(result, task_id='task_015', evidence=_legacy_artifact_evidence(path), artifact_type='board presentation', specs=[{'criterion_id': 'structure__title', 'expected_facts': {'title': 'Q2 Covenant Headroom — Posted vs Pro Forma'}, 'hard_gate_met': has_table, 'hard_gate_evidence': f'real_table={has_table}'}, {'criterion_id': 'status__compliant', 'expected_facts': {'all_three_covenants': 'compliant'}, 'hard_gate_met': numeric_gate, 'hard_gate_evidence': f'all required numeric facts present in table={numeric_gate}'}, {'criterion_id': 'status__proposed', 'expected_facts': {'wip_status': 'proposed'}, 'hard_gate_met': bool(added), 'hard_gate_evidence': f'appended_slide_present={bool(added)}'}, {'criterion_id': 'status__unposted', 'expected_facts': {'wip_status': 'unposted'}, 'hard_gate_met': bool(added), 'hard_gate_evidence': f'appended_slide_present={bool(added)}'}, *[{'criterion_id': f"metric__{_normalize(metric).replace(' ', '_')}__status", 'expected_facts': {metric: 'compliant'}, 'hard_gate_met': numeric_gate, 'hard_gate_evidence': f'all required covenant numeric facts present={numeric_gate}'} for metric in ('leverage', 'fccr', 'tangible net worth')]])

def _document_text(document: Document) -> str:
    chunks = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            chunks.extend((cell.text for cell in row.cells))
    return '\n'.join(chunks)
TASK_001_ARTIFACT = Path('Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx')

def _docx_table_rows(document: Document, required_headers: tuple[str, ...]) -> list[dict[str, str]]:
    required = {_normalize(header) for header in required_headers}
    for table in document.tables:
        if not table.rows:
            continue
        headers = [_normalize(cell.text) for cell in table.rows[0].cells]
        if not required.issubset(set(headers)):
            continue
        rows: list[dict[str, str]] = []
        for row in table.rows[1:]:
            rows.append({header: cell.text.strip() for (header, cell) in zip(headers, row.cells)})
        return rows
    return []

def _docx_named_row(rows: list[dict[str, str]], first_header: str, label: str) -> dict[str, str]:
    wanted = _normalize(label)
    header = _normalize(first_header)
    for row in rows:
        if _normalize(row.get(header)) == wanted:
            return row
    return {}

def _usable_docx_value(value: Any) -> bool:
    text = str(value or '').strip()
    return bool(text) and (not bool(re.search('\\[(?:enter|select|complete|cite|record|explain)\\b', text, flags=re.I)))

def _docx_numeric_value(value: Any) -> Any:
    text = str(value or '').strip()
    return re.sub('\\s*(?:bps|basis points?)\\s*$', '', text, flags=re.I)

def _task_001_document_mapping(document: Document) -> dict[str, Any]:
    mapping: dict[str, Any] = {}
    current_rows = _docx_table_rows(document, ('Metric', 'Result', 'Source / comment'))
    current_keys = {'Current contract': 'current_contract', 'Posted cost': 'posted_cost', 'Billings': 'billings', 'PM ETC': 'pm_etc', 'Finance ETC correction': 'required_etc_adjustment', 'Close ETC': 'close_etc', 'EAC': 'eac', 'Percent complete': 'percent_complete', 'Earned revenue': 'earned_revenue', 'Contract asset': 'contract_asset', 'Contract liability': 'contract_liability', 'Estimated margin at completion': 'estimated_margin', 'Margin rate': 'margin_percent'}
    for (label, key) in current_keys.items():
        value = _docx_named_row(current_rows, 'Metric', label).get('result')
        if _usable_docx_value(value):
            mapping[key] = _docx_numeric_value(value)
    bridge_rows = _docx_table_rows(document, ('Metric', 'May final', 'June close', 'Change', 'Driver'))
    bridge_keys = {'Current contract': ('may_current_contract', 'current_contract_change'), 'Posted cost': ('may_posted_cost', 'posted_cost_change'), 'Remaining cost / ETC': ('may_close_etc', 'close_etc_change'), 'EAC': ('may_eac', 'eac_change'), 'Percent complete': ('may_percent_complete', 'percent_complete_change_bps'), 'Earned revenue': ('may_earned_revenue', 'earned_revenue_change'), 'Billings': ('may_billings', 'billings_change'), 'Contract asset': ('may_contract_asset', 'contract_asset_change'), 'Contract liability': ('may_contract_liability', 'contract_liability_change'), 'Estimated margin at completion': ('may_estimated_margin', 'estimated_margin_change'), 'Margin rate': ('may_margin_percent', 'margin_rate_change_bps')}
    for (label, (may_key, change_key)) in bridge_keys.items():
        row = _docx_named_row(bridge_rows, 'Metric', label)
        may_value = row.get('may final')
        change_value = row.get('change')
        if _usable_docx_value(may_value):
            mapping[may_key] = _docx_numeric_value(may_value)
        if _usable_docx_value(change_value):
            mapping[change_key] = _docx_numeric_value(change_value)
    conclusions = _docx_table_rows(document, ('Question', 'Conclusion', 'Support'))
    contract_treatment = _docx_named_row(conclusions, 'Question', 'Contract value treatment').get('conclusion', '')
    if _usable_docx_value(contract_treatment):
        normalized = _normalize(contract_treatment)
        if 'exclude' in normalized or 'outside' in normalized:
            mapping['draft_change_contract_treatment'] = 'exclude'
        elif 'include' in normalized:
            mapping['draft_change_contract_treatment'] = 'include'
    etc_treatment = _docx_named_row(conclusions, 'Question', 'PM ETC treatment').get('conclusion', '')
    if _usable_docx_value(etc_treatment):
        normalized = _normalize(etc_treatment)
        if 'restore' in normalized or 'add back' in normalized:
            mapping['disputed_recovery_etc_action'] = 'add_back_to_pm_etc'
        elif 'leave' in normalized or 'retain' in normalized:
            mapping['disputed_recovery_etc_action'] = 'leave_in_pm_etc'
    review = _docx_named_row(conclusions, 'Question', 'Controller review').get('conclusion', '')
    if _usable_docx_value(review):
        normalized = _normalize(review)
        mapping['controller_review_required'] = not bool(re.search('\\b(?:not|required no|unnecessary)\\b', normalized))
    return mapping

def _criterion_from_row(row: dict[str, Any]) -> Criterion:
    return Criterion(str(row['id']), str(row['description']), bool(row.get('value')), str(row.get('evidence', '')), category=str(row.get('category') or 'core_finance'), weight=int(row.get('weight', 10)), semantic=bool(row.get('semantic')), failure_cap=None if row.get('failure_cap') is None else float(row['failure_cap']))

def _grade_task_001(workspace_root: Path, answer: Any) -> dict[str, Any]:
    path = workspace_root / TASK_001_ARTIFACT
    document: Document | None = None
    parse_error = ''
    if path.is_file():
        try:
            document = Document(path)
        except Exception as exc:
            parse_error = f'{type(exc).__name__}: {exc}'
    else:
        parse_error = 'artifact missing'
    mapping = dict(_answer_mapping(answer))
    if document is not None:
        mapping.update(_task_001_document_mapping(document))
    legacy = _grade_numeric('task_001', mapping)
    criteria = [_criterion_from_row(row) for row in legacy['criteria'] if not str(row['id']).startswith('format__')]
    text = _document_text(document) if document is not None else ''
    normalized = _normalize(text)
    placeholder_pattern = re.compile('\\[(?:enter|select|complete|cite|record|explain)\\b', flags=re.I)
    placeholders = placeholder_pattern.findall(text)
    metadata_rows = _docx_table_rows(document, ('Field', 'Value')) if document is not None else []
    current_rows = _docx_table_rows(document, ('Metric', 'Result', 'Source / comment')) if document is not None else []
    bridge_rows = _docx_table_rows(document, ('Metric', 'May final', 'June close', 'Change', 'Driver')) if document is not None else []
    conclusion_rows = _docx_table_rows(document, ('Question', 'Conclusion', 'Support')) if document is not None else []
    evidence_rows = _docx_table_rows(document, ('Control / conclusion', 'Controlling source', 'Source fact', 'Status')) if document is not None else []
    action_rows = _docx_table_rows(document, ('Owner', 'Action', 'Completion evidence', 'Status')) if document is not None else []

    def row_text(rows: list[dict[str, str]], header: str, label: str) -> str:
        return ' | '.join(_docx_named_row(rows, header, label).values())
    recommendation = ''
    if document is not None and len(document.tables) >= 2:
        callout = document.tables[1]
        if len(callout.rows) == 1 and len(callout.columns) == 1:
            recommendation = callout.cell(0, 0).text.strip()
    review_text = row_text(conclusion_rows, 'Question', 'Controller review')
    posting_text = row_text(conclusion_rows, 'Question', 'Close and posting status')
    accounting_text = row_text(evidence_rows, 'Control / conclusion', 'Current accounting balances')
    prior_text = row_text(evidence_rows, 'Control / conclusion', 'Prior-close comparison')
    forecast_text = row_text(evidence_rows, 'Control / conclusion', 'Current PM forecast')
    commercial_text = row_text(evidence_rows, 'Control / conclusion', 'Commercial authorization and recovery')
    policy_text = row_text(evidence_rows, 'Control / conclusion', 'Revenue-recognition and review policy')
    finance_action = row_text(action_rows, 'Owner', 'Finance / Project Accounting')
    commercial_action = row_text(action_rows, 'Owner', 'Commercial / Project Team')
    controller_action = row_text(action_rows, 'Owner', 'Controller')
    required_tables = (bool(metadata_rows), len(current_rows) == 13, len(bridge_rows) == 11, len(conclusion_rows) == 5, len(evidence_rows) == 5, len(action_rows) == 3)
    all_log_cells_complete = bool(evidence_rows) and all((all((_usable_docx_value(value) for value in row.values())) for row in evidence_rows))
    criteria.extend([Criterion('artifact__complete', 'The controller sign-off memo is completed without template placeholders', document is not None and (not placeholders), f'path={path}; parse_error={parse_error!r}; placeholder_count={len(placeholders)}', category='integrity', weight=10, semantic=False), Criterion('identity__arm_2409', 'The memo identifies ARM-2409, the project, and the June 30 cutoff', all((token in normalized for token in ('arm 2409', 'northline cold storage expansion', 'june 30 2026'))), 'identity tokens inspected in the submitted document', category='controls', weight=5, semantic=False), Criterion('structure__review_tables', 'The memo retains all required review tables and row populations', all(required_tables), f'required_table_checks={required_tables}', category='controls', weight=5, semantic=False), Criterion('recommendation__policy_corrected_close', 'The recommendation states the policy-corrected close treatment and approval state', _text_contains_number(recommendation, 185000) and all((token in _normalize(recommendation) for token in ('restor', 'disputed recovery', 'etc', 'controller review', 'not approved to post'))), recommendation[:1000], category='decision', weight=10, semantic=True), Criterion('decision__margin_review_trigger', 'The memo connects controller review to the 334.78-basis-point margin decline', _text_contains_number(review_text, 334.78) and contains_concept(review_text, 'controller review required'), review_text[:1000], category='decision', weight=10, semantic=True), Criterion('decision__disputed_estimate_review_trigger', 'The memo connects controller review to the disputed recovery embedded in the PM estimate', contains_concept(review_text, 'disputed recovery') and contains_concept(review_text, 'PM estimate'), review_text[:1000], category='decision', weight=10, semantic=True), Criterion('evidence__accounting_cutoff', 'The evidence log ties current contract, posted cost, and billings to cutoff accounting records', contains_concept(accounting_text, 'company accounting records') and all((token in _normalize(accounting_text) for token in ('current contract', 'posted cost', 'billings'))), accounting_text[:1000], category='provenance', weight=10, semantic=True), Criterion('evidence__prior_close', 'The evidence log identifies the final May WIP source and its comparison purpose', 'wip 5 31 26 final v7 revised nb' in _normalize(prior_text) and contains_concept(prior_text, 'prior close comparison'), prior_text[:1000], category='provenance', weight=10, semantic=True), Criterion('evidence__pm_forecast', 'The evidence log identifies the current PM forecast and the disputed-recovery condition', 'pm etc updates 6 29 530pm combined v3' in _normalize(forecast_text) and _text_contains_number(forecast_text, 584000) and contains_concept(forecast_text, 'disputed recovery'), forecast_text[:1000], category='provenance', weight=10, semantic=True), Criterion('evidence__commercial_authorization', 'The evidence log identifies both commercial sources and the unsupported $185,000 recovery', '2409 pco11 backcharge backup draft2' in _normalize(commercial_text) and 'co log master' in _normalize(commercial_text) and _text_contains_number(commercial_text, 185000) and all((token in _normalize(commercial_text) for token in ('draft', 'unaccepted', 'disputed', 'not an executed contract modification'))), commercial_text[:1000], category='provenance', weight=10, semantic=True), Criterion('evidence__wip_policy', 'The evidence log identifies the signed WIP policy and its applicable treatment', 'revenue recognition wip policy rev11 24 signed' in _normalize(policy_text) and all((token in _normalize(policy_text) for token in ('exclude pending recovery revenue', 'include probable unavoidable cost', 'escalate'))), policy_text[:1000], category='provenance', weight=10, semantic=True), Criterion('decision__future_commercial_evidence', 'The memo states the evidence needed before the recovery may enter contract value', contains_concept(commercial_action, 'executed or accepted resolution') and contains_concept(commercial_action, 'collectability assessment'), commercial_action[:1000], category='decision', weight=10, semantic=True), Criterion('action__finance_correction', 'Finance owns the ETC correction, recalculation, and retained close support', _text_contains_number(finance_action, 185000) and all((token in _normalize(finance_action) for token in ('restore', 'remaining etc', 'recalculate the june wip', 'retain', 'close file'))), finance_action[:1000], category='decision', weight=10, semantic=True), Criterion('action__commercial_resolution', 'The project team owns resolution of PCO-011 and the authorization evidence', 'pco 011' in _normalize(commercial_action) and all((token in _normalize(commercial_action) for token in ('outside contract value', 'executed or accepted resolution', 'collectability'))), commercial_action[:1000], category='decision', weight=10, semantic=True), Criterion('action__controller_disposition', 'The controller owns review and documented disposition before posting', _text_contains_number(controller_action, 334.78) and all((token in _normalize(controller_action) for token in ('disputed recovery treatment', 'approve or return', 'documented controller disposition', 'before any wip posting'))), controller_action[:1000], category='decision', weight=10, semantic=True), Criterion('decision__not_approved_to_post', 'The memo remains a preparer sign-off and does not claim posting approval', 'not approved to post' in _normalize(posting_text) and all((token in _normalize(posting_text) for token in ('controller approval is required', 'before', 'posted'))), posting_text[:1000], category='decision', weight=10, semantic=True), Criterion('evidence__log_complete', 'Every required evidence-log row records a source, source fact, and status', len(evidence_rows) == 5 and all_log_cells_complete, f'rows={len(evidence_rows)}; all_cells_complete={all_log_cells_complete}', category='auditability', weight=10, semantic=False), Criterion('usability__review_ready', 'The memo is organized as a usable controller review record', document is not None and all(required_tables) and (not placeholders) and (len(recommendation.strip()) >= 80), f'tables_complete={all(required_tables)}; recommendation_chars={len(recommendation.strip())}', category='controls', weight=5, semantic=False)])
    result = _result(criteria)
    semantic_ids = {criterion.id for criterion in criteria if criterion.semantic}
    return _attach_semantic_review(result, task_id='task_001', evidence=_legacy_artifact_evidence(path) if document is not None else '', artifact_type='document', specs=[{'criterion_id': criterion.id, 'expected_facts': {'requirement': criterion.description}, 'hard_gate_met': criterion.met, 'hard_gate_evidence': criterion.evidence} for criterion in criteria if criterion.id in semantic_ids], decision_failure_cap=None)

def _display_rounding_tolerance(token: str, *, floor: float) -> float:
    """Accept a number when it rounds to the professional display precision used."""
    clean = token.strip().strip('()')
    suffix_match = re.search('([kmb])\\s*$', clean, flags=re.I)
    multiplier = {'k': 1000.0, 'm': 1000000.0, 'b': 1000000000.0}.get(suffix_match.group(1).lower() if suffix_match else '', 1.0)
    if suffix_match:
        clean = clean[:suffix_match.start()]
    percent = clean.rstrip().endswith('%')
    clean = clean.rstrip('%xX× ')
    decimal_match = re.search('\\.([0-9]+)$', clean.replace(',', ''))
    decimals = len(decimal_match.group(1)) if decimal_match else 0
    tolerance = 0.5 * multiplier * 10 ** (-decimals)
    if percent:
        tolerance /= 100
    return max(floor, tolerance + 1e-09)

def _text_contains_number(text: str, target: float, *, abs_tol: float=0.02) -> bool:
    tokens = re.findall('\\(?-?\\$?[0-9][0-9,]*(?:\\.[0-9]+)?(?:[kmb]|x|%)?\\)?', text, flags=re.I)
    return any((_close(token, target, abs_tol=_display_rounding_tolerance(token, floor=abs_tol), rel_tol=2e-06) for token in tokens))

def grade_apex_task(task_id: str, answer: Any, workspace_root: str | Path) -> dict[str, Any]:
    if task_id == 'task_001':
        result = _grade_task_001(Path(workspace_root), answer)
    elif task_id == 'task_004':
        result = _canonicalize_legacy_file_policy(task_id, _grade_task_004(Path(workspace_root)))
    elif task_id == 'task_015':
        result = _canonicalize_legacy_file_policy(task_id, _grade_task_015(Path(workspace_root)))
    elif task_id in CORPORATE_TASK_IDS:
        from runtime.grading.corporate_finance import grade_corporate_finance_task
        result = grade_corporate_finance_task(task_id, answer, workspace_root)
    else:
        raise KeyError(f'Task is not part of this sample: {task_id}')
    if task_id in TASK_GRADING_REVISIONS:
        result.setdefault('task_grading_revision', dict(TASK_GRADING_REVISIONS[task_id]))
    return apply_reward_policy(task_id, attach_default_policy(result))
