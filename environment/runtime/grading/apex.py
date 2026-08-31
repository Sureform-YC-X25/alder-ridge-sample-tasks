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
from runtime.accounting_mcp.paths import resolve_project_root, resolve_seed_root
CORPORATE_TASK_IDS = frozenset(['task_027', 'task_035', 'task_037', 'task_055', 'task_061', 'task_068', 'task_072', 'task_073', 'task_100'])

GOLD_PATH = Path(__file__).resolve().parent / 'gold' / 'tasks_001_025.json'
SEED_WORKSPACE = resolve_seed_root(__file__) / 'sources'
TASK_GRADING_REVISIONS = {'task_001': {'id': 'task-001-review-ready-controller-signoff-v24', 'effective_date': '2026-08-31', 'basis': 'concise professional request and realistic blank-input controller memo; all material facts are discoverable in the supplied close records and authorized accounting data; deterministic finance checks accept normal rounding, units, sign presentation, reordered tables, and equivalent row and column labels; open-ended conclusions, source authority, actions, and approval boundaries receive criterion-scoped semantic review that accepts normal business wording; June posting completeness is demonstrated through AP, payroll, billing, and aggregate controls rather than a hidden 23-line register; the disputed recovery is graded at its supported total without a required subjective component allocation; seven material workstreams receive stable proportional section weights with no phrase, formatting, or central-section reward cap'}, 'task_004': {'basis': 'a non-presolved controller workpaper removes the source map, policy recipe, and expected control counts while preserving a realistic four-project review shell; project-level WIP finance and release decisions with stable proportional scoring; substantive source and calculation errors remain deductions while harmless workbook formatting changes cannot erase otherwise valid work, and abbreviated professional source references reach independently scoped semantic review with exact submitted cells and source-derived answer keys; the project queue, release decision, and each follow-up are scored independently from separately checked counts and neighboring support fields; a source-backed, formula-driven unbooked commercial-recovery sensitivity distinguishes booked close results from hypothetical fully authorized recoveries; a separately weighted accounting-to-workpaper control rolls May/base balances through June activity for posted job cost, non-voided billings, and original contract plus approved changes, then ties each project and the portfolio total to the WIP review', 'effective_date': '2026-08-29', 'id': 'task-004-evidence-discovery-wip-review-v8'}, 'task_015': {'basis': "an objective-led lender request requires independent discovery and reperformance; all material assumptions, including the downside ladder, remain discoverable in the workspace rather than disclosed as a prompt recipe; source-complete covenant calculations using the working scaffold and executed agreement, an explicit $1,276,325.70 proposed-WIP basis and FCCR/TNW conventions, posted and pro-forma dollar deterioration capacity, binding-covenant identification, normal compliance-status wording, readable native PowerPoint structure, and content-preserving rather than pixel-identical handling of the existing lender deck, with proportional scoring and no single formatting or status-wording failure cap; each status and conclusion is judged from only its relevant slide row plus its disclosed covenant rule, status credit remains independent from separately scored numeric accuracy, and a scoped source-authority judge checks the slide's compact source note. A prompt-visible calculation-support block deterministically reconciles funded debt, posted and pro-forma EBITDA, cash taxes, capex financing, fixed charges, and both FCCR numerators so the displayed ratios can be reperformed from the slide; a prompt-visible 25%/50%/100% proposed-WIP reversal sensitivity independently checks retained adjustment, covenant values, headroom, equations, and scoped scenario status without treating the illustrative cases as posted or submitted", 'effective_date': '2026-08-29', 'id': 'task-015-objective-led-covenant-stress-v11'}, 'task_027': {'id': 'task-027-stable-scenario-model-v4', 'effective_date': '2026-08-26', 'basis': 'controller-visible FY27 source conventions plus formula-driven branch, EBITDA, free-cash-flow, weekly liquidity, lender leverage and fixed-charge coverage release decisions with proportional section-normalized finance and formula-lineage scoring, with professional row-label equivalence and deterministic repeatability'}, 'task_035': {'basis': 'a concise planning request and non-presolved model require the analyst to construct the probability-plan, gross-commitment, executable-priority, contractual-damages, deferred-gross-profit, required-recovery, and release-condition decisions with disclosed sign and status conventions, professionally equivalent labels, deterministic formula-backed model/control checks, independently scoped and non-duplicative monthly, project, portfolio, and recovery judgment rows, equivalent causal recovery actions and release conditions, an explicit non-duplicative current-state boundary and explicit conditional-until equivalence for the final recovery decision, ordinary capacity-clear status equivalents, stable portfolio-release boundaries, source checks aligned to the actual non-MCP source manifest, and proportional section-normalized spreadsheet scoring', 'effective_date': '2026-08-29', 'id': 'task-035-executable-backlog-capacity-decision-v11'}, 'task_037': {'id': 'task-037-professional-capital-model-v2', 'effective_date': '2026-08-26', 'basis': 'source-complete capital-allocation model with formula-linked project and portfolio economics, normal finance labels, deterministic numeric checks, materiality-balanced committee-decision scoring, and no hidden all-or-nothing model cap'}, 'task_055': {'id': 'task-055-professional-accretion-model-v2', 'effective_date': '2026-08-26', 'basis': 'source-grounded acquisition accretion model accepting normal sources-and-uses and income-statement labels, conventional expense signs, and case-and-year-aware sensitivity intersections, a disclosed release rule, and formula-backed finance outcomes without phrase matching'}, 'task_061': {'id': 'task-061-source-complete-tax-provision-v1', 'effective_date': '2026-08-26', 'basis': 'tax-provision model with every graded opening balance and estimated payment disclosed, conventional valuation-allowance presentation, and independent roll-forward checks that preserve the correct economics'}, 'task_068': {'basis': 'a concise executive request requires independent cross-source reconciliation and judgment; source-complete executive performance deck with deterministic posted-revenue, business-unit, cash, backlog, lender, and mitigation-chain checks; the accounting tie and executable mitigation decision carry eighty percent of the reward, unchanged seeded decks receive no deliverable credit, and each semantic review is limited to the slide or slides needed for that criterion', 'effective_date': '2026-08-29', 'id': 'task-068-executive-accounting-mitigation-decision-v10'}, 'task_072': {'id': 'task-072-management-narrative-equivalence-v1', 'effective_date': '2026-08-26', 'basis': 'quarterly finance narrative graded from local paragraph and table context, including common million-dollar presentation and ordinary plan, variance, guidance, and operating-driver labels'}, 'task_073': {'id': 'task-073-disclosed-pro-forma-debt-basis-v1', 'effective_date': '2026-08-26', 'basis': 'credit-metrics model with the controlling pro-forma acquisition-facility draw and document-authority hierarchy stated in the task, so posted and pro-forma debt bases cannot be confused'}, 'task_100': {'id': 'task-100-board-language-equivalence-v2', 'effective_date': '2026-08-26', 'basis': 'source-grounded CFO board deck with deterministic finance checks and independent semantic rescue for normal executive labels, decisions, and status wording while unsupported conclusions remain uncredited and adjacent board-card booleans cannot contaminate one another'}}

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
        specs.append(('commercial_sensitivity__structure', 'Commercial-recovery sensitivity structure'))
        for project_id in ('arm-2318', 'arm-2409', 'arm-2417', 'arm-2506'):
            specs.extend(((f'commercial_sensitivity__{project_id}__{field}', f'{project_id} commercial sensitivity {field}') for field in ('pending_revenue', 'hypothetical_contract', 'close_eac', 'hypothetical_margin', 'hypothetical_margin_rate', 'incremental_margin', 'formula_lineage')))
        specs.extend(((f'commercial_sensitivity__total__{field}', f'Commercial sensitivity total {field}') for field in ('pending_revenue', 'hypothetical_contract', 'close_eac', 'hypothetical_margin', 'hypothetical_margin_rate', 'incremental_margin', 'formula_lineage')))
        specs.append(('commercial_sensitivity__basis', 'Unbooked sensitivity-only basis'))
        specs.append(('accounting_control__structure', 'Accounting-control structure'))
        for scope in ('arm-2318', 'arm-2409', 'arm-2417', 'arm-2506', 'total'):
            for metric in ('job_cost', 'billings', 'current_contract'):
                specs.extend(((f'accounting_control__{scope}__{metric}__{field}', f'{scope} accounting control {metric} {field}') for field in ('base', 'activity', 'close', 'risk_review', 'variance')))
                specs.extend(((f'accounting_control__{scope}__{metric}__equation', f'{scope} accounting control {metric} roll-forward equation'), (f'accounting_control__{scope}__{metric}__formula_lineage', f'{scope} accounting control {metric} formula lineage')))
    else:
        if task_id == 'task_015':
            specs.append(('preservation__slide_count', 'Exactly one slide appended'))
            specs.extend(((f'preservation__slide_{index:02d}', f'Preserve slide {index}') for index in range(1, 6)))
            specs.extend((('structure__title', 'Required slide title'), ('structure__table', 'Required PowerPoint table')))
            specs.extend((('style__visual_system', 'Match the existing deck visual system'), ('style__title_hierarchy', 'Use a clear title hierarchy consistent with the deck'), ('style__table_readability', 'Present the covenant comparison in a reviewable table')))
            for metric in ('leverage', 'fccr', 'tangible_net_worth'):
                specs.extend(((f'metric__{metric}__{field}', f'{metric} {field}') for field in ('posted', 'pro_forma', 'threshold', 'posted_headroom', 'pro_forma_headroom', 'status')))
            specs.extend((('status__proposed', 'Proposed WIP status'), ('status__unposted', 'Unposted WIP status'), ('status__compliant', 'Compliance status'), ('headroom__binding_covenant', 'Binding covenant on like-for-like deterioration capacity'), ('source__authority', 'Source note identifies the controlling calculation and accounting authorities')))
            specs.extend(((f'calculation_support__{field}', f'Calculation support {description}') for (field, description) in (('funded_debt', 'funded debt'), ('posted_adjusted_ebitda', 'posted adjusted EBITDA'), ('proposed_wip_adjustment', 'proposed WIP adjustment'), ('pro_forma_adjusted_ebitda', 'pro-forma adjusted EBITDA'), ('cash_taxes', 'cash taxes'), ('ltm_fixed_asset_additions', 'LTM fixed-asset additions'), ('direct_equipment_financing', 'direct LTM equipment financing'), ('unfunded_capex', 'unfunded capex'), ('cash_interest', 'cash interest'), ('scheduled_principal', 'scheduled principal'), ('fixed_charges', 'fixed charges'), ('posted_fccr_numerator', 'posted FCCR numerator'), ('pro_forma_fccr_numerator', 'pro-forma FCCR numerator'))))
            specs.extend(((f'calculation_support__equation__{field}', f'Calculation support equation {description}') for (field, description) in (('ebitda_bridge', 'posted EBITDA plus proposed WIP equals pro-forma EBITDA'), ('unfunded_capex', 'fixed-asset additions less direct financing equals unfunded capex'), ('fixed_charges', 'cash interest plus scheduled principal equals fixed charges'), ('posted_fccr_numerator', 'posted EBITDA less taxes and unfunded capex equals the posted FCCR numerator'), ('pro_forma_fccr_numerator', 'pro-forma EBITDA less taxes and unfunded capex equals the pro-forma FCCR numerator'))))
            specs.append(('structure__wip_reversal_sensitivity', 'Required proposed-WIP reversal sensitivity table'))
            for reversal in (25, 50, 100):
                specs.extend(((f'sensitivity__reversal_{reversal}__{field}', f"{reversal}% WIP-reversal sensitivity {field.replace('_', ' ')}") for field in ('retained_wip_adjustment', 'adjusted_ebitda', 'tangible_net_worth', 'leverage', 'fccr_numerator', 'fccr', 'leverage_deterioration_capacity', 'fccr_deterioration_capacity', 'tangible_net_worth_headroom', 'status', 'equation_retained_wip', 'equation_leverage', 'equation_fccr')))
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
        return {criterion_id for criterion_id in ids if criterion_id.endswith('__flag') or criterion_id.endswith('__commercial_treatment') or (criterion_id.startswith('review__') and criterion_id.endswith('__source')) or criterion_id.endswith('__close_disposition') or criterion_id.endswith('__required_follow_up') or (criterion_id in {'decision_control__pending_revenue', 'decision_control__review_queue', 'decision_control__close_release', 'commercial_sensitivity__basis'})}
    if task_id == 'task_015':
        return {'structure__title', 'status__compliant', 'status__proposed', 'status__unposted', 'headroom__binding_covenant', 'source__authority', 'metric__leverage__status', 'metric__fccr__status', 'metric__tangible_net_worth__status', 'sensitivity__reversal_25__status', 'sensitivity__reversal_50__status', 'sensitivity__reversal_100__status'}
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
            row.update({'category': 'provenance', 'weight': 10 if task_id == 'task_015' else 3})
            row.pop('failure_cap', None)
        elif task_id == 'task_004':
            if criterion_id.startswith('preservation__'):
                row.update({'category': 'integrity', 'weight': 10})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('source__') or criterion_id.endswith('__source') or criterion_id == 'decision_control__policy_basis':
                row.update({'category': 'provenance', 'weight': 3})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('formula_lineage__') or criterion_id == 'total_formula_lineage' or criterion_id.endswith('__formula_lineage'):
                row.update({'category': 'auditability', 'weight': 5})
                row.pop('failure_cap', None)
            elif criterion_id == 'decision_bridge__structure':
                row.update({'category': 'structure', 'weight': 1, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id == 'commercial_sensitivity__structure':
                row.update({'category': 'structure', 'weight': 1, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('commercial_sensitivity__') and criterion_id.endswith('__formula_lineage'):
                row.update({'category': 'auditability', 'weight': 5, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('commercial_sensitivity__'):
                row.update({'category': 'decision' if criterion_id.endswith('__basis') else 'core_finance', 'weight': 10, 'semantic': criterion_id in semantic_ids})
                row.pop('failure_cap', None)
            elif criterion_id == 'accounting_control__structure':
                row.update({'category': 'structure', 'weight': 1, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('accounting_control__') and criterion_id.endswith('__formula_lineage'):
                row.update({'category': 'auditability', 'weight': 5, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.startswith('accounting_control__'):
                row.update({'category': 'core_finance', 'weight': 10, 'semantic': False})
                row.pop('failure_cap', None)
            elif criterion_id.endswith('__flag') or criterion_id.endswith('__commercial_treatment') or criterion_id.endswith('__close_disposition') or criterion_id.endswith('__required_follow_up') or (criterion_id in {'decision_control__pending_revenue', 'decision_control__review_count', 'decision_control__largest_overlay', 'decision_control__review_queue', 'decision_control__close_release'}):
                row.update({'category': 'decision', 'weight': 10, 'semantic': criterion_id in semantic_ids})
                row.pop('failure_cap', None)
            else:
                row.update({'category': 'core_finance', 'weight': 10, 'semantic': False})
                row.pop('failure_cap', None)
        else:
            if task_id == 'task_015':
                row.pop('failure_cap', None)
                if criterion_id.startswith('preservation__'):
                    row.update({'category': 'integrity', 'weight': 3, 'semantic': False})
                elif criterion_id.startswith('structure__'):
                    row.update({'category': 'structure', 'weight': 1, 'semantic': criterion_id in semantic_ids})
                elif criterion_id.startswith('style__'):
                    row.update({'category': 'presentation_quality', 'weight': 3, 'semantic': False})
                elif criterion_id.startswith('metric__'):
                    is_status = criterion_id.endswith('__status')
                    row.update({'category': 'decision' if is_status else 'core_finance', 'weight': 5 if is_status else 10, 'semantic': criterion_id in semantic_ids})
                elif criterion_id.startswith('status__'):
                    row.update({'category': 'decision', 'weight': 5, 'semantic': criterion_id in semantic_ids})
                elif criterion_id.startswith('headroom__'):
                    row.update({'category': 'decision', 'weight': 10, 'semantic': criterion_id in semantic_ids})
                elif criterion_id.startswith('calculation_support__equation__'):
                    row.update({'category': 'controls', 'weight': 5, 'semantic': False})
                elif criterion_id.startswith('calculation_support__'):
                    row.update({'category': 'core_finance', 'weight': 10, 'semantic': False})
                elif criterion_id == 'structure__wip_reversal_sensitivity':
                    row.update({'category': 'structure', 'weight': 1, 'semantic': False})
                elif criterion_id.startswith('sensitivity__'):
                    is_status = criterion_id.endswith('__status')
                    is_equation = '__equation_' in criterion_id
                    row.update({'category': 'decision' if is_status else 'controls' if is_equation else 'core_finance', 'weight': 5 if is_status or is_equation else 10, 'semantic': is_status})
                else:
                    raise ValueError(f'unclassified task_015 criterion: {criterion_id}')
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

def _attach_semantic_review(result: dict[str, Any], *, task_id: str, evidence: str, artifact_type: str, specs: list[dict[str, Any]], decision_failure_cap: float | None=0.49, always_judge: bool=False, execution_mode: str | None=None) -> dict[str, Any]:
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
        requirement = f"Decide whether the submitted {artifact_type} evidence satisfies: {criterion['description']}" if execution_mode == 'scoped_per_criterion' else semantic_requirement(criterion_id=criterion_id, description=str(criterion['description']), expected_facts=spec.get('expected_facts'), artifact_type=artifact_type)
        review = {'criterion_id': criterion_id, 'requirement': requirement, 'hard_gate_met': bool(spec['hard_gate_met']), 'hard_gate_evidence': str(spec['hard_gate_evidence']), 'legacy_lexical_match': bool(criterion['value']), 'always_judge': bool(spec.get('always_judge', always_judge))}
        if execution_mode == 'scoped_per_criterion':
            review.update({'submitted_evidence': str(spec.get('submitted_evidence') or '')[:12000], 'reference_context': spec.get('reference_context') or spec.get('expected_facts') or {}, 'task_context': spec.get('task_context') or {}, 'evidence_scope': str(spec.get('evidence_scope') or 'criterion_field')})
        reviews.append(review)
    result['semantic_review'] = {'version': 2, 'mode': 'deterministic_hard_gates_plus_bounded_semantic_judge', 'task_id': task_id, 'artifact': None, 'evidence': evidence[:60000], 'criteria': reviews, 'execution_mode': execution_mode or 'legacy_batched', 'policy': 'Semantic wording cannot rescue a failed objective numeric, formula, or structure gate.'}
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
        if any((token in text for token in ('no review', 'review not required', 'review is not required', 'controller review not required', 'no controller review'))):
            return False
        met = text in {'yes', 'required'} or any((token in text for token in ('review required', 'controller review', 'flag', 'escalat', 'hold', 'controller signoff', 'controller sign off')))
    else:
        if any((token in text for token in ('review required', 'controller review required', 'hold for review', 'hold pending review', 'escalate', 'escalation required'))):
            return False
        met = text in {'no', 'monitor'} or any((token in text for token in ('below threshold', 'within threshold', 'no review', 'not required', 'monitor', 'release', 'clear')))
    if expected['netted_claim_or_backcharge']:
        met = met and any((token in text for token in ('backcharge', 'recovery', 'netted')))
    return met

def _task_004_treatment_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = str(value or '')
    normalized = _normalize(text)
    if expected['netted_claim_or_backcharge']:
        wrong_direction = any((token in normalized for token in ('do not add back', 'not add back', 'leave netted', 'remain netted', 'keep netted', 'retain the recovery offset', 'retain recovery offset')))
        treatment_ok = not wrong_direction and (any((token in normalized for token in ('backcharge', 'recovery'))) and any((token in normalized for token in ('do not net', 'not netted', 'de net', 'denet', 'add back', 'added back', 'restore', 'restored', 'remove from pm etc', 'removed from pm etc', 'cost added'))))
        return treatment_ok and _text_contains_number(text, expected['cost_amount'], abs_tol=1.0)
    wrong_revenue_direction = any((token in normalized for token in ('do not exclude', 'not excluded', 'include revenue', 'revenue included', 'recognize revenue', 'recognise revenue', 'revenue recognized', 'revenue recognised')))
    revenue_ok = not wrong_revenue_direction and (any((token in normalized for token in ('exclude', 'excluded', 'omit', 'omitted', 'no revenue', 'not recognize', 'not recognised', 'unapproved', 'pending revenue'))) and any((token in normalized for token in ('revenue', 'recovery', 'change order', 'commercial value', 'pco'))) and _text_contains_number(text, expected['revenue_amount'], abs_tol=1.0))
    cost_ok = any((token in normalized for token in ('include', 'included', 'add', 'added', 'carry', 'carried', 'record', 'recorded', 'overlay', 'probable cost'))) and any((token in normalized for token in ('cost', 'exposure', 'etc', 'overlay')))
    return revenue_ok and cost_ok and _text_contains_number(text, expected['cost_amount'], abs_tol=1.0)

def _task_004_source_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    reference = _normalize(expected['reference'])
    return reference in text and ('fin rev 04' in text or 'signed wip policy' in text or 'wip policy' in text)

def _task_004_has_phrase(text: str, phrase: str) -> bool:
    """Match a normalized word/phrase without substring accidents.

    In particular, ``hold`` must not match the final five letters of
    ``threshold`` in an otherwise correct release conclusion.
    """
    return bool(re.search(f'(?<![a-z0-9]){re.escape(_normalize(phrase))}(?![a-z0-9])', _normalize(text)))

def _task_004_close_disposition_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    if expected['controller_review_required']:
        if any((_task_004_has_phrase(text, token) for token in ('do not hold', 'not held', 'release now', 'release immediately', 'cleared for release', 'clear for release'))):
            return False
        return any((_task_004_has_phrase(text, token) for token in ('hold', 'do not release', 'not release'))) and any((_task_004_has_phrase(text, token) for token in ('review', 'controller', 'signoff', 'sign off', 'approval', 'pending')))
    if any((_task_004_has_phrase(text, token) for token in ('do not release', 'not release', 'release not approved', 'hold'))):
        return False
    return any((_task_004_has_phrase(text, token) for token in ('release', 'clear'))) and any((_task_004_has_phrase(text, token) for token in ('monitor', 'follow up', 'follow-up', 'closeout')))

def _task_004_follow_up_ok(value: Any, expected: dict[str, Any]) -> bool:
    text = _normalize(value)
    return all((any((_normalize(token) in text for token in token_group)) for token_group in expected['required_follow_up_tokens']))

def _task_004_pending_revenue_excluded(value: Any) -> bool:
    if value is False or value == 0:
        return True
    text = _normalize(value)
    if any((token in text for token in ('not excluded', 'do not exclude', 'included', 'include revenue', 'recognize', 'recognise'))):
        return False
    return text in {'no', 'none', 'zero'} or any((token in text for token in ('excluded', 'not included', 'no pending revenue')))

def _task_004_review_queue_ok(value: Any, expected: list[str]) -> bool:
    text = _normalize(value)
    expected_present = all((_normalize(project_id) in text for project_id in expected))
    if not expected_present or 'arm 2318' not in text:
        return expected_present
    return any((token in text for token in ('arm 2318 released', 'release arm 2318', 'arm 2318 cleared', 'clear arm 2318', 'except arm 2318', 'excluding arm 2318')))

def _task_004_close_release_ok(value: Any) -> bool:
    text = _normalize(value)
    if any((_task_004_has_phrase(text, token) for token in ('do not hold', 'not held', 'release now', 'release immediately', 'cleared for release', 'clear for release'))):
        return False
    return _task_004_has_phrase(text, 'hold') and any((_task_004_has_phrase(text, token) for token in ('review', 'controller', 'signoff', 'sign off', 'approval', 'pending')))

def _task_004_review_count(value: Any) -> float | None:
    parsed = _number(value)
    if parsed is not None:
        return parsed
    match = re.search('(?<!\\d)(\\d+)(?!\\d)', str(value or ''))
    return float(match.group(1)) if match else None

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

def _task_004_commercial_sensitivity_checks(formula_sheet: Any, value_sheet: Any, gold: dict[str, Any]) -> tuple[list[Criterion], dict[str, Any]]:
    """Grade the source-backed, explicitly unbooked recovery sensitivity.

    The table may begin on any row below the existing decision controls and may
    use ordinary professional header synonyms. Objective values and formula
    lineage are deterministic; only the accounting-basis statement is semantic.
    """
    expected = gold['commercial_recovery_sensitivity']
    header_aliases = {'project_id': {'project', 'project id', 'job', 'job id'}, 'pending_revenue': {'pending commercial revenue', 'pending revenue', 'commercial revenue', 'pending recovery', 'commercial recovery'}, 'hypothetical_contract': {'hypothetical contract', 'hypothetical contract value', 'authorized contract', 'sensitivity contract', 'pro forma contract'}, 'close_eac': {'close eac', 'close eac booked', 'booked close eac', 'june close eac'}, 'hypothetical_margin': {'hypothetical margin', 'hypothetical total margin', 'sensitivity margin', 'pro forma margin'}, 'hypothetical_margin_rate': {'hypothetical margin rate', 'hypothetical margin %', 'sensitivity margin rate', 'pro forma margin rate', 'pro forma margin %'}, 'incremental_margin': {'incremental margin', 'incremental margin vs booked close', 'margin uplift', 'incremental margin versus booked close'}}
    header_row: int | None = None
    columns: dict[str, int] = {}
    max_scan_row = min(max(formula_sheet.max_row, 26), 80)
    for row in range(26, max_scan_row + 1):
        normalized = {column: _normalize(formula_sheet.cell(row, column).value) for column in range(1, min(max(formula_sheet.max_column, 7), 20) + 1)}
        matched: dict[str, int] = {}
        for (field, aliases) in header_aliases.items():
            column = next((column for (column, text) in normalized.items() if text in aliases), None)
            if column is not None:
                matched[field] = column
        if len(matched) == len(header_aliases):
            header_row = row
            columns = matched
            break
    project_rows: dict[str, int] = {}
    total_row: int | None = None
    section_text = ''
    if header_row is not None:
        project_column = columns['project_id']
        end_row = min(max(formula_sheet.max_row, header_row + 8), header_row + 15)
        text_parts: list[str] = []
        for row in range(max(26, header_row - 2), end_row + 1):
            values = [formula_sheet.cell(row, column).value for column in range(1, 12)]
            text_parts.extend((str(value) for value in values if value not in (None, '')))
            label = _normalize(value_sheet.cell(row, project_column).value)
            if re.fullmatch('arm[ -]?\\d{4}', label):
                project_rows[label.replace(' ', '-')] = row
            elif label in {'total', 'portfolio total', 'aggregate', 'total / control', 'sensitivity total', 'total sensitivity'}:
                total_row = row
        section_text = '\n'.join(text_parts)
    structure_met = header_row is not None and all((_normalize(row['project_id']).replace(' ', '-') in project_rows for row in expected['projects'])) and (total_row is not None)
    criteria = [Criterion('commercial_sensitivity__structure', 'The workbook contains a complete four-project commercial-recovery sensitivity and total row', structure_met, f'header_row={header_row}; project_rows={project_rows}; total_row={total_row}')]
    source_rows = {row['project_id'].casefold(): source_row for (source_row, row) in zip(range(6, 10), expected['projects'], strict=True)}
    numeric_fields = ('pending_revenue', 'hypothetical_contract', 'close_eac', 'hypothetical_margin', 'hypothetical_margin_rate', 'incremental_margin')

    def add_numeric_rows(scope: str, row_number: int | None, targets: dict[str, Any]) -> None:
        for field in numeric_fields:
            actual = value_sheet.cell(row_number, columns[field]).value if row_number is not None and field in columns else None
            tolerance = 5e-05 if field == 'hypothetical_margin_rate' else 0.05
            criteria.append(Criterion(f'commercial_sensitivity__{scope}__{field}', f'Commercial-recovery sensitivity `{scope}` {field} is correct', _close(actual, targets[field], abs_tol=tolerance, rel_tol=0.0), f'actual={actual!r}; expected={targets[field]!r}'))
        lineage_met = False
        lineage_evidence = 'required sensitivity row is missing'
        if row_number is not None and all((field in columns for field in numeric_fields)):
            c_pending = formula_sheet.cell(row_number, columns['pending_revenue'])
            c_contract = formula_sheet.cell(row_number, columns['hypothetical_contract'])
            c_eac = formula_sheet.cell(row_number, columns['close_eac'])
            c_margin = formula_sheet.cell(row_number, columns['hypothetical_margin'])
            c_rate = formula_sheet.cell(row_number, columns['hypothetical_margin_rate'])
            c_incremental = formula_sheet.cell(row_number, columns['incremental_margin'])
            if scope == 'total':
                first_project_row = min(project_rows.values()) if project_rows else row_number - 4
                last_project_row = max(project_rows.values()) if project_rows else row_number - 1
                margin_range = f'{c_margin.column_letter}{first_project_row}:{c_margin.column_letter}{last_project_row}'
                lineage_met = all((_formula_has_refs(c_pending, f'{c_pending.column_letter}{first_project_row}:{c_pending.column_letter}{last_project_row}'), _formula_has_refs(c_contract, f'{c_contract.column_letter}{first_project_row}:{c_contract.column_letter}{last_project_row}'), _formula_has_refs(c_eac, f'{c_eac.column_letter}{first_project_row}:{c_eac.column_letter}{last_project_row}'), _formula_has_refs(c_margin, c_contract.coordinate, c_eac.coordinate) or _formula_has_refs(c_margin, margin_range), _formula_has_refs(c_rate, c_margin.coordinate, c_contract.coordinate), _formula_has_refs(c_incremental, f'{c_incremental.column_letter}{first_project_row}:{c_incremental.column_letter}{last_project_row}')))
            else:
                source_row = source_rows[scope]
                lineage_met = all((_formula_has_refs(c_contract, c_pending.coordinate, f'C{source_row}'), _formula_has_refs(c_eac, f'H{source_row}'), _formula_has_refs(c_margin, c_contract.coordinate, c_eac.coordinate), _formula_has_refs(c_rate, c_margin.coordinate, c_contract.coordinate), _formula_has_refs(c_incremental, c_pending.coordinate) or _formula_has_refs(c_incremental, c_margin.coordinate, f'N{source_row}')))
            lineage_met = lineage_met and all((not _formula_has_large_literal(cell) for cell in (c_contract, c_eac, c_margin, c_rate, c_incremental)))
            lineage_evidence = '; '.join((f'{cell.coordinate}={cell.value!r}' for cell in (c_pending, c_contract, c_eac, c_margin, c_rate, c_incremental)))
        criteria.append(Criterion(f'commercial_sensitivity__{scope}__formula_lineage', f'Commercial-recovery sensitivity `{scope}` responds to its disclosed inputs', lineage_met, lineage_evidence))
    for project in expected['projects']:
        scope = project['project_id'].casefold()
        add_numeric_rows(scope, project_rows.get(scope), project)
    add_numeric_rows('total', total_row, expected['total'])
    basis_present = bool(_normalize(section_text))
    criteria.append(Criterion('commercial_sensitivity__basis', 'The section is clearly labeled as an unbooked sensitivity rather than June-close revenue', basis_present, section_text[:1600] if section_text else 'commercial sensitivity section is missing'))
    return (criteria, {'header_row': header_row, 'project_rows': project_rows, 'total_row': total_row, 'section_text': section_text, 'basis_present': basis_present})

def _task_004_accounting_control_checks(formula_book: Any, value_book: Any, gold: dict[str, Any]) -> tuple[list[Criterion], dict[str, Any]]:
    """Grade the four-project accounting-to-workpaper reconciliation.

    The table can live on an ordinarily named control/reconciliation sheet and
    its columns may use normal controller shorthand. Values and equations are
    checked individually so a missing label or one bad cell cannot erase an
    otherwise correct schedule. No LLM is used for objective accounting facts.
    """
    aliases = {'project_id': {'project', 'project id', 'job', 'job id'}, 'job_cost__base': {'may 31 job cost', 'may job cost', 'job cost base', 'may 31 posted job cost', 'beginning job cost'}, 'job_cost__activity': {'june job cost activity', 'june cost activity', 'job cost activity', 'june posted cost', 'june posted job cost'}, 'job_cost__close': {'june 30 job cost', 'june close job cost', 'job cost close', 'source job cost', 'accounting job cost'}, 'job_cost__risk_review': {'risk review posted cost', 'risk review job cost', 'wip posted cost', 'workpaper job cost', 'risk review cost'}, 'job_cost__variance': {'job cost variance', 'cost variance', 'job cost difference'}, 'billings__base': {'may 31 billings', 'may billings', 'billings base', 'beginning billings'}, 'billings__activity': {'june billings', 'june billing activity', 'billings activity', 'june invoices'}, 'billings__close': {'june 30 billings', 'june close billings', 'billings close', 'source billings', 'accounting billings'}, 'billings__risk_review': {'risk review billings', 'wip billings', 'workpaper billings'}, 'billings__variance': {'billings variance', 'billing variance', 'billings difference'}, 'current_contract__base': {'original contract', 'contract base', 'base contract'}, 'current_contract__activity': {'approved change orders', 'approved changes', 'contract changes', 'approved co', 'approved cos'}, 'current_contract__close': {'june 30 current contract', 'current contract', 'contract close', 'source current contract'}, 'current_contract__risk_review': {'risk review current contract', 'risk review contract', 'wip current contract', 'workpaper current contract'}, 'current_contract__variance': {'current contract variance', 'contract variance', 'contract difference'}}
    best: tuple[int, Any, int, dict[str, int]] | None = None
    for sheet in formula_book.worksheets:
        for row in range(1, min(max(sheet.max_row, 1), 40) + 1):
            normalized = {column: _normalize(sheet.cell(row, column).value) for column in range(1, min(max(sheet.max_column, 1), 40) + 1)}
            matched: dict[str, int] = {}
            for (field, field_aliases) in aliases.items():
                column = next((column for (column, text) in normalized.items() if text in field_aliases), None)
                if column is not None:
                    matched[field] = column
            score = len(matched)
            if 'project_id' in matched and (best is None or score > best[0]):
                best = (score, sheet, row, matched)
    if best is None:
        sheet = None
        value_sheet = None
        header_row = None
        columns: dict[str, int] = {}
    else:
        (_, sheet, header_row, columns) = best
        value_sheet = value_book[sheet.title]
    project_rows: dict[str, int] = {}
    total_row: int | None = None
    if sheet is not None and header_row is not None and ('project_id' in columns):
        project_column = columns['project_id']
        for row in range(header_row + 1, min(sheet.max_row, header_row + 25) + 1):
            displayed = value_sheet.cell(row, project_column).value
            label = _normalize(displayed)
            match = re.search('arm[ -]?(\\d{4})', label)
            if match:
                project_rows[f'arm-{match.group(1)}'] = row
            elif label in {'total', 'portfolio total', 'aggregate', 'total control', 'total / control', 'four project total', 'four project control'}:
                total_row = row
    required_fields = set(aliases)
    structure_met = sheet is not None and required_fields.issubset(columns) and all((_normalize(row['project_id']).replace(' ', '-') in project_rows for row in gold['accounting_control']['rows'])) and (total_row is not None)
    criteria: list[Criterion] = [Criterion('accounting_control__structure', 'A complete four-project accounting-to-workpaper control and total row are present', structure_met, f"sheet={getattr(sheet, 'title', None)!r}; header_row={header_row}; matched_fields={sorted(columns)!r}; project_rows={project_rows!r}; total_row={total_row}")]
    source_rows = {row['project_id'].casefold(): source_row for (source_row, row) in zip(range(6, 10), gold['source_inputs'], strict=True)}
    risk_targets = {row['project_id'].casefold(): {'job_cost': row['cost_to_date'], 'billings': row['billings'], 'current_contract': row['current_contract']} for row in gold['source_inputs']}

    def actual_cell(row_number: int | None, field: str) -> Any | None:
        if row_number is None or value_sheet is None or field not in columns:
            return None
        return value_sheet.cell(row_number, columns[field])

    def formula_cell(row_number: int | None, field: str) -> Any | None:
        if row_number is None or sheet is None or field not in columns:
            return None
        return sheet.cell(row_number, columns[field])

    def add_scope(scope: str, row_number: int | None, targets: dict[str, dict[str, float]], *, source_row: int | None, project_row_numbers: list[int]) -> None:
        for metric in ('job_cost', 'billings', 'current_contract'):
            metric_targets = targets[metric]
            expected_values = {'base': metric_targets['base'], 'activity': metric_targets['activity'], 'close': metric_targets['close'], 'risk_review': metric_targets['risk_review'], 'variance': 0.0}
            actual_values: dict[str, float | None] = {}
            for (field, expected) in expected_values.items():
                cell = actual_cell(row_number, f'{metric}__{field}')
                actual = cell.value if cell is not None else None
                actual_values[field] = _number(actual)
                criteria.append(Criterion(f'accounting_control__{scope}__{metric}__{field}', f'{scope} {metric} accounting-control {field} is correct', _close(actual, expected, abs_tol=0.05, rel_tol=0.0), f'actual={actual!r}; expected={expected}'))
            equation_met = actual_values['base'] is not None and actual_values['activity'] is not None and (actual_values['close'] is not None) and (actual_values['risk_review'] is not None) and (actual_values['variance'] is not None) and _close(actual_values['close'], actual_values['base'] + actual_values['activity'], abs_tol=0.05, rel_tol=0.0) and _close(actual_values['variance'], actual_values['close'] - actual_values['risk_review'], abs_tol=0.05, rel_tol=0.0)
            criteria.append(Criterion(f'accounting_control__{scope}__{metric}__equation', f'{scope} {metric} closes from base plus activity and ties to Risk Review', equation_met, f'submitted={actual_values!r}'))
            base_cell = formula_cell(row_number, f'{metric}__base')
            activity_cell = formula_cell(row_number, f'{metric}__activity')
            close_cell = formula_cell(row_number, f'{metric}__close')
            risk_cell = formula_cell(row_number, f'{metric}__risk_review')
            variance_cell = formula_cell(row_number, f'{metric}__variance')
            lineage_met = False
            if all((cell is not None for cell in (base_cell, activity_cell, close_cell, risk_cell, variance_cell))):
                if scope == 'total':
                    first_row = min(project_row_numbers) if project_row_numbers else row_number or 1
                    last_row = max(project_row_numbers) if project_row_numbers else row_number or 1
                    source_total_column = {'job_cost': 'D', 'billings': 'K', 'current_contract': 'C'}[metric]
                    base_activity_close_met = all((_formula_has_refs(cell, f'{cell.column_letter}{first_row}:{cell.column_letter}{last_row}') for cell in (base_cell, activity_cell, close_cell)))
                    risk_met = _formula_has_refs(risk_cell, f'{risk_cell.column_letter}{first_row}:{risk_cell.column_letter}{last_row}') or _formula_has_refs(risk_cell, f'Risk Review!{source_total_column}11')
                    variance_met = _formula_has_refs(variance_cell, f'{variance_cell.column_letter}{first_row}:{variance_cell.column_letter}{last_row}') or _formula_has_refs(variance_cell, close_cell.coordinate, risk_cell.coordinate)
                    lineage_met = base_activity_close_met and risk_met and variance_met
                else:
                    risk_source_column = {'job_cost': 'D', 'billings': 'K', 'current_contract': 'C'}[metric]
                    lineage_met = all((_formula_has_refs(close_cell, base_cell.coordinate, activity_cell.coordinate), source_row is not None and _formula_has_refs(risk_cell, f'Risk Review!{risk_source_column}{source_row}'), _formula_has_refs(variance_cell, close_cell.coordinate, risk_cell.coordinate)))
                lineage_met = lineage_met and all((not _formula_has_large_literal(cell) for cell in (close_cell, risk_cell, variance_cell)))
            criteria.append(Criterion(f'accounting_control__{scope}__{metric}__formula_lineage', f'{scope} {metric} roll-forward, Risk Review link, variance, and total are formula-driven', lineage_met, '; '.join((f"{getattr(cell, 'coordinate', '?')}={getattr(cell, 'value', None)!r}" for cell in (base_cell, activity_cell, close_cell, risk_cell, variance_cell)))))
    project_row_numbers = sorted(project_rows.values())
    for expected in gold['accounting_control']['rows']:
        scope = expected['project_id'].casefold()
        targets = {metric: {**expected[metric], 'risk_review': risk_targets[scope][metric]} for metric in ('job_cost', 'billings', 'current_contract')}
        add_scope(scope, project_rows.get(scope), targets, source_row=source_rows[scope], project_row_numbers=project_row_numbers)
    total_targets = {metric: {**gold['accounting_control']['total'][metric], 'risk_review': round(sum((row[metric] for row in risk_targets.values())), 2)} for metric in ('job_cost', 'billings', 'current_contract')}
    add_scope('total', total_row, total_targets, source_row=None, project_row_numbers=project_row_numbers)
    return (criteria, {'sheet': getattr(sheet, 'title', None), 'header_row': header_row, 'columns': columns, 'project_rows': project_rows, 'total_row': total_row})

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
    for (sheet, ranges) in (('READ ME first', ['A5:B17']), ('Risk Review', ['A5:T5', 'A6:B9', 'A11:A11']), ('Evidence Map', ['A5:G9'])):
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
        perturbation_met = None if perturbation is None else perturbation[0].get(row_number, False)
        lineage_met = all(cell_checks) if perturbation_met is None else bool(perturbation_met)
        criteria.append(Criterion(f"formula_lineage__{expected['project_id'].casefold()}", f"{expected['project_id']} calculations respond correctly to their source inputs", lineage_met, '; '.join(evidence) + f'; perturbation_check={perturbation_met}'))
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
    perturbation_total_met = None if perturbation is None else perturbation[1]
    total_lineage_met = all(total_lineage_checks) if perturbation_total_met is None else bool(perturbation_total_met)
    criteria.append(Criterion('total_formula_lineage', 'Total/control row responds correctly to the four project rows', total_lineage_met, f"formulas={[formula_ws[f'{column}11'].value for column in 'CDEFGHIJKLMNOPQ']!r}; perturbation_check={perturbation_total_met}"))
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
        perturbation_met = None if perturbation is None else perturbation[2].get(decision_row, False)
        lineage_met = all(lineage_checks) and no_literals if perturbation_met is None else bool(perturbation_met) and no_literals
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__formula_lineage", f"{expected['project_id']} PM-to-close decision bridge responds to source changes", lineage_met, f"formulas={[formula_ws[f'{column}{decision_row}'].value for column in 'BCDEF']!r}; perturbation_check={perturbation_met}"))
        for (label, actual, target) in (('pm_case_margin', value_sheet[f'C{decision_row}'].value, expected['pm_case_margin']), ('margin_impact', value_sheet[f'F{decision_row}'].value, expected['margin_impact'])):
            criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__{label}", f"{expected['project_id']} decision bridge `{label}` is correct", _close(actual, target, abs_tol=0.05, rel_tol=0.0), f'actual={actual!r}; expected={target}'))
        disposition = value_sheet[f'I{decision_row}'].value
        follow_up = value_sheet[f'J{decision_row}'].value
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__close_disposition", f"{expected['project_id']} close disposition distinguishes release from controller hold", _task_004_close_disposition_ok(f"{disposition or ''} {follow_up or ''}", expected), f"disposition={disposition!r}; follow_up={follow_up!r}; review_required={expected['controller_review_required']}"))
        criteria.append(Criterion(f"decision_bridge__{expected['project_id'].casefold()}__required_follow_up", f"{expected['project_id']} required follow-up addresses the unresolved evidence", _task_004_follow_up_ok(follow_up, expected), f"actual={follow_up!r}; required_groups={expected['required_follow_up_tokens']!r}"))
    decision_control_lineage = all((_formula_has_refs(formula_ws['B21'], 'C15:C18'), _formula_has_refs(formula_ws['D21'], 'E15:E18'), _formula_has_refs(formula_ws['F21'], 'D21', 'B21'), _formula_has_refs(formula_ws['D22'], 'H15:H18'))) and all((not _formula_has_large_literal(formula_ws[coordinate]) for coordinate in ('B21', 'D21', 'F21', 'D22')))
    perturbation_decision_control_met = None if perturbation is None else perturbation[3]
    decision_control_lineage_met = decision_control_lineage if perturbation_decision_control_met is None else bool(perturbation_decision_control_met)
    criteria.append(Criterion('decision_control__formula_lineage', 'The controller decision control responds to the project bridge', decision_control_lineage_met, f"formulas={[formula_ws[coordinate].value for coordinate in ('B21', 'D21', 'F21', 'D22')]!r}; perturbation_check={perturbation_decision_control_met}"))
    control = gold['decision_control']
    for (label, coordinate, target) in (('pm_case_margin', 'B21', control['pm_case_margin']), ('margin_impact', 'F21', control['margin_impact'])):
        actual = value_sheet[coordinate].value
        criteria.append(Criterion(f'decision_control__{label}', f'Decision control `{label}` is correct', _close(actual, target, abs_tol=0.05, rel_tol=0.0), f'actual={actual!r}; expected={target}'))
    pending_revenue = value_sheet['B22'].value
    criteria.append(Criterion('decision_control__pending_revenue', 'The control confirms that pending commercial revenue is excluded from the close case', _task_004_pending_revenue_excluded(pending_revenue), f'actual={pending_revenue!r}'))
    review_count = _task_004_review_count(value_sheet['D22'].value)
    criteria.append(Criterion('decision_control__review_count', 'The control identifies the number of projects requiring controller review', review_count is not None and _close(review_count, control['review_required_count'], abs_tol=0.0, rel_tol=0.0), f"actual={review_count!r}; expected={control['review_required_count']}"))
    largest_overlay = _normalize(value_sheet['F22'].value)
    criteria.append(Criterion('decision_control__largest_overlay', 'The control identifies the project with the largest finance overlay', _normalize(control['largest_overlay_project']) in largest_overlay, f"actual={value_sheet['F22'].value!r}; expected={control['largest_overlay_project']}"))
    review_queue = value_sheet['B23'].value
    criteria.append(Criterion('decision_control__review_queue', 'The control identifies the exact controller review queue', _task_004_review_queue_ok(review_queue, control['review_queue']), f"actual={review_queue!r}; expected={control['review_queue']!r}"))
    close_release = _normalize(value_sheet['B24'].value)
    criteria.append(Criterion('decision_control__close_release', 'The aggregate close-release decision holds the flagged jobs for controller review', _task_004_close_release_ok(close_release), f"actual={value_sheet['B24'].value!r}"))
    policy_basis = _normalize(value_sheet['B25'].value)
    criteria.append(Criterion('decision_control__policy_basis', 'The decision control cites the signed WIP policy', 'fin rev 04' in policy_basis or 'signed wip policy' in policy_basis, f"actual={value_sheet['B25'].value!r}"))
    (commercial_sensitivity_criteria, commercial_sensitivity_context) = _task_004_commercial_sensitivity_checks(formula_ws, value_sheet, gold)
    criteria.extend(commercial_sensitivity_criteria)
    (accounting_control_criteria, _) = _task_004_accounting_control_checks(wb, values, gold)
    criteria.extend(accounting_control_criteria)
    result = _result(criteria)
    result['decision_support'] = _task_004_decision_support(value_sheet, gold)
    semantic_specs: list[dict[str, Any]] = []
    task_context = {'assignment': 'Complete the four-project June WIP risk review, controller decision bridge, commercial sensitivity, and accounting-to-workpaper control.', 'grading_boundary': 'Objective amounts, formulas, project identities, totals, and control counts are graded separately. Each semantic judge grades only the exact submitted review field supplied to it and must accept equivalent professional wording.'}
    for (source_row, expected) in zip(range(6, 10), gold['review_rows'], strict=True):
        project_id = expected['project_id']
        project_ok = _normalize(value_sheet[f'A{source_row}'].value) == _normalize(project_id)
        flag = value_sheet[f'R{source_row}'].value
        treatment = str(value_sheet[f'S{source_row}'].value or '')
        source = str(value_sheet[f'T{source_row}'].value or '')
        combined_source_evidence = f'{treatment}\n{source}'.strip()
        source_cells_present = bool(_normalize(combined_source_evidence))
        slug = project_id.casefold()
        semantic_specs.extend([{'criterion_id': f'review__{slug}__flag', 'expected_facts': {'project_id': project_id, 'controller_review_required': expected['controller_review_required'], 'netted_claim_or_backcharge': expected['netted_claim_or_backcharge']}, 'hard_gate_met': project_ok and bool(_normalize(flag)), 'hard_gate_evidence': f'project_associated={project_ok}; review conclusion nonblank={bool(_normalize(flag))}', 'submitted_evidence': f"Project: {project_id}\nSubmitted Controller review: {flag or ''}", 'reference_context': {'controller_review_required': expected['controller_review_required'], 'decision_meaning': 'Hold/escalate for Controller review' if expected['controller_review_required'] else 'No mandatory hold; release with monitoring is appropriate', 'equivalence_rule': 'This is the Controller-review status field. For a required review, Review, Controller review, Required, or an equivalent status with the stated >100 bps or commercial-risk basis is sufficient; the cell need not also say hold or escalate. For a non-required review, Release, Monitor, Within threshold, Supported, or an equivalent no-hold conclusion is sufficient. Reject an unresolved Review status for a non-required project.'}, 'task_context': task_context, 'evidence_scope': f'Risk Review!A{source_row},R{source_row}'}, {'criterion_id': f'review__{slug}__commercial_treatment', 'expected_facts': {'project_id': project_id, 'pending_commercial_amount_excluded': expected['revenue_amount'], 'probable_cost_or_recovery_overlay': expected['cost_amount'], 'disputed_recovery_de_netted': expected['netted_claim_or_backcharge']}, 'hard_gate_met': project_ok and bool(_normalize(treatment)), 'hard_gate_evidence': f'project_associated={project_ok}; treatment_nonblank={bool(_normalize(treatment))}. The finance amounts and calculation chain are independently graded.', 'submitted_evidence': f'Project: {project_id}\nSubmitted commercial treatment / rationale: {treatment}', 'reference_context': {'pending_commercial_amount': expected['revenue_amount'], 'required_treatment': 'exclude pending commercial revenue from the June close case', 'forecast_cost_or_recovery_overlay': expected['cost_amount'], 'overlay_treatment': 'restore/de-net the disputed recovery in ETC' if expected['netted_claim_or_backcharge'] else 'retain/carry the supported probable cost in ETC', 'equivalence_rule': 'The rationale need not repeat amounts already shown in the same deterministically graded project row. Accept accurate professional equivalents; reject the wrong recognition direction or a generic statement with no treatment.'}, 'task_context': task_context, 'evidence_scope': f'Risk Review!S{source_row}'}, {'criterion_id': f'review__{slug}__source', 'expected_facts': {'project_id': project_id, 'commercial_reference': expected['reference'], 'policy_authority': 'signed WIP policy FIN-REV-04'}, 'hard_gate_met': project_ok and source_cells_present, 'hard_gate_evidence': f'project_associated={project_ok}; treatment_and_source_cells_present={source_cells_present}', 'submitted_evidence': f'Project: {project_id}\nSubmitted treatment/source support:\n{combined_source_evidence}', 'reference_context': {'commercial_reference': expected['reference'], 'policy_authority': 'signed WIP policy FIN-REV-04', 'equivalence_rule': 'Accept clear abbreviations and ordinary filename/reference variants. Both the project-specific commercial support and governing policy must be identifiable. A same-row reference to the WIP policy or its relevant section is an acceptable identifier; do not require the literal code FIN-REV-04 when the authority is clear.'}, 'task_context': task_context, 'evidence_scope': f'Risk Review!S{source_row}:T{source_row}'}])
    for (decision_row, expected) in zip(range(15, 19), gold['decision_rows'], strict=True):
        project_id = expected['project_id']
        project_ok = _normalize(value_sheet[f'A{decision_row}'].value) == _normalize(project_id)
        evidence_cells_present = bool(_normalize(value_sheet[f'K{decision_row}'].value) and _normalize(value_sheet[f'L{decision_row}'].value))
        disposition = value_sheet[f'I{decision_row}'].value
        follow_up = value_sheet[f'J{decision_row}'].value
        slug = project_id.casefold()
        semantic_specs.extend([{'criterion_id': f'decision_bridge__{slug}__close_disposition', 'expected_facts': {'project_id': project_id, 'close_disposition': expected['close_disposition'], 'controller_review_required': expected['controller_review_required']}, 'hard_gate_met': project_ok and evidence_cells_present and bool(_normalize(disposition)), 'hard_gate_evidence': f'project_associated={project_ok}; policy_and_commercial_cells_present={evidence_cells_present}; disposition_nonblank={bool(_normalize(disposition))}', 'submitted_evidence': f"Project: {project_id}\nSubmitted close disposition: {disposition or ''}\nPolicy cell: {value_sheet[f'K{decision_row}'].value or ''}\nCommercial support cell: {value_sheet[f'L{decision_row}'].value or ''}", 'reference_context': {'required_disposition': expected['close_disposition'], 'controller_review_required': expected['controller_review_required'], 'policy_reference': expected['policy_reference'], 'commercial_reference': expected['commercial_reference'], 'equivalence_rule': 'Accept any unambiguous professional wording with the same release/hold direction.'}, 'task_context': task_context, 'evidence_scope': f'Risk Review!I{decision_row},K{decision_row}:L{decision_row}'}, {'criterion_id': f'decision_bridge__{slug}__required_follow_up', 'expected_facts': {'project_id': project_id, 'required_follow_up': expected['required_follow_up']}, 'hard_gate_met': project_ok and bool(_normalize(follow_up)), 'hard_gate_evidence': f'project_associated={project_ok}; follow_up_nonblank={bool(_normalize(follow_up))}. Policy and commercial support are evaluated in their own scoped criteria.', 'submitted_evidence': f"Project: {project_id}\nSubmitted required follow-up: {follow_up or ''}", 'reference_context': {'required_follow_up': expected['required_follow_up'], 'project_specific_equivalents': 'For ARM-2318, confirming or monitoring the retest closeout, cost, or disputed recovery resolution is sufficient; the follow-up need not repeat the separately graded accounting treatment. For ARM-2506, PCO-006 pricing/approval and DB-27 executed pricing refer to the same unresolved authorization.', 'equivalence_rule': 'Accept a concise action that obtains the same missing authorization/evidence; do not require the authored sentence or exact verbs.'}, 'task_context': task_context, 'evidence_scope': f'Risk Review!J{decision_row}'}])
    control = gold['decision_control']
    project_rows_ok = all((_normalize(value_sheet[f'A{row}'].value) == _normalize(expected['project_id']) for (row, expected) in zip(range(15, 19), gold['decision_rows'], strict=True)))
    pending_rows_present = all((bool(_normalize(value_sheet[f'G{row}'].value)) for row in range(15, 19)))
    queue_nonblank = bool(_normalize(value_sheet['B23'].value))
    semantic_specs.extend([{'criterion_id': 'decision_control__pending_revenue', 'expected_facts': {'pending_commercial_revenue_included': False}, 'hard_gate_met': project_rows_ok and pending_rows_present, 'hard_gate_evidence': f'project_rows_associated={project_rows_ok}; per-project pending-revenue conclusions present={pending_rows_present}', 'submitted_evidence': '\n'.join((f"{value_sheet[f'A{row}'].value}: pending revenue included = {value_sheet[f'G{row}'].value}" for row in range(15, 19))) + f"\nAggregate control: {value_sheet['B22'].value or ''}", 'reference_context': {'required_conclusion': 'No pending commercial revenue is included in any of the four close cases.', 'equivalence_rule': 'Accept No, excluded, not included, zero, or any unambiguous equivalent.'}, 'task_context': task_context, 'evidence_scope': 'Risk Review!G15:G18,B22'}, {'criterion_id': 'decision_control__review_queue', 'expected_facts': {'held_for_controller_review': control['review_queue'], 'released_with_monitoring': 'ARM-2318'}, 'hard_gate_met': project_rows_ok and queue_nonblank, 'hard_gate_evidence': f'project_rows_associated={project_rows_ok}; review_queue_nonblank={queue_nonblank}. The numeric review count is graded separately and cannot erase an otherwise correct queue.', 'submitted_evidence': f"Submitted review queue: {value_sheet['B23'].value or ''}", 'reference_context': {'held_for_controller_review': control['review_queue'], 'released_with_monitoring': 'ARM-2318', 'review_required_count': control['review_required_count'], 'equivalence_rule': 'Order and punctuation do not matter; the membership and release distinction do.'}, 'task_context': task_context, 'evidence_scope': 'Risk Review!B23'}, {'criterion_id': 'decision_control__close_release', 'expected_facts': {'close_release_decision': control['close_release_decision']}, 'hard_gate_met': project_rows_ok and bool(_normalize(value_sheet['B24'].value)), 'hard_gate_evidence': f"project_rows_associated={project_rows_ok}; decision_nonblank={bool(_normalize(value_sheet['B24'].value))}", 'submitted_evidence': f"Submitted close-release decision: {value_sheet['B24'].value or ''}", 'reference_context': {'required_decision': {'release_with_monitoring': 'ARM-2318', 'hold_for_controller_review': ['ARM-2409', 'ARM-2417', 'ARM-2506']}, 'equivalence_rule': 'Accept any unambiguous release of ARM-2318 with monitoring while ARM-2409, ARM-2417, and ARM-2506 remain held for Controller review. Reject immediate or unconditional release of the three flagged jobs.'}, 'task_context': task_context, 'evidence_scope': 'Risk Review!B24'}, {'criterion_id': 'commercial_sensitivity__basis', 'expected_facts': {'analysis_type': 'unbooked sensitivity only', 'june_close_treatment': 'no hypothetical commercial revenue is authorized, booked, or included'}, 'hard_gate_met': commercial_sensitivity_context['basis_present'], 'hard_gate_evidence': f"commercial-recovery sensitivity section contains a nonblank basis statement={commercial_sensitivity_context['basis_present']}", 'submitted_evidence': commercial_sensitivity_context['section_text'], 'reference_context': {'required_conclusion': 'The commercial-recovery case is a hypothetical sensitivity only. The pending revenue remains unauthorized and unbooked and is not included in the June 30 close.', 'equivalence_rule': 'Accept ordinary professional wording that clearly preserves the same booked-versus-hypothetical boundary. Reject language that presents the recoveries as recorded, approved, or part of the June close.'}, 'task_context': task_context, 'evidence_scope': 'commercial-recovery sensitivity section only'}])
    return _attach_semantic_review(result, task_id='task_004', evidence=_legacy_artifact_evidence(path), artifact_type='controller WIP workbook', specs=semantic_specs, decision_failure_cap=None, always_judge=True, execution_mode='scoped_per_criterion')

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
    inherited_background = added_background is None
    visual_system_ok = (background_matches or inherited_background) and len(matched_colors) >= 2
    visual_evidence = f'background={added_background}; reference_backgrounds={sorted(reference_backgrounds)}; inherited_background={inherited_background}; matched_reference_colors={sorted(matched_colors)}'
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
    original_structure_signatures = [tuple(sorted((str(shape.shape_type) for shape in slide.shapes))) for slide in original.slides]
    current_structure_signatures = [tuple(sorted((str(shape.shape_type) for shape in slide.shapes))) for slide in presentation.slides]
    added = presentation.slides[-1] if len(presentation.slides) > len(original.slides) else None
    added_text = _normalize(_slide_text(added)) if added else ''
    title_ok = 'q2 covenant headroom' in added_text and 'posted' in added_text and ('pro forma' in added_text)
    table_shapes = [shape for shape in added.shapes if getattr(shape, 'has_table', False)] if added else []

    def is_covenant_comparison_table(shape: Any) -> bool:
        rows = [[_normalize(cell.text) for cell in row.cells] for row in shape.table.rows]
        if not rows:
            return False
        headers = rows[0]
        body = ' '.join((value for row in rows[1:] for value in row))
        return any((value in {'posted', 'posted basis', 'recorded', 'recorded basis'} for value in headers)) and any((value in {'pro forma', 'pro forma basis', 'adjusted basis'} for value in headers)) and any(('threshold' in value or 'covenant limit' in value for value in headers)) and any(('status' in value or 'compliance' in value or 'result' in value for value in headers)) and ('leverage' in body or ('funded debt' in body and 'adjusted ebitda' in body)) and ('fccr' in body or 'fixed charge coverage' in body) and ('tangible net worth' in body or 'tnw' in body)
    main_table_shape = next((shape for shape in table_shapes if is_covenant_comparison_table(shape)), None)
    has_table = main_table_shape is not None
    (visual_system_ok, visual_system_evidence, title_hierarchy_ok, title_hierarchy_evidence, table_readability_ok, table_readability_evidence) = _pptx_task_015_style_checks(presentation, original, added, main_table_shape)
    metric_failures: list[str] = []
    metric_results: list[Criterion] = []
    metric_semantic_context: dict[str, dict[str, Any]] = {}
    table_rows: list[list[str]] = []
    joined = ''
    if has_table:
        table = main_table_shape.table
        table_rows = [[cell.text for cell in row.cells] for row in table.rows]
        joined = ' | '.join((value for row in table_rows for value in row))
        headers = [_normalize(value) for value in table_rows[0]] if table_rows else []
        headroom_tokens = ('headroom', 'capacity', 'cushion', 'deterioration')

        def column_index(role: str) -> int | None:
            for (index, header) in enumerate(headers):
                if role == 'posted' and (header in {'posted', 'posted basis', 'recorded', 'recorded basis'} or ('posted' in header and 'pro forma' not in header and (not any((token in header for token in headroom_tokens))))):
                    return index
                if role == 'pro_forma' and (header in {'pro forma', 'pro forma basis', 'adjusted basis'} or ('pro forma' in header and (not any((token in header for token in headroom_tokens))))):
                    return index
                if role == 'threshold' and any((token in header for token in ('threshold', 'covenant limit', 'requirement'))):
                    return index
                if role == 'posted_headroom' and ('posted' in header and 'pro forma' not in header and any((token in header for token in headroom_tokens))):
                    return index
                if role == 'pro_forma_headroom' and ('pro forma' in header and any((token in header for token in headroom_tokens))):
                    return index
                if role == 'status' and any((token in header for token in ('status', 'compliance', 'result'))):
                    return index
            return None
        posted_column = column_index('posted')
        pro_forma_column = column_index('pro_forma')
        threshold_column = column_index('threshold')
        posted_headroom_column = column_index('posted_headroom')
        pro_forma_headroom_column = column_index('pro_forma_headroom')
        status_column = column_index('status')
        capacity = gold['deterioration_capacity']
        metric_targets = {'leverage': (gold['leverage_posted'], gold['leverage_pro_forma'], gold['max_leverage'], capacity['leverage']['posted'], capacity['leverage']['pro_forma']), 'fccr': (gold['fccr_posted'], gold['fccr_pro_forma'], gold['min_fccr'], capacity['fccr']['posted'], capacity['fccr']['pro_forma']), 'tangible net worth': (gold['tangible_net_worth_posted'], gold['tangible_net_worth_pro_forma'], gold['min_tangible_net_worth'], capacity['tangible_net_worth']['posted'], capacity['tangible_net_worth']['pro_forma'])}
        metric_aliases = {'leverage': ('leverage', 'funded debt adjusted ebitda', 'debt adjusted ebitda'), 'fccr': ('fccr', 'fixed charge coverage'), 'tangible net worth': ('tangible net worth', 'tnw')}
        for (metric, expected) in metric_targets.items():
            row = next((values for values in table_rows[1:] if values and any((alias in _normalize(values[0]) for alias in metric_aliases[metric]))), None)
            if row is None or None in (posted_column, pro_forma_column, threshold_column, status_column):
                metric_failures.append(f'{metric}: missing row/columns')
                metric_semantic_context[metric] = {'submitted_evidence': 'required covenant row or column is missing', 'numeric_gate': False, 'numeric_gate_evidence': 'required covenant row or one of Posted/Pro Forma/Threshold/Status columns is missing', 'status_nonblank': False, 'expected': expected[:3]}
                for field in ('posted', 'pro_forma', 'threshold', 'posted_headroom', 'pro_forma_headroom', 'status'):
                    metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__{field}", f'{metric.title()} `{field}` is correctly presented', False, 'missing row or required column'))
                continue
            row_numeric_gate = all((_text_contains_number(row[column], target, abs_tol=0.0001) for (column, target) in zip((posted_column, pro_forma_column, threshold_column), expected[:3], strict=True)))
            displayed_values_parseable = all((_number(re.sub('^[<>=≤≥~\\s]+', '', str(row[column] or ''))) is not None for column in (posted_column, pro_forma_column, threshold_column)))
            metric_semantic_context[metric] = {'submitted_evidence': ' | '.join(row), 'numeric_gate': row_numeric_gate, 'numeric_gate_evidence': f'this exact row contains the correct posted, pro-forma, and threshold values={row_numeric_gate}', 'displayed_values_parseable': displayed_values_parseable, 'status_nonblank': bool(_normalize(row[status_column])), 'expected': expected[:3]}
            for (field, column, target) in zip(('posted', 'pro_forma', 'threshold'), (posted_column, pro_forma_column, threshold_column), expected[:3], strict=True):
                met = _text_contains_number(row[column], target, abs_tol=0.0001)
                if not met:
                    metric_failures.append(f'{metric} {field}: value={row[column]!r}')
                metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__{field}", f'{metric.title()} `{field}` value is correct', met, f'actual={row[column]!r}; expected={target}'))
            for (field, column, target) in (('posted_headroom', posted_headroom_column, expected[3]), ('pro_forma_headroom', pro_forma_headroom_column, expected[4])):
                actual = row[column] if column is not None and column < len(row) else None
                met = actual is not None and _text_contains_number(str(actual), target, abs_tol=1.0)
                if not met:
                    metric_failures.append(f'{metric} {field}: value={actual!r}')
                metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__{field}", f'{metric.title()} `{field}` deterioration capacity is correct', met, f'actual={actual!r}; expected={target}'))
            status_met = any((contains_concept(row[status_column], status) for status in ('compliant', 'pass', 'passes', 'within threshold', 'no breach')))
            if not status_met:
                metric_failures.append(f'{metric}: status={row[status_column]!r}')
            metric_results.append(Criterion(f"metric__{_normalize(metric).replace(' ', '_')}__status", f'{metric.title()} status is correctly reported as compliant', status_met, f'actual={row[status_column]!r}'))
    else:
        for metric in ('leverage', 'fccr', 'tangible net worth'):
            metric_id = _normalize(metric).replace(' ', '_')
            for field in ('posted', 'pro_forma', 'threshold', 'posted_headroom', 'pro_forma_headroom', 'status'):
                metric_results.append(Criterion(f'metric__{metric_id}__{field}', f'{metric.title()} `{field}` is correctly presented', False, 'required appended covenant table is missing'))

    def is_wip_reversal_table(shape: Any) -> bool:
        rows = [[_normalize(cell.text) for cell in row.cells] for row in shape.table.rows]
        if not rows:
            return False
        headers = rows[0]
        return any(('reversal' in value or 'haircut' in value for value in headers)) and any(('retained' in value and 'wip' in value for value in headers)) and any(('adjusted ebitda' in value or value == 'ebitda' for value in headers)) and any(('leverage' in value for value in headers)) and any(('fccr' in value or 'fixed charge coverage' in value for value in headers))
    sensitivity_shape = next((shape for shape in table_shapes if is_wip_reversal_table(shape)), None)
    sensitivity_results: list[Criterion] = [Criterion('structure__wip_reversal_sensitivity', 'The appended slide contains the requested proposed-WIP reversal sensitivity table', sensitivity_shape is not None, f'reversal_sensitivity_table_found={sensitivity_shape is not None}')]
    sensitivity_semantic_context: dict[int, dict[str, Any]] = {}
    if sensitivity_shape is not None:
        sensitivity_rows = [[cell.text for cell in row.cells] for row in sensitivity_shape.table.rows]
        sensitivity_headers = [_normalize(value) for value in sensitivity_rows[0]]

        def sensitivity_column(role: str) -> int | None:
            for (index, header) in enumerate(sensitivity_headers):
                capacity = any((token in header for token in ('capacity', 'headroom', 'cushion', 'deterioration')))
                if role == 'reversal' and ('reversal' in header or 'haircut' in header):
                    return index
                if role == 'retained_wip_adjustment' and 'retained' in header and ('wip' in header or 'adjustment' in header):
                    return index
                if role == 'adjusted_ebitda' and ('adjusted ebitda' in header or header == 'ebitda') and (not capacity):
                    return index
                if role == 'tangible_net_worth' and ('tangible net worth' in header or header == 'tnw') and (not capacity):
                    return index
                if role == 'leverage' and 'leverage' in header and (not capacity):
                    return index
                if role == 'fccr_numerator' and ('fccr' in header or 'fixed charge coverage' in header) and ('numerator' in header):
                    return index
                if role == 'fccr' and (header == 'fccr' or 'fixed charge coverage' in header) and ('numerator' not in header) and (not capacity):
                    return index
                if role == 'leverage_deterioration_capacity' and 'leverage' in header and capacity:
                    return index
                if role == 'fccr_deterioration_capacity' and ('fccr' in header or 'fixed charge coverage' in header) and capacity:
                    return index
                if role == 'tangible_net_worth_headroom' and ('tangible net worth' in header or 'tnw' in header) and capacity:
                    return index
                if role == 'status' and any((token in header for token in ('status', 'compliance', 'result'))):
                    return index
            return None
        sensitivity_columns = {role: sensitivity_column(role) for role in ('reversal', 'retained_wip_adjustment', 'adjusted_ebitda', 'tangible_net_worth', 'leverage', 'fccr_numerator', 'fccr', 'leverage_deterioration_capacity', 'fccr_deterioration_capacity', 'tangible_net_worth_headroom', 'status')}
        for expected in gold['wip_reversal_sensitivity']:
            reversal_percent = int(round(float(expected['reversal_percent']) * 100))
            reversal_col = sensitivity_columns['reversal']
            row = next((values for values in sensitivity_rows[1:] if reversal_col is not None and reversal_col < len(values) and _close(_number(values[reversal_col]), float(expected['reversal_percent']), abs_tol=0.0001, rel_tol=0.0)), None)
            value_fields = ('retained_wip_adjustment', 'adjusted_ebitda', 'tangible_net_worth', 'leverage', 'fccr_numerator', 'fccr', 'leverage_deterioration_capacity', 'fccr_deterioration_capacity', 'tangible_net_worth_headroom')
            parsed_values: dict[str, float | None] = {}
            for field in value_fields:
                column = sensitivity_columns[field]
                actual = row[column] if row is not None and column is not None and (column < len(row)) else None
                tolerance = 0.0001 if field in {'leverage', 'fccr'} else 1.0
                met = actual is not None and _text_contains_number(str(actual), float(expected[field]), abs_tol=tolerance)
                parsed_values[field] = _number(actual) if actual is not None else None
                sensitivity_results.append(Criterion(f'sensitivity__reversal_{reversal_percent}__{field}', f"The {reversal_percent}% reversal sensitivity reports the correct {field.replace('_', ' ')}", met, f'submitted={actual!r}; expected={expected[field]}'))
            status_column = sensitivity_columns['status']
            submitted_status = str(row[status_column]) if row is not None and status_column is not None and (status_column < len(row)) else ''
            required_values_present = all((parsed_values[field] is not None for field in value_fields))
            sensitivity_results.append(Criterion(f'sensitivity__reversal_{reversal_percent}__status', f'The {reversal_percent}% reversal sensitivity reports the correct covenant status', bool(submitted_status.strip()) and required_values_present, submitted_status, semantic=True))
            sensitivity_semantic_context[reversal_percent] = {'submitted_evidence': ' | '.join(row) if row is not None else 'required row is missing', 'hard_gate': bool(submitted_status.strip()) and required_values_present, 'expected': expected}
            retained = parsed_values['retained_wip_adjustment']
            ebitda = parsed_values['adjusted_ebitda']
            leverage = parsed_values['leverage']
            numerator = parsed_values['fccr_numerator']
            fccr = parsed_values['fccr']
            equations = {'retained_wip': (retained is not None and _close(retained, gold['proposed_wip_adjustment'] * (1 - float(expected['reversal_percent'])), abs_tol=1.0, rel_tol=0.0), f"retained={retained!r}; proposed_wip={gold['proposed_wip_adjustment']}; reversal={expected['reversal_percent']}"), 'leverage': (ebitda is not None and leverage is not None and _close(leverage, gold['funded_debt'] / ebitda, abs_tol=0.01, rel_tol=0.0), f"leverage={leverage!r}; funded_debt={gold['funded_debt']}; adjusted_ebitda={ebitda!r}"), 'fccr': (numerator is not None and fccr is not None and _close(fccr, numerator / gold['fixed_charges'], abs_tol=0.01, rel_tol=0.0), f"fccr={fccr!r}; numerator={numerator!r}; fixed_charges={gold['fixed_charges']}")}
            for (equation, (met, evidence)) in equations.items():
                sensitivity_results.append(Criterion(f'sensitivity__reversal_{reversal_percent}__equation_{equation}', f"The {reversal_percent}% reversal {equation.replace('_', ' ')} calculation is internally consistent", met, evidence))
    else:
        for expected in gold['wip_reversal_sensitivity']:
            reversal_percent = int(round(float(expected['reversal_percent']) * 100))
            for field in ('retained_wip_adjustment', 'adjusted_ebitda', 'tangible_net_worth', 'leverage', 'fccr_numerator', 'fccr', 'leverage_deterioration_capacity', 'fccr_deterioration_capacity', 'tangible_net_worth_headroom', 'status', 'equation_retained_wip', 'equation_leverage', 'equation_fccr'):
                sensitivity_results.append(Criterion(f'sensitivity__reversal_{reversal_percent}__{field}', f"The {reversal_percent}% reversal sensitivity {field.replace('_', ' ')} is correct", False, 'required sensitivity table or row is missing', semantic=field == 'status'))
            sensitivity_semantic_context[reversal_percent] = {'submitted_evidence': 'required sensitivity table or row is missing', 'hard_gate': False, 'expected': expected}
    non_table_text_blocks = [str(shape.text).strip() for shape in added.shapes if added is not None and getattr(shape, 'has_text_frame', False) and (not getattr(shape, 'has_table', False)) and str(shape.text or '').strip()] if added is not None else []
    table_row_text_blocks = [' | '.join((str(cell.text or '').strip() for cell in row.cells)) for shape in added.shapes if getattr(shape, 'has_table', False) for row in shape.table.rows] if added is not None else []
    support_segments = [segment.strip() for block in (*non_table_text_blocks, *table_row_text_blocks) for segment in re.split('(?:[\\n;•]+|(?<=[A-Za-z0-9])\\.(?=\\s+[A-Z]))', block) if segment.strip()]

    def support_evidence(*aliases: tuple[str, ...]) -> str:
        return '\n'.join((segment for segment in support_segments if any((all((token in _normalize(segment) for token in alias)) for alias in aliases))))
    calculation_support_specs = {'funded_debt': ('Funded debt', gold['funded_debt'], (('funded', 'debt'),)), 'posted_adjusted_ebitda': ('Posted adjusted EBITDA', gold['ltm_adjusted_ebitda_posted'], (('posted', 'adjusted', 'ebitda'), ('posted', 'ebitda'))), 'proposed_wip_adjustment': ('Proposed WIP adjustment', gold['proposed_wip_adjustment'], (('proposed', 'wip', 'adjustment'), ('june', 'wip', 'adjustment'))), 'pro_forma_adjusted_ebitda': ('Pro-forma adjusted EBITDA', gold['ltm_adjusted_ebitda_pro_forma'], (('pro', 'forma', 'adjusted', 'ebitda'), ('pro', 'forma', 'ebitda'))), 'cash_taxes': ('Cash taxes', gold['cash_taxes'], (('cash', 'tax'),)), 'ltm_fixed_asset_additions': ('LTM fixed-asset additions', gold['ltm_fixed_asset_additions'], (('fixed', 'asset', 'addition'), ('ltm', 'capex', 'addition'))), 'direct_equipment_financing': ('Direct LTM equipment financing', gold['ltm_direct_equipment_financing'], (('direct', 'equipment', 'financing'), ('equipment', 'financing', 'proceeds'))), 'unfunded_capex': ('Unfunded capex', gold['unfunded_capex'], (('unfunded', 'capex'), ('unfunded', 'capital', 'expenditure'))), 'cash_interest': ('Cash interest', gold['cash_interest'], (('cash', 'interest'),)), 'scheduled_principal': ('Scheduled principal', gold['scheduled_principal'], (('scheduled', 'principal'),)), 'fixed_charges': ('Fixed charges', gold['fixed_charges'], (('fixed', 'charge'),)), 'posted_fccr_numerator': ('Posted FCCR numerator', gold['fccr_numerator_posted'], (('posted', 'fccr', 'numerator'), ('posted', 'coverage', 'numerator'))), 'pro_forma_fccr_numerator': ('Pro-forma FCCR numerator', gold['fccr_numerator_pro_forma'], (('pro', 'forma', 'fccr', 'numerator'), ('pro', 'forma', 'coverage', 'numerator')))}
    calculation_support_results: list[Criterion] = []
    calculation_support_met: dict[str, bool] = {}
    for (field, (label, expected, aliases)) in calculation_support_specs.items():
        evidence = support_evidence(*aliases)
        met = _text_contains_number(evidence, expected, abs_tol=1.0)
        if field == 'cash_taxes' and expected == 0:
            met = met or 'zero' in _normalize(evidence)
        calculation_support_met[field] = met
        calculation_support_results.append(Criterion(f'calculation_support__{field}', f"The slide's calculation-support block shows the correct {label}", met, f'submitted={evidence[:800]!r}; expected={expected}'))
    equation_specs = {'ebitda_bridge': (('posted_adjusted_ebitda', 'proposed_wip_adjustment', 'pro_forma_adjusted_ebitda'), '(?:\\+|\\bplus\\b|\\badd(?:ed)?\\s+to\\b)', (('ebitda', 'pro', 'forma'), ('ebitda', 'wip'))), 'unfunded_capex': (('ltm_fixed_asset_additions', 'direct_equipment_financing', 'unfunded_capex'), '(?:\\s[-−]\\s|\\bless\\b|\\bminus\\b)', (('unfunded', 'capex'), ('fixed', 'asset', 'addition'))), 'fixed_charges': (('cash_interest', 'scheduled_principal', 'fixed_charges'), '(?:\\+|\\bplus\\b|\\bsum\\b|\\badd(?:ed)?\\s+to\\b)', (('fixed', 'charge'), ('cash', 'interest'), ('scheduled', 'principal'))), 'posted_fccr_numerator': (('posted_adjusted_ebitda', 'cash_taxes', 'unfunded_capex', 'posted_fccr_numerator'), '(?:\\s[-−]\\s|\\bless\\b|\\bminus\\b)', (('posted', 'fccr', 'numerator'), ('posted', 'coverage', 'numerator'))), 'pro_forma_fccr_numerator': (('pro_forma_adjusted_ebitda', 'cash_taxes', 'unfunded_capex', 'pro_forma_fccr_numerator'), '(?:\\s[-−]\\s|\\bless\\b|\\bminus\\b)', (('pro', 'forma', 'fccr', 'numerator'), ('pro', 'forma', 'coverage', 'numerator')))}
    for (field, (required_fields, operator_pattern, aliases)) in equation_specs.items():
        candidates = [segment for segment in support_segments if any((all((token in _normalize(segment) for token in alias)) for alias in aliases))]

        def equation_segment_result(segment: str) -> tuple[dict[str, bool], bool]:
            values = {name: _text_contains_number(segment, calculation_support_specs[name][1], abs_tol=1.0) or (name == 'cash_taxes' and calculation_support_specs[name][1] == 0 and ('zero' in _normalize(segment))) for name in required_fields}
            operator = bool(re.search(operator_pattern, segment, flags=re.IGNORECASE))
            return (values, operator)
        candidate_results = [(segment, *equation_segment_result(segment)) for segment in candidates]
        matching = next(((segment, values, operator) for (segment, values, operator) in candidate_results if all(values.values()) and operator), None)
        met = matching is not None
        calculation_support_results.append(Criterion(f'calculation_support__equation__{field}', f"The calculation-support block displays the {field.replace('_', ' ')} arithmetic", met, f'same_segment_required=True; matching_segment={matching[0][:1000]!r}' if matching else f'same_segment_required=True; candidates={candidate_results!r}'))
    binding_phrase_blocks = [block for block in non_table_text_blocks if any((contains_concept(block, phrase) for phrase in ('binding', 'most restrictive', 'limiting covenant', 'least headroom')))]
    binding_metric_blocks = [block for block in non_table_text_blocks if any((token in _normalize(block) for token in ('fccr', 'fixed charge coverage', 'adjusted ebitda deterioration', 'smallest like for like', 'least headroom')))]
    binding_evidence = '\n'.join(dict.fromkeys([*binding_phrase_blocks, *binding_metric_blocks]))
    binding_hard_gate = bool(binding_phrase_blocks) and any((token in _normalize(binding_evidence) for token in ('fccr', 'fixed charge coverage')))
    wip_status_evidence = '\n'.join((block for block in non_table_text_blocks if any((token in _normalize(block) for token in ('wip', 'work in progress')))))
    source_note_evidence = '\n'.join((block for block in non_table_text_blocks if any((token in _normalize(block) for token in ('source', 'basis', 'amendment', 'debt support', 'accounting', 'ledger', 'trial balance', 'wip support', 'wip workpaper')))))
    result = _result([Criterion('preservation__slide_count', 'Exactly one slide is appended to the original deck', len(presentation.slides) == len(original.slides) + 1, f'original={len(original.slides)}; current={len(presentation.slides)}'), *[Criterion(f'preservation__slide_{index:02d}', f'Original slide {index} retains its substantive content and structure', index <= len(current_structure_signatures) and current_texts[index - 1] == text and (current_structure_signatures[index - 1] == original_structure_signatures[index - 1]), f'text_match={index <= len(current_texts) and current_texts[index - 1] == text}; shape_types_match={index <= len(current_structure_signatures) and current_structure_signatures[index - 1] == original_structure_signatures[index - 1]}') for (index, text) in enumerate(original_texts, start=1)], Criterion('structure__title', 'The appended slide has the requested title', title_ok, f'title_ok={title_ok}'), Criterion('structure__table', 'The appended slide contains a real PowerPoint covenant-comparison table', has_table, f'table_count={len(table_shapes)}; covenant_comparison_found={has_table}'), Criterion('style__visual_system', 'The appended slide uses the existing deck background and visual palette', visual_system_ok, visual_system_evidence), Criterion('style__title_hierarchy', "The appended slide title follows the deck's established hierarchy", title_hierarchy_ok, title_hierarchy_evidence), Criterion('style__table_readability', 'The covenant comparison is presented as a bounded, clearly headed review table', table_readability_ok, table_readability_evidence), *metric_results, Criterion('status__proposed', 'The WIP entry is identified as proposed', 'proposed' in added_text, added_text[:800]), Criterion('status__unposted', 'The WIP entry is identified as unposted', 'unposted' in added_text or 'not posted' in added_text, added_text[:800]), Criterion('status__compliant', 'Both covenant bases are reported as passing their thresholds', any((contains_concept(added_text, status) for status in ('compliant', 'pass', 'passes', 'within threshold', 'no breach'))), added_text[:800]), Criterion('headroom__binding_covenant', 'The slide identifies FCCR as the binding covenant on the disclosed dollar-deterioration basis', binding_hard_gate, binding_evidence[:1200]), Criterion('source__authority', 'The slide identifies the controlling calculation, agreement, debt, accounting, and WIP authorities', bool(source_note_evidence), source_note_evidence[:1600]), *calculation_support_results, *sensitivity_results])
    status_table_gate = has_table and all((bool(context.get('displayed_values_parseable')) and bool(context.get('status_nonblank')) for context in metric_semantic_context.values()))
    appended_title = ''
    if added is not None:
        title_shape = getattr(added.shapes, 'title', None)
        appended_title = str(getattr(title_shape, 'text', '') or '')
        if not _normalize(appended_title):
            text_candidates = [str(shape.text).strip() for shape in added.shapes if hasattr(shape, 'text') and str(shape.text or '').strip()]
            appended_title = next((value for value in text_candidates if 'covenant' in _normalize(value) and 'headroom' in _normalize(value)), text_candidates[0] if text_candidates else '')
    table_evidence = '\n'.join((' | '.join(row) for row in table_rows))
    task_context = {'assignment': 'Append one Q2 covenant headroom slide to the existing lender update deck.', 'grading_boundary': 'All displayed covenant values are graded separately and deterministically. Each semantic judge sees only the title, status statement, or covenant row relevant to its criterion and must accept ordinary professional wording with the same meaning.', 'reporting_basis': 'Posted is the recorded-company basis. Pro forma includes only the disclosed proposed, unposted $1,276,325.70 WIP entry. Both bases are compared with the executed covenant thresholds. Leverage/FCCR headroom is adjusted-EBITDA deterioration capacity; TNW headroom is the dollar cushion to its minimum.'}
    semantic_specs: list[dict[str, Any]] = [{'criterion_id': 'structure__title', 'expected_facts': {'required_meaning': 'Q2 covenant headroom comparing posted and pro-forma bases'}, 'hard_gate_met': bool(added) and bool(_normalize(appended_title)), 'hard_gate_evidence': f'appended slide and nonblank title present={bool(added) and bool(_normalize(appended_title))}', 'submitted_evidence': f'Submitted appended-slide title: {appended_title}', 'reference_context': {'required_meaning': 'Q2 covenant headroom comparing posted and pro-forma bases', 'equivalence_rule': 'Accept normal title variants; do not require the authored punctuation or exact phrase.'}, 'task_context': task_context, 'evidence_scope': 'appended slide title only'}, {'criterion_id': 'status__compliant', 'expected_facts': {'overall_conclusion': 'all three covenants are compliant on both posted and pro-forma bases'}, 'hard_gate_met': status_table_gate and bool(table_evidence), 'hard_gate_evidence': f'all three required covenant rows contain parseable displayed values and nonblank statuses={status_table_gate}', 'submitted_evidence': table_evidence, 'reference_context': {'required_conclusion': 'Leverage is below its maximum; FCCR and tangible net worth exceed their minimums on both bases, so no covenant breach exists.', 'equivalence_rule': 'Accept compliant, pass/passes, within threshold, no breach, or an equally clear professional conclusion.', 'consistency_rule': 'The conclusion must also agree with the posted, pro-forma, and threshold values displayed in the submitted table. Numeric accuracy itself is scored separately.'}, 'task_context': task_context, 'evidence_scope': 'appended covenant table only'}, {'criterion_id': 'status__proposed', 'expected_facts': {'wip_entry_state': 'proposed'}, 'hard_gate_met': bool(added), 'hard_gate_evidence': f'appended slide present={bool(added)}', 'submitted_evidence': wip_status_evidence, 'reference_context': {'required_meaning': 'the WIP adjustment is proposed rather than recorded or approved', 'equivalence_rule': 'Accept proposed, draft, pending approval, or an unambiguous equivalent.'}, 'task_context': task_context, 'evidence_scope': 'appended slide WIP-status text only'}, {'criterion_id': 'status__unposted', 'expected_facts': {'wip_entry_posting_state': 'unposted'}, 'hard_gate_met': bool(added), 'hard_gate_evidence': f'appended slide present={bool(added)}', 'submitted_evidence': wip_status_evidence, 'reference_context': {'required_meaning': 'the proposed WIP entry has not been posted to the accounting records', 'equivalence_rule': 'Accept unposted, not posted, not yet booked/recorded, or an unambiguous equivalent; reject language that says it is posted.'}, 'task_context': task_context, 'evidence_scope': 'appended slide WIP-status text only'}, {'criterion_id': 'headroom__binding_covenant', 'expected_facts': {'binding_covenant': 'FCCR', 'comparison_basis': 'lowest like-for-like dollar deterioration capacity'}, 'hard_gate_met': binding_hard_gate, 'hard_gate_evidence': 'the appended slide contains a nonblank binding-covenant conclusion that identifies FCCR', 'submitted_evidence': binding_evidence, 'reference_context': {'required_conclusion': 'FCCR is the binding covenant because it has the lowest dollar amount of adjusted-EBITDA deterioration capacity on both posted and pro-forma bases.', 'verified_comparison': 'Posted capacity: FCCR $2,035,453.11 versus leverage $2,604,739.37. Pro-forma capacity: FCCR $3,311,778.81 versus leverage $3,881,065.07. Tangible-net-worth headroom is a separate balance-sheet dollar cushion.', 'equivalence_rule': 'Accept ordinary professional wording that unambiguously identifies FCCR as binding on the disclosed like-for-like deterioration-capacity basis.'}, 'task_context': task_context, 'evidence_scope': 'appended slide binding-covenant callout only'}, {'criterion_id': 'source__authority', 'expected_facts': {'calculation_scaffold': 'Q2 covenant headroom working workbook', 'legal_authority': 'executed U.S. Bank amendment', 'debt_authority': 'current June 30 debt support', 'posted_authority': 'June 30 posted accounting records', 'pro_forma_authority': 'Controller-supported proposed and unposted June WIP adjustment'}, 'hard_gate_met': bool(source_note_evidence), 'hard_gate_evidence': f'nonblank source/basis note present={bool(source_note_evidence)}', 'submitted_evidence': source_note_evidence, 'reference_context': {'required_authorities': 'The note must identify, in ordinary professional wording, the Q2 covenant working/calculation support, executed bank amendment, current June 30 debt support, posted accounting basis, and proposed June WIP support.', 'equivalence_rule': 'Accept shortened filenames, business descriptions, acronyms, and combined source statements when each authority remains unambiguous. Do not require the authored filenames or exact wording.'}, 'task_context': task_context, 'evidence_scope': 'appended slide source/basis note only'}]
    for metric in ('leverage', 'fccr', 'tangible net worth'):
        context = metric_semantic_context.get(metric, {'submitted_evidence': 'required covenant row is missing', 'numeric_gate': False, 'numeric_gate_evidence': 'required covenant row is missing', 'displayed_values_parseable': False, 'status_nonblank': False, 'expected': (None, None, None)})
        (posted, pro_forma, threshold) = context['expected']
        relationship = 'posted and pro-forma values are below the maximum threshold' if metric == 'leverage' else 'posted and pro-forma values are above the minimum threshold'
        semantic_specs.append({'criterion_id': f"metric__{_normalize(metric).replace(' ', '_')}__status", 'expected_facts': {'metric': metric, 'posted': posted, 'pro_forma': pro_forma, 'threshold': threshold, 'required_relationship': relationship, 'required_status': 'compliant'}, 'hard_gate_met': bool(context['displayed_values_parseable']) and bool(context['status_nonblank']), 'hard_gate_evidence': f"displayed_values_parseable={context['displayed_values_parseable']}; status_nonblank={context['status_nonblank']}", 'submitted_evidence': context['submitted_evidence'], 'reference_context': {'required_relationship': relationship, 'required_status': 'compliant', 'equivalence_rule': 'Accept compliant, pass/passes, within threshold, no breach, or an equally clear professional equivalent.', 'consistency_rule': 'The status must agree with the relationship between the posted/pro-forma values and threshold displayed in this submitted row. Numeric accuracy itself is scored separately.'}, 'task_context': task_context, 'evidence_scope': f'{metric} row in appended covenant table only'})
    for reversal_percent in (25, 50, 100):
        context = sensitivity_semantic_context[reversal_percent]
        expected = context['expected']
        semantic_specs.append({'criterion_id': f'sensitivity__reversal_{reversal_percent}__status', 'expected_facts': {'scenario': f'{reversal_percent}% reversal of the proposed WIP adjustment', 'leverage': expected['leverage'], 'maximum_leverage': gold['max_leverage'], 'fccr': expected['fccr'], 'minimum_fccr': gold['min_fccr'], 'tangible_net_worth': expected['tangible_net_worth'], 'minimum_tangible_net_worth': gold['min_tangible_net_worth'], 'required_status': 'all three covenants remain compliant'}, 'hard_gate_met': context['hard_gate'], 'hard_gate_evidence': f"the exact scenario row contains all required parseable values and a nonblank status={context['hard_gate']}", 'submitted_evidence': context['submitted_evidence'], 'reference_context': {'required_relationship': 'Leverage remains below 3.00x; FCCR remains above 1.20x; tangible net worth remains above $2,500,000 in this scenario.', 'equivalence_rule': 'Accept compliant, passes, within threshold, no breach, or an equally clear professional conclusion. Numeric accuracy is scored separately.'}, 'task_context': task_context, 'evidence_scope': f'{reversal_percent}% proposed-WIP reversal sensitivity row only'})
    return _attach_semantic_review(result, task_id='task_015', evidence=_legacy_artifact_evidence(path), artifact_type='board presentation', specs=semantic_specs, always_judge=True, execution_mode='scoped_per_criterion')

def _document_text(document: Document) -> str:
    chunks = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            chunks.extend((cell.text for cell in row.cells))
    return '\n'.join(chunks)
TASK_001_ARTIFACT = Path('Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx')

def _docx_table_rows_by_roles(document: Document, header_aliases: dict[str, tuple[str, ...]], *, expected_row_count: int | None=None, allow_positional_fallback: bool=False) -> list[dict[str, str]]:
    """Return a table using canonical column roles rather than literal headers.

    Word workpapers are often relabeled during review (for example, ``Amount``
    instead of ``Result`` or ``Responsible party`` instead of ``Owner``).  Such
    presentation changes are harmless.  The parser therefore accepts a bounded
    set of professional header equivalents while still requiring one and only
    one physical column for every requested role.
    """
    normalized_aliases = {role: {_normalize(alias) for alias in aliases} for (role, aliases) in header_aliases.items()}
    positional_candidates: list[list[dict[str, str]]] = []
    for table in document.tables:
        if not table.rows:
            continue
        physical_headers = [_normalize(cell.text) for cell in table.rows[0].cells]
        role_indexes: dict[str, int] = {}
        ambiguous = False
        for (role, aliases) in normalized_aliases.items():
            matches = [index for (index, header) in enumerate(physical_headers) if header in aliases]
            if len(matches) != 1:
                ambiguous = True
                break
            role_indexes[role] = matches[0]
        if ambiguous or len(set(role_indexes.values())) != len(role_indexes):
            if allow_positional_fallback and len(table.rows[0].cells) == len(header_aliases) and (expected_row_count is None or len(table.rows) - 1 == expected_row_count):
                roles = list(header_aliases)
                positional_candidates.append([{role: row.cells[index].text.strip() for (index, role) in enumerate(roles)} for row in table.rows[1:]])
            continue
        rows = [{role: row.cells[index].text.strip() for (role, index) in role_indexes.items()} for row in table.rows[1:]]
        if expected_row_count is None or len(rows) >= expected_row_count:
            return rows
    return positional_candidates[0] if len(positional_candidates) == 1 else []

def _usable_docx_value(value: Any) -> bool:
    text = str(value or '').strip()
    return bool(text) and (not bool(re.search('\\[(?:enter|select|complete|cite|record|explain|state|list)\\b', text, flags=re.I)))

def _task_001_has_cutoff_date(text: Any) -> bool:
    """Recognize ordinary business renderings of the June 30, 2026 cutoff."""
    raw = str(text or '')
    normalized = _normalize(raw)
    if any((value in normalized for value in ('june 30 2026', 'jun 30 2026', '30 june 2026', '30 jun 2026'))):
        return True
    return bool(re.search('(?<!\\d)(?:0?6\\s*[/.-]\\s*30\\s*[/.-]\\s*(?:20)?26|2026\\s*[/.-]\\s*0?6\\s*[/.-]\\s*30)(?!\\d)', raw, flags=re.I) or re.search('(?<!\\d)0?6\\s+30\\s+(?:20)?26(?!\\d)', normalized))

def _task_001_metadata_value(metadata_fields: dict[str, str], *aliases: str) -> str:
    """Return a completed metadata value under a professional label alias."""
    for alias in aliases:
        value = metadata_fields.get(_normalize(alias), '')
        if _usable_docx_value(value):
            return value
    return ''

def _document_has_prohibited_signature(document: Document) -> bool:
    """Detect an actual signature block without flagging cited signed sources."""
    paragraphs = list(document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for paragraph in paragraphs:
        marker = _normalize(paragraph.text)
        if re.match('^(?:signed by|signature|electronic signature|digitally signed by)\\b', marker):
            return True
        if re.match('^s [a-z][a-z ]{1,80}$', marker):
            return True
    for part in document.part.package.parts:
        part_name = str(part.partname).casefold()
        content_type = str(part.content_type).casefold()
        if '_xmlsignatures' in part_name or 'digital-signature' in content_type:
            return True
    for element in document.element.body.iter():
        metadata = ' '.join((str(element.get(key, '')) for key in ('name', 'descr', 'title'))).casefold()
        if 'signature' in metadata:
            return True
    return False

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
TASK_001_V24_TABLE_ALIASES: dict[str, dict[str, tuple[str, ...]]] = {'close': {'metric': ('Metric', 'Measure', 'Line item'), 'current': ('Current / submitted', 'Current or submitted', 'Submitted / current', 'Current value', 'Submitted value', 'Current case'), 'recommended': ('Recommended close', 'June close', 'Finance close', 'Adjusted close', 'Policy-corrected close', 'Policy corrected close'), 'change': ('Change / control', 'Change or control', 'Change', 'Variance'), 'source': ('Source or conclusion', 'Source / conclusion', 'Source and conclusion', 'Source / basis', 'Basis / conclusion', 'Source note')}, 'bridge': {'metric': ('Metric', 'Measure', 'Line item'), 'may': ('May final', 'Final May', 'May close', 'Prior close'), 'june': ('June close', 'Recommended close', 'June recommended', 'Current close'), 'change': ('Change', 'June less May', 'Variance', 'Movement'), 'driver': ('Driver', 'Explanation', 'Change driver', 'Comment')}, 'impact': {'metric': ('Metric', 'Measure', 'Line item'), 'submitted': ('Submitted PM case', 'PM case', 'Submitted case', 'Management case'), 'recommended': ('Recommended close', 'Finance close', 'Adjusted close', 'Policy-corrected close', 'Policy corrected close'), 'impact': ('Recommended less submitted', 'Finance less PM', 'Impact', 'Difference', 'Variance')}, 'reconciliation': {'control': ('Control line', 'Control', 'Metric', 'Line item'), 'base': ('May / opening', 'May opening', 'Opening', 'Base', 'May final'), 'activity': ('June activity / approved change', 'June activity', 'Activity / change', 'Current-period activity', 'Approved change'), 'close': ('June close', 'Closing', 'Close', 'Ending balance'), 'variance': ('Variance / status', 'Variance', 'Control status', 'Tie / status')}, 'population': {'source': ('Source / population', 'Source population', 'Source', 'Population'), 'source_count': ('Source records', 'Record count', 'Source count', 'Entries'), 'source_amount': ('Source amount', 'Population amount', 'Source total'), 'journal_count': ('Posted journals', 'Journal count', 'Ledger documents'), 'ledger_amount': ('Ledger amount', 'Posted amount', 'Ledger total'), 'variance': ('Variance', 'Difference', 'Control variance'), 'exceptions': ('Exceptions / resolution', 'Exceptions and resolution', 'Exceptions', 'Resolution / status', 'Control conclusion')}, 'etc': {'component': ('Component', 'ETC component', 'Cost component', 'Line item'), 'pm': ('Submitted PM ETC', 'PM ETC', 'Submitted ETC', 'Management ETC'), 'correction': ('Finance adjustment', 'Finance correction', 'Close adjustment', 'Adjustment'), 'close': ('Recommended close ETC', 'Finance close ETC', 'Adjusted ETC', 'Policy-corrected ETC', 'Policy corrected ETC'), 'source': ('Source / treatment', 'Source and treatment', 'Basis', 'Source / basis')}, 'judgment': {'question': ('Question', 'Issue', 'Decision', 'Assessment'), 'conclusion': ('Conclusion', 'Treatment', 'Decision / conclusion'), 'evidence': ('Controlling evidence / policy', 'Evidence / policy', 'Basis / authority', 'Controlling support', 'Evidence and policy'), 'clearance': ('Clearance or approval needed', 'Clearance / approval', 'Approval needed', 'Required clearance', 'Next gate')}, 'posting': {'stage': ('Stage', 'Posting stage', 'Entry stage'), 'account': ('Account / description', 'Account and description', 'Account', 'GL account'), 'debit': ('Debit', 'Debits'), 'credit': ('Credit', 'Credits'), 'basis': ('Basis / control', 'Basis and control', 'Basis', 'Control / note')}, 'evidence': {'role': ('Source role', 'Control', 'Evidence role', 'Source / control'), 'version': ('Version / cutoff', 'Version and cutoff', 'Version / date', 'Cutoff'), 'authority': ('Authority / status', 'Authority and status', 'Status / authority', 'Authority'), 'fact': ('Fact used and conclusion', 'Fact / conclusion', 'Use / conclusion', 'Fact used')}, 'action': {'owner': ('Owner', 'Responsible', 'Responsible party', 'Function'), 'action': ('Required action', 'Action', 'Next action'), 'evidence': ('Completion evidence', 'Evidence to close', 'Close evidence'), 'due': ('Due', 'Timing', 'Due date'), 'status': ('Status', 'Action status')}, 'disposition': {'reviewer': ('Reviewer', 'Controller / reviewer', 'Approver'), 'disposition': ('Disposition', 'Decision', 'Review status'), 'date': ('Review date', 'Date', 'Disposition date'), 'comments': ('Comments / conditions', 'Comments and conditions', 'Comments', 'Conditions')}}

def _task_001_v24_rows(document: Document | None, table_name: str) -> list[dict[str, str]]:
    if document is None:
        return []
    return _docx_table_rows_by_roles(document, TASK_001_V24_TABLE_ALIASES[table_name], allow_positional_fallback=False)

def _task_001_v24_row(rows: list[dict[str, str]], label_role: str, aliases: tuple[str, ...]) -> dict[str, str]:
    wanted = {_normalize(alias) for alias in aliases}
    exact = [row for row in rows if _normalize(row.get(label_role)) in wanted]
    if len(exact) == 1:
        return exact[0]
    fuzzy = []
    for row in rows:
        label = _normalize(row.get(label_role))
        if not label:
            continue
        if any((len(alias) >= 4 and alias in label or (len(label) >= 4 and label in alias) for alias in wanted)):
            fuzzy.append(row)
    return fuzzy[0] if len(fuzzy) == 1 else {}

def _task_001_v24_row_text(row: dict[str, str]) -> str:
    return ' | '.join((str(value or '').strip() for value in row.values()))

def _task_001_v24_has_value(row: dict[str, str], *roles: str) -> bool:
    return any((_usable_docx_value(row.get(role)) for role in roles))

def _task_001_v24_bps_matches(value: Any, target: float) -> bool:
    text = str(value or '').translate(str.maketrans({'−': '-', '–': '-', '—': '-'}))
    tokens = re.findall('\\(?-?\\$?[0-9][0-9,]*(?:\\.[0-9]+)?(?:[kmb]|x|%)?\\)?', text, flags=re.I)
    normalized = _normalize(text)
    for token in tokens:
        value_number = _number(token)
        if value_number is None:
            continue
        if '%' in token:
            candidate = value_number * 10000
        elif re.search('\\b(?:pp|percentage points?|percent points?)\\b', normalized):
            candidate = value_number * 100
        else:
            candidate = value_number
        if abs(candidate - target) <= 0.15:
            return True
    return False

def _task_001_v24_zero(value: Any, *, allow_blank: bool=False) -> bool:
    text = str(value or '').strip()
    if not text:
        return allow_blank
    normalized = _normalize(text)
    if normalized in {'zero', 'nil', 'none', 'n a', 'not applicable', 'tied', 'ties', 'reconciled', 'reconciles', 'balanced', 'no variance', 'zero variance', 'pass', 'passes', 'ok', 'clear', 'no exceptions'}:
        return True
    tokens = re.findall('\\(?-?\\$?[0-9][0-9,]*(?:\\.[0-9]+)?%?\\)?', text)
    return bool(tokens) and all((_close(token, 0.0, abs_tol=0.02, rel_tol=0.0) for token in tokens))

def _task_001_v24_metadata(document: Document | None) -> dict[str, str]:
    if document is None:
        return {}
    result: dict[str, str] = {}
    for table in document.tables:
        if not table.rows or len(table.rows[0].cells) < 2:
            continue
        headers = [_normalize(cell.text) for cell in table.rows[0].cells]
        if headers[0] not in {'field', 'label', 'item'}:
            continue
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            for index in range(0, len(cells) - 1, 2):
                label = _normalize(cells[index])
                if label:
                    result[label] = cells[index + 1]
        if result:
            return result
    return result

def _task_001_v24_box(document: Document | None, *labels: str) -> str:
    if document is None:
        return ''
    wanted = {_normalize(label) for label in labels}
    for table in document.tables:
        if len(table.rows) < 2 or len(table.rows[0].cells) != 1:
            continue
        heading = _normalize(table.rows[0].cells[0].text)
        if any((label in heading or heading in label for label in wanted if label and heading)):
            return '\n'.join((row.cells[0].text.strip() for row in table.rows[1:])).strip()
    return ''

def _task_001_v24_artifact_review(document: Document | None, path: Path, parse_error: str) -> dict[str, Any]:
    tables: list[dict[str, Any]] = []
    paragraphs: list[str] = []
    if document is not None:
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()][:80]
        for (index, table) in enumerate(document.tables[:30], start=1):
            rows = [[cell.text.strip()[:500] for cell in row.cells] for row in table.rows[:80]]
            tables.append({'table': index, 'rows': rows})
    return {'version': 1, 'artifact': str(TASK_001_ARTIFACT), 'exists': path.is_file(), 'parse_error': parse_error, 'paragraphs': paragraphs, 'tables': tables}

def _grade_task_001_v24(workspace_root: Path, answer: Any) -> dict[str, Any]:
    """Grade the ARM-2409 memo without prescribing wording or hidden detail registers."""
    del answer
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
    gold = load_apex_gold('task_001')
    project = gold['project']
    prior = gold['prior_month']
    bridge_gold = gold['bridge']
    text = _document_text(document) if document is not None else ''
    normalized = _normalize(text)
    metadata = _task_001_v24_metadata(document)
    tables = {name: _task_001_v24_rows(document, name) for name in TASK_001_V24_TABLE_ALIASES}
    criteria: list[Criterion] = []
    semantic_specs: list[dict[str, Any]] = []

    def add(criterion_id: str, description: str, met: bool, evidence: str, *, category: str='core_finance', weight: int=10) -> None:
        criteria.append(Criterion(criterion_id, description, bool(met), evidence[:2000], category=category, weight=weight, semantic=False))

    def add_semantic(criterion_id: str, description: str, submitted_evidence: str, reference_context: Any, *, category: str='decision', weight: int=10, hard_gate: bool | None=None, evidence_scope: str='criterion_field') -> None:
        submitted = str(submitted_evidence or '').strip()
        gate = bool(submitted) if hard_gate is None else bool(hard_gate)
        criteria.append(Criterion(criterion_id, description, False, 'submitted field is non-empty' if gate else 'submitted field is empty', category=category, weight=weight, semantic=True))
        semantic_specs.append({'criterion_id': criterion_id, 'expected_facts': reference_context, 'reference_context': reference_context, 'task_context': {'project': 'ARM-2409 Northline Cold Storage Expansion', 'cutoff': '2026-06-30', 'work_product': 'preparer controller-sign-off memorandum'}, 'submitted_evidence': submitted[:12000], 'evidence_scope': evidence_scope, 'hard_gate_met': gate, 'hard_gate_evidence': 'the exact submitted field or row is present and non-empty' if gate else 'the required submitted field or row is blank or absent', 'always_judge': True})

    def row_for(table_name: str, label_role: str, aliases: tuple[str, ...]) -> dict[str, str]:
        return _task_001_v24_row(tables[table_name], label_role, aliases)

    def add_number(criterion_id: str, description: str, row: dict[str, str], roles: tuple[str, ...], expected: float, *, abs_tol: float=0.02) -> None:
        submitted = ' | '.join((str(row.get(role) or '') for role in roles))
        met = bool(row) and _text_contains_number(submitted, expected, abs_tol=abs_tol)
        add(criterion_id, description, met, f'submitted={submitted!r}; expected={expected}')

    def add_bps(criterion_id: str, description: str, row: dict[str, str], role: str, expected: float) -> None:
        submitted = str(row.get(role) or '')
        add(criterion_id, description, bool(row) and _task_001_v24_bps_matches(submitted, expected), f'submitted={submitted!r}; expected={expected} basis points')

    def add_zero(criterion_id: str, description: str, row: dict[str, str], role: str) -> None:
        submitted = str(row.get(role) or '')
        add(criterion_id, description, bool(row) and _task_001_v24_zero(submitted), f'submitted={submitted!r}; expected zero or an equivalent tied status', category='auditability', weight=5)
    add('artifact__working_document_present', 'The required editable working memorandum exists', path.is_file(), f'path={path}', category='structure', weight=1)
    add('artifact__document_parses', 'The working memorandum is a readable Word document', document is not None, parse_error or 'document parsed', category='structure', weight=1)
    add('artifact__project_id', 'The memorandum identifies ARM-2409', 'arm 2409' in normalized, 'project id present' if 'arm 2409' in normalized else 'project id absent', category='structure', weight=1)
    add('artifact__project_name', 'The memorandum identifies the Northline Cold Storage Expansion', 'northline cold storage expansion' in normalized, 'project name present' if 'northline cold storage expansion' in normalized else 'project name absent', category='structure', weight=1)
    add('artifact__june_30_cutoff', 'The memorandum identifies the June 30, 2026 close cutoff', _task_001_has_cutoff_date(text), 'cutoff found' if _task_001_has_cutoff_date(text) else 'cutoff not found', category='structure', weight=1)
    prepared_by = _task_001_metadata_value(metadata, 'Prepared by', 'Preparer', 'Prepared for')
    prepared_date = _task_001_metadata_value(metadata, 'Prepared date', 'Date prepared', 'Preparation date')
    review_status = _task_001_metadata_value(metadata, 'Review status', 'Status', 'Approval status')
    add('artifact__preparer', 'A preparer is identified', bool(prepared_by), f'prepared_by={prepared_by!r}', category='structure', weight=1)
    add('artifact__prepared_date', 'A preparation date is recorded', bool(prepared_date), f'prepared_date={prepared_date!r}', category='structure', weight=1)
    add_semantic('artifact__review_status', 'The review status makes clear that Controller review or approval is still pending', review_status, {'acceptable_meaning': 'pending Controller review, preparer draft, not approved, returned for review, or another unambiguous not-yet-approved status', 'unacceptable_meaning': 'Controller-approved or authorized-to-post status'}, category='structure', weight=1)
    placeholder_hits = re.findall('\\[(?:enter|select|complete|cite|record|explain|state|list)\\b', text, flags=re.I)
    add('artifact__no_instruction_placeholders', 'The completed memorandum contains no unresolved bracketed instruction prompts', document is not None and (not placeholder_hits), f'unresolved_instruction_prompts={len(placeholder_hits)}', category='structure', weight=1)
    prohibited_signature = bool(document and _document_has_prohibited_signature(document))
    add('artifact__no_simulated_controller_signature', 'The preparer did not simulate a Controller signature or approval', document is not None and (not prohibited_signature), f'prohibited_signature={prohibited_signature}', category='integrity', weight=10)
    close_specs = (('current_contract', ('Current contract', 'Executed contract', 'Contract value'), project['current_contract'], ('current', 'recommended')), ('posted_cost', ('Posted cost to date', 'Posted cost', 'Cost to date'), project['cost_to_date'], ('current', 'recommended')), ('billings', ('Billings to date', 'Billings', 'Cumulative billings'), project['billings'], ('current', 'recommended')), ('pm_etc', ('Submitted PM ETC', 'PM ETC', 'Submitted ETC'), gold['pm_etc'], ('current',)), ('finance_adjustment', ('Finance ETC adjustment', 'Finance adjustment', 'ETC correction'), gold['required_etc_adjustment'], ('change', 'recommended')), ('close_etc', ('Recommended close ETC', 'Close ETC', 'Adjusted ETC'), project['estimated_cost_to_complete'], ('recommended',)), ('eac', ('Estimated cost at completion', 'EAC', 'Total estimated cost'), project['estimated_cost_at_completion'], ('recommended',)), ('percent_complete', ('Percent complete', '% complete', 'POC'), project['percent_complete'], ('recommended',)), ('earned_revenue', ('Earned revenue', 'Revenue earned', 'Recognized revenue'), project['earned_revenue'], ('recommended',)), ('contract_asset', ('Contract asset', 'Underbilling', 'Costs in excess'), project['underbilling'], ('recommended',)), ('contract_liability', ('Contract liability', 'Overbilling', 'Billings in excess'), project['overbilling'], ('recommended',)), ('estimated_margin', ('Estimated margin at completion', 'Estimated margin', 'Total margin'), project['estimated_total_margin'], ('recommended',)), ('margin_rate', ('Margin rate', 'Margin percent', 'Gross margin rate'), project['estimated_margin_percent'], ('recommended',)))
    for (key, aliases, expected, roles) in close_specs:
        row = row_for('close', 'metric', aliases)
        add_number(f'close__{key}', f'The June close position reports the correct {aliases[0].lower()}', row, roles, float(expected))
    recommendation = _task_001_v24_box(document, 'Close recommendation and posting boundary', 'Recommendation and posting boundary', 'Close recommendation')
    add_semantic('close__recommendation', 'The recommendation restores the unsupported recovery to ETC, uses the recalculated June WIP, and holds posting for Controller disposition', recommendation, {'expected_recommendation': gold['signoff']['recommendation']})
    bridge_rows = {'current_contract': row_for('bridge', 'metric', ('Current contract', 'Executed contract', 'Contract value')), 'posted_cost': row_for('bridge', 'metric', ('Posted cost', 'Cost to date', 'Posted cost to date')), 'close_etc': row_for('bridge', 'metric', ('Remaining cost / ETC', 'Remaining cost', 'Close ETC', 'ETC')), 'eac': row_for('bridge', 'metric', ('Estimated cost at completion', 'EAC', 'Total estimated cost')), 'percent_complete': row_for('bridge', 'metric', ('Percent complete', '% complete', 'POC')), 'earned_revenue': row_for('bridge', 'metric', ('Earned revenue', 'Revenue earned', 'Recognized revenue')), 'billings': row_for('bridge', 'metric', ('Billings', 'Cumulative billings')), 'contract_asset': row_for('bridge', 'metric', ('Contract asset', 'Underbilling', 'Costs in excess')), 'contract_liability': row_for('bridge', 'metric', ('Contract liability', 'Overbilling', 'Billings in excess')), 'estimated_margin': row_for('bridge', 'metric', ('Estimated margin at completion', 'Estimated margin', 'Total margin')), 'margin_rate': row_for('bridge', 'metric', ('Margin rate', 'Margin percent', 'Gross margin rate'))}
    bridge_values = {'current_contract': (prior['current_contract'], project['current_contract'], bridge_gold['current_contract_change']), 'posted_cost': (prior['posted_cost'], project['cost_to_date'], bridge_gold['posted_cost_change']), 'close_etc': (prior['close_etc'], project['estimated_cost_to_complete'], bridge_gold['close_etc_change']), 'eac': (prior['eac'], project['estimated_cost_at_completion'], bridge_gold['eac_change']), 'percent_complete': (prior['percent_complete'], project['percent_complete'], bridge_gold['percent_complete_change_bps']), 'earned_revenue': (prior['earned_revenue'], project['earned_revenue'], bridge_gold['earned_revenue_change']), 'billings': (prior['billings'], project['billings'], bridge_gold['billings_change']), 'contract_asset': (prior['contract_asset'], project['underbilling'], bridge_gold['contract_asset_change']), 'contract_liability': (prior['contract_liability'], project['overbilling'], bridge_gold['contract_liability_change']), 'estimated_margin': (prior['estimated_margin'], project['estimated_total_margin'], bridge_gold['estimated_margin_change']), 'margin_rate': (prior['margin_percent'], project['estimated_margin_percent'], bridge_gold['margin_rate_change_bps'])}
    for (key, (may_value, june_value, change_value)) in bridge_values.items():
        row = bridge_rows[key]
        add_number(f'bridge__{key}__may', f"The bridge reports the correct May {key.replace('_', ' ')}", row, ('may',), float(may_value))
        add_number(f'bridge__{key}__june', f"The bridge reports the correct June {key.replace('_', ' ')}", row, ('june',), float(june_value))
        if key in {'percent_complete', 'margin_rate'}:
            add_bps(f'bridge__{key}__change', f"The bridge reports the correct basis-point change in {key.replace('_', ' ')}", row, 'change', float(change_value))
        else:
            add_number(f'bridge__{key}__change', f"The bridge reports the correct change in {key.replace('_', ' ')}", row, ('change',), float(change_value))
    driver_keys = ('posted_cost', 'close_etc', 'contract_asset', 'estimated_margin')
    driver_labels = {'posted_cost': 'Posted cost', 'close_etc': 'Remaining cost / ETC', 'contract_asset': 'Contract asset', 'estimated_margin': 'Estimated margin at completion'}
    for key in driver_keys:
        row = bridge_rows[key]
        submitted = str(row.get('driver') or '')
        add_semantic(f'bridge__{key}__driver', f'The bridge gives a substantively correct driver for {driver_labels[key]}', submitted, {'expected_driver': gold['signoff']['bridge_driver_explanations'][driver_labels[key]], 'metric': driver_labels[key]})
    for impact_row in gold['close_impact_bridge']['rows']:
        metric = str(impact_row['metric'])
        key = _normalize(metric).replace(' ', '_')
        aliases = {'ETC': ('ETC', 'Remaining cost', 'Close ETC'), 'EAC': ('EAC', 'Estimated cost at completion', 'Total estimated cost'), 'Percent complete': ('Percent complete', '% complete', 'POC'), 'Earned revenue': ('Earned revenue', 'Revenue earned', 'Recognized revenue'), 'Contract asset': ('Contract asset', 'Underbilling', 'Costs in excess'), 'Estimated margin': ('Estimated margin', 'Total margin', 'Estimated margin at completion'), 'Margin rate': ('Margin rate', 'Margin percent', 'Gross margin rate')}[metric]
        row = row_for('impact', 'metric', aliases)
        add_number(f'bridge__impact__{key}__submitted', f'The impact bridge reports the correct submitted-PM {metric.lower()}', row, ('submitted',), float(impact_row['pm_case']))
        add_number(f'bridge__impact__{key}__recommended', f'The impact bridge reports the correct recommended-close {metric.lower()}', row, ('recommended',), float(impact_row['policy_corrected']))
        if impact_row['unit'] == 'basis_points':
            add_bps(f'bridge__impact__{key}__difference', f'The impact bridge reports the correct basis-point effect on {metric.lower()}', row, 'impact', float(impact_row['impact']))
        else:
            add_number(f'bridge__impact__{key}__difference', f'The impact bridge reports the correct recommended-less-submitted effect on {metric.lower()}', row, ('impact',), float(impact_row['impact']))
    bridge_basis = _task_001_v24_box(document, 'Basis and authority of each case', 'Case basis and authority', 'Scenario basis')
    add_semantic('bridge__case_basis', 'The case basis distinguishes the submitted PM estimate from the policy-corrected preparer close and does not imply either is Controller-approved', bridge_basis, {'submitted_case_status': gold['close_impact_bridge']['pm_case_status'], 'recommended_case_status': gold['close_impact_bridge']['policy_corrected_status']})
    for source_row in gold['source_reconciliation']['rows']:
        key = str(source_row['key'])
        aliases = {'labor': ('Labor job cost', 'Labor cost', 'Labor'), 'material': ('Material job cost', 'Material cost', 'Materials'), 'subcontract': ('Subcontract job cost', 'Subcontract cost', 'Subcontract'), 'other': ('Other job cost', 'Other cost', 'Other'), 'total_job_cost': ('Total job cost', 'Job cost total', 'Total cost'), 'billings': ('Billings', 'Cumulative billings'), 'current_contract': ('Current contract', 'Executed contract', 'Contract value')}[key]
        row = row_for('reconciliation', 'control', aliases)
        for role in ('base', 'activity', 'close'):
            add_number(f'accounting__{key}__{role}', f"The accounting reconciliation reports the correct {role} value for {source_row['label']}", row, (role,), float(source_row[role]))
        add_zero(f'accounting__{key}__variance', f"The accounting reconciliation shows a zero variance or equivalent tied status for {source_row['label']}", row, 'variance')
    population_expectations = {'ap': {'aliases': ('AP', 'Accounts payable', 'Vendor invoices', 'AP job cost'), 'source_count': 8.0, 'source_amount': 107393.79, 'journal_count': 8.0, 'ledger_amount': 107393.79}, 'payroll': {'aliases': ('PAY', 'Payroll', 'Payroll job cost', 'Labor payroll'), 'source_count': 15.0, 'source_amount': 54498.34, 'journal_count': 5.0, 'ledger_amount': 54498.34}, 'billing': {'aliases': ('Billing', 'Billings', 'June billing', 'AR billing', 'Invoice'), 'source_count': 1.0, 'source_amount': 79513.55, 'journal_count': 1.0, 'ledger_amount': 79513.55}, 'job_cost_total': {'aliases': ('Total / control', 'Total control', 'Job cost total', 'Total job cost'), 'source_count': 23.0, 'source_amount': 161892.13, 'journal_count': 13.0, 'ledger_amount': 161892.13}}
    population_rows: dict[str, dict[str, str]] = {}
    for (key, expectation) in population_expectations.items():
        row = row_for('population', 'source', expectation['aliases'])
        population_rows[key] = row
        for role in ('source_count', 'source_amount', 'journal_count', 'ledger_amount'):
            add_number(f'population__{key}__{role}', f"The {key.replace('_', ' ')} population reports the correct {role.replace('_', ' ')}", row, (role,), float(expectation[role]))
        add_zero(f'population__{key}__variance', f"The {key.replace('_', ' ')} population ties to the ledger with zero variance", row, 'variance')
    population_conclusion = _task_001_v24_box(document, 'Population control conclusion', 'Posting completeness conclusion', 'Population conclusion')
    job_cost_evidence = '\n'.join([population_conclusion, _task_001_v24_row_text(population_rows['ap']), _task_001_v24_row_text(population_rows['payroll']), _task_001_v24_row_text(population_rows['job_cost_total'])]).strip()
    billing_evidence = '\n'.join([population_conclusion, _task_001_v24_row_text(population_rows['billing'])]).strip()
    add_semantic('population__job_cost_conclusion', 'The job-cost conclusion confirms a complete posted AP and payroll population tied to 23 entries, 13 journals, and $161,892.13, or clearly identifies any exception', job_cost_evidence, {'expected_control': '8 AP entries totaling $107,393.79 and 15 payroll entries totaling $54,498.34 reconcile to 13 posted journals and $161,892.13 with no exceptions'}, hard_gate=bool(population_conclusion) or _task_001_v24_has_value(population_rows['job_cost_total'], 'exceptions'))
    add_semantic('population__billing_conclusion', 'The billing conclusion confirms the complete non-voided June billing ties to the posted ledger and treats payment status as a collection matter, or clearly identifies an exception', billing_evidence, {'expected_control': 'one June invoice totaling $79,513.55 ties to the posted billing journal; partial payment does not reduce recognized non-voided billing activity'}, hard_gate=bool(population_conclusion) or _task_001_v24_has_value(population_rows['billing'], 'exceptions'))
    etc_component_specs = (('remaining_field_labor', ('Remaining field labor', 'Field labor', 'Labor'), 210240.0), ('material_equipment', ('Material / equipment', 'Material and equipment', 'Materials / equipment'), 181040.0), ('subcontract', ('Subcontract', 'Subcontracts'), 105120.0), ('commissioning_closeout', ('Commissioning / close-out', 'Commissioning and close-out', 'Close-out'), 52560.0), ('contingency', ('Contingency', 'Remaining contingency'), 35040.0))
    for (key, aliases, expected) in etc_component_specs:
        row = row_for('etc', 'component', aliases)
        add_number(f'commercial__etc__{key}__submitted', f'The ETC schedule reports the submitted-PM {aliases[0].lower()} component', row, ('pm',), expected)
    recovery_row = row_for('etc', 'component', ('Commercial recovery / cost offset', 'Commercial recovery', 'Recovery / cost offset', 'PCO-011 recovery', 'Disputed recovery'))
    add_number('commercial__etc__recovery_adjustment', 'The ETC schedule restores the supported $185,000 commercial recovery or cost offset', recovery_row, ('correction',), 185000.0)
    total_etc_row = row_for('etc', 'component', ('Total / control', 'Total control', 'Total ETC', 'ETC total'))
    for (role, expected) in (('pm', 584000.0), ('correction', 185000.0), ('close', 769000.0)):
        add_number(f'commercial__etc__total__{role}', f"The ETC schedule reports the correct total {role.replace('_', ' ')}", total_etc_row, (role,), expected)
    add_semantic('commercial__etc__recovery_treatment', 'The ETC schedule explains that the disputed recovery is restored to remaining cost and is not treated as authorized recovery revenue or a contract-value increase', _task_001_v24_row_text(recovery_row), {'expected_treatment': gold['etc_composition_schedule']['basis']}, hard_gate=_task_001_v24_has_value(recovery_row, 'source'))
    judgment_specs = (('authorization', ('Authorization at cutoff', 'PCO-011 authorization', 'Commercial authorization'), 'PCO-011 authorization at cutoff'), ('contract_treatment', ('Contract value treatment', 'Contract treatment', 'Revenue treatment'), 'Contract value treatment'), ('etc_treatment', ('ETC treatment', 'PM ETC treatment', 'Cost treatment'), 'PM ETC treatment'), ('controller_review', ('Controller review trigger', 'Controller review', 'Review trigger'), 'Controller review'), ('posting_status', ('Close and posting status', 'Close / posting status', 'Posting status'), 'Close and posting status'))
    conclusions = gold['signoff']['conclusions']
    for (key, aliases, gold_key) in judgment_specs:
        row = row_for('judgment', 'question', aliases)
        submitted = _task_001_v24_row_text(row)
        add_semantic(f'commercial__judgment__{key}', f'The memorandum reaches the correct, supported conclusion for {aliases[0].lower()}', submitted, {'expected_conclusion': conclusions[gold_key], 'expected_support': conclusions[f'{gold_key} - support']}, hard_gate=_task_001_v24_has_value(row, 'conclusion', 'evidence', 'clearance'))
    posting_rows: list[dict[str, str]] = []
    carried_stage = ''
    for raw_row in tables['posting']:
        row = dict(raw_row)
        if _usable_docx_value(row.get('stage')):
            carried_stage = str(row['stage'])
        row['effective_stage'] = carried_stage
        posting_rows.append(row)
    stage_aliases = {'may_reversal': ('Reverse May close', 'Reverse prior close', 'May reversal', 'Prior-close reversal'), 'june_establishment': ('Establish June close', 'Record June close', 'June establishment', 'Establish recommended close'), 'net_june': ('Net June posting', 'Net June entry', 'June net posting', 'Net change')}
    account_aliases = {'account_1200': ('1200', 'Costs and earnings in excess', 'Contract asset'), 'account_2100': ('2100', 'Billings in excess', 'Contract liability'), 'account_4300': ('4300', 'WIP revenue adjustment', 'Revenue adjustment')}

    def posting_row(stage_key: str, account_key: str) -> dict[str, str]:
        wanted_stages = {_normalize(value) for value in stage_aliases[stage_key]}
        wanted_accounts = {_normalize(value) for value in account_aliases[account_key]}
        matches = []
        for row in posting_rows:
            stage_text = _normalize(row.get('effective_stage'))
            account_text = _normalize(row.get('account'))
            stage_met = stage_text in wanted_stages or any((value in stage_text for value in wanted_stages if len(value) >= 4))
            account_met = any((value in account_text for value in wanted_accounts))
            if stage_met and account_met:
                matches.append(row)
        return matches[0] if len(matches) == 1 else {}
    for stage in gold['proposed_journal_bridge']['stages']:
        stage_key = str(stage['stage_key'])
        stage_basis_parts = []
        for journal_row in stage['rows']:
            account_key = str(journal_row['account_key'])
            row = posting_row(stage_key, account_key)
            debit = str(row.get('debit') or '')
            credit = str(row.get('credit') or '')
            expected_debit = float(journal_row['debit'])
            expected_credit = float(journal_row['credit'])
            debit_met = _task_001_v24_zero(debit, allow_blank=True) if expected_debit == 0 else _text_contains_number(debit, expected_debit)
            credit_met = _task_001_v24_zero(credit, allow_blank=True) if expected_credit == 0 else _text_contains_number(credit, expected_credit)
            add(f'journal__{stage_key}__{account_key}', f"The {stage['stage']} row for {journal_row['account']} has the correct debit and credit", bool(row) and debit_met and credit_met, f"account={row.get('account')!r}; debit={debit!r}; credit={credit!r}; expected_debit={expected_debit}; expected_credit={expected_credit}")
            if row.get('basis'):
                stage_basis_parts.append(str(row['basis']))
        add_semantic(f'journal__{stage_key}__basis', f"The {stage['stage']} basis identifies the correct close balance or net-period relationship", '\n'.join(stage_basis_parts), {'expected_stage_basis': stage['basis']}, hard_gate=bool(stage_basis_parts))
    posting_status = _task_001_v24_box(document, 'Posting status, approval boundary, and tie to the May-to-June change', 'Posting status and approval boundary', 'Posting boundary')
    add_semantic('journal__posting_status', 'The posting bridge is clearly proposed and unposted, pending documented Controller approval, and tied to the May-to-June contract-asset change', posting_status, {'expected_status': gold['proposed_journal_bridge']['status'], 'expected_net_effect': 'credit account 1200 and debit account 4300 for $68,625.38, equal to the May-to-June contract-asset decrease; account 2100 remains nil'})
    source_specs = (('accounting', ('Current accounting records', 'Accounting records', 'Current accounting balances'), 0), ('may_close', ('Final May close', 'Prior close', 'May close'), 1), ('pm_forecast', ('Current PM forecast', 'PM forecast', 'PM ETC forecast'), 2), ('commercial', ('Commercial log / support', 'Commercial log and support', 'Commercial evidence', 'PCO support'), 3), ('policy', ('Signed accounting policy', 'WIP policy', 'Revenue recognition policy'), 4))
    evidence_log = gold['signoff']['evidence_log']
    for (key, aliases, gold_index) in source_specs:
        row = row_for('evidence', 'role', aliases)
        add_semantic(f'source__{key}', f'The evidence log identifies the controlling {aliases[0].lower()} and accurately states the fact or conclusion used', _task_001_v24_row_text(row), {'expected_source_control': evidence_log[gold_index]}, category='provenance', weight=3, hard_gate=_task_001_v24_has_value(row, 'version', 'authority', 'fact'))
    action_specs = (('finance', ('Finance / Project Accounting', 'Finance', 'Project Accounting'), 0), ('commercial', ('Commercial / Project Team', 'Commercial', 'Project Team'), 1), ('controller', ('Controller', 'Corporate Controller'), 2))
    expected_actions = gold['signoff']['actions']
    for (key, aliases, gold_index) in action_specs:
        row = row_for('action', 'owner', aliases)
        add_semantic(f'action__{key}', f'The {aliases[0]} action is specific, assigned, supported by completion evidence, and has an open or pending status consistent with the workpaper boundary', _task_001_v24_row_text(row), {'expected_action': expected_actions[gold_index]}, hard_gate=_task_001_v24_has_value(row, 'action', 'evidence', 'status'))
    result = _result(criteria)
    result['artifact_review'] = _task_001_v24_artifact_review(document, path, parse_error)
    result = _attach_semantic_review(result, task_id='task_001', evidence=_legacy_artifact_evidence(path) if document is not None else '', artifact_type='controller sign-off memorandum', specs=semantic_specs, decision_failure_cap=None, always_judge=True, execution_mode='scoped_per_criterion')
    result['task_grading_revision'] = dict(TASK_GRADING_REVISIONS['task_001'])
    return result

def grade_apex_task(task_id: str, answer: Any, workspace_root: str | Path) -> dict[str, Any]:
    if task_id == 'task_001':
        result = _grade_task_001_v24(Path(workspace_root), answer)
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
