from __future__ import annotations
import copy
import re
from typing import Any, Iterable, Mapping
REWARD_SCHEMA_VERSION = 7
ALLOWED_WEIGHTS = frozenset({1, 3, 5, 10})
DECISION_ACCURACY_CRITERIA = {'task_004': ('review__arm-2318__commercial_treatment', 'review__arm-2409__commercial_treatment', 'review__arm-2417__commercial_treatment', 'review__arm-2506__commercial_treatment', 'decision_bridge__arm-2318__close_disposition', 'decision_bridge__arm-2409__close_disposition', 'decision_bridge__arm-2417__close_disposition', 'decision_bridge__arm-2506__close_disposition', 'decision_control__pending_revenue', 'decision_control__review_queue', 'decision_control__close_release')}
SECTION_NORMALIZED_REWARD_POLICIES = {'task_037': {'sections': ({'name': 'artifact_sources_and_controls', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.15}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.1}, {'name': 'project_underwriting', 'prefixes': tuple((f'headline_values__cp_{project:02d}_{suffix}' for project in range(1, 9) for suffix in ('initial_capex', 'technician_capacity', 'mandatory_safety_flag', 'npv', 'downside_npv', 'irr', 'payback_years', 'selected', 'year_'))), 'weight': 0.25}, {'name': 'base_portfolio_decision', 'prefixes': tuple((f'headline_values__{label}' for label in ('selected_portfolio', 'selected_capex', 'portfolio_npv', 'downside_portfolio_npv', 'selected_debt_eligible_basis', 'selected_cash_funding', 'selected_technician_capacity', 'cash_capex_headroom', 'debt_capacity_headroom', 'technician_capacity_headroom', 'mandatory_safety_projects_selected', 'cash_constraint_check', 'debt_constraint_check', 'technician_constraint_check'))) + ('selection_tie',), 'weight': 0.25}, {'name': 'resilience_and_committee_release', 'prefixes': tuple((f'headline_values__{prefix}' for prefix in ('cp_03_unavailable_', 'cp_04_unavailable_', 'cash_capacity_down_20_percent_', 'debt_capacity_down_25_percent_', 'technician_capacity_down_25_percent_', 'downside_npv_objective_'))), 'weight': 0.25})}, 'task_035': {'sections': ({'name': 'artifact_and_sources', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.05}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.08}, {'name': 'probability_plan', 'prefixes': tuple((f'headline_values__{prefix}' for prefix in ('signed_backlog', 'fy27_', 'constrained_', 'first_constrained_', 'maximum_capacity_', 'probability_plan_', 'controlling_capacity_case', '2026_', '2027_'))), 'weight': 0.1}, {'name': 'gross_commitment_stress', 'prefixes': ('headline_values__gross_commitment_',), 'weight': 0.12}, {'name': 'execution_portfolio', 'prefixes': ('headline_values__execution_',), 'weight': 0.15}, {'name': 'earnings_release_decision', 'prefixes': ('headline_values__release_bridge_',), 'weight': 0.2}, {'name': 'executive_recovery_decision', 'prefixes': ('headline_values__executive_recovery_',), 'weight': 0.3})}, 'task_027': {'sections': ({'name': 'artifact_and_sources', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.05}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.08}, {'name': 'scenario_financials', 'prefixes': ('base_scenario_values__', 'downside_scenario_values__', 'upside_scenario_values__', 'scenario_comparison_values__'), 'weight': 0.14}, {'name': 'quarterly_liquidity', 'prefixes': ('base_liquidity_values__', 'downside_liquidity_values__', 'upside_liquidity_values__'), 'weight': 0.12}, {'name': 'scenario_attribution', 'prefixes': ('downside_scenario_attribution_values__', 'upside_scenario_attribution_values__'), 'weight': 0.1}, {'name': 'contingency_selection', 'prefixes': tuple((f'downside_contingency_values__{prefix}' for prefix in ('contingency_capex_deferral_', 'contingency_collections_acceleration_', 'contingency_construction_pricing_', 'contingency_controls_mix_', 'contingency_fixed_opex_freeze_', 'contingency_procurement_savings_', 'contingency_service_productivity_', 'contingency_targeted_overhead_reduction_', 'contingency_selected_', 'contingency_annual_', 'contingency_implementation_', 'contingency_adjusted_', 'contingency_ebitda_margin_headroom', 'contingency_free_cash_flow_headroom'))), 'weight': 0.11}, {'name': 'weekly_cash_support', 'prefixes': ('downside_contingency_values__contingency_week_', 'downside_contingency_values__contingency_13_week_'), 'weight': 0.1}, {'name': 'lender_release_decision', 'prefixes': ('covenant_release_values__',), 'weight': 0.3})}}

def criterion_policy(criterion: Mapping[str, Any]) -> dict[str, Any]:
    weight = int(criterion.get('weight', 10))
    if weight not in ALLOWED_WEIGHTS:
        raise ValueError(f"invalid criterion weight {weight}: {criterion.get('id')}")
    return {'category': str(criterion.get('category') or 'core_finance'), 'weight': weight, 'failure_cap': None if criterion.get('failure_cap') is None else float(criterion['failure_cap']), 'semantic': bool(criterion.get('semantic'))}

def attach_default_policy(result: Mapping[str, Any]) -> dict[str, Any]:
    """Classify legacy hand-written criteria that predate schema-v3 gold."""
    updated = copy.deepcopy(dict(result))
    for row in updated.get('criteria', []):
        if not isinstance(row, dict) or 'weight' in row:
            continue
        text = f"{row.get('id', '')} {row.get('description', '')}".casefold()
        semantic = bool(re.search('status|compliance|classification|timing|decision|recommend|title|required item|complete set|professional conclusion', text))
        if 'preserv' in text:
            (category, weight, cap) = ('integrity', 10, 0.0)
        elif re.search('formula|lineage|audit', text):
            (category, weight, cap) = ('auditability', 5, 0.69)
        elif re.search('source|evidence', text):
            (category, weight, cap) = ('provenance', 3, None)
        elif re.search('chart|format|structure|slide|table|length|page|identity', text):
            (category, weight, cap) = ('structure', 1, None)
        elif semantic:
            (category, weight, cap) = ('decision', 10, 0.49)
        else:
            (category, weight, cap) = ('core_finance', 10, None)
        row.update({'category': category, 'weight': weight, 'semantic': semantic})
        if cap is not None:
            row['failure_cap'] = cap
    return updated

def apply_reward_policy(task_id: str, result: Mapping[str, Any], *, integrity: Mapping[str, Any] | None=None) -> dict[str, Any]:
    """Compute a monotonic training reward and independent quality gates.

    Partial credit must preserve the professional ordering of two valid but
    incomplete submissions.  Non-zero legacy caps therefore remain visible as
    critical quality-gate metadata and strict-pass blockers, but they do not
    flatten distinct weighted scores.  Only declared zero-reward criteria and
    environment-integrity failures invalidate the scalar reward.
    """
    updated = copy.deepcopy(dict(result))
    criteria = [row for row in updated.get('criteria', []) if isinstance(row, dict)]
    total_weight = 0
    earned_weight = 0
    failed_declared_gates: list[tuple[float, str]] = []
    for row in criteria:
        policy = criterion_policy(row)
        row.update(policy)
        value = int(bool(row.get('value')))
        row['value'] = value
        total_weight += policy['weight']
        earned_weight += policy['weight'] * value
        if not value and policy['failure_cap'] is not None:
            failed_declared_gates.append((float(policy['failure_cap']), str(row.get('id'))))
    raw_reward = earned_weight / total_weight if total_weight else 0.0
    reward = raw_reward
    decision_accuracy_adjustment = None
    section_normalized_reward = None
    core_model_failure = None
    critical_ids = DECISION_ACCURACY_CRITERIA.get(task_id)
    if critical_ids:
        by_id = {str(row.get('id')): row for row in criteria}
        missing_ids = [criterion_id for criterion_id in critical_ids if criterion_id not in by_id]
        if missing_ids:
            raise ValueError(f'missing decision-accuracy criteria for {task_id}: {missing_ids}')
        critical_rows = [by_id[criterion_id] for criterion_id in critical_ids]
        critical_met = sum((int(bool(row.get('value'))) for row in critical_rows))
        critical_weight_total = sum((int(row.get('weight', 10)) for row in critical_rows))
        critical_weight_earned = sum((int(row.get('weight', 10)) * int(bool(row.get('value'))) for row in critical_rows))
        supporting_weight_total = total_weight - critical_weight_total
        supporting_weight_earned = earned_weight - critical_weight_earned
        if supporting_weight_total <= 0:
            raise ValueError(f'missing supporting criteria for {task_id}')
        supporting_accuracy = supporting_weight_earned / supporting_weight_total
        support_payload = updated.get('decision_support')
        support_by_id = support_payload.get('criteria', {}) if isinstance(support_payload, Mapping) else {}
        decision_quality_scores: dict[str, float] = {}
        for row in critical_rows:
            criterion_id = str(row.get('id'))
            if bool(row.get('value')):
                score = 1.0
            else:
                support = support_by_id.get(criterion_id, {})
                score = float(support.get('score', 0.0)) if isinstance(support, Mapping) else 0.0
                if not 0.0 <= score <= 0.5:
                    raise ValueError(f'invalid non-passing decision-support score for {criterion_id}: {score}')
            decision_quality_scores[criterion_id] = score
        decision_accuracy = critical_met / len(critical_ids)
        decision_support_quality = sum(decision_quality_scores.values()) / len(critical_ids)
        maximum_reward_factor = 0.25 + 0.75 * decision_support_quality
        decision_adjusted_reward = supporting_accuracy * maximum_reward_factor
        reward = min(raw_reward, decision_adjusted_reward)
        decision_accuracy_adjustment = {'method': 'cap_by_supporting_accuracy_and_decision_quality', 'criteria_ids': list(critical_ids), 'criteria_met': critical_met, 'criteria_total': len(critical_ids), 'decision_accuracy': round(decision_accuracy, 6), 'decision_support_quality': round(decision_support_quality, 6), 'decision_quality_scores': {criterion_id: round(score, 6) for (criterion_id, score) in decision_quality_scores.items()}, 'supporting_weight_earned': supporting_weight_earned, 'supporting_weight_total': supporting_weight_total, 'supporting_accuracy': round(supporting_accuracy, 6), 'maximum_reward_factor': round(maximum_reward_factor, 6), 'decision_adjusted_reward': round(decision_adjusted_reward, 6)}
    section_policy = SECTION_NORMALIZED_REWARD_POLICIES.get(task_id)
    if section_policy:
        sections = tuple(section_policy.get('sections', ()))
        if not sections:
            raise ValueError(f'missing section-normalized policy for {task_id}')
        section_results: dict[str, dict[str, Any]] = {}
        matched_ids: set[str] = set()
        normalized_reward = 0.0
        declared_weight = 0.0
        for section in sections:
            name = str(section['name'])
            prefixes = tuple((str(prefix) for prefix in section['prefixes']))
            rows = [row for row in criteria if str(row.get('id', '')).startswith(prefixes)]
            if not rows:
                raise ValueError(f'missing section-normalized criteria for {task_id}: {name}')
            row_ids = {str(row['id']) for row in rows}
            if matched_ids.intersection(row_ids):
                raise ValueError(f'overlapping section-normalized criteria for {task_id}: {name}')
            matched_ids.update(row_ids)
            total = sum((int(row['weight']) for row in rows))
            earned = sum((int(row['weight']) * int(bool(row.get('value'))) for row in rows))
            accuracy = earned / total
            section_weight = float(section['weight'])
            declared_weight += section_weight
            normalized_reward += section_weight * accuracy
            section_results[name] = {'prefixes': list(prefixes), 'section_weight': section_weight, 'weight_earned': earned, 'weight_total': total, 'accuracy': round(accuracy, 6)}
        all_ids = {str(row['id']) for row in criteria}
        if matched_ids != all_ids:
            raise ValueError(f'unassigned section-normalized criteria for {task_id}: {sorted(all_ids - matched_ids)}')
        if abs(declared_weight - 1.0) > 1e-09:
            raise ValueError(f'invalid section weights for {task_id}')
        reward = normalized_reward
        section_normalized_reward = {'method': 'linear_multi_section_normalization', 'sections': section_results, 'section_normalized_reward': round(reward, 6)}
    if task_id == 'task_037':
        by_id = {str(row.get('id')): row for row in criteria}
        formula_rows = [row for row in criteria if str(row.get('id', '')).startswith('formula_lineage__')]
        governing_decision_ids = ('headline_values__selected_portfolio', 'headline_values__portfolio_npv', 'headline_values__downside_portfolio_npv')
        workbook_integrity_id = 'model_integrity__no_errors'
        missing_decision_ids = [criterion_id for criterion_id in (*governing_decision_ids, workbook_integrity_id) if criterion_id not in by_id]
        if missing_decision_ids:
            raise ValueError(f'missing Task037 governing-decision criteria: {missing_decision_ids}')
        if formula_rows and (not any((bool(row.get('value')) for row in formula_rows))) and (not bool(by_id[workbook_integrity_id].get('value')) or not any((bool(by_id[criterion_id].get('value')) for criterion_id in governing_decision_ids))):
            core_model_failure = {'code': 'nonfunctional_or_missing_governing_portfolio_model', 'formula_criteria_met': 0, 'formula_criteria_total': len(formula_rows), 'governing_decision_ids': list(governing_decision_ids), 'governing_decisions_met': sum((bool(by_id[criterion_id].get('value')) for criterion_id in governing_decision_ids)), 'workbook_integrity_met': bool(by_id[workbook_integrity_id].get('value')), 'reward_effect': 'zero_reward_hard_failure'}
    required_workbook_by_task = {'task_027': 'Shared/Finance/FP&A/FY27 plan/FY27 EBITDA scenarios - WORKING.xlsx', 'task_035': 'Shared/Finance/FP&A/Backlog/FY27 backlog burn and capacity - WORKING.xlsx'}
    if task_id in required_workbook_by_task and integrity is not None:
        required_workbook = required_workbook_by_task[task_id]
        declared_artifact = str(integrity.get('required_artifact') or '')
        changed_files = {str(path) for path in integrity.get('changed_files', []) if path is not None}
        created_files = {str(path) for path in integrity.get('created_files', []) if path is not None}
        if declared_artifact == required_workbook and required_workbook not in changed_files | created_files:
            core_model_failure = {'code': 'missing_required_workbook_modification', 'required_artifact': required_workbook, 'changed_files_count': len(changed_files), 'created_files_count': len(created_files), 'reward_effect': 'zero_reward_hard_failure'}
    hard_failures: list[dict[str, Any]] = []
    if integrity is not None:
        updated['integrity'] = copy.deepcopy(dict(integrity))
        hard_failures = [dict(row) for row in integrity.get('hard_failures', []) if isinstance(row, Mapping)]
        if hard_failures:
            reward = 0.0
    applied_caps = []
    quality_gate_failures = []
    if core_model_failure is not None:
        reward = 0.0
        applied_caps.append({'criterion_id': core_model_failure['code'], 'cap': 0.0})
        quality_gate_failures.append(dict(core_model_failure))
    for (cap, criterion_id) in sorted(failed_declared_gates):
        gate = {'criterion_id': criterion_id, 'declared_threshold': cap}
        if cap == 0.0:
            reward = 0.0
            applied_caps.append({'criterion_id': criterion_id, 'cap': cap})
            gate['reward_effect'] = 'zero_reward_hard_failure'
        else:
            gate['reward_effect'] = 'strict_pass_blocker_only'
        quality_gate_failures.append(gate)
    met_count = sum((int(bool(row.get('value'))) for row in criteria))
    updated.update({'reward': round(reward, 6), 'raw_weighted_reward': round(raw_reward, 6), 'decision_accuracy_adjustment': decision_accuracy_adjustment, 'section_normalized_reward': section_normalized_reward, 'core_model_failure': core_model_failure, 'strict_pass': bool(criteria) and met_count == len(criteria) and (not hard_failures), 'criteria_met': met_count, 'criteria_total': len(criteria), 'weight_earned': earned_weight, 'weight_total': total_weight, 'reward_schema_version': REWARD_SCHEMA_VERSION, 'reward_definition': 'weighted binary criteria using 1/3/5/10 importance; task-declared linear section normalization keeps repeated support schedules proportional to separately weighted management judgment; Task037 submissions that contain neither a formula-driven model nor the governing portfolio decision receive zero; Tasks027, 031, 035, and 041 submissions that do not modify or create their required working workbook receive zero; Tasks079 and 080 blank/default responses with no substantive finance work receive zero; decision-quality caps preserve supporting-work credit, distinguish correct schedule support from omitted facts, and prevent wrong critical business outputs from being masked; non-zero declared quality thresholds block strict pass without clipping partial credit; declared zero-reward criteria and environment-integrity failures score zero', 'applied_reward_caps': applied_caps, 'quality_gate_failures': quality_gate_failures, 'hard_failures': hard_failures})
    return updated
