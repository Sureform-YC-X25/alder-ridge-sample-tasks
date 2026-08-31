from __future__ import annotations
import copy
import re
from typing import Any, Iterable, Mapping
REWARD_SCHEMA_VERSION = 8
ALLOWED_WEIGHTS = frozenset({1, 3, 5, 10})
DECISION_ACCURACY_CRITERIA = {}
SECTION_NORMALIZED_REWARD_POLICIES = {'task_001': {'sections': ({'name': 'artifact_identity_and_structure', 'prefixes': ('artifact__',), 'weight': 0.04}, {'name': 'recommended_june_close', 'prefixes': ('close__',), 'weight': 0.25}, {'name': 'close_bridges_and_drivers', 'prefixes': ('bridge__',), 'weight': 0.15}, {'name': 'accounting_and_population_controls', 'prefixes': ('accounting__', 'population__'), 'weight': 0.18}, {'name': 'etc_commercial_and_policy_judgment', 'prefixes': ('commercial__',), 'weight': 0.17}, {'name': 'proposed_wip_posting_bridge', 'prefixes': ('journal__',), 'weight': 0.13}, {'name': 'source_traceability_and_actions', 'prefixes': ('source__', 'action__'), 'weight': 0.08})}, 'task_004': {'sections': ({'name': 'artifact_preservation', 'prefixes': ('preservation__',), 'weight': 0.03}, {'name': 'source_inputs_and_provenance', 'prefixes': ('source__',), 'weight': 0.08}, {'name': 'wip_finance_and_controls', 'prefixes': ('formula_lineage__', 'result__', 'total_formula_lineage', 'total_result__'), 'weight': 0.22}, {'name': 'commercial_review_and_release_decision', 'prefixes': ('review__', 'decision_bridge__', 'decision_control__'), 'weight': 0.28}, {'name': 'unbooked_commercial_sensitivity', 'prefixes': ('commercial_sensitivity__',), 'weight': 0.14}, {'name': 'accounting_to_workpaper_control', 'prefixes': ('accounting_control__',), 'weight': 0.25})}, 'task_015': {'sections': ({'name': 'artifact_and_presentation', 'criterion_ids': ('structure__title', 'structure__table'), 'prefixes': ('preservation__', 'style__'), 'weight': 0.08}, {'name': 'covenant_reperformance', 'prefixes': ('metric__',), 'weight': 0.25}, {'name': 'compliance_and_binding_decision', 'prefixes': ('status__', 'headroom__'), 'weight': 0.1}, {'name': 'calculation_support', 'prefixes': ('calculation_support__',), 'weight': 0.27}, {'name': 'downside_capacity', 'criterion_ids': ('structure__wip_reversal_sensitivity',), 'prefixes': ('sensitivity__',), 'weight': 0.25}, {'name': 'source_authority', 'prefixes': ('source__',), 'weight': 0.05})}, 'task_068': {'sections': ({'name': 'artifact_sources_and_controls', 'prefixes': ('preservation__', 'sources__', 'controls__'), 'weight': 0.08}, {'name': 'accounting_and_reporting_tie', 'prefixes': tuple((f'headline_values__{label}' for label in ('june_revenue', 'june_ebitda', 'posted_ytd_revenue', 'q2_posted_ytd_revenue_share', 'q2_posted_ytd_revenue_reconciliation_check', 'q2_revenue', 'q2_approved_plan_revenue', 'q2_revenue_variance_to_plan', 'q2_organic_growth', 'q2_adjusted_ebitda', 'q2_approved_plan_adjusted_ebitda', 'q2_adjusted_ebitda_variance_to_plan', 'construction_q2_revenue', 'construction_q2_approved_plan_revenue', 'construction_q2_revenue_variance_to_plan', 'construction_q2_adjusted_ebitda', 'construction_q2_approved_plan_adjusted_ebitda', 'construction_q2_adjusted_ebitda_variance_to_plan', 'service_q2_revenue', 'service_q2_approved_plan_revenue', 'service_q2_revenue_variance_to_plan', 'service_q2_adjusted_ebitda', 'service_q2_approved_plan_adjusted_ebitda', 'service_q2_adjusted_ebitda_variance_to_plan', 'controls_q2_revenue', 'controls_q2_approved_plan_revenue', 'controls_q2_revenue_variance_to_plan', 'controls_q2_adjusted_ebitda', 'controls_q2_approved_plan_adjusted_ebitda', 'controls_q2_adjusted_ebitda_variance_to_plan', 'q2_operating_cash_flow', 'q2_capital_expenditure', 'q2_free_cash_flow', 'q2_free_cash_flow_plan', 'q2_free_cash_flow_variance_to_plan', 'collections_timing_cash_impact', 'q2_unrestricted_cash', 'maximum_revolver', 'latest_full_year_revenue_outlook', 'signed_backlog', 'remaining_fy26_revenue_outlook', 'signed_backlog_coverage_of_remaining_outlook', 'ltm_adjusted_ebitda', 'funded_debt', 'lender_leverage', 'fixed_charge_coverage', 'revenue_guidance_low', 'revenue_guidance_high', 'ebitda_guidance_low', 'ebitda_guidance_high'))), 'weight': 0.4}, {'name': 'mitigation_and_release_decision', 'prefixes': tuple((f'headline_values__{label}' for label in ('identified_ebitda_action_pool', 'approved_mitigation_conversion', 'executable_ebitda_mitigation', 'post_mitigation_combined_stress_ebitda', 'post_mitigation_headroom_to_guidance_low', 'minimum_guidance_update_headroom', 'guidance_update_trigger', 'guidance_release_status', 'construction_gross_profit_impact', 'service_labor_productivity_impact', 'controls_mix_impact', 'largest_downside_driver', 'combined_stress_ebitda', 'ebitda_shortfall_to_guidance_low', 'additional_revenue_decline_to_update_trigger', 'guidance_update_required'))), 'weight': 0.4}, {'name': 'executive_narrative', 'prefixes': ('narrative__',), 'weight': 0.12})}, 'task_037': {'sections': ({'name': 'artifact_sources_and_controls', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.15}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.1}, {'name': 'project_underwriting', 'prefixes': tuple((f'headline_values__cp_{project:02d}_{suffix}' for project in range(1, 9) for suffix in ('initial_capex', 'technician_capacity', 'mandatory_safety_flag', 'npv', 'downside_npv', 'irr', 'payback_years', 'selected', 'year_'))), 'weight': 0.25}, {'name': 'base_portfolio_decision', 'prefixes': tuple((f'headline_values__{label}' for label in ('selected_portfolio', 'selected_capex', 'portfolio_npv', 'downside_portfolio_npv', 'selected_debt_eligible_basis', 'selected_cash_funding', 'selected_technician_capacity', 'cash_capex_headroom', 'debt_capacity_headroom', 'technician_capacity_headroom', 'mandatory_safety_projects_selected', 'cash_constraint_check', 'debt_constraint_check', 'technician_constraint_check'))) + ('selection_tie',), 'weight': 0.25}, {'name': 'resilience_and_committee_release', 'prefixes': tuple((f'headline_values__{prefix}' for prefix in ('cp_03_unavailable_', 'cp_04_unavailable_', 'cash_capacity_down_20_percent_', 'debt_capacity_down_25_percent_', 'technician_capacity_down_25_percent_', 'downside_npv_objective_'))), 'weight': 0.25})}, 'task_035': {'sections': ({'name': 'artifact_and_sources', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.08}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.1}, {'name': 'probability_plan', 'prefixes': tuple((f'headline_values__{prefix}' for prefix in ('signed_backlog', 'fy27_', 'constrained_', 'first_constrained_', 'maximum_capacity_', 'probability_plan_', 'controlling_capacity_case'))), 'weight': 0.14}, {'name': 'gross_commitment_stress', 'prefixes': ('headline_values__gross_commitment_',), 'weight': 0.18}, {'name': 'execution_portfolio', 'prefixes': ('headline_values__execution_',), 'weight': 0.2}, {'name': 'earnings_release_decision', 'prefixes': ('headline_values__release_bridge_',), 'weight': 0.15}, {'name': 'executive_recovery_decision', 'prefixes': ('headline_values__executive_recovery_',), 'weight': 0.15})}, 'task_027': {'sections': ({'name': 'artifact_and_sources', 'prefixes': ('structure__', 'model_integrity__', 'model_content__', 'sources__', 'controls__'), 'weight': 0.05}, {'name': 'formula_auditability', 'prefixes': ('formula_lineage__',), 'weight': 0.08}, {'name': 'scenario_financials', 'prefixes': ('base_scenario_values__', 'downside_scenario_values__', 'upside_scenario_values__', 'scenario_comparison_values__'), 'weight': 0.14}, {'name': 'quarterly_liquidity', 'prefixes': ('base_liquidity_values__', 'downside_liquidity_values__', 'upside_liquidity_values__'), 'weight': 0.12}, {'name': 'scenario_attribution', 'prefixes': ('downside_scenario_attribution_values__', 'upside_scenario_attribution_values__'), 'weight': 0.1}, {'name': 'contingency_selection', 'prefixes': tuple((f'downside_contingency_values__{prefix}' for prefix in ('contingency_capex_deferral_', 'contingency_collections_acceleration_', 'contingency_construction_pricing_', 'contingency_controls_mix_', 'contingency_fixed_opex_freeze_', 'contingency_procurement_savings_', 'contingency_service_productivity_', 'contingency_targeted_overhead_reduction_', 'contingency_selected_', 'contingency_annual_', 'contingency_implementation_', 'contingency_adjusted_', 'contingency_ebitda_margin_headroom', 'contingency_free_cash_flow_headroom'))), 'weight': 0.11}, {'name': 'weekly_cash_support', 'prefixes': ('downside_contingency_values__contingency_week_', 'downside_contingency_values__contingency_13_week_'), 'weight': 0.1}, {'name': 'lender_release_decision', 'prefixes': ('covenant_release_values__',), 'weight': 0.3})}}
CRITICAL_SECTION_GATES = {'task_004': ({'section': 'wip_finance_and_controls', 'minimum_accuracy': 0.5, 'cap': 0.49}, {'section': 'commercial_review_and_release_decision', 'minimum_accuracy': 0.45, 'cap': 0.49}, {'section': 'accounting_to_workpaper_control', 'minimum_accuracy': 0.45, 'cap': 0.49}), 'task_015': ({'section': 'covenant_reperformance', 'minimum_accuracy': 0.55, 'cap': 0.49}, {'section': 'calculation_support', 'minimum_accuracy': 0.45, 'cap': 0.49}, {'section': 'downside_capacity', 'minimum_accuracy': 0.35, 'cap': 0.49}), 'task_035': ({'section': 'gross_commitment_stress', 'minimum_accuracy': 0.45, 'cap': 0.49}, {'section': 'execution_portfolio', 'minimum_accuracy': 0.4, 'cap': 0.49}, {'section': 'executive_recovery_decision', 'minimum_accuracy': 0.45, 'cap': 0.49}), 'task_068': ({'section': 'accounting_and_reporting_tie', 'minimum_accuracy': 0.55, 'cap': 0.49}, {'section': 'mitigation_and_release_decision', 'minimum_accuracy': 0.55, 'cap': 0.49})}

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
    critical_section_gate_failures: list[dict[str, Any]] = []
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
    if section_policy and criteria:
        sections = tuple(section_policy.get('sections', ()))
        if not sections:
            raise ValueError(f'missing section-normalized policy for {task_id}')
        section_results: dict[str, dict[str, Any]] = {}
        matched_ids: set[str] = set()
        normalized_reward = 0.0
        declared_weight = 0.0
        for section in sections:
            name = str(section['name'])
            prefixes = tuple((str(prefix) for prefix in section.get('prefixes', ())))
            criterion_ids = {str(criterion_id) for criterion_id in section.get('criterion_ids', ())}
            rows = [row for row in criteria if str(row.get('id', '')) in criterion_ids or (prefixes and str(row.get('id', '')).startswith(prefixes))]
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
            section_results[name] = {'prefixes': list(prefixes), 'criterion_ids': sorted(criterion_ids), 'section_weight': section_weight, 'weight_earned': earned, 'weight_total': total, 'accuracy': round(accuracy, 6)}
        all_ids = {str(row['id']) for row in criteria}
        if matched_ids != all_ids:
            raise ValueError(f'unassigned section-normalized criteria for {task_id}: {sorted(all_ids - matched_ids)}')
        if abs(declared_weight - 1.0) > 1e-09:
            raise ValueError(f'invalid section weights for {task_id}')
        reward = normalized_reward
        section_normalized_reward = {'method': 'linear_multi_section_normalization', 'sections': section_results, 'section_normalized_reward': round(reward, 6)}
        for gate in CRITICAL_SECTION_GATES.get(task_id, ()):
            section_name = str(gate['section'])
            if section_name not in section_results:
                raise ValueError(f'missing critical section gate for {task_id}: {section_name}')
            accuracy = float(section_results[section_name]['accuracy'])
            minimum_accuracy = float(gate['minimum_accuracy'])
            if accuracy < minimum_accuracy:
                cap = float(gate['cap'])
                reward = min(reward, cap)
                critical_section_gate_failures.append({'criterion_id': f'critical_section__{section_name}', 'section': section_name, 'section_accuracy': round(accuracy, 6), 'minimum_accuracy': minimum_accuracy, 'cap': cap, 'reward_effect': 'central_workstream_pass_cap'})
    required_editable_artifact_by_task = {'task_001': 'Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx', 'task_027': 'Shared/Finance/FP&A/FY27 plan/FY27 EBITDA scenarios - WORKING.xlsx', 'task_035': 'Shared/Finance/FP&A/Backlog/FY27 backlog burn and capacity - WORKING.xlsx', 'task_068': 'Shared/Finance/Reporting/2026/06 June/June executive performance review - WORKING.pptx'}
    if task_id in required_editable_artifact_by_task and integrity is not None:
        required_artifact = required_editable_artifact_by_task[task_id]
        declared_artifact = str(integrity.get('required_artifact') or '')
        changed_files = {str(path) for path in integrity.get('changed_files', []) if path is not None}
        created_files = {str(path) for path in integrity.get('created_files', []) if path is not None}
        if declared_artifact == required_artifact and required_artifact not in changed_files | created_files:
            core_model_failure = {'code': 'missing_required_artifact_modification' if task_id == 'task_001' else 'missing_required_workbook_modification', 'required_artifact': required_artifact, 'changed_files_count': len(changed_files), 'created_files_count': len(created_files), 'reward_effect': 'zero_reward_hard_failure'}
    if task_id == 'task_001':
        substantive_prefixes = ('close__', 'bridge__', 'accounting__', 'population__', 'commercial__', 'journal__', 'source__', 'action__')
        substantive_met_ids = sorted((str(row.get('id')) for row in criteria if str(row.get('id', '')).startswith(substantive_prefixes) and bool(row.get('value'))))
        if not substantive_met_ids:
            core_model_failure = {'code': 'no_substantive_finance_work', 'substantive_criteria_met': 0, 'reward_effect': 'zero_reward_hard_failure'}
    hard_failures: list[dict[str, Any]] = []
    if integrity is not None:
        updated['integrity'] = copy.deepcopy(dict(integrity))
        hard_failures = [dict(row) for row in integrity.get('hard_failures', []) if isinstance(row, Mapping)]
        if hard_failures:
            reward = 0.0
    applied_caps = []
    quality_gate_failures = []
    for gate in critical_section_gate_failures:
        applied_caps.append({'criterion_id': gate['criterion_id'], 'cap': gate['cap']})
        quality_gate_failures.append(dict(gate))
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
    updated.update({'reward': round(reward, 6), 'raw_weighted_reward': round(raw_reward, 6), 'decision_accuracy_adjustment': decision_accuracy_adjustment, 'section_normalized_reward': section_normalized_reward, 'core_model_failure': core_model_failure, 'strict_pass': bool(criteria) and met_count == len(criteria) and (not hard_failures) and (core_model_failure is None) and (not any((cap == 0.0 for (cap, _criterion_id) in failed_declared_gates))), 'criteria_met': met_count, 'criteria_total': len(criteria), 'weight_earned': earned_weight, 'weight_total': total_weight, 'reward_schema_version': REWARD_SCHEMA_VERSION, 'reward_definition': 'weighted binary criteria using 1/3/5/10 importance; task-declared linear section normalization keeps repeated support schedules proportional to separately weighted management judgment; tasks with a declared required editable artifact receive zero when that artifact is not modified or created; Tasks079 and 080 blank/default responses with no substantive finance work receive zero; decision-quality caps preserve supporting-work credit, distinguish correct schedule support from omitted facts, and prevent wrong critical business outputs from being masked; non-zero declared quality thresholds block strict pass, while task-declared central-workstream gates cap only materially incomplete professional modules; non-zero declared quality thresholds otherwise do not clip partial credit; declared zero-reward criteria and environment-integrity failures score zero', 'applied_reward_caps': applied_caps, 'quality_gate_failures': quality_gate_failures, 'hard_failures': hard_failures})
    return updated
