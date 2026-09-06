from __future__ import annotations

import copy
from typing import Any, Mapping

from runtime.grading.rubric import apply_reward_policy, attach_default_policy


HYBRID_TASKS = frozenset(
    {"task_001", "task_004", "task_015", "task_035", "task_068"}
)
# Backward-compatible import for the pilot workflow.  The implementation is
# now benchmark-wide, but keeping the old name avoids breaking archived replay
# commands and already-published workflow files.
PILOT_TASKS = HYBRID_TASKS


def _apply_task_001_semantic_coherence(
    criteria_by_id: Mapping[str, dict[str, Any]],
    audit_by_id: Mapping[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """Resolve only strict logical contradictions among Task001 verdicts.

    These rules do not parse business wording. They operate after the scoped
    semantic judge has interpreted that wording and only where one adjudicated
    criterion necessarily entails another narrower criterion, or where two
    independently failed authority criteria contradict a claimed correct use
    of the same controlling source.
    """

    adjustments: list[dict[str, Any]] = []

    def value(criterion_id: str) -> bool:
        row = criteria_by_id.get(criterion_id)
        return bool(row and row.get("value"))

    def hard_gate(criterion_id: str) -> bool:
        row = audit_by_id.get(criterion_id)
        return bool(row and row.get("hard_gate_met"))

    def override(
        criterion_id: str,
        *,
        final_met: bool,
        rule: str,
        reason: str,
        depends_on: tuple[str, ...],
    ) -> None:
        criterion = criteria_by_id.get(criterion_id)
        audit = audit_by_id.get(criterion_id)
        if criterion is None or audit is None:
            return
        previous = bool(criterion.get("value"))
        if previous == final_met:
            return
        criterion["value"] = int(final_met)
        criterion["grading_method"] = (
            "deterministic_gate_semantic_judge_and_logical_coherence"
        )
        criterion["evidence"] = (
            f"{criterion.get('evidence', '')}; logical_coherence={rule}: {reason}"
        ).strip("; ")
        audit["pre_coherence_final_met"] = previous
        audit["final_met"] = final_met
        audit["coherence_rule"] = rule
        audit["coherence_reason"] = reason
        audit["coherence_depends_on"] = list(depends_on)
        adjustments.append({
            "criterion_id": criterion_id,
            "pre_coherence_final_met": previous,
            "final_met": final_met,
            "rule": rule,
            "reason": reason,
            "depends_on": list(depends_on),
        })

    # The broader June-eligible allocation criterion explicitly requires the
    # complete $46,800 gross / $13,003.31 prior / $33,796.69 due source
    # reconciliation. Once that criterion is semantically MET, failing the
    # narrower source-amount criterion is a judge inconsistency, not a
    # difference in the submitted work.
    source_amount = "population__cutoff_invoice__source_amount"
    june_eligible = "population__cutoff_invoice__june_eligible_amount"
    if value(june_eligible) and not value(source_amount) and hard_gate(source_amount):
        override(
            source_amount,
            final_met=True,
            rule="task001_invoice_allocation_entails_source_reconciliation_v1",
            reason=(
                "the semantically MET June-eligible allocation necessarily includes "
                "the complete cumulative invoice source reconciliation"
            ),
            depends_on=(june_eligible,),
        )

    # Correct use of the executed-credit source requires the same operative
    # condition and cutoff chronology tested by both authority criteria. A MET
    # source-use verdict cannot coexist with both of those semantic predicates
    # being UNMET. This removes no credit merely because a downstream numeric
    # calculation is wrong; it addresses the shared authority proposition only.
    executed_source = "source__executed_credit"
    authority = "commercial__authority_at_cutoff"
    record_conflict = "commercial__record_conflict"
    if value(executed_source) and not value(authority) and not value(record_conflict):
        override(
            executed_source,
            final_met=False,
            rule="task001_executed_source_requires_authority_resolution_v1",
            reason=(
                "both semantic authority predicates reject the operative cutoff "
                "treatment required by this source-use criterion"
            ),
            depends_on=(authority, record_conflict),
        )

    return adjustments


def semantic_requirement(
    *,
    criterion_id: str,
    description: str,
    expected_facts: Mapping[str, Any] | None = None,
    artifact_type: str,
) -> str:
    """Build a narrow, self-contained binary requirement for a semantic judge.

    Objective facts are supplied to the judge only after the deterministic
    grader has verified that they actually occur in the artifact.  The judge's
    role is therefore limited to professional-language association, scenario
    binding, negation, and substantive narrative sufficiency.
    """

    facts = dict(expected_facts or {})
    requirement = (
        f"[{criterion_id}] Decide whether the {artifact_type} substantively satisfies this "
        f"rubric criterion: {description}. Be flexible about professional wording, labels, "
        "abbreviations, plurals, table layout, footnote markers, and whether units are stated "
        "once for a schedule. Be strict about which metric and scenario each value belongs to, "
        "the direction/sign of conclusions, negation, and whether the requested analysis is "
        "actually present. Do not pass keyword lists, copied rubric prose, or an unlabeled pile "
        "of numbers as substantive evidence."
    )
    if facts:
        requirement += (
            " The deterministically verified expected facts that must be professionally "
            f"associated in the artifact are: {facts!r}."
        )
    return requirement


def apply_semantic_judgments(
    result: Mapping[str, Any],
    judgments: Mapping[str, Any],
    *,
    judge_model: str,
) -> dict[str, Any]:
    """Combine semantic verdicts with hard deterministic gates.

    A semantic pass can never rescue a failed numeric, formula, structure, or
    unit gate.  Conversely, exact professional facts are not rejected merely
    because their labels differ from a hand-maintained alias list.
    """

    updated = copy.deepcopy(dict(result))
    review = updated.get("semantic_review")
    if not isinstance(review, dict):
        return updated
    review_criteria = review.get("criteria")
    if not isinstance(review_criteria, list):
        return updated

    criteria_by_id = {
        str(criterion.get("id")): criterion
        for criterion in updated.get("criteria", [])
        if isinstance(criterion, dict)
    }
    audit: list[dict[str, Any]] = []
    for item in review_criteria:
        if not isinstance(item, dict):
            continue
        criterion_id = str(item.get("criterion_id") or "")
        criterion = criteria_by_id.get(criterion_id)
        if criterion is None:
            continue
        hard_gate_met = bool(item.get("hard_gate_met"))
        raw_judgment = judgments.get(criterion_id)
        if isinstance(raw_judgment, dict):
            semantic_met = bool(raw_judgment.get("met"))
            reason = str(raw_judgment.get("reason") or "")
            response_id = raw_judgment.get("response_id")
        elif raw_judgment is None:
            semantic_met = False
            reason = "semantic judge returned no verdict"
            response_id = None
        else:
            semantic_met = bool(raw_judgment)
            reason = ""
            response_id = None
        final_met = hard_gate_met and semantic_met
        lexical_value = int(bool(criterion.get("value")))
        criterion["value"] = int(final_met)
        criterion["grading_method"] = "deterministic_gate_and_semantic_judge"
        criterion["evidence"] = (
            f"hard_gate={hard_gate_met}: {item.get('hard_gate_evidence', '')}; "
            f"semantic_judge={'MET' if semantic_met else 'UNMET'} ({judge_model}): {reason}; "
            f"legacy_lexical_match={lexical_value}"
        )
        audit.append(
            {
                "criterion_id": criterion_id,
                "hard_gate_met": hard_gate_met,
                "semantic_met": semantic_met,
                "final_met": final_met,
                "legacy_lexical_match": lexical_value,
                "reason": reason,
                "response_id": response_id,
            }
        )

    audit_by_id = {
        str(row.get("criterion_id")): row
        for row in audit
        if isinstance(row, dict)
    }
    coherence_adjustments = (
        _apply_task_001_semantic_coherence(criteria_by_id, audit_by_id)
        if str(review.get("task_id") or "") == "task_001"
        else []
    )

    updated["semantic_review_result"] = {
        "mode": "environment_aligned_hybrid_semantic_v3",
        "judge_model": judge_model,
        "criteria": audit,
        "policy": "semantic MET is accepted only when its deterministic hard gate also passes",
        "logical_coherence": {
            "version": "task001-strict-entailment-v1" if coherence_adjustments else None,
            "adjustments": coherence_adjustments,
        },
    }
    task_id = str(review.get("task_id") or "unknown")
    return apply_reward_policy(task_id, attach_default_policy(updated))
