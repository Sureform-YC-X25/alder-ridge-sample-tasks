from __future__ import annotations
import json
import shutil
from pathlib import Path
from typing import Any, Mapping
from runtime.accounting_mcp.paths import resolve_seed_root
from runtime.accounting_mcp.provenance import verified_accounting_exports
from runtime.grading.apex import grade_apex_task
from runtime.grading.integrity import WorkspaceCreationMonitor, assess_integrity, capture_integrity_snapshot
from runtime.grading.production_semantic import submission_integrity_evidence, verify_semantic_review
from runtime.grading.rubric import apply_reward_policy, attach_default_policy
from runtime.task_catalog import TASKS

def _workspace_slug(task_id: str, declared_slug: str | None=None) -> str | None:
    return declared_slug

def prepare_task_workspace(task_id: str, runtime_root: Path, *, task_slug: str | None=None, seed_root: Path | None=None) -> None:
    """The selected sample uses the shared canonical workspace without overlays."""
    return None

def invalidate_on_grading_error(result: dict[str, Any]) -> dict[str, Any]:
    """Turn mandatory-verifier failure into an excluded, zero-reward result."""
    if result.get('grading_error'):
        result['reward'] = 0.0
        result['strict_pass'] = False
        result['infrastructure_valid'] = False
        result.setdefault('hard_failures', []).append({'code': 'semantic_verifier_unavailable', 'message': result['grading_error'], 'reward': 0.0})
    else:
        result['infrastructure_valid'] = True
    return result

def _compact_integrity(integrity: Any) -> Any:
    """Keep HUD transport evidence bounded while preserving decisive signals.

    The complete integrity record is persisted in ``latest_grade_result.json``.
    Agent-created dependency caches can contain thousands of paths, so sending
    those path arrays again over HUD's newline-delimited subprocess protocol can
    exceed its per-message limit even though grading completed successfully.
    """
    if not isinstance(integrity, Mapping):
        return integrity
    compact = {key: integrity.get(key) for key in ('version', 'task_id', 'required_artifact', 'database_changed', 'audit_log_changed', 'artifact_readable', 'artifact_readability_evidence', 'semantic_integrity', 'startup_isolation_attestation') if key in integrity}
    for key in ('allowed_workspace_mutations', 'changed_files', 'created_files', 'deleted_files', 'allowed_auxiliary_files'):
        value = integrity.get(key)
        compact[f'{key}_count'] = len(value) if isinstance(value, list) else 0
    for key in ('forbidden_workspace_mutations', 'prohibited_tool_actions', 'hard_failures'):
        value = integrity.get(key)
        compact[key] = value if isinstance(value, list) else []
    return compact

def _hud_result(task_id: str, result: dict[str, Any]):
    """Expose an IPC-safe rubric summary and governed terminal reward in HUD.

    HUD's local subprocess protocol is newline-delimited.  A fully evidenced
    large workbook rubric can exceed asyncio's 64 KiB default line limit, so
    the full result is persisted in the grader-only state directory and this
    transport object carries only the fields needed to verify every atomic
    outcome.  The frontier runner reads and cross-checks the sidecar before it
    writes the commercial trace.
    """
    from hud.graders import EvaluationResult
    criteria = [{key: criterion[key] for key in ('id', 'description', 'category', 'weight', 'semantic', 'failure_cap', 'value') if key in criterion} for criterion in result['criteria']]
    criterion_transport: dict[str, Any] = {}
    strict_pass = bool(result['strict_pass'])
    semantic_review = result.get('semantic_review_result')
    return EvaluationResult(reward=float(result['reward']), done=True, content=f"{('PASS' if strict_pass else 'INCOMPLETE')} — {result['criteria_met']}/{result['criteria_total']} atomic criteria met; weighted reward={result['reward']:.3f}", info={'task_id': task_id, 'strict_pass': strict_pass, 'criteria_met': result['criteria_met'], 'criteria_total': result['criteria_total'], 'weight_earned': result.get('weight_earned'), 'weight_total': result.get('weight_total'), 'criteria': criteria, **criterion_transport, 'pass_definition': 'strict pass only when every binary criterion equals 1', 'reward_definition': result.get('reward_definition'), 'reward_schema_version': result.get('reward_schema_version'), 'raw_weighted_reward': result.get('raw_weighted_reward'), 'decision_accuracy_adjustment': result.get('decision_accuracy_adjustment'), 'applied_reward_caps': result.get('applied_reward_caps'), 'quality_gate_failures': result.get('quality_gate_failures'), 'hard_failures': result.get('hard_failures'), 'integrity': _compact_integrity(result.get('integrity')), 'semantic_review_result': semantic_review, 'grading_error': result.get('grading_error'), 'infrastructure_valid': result.get('infrastructure_valid', True)}, subscores=[], isError=bool(result.get('grading_error')))

def persist_grade_sidecar(*, task_id: str, result: dict[str, Any], state_root: Path) -> None:
    """Persist full grader evidence outside the agent-visible workspace."""
    path = state_root / 'latest_grade_result.json'
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.json.tmp')
    temporary.write_text(json.dumps({'task_id': task_id, 'result': result}, default=str), encoding='utf-8')
    temporary.replace(path)

def register_task_templates(env: Any, runtime_root: Path, state_root: Path, project_root: Path, *, semantic_credentials: Mapping[str, str | None], included_task_ids: frozenset[str] | None=None) -> dict[str, Any]:
    """Register the requested independently runnable tasks on one company world."""
    registered: dict[str, Any] = {}
    for spec in TASKS:
        if included_task_ids is not None and spec.task_id not in included_task_ids:
            continue
        task_id = spec.task_id
        prompt = spec.prompt
        workspace_slug = _workspace_slug(task_id, spec.slug)

        @env.template(id=task_id, description=spec.title)
        async def task_template(_task_id: str=task_id, _prompt: str=prompt, _workspace_slug: str | None=workspace_slug):
            prepare_task_workspace(_task_id, runtime_root, task_slug=_workspace_slug)
            database_path = state_root / 'accounting.db'
            before = capture_integrity_snapshot(runtime_root, database_path)
            creation_monitor = None
            try:
                answer = (yield _prompt)
            finally:
                observed_workspace_creations = creation_monitor.stop() if creation_monitor is not None else []
            result = grade_apex_task(_task_id, answer, runtime_root)
            semantic_integrity: dict[str, Any] = {}
            try:
                accounting_exports = verified_accounting_exports(state_root=state_root, workspace_root=runtime_root)
                result, semantic_integrity = await verify_semantic_review(task_id=_task_id, prompt=_prompt, result=result, project_root=project_root, hud_api_key=semantic_credentials.get('hud_api_key'), openai_api_key=semantic_credentials.get('openai_api_key'), submission_evidence=submission_integrity_evidence(task_id=_task_id, final_answer=answer, workspace_root=runtime_root), authorized_accounting_exports=accounting_exports)
            except Exception as exc:
                result['grading_error'] = f'production semantic verifier failed: {type(exc).__name__}: {exc}'
            integrity = assess_integrity(task_id=_task_id, before=before, workspace_root=runtime_root, database_path=database_path, final_answer=answer, semantic_integrity=semantic_integrity, observed_workspace_creations=observed_workspace_creations)
            result = apply_reward_policy(_task_id, attach_default_policy(result), integrity=integrity)
            result = invalidate_on_grading_error(result)
            persist_grade_sidecar(task_id=_task_id, result=result, state_root=state_root)
            yield _hud_result(_task_id, result)
        registered[task_id] = task_template
    return registered
