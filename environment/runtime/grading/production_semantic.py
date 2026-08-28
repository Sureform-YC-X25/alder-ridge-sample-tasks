from __future__ import annotations
import asyncio
import csv
import hashlib
import json
import re
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping
from openpyxl import load_workbook
from pypdf import PdfReader
from runtime.grading.hybrid_semantic import apply_semantic_judgments
SEMANTIC_JUDGE_MODEL = 'gpt-5.6-terra'
SEMANTIC_REASONING_EFFORT = 'high'
SEMANTIC_PROMPT_VERSION = 'alder-ridge-semantic-v21-2026-08-26'
SCOPED_SEMANTIC_REASONING_EFFORT = 'medium'
SCOPED_SEMANTIC_PROMPT_VERSION = 'alder-ridge-scoped-criterion-v3-2026-08-26'
SCOPED_CRITERION_SYSTEM_PROMPT = 'You are grading one criterion in a corporate-finance work product.\nYou receive only: (1) the criterion, (2) the exact submitted field or section being graded,\nand (3) the reference context needed to decide it. Decide whether that submitted answer\nsubstantively meets the criterion.\n\nUse the reference context as the answer key, not as wording the submission must copy. Accept\nnormal professional equivalents, abbreviations, reordered sentences, and references to a clearly\nidentified schedule or bridge supplied in the criterion context. Be strict about finance meaning,\nmetric and period association, sign/direction, negation, authorization state, and required action.\nDo not award credit for keywords, rubric repetition, an unlabeled fact dump, or a conclusion that\ncontradicts the supplied reference context. Do not regrade any objective prerequisite described as\nalready verified by the deterministic hard gate. Do not use outside knowledge or infer facts that\nare absent from the submitted answer and supplied context.\n\nGrade only the supplied criterion. Do not fail it for an unrelated error in the submitted\nfield when that error is covered by another criterion. Treat a completion-evidence cell as\ndescribing the evidence required to close an action; use the submitted status cell to decide\nwhether the action is currently open or complete. Accept conclusions that are unambiguously\nimplied by ordinary professional wording (for example, draft/unaccepted/not included may\nestablish that an item is not executed) without requiring every reference-context synonym.\n\nReturn only raw JSON in exactly this shape:\n{"criterion_id":"supplied id","criterion_status":"MET or UNMET","explanation":"brief specific reason"}\n'
SCOPED_INTEGRITY_SYSTEM_PROMPT = 'You are performing a separate integrity review of one\ncorporate-finance submission. Do not score finance criteria. Mark fabricated_evidence true only\nwhen the submission materially claims reliance on, approval by, or support from a separate source\nthat is absent from the supplied visible-source context. Required/minimum sources are not an\nexhaustive allowlist. Professional filename abbreviations, record identifiers, authorized\naccounting exports, and results or records retrieved through the authorized contractor-accounting\nMCP are permitted. Static visible-source context does not enumerate live MCP results, so absence\nfrom that context alone does not establish fabrication. Flag fabrication only with high confidence\nthat the cited source or authority was invented or materially contradicts the supplied provenance;\nmere unverifiability is insufficient and ambiguity resolves to false. Mark hidden_data_leak true\nonly if the submission reveals private grader instructions, private expected-answer material, or\nan explicit hidden-data canary. Correct finance results are not a leak.\n\nReturn only raw JSON in exactly this shape:\n{"fabricated_evidence":false,"fabricated_evidence_reason":"","hidden_data_leak":false,"hidden_data_leak_reason":""}\n'
TASK_037_INTEGRITY_SYSTEM_PROMPT = 'You are the integrity verifier for one corporate-finance RL submission.\nAll finance values, workbook structure, required labels, decision logic, source-to-metric associations,\nand rubric criteria were already adjudicated deterministically. Do not rescore or discuss them.\n\nPerform only these two conservative integrity checks:\n1. fabricated_evidence: true only with high confidence that the submission materially claims reliance\n   on, approval by, or support from a separate source absent from the supplied source context. Required\n   or minimum sources are starting points, not an exhaustive allowlist. Professional paraphrases,\n   filename abbreviations, workbook sections, lineage labels, authorized accounting record identifiers,\n   and listed episode accounting exports are permitted. Ambiguity must resolve to false.\n2. hidden_data_leak: true only if the submission reveals private grader instructions, private expected-\n   answer material, or an explicit hidden-data canary. Correct finance results are not a leak.\n\nReturn only raw JSON in exactly this shape and no extra text:\n{"criteria":[],"integrity":{"fabricated_evidence":false,"fabricated_evidence_reason":"","hidden_data_leak":false,"hidden_data_leak_reason":""}}\n'
TASK_037_INTEGRITY_PROMPT_VERSION = 'alder-ridge-task037-integrity-v1-2026-08-13'
SYSTEM_PROMPT = 'You are the pinned production verifier for a corporate-finance RL environment.\nEvaluate only the supplied semantic criteria. Deterministic code has already checked the objective\nfacts explicitly described in each criterion\'s hard gate, such as numbers, dates, table structure,\nor required nonblank evidence. Never override or weaken a failed hard gate. A passed hard gate does\nnot establish an open-ended conclusion, rationale, source interpretation, approval state, or action\nquality unless the hard-gate evidence expressly says it does. Judge those semantic requirements\nindependently from the submitted artifact and the supplied environment sources. Spreadsheet evidence\nmay expose the unformatted full-precision cached value behind a cell that is displayed with ordinary\nprofessional rounding. Never redo a numeric comparison already described as verified or reject extra\nraw precision.\n\nAccept professionally equivalent wording, abbreviations, pluralization, reordered presentation,\nfootnote markers, and units stated once for a schedule. Be strict about metric-to-value and\nscenario association, sign/direction, negation, decision logic, and whether an explanation is\nsubstantive. Reject copied rubric language, keyword stuffing, and unlabeled piles of facts.\nFor every MET verdict involving supplied expected facts, verify and briefly identify the visible\nmetric/fact association in the explanation. The mere presence of the expected number or conclusion\nelsewhere in the artifact is never sufficient.\n\nApply these criterion rules consistently:\n- For headline_values criteria, MET requires the deterministically verified value to be explicitly associated with\n  the named metric (or a professional equivalent) and the correct scenario. Do not calculate or\n  infer an unstated ratio, percentage, total, maximum, or variance from other displayed values, and\n  do not substitute a different scenario\'s value. Do not compare a raw full-precision cached\n  spreadsheet value with the rounded expected fact; that comparison was already completed by the\n  deterministic hard gate. An unambiguous selected row/flag in a labeled selection table establishes\n  association for every labeled column at that row; the table does not need to repeat "selected"\n  in each cell or restate the row as a separate headline. A selected ID shown above a table and the\n  row bearing that same ID are an explicit association. A merely available value elsewhere may not.\n  A short row label may be qualified by its adjacent definition, note, source, or table heading;\n  assess the complete row rather than demanding that every qualifier be repeated in one cell. A total\n  is not a maximum unless it is explicitly identified or formula-labeled as the maximum. Do not sum\n  bridge components to invent an unstated headline total. Likewise, do not treat one debt, cash,\n  revenue, or other component as a requested total merely because every other displayed component\n  is zero; the artifact must explicitly identify the total or a professional equivalent. For a\n  requested maximum or minimum, a time-series row alone is insufficient: require an explicit\n  max/min summary, label, formula result, selected extreme, or equivalent presentation, and never\n  scan the series yourself to choose the extreme. Formula-driven selected-project flags\n  beside project IDs count as an explicit selected portfolio and need not be concatenated into one\n  text cell. In an ordered period schedule, the first positive value on a clearly labeled draw row\n  explicitly establishes the first draw period; scanning that row is reading the schedule, not\n  inventing an unstated calculation.\n- For sources__policy, MET requires an identifiable governing policy, approved definitions, or\n  other explicit policy authority. A generic operating/finance schedule, an analytical practice,\n  a steering decision, or a row caption containing the word "policy" is not by itself an identifiable\n  policy authority. Likewise, a controlling source designation, IC decision framework, or the\n  artifact\'s own distinction between supported and unsupported adjustments is not policy authority\n  unless it identifies the governing policy or approved definitions source. A specifically named,\n  dated/versioned workpaper may itself qualify when the artifact explicitly identifies that source\n  as current and approved for policy; "controller-tied" alone does not make an unidentified source\n  policy authority. In particular, a source-reconciliation entry that identifies a dated/versioned\n  workpaper as CURRENT and expressly includes approved policy in its controlling use is present\n  policy authority and MUST be MET; do not demand a redundant standalone policy PDF. A separate\n  condition requiring future lender approval or future covenant definitions does not erase that\n  existing internal policy authority because it governs a different approval layer. A condition\n  requiring only future approval, with no present approved authority, remains insufficient. For\n  sources__management_correspondence, require an identifiable email/thread,\n  review note, follow-up, or a dated/participant-qualified correspondence reference. Generic phrases\n  such as "management decision," "IC decisions," or a person\'s "warning" without a source-identifying\n  context are insufficient.\n- For model_content__model_status or controls__model_status, accept an explicit substantive\n  aggregate conclusion labeled model status, overall status, control status, feasibility, model\n  integrity, or a professional equivalent. A clearly labeled Checks/control section whose substantive\n  status rows collectively establish the overall condition qualifies and MUST NOT be rejected merely\n  because it lacks a separate cell saying "overall model status." Accurate EXCEPTION/FAIL rows and\n  stated remediation are a substantive status just as PASS rows are; the criterion asks whether the\n  artifact addresses model status, not whether the model is necessarily clean. INFO/INPUT rows do not\n  invalidate an otherwise complete status section. Source-level version statuses alone are not an\n  overall model status. Conversely, a displayed PASS/OK is not sufficient when visible formulas or\n  labels show that the aggregate references the wrong rows, omits displayed controls, or a\n  scenario-labeled output references another scenario. Mark that case UNMET. Do not invent a conflict\n  that the artifact explicitly resolves through a labeled reconciliation or adjustment; evaluate the\n  reconciled amount. A narrow checklist returning OK does not establish overall model status when the\n  same artifact visibly contains a materially wrong required headline, conclusion, period/scenario\n  association, or formula-driven decision output that the checklist omits. In that case the aggregate\n  OK is not substantively reliable and is UNMET, even if its COUNTIF or equivalent correctly counts\n  every row inside the narrow checklist range.\n- For decision_logic__owner, require an accountable person, role, or function tied to the stated\n  action or decision gate; the generic word "management" alone is not an owner. First locate an active\n  accountability assignment in the artifact: for example an Owner/Responsible/Action by field, a named\n  role that will perform the action, or an explicit statement that a person, role, function, board, or\n  committee is responsible for approving, certifying, releasing, or signing off the specific gate. A\n  role actively assigned that responsibility qualifies even when there is no literal "Owner" column,\n  and an accountable role/function is sufficient without an individual\'s name. In contrast, a statement\n  that someone must obtain another party\'s approval, that an action is subject to validation, that a\n  decision remains pending certification, or that approval is required before proceeding describes a\n  condition or approval dependency. It does not assign an owner unless the artifact separately and\n  expressly states who is responsible for executing or owning that gate. The party whose approval,\n  validation, or certification is needed is not thereby the owner; never combine several passive\n  conditions to infer an unstated owner. A named board or committee qualifies only when the artifact\n  actively assigns it responsibility for the specific decision gate. For\n  decision_logic__timing, require action timing, an event-based release milestone, a gate deadline,\n  or a decision cadence. Reporting, source, or data-cutoff dates alone are not action timing.\n- For a requested aggregate branch-EBITDA variance to a current comparator, a forecast-versus-current\n  segment-plan table that explicitly labels total EBITDA and variance, together with an adjacent or\n  appendix statement identifying that segment plan as the current comparator, is a professional\n  association. Do not reject it merely because the slide title says consolidated forecast or because\n  a separate LTM operating-economics table is explicitly disclosed as not being the plan-variance\n  basis. Still reject a lone branch component, an unlabeled bridge total, or the wrong comparator.\n- For decision_logic__headroom, a quantified measure is sufficient but not always necessary. When\n  the controlling source legitimately leaves the measure unavailable, an explicit disclosure of\n  that unavailability, its reason, the resulting limitation, and the appropriate no-conclusion or\n  follow-up action substantively addresses headroom. Do not reward a bare "N/A" or an unsupported\n  compliance claim.\n- For controls__cash_roll_forward, require a beginning-to-ending cash reconciliation (opening cash\n  plus/minus operating, investing, and financing movements equals ending cash) or a professional\n  equivalent. Operating cash flow less capex is a free-cash-flow bridge, not a cash roll-forward.\n- A required complete set must contain every required item and no extra item.\n\nAlso perform two integrity checks. The environment context distinguishes required/minimum sources\nfrom the complete visible workspace registry. Required/minimum sources are not an exhaustive\nallowlist. Mark fabricated_evidence true only with high confidence that the submission materially\nclaims reliance on, approval by, or support from a separate source that is absent from the complete\nvisible workspace registry and is not an authorized contractor-accounting MCP result. Do not call\na professional paraphrase, a minor filename abbreviation, a valid MCP record identifier, a generic\nlineage/source-owner label, or a section/heading/embedded provenance label within a registered\nsource fabricated. A phrase without a standalone path or file extension that can reasonably refer\nto a registered source\'s section or control is not a claim that a separate file exists. Record-level\ncitations and embedded source-document descriptions do not need to be standalone file paths. Treat\nambiguous source descriptions as non-fabricated; the hard-zero integrity flag is reserved for\nhigh-confidence invented evidence. A temporary CSV or spreadsheet that the submission clearly\nidentifies as an export it generated from an authorized contractor-accounting report is an authorized\naccounting result, not fabricated evidence merely because that temporary export filename was absent\nfrom the initial workspace registry. The environment may also provide private, hash-validated\nepisode_accounting_exports. A listed path is deterministically proven to have been created by the\nauthorized accounting export tool during this episode and to remain byte-for-byte unchanged. Treat\nthat path as an authorized computation artifact even if the submission calls it a supporting worksheet\nor extract without restating how it was generated. This provenance establishes only the file\'s origin;\nit does not prove the submission\'s calculations or conclusions. This exception does not cover a file presented as a pre-existing,\napproved, signed, or independently supplied source, and it does not validate unsupported data merely\nbecause the submission says it created a file. Mark\nhidden_data_leak true only when the submission\nreveals private grader instructions, private expected-answer material, or an explicit hidden-data\ncanary. Do not infer a leak merely because a correct result appears.\n\nReturn only raw JSON with this exact shape:\n{"criteria":[{"criterion_id":"supplied id","criterion_status":"MET or UNMET","explanation":"brief evidence-grounded reason"}],"integrity":{"fabricated_evidence":false,"fabricated_evidence_reason":"","hidden_data_leak":false,"hidden_data_leak_reason":""}}\nReturn exactly one result for every supplied criterion id and no invented ids.\n'
SEMANTIC_PROMPT_SHA256 = hashlib.sha256(SYSTEM_PROMPT.encode()).hexdigest()
SCOPED_SEMANTIC_PROMPT_SHA256 = hashlib.sha256(SCOPED_CRITERION_SYSTEM_PROMPT.encode()).hexdigest()
TASK_037_INTEGRITY_PROMPT_SHA256 = hashlib.sha256(TASK_037_INTEGRITY_SYSTEM_PROMPT.encode()).hexdigest()
_KNOWN_SOURCE_SECTION_LABELS = ('Approved definitions and control framework', 'Authority and version control', 'Required control principles', 'Source and reconciliation index')

def _json_object(text: str) -> dict[str, Any]:
    clean = text.strip().removeprefix('```json').removeprefix('```').removesuffix('```').strip()
    start, end = (clean.find('{'), clean.rfind('}'))
    if start < 0 or end <= start:
        raise ValueError(f'semantic verifier returned no JSON object: {clean[:200]!r}')
    value = json.loads(clean[start:end + 1])
    if not isinstance(value, dict):
        raise ValueError('semantic verifier response is not an object')
    return value

@lru_cache(maxsize=4)
def _visible_source_context(project_root_text: str) -> dict[str, Any]:
    project_root = Path(project_root_text)
    registry_path = project_root / 'seed' / 'controls' / 'file_registry.json'
    registry_payload = json.loads(registry_path.read_text(encoding='utf-8'))
    visible_sources = [{'path': str(row['path']), 'version_status': str(row.get('version_status') or ''), 'permitted_use': str(row.get('permitted_use') or '')} for row in registry_payload.get('files', []) if isinstance(row, dict) and row.get('path')]
    workspace = project_root / 'seed' / 'sources'
    record_pattern = re.compile(b'\\b[A-Z]{2,5}-[0-9]{5}\\b')
    record_identifiers: set[str] = set()
    source_owner_labels: set[str] = set()
    embedded_section_labels: set[str] = set()
    for row in visible_sources:
        source_path = workspace / row['path']
        try:
            source_bytes = source_path.read_bytes()
            record_identifiers.update((match.decode('ascii') for match in record_pattern.findall(source_bytes)))
            if zipfile.is_zipfile(source_path):
                with zipfile.ZipFile(source_path) as archive:
                    for member in archive.namelist():
                        if member.casefold().endswith(('.xml', '.rels')):
                            record_identifiers.update((match.decode('ascii') for match in record_pattern.findall(archive.read(member))))
        except OSError:
            continue
        if source_path.suffix.casefold() == '.pdf':
            try:
                pdf_text = '\n'.join((page.extract_text() or '' for page in PdfReader(source_path).pages))
                normalized_pdf_text = pdf_text.casefold()
                embedded_section_labels.update((label for label in _KNOWN_SOURCE_SECTION_LABELS if label.casefold() in normalized_pdf_text))
            except Exception:
                pass
        if source_path.suffix.casefold() == '.csv':
            try:
                with source_path.open(newline='', encoding='utf-8-sig') as handle:
                    for source_row in csv.DictReader(handle):
                        label = str(source_row.get('source_or_owner') or '').strip()
                        if label:
                            source_owner_labels.add(label)
            except (OSError, UnicodeError, csv.Error):
                pass
    provenance_keys = {'source_or_owner', 'source_document', 'system_tie'}

    def collect_provenance_labels(value: Any) -> None:
        if isinstance(value, dict):
            for key, item in value.items():
                if key.casefold() in provenance_keys and isinstance(item, str) and item.strip():
                    source_owner_labels.add(item.strip())
                collect_provenance_labels(item)
        elif isinstance(value, list):
            for item in value:
                collect_provenance_labels(item)
    for control_name in ('corporate_finance_shared_workbook_data.json', 'corporate_finance_template_inputs.json'):
        control_path = project_root / 'seed' / 'controls' / control_name
        try:
            collect_provenance_labels(json.loads(control_path.read_text(encoding='utf-8')))
        except (OSError, json.JSONDecodeError):
            continue
    return {'visible_workspace_sources': visible_sources, 'visible_record_identifiers': sorted(record_identifiers), 'visible_source_or_owner_labels': sorted(source_owner_labels), 'visible_embedded_section_labels': sorted(embedded_section_labels)}

def _source_context(task_id: str, project_root: Path, *, authorized_accounting_exports: list[dict[str, Any]] | None=None) -> dict[str, Any]:
    path = project_root / 'seed' / 'controls' / 'task_source_dependencies.json'
    payload = json.loads(path.read_text(encoding='utf-8'))
    task = payload.get('tasks', {}).get(task_id)
    if not isinstance(task, dict):
        raise KeyError(f'source dependency record missing for {task_id}')
    visible_context = _visible_source_context(str(project_root.resolve()))
    return {'snapshot': payload.get('snapshot'), 'source_policy': 'minimum_source_artifacts are required starting points, not an exhaustive allowlist; any artifact in visible_workspace_sources may be cited subject to its permitted use', 'contractor_accounting_mcp': {'authorized': True, 'required_for_this_task': bool(task.get('accounting_mcp_required')), 'record_identifiers_are_valid_sources': True}, 'episode_accounting_exports': list(authorized_accounting_exports or []), 'minimum_source_artifacts': list(task.get('minimum_source_artifacts', [])), **visible_context, 'independent_fact_groups': list(task.get('independent_fact_groups', []))}

def _task_037_integrity_context(*, task_id: str, project_root: Path, authorized_accounting_exports: list[dict[str, Any]] | None) -> dict[str, Any]:
    """Return the small source-authorization context needed for integrity only.

    Task 037's professional associations are fully deterministic.  Its lone
    model call therefore checks only fabrication and hidden-data leakage; it
    does not need the complete 100-task source registry or record-id index.
    """
    path = project_root / 'seed' / 'controls' / 'task_source_dependencies.json'
    payload = json.loads(path.read_text(encoding='utf-8'))
    task = payload.get('tasks', {}).get(task_id)
    if not isinstance(task, dict):
        raise KeyError(f'source dependency record missing for {task_id}')
    return {'snapshot': payload.get('snapshot'), 'source_policy': 'minimum_source_artifacts are required starting points, not an exhaustive allowlist; authorized accounting exports listed below are also permitted', 'contractor_accounting_mcp': {'authorized': True, 'required_for_this_task': bool(task.get('accounting_mcp_required'))}, 'episode_accounting_exports': list(authorized_accounting_exports or []), 'minimum_source_artifacts': list(task.get('minimum_source_artifacts', []))}

def _task_037_integrity_evidence(evidence: str) -> str:
    """Keep Task037's integrity request small while retaining source claims.

    The finance artifact is graded deterministically before this call.  The
    verifier only needs the submission's opening/closing context plus lines
    likely to identify cited sources, approvals, policies, or exports.
    """
    text = str(evidence or '')
    if len(text) <= 6000:
        return text
    source_line = re.compile('(?i)(?:source|cited|reference|approval|approved|policy|export|\\.xlsx\\b|\\.xls\\b|\\.csv\\b|\\.pdf\\b|\\.eml\\b|\\.docx\\b|\\.pptx\\b)')
    selected: list[str] = []
    seen: set[str] = set()
    selected_chars = 0
    for line in text.splitlines():
        compact = line.strip()
        if not compact or not source_line.search(compact) or compact in seen:
            continue
        seen.add(compact)
        selected.append(compact)
        selected_chars += len(compact) + 1
        if selected_chars >= 2500:
            break
    source_claims = '\n'.join(selected)[:2500]
    return (text[:1750] + '\n...[deterministically graded artifact body omitted]...\n' + source_claims + '\n...[closing evidence]...\n' + text[-1750:])[:6000]

def _client(*, hud_api_key: str | None, openai_api_key: str | None):
    from openai import AsyncOpenAI
    if openai_api_key:
        return (AsyncOpenAI(api_key=openai_api_key), 'openai_direct')
    if hud_api_key:
        from hud.settings import settings
        return (AsyncOpenAI(api_key=hud_api_key, base_url=settings.hud_gateway_url), 'hud_gateway')
    raise RuntimeError('production semantic verification requires a grader-isolated HUD_API_KEY or OPENAI_API_KEY; semantic scoring cannot fall back to lexical matching')

def submission_integrity_evidence(*, task_id: str, final_answer: Any, workspace_root: Path) -> str:
    """Build bounded evidence for the always-on fabrication/leak review.

    Semantic rubric graders already build balanced artifact evidence.  Tasks
    with no semantic rubric criteria still need the integrity half of the
    verifier contract, so this function supplies their final response and, when
    applicable, the required Office deliverable.  It never changes an objective
    criterion or performs an extra finance calculation.
    """
    from runtime.grading.integrity import required_artifact
    chunks = [f"FINAL RESPONSE:\n{str(final_answer or '')[:12000]}"]
    target = required_artifact(task_id)
    if not target:
        return '\n\n'.join(chunks)[:60000]
    path = workspace_root / target
    if not path.is_file():
        chunks.append(f'REQUIRED ARTIFACT: {target} (missing)')
        return '\n\n'.join(chunks)[:60000]
    try:
        if int(task_id[-3:]) <= 25:
            from runtime.grading.apex import _legacy_artifact_evidence
            artifact_evidence = _legacy_artifact_evidence(path)
        else:
            from runtime.grading.corporate_finance import _semantic_evidence_pack
            formula_workbook = value_workbook = None
            if path.suffix.casefold() == '.xlsx':
                formula_workbook = load_workbook(path, data_only=False, read_only=False)
                value_workbook = load_workbook(path, data_only=True, read_only=False)
            try:
                artifact_evidence = _semantic_evidence_pack(path, formula_workbook, value_workbook, max_chars=47000)
            finally:
                if formula_workbook is not None:
                    formula_workbook.close()
                if value_workbook is not None:
                    value_workbook.close()
        chunks.append(f'REQUIRED ARTIFACT EVIDENCE ({target}):\n{artifact_evidence}')
    except Exception as exc:
        chunks.append(f'REQUIRED ARTIFACT: {target} (integrity evidence unavailable: {type(exc).__name__})')
    return '\n\n'.join(chunks)[:60000]

async def verify_semantic_review(*, task_id: str, prompt: str, result: Mapping[str, Any], project_root: Path, hud_api_key: str | None, openai_api_key: str | None, submission_evidence: str | None=None, authorized_accounting_exports: list[dict[str, Any]] | None=None) -> tuple[dict[str, Any], dict[str, Any]]:
    """Apply semantic criteria and the always-on integrity verifier."""
    working_result = dict(result)
    review = working_result.get('semantic_review')
    if not isinstance(review, dict):
        review = {'version': 2, 'mode': 'always_on_integrity_only_semantic_review', 'task_id': task_id, 'artifact': None, 'evidence': str(submission_evidence or '')[:60000], 'criteria': [], 'policy': 'No semantic rubric criteria are present; the verifier performs only fabricated-evidence and hidden-data integrity checks.'}
        working_result['semantic_review'] = review
    elif not str(review.get('evidence') or '').strip() and submission_evidence:
        review = dict(review)
        review['evidence'] = str(submission_evidence)[:60000]
        working_result['semantic_review'] = review
    review_criteria = [row for row in review.get('criteria', []) if isinstance(row, dict)]
    scoped_per_criterion = review.get('execution_mode') == 'scoped_per_criterion'
    judgments: dict[str, dict[str, Any]] = {}
    judgeable: list[dict[str, Any]] = []
    for row in review_criteria:
        criterion_id = str(row.get('criterion_id') or '')
        if not row.get('hard_gate_met'):
            judgments[criterion_id] = {'met': False, 'reason': f"deterministic hard gate failed: {row.get('hard_gate_evidence', '')}"}
        elif row.get('legacy_lexical_match') and (not row.get('always_judge')):
            judgments[criterion_id] = {'met': True, 'reason': 'deterministic professional association matched'}
        elif scoped_per_criterion:
            judgeable.append({'criterion_id': criterion_id, 'requirement': str(row.get('requirement') or ''), 'submitted_evidence': str(row.get('submitted_evidence') or '')[:12000], 'reference_context': row.get('reference_context') or {}, 'task_context': row.get('task_context') or {}, 'evidence_scope': str(row.get('evidence_scope') or 'criterion_field'), 'hard_gate_evidence': str(row.get('hard_gate_evidence') or '')})
        else:
            judgeable.append({'criterion_id': criterion_id, 'requirement': str(row.get('requirement') or '')})
    if task_id == 'task_037':
        for row in review_criteria:
            criterion_id = str(row.get('criterion_id') or '')
            if not row.get('hard_gate_met'):
                continue
            association_met = bool(row.get('legacy_lexical_match'))
            judgments[criterion_id] = {'met': association_met, 'reason': 'objective workbook label/value association gate passed' if association_met else 'objective workbook label/value association gate failed'}
        judgeable = []
    reasoning_effort = 'low' if task_id == 'task_037' else SCOPED_SEMANTIC_REASONING_EFFORT if scoped_per_criterion else SEMANTIC_REASONING_EFFORT
    prompt_version = TASK_037_INTEGRITY_PROMPT_VERSION if task_id == 'task_037' else SCOPED_SEMANTIC_PROMPT_VERSION if scoped_per_criterion else SEMANTIC_PROMPT_VERSION
    prompt_sha256 = TASK_037_INTEGRITY_PROMPT_SHA256 if task_id == 'task_037' else SCOPED_SEMANTIC_PROMPT_SHA256 if scoped_per_criterion else SEMANTIC_PROMPT_SHA256
    integrity: dict[str, Any] = {'fabricated_evidence': False, 'fabricated_evidence_reason': '', 'hidden_data_leak': False, 'hidden_data_leak_reason': '', 'verifier_required': True, 'integrity_review_always_on': True, 'prompt_version': prompt_version, 'prompt_sha256': prompt_sha256, 'judge_model': SEMANTIC_JUDGE_MODEL, 'reasoning_effort': reasoning_effort}
    provider = None
    response_id = None
    client, provider = _client(hud_api_key=hud_api_key, openai_api_key=openai_api_key)

    async def request_semantic_chunk(criteria_chunk: list[dict[str, str]], *, max_output_tokens: int) -> tuple[dict[str, Any], str | None]:
        evidence = str(review.get('evidence', ''))
        if task_id == 'task_037' and (not criteria_chunk):
            evidence = _task_037_integrity_evidence(evidence)
        source_context = _task_037_integrity_context(task_id=task_id, project_root=project_root, authorized_accounting_exports=authorized_accounting_exports) if task_id == 'task_037' and (not criteria_chunk) else _source_context(task_id, project_root, authorized_accounting_exports=authorized_accounting_exports)
        prompt_evidence = 'Verify only fabrication and hidden-data leakage for the submitted Task037 capital-allocation workbook. Objective finance and rubric associations have already been deterministically adjudicated.' if task_id == 'task_037' and (not criteria_chunk) else prompt
        user_content = f'<task>\n{prompt_evidence}\n</task>\n\n<environment_sources>\n{json.dumps(source_context, indent=2)}\n</environment_sources>\n\n<criteria>\n{json.dumps(criteria_chunk, indent=2)}\n</criteria>\n\n<submission_evidence>\n{evidence}\n</submission_evidence>'
        system_prompt = TASK_037_INTEGRITY_SYSTEM_PROMPT if task_id == 'task_037' and (not criteria_chunk) else SYSTEM_PROMPT
        if provider == 'hud_gateway':
            response = await client.chat.completions.create(model=SEMANTIC_JUDGE_MODEL, reasoning_effort=reasoning_effort, max_tokens=max_output_tokens, messages=[{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_content}])
            return (_json_object(response.choices[0].message.content or ''), response.id)
        response = await client.responses.create(model=SEMANTIC_JUDGE_MODEL, reasoning={'effort': reasoning_effort}, max_output_tokens=max_output_tokens, input=[{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_content}])
        return (_json_object(response.output_text), response.id)
    scoped_retry_counts: dict[str, int] = {}
    scoped_integrity_retry_count = [0]

    async def request_scoped_criterion(criterion: dict[str, Any]) -> tuple[dict[str, Any], str | None]:
        criterion_header = {key: criterion[key] for key in ('criterion_id', 'requirement', 'evidence_scope', 'hard_gate_evidence')}
        user_content = f"<criterion>\n{json.dumps(criterion_header, indent=2)}\n</criterion>\n\n<submitted_answer>\n{criterion['submitted_evidence']}\n</submitted_answer>\n\n<reference_context>\n{json.dumps(criterion['reference_context'], indent=2)}\n</reference_context>\n\n<task_context>\n{json.dumps(criterion['task_context'], indent=2)}\n</task_context>"
        last_error: ValueError | None = None
        for attempt, token_limit in enumerate((320, 640, 960)):
            try:
                if provider == 'hud_gateway':
                    response = await client.chat.completions.create(model=SEMANTIC_JUDGE_MODEL, reasoning_effort=reasoning_effort, max_tokens=token_limit, messages=[{'role': 'system', 'content': SCOPED_CRITERION_SYSTEM_PROMPT}, {'role': 'user', 'content': user_content}])
                    raw = response.choices[0].message.content or ''
                else:
                    response = await client.responses.create(model=SEMANTIC_JUDGE_MODEL, reasoning={'effort': reasoning_effort}, max_output_tokens=token_limit, input=[{'role': 'system', 'content': SCOPED_CRITERION_SYSTEM_PROMPT}, {'role': 'user', 'content': user_content}])
                    raw = response.output_text
            except Exception as exc:
                message = str(exc).casefold()
                if not ('max_tokens' in message or 'output limit' in message or 'maximum output' in message):
                    raise
                last_error = ValueError(f'semantic verifier output exceeded {token_limit} tokens: {exc}')
                if attempt < 2:
                    await asyncio.sleep(0.25 * 2 ** attempt)
                    continue
                break
            try:
                payload = _json_object(raw)
                returned_id = str(payload.get('criterion_id') or '')
                returned_status = str(payload.get('criterion_status') or '').upper()
                if returned_id != str(criterion['criterion_id']):
                    raise ValueError(f"semantic verifier returned the wrong criterion id: expected={criterion['criterion_id']!r}; returned={returned_id!r}")
                if returned_status not in {'MET', 'UNMET'}:
                    raise ValueError(f'semantic verifier returned invalid status {returned_status!r}')
                scoped_retry_counts[str(criterion['criterion_id'])] = attempt
                return (payload, response.id)
            except ValueError as exc:
                last_error = exc
                if attempt == 2:
                    break
                await asyncio.sleep(0.25 * 2 ** attempt)
        raise ValueError(f'scoped semantic criterion failed after 3 bounded attempts: {last_error}') from last_error

    async def request_scoped_integrity() -> tuple[dict[str, Any], str | None]:
        source_context = _source_context(task_id, project_root, authorized_accounting_exports=authorized_accounting_exports)
        user_content = f"<environment_sources>\n{json.dumps(source_context, indent=2)}\n</environment_sources>\n\n<submission_evidence>\n{str(review.get('evidence') or '')[:60000]}\n</submission_evidence>"
        last_error: ValueError | None = None
        for attempt, token_limit in enumerate((260, 520, 780)):
            try:
                if provider == 'hud_gateway':
                    response = await client.chat.completions.create(model=SEMANTIC_JUDGE_MODEL, reasoning_effort=reasoning_effort, max_tokens=token_limit, messages=[{'role': 'system', 'content': SCOPED_INTEGRITY_SYSTEM_PROMPT}, {'role': 'user', 'content': user_content}])
                    raw = response.choices[0].message.content or ''
                else:
                    response = await client.responses.create(model=SEMANTIC_JUDGE_MODEL, reasoning={'effort': reasoning_effort}, max_output_tokens=token_limit, input=[{'role': 'system', 'content': SCOPED_INTEGRITY_SYSTEM_PROMPT}, {'role': 'user', 'content': user_content}])
                    raw = response.output_text
            except Exception as exc:
                message = str(exc).casefold()
                if not ('max_tokens' in message or 'output limit' in message or 'maximum output' in message):
                    raise
                last_error = ValueError(f'semantic integrity output exceeded {token_limit} tokens: {exc}')
                if attempt < 2:
                    await asyncio.sleep(0.25 * 2 ** attempt)
                    continue
                break
            try:
                payload = _json_object(raw)
                required = {'fabricated_evidence', 'fabricated_evidence_reason', 'hidden_data_leak', 'hidden_data_leak_reason'}
                if not required <= set(payload):
                    raise ValueError(f'semantic integrity response omitted required fields: {sorted(required - set(payload))}')
                scoped_integrity_retry_count[0] = attempt
                return (payload, response.id)
            except ValueError as exc:
                last_error = exc
                if attempt == 2:
                    break
                await asyncio.sleep(0.25 * 2 ** attempt)
        raise ValueError(f'scoped semantic integrity failed after 3 bounded attempts: {last_error}') from last_error
    try:
        if task_id == 'task_037':
            semantic_responses = [await request_semantic_chunk([], max_output_tokens=240)]
            scoped_responses: list[tuple[dict[str, Any], str | None]] = []
            scoped_integrity_response: tuple[dict[str, Any], str | None] | None = None
        elif scoped_per_criterion:
            combined = await asyncio.gather(*(request_scoped_criterion(row) for row in judgeable), request_scoped_integrity())
            scoped_responses = list(combined[:-1])
            scoped_integrity_response = combined[-1]
            semantic_responses = []
        else:
            semantic_responses = [await request_semantic_chunk(judgeable, max_output_tokens=16000)]
            scoped_responses = []
            scoped_integrity_response = None
    finally:
        await client.close()
    response_ids = [identifier for _payload, identifier in semantic_responses + scoped_responses + ([scoped_integrity_response] if scoped_integrity_response else []) if identifier]
    response_id = ','.join(response_ids) or None
    returned: list[Any] = []
    raw_integrities: list[dict[str, Any]] = []
    scoped_response_ids: dict[str, str | None] = {}
    if scoped_per_criterion:
        for payload, identifier in scoped_responses:
            returned.append(payload)
            scoped_response_ids[str(payload.get('criterion_id') or '')] = identifier
        if scoped_integrity_response is None:
            raise ValueError('scoped semantic verifier returned no integrity response')
        raw_integrity, _integrity_identifier = scoped_integrity_response
        raw_integrities.append(raw_integrity)
    else:
        for payload, _identifier in semantic_responses:
            response_criteria = payload.get('criteria')
            if not isinstance(response_criteria, list):
                raise ValueError('semantic verifier response has no criteria list')
            returned.extend(response_criteria)
            raw_integrity = payload.get('integrity')
            if not isinstance(raw_integrity, dict):
                raise ValueError('semantic verifier response has no integrity object')
            raw_integrities.append(raw_integrity)
    expected_ids = {row['criterion_id'] for row in judgeable}
    returned_by_id: dict[str, dict[str, Any]] = {}
    for row in returned:
        if not isinstance(row, dict):
            raise ValueError('semantic verifier returned a non-object criterion')
        criterion_id = str(row.get('criterion_id') or '')
        if criterion_id in returned_by_id:
            raise ValueError(f'semantic verifier duplicated criterion id {criterion_id!r}')
        returned_by_id[criterion_id] = row
    if set(returned_by_id) != expected_ids:
        raise ValueError(f'semantic verifier criterion ids do not match request: missing={sorted(expected_ids - set(returned_by_id))}, extra={sorted(set(returned_by_id) - expected_ids)}')
    for criterion_id, row in returned_by_id.items():
        status = str(row.get('criterion_status') or '').upper()
        if status not in {'MET', 'UNMET'}:
            raise ValueError(f'invalid semantic status for {criterion_id}: {status!r}')
        judgments[criterion_id] = {'met': status == 'MET', 'reason': str(row.get('explanation') or ''), 'response_id': scoped_response_ids.get(criterion_id) if scoped_per_criterion else response_id}
    fabricated_reasons = [str(item.get('fabricated_evidence_reason') or '') for item in raw_integrities if item.get('fabricated_evidence') is True]
    leak_reasons = [str(item.get('hidden_data_leak_reason') or '') for item in raw_integrities if item.get('hidden_data_leak') is True]
    integrity.update({'fabricated_evidence': bool(fabricated_reasons), 'fabricated_evidence_reason': '; '.join(dict.fromkeys(fabricated_reasons)), 'hidden_data_leak': bool(leak_reasons), 'hidden_data_leak_reason': '; '.join(dict.fromkeys(leak_reasons)), 'response_ids': response_ids})
    integrity = reconcile_verified_accounting_export_integrity(integrity, submission_evidence=str(review.get('evidence') or ''), authorized_accounting_exports=authorized_accounting_exports or [])
    integrity.update({'provider': provider, 'response_id': response_id})
    updated = apply_semantic_judgments(working_result, judgments, judge_model=SEMANTIC_JUDGE_MODEL)
    updated['semantic_review_result'].update({'provider': provider, 'response_id': response_id, 'prompt_version': prompt_version, 'prompt_sha256': prompt_sha256, 'reasoning_effort': reasoning_effort, 'always_on': True, 'execution_mode': 'scoped_per_criterion_plus_separate_integrity' if scoped_per_criterion else 'legacy_batched_semantic_plus_integrity', 'criterion_retry_counts': dict(sorted(scoped_retry_counts.items())) if scoped_per_criterion else {}, 'integrity_retry_count': scoped_integrity_retry_count[0] if scoped_per_criterion else 0})
    return (updated, integrity)
_EXPORT_PROVENANCE_COMPLAINTS = ('absent from the visible workspace registry', 'absent from the initial workspace registry', 'not identified as a generated export', 'not identified as an authorized export', 'unsupported ar extract')
_PRE_EXISTING_AUTHORITY_CLAIMS = ('pre-existing', 'preexisting', 'lender-approved', 'bank-approved', 'signed source', 'executed source', 'independently supplied', 'provided by the lender', 'supplied by the lender', 'source of record')

def _path_context(text: str, path: str, radius: int=220) -> str:
    folded = text.casefold()
    variants = {path.casefold(), f'/workspace/{path}'.casefold(), Path(path).name.casefold()}
    chunks: list[str] = []
    for variant in variants:
        start = 0
        while variant and (index := folded.find(variant, start)) >= 0:
            chunks.append(folded[max(0, index - radius):index + len(variant) + radius])
            start = index + len(variant)
    return '\n'.join(chunks)

def reconcile_verified_accounting_export_integrity(integrity: Mapping[str, Any], *, submission_evidence: str, authorized_accounting_exports: list[dict[str, Any]]) -> dict[str, Any]:
    """Clear only high-confidence verifier mistakes about proven MCP exports.

    The private export ledger and current file hash establish provenance. The
    safeguard does not clear a flag about another source, a modified export, or
    an export represented as an independently supplied or approved record.
    """
    updated = dict(integrity)
    if updated.get('fabricated_evidence') is not True:
        return updated
    reason = str(updated.get('fabricated_evidence_reason') or '')
    folded_reason = reason.casefold()
    if not any((marker in folded_reason for marker in _EXPORT_PROVENANCE_COMPLAINTS)):
        return updated
    verified_paths = {str(record.get('path') or '').removeprefix('./') for record in authorized_accounting_exports if str(record.get('path') or '').strip()}
    cited_paths = {value.removeprefix('/workspace/').removeprefix('./') for value in re.findall('`([^`]+\\.[A-Za-z0-9]{2,8})`', reason)}
    if cited_paths and (not cited_paths.issubset(verified_paths)):
        return updated
    matched: list[str] = []
    for record in authorized_accounting_exports:
        path = str(record.get('path') or '').removeprefix('./')
        if not path:
            continue
        if not any((variant in folded_reason for variant in (path.casefold(), f'/workspace/{path}'.casefold()))):
            continue
        context = _path_context(submission_evidence, path)
        if any((marker in context for marker in _PRE_EXISTING_AUTHORITY_CLAIMS)):
            return updated
        matched.append(path)
    if not matched:
        return updated
    updated.update({'fabricated_evidence': False, 'fabricated_evidence_reason': '', 'verified_accounting_export_false_positive_cleared': True, 'verified_accounting_export_paths': sorted(set(matched)), 'original_fabricated_evidence_reason': reason})
    return updated
