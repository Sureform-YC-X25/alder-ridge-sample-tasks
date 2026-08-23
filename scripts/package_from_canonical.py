#!/usr/bin/env python3
"""Build the buyer-facing Alder Ridge twelve-task sample from a canonical tree.

The output is intentionally a clean repository rather than a branch or fork.
Only the selected task definitions, their declared source closure, the shared
accounting system of record, and reachable verifier code are copied.
"""

from __future__ import annotations

import argparse
import ast
import dataclasses
import hashlib
import importlib.util
import json
import os
import re
import shutil
import sys
from collections import Counter, defaultdict, deque
from pathlib import Path
from typing import Any, Iterable


CANONICAL_COMMIT = "0202b6e1bf789a20d3770335922e4ee34f231408"
SELECTED_TASK_IDS = (
    "task_001",
    "task_004",
    "task_015",
    "task_027",
    "task_035",
    "task_037",
    "task_055",
    "task_061",
    "task_068",
    "task_072",
    "task_073",
    "task_100",
)
SELECTED_TASK_SET = frozenset(SELECTED_TASK_IDS)
TASK_ID_PATTERN = re.compile(r"task_\d{3}")
FULL_TASK_ID_PATTERN = re.compile(r"^task_\d{3}$")

APEX_TASK_IDS = frozenset({"task_001", "task_004", "task_015"})
CORPORATE_TASK_IDS = frozenset(SELECTED_TASK_SET - APEX_TASK_IDS)

TARGET_OVERRIDES = {
    "task_001": (
        "Shared/Finance/Close/2026/06 June/4 WIP/"
        "ARM-2409 June WIP controller sign-off - WORKING.docx"
    ),
}

GRADING_REVISIONS = {
    "task_001": "weighted-atomic-hybrid-v4",
    "task_004": "weighted-atomic-hybrid-v4",
    "task_015": (
        "weighted-atomic-hybrid-v7-covenant-slide-presentation-quality"
    ),
    "task_027": "weighted-atomic-hybrid-v12-covenant-release-decision",
    "task_035": "weighted-atomic-hybrid-v12-executive-recovery-decision",
    "task_037": "weighted-atomic-hybrid-v7-portfolio-resilience-decision",
    "task_055": "weighted-atomic-hybrid-v6-transaction-sensitivity-release",
    "task_061": "weighted-atomic-hybrid-v6-tax-close-journal-release",
    "task_068": "weighted-atomic-hybrid-v7-guidance-mitigation-release",
    "task_072": "weighted-atomic-hybrid-v5-deterministic-financial-narrative",
    "task_073": "weighted-atomic-hybrid-v5-pro-forma-credit-capacity",
    "task_100": "weighted-atomic-hybrid-v6-layout-aware-branch-risk-schedules",
}


def _json_dump(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            payload,
            indent=2,
            sort_keys=True,
            ensure_ascii=False,
            allow_nan=False,
        )
        + "\n",
        encoding="utf-8",
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def _load_catalog(source: Path) -> list[Any]:
    module_path = source / "task_catalog.py"
    spec = importlib.util.spec_from_file_location(
        "alder_sample_canonical_task_catalog", module_path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load canonical catalog: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.path.insert(0, str(source))
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path.remove(str(source))
        sys.modules.pop(spec.name, None)
    selected = [task for task in module.TASKS if task.task_id in SELECTED_TASK_SET]
    if tuple(task.task_id for task in selected) != SELECTED_TASK_IDS:
        raise RuntimeError(
            "Canonical catalog does not contain the locked sample in numeric order"
        )
    return selected


def _render_catalog(tasks: list[Any]) -> str:
    rows = []
    for task in tasks:
        values = dataclasses.asdict(task)
        rows.append(
            "    TaskSpec(\n"
            f"        task_id={values['task_id']!r},\n"
            f"        slug={values['slug']!r},\n"
            f"        title={values['title']!r},\n"
            f"        workflow={values['workflow']!r},\n"
            f"        output_mode={values['output_mode']!r},\n"
            f"        prompt={values['prompt']!r},\n"
            f"        difficulty={values['difficulty']!r},\n"
            "    ),"
        )
    return (
        "from __future__ import annotations\n\n"
        "from dataclasses import dataclass\n\n\n"
        "@dataclass(frozen=True)\n"
        "class TaskSpec:\n"
        "    task_id: str\n"
        "    slug: str\n"
        "    title: str\n"
        "    workflow: str\n"
        "    output_mode: str\n"
        "    prompt: str\n"
        "    difficulty: str = 'advanced'\n\n\n"
        "TASKS: tuple[TaskSpec, ...] = (\n"
        + "\n".join(rows)
        + "\n)\n\n"
        "TASK_BY_ID = {task.task_id: task for task in TASKS}\n"
        "TASK_BY_SLUG = {task.slug: task for task in TASKS}\n"
    )


def _task_ids(node: ast.AST | None) -> set[str]:
    if node is None:
        return set()
    return {
        candidate.value
        for candidate in ast.walk(node)
        if isinstance(candidate, ast.Constant)
        and isinstance(candidate.value, str)
        and FULL_TASK_ID_PATTERN.fullmatch(candidate.value)
    }


class _SelectedTaskTransformer(ast.NodeTransformer):
    """Remove branches and named helpers that belong only to other tasks."""

    def visit_If(self, node: ast.If):  # noqa: N802
        referenced = _task_ids(node.test)
        if referenced and not referenced.intersection(SELECTED_TASK_SET):
            return [self.visit(child) for child in node.orelse]
        return self.generic_visit(node)

    def visit_IfExp(self, node: ast.IfExp):  # noqa: N802
        referenced = _task_ids(node.test)
        if referenced and not referenced.intersection(SELECTED_TASK_SET):
            return self.visit(node.orelse)
        return self.generic_visit(node)

    def visit_FunctionDef(self, node: ast.FunctionDef):  # noqa: N802
        # Reachability pruning below removes unused task-specific helpers. A
        # few historically named helpers were later generalized and are used
        # by selected tasks, so do not discard definitions by name alone.
        return self.generic_visit(node)

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_Dict(self, node: ast.Dict):  # noqa: N802
        node = self.generic_visit(node)
        keys: list[ast.expr | None] = []
        values: list[ast.expr] = []
        for key, value in zip(node.keys, node.values):
            referenced = _task_ids(key) | _task_ids(value)
            if referenced and not referenced.intersection(SELECTED_TASK_SET):
                continue
            keys.append(key)
            values.append(value)
        node.keys = keys
        node.values = values
        return node

    def _visit_sequence(self, node: ast.List | ast.Set | ast.Tuple):
        node = self.generic_visit(node)
        node.elts = [
            element
            for element in node.elts
            if not (
                isinstance(element, ast.Constant)
                and (_task_ids(element) - SELECTED_TASK_SET)
            )
        ]
        return node

    visit_List = _visit_sequence
    visit_Set = _visit_sequence
    visit_Tuple = _visit_sequence


def _defined_names(node: ast.stmt) -> list[str]:
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        return [node.name]
    if isinstance(node, ast.Assign):
        return [
            candidate.id
            for target in node.targets
            for candidate in ast.walk(target)
            if isinstance(candidate, ast.Name)
            and isinstance(candidate.ctx, ast.Store)
        ]
    if isinstance(node, ast.AnnAssign):
        return [
            candidate.id
            for candidate in ast.walk(node.target)
            if isinstance(candidate, ast.Name)
            and isinstance(candidate.ctx, ast.Store)
        ]
    return []


def _replace_function(tree: ast.Module, name: str, source: str) -> None:
    replacement = ast.parse(source).body[0]
    tree.body = [
        replacement
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name == name
        else node
        for node in tree.body
    ]


def _reachable_module(
    source_path: Path,
    roots: Iterable[str],
    *,
    function_replacements: dict[str, str] | None = None,
) -> str:
    tree = ast.parse(source_path.read_text(encoding="utf-8"))
    tree = _SelectedTaskTransformer().visit(tree)
    ast.fix_missing_locations(tree)
    for name, replacement in (function_replacements or {}).items():
        _replace_function(tree, name, replacement)
    ast.fix_missing_locations(tree)

    definitions: dict[str, list[ast.stmt]] = defaultdict(list)
    for node in tree.body:
        for name in _defined_names(node):
            definitions[name].append(node)

    wanted = deque(roots)
    visited: set[str] = set()
    keep_nodes: set[int] = set()
    while wanted:
        name = wanted.popleft()
        if name in visited:
            continue
        visited.add(name)
        for node in definitions.get(name, []):
            keep_nodes.add(id(node))
            for candidate in ast.walk(node):
                if (
                    isinstance(candidate, ast.Name)
                    and isinstance(candidate.ctx, ast.Load)
                    and candidate.id in definitions
                    and candidate.id not in visited
                ):
                    wanted.append(candidate.id)

    tree.body = [
        node
        for node in tree.body
        if isinstance(node, (ast.Import, ast.ImportFrom))
        or id(node) in keep_nodes
        or (
            isinstance(node, ast.Expr)
            and isinstance(node.value, ast.Constant)
            and isinstance(node.value.value, str)
        )
    ]
    ast.fix_missing_locations(tree)
    rendered = ast.unparse(tree) + "\n"
    leaked = sorted(set(TASK_ID_PATTERN.findall(rendered)) - SELECTED_TASK_SET)
    # Some reusable parsers retain an old task-number prefix in their Python
    # identifier even though the implementation is generic. Rename only
    # those reachable identifier/docstring remnants; task branches and data
    # rows for excluded tasks were already removed above.
    for task_id in leaked:
        neutral = "shared_helper_" + hashlib.sha256(
            task_id.encode("utf-8")
        ).hexdigest()[:8]
        rendered = rendered.replace(task_id, neutral)
        rendered = rendered.replace(task_id.upper(), neutral.upper())
    leaked = sorted(set(TASK_ID_PATTERN.findall(rendered)) - SELECTED_TASK_SET)
    if leaked:
        raise RuntimeError(f"Task-id leakage in {source_path.name}: {leaked}")
    compile(rendered, str(source_path), "exec")
    return rendered


APEX_ENTRY = '''\
def grade_apex_task(task_id: str, answer: Any, workspace_root: str | Path) -> dict[str, Any]:
    if task_id == "task_001":
        result = _grade_task_001(Path(workspace_root), answer)
    elif task_id == "task_004":
        result = _canonicalize_legacy_file_policy(
            task_id, _grade_task_004(Path(workspace_root))
        )
    elif task_id == "task_015":
        result = _canonicalize_legacy_file_policy(
            task_id, _grade_task_015(Path(workspace_root))
        )
    elif task_id in CORPORATE_TASK_IDS:
        from graders.corporate_finance import grade_corporate_finance_task
        result = grade_corporate_finance_task(task_id, answer, workspace_root)
    else:
        raise KeyError(f"Task is not part of this sample: {task_id}")
    if task_id in TASK_GRADING_REVISIONS:
        result.setdefault(
            "task_grading_revision", dict(TASK_GRADING_REVISIONS[task_id])
        )
    return apply_reward_policy(task_id, attach_default_policy(result))
'''


def _patch_apex_header(rendered: str) -> str:
    marker = "from graders.rubric import"
    if marker not in rendered:
        raise RuntimeError("Pruned apex module lost rubric imports")
    insertion = "CORPORATE_TASK_IDS = frozenset(" + repr(sorted(CORPORATE_TASK_IDS)) + ")\n\n"
    lines = rendered.splitlines(keepends=True)
    position = 0
    while position < len(lines) and (
        lines[position].startswith("from __future__")
        or lines[position].strip() == ""
        or lines[position].startswith("import ")
        or lines[position].startswith("from ")
        or lines[position].startswith("    ")
        or lines[position].rstrip().endswith("(")
    ):
        position += 1
    lines.insert(position, insertion)
    return "".join(lines)


def _write_pruned_graders(source: Path, output: Path) -> None:
    grader_output = output / "graders"
    grader_output.mkdir(parents=True, exist_ok=True)
    for name in ("__init__.py", "common.py", "hybrid_semantic.py", "semantic.py"):
        _copy(source / "graders" / name, grader_output / name)

    apex = _reachable_module(
        source / "graders" / "apex.py",
        roots=(
            "grade_apex_task",
            "load_apex_gold",
            "TASK_GRADING_REVISIONS",
        ),
        function_replacements={"grade_apex_task": APEX_ENTRY},
    )
    apex = _patch_apex_header(apex)
    (grader_output / "apex.py").write_text(apex, encoding="utf-8")

    modules = {
        "corporate_finance.py": (
            "grade_corporate_finance_task",
            "load_corporate_finance_gold",
            "_semantic_evidence_pack",
        ),
        "integrity.py": (
            "WorkspaceCreationMonitor",
            "assess_integrity",
            "capture_integrity_snapshot",
        ),
        "production_semantic.py": (
            "submission_integrity_evidence",
            "verify_semantic_review",
        ),
        "rubric.py": ("apply_reward_policy", "attach_default_policy"),
    }
    for name, roots in modules.items():
        rendered = _reachable_module(source / "graders" / name, roots)
        (grader_output / name).write_text(rendered, encoding="utf-8")


def _write_gold(source: Path, output: Path) -> dict[str, Any]:
    first = json.loads(
        (source / "graders" / "gold" / "tasks_001_025.json").read_text()
    )
    later = json.loads(
        (source / "graders" / "gold" / "tasks_026_100.json").read_text()
    )
    selected = {task_id: (first | later)[task_id] for task_id in SELECTED_TASK_IDS}
    _json_dump(
        output / "graders" / "gold" / "tasks_001_025.json",
        {task_id: selected[task_id] for task_id in SELECTED_TASK_IDS if task_id in APEX_TASK_IDS},
    )
    _json_dump(
        output / "graders" / "gold" / "tasks_026_100.json",
        {task_id: selected[task_id] for task_id in SELECTED_TASK_IDS if task_id in CORPORATE_TASK_IDS},
    )
    return selected


def _write_sources(
    source: Path,
    output: Path,
    tasks: list[Any],
    gold: dict[str, Any],
) -> tuple[set[str], dict[str, Any]]:
    seed = source / "world" / "seed"
    destination_seed = output / "world" / "seed"
    dependencies = json.loads(
        (seed / "control" / "task_source_dependencies.json").read_text()
    )
    selected_dependencies = {
        **{key: value for key, value in dependencies.items() if key != "tasks"},
        "tasks": {
            task_id: dependencies["tasks"][task_id]
            for task_id in SELECTED_TASK_IDS
        },
    }
    _json_dump(
        destination_seed / "control" / "task_source_dependencies.json",
        selected_dependencies,
    )

    included: set[str] = set()
    for task_id, payload in selected_dependencies["tasks"].items():
        included.update(payload["minimum_source_artifacts"])
        if task_id in TARGET_OVERRIDES:
            included.add(TARGET_OVERRIDES[task_id])
        artifact = gold[task_id].get("artifact")
        if isinstance(artifact, dict) and artifact.get("edit"):
            included.add(artifact["path"])

    for relative in sorted(included):
        source_file = seed / "workspace" / relative
        if not source_file.is_file():
            raise FileNotFoundError(f"Selected source artifact is missing: {relative}")
        _copy(source_file, destination_seed / "workspace" / relative)
    _copy(seed / "accounting_seed.db", destination_seed / "accounting_seed.db")

    registry = json.loads((seed / "control" / "file_registry.json").read_text())
    by_path = {row["path"]: row for row in registry["files"]}
    missing_registry = sorted(included - by_path.keys())
    if missing_registry:
        raise RuntimeError(f"Selected artifacts missing from registry: {missing_registry}")
    selected_rows = []
    for path in sorted(included):
        row = dict(by_path[path])
        packaged = destination_seed / "workspace" / path
        row["bytes"] = packaged.stat().st_size
        row["sha256"] = _sha256(packaged)
        selected_rows.append(row)
    selected_registry = {
        **{key: value for key, value in registry.items() if key != "files"},
        "files": selected_rows,
    }
    _json_dump(
        destination_seed / "control" / "file_registry.json",
        selected_registry,
    )

    source_map = json.loads(
        (seed / "control" / "corporate_finance_source_map.json").read_text()
    )
    _json_dump(
        destination_seed / "control" / "corporate_finance_source_map.json",
        {
            **{key: value for key, value in source_map.items() if key != "tasks"},
            "tasks": {
                task_id: source_map["tasks"][task_id]
                for task_id in SELECTED_TASK_IDS
                if task_id in source_map["tasks"]
            },
        },
    )

    template_inputs = json.loads(
        (seed / "control" / "corporate_finance_template_inputs.json").read_text()
    )
    _json_dump(
        destination_seed / "control" / "corporate_finance_template_inputs.json",
        {
            task_id: template_inputs[task_id]
            for task_id in SELECTED_TASK_IDS
            if task_id in template_inputs
        },
    )

    # The canonical shared-workbook build cache contains authoring records for
    # the entire 100-task environment. It is not required by the runtime or by
    # any selected verifier, so deliberately omit it from this clean sample.
    (destination_seed / "control" / "corporate_finance_shared_workbook_data.json").unlink(
        missing_ok=True
    )

    inventory = Counter(Path(path).suffix.lower().lstrip(".") or "text" for path in included)
    manifest = {
        "world_id": "alder-ridge-mechanical-sample-v1",
        "snapshot": "2026-06-30-pre-close",
        "canonical_source_commit": CANONICAL_COMMIT,
        "sample_task_count": len(SELECTED_TASK_IDS),
        "shared_accounting_system_included": True,
        "agent_workspace": "/workspace",
        "mutable_deliverables_root": "/workspace/Deliverables",
        "accounting_agent_interface": (
            "Company accounting tools exposed through the contractor_accounting MCP; "
            "raw SQLite storage is not agent-facing at runtime."
        ),
        "source_inventory": dict(sorted(inventory.items())),
        "source_file_count": len(included),
        "notes": (
            "This manifest contains only the declared source closure for the twelve "
            "sample tasks. The shared accounting snapshot is retained unchanged so "
            "the selected tasks reproduce the same system-of-record facts as the "
            "canonical environment."
        ),
    }
    _json_dump(destination_seed / "source_manifest.json", manifest)
    return included, selected_dependencies


def _write_task_packages(
    output: Path,
    tasks: list[Any],
    gold: dict[str, Any],
    dependencies: dict[str, Any],
) -> None:
    for task in tasks:
        root = output / "tasks" / task.slug
        reference = gold[task.task_id]
        artifact = reference.get("artifact")
        metadata = {
            **dataclasses.asdict(task),
            "grading_revision": GRADING_REVISIONS[task.task_id],
            "artifact": artifact,
            "canonical_source_commit": CANONICAL_COMMIT,
        }
        _json_dump(root / "task.json", metadata)
        (root / "prompt.md").write_text(
            f"# {task.title}\n\n{task.prompt.strip()}\n",
            encoding="utf-8",
        )
        _json_dump(root / "gold.json", reference)
        rubric = {
            "task_id": task.task_id,
            "slug": task.slug,
            "grading_revision": GRADING_REVISIONS[task.task_id],
            "reward_type": "weighted_partial_credit",
            "strict_pass": "all atomic criteria equal 1",
            "deterministic_verifier": (
                "graders/apex.py"
                if task.task_id in APEX_TASK_IDS
                else "graders/corporate_finance.py"
            ),
            "semantic_verifier": "graders/production_semantic.py",
            "integrity_verifier": "graders/integrity.py",
            "reward_policy": "graders/rubric.py",
            "criteria": reference.get("criteria"),
            "reference_contract": (
                None if reference.get("criteria") else reference
            ),
        }
        _json_dump(root / "rubric.json", rubric)
        _json_dump(
            root / "source-dependencies.json",
            dependencies["tasks"][task.task_id],
        )


def _write_runtime_files(source: Path, output: Path) -> None:
    for relative in (
        "alder_ridge_world/__init__.py",
        "alder_ridge_world/accounting.py",
        "alder_ridge_world/export_provenance.py",
        "alder_ridge_world/mcp_server.py",
        "alder_ridge_world/paths.py",
        "alder_ridge_world/reset.py",
        "alder_ridge_world/workspace_security.py",
        "scripts/prepare_agent_runtime.py",
    ):
        _copy(source / relative, output / relative)

    task_templates = _reachable_module(
        source / "task_templates.py",
        roots=("register_task_templates",),
        function_replacements={
            "_workspace_slug": '''\
def _workspace_slug(task_id: str, declared_slug: str | None = None) -> str | None:
    return declared_slug
''',
            "prepare_task_workspace": '''\
def prepare_task_workspace(
    task_id: str,
    runtime_root: Path,
    *,
    task_slug: str | None = None,
    seed_root: Path | None = None,
) -> None:
    """The selected sample uses the shared canonical workspace without overlays."""
    return None
''',
        },
    ).replace(
        "Register 100 independently runnable APEX-style tasks on one company world.",
        "Register twelve independently runnable tasks on one company world.",
    )
    (output / "task_templates.py").write_text(task_templates, encoding="utf-8")

    env = (source / "env.py").read_text(encoding="utf-8").replace(
        'DEFAULT_ENVIRONMENT_NAME = "alder-ridge-corporate-finance"',
        'DEFAULT_ENVIRONMENT_NAME = "sample-alder-ridge-corporate-finance-environment"',
    )
    (output / "env.py").write_text(env, encoding="utf-8")

    tasks_source = (
        "from __future__ import annotations\n\n"
        "from env import TASK_TEMPLATES\n"
        "from task_catalog import TASKS\n\n\n"
        f"GRADING_REVISIONS = {GRADING_REVISIONS!r}\n\n"
        "tasks = []\n"
        "for spec in TASKS:\n"
        "    task = TASK_TEMPLATES[spec.task_id]()\n"
        "    task.slug = spec.slug\n"
        "    task.columns = {\n"
        "        'world': 'alder-ridge-mechanical',\n"
        "        'workflow': spec.workflow,\n"
        "        'level': 'senior-finance-analyst',\n"
        "        'output_mode': spec.output_mode,\n"
        "        'difficulty': spec.difficulty,\n"
        "        'snapshot': '2026-06-30-pre-close',\n"
        "        'grading': GRADING_REVISIONS[spec.task_id],\n"
        "    }\n"
        "    tasks.append(task)\n\n"
        "del task, spec\n"
    )
    (output / "tasks.py").write_text(tasks_source, encoding="utf-8")


def _verify_scope(output: Path, included_sources: set[str]) -> dict[str, Any]:
    forbidden_ids: dict[str, list[str]] = {}
    scan_suffixes = {".py", ".json", ".md", ".toml", ".yml", ".yaml", ".txt"}
    for path in output.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in scan_suffixes:
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        leaked = sorted(set(TASK_ID_PATTERN.findall(text)) - SELECTED_TASK_SET)
        if leaked:
            forbidden_ids[path.relative_to(output).as_posix()] = leaked
    if forbidden_ids:
        raise RuntimeError(f"Non-sample task identifiers leaked: {forbidden_ids}")

    catalog_path = output / "task_catalog.py"
    catalog_text = catalog_path.read_text(encoding="utf-8")
    catalog_ids = tuple(TASK_ID_PATTERN.findall(catalog_text))
    if tuple(dict.fromkeys(catalog_ids)) != SELECTED_TASK_IDS:
        raise RuntimeError("Generated catalog does not contain exactly the locked tasks")

    actual_sources = {
        path.relative_to(output / "world" / "seed" / "workspace").as_posix()
        for path in (output / "world" / "seed" / "workspace").rglob("*")
        if path.is_file()
    }
    if actual_sources != included_sources:
        raise RuntimeError("Workspace source closure differs from the declared sample")

    payload = {
        "schema_version": 1,
        "canonical_source_commit": CANONICAL_COMMIT,
        "selected_task_ids": list(SELECTED_TASK_IDS),
        "task_count": len(SELECTED_TASK_IDS),
        "source_file_count": len(actual_sources),
        "source_files": sorted(actual_sources),
        "accounting_seed_sha256": _sha256(
            output / "world" / "seed" / "accounting_seed.db"
        ),
        "non_sample_task_identifiers_found": [],
        "scope_verified": True,
    }
    _json_dump(output / "SAMPLE_SCOPE_MANIFEST.json", payload)
    return payload


def package(source: Path, output: Path) -> dict[str, Any]:
    source = source.resolve()
    output = output.resolve()
    if not (source / "task_catalog.py").is_file():
        raise FileNotFoundError(f"Canonical source tree is invalid: {source}")
    output.mkdir(parents=True, exist_ok=True)

    tasks = _load_catalog(source)
    (output / "task_catalog.py").write_text(
        _render_catalog(tasks), encoding="utf-8"
    )
    gold = _write_gold(source, output)
    _write_pruned_graders(source, output)
    _write_runtime_files(source, output)
    included, dependencies = _write_sources(source, output, tasks, gold)
    _write_task_packages(output, tasks, gold, dependencies)
    return _verify_scope(output, included)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = package(args.source, args.output)
    print(json.dumps(payload, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
