#!/usr/bin/env python3
"""Validate the exact five-task Alder Ridge sample without external services."""

from __future__ import annotations

import ast
import hashlib
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from pypdf import PdfReader

from runtime.accounting_mcp.accounting import AccountingRepository
from runtime.grading.apex import (
    SAMPLE_TASK_IDS,
    TASK_GRADING_REVISIONS,
    grade_apex_task,
    load_apex_gold,
)
from runtime.grading.corporate_finance import (
    SAMPLE_CORPORATE_TASK_IDS,
    load_corporate_finance_gold,
)
from runtime.grading.hybrid_semantic import HYBRID_TASKS
from runtime.task_catalog import TASKS


ENVIRONMENT_ROOT = Path(__file__).resolve().parent
REPOSITORY_ROOT = ENVIRONMENT_ROOT.parent
SEED_ROOT = ENVIRONMENT_ROOT / "seed"
SOURCES_ROOT = SEED_ROOT / "sources"
CONTROLS_ROOT = SEED_ROOT / "controls"
TASKS_ROOT = REPOSITORY_ROOT / "tasks"

SELECTED = (
    "task_001",
    "task_004",
    "task_015",
    "task_035",
    "task_068",
)
EXPECTED = {
    "task_001": {
        "slug": "arm-2409-wip-backcharge",
        "prompt_sha256": "2d532c8ddb41e1d16ead95a58966655166c8560258c2654f0635d44869e974f2",
        "grading_revision": "weighted-atomic-hybrid-v53-proportional-integrity",
        "catalog_revision": "task-001-conditional-credit-effectiveness-v44",
        "criteria": 94,
        "semantic": 87,
        "weight": 720,
    },
    "task_004": {
        "slug": "complete-june-wip-risk-template",
        "prompt_sha256": "a840cb78e51fcaab2076ac996e960c482188287c61194eefc2cbcf26a61bd223",
        "grading_revision": "weighted-atomic-hybrid-v20-proportional-integrity",
        "catalog_revision": "task-004-close-status-release-judgment-v16",
        "criteria": 98,
        "semantic": 98,
        "weight": 927,
    },
    "task_015": {
        "slug": "append-q2-covenant-slide",
        "prompt_sha256": "554e4c82f7aa37896ed79e85858695c908e80e4aa4c4e3fc61673734d2f4833a",
        "grading_revision": "weighted-atomic-hybrid-v20-proportional-integrity-and-prompt-boundary",
        "catalog_revision": "task-015-current-close-covenant-v15",
        "criteria": 74,
        "semantic": 65,
        "weight": 518,
    },
    "task_035": {
        "slug": "complete-backlog-capacity-model",
        "prompt_sha256": "4e40f6100f92c1026099662397793453ec6b37ee7225393f8c7c1bcf5dee6cb6",
        "grading_revision": "weighted-atomic-hybrid-v26-preferred-sheet-scenario-lineage",
        "catalog_revision": "task-035-authentic-backlog-capacity-decision-v15",
        "criteria": 130,
        "semantic": 106,
        "weight": 807,
    },
    "task_068": {
        "slug": "complete-executive-performance-deck",
        "prompt_sha256": "c195103cee0a52c1dfd3a0fa84098d72b2725672c5067171c62cdf323fbf45ae",
        "grading_revision": "weighted-atomic-hybrid-v30-recursive-ppt-evidence",
        "catalog_revision": "task-068-authentic-executive-performance-decision-v21",
        "criteria": 88,
        "semantic": 87,
        "weight": 628,
    },
}
EXPECTED_ACCOUNTING_SHA256 = (
    "00ddfe914af7abb109a05f703623f43c241ac5f90d73e0f4f5e46d9459b76e7d"
)
EXPECTED_CANONICAL_SOURCE_COMMIT = (
    "d03ed15caa1f66008ab767fbd203f656a32425eb"
)
EXPECTED_SOURCE_COUNT = 144
TASK_FILE_SET = {
    "gold.json",
    "prompt.md",
    "rubric.json",
    "source_manifest.json",
    "task.json",
}
CONTROL_FILE_SET = {
    "corporate_finance_shared_workbook_data.json",
    "corporate_finance_source_map.json",
    "corporate_finance_template_inputs.json",
    "fact_registry.json",
    "file_registry.json",
    "task_source_dependencies.json",
}
FORBIDDEN_SOURCE_LABELS = (
    "controlled finance document",
    "internal - finance and operations",
    "internal — finance and operations",
)
PROMPT_LEAK_TERMS = (
    "hidden trick",
    "hidden exception",
    "gotcha",
    "stump the agent",
    "grading rubric",
)


def _json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _literal_assignment(path: Path, name: str) -> Any:
    module = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in module.body:
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            if any(isinstance(target, ast.Name) and target.id == name for target in targets):
                return ast.literal_eval(node.value)
    raise AssertionError(f"Missing literal assignment {name} in {path}")


def _office_text(path: Path) -> str:
    chunks: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.casefold().endswith(".xml"):
                continue
            try:
                chunks.extend(ElementTree.fromstring(archive.read(name)).itertext())
            except ElementTree.ParseError:
                continue
    return " ".join(chunks)


def _searchable_text(path: Path) -> str:
    suffix = path.suffix.casefold()
    if suffix in {".docx", ".pptx", ".xlsx"}:
        return _office_text(path)
    if suffix == ".pdf":
        try:
            return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages)
        except Exception:
            return ""
    if suffix in {".csv", ".eml", ".md", ".txt"}:
        return path.read_text(encoding="utf-8", errors="ignore")
    return ""


def _criterion_contract(rows: list[dict[str, Any]]) -> list[tuple[Any, ...]]:
    return [
        (
            row.get("id"),
            row.get("category"),
            row.get("weight"),
            bool(row.get("semantic")),
        )
        for row in rows
    ]


def main() -> int:
    expected_ids = frozenset(SELECTED)
    manifest = _json(SEED_ROOT / "sample_manifest.json")
    assert {
        path.name for path in CONTROLS_ROOT.iterdir() if path.is_file()
    } == CONTROL_FILE_SET
    registry = _json(CONTROLS_ROOT / "file_registry.json")
    dependencies = _json(CONTROLS_ROOT / "task_source_dependencies.json")

    assert tuple(manifest["selected_task_ids"]) == SELECTED
    assert manifest["task_count"] == len(SELECTED) == 5
    assert manifest["canonical_source_commit"] == EXPECTED_CANONICAL_SOURCE_COMMIT
    assert tuple(task.task_id for task in TASKS) == SELECTED
    assert SAMPLE_TASK_IDS == HYBRID_TASKS == expected_ids
    assert SAMPLE_CORPORATE_TASK_IDS == frozenset({"task_035", "task_068"})
    assert set(load_apex_gold()) == set(SELECTED[:3])
    assert set(load_corporate_finance_gold()) == set(SELECTED[3:])
    assert set(dependencies["tasks"]) == expected_ids
    assert set(_json(CONTROLS_ROOT / "corporate_finance_source_map.json")["tasks"]) == {
        "task_035",
        "task_068",
    }
    assert set(_json(CONTROLS_ROOT / "corporate_finance_template_inputs.json")) == {
        "task_035"
    }
    assert set(
        _json(CONTROLS_ROOT / "corporate_finance_shared_workbook_data.json")[
            "task_records"
        ]
    ) == {"task_035", "task_068"}

    assert manifest["complete_shared_company_world_included"] is True
    assert manifest["container_includes_complete_seed_world"] is True
    assert manifest["other_task_definitions_included"] is False
    assert manifest["non_sample_task_seed_overlays_included"] is False
    assert not (SEED_ROOT / "private_task_overlays").exists()

    actual_sources = {
        path.relative_to(SOURCES_ROOT).as_posix(): path
        for path in SOURCES_ROOT.rglob("*")
        if path.is_file() and path.name != ".DS_Store"
    }
    registered_sources = {row["path"]: row for row in registry["files"]}
    assert len(actual_sources) == EXPECTED_SOURCE_COUNT
    assert len(registered_sources) == EXPECTED_SOURCE_COUNT
    assert set(actual_sources) == set(registered_sources) == set(manifest["source_files"])
    assert manifest["source_file_count"] == EXPECTED_SOURCE_COUNT
    for relative, source in actual_sources.items():
        assert _sha256(source) == registered_sources[relative]["sha256"], relative

    accounting_sha = _sha256(SEED_ROOT / "accounting.db")
    assert accounting_sha == EXPECTED_ACCOUNTING_SHA256
    assert accounting_sha == manifest["accounting_seed_sha256"]

    task_folders = {path.name for path in TASKS_ROOT.iterdir() if path.is_dir()}
    assert task_folders == {task.slug for task in TASKS}
    runtime_gold = {**load_apex_gold(), **load_corporate_finance_gold()}
    task_revision_map = _literal_assignment(
        ENVIRONMENT_ROOT / "runtime" / "tasks.py", "GRADING_REVISIONS"
    )
    assert task_revision_map == {
        task_id: EXPECTED[task_id]["grading_revision"] for task_id in SELECTED
    }

    blank_scores: dict[str, float] = {}
    for task in TASKS:
        expected = EXPECTED[task.task_id]
        folder = TASKS_ROOT / task.slug
        assert {path.name for path in folder.iterdir() if path.is_file()} == TASK_FILE_SET
        metadata = _json(folder / "task.json")
        rubric = _json(folder / "rubric.json")
        source_manifest = _json(folder / "source_manifest.json")
        packaged_gold = _json(folder / "gold.json")

        assert task.slug == expected["slug"]
        assert metadata["task_id"] == task.task_id
        assert metadata["slug"] == task.slug
        assert metadata["canonical_source_commit"] == EXPECTED_CANONICAL_SOURCE_COMMIT
        assert metadata["prompt"] == task.prompt
        assert metadata["prompt_sha256"] == expected["prompt_sha256"]
        assert _sha256_bytes(task.prompt.encode()) == expected["prompt_sha256"]
        assert metadata["grading_revision"] == expected["grading_revision"]
        assert metadata["grading_contract"]["id"] == expected["catalog_revision"]
        assert TASK_GRADING_REVISIONS[task.task_id]["id"] == expected["catalog_revision"]
        assert metadata["grading_contract"] == TASK_GRADING_REVISIONS[task.task_id]
        assert (folder / "prompt.md").read_text(encoding="utf-8") == (
            f"# {task.title}\n\n{task.prompt.rstrip()}\n"
        )
        assert task.prompt.startswith("Please ")
        prompt_folded = task.prompt.casefold()
        assert not any(term in prompt_folded for term in PROMPT_LEAK_TERMS)

        assert rubric["task_id"] == task.task_id
        assert rubric["grading_revision"] == expected["grading_revision"]
        assert rubric["grading_contract"]["id"] == expected["catalog_revision"]
        assert rubric["grading_contract"] == TASK_GRADING_REVISIONS[task.task_id]
        assert rubric["criterion_count"] == len(rubric["criteria"]) == expected["criteria"]
        assert rubric["semantic_criterion_count"] == expected["semantic"]
        assert sum(int(row["weight"]) for row in rubric["criteria"]) == expected["weight"]
        assert len({row["id"] for row in rubric["criteria"]}) == expected["criteria"]
        capped = {
            row["id"]: row["failure_cap"]
            for row in rubric["criteria"]
            if row.get("failure_cap") is not None
        }
        assert capped == {}
        assert packaged_gold == runtime_gold[task.task_id]
        assert source_manifest == dependencies["tasks"][task.task_id]
        for relative in source_manifest["minimum_source_artifacts"]:
            assert relative in actual_sources, (task.task_id, relative)

        result = grade_apex_task(task.task_id, "", SOURCES_ROOT)
        assert result["reward"] == 0.0
        assert result["criteria_total"] == expected["criteria"]
        assert result["weight_total"] == expected["weight"]
        assert _criterion_contract(result["criteria"]) == _criterion_contract(
            rubric["criteria"]
        )
        blank_scores[task.task_id] = result["reward"]

    try:
        grade_apex_task("task_027", "", SOURCES_ROOT)
    except KeyError:
        pass
    else:
        raise AssertionError("A removed task remains gradeable")

    repository = AccountingRepository(SEED_ROOT / "accounting.db")
    ledger_rows = repository.general_ledger(
        "2026-06-01", "2026-06-30", project_id="ARM-2409", limit=1
    )
    job_cost_rows = repository.job_cost_detail(
        "ARM-2409", "2026-06-30", limit=1, start_date="2026-06-01"
    )
    assert ledger_rows and {"journal_id", "journal_line_id"} <= set(ledger_rows[0])
    assert job_cost_rows and "job_cost_entry_id" in job_cost_rows[0]

    forbidden_hits: dict[str, list[str]] = {}
    for relative, source in actual_sources.items():
        normalized = re.sub(r"\s+", " ", _searchable_text(source)).casefold()
        hits = [label for label in FORBIDDEN_SOURCE_LABELS if label in normalized]
        if hits:
            forbidden_hits[relative] = hits
    assert forbidden_hits == {}, forbidden_hits

    print(
        json.dumps(
            {
                "accounting_seed_sha256": accounting_sha,
                "blank_rewards": blank_scores,
                "grader_scope": list(SELECTED),
                "private_task_overlays": False,
                "source_file_count": len(actual_sources),
                "source_registry_exact": True,
                "task_count": len(TASKS),
            },
            indent=2,
            sort_keys=True,
        )
    )
    return 0


def _sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


if __name__ == "__main__":
    raise SystemExit(main())
