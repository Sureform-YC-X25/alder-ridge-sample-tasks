#!/usr/bin/env python3
"""Validate the complete 12-task sample package without external services."""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

from runtime.grading.apex import grade_apex_task, load_apex_gold
from runtime.grading.corporate_finance import load_corporate_finance_gold
from runtime.task_catalog import TASKS


ENVIRONMENT_ROOT = Path(__file__).resolve().parent
REPOSITORY_ROOT = ENVIRONMENT_ROOT.parent
SEED_ROOT = ENVIRONMENT_ROOT / "seed"
SOURCES_ROOT = SEED_ROOT / "sources"
OVERLAYS_ROOT = SEED_ROOT / "private_task_overlays"
SELECTED = (
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


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    manifest = json.loads((SEED_ROOT / "sample_manifest.json").read_text())
    assert tuple(manifest["selected_task_ids"]) == SELECTED
    assert tuple(task.task_id for task in TASKS) == SELECTED
    assert set(load_apex_gold()) == set(SELECTED[:3])
    assert set(load_corporate_finance_gold()) == set(SELECTED[3:])
    assert manifest["complete_shared_company_world_included"] is True
    assert manifest["container_includes_complete_seed_world"] is True
    assert manifest["other_task_definitions_included"] is False

    actual_sources = {
        path.relative_to(SOURCES_ROOT).as_posix()
        for path in SOURCES_ROOT.rglob("*")
        if path.is_file()
    }
    assert actual_sources == set(manifest["source_files"])
    assert len(actual_sources) == manifest["source_file_count"] == 132
    assert sha256(SEED_ROOT / "accounting.db") == manifest[
        "accounting_seed_sha256"
    ]

    actual_overlays = {
        path.relative_to(OVERLAYS_ROOT).as_posix()
        for path in OVERLAYS_ROOT.rglob("*")
        if path.is_file()
    }
    assert actual_overlays == set(manifest["private_task_overlay_files"])
    assert len(actual_overlays) == manifest["private_task_overlay_file_count"] == 1

    registry = json.loads(
        (SEED_ROOT / "controls" / "file_registry.json").read_text()
    )
    for row in registry["files"]:
        source = SOURCES_ROOT / row["path"]
        assert source.is_file(), row["path"]
        assert sha256(source) == row["sha256"], row["path"]

    task_folders = {path.name for path in (REPOSITORY_ROOT / "tasks").iterdir()}
    assert task_folders == {task.slug for task in TASKS}
    for task in TASKS:
        folder = REPOSITORY_ROOT / "tasks" / task.slug
        assert {
            "prompt.md",
            "task.json",
            "rubric.json",
            "gold.json",
            "source_manifest.json",
        }.issubset({path.name for path in folder.iterdir()})
        metadata = json.loads((folder / "task.json").read_text())
        assert metadata["task_id"] == task.task_id
        assert metadata["slug"] == task.slug
        assert metadata["prompt"].strip() == task.prompt.strip()
        assert task.prompt.strip() in (folder / "prompt.md").read_text()

    leaked: dict[str, list[str]] = {}
    selected = set(SELECTED)
    for path in REPOSITORY_ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts:
            continue
        if path.suffix.lower() not in {".py", ".json", ".md", ".toml", ".yaml"}:
            continue
        unexpected = sorted(
            set(re.findall(r"task_\d{3}", path.read_text(errors="ignore")))
            - selected
        )
        if unexpected:
            leaked[path.relative_to(REPOSITORY_ROOT).as_posix()] = unexpected
    assert leaked == {}, leaked

    console_gold = load_corporate_finance_gold("task_073")
    oracle = grade_apex_task("task_073", console_gold["answer"], SOURCES_ROOT)
    assert oracle["reward"] == 1.0
    assert oracle["strict_pass"] is True
    assert oracle["criteria_met"] == oracle["criteria_total"] == 24

    for task_id in ("task_001", "task_004", "task_015", "task_027"):
        result = grade_apex_task(task_id, "", SOURCES_ROOT)
        assert result["criteria_total"] > 0
        assert result["weight_total"] > 0
        assert 0.0 <= result["reward"] <= 1.0
        assert len(result["criteria"]) == result["criteria_total"]

    print(
        json.dumps(
            {
                "accounting_seed_sha256": manifest["accounting_seed_sha256"],
                "non_sample_task_identifiers_found": [],
                "oracle_reward": oracle["reward"],
                "private_task_overlay_file_count": len(actual_overlays),
                "scope_verified": True,
                "source_file_count": len(actual_sources),
                "task_count": len(TASKS),
            },
            indent=2,
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
