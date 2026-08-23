from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

from graders.apex import grade_apex_task, load_apex_gold
from graders.corporate_finance import load_corporate_finance_gold
from task_catalog import TASKS


ROOT = Path(__file__).resolve().parents[1]
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


def test_catalog_and_gold_are_exactly_the_selected_sample() -> None:
    assert tuple(task.task_id for task in TASKS) == SELECTED
    assert set(load_apex_gold()) == set(SELECTED[:3])
    assert set(load_corporate_finance_gold()) == set(SELECTED[3:])


def test_declared_source_closure_is_complete_and_hash_bound() -> None:
    seed = ROOT / "world" / "seed"
    dependencies = json.loads(
        (seed / "control" / "task_source_dependencies.json").read_text()
    )
    registry = json.loads((seed / "control" / "file_registry.json").read_text())
    assert set(dependencies["tasks"]) == set(SELECTED)
    for row in registry["files"]:
        path = seed / "workspace" / row["path"]
        assert path.is_file(), row["path"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == row["sha256"]


def test_machine_scope_manifest_matches_disk() -> None:
    manifest = json.loads((ROOT / "SAMPLE_SCOPE_MANIFEST.json").read_text())
    assert tuple(manifest["selected_task_ids"]) == SELECTED
    assert manifest["task_count"] == 12
    assert manifest["scope_verified"] is True
    actual = {
        path.relative_to(ROOT / "world" / "seed" / "workspace").as_posix()
        for path in (ROOT / "world" / "seed" / "workspace").rglob("*")
        if path.is_file()
    }
    assert actual == set(manifest["source_files"])


def test_no_other_task_identifier_is_exposed() -> None:
    selected = set(SELECTED)
    leaked: dict[str, list[str]] = {}
    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts:
            continue
        if path.suffix.lower() not in {".py", ".json", ".md", ".toml", ".yaml"}:
            continue
        found = set(re.findall(r"task_\d{3}", path.read_text(errors="ignore")))
        unexpected = sorted(found - selected)
        if unexpected:
            leaked[path.relative_to(ROOT).as_posix()] = unexpected
    assert leaked == {}


def test_console_oracle_earns_full_credit() -> None:
    gold = load_corporate_finance_gold("task_073")
    result = grade_apex_task(
        "task_073",
        gold["answer"],
        ROOT / "world" / "seed" / "workspace",
    )
    assert result["reward"] == 1.0
    assert result["strict_pass"] is True
    assert result["criteria_met"] == result["criteria_total"] == 24


def test_selected_artifact_graders_return_weighted_atomic_results() -> None:
    for task_id in ("task_001", "task_004", "task_015", "task_027"):
        result = grade_apex_task(
            task_id,
            "",
            ROOT / "world" / "seed" / "workspace",
        )
        assert result["criteria_total"] > 0
        assert result["weight_total"] > 0
        assert 0.0 <= result["reward"] <= 1.0
        assert len(result["criteria"]) == result["criteria_total"]
