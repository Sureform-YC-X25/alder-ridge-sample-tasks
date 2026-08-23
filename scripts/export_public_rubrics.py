#!/usr/bin/env python3
"""Export stable public rubric rows without outcome/evidence from a rollout."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from graders.apex import grade_apex_task
from graders.corporate_finance import load_corporate_finance_gold
from task_catalog import TASKS


ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT / "world" / "seed" / "workspace"
APEX = {"task_001", "task_004", "task_015"}
PUBLIC_FIELDS = (
    "id",
    "description",
    "category",
    "weight",
    "semantic",
    "failure_cap",
)


def _rows(task_id: str) -> list[dict[str, Any]]:
    if task_id in APEX:
        rows = grade_apex_task(task_id, "", WORKSPACE)["criteria"]
    else:
        rows = load_corporate_finance_gold(task_id)["criteria"]
    return [
        {key: row.get(key) for key in PUBLIC_FIELDS if key in row}
        for row in rows
    ]


def main() -> int:
    for task in TASKS:
        path = ROOT / "tasks" / task.slug / "rubric.json"
        payload = json.loads(path.read_text(encoding="utf-8"))
        payload["criteria"] = _rows(task.task_id)
        payload.pop("reference_contract", None)
        path.write_text(
            json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
    print(f"Exported public atomic rubrics for {len(TASKS)} tasks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
