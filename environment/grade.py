#!/usr/bin/env python3
"""Run the exact sample grader against a response and/or artifact workspace."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from runtime.grading.apex import grade_apex_task
from runtime.task_catalog import TASK_BY_ID, TASK_BY_SLUG


ROOT = Path(__file__).resolve().parent


def _answer(args: argparse.Namespace) -> Any:
    if args.answer_file is None:
        return args.answer or ""
    text = args.answer_file.read_text(encoding="utf-8")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--task", required=True, help="Task ID or slug")
    parser.add_argument("--answer")
    parser.add_argument("--answer-file", type=Path)
    parser.add_argument(
        "--workspace",
        type=Path,
        default=ROOT / "seed" / "sources",
    )
    args = parser.parse_args()
    spec = TASK_BY_ID.get(args.task) or TASK_BY_SLUG.get(args.task)
    if spec is None:
        parser.error(f"Task is not part of this sample: {args.task}")
    result = grade_apex_task(spec.task_id, _answer(args), args.workspace)
    print(json.dumps(result, indent=2, sort_keys=True, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
