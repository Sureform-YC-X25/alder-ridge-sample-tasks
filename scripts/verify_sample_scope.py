#!/usr/bin/env python3
"""Verify the published sample against its machine-readable scope manifest."""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    manifest = json.loads((ROOT / "SAMPLE_SCOPE_MANIFEST.json").read_text())
    selected = set(manifest["selected_task_ids"])
    workspace = ROOT / "world" / "seed" / "workspace"
    actual_sources = {
        path.relative_to(workspace).as_posix()
        for path in workspace.rglob("*")
        if path.is_file()
    }
    assert actual_sources == set(manifest["source_files"])
    accounting = ROOT / "world" / "seed" / "accounting_seed.db"
    assert hashlib.sha256(accounting.read_bytes()).hexdigest() == manifest[
        "accounting_seed_sha256"
    ]

    leaked: dict[str, list[str]] = {}
    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts:
            continue
        if path.suffix.lower() not in {".py", ".json", ".md", ".toml", ".yaml"}:
            continue
        unexpected = sorted(
            set(re.findall(r"task_\d{3}", path.read_text(errors="ignore")))
            - selected
        )
        if unexpected:
            leaked[path.relative_to(ROOT).as_posix()] = unexpected
    assert leaked == {}, leaked
    print(
        json.dumps(
            {
                "scope_verified": True,
                "task_count": len(selected),
                "source_file_count": len(actual_sources),
                "non_sample_task_identifiers_found": [],
                "accounting_seed_sha256": manifest["accounting_seed_sha256"],
            },
            indent=2,
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
