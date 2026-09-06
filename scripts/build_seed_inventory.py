#!/usr/bin/env python3
"""Build or verify the deterministic Alder Ridge seed inventory."""

from __future__ import annotations

import argparse
import hashlib
import json
import sqlite3
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SEED = ROOT / "environment" / "seed"
SOURCES = SEED / "sources"
CONTROLS = SEED / "controls"
OUTPUT = ROOT / "SEED_DATA_INVENTORY.md"


def _json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _size(value: int) -> str:
    if value < 1024:
        return f"{value} B"
    if value < 1024**2:
        return f"{value / 1024:.2f} KiB"
    return f"{value / 1024**2:.2f} MiB"


def _database_counts(path: Path) -> list[tuple[str, int]]:
    with sqlite3.connect(f"file:{path}?mode=ro", uri=True) as connection:
        names = [
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master "
                "WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            )
        ]
        return [
            (name, int(connection.execute(f'SELECT COUNT(*) FROM "{name}"').fetchone()[0]))
            for name in names
        ]


def build() -> str:
    manifest = _json(SEED / "sample_manifest.json")
    registry = _json(CONTROLS / "file_registry.json")
    dependencies = _json(CONTROLS / "task_source_dependencies.json")["tasks"]
    by_path = {row["path"]: row for row in registry["files"]}
    files = [SOURCES / relative for relative in manifest["source_files"]]
    type_counts = Counter(path.suffix.casefold().removeprefix(".") for path in files)
    status_counts = Counter(str(by_path[path.relative_to(SOURCES).as_posix()].get("version_status") or "unspecified") for path in files)
    database = SEED / "accounting.db"
    db_counts = _database_counts(database)

    lines = [
        "# Alder Ridge Seed Data Inventory",
        "",
        "This inventory is generated from the packaged five-task manifest, source registry, and read-only accounting snapshot. Run `python scripts/build_seed_inventory.py --check` to verify it.",
        "",
        "## Summary",
        "",
        "| Measure | Value |",
        "|---|---:|",
        f"| Selected tasks | {manifest['task_count']} |",
        f"| Company-world source files | {len(files)} |",
        f"| Company-world source size | {_size(sum(path.stat().st_size for path in files))} |",
        f"| Accounting database size | {_size(database.stat().st_size)} |",
        f"| Accounting business tables | {len(db_counts)} |",
        f"| Accounting rows | {sum(count for _, count in db_counts):,} |",
        f"| Accounting SHA256 | `{_sha256(database)}` |",
        f"| Source world | `{manifest['source_world_id']}` |",
        f"| Snapshot | `{manifest['snapshot']}` |",
        "",
        "### File types",
        "",
        "| Type | Files |",
        "|---|---:|",
    ]
    lines.extend(f"| `{kind or 'none'}` | {count} |" for kind, count in sorted(type_counts.items()))
    lines.extend(["", "### Source authority/version status", "", "| Status | Files |", "|---|---:|"])
    lines.extend(f"| `{status}` | {count} |" for status, count in sorted(status_counts.items()))

    lines.extend(["", "## Selected task source contracts", "", "| Task | Minimum source artifacts | Accounting MCP |", "|---|---:|---|"])
    for task_id in manifest["selected_task_ids"]:
        record = dependencies[task_id]
        lines.append(
            f"| `{task_id}` | {len(record['minimum_source_artifacts'])} | "
            f"{'required' if record.get('accounting_mcp_required') else 'not required'} |"
        )

    lines.extend(["", "## Accounting database", "", "| Table | Rows |", "|---|---:|"])
    lines.extend(f"| `{name}` | {count:,} |" for name, count in db_counts)
    lines.append(f"| **Total** | **{sum(count for _, count in db_counts):,}** |")

    lines.extend(["", "## Company-world source files", "", "| Path | Type | Size | Status | SHA256 |", "|---|---|---:|---|---|"])
    for path in files:
        relative = path.relative_to(SOURCES).as_posix()
        row = by_path[relative]
        lines.append(
            f"| `{relative}` | `{path.suffix.casefold().removeprefix('.')}` | "
            f"{_size(path.stat().st_size)} | `{row.get('version_status') or 'unspecified'}` | "
            f"`{_sha256(path)}` |"
        )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    rendered = build()
    if args.check:
        if not OUTPUT.is_file() or OUTPUT.read_text(encoding="utf-8") != rendered:
            raise SystemExit("SEED_DATA_INVENTORY.md is stale; rebuild it")
        print("SEED_DATA_INVENTORY.md is current")
        return 0
    OUTPUT.write_text(rendered, encoding="utf-8")
    print(f"Wrote {OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
