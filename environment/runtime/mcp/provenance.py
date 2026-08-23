from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Mapping


EXPORT_PROVENANCE_FILENAME = "accounting_export_provenance.jsonl"
EXPORT_PROVENANCE_SCHEMA_VERSION = 1


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _workspace_path(workspace_root: Path, relative_path: str) -> tuple[str, Path]:
    normalized = Path(relative_path).as_posix()
    while normalized.startswith("./"):
        normalized = normalized[2:]
    target = (workspace_root / normalized).resolve()
    root = workspace_root.resolve()
    if not normalized or root not in target.parents:
        raise ValueError("accounting export path must stay inside the workspace")
    return normalized, target


def record_accounting_export(
    *,
    state_root: Path,
    workspace_root: Path,
    relative_path: str,
    report: str,
    parameters: Mapping[str, Any],
    row_count: int,
) -> dict[str, Any]:
    """Record successful accounting-export lineage outside the agent workspace."""

    normalized, target = _workspace_path(workspace_root, relative_path)
    if not target.is_file():
        raise FileNotFoundError(target)
    record = {
        "schema_version": EXPORT_PROVENANCE_SCHEMA_VERSION,
        "path": normalized,
        "report": str(report),
        "parameters": dict(parameters),
        "row_count": int(row_count),
        "sha256": _sha256(target),
    }
    provenance_path = state_root / EXPORT_PROVENANCE_FILENAME
    provenance_path.parent.mkdir(parents=True, exist_ok=True)
    with provenance_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, sort_keys=True, separators=(",", ":")) + "\n")
    return record


def verified_accounting_exports(
    *, state_root: Path, workspace_root: Path
) -> list[dict[str, Any]]:
    """Return only private records whose current workspace export is unchanged."""

    provenance_path = state_root / EXPORT_PROVENANCE_FILENAME
    if not provenance_path.is_file():
        return []
    verified_by_path: dict[str, dict[str, Any]] = {}
    for line in provenance_path.read_text(encoding="utf-8").splitlines():
        try:
            record = json.loads(line)
            if not isinstance(record, dict):
                continue
            if record.get("schema_version") != EXPORT_PROVENANCE_SCHEMA_VERSION:
                continue
            normalized, target = _workspace_path(
                workspace_root, str(record.get("path") or "")
            )
            expected_sha256 = str(record.get("sha256") or "")
            if not target.is_file() or len(expected_sha256) != 64:
                continue
            if _sha256(target) != expected_sha256:
                continue
            verified_by_path[normalized] = {
                "path": normalized,
                "report": str(record.get("report") or ""),
                "parameters": dict(record.get("parameters") or {}),
                "row_count": int(record.get("row_count") or 0),
                "sha256": expected_sha256,
                "provenance": "authorized_accounting_export",
            }
        except (json.JSONDecodeError, OSError, TypeError, ValueError):
            continue
    return [verified_by_path[path] for path in sorted(verified_by_path)]
