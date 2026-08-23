from __future__ import annotations

import os
from pathlib import Path


def resolve_project_root(anchor: str | Path = __file__) -> Path:
    """Locate the checked-out world when code is installed non-editably.

    Production images keep the immutable world beside the installed package,
    while a non-editable wheel lives under ``site-packages``.  Never infer the
    world from the wheel location alone; prefer the explicit release root and
    then a checked-out ancestor of the current working directory.
    """

    candidates: list[Path] = []
    configured = os.environ.get("ALDER_RIDGE_PROJECT_ROOT")
    if configured:
        candidates.append(Path(configured).expanduser())
    current = Path.cwd().resolve()
    candidates.extend((current, *current.parents))
    module = Path(anchor).resolve()
    candidates.extend(module.parents)

    seen: set[Path] = set()
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate in seen:
            continue
        seen.add(candidate)
        if (candidate / "runtime" / "tasks.py").is_file() and (
            candidate / "seed"
        ).is_dir():
            return candidate
    raise RuntimeError(
        "Unable to locate the Alder Ridge project root. Set "
        "ALDER_RIDGE_PROJECT_ROOT to the immutable checked-out release."
    )


def resolve_seed_root(anchor: str | Path = __file__) -> Path:
    configured = os.environ.get("WORLD_SEED_ROOT")
    if configured:
        return Path(configured).expanduser().resolve()
    return resolve_project_root(anchor) / "seed"
