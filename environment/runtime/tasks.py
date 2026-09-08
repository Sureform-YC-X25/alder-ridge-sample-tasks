from __future__ import annotations

from runtime.env import TASK_TEMPLATES
from runtime.task_catalog import TASKS


GRADING_REVISIONS = {
    "task_001": "weighted-atomic-hybrid-v52-release-boundary-semantic-fidelity",
    "task_004": "weighted-atomic-hybrid-v19-source-action-conjunct-fidelity",
    "task_015": "weighted-atomic-hybrid-v19-aligned-matrix-unit-scoping",
    "task_035": "weighted-atomic-hybrid-v24-semantic-metric-association",
    "task_068": "weighted-atomic-hybrid-v28-risk-basis-and-semantic-association",
}

tasks = []
for spec in TASKS:
    task = TASK_TEMPLATES[spec.task_id]()
    task.slug = spec.slug
    task.columns = {
        'world': 'alder-ridge-mechanical',
        'workflow': spec.workflow,
        'level': 'senior-finance-analyst',
        'output_mode': spec.output_mode,
        'difficulty': spec.difficulty,
        'snapshot': '2026-06-30-pre-close',
        'grading': GRADING_REVISIONS[spec.task_id],
    }
    tasks.append(task)

del task, spec
