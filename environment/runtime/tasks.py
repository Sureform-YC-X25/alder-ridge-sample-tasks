from __future__ import annotations

from runtime.env import TASK_TEMPLATES
from runtime.task_catalog import TASKS


GRADING_REVISIONS = {'task_001': 'weighted-atomic-hybrid-v27-full-population-register', 'task_004': 'weighted-atomic-hybrid-v11-accounting-control-wip-review', 'task_015': 'weighted-atomic-hybrid-v15-wip-reversal-sensitivity', 'task_027': 'weighted-atomic-hybrid-v13-stable-scenario-model', 'task_035': 'weighted-atomic-hybrid-v21-source-aligned-proportional-status', 'task_037': 'weighted-atomic-hybrid-v9-proportional-capital-portfolio', 'task_055': 'weighted-atomic-hybrid-v8-case-aware-accretion-sensitivity', 'task_061': 'weighted-atomic-hybrid-v7-disclosed-tax-close-inputs', 'task_068': 'weighted-atomic-hybrid-v16-source-complete-central-work', 'task_072': 'weighted-atomic-hybrid-v6-management-narrative-equivalence', 'task_073': 'weighted-atomic-hybrid-v6-disclosed-pro-forma-debt-basis', 'task_100': 'weighted-atomic-hybrid-v8-local-board-card-association'}

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
