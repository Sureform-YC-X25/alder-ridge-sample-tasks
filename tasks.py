from __future__ import annotations

from env import TASK_TEMPLATES
from task_catalog import TASKS


GRADING_REVISIONS = {'task_001': 'weighted-atomic-hybrid-v4', 'task_004': 'weighted-atomic-hybrid-v4', 'task_015': 'weighted-atomic-hybrid-v7-covenant-slide-presentation-quality', 'task_027': 'weighted-atomic-hybrid-v12-covenant-release-decision', 'task_035': 'weighted-atomic-hybrid-v12-executive-recovery-decision', 'task_037': 'weighted-atomic-hybrid-v7-portfolio-resilience-decision', 'task_055': 'weighted-atomic-hybrid-v6-transaction-sensitivity-release', 'task_061': 'weighted-atomic-hybrid-v6-tax-close-journal-release', 'task_068': 'weighted-atomic-hybrid-v7-guidance-mitigation-release', 'task_072': 'weighted-atomic-hybrid-v5-deterministic-financial-narrative', 'task_073': 'weighted-atomic-hybrid-v5-pro-forma-credit-capacity', 'task_100': 'weighted-atomic-hybrid-v6-layout-aware-branch-risk-schedules'}

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
