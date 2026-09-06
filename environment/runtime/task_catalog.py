from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    slug: str
    title: str
    workflow: str
    output_mode: str
    prompt: str
    difficulty: str = "advanced"


TASKS: tuple[TaskSpec, ...] = (
    TaskSpec(
        task_id="task_001",
        slug="arm-2409-wip-backcharge",
        title="ARM-2409 June WIP controller sign-off memorandum",
        workflow="project-accounting",
        output_mode="document_edit",
        prompt="""Please finalize the ARM-2409 June WIP sign-off memorandum at `Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx`.

Prepare a review-ready June 30 recommendation from records available through July 2 and the current close support. Reconcile the project to final May WIP and posted June activity, quantify its proposed June WIP contribution, and document unresolved matters, owners, and evidence needed for Controller disposition.

Complete only the working memorandum and save it in place. Do not post a journal, imply Controller approval, change another close file, or create support exports. A brief completion note is enough.
""",
    ),
    TaskSpec(
        task_id="task_004",
        slug="complete-june-wip-risk-template",
        title="Complete the four-project WIP risk template",
        workflow="controllership",
        output_mode="spreadsheet_edit",
        prompt="""Please complete the Controller's four-project June WIP risk review at
`Shared/Finance/Close/2026/06 June/4 WIP/WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx`.

Give the Controller a review-ready June 30 recommendation using the records available as of July 1. Reconcile each project to the posted accounting records and prior close, support the current estimates from the available project and commercial records, and identify any matter that still needs an owner or approval before release.

Complete only the working workbook and save it in place. Do not post or imply approval. A brief completion note is sufficient in the final response.
""",
    ),
    TaskSpec(
        task_id="task_015",
        slug="append-q2-covenant-slide",
        title="Append one Q2 covenant-headroom slide",
        workflow="lender-reporting",
        output_mode="presentation_edit",
        prompt="""Please append one decision-ready June 30 covenant-headroom slide to
`Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 lender update - review working v3.pptx`.

Reperform the covenants from the executed agreement and current support. Show the posted and proposed-WIP bases, the meaningful downside capacity, the calculation/source basis, and the resulting compliance and circulation conclusion. Keep the proposal visibly unposted.

Preserve the five existing slides and their order, match the deck's visual system, append exactly one slide, and save in place. This remains a working deck, do not sign, submit, or describe it as submitted.
""",
    ),
    TaskSpec(
        task_id="task_035",
        slug="complete-backlog-capacity-model",
        title="Complete the backlog burn and capacity model",
        workflow="fpa-operations",
        output_mode="spreadsheet_edit",
        prompt="""Please finalize the FY27 backlog burn and labor-capacity model at
`Shared/Finance/FP&A/Backlog/FY27 backlog burn and capacity - WORKING.xlsx` for the next planning review.

Use the available planning, backlog, labor, and contract records to produce a formula-driven, auditable delivery outlook and management recommendation. Reconcile the schedules, quantify the financial exposure, and apply the approved definitions and authority evidence consistently.

Preserve the three input schedules, recalculate the workbook, and save it in place. Leave no spreadsheet errors or temporary files in the shared workspace.""",
        difficulty="expert-long-horizon",
    ),
    TaskSpec(
        task_id="task_068",
        slug="complete-executive-performance-deck",
        title="Complete the June executive performance review",
        workflow="fpa-investor-relations",
        output_mode="presentation_edit",
        prompt="""Please finalize the June executive performance review at
`Shared/Finance/Reporting/2026/06 June/June executive performance review - WORKING.pptx` for the leadership meeting.

Give leadership a source-backed view of Q2 performance, liquidity, outlook, and the resulting guidance recommendation. Reconcile the accounting records to the planning records, explain the material drivers and risks, and make the recommendation and accountable next steps auditable.

Preserve the nine-slide template, master, section order, and title slide. Keep the deck readable and source-footnoted, render-check it, and save it in place. Do not create separate deliverables or temporary exports in the shared workspace.""",
        difficulty="advanced-long-horizon",
    ),
)

TASK_BY_ID = {task.task_id: task for task in TASKS}
TASK_BY_SLUG = {task.slug: task for task in TASKS}
