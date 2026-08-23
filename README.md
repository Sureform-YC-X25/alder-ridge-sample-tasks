# Alder Ridge Sample Tasks

A runnable company environment for evaluating agents on long-horizon corporate finance work. Alder Ridge Mechanical is a simulated specialty mechanical contractor with project-based accounting, percentage-of-completion revenue recognition, WIP, change orders, service operations, treasury, and contractor-accounting records.

Tasks in this dataset require agents to use company accounting tools via MCP, along with spreadsheets, documents, presentations, PDFs, emails, policies, and operating extracts. The tasks test controllership, FP&A, treasury, working capital, manufacturing and project accounting, tax, corporate development, capital allocation, and lender and board reporting. Assignments range from reconciliations and analyses to creating/editing full finance deliverables.

Task Type Breakdown

- Controllership & Project Accounting: 9 tasks (2 in sample)
- FP&A & Operational Finance: 24 tasks (2 in sample)
- Treasury, Working Capital & Lender Finance: 25 tasks (2 in sample)
- Strategic Finance & Capital Allocation: 13 tasks (1 in sample)
- Corporate Development & M&A: 16 tasks (1 in sample)
- Corporate Tax: 5 tasks (1 in sample)
- Investor Relations, Board & External Reporting: 6 tasks (2 in sample)
- CFO Synthesis & Executive Decision Support: 2 tasks (1 in sample)

This sample contains 12 of the 100 tasks in the full taskset.

All tasks and source files were authored by domain experts and informed by Sureform's partner engagements with a specialty mechanical contractor. Alder Ridge Mechanical, including its personnel, counterparties, communications, transactions, and records, are simulated. The package does not contain identifiable client, partner, employee, insurance, banking, or customer data.

## Start here

The repository has two parts:

```text
environment/
  Dockerfile
  runtime/              runnable environment, accounting MCP and grading code
  seed/
    accounting.db       simulated company accounting system
    ACCOUNTING_MCP.md   accounting records and MCP tool reference
    sources/            complete shared company document world

tasks/
  <task-slug>/
    prompt.md
    task.json
    rubric.json
    gold.json
    source_manifest.json
```

Open `tasks/` to review the 12 assignments. Open `environment/seed/sources/` to inspect the complete 132-file shared company world: spreadsheets, documents, presentations, PDFs, emails and operating extracts. This is the same evidence universe used by the finalized 100-task environment, including relevant files, distractors, version history and conflicting working drafts. `environment/seed/accounting.db` is the unchanged full company accounting snapshot exposed through the accounting MCP.

The `environment/runtime/grading/` directory is the executable verifier used for the tasks. It checks financial values, dates, formulas, document structure, source lineage and required deliverables deterministically. Semantic review is limited to criteria that allow professionally equivalent wording and cannot override failed deterministic finance checks.

The repository does not contain task definitions, gold data, rubrics, graders or task-specific seed overlays for the other 88 tasks. It intentionally contains the full shared company evidence universe and accounting system so reviewers can assess each sample task against the real navigation and reconciliation burden of the complete world.

## Run the environment

Docker with at least 8 GB of memory is recommended.

```bash
docker build -f environment/Dockerfile -t alder-ridge-sample-tasks:1.1.1 .
docker run --rm \
  --cap-add SYS_ADMIN \
  --security-opt seccomp=unconfined \
  --security-opt apparmor=unconfined \
  --security-opt systempaths=unconfined \
  -p 8765:8765 \
  alder-ridge-sample-tasks:1.1.1
```

The service listens on port `8765`. The listed Linux container permissions allow the environment to create its nested Bubblewrap sandbox; they do not expose the hidden evaluation files to the task agent. Each task receives an isolated `/workspace` and the `contractor_accounting` MCP capability. The agent cannot access the accounting database, gold values, rubric implementation or environment source code directly.

## Validate or grade locally

Python 3.12 and `uv` are required.

```bash
uv sync --project environment --frozen
uv run --project environment python environment/validate.py
```

To grade an answer or completed artifact workspace:

```bash
uv run --project environment python environment/grade.py \
  --task credit-metrics-debt-capacity \
  --answer-file answer.json \
  --workspace environment/seed/sources
```

## Access and licensing

This repository and its container image are private commercial-evaluation materials. Add reviewers as read-only GitHub outside collaborators. Grant container access separately with repository-level GCP `Artifact Registry Reader` permission. Read-only reviewers cannot manage the access list.

The repository is not open source and may not be redistributed, used for model training or used in production without a separate signed agreement. See [LICENSE](LICENSE).

## Provenance

The sample was clean-room exported from canonical finalized commit `0202b6e1bf789a20d3770335922e4ee34f231408`. The complete shared source inventory, selected task set and accounting-seed digest are recorded in `environment/seed/sample_manifest.json`.
