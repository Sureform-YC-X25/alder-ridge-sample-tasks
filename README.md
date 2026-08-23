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

## What is included

This private repository is the complete technical-diligence package for the sample. It contains:

- the exact 12 task prompts and metadata;
- the task-specific source documents and editable starting artifacts;
- the shared June 30 accounting snapshot and the MCP accounting interface;
- complete gold/reference contracts and weighted rubrics;
- deterministic file, formula, finance, integrity, and source checks;
- the separately invoked semantic-review layer used for professionally equivalent answers;
- reproducible scope, source-integrity, and grader tests; and
- the Docker build used to run the sample environment.

It does not contain task definitions, gold data, rubrics, or exclusive source-document packages for the other 88 tasks. The shared accounting database is intentionally unchanged because it is the common system of record required to reproduce the selected tasks exactly.

## Repository map

| Path | Purpose |
| --- | --- |
| `tasks/<slug>/` | Prompt, task metadata, full gold contract, rubric, and source dependency declaration |
| `world/seed/workspace/` | The 71 source files needed by the selected tasks |
| `world/seed/accounting_seed.db` | Shared, read-only seed for the accounting MCP |
| `alder_ridge_world/` | Accounting repository, MCP tools, reset logic, and shell isolation |
| `graders/` | Deterministic, semantic, integrity, and weighted reward implementations |
| `tests/` | Scope, source-integrity, runtime, and oracle checks |
| `SAMPLE_SCOPE_MANIFEST.json` | Machine-readable proof of the selected task and source closure |

## Run locally

Requirements: Docker with at least 8 GB of memory available.

```bash
docker build -t alder-ridge-sample-tasks:1.0.0 .
docker run --rm -p 8765:8765 alder-ridge-sample-tasks:1.0.0
```

The environment exposes the HUD service on port `8765`. Each task receives an isolated `/workspace`, plus the `contractor_accounting` MCP capability. Raw accounting storage, grader code, and gold data are not mounted into the agent shell.

## Inspect and test without Docker

Python 3.12 and `uv` are required.

```bash
uv sync --frozen --extra dev
uv run pytest
uv run python scripts/verify_sample_scope.py
```

To grade an answer or completed artifact workspace directly:

```bash
uv run python scripts/grade_task.py \
  --task credit-metrics-debt-capacity \
  --answer-file answer.json \
  --workspace world/seed/workspace
```

## Grading model

Checkable facts are graded deterministically: financial values, dates, source lineage, formulas, workbook controls, document structure, and required artifacts. Semantic review is used only for criteria that permit professionally equivalent wording. The semantic verifier receives source-grounded evidence and cannot overwrite deterministic financial failures. Fabricated citations, prohibited writes, missing required artifact changes, and grader/infrastructure failures are handled by separate fail-closed gates. See [docs/GRADING.md](docs/GRADING.md).

## Access and licensing

This repository and its container image are private commercial-evaluation materials. Access is granted to named reviewers only. The repository is not open source and may not be redistributed, used for model training, or used in production without a separate signed agreement. See [LICENSE](LICENSE) and [docs/ACCESS.md](docs/ACCESS.md).

## Published sample image

The private GCP Artifact Registry image was built from repository commit `e22913e6851ef420dca3c483794aee236f937fa8`:

```text
us-central1-docker.pkg.dev/sureform-479706/alder-ridge-samples/alder-ridge-sample-tasks:1.0.0
us-central1-docker.pkg.dev/sureform-479706/alder-ridge-samples/alder-ridge-sample-tasks@sha256:685e0e90908b05fdd54e3b6f83cb6d06978b0a5cb67fa23c85d1961eec63e2fb
```

Use the digest-qualified reference for an immutable evaluation. Pull access is granted per buyer as described in [docs/ACCESS.md](docs/ACCESS.md).

## Provenance

The sample was clean-room exported from canonical finalized commit `0202b6e1bf789a20d3770335922e4ee34f231408`. The selected task/source closure and accounting-seed digest are recorded in `SAMPLE_SCOPE_MANIFEST.json`.
