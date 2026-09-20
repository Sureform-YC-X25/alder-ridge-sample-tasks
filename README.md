# Alder Ridge Sample Tasks

A runnable five-task company environment for evaluating agents on long-horizon corporate-finance work. Alder Ridge Mechanical is a simulated specialty mechanical contractor with project-based accounting, percentage-of-completion revenue recognition, WIP, change orders, service operations, treasury, and contractor-accounting records.

The sample contains:

- ARM-2409 June WIP controller sign-off memorandum
- four-project June WIP risk review
- Q2 covenant-headroom slide
- FY27 backlog burn and labor-capacity model
- June executive performance review

All tasks and source files were authored by domain experts and informed by Sureform's partner engagements with a specialty mechanical contractor. Alder Ridge Mechanical, including its personnel, counterparties, communications, transactions, and records, is simulated. The package does not contain identifiable client, partner, employee, insurance, banking, or customer data.

## Repository structure

```text
environment/
  Dockerfile
  runtime/              runnable environment, accounting MCP, graders, and verifiers
  seed/
    accounting.db       simulated company accounting system
    ACCOUNTING_MCP.md   accounting records and MCP tool reference
    controls/           
    sources/            seed files (spreadsheets, documents, etc.)

tasks/
  <task-slug>/
    prompt.md
    task.json
    rubric.json
    gold.json
    source_manifest.json
```

## Run the environment

```bash
docker build -f environment/Dockerfile -t alder-ridge-sample-tasks:1.0.0 environment
docker run --rm \
  --cap-add SYS_ADMIN \
  --security-opt seccomp=unconfined \
  --security-opt apparmor=unconfined \
  --security-opt systempaths=unconfined \
  -p 8765:8765 \
  alder-ridge-sample-tasks:1.0.0
```

The service listens on port `8765`. The listed Linux container permissions allow the environment to create its nested Bubblewrap sandbox; they do not expose hidden evaluation files to the task agent. Each task receives an isolated `/workspace` and the `contractor_accounting` MCP capability. The agent cannot access the accounting database, gold values, rubric implementation, verifier controls, or environment source code directly.

## Validate or grade locally

Python 3.12 and `uv` are required.

```bash
uv sync --project environment --frozen
uv run --project environment python environment/validate.py
```

To grade an answer or completed artifact workspace:

```bash
uv run --project environment python environment/grade.py \
  --task complete-executive-performance-deck \
  --workspace environment/seed/sources
```

## Access and licensing

This repository and its container image are private commercial-evaluation materials. Add reviewers as read-only GitHub outside collaborators.

The repository is not open source and may not be redistributed, used for model training, or used in production without a separate signed agreement. See [LICENSE](LICENSE).
