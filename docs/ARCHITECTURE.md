# Runtime architecture

The sample uses one company world and twelve independently runnable tasks.

1. `reset_world()` copies the immutable selected source package into a clean
   `/workspace` and copies the shared accounting seed into isolated mutable
   state.
2. The agent receives a restricted shell rooted at `/workspace` and the
   `contractor_accounting` MCP capability.
3. The agent cannot browse grader code, gold data, raw SQLite storage, host Git
   metadata, or orchestration state from the task shell.
4. On completion, the deterministic grader reads the response and/or required
   artifact, then semantic review evaluates only explicitly semantic criteria.
5. Integrity and reward-policy gates are applied after criterion scoring, and
   the complete result is stored outside the agent workspace.

The image contains the private verifier so a reviewer can reproduce grading.
Runtime isolation prevents that verifier from becoming visible to the model
being evaluated.

## Shared accounting system

`world/seed/accounting_seed.db` is the canonical company system-of-record
snapshot shared by the full environment. It is retained unchanged to preserve
the exact posted facts and MCP behavior used by the selected tasks. It contains
no task prompts, rubrics, gold answers, or task mapping. Only the selected
tasks' 71 workspace files are included.
