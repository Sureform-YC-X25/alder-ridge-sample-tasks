# Grading design

## Deterministic checks

Every checkable fact is evaluated deterministically. Depending on the task,
the verifier checks financial values and tolerances, dates, IDs, formulas,
roll-forwards, source lineage, native workbook objects, document sections,
slide titles and counts, artifact readability, and spreadsheet error states.
Weighted atomic criteria provide partial credit.

## Semantic checks

Criteria marked `semantic: true` permit professionally equivalent wording or
presentation. They are reviewed separately from numeric checks. The semantic
review receives criterion-specific expected facts and authorized source
evidence. It cannot rescue a deterministic numerical contradiction or a failed
hard gate. The reference-answer construction path is not used as the judge.

## Integrity and anti-bluff controls

The final reward is additionally governed by checks for required artifact
creation/modification, prohibited accounting mutations, fabricated or
unauthorized source claims, grader leakage, response-only boundary violations,
unreadable files, and missing governing model decisions. A response that merely
repeats prompt language without the requested finance work earns only the
criteria it actually satisfies.

## Audit surfaces

Each `tasks/<slug>/` folder contains the prompt, complete gold/reference
contract, public weighted rubric, and source dependency declaration. The exact
executable implementations are in `graders/`; the scope manifest binds the
package to the canonical source commit and accounting-seed digest.
