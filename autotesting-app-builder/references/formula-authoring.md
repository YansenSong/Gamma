# Authoring scoring formulas into a bare workbook

`scoring-spec.md` assumes the workbook already contains the scoring formulas.
This document covers the other case: the supplied workbook is a *tally sheet*
whose per-item score cells are empty or manually typed, and the scoring rules
exist only in the charter PDF. The Skill derives the formulas from the charter
and writes them into an authorized copy.

The authority does not move. Excel still evaluates every score. AI performs
*authoring*, which is a reviewable edit to the workbook, not *evaluation*. A
Python expression that computes a score remains forbidden.

## Detect the case

Run `scripts/inspect_excel.py` and inspect the cells the analyzer marks
`scoring_value_candidate` and `formula_output_candidate`.

- Every score-bearing cell has `data_type == 'f'` → the workbook is already an
  engine. Use `scoring-spec.md` and make no edits.
- Score cells hold literal numbers, or the formula inventory covers only
  aggregate `SUM` rows while the per-item cells are bare → this document applies.

A workbook whose only formulas are aggregates is the common case, not an
exception. Aggregates tell you the total is authoritative; they do not tell you
how each item was scored.

## Separate inputs from formulas

Before writing anything, classify every bare score cell:

- **Derivable** — the charter states a deterministic rule over values the test
  produced. This becomes a formula.
- **Judgement** — the charter requires a human verdict the workbook cannot
  observe (for example whether the subject vehicle passed a scene unaided).
  This stays a declared `scoring.inputs` entry that the application collects.
  Never encode a human verdict as a formula.
- **Unresolved** — the charter admits more than one reading. This becomes a
  `business_decisions` entry with `status: pending`. Do not pick silently.

The working shape is: the application collects judgements and measurements into
input cells; Excel formulas turn those into scores; the Skill reads only the
formula results.

## Boundaries on derived formulas

Every formula must be a faithful transcription of a charter provision. The
following are violations, not judgement calls:

- A threshold, cap, rounding rule, tier boundary or default that the charter
  does not state. Caps such as "每项罚分项最多扣3分" and "总罚分最多不超过10分"
  are charter facts — transcribe them. Never invent a cap to make a range behave.
- A widened range or an added `IFERROR` branch that silently converts an
  out-of-range input into a valid score. Out-of-range input must surface as an
  error, not a plausible number.
- Rewriting an existing formula. Cells that already carry a formula are
  authoritative. Changing one requires its own explicit per-cell authorization
  and a new approval; it is never part of a bulk fill.
- Filling cells the charter does not cover. Leave them bare and list them as
  blockers.

Where the charter contradicts itself — two tables giving different rates for the
same conduct, an ambiguous equality boundary, a unit that changes meaning between
sections — record the conflict and the candidate readings with the score
difference each produces. That is exactly the material `business_decisions`
exists for.

## Authorization

Writing formulas is a template cleanup operation under `SKILL.md`. It requires:

1. A separate copy, hash-bound to the pristine original.
2. An explicit per-cell authorization record: cell reference, prior content, the
   new formula, the charter `source_refs` it transcribes, and the reviewer,
   time and reason.
3. The pristine original left byte-identical. Never write into the supplied file.

Cells with no authorization stay unchanged, and their absence blocks a business
verdict.

## Which file becomes the workbook source

`validate_project_spec.py` requires `excel_output.template` to carry the exact
SHA-256 of the `role: workbook` source. A formula-authored copy cannot share the
original's hash, so the copy becomes the canonical source:

- `sources[role=workbook]` and `excel_output.template` both point at the
  **authorized copy**, and share its hash.
- The pristine original is **not** a `sources` entry. Its hash is recorded in the
  analysis inventory, the authorization record and the delivery manifest under a
  name that says what it is, such as `upstream_source_sha256`.

Do not work around this by loosening the hash check. The check is what proves the
application scores with the workbook a human approved.

## Verify

- `scripts/diff_workbooks.py <original> <authored>` — confirm the changed set is
  exactly the authorized cells and that every non-target cell, style, merge,
  validation and sheet property is untouched. An unexpected difference is a
  failure even when the scores look right.
- `scripts/validate_workbook.py <authored> --expected-sha256 <hash>` then
  `--require-recalculated` after native Excel recalculation. openpyxl does not
  calculate; a missing cache is a blocker, never a Python fallback.
- Test declared boundary values through native Excel and compare against the
  charter's stated tiers. A formula that has never been recalculated by Excel is
  an untested formula.
- Confirm the authored copy opens without repair prompts and that the target
  environment's Excel version is recorded.

## Report

State, per cell: the rule, its charter citation, the formula written, the
authorization, and the verification outcome. Report the aggregate, the
pristine-original hash and the authored-copy hash together. A workbook whose
per-item formulas were derived by AI is an `engineering_sample` until named
reviewers approve both the derivations and the calculation authority.
