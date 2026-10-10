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

## Original-path writeback

The user authorizes unambiguous missing formulas as part of this workflow. Existing formulas are preserved, including valid SUM aggregates. Overwrites, conflicts and ambiguous business rules require explicit review and are blocked until resolved.

Create a per-cell JSON plan with `workbook`, `source_sha256`, `formulas`, and `native_excel_validation`. Each formula entry includes `sheet`, `cell`, `original: null`, `formula`, `inputs`, `pdf_refs` (page and location), `rule_explanation`, `boundary_checks`, and risks. Run native Microsoft Excel recalculation against the staged formulas first; the proof contains `status: passed`, source hash, Excel version and an evidence file path. A proof claim without actual run evidence is invalid.

`write_scoring_formulas.py` creates a unique backup, patches worksheet XML while keeping other XLSX package members, verifies the staged workbook, then atomically replaces the original path. It rejects unsupported features and nonblank or protected targets. Reopen the original path and retain the receipt with original, backup and new SHA-256. If a write fails before replacement, the original stays unchanged; the backup can restore it. Use the rewritten workbook as the controlled runtime template and retain the pre-write hash as provenance.

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
