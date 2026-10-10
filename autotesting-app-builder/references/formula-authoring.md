# Authoring scoring formulas into a bare workbook

`scoring-spec.md` assumes the workbook already contains the scoring formulas.
This document covers the other case: the supplied workbook is a *tally sheet*
whose per-item score cells are empty or manually typed, and the scoring rules
exist only in the charter PDF. The Skill derives the formulas from the charter
and writes them back to the supplied Excel path only after native validation.

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

## Project-copy writeback

Unambiguous missing formulas may be filled in the project copy without another per-cell permission prompt. Point the plan at `../templates/score.xlsx` inside the `init_project.py` workspace; do not point it at the user's original source. Existing formulas, including SUM aggregates, remain intact. Conflicts and ambiguous business rules are blockers until explicitly reviewed.

The JSON plan requires `workbook`, `source_sha256`, `formulas` and `boundary_cases`. Each formula has `sheet`, `cell`, `original: null`, `formula`, `inputs`, `pdf_refs` with integer page and location, `rule_explanation`, `boundary_checks` and risks. Each boundary case has concrete `inputs` and `expected` formula results. These are reviewed engineering expectations from the PDF, not Python-calculated scores.

Run `write_scoring_formulas.py prepare <plan.json> --staged <stage.xlsx> --manifest <stage.json>`. This checks syntax, references and cycles, inserts only the target formulas into an independent XLSX and records hashes. Next run `write_scoring_formulas.py commit <plan.json> --manifest <stage.json> --evidence <native.json> --receipt <receipt.json>` on Windows with Microsoft Excel installed. Commit invokes `validate_formula_stage.ps1`, which recalculates every boundary case in Excel. It checks the plan, source and stage hashes before creating a unique backup, atomically replacing the original Excel path and writing the receipt. Failed validation leaves the original intact; a post-replacement verification failure restores it from backup.

A constructed test double is valid only for engineering tests of the transaction. It is not Microsoft Excel evidence and cannot authorize a real delivery. Keep the rewritten Excel as the program's controlled scoring template and record the pre-write hash separately.

## Verify

- `scripts/diff_workbooks.py <original> <authored>` — confirm the changed set is
  exactly the planned cells and that every non-target cell, style, merge,
  validation and sheet property is untouched. An unexpected difference is a
  failure even when the scores look right.
- `scripts/validate_workbook.py <authored> --expected-sha256 <hash>` then
  `--require-recalculated` after native Excel recalculation. openpyxl does not
  calculate; a missing cache is a blocker, never a Python fallback.
- Test declared boundary values through native Excel and compare against the
  charter's stated tiers. A formula that has never been recalculated by Excel is
  an untested formula.
- Confirm the rewritten workbook opens without repair prompts and that the target
  environment's Excel version is recorded.

## Report

State, per cell: the rule, its charter citation, the formula written, the
authorization, and the verification outcome. Report the aggregate, the
pre-write hash, backup hash and rewritten-workbook hash together. A workbook whose
per-item formulas were derived by AI is an `engineering_sample` until named
reviewers approve both the derivations and the calculation authority.
