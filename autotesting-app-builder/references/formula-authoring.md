# Authoring scoring formulas into a bare workbook

`scoring-spec.md` assumes the workbook already contains the scoring formulas.
This document covers the other case: the supplied workbook is a *tally sheet*
whose per-item score cells are empty or manually typed. AI first inventories the
workbook's score targets, inputs and existing formulas. For each target, identify
the formula question and locate its PDF evidence; resolve score-changing ambiguity
only where it affects a workbook target. Then author `评分细则.md` and draft per-cell
Excel formulas. The PDF remains the source of evidence; the reviewed specification
is the working interpretation. Write formulas only to the
project workbook copy after native validation.

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

Before writing anything, list every workbook score row, heading and input column,
then compare each target with its relevant PDF evidence. Classify every bare score cell:

- **Derivable** — the reviewed specification states a deterministic rule over values the test
  produced. This becomes a formula.
- **Judgement** — the rule requires a human verdict the workbook cannot observe
  (for example whether the subject vehicle passed a scene unaided). The operator
  records the verdict in an explicitly mapped workbook input cell before upload;
  the generated application's default path reads the completed workbook. Never
  encode an unobserved human verdict as a guessed formula result.
- **Unresolved** — the specification or Excel mapping admits more than one
  reading. Record the exact question and keep the formula candidate pending.
  Do not pick silently.

Keep workbook coverage separate from charter inventory. A PDF clause without an
Excel target may be noted as out of scope; it does not create a new score cell,
question, or release blocker. A workbook score target without adequate PDF support
remains blocked and must not receive an invented formula.

## Keep the workbook's input contract

Inventory the workbook's existing input cells before applying charter clauses. A
charter event is not automatically a new input field. For every apparent missing
fact, check the workbook's existing mutually exclusive states, score-tier counts,
zero-point category, and already derived measurements. A charter override can be
an instruction for how an operator classifies an occurrence into an existing
category before entering its count. Count zero-point occurrences in the existing
category's denominator when the workbook supplies one; never silently omit them.

If a required target remains uncomputable after that check, record the target,
the exact missing dependency, and the attempted existing-cell mappings as a
blocker. Do not ask the user to add columns or records as a routine clarification.
Changing the workbook schema requires a separate, explicit user request or
authorization. Keep source workbook and project copy provenance distinct.

The working shape is: people enter measurements and judgements into the workbook;
Excel formulas turn those into scores; the application recalculates a safe copy
of this uploaded workbook and reads only its formula results.

Before drafting a formula, use the `excel-mapping.md` entry contract for each
human-filled input. Distinguish an unperformed item from omitted alternatives
inside a recorded result according to this workbook's confirmed convention.
Do not implement a blanket blank-to-zero shortcut or require every unused
choice to be filled. Test the plausible forms of each entered result on filled
workbook copies before committing formulas.

Decide and record whether users will calculate by opening the workbook directly
or only by uploading it to the application. For direct Excel use, formulas must
interpret the accepted raw forms themselves. For example, a flag formula that
compares only `=1` treats a typed `是` as false even if the application later
normalizes `是` to 1. Test both paths separately with the same raw inputs. If a
form is accepted only by the application, label that limitation in the workbook
and delivery instructions, and do not claim its direct Excel result is valid.

For every count group, distinguish an untouched blank row, an explicit zero
total, positive counts, and a counted occurrence in a zero-point category.
These are different business states. A denominator of zero needs a reviewed
rule; `NA()`, `IFERROR`, zero, full marks, or omitting the item are all
score-changing choices. If the PDF and workbook do not decide which applies,
ask the user before confirming that formula. A request to support multiple
entry formats does not resolve this scoring decision.

## Boundaries on derived formulas

Every formula must faithfully implement the reviewed specification and remain
traceable to the PDF provision and any decision that resolved an ambiguity. The
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

Unambiguous missing formulas may be staged and confirmed by AI in the project copy without a per-cell permission prompt. Point the plan at `../templates/score.xlsx` inside the `init_project.py` workspace; do not point it at the user's original source. Existing formulas, including SUM aggregates, remain intact. Record score-changing business or cell-mapping questions in the project review document, then ask the user interactively in batches of three or four when available. Explain each alternative and its score or cell effect; offer concrete choices and a free-text correction. Pending questions may be tested on a staged copy but block `commit` to the project template.

The JSON plan requires `workbook`, `source_sha256`, `review`, `formulas` and `boundary_cases`. `review.scoring_spec` and `review.rules` each contain a path and SHA-256 binding the exact AI-authored Markdown and structured rule file. Each formula has `sheet`, `cell`, `original: null`, `formula`, `inputs`, `pdf_refs` with integer page and location, `rule_ids`, `rule_explanation`, `boundary_checks`, risks and `review_status`. A formula can be committed only when its referenced rules are confirmed and it has `review_status: confirmed`, `reviewer` and `reviewed_at`. For a clear derivation, identify AI as reviewer and record its source evidence; for an ambiguity, bind the user's decision. Each boundary case has concrete `inputs` and `expected` formula results. These are reviewed engineering expectations, not Python-calculated scores.

Run `write_scoring_formulas.py prepare <plan.json> --staged <stage.xlsx> --manifest <stage.json>`. This checks syntax, references and cycles, inserts only the target formulas into an independent XLSX and records hashes. `prepare` accepts pending candidates for review and testing. Next run `write_scoring_formulas.py commit <plan.json> --manifest <stage.json> --evidence <native.json> --receipt <receipt.json>` on Windows with Microsoft Excel installed. Commit first checks the specification and rule hashes, rule IDs and named per-formula confirmation. It then invokes `validate_formula_stage.ps1`, which recalculates every boundary case in Excel. It checks the plan, source and stage hashes before creating a unique backup, atomically replacing the project workbook path and writing the receipt. Failed validation leaves the original intact; a post-replacement verification failure restores it from backup.

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
per-item formulas were derived by AI remains an `engineering_sample` until the derivations, source-bound confirmation records, native Excel calculation authority and real evidence pass the release gates. Attribute AI confirmations to AI, not to a human.
