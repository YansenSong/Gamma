# Excel analysis and mapping

Use the workbook analyzer as a read-only inventory. Record the source SHA-256, worksheets, cells, formulas, merged ranges, validations, names, protection, and unsupported features. AI checks labels, units, cell role, type and scoring relationship, and directly confirms a unique, source-supported match. Ask the user only when multiple plausible mappings remain or the materials omit a required meaning.

## First deliverable: Excel scoring map

Create a human-readable table and a machine-readable companion under `analysis/`
before deriving rules from the PDF. For every workbook score target, record:

| Field | What to capture |
| --- | --- |
| Evaluation item | Workbook's section or scoring dimension. |
| Specific item | Scenario, route, penalty behavior, or other row label; include relevant group headings from merged cells. |
| Performance/result | What each existing input column means, such as a specific observed outcome, occurrence count, elapsed time, or judged condition. |
| Bindings | Sheet and address of each input, intermediate rate, item score, and aggregate score. |
| Current behavior | Existing value/formula, unit and number format; mark a missing score formula as pending. |
| Formula question | What must be calculated from those inputs; leave the actual rate, threshold, and formula pending until the PDF confirms them. |
| Entry contract | Plausible operator entries, accepted representations, canonical value, the meaning of an empty cell, and invalid cases for this specific input or input group; mark uncertainty pending. |

Do not collapse a section into a list of score addresses: a reviewer must be able
to trace each concrete performance choice to an input cell and the score cell it
affects. Include intermediate rates and totals when the workbook asks for them.

## Think through the actual filled workbook

For **every field a person must fill**, reason from its row and column labels,
unit, cell format, validation list, neighboring fields, and any supplied filled
examples. Ask what the operator is recording, how they are likely to enter it
in Excel, and how the app will distinguish equivalent entries from missing,
contradictory, or invalid ones. Cover measured values, counts, categories,
judgements, dates, durations, text and any other input type the workbook uses;
do not start from a fixed list of formats. Record the result as a cell-specific
or group-specific input contract, with evidence and unresolved assumptions.

For each contract, specify the accepted raw forms and units, canonical form for
Excel calculation, normalization rule, valid range or choices, meaning of an
empty cell, difference between empty and an explicit zero/negative response
where relevant, cross-field consistency rules, and the error shown for invalid
data. Inspect how Excel actually stores a user's entry, since visual formatting
can hide a different underlying type or unit. An empty alternative in a
recorded outcome may mean “did not occur”; an entirely empty item may mean
“not tested.” Confirm that distinction for this workbook rather than applying
it universally. Examples such as `是`/`否` versus `1`/`0`, or several ways to
write a duration, illustrate the analysis only; they are not required formats
for every workbook.

Normalize only confirmed, unambiguous representations on a safe copy before
native Excel recalculation. Do not guess a unit, date order, category code, or
meaning of blank from a convenient example. Reject ambiguous, malformed, or
out-of-range entries with the exact sheet/cell and expected forms. Keep the
uploaded workbook unchanged, and audit the raw value, canonical value,
normalization rule and calculation-copy hash. Python may parse and validate
inputs; Excel alone evaluates scores.

Record two calculation paths in the input contract: what native Excel does with
the raw cell when a person edits the workbook, and what it does after the app
normalizes a safe copy. Test both for each accepted representation. If the
workbook is delivered for direct scoring, these paths must agree. A formula
comparison against numeric 1 does not recognize typed `是`; a normalizer that
does recognize it can hide this defect in application-only tests.

Do not classify a row containing explicit `0` as an untouched row. For a
counted scenario, record whether zero means an observed zero-point occurrence,
no occurrence during the test, or an unperformed required test. Those states
may need different scores or a pending result. If the source materials do not
resolve a score-changing distinction, ask the user with the affected row and
total score consequence before confirming formulas.

The bundled `run_job_from_workbook` is a structural starting point; it does not
provide this project-specific normalization by itself. Implement the confirmed
entry contract in the generated app, before its native recalculation call, and
test the filled-workbook route rather than only the legacy facts/JSON route.

Keep mappings hash-bound to the exact workbook and rule specification. Outputs are written to a new copy. Preserve styles and unrelated values. Reject macro-enabled workbooks or workbook features the selected writer cannot preserve.

openpyxl does not calculate formulas. Microsoft Excel formulas are the only scoring authority: require native full recalculation, save and reopen the file, validate mapped formula results and record the Excel version. Missing results or errors block output; Python must not fill in substitute scores.

The generated project must make the selected calculation authority visible in its config and audit record.
