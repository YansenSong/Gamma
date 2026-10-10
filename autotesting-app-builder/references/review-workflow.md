# Business review and incremental evidence revalidation

Use this workflow when an engineering project needs named business opinions or
new trial evidence. Keep scoring code and accepted business decisions separate
from submitted opinions. A CSV is a handoff form, not an approval authority.

## Prepare the review package

For a project without mappings, explicitly declare optional identity fields in
`review_preview.metadata_fields`, then run `scripts/build_review_inventory.py
<project-root>`. All other fields remain D (missing trial or conclusion evidence).
This helper does not identify Excel semantic matches. Preserve B proposals only
when an actual candidate source and its semantics have been separately analyzed.

Run `scripts/review_workflow.py export <project-root> --out
<project-root>/analysis/review-package`. It includes current report variables,
identified charter rows, business questions and known blockers. Coverage rows are
included only if the project has them; question inventories do not prove full
charter coverage. Historical v1.1 Python comparison rows remain audit history, not
a live scoring engine. New score verification uses native Excel only. Copy the exported package JSON to
`business_review.package` when enabling bound CSV export in the desktop dialog.
Refresh both after changing project configuration.

Each opinion binds the entire package SHA-256 and the individual subject hash.
The package binds source files, project configuration and four approval content
projections. Changing rules, mapping, calculation authority, source inventory or
review contents requires a new package. Retain old packages as historical records.

## Collect opinions and evidence

The business reviewer fills `opinions.csv` with a real name, timezone-bearing
ISO timestamp, decision, reason, evidence paths, SHA-256 and record identifiers.
Evidence arrays are JSON arrays with equal lengths. New files live under the
independent project's `evidence/submissions/`. Do not overwrite original templates
or analysis records. `accept`, `reject`, `needs_evidence`, `not_applicable` are
submitted opinions, never self-authenticating approvals. Keep private trial data
out of public Git by default; the project controls access to those submissions.

Run `scripts/review_workflow.py revalidate <project-root> --package <package.json>
--opinions <opinions.csv> --previous <prior-result.json> --out <result.json>`.
The script reruns project validation, compares the template's raw XML variables,
checks opinion bindings and evidence identity, and inspects workbook caches or
DOCX readability. Previous records identify changed subjects; they never bypass
checks. The result lists the project-specific rule, fact, mapping and layout
checks still pending. Structure checks alone cannot establish trial truth, native
Excel recalculation, business semantics, or authenticity of the reviewer's name.

After an opinion is independently accepted, update only the affected project
code, facts and mapping, rerun its relevant comparisons, and collect current
content-bound approvals. Keep unresolved Q1–Q10 and threshold questions pending.
The workflow does not automatically import CSV opinions into project approvals,
change scores, or promote a release. A complete project does not need regeneration
when evidence is supplemented. `verify_release.py` remains the release gate.

## Conditional empty formula results

Record each questionable output, its selection input, the workbook hash, exact
charter provision, selected-project trial record, total-score treatment and named
review. Empty caches, an empty selection input, or a manually typed
`not_applicable` opinion do not establish permission to omit a score.
`config/formula-empty-review.json` is a proposed disposition inventory. It is not
an exception list consumed by `validate_workbook.py`. Keep strict cache failure
until a separately approved calculation/mapping policy implements the agreed
conditions and has native recalculation evidence. Do not modify a legacy formula
merely to make an empty cache disappear.

## Unscored engineering preview

Blocked review projects may configure `review_preview` and the reusable
`review_preview.py`, `report_safety.py`, `field_review_ui.py` modules. The preview
produces a DOCX and audit in one staged directory. It does not score or create a
workbook. Optional identity entries are marked user-supplied/unverified; all
unmapped result fields remain “未提供”. No grade, collision, completion, signature,
report number or signing date is inferred from examples. General scoring output
buttons remain blocked. Use the exact current template hash and visibly mark
the first page, every page and any signature boxes. Review copied static text for
claims about formal scoring authority and override them explicitly in the copy's
safety policy when necessary. Render and inspect every current preview page.
