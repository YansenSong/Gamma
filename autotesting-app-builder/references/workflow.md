# Generation workflow

## Inputs and identity

Run `scripts/init_project.py <pdf> <xlsx> <docx> --out <new-project> --name <name>`
first. The destination must be new. The result copies the three originals under
`sources/`, copies the workbook to `templates/score.xlsx`, records hashes and
inventories under `analysis/`, and creates empty plans under `review/`. The
original inputs are never rewritten by this flow. It is deliberately
`review_required` until the business meaning is supplied and checked.

Record each source file's absolute input role, SHA-256, format, size, and readability in config/project.json. The DOCX must be a blank template whose fill-in positions are marked with `{{ }}`; reject a markerless document as a completed report, which is not an accepted input. Inventory the marker set and carry it forward as the report field list. Keep business names and scoring claims grounded in the supplied material.

## Analysis and review

AI starts from the workbook. Use `inspect_excel.py` to enumerate every score cell, intermediate result and total, identify the existing formulas and input cells, and record the full target set in `review/rule-plan.json` as `scope.score_targets`. This set, including score cells whose PDF evidence is still missing, is the coverage denominator. Then search the charter for the rule each target requires. Read relevant PDF pages using selectable text, rendering, table extraction or OCR as needed; check equations, symbols and tables against the page image. `inspect_pdf.py` is optional, and its text is never assumed complete. Use `inspect_report.py` for the report inventory. Preserve PDF page citations and short verified excerpts beside each workbook target.

Before the PDF search, write a human-readable and machine-readable Excel scoring
map as described in `excel-mapping.md`. It must identify the evaluation item,
specific scenario, performance/result represented by each input column, exact
input and output addresses, and the formula question for every score target.
Record real-entry formats and blank semantics as confirmed or pending. The PDF
then answers how the mapped performance becomes a score; it must not silently
redefine the workbook's fields.

Treat the workbook's existing input cells as the input contract. Before recording
an input-mapping question, inspect whether existing tiers or status columns
already express the PDF case, including zero-score outcomes. Do not convert a
PDF exception into a new input requirement merely because it has no dedicated
column. If no existing representation works for a required Excel target, leave
that target pending and document the specific gap; a proposed new cell or
external record is a schema change, not an automatic review option.

After defining the Excel scope, write a rule plan with `source_sha256`, `scope` (`type: workbook`, `workbook_sha256`, and `score_targets`) and a
nonempty `rules` array. Each rule needs `id`, `statement`, `target_cells`, `status` (`pending`
or `confirmed`), and `source_refs` from the inventory; use `title` and
`questions` where useful. For a pending business choice, use a structured question with `question`, `explanation`, and two or three `options`, each with `label` and `effect`; this keeps the Markdown record and chat choices aligned. Confirmed rules additionally need `reviewer` and
`reviewed_at`. AI may identify itself as reviewer when the PDF clearly supports the rule; do not attribute an AI confirmation to a person. Publish it with `publish_scoring_rules.py` to produce
`scoring_rules.json` and `open_questions.json`; use the readable `业务评审.md` as a user handoff only when substantive questions remain. These files carry the source hash
and `release_blocked` whenever a rule remains pending, a workbook score target is uncovered, or a cited PDF page needs review. The plan's `reviewed_pages` lists relevant pages actually checked. An unrelated PDF page or clause is not a missing Excel requirement and does not trigger a user question. This publisher checks
provenance, while AI remains responsible for source-supported interpretations and the user decides unresolved business alternatives. A rule plan may provide `evidence` records with a unique `ref`, one-based `page`, and short `text` excerpt read by AI. The publisher checks that the page exists, but cannot prove the excerpt matches the visible PDF; AI must verify it. Keep page coverage progress separate from questions that need a business decision.

After resolving decisions needed by the workbook targets, AI writes `评分细则.md` itself.
Follow `scoring-spec-writing.md`: organize it by the Excel score targets and use
only their relevant PDF evidence. A user-provided example may guide presentation,
but cannot add PDF-only scoring items to the workbook-defined scope or supply
another project's facts. A script does not generate this Markdown. Run
`validate_scoring_spec.py` to detect a stale PDF hash, missing rule IDs, or an
unmarked pending draft, then review the actual content, calculations and decisions
against the PDF and Excel map.

Use the already enumerated workbook targets to create a per-cell formula plan.
Verify measured and judged inputs, units, row/column semantics and any existing formula against the cited rule.
Bind the plan to the specification and structured rule hashes. Pending rules or
uncertain workbook mappings may be staged for analysis, but cannot be committed
to the project template. See `formula-authoring.md` for the commit gate and
native Excel boundary checks.

Create a concise business-facing review that includes:

- clause/table citation and the proposed rule;
- required input, unit, accepted range, missing-value behavior, thresholds, rounding, caps, overrides, and exceptions;
- current Excel behavior and the proposed field/output locations;
- report values and their source lineage;
- unresolved alternatives with example score differences when calculable.

Ask the user only where the supplied documents and verified evidence leave multiple plausible alternatives that change scores or report claims. At each of the PDF-rule, Excel-formula, and report-mapping stages, first write the source-bound questions to the project's review document, then ask them in the active conversation. Group three or four related questions in one batch when that many remain; use fewer when appropriate. Each question states the exact uncertainty, citation or cell/marker position, why it matters, and two or three mutually exclusive choices with their effects on score, formula or report. Include a free-text path if none fits. Do not treat a suggested choice as user consent. A user decision records the person, time, reason, chosen interpretation, and applicable SHA-256 values; preserve the exact text of a custom answer. Update the review document and structured decision record before continuing dependent work. AI confirms clear cases itself and records its identity, time, evidence and reasoning; it must never impersonate a human reviewer.

Before setting `release_status` to `business_validated`, record separate named, dated, reasoned confirmation records bound to all source hashes for scoring rules, Excel mapping, report mapping, and calculation authority. AI may record these for unambiguous, independently verified content under its own identity; this is automated provenance, not a claim of human sign-off. Every user-resolved business decision retains the user's identity, decision time, reason and source hashes. Missing real evidence or native Excel verification still blocks release.

## Generated project boundary

After formula authoring and native Excel validation, complete
`review/excel-map-plan.json` with exact input/output addresses and source refs.
Complete `review/report-map-plan.json` for every marker; use
`review/report-positions.csv` for table row/column context, then call
`map_report_template.py`. AI gives a position `review_status: confirmed` after verifying one clear meaning and source; ask the user only for unresolved alternatives. Complete `review/project-plan.json` with
the application ID and formula bindings. `freeze_project.py <project-root>`
checks the reviewed mappings, actual workbook formulas and named report before
writing a validated `engineering_sample` config. Windows + Excel calculation,
source-bound confirmation records and isolated onedir acceptance are separate later gates.

Create generated_apps/<project-name>/ with self-contained src/, config/, templates/, tests/, build.ps1, README.md, and delivery-manifest.json. Copy only required runtime components and templates. The application must run after the AutoTesting checkout is absent.

During generation, call the existing PDF, Excel, and DOCX analysis functions through the wrappers. Do not copy all of ivista_engine/ into the generated app. The app reads only native Excel formula results. It must not contain a Python scoring implementation.

## Completion status

Use separate statuses:

- engineering_sample: code/tests work on synthetic data; source-bound confirmations, real records, or target Excel recalculation remain missing.
- business_validated: named rule/mapping/report reviews are complete, the required real evidence is checked, output values are independently verified, and deployment conditions pass.

Never promote an engineering sample to business validated by changing a label.
