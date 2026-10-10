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

Extract the charter with inspect_pdf.py, the workbook with inspect_excel.py, and the report document with inspect_report.py. Preserve the extraction inventory and SourceRefs beside the generated project. AI may summarize and propose interpretations; scripts supply structural facts and enforce deterministic checks.

After reviewing the PDF inventory, write a rule plan with `source_sha256` and a
nonempty `rules` array. Each rule needs `id`, `statement`, `status` (`pending`
or `confirmed`), and `source_refs` from the inventory; use `title` and
`questions` where useful. Confirmed rules additionally need `reviewer` and
`reviewed_at`. Publish it with `publish_scoring_rules.py` to produce
`评分细则.md`, `scoring_rules.json`, and `open_questions.json`; the latter files carry the inventory hash
and `release_blocked` whenever any rule remains pending, a PDF page is unreviewed,
or an image-only page needs inspection. The plan's `reviewed_pages` lists pages
actually checked for rule coverage. This publisher checks
provenance, while AI and named reviewers remain responsible for rule meaning.

Create a concise business-facing review that includes:

- clause/table citation and the proposed rule;
- required input, unit, accepted range, missing-value behavior, thresholds, rounding, caps, overrides, and exceptions;
- current Excel behavior and the proposed field/output locations;
- report values and their source lineage;
- unresolved alternatives with example score differences when calculable.

Ask for a human decision wherever alternatives can change scores or report claims. A reviewer decision records the person, time, reason, chosen interpretation, and the SHA-256 values it applies to. A custom interpretation includes its exact text.

Before setting `release_status` to `business_validated`, also record named, dated, reasoned approvals bound to all source hashes for scoring rules, Excel mapping, report mapping, and calculation authority. Each resolved business decision must carry its own reviewer, decision time, reason, and source hashes.

## Generated project boundary

After formula authoring and native Excel validation, complete
`review/excel-map-plan.json` with exact input/output addresses and source refs.
Complete `review/report-map-plan.json` for every marker; use
`review/report-positions.csv` for table row/column context, then call
`map_report_template.py`. Give each position `review_status: confirmed` only
after checking its meaning and source. Complete `review/project-plan.json` with
the application ID and formula bindings. `freeze_project.py <project-root>`
checks the reviewed mappings, actual workbook formulas and named report before
writing a validated `engineering_sample` config. Windows + Excel calculation,
business approvals and isolated onedir acceptance are separate later gates.

Create generated_apps/<project-name>/ with self-contained src/, config/, templates/, tests/, build.ps1, README.md, and delivery-manifest.json. Copy only required runtime components and templates. The application must run after the AutoTesting checkout is absent.

During generation, call the existing PDF, Excel, and DOCX analysis functions through the wrappers. Do not copy all of ivista_engine/ into the generated app. The app reads only native Excel formula results. It must not contain a Python scoring implementation.

## Completion status

Use separate statuses:

- engineering_sample: code/tests work on synthetic data; business approvals, real records, or target Excel recalculation remain missing.
- business_validated: named rule/mapping/report reviews are complete, the required real evidence is checked, output values are independently verified, and deployment conditions pass.

Never promote an engineering sample to business validated by changing a label.
