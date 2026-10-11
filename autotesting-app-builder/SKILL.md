---
name: autotesting-app-builder
description: Analyze a user-designed Excel scoring workbook, its PDF scoring charter, and a blank DOCX inspection report template marked with {{ }}, then generate and validate a dedicated offline Windows scoring application. Start from every Excel score target, locate its PDF evidence, author the needed scoring specification and stage native Excel formulas in a protected project copy.
---

# AutoTesting App Builder

Use this skill when a user wants a Windows desktop scoring program generated from their own business documents. Do not invoke it for ordinary edits to the existing IVISTA scoring workflows.

This skill ships standalone. Its analysis helpers run against a vendored copy of the AutoTesting modules rather than the repository. The Skill folder itself holds plain files only, so it stays byte-for-byte re-copyable; the Python environment it works in is created separately, before any work starts. Read references/environment.md and create that environment first.

## Product boundary

The skill guides AI during application generation. The generated application runs deterministically without an AI service, network access, or the AutoTesting repository. Treat existing IVISTA RulePacks and fixed report layouts as compatibility examples unless the user explicitly selects them.

Excel formulas are the sole scoring authority. Python only validates inputs,
automates native Excel recalculation, reads results, fills DOCX and records audit
evidence. Do not generate a Python scoring evaluator, a parallel reference-score
algorithm, or a fallback score when Excel is unavailable. Preserve charter review
questions as evidence. AI confirms rules and formulas with clear source support; ask the user only when the materials leave a score-changing interpretation unresolved.

The delivered workbook is itself an operator-facing scoring artifact. If people
can edit it directly in Excel, its formulas must accept every documented raw
entry form without the app's Python normalization. A successful app run on a
normalized copy does not establish that the workbook calculates correctly by
itself. If the workbook is intentionally usable only through the app, make that
restriction explicit in the handoff and prevent direct-entry claims; do not
silently rely on preprocessing to rescue formulas that score common entries
incorrectly.

A supplied workbook is often a tally sheet whose per-item score cells are empty
or manually typed, with the scoring rules living only in the charter. In that case
first enumerate every Excel scoring item, its specific scenario/performance columns,
existing input cells, score target, and formula question. Document how an operator
will fill those inputs in a real submitted workbook before interpreting blanks or
writing formulas. Read references/excel-mapping.md.
For each target, find the relevant PDF clause, table and prerequisites; then author
`评分细则.md` around those targets and draft formulas for the project copy. Commit only confirmed,
native-validated formulas after an immutable backup. This
is authoring, not evaluation: Excel still computes every score, and a Python score
remains forbidden. Never invent a threshold, cap, tier or default the charter does
not state. Read references/formula-authoring.md.

The DOCX input is a blank template whose fill-in positions the user marked with
`{{ }}`. The brackets say "fill this in automatically" and carry no name of their
own; the mapping step assigns one. A DOCX with no markers is a completed report,
not a template, and is rejected. The named marker set is the authoritative list
of what gets filled and must equal the keys of `report_output.fields`. Read
references/report-mapping.md.
Generated reports may change only those mapped marker contents. Preserve every
other word, style, header, footer, section, image and signature area. Do not
inject engineering labels or disclaimers into the DOCX; record release and
review status in the separate audit and application UI.

Analyze the three inputs read-only before formula authoring. Start a new workspace with `scripts/init_project.py` and author formulas only in its copied `templates/score.xlsx`; keep the user's source workbook unchanged. The project copy is changed only by the two-step validated writeback; keep its unique backup and receipt. Naming a template's slots
writes a separate, hash-bound copy, so the supplied template keeps its SHA-256. A
completed report is not an accepted input at all, so its values never enter
scoring, inputs, or outputs.

## Before you start: create the working environment

This skill works in a Python environment of its own, and nothing else in this
document runs until that environment exists. **Create it before you touch any
user document**, not on demand partway through a run.

It is deliberately not shipped inside the Skill folder: a virtual environment
bakes in absolute paths from the machine that built it, and the Skill folder has
to stay re-copyable. Create it as a sibling of `<skill-root>` — run both commands
from the parent directory that holds the Skill:

```
uv venv .venv --python 3.13
uv pip install --python .venv/Scripts/python.exe -r autotesting-app-builder/requirements.txt
```

Creating it needs network access the first time, because `uv` resolves and
downloads the wheels; everything afterwards runs offline. Then confirm it before
starting real work:

```
../.venv/Scripts/python.exe -B scripts/check_environment.py
```

Invoke every wrapper the same way — `../.venv/Scripts/python.exe -B
scripts/<wrapper>.py` from inside `<skill-root>`. The `-B` is not optional: it
stops Python writing `__pycache__` into the sealed folder. A blocker from this
preflight invalidates every step downstream, so fix the environment before
proceeding. references/environment.md covers the same ground in more detail,
including what to do when the Skill moves to another host.

## Workflow

The Excel score-target **and existing input-cell** inventory defines the scoring scope. The PDF supplies calculation rules for those cells; it does not enlarge the workbook schema. Do not add an input column, score target, trial record, or operator question merely because the PDF names a separate event or exception. First check whether the workbook already represents it through a status choice, tier count (including a zero-point tier), or a derived input. If an exception is decided by the operator before entering an existing count, document that classification rule instead of adding a field. If an existing Excel target truly cannot be calculated from existing inputs, record the exact target and missing dependency as a blocker; propose a workbook change only when the user requests a schema change or explicitly authorizes one after seeing that evidence. Treat the three review gates below in that order: workbook cell mapping, PDF interpretation needed by those cells, then report field mapping. An uncited PDF page outside this scope is not a release blocker; an uncovered Excel score target is.

AI owns the routine analysis, source-grounded confirmations, implementation and verification. There are conditional user decision points for Excel target/input mapping, PDF interpretation needed by those targets, and Excel-result-to-DOCX field mapping. At each point, exhaust the supplied evidence first and continue directly when no material ambiguity remains. For unresolved points, first save a source-bound question record in the project, then ask the user interactively in the current conversation; a file alone is not a handoff. Present three or four questions per batch when available (fewer if fewer exist). For each, explain the ambiguity and its source, offer two or three concrete, mutually exclusive choices with the consequence of each, and allow a free-text correction. Do not preselect an answer merely to keep work moving. Record the selected or custom answer, its rationale and applicable source hashes, update the question record, then resume dependent work. Continue independent work while waiting. Never manufacture missing measurements or trial records; request those only when the task actually needs them. Record AI confirmations as AI, never as human sign-off.

1. Check the external working environment and native Microsoft Excel availability; keep environments and generated projects outside this Skill folder. Use `init_project.py` to create a new review workspace from any three conforming documents, without domain presets.
2. Inspect all three inputs read-only and record source hashes. Start with Excel: make a row-by-row scoring map of evaluation item, specific test item/scenario, performance or outcome represented by each input column, bound input cells, score-rate and score cells, units, existing formulas, and missing formulas. Record all score targets in `review/rule-plan.json` under `scope.score_targets`; treat this list as the completeness denominator. Include intermediate scores and totals the workbook actually requires. For **every human-filled input**, work out what an operator is likely to enter in a completed copy of this particular workbook, which representations are equivalent, what an empty cell means in context, and what should be rejected. Use workbook labels, formats, validation rules, examples and available filled records as evidence; mark uncertainty pending. Freeze this input contract alongside the cell map before reading PDF rules. Do not infer a score target or required input solely from a PDF clause. Read references/excel-mapping.md.
3. For each Excel target, state what a formula needs to decide using the inventoried inputs, then locate its PDF clause, table, prerequisites and exceptions. AI chooses how to read the relevant PDF pages: selectable text, rendered pages, tables, or OCR where needed. Check equations and symbols against page images. Record page numbers and short excerpts. Write a detailed `评分细则.md` organized around workbook targets, confirm source-supported rules directly, and draft formulas when inputs and the rule support one result. Bind each formula to target cells, rule IDs, the exact scoring-specification hash, PDF locations, inputs, boundary cases and risks. Preserve existing formulas. Before asking about a supposedly missing input, show why no existing cell or category can encode the needed result and why the target cannot otherwise be calculated. Ask the user only about unresolved choices that change a workbook score, required report claim, or indispensable mapping of an existing input. Log PDF rules outside workbook scope as out of scope; their absence from this workbook is not itself a question or release blocker.
4. Validate syntax, references, cycles, range boundaries and native Microsoft Excel recalculation on a staged workbook. Python must never calculate final scores. Run `write_scoring_formulas.py prepare`; AI records its own confirmation for each unambiguous formula with source evidence. Ask the user only about formula or cell-mapping ambiguities the materials cannot resolve. Include representative completed-workbook entries for every distinct input contract: equivalent valid representations, meaningful empty cells, missing required data, invalid entries and boundary values. `commit` requires the bound specification and structured rules, confirmed rule IDs, and recorded per-formula review. It invokes Microsoft Excel for all boundary cases and only on success creates a unique backup, writes the project workbook copy and emits a hash receipt. Keep the original hash as provenance and bind the rewritten file hash to `sources[role=workbook]` and `excel_output.template`.
   Before commit, run each documented raw entry form directly in the staged workbook through native Excel, and compare its result with the app's normalized-copy path. Include a whole-row zero, a wholly blank row, a zero-point occurrence, and unused alternatives wherever those states are possible. A score-changing state absent from the charter, such as zero occurrences of a required scenario, stays pending until the user decides its meaning; permission to consider multiple input formats is not permission to choose a business score. Check every score and aggregate for Excel error values, not only the selected cell.
5. Inspect DOCX `{{}}`, `{{ }}` and named markers across body, tables, headers, footers and split runs. Reject markerless templates. Use `report_markers.py` position IDs and `map_report_template.py` to map every named and unnamed position to a field and declared source, then create a named DOCX copy. AI confirms mappings supported by context and source lineage directly; ask the user only when multiple plausible meanings or sources remain. Unresolved sources block release.
6. Generate the standalone Windows app with a controlled copy of the rewritten workbook and named report template. The default runtime path accepts the operator's filled workbook and calls `run_job_from_workbook`; adapt the copied runtime to validate and normalize its declared input cells on a safe copy according to the project-specific input contract before native Excel recalculation. The bundled generic runtime cannot infer that contract. Read report scores only from the recalculated copy, and audit raw and normalized input values. `run_job_from_facts` is a legacy form/JSON route. Runtime must not depend on the original input paths or the Skill directory.
7. Test two differently filled workbooks and the realistic entry variants found for every human-filled field in this workbook. Verify equivalent entries produce equivalent Excel results, meaningful blanks behave as documented, invalid or ambiguous entries fail at the exact cell, normalization changes only a safe copy, and raw/canonical values are audited. Also test unchanged upload hashes, formula tampering, missing scores, native Excel boundaries, report rendering and isolated Windows EXE execution. A mock recalculation is engineering evidence only. Keep engineering samples blocked until business review and all required build evidence exist. Audit hashes, differences and source mappings before release.

## 交付文件定位

完成一个生成项目后，在最终答复中分别给出可点击的**实际绝对路径**：

- Windows 程序：`generated_apps/<项目名>/dist/<应用名>/<应用名>.exe`。同时说明必须交付整个 `dist/<应用名>/` 目录，不能只复制 EXE。
- 已填公式的评分表：读取该项目 `config/project.json` 的 `excel_output.template`，给出所指向的 `generated_apps/<项目名>/templates/*.xlsx` 实际文件。不能假定文件名恒为 `score.xlsx`。

核对两处文件确实存在，且评分表的 SHA-256 与项目配置一致。若某项尚未生成或验证失败，要明确说明，不能给出示例路径冒充交付路径。应用、Excel 模板、报告及审计文件的用途也应写入生成项目的 `README.md` 或交付清单，方便使用者查找。

## Tools in this skill

For business review handoff and evidence supplementation, read
references/review-workflow.md. Use scripts/review_workflow.py to export a
content-bound review package and revalidate submitted opinions without granting
approvals. Unscored review previews are available for blocked projects; keep
strict formula-cache validation while conditional empty results lack approval.

Run scripts from the Skill's scripts/ directory with the environment interpreter
described above. Start with the preflight:

- scripts/check_environment.py [--require-excel] — verify the vendored modules, declared dependencies, and whether an Excel calculation authority is reachable. A blocker here invalidates everything downstream, so read the `remediation` field it returns and follow its steps in order before running anything else.
- scripts/inspect_pdf.py <pdf> --out <json> — optional text/table inventory when it helps; verify its output against the PDF where layout matters.
- scripts/init_project.py <pdf> <xlsx> <docx> --out <new-project-directory> --name <project-name> — copy three sources, produce inventories and unresolved review plans, and copy the generic runtime template. This is `review_required`, not a generated release.
- scripts/export_report_positions.py <docx> --json <positions.json> --csv <positions.csv> — include table row, column and neighboring text for each anonymous marker; all mapping columns remain pending.
- scripts/publish_scoring_rules.py <pdf-source.json> <reviewed-rule-plan.json> --rules <scoring_rules.json> --questions <open_questions.json> — publish source-bound rules and verify coverage of every `scope.score_targets` Excel cell. In workbook scope, only cited PDF pages require review; each rule lists `target_cells`. It creates a readable `业务评审.md` only when questions remain. For interactive review, write each question as an object with `question`, `explanation`, and two or three `options` containing `label` and `effect`; the same choices appear in the document and should be asked in chat. It does not generate `评分细则.md`. AI may add `evidence` entries containing `ref`, `page`, and a short verified `text` excerpt to the rule plan. The tool checks page identity and references but cannot independently verify an AI transcription. `confirmed` requires reviewer and reviewed_at; AI may use its own identity for clear cases. Pending rules and uncovered workbook targets block release.
- AI writes `评分细则.md` directly from the Excel target inventory, corresponding PDF evidence, rule plan and completed business review. Organize the document by the workbook's scoring sections and target cells; explain only the PDF clauses needed to calculate those targets. Do not turn a detailed PDF inventory or another project's scoring document into the outline. Read `references/scoring-spec-writing.md`. A user-provided example may guide presentation only and cannot enlarge the Excel-defined scope or supply scoring facts. Run `scripts/validate_scoring_spec.py <评分细则.md> <scoring_rules.json>` for source and coverage checks; it does not judge business accuracy.
- scripts/freeze_project.py <project-root> — after formula workbook and all report positions have confirmed mappings, bind their hashes into a validated `engineering_sample` project.json. It never grants business approval.
- scripts/inspect_excel.py <xlsx> --out <json>
- scripts/inspect_report.py <docx> --out <json> — rejects a DOCX with no {{ }} markers and reports the marker set, so the positions to fill are visible up front.
- scripts/name_report_template.py <plan.json> [--dry-run] — assign field names to a blank template's {{ }} slots; reports every unnamed slot as a blocker.
- scripts/write_scoring_formulas.py prepare <plan.json> --staged <stage.xlsx> --manifest <stage.json>; then `commit <plan.json> --manifest <stage.json> --evidence <native.json> --receipt <receipt.json>`. Commit invokes Windows PowerShell and Microsoft Excel; Linux can prepare and test blocking paths, but cannot claim native scoring validation.
- scripts/map_report_template.py <plan.json> — require a mapping for every DOCX position, validate each source, and write a named template plus manifest.
- scripts/diff_workbooks.py <original.xlsx> <authored.xlsx> --expect-changed Sheet1!N4,... — prove an authorized formula edit touched only the intended cells.
- scripts/validate_project_spec.py <project.json>
- scripts/validate_workbook.py <source.xlsx> --expected-sha256 <hash>. Pass --require-recalculated whenever scores come from formulas — the normal case — because that is what turns a missing recalculated cache into a hard failure instead of a note in the report. Omit it only for a workbook with no formula score cells at all (a blocked review project with empty formula_bindings), where no cache exists to require.
- scripts/validate_report.py <report.docx> --rendered for a final generated report.
- scripts/write_manifest.py <project-root> --tests-status passed|failed|not_run --tests-command "..." --build-status passed|failed|not_available|not_run [--smoke-status passed|failed|not_run] [--executable dist/App/App.exe] [--known-limit "..."]
- scripts/verify_release.py <project-root>

The Skill has its own test suite in tests/, covering these wrappers and the frozen
vendored subset. Run it with `../.venv/Scripts/python.exe -B -m unittest discover
-s tests -t tests` after changing any script, and always with `-B` so no
`__pycache__` lands in the sealed folder. See tests/README.md.

Read environment.md before the first run, workflow.md for the full handoff, formula-authoring.md when the workbook has no per-item scoring formulas, scoring-spec.md for rule/decision consistency, excel-mapping.md for template and recalculation constraints, report-mapping.md for the DOCX marker rules and mapping, acceptance.md before delivery, and migration.md to preserve the IVISTA compatibility boundary. Real repository PDF/Excel/DOCX materials are the primary validation evidence; read references/real-material-validation.md for separate Excel, charter and report verdicts and driving Q1–Q10 handling. Keep synthetic test fixtures only as engineering regression cases, never as sources for a new project's scoring rules.
