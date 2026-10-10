---
name: autotesting-app-builder
description: Analyze a scoring charter PDF, user-designed Excel workbook, and a blank DOCX inspection report template whose fill-in positions are marked with {{ }}, then help generate and validate a dedicated offline Windows scoring application. When the workbook does not yet contain per-item scoring formulas, derive them from the charter and write validated formulas back to the supplied Excel path after an immutable backup.
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
questions as evidence for humans to approve the workbook's formulas.

A supplied workbook is often a tally sheet whose per-item score cells are empty
or manually typed, with the scoring rules living only in the charter. In that case
derive the formulas from the charter and write validated formulas back to the supplied Excel path after an immutable backup. This
is authoring, not evaluation: Excel still computes every score, and a Python score
remains forbidden. Never invent a threshold, cap, tier or default the charter does
not state. Read references/formula-authoring.md.

The DOCX input is a blank template whose fill-in positions the user marked with
`{{ }}`. The brackets say "fill this in automatically" and carry no name of their
own; the mapping step assigns one. A DOCX with no markers is a completed report,
not a template, and is rejected. The named marker set is the authoritative list
of what gets filled and must equal the keys of `report_output.fields`. Read
references/report-mapping.md.

Analyze the three inputs read-only before formula authoring. The Excel source is changed only by the two-step validated writeback; keep its unique backup and receipt. Naming a template's slots
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

1. Check the external working environment and native Microsoft Excel availability; keep environments and generated projects outside this Skill folder.
2. Inspect all three inputs read-only and record source hashes. Extract PDF clauses with page and block references; map Excel inputs, blank scoring cells, existing detail formulas and aggregate formulas. Missing scoring formulas are normal input.
3. Derive formulas only for unambiguous rules. Record per-cell old content, formula, inputs, PDF references, explanation, boundary cases and risks. Preserve existing formulas; conflicts, ambiguous boundaries and missing inputs become blocking review questions.
4. Validate syntax, references, cycles, range boundaries and native Microsoft Excel recalculation on a staged workbook. Python must never calculate final scores. Run `write_scoring_formulas.py prepare` to create the staged workbook and hash manifest; then `commit` invokes Microsoft Excel directly for all boundary cases and only on success creates a unique backup, writes the original XLSX path and emits a hash receipt. Keep the original hash as provenance and bind the rewritten file hash to `sources[role=workbook]` and `excel_output.template`.
5. Inspect DOCX `{{}}`, `{{ }}` and named markers across body, tables, headers, footers and split runs. Reject markerless templates. Use `report_markers.py` position IDs and `map_report_template.py` to map every named and unnamed position to a field and declared source, then create a named DOCX copy. Unresolved sources block release.
6. Generate the standalone Windows app with a controlled copy of the rewritten workbook and named report template. Runtime must not depend on the original input paths or the Skill directory.
7. Test native Excel scoring boundaries, report rendering and isolated Windows EXE execution. Keep engineering samples blocked until business review and all required build evidence exist. Audit hashes, differences and source mappings before release.

## Tools in this skill

For business review handoff and evidence supplementation, read
references/review-workflow.md. Use scripts/review_workflow.py to export a
content-bound review package and revalidate submitted opinions without granting
approvals. Unscored review previews are available for blocked projects; keep
strict formula-cache validation while conditional empty results lack approval.

Run scripts from the Skill's scripts/ directory with the environment interpreter
described above. Start with the preflight:

- scripts/check_environment.py [--require-excel] — verify the vendored modules, declared dependencies, and whether an Excel calculation authority is reachable. A blocker here invalidates everything downstream, so read the `remediation` field it returns and follow its steps in order before running anything else.
- scripts/inspect_pdf.py <pdf> --out <json>
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

Read environment.md before the first run, workflow.md for the full handoff, formula-authoring.md when the workbook has no per-item scoring formulas, scoring-spec.md for rule/decision consistency, excel-mapping.md for template and recalculation constraints, report-mapping.md for the DOCX marker rules and mapping, acceptance.md before delivery, and migration.md to preserve the IVISTA compatibility boundary. Real repository PDF/Excel/DOCX materials are the primary validation evidence; read references/real-material-validation.md for separate Excel, charter and report verdicts and driving Q1–Q10 handling. Preserve synthetic examples as baseline tests.

The upstream repository keeps a worked example at
`generated_apps/autotesting-synthetic-demo/`. It is synthetic training data:
it demonstrates engineering flow only and is not approved for business scoring.
It is not part of this standalone copy, and `generated_apps/` does not exist here
until the first project is generated.
