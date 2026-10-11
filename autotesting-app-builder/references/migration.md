# Reuse and migration map

## Reused directly

- `ivista_engine.charter.pdf_extractor.PdfPlumberExtractor` is called by the Skill's PDF inspection wrapper. It retains page anchors and the source hash.
- `ivista_engine.authoring.excel_analyzer` is called by the workbook wrapper for sheet structure, formulas, validation, protection, merged cells, and unsupported features. `validate_workbook.py` also uses its `_archive_features` helper.
- The existing DOCX template inspector (`ivista_reporting.template_inspector`) is exposed through a wrapper. It reports structure and placeholder candidates; `scripts/report_markers.py` adds the bare `{{ }}` slots the frozen inspector cannot see.
- The Windows build flow reuses the repository's PyInstaller approach and Qt dependency family through a standalone per-project build script.

These four modules are reached through the frozen copy in `<skill-root>/vendor/`, not a live checkout; `scripts/_bootstrap.py` resolves that directory. The Skill no longer assumes it sits inside the AutoTesting repository. See environment.md.

## Abstracted for generated projects

- IVISTA front-end colors, card layouts and drop-area interaction are extracted into the standalone UI design system. Names, logos, categories and dynamic fields come from each project; no fixed scoring/report business code is reused.

- Scoring uses only native Excel formulas. Input validation, Excel automation and report generation are standalone; no Python scorer is copied.
- Excel output copies a user-reviewed template and writes only explicit mappings into a separate output file. Formula workbooks require native Excel recalculation before their cached formulas can be trusted.
- When the supplied workbook has no per-item scoring formulas, the Skill derives them from the charter and writes them only to the `init_project.py` workspace copy at `templates/score.xlsx` after staged validation, native Excel boundary checks and a unique backup. The user's source workbook remains unchanged; the formulas remain Excel's, not Python's. Read formula-authoring.md.
- DOCX output uses explicit source mappings and Jinja placeholders. The report input is a blank template marked with `{{ }}`; a completed report is not accepted. Each marker is mapped and reviewed before it can be filled.
- Project validation binds each approved business choice to a canonical Excel formula binding hash. Changing the selected choice, formula or mapping invalidates approval.

## Legacy-only behavior

- `ivista_desktop`, its fixed vehicle categories, IVISTA-specific report layout, existing RulePacks, and the current EXE entry point remain compatibility workflows. They are not copied into generated applications and should not be generalized by changing their fixed business behavior.

## Migration path

1. Retain old IVISTA code and artifacts as historical compatibility references; do not use their Python scoring path in the Skill or generated applications.
2. Create a separate `generated_apps/<project-id>/` for each new charter/workbook/report combination.
3. Review rule interpretation, Excel cell mapping, report fields, calculation authority, and source hashes before marking any decision confirmed.
4. Run generated tests and inspect outputs. Promote to `business_validated` only after real evidence and named reviews meet `references/acceptance.md`.
