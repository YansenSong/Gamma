# Project acceptance

## Engineering checks

- Generated applications use the IVISTA visual design system and pass native screenshot and interaction checks in ui-design-system.md.
- Primary acceptance uses real materials with separate legacy Excel, charter correctness and report fidelity verdicts. Pending mappings/decisions remain blocked; synthetic examples remain baseline tests.

- Project specification parses and binds to source hashes.
- Every mapped scoring output binds the exact formula in the project's `templates/score.xlsx` copy; Python scoring modes and expressions are rejected.
- Formulas the Skill authored into a bare workbook are listed with the charter source reference each transcribes and a per-cell formula plan. `scripts/diff_workbooks.py` shows the changed set is exactly the authorized cells, with every non-target cell, merged range, style, validation and sheet property unchanged. The project-copy write receipt records pre-write, backup and rewritten hashes; the rewritten hash binds the delivered workbook template. The user's source workbook retains its original hash.
- The report source is a blank template whose `{{ }}` marker set is non-empty, and after naming it equals `report_output.fields` exactly. The supplied template keeps its SHA-256; the named copy carries the bound `report_template` hash. A markerless DOCX, or one supplied as a completed report, is rejected. A field with no confirmed source stays visibly unfilled and blocks the verdict.
- Required, missing, invalid, and boundary values have explicit tests.
- Formula authoring updates only the `init_project.py` workspace copy at `templates/score.xlsx` after staged plan validation, native Excel boundary checks and a unique backup; the user's source Excel path remains untouched. Application runtime output is a separate copy. Verify exact target cells and unchanged non-target package members.
- Excel formulas are the sole authority. Perform native recalculation, verify saved results and compare against reviewed fixed boundary expectations. Python scoring and fallback results are absent.
- Stage Excel, DOCX, and audit output together and expose the bundle only after all three are complete.
- Report fields resolve only to declared, reviewed data sources, and every declared field has a marker in the template.
- Generated reports contain no unrendered placeholders.
- Every dynamic field has a traceable review disposition and readable checklist. Missing-value display and user entry do not constitute mapping approval; unsupported not-applicable claims remain blocked.
- Engineering or blocked reports show visible draft identity and result provenance on the first page and every subsequent page. Signature boxes remain unapproved. Review the latest rendered copy and retain the original template hash.
- Summarize independent charter statuses from per-rule rows, separately from selected legacy-output reproduction. Unknown boundaries and missing trial facts must remain visible blockers.
- Windows onedir package includes the Python runtime, configured templates, project configuration, the Qt Windows platform plugin, and any required native recalculation scripts; launch it from a directory independent of the AutoTesting source. Include an application icon when the project defines one.
- Release verification reruns project-spec validation and checks the current project hash, release state, sources, decisions, approvals, and manifest validation result.
- For a business-validated release, separately hash-bind confirmation records for scoring rules, Excel mappings, report mappings, and calculation authority. AI may confirm clear, independently verified items under its own identity; user decisions are required only for unresolved business alternatives. A content change invalidates the corresponding confirmation.

## Delivery manifest

Include project/app version, source-file hashes, business-decision and mapping reviewer details, rule/test coverage, full test result, workbook/report validation, build result, target OS, and known limitations. Mark an unapproved synthetic sample as engineering_sample.

Do not describe the result as business-ready when real evidence, a required native Excel recalc, or a necessary user decision is missing. Keep the output folder usable without the AutoTesting repository.
# Filled-workbook release gate

The default generated application must accept the operator's filled `.xlsx`
workbook. Run `run_job_from_workbook` on Windows with installed Microsoft Excel
and at least two copies whose input cells differ. Check that the report values
match the recalculated output cells in each submitted workbook copy, that the
reports differ as expected, and that each original upload's SHA-256 is unchanged.
Tamper a bound score formula, remove a required cached result, and make native
recalculation fail; all three cases must stop without a report bundle. Record
the real Excel version, executable hash, and isolated onedir run. A patched
`_excel_native_recalculate` is only a service contract test, never this gate.
