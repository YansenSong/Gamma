# Windows scoring app template

Copy this directory into a new generated project, then replace `config/project.json`, add `templates/`, and fill in the tests and examples. The runtime does not import AutoTesting modules.

## Filling in a new project

| Path | What to put there |
|---|---|
| `config/project.json` | sources, `scoring.inputs`, `excel_output`, `report_output` |
| `templates/` | the score workbook and the DOCX report template |
| `examples/input.json` | one realistic input record |
| `examples/boundary_cases.json` | optional — the charter's tier boundaries |
| `tests/` | keep `test_generated_application.py`; add project-specific cases |

`tests/test_generated_application.py` ships as a skeleton: it states the checks
every generated application must satisfy and **skips** until `config/project.json`
is filled in, so an unfilled copy reports "8 skipped" rather than failing. Run it
with the command the upstream generated apps record:

```
uv run --with pytest --with openpyxl --with docxtpl --with python-docx python -m pytest tests -q
```

The behavioural tests drive native Excel recalculation and therefore need
Microsoft Excel on the host — the same requirement the application itself has.


The app accepts a JSON test record, validates required values, writes them into a copy of the workbook, drives native Excel recalculation (`scripts/recalculate_excel.ps1`, which refuses to overwrite a cell that holds a formula), reads the resulting formula values back, renders a DOCX report, and records output hashes.

Excel formulas are the only scoring authority. Only `scoring.mode = "excel_native"` is accepted: the runtime contains no Python scorer, no expression evaluator and no fallback, and a mapped cell whose recalculated formula cache is missing is a blocker rather than a substituted number.
