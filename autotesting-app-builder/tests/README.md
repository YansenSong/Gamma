# Skill tests

Tests for the Skill's own scripts and for the frozen `vendor/` subset.

## Running

From `<skill-root>`:

```
../.venv/Scripts/python.exe -B -m unittest discover -s tests -t tests
```

`-B` is not optional. Without it Python writes `__pycache__` into the Skill
folder, and the folder has to stay byte-for-byte re-copyable
(`references/environment.md`).

Standard library `unittest` only, matching the command the upstream AutoTesting
repository documents for its own suite (`uv run python -m unittest discover -s
tests`). No test dependency is added to `requirements.txt`.

## What is covered

| File | Covers |
|---|---|
| `test_docx_inspector.py` | vendored `ivista_reporting.template_inspector` + `template_contract` |
| `test_excel_authoring.py` | vendored `ivista_engine.authoring.excel_analyzer` + `errors` |
| `test_project_spec_validator.py` | `scripts/validate_project_spec.py` |
| `test_skill_review_workflow.py` | `scripts/review_workflow.py`, `scripts/build_review_inventory.py`, `assets/windows-app-template/src/review_preview.py` |
| `test_report_markers.py` | `scripts/report_markers.py` — bare-slot detection and the template-error rules |
| `test_name_report_template.py` | `scripts/name_report_template.py` — slot naming, and the original keeping its hash |

The two vendored-module tests pin the frozen contract on purpose: `vendor/` is a
copy that must stay byte-identical to upstream (`vendor/VENDOR.md`), so a failure
there means someone edited the frozen subset.

## Provenance and adaptations

The first four files are ported from the upstream repository's `tests/` directory;
each names its upstream original in its module docstring. The last two cover
scripts that are Skill-only — `report_markers.py` and `name_report_template.py`
do not exist upstream — so those tests are new.

For the ported files, three kinds of change were required and nothing else was
touched:

1. **Path resolution.** Upstream located the Skill through the checkout at
   `.agents/skills/autotesting-app-builder/`, which does not exist once the
   Skill ships standalone. Paths now resolve inside `<skill-root>`.
2. **`sys.path`.** The two vendored-module tests need `vendor/` importable.
3. **Fixtures the newer validator rejects.** The Skill's validator gained
   `docx_markers` / `validate_report_template` after these tests were written,
   so it now parses the report template and requires a non-empty
   `report_output.fields`. Two upstream fixtures stubbed the DOCX as raw bytes
   or declared `fields: {}`; both were updated, and each says so in its
   docstring.

## Not ported

Upstream tests that depend on the repository layout, its rule engine, or the
real PDF corpus cannot come along: `test_authoring_review.py`,
`test_charter_compiler.py`, `test_production_boundary.py`. There is no isolated
upstream test of `vendor/.../pdf_extractor.py` at all — it is only exercised
through the full charter pipeline.

Two upstream tests cover Skill scripts but are written for **pytest**
(`tmp_path` fixture and `@pytest.mark.parametrize`): `test_release_verification.py`
and `test_skill_report_review.py`. They are not ported yet, so
`scripts/verify_release.py`, `scripts/write_manifest.py`,
`scripts/export_field_review.py` and `assets/.../src/report_safety.py`
currently have no coverage here.
