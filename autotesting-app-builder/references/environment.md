# Environment and self-containment

The Skill was originally authored inside the AutoTesting checkout, where the
analysis helpers could import `ivista_engine` and `ivista_reporting` from the
repository root. It now ships standalone. Read this before running anything.

## Layout

The Skill folder holds plain files only: the instructions, the vendored subset,
the wrappers and the assets. Copying that one directory to another host carries
the whole Skill. The single thing that cannot travel is the interpreter, so it is
deliberately kept outside the folder:

```
<parent>/                      working directory
  .venv/                       toolchain interpreter (uv-managed, NOT portable)
  autotesting-app-builder/     <skill-root> — the sealed, re-copyable folder
    SKILL.md
    .gitignore
    agents/
    assets/
    references/
    scripts/
    vendor/                    frozen AutoTesting subset
      VENDOR.md                bundled module inventory and per-file SHA-256
      ivista_engine/
      ivista_reporting/
    requirements.txt
```

`<skill-root>` is the Skill folder. User documents (charter PDF, workbook,
report) are deliberately *not* part of it — they arrive from the user and are
analysed read-only, so keep them wherever the user keeps their material.

Nothing generated may be written inside `<skill-root>`: no `.venv/`, no
`__pycache__/`, no `generated_apps/`, no `build/` or `dist/`. `.gitignore`
enforces this. A virtual environment bakes in absolute paths and a reference to
the base interpreter it was built from, so copying it to another machine or user
profile yields an interpreter that resolves to the wrong paths — which is why it
lives beside the Skill rather than in it. The vendored modules, scripts and
references are plain files and move freely.

## How the helpers find their imports

`scripts/_bootstrap.py` walks up from its own location until it finds a `vendor/`
directory holding `ivista_engine/`, inserts it on `sys.path`, and exposes it as
`REPO_ROOT`. It raises `ModuleNotFoundError` when no such directory exists. Do
not reintroduce a hard-coded `parents[N]`: the Skill must keep working when the
whole `<skill-root>` tree is moved or copied.

`validate_workbook.py` and `validate_report.py` re-insert `REPO_ROOT` themselves
because they import the analysis modules directly rather than through the
inspection wrappers.

## Vendored subset

`vendor/` holds eight inspection modules plus package initialisers.
`vendor/VENDOR.md` lists every implementation file, its SHA-256, and the two
initialisers with deliberately limited imports.

Two upstream initialisers were emptied because they re-exported the rule engine
(`ivista_engine.rules.*`, `CharterCompiler`), which the Skill must never carry:
generated applications are forbidden from containing a Python scoring path. Only
`charter.pdf_extractor`, `authoring.excel_analyzer` and
`reporting.template_inspector` are needed.

`vendor/` is a frozen copy, not a live dependency. When upstream changes, re-copy
the listed files and update `VENDOR.md`. Never vendor rule-engine modules.

## Creating the working environment

This is the first thing to do in a session, before any user document is read.
The Skill works in a dedicated interpreter of its own — never a global one — and
that interpreter is not shipped with the Skill. Create it as a sibling of the
Skill folder so the sealed folder stays clean. Run both commands from the parent
directory that holds `<skill-root>`:

```
uv venv .venv --python 3.13
uv pip install --python .venv/Scripts/python.exe -r autotesting-app-builder/requirements.txt
```

The first command builds the interpreter; the second needs network access,
because `uv` resolves and downloads the wheels. After that the Skill runs
offline.

Then invoke wrappers from inside `<skill-root>` as
`../.venv/Scripts/python.exe -B scripts/<wrapper>.py`. The `-B` matters: it stops
Python writing `__pycache__` into the sealed folder.

`requirements.txt` covers the inspection wrappers (pdfplumber, openpyxl,
python-docx), the generated application runtime (PySide6, docxtpl) and packaging
(PyInstaller). Every PySide6 and PyInstaller pin in this tree names the same
version, so the interpreter that smoke-tests a generated app is the one that
builds it.

Anywhere outside `<skill-root>` is acceptable for this environment; `../.venv` is
the convention the rest of these documents assume. Only PySide6 and PyInstaller
are needed at packaging time — the inspection wrappers alone need pdfplumber,
openpyxl and python-docx.

## Preflight

Run `scripts/check_environment.py` before a session, from the environment just
created. It reports the interpreter, the resolved `REPO_ROOT`, whether both
vendored packages import, which declared dependencies are missing, and whether
Microsoft Excel automation is reachable. Treat a missing dependency or an
unreachable calculation authority as a blocker, not a warning: several acceptance
checks cannot be satisfied without them.

Check the reported `python.executable`. If it is not the environment you created,
the wrappers are running under the wrong interpreter and every downstream verdict
is suspect.

The preflight also returns a `remediation` object carrying the interpreter it
actually used, the one it expected, and an ordered list of fix steps. When
`status` is `blocked`, work through those steps rather than guessing — they are
written for the failure that was actually detected, so a vendored import failure
caused by a missing package is not reported as a damaged vendor file.

## What self-containment does not cover

These modules are bundled copies, so tests or conclusions from another checkout
do not automatically transfer. Record the bundled file hashes from `VENDOR.md`
in delivery evidence when relevant, and re-run `check_environment.py` when the
Skill moves to another host.
