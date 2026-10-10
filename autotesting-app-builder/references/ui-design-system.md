# Default desktop design

The visual authority is `ivista_desktop/main.py`, `settings.py`, and actual running screenshots. The reusable template contains only visual components and project-scoped settings; it has no IVISTA scoring imports.

## Assets and configuration

- `src/ui_design_system/workbench.qss`: extracted IVISTA colors and control selectors; hero #12325b, background #f3f6fb, primary #246fdb, white cards with 14px corners.
- `components.py`: keyboard-accessible clickable drop area, drag state and local-file handling.
- `settings.py`: per-project output directory, completion action and always-on-top preferences.
- `application.ui.brand`, `category`, `subtitle` and optional `application.logo` are project values. Bundle configured assets; do not invent a business identity or copy IVISTA branding automatically.
- Two-column workbench: dynamic inputs and data import on the left; queued files, progress and outputs on the right. Preserve the 32/28px outer margins, 18px card gaps, 22/18px card padding and typography from the reference.

## Interaction acceptance

Run native Windows widgets and save screenshots of the original workbench, generated workbench, settings and completed/error states. Review images for clipping, colors, hierarchy and spacing. Widget construction alone is insufficient.

Exercise file picker, real Qt drag/drop events, duplicate/unsupported files, list clearing, import/edit/current-form generation, multiple-file processing, progress updates, failure recovery, preference persistence, result opening and output-directory opening. Long work runs off the GUI thread. Explain that batch jobs use file inputs while single-record jobs use current edited form values. Disable queue editing during work, and prevent closing a running worker.

Keep each complete workbook/report/audit bundle in a new numbered directory. Never overwrite existing outputs. Report batch failures individually and do not label failed records successful.
