# DOCX report mapping

The DOCX input is a blank template whose fill-in positions are marked with `{{ }}`. The brackets are the author's way of saying "this position has to be filled automatically"; they carry no variable name until the mapping step assigns one. The inspector is read-only and reports paragraphs, tables, section headers/footers, content controls, the marker set, and a source hash. A DOCX with no markers is a completed report, not a blank template, and is not an accepted input.

For each marked position, propose a mapping and assign the field name that replaces the bare `{{ }}`. scripts/name_report_template.py writes the named copy as a new file, so the supplied template keeps its hash. A field is reportable only after its source is identified in the reviewed project specification (for example, a confirmed Excel score cell or a measured input). The template's named marker set is the authoritative list of what gets filled: it must equal the keys of `report_output.fields` exactly, in both directions.

Use docxtpl for supported Jinja fields such as {{ score }}. Only simple `{{ field_id }}` markers are supported — no `{% %}` control statements and no dotted or computed expressions. Render a copy, validate that no template markers remain, reopen it with python-docx, and visually review pages when a renderer is available. Preserve static text and layout. Do not silently replace an unrecognized field with an empty string or invented value.

## Markers

`{{ }}` is the only accepted marker. The other bracket families the inspector reports are treated differently on purpose: `[[ ]]` and `<< >>` are placeholder syntax belonging to another system and must be converted to `{{ }}` or removed, while `【 】` is ordinary Chinese punctuation (`【注意事项】`) and is reported for information only, never as a blocker. `{% %}` and `{# #}` are rejected outright.

An unnamed `{{ }}` slot reaching a rendered report is a blocker, not a blank to tidy away: docxtpl would emit the brackets as literal text and the position would stay visibly unfilled.

## Verify before the template is released

- The supplied template keeps its SHA-256; the named copy is a separate file and carries the bound `report_template` hash. The pristine original's hash is recorded in the analysis inventory, outside `sources`.
- The named marker set equals `report_output.fields` exactly. A marker no field declares would render as an empty string; a declared field with no marker is a value the author believes is reported but that never appears.
- No unnamed `{{ }}` slot remains.
- Render a copy with representative data, reopen it, and confirm no marker survives and no position rendered empty. Review the rendered pages, not just the JSON.
- Static text, tables, merged cells, section breaks, headers and footers, page-count text, images and signature blocks are intact from the supplied template.

## Field review inventory

Inventory every variable once, retaining every Word location. Each item must include field_id, word_locations, meaning, suggested_source, source, evidence, affects_score, affects_conclusion, review_status, missing_reason, blocks_official, question and category. Keep baseline_missing separately when optional user entry is added. Export JSON plus Markdown/CSV with scripts/export_field_review.py; CSV reviewer/decision/reason cells are proposals for subsequent review, not imported approvals.

Use A for real information to be supplied by the user, B for explicitly identified existing-file candidates pending semantic/unit/source review, C for not-applicable dispositions with a named reviewer and evidence, and D for missing trial/conclusion facts. Category and official blocking are separate axes: A and B can still block. Never classify absent results as not applicable without evidence. Scores cannot establish completion or collision facts. Optional form entries remain user-supplied unverified data until reviewed; reset them when importing another record.

## Visible engineering copy

For engineering_sample or blocked projects, use the template's report_safety helper on generated copies: visible engineering identity on the first page, persistent header/footer notices on every page, and a source disclaimer alongside key scores/grades. Configure the exact source disclaimer in each project; read scoring values only from native Excel formula results. Mark signature boxes as unapproved; never fill reviewers, conclusions or dates from examples. The original DOCX must keep its SHA-256. Record changed parts and any preserved erroneous page-count text. Render the latest copy with the selected native tool and inspect every page; a JSON release flag is insufficient.

Expose report_review.path in project configuration to show the read-only, searchable FieldReviewDialog and export the checklist. Optional data-entry fields are declared in scoring.inputs with required=false; they must not become scoring facts or approval records merely through user entry.
