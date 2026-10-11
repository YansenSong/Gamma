# Excel analysis and mapping

Use the workbook analyzer as a read-only inventory. Record the source SHA-256, worksheets, cells, formulas, merged ranges, validations, names, protection, and unsupported features. AI checks labels, units, cell role, type and scoring relationship, and directly confirms a unique, source-supported match. Ask the user only when multiple plausible mappings remain or the materials omit a required meaning.

Keep mappings hash-bound to the exact workbook and rule specification. Outputs are written to a new copy. Preserve styles and unrelated values. Reject macro-enabled workbooks or workbook features the selected writer cannot preserve.

openpyxl does not calculate formulas. Microsoft Excel formulas are the only scoring authority: require native full recalculation, save and reopen the file, validate mapped formula results and record the Excel version. Missing results or errors block output; Python must not fill in substitute scores.

The generated project must make the selected calculation authority visible in its config and audit record.
