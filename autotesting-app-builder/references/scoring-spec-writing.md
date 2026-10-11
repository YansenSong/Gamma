# AI writing of the scoring specification

`评分细则.md` is an authored business document organized by the supplied Excel
score targets. AI first lists each target and its formula needs, then reads the
corresponding PDF evidence, rule plan and completed business review. AI confirms
clear interpretations directly and asks the user only about ambiguity that changes
a target score, indispensable input mapping or required report claim. No script
assembles its paragraphs, tables, formulas, ordering, or level of detail.

Read `examples/README.md` and the bundled
`examples/IVISTA_2026_智能行车评分细则整理.md` as a style and depth reference,
not as a source of facts for a new project. If the user supplies a different
example, follow that preference. The bundled example's useful
pattern is a complete score framework followed by prerequisites, input and unit
definitions, detailed per-item decisions, formulas, boundaries, caps, exceptions,
and concrete tables for tiers and scenarios. Adapt the outline to the actual
charter; do not force unrelated charters into those headings.

For every workbook score target, make it possible for a reviewer to determine:

- the exact sheet/cell, its input addresses, and the PDF page/table supporting it;

- what is measured, which valid test event supplies it, and its unit;
- how each threshold and equality boundary is treated;
- how scores, deductions, rounding, caps and totals are calculated;
- what happens when an input is absent, invalid, exempt or contradictory;
- which PDF page supports the interpretation and which business review decision
  resolves any ambiguity.

Record other PDF scoring clauses as out of scope if useful for audit, without
turning their absence from this workbook into a clarification request. A target
whose rule is missing or unsupported stays visibly pending.

Use tables for repeated scenarios or tiers when they improve readability. Include
the PDF SHA-256 and stable rule IDs somewhere in the document so a separate check
can detect stale or omitted content. A pending decision must be visibly marked;
do not present the document as final while the relevant rule remains unresolved.

Run `scripts/validate_scoring_spec.py <评分细则.md> <scoring_rules.json>` after
writing. It checks source identity, rule ID coverage and a visible draft label.
It cannot assess prose quality, correct formulas, or whether a reviewer decision
was faithfully applied. Review those against the source PDF and signed opinions.
