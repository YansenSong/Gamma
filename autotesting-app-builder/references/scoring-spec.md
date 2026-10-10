# Excel scoring specification

Each generated project uses `scoring.mode = excel_native`. There is no Python
score evaluator, expression tree, reference score or fallback authority.

A supplied workbook often contains no per-item scoring formulas at all: the score
cells are bare and only aggregate `SUM` rows carry formulas, if anything. Bindings
cannot exist until the formulas do. Derive them from the charter and write them
into an authorized copy, then bind as usual — read formula-authoring.md. Authoring
does not shift authority: Excel evaluates every score either way, and a formula
the Skill wrote is subject to the same review, source references and approval as
one the user supplied.

Declare inputs for data validation and `scoring.formula_bindings` for every
`excel_output.output_cells` entry. Each binding records output ID, sheet, cell,
exact formula and charter/source references. The validator compares that formula
with the hash-bound workbook. Empty bindings are allowed only in blocked review
projects without output mappings; such projects cannot perform scoring.

```json
{
  "mode": "excel_native",
  "authority": {"type": "Microsoft Excel", "require_full_recalculation": true},
  "inputs": [{"id": "latency_ms", "type": "number", "minimum": 0}],
  "formula_bindings": [{
    "output": "score", "sheet": "Score", "cell": "B6",
    "formula": "=IF(B4<=200,100,IF(B4<=300,80,60))",
    "source_refs": ["synthetic charter, page 1"]
  }]
}
```

This is a synthetic illustration, not approved business logic. Human decisions
bind the canonical SHA-256 of sheet/cell/formula through `decision_bindings`.
Source hashes and four scoped approval content hashes must cover current material.
A reviewer label does not alter a formula. Update an explicitly authorized
workbook copy and obtain new approvals when a business decision changes formulas.

Read scoring results only after native Excel full recalculation, workbook save,
and reopening. Record Excel version and actual workbook/report hashes. Reject
missing, nonfinite, erroneous or wrong-type mapped results. Never replace a
failed Excel calculation with a computed Python number or a sample value.

Charter review may identify threshold, unit, cap, exception and rounding conflicts.
It records human-review evidence; it does not run a second scoring implementation.
