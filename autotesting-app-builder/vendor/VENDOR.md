# Bundled inspection modules

This directory holds the small set of inspection modules required by the Skill.
`scripts/_bootstrap.py` resolves them from this directory, so a live source
checkout is not needed.

It sits inside the Skill folder (`autotesting-app-builder/vendor/`). Nothing
here is generated during project work; the Skill reads these frozen modules but
does not write to them.

## Why each file is here

| Vendored file | Required by |
|---|---|
| `ivista_engine/charter/pdf_extractor.py` | `scripts/inspect_pdf.py` |
| `ivista_engine/charter/errors.py` | `pdf_extractor` |
| `ivista_engine/charter/models.py` | `pdf_extractor` |
| `ivista_engine/charter/normalizer.py` | `pdf_extractor` |
| `ivista_engine/authoring/excel_analyzer.py` | `scripts/inspect_excel.py`, `scripts/validate_workbook.py` |
| `ivista_engine/authoring/errors.py` | `excel_analyzer` |
| `ivista_reporting/template_inspector.py` | `scripts/inspect_report.py`, `scripts/validate_report.py` |
| `ivista_reporting/template_contract.py` | `template_inspector` |

`ivista_engine/authoring/__init__.py` and `ivista_reporting/__init__.py` are
empty package initialisers.

## Deliberately limited imports

Two package initialisers expose only the inspection modules needed here:

- `ivista_engine/__init__.py` — dropped the `evaluate` / `load_rulepack`
  re-exports. Those imports pull in `ivista_engine.rules.*`, which includes
  the whole rule engine and RulePack loader; the Skill's wrappers never call
  them, and the Skill is explicitly forbidden from shipping a Python scoring
  path (`references/scoring-spec.md`).
- `ivista_engine/charter/__init__.py` — dropped the `reviewer` and
  `CharterCompiler` re-exports for the same reason (rule synthesis, markdown
  services). `PdfExtractor` / `PdfPlumberExtractor` are retained.

SHA-256 values for the eight bundled implementation files:

```
eaef5ba4656094fae9b60edc3f8c35586f179d85acb1474f48ed84491fc34b56  ivista_engine/charter/pdf_extractor.py
1a72a3a4492d5e6d5a41effa103536615b445b0b5b7511404f14e595a5a2a552  ivista_engine/charter/errors.py
7ece8dc8b635305d69e9802e88f8d2842748489b8068ba13e35ce9e8e213abf0  ivista_engine/charter/models.py
d6dd6999c00773e0a3c22d7232ed61291ddac411c146ead3f17af0f4d60cfce5  ivista_engine/charter/normalizer.py
3ca39c57847832f60c45c952caf1e46e6fcda31a82130bfea8031388119b8b69  ivista_engine/authoring/excel_analyzer.py
69c1343a7833feb295b00a0cbd39dc3ce85212dab3695b9186202411480ad6f1  ivista_engine/authoring/errors.py
ff80fa913372441c2071d69048f412e346a64985cab35662329427966edf6ec1  ivista_reporting/template_inspector.py
a991632b4348819d55db495de0fe4168cc47d01fb138a01e4edbadcbd5d7295f  ivista_reporting/template_contract.py
```

## Currency

This is a frozen copy. If these modules are updated, verify their imports and
update the hashes above. Do not add modules that bring in a Python scoring
implementation; generated applications use Excel formulas as their authority.
