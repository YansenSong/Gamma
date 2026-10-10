# Vendored AutoTesting subset

This directory holds the minimum subset of the AutoTesting repository that the
Skill's inspection wrappers need. It exists because `autotesting-app-builder/`
was copied out of its parent checkout: `scripts/_bootstrap.py` used to resolve
the package root as `parents[4]`, which pointed at the repository root when the
Skill lived inside it.

It sits inside the Skill folder (`autotesting-app-builder/vendor/`), so the
Skill is one self-contained directory. Nothing here is generated: it is a frozen
copy, and the Skill reads it but never writes to it.

- Upstream repository: `D:/实习/汽研/测评`
- Upstream revision: `8d65005f66bc0bb942589613d3b65fb2d724fa6a` (2026-10-10)
- Vendored on: 2026-10-10

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
copied verbatim (both are empty upstream).

## Modifications from upstream

Only two files differ, both package initialisers that had to be trimmed:

- `ivista_engine/__init__.py` — dropped the `evaluate` / `load_rulepack`
  re-exports. Upstream's versions import `ivista_engine.rules.*`, which pulls in
  the whole rule engine and RulePack loader; the Skill's wrappers never call
  them, and the Skill is explicitly forbidden from shipping a Python scoring
  path (`references/scoring-spec.md`).
- `ivista_engine/charter/__init__.py` — dropped the `reviewer` and
  `CharterCompiler` re-exports for the same reason (rule synthesis, markdown
  services). `PdfExtractor` / `PdfPlumberExtractor` are retained.

Everything else is byte-identical to upstream. Upstream SHA-256 values:

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

This is a frozen copy. If the upstream packages change, re-copy the files above
and update this document. Upstream files that depend on the rule engine should
not be vendored: the generated applications must not contain a Python scoring
implementation.
