# Generation workflow

## Inputs and identity

Record each source file's absolute input role, SHA-256, format, size, and readability in config/project.json. The DOCX must be a blank template whose fill-in positions are marked with `{{ }}`; reject a markerless document as a completed report, which is not an accepted input. Inventory the marker set and carry it forward as the report field list. Keep business names and scoring claims grounded in the supplied material.

## Analysis and review

Extract the charter with inspect_pdf.py, the workbook with inspect_excel.py, and the report document with inspect_report.py. Preserve the extraction inventory and SourceRefs beside the generated project. AI may summarize and propose interpretations; scripts supply structural facts and enforce deterministic checks.

Create a concise business-facing review that includes:

- clause/table citation and the proposed rule;
- required input, unit, accepted range, missing-value behavior, thresholds, rounding, caps, overrides, and exceptions;
- current Excel behavior and the proposed field/output locations;
- report values and their source lineage;
- unresolved alternatives with example score differences when calculable.

Ask for a human decision wherever alternatives can change scores or report claims. A reviewer decision records the person, time, reason, chosen interpretation, and the SHA-256 values it applies to. A custom interpretation includes its exact text.

Before setting `release_status` to `business_validated`, also record named, dated, reasoned approvals bound to all source hashes for scoring rules, Excel mapping, report mapping, and calculation authority. Each resolved business decision must carry its own reviewer, decision time, reason, and source hashes.

## Generated project boundary

Create generated_apps/<project-name>/ with self-contained src/, config/, templates/, tests/, build.ps1, README.md, and delivery-manifest.json. Copy only required runtime components and templates. The application must run after the AutoTesting checkout is absent.

During generation, call the existing PDF, Excel, and DOCX analysis functions through the wrappers. Do not copy all of ivista_engine/ into the generated app. The app reads only native Excel formula results. It must not contain a Python scoring implementation.

## Completion status

Use separate statuses:

- engineering_sample: code/tests work on synthetic data; business approvals, real records, or target Excel recalculation remain missing.
- business_validated: named rule/mapping/report reviews are complete, the required real evidence is checked, output values are independently verified, and deployment conditions pass.

Never promote an engineering sample to business validated by changing a label.
