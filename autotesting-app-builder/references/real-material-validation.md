# Real business evidence

Use existing real charter/workbook/report files as the primary acceptance evidence. Preserve synthetic cases as baseline tests. Start with parking, then security, then driving when these repository materials are selected.

Record three separate verdicts:

1. Legacy Excel consistency: source SHA, original formulas/cache and native Microsoft Excel version/results, using independent copies only.
2. Charter correctness: independent calculations traceable to PDF page and block/table anchors. List conflicting thresholds, caps, missing raw observations, rounding and interpretation questions. Legacy formulas are a comparison baseline and do not resolve conflicts.
3. Report fidelity: each variable's source, missing fields, static XML/table/header/footer preservation, old/new output differences and rendered page review.

Existing values lack verified trial provenance unless supplied evidence establishes it. Classify them as existing-workbook evidence; never invent trial dates, collisions, completion flags, approvals or metadata. Blank optional trials must remain distinct from zero scores. Boundary inputs on copies are explicitly labeled constructed tests.

For driving, retain all Q1–Q10 in `docs/评分细则候选/driving/business_decisions.json`, preserving pending decisions. Changing calculation authority cannot silently resolve them.

Generate project code from the three source documents and analysis. Old adapters/RulePacks may run only as comparison references. Missing mappings or review decisions require a blocked validation state and an actionable checklist. An engineering preview may fill mapped placeholders with “未提供”, but must refuse official release and disclose incomplete comparisons in its separate audit and UI. Do not add labels or alter static content in the DOCX. Never mark a whole category validated after checking a subset.

Keep native Excel results separate from LibreOffice-only regression results. Missing LibreOffice is an environment error, never a passed test. Deliver source hashes, per-cell comparisons, process logs, screenshots, page checks, self-contained onedir package and honest per-category status.
