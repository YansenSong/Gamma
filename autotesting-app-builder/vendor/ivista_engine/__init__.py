"""RulePack driven scoring engine.

VENDORED SUBSET — not the full AutoTesting package.
The upstream __init__ re-exported ``evaluate`` / ``load_rulepack`` from
``.rules``, which drags in the whole rule engine and the rulepack loader.
The Skill's inspection wrappers only need ``charter.pdf_extractor`` and
``authoring.excel_analyzer``, so those re-exports were removed here to keep
this copy self-contained. Import the submodules directly.
"""

__all__: list[str] = []
