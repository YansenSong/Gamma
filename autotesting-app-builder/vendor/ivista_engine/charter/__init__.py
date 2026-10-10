"""Deterministic, evidence-preserving PDF charter compiler.

VENDORED SUBSET — not the full AutoTesting package.
Source: D:/实习/汽研/测评/ivista_engine/charter

Upstream also re-exported ``reviewer`` and ``service`` (CharterCompiler),
which pull in rule synthesis and markdown services that the Skill does not
use. Only the PDF extraction path is kept here.
"""

from .pdf_extractor import PdfExtractor, PdfPlumberExtractor

__all__ = ["PdfExtractor", "PdfPlumberExtractor"]
