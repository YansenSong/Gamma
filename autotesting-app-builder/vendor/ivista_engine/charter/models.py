"""JSON-safe, frozen artifacts for protocol authoring."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any


SOURCE_SCHEMA = "charter-source-v1"
CLAUSE_SCHEMA = "charter-clause-v1"
CANDIDATE_SCHEMA = "charter-candidate-v1"
COMPILER_VERSION = "charter-compiler-v1"
BBox = tuple[float, float, float, float]


@dataclass(frozen=True)
class SourceWord:
    ref: str
    page: int
    bbox: BBox
    raw_text: str
    normalized_text: str


@dataclass(frozen=True)
class SourceBlock:
    ref: str
    page: int
    block_index: int
    bbox: BBox
    raw_text: str
    normalized_text: str
    role: str = "body"
    heading_path: tuple[str, ...] = ()
    font_size: float | None = None
    font_name: str | None = None


@dataclass(frozen=True)
class SourceCell:
    ref: str
    page: int
    bbox: BBox | None
    row: int
    column: int
    raw_text: str
    normalized_text: str


@dataclass(frozen=True)
class SourceTable:
    ref: str
    page: int
    table_index: int
    bbox: BBox
    raw_matrix: tuple[tuple[str | None, ...], ...]
    cells: tuple[SourceCell, ...]
    continued_from: str | None = None


@dataclass(frozen=True)
class SourcePage:
    page_index: int
    page_label: str
    width: float
    height: float
    blocks: tuple[SourceBlock, ...]
    words: tuple[SourceWord, ...]
    tables: tuple[SourceTable, ...]
    image_count: int = 0


@dataclass(frozen=True)
class SourceDocument:
    filename: str
    sha256: str
    page_count: int
    metadata: dict[str, str]
    pages: tuple[SourcePage, ...]
    schema_version: str = SOURCE_SCHEMA
    compiler_version: str = COMPILER_VERSION


@dataclass(frozen=True)
class Clause:
    clause_id: str
    kind: str
    classification: str
    heading_path: tuple[str, ...]
    source_refs: tuple[str, ...]
    raw_text: str
    normalized_text: str
    table_matrix: tuple[tuple[str | None, ...], ...] = ()


@dataclass(frozen=True)
class Issue:
    code: str
    severity: str
    source_refs: tuple[str, ...]
    detail: str


@dataclass(frozen=True)
class RuleCandidate:
    candidate_id: str
    candidate_type: str
    title: str
    heading_path: tuple[str, ...]
    source_refs: tuple[str, ...]
    raw_evidence: tuple[str, ...]
    normalized_statement: str
    proposed_inputs: tuple[dict[str, Any], ...]
    proposed_calculation: dict[str, Any] | None
    proposed_output: str | None
    status: str  # proposed, unresolved, unsupported
    issues: tuple[str, ...] = ()


def jsonable(value: Any) -> Any:
    """Convert an artifact dataclass into plain JSON-compatible values."""
    return asdict(value)
