"""Only module that imports or touches third-party PDF objects."""

from __future__ import annotations

import hashlib
from collections import Counter
from dataclasses import replace
from pathlib import Path
from typing import Protocol

import pdfplumber

from .errors import CharterError
from .models import SourceBlock, SourceCell, SourceDocument, SourcePage, SourceTable, SourceWord
from .normalizer import normalize_text


def _box(value) -> tuple[float, float, float, float]:
    return tuple(round(float(number), 3) for number in value)


class PdfExtractor(Protocol):
    def extract(self, path: Path) -> SourceDocument: ...


class PdfPlumberExtractor:
    def __init__(self, *, max_file_size: int = 100_000_000, max_pages: int = 500, password: str | None = None):
        self.max_file_size = max_file_size
        self.max_pages = max_pages
        self.password = password

    def extract(self, path: Path) -> SourceDocument:
        path = Path(path)
        if not path.is_file() or path.suffix.lower() != ".pdf":
            raise CharterError("PDF_INPUT_INVALID", "extract", "输入必须是可读取的 PDF 文件", {"path": str(path)})
        if path.stat().st_size > self.max_file_size:
            raise CharterError("PDF_LIMIT_EXCEEDED", "extract", "PDF 文件超过大小限制", {"size": path.stat().st_size})
        source_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        try:
            with pdfplumber.open(path, password=self.password) as pdf:
                if len(pdf.pages) > self.max_pages:
                    raise CharterError("PDF_LIMIT_EXCEEDED", "extract", "PDF 页数超过限制", {"pages": len(pdf.pages)})
                pages = tuple(self._page(page, index + 1) for index, page in enumerate(pdf.pages))
                metadata = {str(key): str(value) for key, value in (pdf.metadata or {}).items() if value is not None}
        except CharterError:
            raise
        except Exception as error:
            name = type(error).__name__.lower()
            if "password" in name or "encrypted" in name:
                raise CharterError("PDF_PASSWORD_REQUIRED", "extract", "PDF 已加密，需要密码") from error
            raise CharterError("PDF_EXTRACTION_FAILED", "extract", "PDF 提取失败", {"errorType": type(error).__name__}) from error
        return self._mark_repeated_margins(SourceDocument(path.name, source_hash, len(pages), metadata, pages))

    @staticmethod
    def _page(page, page_number: int) -> SourcePage:
        prefix = f"p{page_number:04d}"
        lines = page.extract_text_lines(layout=False, return_chars=True) or []
        blocks = tuple(
            SourceBlock(
                ref=f"{prefix}-b{index:04d}", page=page_number, block_index=index,
                bbox=_box((line["x0"], line["top"], line["x1"], line["bottom"])),
                raw_text=str(line.get("text", "")), normalized_text=normalize_text(str(line.get("text", ""))),
                font_size=(round(float(line["chars"][0].get("size", 0)), 2) if line.get("chars") else None),
                font_name=(str(line["chars"][0].get("fontname", "")) if line.get("chars") else None),
            )
            for index, line in enumerate(lines, 1)
        )
        words = tuple(
            SourceWord(
                ref=f"{prefix}-w{index:04d}", page=page_number,
                bbox=_box((word["x0"], word["top"], word["x1"], word["bottom"])),
                raw_text=str(word["text"]), normalized_text=normalize_text(str(word["text"])),
            )
            for index, word in enumerate(page.extract_words(), 1)
        )
        tables = []
        for table_index, table in enumerate(page.find_tables(), 1):
            table_ref = f"{prefix}-t{table_index:04d}"
            matrix = tuple(tuple(str(cell) if cell is not None else None for cell in row) for row in table.extract())
            cells = []
            for row_index, row in enumerate(matrix, 1):
                pdf_row = table.rows[row_index - 1] if row_index - 1 < len(table.rows) else None
                for column_index, text in enumerate(row, 1):
                    if text is None:
                        continue
                    pdf_box = (pdf_row.cells[column_index - 1] if pdf_row and column_index - 1 < len(pdf_row.cells) else None)
                    cells.append(SourceCell(
                        ref=f"{table_ref}-r{row_index:04d}-c{column_index:04d}",
                        page=page_number, bbox=_box(pdf_box or table.bbox), row=row_index,
                        column=column_index, raw_text=text, normalized_text=normalize_text(text),
                    ))
            tables.append(SourceTable(table_ref, page_number, table_index, _box(table.bbox), matrix, tuple(cells)))
        return SourcePage(page_number, str(page_number), round(float(page.width), 3), round(float(page.height), 3), blocks, words, tuple(tables), len(page.images))

    @staticmethod
    def _mark_repeated_margins(document: SourceDocument) -> SourceDocument:
        occurrences: Counter[str] = Counter()
        for page in document.pages:
            for block in page.blocks:
                top, bottom = block.bbox[1], block.bbox[3]
                if top < page.height * .12 or bottom > page.height * .88:
                    occurrences[block.normalized_text] += 1
        threshold = max(3, int(document.page_count * .2))
        pages = []
        for page in document.pages:
            blocks = []
            for block in page.blocks:
                top, bottom = block.bbox[1], block.bbox[3]
                role = block.role
                if top < page.height * .12 and occurrences[block.normalized_text] >= threshold:
                    role = "header"
                elif bottom > page.height * .88 and (occurrences[block.normalized_text] >= threshold or block.normalized_text.isdigit()):
                    role = "footer"
                blocks.append(replace(block, role=role))
            pages.append(replace(page, blocks=tuple(blocks)))
        return replace(document, pages=tuple(pages))
