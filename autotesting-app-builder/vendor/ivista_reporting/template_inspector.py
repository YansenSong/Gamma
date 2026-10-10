"""Read-only inspection of user-supplied DOCX report templates and examples."""
from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import re
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

from docx import Document
from docx.document import Document as DocumentObject
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.section import _Header, _Footer

from .template_contract import TemplateContractError

MAX_DOCX_BYTES = 100_000_000
VARIABLE = re.compile(r"{{\s*([A-Za-z_][A-Za-z0-9_.]*)\s*}}")
PLACEHOLDER = re.compile(r"\[\[([^\]]+)\]\]|<<([^>]+)>>|【([^】]+)】")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _text_blocks(container, prefix: str) -> list[dict]:
    element = container.element.body if isinstance(container, DocumentObject) else container._element
    blocks: list[dict] = []
    paragraph_index = 0
    table_index = 0
    for child in element.iterchildren():
        if child.tag.endswith("}p"):
            paragraph_index += 1
            paragraph = Paragraph(child, container)
            text = paragraph.text
            if text.strip():
                blocks.append({"location": f"{prefix}.paragraph[{paragraph_index}]",
                               "style": paragraph.style.name if paragraph.style else None,
                               "text": text})
        elif child.tag.endswith("}tbl"):
            table_index += 1
            table = Table(child, container)
            rows = []
            for row_index, row in enumerate(table.rows, 1):
                cells = []
                for cell_index, cell in enumerate(row.cells, 1):
                    cells.append({"cell": cell_index, "text": cell.text,
                                  "nested_tables": _nested_tables(cell, f"{prefix}.table[{table_index}].row[{row_index}].cell[{cell_index}]")})
                rows.append({"row": row_index, "cells": cells})
            blocks.append({"location": f"{prefix}.table[{table_index}]", "rows": rows})
    return blocks


def _nested_tables(cell: _Cell, prefix: str) -> list[dict]:
    found = []
    for table_index, table in enumerate(cell.tables, 1):
        rows = []
        for row_index, row in enumerate(table.rows, 1):
            rows.append({"row": row_index, "cells": [
                {"cell": cell_index, "text": nested_cell.text}
                for cell_index, nested_cell in enumerate(row.cells, 1)
            ]})
        found.append({"location": f"{prefix}.nested_table[{table_index}]", "rows": rows})
    return found


def _paragraph_text(block: dict) -> list[tuple[str, str]]:
    result = []
    if "text" in block:
        result.append((block["location"], block["text"]))
    for row in block.get("rows", []):
        for cell in row["cells"]:
            result.append((block["location"] + f".row[{row['row']}].cell[{cell['cell']}]", cell["text"]))
            for table in cell.get("nested_tables", []):
                for nested_row in table["rows"]:
                    for nested_cell in nested_row["cells"]:
                        result.append((table["location"] + f".row[{nested_row['row']}].cell[{nested_cell['cell']}]", nested_cell["text"]))
    return result


def inspect_docx_template(path: Path | str, *, document_role: str) -> dict:
    """Inventory readable DOCX structure. Extracted sample values are never facts."""
    source = Path(path).expanduser().resolve()
    if source.suffix.lower() != ".docx" or not source.is_file():
        raise TemplateContractError("输入必须是可读的 .docx 文件")
    if document_role not in {"blank_template", "reference_sample"}:
        raise TemplateContractError("document_role 必须是 blank_template 或 reference_sample")
    size = source.stat().st_size
    if size > MAX_DOCX_BYTES:
        raise TemplateContractError("DOCX 文件超过读取上限")
    try:
        with ZipFile(source) as package:
            bad_part = package.testzip()
            if bad_part:
                raise TemplateContractError(f"DOCX 内部部件损坏：{bad_part}")
            names = package.namelist()
            xml_parts = {name: package.read(name) for name in names
                         if name.startswith("word/") and name.endswith(".xml")}
    except (OSError, BadZipFile) as error:
        raise TemplateContractError(f"DOCX 包无法读取：{error}") from error
    try:
        document = Document(source)
    except Exception as error:
        raise TemplateContractError(f"DOCX 无法解析：{error}") from error

    sections = []
    for section_index, section in enumerate(document.sections, 1):
        header = _text_blocks(section.header, f"section[{section_index}].header")
        footer = _text_blocks(section.footer, f"section[{section_index}].footer")
        sections.append({"section": section_index, "header": header, "footer": footer})
    body = _text_blocks(document, "body")
    all_text = []
    for block in body:
        all_text.extend(_paragraph_text(block))
    for section in sections:
        for block in (*section["header"], *section["footer"]):
            all_text.extend(_paragraph_text(block))

    variables: set[str] = set()
    placeholders: dict[str, list[str]] = {}
    for location, text in all_text:
        variables.update(VARIABLE.findall(text))
        for match in PLACEHOLDER.finditer(text):
            label = next((part.strip() for part in match.groups() if part), "")
            if label:
                placeholders.setdefault(label, []).append(location)
    # Scan each XML paragraph as its text nodes may split one variable across runs.
    jinja_xml = set()
    for part_name, xml in xml_parts.items():
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            continue
        for paragraph in root.iter(f"{{{W_NS}}}p"):
            joined = "".join(node.text or "" for node in paragraph.iter(f"{{{W_NS}}}t"))
            jinja_xml.update(VARIABLE.findall(joined))
    variables.update(jinja_xml)

    controls = []
    for part_name, xml in xml_parts.items():
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            continue
        for sdt in root.findall(f".//{{{W_NS}}}sdt"):
            alias = sdt.find(f".//{{{W_NS}}}alias")
            tag = sdt.find(f".//{{{W_NS}}}tag")
            controls.append({"part": part_name, "alias": alias.get(f"{{{W_NS}}}val") if alias is not None else None,
                             "tag": tag.get(f"{{{W_NS}}}val") if tag is not None else None})

    fields = [{"name": name, "kind": "jinja_variable", "locations": [],
               "data_source": None, "mapping_status": "needs_human_mapping"} for name in sorted(variables)]
    by_name = {item["name"]: item for item in fields}
    for location, text in all_text:
        for name in set(VARIABLE.findall(text)):
            if location not in by_name[name]["locations"]:
                by_name[name]["locations"].append(location)
    for name, locations in sorted(placeholders.items()):
        fields.append({"name": name, "kind": "visual_placeholder", "locations": sorted(set(locations)),
                       "data_source": None, "mapping_status": "needs_human_mapping"})

    image_parts = sorted(name for name in names if name.startswith("word/media/"))
    return {
        "schema_version": "autotesting-docx-template-inventory-v1",
        "source_file": source.name,
        "source_sha256": sha256(source.read_bytes()).hexdigest(),
        "file_size": size,
        "document_role": document_role,
        "sample_values_are_facts": False,
        "analysis_mode": "read_only_no_save",
        "paragraph_count": len(document.paragraphs),
        "table_count": len(document.tables),
        "section_count": len(document.sections),
        "body": body,
        "sections": sections,
        "field_candidates": fields,
        "content_controls": controls,
        "image_parts": image_parts,
        "has_macro_parts": any(name.lower().endswith("vbaproject.bin") for name in names),
        "mapping_requires_human_confirmation": True,
    }


__all__ = ["inspect_docx_template"]
