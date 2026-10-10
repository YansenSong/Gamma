"""Read-only inventory for human designed .xlsx assessment templates."""

from __future__ import annotations

from datetime import date, datetime, time
import hashlib
import json
import re
from pathlib import Path
import zipfile
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from .errors import AuthoringError


_CELL_REF = re.compile(r"(?:(?:'([^']+)'|([A-Za-z0-9_][A-Za-z0-9_. ]*))!)?(\$?[A-Z]{1,3}\$?\d+)")
_OUTPUT_WORDS = ("得分", "总分", "合计", "汇总", "平均", "评分", "分值", "分数", "等级")
_INPUT_WORDS = ("输入", "次数", "成绩", "结果", "距离", "时间", "速度", "数值", "是否", "实测", "记录", "判定")
_MAX_FILE_SIZE = 100_000_000
_MAX_CELLS = 500_000


def _plain(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _color(color):
    if color is None:
        return None
    if color.type == "rgb":
        return color.rgb
    if color.type == "indexed":
        return {"indexed": color.indexed}
    if color.type == "theme":
        return {"theme": color.theme, "tint": color.tint}
    return {"type": color.type}


def _formula_references(formula: str) -> list[dict]:
    refs = []
    for match in _CELL_REF.finditer(formula):
        sheet = match.group(1) or match.group(2)
        refs.append({"sheet": sheet.replace("''", "'").strip() if sheet else None, "cell": match.group(3).replace("$", "")})
    return refs


def _archive_features(path: Path) -> list[str]:
    features = []
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
    except (OSError, zipfile.BadZipFile) as error:
        raise AuthoringError("EXCEL_ARCHIVE_INVALID", "inspect", "输入不是有效的 xlsx ZIP 工作簿") from error
    checks = (
        ("xl/vbaProject.bin", "VBA_MACRO"),
        ("xl/externalLinks/", "EXTERNAL_LINKS"),
        ("xl/pivotTables/", "PIVOT_TABLES"),
        ("xl/slicers/", "SLICERS"),
        ("xl/activeX/", "ACTIVEX_CONTROLS"),
        ("xl/embeddings/", "EMBEDDED_OBJECTS"),
        ("xl/queryTables/", "QUERY_TABLES"),
        ("xl/connections.xml", "DATA_CONNECTIONS"),
        ("xl/charts/", "CHARTS"),
        ("xl/drawings/", "DRAWINGS"),
        ("xl/media/", "IMAGES_OR_MEDIA"),
        ("xl/comments", "COMMENTS"),
        ("xl/threadedComments/", "THREADED_COMMENTS"),
    )
    for prefix, code in checks:
        if any(name.startswith(prefix) for name in names):
            features.append(code)
    return features


def _validation_for(ws, coordinate: str) -> list[dict]:
    found = []
    row, column = coordinate_to_tuple(coordinate)
    for validation in ws.data_validations.dataValidation:
        applies = False
        for area in validation.sqref.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(area))
            if min_row <= row <= max_row and min_col <= column <= max_col:
                applies = True
                break
        if applies:
            found.append({
                "type": validation.type, "operator": validation.operator,
                "formula1": validation.formula1, "formula2": validation.formula2,
                "allow_blank": validation.allow_blank, "error": validation.error,
                "prompt": validation.prompt,
            })
    return found


def _merge_context(ws, coordinate: str, merged_lookup: dict[str, str]) -> str | None:
    anchor = merged_lookup.get(coordinate)
    if not anchor:
        return None
    return _plain(ws[anchor].value)


def analyze_workbook(path: Path, *, max_file_size: int = _MAX_FILE_SIZE, max_cells: int = _MAX_CELLS) -> dict:
    """Extract workbook structure without mutating or saving the source file."""
    path = Path(path)
    if not path.is_file() or path.suffix.lower() != ".xlsx":
        raise AuthoringError("EXCEL_INPUT_INVALID", "inspect", "只支持可读取的 .xlsx 文件")
    size = path.stat().st_size
    if size > max_file_size:
        raise AuthoringError("EXCEL_LIMIT_EXCEEDED", "inspect", "Excel 文件超过大小限制", {"size": size})
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    feature_flags = _archive_features(path)
    try:
        workbook = load_workbook(path, data_only=False, read_only=False, keep_links=True)
    except Exception as error:
        raise AuthoringError("EXCEL_READ_FAILED", "inspect", "Excel 文件无法由当前读取器安全解析", {"errorType": type(error).__name__}) from error

    sheets = []
    all_cells = []
    cell_count = 0
    for sheet_index, ws in enumerate(workbook.worksheets, 1):
        declared_cells = ws.max_row * ws.max_column
        if declared_cells + cell_count > max_cells:
            workbook.close()
            raise AuthoringError("EXCEL_LIMIT_EXCEEDED", "inspect", "工作簿有效区域超过单元格扫描上限", {"limit": max_cells, "sheet": ws.title})
        merged_lookup: dict[str, str] = {}
        merged_ranges = []
        for merged in ws.merged_cells.ranges:
            bounds = range_boundaries(str(merged))
            min_col, min_row, max_col, max_row = bounds
            anchor = ws.cell(min_row, min_col).coordinate
            merged_ranges.append(str(merged))
            merged_lookup[anchor] = anchor
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    merged_lookup[ws.cell(row, col).coordinate] = anchor
        table_refs = {table.ref: table.displayName for table in ws.tables.values()}
        validation_cells: set[str] = set()
        validations = []
        for validation in ws.data_validations.dataValidation:
            validations.append({
                "type": validation.type, "operator": validation.operator,
                "formula1": validation.formula1, "formula2": validation.formula2,
                "allow_blank": validation.allow_blank, "sqref": str(validation.sqref),
                "error": validation.error, "prompt": validation.prompt,
            })
            for area in validation.sqref.ranges:
                min_col, min_row, max_col, max_row = range_boundaries(str(area))
                if (max_row - min_row + 1) * (max_col - min_col + 1) <= 5000:
                    for row in range(min_row, max_row + 1):
                        for col in range(min_col, max_col + 1):
                            validation_cells.add(ws.cell(row, col).coordinate)

        sheet_cells = []
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                coordinate = cell.coordinate
                validations_here = _validation_for(ws, coordinate) if coordinate in validation_cells else []
                if value is None and not cell.has_style and not validations_here:
                    continue
                cell_count += 1
                is_formula = isinstance(value, str) and value.startswith("=")
                text_value = str(value) if isinstance(value, str) else ""
                header_above = []
                for row_index in range(cell.row - 1, max(0, cell.row - 6), -1):
                    above = ws.cell(row_index, cell.column).value
                    if above is not None:
                        header_above.append(str(above)[:240])
                labels_left = []
                for col_index in range(cell.column - 1, max(0, cell.column - 13), -1):
                    left = ws.cell(cell.row, col_index).value
                    if left is not None:
                        labels_left.append(str(left)[:240])
                merged_label = _merge_context(ws, coordinate, merged_lookup)
                context_parts = [str(item) for item in [merged_label, *reversed(header_above), *reversed(labels_left), value] if item is not None]
                context = " | ".join(context_parts)
                context_lower = context.casefold()
                if is_formula:
                    role = "formula_output_candidate"
                elif text_value and any(word in text_value for word in _OUTPUT_WORDS):
                    role = "score_label_candidate"
                elif validations_here or value is None and any(word in context for word in _INPUT_WORDS):
                    role = "input_candidate"
                elif value is not None and any(word in context for word in _OUTPUT_WORDS):
                    role = "scoring_value_candidate"
                else:
                    role = "other"
                cell_record = {
                    "field_id": f"excel.{digest[:12]}.s{sheet_index:03d}.{coordinate}",
                    "sheet": ws.title, "sheet_index": sheet_index, "cell": coordinate,
                    "value": _plain(value), "value_type": "formula" if is_formula else type(value).__name__ if value is not None else "blank",
                    "formula_references": _formula_references(value) if is_formula else [],
                    "number_format": cell.number_format if cell.has_style else None,
                    "style": {"bold": cell.font.bold, "font_size": cell.font.sz,
                              "fill": _color(cell.fill.fgColor), "horizontal_alignment": cell.alignment.horizontal,
                              "locked": cell.protection.locked} if cell.has_style else None,
                    "merged_anchor": merged_lookup.get(coordinate),
                    "merged_label": _plain(merged_label),
                    "headers_above": list(reversed(header_above)), "labels_left": list(reversed(labels_left)),
                    "validations": validations_here, "role_candidate": role,
                    "context": context[:3000], "search_text": context_lower[:3000],
                }
                sheet_cells.append(cell_record)
                all_cells.append(cell_record)
        sheets.append({
            "sheet": ws.title, "sheet_index": sheet_index, "state": ws.sheet_state,
            "max_row": ws.max_row, "max_column": ws.max_column,
            "merged_ranges": merged_ranges, "tables": table_refs,
            "data_validations": validations,
            "protection": {"sheet": ws.protection.sheet, "objects": ws.protection.objects, "scenarios": ws.protection.scenarios},
            "hidden_rows": [str(index) for index, dim in ws.row_dimensions.items() if dim.hidden],
            "hidden_columns": [key for key, dim in ws.column_dimensions.items() if dim.hidden],
            "cells": sheet_cells,
        })

    defined_names = []
    for name in workbook.defined_names.values():
        defined_names.append({"name": name.name, "formula": name.attr_text, "scope": name.localSheetId, "hidden": name.hidden})
    workbook.close()
    duplicate_index: dict[str, list[dict]] = defaultdict(list)
    formula_edges = []
    for cell in all_cells:
        value = cell.get("value")
        if isinstance(value, str) and cell.get("role_candidate") == "score_label_candidate":
            normalized = " ".join(value.casefold().split())
            if len(normalized) >= 2:
                duplicate_index[normalized].append({"sheet": cell["sheet"], "cell": cell["cell"], "field_id": cell["field_id"]})
        for reference in cell.get("formula_references", []):
            formula_edges.append({"from": {"sheet": cell["sheet"], "cell": cell["cell"]},
                                  "to": {"sheet": reference.get("sheet") or cell["sheet"], "cell": reference["cell"]}})
    duplicate_labels = [{"label": label, "cells": locations} for label, locations in sorted(duplicate_index.items()) if len(locations) > 1]
    return {
        "schema_version": "assessment-workbook-inventory-v1", "source_file": path.name,
        "source_sha256": digest, "file_size": size, "reader": "openpyxl",
        "analysis_mode": "read_only_no_save", "write_supported": False,
        "unsupported_feature_flags": feature_flags,
        "sheet_count": len(sheets), "nonempty_or_styled_cell_count": cell_count,
        "defined_names": defined_names, "duplicate_labels": duplicate_labels,
        "formula_dependency_edges": formula_edges, "sheets": sheets,
        "cells": all_cells,
    }


def write_inventory(path: Path, output: Path, *, max_file_size: int = _MAX_FILE_SIZE, max_cells: int = _MAX_CELLS) -> dict:
    inventory = analyze_workbook(path, max_file_size=max_file_size, max_cells=max_cells)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(inventory, ensure_ascii=False, sort_keys=True, indent=2) + "\n", encoding="utf-8", newline="\n")
    return inventory
