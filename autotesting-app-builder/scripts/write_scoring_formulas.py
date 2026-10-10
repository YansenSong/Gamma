"""Apply an externally reviewed formula plan to the original XLSX path.

The plan contains formula text and PDF evidence; this tool never derives business
rules or calculates scores. A native Excel validation receipt is mandatory.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
import uuid
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _sheet_parts(package):
    wb = ET.fromstring(package.read("xl/workbook.xml"))
    rels = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    result = {}
    for sheet in wb.findall("m:sheets/m:sheet", NS):
        target = targets[sheet.attrib[f"{{{NS['r']}}}id"]]
        result[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    return result


def _patch_xlsx(source, staged, entries):
    """Patch only worksheet XML members, retaining every other ZIP member byte content."""
    with ZipFile(source) as original, ZipFile(staged, "w") as output:
        parts = _sheet_parts(original)
        by_part = {}
        for entry in entries:
            if entry["sheet"] not in parts:
                raise ValueError(f"unknown sheet: {entry['sheet']}")
            by_part.setdefault(parts[entry["sheet"]], []).append(entry)
        for member in original.infolist():
            payload = original.read(member.filename)
            if member.filename in by_part:
                root = ET.fromstring(payload)
                for entry in by_part[member.filename]:
                    cell = root.find(f".//m:sheetData/m:row/m:c[@r='{entry['cell']}' ]", NS)
                    if cell is None:
                        raise ValueError(f"target cell {entry['sheet']}!{entry['cell']} has no XML record; style it in the template")
                    if cell.find("m:f", NS) is not None or cell.find("m:v", NS) is not None or cell.find("m:is", NS) is not None:
                        raise ValueError(f"target is not blank: {entry['sheet']}!{entry['cell']}")
                    cell.attrib.pop("t", None)
                    ET.SubElement(cell, f"{{{NS['m']}}}f").text = entry["formula"][1:]
                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            output.writestr(member, payload)


def apply(plan_path, receipt_path):
    plan = json.loads(Path(plan_path).read_text(encoding="utf-8"))
    path = Path(plan["workbook"]).resolve()
    if path.suffix.lower() != ".xlsx" or not path.is_file() or not os.access(path, os.W_OK):
        raise ValueError("workbook must be an existing writable .xlsx")
    before = digest(path)
    if before != plan.get("source_sha256"):
        raise ValueError("source SHA-256 changed")
    proof = plan.get("native_excel_validation", {})
    if proof.get("status") != "passed" or proof.get("source_sha256") != before or not proof.get("excel_version") or not proof.get("evidence_path") or not Path(proof["evidence_path"]).is_file():
        raise ValueError("Microsoft Excel native recalculation evidence is required")
    entries = plan.get("formulas", [])
    if not entries:
        raise ValueError("no formulas in plan")
    seen = set()
    book = load_workbook(path, read_only=False, data_only=False)
    try:
        for entry in entries:
            key = (entry["sheet"], entry["cell"])
            if key in seen or entry.get("original") is not None or not entry.get("pdf_refs") or not entry.get("inputs") or not entry.get("rule_explanation") or not entry.get("boundary_checks"):
                raise ValueError(f"incomplete or duplicate plan entry: {key}")
            seen.add(key)
            ws = book[entry["sheet"]]
            if ws.protection.sheet or ws[entry["cell"]].value is not None or ws[entry["cell"]].__class__.__name__ == "MergedCell":
                raise ValueError(f"protected, merged or occupied target: {key}")
            formula = entry["formula"]
            if not isinstance(formula, str) or not formula.startswith("="):
                raise ValueError(f"invalid formula: {key}")
            Tokenizer(formula)
    finally:
        book.close()
    with ZipFile(path) as archive:
        names = archive.namelist()
        if any(name.startswith(("xl/vbaProject", "xl/activeX/", "xl/externalLinks/", "xl/pivotTables/", "xl/slicers/", "xl/embeddings/", "xl/queryTables/")) or name == "xl/connections.xml" for name in names):
            raise ValueError("unsupported workbook feature; original retained")
    backup = path.with_name(path.name + f".backup-{uuid.uuid4().hex}")
    with path.open("rb") as inp, backup.open("xb") as out:
        shutil.copyfileobj(inp, out)
        out.flush(); os.fsync(out.fileno())
    if digest(backup) != before:
        raise ValueError("backup verification failed")
    staged = Path(tempfile.mkstemp(prefix=".gamma-formulas-", suffix=".xlsx", dir=path.parent)[1])
    try:
        _patch_xlsx(path, staged, entries)
        with ZipFile(path) as old, ZipFile(staged) as new:
            if new.testzip() is not None or set(old.namelist()) != set(new.namelist()):
                raise ValueError("staged archive integrity failed")
            changed = {name for name in old.namelist() if old.read(name) != new.read(name)}
            if changed != set(_sheet_parts(old)[entry["sheet"]] for entry in entries):
                raise ValueError("unexpected package member changed")
        check = load_workbook(staged, data_only=False)
        try:
            for entry in entries:
                if check[entry["sheet"]][entry["cell"]].value != entry["formula"]:
                    raise ValueError("staged formula mismatch")
        finally:
            check.close()
        if digest(path) != before:
            raise ValueError("source changed during staging")
        os.replace(staged, path)
        after = digest(path)
        check = load_workbook(path, read_only=True, data_only=False)
        try:
            if any(check[e["sheet"]][e["cell"]].value != e["formula"] for e in entries):
                raise ValueError("post-write formula mismatch")
        finally:
            check.close()
        receipt = {"status": "engineering_sample", "original_path": str(path), "original_sha256": before,
                   "backup_path": str(backup), "backup_sha256": digest(backup), "new_sha256": after,
                   "modified_cells": [{"sheet": e["sheet"], "cell": e["cell"], "formula": e["formula"], "pdf_refs": e["pdf_refs"]} for e in entries],
                   "native_excel_validation": proof}
        Path(receipt_path).write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return receipt
    finally:
        staged.unlink(missing_ok=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("plan", type=Path)
    parser.add_argument("--receipt", required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(apply(args.plan, args.receipt), ensure_ascii=False, indent=2))
