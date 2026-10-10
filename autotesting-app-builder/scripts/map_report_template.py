"""Bind every DOCX marker position to a named field and a declared data source."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree as ET
from report_markers import W_NS, extract_markers, sha256_file, template_error

FIELD = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")


def _replace(paragraph, offset, old, new):
    nodes = list(paragraph.iter(f"{{{W_NS}}}t"))
    texts = [node.text or "" for node in nodes]
    full = "".join(texts)
    if full[offset:offset + len(old)] != old:
        raise ValueError("DOCX position changed during naming")
    start, end = offset, offset + len(old)
    first = last = None
    position = 0
    for index, value in enumerate(texts):
        stop = position + len(value)
        if first is None and stop > start:
            first = index
            first_offset = start - position
        if position < end and stop >= end:
            last = index
            last_offset = end - position
            break
        position = stop
    if first is None or last is None:
        raise ValueError("DOCX marker has no text node")
    if first == last:
        nodes[first].text = texts[first][:first_offset] + new + texts[first][last_offset:]
    else:
        nodes[first].text = texts[first][:first_offset] + new
        for index in range(first + 1, last):
            nodes[index].text = ""
        nodes[last].text = texts[last][last_offset:]


def map_template(plan_path):
    plan = json.loads(Path(plan_path).read_text(encoding="utf-8"))
    source = Path(plan["template"]).resolve()
    output = Path(plan["output"]).resolve()
    manifest = Path(plan["manifest"]).resolve()
    if source == output or output.exists() or manifest.exists():
        raise ValueError("named template and manifest must be new paths")
    report = extract_markers(source)
    if report["unsupported"] or report["has_control_syntax"] or any(x["family"] != "corner" for x in report["bracketed"]):
        raise ValueError("unsupported template markers")
    if not report["positions"]:
        raise ValueError(template_error(report))
    if plan.get("source_sha256") != report["source_sha256"]:
        raise ValueError("template source SHA-256 mismatch")
    mapping = plan.get("positions")
    if not isinstance(mapping, list) or len(mapping) != len(report["positions"]):
        raise ValueError("every marker position must have exactly one mapping")
    by_id = {item["position_id"]: item for item in mapping}
    expected_ids = {item["position_id"] for item in report["positions"]}
    if len(by_id) != len(mapping) or set(by_id) != expected_ids:
        raise ValueError("unknown, duplicate or unmapped DOCX position")
    allowed = plan.get("available_sources", {})
    fields = {}
    for position in report["positions"]:
        item = by_id[position["position_id"]]
        field = item.get("field_id")
        provenance = item.get("source", {})
        if not isinstance(field, str) or not FIELD.fullmatch(field) or position["field_id"] not in (None, field):
            raise ValueError(f"invalid or conflicting field at {position['position_id']}")
        kind, key = provenance.get("kind"), provenance.get("id")
        if kind not in ("input", "score", "system") or key not in allowed.get(kind, []):
            raise ValueError(f"unresolved data source at {position['position_id']}")
        if field in fields and fields[field] != provenance:
            raise ValueError(f"field {field} has conflicting data sources")
        fields[field] = provenance
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(source) as old, ZipFile(output, "w") as new:
        for info in old.infolist():
            payload = old.read(info.filename)
            positions = [p for p in report["positions"] if p["part"] == info.filename]
            if positions:
                root = ET.fromstring(payload)
                paragraphs = list(root.iter(f"{{{W_NS}}}p"))
                for position in sorted(positions, key=lambda p: (p["paragraph"], p["offset"]), reverse=True):
                    item = by_id[position["position_id"]]
                    _replace(paragraphs[position["paragraph"]], position["offset"], position["marker"], "{{ " + item["field_id"] + " }}")
                payload = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
            new.writestr(info, payload)
    named = extract_markers(output)
    if named["slots"] or named["unsupported"] or set(named["names"]) != {item["field_id"] for item in mapping}:
        output.unlink(missing_ok=True)
        raise ValueError("named template failed marker verification")
    result = {"status": "engineering_sample", "template_sha256": report["source_sha256"],
              "named_template_sha256": sha256_file(output), "named_template": str(output),
              "fields": fields,
              "positions": [{"position_id": p["position_id"], "context": p["context"],
                             "field_id": by_id[p["position_id"]]["field_id"],
                             "source": by_id[p["position_id"]]["source"]} for p in report["positions"]]}
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("plan", type=Path)
    args = parser.parse_args()
    print(json.dumps(map_template(args.plan), ensure_ascii=False, indent=2))
