"""Canonical detector for the DOCX report template's fill-in markers.

The Skill accepts exactly one DOCX input: a *blank template* whose fill-in
positions are marked with ``{{ }}``. The brackets are the user's way of saying
"this part has to be filled automatically" — they carry no variable name. The
name is assigned later, in the mapping step, producing ``{{ score }}`` which is
what docxtpl actually renders.

    {{}}          bare slot   — a position the user marked as "fill me in"
    {{ score }}   named slot  — the same position after mapping assigned a name

The vendored inspector cannot see a bare slot: its pattern
(``vendor/ivista_reporting/template_inspector.py``) requires at least one
identifier character between the braces. That module is a frozen copy the Skill
never writes to, so bare-slot detection lives here instead.

Standard library only. ``assets/windows-app-template/src/project_validation.py``
carries a mirrored copy of the extractor because the generated application must
run with the Skill tree absent — it cannot import this module.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Deliberately identical in meaning to the frozen engine's pattern, except that
# dots are excluded: review_preview.py rejects anything but simple fields, so a
# dotted name would render in the app but fail in the engineering preview.
VARIABLE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")
SLOT = re.compile(r"\{\{\s*\}\}")
CONTROL = re.compile(r"\{%|\{#")

# Bracketed markers the engine's PLACEHOLDER pattern also recognises. Kept in
# separate families because they do not deserve the same treatment: [[ ]] and
# << >> are deliberate placeholder syntax and must be converted to {{ }}, while
# 【 】 is ordinary Chinese punctuation (【注意事项】) and is advisory only.
BRACKETED_FAMILIES = {
    "square": re.compile(r"\[\[([^\]]+)\]\]"),
    "angle": re.compile(r"<<([^>]+)>>"),
    "corner": re.compile(r"【([^】]+)】"),
}

_ANY_BRACE = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)
_IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _word_parts(path: Path):
    """Yield (part_name, bytes) for every WordprocessingML part in the package."""
    with ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                yield name, package.read(name)


def _paragraph_texts(payload: bytes):
    """Yield (ordinal, text) per ``w:p``, joining every ``w:t`` inside it.

    Word splits a sentence across runs at every formatting change, so a marker
    can be spread over several ``w:t`` nodes. The engine joins them for the same
    reason; reading run-by-run would miss exactly the markers users hand-type.
    """
    root = ElementTree.fromstring(payload)
    for ordinal, paragraph in enumerate(root.iter(f"{{{W_NS}}}p")):
        yield ordinal, "".join(node.text or "" for node in paragraph.iter(f"{{{W_NS}}}t"))


def extract_markers(path: Path) -> dict:
    """Inventory every fill-in marker in a DOCX.

    Returns ``slots`` (bare, unnamed markers), ``names`` (assigned variables),
    ``unsupported`` (brace spans that are neither), ``bracketed`` and
    ``has_control_syntax``. Raises ``BadZipFile`` when the file is not a package.
    """
    document = Path(path)
    slots: list = []
    names: set = set()
    positions: list = []
    unsupported: list = []
    bracketed: list = []
    has_control = False

    for part, payload in _word_parts(document):
        for ordinal, text in _paragraph_texts(payload):
            if CONTROL.search(text):
                has_control = True
            for match in _ANY_BRACE.finditer(text):
                inner = match.group(1).strip()
                locator = f"{part}:p[{ordinal}]@{match.start()}"
                if inner == "" or _IDENTIFIER.fullmatch(inner):
                    positions.append({"position_id": locator, "part": part, "paragraph": ordinal,
                                      "offset": match.start(), "marker": match.group(0),
                                      "field_id": inner or None,
                                      "context": text[max(0, match.start()-80):match.end()+80]})
                if inner == "":
                    slots.append({"part": part, "paragraph": ordinal,
                                  "offset": match.start(), "locator": locator})
                elif _IDENTIFIER.fullmatch(inner):
                    names.add(inner)
                else:
                    # A dotted or otherwise complex name renders under docxtpl but
                    # is invisible to every Skill-side walker. Surface it instead
                    # of letting it pass unremarked.
                    unsupported.append({"expression": match.group(0),
                                        "locator": locator})
            for family, pattern in BRACKETED_FAMILIES.items():
                for match in pattern.finditer(text):
                    bracketed.append({"family": family, "text": match.group(1),
                                      "locator": f"{part}:p[{ordinal}]@{match.start()}"})

    return {
        "source_sha256": sha256_file(document),
        "slots": slots,
        "names": sorted(names),
        "positions": positions,
        "unsupported": unsupported,
        "bracketed": bracketed,
        "has_control_syntax": has_control,
    }


def template_field_names(path: Path) -> frozenset:
    """The named markers only — what docxtpl will actually substitute."""
    return frozenset(extract_markers(path)["names"])


def has_fill_markers(report: dict) -> bool:
    return bool(report["slots"]) or bool(report["names"])


def blocking_bracketed(report: dict) -> list:
    """Bracketed markers that must be converted to ``{{ }}`` or removed.

    Corner brackets are excluded on purpose: they are ordinary punctuation.
    """
    return [item for item in report["bracketed"] if item["family"] != "corner"]


def template_error(report: dict) -> str | None:
    """Return the prose reason a template is unusable, or None when it is fine."""
    if not has_fill_markers(report):
        return ("DOCX 不是空白模板：未发现任何 {{}} 填充标记。"
                "本技能只接受带 {{}} 标记的空白模板；已完成样例报告不再作为输入。"
                "请在模板中用 {{}} 标出每个待填位置。")
    blocking = blocking_bracketed(report)
    if blocking:
        families = sorted({item["family"] for item in blocking})
        return ("模板中存在非 {{}} 的占位符标记（" + ", ".join(families) + "）。"
                "请将其改为 {{}} 或删除，以保证待填位置只有一种标记方式。")
    if report["has_control_syntax"]:
        return "模板中发现了 {% %} 或 {# #} 控制语法；报告模板只支持简单的 {{}} 填充标记。"
    return None


def summarize(report: dict) -> dict:
    """The compact, human-facing answer to "which parts need filling?"."""
    return {
        "slots": len(report["slots"]),
        "markers": report["names"],
        "slotLocations": [item["locator"] for item in report["slots"]],
        "positions": report["positions"],
        "unsupported": report["unsupported"],
        "bracketed": report["bracketed"],
    }


def main() -> int:
    import argparse
    parser = argparse.ArgumentParser(description="Inventory the {{}} fill-in markers of a DOCX template.")
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    try:
        report = extract_markers(args.docx)
    except (OSError, BadZipFile) as error:
        print(json.dumps({"status": "error", "error": str(error)}, ensure_ascii=False))
        return 1
    error = template_error(report)
    print(json.dumps({"status": "error" if error else "ok", "error": error,
                      "sourceSha256": report["source_sha256"], **summarize(report)},
                     ensure_ascii=False, indent=2))
    return 2 if error else 0


if __name__ == "__main__":
    sys.exit(main())
