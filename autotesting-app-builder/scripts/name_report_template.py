"""Assign field names to the bare ``{{}}`` slots of a blank DOCX template.

The user marks fill-in positions with bare ``{{}}``; the mapping step decides
what each position means and writes the name in, turning it into the
``{{ score }}`` that docxtpl renders. The supplied template is never modified:
a named copy is written, so the original keeps its SHA-256 and the named copy
becomes the project's ``role: report_template`` source.
"""
from __future__ import annotations
import argparse, hashlib, json, re, shutil
from pathlib import Path
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SLOT = re.compile(r"\{\{\s*\}\}")
IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def paragraph_text(paragraph: Paragraph) -> str:
    return "".join(run.text or "" for run in paragraph.runs)

def _replace_nth(paragraph: Paragraph, find: str, replacement: str, occurrence: int = 1) -> bool:
    """Replace the `occurrence`-th match, preserving surrounding runs.

    python-docx splits a sentence across runs at every formatting change, so a
    naive run-by-run replace misses any target that spans a boundary. Rewrite
    the first overlapping run and clear the covered tail of the last one.
    """
    runs = paragraph.runs
    texts = [run.text or "" for run in runs]
    full = "".join(texts)
    index, search_from = -1, 0
    for _ in range(occurrence):
        index = full.find(find, search_from)
        if index < 0:
            return False
        search_from = index + len(find)
    end = index + len(find)
    position, first, last = 0, 0, 0
    first_offset = last_offset = 0
    located = False
    for i, text in enumerate(texts):
        start, stop = position, position + len(text)
        if not located and stop > index:
            first, first_offset, located = i, index - start, True
        if start < end:
            last, last_offset = i, end - start
        position = stop
    if not located:
        return False
    prefix = texts[first][:first_offset]
    suffix = texts[last][last_offset:]
    if first == last:
        runs[first].text = prefix + replacement + suffix
    else:
        runs[first].text = prefix + replacement
        runs[last].text = suffix
        for i in range(first + 1, last):
            runs[i].text = ""
    return True

# lxml builds Python proxy objects on demand and frees them immediately, so
# id(element) is not stable across traversals and freed ids get reused. Holding
# a reference keeps every proxy alive for the lifetime of the run, which makes
# the `seen` set below reliable. Without this, merged cells and repeated walks
# collide and most paragraphs are wrongly skipped.
_PROXY_KEEPALIVE: list = []


def iter_paragraphs(container, path: str, seen: set):
    """Yield (paragraph, location) for body paragraphs and tables, recursively."""
    for index, paragraph in enumerate(getattr(container, "paragraphs", []) or []):
        element = paragraph._p
        key = id(element)
        if key not in seen:
            seen.add(key)
            _PROXY_KEEPALIVE.append(element)
            yield paragraph, f"{path}/p[{index}]"
    for index, table in enumerate(getattr(container, "tables", []) or []):
        if not isinstance(table, Table):
            continue
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                yield from iter_paragraphs(cell, f"{path}/table[{index}]/r{row_index}c{cell_index}", seen)

def iter_document_paragraphs(document: Document):
    seen: set = set()
    yield from iter_paragraphs(document, "body", seen)
    for section_index, section in enumerate(document.sections):
        for name, part in (("header", section.header), ("footer", section.footer),
                           ("first_header", section.first_page_header),
                           ("first_footer", section.first_page_footer),
                           ("even_header", section.even_page_header),
                           ("even_footer", section.even_page_footer)):
            yield from iter_paragraphs(part, f"section[{section_index}]/{name}", seen)

def scan_slots(document: Document) -> list:
    """Every bare slot in document order.

    Slots are textually identical, so a slot is identified by its position:
    `index` is its ordinal across the whole document, `nth` its ordinal among
    the identical matches inside one paragraph.
    """
    found: list = []
    for paragraph, location in iter_document_paragraphs(document):
        nth = {}
        for match in SLOT.finditer(paragraph_text(paragraph)):
            raw = match.group(0)
            nth[raw] = nth.get(raw, 0) + 1
            found.append({"paragraph": paragraph, "location": location, "find": raw, "nth": nth[raw]})
    return found

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Name the bare {{}} slots of a blank DOCX template, writing a new file.")
    parser.add_argument("plan", type=Path, help="JSON plan: {template, output, slots:[{occurrence, field_id}]}")
    parser.add_argument("--dry-run", action="store_true", help="report the naming without writing the template")
    args = parser.parse_args()

    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    source = Path(plan["template"])
    output = Path(plan["output"])
    if not source.is_file():
        print(json.dumps({"status": "error", "error": f"template not found: {source}"}, ensure_ascii=False))
        return 2
    if source.resolve() == output.resolve():
        print(json.dumps({"status": "error",
                          "error": "output must differ from the template; the original keeps its hash"},
                         ensure_ascii=False))
        return 2
    entries = plan.get("slots", [])
    for entry in entries:
        if not IDENTIFIER.fullmatch(str(entry.get("field_id", ""))):
            print(json.dumps({"status": "error",
                              "error": f"invalid field_id: {entry.get('field_id')!r}; "
                                       "use a plain identifier such as score or vehicle_id"},
                             ensure_ascii=False))
            return 2

    document = Document(str(source))
    available = scan_slots(document)
    total = len(available)
    for entry in entries:
        if not 0 <= int(entry["occurrence"]) < total:
            print(json.dumps({"status": "error",
                              "error": f"occurrence {entry['occurrence']} is out of range; "
                                       f"the template has {total} slots"},
                             ensure_ascii=False))
            return 2

    # Name the later slots first: replacing slot 5 leaves slots 0-4 where they
    # were, so every occurrence index stays valid without recomputing offsets.
    results = []
    for entry in sorted(entries, key=lambda item: int(item["occurrence"]), reverse=True):
        slot = scan_slots(document)[int(entry["occurrence"])]
        placeholder = f"{{{{ {entry['field_id']} }}}}"
        if not _replace_nth(slot["paragraph"], slot["find"], placeholder, slot["nth"]):
            print(json.dumps({"status": "error",
                              "error": f"could not write {placeholder} at occurrence {entry['occurrence']}"},
                             ensure_ascii=False))
            return 2
        results.append({"occurrence": int(entry["occurrence"]), "field_id": entry["field_id"],
                        "placeholder": placeholder, "location": slot["location"]})

    remaining = scan_slots(document)
    unnamed = [{"occurrence": index, "location": slot["location"]} for index, slot in enumerate(remaining)]

    if not args.dry_run:
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists() and sha256(output) != sha256(source):
            backup = output.with_suffix(output.suffix + ".previous")
            shutil.copy2(output, backup)
        document.save(str(output))

    result = {
        "status": "ok" if not unnamed else "incomplete",
        "dryRun": args.dry_run,
        "template": {"path": str(source.resolve()), "sha256": sha256(source), "slots": total},
        "output": None if args.dry_run else {"path": str(output.resolve()), "sha256": sha256(output)},
        "replacements": sorted(results, key=lambda item: item["occurrence"]),
        "unnamedSlots": unnamed,
        "note": ("An unnamed slot is a blocker: it reaches docxtpl as literal text and the position "
                 "stays visibly unfilled. Name every {{}} slot before the template can be released."),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not unnamed else 1

if __name__ == "__main__":
    raise SystemExit(main())
