from __future__ import annotations
import argparse, hashlib, json, sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

MAX_LISTED = 500

def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def structure(ws) -> dict:
    validations = sorted(
        (str(dv.sqref), dv.type, str(dv.formula1), str(dv.operator), bool(dv.allow_blank))
        for dv in ws.data_validations.dataValidation
    )
    return {
        "state": ws.sheet_state,
        "merged": sorted(str(r) for r in ws.merged_cells.ranges),
        "validations": validations,
        "protection": {"sheet": bool(ws.protection.sheet), "password": str(ws.protection.password)},
        "freezePanes": str(ws.freeze_panes) if ws.freeze_panes else None,
        "autoFilter": str(ws.auto_filter.ref) if ws.auto_filter and ws.auto_filter.ref else None,
        "columnWidths": {k: v.width for k, v in ws.column_dimensions.items() if v.width is not None},
        "rowHeights": {k: v.height for k, v in ws.row_dimensions.items() if v.height is not None},
    }

def cell_value(ws, coordinate: str):
    cell = ws[coordinate]
    return cell.value if cell.data_type == "f" else cell.value, cell.data_type, cell.number_format

def compare_cells(original_ws, authored_ws, include_caches: bool) -> list[dict]:
    """Compare the populated area of both sheets in formula view."""
    diffs = []
    rows = max(original_ws.max_row, authored_ws.max_row, 0)
    cols = max(original_ws.max_column, authored_ws.max_column, 0)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            coordinate = f"{get_column_letter(c)}{r}"
            before = original_ws[coordinate]
            after = authored_ws[coordinate]
            before_value, before_type, before_format = before.value, before.data_type, before.number_format
            after_value, after_type, after_format = after.value, after.data_type, after.number_format
            if before_value == after_value and before_type == after_type and before_format == after_format:
                continue
            entry = {
                "cell": coordinate,
                "before": {"value": before_value, "type": before_type},
                "after": {"value": after_value, "type": after_type},
            }
            if before_format != after_format:
                entry["numberFormat"] = {"before": before_format, "after": after_format}
            if before_type == "f" or after_type == "f":
                entry["kind"] = "formula"
            elif before_value is None or after_value is None:
                entry["kind"] = "presence"
            else:
                entry["kind"] = "value"
            diffs.append(entry)
    return diffs

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Prove an authorized workbook edit touched only the intended cells.")
    parser.add_argument("original", type=Path, help="pristine supplied workbook")
    parser.add_argument("authored", type=Path, help="workbook after formula authoring")
    parser.add_argument("--expect-changed", default="",
                        help="comma-separated Sheet1!N4,Sheet1!M22 list of cells authorized to change")
    parser.add_argument("--include-caches", action="store_true",
                        help="also compare cached formula results (they change after recalculation)")
    args = parser.parse_args()

    expected = {item.strip() for item in args.expect_changed.split(",") if item.strip()}

    original_formulas = load_workbook(args.original, data_only=False)
    authored_formulas = load_workbook(args.authored, data_only=False)
    try:
        structural: list[dict] = []
        cell_changes: list[dict] = []

        original_names, authored_names = original_formulas.sheetnames, authored_formulas.sheetnames
        if original_names != authored_names:
            structural.append({"scope": "workbook", "item": "sheetnames",
                               "before": original_names, "after": authored_names})

        for name in original_names:
            if name not in authored_names:
                continue
            before_ws, after_ws = original_formulas[name], authored_formulas[name]
            before_structure, after_structure = structure(before_ws), structure(after_ws)
            for key in before_structure:
                if before_structure[key] != after_structure[key]:
                    structural.append({"scope": name, "item": key,
                                       "before": before_structure[key], "after": after_structure[key]})
            for diff in compare_cells(before_ws, after_ws, args.include_caches):
                diff["sheet"] = name
                diff["qualified"] = f"{name}!{diff['cell']}"
                diff["authorized"] = diff["qualified"] in expected
                cell_changes.append(diff)

        if args.include_caches:
            original_values = load_workbook(args.original, data_only=True)
            authored_values = load_workbook(args.authored, data_only=True)
            try:
                for name in original_names:
                    if name not in authored_values.sheetnames:
                        continue
                    before_ws, after_ws = original_values[name], authored_values[name]
                    for r in range(1, max(before_ws.max_row, after_ws.max_row, 0) + 1):
                        for c in range(1, max(before_ws.max_column, after_ws.max_column, 0) + 1):
                            coordinate = f"{get_column_letter(c)}{r}"
                            before_cell, after_cell = before_ws[coordinate], after_ws[coordinate]
                            if before_cell.data_type != "f" and after_cell.data_type != "f":
                                continue
                            if before_cell.value == after_cell.value:
                                continue
                            qualified = f"{name}!{coordinate}"
                            cell_changes.append({
                                "sheet": name, "cell": coordinate, "qualified": qualified,
                                "kind": "cachedValue", "authorized": qualified in expected,
                                "before": {"value": before_cell.value}, "after": {"value": after_cell.value},
                            })
            finally:
                original_values.close()
                authored_values.close()

        changed_cells = {d["qualified"] for d in cell_changes}
        unexpected = [d for d in cell_changes if not d["authorized"]]
        missing = sorted(expected - changed_cells)

        def listed(items):
            return items[:MAX_LISTED]

        result = {
            "status": "failed" if (unexpected or structural or missing) else "passed",
            "original": {"path": str(args.original.resolve()), "sha256": sha256(args.original)},
            "authored": {"path": str(args.authored.resolve()), "sha256": sha256(args.authored)},
            "summary": {
                "expectedCells": len(expected),
                "changedCells": len(changed_cells),
                "authorizedChanges": len(cell_changes) - len(unexpected),
                "unexpectedChanges": len(unexpected),
                "structuralChanges": len(structural),
                "authorizedCellsUnchanged": len(missing),
            },
            "unexpectedChanges": listed(unexpected),
            "structuralChanges": listed(structural),
            "authorizedCellsUnchanged": missing[:MAX_LISTED],
            "truncated": (len(unexpected) > MAX_LISTED or len(structural) > MAX_LISTED
                          or len(missing) > MAX_LISTED),
            "note": "A change outside --expect-changed is a failure even when the scores look correct.",
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "passed" else 1
    finally:
        original_formulas.close()
        authored_formulas.close()

if __name__ == "__main__":
    raise SystemExit(main())
