from __future__ import annotations
import argparse, importlib, json, os, platform, sys
from pathlib import Path

DECLARED = [
    ("pdfplumber", "charter PDF extraction", True),
    ("openpyxl", "workbook inspection and writing", True),
    ("docx", "DOCX template inspection", True),
    ("docxtpl", "report rendering", False),
    ("PySide6", "generated desktop application runtime", False),
    ("PyInstaller", "onedir packaging", False),
]
VENDORED = [
    ("ivista_engine.charter.pdf_extractor", "PdfPlumberExtractor"),
    ("ivista_engine.authoring.excel_analyzer", "analyze_workbook"),
    ("ivista_reporting.template_inspector", "inspect_docx_template"),
]

def probe_vendor_root() -> dict:
    try:
        from _bootstrap import REPO_ROOT
    except ModuleNotFoundError as error:
        return {"resolved": False, "path": None, "error": str(error)}
    return {"resolved": True, "path": str(REPO_ROOT), "error": None}

def probe_vendored() -> list[dict]:
    results = []
    for module_name, symbol in VENDORED:
        entry = {"module": module_name, "symbol": symbol, "importable": False, "error": None}
        try:
            module = importlib.import_module(module_name)
            if not hasattr(module, symbol):
                entry["error"] = f"module imported but has no {symbol!r}"
            else:
                entry["importable"] = True
        except Exception as error:  # noqa: BLE001 - report any import failure verbatim
            entry["error"] = f"{type(error).__name__}: {error}"
        results.append(entry)
    return results

def probe_declared() -> list[dict]:
    results = []
    for module_name, purpose, required in DECLARED:
        entry = {"module": module_name, "purpose": purpose, "required": required,
                 "installed": False, "version": None, "error": None}
        try:
            module = importlib.import_module(module_name)
            entry["installed"] = True
            entry["version"] = str(getattr(module, "__version__", None) or "unknown")
        except Exception as error:  # noqa: BLE001
            entry["error"] = f"{type(error).__name__}: {error}"
        results.append(entry)
    return results

def probe_excel() -> dict:
    """Detect a registered Excel COM server.

    A registry presence check only proves the ProgID is registered; it does not
    start Excel. Treat a positive result as "probably available" and let the
    actual recalculation decide.
    """
    if platform.system() != "Windows":
        return {"platform": platform.system(), "registered": False,
                "detail": "Excel COM automation requires Windows"}
    try:
        import winreg
    except ImportError:
        return {"platform": "Windows", "registered": None, "detail": "winreg unavailable"}
    for hive, path in ((winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"),
                       (winreg.HKEY_CLASSES_ROOT, r"Excel.Application")):
        try:
            with winreg.OpenKey(hive, path):
                return {"platform": "Windows", "registered": True, "detail": f"HKCR\\{path}"}
        except OSError:
            continue
    return {"platform": "Windows", "registered": False,
            "detail": "HKCR\\Excel.Application not found; native recalculation will fail"}

def skill_paths() -> tuple[Path, Path]:
    """Return (skill_root, parent) derived from this file's own location.

    Derived rather than hard-coded so the Skill keeps working when its folder is
    renamed or moved to another host.
    """
    skill_root = Path(__file__).resolve().parents[1]
    return skill_root, skill_root.parent


def build_remediation(vendor_root: dict, vendored: list[dict], declared: list[dict],
                      excel: dict, require_excel: bool) -> dict:
    """Explain how to fix a blocked environment, in the order the fixes apply.

    A blocker is only useful if the reader knows the next command to run, so the
    preflight carries its own remedy rather than deferring to the documents.
    """
    skill_root, parent = skill_paths()
    expected = parent / ".venv" / "Scripts" / "python.exe"
    in_use = Path(sys.executable)
    under_expected = os.path.normcase(str(in_use)) == os.path.normcase(str(expected))

    missing = [e["module"] for e in declared if e["required"] and not e["installed"]]
    broken = [f"{e['module']}.{e['symbol']}" for e in vendored if not e["importable"]]
    excel_blocked = require_excel and excel.get("registered") is not True
    steps: list[str] = []

    if missing or broken:
        steps.append(f"Build the Skill's working environment. Run these from {parent}; "
                     "the install needs network access the first time.")
        steps.append("uv venv .venv --python 3.13")
        steps.append(f"uv pip install --python .venv/Scripts/python.exe -r {skill_root.name}/requirements.txt")
        if not under_expected:
            steps.append(f"You are running under {in_use}, which is not the interpreter this "
                         f"Skill expects ({expected}). Re-run the preflight with the "
                         "environment interpreter before trusting any other result.")
    if not vendor_root["resolved"]:
        steps.append("This file was run from outside the Skill, or vendor/ was not copied. "
                     "Restore vendor/ivista_engine and vendor/ivista_reporting beside "
                     "scripts/, then re-run.")
    if broken and not missing:
        # Dependencies are all present, so a failing import really does mean a
        # damaged or outdated vendored file.
        steps.append("Vendored modules are present but do not import: " + ", ".join(broken) +
                     ". Re-copy the files listed in vendor/VENDOR.md and check their SHA-256.")
    elif broken:
        # The vendored code imports the very packages that are missing, so the
        # failure is downstream of the first step. Do not send the reader off to
        # re-copy files that are intact.
        steps.append("Vendored import failures above are downstream of the missing "
                     "dependencies; installing them fixes both. Only if they persist "
                     "afterwards are the vendored files themselves damaged.")
    if excel_blocked:
        steps.append("Microsoft Excel is unreachable, so native recalculation cannot run and "
                     "no score can be read. Install or repair Excel, or drop --require-excel "
                     "if this session needs no calculation authority.")
    if not steps:
        steps.append("Environment is ready. Invoke wrappers as "
                     f"../.venv/Scripts/python.exe -B scripts/<wrapper>.py from {skill_root}.")

    return {
        "required": bool(not vendor_root["resolved"] or missing or broken or excel_blocked),
        "interpreterInUse": str(in_use),
        "expectedInterpreter": str(expected),
        "runningUnderExpectedInterpreter": under_expected,
        "steps": steps,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Preflight the Skill toolchain and vendored analysis modules.")
    parser.add_argument("--require-excel", action="store_true",
                        help="treat an unreachable Excel calculation authority as a blocker")
    args = parser.parse_args()

    vendor_root = probe_vendor_root()
    if vendor_root["resolved"]:
        sys.path.insert(0, vendor_root["path"])
    vendored = probe_vendored() if vendor_root["resolved"] else []
    declared = probe_declared()
    excel = probe_excel()

    blockers: list[str] = []
    if not vendor_root["resolved"]:
        blockers.append("vendor root not found: " + str(vendor_root["error"]))
    blockers += [f"vendored import failed: {e['module']}.{e['symbol']}" for e in vendored if not e["importable"]]
    blockers += [f"required dependency missing: {e['module']}" for e in declared if e["required"] and not e["installed"]]
    if args.require_excel and excel.get("registered") is not True:
        blockers.append("Excel calculation authority unreachable: " + str(excel.get("detail")))

    optional_missing = [e["module"] for e in declared if not e["required"] and not e["installed"]]
    remediation = build_remediation(vendor_root, vendored, declared, excel, args.require_excel)
    result = {
        "status": "blocked" if blockers else "ok",
        "python": {"version": platform.python_version(), "executable": sys.executable},
        "platform": platform.platform(),
        "vendorRoot": vendor_root,
        "vendoredModules": vendored,
        "dependencies": declared,
        "optionalMissing": optional_missing,
        "excel": excel,
        "blockers": blockers,
        "remediation": remediation,
        "note": "Optional dependencies are required only for report rendering, packaging or the desktop runtime.",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if blockers else 0

if __name__ == "__main__":
    raise SystemExit(main())
