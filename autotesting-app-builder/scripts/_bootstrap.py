from __future__ import annotations
import sys
from pathlib import Path


def _find_vendor_root() -> Path:
    """Locate the vendored ivista_engine/ivista_reporting subset.

    Upstream hard-coded ``parents[4]`` because the Skill lived inside the
    AutoTesting checkout. This standalone copy keeps the needed modules in a
    ``vendor/`` directory beside the Skill instead. Walk up from this file so
    the scripts keep working if the Skill is moved.
    """
    for parent in Path(__file__).resolve().parents:
        candidate = parent / "vendor"
        if (candidate / "ivista_engine").is_dir():
            return candidate
    raise ModuleNotFoundError(
        "Could not locate the 'vendor' directory holding ivista_engine/"
        "ivista_reporting. Expected <skill-root>/vendor beside the Skill."
    )


REPO_ROOT = _find_vendor_root()
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# The only document role this Skill accepts. The vendored inspector still
# declares a blank_template/reference_sample pair, but the reference-sample path
# is gone: a completed report is no longer an accepted input. The engine treats
# the value purely as a label -- it validates the string and echoes it back, and
# branches on it nowhere -- so this constant exists to name the surviving value
# in one place rather than to steer engine behaviour.
DOCUMENT_ROLE = "blank_template"

def write_json(path: Path, value: dict) -> None:
    import json
    path=Path(path); path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(value,ensure_ascii=False,sort_keys=True,indent=2)+"\n",encoding="utf-8",newline="\n")
