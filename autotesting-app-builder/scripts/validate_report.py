from __future__ import annotations
import argparse, json
from pathlib import Path
from _bootstrap import DOCUMENT_ROLE, REPO_ROOT
from ivista_reporting.template_inspector import inspect_docx_template

def main() -> int:
    parser=argparse.ArgumentParser(description="Validate DOCX report output and unresolved template markers.")
    parser.add_argument("report",type=Path); parser.add_argument("--rendered",action="store_true")
    args=parser.parse_args()
    inventory=inspect_docx_template(args.report,document_role=DOCUMENT_ROLE)
    unresolved=[x for x in inventory["field_candidates"] if x["kind"]=="jinja_variable"]
    # Without --rendered this checks a template, where unresolved markers are the
    # point. With --rendered it checks a finished report, where any survivor means
    # the render silently left a position unfilled.
    status="passed" if not (args.rendered and unresolved) else "failed"
    print(json.dumps({"status":status,"sourceSha256":inventory["source_sha256"],
                      "unresolvedVariables":unresolved},ensure_ascii=False))
    return 0 if status=="passed" else 1
if __name__=="__main__": raise SystemExit(main())
