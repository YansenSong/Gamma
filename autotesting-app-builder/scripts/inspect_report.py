from __future__ import annotations
import argparse, json
from pathlib import Path
from _bootstrap import DOCUMENT_ROLE, write_json
from report_markers import extract_markers, summarize, template_error
from ivista_reporting.template_inspector import inspect_docx_template

def main() -> int:
    parser=argparse.ArgumentParser(description="Read-only DOCX blank-template structure inventory.")
    parser.add_argument("docx",type=Path); parser.add_argument("--out",required=True,type=Path)
    args=parser.parse_args()
    markers=extract_markers(args.docx); error=template_error(markers)
    if error:
        print(json.dumps({"status":"error","error":error,"sourceSha256":markers["source_sha256"],
                          **summarize(markers)},ensure_ascii=False))
        return 2
    result=inspect_docx_template(args.docx,document_role=DOCUMENT_ROLE); write_json(args.out,result)
    print(json.dumps({"status":"ok","paragraphs":result["paragraph_count"],"tables":result["table_count"],
                      "fieldCandidates":len(result["field_candidates"]),"sourceSha256":result["source_sha256"],
                      "output":str(args.out),**summarize(markers)},ensure_ascii=False))
    return 0
if __name__=="__main__": raise SystemExit(main())
