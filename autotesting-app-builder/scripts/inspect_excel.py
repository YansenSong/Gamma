from __future__ import annotations
import argparse, json
from pathlib import Path
from _bootstrap import write_json
from ivista_engine.authoring.excel_analyzer import analyze_workbook

def main() -> int:
    parser=argparse.ArgumentParser(description="Read-only inventory of an Excel scoring template.")
    parser.add_argument("workbook",type=Path); parser.add_argument("--out",required=True,type=Path)
    args=parser.parse_args(); result=analyze_workbook(args.workbook); write_json(args.out,result)
    print(json.dumps({"status":"ok","sheets":result["sheet_count"],"cells":result["nonempty_or_styled_cell_count"],
                      "sourceSha256":result["source_sha256"],"unsupportedFeatures":result["unsupported_feature_flags"],
                      "output":str(args.out)},ensure_ascii=False))
    return 0
if __name__=="__main__": raise SystemExit(main())
