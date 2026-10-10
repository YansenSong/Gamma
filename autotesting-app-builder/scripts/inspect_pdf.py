from __future__ import annotations
import argparse, json
from dataclasses import asdict
from pathlib import Path
from _bootstrap import REPO_ROOT, write_json
from ivista_engine.charter.pdf_extractor import PdfPlumberExtractor

def main() -> int:
    parser=argparse.ArgumentParser(description="Extract PDF charter text, tables, and stable source references.")
    parser.add_argument("pdf",type=Path); parser.add_argument("--out",required=True,type=Path); parser.add_argument("--password")
    args=parser.parse_args()
    document=PdfPlumberExtractor(password=args.password).extract(args.pdf)
    result={"schema_version":"autotesting-charter-inventory-v1","source_file":str(args.pdf.resolve()),"source_sha256":document.sha256,
            "page_count":document.page_count,"metadata":document.metadata,"pages":[asdict(page) for page in document.pages]}
    write_json(args.out,result)
    print(json.dumps({"status":"ok","pages":document.page_count,"sourceSha256":document.sha256,"output":str(args.out)},ensure_ascii=False))
    return 0
if __name__=="__main__": raise SystemExit(main())
