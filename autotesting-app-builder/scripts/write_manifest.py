from __future__ import annotations
import argparse, hashlib, json
from datetime import datetime, timezone
from pathlib import Path
from _bootstrap import write_json
from validate_project_spec import validate as validate_project

EXCLUDED_PARTS={".git",".venv","__pycache__","build","dist"}

def create_manifest(root: Path, *, tests_status: str, tests_command: str,
                    build_status: str, smoke_status: str, executable: str | None,
                    limitations: list[str]) -> dict:
    root=root.resolve()
    spec_path=root/"config"/"project.json"
    spec=json.loads(spec_path.read_text(encoding="utf-8"))
    spec_result=validate_project(spec_path)
    if build_status=="passed" and not executable:
        raise ValueError("build is marked passed but no executable path was supplied")
    artifacts={}
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.name=="delivery-manifest.json":
            continue
        relative=path.relative_to(root)
        if any(part in EXCLUDED_PARTS for part in relative.parts):
            continue
        artifacts[relative.as_posix()]=hashlib.sha256(path.read_bytes()).hexdigest()
    executable_info=None
    if executable:
        exe=(root/executable).resolve()
        try: exe.relative_to(root)
        except ValueError: raise ValueError("executable path escapes project root")
        if build_status=="passed" and not exe.is_file(): raise ValueError("build is marked passed but executable is missing")
        executable_info={"path":exe.relative_to(root).as_posix(),
                         "exists":exe.is_file(),
                         "sha256":hashlib.sha256(exe.read_bytes()).hexdigest() if exe.is_file() else None}
    app=spec.get("application",{})
    decisions=[]
    for item in spec.get("business_decisions",[]):
        decisions.append({key:item.get(key) for key in ("id","status","reviewer","decided_at","choice","reason","source_hashes") if key in item})
    return {"schema_version":"autotesting-delivery-manifest-v1",
      "generated_at_utc":datetime.now(timezone.utc).isoformat(timespec="seconds"),
      "project_id":app.get("id"),"release_status":spec.get("release_status"),
      "project_spec_sha256":hashlib.sha256(spec_path.read_bytes()).hexdigest(),
      "application":{"id":app.get("id"),"name":app.get("name"),"version":app.get("version"),
                      "executable_name":Path(executable).stem if executable else None},
      "sources":spec.get("sources",[]),"business_decisions":decisions,"approvals":spec.get("approvals",{}),
      "validation":{"tests":{"status":tests_status,"command":tests_command},
                    "project_spec":{"status":"validated","result":spec_result},
                    "build":{"status":build_status,"startup_smoke_status":smoke_status,"executable":executable_info}},
      "known_limitations":limitations,"artifacts":artifacts}

def main() -> int:
    parser=argparse.ArgumentParser(description="Write a source-hash and validation delivery manifest.")
    parser.add_argument("project_root",type=Path)
    parser.add_argument("--tests-status",choices=("passed","failed","not_run"),required=True)
    parser.add_argument("--tests-command",required=True)
    parser.add_argument("--build-status",choices=("passed","failed","not_available","not_run"),required=True)
    parser.add_argument("--smoke-status",choices=("passed","failed","not_run"),default="not_run")
    parser.add_argument("--executable",help="Executable path relative to the project root")
    parser.add_argument("--known-limit",action="append",default=[])
    args=parser.parse_args()
    try:
        root=args.project_root.resolve()
        manifest=create_manifest(root,tests_status=args.tests_status,tests_command=args.tests_command,
            build_status=args.build_status,smoke_status=args.smoke_status,
            executable=args.executable,limitations=args.known_limit)
        write_json(root/"delivery-manifest.json",manifest)
    except (OSError,ValueError,json.JSONDecodeError) as error:
        print(json.dumps({"status":"error","error":str(error)},ensure_ascii=False)); return 1
    print(json.dumps({"status":"written","artifactCount":len(manifest["artifacts"]),
        "releaseStatus":manifest["release_status"],"buildStatus":args.build_status},ensure_ascii=False))
    return 0
if __name__=="__main__": raise SystemExit(main())
