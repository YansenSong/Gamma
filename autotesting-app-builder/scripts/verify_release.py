from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
from validate_project_spec import SpecError, validate as validate_project_spec

def decision_snapshot(spec):
    keys=("id","status","reviewer","decided_at","choice","reason","source_hashes")
    return [{key:item[key] for key in keys if key in item} for item in spec.get("business_decisions",[])]

def main() -> int:
    parser=argparse.ArgumentParser(description="Verify generated project delivery files and manifest hashes.")
    parser.add_argument("project_root",type=Path); parser.add_argument("--require-exe",action="store_true")
    args=parser.parse_args(); root=args.project_root.resolve(); manifest_path=root/"delivery-manifest.json"
    try:
        manifest=json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError,json.JSONDecodeError) as error:
        print(json.dumps({"status":"error","error":f"delivery manifest unavailable: {error}"})); return 1
    errors=[]
    spec_path=root/"config"/"project.json"
    try:
        spec=json.loads(spec_path.read_text(encoding="utf-8"))
        spec_validation=validate_project_spec(spec_path)
    except (OSError,json.JSONDecodeError,SpecError,ValueError) as error:
        spec={}; spec_validation=None; errors.append({"path":"config/project.json","error":f"project specification invalid: {error}"})
    if spec_validation is not None:
        if manifest.get("release_status")!=spec.get("release_status"):
            errors.append({"path":"delivery-manifest.json","error":"release_status differs from project.json"})
        if manifest.get("project_spec_sha256")!=hashlib.sha256(spec_path.read_bytes()).hexdigest():
            errors.append({"path":"delivery-manifest.json","error":"project_spec_sha256 is stale"})
        if manifest.get("sources")!=spec.get("sources"):
            errors.append({"path":"delivery-manifest.json","error":"source inventory differs from project.json"})
        if manifest.get("approvals",{})!=spec.get("approvals",{}):
            errors.append({"path":"delivery-manifest.json","error":"approvals differ from project.json"})
        if manifest.get("business_decisions",[])!=decision_snapshot(spec):
            errors.append({"path":"delivery-manifest.json","error":"business decisions differ from project.json"})
        if manifest.get("validation",{}).get("project_spec",{}).get("result")!=spec_validation:
            errors.append({"path":"delivery-manifest.json","error":"recorded project validation result is stale"})
    required=["README.md","build.ps1","config/project.json","src/main.py","tests"]
    missing=[name for name in required if not (root/name).exists()]
    # An empty tests/ directory satisfies the existence check above while proving
    # nothing, and tests.status is supplied by the caller rather than measured.
    # Require at least one test module so `mkdir tests` cannot stand in for a suite.
    if (root/"tests").is_dir() and not any((root/"tests").glob("test_*.py")):
        missing.append("tests/test_*.py (tests/ declares no test module)")
    artifacts=manifest.get("artifacts",{})
    mismatches=list(errors)
    for name,expected in artifacts.items():
        path=(root/name).resolve()
        try: path.relative_to(root)
        except ValueError: mismatches.append({"path":name,"error":"path escapes project root"}); continue
        if not path.is_file(): mismatches.append({"path":name,"error":"missing"})
        elif hashlib.sha256(path.read_bytes()).hexdigest()!=expected: mismatches.append({"path":name,"error":"sha256 mismatch"})
    app=manifest.get("application",{})
    app_name=app.get("executable_name") or app.get("name")
    exe=root/"dist"/app_name/f"{app_name}.exe" if app_name else None
    validation=manifest.get("validation",{})
    tests=validation.get("tests",{})
    if tests.get("status")!="passed": missing.append("validation.tests.status=passed")
    if validation.get("project_spec",{}).get("status")!="validated": missing.append("validation.project_spec.status=validated")
    build=validation.get("build",{})
    if build.get("status")=="failed": missing.append("validation.build.status is failed")
    if manifest.get("release_status")=="business_validated" and build.get("status")!="passed":
        missing.append("business_validated release requires a passed Windows build")
    if build.get("status")=="passed" and build.get("startup_smoke_status")!="passed":
        missing.append("passed build has no passed startup smoke check")
    executable_info=build.get("executable") or {}
    if build.get("status")=="passed" and (not executable_info.get("exists") or not executable_info.get("sha256")):
        missing.append("passed build has no verified executable")
    executable_path=executable_info.get("path")
    if executable_path:
        relative=Path(executable_path)
        if relative.is_absolute() or ".." in relative.parts:
            mismatches.append({"path":executable_path,"error":"executable path escapes project root"})
        else:
            built_exe=(root/relative).resolve()
            try: built_exe.relative_to(root)
            except ValueError: mismatches.append({"path":executable_path,"error":"executable path escapes project root"})
            else:
                if not built_exe.is_file(): mismatches.append({"path":executable_path,"error":"missing"})
                elif hashlib.sha256(built_exe.read_bytes()).hexdigest()!=executable_info.get("sha256"):
                    mismatches.append({"path":executable_path,"error":"sha256 mismatch"})
                elif build.get("status")=="passed":
                    bundle=built_exe.parent
                    runtime=bundle/"_internal" if (bundle/"_internal").is_dir() else bundle
                    try: project_spec=json.loads((root/"config"/"project.json").read_text(encoding="utf-8"))
                    except (OSError,json.JSONDecodeError): project_spec={}
                    runtime_files=["config/project.json"]
                    if (root/"src/ui_design_system/workbench.qss").is_file():
                        runtime_files.append("src/ui_design_system/workbench.qss")
                    runtime_files.extend(item for item in (
                        project_spec.get("excel_output",{}).get("template"),
                        project_spec.get("report_output",{}).get("template")) if isinstance(item,str))
                    if project_spec.get("scoring",{}).get("mode")=="excel_native":
                        recalc=project_spec.get("scoring",{}).get("authority",{}).get("recalculation_script","scripts/recalculate_excel.ps1")
                        if isinstance(recalc,str): runtime_files.append(recalc)
                    for name in runtime_files:
                        relative=Path(name)
                        if relative.is_absolute() or ".." in relative.parts or not (runtime/relative).is_file():
                            mismatches.append({"path":f"{executable_path} bundle/{name}","error":"required runtime file missing or unsafe"})
                    if not any(bundle.rglob("qwindows.dll")):
                        mismatches.append({"path":executable_path,"error":"Qt Windows platform plugin qwindows.dll is missing"})
    if args.require_exe and (not exe or not exe.is_file()): missing.append("dist/<AppName>/<AppName>.exe")
    errors=missing+[f"{x['path']}: {x['error']}" for x in mismatches]
    status="passed" if not errors else "failed"
    print(json.dumps({"status":status,"releaseStatus":manifest.get("release_status"),"missing":missing,
                      "artifactMismatches":mismatches,"exePresent":bool(exe and exe.is_file())},ensure_ascii=False))
    return 0 if status=="passed" else 1
if __name__=="__main__": raise SystemExit(main())
