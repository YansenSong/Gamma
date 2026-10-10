from __future__ import annotations
import argparse, hashlib, json, re, sys
from pathlib import Path

HEX=re.compile(r"^[0-9a-f]{64}$")

class SpecError(ValueError): pass

# The DOCX fill-in marker extractor, mirrored from scripts/report_markers.py.
# This file is also shipped as assets/windows-app-template/src/project_validation.py,
# and the generated application must run once the Skill tree is gone -- it cannot
# import the Skill's scripts. Keep the two copies in step; `diff` them to check.
_W_NS="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MARKER=re.compile(r"\{\{(.*?)\}\}",re.DOTALL)
_BRACKETED=(re.compile(r"\[\[([^\]]+)\]\]"),re.compile(r"<<([^>]+)>>"))
_CONTROL=re.compile(r"\{%|\{#")
_IDENTIFIER=re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

def docx_markers(path: Path) -> dict:
    from xml.etree import ElementTree
    from zipfile import ZipFile
    names,slots,unsupported,bracketed,control=set(),0,[],[],False
    with ZipFile(path) as package:
        for part in package.namelist():
            if not part.startswith("word/") or not part.endswith(".xml"):continue
            root=ElementTree.fromstring(package.read(part))
            for paragraph in root.iter(f"{{{_W_NS}}}p"):
                text="".join(node.text or "" for node in paragraph.iter(f"{{{_W_NS}}}t"))
                if _CONTROL.search(text):control=True
                for match in _MARKER.finditer(text):
                    inner=match.group(1).strip()
                    if not inner:slots+=1
                    elif _IDENTIFIER.fullmatch(inner):names.add(inner)
                    else:unsupported.append(match.group(0))
                for pattern in _BRACKETED:
                    bracketed.extend(item.group(0) for item in pattern.finditer(text))
    return {"names":names,"slots":slots,"unsupported":unsupported,"bracketed":bracketed,"control":control}

def validate_report_template(root: Path, report: dict) -> None:
    """The template's {{}} markers are the authoritative list of what to fill.

    A marker the configuration does not declare would reach docxtpl with no
    value and silently render as an empty string; a declared field with no
    marker is a value the author believes is reported but never appears.
    """
    declared=set(report.get("fields",{}))
    if not declared:raise SpecError("report_output.fields must declare at least one fill-in field")
    name=Path(report.get("template",""))
    if name.is_absolute() or ".." in name.parts:raise SpecError('Report template path escapes project root')
    template=(root/name).resolve()
    try:template.relative_to(root.resolve())
    except ValueError:raise SpecError('Report template path escapes project root')
    if not template.is_file():raise SpecError(f"report template is missing: {report.get('template','')}")
    markers=docx_markers(template)
    if markers["slots"]:raise SpecError("Report template still contains unnamed {{}} slots; name them before release")
    if markers["unsupported"]:raise SpecError("Report template has unsupported marker expressions: "+", ".join(sorted(set(markers["unsupported"]))))
    if markers["bracketed"]:raise SpecError("Report template uses non-{{}} placeholder markers: "+", ".join(sorted(set(markers["bracketed"]))))
    if markers["control"]:raise SpecError("Report template uses {% %} control syntax; only simple {{}} markers are supported")
    undeclared=sorted(markers["names"]-declared);unmarked=sorted(declared-markers["names"])
    if undeclared:raise SpecError("Report template has markers no field declares: "+", ".join(undeclared))
    if unmarked:raise SpecError("Report fields have no marker in the template: "+", ".join(unmarked))

def digest(value: dict) -> str:
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(",",":")).encode("utf-8")).hexdigest()

def approval_content_hash(spec: dict, approval_type: str) -> str:
    scoring=spec.get("scoring",{})
    projections={
        "scoring_rules":{"inputs":scoring.get("inputs",[]),"formula_bindings":scoring.get("formula_bindings",[])},
        "excel_mapping":spec.get("excel_output",{}),
        "report_mapping":spec.get("report_output",{}),
        "calculation_authority":{"mode":scoring.get("mode"),"authority":scoring.get("authority",{})},
    }
    if approval_type not in projections: raise SpecError(f"unknown approval type: {approval_type}")
    return digest(projections[approval_type])

def validate(path: Path) -> dict:
    spec=json.loads(path.read_text(encoding="utf-8"))
    if spec.get("schema_version")!="autotesting-project-v1": raise SpecError("schema_version must be autotesting-project-v1")
    app=spec.get("application",{})
    if any(not isinstance(app.get(k),str) or not app[k].strip() for k in ("id","name","version")): raise SpecError("application id, name, and version are required")
    root=path.parent.parent if path.parent.name=="config" else path.parent
    sources=spec.get("sources")
    if not isinstance(sources,list) or not sources: raise SpecError("sources must be a non-empty list")
    roles=set()
    for source in sources:
        if not isinstance(source,dict) or source.get("role") not in {"charter","workbook","report_template"}: raise SpecError("source role is invalid")
        roles.add(source["role"]); rel=Path(source.get("path",""))
        if rel.is_absolute() or ".." in rel.parts: raise SpecError("source paths must stay inside the project")
        if not HEX.fullmatch(str(source.get("sha256",""))): raise SpecError("source SHA-256 is invalid")
        file=(root/rel).resolve()
        try: file.relative_to(root.resolve())
        except ValueError: raise SpecError(f"source path escapes project root: {source['path']}")
        if not file.is_file(): raise SpecError(f"source file is missing: {source['path']}")
        if hashlib.sha256(file.read_bytes()).hexdigest()!=source["sha256"]: raise SpecError(f"source hash mismatch: {source['path']}")
    if roles!={"charter","workbook","report_template"}: raise SpecError("charter, workbook, and report_template sources are required")
    report=spec.get("report_output",{})
    validate_report_template(root,report)
    report_source=next(x for x in sources if x['role']=='report_template')
    if hashlib.sha256((root/Path(report.get("template",""))).read_bytes()).hexdigest()!=report_source['sha256']:
        raise SpecError('Report template must have the bound report_template source hash')
    inputs=spec.get("scoring",{}).get("inputs",[])
    if not isinstance(inputs,list): raise SpecError("scoring inputs must be a list")
    field_ids=[x.get("id") for x in inputs if isinstance(x,dict)]
    if len(field_ids)!=len(inputs) or len(set(field_ids))!=len(field_ids): raise SpecError("input IDs must be unique")
    fields=set(field_ids)
    business_decisions=spec.get("business_decisions",[])
    if not isinstance(business_decisions,list) or any(not isinstance(x,dict) for x in business_decisions):
        raise SpecError("business_decisions must be a list of objects")
    decision_ids=[x.get("id") for x in business_decisions]
    if len(set(decision_ids))!=len(decision_ids): raise SpecError("business decision IDs must be unique")
    decisions={x.get("id"):x for x in business_decisions}
    scoring=spec.get('scoring',{})
    if scoring.get('mode')!='excel_native':raise SpecError('scoring.mode must be excel_native; Python scoring is removed')
    if scoring.get('rules'):raise SpecError('Python rule expressions are removed; bind Excel formulas instead')
    bindings=scoring.get('formula_bindings',[])
    if not isinstance(bindings,list):raise SpecError('formula_bindings must be a list')
    outputs=spec.get('excel_output',{}).get('output_cells',{})
    if {x.get('output') for x in bindings}!=set(outputs):raise SpecError('Every Excel output requires a formula binding')
    if len(bindings)!=len(outputs):raise SpecError('Formula binding outputs must be unique')
    from openpyxl import load_workbook
    workbook_source=next(x for x in sources if x['role']=='workbook')
    template_name=Path(spec.get('excel_output',{}).get('template',''))
    if template_name.is_absolute() or '..' in template_name.parts:raise SpecError('Excel template path escapes project root')
    workbook_template=(root/template_name).resolve()
    try:workbook_template.relative_to(root.resolve())
    except ValueError:raise SpecError('Excel template path escapes project root')
    if not workbook_template.is_file() or hashlib.sha256(workbook_template.read_bytes()).hexdigest()!=workbook_source['sha256']:
        raise SpecError('Excel template must have the bound workbook source hash')
    if bindings:
        workbook=load_workbook(workbook_template,data_only=False,read_only=True)
        try:
            for binding in bindings:
                output=binding['output'];mapping=outputs[output]
                if binding.get('sheet')!=mapping.get('sheet') or binding.get('cell')!=mapping.get('cell'):raise SpecError('Formula binding mapping differs: '+output)
                try:cell=workbook[binding['sheet']][binding['cell']]
                except (KeyError,ValueError):raise SpecError('Formula binding cell is invalid: '+output)
                if cell.data_type!='f' or cell.value!=binding.get('formula'):raise SpecError('Formula binding differs from workbook: '+output)
                if not binding.get('source_refs'):raise SpecError('Formula binding requires source references: '+output)
                for decision_id,choices in binding.get('decision_bindings',{}).items():
                    decision=decisions.get(decision_id)
                    if not decision:raise SpecError('Formula binding references unknown decision: '+decision_id)
                    if decision.get('status') in {'confirmed','decided'}:
                        formula_hash=digest({'sheet':binding['sheet'],'cell':binding['cell'],'formula':binding['formula']})
                        if choices.get(decision.get('choice'))!=formula_hash:raise SpecError('Selected business decision is not implemented by Excel formula: '+decision_id)
        finally:workbook.close()
    status=spec.get("release_status","engineering_sample")
    if status not in {"engineering_sample","business_validated"}: raise SpecError("release_status is invalid")
    if status=="business_validated":
        if not bindings:raise SpecError("business_validated release requires mapped Excel formulas")
        state=spec.get('validation_state',{})
        if state.get('status')!='approved' or state.get('official_output_allowed') is not True:
            raise SpecError('business_validated release requires approved validation_state')
        pending=[x.get("id") for x in business_decisions if x.get("status") not in {"confirmed","decided","not_applicable"}]
        if pending: raise SpecError("business_validated project still has pending decisions: "+", ".join(pending))
        source_hashes={x["sha256"] for x in sources}
        for decision in business_decisions:
            if any(not isinstance(decision.get(key),str) or not decision[key].strip() for key in ("reviewer","reason","decided_at")):
                raise SpecError(f"business decision {decision.get('id')} lacks reviewer, reason, or decision time")
            if not isinstance(decision.get("source_hashes"),list) or not decision["source_hashes"] or not all(isinstance(value,str) for value in decision["source_hashes"]) or not set(decision["source_hashes"])<=source_hashes:
                raise SpecError(f"business decision {decision.get('id')} is not bound to project source hashes")
        approvals=spec.get("approvals",{})
        required_approvals={"scoring_rules","excel_mapping","report_mapping","calculation_authority"}
        if not isinstance(approvals,dict) or not required_approvals<=set(approvals):
            raise SpecError("business_validated project requires named approvals for rules, Excel mapping, report mapping, and calculation authority")
        for name in required_approvals:
            approval=approvals[name]
            if not isinstance(approval,dict) or any(not isinstance(approval.get(key),str) or not approval[key].strip() for key in ("reviewer","reason","decided_at")):
                raise SpecError(f"{name} approval lacks reviewer, reason, or decision time")
            if not isinstance(approval.get("source_hashes"),list) or not all(isinstance(value,str) for value in approval["source_hashes"]) or not source_hashes<=set(approval["source_hashes"]):
                raise SpecError(f"{name} approval does not cover every project source hash")
            expected_content_hash=approval_content_hash(spec,name)
            if approval.get("content_sha256")!=expected_content_hash:
                raise SpecError(f"{name} approval content hash is stale")
    return {"status":"valid","application":app["id"],"sourceCount":len(sources),"formulaCount":len(bindings),"releaseStatus":status}

def main() -> int:
    parser=argparse.ArgumentParser(description="Validate project schema, hashes, and decision-to-Excel-formula consistency.")
    parser.add_argument("project_spec",type=Path); args=parser.parse_args()
    try: result=validate(args.project_spec)
    except (OSError,json.JSONDecodeError,SpecError) as error:
        print(json.dumps({"status":"error","error":str(error)},ensure_ascii=False)); return 1
    print(json.dumps(result,ensure_ascii=False)); return 0
if __name__=="__main__": raise SystemExit(main())
