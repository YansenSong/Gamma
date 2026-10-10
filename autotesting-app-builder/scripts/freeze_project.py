"""Freeze reviewed mappings into an engineering project configuration.

This only binds artifacts and validates formulas. It never grants business
approval or claims native Excel / Windows acceptance.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from openpyxl import load_workbook
from _bootstrap import write_json
from validate_project_spec import validate as validate_project
from report_markers import sha256_file


def freeze(root:Path):
    root=Path(root).resolve()
    spec_path=root/'config/project.json'
    current=json.loads(spec_path.read_text(encoding='utf-8'))
    if current.get('sources'):
        raise ValueError('project.json is already configured; existing configuration is preserved')
    source=json.loads((root/'analysis/source-manifest.json').read_text(encoding='utf-8'))
    rules=json.loads((root/'analysis/scoring_rules.json').read_text(encoding='utf-8'))
    excel=json.loads((root/'review/excel-map-plan.json').read_text(encoding='utf-8'))
    report=json.loads((root/'analysis/report-map.json').read_text(encoding='utf-8'))
    plan=json.loads((root/'review/project-plan.json').read_text(encoding='utf-8'))
    originals={item['role']:item for item in source['sources']}
    if set(originals)!={'charter','workbook','report_template'}:
        raise ValueError('three original sources are required')
    if rules.get('source_sha256')!=originals['charter']['sha256']:
        raise ValueError('scoring rules refer to a different PDF')
    if excel.get('workbook_sha256')!=originals['workbook']['sha256']:
        raise ValueError('Excel mapping refers to a different original workbook')
    if report.get('template_sha256')!=originals['report_template']['sha256']:
        raise ValueError('report mapping refers to a different original DOCX')
    if not report.get('positions') or any(item.get('review_status')!='confirmed' for item in report['positions']):
        raise ValueError('every report position needs explicit confirmed review before project freeze')
    application=plan.get('application',{})
    if application.get('id') in (None,'','replace-me'):
        raise ValueError('set a project-specific application.id')
    inputs=excel.get('input_cells',{});outputs=excel.get('output_cells',{})
    if not isinstance(inputs,dict) or not isinstance(outputs,dict) or not outputs:
        raise ValueError('confirmed Excel input/output mapping is required')
    for kind,mappings in (('input',inputs),('output',outputs)):
        for name,item in mappings.items():
            if item.get('review_status')!='confirmed' or not item.get('source_refs'):
                raise ValueError(f'{kind} mapping {name} lacks confirmed source review')
    bindings=plan.get('formula_bindings',[])
    if not isinstance(bindings,list) or {x.get('output') for x in bindings}!=set(outputs) or len(bindings)!=len(outputs):
        raise ValueError('every output needs one formula binding')
    workbook_path=root/'templates/score.xlsx'
    workbook=load_workbook(workbook_path,data_only=False,read_only=True)
    try:
        for binding in bindings:
            name=binding['output'];mapping=outputs[name]
            if (binding.get('sheet'),binding.get('cell'))!=(mapping.get('sheet'),mapping.get('cell')):
                raise ValueError('formula binding location differs: '+name)
            if not binding.get('source_refs') or workbook[binding['sheet']][binding['cell']].value!=binding.get('formula'):
                raise ValueError('formula binding lacks PDF source or workbook formula: '+name)
    finally:workbook.close()
    fields=report.get('fields',{})
    for name,field in fields.items():
        kind,key=field.get('kind'),field.get('id')
        if (kind=='score' and key not in outputs) or (kind=='input' and key not in inputs) or (kind=='system' and key not in {'audit_id','excel_version'}) or kind not in {'score','input','system'}:
            raise ValueError('report field has no valid workbook/system source: '+name)
    named=root/'templates/report.docx'
    if sha256_file(named)!=report.get('named_template_sha256'):
        raise ValueError('named DOCX changed after mapping')
    spec={'schema_version':'autotesting-project-v1','application':application,'release_status':'engineering_sample',
          'validation_state':{'status':'pending','official_output_allowed':False},
          'sources':[{'role':'charter','path':'sources/charter.pdf','sha256':originals['charter']['sha256']},
                     {'role':'workbook','path':'templates/score.xlsx','sha256':sha256_file(workbook_path)},
                     {'role':'report_template','path':'templates/report.docx','sha256':sha256_file(named)}],
          'business_decisions':plan.get('business_decisions',[]),
          'scoring':{'mode':'excel_native','inputs':plan.get('scoring_inputs',[]),'formula_bindings':bindings,
                     'authority':{'type':'Microsoft Excel','require_full_recalculation':True,
                                  'recalculation_script':'scripts/recalculate_excel.ps1'}},
          'excel_output':{'template':'templates/score.xlsx','input_cells':inputs,'output_cells':outputs},
          'report_output':{'template':'templates/report.docx','fields':fields},
          'report_safety':{'provenance':'工程核对稿：规则和真实业务数据尚未完成正式审批。'}}
    # Validation runs against a temporary config, leaving the placeholder intact on error.
    staged=root/'config/project.pending.json';write_json(staged,spec)
    try:validate_project(staged)
    except Exception:
        staged.unlink(missing_ok=True);raise
    staged.replace(spec_path)
    return {'status':'engineering_sample','project':str(root),'output_count':len(outputs),'report_field_count':len(fields),
            'business_review_blocked':rules.get('release_blocked',True)}


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('project_root',type=Path)
    args=parser.parse_args()
    try:result=freeze(args.project_root)
    except (OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False));return 1
    print(json.dumps(result,ensure_ascii=False));return 0


if __name__=='__main__':raise SystemExit(main())
