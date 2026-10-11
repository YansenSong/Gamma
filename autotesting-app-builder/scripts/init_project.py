"""Start a project from any charter PDF, scoring XLSX and marked DOCX.

This stage inventories sources and creates review plans. It never guesses rules,
cell mappings or report fields, and it does not produce a runnable release.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import tempfile

from _bootstrap import DOCUMENT_ROLE, write_json
from ivista_engine.authoring.excel_analyzer import analyze_workbook
import pdfplumber
from ivista_reporting.template_inspector import inspect_docx_template
from report_markers import extract_markers, template_error, sha256_file
from export_report_positions import export as export_report_positions

SKILL_ROOT=Path(__file__).resolve().parents[1]
TEMPLATE=SKILL_ROOT/'assets/windows-app-template'


def initialize(pdf:Path,workbook:Path,report:Path,destination:Path,*,project_name:str):
    paths={'charter':Path(pdf).resolve(),'workbook':Path(workbook).resolve(),'report_template':Path(report).resolve()}
    for role,path in paths.items():
        expected={'charter':'.pdf','workbook':'.xlsx','report_template':'.docx'}[role]
        if not path.is_file() or path.suffix.lower()!=expected:
            raise ValueError(f'{role} must be a readable {expected} file')
    destination=Path(destination).resolve()
    if destination.exists():raise ValueError('project destination already exists; existing work is preserved')
    if any(destination==path or destination in path.parents for path in paths.values()):
        raise ValueError('project destination must not contain an input document')
    if not project_name.strip():raise ValueError('project name is required')
    # Analyze originals before writing anything; a malformed input leaves no project.
    pdf_hash=hashlib.sha256(paths['charter'].read_bytes()).hexdigest()
    with pdfplumber.open(paths['charter']) as source_pdf:
        page_count=len(source_pdf.pages)
    pdf_inventory={'schema_version':'gamma-charter-source-v1','source_file':str(paths['charter']),
                   'source_sha256':pdf_hash,'page_count':page_count,
                   'pages':[{'page_index':index} for index in range(1,page_count+1)]}
    excel_inventory=analyze_workbook(paths['workbook'])
    markers=extract_markers(paths['report_template'])
    issue=template_error(markers)
    if issue:raise ValueError(issue)
    report_inventory=inspect_docx_template(paths['report_template'],document_role=DOCUMENT_ROLE)
    destination.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.gamma-init-',dir=destination.parent) as staging:
        root=Path(staging)/'project';shutil.copytree(TEMPLATE,root,ignore=shutil.ignore_patterns('__pycache__','*.pyc','dist','build','.pytest_cache'))
        (root/'sources').mkdir();(root/'analysis').mkdir(exist_ok=True);(root/'review').mkdir()
        source_manifest=[]
        for role,path in paths.items():
            relative=Path('sources')/(role+path.suffix.lower())
            shutil.copy2(path,root/relative)
            expected_hash={'charter':pdf_hash,'workbook':excel_inventory['source_sha256'],
                           'report_template':markers['source_sha256']}[role]
            if sha256_file(root/relative)!=expected_hash or sha256_file(path)!=expected_hash:
                raise ValueError(f'{role} changed during project initialization')
            source_manifest.append({'role':role,'original_path':str(path),'project_path':str(relative),
                                    'sha256':expected_hash,'size_bytes':path.stat().st_size})
        (root/'templates').mkdir(exist_ok=True)
        shutil.copy2(root/'sources/workbook.xlsx',root/'templates/score.xlsx')
        write_json(root/'analysis/pdf.json',pdf_inventory)
        write_json(root/'analysis/excel.json',excel_inventory)
        write_json(root/'analysis/report.json',report_inventory)
        write_json(root/'analysis/report-markers.json',markers)
        report_positions=export_report_positions(root/'sources/report_template.docx',root/'analysis/report-positions.json',root/'review/report-positions.csv')
        write_json(root/'analysis/source-manifest.json',{'schema_version':'gamma-source-manifest-v1','project_name':project_name,
                                                           'sources':source_manifest,'status':'review_required'})
        write_json(root/'review/rule-plan.json',{'source_sha256':pdf_hash,'reviewed_pages':[],
                                                 'scope':{'type':'workbook','workbook_sha256':excel_inventory['source_sha256'],
                                                          'score_targets':[]},
                                                 'rules':[],'status':'pending'})
        write_json(root/'review/excel-map-plan.json',{'workbook_sha256':excel_inventory['source_sha256'],
                                                      'input_cells':{},'output_cells':{},'status':'pending',
                                                      'instruction':'First enumerate every Excel score output and its input cells; then find PDF evidence for each required formula.'})
        write_json(root/'review/formula-plan.json',{'workbook':'../templates/score.xlsx',
                                                    'source_sha256':excel_inventory['source_sha256'],
                                                    'review':{'scoring_spec':{'path':'../analysis/评分细则.md','sha256':None},
                                                              'rules':{'path':'../analysis/scoring_rules.json','sha256':None}},
                                                    'formulas':[],'boundary_cases':[],'status':'pending'})
        write_json(root/'review/project-plan.json',{'application':{'id':'replace-me','name':project_name,'version':'0.1.0'},
                                                    'scoring_inputs':[],'formula_bindings':[],
                                                    'business_decisions':[],'status':'pending'})
        write_json(root/'review/report-map-plan.json',{'template':'../sources/report_template.docx',
                                                       'source_sha256':markers['source_sha256'],
                                                       'output':'../templates/report.docx',
                                                       'manifest':'../analysis/report-map.json',
                                                       'available_sources':{'input':[],'score':[],'system':[]},
                                                       'positions':[{'position_id':item['position_id'],
                                                                     'context':{key:item.get(key) for key in ('context','table_row','table_column','row_text','previous_row_text','preceding_text')},
                                                                     'field_id':None,'source':None,'review_status':'pending'}
                                                                    for item in report_positions['positions']]})
        (root/'PROJECT_STAGE.md').write_text(
            '# 项目构建状态\n\n当前为 `review_required`，不是可运行或已批准的评分应用。\n\n'
            '1. 核对 `analysis/` 中的三文件盘点及原始文件哈希。\n'
            '2. 先逐格盘点 Excel 的输入、得分及汇总单元格，将全部得分目标写入 `review/rule-plan.json` 的 `scope.score_targets`。\n'
            '3. 按每个得分格的公式需求查找 PDF 条款并记录页码、原文证据；据此撰写面向 Excel 的 `评分细则.md` 和公式计划。在 Windows + Excel 上验证后形成带公式评分模板。\n'
            '4. 逐个解释 Word 标记并填写 `review/report-map-plan.json`；生成具名模板。\n'
            '5. 冻结项目配置、运行生成项目测试，再执行 Windows onedir 构建与隔离验收。\n',encoding='utf-8')
        root.replace(destination)
    return {'status':'review_required','project':str(destination),'pages':page_count,
            'workbook_cells':excel_inventory['nonempty_or_styled_cell_count'],'report_positions':len(markers['positions'])}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf',type=Path);parser.add_argument('workbook',type=Path);parser.add_argument('report',type=Path)
    parser.add_argument('--out',required=True,type=Path);parser.add_argument('--name',required=True)
    args=parser.parse_args()
    try:result=initialize(args.pdf,args.workbook,args.report,args.out,project_name=args.name)
    except (OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False));return 1
    print(json.dumps(result,ensure_ascii=False));return 0


if __name__=='__main__':raise SystemExit(main())
