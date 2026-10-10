"""Build an unmapped inventory; optional user inputs must be explicitly declared."""
import argparse
import json
import re
from pathlib import Path
from export_field_review import docx_variable_locations, export_review


def build(root):
    spec=json.loads((root/'config/project.json').read_text(encoding='utf-8'))
    analysis=json.loads((root/'analysis/report.json').read_text(encoding='utf-8'))
    locations=docx_variable_locations(root/spec['report_output']['template'])
    metadata=spec.get('review_preview',{}).get('metadata_fields',{})
    meanings={}
    for block in analysis['body']:
        for row in block.get('rows',[]):
            text=' / '.join(dict.fromkeys(c['text'] for c in row['cells']))
            for field in re.findall(r'\{\{\s*(\w+)\s*\}\}',text):
                meanings.setdefault(field,re.sub(r'\{\{.*?\}\}','[动态字段]',text))
        for field in re.findall(r'\{\{\s*(\w+)\s*\}\}',block.get('text','')):
            meanings.setdefault(field,re.sub(r'\{\{.*?\}\}','[动态字段]',block['text']))
    entries=[]
    for field in sorted(locations):
        meta=field in metadata
        entries.append({'field_id':field,'word_locations':locations[field],
          'meaning':metadata[field]['label'] if meta else meanings.get(field,field),
          'suggested_source':'用户提交真实委托/车辆/试验身份记录' if meta else '原始试验记录及经批准的规则与映射',
          'source':{'input_id':metadata[field]['input_id'],'kind':'user_supplied_unverified'} if meta else None,
          'evidence':['DOCX '+analysis['source_sha256']], 'affects_score':not meta,
          'affects_conclusion':True,'review_status':'pending','missing_reason':'缺少具名来源与字段语义确认；未匹配Excel单元格',
          'blocks_official':True,'question':'请确认含义、单位、工况与来源记录；补录不授予审批，不得由分数推断试验事实。',
          'category':'A' if meta else 'D','baseline_missing':True})
    unknown=set(metadata)-set(locations)
    if unknown:raise ValueError('Metadata fields absent from original template: '+str(unknown))
    return {'schema_version':'autotesting-field-review-v1','template_sha256':analysis['source_sha256'],
            'business_status':'blocked','categories_are_proposals':True,'not_applicable_approved':0,
            'static_review_questions':['正式报告编号、日期和签名须由真实业务审批提供。','静态模板内容须逐页审核，动态字段清单不构成全模板业务验收。'],
            'fields':entries}


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('project_root',type=Path)
    args=parser.parse_args();root=args.project_root.resolve();review=build(root)
    review['summary']=export_review(review,root/'analysis/report-field-review')
    text=json.dumps(review,ensure_ascii=False,indent=2)+'\n'
    for relative in ('analysis/report-mapping-review.json','config/report-field-review.json'):
        (root/relative).write_text(text,encoding='utf-8',newline='\n')
    print(json.dumps(review['summary'],ensure_ascii=False))
