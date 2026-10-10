"""Unscored draft preview for blocked review projects. No Excel or business scores."""
import hashlib
import json
import shutil
import tempfile
import uuid
import re
from zipfile import ZipFile
from pathlib import Path
from lxml import etree
from report_safety import mark_engineering_copy


def render_review_template(template, destination, context):
    """Replace simple review variables, including runs and merged-cell XML.

    Keep surrounding runs, drawings, fields and package members unchanged.
    This preview deliberately rejects Jinja expressions/control statements.
    """
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    seen=set()
    with ZipFile(template) as source, ZipFile(destination,'w') as target:
        for member in source.infolist():
            data=source.read(member.filename)
            if member.filename.startswith('word/') and member.filename.endswith('.xml'):
                xml=etree.fromstring(data)
                for paragraph in xml.findall('.//w:p',ns):
                    nodes=paragraph.findall('.//w:t',ns)
                    text=''.join(node.text or '' for node in nodes)
                    matches=list(re.finditer(r'\{\{\s*(\w+)\s*\}\}',text))
                    residue=re.sub(r'\{\{\s*(\w+)\s*\}\}','',text)
                    if '{{' in residue or '{%' in residue:
                        raise ValueError('工程预览仅支持简单字段；发现未支持的模板表达式')
                    offsets=[];position=0
                    for node in nodes:
                        offsets.append((position,position+len(node.text or ''),node));position=offsets[-1][1]
                    for match in reversed(matches):
                        name=match.group(1)
                        if name not in context:raise ValueError('未知报告字段：'+name)
                        seen.add(name);affected=[entry for entry in offsets if entry[0]<match.end() and entry[1]>match.start()]
                        first=True
                        for start,end,node in affected:
                            value=node.text or ''
                            left=max(0,match.start()-start);right=min(end-start,match.end()-start)
                            node.text=value[:left]+(context[name] if first else '')+value[right:]
                            node.set('{http://www.w3.org/XML/1998/namespace}space','preserve');first=False
                data=etree.tostring(xml,xml_declaration=True,encoding='UTF-8',standalone=True)
            target.writestr(member,data)
    if seen!=set(context):raise ValueError('审核清单与实际模板字段不一致')


def create_review_preview(spec, facts, output_dir, resource_path):
    if not isinstance(facts,dict):raise ValueError('审核补录记录须为 JSON 对象')
    config=spec['review_preview']
    template=resource_path(spec['report_output']['template'])
    source=next(x for x in spec['sources'] if x['role']=='report_template')
    actual=hashlib.sha256(template.read_bytes()).hexdigest()
    if actual!=source['sha256']:raise ValueError('报告模板哈希不一致')
    review=json.loads(resource_path(spec['report_review']['path']).read_text(encoding='utf-8'))
    if review['template_sha256']!=actual:raise ValueError('审核清单与模板不一致')
    context={item['field_id']:'未提供' for item in review['fields']}
    supplied={}
    for field, definition in config['metadata_fields'].items():
        value=facts.get(definition['input_id'],'')
        if not isinstance(value,str):raise ValueError('补录信息须为文本：'+field)
        if value.strip():
            context[field]=value.strip()+'（用户补录，未核验）';supplied[field]=value.strip()
    output=Path(output_dir).resolve();output.mkdir(parents=True,exist_ok=True)
    stage=Path(tempfile.mkdtemp(prefix='.review-staging-',dir=output))
    try:
        report=stage/'engineering-review.docx'
        render_review_template(template,report,context)
        protection=mark_engineering_copy(report,config['report_safety'])
        audit={'schema_version':'autotesting-unscored-review-preview-v1','project_id':spec['application']['id'],
               'project_spec_sha256':hashlib.sha256(resource_path('config/project.json').read_bytes()).hexdigest(),
               'business_status':'blocked','official_output_allowed':False,'scoring_performed':False,
               'template_sha256':actual,'user_supplied_unverified':supplied,
               'missing_fields':[key for key in context if key not in supplied],
               'report_safety':protection,'report_sha256':hashlib.sha256(report.read_bytes()).hexdigest(),
               'review_note_unverified':facts.get('review_note',''),
               'notice':'工程预览仅为字段审核；没有评分表、评分结论或正式审批。'}
        (stage/'engineering-review-audit.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2)+'\n',encoding='utf-8',newline='\n')
        final=output/('review-preview-'+uuid.uuid4().hex)
        stage.rename(final)
        return {'bundle':str(final),'report':str(final/report.name),'audit':str(final/'engineering-review-audit.json')}
    except Exception:
        shutil.rmtree(stage);raise
