"""Publish reviewed, source-bound scoring rules from a PDF inventory and rule plan.

The AI/reviewer authors the plan; this tool validates references and refuses to
present unresolved interpretations as confirmed rules.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path


def _refs(inventory):
    found={}
    for page in inventory['pages']:
        for kind in ('blocks','tables'):
            for item in page.get(kind,[]):
                found[item['ref']]={'page':page['page_index'],'text':item.get('raw_text') or item.get('normalized_text') or '', 'kind':kind}
                if kind=='tables':
                    for cell in item.get('cells',[]):
                        found[cell['ref']]={'page':page['page_index'],'text':cell.get('raw_text') or '', 'kind':'cell'}
    return found


def publish(inventory_path:Path,plan_path:Path,markdown_path:Path,rules_path:Path,questions_path:Path|None=None):
    inventory=json.loads(inventory_path.read_text(encoding='utf-8'))
    plan=json.loads(plan_path.read_text(encoding='utf-8'))
    refs=_refs(inventory)
    if inventory.get('source_sha256')!=plan.get('source_sha256'):
        raise ValueError('PDF source SHA-256 does not match rule plan')
    rules=plan.get('rules')
    if not isinstance(rules,list) or not rules:raise ValueError('Rule plan must contain rules')
    ids=set();published=[];questions=[]
    for rule in rules:
        rule_id=rule.get('id');status=rule.get('status')
        if not isinstance(rule_id,str) or not rule_id or rule_id in ids:raise ValueError('Rule IDs must be unique and nonempty')
        ids.add(rule_id)
        if status not in {'pending','confirmed'}:raise ValueError(f'Invalid review status: {rule_id}')
        citations=rule.get('source_refs')
        if not isinstance(citations,list) or not citations or any(ref not in refs for ref in citations):
            raise ValueError(f'Missing or unknown PDF source reference: {rule_id}')
        if not isinstance(rule.get('statement'),str) or not rule['statement'].strip():raise ValueError(f'Missing rule statement: {rule_id}')
        if status=='confirmed' and (not rule.get('reviewer') or not rule.get('reviewed_at')):
            raise ValueError(f'Confirmed rule lacks named review: {rule_id}')
        for question in rule.get('questions',[]):
            questions.append({'rule_id':rule_id,'question':question,'source_refs':citations,'status':'pending'})
        published.append({**rule,'evidence':[{'ref':ref,**refs[ref]} for ref in citations]})
    reviewed_pages=plan.get('reviewed_pages',[])
    all_pages={page['page_index'] for page in inventory['pages']}
    if not isinstance(reviewed_pages,list) or any(not isinstance(page,int) for page in reviewed_pages) or not set(reviewed_pages)<=all_pages:
        raise ValueError('reviewed_pages must contain valid PDF page numbers')
    for page in sorted(all_pages-set(reviewed_pages)):
        questions.append({'rule_id':None,'question':f'第 {page} 页尚未完成条款覆盖审查','source_refs':[],'status':'pending'})
    for page in inventory['pages']:
        if page.get('image_count',0)>0 and not page.get('blocks') and not page.get('tables'):
            questions.append({'rule_id':None,'question':f"第 {page['page_index']} 页可能是扫描页，未提取到文本或表格",'source_refs':[],'status':'pending'})
    result={'schema_version':'gamma-scoring-rules-v1','source_sha256':inventory['source_sha256'],
            'inventory_sha256':hashlib.sha256(inventory_path.read_bytes()).hexdigest(),'rules':published,
            'release_blocked':any(rule['status']!='confirmed' for rule in published) or bool(questions),
            'open_question_count':len(questions)}
    lines=['# 评分细则','','PDF SHA-256：`'+result['source_sha256']+'`','',
           '状态：'+('待审，不得作为正式评分依据' if result['release_blocked'] else '规则已具名审核；公式和项目仍需单独验收'),'']
    for rule in published:
        lines += [f"## {rule['id']} {rule.get('title','')}",'',rule['statement'].strip(),'',
                  '审核状态：'+rule['status'],'',
                  '来源：'+', '.join(f"第{item['page']}页 `{item['ref']}`" for item in rule['evidence']),'']
        if rule.get('questions'):lines += ['待解决：'+ '；'.join(rule['questions']),'']
    markdown_path.parent.mkdir(parents=True,exist_ok=True);rules_path.parent.mkdir(parents=True,exist_ok=True)
    markdown_path.write_text('\n'.join(lines),encoding='utf-8')
    rules_path.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    if questions_path is not None:
        questions_path.parent.mkdir(parents=True,exist_ok=True)
        questions_path.write_text(json.dumps({'schema_version':'gamma-open-questions-v1','source_sha256':inventory['source_sha256'],'questions':questions},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    return result


def main():
    parser=argparse.ArgumentParser();parser.add_argument('inventory',type=Path);parser.add_argument('plan',type=Path)
    parser.add_argument('--markdown',required=True,type=Path);parser.add_argument('--rules',required=True,type=Path);parser.add_argument('--questions',type=Path)
    args=parser.parse_args()
    try:result=publish(args.inventory,args.plan,args.markdown,args.rules,args.questions)
    except (OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False));return 1
    print(json.dumps({'status':'ok','rule_count':len(result['rules']),'release_blocked':result['release_blocked']},ensure_ascii=False));return 0


if __name__=='__main__':raise SystemExit(main())
