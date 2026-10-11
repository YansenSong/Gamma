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


def publish(inventory_path:Path,plan_path:Path,rules_path:Path,questions_path:Path|None=None):
    inventory=json.loads(inventory_path.read_text(encoding='utf-8'))
    plan=json.loads(plan_path.read_text(encoding='utf-8'))
    refs=_refs(inventory)
    all_pages={page['page_index'] for page in inventory['pages']}
    for item in plan.get('evidence',[]):
        ref=item.get('ref');page=item.get('page');excerpt=item.get('text')
        if (not isinstance(ref,str) or not ref or ref in refs or not isinstance(page,int)
                or page not in all_pages or not isinstance(excerpt,str) or not excerpt.strip()):
            raise ValueError('AI evidence needs a unique ref, valid PDF page and nonempty excerpt')
        refs[ref]={'page':page,'text':excerpt.strip(),'kind':'ai_reviewed_excerpt'}
    if inventory.get('source_sha256')!=plan.get('source_sha256'):
        raise ValueError('PDF source SHA-256 does not match rule plan')
    rules=plan.get('rules')
    if not isinstance(rules,list) or not rules:raise ValueError('Rule plan must contain rules')
    scope=plan.get('scope',{'type':'full_pdf'})
    if not isinstance(scope,dict) or scope.get('type') not in {'full_pdf','workbook'}:
        raise ValueError('scope.type must be full_pdf or workbook')
    workbook_targets=scope.get('score_targets',[])
    if scope['type']=='workbook' and (not isinstance(scope.get('workbook_sha256'),str)
            or len(scope['workbook_sha256'])!=64 or not isinstance(workbook_targets,list)
            or not workbook_targets or any(not isinstance(cell,str) or '!' not in cell for cell in workbook_targets)
            or len(set(workbook_targets))!=len(workbook_targets)):
        raise ValueError('Workbook scope needs workbook_sha256 and unique score_targets')
    ids=set();published=[];questions=[];covered_targets=set();cited_pages=set()
    for rule in rules:
        rule_id=rule.get('id');status=rule.get('status')
        if not isinstance(rule_id,str) or not rule_id or rule_id in ids:raise ValueError('Rule IDs must be unique and nonempty')
        ids.add(rule_id)
        if status not in {'pending','confirmed'}:raise ValueError(f'Invalid review status: {rule_id}')
        citations=rule.get('source_refs')
        if not isinstance(citations,list) or not citations or any(ref not in refs for ref in citations):
            raise ValueError(f'Missing or unknown PDF source reference: {rule_id}')
        if not isinstance(rule.get('statement'),str) or not rule['statement'].strip():raise ValueError(f'Missing rule statement: {rule_id}')
        targets=rule.get('target_cells',[])
        if scope['type']=='workbook':
            if (not isinstance(targets,list) or not targets
                    or any(target not in workbook_targets for target in targets)):
                raise ValueError(f'Rule {rule_id} needs score targets in workbook scope')
            covered_targets.update(targets)
        cited_pages.update(refs[ref]['page'] for ref in citations)
        if status=='confirmed' and (not rule.get('reviewer') or not rule.get('reviewed_at')):
            raise ValueError(f'Confirmed rule lacks named review: {rule_id}')
        rule_questions=rule.get('questions',[])
        if status=='pending' and not rule_questions:
            rule_questions=['该规则仍待确认；请明确尚未解决的业务歧义或所需材料。']
        for question in rule_questions:
            if isinstance(question,str):
                item={'question':question}
            elif isinstance(question,dict):
                item={key:question.get(key) for key in ('question','explanation','options') if key in question}
                options=item.get('options',[])
                if (not isinstance(item.get('question'),str) or not item['question'].strip()
                        or not isinstance(item.get('explanation'),str) or not item['explanation'].strip()
                        or not isinstance(options,list) or not 2<=len(options)<=3
                        or any(not isinstance(option,dict) or not isinstance(option.get('label'),str)
                               or not option['label'].strip() or not isinstance(option.get('effect'),str)
                               or not option['effect'].strip() for option in options)):
                    raise ValueError(f'Question for {rule_id} needs explanation and 2-3 labeled options with effects')
                if len({option['label'] for option in options})!=len(options):
                    raise ValueError(f'Question for {rule_id} has duplicate option labels')
            else:
                raise ValueError(f'Invalid question for {rule_id}')
            questions.append({'rule_id':rule_id,**item,'source_refs':citations,'status':'pending'})
        published.append({**rule,'evidence':[{'ref':ref,**refs[ref]} for ref in citations]})
    reviewed_pages=plan.get('reviewed_pages',[])
    if not isinstance(reviewed_pages,list) or any(not isinstance(page,int) for page in reviewed_pages) or not set(reviewed_pages)<=all_pages:
        raise ValueError('reviewed_pages must contain valid PDF page numbers')
    required_pages=cited_pages if scope['type']=='workbook' else all_pages
    coverage_gaps=sorted(required_pages-set(reviewed_pages))
    target_gaps=sorted(set(workbook_targets)-covered_targets)
    for page in inventory['pages']:
        if (page['page_index'] in required_pages and page.get('image_count',0)>0
                and not page.get('blocks') and not page.get('tables')):
            questions.append({'rule_id':None,'question':f"第 {page['page_index']} 页可能是扫描页，未提取到文本或表格",'source_refs':[],'status':'pending'})
    result={'schema_version':'gamma-scoring-rules-v1','source_sha256':inventory['source_sha256'],
            'inventory_sha256':hashlib.sha256(inventory_path.read_bytes()).hexdigest(),'rules':published,
            'scope':scope,'uncovered_score_targets':target_gaps,
            'release_blocked':any(rule['status']!='confirmed' for rule in published) or bool(questions) or bool(coverage_gaps) or bool(target_gaps),
            'open_question_count':len(questions),'coverage_gap_count':len(coverage_gaps)}
    rules_path.parent.mkdir(parents=True,exist_ok=True)
    rules_path.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    if questions_path is not None:
        questions_path.parent.mkdir(parents=True,exist_ok=True)
        questions_path.write_text(json.dumps({'schema_version':'gamma-open-questions-v1','source_sha256':inventory['source_sha256'],
            'questions':questions,'coverage':{'reviewed_pages':sorted(set(reviewed_pages)),
            'unreviewed_pages':coverage_gaps,'uncovered_score_targets':target_gaps,
            'status':'pending' if coverage_gaps or target_gaps else 'complete'}},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        review_path=questions_path.with_name('业务评审.md')
        if questions:
            review_lines=['# 评分规则业务评审','',
                          '请对下面存在歧义或读取障碍的事项给出结论和依据；填写此文件不会自动批准规则。','']
            for index,question in enumerate(questions,1):
                rule=next((item for item in published if item['id']==question['rule_id']),None)
                review_lines += [f'## {index}. {rule.get("title") or rule["id"] if rule else "PDF 读取问题"}', '',
                                 '待确认：'+question['question'],'']
                if question.get('explanation'):
                    review_lines += ['疑点与影响：'+question['explanation'],'']
                if rule:
                    pages=sorted({item['page'] for item in rule['evidence']})
                    review_lines += ['当前理解：'+rule['statement'],'',
                                     '原文位置：PDF 第 '+ '、'.join(str(page) for page in pages)+' 页','']
                for option_index,option in enumerate(question.get('options',[]),1):
                    review_lines += [f'选项 {option_index}：{option["label"]} — {option["effect"]}']
                if question.get('options'):
                    review_lines += ['其他解释（请写明）：________________','']
                review_lines += ['您的判断（可直接修改上面的理解）：________________',
                                 '判断依据／还需要的材料：________________','审核人及日期：________________','']
            review_path.write_text('\n'.join(review_lines),encoding='utf-8')
        else:
            review_path.unlink(missing_ok=True)
    return result


def main():
    parser=argparse.ArgumentParser();parser.add_argument('inventory',type=Path);parser.add_argument('plan',type=Path)
    parser.add_argument('--rules',required=True,type=Path);parser.add_argument('--questions',type=Path)
    args=parser.parse_args()
    try:result=publish(args.inventory,args.plan,args.rules,args.questions)
    except (OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False));return 1
    print(json.dumps({'status':'ok','rule_count':len(result['rules']),'release_blocked':result['release_blocked']},ensure_ascii=False));return 0


if __name__=='__main__':raise SystemExit(main())
