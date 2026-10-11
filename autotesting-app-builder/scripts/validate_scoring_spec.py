"""Check identity and rule coverage of an AI-authored scoring specification.

This cannot judge whether the prose, tables, formulas or business interpretation
are correct. It deliberately does not generate or prescribe the Markdown layout.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


def validate(markdown_path: Path, rules_path: Path) -> dict:
    document=markdown_path.read_text(encoding='utf-8')
    rules=json.loads(rules_path.read_text(encoding='utf-8'))
    issues=[]
    pdf_hash=rules.get('source_sha256')
    if not isinstance(pdf_hash,str) or pdf_hash not in document:
        issues.append('文档缺少当前 PDF 的 SHA-256，无法确认来源版本')
    for rule in rules.get('rules',[]):
        identifier=rule.get('id')
        if not isinstance(identifier,str) or not re.search(r'(?<![\w-])'+re.escape(identifier)+r'(?![\w-])',document):
            issues.append(f'文档未覆盖规则编号：{identifier}')
    if rules.get('release_blocked') and not re.search(r'待审|草稿|未获批准|尚未确认',document):
        issues.append('规则仍有待审项，文档须明确标注草稿状态')
    return {'status':'passed' if not issues else 'failed','rule_count':len(rules.get('rules',[])),
            'issues':issues,'note':'仅校验来源身份、编号覆盖与待审标记；内容深度和业务正确性须由 AI 与业务审核人核对'}


def main() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('markdown',type=Path);parser.add_argument('rules',type=Path)
    args=parser.parse_args()
    try:result=validate(args.markdown,args.rules)
    except (OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print(json.dumps({'status':'failed','issues':[str(exc)]},ensure_ascii=False));return 1
    print(json.dumps(result,ensure_ascii=False,indent=2));return 0 if result['status']=='passed' else 1


if __name__=='__main__':raise SystemExit(main())
