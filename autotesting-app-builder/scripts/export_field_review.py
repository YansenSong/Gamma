"""Export a proposed per-variable review inventory without granting approvals."""
import argparse
import csv
import json
from collections import Counter
from pathlib import Path
import re
from zipfile import ZipFile
from lxml import etree

REQUIRED = ('field_id', 'word_locations', 'meaning', 'suggested_source', 'source',
            'evidence', 'affects_score', 'affects_conclusion', 'review_status',
            'missing_reason', 'blocks_official', 'question', 'category')

def docx_variable_locations(path):
    """Raw XML catches fields hidden by python-docx's merged-cell projection."""
    locations={}
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'):continue
            tree=etree.fromstring(package.read(name)).getroottree()
            for paragraph in tree.findall('.//w:p',ns):
                text=''.join(paragraph.itertext())
                for field in re.findall(r'\{\{\s*(\w+)\s*\}\}',text):
                    locations.setdefault(field,[]).append(name+':'+tree.getpath(paragraph))
    return locations

def validate_review(review):
    fields = review['fields']
    ids = [x['field_id'] for x in fields]
    if len(ids) != len(set(ids)):
        raise ValueError('Duplicate report field IDs')
    for item in fields:
        if set(REQUIRED) - set(item):
            raise ValueError('Incomplete review: ' + item['field_id'])
        if item['category'] not in ('A', 'B', 'C', 'D', 'mapped_candidate', 'engineering_literal'):
            raise ValueError('Unknown category: ' + item['field_id'])
        if item['review_status'] != 'approved' and not item['blocks_official']:
            if item['category'] != 'engineering_literal':
                raise ValueError('Unreviewed field cannot unblock official report')
        if item['category'] == 'C' and item['review_status'] == 'approved' and not item.get('reviewer'):
            raise ValueError('Not-applicable disposition requires named review')
    return {'total': len(fields), 'baseline_missing': sum(x.get('baseline_missing', False) for x in fields),
            'categories': dict(Counter(x['category'] for x in fields)),
            'statuses': dict(Counter(x['review_status'] for x in fields)),
            'formal_blockers': sum(x['blocks_official'] for x in fields)}

def export_review(review, destination):
    summary = validate_review(review)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fields = review['fields']
    with destination.with_suffix('.csv').open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=REQUIRED + ('baseline_missing', 'reviewer', 'decision', 'reason'),lineterminator='\n')
        writer.writeheader()
        for item in fields:
            row = {k: item.get(k, '') for k in writer.fieldnames}
            for key, value in row.items():
                if isinstance(value, (dict, list)):
                    row[key] = json.dumps(value, ensure_ascii=False)
            writer.writerow(row)
    lines = ['# 报告字段审核清单', '', '这是候选来源清单。填写 CSV 的 reviewer / decision / reason 供后续具名审核；填写或导出不授予批准。',
             '', 'A 用户补录；B 已有文件候选；C 有依据的不适用；D 结论或事实阻断。类别与是否阻断分别记录。',
             '', '统计：`' + json.dumps(summary, ensure_ascii=False) + '`', '',
             '|字段 / 含义|Word位置|分类 / 状态|候选来源 / 证据|缺失原因 / 待确认|正式阻断|',
             '|---|---|---|---|---|---|']
    def safe(value):
        return str(value).replace('|', '\\|').replace('\n', ' ')
    for item in fields:
        values = [item['field_id'] + ' / ' + item['meaning'], '; '.join(item['word_locations']),
                  item['category'] + ' / ' + item['review_status'],
                  item['suggested_source'] + ' / ' + json.dumps({k:v for k,v in (item['source'] or {}).items() if k!='observations'}, ensure_ascii=False) + ' / ' + '; '.join(item['evidence']),
                  item['missing_reason'] + ' / ' + item['question'], '是' if item['blocks_official'] else '工程说明']
        lines.append('|' + '|'.join(safe(x) for x in values) + '|')
    destination.with_suffix('.md').write_text('\n'.join(lines) + '\n', encoding='utf-8',newline='\n')
    return summary

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('review', type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(export_review(json.loads(args.review.read_text(encoding='utf-8')), args.out), ensure_ascii=False))
