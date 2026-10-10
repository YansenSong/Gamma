"""Export contextual DOCX marker positions for semantic review, without mapping."""
from __future__ import annotations
import argparse
import csv
import json
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from report_markers import extract_markers

NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _text(node):
    return ''.join(node.itertext()).strip()


def positions_with_context(template):
    report=extract_markers(template)
    grouped={}
    for item in report['positions']:grouped.setdefault(item['part'],[]).append(item)
    enriched=[]
    with ZipFile(template) as package:
        for part,positions in grouped.items():
            root=etree.fromstring(package.read(part));paragraphs=list(root.iter(f'{{{NS}}}p'))
            for item in positions:
                p=paragraphs[item['paragraph']]
                cells=p.xpath('ancestor::w:tc[1]',namespaces={'w':NS})
                table_row=table_column=None;row_text='';previous_row=''
                if cells:
                    cell=cells[0];rows=cell.xpath('ancestor::w:tr[1]',namespaces={'w':NS})
                    if rows:
                        row=rows[0];siblings=row.getparent().findall(f'{{{NS}}}tr')
                        table_row=siblings.index(row)+1
                        row_cells=row.findall(f'{{{NS}}}tc');table_column=row_cells.index(cell)+1
                        row_text=' | '.join(_text(node) for node in row_cells)
                        if table_row>1:previous_row=' | '.join(_text(node) for node in siblings[table_row-2].findall(f'{{{NS}}}tc'))
                preceding=''
                for prior in reversed(paragraphs[max(0,item['paragraph']-8):item['paragraph']]):
                    candidate=_text(prior)
                    if candidate and '{{' not in candidate:
                        preceding=candidate[:200];break
                enriched.append({**item,'table_row':table_row,'table_column':table_column,
                                 'row_text':row_text[:500],'previous_row_text':previous_row[:500],
                                 'preceding_text':preceding})
    return {'schema_version':'gamma-report-positions-v1','template_sha256':report['source_sha256'],
            'position_count':len(enriched),'positions':enriched}


def export(template,output_json,output_csv):
    result=positions_with_context(template)
    output_json=Path(output_json);output_csv=Path(output_csv)
    output_json.parent.mkdir(parents=True,exist_ok=True);output_csv.parent.mkdir(parents=True,exist_ok=True)
    output_json.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    columns=('position_id','part','paragraph','context','table_row','table_column','row_text','previous_row_text','preceding_text',
             'field_id','source_kind','source_id','review_status')
    with output_csv.open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=columns);writer.writeheader()
        for item in result['positions']:
            row={key:item.get(key,'') for key in columns}
            row.update({'field_id':'','source_kind':'','source_id':'','review_status':'pending'})
            writer.writerow(row)
    return result


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('template',type=Path)
    parser.add_argument('--json',required=True,type=Path);parser.add_argument('--csv',required=True,type=Path)
    args=parser.parse_args()
    try:result=export(args.template,args.json,args.csv)
    except (OSError,ValueError,KeyError) as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False));return 1
    print(json.dumps({'status':'ok','positions':result['position_count']},ensure_ascii=False));return 0


if __name__=='__main__':raise SystemExit(main())
