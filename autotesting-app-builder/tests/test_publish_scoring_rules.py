from __future__ import annotations
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from publish_scoring_rules import publish


class RulePublisherTest(unittest.TestCase):
    def test_workbook_scope_ignores_unrelated_pdf_pages_but_requires_every_score_cell(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json'
            inventory.write_text(json.dumps({'source_sha256':'f'*64,'pages':[
                {'page_index':1,'blocks':[{'ref':'p0001-b0001','raw_text':'Score rule'}],'tables':[]},
                {'page_index':2,'blocks':[],'tables':[],'image_count':1}]}),encoding='utf-8')
            data={'source_sha256':'f'*64,'reviewed_pages':[1],
                  'scope':{'type':'workbook','workbook_sha256':'a'*64,
                           'score_targets':['Score!B2','Score!B3']},
                  'rules':[{'id':'R1','statement':'Score rule','status':'confirmed',
                            'reviewer':'AI','reviewed_at':'2026-10-11',
                            'source_refs':['p0001-b0001'],'target_cells':['Score!B2']}]}
            plan.write_text(json.dumps(data),encoding='utf-8')
            result=publish(inventory,plan,root/'rules.json',root/'questions.json')
            self.assertEqual(result['coverage_gap_count'],0)
            self.assertEqual(result['uncovered_score_targets'],['Score!B3'])
            self.assertTrue(result['release_blocked'])
            data['rules'][0]['target_cells'].append('Score!B3')
            plan.write_text(json.dumps(data),encoding='utf-8')
            result=publish(inventory,plan,root/'rules.json',root/'questions.json')
            self.assertFalse(result['release_blocked'])

    def test_structured_question_renders_choices_in_review_document(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json'
            inventory.write_text(json.dumps({'source_sha256':'e'*64,'pages':[
                {'page_index':1,'blocks':[{'ref':'p0001-b0001','raw_text':'Boundary unclear'}],'tables':[]}]}),encoding='utf-8')
            plan.write_text(json.dumps({'source_sha256':'e'*64,'reviewed_pages':[1],
                'rules':[{'id':'R1','statement':'边界待定','status':'pending','source_refs':['p0001-b0001'],
                          'questions':[{'question':'等于阈值如何计分？','explanation':'两种解释会改变得分。',
                                        'options':[{'label':'计入低档','effect':'等于阈值时得 2 分'},
                                                   {'label':'计入高档','effect':'等于阈值时得 3 分'}]}]}]}),encoding='utf-8')
            publish(inventory,plan,root/'rules.json',root/'open_questions.json')
            review=(root/'业务评审.md').read_text(encoding='utf-8')
            self.assertIn('疑点与影响：两种解释会改变得分。',review)
            self.assertIn('选项 1：计入低档 — 等于阈值时得 2 分',review)
            self.assertIn('选项 2：计入高档 — 等于阈值时得 3 分',review)
            self.assertIn('其他解释',review)

    def test_clear_ai_confirmed_rule_needs_no_user_review_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json'
            inventory.write_text(json.dumps({'source_sha256':'d'*64,'pages':[
                {'page_index':1,'blocks':[{'ref':'p0001-b0001','raw_text':'Score 2 if <=10'}],'tables':[]}]}),encoding='utf-8')
            plan.write_text(json.dumps({'source_sha256':'d'*64,'reviewed_pages':[1],
                'rules':[{'id':'R1','statement':'输入不大于 10 时得 2 分','status':'confirmed',
                          'reviewer':'AI','reviewed_at':'2026-10-11','source_refs':['p0001-b0001']}]}),encoding='utf-8')
            review=root/'业务评审.md';review.write_text('stale',encoding='utf-8')
            result=publish(inventory,plan,root/'rules.json',root/'open_questions.json')
            self.assertFalse(result['release_blocked'])
            self.assertFalse(review.exists())

    def test_pending_rule_has_page_evidence_and_blocks_release(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json';out=root/'rules.json'
            inventory.write_text(json.dumps({'source_sha256':'a'*64,'pages':[{'page_index':3,'blocks':[{'ref':'p0003-b0001','raw_text':'Score 2 if <=10'}],'tables':[]}]}),encoding='utf-8')
            plan.write_text(json.dumps({'source_sha256':'a'*64,'rules':[{'id':'R1','statement':'输入不大于 10 时得 2 分','status':'pending','source_refs':['p0003-b0001'],'questions':['等于边界是否包含？']}]}),encoding='utf-8')
            questions=root/'open_questions.json'
            result=publish(inventory,plan,out,questions)
            self.assertTrue(result['release_blocked'])
            self.assertGreaterEqual(len(json.loads(questions.read_text())['questions']),1)
            self.assertFalse((root/'评分细则.md').exists())
            self.assertEqual(result['rules'][0]['evidence'][0]['text'],'Score 2 if <=10')
            self.assertEqual(json.loads(questions.read_text())['coverage']['unreviewed_pages'],[3])
            self.assertEqual(result['coverage_gap_count'],1)
            changed=json.loads(plan.read_text(encoding='utf-8'));changed['rules'][0]['source_refs']=['missing']
            plan.write_text(json.dumps(changed),encoding='utf-8')
            with self.assertRaisesRegex(ValueError,'unknown PDF source reference'):
                publish(inventory,plan,out)

    def test_confirmed_rule_needs_page_review_and_scan_page_remains_blocked(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json'
            inventory.write_text(json.dumps({'source_sha256':'b'*64,'pages':[
                {'page_index':1,'blocks':[{'ref':'p0001-b0001','raw_text':'Rule'}],'tables':[],'image_count':0},
                {'page_index':2,'blocks':[],'tables':[],'image_count':1}]}),encoding='utf-8')
            save={'source_sha256':'b'*64,'reviewed_pages':[1,2],
                  'rules':[{'id':'R1','statement':'Rule','status':'confirmed','reviewer':'Person',
                            'reviewed_at':'2026-10-10','source_refs':['p0001-b0001']}]}
            plan.write_text(json.dumps(save),encoding='utf-8')
            result=publish(inventory,plan,root/'rules.json',root/'questions.json')
            self.assertTrue(result['release_blocked'])
            self.assertIn('扫描页',json.dumps(json.loads((root/'questions.json').read_text()),ensure_ascii=False))

    def test_ai_page_evidence_and_human_review_are_published_without_text_extractor(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json'
            inventory.write_text(json.dumps({'source_sha256':'c'*64,'pages':[{'page_index':1},{'page_index':2}]}),encoding='utf-8')
            plan.write_text(json.dumps({'source_sha256':'c'*64,'reviewed_pages':[1],
                'evidence':[{'ref':'ai-p1-rule','page':1,'text':'得分不超过 3 分'}],
                'rules':[{'id':'R1','title':'上限','statement':'该项最多 3 分','status':'pending',
                          'source_refs':['ai-p1-rule'],'questions':['出现两次时如何计算？']}]}),encoding='utf-8')
            result=publish(inventory,plan,root/'rules.json',root/'open_questions.json')
            self.assertEqual(result['coverage_gap_count'],1)
            self.assertEqual(result['rules'][0]['evidence'][0]['kind'],'ai_reviewed_excerpt')
            review=(root/'业务评审.md').read_text(encoding='utf-8')
            self.assertIn('出现两次时如何计算？',review)
            self.assertIn('原文位置：PDF 第 1 页',review)
            self.assertNotIn('页未记录逐页检查',review)


if __name__=='__main__':unittest.main()
