from __future__ import annotations
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from publish_scoring_rules import publish


class RulePublisherTest(unittest.TestCase):
    def test_pending_rule_has_page_evidence_and_blocks_release(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);inventory=root/'pdf.json';plan=root/'plan.json';md=root/'评分细则.md';out=root/'rules.json'
            inventory.write_text(json.dumps({'source_sha256':'a'*64,'pages':[{'page_index':3,'blocks':[{'ref':'p0003-b0001','raw_text':'Score 2 if <=10'}],'tables':[]}]}),encoding='utf-8')
            plan.write_text(json.dumps({'source_sha256':'a'*64,'rules':[{'id':'R1','statement':'输入不大于 10 时得 2 分','status':'pending','source_refs':['p0003-b0001'],'questions':['等于边界是否包含？']}]}),encoding='utf-8')
            questions=root/'open_questions.json'
            result=publish(inventory,plan,md,out,questions)
            self.assertTrue(result['release_blocked'])
            self.assertGreaterEqual(len(json.loads(questions.read_text())['questions']),1)
            self.assertIn('第3页 `p0003-b0001`',md.read_text(encoding='utf-8'))
            self.assertEqual(result['rules'][0]['evidence'][0]['text'],'Score 2 if <=10')
            changed=json.loads(plan.read_text(encoding='utf-8'));changed['rules'][0]['source_refs']=['missing']
            plan.write_text(json.dumps(changed),encoding='utf-8')
            with self.assertRaisesRegex(ValueError,'unknown PDF source reference'):
                publish(inventory,plan,md,out)

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
            result=publish(inventory,plan,root/'rules.md',root/'rules.json',root/'questions.json')
            self.assertTrue(result['release_blocked'])
            self.assertIn('扫描页',json.dumps(json.loads((root/'questions.json').read_text()),ensure_ascii=False))


if __name__=='__main__':unittest.main()
