from __future__ import annotations
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from validate_scoring_spec import validate


class ScoringSpecValidationTest(unittest.TestCase):
    def test_checks_identity_and_coverage_without_fixing_format(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);doc=root/'评分细则.md';rules=root/'rules.json'
            rules.write_text(json.dumps({'source_sha256':'a'*64,'release_blocked':True,
                'rules':[{'id':'R1'},{'id':'R2'}]}),encoding='utf-8')
            doc.write_text('# 待审草稿\n\nPDF: '+ 'a'*64 +'\n\n|规则|说明|\n|---|---|\n|R1|文字|\n|R2|文字|',encoding='utf-8')
            self.assertEqual(validate(doc,rules)['status'],'passed')
            doc.write_text('# 正式细则\n\nR1',encoding='utf-8')
            result=validate(doc,rules)
            self.assertEqual(result['status'],'failed')
            self.assertEqual(len(result['issues']),3)


if __name__=='__main__':unittest.main()
