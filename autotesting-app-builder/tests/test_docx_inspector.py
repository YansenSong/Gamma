"""Contract of the frozen vendored DOCX template inspector.

Ported from the upstream AutoTesting repository's tests/test_docx_inspector.py.
`vendor/` is a frozen copy that must stay byte-identical to upstream, so this
test pins the vendored contract: if anyone edits the vendored inspector, these
assertions fail.
"""
from __future__ import annotations
import sys, tempfile, unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "vendor"))

from docx import Document
from ivista_reporting.template_inspector import inspect_docx_template
from ivista_reporting.template_contract import TemplateContractError

class DocxTemplateInspectorTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.path=Path(self.temp.name)/"reference.docx"
        document=Document(); document.add_paragraph("车辆：{{ vehicle_name }}")
        split=document.add_paragraph(); split.add_run("拆分标记：{{ split_"); split.add_run("field }}")
        table=document.add_table(rows=2,cols=2)
        table.cell(0,0).text="得分"; table.cell(0,1).text="{{ score }}"
        table.cell(1,0).text="报告编号"; table.cell(1,1).text="[[report_id]]"
        document.sections[0].header.paragraphs[0].text="检验报告"
        document.save(self.path)
    def tearDown(self): self.temp.cleanup()
    def test_reference_sample_values_are_never_promoted_to_facts(self):
        result=inspect_docx_template(self.path,document_role="reference_sample")
        self.assertFalse(result["sample_values_are_facts"])
        self.assertEqual({x["name"] for x in result["field_candidates"]},{"score","vehicle_name","split_field","report_id"})
        self.assertTrue(all(x["mapping_status"]=="needs_human_mapping" for x in result["field_candidates"]))
        self.assertTrue(any("header" in x["location"] for x in result["body"]+result["sections"][0]["header"]))
    def test_role_is_required_and_constrained(self):
        with self.assertRaises(TemplateContractError):
            inspect_docx_template(self.path,document_role="completed_truth")

if __name__=="__main__": unittest.main()
