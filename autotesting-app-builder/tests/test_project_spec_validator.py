"""Validation contract of scripts/validate_project_spec.py.

Ported from the upstream AutoTesting repository's tests/test_project_spec_validator.py.
The only change from upstream is the SCRIPT path below: upstream resolved the
Skill through the repository checkout at `.agents/skills/autotesting-app-builder/`,
which does not exist once the Skill ships standalone.
"""
from __future__ import annotations
import hashlib, importlib.util, json, tempfile, unittest
from pathlib import Path
from unittest.mock import patch
from docx import Document
from openpyxl import Workbook

ROOT=Path(__file__).resolve().parents[1]
SCRIPT=ROOT/"scripts"/"validate_project_spec.py"
spec_loader=importlib.util.spec_from_file_location("autotesting_spec_validator",SCRIPT)
validator=importlib.util.module_from_spec(spec_loader); spec_loader.loader.exec_module(validator)

class ProjectSpecValidatorTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.root=Path(self.temp.name); (self.root/"config").mkdir()
        (self.root/"inputs").mkdir(parents=True)
        # Upstream wrote the three sources as raw stub bytes. The Skill's validator
        # now parses the report template to extract its {{ }} markers
        # (docx_markers / validate_report_template), so this fixture must build a
        # real DOCX whose marker set equals report_output.fields exactly.
        (self.root/"inputs"/"charter.pdf").write_bytes(b"charter")
        document=Document(); document.add_paragraph("得分：{{ score }}"); document.save(self.root/'inputs/report.docx')
        book=Workbook();book.active.title='Inputs';book['Inputs']['B4']='=IF(B2<=200,100,0)';book.save(self.root/'inputs/score.xlsx');book.close()
        sources=[{"role":role,"path":f"inputs/{name}","sha256":hashlib.sha256((self.root/"inputs"/name).read_bytes()).hexdigest()}
                 for role,name in (("charter","charter.pdf"),("workbook","score.xlsx"),("report_template","report.docx"))]
        self.formula={'sheet':'Inputs','cell':'B4','formula':'=IF(B2<=200,100,0)'}
        self.expression_hash=validator.digest(self.formula)
        self.spec={"schema_version":"autotesting-project-v1","application":{"id":"sample","name":"Sample","version":"0.1"},
          "release_status":"engineering_sample","sources":sources,"validation_state":{"status":"approved","official_output_allowed":True},
          "business_decisions":[{"id":"Q1","choice":"absolute","status":"confirmed","reviewer":"reviewer","reason":"checked","decided_at":"now"}],
          "scoring":{"mode":"excel_native","inputs":[{"id":"latency","type":"number","required":True}],
                     "formula_bindings":[{**self.formula,"output":"score","source_refs":["p1"],
                               "decision_bindings":{"Q1":{"absolute":self.expression_hash}}}]},
          "excel_output":{"template":"inputs/score.xlsx","input_cells":{"latency":{"sheet":"Inputs","cell":"B2"}},"output_cells":{"score":{"sheet":"Inputs","cell":"B4"}}},
          "report_output":{"template":"inputs/report.docx","fields":{"score":{"kind":"score","id":"score"}}}}
        self.path=self.root/"config"/"project.json"
    def tearDown(self): self.temp.cleanup()
    def test_matching_decision_and_expression_is_valid(self):
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        self.assertEqual(validator.validate(self.path)["status"],"valid")
    def test_python_mode_and_expressions_are_rejected(self):
        self.spec['scoring']['mode']='deterministic_python'
        self.path.write_text(json.dumps(self.spec),encoding='utf-8')
        with self.assertRaisesRegex(validator.SpecError,'Python scoring is removed'):validator.validate(self.path)
        self.spec['scoring']['mode']='excel_native';self.spec['scoring']['rules']=[{'expression':{'op':'value','value':100}}]
        self.path.write_text(json.dumps(self.spec),encoding='utf-8')
        with self.assertRaisesRegex(validator.SpecError,'Python rule expressions are removed'):validator.validate(self.path)
    def test_selected_decision_with_different_expression_is_blocked(self):
        self.spec["business_decisions"][0]["choice"]="reference"
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        with self.assertRaisesRegex(validator.SpecError,"not implemented by Excel formula"):
            validator.validate(self.path)
    def test_changed_expression_invalidates_binding_hash(self):
        self.spec["scoring"]["formula_bindings"][0]["formula"]="=IF(B2<=200,90,0)"
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        with self.assertRaisesRegex(validator.SpecError,"Formula binding differs"):
            validator.validate(self.path)
    def test_missing_source_file_is_blocked(self):
        (self.root/"inputs"/"report.docx").unlink()
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        with self.assertRaisesRegex(validator.SpecError,"source file is missing"):
            validator.validate(self.path)
    def test_symlink_source_cannot_escape_project_root(self):
        outside=self.root.parent/(self.root.name+"-outside.pdf"); outside.write_bytes(b"outside")
        link=self.root/"inputs"/"escape.pdf"
        self.spec["sources"][0]={"role":"charter","path":"inputs/escape.pdf","sha256":hashlib.sha256(outside.read_bytes()).hexdigest()}
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        real_resolve=Path.resolve
        def resolve_with_external_target(path,*args,**kwargs):
            if path==link: return real_resolve(outside)
            return real_resolve(path,*args,**kwargs)
        with patch.object(Path,"resolve",resolve_with_external_target):
            with self.assertRaisesRegex(validator.SpecError,"escapes project root"):
                validator.validate(self.path)
        outside.unlink(missing_ok=True)
    def test_business_validated_requires_hash_bound_named_reviews(self):
        self.spec["release_status"]="business_validated"
        source_hashes=[item["sha256"] for item in self.spec["sources"]]
        self.spec["business_decisions"][0]["source_hashes"]=source_hashes
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        with self.assertRaisesRegex(validator.SpecError,"requires named approvals"):
            validator.validate(self.path)
        for approval in ("scoring_rules","excel_mapping","report_mapping","calculation_authority"):
            self.spec.setdefault("approvals",{})[approval]={"reviewer":"Reviewer One","reason":"Reviewed","decided_at":"2026-01-01T00:00:00Z","source_hashes":source_hashes,
                "content_sha256":validator.approval_content_hash(self.spec,approval)}
        decision=self.spec["business_decisions"][0]
        decision["source_hashes"]=source_hashes
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        self.assertEqual(validator.validate(self.path)["releaseStatus"],"business_validated")
        self.spec["excel_output"]["input_cells"]["latency"]["cell"]="B3"
        self.path.write_text(json.dumps(self.spec),encoding="utf-8")
        with self.assertRaisesRegex(validator.SpecError,"excel_mapping approval content hash is stale"):
            validator.validate(self.path)
    def test_each_scoped_approval_is_invalidated_by_its_content(self):
        self.spec["release_status"]="business_validated"
        source_hashes=[item["sha256"] for item in self.spec["sources"]]
        self.spec["business_decisions"][0]["source_hashes"]=source_hashes
        self.spec["approvals"]={name:{"reviewer":"Reviewer One","reason":"Reviewed","decided_at":"2026-01-01T00:00:00Z",
            "source_hashes":source_hashes,"content_sha256":validator.approval_content_hash(self.spec,name)}
            for name in ("scoring_rules","excel_mapping","report_mapping","calculation_authority")}
        mutations={
            "scoring_rules":lambda: self.spec["scoring"]["inputs"][0].update({"label":"changed"}),
            "excel_mapping":lambda: self.spec["excel_output"]["input_cells"]["latency"].update({"cell":"B3"}),
            "report_mapping":lambda: self.spec["report_output"]["fields"]["score"].update({"label":"changed"}),
            "calculation_authority":lambda: self.spec["scoring"].update({"authority":{"type":"changed"}}),
        }
        for approval,mutate in mutations.items():
            with self.subTest(approval=approval):
                snapshot=json.loads(json.dumps(self.spec)); mutate(); self.path.write_text(json.dumps(self.spec),encoding="utf-8")
                with self.assertRaisesRegex(validator.SpecError,f"{approval} approval content hash is stale"):
                    validator.validate(self.path)
                self.spec=snapshot
if __name__=="__main__": unittest.main()
