"""Naming the bare `{{ }}` slots of a blank DOCX template.

`scripts/name_report_template.py` is Skill-only: it does not exist in the
upstream AutoTesting repository, so these tests are new rather than ported.

The central guarantee under test is that the *supplied* template is never
modified — naming writes a separate copy, so the original keeps its SHA-256 and
the named copy becomes the project's `role: report_template` source.
"""
from __future__ import annotations
import hashlib, io, json, sys, tempfile, unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from docx import Document
import name_report_template, report_markers


class NameReportTemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def build(self, texts=("编号：{{}}",), name="template.docx"):
        path = self.root / name
        document = Document()
        for text in texts:
            document.add_paragraph(text)
        document.save(path)
        return path

    def build_split(self, name="split.docx"):
        path = self.root / name
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("编号：{{")
        paragraph.add_run("}}")
        document.save(path)
        return path

    def run_main(self, template, output, slots, *extra):
        plan = self.root / "plan.json"
        plan.write_text(json.dumps({"template": str(template), "output": str(output), "slots": slots}),
                        encoding="utf-8")
        with patch.object(sys, "argv", ["name_report_template.py", str(plan), *extra]):
            buffer = io.StringIO()
            with redirect_stdout(buffer):
                code = name_report_template.main()
        return code, json.loads(buffer.getvalue())

    def sha(self, path):
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()

    def test_names_a_slot_and_leaves_the_original_untouched(self):
        template = self.build()
        before = self.sha(template)
        output = self.root / "named.docx"
        code, result = self.run_main(template, output, [{"occurrence": 0, "field_id": "score"}])
        self.assertEqual(code, 0)
        self.assertEqual(result["status"], "ok")
        self.assertEqual(self.sha(template), before)
        self.assertEqual(result["template"]["sha256"], before)
        self.assertNotEqual(output, template)
        # scan through the Skill's own marker inventory, not python-docx's view
        self.assertIn("score", report_markers.extract_markers(output)["names"])

    def test_multiple_slots_in_one_paragraph_are_named_by_occurrence(self):
        template = self.build(["{{}} 与 {{}}"])
        output = self.root / "named.docx"
        code, result = self.run_main(template, output, [{"occurrence": 0, "field_id": "first"},
                                                       {"occurrence": 1, "field_id": "second"}])
        self.assertEqual(code, 0)
        self.assertEqual(result["status"], "ok")
        self.assertEqual(report_markers.extract_markers(output)["names"], ["first", "second"])

    def test_slot_split_across_runs_is_named(self):
        # python-docx splits text across runs at every formatting change, so a
        # naive run-by-run replace would miss a slot spanning a boundary.
        template = self.build_split()
        output = self.root / "named.docx"
        code, _ = self.run_main(template, output, [{"occurrence": 0, "field_id": "score"}])
        self.assertEqual(code, 0)
        self.assertEqual(report_markers.extract_markers(output)["names"], ["score"])

    def test_unnamed_slot_is_reported_as_a_blocker(self):
        template = self.build(["{{}} 与 {{}}"])
        output = self.root / "named.docx"
        code, result = self.run_main(template, output, [{"occurrence": 0, "field_id": "first"}])
        self.assertEqual(code, 1)
        self.assertEqual(result["status"], "incomplete")
        self.assertEqual(len(result["unnamedSlots"]), 1)

    def test_dry_run_reports_without_writing(self):
        template = self.build()
        output = self.root / "named.docx"
        code, result = self.run_main(template, output, [{"occurrence": 0, "field_id": "score"}], "--dry-run")
        self.assertEqual(code, 0)
        self.assertTrue(result["dryRun"])
        self.assertIsNone(result["output"])
        self.assertFalse(output.exists())

    def test_output_must_differ_from_the_template(self):
        template = self.build()
        code, result = self.run_main(template, template, [{"occurrence": 0, "field_id": "score"}])
        self.assertEqual(code, 2)
        self.assertIn("original keeps its hash", result["error"])

    def test_invalid_field_id_is_rejected(self):
        template = self.build()
        for bad in ("vehicle.name", "", "2score", "有中文"):
            with self.subTest(field_id=bad):
                code, result = self.run_main(template, self.root / "named.docx",
                                             [{"occurrence": 0, "field_id": bad}])
                self.assertEqual(code, 2)
                self.assertIn("invalid field_id", result["error"])

    def test_occurrence_out_of_range_is_rejected(self):
        template = self.build()
        code, result = self.run_main(template, self.root / "named.docx",
                                     [{"occurrence": 3, "field_id": "score"}])
        self.assertEqual(code, 2)
        self.assertIn("out of range", result["error"])

    def test_missing_template_is_rejected(self):
        code, result = self.run_main(self.root / "absent.docx", self.root / "named.docx",
                                     [{"occurrence": 0, "field_id": "score"}])
        self.assertEqual(code, 2)
        self.assertIn("template not found", result["error"])

    def test_existing_output_is_backed_up_before_overwrite(self):
        template = self.build()
        output = self.root / "named.docx"
        output.write_bytes(b"previous contents")
        code, _ = self.run_main(template, output, [{"occurrence": 0, "field_id": "score"}])
        self.assertEqual(code, 0)
        backup = output.with_suffix(output.suffix + ".previous")
        self.assertTrue(backup.is_file())
        self.assertEqual(backup.read_bytes(), b"previous contents")


if __name__ == "__main__":
    unittest.main()
