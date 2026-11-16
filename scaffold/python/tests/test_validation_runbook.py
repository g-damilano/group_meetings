from __future__ import annotations

import shutil
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import release_audit


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ValidationRunbookTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_vbs_runner_is_release_required(self) -> None:
        result = release_audit.audit(ROOT)
        required = set(result["required_release_files"])
        self.assertIn("vbs/tests/RunAllSmokeTests.vbs", required)
        self.assertIn("docs/VALIDATION_RUNBOOK.md", required)

    def test_vbs_runner_owns_no_office_or_mailbox_actions(self) -> None:
        text = (ROOT / "vbs" / "tests" / "RunAllSmokeTests.vbs").read_text(encoding="utf-8")
        forbidden = [
            'CreateObject("Outlook.Application")',
            'CreateObject("Excel.Application")',
            "GetDefaultFolder",
            ".Send",
            ".Delete",
        ]
        for token in forbidden:
            self.assertNotIn(token, text)

    def test_validation_runbook_documents_safe_order(self) -> None:
        text = (ROOT / "docs" / "VALIDATION_RUNBOOK.md").read_text(encoding="utf-8")
        self.assertIn("RunAllSmokeTests.vbs", text)
        self.assertIn("validate_scaffold.py", text)
        self.assertIn("legacy_mirror_audit.py", text)
        self.assertIn("legacy_sanitisation_plan.py", text)
        self.assertIn("release_audit.py", text)
        self.assertIn("release_candidate_plan.py", text)
        self.assertIn("release_export_plan.py", text)
        self.assertIn("release_gate.py", text)
        self.assertIn("does not perform real email sending", text)


if __name__ == "__main__":
    unittest.main()
