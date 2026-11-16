from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import contract_fixture_report


class ContractFixtureReportTests(unittest.TestCase):
    def test_report_summarizes_shared_fixture(self) -> None:
        report = contract_fixture_report.build_report(ROOT)

        self.assertEqual(report["schema"], contract_fixture_report.SCHEMA)
        self.assertEqual(report["message_count"], 2)
        self.assertEqual(report["dispatch_result_count"], 2)
        self.assertTrue(report["dry_run"])
        self.assertTrue(report["example_address_only"])
        self.assertTrue(report["dispatch_state_parity"])
        self.assertEqual(report["validation_errors"], [])
        self.assertIn("group-meetings", report["workflow_references"])
        self.assertIn("induction-mailer", report["workflow_references"])

    def test_cli_json_report_is_report_only(self) -> None:
        run = subprocess.run(
            [
                sys.executable,
                "-B",
                str(ROOT / "python" / "contract_fixture_report.py"),
                "--root",
                str(ROOT),
                "--format",
                "json",
                "--fail-if-invalid",
            ],
            check=True,
            capture_output=True,
            text=True,
        )

        self.assertIn('"mutations_performed": false', run.stdout)
        self.assertIn('"real_email_sending_performed": false', run.stdout)
        self.assertIn('"validation_errors": []', run.stdout)


if __name__ == "__main__":
    unittest.main()
