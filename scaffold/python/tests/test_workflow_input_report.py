from __future__ import annotations

import json
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import workflow_input_report


class WorkflowInputReportTests(unittest.TestCase):
    def test_public_examples_are_route_ready(self) -> None:
        report = workflow_input_report.build_report(ROOT)

        self.assertEqual(report["schema"], workflow_input_report.SCHEMA)
        self.assertTrue(report["report_only"])
        self.assertFalse(report["mutations_performed"])
        self.assertFalse(report["real_email_sending_performed"])
        self.assertEqual(report["validation_errors"], [])

        group = report["workflow_inputs"]["group_meetings"]
        self.assertTrue(group["dry_run_route_ready"])
        self.assertEqual(group["row_count"], 2)
        self.assertEqual(group["configured_meeting_count"], 1)
        self.assertEqual(group["no_meeting_count"], 1)
        self.assertEqual(group["first_route_planned_message_estimate"], 6)

        induction = report["workflow_inputs"]["induction"]
        self.assertTrue(induction["dry_run_route_ready"])
        self.assertEqual(induction["recipients"]["active_count"], 2)
        self.assertEqual(induction["template_manifest"]["enabled_count"], 2)
        self.assertEqual(induction["planned_message_estimate"], 3)

    def test_invalid_recipient_fixture_reports_active_missing_email(self) -> None:
        report = workflow_input_report.build_report(
            ROOT,
            induction_recipients="vbs/tests/fixtures/induction_recipients_invalid_fixture.csv",
            induction_templates="vbs/tests/fixtures/induction_template_manifest_fixture.csv",
        )

        self.assertFalse(report["workflow_inputs"]["induction"]["dry_run_route_ready"])
        self.assertTrue(
            any("ACTIVE_RECIPIENT_MISSING_EMAIL" in error for error in report["validation_errors"]),
            report["validation_errors"],
        )

    def test_invalid_template_fixture_reports_enabled_row_issues(self) -> None:
        report = workflow_input_report.build_report(
            ROOT,
            induction_recipients="vbs/tests/fixtures/induction_recipients_fixture.csv",
            induction_templates="vbs/tests/fixtures/induction_template_manifest_invalid_fixture.csv",
        )

        errors = "\n".join(report["validation_errors"])
        self.assertFalse(report["workflow_inputs"]["induction"]["dry_run_route_ready"])
        self.assertIn("TEMPLATE_MODE_INVALID", errors)
        self.assertIn("DELIVERY_ACTION_INVALID", errors)
        self.assertIn("DEFER_MINUTES_INVALID", errors)

    def test_cli_json_fail_if_invalid_succeeds_for_public_examples(self) -> None:
        completed = subprocess.run(
            [
                sys.executable,
                "-B",
                str(ROOT / "python" / "workflow_input_report.py"),
                "--root",
                str(ROOT),
                "--format",
                "json",
                "--fail-if-invalid",
            ],
            check=False,
            capture_output=True,
            text=True,
        )

        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        payload = json.loads(completed.stdout)
        self.assertEqual(payload["schema"], workflow_input_report.SCHEMA)
        self.assertEqual(payload["validation_errors"], [])


if __name__ == "__main__":
    unittest.main()
