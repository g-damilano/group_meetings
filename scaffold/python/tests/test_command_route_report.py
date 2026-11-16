from __future__ import annotations

import json
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import command_route_report


class CommandRouteReportTests(unittest.TestCase):
    def test_public_routes_match_input_report_counts(self) -> None:
        report = command_route_report.build_report(ROOT)

        self.assertEqual(report["schema"], command_route_report.SCHEMA)
        self.assertEqual(report["quickstart_fixture"]["schema"], command_route_report.QUICKSTART_ROUTE_SCHEMA)
        self.assertTrue(report["report_only"])
        self.assertFalse(report["mutations_performed"])
        self.assertFalse(report["real_email_sending_performed"])
        self.assertEqual(report["validation_errors"], [])

        group_fields = report["routes"]["group_meetings"]["fields"]
        induction_fields = report["routes"]["induction"]["fields"]
        watchdog_fields = report["routes"]["watchdog"]["fields"]

        self.assertEqual(report["routes"]["help"]["exit_code"], 0)
        self.assertIn("Exit codes:", report["routes"]["help"]["output"])
        self.assertEqual(report["routes"]["validate"]["exit_code"], 0)
        self.assertEqual(report["routes"]["validate"]["fields"]["scaffold_command"], "validate")
        self.assertIn("Validation passed.", report["routes"]["validate"]["output"])
        self.assertEqual(group_fields["state"], "COMPOSED")
        self.assertEqual(group_fields["planned_count"], "6")
        self.assertEqual(group_fields["dry_run"], "true")
        self.assertEqual(induction_fields["state"], "COMPOSED")
        self.assertEqual(induction_fields["planned_count"], "3")
        self.assertEqual(induction_fields["dry_run"], "true")
        self.assertEqual(watchdog_fields["state"], "COMPOSED")
        self.assertEqual(watchdog_fields["report_only"], "true")
        self.assertEqual(watchdog_fields["repair_actions_taken"], "0")

        self.assertEqual(report["blocked_routes"]["group_meetings_unsafe_send_blocked"]["exit_code"], 3)
        self.assertEqual(report["blocked_routes"]["watchdog_repair_blocked"]["exit_code"], 3)
        self.assertEqual(
            report["blocked_routes"]["group_meetings_unsafe_send_blocked"]["fields"]["state"],
            "FAILED_SAFE",
        )
        self.assertEqual(report["blocked_routes"]["watchdog_repair_blocked"]["fields"]["state"], "FAILED_SAFE")

    def test_quickstart_route_fixture_is_valid(self) -> None:
        self.assertEqual(command_route_report.validate_quickstart_fixture(ROOT), [])

    def test_parser_ignores_non_key_value_lines(self) -> None:
        parsed = command_route_report.parse_key_value_output("hello\nstate=COMPOSED\nplanned_count=3\n")

        self.assertEqual(parsed, {"state": "COMPOSED", "planned_count": "3"})

    def test_cli_json_fail_if_invalid_succeeds_for_public_routes(self) -> None:
        completed = subprocess.run(
            [
                sys.executable,
                "-B",
                str(ROOT / "python" / "command_route_report.py"),
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
        self.assertEqual(payload["schema"], command_route_report.SCHEMA)
        self.assertEqual(payload["validation_errors"], [])


if __name__ == "__main__":
    unittest.main()
