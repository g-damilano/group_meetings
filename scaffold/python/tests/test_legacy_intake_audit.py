from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = ROOT.parent
sys.path.insert(0, str(ROOT / "python"))

import legacy_intake_audit


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class LegacyIntakeAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_current_legacy_intake_audit_passes(self) -> None:
        result = legacy_intake_audit.audit(REPO_ROOT, ROOT)
        self.assertTrue(result["ready"], result["errors"])
        self.assertGreaterEqual(result["entries_count"], 10)
        self.assertFalse(result["raw_copy_allowed"])

    def test_current_audit_is_json_serializable(self) -> None:
        result = legacy_intake_audit.audit(REPO_ROOT, ROOT)
        payload = json.loads(json.dumps(result))
        self.assertEqual(payload["schema"], legacy_intake_audit.SCHEMA)
        self.assertEqual(payload["scaffold_root"], "scaffold")

    def test_private_source_signal_is_reported_without_blocking_scaffold_target(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo_root = Path(temp_dir)
            scaffold_root = repo_root / "scaffold"
            (scaffold_root / "migration").mkdir(parents=True)
            (scaffold_root / "fixtures" / "sanitisation").mkdir(parents=True)
            (scaffold_root / "vbs").mkdir()

            required_wrappers = [
                "RunGroupMeetings.vbs",
                "RunInductionMailer.vbs",
                "OutlookWatchdog.vbs",
                "ValidateWorkflow.vbs",
            ]
            for wrapper in required_wrappers:
                (repo_root / wrapper).write_text("' legacy wrapper\n", encoding="utf-8")
                (scaffold_root / wrapper).write_text("' scaffold wrapper\n", encoding="utf-8")

            (repo_root / "Legacy.vbs").write_text(
                "Const PRIVATE_LINK = \"blocked-intranet-host/internal\"\n",
                encoding="utf-8",
            )
            (scaffold_root / "Alternate.vbs").write_text("' sanitized alternate\n", encoding="utf-8")
            (scaffold_root / "fixtures" / "sanitisation" / "blocked_patterns.json").write_text(
                json.dumps(
                    {
                        "allowed_email_domains": ["example.invalid"],
                        "blocked_text_patterns": ["blocked-intranet-host"],
                        "blocked_generated_suffixes": [".log", ".pyc", ".pyo", ".zip", ".msg"],
                        "blocked_generated_directory_names": ["__pycache__"],
                    }
                ),
                encoding="utf-8",
            )
            (scaffold_root / "migration" / "legacy_alternate_map.json").write_text(
                json.dumps(
                    {
                        "schema": "office-communication-scaffold.legacy-alternate-map.v1",
                        "policy": {
                            "copying_into_scaffold_authorised": True,
                            "moving_legacy_files_allowed": False,
                            "outside_root_edits_allowed": False,
                            "raw_private_content_allowed": False,
                            "live_send_allowed": False,
                            "destructive_mailbox_actions_allowed": False,
                        },
                        "entries": [
                            {
                                "legacy_path": "Legacy.vbs",
                                "scaffold_path": "Alternate.vbs",
                                "status": "ported_alternate",
                                "role": "test alternate",
                                "safety": "dry-run only",
                            }
                        ]
                        + [
                            {
                                "legacy_path": wrapper,
                                "scaffold_path": wrapper,
                                "status": "ported_alternate",
                                "role": "required wrapper",
                                "safety": "dry-run only",
                            }
                            for wrapper in required_wrappers
                        ],
                    }
                ),
                encoding="utf-8",
            )

            result = legacy_intake_audit.audit(repo_root, scaffold_root)

            self.assertTrue(result["ready"], result["errors"])
            self.assertEqual(result["sources_requiring_sanitisation_count"], 1)
            self.assertEqual(result["source_sanitisation_required"], ["Legacy.vbs"])
            self.assertFalse(result["entries"][0]["scaffold_target_requires_sanitisation"])


if __name__ == "__main__":
    unittest.main()
