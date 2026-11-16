from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import legacy_mirror_audit


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


def write_sanitisation_config(scaffold_root: Path) -> None:
    target = scaffold_root / "fixtures" / "sanitisation"
    target.mkdir(parents=True)
    (target / "blocked_patterns.json").write_text(
        json.dumps(
            {
                "allowed_email_domains": ["example.invalid"],
                "blocked_text_patterns": ["blocked-intranet-host"],
                "blocked_generated_suffixes": [".log", ".pyc", ".pyo", ".zip", ".msg"],
                "blocked_generated_directory_names": ["__pycache__"],
                "release_excluded_directory_names": ["legacy"],
            }
        ),
        encoding="utf-8",
    )


def write_legacy_manifest(scaffold_root: Path) -> None:
    legacy_root = scaffold_root / "legacy"
    legacy_root.mkdir(parents=True, exist_ok=True)
    (legacy_root / "manifest.json").write_text(
        json.dumps(
            {
                "schema": "office-communication-scaffold.legacy-mirror.v1",
                "excluded_from_public_release": True,
                "copy_mode": "copy-only",
                "moved_files": False,
                "file_count": 3,
                "directory_count": 1,
            }
        ),
        encoding="utf-8",
    )


class LegacyMirrorAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_missing_mirror_is_quarantine_ready(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            write_sanitisation_config(root)

            result = legacy_mirror_audit.audit(root)

            self.assertTrue(result["quarantine_ready"], result)
            self.assertFalse(result["mirror_present"])
            self.assertEqual(result["findings"], [])

    def test_mirror_findings_are_report_only_and_content_free(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            write_sanitisation_config(root)
            write_legacy_manifest(root)
            mirror = root / "legacy" / "repo_snapshot"
            mirror.mkdir(parents=True)
            (mirror / "workflow.log").write_text("generated", encoding="utf-8")
            (mirror / "archive.zip").write_bytes(b"zip")
            (mirror / "notes.md").write_text(
                "blocked-intranet-host\n" + "person" + "@" + "blocked.invalid\n",
                encoding="utf-8",
            )
            cache = mirror / "tool" / "__pycache__"
            cache.mkdir(parents=True)
            (cache / "module.pyc").write_bytes(b"compiled")

            result = legacy_mirror_audit.audit(root)
            payload = json.loads(json.dumps(result))
            kinds = {finding["kind"] for finding in payload["findings"]}
            joined = json.dumps(payload)

            self.assertTrue(payload["quarantine_ready"], payload)
            self.assertTrue(payload["requires_sanitisation"])
            self.assertTrue(payload["report_only"])
            self.assertFalse(payload["mutations_performed"])
            self.assertIn("blocked_generated_artifact", kinds)
            self.assertIn("blocked_text_patterns", kinds)
            self.assertIn("non_example_email_addresses", kinds)
            self.assertIn("generated_directory", kinds)
            self.assertNotIn("blocked-intranet-host/internal", joined)
            self.assertNotIn("person" + "@" + "blocked.invalid", joined)

    def test_guardrail_errors_catch_non_copy_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            write_sanitisation_config(root)
            legacy_root = root / "legacy"
            (legacy_root / "repo_snapshot").mkdir(parents=True)
            (legacy_root / "manifest.json").write_text(
                json.dumps(
                    {
                        "excluded_from_public_release": False,
                        "copy_mode": "move",
                        "moved_files": True,
                    }
                ),
                encoding="utf-8",
            )

            result = legacy_mirror_audit.audit(root)

            self.assertFalse(result["quarantine_ready"])
            self.assertGreaterEqual(len(result["guardrail_errors"]), 3)

    def test_json_report_writes_only_under_legacy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            write_sanitisation_config(root)
            write_legacy_manifest(root)
            result = legacy_mirror_audit.audit(root)

            report = legacy_mirror_audit.write_json_report(root, Path("legacy/reports/audit.json"), result)

            self.assertTrue(report.is_file())
            with self.assertRaises(ValueError):
                legacy_mirror_audit.report_target_path(root, Path("reports/audit.json"))

    def test_current_mirror_guardrails_are_valid_when_present(self) -> None:
        result = legacy_mirror_audit.audit(ROOT)
        self.assertTrue(result["quarantine_ready"], result["guardrail_errors"])
        self.assertFalse(result["mutations_performed"])


if __name__ == "__main__":
    unittest.main()
