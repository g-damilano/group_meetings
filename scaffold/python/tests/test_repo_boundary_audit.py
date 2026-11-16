from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import repo_boundary_audit


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
            }
        ),
        encoding="utf-8",
    )


class RepoBoundaryAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_clean_boundary_passes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo_root = Path(temp_dir)
            scaffold_root = repo_root / "scaffold"
            write_sanitisation_config(scaffold_root)
            (repo_root / "README.md").write_text("Public overview\n", encoding="utf-8")
            (scaffold_root / "private.log").write_text("ignored inside scaffold\n", encoding="utf-8")

            result = repo_boundary_audit.audit(repo_root, scaffold_root)

            self.assertTrue(result["ready"], result)
            self.assertTrue(result["report_only"])
            self.assertFalse(result["outside_root_edits_performed"])
            self.assertEqual(result["findings"], [])

    def test_boundary_findings_are_paths_and_categories_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo_root = Path(temp_dir)
            scaffold_root = repo_root / "scaffold"
            write_sanitisation_config(scaffold_root)
            (repo_root / "workflow.log").write_text("generated\n", encoding="utf-8")
            (repo_root / "archive.zip").write_bytes(b"zip")
            (repo_root / "Attachment.pdf").write_bytes(b"pdf")
            (repo_root / "notes.md").write_text(
                "blocked-intranet-host\n" + "person" + "@blocked.invalid\n",
                encoding="utf-8",
            )
            cache = repo_root / "tool" / "__pycache__"
            cache.mkdir(parents=True)
            (cache / "module.pyc").write_bytes(b"compiled")

            result = repo_boundary_audit.audit(repo_root, scaffold_root)
            payload = json.loads(json.dumps(result))
            kinds = {finding["kind"] for finding in payload["findings"]}
            joined = json.dumps(payload)

            self.assertFalse(payload["ready"])
            self.assertIn("blocked_generated_artifact", kinds)
            self.assertIn("high_risk_binary_artifact", kinds)
            self.assertIn("blocked_text_patterns", kinds)
            self.assertIn("non_example_email_addresses", kinds)
            self.assertIn("generated_directory", kinds)
            self.assertNotIn("blocked-intranet-host/internal", joined)
            self.assertNotIn("person" + "@blocked.invalid", joined)

    def test_current_repository_boundary_audit_is_report_only(self) -> None:
        result = repo_boundary_audit.audit(ROOT.parent, ROOT)
        self.assertTrue(result["report_only"])
        self.assertFalse(result["outside_root_edits_performed"])
        self.assertGreaterEqual(result["scanned_paths"], 1)


if __name__ == "__main__":
    unittest.main()
