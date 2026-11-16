from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import legacy_sanitisation_plan


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
                "file_count": 4,
                "directory_count": 1,
            }
        ),
        encoding="utf-8",
    )


class LegacySanitisationPlanTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_plan_groups_findings_without_source_content(self) -> None:
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

            plan = legacy_sanitisation_plan.build_plan(root)
            payload = json.loads(json.dumps(plan))
            actions = {entry["recommended_action"] for entry in payload["entries"]}
            joined = json.dumps(payload)

            self.assertEqual(payload["schema"], legacy_sanitisation_plan.SCHEMA)
            self.assertTrue(payload["quarantine_ready"], payload)
            self.assertTrue(payload["requires_sanitisation"])
            self.assertTrue(payload["report_only"])
            self.assertFalse(payload["mutations_performed"])
            self.assertIn("remove_generated_artifact", actions)
            self.assertIn("remove_generated_cache", actions)
            self.assertIn("rewrite_or_redact_text", actions)
            self.assertNotIn("blocked-intranet-host/internal", joined)
            self.assertNotIn("person" + "@" + "blocked.invalid", joined)

    def test_action_priority_prefers_cache_then_generated_then_text_review(self) -> None:
        self.assertEqual(
            legacy_sanitisation_plan.action_for_kinds({"inside_generated_directory", "blocked_generated_artifact"}),
            "remove_generated_cache",
        )
        self.assertEqual(
            legacy_sanitisation_plan.action_for_kinds({"blocked_generated_artifact", "non_example_email_addresses"}),
            "remove_generated_artifact",
        )
        self.assertEqual(
            legacy_sanitisation_plan.action_for_kinds({"high_risk_binary_artifact", "non_example_email_addresses"}),
            "manual_binary_review",
        )
        self.assertEqual(
            legacy_sanitisation_plan.action_for_kinds({"non_example_email_addresses"}),
            "rewrite_or_redact_text",
        )

    def test_plan_writes_only_under_legacy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            write_sanitisation_config(root)
            write_legacy_manifest(root)
            (root / "legacy" / "repo_snapshot").mkdir(parents=True)
            plan = legacy_sanitisation_plan.build_plan(root)

            target = legacy_sanitisation_plan.write_json_plan(root, Path("legacy/reports/plan.json"), plan)

            self.assertTrue(target.is_file())
            with self.assertRaises(ValueError):
                legacy_sanitisation_plan.write_json_plan(root, Path("reports/plan.json"), plan)

    def test_current_plan_is_report_only_when_mirror_exists(self) -> None:
        plan = legacy_sanitisation_plan.build_plan(ROOT)
        self.assertTrue(plan["quarantine_ready"], plan["guardrail_errors"])
        self.assertTrue(plan["report_only"])
        self.assertFalse(plan["mutations_performed"])


if __name__ == "__main__":
    unittest.main()
