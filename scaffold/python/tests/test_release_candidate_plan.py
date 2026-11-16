from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import release_candidate_plan


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ReleaseCandidatePlanTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_current_repo_plan_is_report_only_and_boundary_advisory(self) -> None:
        plan = release_candidate_plan.build_plan(ROOT)

        self.assertEqual(plan["schema"], release_candidate_plan.SCHEMA)
        self.assertEqual(plan["decision"], release_candidate_plan.DECISION_READY)
        self.assertTrue(plan["report_only"])
        self.assertFalse(plan["publishing_performed"])
        self.assertFalse(plan["mutations_performed"])
        self.assertTrue(plan["scaffold_release_ready"])
        self.assertFalse(plan["repo_boundary_ready"])
        self.assertTrue(plan["repo_boundary_advisory"])
        self.assertNotIn("repo_boundary_findings_present", plan["blockers"])
        self.assertIn("repo_boundary_findings_present", plan["advisories"])
        self.assertEqual(plan["release_manifest"]["legacy_paths_in_files"], 0)

    def test_plan_is_json_serializable_without_private_payloads(self) -> None:
        plan = release_candidate_plan.build_plan(ROOT)
        payload = json.loads(json.dumps(plan))
        joined = json.dumps(payload)

        self.assertEqual(payload["schema"], release_candidate_plan.SCHEMA)
        self.assertNotIn("findings", payload["repo_boundary"])
        self.assertNotIn("findings", payload["legacy_mirror"])
        self.assertNotIn("entries", payload["legacy_sanitisation_plan"])
        self.assertNotIn("person" + "@" + "blocked.invalid", joined)

    def test_write_plan_is_limited_to_private_legacy_reports(self) -> None:
        plan = release_candidate_plan.build_plan(ROOT)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            (temp_root / "legacy").mkdir()
            target = release_candidate_plan.write_json_plan(temp_root, Path("legacy/reports/plan.json"), plan)

            self.assertTrue(target.is_file())
            with self.assertRaises(ValueError):
                release_candidate_plan.write_json_plan(temp_root, Path("reports/plan.json"), plan)

    def test_required_human_actions_ignore_advisory_boundary_state(self) -> None:
        scaffold = {"ready": True}
        boundary = {"ready": False}
        sanitisation = {"requires_sanitisation": True}

        actions = release_candidate_plan.required_human_actions(scaffold, boundary, sanitisation)

        self.assertEqual(actions, ["perform_manual_public_release_review"])

    def test_advisory_actions_reflect_boundary_and_private_mirror_state(self) -> None:
        boundary = {"ready": False}
        sanitisation = {"requires_sanitisation": True}

        actions = release_candidate_plan.advisory_actions(boundary, sanitisation)

        self.assertIn("repo_wide_public_cleanup_outside_scaffold_is_advisory_for_scaffold_development", actions)
        self.assertIn("private_legacy_mirror_sanitisation_is_advisory_until_mirror_content_is_promoted", actions)


if __name__ == "__main__":
    unittest.main()
