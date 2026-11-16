from __future__ import annotations

import json
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


class ReleaseAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_release_audit_passes_for_current_scaffold(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertTrue(result["ready"], result["errors"])

    def test_release_audit_is_json_serializable(self) -> None:
        result = release_audit.audit(ROOT)
        payload = json.loads(json.dumps(result))
        self.assertEqual(payload["stage"], release_audit.STAGE)

    def test_release_audit_tracks_expected_stage_7_files(self) -> None:
        result = release_audit.audit(ROOT)
        required = set(result["required_release_files"])
        self.assertIn("LICENSE", required)
        self.assertIn("docs/PUBLIC_RELEASE_CHECKLIST.md", required)
        self.assertIn("docs/PUBLICATION_BOUNDARY.md", required)
        self.assertIn("docs/RELEASE_AUDIT.md", required)
        self.assertIn("docs/RELEASE_CANDIDATE_PLAN.md", required)
        self.assertIn("docs/RELEASE_EXPORT_PLAN.md", required)
        self.assertIn("docs/RELEASE_GATE.md", required)
        self.assertIn("docs/LEGACY_ALTERNATE_MAP.md", required)
        self.assertIn("docs/LEGACY_MIRROR.md", required)
        self.assertIn("docs/REPO_BOUNDARY_AUDIT.md", required)
        self.assertIn("docs/RELEASE_NOTES_TEMPLATE.md", required)
        self.assertIn("migration/legacy_alternate_map.json", required)
        self.assertIn("python/release_candidate_plan.py", required)
        self.assertIn("python/release_export_plan.py", required)
        self.assertIn("python/release_gate.py", required)
        self.assertIn("python/legacy_intake_audit.py", required)
        self.assertIn("python/legacy_mirror_audit.py", required)
        self.assertIn("python/legacy_sanitisation_plan.py", required)
        self.assertIn("python/repo_boundary_audit.py", required)
        self.assertIn("python/tests/test_legacy_intake_audit.py", required)
        self.assertIn("python/tests/test_legacy_mirror_audit.py", required)
        self.assertIn("python/tests/test_legacy_sanitisation_plan.py", required)
        self.assertIn("python/tests/test_release_candidate_plan.py", required)
        self.assertIn("python/tests/test_release_export_plan.py", required)
        self.assertIn("python/tests/test_release_gate.py", required)
        self.assertIn("python/tests/test_repo_boundary_audit.py", required)
        self.assertIn("RunGroupMeetings.vbs", required)
        self.assertIn("RunInductionMailer.vbs", required)
        self.assertIn("OutlookWatchdog.vbs", required)
        self.assertIn("ValidateWorkflow.vbs", required)
        self.assertIn("vbs/entrypoints/scaffold_launcher.vbs", required)
        self.assertIn("vbs/tests/PublicReleaseGateSmokeTest.vbs", required)
        self.assertIn("vbs/tests/ScaffoldAlternateEntrypointsSmokeTest.vbs", required)

    def test_workflow_safety_scan_passes(self) -> None:
        self.assertEqual(release_audit.validate_workflow_safety(ROOT), [])

    def test_runtime_paths_do_not_reference_legacy_mirror(self) -> None:
        self.assertEqual(release_audit.validate_legacy_mirror_runtime_isolation(ROOT), [])

    def test_release_audit_includes_legacy_intake_summary(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertTrue(result["checks"]["legacy_intake"])
        self.assertEqual(result["legacy_intake"]["schema"], "office-communication-scaffold.legacy-intake-audit.v1")
        self.assertFalse(result["legacy_intake"]["raw_copy_allowed"])
        self.assertGreaterEqual(result["legacy_intake"]["entries_count"], 10)

    def test_release_audit_registers_repo_boundary_tool_without_running_repo_gate(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertTrue(result["checks"]["repo_boundary_tool"])
        self.assertEqual(result["repo_boundary_tool"]["schema"], "office-communication-scaffold.repo-boundary-audit.v1")
        self.assertFalse(result["repo_boundary_tool"]["included_in_scaffold_release_gate"])
        self.assertTrue(result["repo_boundary_tool"]["used_by_publication_gate"])
        self.assertTrue(result["repo_boundary_tool"]["legacy_mirror_excluded_from_release"])

    def test_release_audit_excludes_legacy_mirror_from_manifest_summary(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertIn("legacy", result["manifest"]["excluded_release_directories"])

    def test_release_audit_registers_legacy_mirror_quarantine(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertTrue(result["checks"]["legacy_mirror"])
        self.assertEqual(result["legacy_mirror"]["schema"], "office-communication-scaffold.legacy-mirror-audit.v1")
        self.assertTrue(result["legacy_mirror"]["report_only"])
        self.assertFalse(result["legacy_mirror"]["mutations_performed"])

    def test_release_audit_registers_legacy_sanitisation_plan(self) -> None:
        result = release_audit.audit(ROOT)
        self.assertTrue(result["checks"]["legacy_runtime_isolation"])
        self.assertTrue(result["checks"]["legacy_sanitisation_plan"])
        self.assertEqual(
            result["legacy_sanitisation_plan"]["schema"],
            "office-communication-scaffold.legacy-sanitisation-plan.v1",
        )
        self.assertTrue(result["legacy_sanitisation_plan"]["report_only"])
        self.assertFalse(result["legacy_sanitisation_plan"]["mutations_performed"])


if __name__ == "__main__":
    unittest.main()
