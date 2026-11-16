from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import release_gate


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ReleaseGateTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_current_gate_is_report_only_and_repo_boundary_is_advisory(self) -> None:
        gate = release_gate.build_gate(ROOT)

        self.assertEqual(gate["schema"], release_gate.SCHEMA)
        self.assertEqual(gate["gate"], release_gate.GATE_PASS)
        self.assertTrue(gate["report_only"])
        self.assertFalse(gate["publishing_performed"])
        self.assertFalse(gate["packaging_performed"])
        self.assertFalse(gate["mutations_performed"])
        self.assertFalse(gate["real_email_sending_performed"])
        self.assertTrue(gate["scaffold_release_ready"])
        self.assertFalse(gate["repo_boundary_ready"])
        self.assertNotIn("repo_boundary_findings_present", gate["blockers"])
        self.assertIn("repo_boundary_findings_present", gate["advisories"])
        self.assertEqual(gate["release_manifest"]["legacy_paths_in_files"], 0)

    def test_gate_payload_omits_private_finding_entries(self) -> None:
        gate = release_gate.build_gate(ROOT)
        payload = json.loads(json.dumps(gate))
        joined = json.dumps(payload)

        self.assertNotIn("findings", payload["repo_boundary"])
        self.assertNotIn("findings", payload["legacy_mirror"])
        self.assertNotIn("entries", payload["legacy_sanitisation_plan"])
        self.assertNotIn("person" + "@" + "blocked.invalid", joined)

    def test_write_gate_is_limited_to_private_legacy_reports(self) -> None:
        gate = release_gate.build_gate(ROOT)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            (temp_root / "legacy").mkdir()
            target = release_gate.write_json_gate(temp_root, Path("legacy/reports/gate.json"), gate)

            self.assertTrue(target.is_file())
            with self.assertRaises(ValueError):
                release_gate.write_json_gate(temp_root, Path("reports/gate.json"), gate)

    def test_fail_if_blocked_exit_code_is_explicit(self) -> None:
        blocked = {"gate": release_gate.GATE_BLOCKED}
        passed = {"gate": release_gate.GATE_PASS}

        self.assertEqual(release_gate.exit_code_for_gate(blocked, True), 1)
        self.assertEqual(release_gate.exit_code_for_gate(blocked, False), 0)
        self.assertEqual(release_gate.exit_code_for_gate(passed, True), 0)


if __name__ == "__main__":
    unittest.main()
