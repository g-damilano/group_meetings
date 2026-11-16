from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import release_export_plan


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ReleaseExportPlanTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_export_plan_is_public_scaffold_only(self) -> None:
        plan = release_export_plan.build_export_plan(ROOT)
        paths = [item["path"] for item in plan["files"]]

        self.assertEqual(plan["schema"], release_export_plan.SCHEMA)
        self.assertEqual(release_export_plan.validate_export_plan(plan), [])
        self.assertTrue(plan["report_only"])
        self.assertFalse(plan["publishing_performed"])
        self.assertFalse(plan["packaging_performed"])
        self.assertFalse(plan["archive_created"])
        self.assertEqual(plan["legacy_paths_in_files"], 0)
        self.assertIn("legacy", plan["excluded_release_directories"])
        self.assertIn("python/release_export_plan.py", paths)
        self.assertIn("python/release_gate.py", paths)
        self.assertIn("docs/RELEASE_EXPORT_PLAN.md", paths)
        self.assertIn("docs/RELEASE_GATE.md", paths)
        self.assertFalse(any(path.startswith("legacy/") for path in paths))

    def test_export_plan_is_json_serializable_without_private_payloads(self) -> None:
        plan = release_export_plan.build_export_plan(ROOT)
        payload = json.loads(json.dumps(plan))
        joined = json.dumps(payload)

        self.assertEqual(payload["schema"], release_export_plan.SCHEMA)
        self.assertNotIn("legacy/repo_snapshot", joined)
        self.assertNotIn("person" + "@" + "blocked.invalid", joined)

    def test_kind_summary_matches_file_entries(self) -> None:
        plan = release_export_plan.build_export_plan(ROOT)

        self.assertEqual(sum(plan["kind_summary"].values()), plan["file_count"])
        self.assertEqual(plan["file_count"], len(plan["files"]))
        self.assertGreater(plan["total_bytes"], 0)

    def test_write_export_plan_is_limited_to_private_legacy_reports(self) -> None:
        plan = release_export_plan.build_export_plan(ROOT)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            (temp_root / "legacy").mkdir()
            target = release_export_plan.write_json_plan(temp_root, Path("legacy/reports/export.json"), plan)

            self.assertTrue(target.is_file())
            with self.assertRaises(ValueError):
                release_export_plan.write_json_plan(temp_root, Path("reports/export.json"), plan)


if __name__ == "__main__":
    unittest.main()
