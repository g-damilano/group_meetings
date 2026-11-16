from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import release_manifest


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ReleaseManifestTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_manifest_paths_are_relative_sorted_and_hashed(self) -> None:
        manifest = release_manifest.build_manifest(ROOT)
        self.assertEqual(release_manifest.validate_manifest(manifest), [])
        paths = [item["path"] for item in manifest["files"]]
        self.assertEqual(paths, sorted(paths))
        self.assertIn(".gitignore", paths)
        self.assertIn("LICENSE", paths)
        self.assertIn("migration/legacy_alternate_map.json", paths)
        self.assertIn("python/legacy_intake_audit.py", paths)
        self.assertIn("python/release_export_plan.py", paths)
        self.assertIn("python/release_gate.py", paths)
        self.assertIn("python/release_manifest.py", paths)
        self.assertIn("docs/RELEASE_EXPORT_PLAN.md", paths)
        self.assertIn("docs/RELEASE_GATE.md", paths)
        self.assertIn("docs/RELEASE_MANIFEST.md", paths)
        for path in paths:
            self.assertFalse(path.startswith("/"))
            self.assertNotIn("\\", path)
            self.assertNotIn(":", path)

    def test_manifest_is_json_serializable(self) -> None:
        manifest = release_manifest.build_manifest(ROOT)
        payload = json.loads(json.dumps(manifest))
        self.assertEqual(payload["schema"], release_manifest.SCHEMA)
        self.assertEqual(payload["root"], "scaffold")

    def test_generated_artifacts_are_excluded_from_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "docs").mkdir()
            (root / "docs" / "guide.md").write_text("guide", encoding="utf-8")
            (root / "__pycache__").mkdir()
            (root / "__pycache__" / "ignored.pyc").write_bytes(b"compiled")
            (root / "run.log").write_text("generated", encoding="utf-8")
            (root / "legacy").mkdir()
            (root / "legacy" / "private.md").write_text("person" + "@" + "contoso.invalid", encoding="utf-8")
            (root / "legacy" / "legacy.zip").write_bytes(b"archive")

            manifest = release_manifest.build_manifest(root)
            paths = [item["path"] for item in manifest["files"]]

            self.assertEqual(paths, ["docs/guide.md"])
            self.assertIn("legacy", manifest["excluded_release_directories"])
            self.assertNotIn("legacy/private.md", paths)
            self.assertNotIn("legacy/legacy.zip", manifest["excluded_generated_artifacts"])
            self.assertIn("__pycache__", manifest["excluded_generated_artifacts"])
            self.assertIn("__pycache__/ignored.pyc", manifest["excluded_generated_artifacts"])
            self.assertIn("run.log", manifest["excluded_generated_artifacts"])

    def test_file_kind_classification(self) -> None:
        self.assertEqual(release_manifest.file_kind("LICENSE"), "license")
        self.assertEqual(release_manifest.file_kind("docs/RELEASE_MANIFEST.md"), "documentation")
        self.assertEqual(release_manifest.file_kind("vbs/core/contracts.vbs"), "vbs-core")
        self.assertEqual(release_manifest.file_kind("vbs/workflows/group.vbs"), "vbs-workflow")
        self.assertEqual(release_manifest.file_kind("python/release_manifest.py"), "python-tool")
        self.assertEqual(release_manifest.file_kind("migration/legacy_alternate_map.json"), "migration")


if __name__ == "__main__":
    unittest.main()
