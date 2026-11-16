from __future__ import annotations

import unittest
from pathlib import Path
import shutil
import tempfile

import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

import validate_scaffold


def remove_generated_artifacts() -> None:
    legacy_root = ROOT / "legacy"
    for path in ROOT.rglob("__pycache__"):
        if legacy_root in path.parents:
            continue
        shutil.rmtree(path, ignore_errors=True)


class ValidateScaffoldTests(unittest.TestCase):
    def setUp(self) -> None:
        remove_generated_artifacts()

    def tearDown(self) -> None:
        remove_generated_artifacts()

    @classmethod
    def tearDownClass(cls) -> None:
        remove_generated_artifacts()

    def test_dispatch_states_match_contract(self) -> None:
        self.assertEqual(validate_scaffold.validate_dispatch_states(ROOT), [])

    def test_dispatch_states_match_across_vbs_schema_and_fixture(self) -> None:
        self.assertEqual(validate_scaffold.extract_vbs_dispatch_states(ROOT), validate_scaffold.EXPECTED_STATES)
        self.assertEqual(validate_scaffold.workflow_schema_dispatch_states(ROOT), validate_scaffold.EXPECTED_STATES)
        self.assertEqual(validate_scaffold.validate_dispatch_state_parity(ROOT), [])

    def test_workflow_schema_names_all_canonical_definitions(self) -> None:
        self.assertEqual(validate_scaffold.validate_workflow_schema(ROOT), [])

    def test_shared_contract_fixture_is_valid(self) -> None:
        self.assertEqual(validate_scaffold.validate_shared_contract_fixture(ROOT), [])

    def test_watchdog_report_fixture_is_valid(self) -> None:
        self.assertEqual(validate_scaffold.validate_watchdog_report_fixture(ROOT), [])

    def test_watchdog_report_fixture_blocks_unmarked_mutation(self) -> None:
        path = ROOT / "fixtures" / "watchdog" / "report_only_health_fixture.json"
        data = validate_scaffold.load_json(path)
        self.assertIsInstance(data, dict)
        assert isinstance(data, dict)
        items = data["items"]
        assert isinstance(items, list)
        for item in items:
            if isinstance(item, dict) and item.get("item_id") == "unmarked-user-mail":
                item["would_modify"] = True
                break

        errors = validate_scaffold.validate_watchdog_report_fixture_data(ROOT, "fixture.json", data)

        self.assertTrue(
            any("unmarked mail must never be repair-eligible or mutable" in error for error in errors),
            errors,
        )

    def test_watchdog_report_fixture_blocks_uncertain_repair(self) -> None:
        path = ROOT / "fixtures" / "watchdog" / "report_only_health_fixture.json"
        data = validate_scaffold.load_json(path)
        self.assertIsInstance(data, dict)
        assert isinstance(data, dict)
        items = data["items"]
        assert isinstance(items, list)
        for item in items:
            if isinstance(item, dict) and item.get("item_id") == "managed-uncertain-age":
                item["dispatch_state"] = "QUEUED_OUTBOX"
                item["replace_eligible"] = True
                break

        errors = validate_scaffold.validate_watchdog_report_fixture_data(ROOT, "fixture.json", data)

        self.assertTrue(
            any("uncertain watchdog state must be ACTION_REQUIRED or FAILED_SAFE" in error for error in errors),
            errors,
        )
        self.assertTrue(
            any("uncertain watchdog state must not be replace-eligible" in error for error in errors),
            errors,
        )

    def test_watchdog_repair_fixture_is_valid(self) -> None:
        self.assertEqual(validate_scaffold.validate_watchdog_repair_fixture(ROOT), [])

    def test_watchdog_repair_fixture_blocks_unmarked_mutation(self) -> None:
        path = ROOT / "fixtures" / "watchdog" / "managed_repair_guardrail_fixture.json"
        data = validate_scaffold.load_json(path)
        self.assertIsInstance(data, dict)
        assert isinstance(data, dict)
        cases = data["cases"]
        assert isinstance(cases, list)
        for repair_case in cases:
            if isinstance(repair_case, dict) and repair_case.get("case_id") == "commit-unmarked-mail":
                actions = repair_case["actions"]
                assert isinstance(actions, list)
                action = actions[0]
                assert isinstance(action, dict)
                action["would_modify_fake_object"] = True
                action["stale_deleted"] = True
                break

        errors = validate_scaffold.validate_watchdog_repair_fixture_data(ROOT, "repair-fixture.json", data)

        self.assertTrue(
            any("unmarked mail must never be repair-eligible or mutable" in error for error in errors),
            errors,
        )

    def test_watchdog_repair_fixture_blocks_preview_mutation(self) -> None:
        path = ROOT / "fixtures" / "watchdog" / "managed_repair_guardrail_fixture.json"
        data = validate_scaffold.load_json(path)
        self.assertIsInstance(data, dict)
        assert isinstance(data, dict)
        cases = data["cases"]
        assert isinstance(cases, list)
        for repair_case in cases:
            if isinstance(repair_case, dict) and repair_case.get("case_id") == "preview-stale-managed":
                actions = repair_case["actions"]
                assert isinstance(actions, list)
                action = actions[0]
                assert isinstance(action, dict)
                action["replacement_created"] = True
                break

        errors = validate_scaffold.validate_watchdog_repair_fixture_data(ROOT, "repair-fixture.json", data)

        self.assertTrue(
            any("preview mode must not create, delete, or mutate" in error for error in errors),
            errors,
        )

    def test_watchdog_repair_fixture_blocks_live_outlook_behaviour(self) -> None:
        path = ROOT / "fixtures" / "watchdog" / "managed_repair_guardrail_fixture.json"
        data = validate_scaffold.load_json(path)
        self.assertIsInstance(data, dict)
        assert isinstance(data, dict)
        cases = data["cases"]
        assert isinstance(cases, list)
        first_case = cases[0]
        assert isinstance(first_case, dict)
        actions = first_case["actions"]
        assert isinstance(actions, list)
        action = actions[0]
        assert isinstance(action, dict)
        action["send_committed"] = True
        action["folder_accessed"] = True
        action["outlook_created"] = True

        errors = validate_scaffold.validate_watchdog_repair_fixture_data(ROOT, "repair-fixture.json", data)

        self.assertTrue(any("must not commit send" in error for error in errors), errors)
        self.assertTrue(any("must not access Outlook folders" in error for error in errors), errors)
        self.assertTrue(any("must not create Outlook.Application" in error for error in errors), errors)

    def test_csv_headers_are_present(self) -> None:
        self.assertEqual(validate_scaffold.validate_csv_headers(ROOT), [])

    def test_manifest_template_paths_resolve(self) -> None:
        self.assertEqual(
            validate_scaffold.validate_manifest_template_paths(
                ROOT, "examples/workbooks/induction_template_manifest_example.csv"
            ),
            [],
        )

    def test_induction_manifest_assets_resolve(self) -> None:
        self.assertEqual(
            validate_scaffold.validate_induction_template_manifest_assets(
                ROOT, "examples/workbooks/induction_template_manifest_example.csv"
            ),
            [],
        )
        self.assertEqual(
            validate_scaffold.validate_induction_template_manifest_assets(
                ROOT, "vbs/tests/fixtures/induction_template_manifest_fixture.csv"
            ),
            [],
        )

    def test_workflow_input_examples_are_validated(self) -> None:
        self.assertEqual(validate_scaffold.workflow_input_report.validate_default_examples(ROOT), [])

    def test_quickstart_routes_fixture_is_valid(self) -> None:
        self.assertEqual(validate_scaffold.command_route_report.validate_quickstart_fixture(ROOT), [])

    def test_scaffold_command_routes_are_validated(self) -> None:
        self.assertEqual(validate_scaffold.command_route_report.validate_default_routes(ROOT), [])

    def test_full_validator_passes(self) -> None:
        self.assertEqual(validate_scaffold.validate(ROOT), [])

    def test_generated_artifacts_are_absent(self) -> None:
        self.assertEqual(validate_scaffold.validate_generated_artifacts(ROOT), [])

    def test_sanitisation_config_allows_only_example_domain(self) -> None:
        self.assertEqual(validate_scaffold.allowed_email_domains(ROOT), {"example.invalid"})

    def test_legacy_alternate_map_is_valid(self) -> None:
        self.assertEqual(validate_scaffold.validate_legacy_alternate_map(ROOT), [])

    def test_legacy_mirror_is_excluded_from_public_validation(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "docs").mkdir()
            (root / "docs" / "public.md").write_text("public scaffold text", encoding="utf-8")
            (root / "legacy").mkdir()
            private_text = "private" + "@" + "contoso.invalid"
            private_text = private_text + " C:/Users/Private/" + "OneDrive" + " - Example"
            (root / "legacy" / "private.md").write_text(
                private_text, encoding="utf-8"
            )
            (root / "legacy" / "bundle.zip").write_bytes(b"archive")

            public_paths = [path.relative_to(root).as_posix() for path in validate_scaffold.iter_public_text_files(root)]

            self.assertEqual(public_paths, ["docs/public.md"])
            self.assertTrue(validate_scaffold.is_release_excluded_path(root / "legacy" / "private.md", root))
            self.assertEqual(validate_scaffold.validate_generated_artifacts(root), [])


if __name__ == "__main__":
    unittest.main()
