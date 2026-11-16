from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path


PACKAGE_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PACKAGE_ROOT.parent
sys.path.insert(0, str(PACKAGE_ROOT))

from induction_mailer.contracts import DispatchState, build_managed_email_key  # noqa: E402
from induction_mailer.scripts.validate_workflow import validate_contract_files, validate_dispatch_states  # noqa: E402


class ContractTests(unittest.TestCase):
    def test_dispatch_state_fixture_matches_enum(self) -> None:
        self.assertEqual(validate_dispatch_states(REPO_ROOT), [])

    def test_required_contract_files_exist(self) -> None:
        self.assertEqual(validate_contract_files(REPO_ROOT), [])

    def test_managed_email_key_is_deterministic(self) -> None:
        key = build_managed_email_key("demo org", "meetings", "2026-05-29", "regular", "group list")
        self.assertEqual(key, "DEMO-ORG|MEETINGS|2026-05-29|REGULAR|GROUP-LIST")

    def test_fixture_contains_all_public_states(self) -> None:
        fixture = json.loads((REPO_ROOT / "fixtures" / "dispatch_states" / "states.json").read_text(encoding="utf-8"))
        self.assertEqual(set(fixture["states"]), {state.value for state in DispatchState})


if __name__ == "__main__":
    unittest.main()
