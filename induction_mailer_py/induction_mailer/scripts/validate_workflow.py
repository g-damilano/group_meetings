from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Iterable

from ..contracts import DispatchState


DEFAULT_REQUIRED_SHEETS = ("Settings", "People", "Messages")


def repo_root_from_here() -> Path:
    return Path(__file__).resolve().parents[3]


def normalize_repo_root(root: Path) -> Path:
    resolved = root.resolve()
    states_path = resolved / "fixtures" / "dispatch_states" / "states.json"
    if states_path.exists():
        return resolved

    parent_states_path = resolved.parent / "fixtures" / "dispatch_states" / "states.json"
    if parent_states_path.exists():
        return resolved.parent

    return resolved


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def validate_dispatch_states(root: Path) -> list[str]:
    errors: list[str] = []
    fixture = load_json(root / "fixtures" / "dispatch_states" / "states.json")
    fixture_states = set(fixture.get("states", []))
    enum_states = {state.value for state in DispatchState}
    missing_from_fixture = sorted(enum_states - fixture_states)
    unknown_in_fixture = sorted(fixture_states - enum_states)
    if missing_from_fixture:
        errors.append(f"Dispatch fixture missing states: {', '.join(missing_from_fixture)}")
    if unknown_in_fixture:
        errors.append(f"Dispatch fixture has unknown states: {', '.join(unknown_in_fixture)}")
    return errors


def validate_contract_files(root: Path) -> list[str]:
    required = [
        root / "docs" / "ARCHITECTURE.md",
        root / "docs" / "CONTRACTS.md",
        root / "docs" / "DISPATCH_STATE_MACHINE.md",
        root / "docs" / "WATCHDOG_SAFETY_DOCTRINE.md",
        root / "contracts" / "workflow_schema.json",
        root / "fixtures" / "dispatch_states" / "states.json",
    ]
    return [f"Missing required contract file: {path}" for path in required if not path.exists()]


def validate_workbook(path: Path, required_sheets: Iterable[str]) -> list[str]:
    errors: list[str] = []
    try:
        from openpyxl import load_workbook
    except Exception as exc:
        return [f"openpyxl unavailable; cannot validate workbook: {exc}"]

    if not path.exists():
        return [f"Workbook not found: {path}"]

    wb = load_workbook(path, read_only=True, data_only=True)
    sheet_names = set(wb.sheetnames)
    for sheet in required_sheets:
        if sheet not in sheet_names:
            errors.append(f"Workbook missing required sheet: {sheet}")
    wb.close()
    return errors


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="validate-workflow-contracts",
        description="Validate shared Office communication automation contracts and optional workflow workbooks.",
    )
    parser.add_argument("--root", type=Path, default=None, help="Repository root. Defaults to auto-detected root.")
    parser.add_argument("--workbook", type=Path, default=None, help="Optional workflow workbook to validate.")
    parser.add_argument(
        "--required-sheet",
        action="append",
        dest="required_sheets",
        default=None,
        help="Required workbook sheet. Can be supplied more than once.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = normalize_repo_root(args.root or repo_root_from_here())
    errors: list[str] = []
    errors.extend(validate_contract_files(root))
    errors.extend(validate_dispatch_states(root))

    if args.workbook is not None:
        required = tuple(args.required_sheets or DEFAULT_REQUIRED_SHEETS)
        errors.extend(validate_workbook(args.workbook, required))

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1

    print("Workflow contracts validated.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
