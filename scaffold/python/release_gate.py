"""Formal report-only publication gate for the scaffold.

The gate consumes the release candidate plan and converts it into an explicit
PASS/BLOCKED result. It does not package, publish, tag, send mail, or write
outside the private ``legacy/`` reporting area.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import legacy_mirror_audit
import release_candidate_plan


SCHEMA = "office-communication-scaffold.release-gate.v1"
GATE_PASS = "PASS"
GATE_BLOCKED = "BLOCKED"


def build_gate(root: Path, repo_root: Path | None = None) -> dict[str, Any]:
    plan = release_candidate_plan.build_plan(root, repo_root)
    gate = GATE_PASS if plan["decision"] == release_candidate_plan.DECISION_READY else GATE_BLOCKED

    return {
        "schema": SCHEMA,
        "gate": gate,
        "candidate_plan_schema": plan["schema"],
        "candidate_decision": plan["decision"],
        "report_only": True,
        "mutations_performed": False,
        "publishing_performed": False,
        "packaging_performed": False,
        "tagging_performed": False,
        "real_email_sending_performed": False,
        "scaffold_root": plan["scaffold_root"],
        "repo_root": plan["repo_root"],
        "scaffold_release_ready": plan["scaffold_release_ready"],
        "repo_boundary_ready": plan["repo_boundary_ready"],
        "blockers": plan["blockers"],
        "advisories": plan.get("advisories", []),
        "required_human_actions": plan["required_human_actions"],
        "advisory_actions": plan.get("advisory_actions", []),
        "forbidden_actions": plan["forbidden_actions"],
        "release_manifest": plan["release_manifest"],
        "repo_boundary": plan["repo_boundary"],
        "legacy_mirror": plan["legacy_mirror"],
        "legacy_sanitisation_plan": plan["legacy_sanitisation_plan"],
    }


def write_json_gate(root: Path, requested_path: Path, gate: dict[str, Any]) -> Path:
    target = legacy_mirror_audit.report_target_path(root, requested_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(gate, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def exit_code_for_gate(gate: dict[str, Any], fail_if_blocked: bool) -> int:
    if fail_if_blocked and gate["gate"] != GATE_PASS:
        return 1
    return 0


def print_text_gate(gate: dict[str, Any]) -> None:
    print(f"Release gate: {gate['gate']}")
    print(f"schema: {gate['schema']}")
    print(f"candidate_decision: {gate['candidate_decision']}")
    print(f"report_only: {str(gate['report_only']).lower()}")
    print(f"publishing_performed: {str(gate['publishing_performed']).lower()}")
    print(f"packaging_performed: {str(gate['packaging_performed']).lower()}")
    print(f"mutations_performed: {str(gate['mutations_performed']).lower()}")
    print(f"real_email_sending_performed: {str(gate['real_email_sending_performed']).lower()}")
    print(f"scaffold_release_ready: {str(gate['scaffold_release_ready']).lower()}")
    print(f"repo_boundary_ready: {str(gate['repo_boundary_ready']).lower()}")
    print("release_manifest:")
    for name, value in gate["release_manifest"].items():
        print(f"- {name}: {value}")
    if gate["blockers"]:
        print("blockers:")
        for blocker in gate["blockers"]:
            print(f"- {blocker}")
    if gate["advisories"]:
        print("advisories:")
        for advisory in gate["advisories"]:
            print(f"- {advisory}")
    print("required_human_actions:")
    for action in gate["required_human_actions"]:
        print(f"- {action}")
    if gate["advisory_actions"]:
        print("advisory_actions:")
        for action in gate["advisory_actions"]:
            print(f"- {action}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate the scaffold public-release gate.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--write-json", type=Path)
    parser.add_argument("--fail-if-blocked", action="store_true")
    args = parser.parse_args()

    gate = build_gate(args.root, args.repo_root)
    if args.write_json:
        try:
            target = write_json_gate(args.root, args.write_json, gate)
        except ValueError as exc:
            print(f"FAILED: {exc}", file=sys.stderr)
            return 2
        gate["written_gate"] = legacy_mirror_audit.relative(target, Path(args.root).resolve())

    if args.format == "json":
        print(json.dumps(gate, indent=2, sort_keys=True))
    else:
        print_text_gate(gate)
        if args.write_json:
            print(f"written_gate: {gate['written_gate']}")

    return exit_code_for_gate(gate, args.fail_if_blocked)


if __name__ == "__main__":
    raise SystemExit(main())
