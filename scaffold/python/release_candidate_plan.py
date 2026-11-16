"""Scaffold-only public release candidate planner.

This planner combines scaffold release evidence with advisory repository
boundary and legacy mirror status. Parent-repository findings are not scaffold
development blockers by default. It never publishes, packages, tags, sends
mail, or writes outside the private ``legacy/`` reporting area.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import legacy_mirror_audit
import release_audit
import release_manifest
import repo_boundary_audit


SCHEMA = "office-communication-scaffold.release-candidate-plan.v1"
DECISION_READY = "READY_FOR_AUTHORISED_RELEASE"
DECISION_BLOCKED = "BLOCKED_BY_REPO_BOUNDARY"
DECISION_SCAFFOLD_FAILED = "BLOCKED_BY_SCAFFOLD_VALIDATION"


def build_plan(root: Path, repo_root: Path | None = None) -> dict[str, Any]:
    root = root.resolve()
    repo_root = (repo_root or root.parent).resolve()

    scaffold = release_audit.audit(root)
    manifest = release_manifest.build_manifest(root)
    boundary = repo_boundary_audit.audit(repo_root, root)
    mirror = scaffold["legacy_mirror"]
    sanitisation = scaffold["legacy_sanitisation_plan"]

    legacy_paths_in_files = sum(1 for item in manifest["files"] if str(item["path"]).startswith("legacy/"))

    if not scaffold["ready"] or legacy_paths_in_files > 0:
        decision = DECISION_SCAFFOLD_FAILED
    else:
        decision = DECISION_READY

    blockers: list[str] = []
    advisories: list[str] = []
    if not scaffold["ready"]:
        blockers.append("scaffold_release_audit_failed")
    if not boundary["ready"]:
        advisories.append("repo_boundary_findings_present")
    if legacy_paths_in_files > 0:
        blockers.append("release_manifest_includes_private_legacy_mirror")

    return {
        "schema": SCHEMA,
        "decision": decision,
        "report_only": True,
        "mutations_performed": False,
        "publishing_performed": False,
        "real_email_sending_performed": False,
        "scaffold_root": "scaffold",
        "repo_root": ".",
        "scaffold_release_ready": scaffold["ready"],
        "repo_boundary_ready": boundary["ready"],
        "repo_boundary_advisory": True,
        "blockers": blockers,
        "advisories": advisories,
        "release_manifest": {
            "schema": manifest["schema"],
            "file_count": manifest["file_count"],
            "excluded_release_directories": manifest["excluded_release_directories"],
            "legacy_paths_in_files": legacy_paths_in_files,
            "excluded_generated_artifacts_count": len(manifest["excluded_generated_artifacts"]),
        },
        "repo_boundary": {
            "schema": boundary["schema"],
            "report_only": boundary["report_only"],
            "outside_root_edits_performed": boundary["outside_root_edits_performed"],
            "findings_count": boundary["findings_count"],
            "summary": boundary["summary"],
        },
        "legacy_mirror": {
            "schema": mirror["schema"],
            "mirror_present": mirror["mirror_present"],
            "quarantine_ready": mirror["quarantine_ready"],
            "requires_sanitisation": mirror["requires_sanitisation"],
            "findings_count": mirror["findings_count"],
        },
        "legacy_sanitisation_plan": {
            "schema": sanitisation["schema"],
            "quarantine_ready": sanitisation["quarantine_ready"],
            "requires_sanitisation": sanitisation["requires_sanitisation"],
            "plan_item_count": sanitisation["plan_item_count"],
            "action_summary": sanitisation["action_summary"],
        },
        "required_human_actions": required_human_actions(scaffold, boundary, sanitisation),
        "advisory_actions": advisory_actions(boundary, sanitisation),
        "forbidden_actions": [
            "publish_legacy_mirror",
            "edit_legacy_outside_scaffold",
            "perform_real_email_send",
            "modify_unmarked_mail",
            "run_destructive_mailbox_action",
        ],
    }


def required_human_actions(
    scaffold: dict[str, Any],
    boundary: dict[str, Any],
    sanitisation: dict[str, Any],
) -> list[str]:
    actions: list[str] = []
    if not scaffold["ready"]:
        actions.append("fix_scaffold_release_audit_inside_scaffold")
    if not actions:
        actions.append("perform_manual_public_release_review")
    return actions


def advisory_actions(
    boundary: dict[str, Any],
    sanitisation: dict[str, Any],
) -> list[str]:
    actions: list[str] = []
    if not boundary["ready"]:
        actions.append("repo_wide_public_cleanup_outside_scaffold_is_advisory_for_scaffold_development")
    if sanitisation["requires_sanitisation"]:
        actions.append("private_legacy_mirror_sanitisation_is_advisory_until_mirror_content_is_promoted")
    return actions


def write_json_plan(root: Path, requested_path: Path, plan: dict[str, Any]) -> Path:
    target = legacy_mirror_audit.report_target_path(root, requested_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def print_text_plan(plan: dict[str, Any]) -> None:
    print(f"Release candidate decision: {plan['decision']}")
    print(f"schema: {plan['schema']}")
    print(f"report_only: {str(plan['report_only']).lower()}")
    print(f"publishing_performed: {str(plan['publishing_performed']).lower()}")
    print(f"mutations_performed: {str(plan['mutations_performed']).lower()}")
    print(f"scaffold_release_ready: {str(plan['scaffold_release_ready']).lower()}")
    print(f"repo_boundary_ready: {str(plan['repo_boundary_ready']).lower()}")
    print(f"repo_boundary_advisory: {str(plan['repo_boundary_advisory']).lower()}")
    print("release_manifest:")
    for name, value in plan["release_manifest"].items():
        print(f"- {name}: {value}")
    print("repo_boundary:")
    for name, value in plan["repo_boundary"].items():
        print(f"- {name}: {value}")
    print("legacy_sanitisation_plan:")
    for name, value in plan["legacy_sanitisation_plan"].items():
        print(f"- {name}: {value}")
    print("required_human_actions:")
    for action in plan["required_human_actions"]:
        print(f"- {action}")
    if plan["advisories"]:
        print("advisories:")
        for advisory in plan["advisories"]:
            print(f"- {advisory}")
    if plan["advisory_actions"]:
        print("advisory_actions:")
        for action in plan["advisory_actions"]:
            print(f"- {action}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a scaffold-only release candidate decision plan.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--write-json", type=Path)
    parser.add_argument("--fail-if-blocked", action="store_true")
    args = parser.parse_args()

    plan = build_plan(args.root, args.repo_root)
    if args.write_json:
        try:
            target = write_json_plan(args.root, args.write_json, plan)
        except ValueError as exc:
            print(f"FAILED: {exc}", file=sys.stderr)
            return 2
        plan["written_plan"] = legacy_mirror_audit.relative(target, Path(args.root).resolve())

    if args.format == "json":
        print(json.dumps(plan, indent=2, sort_keys=True))
    else:
        print_text_plan(plan)
        if args.write_json:
            print(f"written_plan: {plan['written_plan']}")

    if args.fail_if_blocked and plan["decision"] != DECISION_READY:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
