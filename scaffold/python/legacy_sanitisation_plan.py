"""Build a content-free sanitisation plan for the private legacy mirror.

The plan is a migration aid. It groups mirror audit findings by path and
suggests generic remediation actions without changing files or exposing matched
source text.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import legacy_mirror_audit


SCHEMA = "office-communication-scaffold.legacy-sanitisation-plan.v1"

ACTION_PRIORITY = [
    "remove_generated_cache",
    "remove_generated_artifact",
    "manual_binary_review",
    "rewrite_or_redact_text",
    "review",
]

KIND_ACTIONS = {
    "generated_directory": "remove_generated_cache",
    "inside_generated_directory": "remove_generated_cache",
    "blocked_generated_artifact": "remove_generated_artifact",
    "high_risk_binary_artifact": "manual_binary_review",
    "blocked_text_patterns": "rewrite_or_redact_text",
    "non_example_email_addresses": "rewrite_or_redact_text",
}


def action_for_kinds(kinds: set[str]) -> str:
    actions = {KIND_ACTIONS.get(kind, "review") for kind in kinds}
    for action in ACTION_PRIORITY:
        if action in actions:
            return action
    return "review"


def summarize_entries(entries: list[dict[str, Any]], field: str) -> dict[str, int]:
    summary: dict[str, int] = {}
    for entry in entries:
        key = str(entry.get(field, ""))
        if not key:
            key = "."
        summary[key] = summary.get(key, 0) + int(entry.get("finding_count", 0))
    return dict(sorted(summary.items()))


def aggregate_findings(findings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for finding in findings:
        path = str(finding.get("path", ""))
        if not path:
            continue
        entry = grouped.setdefault(
            path,
            {
                "path": path,
                "top_level": path.split("/", 1)[0],
                "finding_count": 0,
                "finding_kinds": set(),
            },
        )
        entry["finding_count"] += int(finding.get("count", 1))
        entry["finding_kinds"].add(str(finding.get("kind", "review")))

    entries: list[dict[str, Any]] = []
    for entry in grouped.values():
        kinds = set(entry["finding_kinds"])
        action = action_for_kinds(kinds)
        entries.append(
            {
                "path": entry["path"],
                "top_level": entry["top_level"],
                "finding_count": entry["finding_count"],
                "finding_kinds": sorted(kinds),
                "recommended_action": action,
                "public_release_allowed": False,
                "requires_manual_review": action in {"manual_binary_review", "rewrite_or_redact_text", "review"},
            }
        )
    return sorted(entries, key=lambda item: (str(item["recommended_action"]), str(item["path"])))


def build_plan(root: Path, mirror_root: Path | None = None) -> dict[str, Any]:
    root = root.resolve()
    mirror = legacy_mirror_audit.audit(root, mirror_root)
    entries = aggregate_findings(mirror["findings"])

    return {
        "schema": SCHEMA,
        "scaffold_root": str(root),
        "mirror_root": mirror["mirror_root"],
        "source_audit_schema": mirror["schema"],
        "quarantine_ready": mirror["quarantine_ready"],
        "requires_sanitisation": mirror["requires_sanitisation"],
        "report_only": True,
        "mutations_performed": False,
        "guardrail_errors": mirror["guardrail_errors"],
        "finding_count": mirror["findings_count"],
        "plan_item_count": len(entries),
        "action_summary": summarize_entries(entries, "recommended_action"),
        "top_level_summary": summarize_entries(entries, "top_level"),
        "entries": entries,
    }


def write_json_plan(root: Path, requested_path: Path, plan: dict[str, Any]) -> Path:
    target = legacy_mirror_audit.report_target_path(root, requested_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def print_text_plan(plan: dict[str, Any], limit: int) -> None:
    status = "READY" if plan["quarantine_ready"] else "FAILED"
    print(f"Legacy sanitisation plan: {status}")
    print(f"schema: {plan['schema']}")
    print(f"mirror_root: {plan['mirror_root']}")
    print(f"report_only: {str(plan['report_only']).lower()}")
    print(f"mutations_performed: {str(plan['mutations_performed']).lower()}")
    print(f"requires_sanitisation: {str(plan['requires_sanitisation']).lower()}")
    print(f"finding_count: {plan['finding_count']}")
    print(f"plan_item_count: {plan['plan_item_count']}")
    if plan["guardrail_errors"]:
        print("guardrail_errors:")
        for error in plan["guardrail_errors"]:
            print(f"- {error}")
    if plan["action_summary"]:
        print("action_summary:")
        for action, count in plan["action_summary"].items():
            print(f"- {action}: {count}")
    if plan["entries"]:
        print(f"entries_first_{limit}:")
        for entry in plan["entries"][:limit]:
            print(
                f"- {entry['path']}: {entry['recommended_action']} "
                f"({','.join(entry['finding_kinds'])}; {entry['finding_count']})"
            )


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a report-only sanitisation plan for the legacy mirror.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--mirror-root", type=Path)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--write-json", type=Path)
    parser.add_argument("--fail-on-guardrails", action="store_true")
    args = parser.parse_args()

    plan = build_plan(args.root, args.mirror_root)
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
        print_text_plan(plan, max(0, args.limit))
        if args.write_json:
            print(f"written_plan: {plan['written_plan']}")

    if args.fail_on_guardrails and not plan["quarantine_ready"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
