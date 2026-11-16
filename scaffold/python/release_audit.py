"""Stage 7 release-readiness audit for the scaffold.

This is a development-side gate. Operational use remains VBS-first; this
script summarizes the scaffold-only checks that should pass before any public
release work happens outside this folder.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import validate_scaffold
import release_manifest
import legacy_intake_audit
import legacy_mirror_audit
import legacy_sanitisation_plan
import repo_boundary_audit


STAGE = "Stage 7 - public sanitisation and release readiness"

REQUIRED_RELEASE_FILES = [
    ".gitignore",
    "LICENSE",
    "README.md",
    "RunGroupMeetings.vbs",
    "RunInductionMailer.vbs",
    "OutlookWatchdog.vbs",
    "ValidateWorkflow.vbs",
    "docs/PUBLIC_RELEASE_CHECKLIST.md",
    "docs/PUBLICATION_BOUNDARY.md",
    "docs/RELEASE_AUDIT.md",
    "docs/RELEASE_CANDIDATE_PLAN.md",
    "docs/RELEASE_EXPORT_PLAN.md",
    "docs/RELEASE_GATE.md",
    "docs/LEGACY_ALTERNATE_MAP.md",
    "docs/LEGACY_MIRROR.md",
    "docs/REPO_BOUNDARY_AUDIT.md",
    "docs/RELEASE_MANIFEST.md",
    "docs/RELEASE_NOTES_TEMPLATE.md",
    "docs/VALIDATION_RUNBOOK.md",
    "fixtures/sanitisation/blocked_patterns.json",
    "migration/legacy_alternate_map.json",
    "python/validate_scaffold.py",
    "python/release_audit.py",
    "python/release_candidate_plan.py",
    "python/release_export_plan.py",
    "python/release_gate.py",
    "python/release_manifest.py",
    "python/legacy_intake_audit.py",
    "python/legacy_mirror_audit.py",
    "python/legacy_sanitisation_plan.py",
    "python/repo_boundary_audit.py",
    "python/tests/test_validate_scaffold.py",
    "python/tests/test_release_audit.py",
    "python/tests/test_release_candidate_plan.py",
    "python/tests/test_release_export_plan.py",
    "python/tests/test_release_gate.py",
    "python/tests/test_release_manifest.py",
    "python/tests/test_validation_runbook.py",
    "python/tests/test_legacy_intake_audit.py",
    "python/tests/test_legacy_mirror_audit.py",
    "python/tests/test_legacy_sanitisation_plan.py",
    "python/tests/test_repo_boundary_audit.py",
    "vbs/tests/ValidateScaffold.vbs",
    "vbs/tests/PublicReleaseGateSmokeTest.vbs",
    "vbs/tests/ReleaseManifestSmokeTest.vbs",
    "vbs/tests/RunAllSmokeTests.vbs",
    "vbs/tests/ScaffoldAlternateEntrypointsSmokeTest.vbs",
    "vbs/entrypoints/scaffold_launcher.vbs",
]

WORKFLOW_SAFETY_TARGETS = [
    "vbs/workflows",
    "vbs/entrypoints",
]

ALTERNATE_ENTRYPOINT_FILES = [
    "RunGroupMeetings.vbs",
    "RunInductionMailer.vbs",
    "OutlookWatchdog.vbs",
    "ValidateWorkflow.vbs",
]

WORKFLOW_FORBIDDEN_TOKENS = [
    'CreateObject("Outlook.Application")',
    'CreateObject("Excel.Application")',
    "GetDefaultFolder",
    ".Send",
    ".Delete",
]

WATCHDOG_REPORT_ONLY_FORBIDDEN_TOKENS = [
    'CreateObject("Outlook.Application")',
    'CreateObject("Excel.Application")',
    "GetDefaultFolder",
    ".Send",
    ".Delete",
]

WATCHDOG_REPAIR_FORBIDDEN_TOKENS = [
    'CreateObject("Outlook.Application")',
    'CreateObject("Excel.Application")',
    "GetDefaultFolder",
    ".Send",
]

RUNTIME_LEGACY_MIRROR_FORBIDDEN_TOKENS = [
    "legacy/repo_snapshot",
    "legacy\\repo_snapshot",
]


def relative(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def existing_files(root: Path, relative_paths: list[str]) -> list[str]:
    return [path for path in relative_paths if (root / path).is_file()]


def missing_files(root: Path, relative_paths: list[str]) -> list[str]:
    return [path for path in relative_paths if not (root / path).is_file()]


def count_files(root: Path, pattern: str) -> int:
    return sum(1 for path in root.rglob(pattern) if path.is_file())


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def contains_token(path: Path, token: str) -> bool:
    return token.lower() in read_text(path).lower()


def validate_workflow_safety(root: Path) -> list[str]:
    errors: list[str] = []
    for relative_dir in WORKFLOW_SAFETY_TARGETS:
        directory = root / relative_dir
        if not directory.exists():
            errors.append(f"missing safety scan directory: {relative_dir}")
            continue
        for path in sorted(directory.glob("*.vbs")):
            for token in WORKFLOW_FORBIDDEN_TOKENS:
                if contains_token(path, token):
                    errors.append(f"{relative(path, root)} owns forbidden low-level token: {token}")

    for relative_path in ALTERNATE_ENTRYPOINT_FILES:
        path = root / relative_path
        if not path.is_file():
            errors.append(f"missing alternate entrypoint safety target: {relative_path}")
            continue
        for token in WORKFLOW_FORBIDDEN_TOKENS:
            if contains_token(path, token):
                errors.append(f"{relative(path, root)} owns forbidden low-level token: {token}")

    sidecar = root / "vbs" / "watchdog" / "outlook_watchdog_sidecar.vbs"
    if sidecar.is_file():
        for token in WATCHDOG_REPORT_ONLY_FORBIDDEN_TOKENS:
            if contains_token(sidecar, token):
                errors.append(f"{relative(sidecar, root)} is report-only but contains: {token}")
    else:
        errors.append("missing watchdog report-only sidecar")

    repair = root / "vbs" / "watchdog" / "managed_outbox_repair.vbs"
    if repair.is_file():
        for token in WATCHDOG_REPAIR_FORBIDDEN_TOKENS:
            if contains_token(repair, token):
                errors.append(f"{relative(repair, root)} contains forbidden repair token: {token}")
        if not contains_token(repair, ".Delete"):
            errors.append("managed repair module no longer contains its guarded delete path")
    else:
        errors.append("missing managed outbox repair module")

    return errors


def validate_legacy_mirror_runtime_isolation(root: Path) -> list[str]:
    errors: list[str] = []
    runtime_targets = [
        root / "RunGroupMeetings.vbs",
        root / "RunInductionMailer.vbs",
        root / "OutlookWatchdog.vbs",
        root / "ValidateWorkflow.vbs",
        root / "vbs" / "core",
        root / "vbs" / "entrypoints",
        root / "vbs" / "watchdog",
        root / "vbs" / "workflows",
    ]

    for target in runtime_targets:
        if target.is_file():
            paths = [target]
        elif target.is_dir():
            paths = sorted(target.glob("*.vbs"))
        else:
            continue

        for path in paths:
            for token in RUNTIME_LEGACY_MIRROR_FORBIDDEN_TOKENS:
                if contains_token(path, token):
                    errors.append(f"{relative(path, root)} references private legacy mirror: {token}")

    return errors


def collect_counts(root: Path) -> dict[str, int]:
    return {
        "files": len(release_manifest.iter_release_files(root)),
        "docs": count_files(root / "docs", "*.md"),
        "vbs_core_modules": count_files(root / "vbs" / "core", "*.vbs"),
        "vbs_workflow_modules": count_files(root / "vbs" / "workflows", "*.vbs"),
        "vbs_entrypoints": count_files(root / "vbs" / "entrypoints", "*.vbs"),
        "vbs_tests": count_files(root / "vbs" / "tests", "*.vbs"),
        "python_modules": count_files(root / "python", "*.py"),
        "example_workbooks": count_files(root / "examples" / "workbooks", "*.csv"),
        "example_templates": count_files(root / "examples" / "templates", "*.html"),
    }


def audit(root: Path) -> dict[str, Any]:
    root = root.resolve()
    validation_errors = validate_scaffold.validate(root)
    missing_release_files = missing_files(root, REQUIRED_RELEASE_FILES)
    safety_errors = validate_workflow_safety(root)
    legacy_runtime_errors = validate_legacy_mirror_runtime_isolation(root)
    manifest = release_manifest.build_manifest(root)
    manifest_errors = release_manifest.validate_manifest(manifest)
    intake = legacy_intake_audit.audit(root.parent, root)
    mirror = legacy_mirror_audit.audit(root)
    mirror_plan = legacy_sanitisation_plan.build_plan(root)
    boundary_tool_schema = repo_boundary_audit.SCHEMA

    errors = {
        "scaffold_validation": validation_errors,
        "required_release_files": missing_release_files,
        "workflow_safety": safety_errors,
        "legacy_runtime_isolation": legacy_runtime_errors,
        "release_manifest": manifest_errors,
        "legacy_intake": [] if intake["ready"] else [
            f"{group}: {error}"
            for group, group_errors in intake["errors"].items()
            for error in group_errors
        ],
        "legacy_mirror": mirror["guardrail_errors"],
        "legacy_sanitisation_plan": mirror_plan["guardrail_errors"],
        "repo_boundary_tool": [],
    }

    return {
        "stage": STAGE,
        "scaffold_root": str(root),
        "ready": not any(errors.values()),
        "checks": {
            "scaffold_validation": len(validation_errors) == 0,
            "required_release_files": len(missing_release_files) == 0,
            "workflow_safety": len(safety_errors) == 0,
            "legacy_runtime_isolation": len(legacy_runtime_errors) == 0,
            "release_manifest": len(manifest_errors) == 0,
            "legacy_intake": intake["ready"],
            "legacy_mirror": mirror["quarantine_ready"] and mirror["schema"] == legacy_mirror_audit.SCHEMA,
            "legacy_sanitisation_plan": (
                mirror_plan["quarantine_ready"] and mirror_plan["schema"] == legacy_sanitisation_plan.SCHEMA
            ),
            "repo_boundary_tool": boundary_tool_schema == repo_boundary_audit.SCHEMA,
        },
        "errors": errors,
        "counts": collect_counts(root),
        "manifest": {
            "schema": manifest["schema"],
            "file_count": manifest["file_count"],
            "excluded_release_directories": manifest["excluded_release_directories"],
            "excluded_generated_artifacts_count": len(manifest["excluded_generated_artifacts"]),
        },
        "legacy_intake": {
            "schema": intake["schema"],
            "entries_count": intake["entries_count"],
            "source_present_count": intake["source_present_count"],
            "scaffold_target_present_count": intake["scaffold_target_present_count"],
            "sources_requiring_sanitisation_count": intake["sources_requiring_sanitisation_count"],
            "raw_copy_allowed": intake["raw_copy_allowed"],
        },
        "legacy_mirror": {
            "schema": mirror["schema"],
            "mirror_present": mirror["mirror_present"],
            "quarantine_ready": mirror["quarantine_ready"],
            "requires_sanitisation": mirror["requires_sanitisation"],
            "report_only": mirror["report_only"],
            "mutations_performed": mirror["mutations_performed"],
            "findings_count": mirror["findings_count"],
            "summary": mirror["summary"],
        },
        "legacy_sanitisation_plan": {
            "schema": mirror_plan["schema"],
            "quarantine_ready": mirror_plan["quarantine_ready"],
            "requires_sanitisation": mirror_plan["requires_sanitisation"],
            "report_only": mirror_plan["report_only"],
            "mutations_performed": mirror_plan["mutations_performed"],
            "plan_item_count": mirror_plan["plan_item_count"],
            "action_summary": mirror_plan["action_summary"],
        },
        "repo_boundary_tool": {
            "schema": boundary_tool_schema,
            "report_only_by_default": True,
            "included_in_scaffold_release_gate": False,
            "used_by_publication_gate": True,
            "legacy_mirror_excluded_from_release": True,
        },
        "required_release_files": existing_files(root, REQUIRED_RELEASE_FILES),
    }


def print_text_report(result: dict[str, Any]) -> None:
    status = "PASSED" if result["ready"] else "FAILED"
    print(f"Release audit: {status}")
    print(f"stage: {result['stage']}")
    print(f"scaffold_root: {result['scaffold_root']}")
    print("checks:")
    for name, passed in result["checks"].items():
        print(f"- {name}: {'OK' if passed else 'FAILED'}")
    print("counts:")
    for name, value in result["counts"].items():
        print(f"- {name}: {value}")
    print("manifest:")
    for name, value in result["manifest"].items():
        print(f"- {name}: {value}")
    print("legacy_intake:")
    for name, value in result["legacy_intake"].items():
        print(f"- {name}: {value}")
    print("legacy_mirror:")
    for name, value in result["legacy_mirror"].items():
        print(f"- {name}: {value}")
    print("legacy_sanitisation_plan:")
    for name, value in result["legacy_sanitisation_plan"].items():
        print(f"- {name}: {value}")
    print("repo_boundary_tool:")
    for name, value in result["repo_boundary_tool"].items():
        print(f"- {name}: {value}")

    if not result["ready"]:
        print("errors:")
        for group, errors in result["errors"].items():
            for error in errors:
                print(f"- {group}: {error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit scaffold public-release readiness.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    result = audit(args.root)
    if args.format == "json":
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        print_text_report(result)
    return 0 if result["ready"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
