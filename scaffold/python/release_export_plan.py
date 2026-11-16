"""Deterministic public export plan for scaffold contents.

This is an inventory and review aid. It never copies files, creates archives,
publishes releases, tags commits, sends mail, or writes outside the private
``legacy/`` reporting area.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import legacy_mirror_audit
import release_manifest


SCHEMA = "office-communication-scaffold.release-export-plan.v1"


def summarize_kinds(files: list[dict[str, Any]]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for item in files:
        kind = str(item["kind"])
        summary[kind] = summary.get(kind, 0) + 1
    return dict(sorted(summary.items()))


def build_export_plan(root: Path) -> dict[str, Any]:
    manifest = release_manifest.build_manifest(root)
    files = manifest["files"]
    legacy_paths = [item["path"] for item in files if str(item["path"]).startswith("legacy/")]
    manifest_errors = release_manifest.validate_manifest(manifest)

    return {
        "schema": SCHEMA,
        "source_root": "scaffold",
        "export_scope": "public-scaffold-only",
        "release_manifest_schema": manifest["schema"],
        "report_only": True,
        "mutations_performed": False,
        "publishing_performed": False,
        "packaging_performed": False,
        "archive_created": False,
        "real_email_sending_performed": False,
        "legacy_paths_in_files": len(legacy_paths),
        "legacy_paths": legacy_paths,
        "file_count": manifest["file_count"],
        "total_bytes": sum(int(item["bytes"]) for item in files),
        "kind_summary": summarize_kinds(files),
        "excluded_release_directories": manifest["excluded_release_directories"],
        "excluded_generated_artifacts_count": len(manifest["excluded_generated_artifacts"]),
        "manifest_errors": manifest_errors,
        "files": files,
    }


def validate_export_plan(plan: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if plan.get("schema") != SCHEMA:
        errors.append("export plan schema is unknown")
    if plan.get("source_root") != "scaffold":
        errors.append("export source root must be scaffold")
    if plan.get("export_scope") != "public-scaffold-only":
        errors.append("export scope must be public-scaffold-only")
    for flag in [
        "report_only",
    ]:
        if plan.get(flag) is not True:
            errors.append(f"{flag} must be true")
    for flag in [
        "mutations_performed",
        "publishing_performed",
        "packaging_performed",
        "archive_created",
        "real_email_sending_performed",
    ]:
        if plan.get(flag) is not False:
            errors.append(f"{flag} must be false")
    if plan.get("legacy_paths_in_files") != 0:
        errors.append("export plan includes private legacy paths")
    if plan.get("manifest_errors"):
        errors.extend(f"manifest: {error}" for error in plan["manifest_errors"])
    paths = [str(item.get("path", "")) for item in plan.get("files", [])]
    if paths != sorted(paths):
        errors.append("export plan file paths are not sorted")
    if plan.get("file_count") != len(paths):
        errors.append("export plan file_count does not match files")
    return errors


def write_json_plan(root: Path, requested_path: Path, plan: dict[str, Any]) -> Path:
    target = legacy_mirror_audit.report_target_path(root, requested_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def print_text_plan(plan: dict[str, Any]) -> None:
    print(f"Release export plan: {plan['schema']}")
    print(f"source_root: {plan['source_root']}")
    print(f"export_scope: {plan['export_scope']}")
    print(f"report_only: {str(plan['report_only']).lower()}")
    print(f"publishing_performed: {str(plan['publishing_performed']).lower()}")
    print(f"packaging_performed: {str(plan['packaging_performed']).lower()}")
    print(f"archive_created: {str(plan['archive_created']).lower()}")
    print(f"file_count: {plan['file_count']}")
    print(f"total_bytes: {plan['total_bytes']}")
    print(f"legacy_paths_in_files: {plan['legacy_paths_in_files']}")
    print("kind_summary:")
    for kind, count in plan["kind_summary"].items():
        print(f"- {kind}: {count}")
    print("excluded_release_directories:")
    for path in plan["excluded_release_directories"]:
        print(f"- {path}")
    print("files:")
    for item in plan["files"]:
        print(f"- {item['path']} {item['bytes']} {item['sha256']} {item['kind']}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a deterministic public scaffold export plan.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--write-json", type=Path)
    parser.add_argument("--fail-if-invalid", action="store_true")
    args = parser.parse_args()

    plan = build_export_plan(args.root)
    errors = validate_export_plan(plan)
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
        if errors:
            print("errors:")
            for error in errors:
                print(f"- {error}")

    if args.fail_if_invalid and errors:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
