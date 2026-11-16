"""Report-only audit for the private legacy mirror under scaffold/legacy.

The mirror is a local migration aid, not public scaffold content. This tool
summarizes copied raw material by path and category only, without modifying the
mirror or exposing matched source text.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import repo_boundary_audit


SCHEMA = "office-communication-scaffold.legacy-mirror-audit.v1"
DEFAULT_MIRROR_RELATIVE_PATH = Path("legacy") / "repo_snapshot"
DEFAULT_MANIFEST_RELATIVE_PATH = Path("legacy") / "manifest.json"


def relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def load_manifest(root: Path) -> dict[str, Any]:
    manifest_path = root / DEFAULT_MANIFEST_RELATIVE_PATH
    if not manifest_path.is_file():
        return {}
    data = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    return data if isinstance(data, dict) else {}


def guardrail_errors(root: Path, mirror_root: Path, manifest: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    legacy_root = root / "legacy"

    if not is_within(mirror_root, legacy_root):
        errors.append("mirror root must stay under legacy/")

    if mirror_root.exists():
        if manifest.get("excluded_from_public_release") is not True:
            errors.append("legacy manifest must mark mirror excluded_from_public_release=true")
        if manifest.get("copy_mode") != "copy-only":
            errors.append("legacy manifest must keep copy_mode=copy-only")
        if manifest.get("moved_files") is not False:
            errors.append("legacy manifest must keep moved_files=false")

    return errors


def iter_mirror_paths(mirror_root: Path) -> list[Path]:
    if not mirror_root.exists():
        return []
    return sorted(mirror_root.rglob("*"), key=lambda item: item.as_posix().lower())


def top_level_summary(findings: list[dict[str, Any]]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for finding in findings:
        path = str(finding.get("path", ""))
        top_level = path.split("/", 1)[0] if path else "."
        summary[top_level] = summary.get(top_level, 0) + int(finding.get("count", 1))
    return dict(sorted(summary.items()))


def audit(root: Path, mirror_root: Path | None = None) -> dict[str, Any]:
    root = root.resolve()
    mirror_root = (mirror_root or (root / DEFAULT_MIRROR_RELATIVE_PATH)).resolve()
    manifest = load_manifest(root)
    findings: list[dict[str, Any]] = []
    scanned_paths = 0
    scanned_files = 0

    for path in iter_mirror_paths(mirror_root):
        scanned_paths += 1
        if path.is_file():
            scanned_files += 1
        findings.extend(repo_boundary_audit.path_findings(path, mirror_root, root))

    errors = guardrail_errors(root, mirror_root, manifest)
    findings = sorted(findings, key=lambda item: (str(item["path"]), str(item["kind"])))

    return {
        "schema": SCHEMA,
        "scaffold_root": str(root),
        "mirror_root": relative(mirror_root, root) if is_within(mirror_root, root) else str(mirror_root),
        "mirror_present": mirror_root.exists(),
        "manifest_present": bool(manifest),
        "manifest_file_count": manifest.get("file_count"),
        "manifest_directory_count": manifest.get("directory_count"),
        "excluded_from_public_release": manifest.get("excluded_from_public_release"),
        "copy_mode": manifest.get("copy_mode"),
        "moved_files": manifest.get("moved_files"),
        "report_only": True,
        "mutations_performed": False,
        "quarantine_ready": len(errors) == 0,
        "requires_sanitisation": len(findings) > 0,
        "guardrail_errors": errors,
        "scanned_paths": scanned_paths,
        "scanned_files": scanned_files,
        "findings_count": len(findings),
        "summary": repo_boundary_audit.summarize(findings),
        "top_level_summary": top_level_summary(findings),
        "findings": findings,
    }


def report_target_path(root: Path, requested_path: Path) -> Path:
    root = root.resolve()
    target = requested_path if requested_path.is_absolute() else root / requested_path
    target = target.resolve()
    legacy_root = (root / "legacy").resolve()
    if not is_within(target, legacy_root):
        raise ValueError("mirror audit reports may only be written under scaffold/legacy/")
    return target


def write_json_report(root: Path, requested_path: Path, result: dict[str, Any]) -> Path:
    target = report_target_path(root, requested_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def print_text_report(result: dict[str, Any], limit: int) -> None:
    status = "QUARANTINED" if result["quarantine_ready"] else "FAILED"
    print(f"Legacy mirror audit: {status}")
    print(f"schema: {result['schema']}")
    print(f"mirror_root: {result['mirror_root']}")
    print(f"mirror_present: {str(result['mirror_present']).lower()}")
    print(f"report_only: {str(result['report_only']).lower()}")
    print(f"mutations_performed: {str(result['mutations_performed']).lower()}")
    print(f"requires_sanitisation: {str(result['requires_sanitisation']).lower()}")
    print(f"scanned_paths: {result['scanned_paths']}")
    print(f"scanned_files: {result['scanned_files']}")
    print(f"findings_count: {result['findings_count']}")
    if result["guardrail_errors"]:
        print("guardrail_errors:")
        for error in result["guardrail_errors"]:
            print(f"- {error}")
    if result["summary"]:
        print("summary:")
        for kind, count in result["summary"].items():
            print(f"- {kind}: {count}")
    if result["findings"]:
        print(f"findings_first_{limit}:")
        for finding in result["findings"][:limit]:
            print(f"- {finding['path']}: {finding['kind']} ({finding['count']})")


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit the private scaffold legacy mirror in report-only mode.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--mirror-root", type=Path)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--write-json", type=Path)
    parser.add_argument("--fail-on-guardrails", action="store_true")
    parser.add_argument("--fail-on-findings", action="store_true")
    args = parser.parse_args()

    result = audit(args.root, args.mirror_root)
    if args.write_json:
        try:
            target = write_json_report(args.root, args.write_json, result)
        except ValueError as exc:
            print(f"FAILED: {exc}", file=sys.stderr)
            return 2
        result["written_report"] = relative(target, Path(args.root).resolve())

    if args.format == "json":
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        print_text_report(result, max(0, args.limit))
        if args.write_json:
            print(f"written_report: {result['written_report']}")

    if args.fail_on_guardrails and not result["quarantine_ready"]:
        return 1
    if args.fail_on_findings and result["requires_sanitisation"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
