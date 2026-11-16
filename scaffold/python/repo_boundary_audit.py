"""Read-only repository boundary audit for public-release preparation.

The scaffold can be release-ready while the surrounding repository still
contains frozen legacy/reference or generated material. This tool scans the
repository outside ``scaffold/`` without modifying it and reports release
blockers by path and category only.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Iterable

sys.dont_write_bytecode = True

import validate_scaffold


SCHEMA = "office-communication-scaffold.repo-boundary-audit.v1"
TEXT_SUFFIXES = validate_scaffold.TEXT_SUFFIXES | {".htm"}
SKIPPED_TOP_LEVEL_DIRS = {".git", "scaffold"}
HIGH_RISK_BINARY_SUFFIXES = {
    ".doc",
    ".docx",
    ".pdf",
    ".ppt",
    ".pptx",
    ".xls",
    ".xlsm",
    ".xlsx",
}


def is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def should_skip(path: Path, repo_root: Path, scaffold_root: Path) -> bool:
    if path == repo_root:
        return False
    if is_within(path, scaffold_root):
        return True
    try:
        parts = path.resolve().relative_to(repo_root.resolve()).parts
    except ValueError:
        return True
    return bool(parts and parts[0].lower() in SKIPPED_TOP_LEVEL_DIRS)


def iter_boundary_paths(repo_root: Path, scaffold_root: Path) -> Iterable[Path]:
    repo_root = repo_root.resolve()
    scaffold_root = scaffold_root.resolve()
    for path in repo_root.rglob("*"):
        if should_skip(path, repo_root, scaffold_root):
            continue
        yield path


def blocked_suffixes(scaffold_root: Path) -> set[str]:
    data = validate_scaffold.sanitisation_config(scaffold_root).get("blocked_generated_suffixes", [])
    suffixes = {str(item).lower() for item in data} if isinstance(data, list) else set()
    return suffixes or set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_SUFFIXES)


def blocked_directory_names(scaffold_root: Path) -> set[str]:
    data = validate_scaffold.sanitisation_config(scaffold_root).get("blocked_generated_directory_names", [])
    names = {str(item).lower() for item in data} if isinstance(data, list) else set()
    return names or set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_DIRECTORY_NAMES)


def text_counts(path: Path, scaffold_root: Path) -> dict[str, int]:
    counts = {
        "blocked_text_patterns": 0,
        "non_example_email_addresses": 0,
    }
    if not path.is_file() or path.suffix.lower() not in TEXT_SUFFIXES:
        return counts

    text = path.read_text(encoding="utf-8", errors="ignore")
    for pattern in validate_scaffold.blocked_text_regexes(scaffold_root):
        if pattern.search(text):
            counts["blocked_text_patterns"] += 1

    allowed_domains = validate_scaffold.allowed_email_domains(scaffold_root)
    for match in validate_scaffold.EMAIL_RE.finditer(text):
        if match.group(1).lower() not in allowed_domains:
            counts["non_example_email_addresses"] += 1

    return counts


def path_findings(path: Path, repo_root: Path, scaffold_root: Path) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    path_text = relative(path, repo_root)
    blocked_dirs = blocked_directory_names(scaffold_root)
    lower_parts = {part.lower() for part in path.parts}

    if path.is_dir() and path.name.lower() in blocked_dirs:
        findings.append({"path": path_text, "kind": "generated_directory", "count": 1})
    elif lower_parts & blocked_dirs:
        findings.append({"path": path_text, "kind": "inside_generated_directory", "count": 1})

    if path.is_file():
        suffix = path.suffix.lower()
        if suffix in blocked_suffixes(scaffold_root):
            findings.append({"path": path_text, "kind": "blocked_generated_artifact", "count": 1})
        if suffix in HIGH_RISK_BINARY_SUFFIXES:
            findings.append({"path": path_text, "kind": "high_risk_binary_artifact", "count": 1})

        counts = text_counts(path, scaffold_root)
        for kind, count in counts.items():
            if count > 0:
                findings.append({"path": path_text, "kind": kind, "count": count})

    return findings


def summarize(findings: list[dict[str, Any]]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for finding in findings:
        kind = str(finding["kind"])
        summary[kind] = summary.get(kind, 0) + int(finding.get("count", 1))
    return dict(sorted(summary.items()))


def audit(repo_root: Path, scaffold_root: Path) -> dict[str, Any]:
    repo_root = repo_root.resolve()
    scaffold_root = scaffold_root.resolve()
    findings: list[dict[str, Any]] = []
    scanned_paths = 0
    scanned_files = 0

    for path in iter_boundary_paths(repo_root, scaffold_root):
        scanned_paths += 1
        if path.is_file():
            scanned_files += 1
        findings.extend(path_findings(path, repo_root, scaffold_root))

    return {
        "schema": SCHEMA,
        "repo_root": ".",
        "scaffold_root": relative(scaffold_root, repo_root),
        "ready": len(findings) == 0,
        "report_only": True,
        "outside_root_edits_performed": False,
        "scanned_paths": scanned_paths,
        "scanned_files": scanned_files,
        "findings_count": len(findings),
        "summary": summarize(findings),
        "findings": sorted(findings, key=lambda item: (str(item["path"]), str(item["kind"]))),
    }


def print_text_report(result: dict[str, Any], limit: int) -> None:
    print(f"Repository boundary audit: {'PASSED' if result['ready'] else 'FINDINGS'}")
    print(f"schema: {result['schema']}")
    print(f"repo_root: {result['repo_root']}")
    print(f"scaffold_root: {result['scaffold_root']}")
    print(f"report_only: {str(result['report_only']).lower()}")
    print(f"outside_root_edits_performed: {str(result['outside_root_edits_performed']).lower()}")
    print(f"scanned_paths: {result['scanned_paths']}")
    print(f"scanned_files: {result['scanned_files']}")
    print(f"findings_count: {result['findings_count']}")
    if result["summary"]:
        print("summary:")
        for kind, count in result["summary"].items():
            print(f"- {kind}: {count}")
    if result["findings"]:
        print(f"findings_first_{limit}:")
        for finding in result["findings"][:limit]:
            print(f"- {finding['path']}: {finding['kind']} ({finding['count']})")


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit repository content outside scaffold for release blockers.")
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--scaffold-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--fail-on-findings", action="store_true")
    args = parser.parse_args()

    result = audit(args.repo_root, args.scaffold_root)
    if args.format == "json":
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        print_text_report(result, max(0, args.limit))

    if args.fail_on_findings and not result["ready"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
