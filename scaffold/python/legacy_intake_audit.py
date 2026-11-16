"""Read-only audit for legacy-to-scaffold alternate mappings.

This tool inspects frozen legacy/reference source paths without copying,
rewriting, or exposing source content. It reports whether scaffold-local
alternates exist and whether a source would require sanitisation before any
future copy/port work.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import validate_scaffold


SCHEMA = "office-communication-scaffold.legacy-intake-audit.v1"
SOURCE_SCAN_TEXT_SUFFIXES = {".bas", ".cls", ".csv", ".htm", ".html", ".json", ".md", ".py", ".txt", ".vba", ".vbs"}


def load_map(scaffold_root: Path) -> dict[str, Any]:
    data = validate_scaffold.load_json(scaffold_root / "migration" / "legacy_alternate_map.json")
    if not isinstance(data, dict):
        return {}
    return data


def safe_relative(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return path.name


def has_blocked_directory(path: Path, scaffold_root: Path) -> bool:
    blocked_names = {
        str(item).lower()
        for item in validate_scaffold.sanitisation_config(scaffold_root).get("blocked_generated_directory_names", [])
    } or set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_DIRECTORY_NAMES)
    return any(part.lower() in blocked_names for part in path.parts)


def has_blocked_suffix(path: Path, scaffold_root: Path) -> bool:
    blocked_suffixes = {
        str(item).lower()
        for item in validate_scaffold.sanitisation_config(scaffold_root).get("blocked_generated_suffixes", [])
    } or set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_SUFFIXES)
    return path.suffix.lower() in blocked_suffixes


def text_signal_counts(path: Path, scaffold_root: Path) -> dict[str, int]:
    counts = {
        "blocked_text_patterns": 0,
        "non_example_email_addresses": 0,
    }
    if not path.is_file() or path.suffix.lower() not in SOURCE_SCAN_TEXT_SUFFIXES:
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


def source_signals(path: Path, scaffold_root: Path) -> dict[str, Any]:
    signals: dict[str, Any] = {
        "blocked_generated_suffix": has_blocked_suffix(path, scaffold_root),
        "blocked_generated_directory": has_blocked_directory(path, scaffold_root),
        "blocked_text_patterns": 0,
        "non_example_email_addresses": 0,
    }
    if path.exists():
        signals.update(text_signal_counts(path, scaffold_root))
    signals["requires_sanitisation"] = any(
        [
            signals["blocked_generated_suffix"],
            signals["blocked_generated_directory"],
            signals["blocked_text_patterns"] > 0,
            signals["non_example_email_addresses"] > 0,
        ]
    )
    return signals


def entry_report(entry: dict[str, Any], repo_root: Path, scaffold_root: Path) -> dict[str, Any]:
    legacy_path = str(entry.get("legacy_path", ""))
    scaffold_path = str(entry.get("scaffold_path", ""))
    route_path = entry.get("route_path")

    source = repo_root / legacy_path
    target = scaffold_root / scaffold_path
    route = scaffold_root / str(route_path) if route_path else None
    signals = source_signals(source, scaffold_root)

    target_signals = source_signals(target, scaffold_root)
    return {
        "legacy_path": legacy_path,
        "scaffold_path": scaffold_path,
        "route_path": str(route_path) if route_path else "",
        "status": str(entry.get("status", "")),
        "source_exists": source.exists(),
        "scaffold_target_exists": target.exists(),
        "route_target_exists": bool(route.exists()) if route else True,
        "source_requires_sanitisation_before_copy": bool(signals["requires_sanitisation"]),
        "source_signal_counts": {
            "blocked_text_patterns": int(signals["blocked_text_patterns"]),
            "non_example_email_addresses": int(signals["non_example_email_addresses"]),
        },
        "source_is_generated_or_archive": bool(signals["blocked_generated_suffix"] or signals["blocked_generated_directory"]),
        "scaffold_target_requires_sanitisation": bool(target_signals["requires_sanitisation"]),
        "copy_decision": "port_or_sanitise_only",
    }


def audit(repo_root: Path, scaffold_root: Path) -> dict[str, Any]:
    repo_root = repo_root.resolve()
    scaffold_root = scaffold_root.resolve()
    map_errors = validate_scaffold.validate_legacy_alternate_map(scaffold_root)
    data = load_map(scaffold_root)
    entries = data.get("entries", []) if isinstance(data.get("entries"), list) else []
    reports = [
        entry_report(entry, repo_root, scaffold_root)
        for entry in entries
        if isinstance(entry, dict)
    ]

    source_missing = [item["legacy_path"] for item in reports if not item["source_exists"]]
    scaffold_missing = [item["scaffold_path"] for item in reports if not item["scaffold_target_exists"]]
    route_missing = [item["route_path"] for item in reports if item["route_path"] and not item["route_target_exists"]]
    scaffold_sanitisation = [
        item["scaffold_path"] for item in reports if item["scaffold_target_requires_sanitisation"]
    ]
    source_sanitisation = [
        item["legacy_path"] for item in reports if item["source_requires_sanitisation_before_copy"]
    ]

    errors = {
        "map": map_errors,
        "source_missing": source_missing,
        "scaffold_target_missing": scaffold_missing,
        "route_target_missing": route_missing,
        "scaffold_target_sanitisation": scaffold_sanitisation,
    }

    return {
        "schema": SCHEMA,
        "repo_root": ".",
        "scaffold_root": safe_relative(scaffold_root, repo_root),
        "ready": not any(errors.values()),
        "raw_copy_allowed": False,
        "entries_count": len(reports),
        "source_present_count": len(reports) - len(source_missing),
        "scaffold_target_present_count": len(reports) - len(scaffold_missing),
        "route_target_present_count": len([item for item in reports if item["route_target_exists"]]),
        "sources_requiring_sanitisation_count": len(source_sanitisation),
        "errors": errors,
        "source_sanitisation_required": source_sanitisation,
        "entries": reports,
    }


def print_text_report(result: dict[str, Any]) -> None:
    print(f"Legacy intake audit: {'PASSED' if result['ready'] else 'FAILED'}")
    print(f"schema: {result['schema']}")
    print(f"repo_root: {result['repo_root']}")
    print(f"scaffold_root: {result['scaffold_root']}")
    print(f"raw_copy_allowed: {str(result['raw_copy_allowed']).lower()}")
    print(f"entries: {result['entries_count']}")
    print(f"source_present: {result['source_present_count']}")
    print(f"scaffold_targets_present: {result['scaffold_target_present_count']}")
    print(f"routes_present: {result['route_target_present_count']}")
    print(f"sources_requiring_sanitisation: {result['sources_requiring_sanitisation_count']}")
    if not result["ready"]:
        print("errors:")
        for group, errors in result["errors"].items():
            for error in errors:
                print(f"- {group}: {error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit legacy source intake into scaffold-local alternates.")
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--scaffold-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    result = audit(args.repo_root, args.scaffold_root)
    if args.format == "json":
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        print_text_report(result)
    return 0 if result["ready"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
