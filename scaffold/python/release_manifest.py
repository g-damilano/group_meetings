"""Deterministic scaffold release manifest generator.

The manifest is printed to stdout and contains only relative scaffold paths.
It does not package, publish, tag, or write artifacts by default.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

import validate_scaffold


SCHEMA = "office-communication-scaffold.release-manifest.v1"


def relative(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def blocked_generated_suffixes(root: Path) -> set[str]:
    config = validate_scaffold.sanitisation_config(root)
    suffixes = config.get("blocked_generated_suffixes", [])
    if isinstance(suffixes, list) and suffixes:
        return {str(item).lower() for item in suffixes}
    return set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_SUFFIXES)


def blocked_generated_directory_names(root: Path) -> set[str]:
    config = validate_scaffold.sanitisation_config(root)
    directory_names = config.get("blocked_generated_directory_names", [])
    if isinstance(directory_names, list) and directory_names:
        return {str(item).lower() for item in directory_names}
    return set(validate_scaffold.DEFAULT_BLOCKED_GENERATED_DIRECTORY_NAMES)


def is_generated_path(path: Path, root: Path) -> bool:
    relative_parts = path.relative_to(root).parts
    blocked_dirs = blocked_generated_directory_names(root)
    if any(part.lower() in blocked_dirs for part in relative_parts):
        return True
    return path.is_file() and path.suffix.lower() in blocked_generated_suffixes(root)


def is_release_excluded_path(path: Path, root: Path) -> bool:
    return validate_scaffold.is_release_excluded_path(path, root)


def file_kind(relative_path: str) -> str:
    if relative_path == ".gitignore":
        return "release-control"
    if relative_path == "LICENSE":
        return "license"
    if relative_path.startswith("docs/"):
        return "documentation"
    if relative_path.startswith("contracts/"):
        return "contract"
    if relative_path.startswith("fixtures/"):
        return "fixture"
    if relative_path.startswith("migration/"):
        return "migration"
    if relative_path.startswith("examples/"):
        return "example"
    if relative_path.startswith("python/tests/"):
        return "python-test"
    if relative_path.startswith("python/"):
        return "python-tool"
    if relative_path.startswith("vbs/tests/fixtures/"):
        return "test-fixture"
    if relative_path.startswith("vbs/tests/"):
        return "vbs-test"
    if relative_path.startswith("vbs/core/"):
        return "vbs-core"
    if relative_path.startswith("vbs/workflows/"):
        return "vbs-workflow"
    if relative_path.startswith("vbs/entrypoints/"):
        return "vbs-entrypoint"
    if relative_path.startswith("vbs/watchdog/"):
        return "vbs-watchdog"
    return "scaffold"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


def iter_release_files(root: Path) -> list[Path]:
    root = root.resolve()
    files = []
    for path in root.rglob("*"):
        if is_release_excluded_path(path, root):
            continue
        if path.is_file() and not is_generated_path(path, root):
            files.append(path)
    return sorted(files, key=lambda item: relative(item, root))


def excluded_generated_artifacts(root: Path) -> list[str]:
    root = root.resolve()
    excluded = []
    for path in root.rglob("*"):
        if is_release_excluded_path(path, root):
            continue
        if is_generated_path(path, root):
            excluded.append(relative(path, root))
    return sorted(excluded)


def build_manifest(root: Path) -> dict[str, Any]:
    root = root.resolve()
    files = []
    for path in iter_release_files(root):
        relative_path = relative(path, root)
        files.append(
            {
                "path": relative_path,
                "kind": file_kind(relative_path),
                "bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )

    return {
        "schema": SCHEMA,
        "root": "scaffold",
        "file_count": len(files),
        "files": files,
        "excluded_release_directories": sorted(validate_scaffold.release_excluded_directory_names(root)),
        "excluded_generated_artifacts": excluded_generated_artifacts(root),
    }


def validate_manifest(manifest: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    paths = []
    for item in manifest.get("files", []):
        path = str(item.get("path", ""))
        paths.append(path)
        if not path or path.startswith("/") or ":" in path or "\\" in path:
            errors.append(f"manifest path is not release-relative: {path}")
        if len(str(item.get("sha256", ""))) != 64:
            errors.append(f"manifest path has invalid sha256: {path}")
        if int(item.get("bytes", -1)) < 0:
            errors.append(f"manifest path has invalid byte count: {path}")

    if len(paths) != len(set(paths)):
        errors.append("manifest contains duplicate paths")
    if paths != sorted(paths):
        errors.append("manifest paths are not sorted")
    if manifest.get("schema") != SCHEMA:
        errors.append("manifest schema is unknown")
    if manifest.get("root") != "scaffold":
        errors.append("manifest root is not scaffold")
    if manifest.get("file_count") != len(paths):
        errors.append("manifest file_count does not match file entries")
    excluded_dirs = {
        str(item).strip("/\\").lower()
        for item in manifest.get("excluded_release_directories", [])
        if str(item).strip("/\\")
    }
    for path in paths:
        first_part = path.split("/", 1)[0].lower()
        if first_part in excluded_dirs:
            errors.append(f"manifest includes release-excluded path: {path}")
    return errors


def print_text_manifest(manifest: dict[str, Any]) -> None:
    print(f"Release manifest: {manifest['schema']}")
    print(f"root: {manifest['root']}")
    print(f"file_count: {manifest['file_count']}")
    if manifest["excluded_generated_artifacts"]:
        print("excluded_generated_artifacts:")
        for path in manifest["excluded_generated_artifacts"]:
            print(f"- {path}")
    if manifest["excluded_release_directories"]:
        print("excluded_release_directories:")
        for path in manifest["excluded_release_directories"]:
            print(f"- {path}")
    print("files:")
    for item in manifest["files"]:
        print(f"- {item['path']} {item['bytes']} {item['sha256']} {item['kind']}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a deterministic scaffold release manifest.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    manifest = build_manifest(args.root)
    errors = validate_manifest(manifest)
    if errors:
        for error in errors:
            print(f"FAILED: {error}", file=sys.stderr)
        return 1

    if args.format == "json":
        print(json.dumps(manifest, indent=2, sort_keys=True))
    else:
        print_text_manifest(manifest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
