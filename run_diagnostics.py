#!/usr/bin/env python3
"""Diagnostics checklist for the induction mailer repository."""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

from typing import Dict, List, Sequence, Tuple

BASE_DIR = Path(__file__).resolve().parent
RESULTS_DIR = BASE_DIR / "diagnostics_results"


@dataclass
class CheckResult:
    """Represents the outcome of an individual diagnostic check."""

    name: str
    status: str
    details: str

    def as_dict(self) -> Dict[str, str]:
        return {"name": self.name, "status": self.status, "details": self.details}


STATUS_PASS = "pass"
STATUS_FAIL = "fail"
STATUS_WARN = "warn"


def check_required_structure(base_dir: Path) -> CheckResult:
    """Ensure the key scripts and folders referenced by the workflow exist."""

    required_entries: Sequence[Tuple[str, str]] = (
        ("InductionMailer.vbs", "file"),
        ("utilities/mail_collections.vbs", "file"),
        ("utilities/mail_template_utils.vbs", "file"),
        ("utilities/mail_delivery_utils.vbs", "file"),
        ("utilities/mail_input_validation.vbs", "file"),
        ("utilities/name_resolution/MNS_text.vbs", "file"),
        ("utilities/name_resolution/MNS_directory.vbs", "file"),
        ("utilities/name_resolution/MNS_core.vbs", "file"),
        ("templates", "directory"),
        (
            "mail_utils_utf8_plus_dynamiclinks_ext/mail_utils_mailer_dynamiclinks.vbs",
            "file",
        ),
        ("links_map.txt", "file"),
    )

    missing: List[str] = []
    unexpected_types: List[str] = []

    for relative_path, expected_kind in required_entries:
        candidate = base_dir / relative_path
        if expected_kind == "file":
            if not candidate.is_file():
                missing.append(relative_path)
        elif expected_kind == "directory":
            if not candidate.is_dir():
                missing.append(relative_path)
        else:
            unexpected_types.append(f"{relative_path} (unknown expectation {expected_kind})")

    if unexpected_types:
        return CheckResult(
            name="Repository structure expectations",
            status=STATUS_FAIL,
            details=(
                "Internal error: unexpected expectation type(s): "
                + ", ".join(unexpected_types)
            ),
        )

    if missing:
        return CheckResult(
            name="Repository structure expectations",
            status=STATUS_FAIL,
            details="Missing required entries: " + ", ".join(sorted(missing)),
        )

    return CheckResult(
        name="Repository structure expectations",
        status=STATUS_PASS,
        details="All required scripts and directories are present.",
    )


def _read_text(path: Path) -> str | None:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return None


def check_template_headers(base_dir: Path) -> CheckResult:
    """Validate the metadata header blocks at the top of each template."""

    template_dir = base_dir / "templates"
    if not template_dir.exists():
        return CheckResult(
            name="Template header metadata",
            status=STATUS_FAIL,
            details="Template directory is missing.",
        )

    template_paths = [
        path
        for path in template_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in {".htm", ".html", ".txt"}
    ]

    if not template_paths:
        return CheckResult(
            name="Template header metadata",
            status=STATUS_WARN,
            details="No template files discovered for validation.",
        )

    header_regex = re.compile(r"<!--(.*?)-->", re.DOTALL)
    kv_regex = re.compile(r"^([A-Za-z0-9 _-]+)\s*:\s*(.*)$")
    recognised_keys = {
        "to",
        "cc",
        "bcc",
        "subject",
        "attachments",
        "stage",
        "code",
        "delivery",
    }

    missing_headers: List[str] = []
    malformed_lines: Dict[str, List[str]] = {}
    missing_metadata: Dict[str, List[str]] = {}

    for path in template_paths:
        text = _read_text(path)
        if text is None:
            missing_headers.append(f"{path.relative_to(base_dir)} (unreadable)")
            continue

        header_match = header_regex.search(text)
        if not header_match:
            missing_headers.append(str(path.relative_to(base_dir)))
            continue

        header_content = header_match.group(1)
        metadata: Dict[str, str] = {}
        malformed: List[str] = []

        for raw_line in header_content.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            kv_match = kv_regex.match(line)
            if not kv_match:
                malformed.append(line)
                continue
            key = kv_match.group(1).strip().lower()
            value = kv_match.group(2).strip()
            metadata[key] = value

        if malformed:
            malformed_lines[str(path.relative_to(base_dir))] = malformed

        if not recognised_keys.intersection(metadata):
            missing_metadata.setdefault(str(path.relative_to(base_dir)), []).append(
                "No recognised metadata keys (To/Cc/Bcc/Subject/Attachments/Stage/Code/Delivery)"
            )

    if missing_headers:
        return CheckResult(
            name="Template header metadata",
            status=STATUS_FAIL,
            details=(
                "Templates with missing or unreadable headers: "
                + "; ".join(sorted(missing_headers))
            ),
        )

    issues: List[str] = []
    if malformed_lines:
        for template, lines in sorted(malformed_lines.items()):
            issues.append(f"{template}: malformed header lines -> {', '.join(lines)}")
    if missing_metadata:
        for template, notes in sorted(missing_metadata.items()):
            issues.append(f"{template}: {', '.join(notes)}")

    if issues:
        return CheckResult(
            name="Template header metadata",
            status=STATUS_WARN,
            details="; ".join(issues),
        )

    return CheckResult(
        name="Template header metadata",
        status=STATUS_PASS,
        details=f"Validated {len(template_paths)} template header block(s).",
    )


def load_link_map(base_dir: Path) -> Tuple[CheckResult, Dict[str, str] | None]:
    link_map_path = base_dir / "links_map.txt"
    if not link_map_path.is_file():
        return (
            CheckResult(
                name="Dynamic link map",
                status=STATUS_FAIL,
                details="links_map.txt is missing.",
            ),
            None,
        )

    raw_text = _read_text(link_map_path)
    if raw_text is None:
        return (
            CheckResult(
                name="Dynamic link map",
                status=STATUS_FAIL,
                details="links_map.txt could not be read.",
            ),
            None,
        )

    entries: Dict[str, str] = {}
    invalid_lines: List[str] = []
    for line in raw_text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if "=" not in stripped:
            invalid_lines.append(stripped)
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip()
        if not key or not value:
            invalid_lines.append(stripped)
            continue
        entries[key] = value

    if invalid_lines:
        return (
            CheckResult(
                name="Dynamic link map",
                status=STATUS_WARN,
                details="Invalid line(s): " + "; ".join(invalid_lines),
            ),
            entries,
        )

    return (
        CheckResult(
            name="Dynamic link map",
            status=STATUS_PASS,
            details=f"Loaded {len(entries)} dynamic link substitution(s).",
        ),
        entries,
    )


def check_link_placeholders(base_dir: Path, link_map: Dict[str, str]) -> CheckResult:
    template_dir = base_dir / "templates"
    template_paths = [
        path
        for path in template_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in {".htm", ".html", ".txt"}
    ]

    missing_links: Dict[str, List[str]] = {}
    available_keys = set(link_map)

    square_pattern = re.compile(r"\[([A-Z0-9_]+)\]")
    brace_pattern = re.compile(r"\{\{([^{}]+)\}\}")

    for path in template_paths:
        text = _read_text(path)
        if text is None:
            continue
        tokens = set()
        for match in square_pattern.findall(text):
            if match.startswith("LINK_"):
                tokens.add(match)
        for match in brace_pattern.findall(text):
            token = match.strip()
            if token.startswith("LINK_"):
                tokens.add(token)
        undefined = sorted(token for token in tokens if token not in available_keys)
        if undefined:
            missing_links[str(path.relative_to(base_dir))] = undefined

    if missing_links:
        details = ", ".join(
            f"{template}: missing {', '.join(tokens)}"
            for template, tokens in sorted(missing_links.items())
        )
        return CheckResult(
            name="Template dynamic links",
            status=STATUS_FAIL,
            details=details,
        )

    return CheckResult(
        name="Template dynamic links",
        status=STATUS_PASS,
        details="All LINK_* placeholders are backed by links_map.txt entries.",
    )


def run_checklist(base_dir: Path) -> List[CheckResult]:
    results: List[CheckResult] = []

    structure_result = check_required_structure(base_dir)
    results.append(structure_result)

    header_result = check_template_headers(base_dir)
    results.append(header_result)

    link_map_result, link_map = load_link_map(base_dir)
    results.append(link_map_result)

    if link_map is not None:
        link_placeholder_result = check_link_placeholders(base_dir, link_map)
        results.append(link_placeholder_result)

    return results


def determine_variations(
    previous: Dict[str, object] | None, current: Dict[str, object]
) -> List[str]:
    if previous is None:
        return ["No previous diagnostic run to compare against."]

    prev_checks = {
        check["name"]: check for check in previous.get("checklist", []) if isinstance(check, dict)
    }
    variations: List[str] = []

    for check in current.get("checklist", []):
        name = check.get("name")
        if name is None:
            continue
        previous_check = prev_checks.pop(name, None)
        if previous_check is None:
            variations.append(f"New check added: {name} -> {check.get('status')}")
            continue
        if previous_check.get("status") != check.get("status"):
            variations.append(
                f"Status change for {name}: {previous_check.get('status')} -> {check.get('status')}"
            )
        elif previous_check.get("details") != check.get("details"):
            variations.append(
                f"Details changed for {name}: '{previous_check.get('details')}' -> '{check.get('details')}'"
            )

    for missing_name, prev_check in prev_checks.items():
        variations.append(
            f"Check removed since last run: {missing_name} (previous status {prev_check.get('status')})"
        )

    if not variations:
        return ["No variations detected compared to the previous run."]

    return variations


def save_results(results: List[CheckResult], output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = str(int(time.time()))
    output_path = output_dir / f"{timestamp}.json"
    payload = {
        "timestamp": timestamp,
        "checklist": [result.as_dict() for result in results],
    }
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return output_path


def load_previous_results(
    results_dir: Path, current_path: Path
) -> Dict[str, object] | None:
    candidates = sorted(p for p in results_dir.glob("*.json") if p != current_path)
    if not candidates:
        return None
    previous_path = candidates[-1]
    try:
        return json.loads(previous_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def format_results_table(results: Sequence[CheckResult]) -> str:
    name_width = max((len(result.name) for result in results), default=4)
    status_width = max((len(result.status) for result in results), default=6)
    header = f"{'Check':{name_width}} | {'Status':{status_width}} | Details"
    divider = "-" * len(header)
    rows = [header, divider]
    for result in results:
        rows.append(f"{result.name:{name_width}} | {result.status:{status_width}} | {result.details}")
    return "\n".join(rows)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run repository diagnostics and persist the checklist outcome with a POSIX timestamp filename."
        )
    )
    parser.add_argument(
        "--results-dir",
        type=Path,
        default=RESULTS_DIR,
        help="Directory where diagnostic results are stored (default: %(default)s)",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    checklist_results = run_checklist(BASE_DIR)
    print(format_results_table(checklist_results))

    output_path = save_results(checklist_results, args.results_dir)
    print(f"\nSaved diagnostic report to {output_path}")

    previous_payload = load_previous_results(args.results_dir, output_path)
    variations = determine_variations(previous_payload, {
        "timestamp": output_path.stem,
        "checklist": [result.as_dict() for result in checklist_results],
    })
    print("\nComparison with previous run:")
    for line in variations:
        print(f"- {line}")

    exit_status = 0 if all(result.status != STATUS_FAIL for result in checklist_results) else 1
    return exit_status


if __name__ == "__main__":
    sys.exit(main())
