"""Report on the shared contract fixture without touching Outlook or mailboxes."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import validate_scaffold


SCHEMA = "office-communication-scaffold.contract-fixture-report.v1"
FIXTURE_RELATIVE_PATH = "fixtures/contracts/stage4_shared_contract_fixture.json"
CANONICAL_CONCEPTS = [
    "MessageSpec",
    "RecipientSpec",
    "DeliveryPolicy",
    "ManagedEmailKey",
    "DispatchResult",
    "WorkflowContext",
    "HealthReport",
]


def load_fixture(root: Path) -> dict[str, Any]:
    fixture = validate_scaffold.load_json(root / FIXTURE_RELATIVE_PATH)
    if not isinstance(fixture, dict):
        raise ValueError(f"{FIXTURE_RELATIVE_PATH} is not a JSON object")
    return fixture


def unique_ordered(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def fixture_email_domains(fixture: dict[str, Any]) -> list[str]:
    domains: list[str] = []
    for message in fixture.get("messages", []):
        if not isinstance(message, dict):
            continue
        for field in ["to", "cc", "bcc"]:
            value = message.get(field)
            if isinstance(value, str):
                for match in validate_scaffold.EMAIL_RE.finditer(value):
                    domains.append(match.group(1).lower())
        recipients = message.get("recipients", {})
        if not isinstance(recipients, dict):
            continue
        for bucket in ["to", "cc", "bcc"]:
            values = recipients.get(bucket, [])
            if not isinstance(values, list):
                continue
            for recipient in values:
                if not isinstance(recipient, dict):
                    continue
                smtp = recipient.get("smtp_address", "")
                if isinstance(smtp, str):
                    for match in validate_scaffold.EMAIL_RE.finditer(smtp):
                        domains.append(match.group(1).lower())
    return unique_ordered(domains)


def build_report(root: Path) -> dict[str, Any]:
    root = root.resolve()
    fixture = load_fixture(root)
    messages = [item for item in fixture.get("messages", []) if isinstance(item, dict)]
    dispatch_results = [item for item in fixture.get("dispatch_results", []) if isinstance(item, dict)]
    validation_errors = validate_scaffold.validate_shared_contract_fixture(root)
    validation_errors.extend(validate_scaffold.validate_dispatch_state_parity(root))
    email_domains = fixture_email_domains(fixture)
    allowed_domains = sorted(validate_scaffold.allowed_email_domains(root))

    return {
        "schema": SCHEMA,
        "fixture_path": FIXTURE_RELATIVE_PATH,
        "fixture_schema": fixture.get("schema", ""),
        "workflow_name": fixture.get("workflow_name", ""),
        "dry_run": fixture.get("dry_run") is True,
        "canonical_concepts": CANONICAL_CONCEPTS,
        "message_count": len(messages),
        "dispatch_result_count": len(dispatch_results),
        "workflow_references": unique_ordered(
            [
                str(message.get("metadata", {}).get("workflow_reference", ""))
                for message in messages
                if isinstance(message.get("metadata"), dict)
            ]
        ),
        "message_types": unique_ordered([str(message.get("message_type", "")) for message in messages]),
        "managed_keys": [str(message.get("message_key", "")) for message in messages],
        "dispatch_states": unique_ordered([str(result.get("state", "")) for result in dispatch_results]),
        "email_domains": email_domains,
        "example_address_only": all(domain in allowed_domains for domain in email_domains),
        "dispatch_state_parity": validate_scaffold.validate_dispatch_state_parity(root) == [],
        "validation_errors": validation_errors,
        "report_only": True,
        "mutations_performed": False,
        "real_email_sending_performed": False,
    }


def print_text(report: dict[str, Any]) -> None:
    print("Contract fixture report")
    print(f"schema: {report['schema']}")
    print(f"fixture_path: {report['fixture_path']}")
    print(f"fixture_schema: {report['fixture_schema']}")
    print(f"workflow_name: {report['workflow_name']}")
    print(f"dry_run: {str(report['dry_run']).lower()}")
    print(f"message_count: {report['message_count']}")
    print(f"dispatch_result_count: {report['dispatch_result_count']}")
    print(f"workflow_references: {', '.join(report['workflow_references'])}")
    print(f"dispatch_states: {', '.join(report['dispatch_states'])}")
    print(f"example_address_only: {str(report['example_address_only']).lower()}")
    print(f"dispatch_state_parity: {str(report['dispatch_state_parity']).lower()}")
    if report["validation_errors"]:
        print("validation_errors:")
        for error in report["validation_errors"]:
            print(f"- {error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Summarize the shared scaffold contract fixture.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--fail-if-invalid", action="store_true")
    args = parser.parse_args()

    report = build_report(args.root)
    if args.format == "json":
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print_text(report)

    if args.fail_if_invalid and report["validation_errors"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
