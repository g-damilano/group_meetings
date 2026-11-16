"""Report on workflow input examples without becoming the runtime path."""

from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import date
from pathlib import Path, PureWindowsPath
from typing import Any


SCHEMA = "office-communication-scaffold.workflow-input-report.v1"
GROUP_ROTA_RELATIVE_PATH = "examples/workbooks/group_meetings_rota_example.csv"
INDUCTION_RECIPIENTS_RELATIVE_PATH = "examples/workbooks/induction_recipients_example.csv"
INDUCTION_TEMPLATE_MANIFEST_RELATIVE_PATH = "examples/workbooks/induction_template_manifest_example.csv"

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@([A-Z0-9.-]+\.[A-Z]{2,63})\b", re.I)
ALLOWED_EMAIL_DOMAINS = {"example.invalid"}
TRUE_VALUES = {"true", "yes", "y", "1", "on"}
FALSE_VALUES = {"false", "no", "n", "0", "off"}
VALID_TEMPLATE_MODES = {"broadcast", "personal"}
VALID_DELIVERY_ACTIONS = {"", "display", "draft", "send"}

GROUP_REQUIRED_HEADERS = [
    "meeting_date",
    "literature_presenter",
    "results_presenter",
    "alternate_text",
    "room",
    "lab_manager",
    "meeting_lead",
    "group_recipient",
    "manager_recipient",
    "include_same_day_reminder",
    "send_presenter_reminders",
    "send_meeting_lead_reminder",
    "send_lab_manager_reminder",
]
INDUCTION_RECIPIENT_REQUIRED_HEADERS = [
    "alias",
    "email",
    "first_name",
    "last_name",
    "full_name",
    "start_date",
    "cohort",
    "active",
]
INDUCTION_TEMPLATE_REQUIRED_HEADERS = [
    "template_id",
    "mode",
    "target_key",
    "display_name",
    "template_path",
    "inline_image_folder",
    "attachment_base_folder",
    "attachments",
    "cc",
    "bcc",
    "delivery_action",
    "defer_minutes",
    "urgent_verify_after_submit",
    "enabled",
]


def read_csv_rows(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = list(reader.fieldnames or [])
        rows: list[dict[str, str]] = []
        for row in reader:
            normalized = {str(key): "" if value is None else str(value) for key, value in row.items() if key is not None}
            normalized["_line"] = str(reader.line_num)
            rows.append(normalized)
    return headers, rows


def add_issue(issues: list[dict[str, Any]], relative_path: str, line: str, code: str, message: str) -> None:
    issues.append(
        {
            "path": relative_path,
            "line": line,
            "code": code,
            "message": message,
        }
    )


def issue_text(issue: dict[str, Any]) -> str:
    location = str(issue["path"])
    if issue.get("line"):
        location = f"{location}:{issue['line']}"
    return f"{location} {issue['code']}: {issue['message']}"


def missing_headers(headers: list[str], required: list[str]) -> list[str]:
    return [header for header in required if header not in headers]


def bool_value(value: str, default: bool) -> tuple[bool, bool]:
    normalized = value.strip().lower()
    if normalized == "":
        return default, True
    if normalized in TRUE_VALUES:
        return True, True
    if normalized in FALSE_VALUES:
        return False, True
    return default, False


def is_iso_date(value: str) -> bool:
    try:
        date.fromisoformat(value.strip())
    except ValueError:
        return False
    return True


def has_example_email_domain(value: str) -> bool:
    domains = [match.group(1).lower() for match in EMAIL_RE.finditer(value)]
    return all(domain in ALLOWED_EMAIL_DOMAINS for domain in domains)


def split_asset_list(value: str) -> list[str]:
    normalized = value.replace(",", ";")
    return [item.strip() for item in normalized.split(";") if item.strip()]


def resolve_manifest_path(manifest_path: Path, path_text: str) -> Path:
    candidate = Path(path_text)
    if candidate.is_absolute() or PureWindowsPath(path_text).is_absolute():
        return candidate
    return manifest_path.parent / candidate


def first_group_row_plan_estimate(rows: list[dict[str, str]]) -> int:
    for row in rows:
        if row.get("alternate_text", "").strip():
            return 1
        if not row.get("meeting_date", "").strip():
            continue
        estimate = 1
        presenter_reminders, _ = bool_value(row.get("send_presenter_reminders", ""), True)
        same_day, _ = bool_value(row.get("include_same_day_reminder", ""), False)
        lead_reminder, _ = bool_value(row.get("send_meeting_lead_reminder", ""), True)
        manager_reminder, _ = bool_value(row.get("send_lab_manager_reminder", ""), True)
        if presenter_reminders:
            estimate += 2
        if same_day:
            estimate += 1
        if lead_reminder:
            estimate += 1
        if manager_reminder:
            estimate += 1
        return estimate
    return 0


def validate_group_meetings(root: Path, relative_path: str = GROUP_ROTA_RELATIVE_PATH) -> dict[str, Any]:
    path = root / relative_path
    issues: list[dict[str, Any]] = []
    headers, rows = read_csv_rows(path)

    for header in missing_headers(headers, GROUP_REQUIRED_HEADERS):
        add_issue(issues, relative_path, "", "HEADER_MISSING", f"Missing required header: {header}")

    configured_count = 0
    no_meeting_count = 0
    boolean_fields = [
        "include_same_day_reminder",
        "send_presenter_reminders",
        "send_meeting_lead_reminder",
        "send_lab_manager_reminder",
    ]

    for row in rows:
        line = row.get("_line", "")
        if not row.get("meeting_date", "").strip():
            add_issue(issues, relative_path, line, "MEETING_DATE_MISSING", "meeting_date is required.")
        elif not is_iso_date(row.get("meeting_date", "")):
            add_issue(issues, relative_path, line, "MEETING_DATE_INVALID", "meeting_date must use YYYY-MM-DD.")

        for field in ["group_recipient", "manager_recipient"]:
            if not row.get(field, "").strip():
                add_issue(issues, relative_path, line, "RECIPIENT_MISSING", f"{field} is required.")

        if row.get("alternate_text", "").strip():
            no_meeting_count += 1
        else:
            configured_count += 1
            for field in ["literature_presenter", "results_presenter", "lab_manager"]:
                if not row.get(field, "").strip():
                    add_issue(issues, relative_path, line, "CONFIGURED_MEETING_FIELD_MISSING", f"{field} is required.")

        for field in boolean_fields:
            _, valid = bool_value(row.get(field, ""), True)
            if not valid:
                add_issue(issues, relative_path, line, "BOOLEAN_INVALID", f"{field} must be boolean-like.")

    return {
        "source": relative_path,
        "row_count": len(rows),
        "configured_meeting_count": configured_count,
        "no_meeting_count": no_meeting_count,
        "first_route_planned_message_estimate": first_group_row_plan_estimate(rows),
        "dry_run_route_ready": len(rows) > 0 and not issues,
        "issues": issues,
    }


def validate_induction_recipients(root: Path, relative_path: str = INDUCTION_RECIPIENTS_RELATIVE_PATH) -> dict[str, Any]:
    path = root / relative_path
    issues: list[dict[str, Any]] = []
    headers, rows = read_csv_rows(path)
    for header in missing_headers(headers, INDUCTION_RECIPIENT_REQUIRED_HEADERS):
        add_issue(issues, relative_path, "", "HEADER_MISSING", f"Missing required header: {header}")

    active_count = 0
    inactive_count = 0
    directory_keys: set[str] = set()

    for row in rows:
        line = row.get("_line", "")
        active, active_valid = bool_value(row.get("active", ""), True)
        if not active_valid:
            add_issue(issues, relative_path, line, "ACTIVE_INVALID", "active must be boolean-like.")
        if active:
            active_count += 1
            if not row.get("email", "").strip():
                add_issue(issues, relative_path, line, "ACTIVE_RECIPIENT_MISSING_EMAIL", "Active recipient rows require email.")
        else:
            inactive_count += 1

        email = row.get("email", "").strip()
        if email and not has_example_email_domain(email):
            add_issue(issues, relative_path, line, "EMAIL_DOMAIN_INVALID", "Public examples must use example.invalid addresses.")

        for field in ["alias", "full_name"]:
            value = row.get(field, "").strip().lower()
            if value:
                directory_keys.add(value)

    return {
        "source": relative_path,
        "row_count": len(rows),
        "active_count": active_count,
        "inactive_count": inactive_count,
        "directory_key_count": len(directory_keys),
        "dry_run_route_ready": active_count > 0 and not issues,
        "issues": issues,
        "directory_keys": sorted(directory_keys),
    }


def validate_induction_manifest(
    root: Path,
    recipient_report: dict[str, Any],
    relative_path: str = INDUCTION_TEMPLATE_MANIFEST_RELATIVE_PATH,
) -> dict[str, Any]:
    manifest_path = root / relative_path
    issues: list[dict[str, Any]] = []
    headers, rows = read_csv_rows(manifest_path)
    for header in missing_headers(headers, INDUCTION_TEMPLATE_REQUIRED_HEADERS):
        add_issue(issues, relative_path, "", "HEADER_MISSING", f"Missing required header: {header}")

    enabled_count = 0
    enabled_existing_template_count = 0
    broadcast_count = 0
    personal_count = 0
    planned_message_estimate = 0
    active_recipient_count = int(recipient_report.get("active_count", 0))
    directory_keys = set(recipient_report.get("directory_keys", []))

    for row in rows:
        line = row.get("_line", "")
        enabled, enabled_valid = bool_value(row.get("enabled", ""), True)
        if not enabled_valid:
            add_issue(issues, relative_path, line, "ENABLED_INVALID", "enabled must be boolean-like.")
        if not enabled:
            continue

        enabled_count += 1
        template_id = row.get("template_id", "").strip()
        if not template_id:
            add_issue(issues, relative_path, line, "TEMPLATE_ID_MISSING", "Enabled template rows require template_id.")

        mode = row.get("mode", "").strip().lower()
        if mode not in VALID_TEMPLATE_MODES:
            add_issue(issues, relative_path, line, "TEMPLATE_MODE_INVALID", "mode must be broadcast or personal.")

        template_text = row.get("template_path", "").strip()
        template_exists = False
        if not template_text:
            add_issue(issues, relative_path, line, "TEMPLATE_PATH_MISSING", "Enabled template rows require template_path.")
        else:
            template_exists = resolve_manifest_path(manifest_path, template_text).is_file()
            if not template_exists:
                add_issue(issues, relative_path, line, "TEMPLATE_MISSING", f"Template path does not exist: {template_text}")
        if template_exists:
            enabled_existing_template_count += 1

        delivery_action = row.get("delivery_action", "").strip().lower()
        if delivery_action not in VALID_DELIVERY_ACTIONS:
            add_issue(issues, relative_path, line, "DELIVERY_ACTION_INVALID", "delivery_action must be blank, display, draft, or send.")

        defer_minutes = row.get("defer_minutes", "").strip()
        if defer_minutes and not defer_minutes.isdigit():
            add_issue(issues, relative_path, line, "DEFER_MINUTES_INVALID", "defer_minutes must be a non-negative integer.")

        _, urgent_valid = bool_value(row.get("urgent_verify_after_submit", ""), False)
        if not urgent_valid:
            add_issue(issues, relative_path, line, "URGENT_VERIFY_INVALID", "urgent_verify_after_submit must be boolean-like.")

        inline_folder = row.get("inline_image_folder", "").strip()
        if inline_folder and not resolve_manifest_path(manifest_path, inline_folder).is_dir():
            add_issue(issues, relative_path, line, "INLINE_IMAGE_FOLDER_MISSING", f"Inline image folder does not exist: {inline_folder}")

        attachments = split_asset_list(row.get("attachments", ""))
        attachment_base = row.get("attachment_base_folder", "").strip()
        if attachments and not attachment_base:
            add_issue(issues, relative_path, line, "ATTACHMENT_BASE_MISSING", "attachments require attachment_base_folder.")
        elif attachment_base:
            base_path = resolve_manifest_path(manifest_path, attachment_base)
            if not base_path.is_dir():
                add_issue(issues, relative_path, line, "ATTACHMENT_BASE_MISSING", f"Attachment base folder does not exist: {attachment_base}")
            else:
                for attachment in attachments:
                    if not (base_path / attachment).is_file():
                        add_issue(issues, relative_path, line, "ATTACHMENT_MISSING", f"Attachment does not exist: {attachment}")

        for bucket in ["cc", "bcc"]:
            for recipient in split_asset_list(row.get(bucket, "")):
                if EMAIL_RE.search(recipient):
                    if not has_example_email_domain(recipient):
                        add_issue(issues, relative_path, line, "EMAIL_DOMAIN_INVALID", "Public examples must use example.invalid addresses.")
                elif recipient.strip().lower() not in directory_keys:
                    add_issue(issues, relative_path, line, "RECIPIENT_ALIAS_UNRESOLVED", f"{bucket} alias is not present in recipients: {recipient}")

        if template_exists and mode == "broadcast":
            broadcast_count += 1
            planned_message_estimate += 1
        elif template_exists and mode == "personal":
            personal_count += 1
            planned_message_estimate += active_recipient_count

    return {
        "source": relative_path,
        "row_count": len(rows),
        "enabled_count": enabled_count,
        "enabled_existing_template_count": enabled_existing_template_count,
        "broadcast_template_count": broadcast_count,
        "personal_template_count": personal_count,
        "planned_message_estimate": planned_message_estimate,
        "dry_run_route_ready": enabled_count > 0 and enabled_existing_template_count > 0 and not issues,
        "issues": issues,
    }


def build_report(
    root: Path,
    group_rota: str = GROUP_ROTA_RELATIVE_PATH,
    induction_recipients: str = INDUCTION_RECIPIENTS_RELATIVE_PATH,
    induction_templates: str = INDUCTION_TEMPLATE_MANIFEST_RELATIVE_PATH,
) -> dict[str, Any]:
    root = root.resolve()
    group_report = validate_group_meetings(root, group_rota)
    recipient_report = validate_induction_recipients(root, induction_recipients)
    manifest_report = validate_induction_manifest(root, recipient_report, induction_templates)
    issues = []
    issues.extend(group_report["issues"])
    issues.extend(recipient_report["issues"])
    issues.extend(manifest_report["issues"])

    return {
        "schema": SCHEMA,
        "report_only": True,
        "mutations_performed": False,
        "real_email_sending_performed": False,
        "workflow_inputs": {
            "group_meetings": group_report,
            "induction": {
                "recipients": recipient_report,
                "template_manifest": manifest_report,
                "planned_message_estimate": manifest_report["planned_message_estimate"],
                "dry_run_route_ready": recipient_report["dry_run_route_ready"] and manifest_report["dry_run_route_ready"],
            },
        },
        "commands": {
            "group_meetings": f"cscript.exe //NoLogo .\\Scaffold.vbs group-meetings --dry-run --input .\\{group_rota.replace('/', '\\')}",
            "induction": (
                "cscript.exe //NoLogo .\\Scaffold.vbs induction --dry-run "
                f"--recipients .\\{induction_recipients.replace('/', '\\')} "
                f"--templates .\\{induction_templates.replace('/', '\\')}"
            ),
        },
        "validation_errors": [issue_text(issue) for issue in issues],
    }


def validate_default_examples(root: Path) -> list[str]:
    return list(build_report(root)["validation_errors"])


def print_text(report: dict[str, Any]) -> None:
    group = report["workflow_inputs"]["group_meetings"]
    induction = report["workflow_inputs"]["induction"]
    recipients = induction["recipients"]
    manifest = induction["template_manifest"]

    print("Workflow input report")
    print(f"schema: {report['schema']}")
    print(f"report_only: {str(report['report_only']).lower()}")
    print(f"group_meetings_rows: {group['row_count']}")
    print(f"group_meetings_route_ready: {str(group['dry_run_route_ready']).lower()}")
    print(f"group_meetings_first_route_planned_message_estimate: {group['first_route_planned_message_estimate']}")
    print(f"induction_recipient_rows: {recipients['row_count']}")
    print(f"induction_active_recipients: {recipients['active_count']}")
    print(f"induction_enabled_templates: {manifest['enabled_count']}")
    print(f"induction_planned_message_estimate: {induction['planned_message_estimate']}")
    print(f"induction_route_ready: {str(induction['dry_run_route_ready']).lower()}")
    if report["validation_errors"]:
        print("validation_errors:")
        for error in report["validation_errors"]:
            print(f"- {error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate scaffold workflow input examples.")
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
