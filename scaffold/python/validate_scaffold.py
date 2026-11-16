"""Standard-library scaffold validation helpers.

Python is a development-side checker in this repository. The operational
architecture remains VBS-first and contract-driven.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Iterable

import workflow_input_report
import command_route_report


EXPECTED_STATES = [
    "NOT_REQUESTED",
    "COMPOSED",
    "DISPLAYED",
    "DRAFTED",
    "SEND_REQUESTED",
    "QUEUED_OUTBOX",
    "SENT_CONFIRMED",
    "DUPLICATE_SUPPRESSED",
    "WAITING_DEFERRED",
    "BLOCKED_OFFLINE",
    "BLOCKED_AUTH",
    "STALE_MANAGED_OUTBOX",
    "REPLACED_MANAGED_OUTBOX",
    "ACTION_REQUIRED",
    "FAILED_SAFE",
]

CSV_REQUIREMENTS = {
    "examples/workbooks/group_meetings_rota_example.csv": [
        "meeting_date",
        "literature_presenter",
        "results_presenter",
        "group_recipient",
        "manager_recipient",
    ],
    "examples/workbooks/induction_recipients_example.csv": [
        "alias",
        "email",
        "first_name",
        "last_name",
        "active",
    ],
    "examples/workbooks/induction_template_manifest_example.csv": [
        "template_id",
        "mode",
        "target_key",
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
    ],
    "vbs/tests/fixtures/induction_recipients_fixture.csv": [
        "alias",
        "email",
        "first_name",
        "last_name",
        "active",
    ],
    "vbs/tests/fixtures/induction_template_manifest_fixture.csv": [
        "template_id",
        "mode",
        "target_key",
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
    ],
}

TEXT_SUFFIXES = {".md", ".txt", ".csv", ".json", ".vbs", ".html", ".py", ".xml"}
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@([A-Z0-9.-]+\.[A-Z]{2,63})\b", re.I)
DEFAULT_ALLOWED_EMAIL_DOMAINS = {"example.invalid"}
DEFAULT_BLOCKED_TEXT_PATTERNS = [
    r"OneDrive\s+-",
    r"Imperial\s+College",
    r"https?://[^\s\"']*sharepoint[^\s\"']*",
]
DEFAULT_BLOCKED_GENERATED_SUFFIXES = {".log", ".pyc", ".pyo", ".zip", ".msg"}
DEFAULT_BLOCKED_GENERATED_DIRECTORY_NAMES = {"__pycache__"}
DEFAULT_RELEASE_EXCLUDED_DIRECTORY_NAMES = {"legacy"}
LEGACY_ALTERNATE_MAP_SCHEMA = "office-communication-scaffold.legacy-alternate-map.v1"
LEGACY_ALTERNATE_STATUSES = {"ported_alternate", "behaviour_mapped", "sanitised_fixture"}
WATCHDOG_REPORT_SCHEMA = "office-communication-scaffold.watchdog-report-fixture.v1"
WATCHDOG_REPAIR_SCHEMA = "office-communication-scaffold.watchdog-managed-repair-guardrail-fixture.v1"


def load_json(path: Path) -> object:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def validate_dispatch_states(root: Path) -> list[str]:
    errors: list[str] = []
    states_path = root / "fixtures" / "dispatch_states" / "states.json"
    data = load_json(states_path)
    states = data.get("states") if isinstance(data, dict) else None
    if states != EXPECTED_STATES:
        errors.append("dispatch state fixture does not match expected state machine")
    return errors


def extract_vbs_dispatch_states(root: Path) -> list[str]:
    contracts_path = root / "vbs" / "core" / "contracts.vbs"
    text = contracts_path.read_text(encoding="utf-8", errors="ignore")
    matches = re.findall(r'Const\s+SC_STATE_[A-Z0-9_]+\s*=\s*"([^"]+)"', text, flags=re.I)
    return [match.strip() for match in matches]


def workflow_schema_dispatch_states(root: Path) -> list[str]:
    schema = load_json(root / "contracts" / "workflow_schema.json")
    if not isinstance(schema, dict):
        return []
    defs = schema.get("$defs", {})
    if not isinstance(defs, dict):
        return []
    dispatch_state = defs.get("dispatch_state", {})
    if not isinstance(dispatch_state, dict):
        return []
    values = dispatch_state.get("enum", [])
    if not isinstance(values, list):
        return []
    return [str(value) for value in values]


def validate_dispatch_state_parity(root: Path) -> list[str]:
    errors: list[str] = []
    data = load_json(root / "fixtures" / "dispatch_states" / "states.json")
    fixture_states = data.get("states") if isinstance(data, dict) else None
    vbs_states = extract_vbs_dispatch_states(root)
    schema_states = workflow_schema_dispatch_states(root)
    if vbs_states != EXPECTED_STATES:
        errors.append("VBS dispatch state constants do not match expected state machine")
    if schema_states != EXPECTED_STATES:
        errors.append("workflow schema dispatch_state enum does not match expected state machine")
    if fixture_states != EXPECTED_STATES:
        errors.append("dispatch state fixture does not match expected state machine")
    if vbs_states != schema_states or vbs_states != fixture_states:
        errors.append("dispatch states are not in parity across VBS, schema, and fixture")
    return errors


def validate_workflow_schema(root: Path) -> list[str]:
    errors: list[str] = []
    schema = load_json(root / "contracts" / "workflow_schema.json")
    if not isinstance(schema, dict):
        return ["workflow schema is not a JSON object"]
    defs = schema.get("$defs", {})
    for definition in [
        "message_spec",
        "recipient_spec",
        "delivery_policy",
        "managed_email_key",
        "dispatch_result",
        "workflow_context",
        "health_report",
    ]:
        if definition not in defs:
            errors.append(f"workflow schema missing {definition} definition")
    return errors


def managed_key_is_portable(value: object) -> bool:
    if not isinstance(value, str):
        return False
    parts = value.split("|")
    return len(parts) == 5 and all(part.strip() for part in parts)


def validate_recipient_specs(root: Path, relative_prefix: str, recipients: object) -> list[str]:
    errors: list[str] = []
    if not isinstance(recipients, dict):
        return [f"{relative_prefix} recipients is not an object"]
    allowed_domains = allowed_email_domains(root)
    for bucket in ["to", "cc", "bcc"]:
        values = recipients.get(bucket, [])
        if not isinstance(values, list):
            errors.append(f"{relative_prefix} recipients.{bucket} is not a list")
            continue
        for index, spec in enumerate(values, start=1):
            if not isinstance(spec, dict):
                errors.append(f"{relative_prefix} recipients.{bucket}[{index}] is not an object")
                continue
            if not str(spec.get("raw_value", "")).strip():
                errors.append(f"{relative_prefix} recipients.{bucket}[{index}] missing raw_value")
            status = spec.get("resolution_status")
            if status not in {"direct", "resolved", "unresolved", "fallback"}:
                errors.append(f"{relative_prefix} recipients.{bucket}[{index}] has unsupported resolution_status")
            for name_key in [
                "display_name",
                "first_name",
                "last_name",
                "surname",
                "full_name",
                "name_parse_status",
                "name_parse_source",
                "name_parse_order",
            ]:
                if name_key in spec and not isinstance(spec.get(name_key), str):
                    errors.append(f"{relative_prefix} recipients.{bucket}[{index}] {name_key} is not a string")
            parse_status = spec.get("name_parse_status")
            if parse_status is not None and parse_status not in {"not_attempted", "parsed", "partial"}:
                errors.append(f"{relative_prefix} recipients.{bucket}[{index}] has unsupported name_parse_status")
            parse_order = spec.get("name_parse_order")
            if parse_order is not None and parse_order not in {"", "given-family", "family-given", "single-token"}:
                errors.append(f"{relative_prefix} recipients.{bucket}[{index}] has unsupported name_parse_order")
            smtp = str(spec.get("smtp_address", "")).strip()
            if smtp:
                match = EMAIL_RE.fullmatch(smtp)
                if not match:
                    errors.append(f"{relative_prefix} recipients.{bucket}[{index}] smtp_address is not an email address")
                elif match.group(1).lower() not in allowed_domains:
                    errors.append(f"{relative_prefix} recipients.{bucket}[{index}] smtp_address is not an example address")
    return errors


def validate_shared_contract_fixture(root: Path) -> list[str]:
    errors: list[str] = []
    relative_path = "fixtures/contracts/stage4_shared_contract_fixture.json"
    path = root / relative_path
    if not path.exists():
        return [f"missing shared contract fixture: {relative_path}"]

    data = load_json(path)
    if not isinstance(data, dict):
        return [f"{relative_path} is not a JSON object"]
    if data.get("schema") != "office-communication-scaffold.shared-contract-fixture.v1":
        errors.append(f"{relative_path} schema mismatch")
    if data.get("dry_run") is not True:
        errors.append(f"{relative_path} must remain dry_run=true")

    context = data.get("workflow_context")
    if not isinstance(context, dict):
        errors.append(f"{relative_path} missing workflow_context object")
    elif context.get("dry_run") is not True:
        errors.append(f"{relative_path} workflow_context must remain dry_run=true")

    messages = data.get("messages")
    if not isinstance(messages, list) or not messages:
        errors.append(f"{relative_path} must contain messages")
        messages = []

    message_keys: set[str] = set()
    for index, message in enumerate(messages, start=1):
        prefix = f"{relative_path}:messages[{index}]"
        if not isinstance(message, dict):
            errors.append(f"{prefix} is not an object")
            continue
        message_key = message.get("message_key")
        if not managed_key_is_portable(message_key):
            errors.append(f"{prefix} has non-portable message_key")
        else:
            message_keys.add(str(message_key))
        for required in ["message_type", "to", "subject"]:
            if not str(message.get(required, "")).strip():
                errors.append(f"{prefix} missing {required}")
        policy = message.get("delivery_policy")
        if not isinstance(policy, dict):
            errors.append(f"{prefix} missing delivery_policy object")
        else:
            if policy.get("action") not in {"display", "draft", "save", "send"}:
                errors.append(f"{prefix} has unsupported delivery_policy.action")
            defer = policy.get("defer_minutes", 0)
            if not isinstance(defer, int) or defer < 0:
                errors.append(f"{prefix} has invalid delivery_policy.defer_minutes")
        attachments = message.get("attachments", [])
        if not isinstance(attachments, list) or not all(isinstance(item, str) for item in attachments):
            errors.append(f"{prefix} attachments must be a list of strings")
        if "recipients" in message:
            errors.extend(validate_recipient_specs(root, prefix, message["recipients"]))

    dispatch_results = data.get("dispatch_results", [])
    if not isinstance(dispatch_results, list):
        errors.append(f"{relative_path} dispatch_results is not a list")
    else:
        for index, result in enumerate(dispatch_results, start=1):
            prefix = f"{relative_path}:dispatch_results[{index}]"
            if not isinstance(result, dict):
                errors.append(f"{prefix} is not an object")
                continue
            if result.get("state") not in EXPECTED_STATES:
                errors.append(f"{prefix} has unsupported state")
            managed_key = result.get("managed_key")
            if managed_key and managed_key not in message_keys:
                errors.append(f"{prefix} managed_key does not reference a fixture message")

    health_report = data.get("health_report")
    if not isinstance(health_report, dict):
        errors.append(f"{relative_path} missing health_report object")
    else:
        for count_key in ["outbox_count", "managed_outbox_count", "unmanaged_outbox_count"]:
            count_value = health_report.get(count_key, 0)
            if not isinstance(count_value, int) or count_value < 0:
                errors.append(f"{relative_path} health_report.{count_key} is not a non-negative integer")
        if not isinstance(health_report.get("action_required_messages", []), list):
            errors.append(f"{relative_path} health_report.action_required_messages is not a list")

    return errors


def validate_watchdog_report_fixture_data(root: Path, relative_path: str, data: object) -> list[str]:
    errors: list[str] = []
    if not isinstance(data, dict):
        return [f"{relative_path} is not a JSON object"]
    if data.get("schema") != WATCHDOG_REPORT_SCHEMA:
        errors.append(f"{relative_path} schema mismatch")
    if data.get("report_only") is not True:
        errors.append(f"{relative_path} must remain report_only=true")
    if data.get("live_outlook_required") is not False:
        errors.append(f"{relative_path} must remain live_outlook_required=false")
    if data.get("mailbox_mutation_allowed") is not False:
        errors.append(f"{relative_path} must keep mailbox_mutation_allowed=false")
    if data.get("repair_committed") is not False:
        errors.append(f"{relative_path} must keep repair_committed=false")
    if data.get("state") not in EXPECTED_STATES:
        errors.append(f"{relative_path} has unsupported top-level state")

    destructive_actions = data.get("destructive_actions")
    if not isinstance(destructive_actions, dict):
        errors.append(f"{relative_path} missing destructive_actions object")
    else:
        for action_key in ["send_committed", "delete_committed", "replace_committed", "repair_committed"]:
            if destructive_actions.get(action_key) is not False:
                errors.append(f"{relative_path} destructive_actions.{action_key} must be false")

    transport = data.get("transport")
    if not isinstance(transport, dict):
        errors.append(f"{relative_path} missing transport object")
    else:
        if transport.get("blocked_or_auth_suspected") is True:
            errors.append(f"{relative_path} report fixture should not mark blocked transport as repairable")

    summary = data.get("outbox_summary")
    if not isinstance(summary, dict):
        errors.append(f"{relative_path} missing outbox_summary object")
        summary = {}

    items = data.get("items")
    if not isinstance(items, list) or not items:
        errors.append(f"{relative_path} must contain watchdog items")
        items = []

    managed_count = 0
    unmanaged_count = 0
    action_required_count = 0
    replace_eligible_count = 0

    for index, item in enumerate(items, start=1):
        prefix = f"{relative_path}:items[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{prefix} is not an object")
            continue

        state = item.get("dispatch_state")
        if state not in EXPECTED_STATES:
            errors.append(f"{prefix} has unsupported dispatch_state")

        managed = item.get("managed")
        if not isinstance(managed, bool):
            errors.append(f"{prefix} managed is not boolean")
            managed = False
        if managed:
            managed_count += 1
        else:
            unmanaged_count += 1

        managed_key = str(item.get("managed_key", "")).strip()
        managed_key_valid = item.get("managed_key_valid")
        if not isinstance(managed_key_valid, bool):
            errors.append(f"{prefix} managed_key_valid is not boolean")
            managed_key_valid = False
        if managed and managed_key_valid and not managed_key_is_portable(managed_key):
            errors.append(f"{prefix} managed_key is not portable")
        if not managed and managed_key:
            errors.append(f"{prefix} unmarked item must not carry a managed_key")
        if not managed and managed_key_valid:
            errors.append(f"{prefix} unmarked item must not have managed_key_valid=true")

        action_required = item.get("action_required")
        replace_eligible = item.get("replace_eligible")
        would_modify = item.get("would_modify")
        repair_action = item.get("repair_action")
        if not isinstance(action_required, bool):
            errors.append(f"{prefix} action_required is not boolean")
            action_required = False
        if not isinstance(replace_eligible, bool):
            errors.append(f"{prefix} replace_eligible is not boolean")
            replace_eligible = False
        if not isinstance(would_modify, bool):
            errors.append(f"{prefix} would_modify is not boolean")
            would_modify = True
        if action_required:
            action_required_count += 1
        if replace_eligible:
            replace_eligible_count += 1

        if would_modify:
            errors.append(f"{prefix} report-only fixture must not modify mailbox items")
        if repair_action != "none":
            errors.append(f"{prefix} repair_action must be none")
        if not managed and (replace_eligible or would_modify):
            errors.append(f"{prefix} unmarked mail must never be repair-eligible or mutable")

        deferred_delivery_pending = item.get("deferred_delivery_pending")
        if not isinstance(deferred_delivery_pending, bool):
            errors.append(f"{prefix} deferred_delivery_pending is not boolean")
            deferred_delivery_pending = False
        if deferred_delivery_pending and replace_eligible:
            errors.append(f"{prefix} future deferred mail must not be replace-eligible")

        transport_allows_repair = item.get("transport_allows_repair")
        if not isinstance(transport_allows_repair, bool):
            errors.append(f"{prefix} transport_allows_repair is not boolean")
            transport_allows_repair = False

        if replace_eligible:
            if not managed:
                errors.append(f"{prefix} replacement eligibility requires managed=true")
            if not managed_key_valid:
                errors.append(f"{prefix} replacement eligibility requires a valid managed key")
            if state != "STALE_MANAGED_OUTBOX":
                errors.append(f"{prefix} replacement eligibility requires STALE_MANAGED_OUTBOX")
            if not transport_allows_repair:
                errors.append(f"{prefix} replacement eligibility requires connected transport")
            if not action_required:
                errors.append(f"{prefix} replacement eligibility must be action-required")

        age_minutes = item.get("age_minutes")
        if age_minutes is not None and (not isinstance(age_minutes, int) or age_minutes < 0):
            errors.append(f"{prefix} age_minutes must be null or a non-negative integer")

        uncertainty_reason = str(item.get("uncertainty_reason", "")).strip()
        if uncertainty_reason:
            if state not in {"ACTION_REQUIRED", "FAILED_SAFE"}:
                errors.append(f"{prefix} uncertain watchdog state must be ACTION_REQUIRED or FAILED_SAFE")
            if replace_eligible:
                errors.append(f"{prefix} uncertain watchdog state must not be replace-eligible")

    expected_counts = {
        "outbox_count": len(items),
        "managed_outbox_count": managed_count,
        "unmanaged_outbox_count": unmanaged_count,
        "action_required_count": action_required_count,
        "replace_eligible_count": replace_eligible_count,
        "mutations_performed": 0,
    }
    for key, expected in expected_counts.items():
        value = summary.get(key)
        if value != expected:
            errors.append(f"{relative_path} outbox_summary.{key} expected {expected}, got {value}")

    return errors


def validate_watchdog_report_fixture(root: Path) -> list[str]:
    relative_path = "fixtures/watchdog/report_only_health_fixture.json"
    path = root / relative_path
    if not path.exists():
        return [f"missing watchdog report fixture: {relative_path}"]
    return validate_watchdog_report_fixture_data(root, relative_path, load_json(path))


def validate_watchdog_repair_fixture_data(root: Path, relative_path: str, data: object) -> list[str]:
    errors: list[str] = []
    if not isinstance(data, dict):
        return [f"{relative_path} is not a JSON object"]
    if data.get("schema") != WATCHDOG_REPAIR_SCHEMA:
        errors.append(f"{relative_path} schema mismatch")
    if data.get("live_outlook_required") is not False:
        errors.append(f"{relative_path} must remain live_outlook_required=false")
    if data.get("uses_fake_outbox_items") is not True:
        errors.append(f"{relative_path} must use fake Outbox items")
    if data.get("real_mailbox_mutation_allowed") is not False:
        errors.append(f"{relative_path} must keep real_mailbox_mutation_allowed=false")
    if data.get("send_allowed") is not False:
        errors.append(f"{relative_path} must keep send_allowed=false")
    if data.get("unmarked_mail_mutation_allowed") is not False:
        errors.append(f"{relative_path} must keep unmarked_mail_mutation_allowed=false")

    cases = data.get("cases")
    if not isinstance(cases, list) or not cases:
        return errors + [f"{relative_path} must contain repair cases"]

    for case_index, repair_case in enumerate(cases, start=1):
        case_prefix = f"{relative_path}:cases[{case_index}]"
        if not isinstance(repair_case, dict):
            errors.append(f"{case_prefix} is not an object")
            continue
        if not str(repair_case.get("case_id", "")).strip():
            errors.append(f"{case_prefix} missing case_id")
        if repair_case.get("expected_state") not in EXPECTED_STATES:
            errors.append(f"{case_prefix} has unsupported expected_state")

        commit_repair = repair_case.get("commit_repair")
        if not isinstance(commit_repair, bool):
            errors.append(f"{case_prefix} commit_repair is not boolean")
            commit_repair = False

        real_mutations = repair_case.get("real_mailbox_mutations_performed")
        fake_mutations = repair_case.get("fake_mutations_performed")
        if real_mutations != 0:
            errors.append(f"{case_prefix} must perform zero real mailbox mutations")
        if not isinstance(fake_mutations, int) or fake_mutations < 0:
            errors.append(f"{case_prefix} fake_mutations_performed must be a non-negative integer")
            fake_mutations = 0

        actions = repair_case.get("actions")
        if not isinstance(actions, list) or not actions:
            errors.append(f"{case_prefix} must contain actions")
            actions = []

        available_count = 0
        taken_count = 0
        failed_safe_count = 0
        computed_fake_mutations = 0

        for action_index, action in enumerate(actions, start=1):
            action_prefix = f"{case_prefix}:actions[{action_index}]"
            if not isinstance(action, dict):
                errors.append(f"{action_prefix} is not an object")
                continue

            managed = action.get("managed")
            if not isinstance(managed, bool):
                errors.append(f"{action_prefix} managed is not boolean")
                managed = False
            managed_key_valid = action.get("managed_key_valid")
            if not isinstance(managed_key_valid, bool):
                errors.append(f"{action_prefix} managed_key_valid is not boolean")
                managed_key_valid = False

            managed_key = str(action.get("managed_key", "")).strip()
            if managed and managed_key_valid and not managed_key_is_portable(managed_key):
                errors.append(f"{action_prefix} managed_key is not portable")
            if not managed and managed_key:
                errors.append(f"{action_prefix} unmarked item must not carry a managed_key")
            if not managed and managed_key_valid:
                errors.append(f"{action_prefix} unmarked item must not have managed_key_valid=true")

            decision_state = action.get("decision_state")
            result_state = action.get("result_state")
            if decision_state not in EXPECTED_STATES:
                errors.append(f"{action_prefix} has unsupported decision_state")
            if result_state not in EXPECTED_STATES:
                errors.append(f"{action_prefix} has unsupported result_state")

            repair_available = action.get("repair_available")
            replacement_created = action.get("replacement_created")
            stale_deleted = action.get("stale_deleted")
            unmarked_deleted = action.get("unmarked_deleted")
            would_modify_fake_object = action.get("would_modify_fake_object")
            send_committed = action.get("send_committed")
            folder_accessed = action.get("folder_accessed")
            outlook_created = action.get("outlook_created")
            transport_allows_repair = action.get("transport_allows_repair")

            for key, value in [
                ("repair_available", repair_available),
                ("replacement_created", replacement_created),
                ("stale_deleted", stale_deleted),
                ("unmarked_deleted", unmarked_deleted),
                ("would_modify_fake_object", would_modify_fake_object),
                ("send_committed", send_committed),
                ("folder_accessed", folder_accessed),
                ("outlook_created", outlook_created),
                ("transport_allows_repair", transport_allows_repair),
            ]:
                if not isinstance(value, bool):
                    errors.append(f"{action_prefix} {key} is not boolean")

            repair_available = bool(repair_available)
            replacement_created = bool(replacement_created)
            stale_deleted = bool(stale_deleted)
            unmarked_deleted = bool(unmarked_deleted)
            would_modify_fake_object = bool(would_modify_fake_object)
            send_committed = bool(send_committed)
            folder_accessed = bool(folder_accessed)
            outlook_created = bool(outlook_created)
            transport_allows_repair = bool(transport_allows_repair)

            if repair_available:
                available_count += 1
            if result_state == "REPLACED_MANAGED_OUTBOX":
                taken_count += 1
            if result_state == "FAILED_SAFE":
                failed_safe_count += 1
            if replacement_created:
                computed_fake_mutations += 1
            if stale_deleted:
                computed_fake_mutations += 1

            if send_committed:
                errors.append(f"{action_prefix} must not commit send")
            if folder_accessed:
                errors.append(f"{action_prefix} must not access Outlook folders")
            if outlook_created:
                errors.append(f"{action_prefix} must not create Outlook.Application")
            if unmarked_deleted:
                errors.append(f"{action_prefix} must never delete unmarked mail")

            if not commit_repair and (replacement_created or stale_deleted or would_modify_fake_object):
                errors.append(f"{action_prefix} preview mode must not create, delete, or mutate")
            if not managed and (repair_available or replacement_created or stale_deleted or would_modify_fake_object):
                errors.append(f"{action_prefix} unmarked mail must never be repair-eligible or mutable")

            if stale_deleted or replacement_created or would_modify_fake_object:
                if not commit_repair:
                    errors.append(f"{action_prefix} mutation requires commit_repair=true")
                if not managed:
                    errors.append(f"{action_prefix} mutation requires managed mail")
                if not managed_key_valid:
                    errors.append(f"{action_prefix} mutation requires a valid managed key")
                if decision_state != "STALE_MANAGED_OUTBOX":
                    errors.append(f"{action_prefix} mutation requires STALE_MANAGED_OUTBOX decision")
                if result_state != "REPLACED_MANAGED_OUTBOX":
                    errors.append(f"{action_prefix} mutation requires REPLACED_MANAGED_OUTBOX result")
                if not transport_allows_repair:
                    errors.append(f"{action_prefix} mutation requires repair-allowed transport")

            if stale_deleted and not replacement_created:
                errors.append(f"{action_prefix} stale delete requires replacement creation first")

            uncertainty_reason = str(action.get("uncertainty_reason", "")).strip()
            if uncertainty_reason and uncertainty_reason not in {"blocked-transport", "unmarked-user-mail"}:
                if result_state not in {"ACTION_REQUIRED", "FAILED_SAFE"}:
                    errors.append(f"{action_prefix} uncertain repair must become ACTION_REQUIRED or FAILED_SAFE")
                if replacement_created or stale_deleted or would_modify_fake_object:
                    errors.append(f"{action_prefix} uncertain repair must not mutate")

        expected_available = repair_case.get("repair_available_count")
        expected_taken = repair_case.get("repair_actions_taken")
        expected_failed_safe = repair_case.get("failed_safe_count")
        if expected_available != available_count:
            errors.append(f"{case_prefix} repair_available_count expected {available_count}, got {expected_available}")
        if expected_taken != taken_count:
            errors.append(f"{case_prefix} repair_actions_taken expected {taken_count}, got {expected_taken}")
        if expected_failed_safe != failed_safe_count:
            errors.append(f"{case_prefix} failed_safe_count expected {failed_safe_count}, got {expected_failed_safe}")
        if fake_mutations != computed_fake_mutations:
            errors.append(f"{case_prefix} fake_mutations_performed expected {computed_fake_mutations}, got {fake_mutations}")

    return errors


def validate_watchdog_repair_fixture(root: Path) -> list[str]:
    relative_path = "fixtures/watchdog/managed_repair_guardrail_fixture.json"
    path = root / relative_path
    if not path.exists():
        return [f"missing watchdog repair fixture: {relative_path}"]
    return validate_watchdog_repair_fixture_data(root, relative_path, load_json(path))


def csv_headers(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        return next(reader)


def validate_csv_headers(root: Path) -> list[str]:
    errors: list[str] = []
    for relative_path, required_headers in CSV_REQUIREMENTS.items():
        path = root / relative_path
        if not path.exists():
            errors.append(f"missing CSV file: {relative_path}")
            continue
        headers = csv_headers(path)
        missing = [header for header in required_headers if header not in headers]
        if missing:
            errors.append(f"{relative_path} missing headers: {', '.join(missing)}")
    return errors


def validate_manifest_template_paths(root: Path, manifest_relative_path: str) -> list[str]:
    errors: list[str] = []
    manifest_path = root / manifest_relative_path
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for index, row in enumerate(reader, start=2):
            template_path = (row.get("template_path") or "").strip()
            if not template_path:
                errors.append(f"{manifest_relative_path}:{index} has empty template_path")
                continue
            resolved = (manifest_path.parent / template_path).resolve()
            if not resolved.exists():
                errors.append(f"{manifest_relative_path}:{index} template does not exist: {template_path}")
    return errors


def bool_text_is_valid(value: str) -> bool:
    return value.strip().lower() in {"", "true", "yes", "y", "1", "on", "false", "no", "n", "0", "off"}


def split_asset_list(value: str) -> list[str]:
    normalized = value.replace(",", ";")
    return [item.strip() for item in normalized.split(";") if item.strip()]


def resolve_manifest_path(manifest_path: Path, path_text: str) -> Path:
    candidate = Path(path_text)
    if candidate.is_absolute() or PureWindowsPath(path_text).is_absolute():
        return candidate
    return manifest_path.parent / candidate


def validate_induction_template_manifest_assets(root: Path, manifest_relative_path: str) -> list[str]:
    errors: list[str] = []
    manifest_path = root / manifest_relative_path
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for index, row in enumerate(reader, start=2):
            mode = (row.get("mode") or "").strip().lower()
            if mode not in {"broadcast", "personal"}:
                errors.append(f"{manifest_relative_path}:{index} has unsupported mode: {mode}")

            delivery_action = (row.get("delivery_action") or "").strip().lower()
            if delivery_action not in {"", "display", "draft", "send"}:
                errors.append(f"{manifest_relative_path}:{index} has unsupported delivery_action: {delivery_action}")

            defer_minutes = (row.get("defer_minutes") or "").strip()
            if defer_minutes:
                if not defer_minutes.isdigit():
                    errors.append(f"{manifest_relative_path}:{index} defer_minutes is not a non-negative integer")

            urgent = row.get("urgent_verify_after_submit") or ""
            if not bool_text_is_valid(urgent):
                errors.append(f"{manifest_relative_path}:{index} urgent_verify_after_submit is not boolean-like")

            inline_folder = (row.get("inline_image_folder") or "").strip()
            if inline_folder:
                resolved_inline = resolve_manifest_path(manifest_path, inline_folder)
                if not resolved_inline.is_dir():
                    errors.append(f"{manifest_relative_path}:{index} inline image folder does not exist: {inline_folder}")

            attachments = split_asset_list(row.get("attachments") or "")
            attachment_base = (row.get("attachment_base_folder") or "").strip()
            if attachments and not attachment_base:
                errors.append(f"{manifest_relative_path}:{index} attachments require attachment_base_folder")
                continue
            if attachment_base:
                resolved_base = resolve_manifest_path(manifest_path, attachment_base)
                if not resolved_base.is_dir():
                    errors.append(f"{manifest_relative_path}:{index} attachment base folder does not exist: {attachment_base}")
                    continue
                for attachment in attachments:
                    if not (resolved_base / attachment).is_file():
                        errors.append(f"{manifest_relative_path}:{index} attachment does not exist: {attachment}")
    return errors


def release_excluded_directory_names(root: Path) -> set[str]:
    data = sanitisation_config(root).get("release_excluded_directory_names", [])
    if isinstance(data, list) and data:
        return {str(item).strip("/\\").lower() for item in data if str(item).strip("/\\")}
    return set(DEFAULT_RELEASE_EXCLUDED_DIRECTORY_NAMES)


def is_release_excluded_path(path: Path, root: Path) -> bool:
    try:
        relative_parts = path.resolve().relative_to(root.resolve()).parts
    except ValueError:
        return False
    if not relative_parts:
        return False
    return relative_parts[0].lower() in release_excluded_directory_names(root)


def iter_public_text_files(root: Path) -> Iterable[Path]:
    root = root.resolve()
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in TEXT_SUFFIXES:
            continue
        if is_release_excluded_path(path, root):
            continue
        if ".log" in path.name.lower():
            continue
        yield path


def sanitisation_config(root: Path) -> dict[str, object]:
    config_path = root / "fixtures" / "sanitisation" / "blocked_patterns.json"
    if not config_path.exists():
        return {}
    data = load_json(config_path)
    return data if isinstance(data, dict) else {}


def allowed_email_domains(root: Path) -> set[str]:
    data = sanitisation_config(root).get("allowed_email_domains", [])
    if isinstance(data, list):
        return {str(item).lower() for item in data}
    return set(DEFAULT_ALLOWED_EMAIL_DOMAINS)


def blocked_text_regexes(root: Path) -> list[re.Pattern[str]]:
    data = sanitisation_config(root).get("blocked_text_patterns", [])
    patterns = data if isinstance(data, list) else DEFAULT_BLOCKED_TEXT_PATTERNS
    return [re.compile(str(pattern), re.I) for pattern in patterns]


def validate_sanitisation(root: Path) -> list[str]:
    errors: list[str] = []
    allowed_domains = allowed_email_domains(root)
    blocked_patterns = blocked_text_regexes(root)
    for path in iter_public_text_files(root):
        text = path.read_text(encoding="utf-8", errors="ignore")
        for match in EMAIL_RE.finditer(text):
            if match.group(1).lower() not in allowed_domains:
                errors.append(f"{path.relative_to(root)} contains non-example email address")
                break
        for blocked_pattern in blocked_patterns:
            if blocked_pattern.search(text):
                errors.append(f"{path.relative_to(root)} contains private path or live SharePoint-style URL")
                break
    return errors


def validate_generated_artifacts(root: Path) -> list[str]:
    root = root.resolve()
    errors: list[str] = []
    config = sanitisation_config(root)
    suffixes = config.get("blocked_generated_suffixes", [])
    directory_names = config.get("blocked_generated_directory_names", [])
    blocked_suffixes = {str(item).lower() for item in suffixes} or set(DEFAULT_BLOCKED_GENERATED_SUFFIXES)
    blocked_directory_names = {str(item).lower() for item in directory_names} or set(DEFAULT_BLOCKED_GENERATED_DIRECTORY_NAMES)

    for path in root.rglob("*"):
        if is_release_excluded_path(path, root):
            continue
        relative = path.relative_to(root)
        if path.is_dir() and path.name.lower() in blocked_directory_names:
            errors.append(f"generated directory is present: {relative}")
        elif path.is_file() and path.suffix.lower() in blocked_suffixes:
            errors.append(f"generated or release-blocked artifact is present: {relative}")
    return errors


def is_relative_portable_path(path_text: object) -> bool:
    if not isinstance(path_text, str) or not path_text.strip():
        return False
    normalized = path_text.replace("\\", "/")
    if Path(path_text).is_absolute() or PureWindowsPath(path_text).is_absolute():
        return False
    return ".." not in PurePosixPath(normalized).parts


def validate_legacy_alternate_map(root: Path) -> list[str]:
    errors: list[str] = []
    map_path = root / "migration" / "legacy_alternate_map.json"
    if not map_path.exists():
        return ["missing legacy alternate map: migration/legacy_alternate_map.json"]

    data = load_json(map_path)
    if not isinstance(data, dict):
        return ["legacy alternate map is not a JSON object"]
    if data.get("schema") != LEGACY_ALTERNATE_MAP_SCHEMA:
        errors.append("legacy alternate map schema mismatch")

    policy = data.get("policy")
    if not isinstance(policy, dict):
        errors.append("legacy alternate map missing policy object")
    else:
        if policy.get("copying_into_scaffold_authorised") is not True:
            errors.append("legacy alternate map must explicitly allow copying into scaffold")
        for false_key in [
            "moving_legacy_files_allowed",
            "outside_root_edits_allowed",
            "raw_private_content_allowed",
            "live_send_allowed",
            "destructive_mailbox_actions_allowed",
        ]:
            if policy.get(false_key) is not False:
                errors.append(f"legacy alternate map policy must keep {false_key}=false")

    entries = data.get("entries")
    if not isinstance(entries, list) or not entries:
        errors.append("legacy alternate map must contain entries")
        return errors

    seen_scaffold_paths: set[str] = set()
    for index, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict):
            errors.append(f"legacy alternate map entry {index} is not an object")
            continue

        legacy_path = entry.get("legacy_path")
        scaffold_path = entry.get("scaffold_path")
        route_path = entry.get("route_path")
        status = entry.get("status")

        if not is_relative_portable_path(legacy_path):
            errors.append(f"legacy alternate map entry {index} has non-relative legacy_path")
        if not is_relative_portable_path(scaffold_path):
            errors.append(f"legacy alternate map entry {index} has non-relative scaffold_path")
        elif not (root / str(scaffold_path)).exists():
            errors.append(f"legacy alternate map scaffold target is missing: {scaffold_path}")
        else:
            seen_scaffold_paths.add(str(scaffold_path).replace("\\", "/"))

        if route_path is not None:
            if not is_relative_portable_path(route_path):
                errors.append(f"legacy alternate map entry {index} has non-relative route_path")
            elif not (root / str(route_path)).exists():
                errors.append(f"legacy alternate map route target is missing: {route_path}")

        if status not in LEGACY_ALTERNATE_STATUSES:
            errors.append(f"legacy alternate map entry {index} has unsupported status: {status}")

    for required in [
        "RunGroupMeetings.vbs",
        "RunInductionMailer.vbs",
        "OutlookWatchdog.vbs",
        "ValidateWorkflow.vbs",
    ]:
        if required not in seen_scaffold_paths:
            errors.append(f"legacy alternate map missing scaffold wrapper: {required}")

    return errors


def validate(root: Path) -> list[str]:
    root = root.resolve()
    errors: list[str] = []
    errors.extend(validate_workflow_schema(root))
    errors.extend(validate_shared_contract_fixture(root))
    errors.extend(validate_watchdog_report_fixture(root))
    errors.extend(validate_watchdog_repair_fixture(root))
    errors.extend(validate_dispatch_states(root))
    errors.extend(validate_dispatch_state_parity(root))
    errors.extend(validate_csv_headers(root))
    errors.extend(validate_manifest_template_paths(root, "examples/workbooks/induction_template_manifest_example.csv"))
    errors.extend(validate_manifest_template_paths(root, "vbs/tests/fixtures/induction_template_manifest_fixture.csv"))
    errors.extend(validate_induction_template_manifest_assets(root, "examples/workbooks/induction_template_manifest_example.csv"))
    errors.extend(validate_induction_template_manifest_assets(root, "vbs/tests/fixtures/induction_template_manifest_fixture.csv"))
    errors.extend(workflow_input_report.validate_default_examples(root))
    errors.extend(command_route_report.validate_default_routes(root))
    errors.extend(validate_sanitisation(root))
    errors.extend(validate_generated_artifacts(root))
    errors.extend(validate_legacy_alternate_map(root))
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate Office communication scaffold contracts and examples.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()

    errors = validate(args.root)
    if errors:
        for error in errors:
            print(f"FAILED: {error}")
        return 1

    print("Validation passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
