from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional


class DispatchState(str, Enum):
    NOT_REQUESTED = "NOT_REQUESTED"
    COMPOSED = "COMPOSED"
    DISPLAYED = "DISPLAYED"
    DRAFTED = "DRAFTED"
    SEND_REQUESTED = "SEND_REQUESTED"
    QUEUED_OUTBOX = "QUEUED_OUTBOX"
    SENT_CONFIRMED = "SENT_CONFIRMED"
    DUPLICATE_SUPPRESSED = "DUPLICATE_SUPPRESSED"
    WAITING_DEFERRED = "WAITING_DEFERRED"
    BLOCKED_OFFLINE = "BLOCKED_OFFLINE"
    BLOCKED_AUTH = "BLOCKED_AUTH"
    STALE_MANAGED_OUTBOX = "STALE_MANAGED_OUTBOX"
    REPLACED_MANAGED_OUTBOX = "REPLACED_MANAGED_OUTBOX"
    ACTION_REQUIRED = "ACTION_REQUIRED"
    FAILED_SAFE = "FAILED_SAFE"


class DeliveryAction(str, Enum):
    DISPLAY = "display"
    DRAFT = "draft"
    SEND = "send"
    SAVE = "save"


@dataclass
class DeliveryPolicy:
    action: DeliveryAction = DeliveryAction.DISPLAY
    defer_minutes: int = 0
    urgent_verify_after_submit: bool = False
    managed_dispatch_required: bool = True


@dataclass
class RecipientSpec:
    raw_value: str
    display_name: str = ""
    smtp_address: str = ""
    resolution_status: str = "unresolved"
    directory_source: str = ""
    fallback_used: bool = False


@dataclass
class MessageSpec:
    message_key: str
    message_type: str
    to: str
    subject: str
    body_text: str = ""
    body_html: str = ""
    cc: str = ""
    bcc: str = ""
    delivery_policy: DeliveryPolicy = field(default_factory=DeliveryPolicy)
    attachments: List[str] = field(default_factory=list)
    inline_image_folder: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class DispatchResult:
    state: DispatchState
    reason: str = ""
    managed_key: str = ""
    folder: str = ""
    outlook_entry_id: str = ""
    action_required: bool = False
    created_new_mail: bool = False
    replaced_existing_mail: bool = False
    safe_to_retry: bool = False


@dataclass
class WorkflowContext:
    workflow_name: str
    run_id: str
    source_workbook: Optional[Path] = None
    template_root: Optional[Path] = None
    log_path: Optional[Path] = None
    dry_run: bool = False
    operator: str = ""
    settings: Dict[str, Any] = field(default_factory=dict)


@dataclass
class HealthReport:
    offline: Optional[bool] = None
    exchange_connection_mode: Optional[int] = None
    exchange_connection_mode_name: str = ""
    blocked_or_auth_suspected: bool = False
    outbox_count: int = 0
    managed_outbox_count: int = 0
    unmanaged_outbox_count: int = 0
    action_required_messages: List[str] = field(default_factory=list)


def build_managed_email_key(org_code: str, workflow_name: str, logical_date: str, message_type: str, target_key: str) -> str:
    return "|".join(
        _normalize_key_part(part)
        for part in (org_code, workflow_name, logical_date, message_type, target_key)
    )


def _normalize_key_part(value: str) -> str:
    text = str(value or "").strip().upper().replace("_", "-").replace(" ", "-")
    while "--" in text:
        text = text.replace("--", "-")
    return text

