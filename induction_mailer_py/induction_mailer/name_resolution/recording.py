from __future__ import annotations

from typing import Any, List, Optional

from ..utils import safe_text
from .directory_matching import describe_address_entry


_g_recorder: Optional[List[str]] = None


def set_lookup_response_recorder(recorder: Optional[List[str]]) -> None:
    global _g_recorder
    _g_recorder = recorder


def clear_lookup_response_recorder() -> None:
    global _g_recorder
    _g_recorder = None


def record_lookup_response(list_name: str, field_name: str, entry: Any) -> None:
    if _g_recorder is None:
        return
    _g_recorder.append(build_lookup_response_summary(list_name, field_name, entry))


def record_recipient_resolution_response(label: str, recipient: Any) -> None:
    if _g_recorder is None:
        return
    entry_summary = "(no address entry)"
    resolved_name = ""
    resolved_address = ""
    try:
        entry = recipient.AddressEntry if recipient is not None else None
    except Exception:
        entry = None
    if entry is not None:
        entry_summary = describe_address_entry(entry)
    try:
        resolved_name = safe_text(recipient.Name)
    except Exception:
        resolved_name = ""
    try:
        resolved_address = safe_text(recipient.Address)
    except Exception:
        resolved_address = ""

    detail = label.strip()
    if resolved_name:
        detail += f" -> Name: {resolved_name}"
    if resolved_address:
        detail += f", Address: {resolved_address}"
    detail += f" ({entry_summary})"
    _g_recorder.append(detail)


def build_lookup_response_summary(list_name: str, field_name: str, entry: Any) -> str:
    effective_list = safe_text(list_name).strip() or "(unknown list)"
    effective_field = safe_text(field_name).strip() or "(unknown field)"
    entry_summary = describe_address_entry(entry) or "(no entry information)"
    return f"List '{effective_list}' field '{effective_field}' => {entry_summary}"


def record_directory_attempt(info: Any, label: str) -> None:
    """Append a label to info.directory_attempted if not already present."""
    if info is None:
        return
    attempt = safe_text(label).strip()
    if not attempt:
        return

    existing = getattr(info, "directory_attempted", None)
    parts: List[str] = []

    if isinstance(existing, list):
        parts = [safe_text(p).strip() for p in existing if safe_text(p).strip()]
    elif isinstance(existing, str):
        parts = [p.strip() for p in existing.split(",") if p.strip()]

    if attempt.lower() in {p.lower() for p in parts}:
        return
    parts.append(attempt)
    info.directory_attempted = parts
