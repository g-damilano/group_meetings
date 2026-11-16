from __future__ import annotations

from typing import Any, List, Sequence

from ..utils import safe_lower, safe_text


def build_match_token_list(tokens: Sequence[str]) -> List[str]:
    out: List[str] = []
    for t in tokens or []:
        lowered = safe_lower(t)
        if lowered:
            out.append(lowered)
    return out


def contains_all_tokens(text: Any, tokens: Sequence[str]) -> bool:
    if not tokens:
        return False
    haystack = safe_lower(text)
    if not haystack:
        return False
    for t in tokens:
        needle = safe_lower(t)
        if not needle:
            return False
        if needle not in haystack:
            return False
    return True


def describe_item_class(item: Any) -> str:
    if item is None:
        return "unknown"
    try:
        class_value = getattr(item, "Class")
    except Exception:
        class_value = "?"
    try:
        message_class = safe_text(getattr(item, "MessageClass"))
    except Exception:
        message_class = ""
    if message_class:
        return f"Class={class_value}, MessageClass='{message_class}'"
    return f"Class={class_value}"


def is_contact_like_item(item: Any) -> bool:
    if item is None:
        return False
    try:
        class_value = int(getattr(item, "Class"))
    except Exception:
        class_value = -1
    try:
        message_class = safe_text(getattr(item, "MessageClass"))
    except Exception:
        message_class = ""
    if class_value in {40, 69}:
        return True
    lowered = message_class.lower()
    if not lowered:
        return False
    if lowered.startswith("ipm.contact"):
        return True
    if lowered == "ipm.distlist":
        return True
    return False


def contact_field_value(contact: Any, field_name: str) -> str:
    if contact is None:
        return ""
    field = (field_name or "").lower()
    try:
        if field == "fullname":
            return safe_text(contact.FullName)
        if field == "fileas":
            return safe_text(contact.FileAs)
        if field == "firstname":
            return safe_text(contact.FirstName)
        if field == "lastname":
            return safe_text(contact.LastName)
        if field == "email1address":
            return safe_text(contact.Email1Address)
        if field == "email2address":
            return safe_text(contact.Email2Address)
        if field == "email3address":
            return safe_text(contact.Email3Address)
        if field == "companyname":
            return safe_text(contact.CompanyName)
        if field == "department":
            return safe_text(contact.Department)
    except Exception:
        return ""
    return ""


def contact_item_matches_tokens(contact: Any, tokens: Sequence[str]) -> bool:
    if contact is None or not tokens:
        return False
    if not is_contact_like_item(contact):
        return False

    fields = [
        safe_lower(contact_field_value(contact, "FullName")),
        safe_lower(contact_field_value(contact, "FileAs")),
        safe_lower(contact_field_value(contact, "FirstName")),
        safe_lower(contact_field_value(contact, "LastName")),
        safe_lower(contact_field_value(contact, "Email1Address")),
        safe_lower(contact_field_value(contact, "Email2Address")),
        safe_lower(contact_field_value(contact, "Email3Address")),
        safe_lower(contact_field_value(contact, "CompanyName")),
        safe_lower(contact_field_value(contact, "Department")),
    ]

    for f in fields:
        if contains_all_tokens(f, tokens):
            return True

    first_lower = safe_lower(contact_field_value(contact, "FirstName"))
    last_lower = safe_lower(contact_field_value(contact, "LastName"))
    if len(tokens) == 2:
        t0 = safe_lower(tokens[0])
        t1 = safe_lower(tokens[1])
        if t0 in first_lower and t1 in last_lower:
            return True
        if t1 in first_lower and t0 in last_lower:
            return True

    return False
