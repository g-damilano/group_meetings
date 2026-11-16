from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Optional

from ..utils import safe_text, safe_lower


def collect_name_tokens(text: Any) -> List[str]:
    raw = safe_text(text)
    trimmed = raw.strip()
    if not trimmed:
        return []
    return [t for t in trimmed.split(" ") if t.strip()]


def collect_flexible_name_tokens(text: Any) -> List[str]:
    normalized = safe_text(text).strip()
    if not normalized:
        return []
    for ch in [",", ".", "_", "-", "/", "\\", "\t", chr(160)]:
        normalized = normalized.replace(ch, " ")
    while "  " in normalized:
        normalized = normalized.replace("  ", " ")
    return collect_name_tokens(normalized)


_EMAIL_CLEAN_RE = re.compile(r"[\s\"'<>]" )


def normalize_email_text(value: Any) -> str:
    text = safe_text(value).strip()
    if not text:
        return ""
    text = _EMAIL_CLEAN_RE.sub("", text)
    return text


def extract_display_name_and_email_info(text: Any) -> Dict[str, Any]:
    """Parse strings like "Display Name <email@domain>".

    Returns a VBScript-style dict (useful for diagnostics and parity with VBS helpers).
    """
    raw = safe_text(text).strip()
    result: Dict[str, Any] = {
        "Display": raw,
        "Email": "",
        "HasAngleFormat": False,
    }
    if not raw:
        return result

    start = raw.find("<")
    end = raw.rfind(">")
    if start >= 0 and end > start:
        between = raw[start + 1 : end]
        email_candidate = normalize_email_text(between)
        if email_candidate:
            result["Email"] = email_candidate
            result["HasAngleFormat"] = True
            before_part = raw[:start].strip()
            after_part = raw[end + 1 :].strip()
            merged = before_part or after_part
            if merged and after_part and before_part:
                merged = f"{before_part} {after_part}".strip()
            result["Display"] = merged.strip()
    return result


def extract_display_name_and_email(text: Any) -> tuple[str, str]:
    """Compatibility helper used by NameResolver.

    Returns (display, email). This avoids dict-unpacking pitfalls.
    """
    info = extract_display_name_and_email_info(text)
    return safe_text(info.get("Display", "")).strip(), safe_text(info.get("Email", "")).strip()


def join_tokens(items: Iterable[Any], delimiter: str) -> str:
    parts: List[str] = []
    for v in items:
        t = safe_text(v).strip()
        if t:
            parts.append(t)
    return delimiter.join(parts)


def build_reversed_token_display(tokens: List[str]) -> str:
    if not tokens:
        return ""
    rev = list(tokens)
    rev.reverse()
    return " ".join(rev)


def add_candidate_value(target: Dict[str, bool], value: Any) -> None:
    v = safe_text(value).strip()
    if not v:
        return
    target[v] = True


def parse_full_name_components(full_name: Any) -> Dict[str, str]:
    text = safe_text(full_name).strip()
    components: Dict[str, str] = {
        "Display": text,
        "Normalized": text,
        "ReverseDisplay": "",
        "First": "",
        "Last": "",
        "AlternateFirst": "",
        "AlternateLast": "",
    }
    if not text:
        return components

    normalized = text.replace(",", " ")
    while "  " in normalized:
        normalized = normalized.replace("  ", " ")
    components["Normalized"] = normalized.strip()

    tokens = collect_flexible_name_tokens(normalized)
    if len(tokens) >= 2:
        components["First"] = tokens[0]
        components["Last"] = tokens[-1]
        components["ReverseDisplay"] = f"{tokens[-1]} {tokens[0]}".strip()
    if "," in text:
        # guess Last, First
        parts = [p.strip() for p in text.split(",") if p.strip()]
        if len(parts) >= 2:
            components["AlternateLast"] = parts[0]
            components["AlternateFirst"] = parts[1].split(" ")[0] if parts[1] else ""
    return components


def extract_email_local_part(email_address: Any) -> str:
    email = safe_text(email_address)
    if "@" in email:
        return email.split("@", 1)[0]
    return email


def apply_email_candidate(info: "ResolutionInfo", candidate: Any) -> None:
    from ..core_types import ResolutionInfo

    if not isinstance(info, ResolutionInfo):
        return
    cand = normalize_email_text(candidate)
    if not cand:
        return
    if not info.email:
        info.email = cand
    else:
        # prefer something that looks like SMTP
        if "@" in cand and "@" not in info.email:
            info.email = cand


def apply_full_name_candidate(info: "ResolutionInfo", candidate: Any, email_address: Any = "") -> None:
    from ..core_types import ResolutionInfo

    if not isinstance(info, ResolutionInfo):
        return
    cand = safe_text(candidate).strip()
    if not cand:
        return
    # prefer candidate that contains the email local part
    if not info.full:
        info.full = cand
        return
    alias = extract_email_local_part(email_address)
    if alias and alias.lower() in cand.lower() and alias.lower() not in info.full.lower():
        info.full = cand


def should_replace_first_name(existing_value: str, new_value: str, alias_value: str, email_address: str) -> bool:
    existing = existing_value.strip()
    new = new_value.strip()
    if not existing:
        return True
    if not new:
        return False

    alias = (alias_value or extract_email_local_part(email_address)).strip()
    if alias:
        # prefer name that is not just the alias
        if existing.lower() == alias.lower() and new.lower() != alias.lower():
            return True

    # prefer longer / more human-looking
    if len(new) > len(existing):
        return True
    return False


def apply_first_name_candidate(info: "ResolutionInfo", candidate: Any, alias_value: Any = "", email_address: Any = "") -> None:
    from ..core_types import ResolutionInfo

    if not isinstance(info, ResolutionInfo):
        return
    cand = safe_text(candidate).strip()
    if not cand:
        return
    if should_replace_first_name(info.first, cand, safe_text(alias_value), safe_text(email_address)):
        info.first = cand


def get_accessor_string(accessor: Any, property_uris: List[str]) -> str:
    """Try a list of MAPI property URIs via PropertyAccessor.GetProperty.

    This mirrors MNS_text.GetAccessorString but is intentionally best-effort.
    """
    if accessor is None:
        return ""
    for uri in property_uris:
        try:
            val = accessor.GetProperty(uri)
            s = safe_text(val).strip()
            if s:
                return s
        except Exception:
            continue
    return ""
