from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, List, Optional

from ..core_types import ResolutionInfo
from ..utils import safe_lower, safe_text
from ..validation import is_valid_email
from ..outlook.client import OutlookClient
from .directory_population import (
    populate_name_from_address_lists,
    populate_name_from_address_lists_by_display_name,
    populate_name_from_address_lists_by_full_name,
)
from .recording import record_recipient_resolution_response
from .text import (
    apply_email_candidate,
    apply_full_name_candidate,
    extract_display_name_and_email,
    normalize_email_text,
    parse_full_name_components,
)


@dataclass
class NameResolver:
    """Resolve a person from an input string using Outlook directories.

    This wraps the bulk of the VBS `MNS_*` logic in a single OOP entry point.
    """

    client: OutlookClient

    def resolve(self, raw_value: str, lookup_field: str = "auto") -> ResolutionInfo:
        info = ResolutionInfo(found=False, preferred_address=safe_text(raw_value).strip())
        if not info.preferred_address:
            return info

        ns = self.client.namespace

        display, email = extract_display_name_and_email(info.preferred_address)
        if email:
            apply_email_candidate(info, email)
        elif is_valid_email(info.preferred_address):
            apply_email_candidate(info, info.preferred_address)

        # 1) Attempt Recipient.Resolve (alias/email/full string)
        if lookup_field.lower() in {"auto", "email"} and info.email:
            if self._attempt_recipient_resolve(ns, info.email, info, descriptor="Email"):
                return info

        if lookup_field.lower() in {"auto", "display", "fullname"}:
            candidate = display or info.preferred_address
            if candidate:
                if self._attempt_recipient_resolve(ns, candidate, info, descriptor="Display"):
                    return info

        # 2) Directory search
        if info.email:
            populate_name_from_address_lists(ns, normalize_email_text(info.email), info)
            if info.found:
                return info

        components = parse_full_name_components(display or info.preferred_address)
        full_name = components.get("Display") or info.preferred_address
        if lookup_field.lower() in {"auto", "fullname"}:
            populate_name_from_address_lists_by_full_name(ns, full_name, info, email_address=info.email)
            if info.found:
                return info

        populate_name_from_address_lists_by_display_name(ns, full_name, info, email_address=info.email)
        if info.found:
            return info

        # 3) fallback: set full from display, or guess from email local
        if not info.full:
            apply_full_name_candidate(info, full_name, info.email)

        if not info.first and not info.last and info.email:
            # naive: local part split on '.' '_' '-' and take first two tokens
            local = safe_text(info.email).split("@", 1)[0]
            parts = [p for p in re.split(r"[._\-]+", local) if p]
            if parts:
                info.first = parts[0]
                if len(parts) > 1:
                    info.last = parts[1]
                info.full = (info.first + " " + info.last).strip()

        return info

    def resolve_many(self, raw_values: List[str], lookup_field: str = "auto") -> List[ResolutionInfo]:
        return [self.resolve(v, lookup_field=lookup_field) for v in raw_values]

    def _attempt_recipient_resolve(self, ns: Any, value: str, info: ResolutionInfo, descriptor: str) -> bool:
        value = safe_text(value).strip()
        if not value:
            return False
        try:
            recip = ns.CreateRecipient(value)
        except Exception:
            return False

        try:
            ok = bool(recip.Resolve())
        except Exception:
            ok = False

        record_recipient_resolution_response(f"Resolve attempt ({descriptor}) '{value}'", recip)

        if not ok:
            return False

        try:
            entry = recip.AddressEntry
        except Exception:
            entry = None
        if entry is None:
            return False

        info.found = True
        try:
            info.directory_source = "Recipient.Resolve"
        except Exception:
            pass

        # populate info via AddressEntry
        from .directory_population import populate_name_from_property_accessor

        exch = None
        try:
            exch = entry.GetExchangeUser()
        except Exception:
            exch = None
        if exch is not None:
            try:
                apply_full_name_candidate(info, exch.Name, info.email)
            except Exception:
                pass
            try:
                info.first = info.first or safe_text(exch.FirstName)
            except Exception:
                pass
            try:
                info.last = info.last or safe_text(exch.LastName)
            except Exception:
                pass
            try:
                apply_email_candidate(info, exch.PrimarySmtpAddress)
            except Exception:
                pass

        populate_name_from_property_accessor(entry, info, alias_value=value, email_address=info.email)

        try:
            info.full = info.full or safe_text(recip.Name)
        except Exception:
            pass

        return True
