from __future__ import annotations

from typing import Any, Dict, List, Optional

from ..core_types import ResolutionInfo
from ..utils import safe_lower, safe_text
from .contacts import is_contact_like_item
from .directory_lists import build_directory_search_list
from .directory_matching import (
    choose_best_address_entry,
    collect_address_entries_by_fields,
    collect_address_entries_by_full_name,
)
from .recording import record_directory_attempt, record_lookup_response
from .text import (
    apply_email_candidate,
    apply_first_name_candidate,
    apply_full_name_candidate,
    get_accessor_string,
    normalize_email_text,
    parse_full_name_components,
)


def contact_matches_email(contact: Any, lowered_target: str) -> bool:
    if contact is None or not lowered_target:
        return False
    try:
        candidates = [
            safe_lower(getattr(contact, "Email1Address", "")),
            safe_lower(getattr(contact, "Email2Address", "")),
            safe_lower(getattr(contact, "Email3Address", "")),
        ]
    except Exception:
        candidates = []
    for c in candidates:
        if c and c == lowered_target:
            return True
    return False


def find_contact_in_address_list(address_list: Any, email_address: str) -> Any:
    """Iterate an address list and return the first contact item matching email.

    This can be expensive; it is intended for small local Contacts lists.
    """
    if address_list is None:
        return None
    target = safe_lower(normalize_email_text(email_address))
    if not target:
        return None
    try:
        entries = address_list.AddressEntries
    except Exception:
        return None

    for entry in entries:
        contact = None
        try:
            contact = entry.GetContact()
        except Exception:
            contact = None
        if contact is not None and contact_matches_email(contact, target):
            return contact
    return None


def populate_name_from_contact_item(contact: Any, info: ResolutionInfo) -> None:
    if contact is None:
        return
    try:
        full = safe_text(contact.FullName)
        if full:
            apply_full_name_candidate(info, full, info.email)
    except Exception:
        pass
    try:
        first = safe_text(contact.FirstName)
        if first:
            apply_first_name_candidate(info, first, "", info.email)
    except Exception:
        pass
    try:
        last = safe_text(contact.LastName)
        if last and not info.last:
            info.last = last
    except Exception:
        pass

    # emails
    for field in ["Email1Address", "Email2Address", "Email3Address"]:
        try:
            apply_email_candidate(info, getattr(contact, field, ""))
        except Exception:
            continue

    # additional via PropertyAccessor if available
    try:
        populate_name_from_property_accessor(contact, info, "", "")
    except Exception:
        pass


def populate_name_from_property_accessor(address_entry: Any, info: ResolutionInfo, alias_value: str, email_address: str) -> None:
    if address_entry is None:
        return
    try:
        accessor = address_entry.PropertyAccessor
    except Exception:
        accessor = None
    if accessor is None:
        return

    found_any = False

    # Given name
    given = get_accessor_string(
        accessor,
        [
            "http://schemas.microsoft.com/mapi/proptag/0x3A06001F",
            "http://schemas.microsoft.com/mapi/proptag/0x3A06001E",
        ],
    )
    if given:
        apply_first_name_candidate(info, given, alias_value, email_address)
        found_any = True

    # Surname
    surname = get_accessor_string(
        accessor,
        [
            "http://schemas.microsoft.com/mapi/proptag/0x3A11001F",
            "http://schemas.microsoft.com/mapi/proptag/0x3A11001E",
        ],
    )
    if surname and not info.last:
        info.last = surname
        found_any = True

    # Display
    display = get_accessor_string(
        accessor,
        [
            "http://schemas.microsoft.com/mapi/proptag/0x3001001F",
            "http://schemas.microsoft.com/mapi/proptag/0x3001001E",
        ],
    )
    if display:
        before = info.full
        apply_full_name_candidate(info, display, email_address)
        if safe_lower(info.full).strip() != safe_lower(before).strip():
            found_any = True

    # Email candidates
    email_value = get_accessor_string(
        accessor,
        [
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F",
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E",
            "http://schemas.microsoft.com/mapi/proptag/0x3003001F",
            "http://schemas.microsoft.com/mapi/proptag/0x3003001E",
        ],
    )
    apply_email_candidate(info, email_value)

    if not safe_text(info.email).strip():
        proxy_value = get_accessor_string(
            accessor,
            [
                "http://schemas.microsoft.com/mapi/proptag/0x39FF001F",
                "http://schemas.microsoft.com/mapi/proptag/0x39FF001E",
            ],
        )
        apply_email_candidate(info, proxy_value)

    if not safe_text(info.email).strip():
        try:
            apply_email_candidate(info, address_entry.Address)
        except Exception:
            pass

    if found_any:
        info.found = True


def _populate_from_exchange_user(exch: Any, info: ResolutionInfo) -> None:
    if exch is None:
        return
    try:
        apply_full_name_candidate(info, exch.Name, info.email)
    except Exception:
        pass
    try:
        apply_first_name_candidate(info, exch.FirstName, "", info.email)
    except Exception:
        pass
    try:
        last = safe_text(exch.LastName)
        if last and not info.last:
            info.last = last
    except Exception:
        pass
    try:
        apply_email_candidate(info, exch.PrimarySmtpAddress)
    except Exception:
        pass


def populate_name_from_address_lists(ns: Any, email_address: str, info: ResolutionInfo, verbose: bool = False) -> None:
    if ns is None:
        return
    normalized_email = normalize_email_text(email_address)
    if not normalized_email:
        return

    list_names = build_directory_search_list(ns)
    profile = None

    for name in list_names:
        try:
            address_list = ns.AddressLists(name)
        except Exception:
            address_list = None
        if address_list is None:
            continue

        record_directory_attempt(info, name)

        entry = None
        try:
            entry = address_list.AddressEntries.Find(f"[SMTPAddress] = '{normalized_email}'")
            if entry is not None:
                record_lookup_response(name, "SMTPAddress", entry)
        except Exception:
            entry = None

        if entry is None:
            try:
                entry = address_list.AddressEntries.Find(f"[EmailAddress] = '{normalized_email}'")
                if entry is not None:
                    record_lookup_response(name, "EmailAddress", entry)
            except Exception:
                entry = None

        if entry is None:
            contact = find_contact_in_address_list(address_list, normalized_email)
            if contact is not None:
                info.found = True
                info.directory_source = name
                populate_name_from_contact_item(contact, info)
                if info.found:
                    return

        if entry is not None:
            info.found = True
            info.directory_source = name
            # Try ExchangeUser
            exch = None
            try:
                exch = entry.GetExchangeUser()
            except Exception:
                exch = None
            _populate_from_exchange_user(exch, info)
            populate_name_from_property_accessor(entry, info, "", normalized_email)
            if info.found:
                return


def populate_name_from_address_lists_by_full_name(ns: Any, full_name: str, info: ResolutionInfo, email_address: str = "") -> None:
    if ns is None:
        return
    comps = parse_full_name_components(full_name)
    first = comps.get("First", "")
    last = comps.get("Last", "")
    if not first or not last:
        return

    profile = None
    try:
        from .directory_matching import get_current_user_profile_info

        profile = get_current_user_profile_info(ns)
    except Exception:
        profile = {}

    best_entry = None
    best_details: Dict[str, Any] = {}
    best_source = ""

    for list_name in build_directory_search_list(ns):
        try:
            address_list = ns.AddressLists(list_name)
        except Exception:
            continue
        record_directory_attempt(info, list_name)
        entries = collect_address_entries_by_full_name(address_list, first, last)
        if not entries:
            continue
        entry, details = choose_best_address_entry(entries, profile)
        if entry is None:
            continue
        best_entry = entry
        best_details = details
        best_source = list_name
        break

    if best_entry is None:
        return

    info.found = True
    info.directory_source = best_source
    info.match_details = best_details
    exch = None
    try:
        exch = best_entry.GetExchangeUser()
    except Exception:
        exch = None
    _populate_from_exchange_user(exch, info)
    populate_name_from_property_accessor(best_entry, info, "", email_address)


def populate_name_from_address_lists_by_display_name(ns: Any, display_name: str, info: ResolutionInfo, email_address: str = "") -> None:
    if ns is None:
        return
    value = safe_text(display_name).strip()
    if not value:
        return

    profile = None
    try:
        from .directory_matching import get_current_user_profile_info

        profile = get_current_user_profile_info(ns)
    except Exception:
        profile = {}

    # Try exact display-name like fields
    field_sets = [
        ["Name"],
        ["DisplayName"],
        ["FullName"],
    ]

    for list_name in build_directory_search_list(ns):
        try:
            address_list = ns.AddressLists(list_name)
        except Exception:
            continue
        record_directory_attempt(info, list_name)

        aggregate: List[Any] = []
        for fields in field_sets:
            aggregate.extend(collect_address_entries_by_fields(address_list, fields, value))
        if not aggregate:
            continue

        entry, details = choose_best_address_entry(aggregate, profile)
        if entry is None:
            continue
        info.found = True
        info.directory_source = list_name
        info.match_details = details

        exch = None
        try:
            exch = entry.GetExchangeUser()
        except Exception:
            exch = None
        _populate_from_exchange_user(exch, info)
        populate_name_from_property_accessor(entry, info, "", email_address)
        return


def attempt_recipient_resolve_for_value(ns: Any, raw_value: str, info: ResolutionInfo, descriptor: str = "") -> None:
    if ns is None:
        return
    value = safe_text(raw_value).strip()
    if not value:
        return
    try:
        recip = ns.CreateRecipient(value)
    except Exception:
        recip = None
    if recip is None:
        return
    try:
        ok = bool(recip.Resolve())
    except Exception:
        ok = False
    if not ok:
        return

    try:
        entry = recip.AddressEntry
    except Exception:
        entry = None

    info.found = True
    if descriptor and not info.directory_source:
        info.directory_source = descriptor

    if entry is not None:
        exch = None
        try:
            exch = entry.GetExchangeUser()
        except Exception:
            exch = None
        _populate_from_exchange_user(exch, info)
        populate_name_from_property_accessor(entry, info, "", info.email)


def attempt_alias_recipient_resolve(ns: Any, alias_value: str, info: ResolutionInfo, email_address: str = "") -> None:
    attempt_recipient_resolve_for_value(ns, alias_value, info, descriptor="AliasResolve")


def maybe_refine_name_using_given_name(email_address: str, info: ResolutionInfo) -> None:
    # If first name looks like alias/local-part, try to refine from full name or email
    if not info.first and info.full:
        comps = parse_full_name_components(info.full)
        if comps.get("First"):
            info.first = comps["First"]
        if comps.get("Last") and not info.last:
            info.last = comps["Last"]

    if not info.full and info.first and info.last:
        info.full = f"{info.first} {info.last}".strip()
