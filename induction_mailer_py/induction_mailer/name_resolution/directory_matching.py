from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple

from ..utils import safe_lower, safe_text
from .text import get_accessor_string


def describe_address_entry(entry: Any) -> str:
    if entry is None:
        return "(none)"
    try:
        name_part = safe_text(entry.Name)
    except Exception:
        name_part = ""
    try:
        addr_part = safe_text(entry.Address)
    except Exception:
        addr_part = ""
    summary = name_part.strip() or "(unnamed entry)"
    if addr_part:
        summary += f" <{addr_part}>"
    return summary


def build_address_entry_tracking_key(entry: Any) -> str:
    if entry is None:
        return ""
    try:
        key = safe_lower(entry.ID)
    except Exception:
        key = ""
    if not key:
        key = (safe_lower(getattr(entry, "Name", "")) + "|" + safe_lower(getattr(entry, "Address", ""))).strip("|")
    return key


def add_address_entry_to_collection(target: List[Any], entry: Any, seen: Optional[set] = None) -> None:
    if entry is None:
        return
    if seen is None:
        target.append(entry)
        return
    key = build_address_entry_tracking_key(entry)
    if key and key in seen:
        return
    if key:
        seen.add(key)
    target.append(entry)


def merge_address_entry_matches(target: List[Any], matches: List[Any], seen: Optional[set], source_map: Dict[str, str], list_name: str) -> None:
    for e in matches or []:
        before = len(target)
        add_address_entry_to_collection(target, e, seen)
        if len(target) > before:
            key = build_address_entry_tracking_key(e)
            if key and key not in source_map:
                source_map[key] = list_name


def escape_restriction_value(value: str) -> str:
    return safe_text(value).replace("'", "''")


def collect_address_entries_by_filters(address_list: Any, filters: Sequence[str]) -> List[Any]:
    results: List[Any] = []
    if address_list is None:
        return results
    try:
        entries = address_list.AddressEntries
    except Exception:
        return results

    seen: set = set()
    for filt in filters:
        try:
            entry = entries.Find(filt)
        except Exception:
            entry = None
        while entry is not None:
            add_address_entry_to_collection(results, entry, seen)
            try:
                entry = entries.FindNext()
            except Exception:
                entry = None
    return results


def collect_address_entries_by_fields(address_list: Any, field_names: Sequence[str], search_value: str) -> List[Any]:
    if not search_value:
        return []
    escaped = escape_restriction_value(search_value)
    filters = [f"[{f}] = '{escaped}'" for f in field_names]
    return collect_address_entries_by_filters(address_list, filters)


def find_address_entry_by_fields(address_list: Any, field_names: Sequence[str], search_value: str) -> Any:
    matches = collect_address_entries_by_fields(address_list, field_names, search_value)
    return matches[0] if matches else None


def collect_address_entries_by_full_name(address_list: Any, first_name: str, last_name: str) -> List[Any]:
    f = safe_text(first_name).strip()
    l = safe_text(last_name).strip()
    if not f or not l:
        return []
    ef = escape_restriction_value(f)
    el = escape_restriction_value(l)
    filters = [
        f"[FirstName] = '{ef}' AND [LastName] = '{el}'",
        f"[GivenName] = '{ef}' AND [Surname] = '{el}'",
        f"[GivenName] = '{ef}' AND [LastName] = '{el}'",
        f"[FirstName] = '{ef}' AND [Surname] = '{el}'",
    ]
    return collect_address_entries_by_filters(address_list, filters)


def find_address_entry_by_full_name(address_list: Any, first_name: str, last_name: str) -> Any:
    matches = collect_address_entries_by_full_name(address_list, first_name, last_name)
    return matches[0] if matches else None


def build_address_entry_candidate(entry: Any) -> Optional[Dict[str, Any]]:
    if entry is None:
        return None
    cand: Dict[str, Any] = {
        "Entry": entry,
        "Department": "",
        "Company": "",
        "MatchDetails": [],
    }

    exch = None
    try:
        exch = entry.GetExchangeUser()
    except Exception:
        exch = None
    if exch is not None:
        cand["Exchange"] = exch
        try:
            cand["Department"] = safe_text(exch.Department).strip()
        except Exception:
            cand["Department"] = ""
        try:
            cand["Company"] = safe_text(exch.CompanyName).strip()
        except Exception:
            cand["Company"] = ""

    if not cand.get("Department") or not cand.get("Company"):
        try:
            accessor = entry.PropertyAccessor
        except Exception:
            accessor = None
        if accessor is not None:
            if not cand.get("Department"):
                cand["Department"] = safe_text(
                    get_accessor_string(
                        accessor,
                        [
                            "http://schemas.microsoft.com/mapi/proptag/0x3A18001F",
                            "http://schemas.microsoft.com/mapi/proptag/0x3A18001E",
                        ],
                    )
                ).strip()
            if not cand.get("Company"):
                cand["Company"] = safe_text(
                    get_accessor_string(
                        accessor,
                        [
                            "http://schemas.microsoft.com/mapi/proptag/0x3A16001F",
                            "http://schemas.microsoft.com/mapi/proptag/0x3A16001E",
                        ],
                    )
                ).strip()

    return cand


def caller_profile_has_org_details(profile: Dict[str, str]) -> bool:
    if not profile:
        return False
    dept = safe_text(profile.get("Department", "")).strip()
    comp = safe_text(profile.get("Company", "")).strip()
    return bool(dept or comp)


def score_entry_candidate(candidate: Dict[str, Any], profile: Dict[str, str]) -> int:
    if not candidate or not profile or not caller_profile_has_org_details(profile):
        return 0
    dept = safe_text(candidate.get("Department", "")).strip().lower()
    comp = safe_text(candidate.get("Company", "")).strip().lower()
    pdept = safe_text(profile.get("Department", "")).strip().lower()
    pcomp = safe_text(profile.get("Company", "")).strip().lower()

    score = 0
    details: List[str] = []
    if pdept and dept and pdept == dept:
        score += 10
        details.append("Department match")
    if pcomp and comp and pcomp == comp:
        score += 5
        details.append("Company match")

    if details:
        candidate.setdefault("MatchDetails", []).extend(details)
    return score


def get_current_user_profile_info(ns: Any) -> Dict[str, str]:
    profile: Dict[str, str] = {"HasProfile": "False", "Department": "", "Company": ""}
    if ns is None:
        return profile
    try:
        curr_user = ns.CurrentUser
    except Exception:
        curr_user = None
    if curr_user is None:
        return profile
    try:
        addr_entry = curr_user.AddressEntry
    except Exception:
        addr_entry = None
    if addr_entry is None:
        return profile

    dept = ""
    comp = ""
    exch = None
    try:
        exch = addr_entry.GetExchangeUser()
    except Exception:
        exch = None
    if exch is not None:
        try:
            dept = safe_text(exch.Department)
        except Exception:
            dept = ""
        try:
            comp = safe_text(exch.CompanyName)
        except Exception:
            comp = ""

    if (not dept.strip()) or (not comp.strip()):
        try:
            accessor = addr_entry.PropertyAccessor
        except Exception:
            accessor = None
        if accessor is not None:
            if not dept.strip():
                dept = safe_text(
                    get_accessor_string(
                        accessor,
                        [
                            "http://schemas.microsoft.com/mapi/proptag/0x3A18001F",
                            "http://schemas.microsoft.com/mapi/proptag/0x3A18001E",
                        ],
                    )
                )
            if not comp.strip():
                comp = safe_text(
                    get_accessor_string(
                        accessor,
                        [
                            "http://schemas.microsoft.com/mapi/proptag/0x3A16001F",
                            "http://schemas.microsoft.com/mapi/proptag/0x3A16001E",
                        ],
                    )
                )

    profile["Department"] = dept.strip()
    profile["Company"] = comp.strip()
    if profile["Department"] or profile["Company"]:
        profile["HasProfile"] = "True"
    return profile


def choose_best_address_entry(entries: List[Any], profile: Dict[str, str] | None = None) -> Tuple[Optional[Any], Dict[str, Any]]:
    """Choose a best entry using profile org scoring.

    Returns (entry, details)
    """
    if not entries:
        return None, {}
    profile = profile or {}
    best_entry = entries[0]
    best_cand = build_address_entry_candidate(best_entry) or {}
    best_score = score_entry_candidate(best_cand, profile)

    for e in entries[1:]:
        cand = build_address_entry_candidate(e)
        if cand is None:
            continue
        score = score_entry_candidate(cand, profile)
        if score > best_score:
            best_entry = e
            best_score = score
            best_cand = cand

    details = {
        "Score": best_score,
        "Department": safe_text(best_cand.get("Department", "")),
        "Company": safe_text(best_cand.get("Company", "")),
        "MatchDetails": list(best_cand.get("MatchDetails", [])),
    }
    return best_entry, details
