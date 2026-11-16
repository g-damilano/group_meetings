from __future__ import annotations

from typing import Any, List

from ..utils import safe_lower, safe_text


OL_DEFAULT_CONTACTS_FOLDER = 10
OL_OUTLOOK_ADDRESS_LIST = 4


def get_default_contacts_address_list_name(ns: Any) -> str:
    if ns is None:
        return ""
    try:
        folder = ns.GetDefaultFolder(OL_DEFAULT_CONTACTS_FOLDER)
    except Exception:
        folder = None
    if folder is None:
        return ""
    try:
        name = safe_text(getattr(folder, "AddressBookName", "")).strip()
        if not name:
            name = safe_text(getattr(folder, "Name", "")).strip()
        return name
    except Exception:
        return ""


def should_treat_as_local_contacts_list(address_list: Any) -> bool:
    if address_list is None:
        return False
    try:
        list_type = int(address_list.AddressListType)
    except Exception:
        list_type = -1
    if list_type == OL_OUTLOOK_ADDRESS_LIST:
        return True

    try:
        name = safe_text(address_list.Name)
    except Exception:
        name = ""
    lowered = safe_lower(name)
    return lowered in {"contacts", "contacts (this computer only)"}


def append_available_contact_list(target: List[str], ns: Any, list_name: str) -> None:
    name = safe_text(list_name).strip()
    if not name:
        return
    if name.lower() in {t.lower() for t in target}:
        return
    try:
        address_list = ns.AddressLists(name)
    except Exception:
        address_list = None
    if address_list is None:
        return
    if not should_treat_as_local_contacts_list(address_list):
        return
    target.append(name)


def collect_local_contacts_list_names(ns: Any) -> List[str]:
    out: List[str] = []
    if ns is None:
        return out
    default_name = get_default_contacts_address_list_name(ns)
    append_available_contact_list(out, ns, default_name)
    append_available_contact_list(out, ns, "Contacts")
    append_available_contact_list(out, ns, "Contacts (This computer only)")
    return out


def build_directory_search_list(ns: Any) -> List[str]:
    out: List[str] = []

    def add(name: str) -> None:
        n = safe_text(name).strip()
        if not n:
            return
        if n.lower() in {x.lower() for x in out}:
            return
        out.append(n)

    add("Global Address List")
    add("Offline Global Address List")

    for n in collect_local_contacts_list_names(ns):
        add(n)

    return out
