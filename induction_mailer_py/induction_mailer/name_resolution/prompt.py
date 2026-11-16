from __future__ import annotations

from typing import List

from ..core_types import RecipientPackage
from ..utils import safe_text
from .resolver import NameResolver


def prompt_recipient_resolution_package(resolver: NameResolver) -> RecipientPackage:
    """Interactive prompt mirroring VBS `PromptRecipientResolutionPackage`.

    Returns a RecipientPackage containing raw values plus resolved entries.
    """
    print("\nRecipient input")
    print("  Enter recipients separated by ';' or ',' (e.g. emails, aliases, 'Name <email>')")
    raw_line = input("Recipients: ").strip()

    raw_values: List[str] = []
    if raw_line:
        raw_values = [v.strip() for v in raw_line.replace(",", ";").split(";") if v.strip()]

    print("\nLookup field")
    print("  1) auto\n  2) email\n  3) fullname\n  4) display")
    choice = input("Lookup (1-4, default 1): ").strip()
    lookup_field = "auto"
    if choice == "2":
        lookup_field = "email"
    elif choice == "3":
        lookup_field = "fullname"
    elif choice == "4":
        lookup_field = "display"

    resolved = resolver.resolve_many(raw_values, lookup_field=lookup_field)
    addresses = [
        safe_text(i.email or i.preferred_address).strip()
        for i in resolved
        if safe_text(i.email or i.preferred_address).strip()
    ]

    return RecipientPackage(
        lookup_field=lookup_field,
        raw_values=raw_values,
        resolved_addresses=addresses,
        entries=resolved,
    )
