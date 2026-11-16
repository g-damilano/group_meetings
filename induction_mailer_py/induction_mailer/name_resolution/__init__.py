"""Recipient/name resolution (GAL, Contacts)"""

from .resolver import NameResolver
from .prompt import prompt_recipient_resolution_package

__all__ = ["NameResolver", "prompt_recipient_resolution_package"]
