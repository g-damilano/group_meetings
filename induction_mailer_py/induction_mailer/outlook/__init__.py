"""Outlook COM wrappers (pywin32)."""

from .client import OutlookClient
from .composer import MailComposer

__all__ = ["OutlookClient", "MailComposer"]
