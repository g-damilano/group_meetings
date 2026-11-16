"""Group meetings manager (ported from group_meetings_manager.vbs).

This package reuses the same Outlook COM abstractions as InductionMailer.
"""

from .config import GroupMeetingsConfig
from .manager import GroupMeetingsManager
from .workbook import RotaPerson

__all__ = ["GroupMeetingsConfig", "GroupMeetingsManager", "RotaPerson"]
