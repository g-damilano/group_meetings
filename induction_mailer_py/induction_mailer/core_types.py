from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional


class DeliveryMode(str, Enum):
    BROADCAST = "broadcast"
    PERSONAL = "personal"


class ActionMode(str, Enum):
    DISPLAY = "display"
    SEND = "send"
    SAVE = "save"


@dataclass
class Template:
    """A parsed email template.

    Mirrors the VBS dictionary shape:
    - Path, Folder, FileName, Meta, Body, DisplayName, Stage, Code
    """

    path: Path
    body: str
    meta: Dict[str, str] = field(default_factory=dict)
    display_name: str = ""
    stage: str = ""
    code: str = ""

    @property
    def folder(self) -> Path:
        return self.path.parent

    @property
    def file_name(self) -> str:
        return self.path.name


@dataclass
class ResolutionInfo:
    """Result of resolving a person/address.

    Port of the VBS `CreateEmptyResolutionInfo()` dictionary.
    """

    # Input / fallback address used during resolution (alias/email/display name).
    preferred_address: str = ""

    # Resolution results
    found: bool = False
    first: str = ""
    last: str = ""
    full: str = ""
    email: str = ""
    directory_source: str = ""
    directory_attempted: List[str] = field(default_factory=list)

    def ensure_full(self) -> None:
        if not self.full:
            combined = (self.first + " " + self.last).strip()
            if combined:
                self.full = combined


@dataclass
class RecipientResolutionEntry:
    """One recipient provided by the user (potentially resolved via GAL/Exchange)."""

    raw_value: str
    lookup_field: str = "auto"  # auto, email, first, surname, display, full...
    preferred_address: str = ""
    resolution: ResolutionInfo = field(default_factory=ResolutionInfo)


@dataclass
class RecipientPackage:
    """A set of recipient inputs + their resolved outputs."""

    # Mirrors the VBScript package shape used by both scripts.
    lookup_field: str = "auto"
    raw_values: List[str] = field(default_factory=list)
    resolved_addresses: List[str] = field(default_factory=list)
    entries: List[ResolutionInfo] = field(default_factory=list)

    def count(self) -> int:
        return len(self.entries)

    def addresses(self) -> List[str]:
        # Back-compat helper: if callers didn't precompute addresses, derive them.
        if self.resolved_addresses:
            return list(self.resolved_addresses)
        out: List[str] = []
        for e in self.entries:
            addr = (e.email or e.preferred_address).strip()
            if addr:
                out.append(addr)
        return out


SignatureProvider = Callable[[Any], str]
BodyTransform = Callable[[str], str]
InlineCidTransform = Callable[[str], str]


@dataclass
class MailOptions:
    """Options fed to the composer (similar to VBS dictionary)."""

    to: str = ""
    cc: str = ""
    bcc: str = ""
    subject: str = ""
    body_html: str = ""
    # Optional plain-text body. If provided and `prefer_plain_text` is True,
    # the composer will set Outlook MailItem.Body instead of HTMLBody.
    body_text: str = ""
    prefer_plain_text: bool = False
    action: ActionMode = ActionMode.DISPLAY
    save_path: str = ""
    defer_minutes: int = 0
    log_path: str = ""

    head_content: str = ""
    body_attributes: str = ""
    body_transforms: List[BodyTransform] = field(default_factory=list)
    inline_image_folder: str = ""
    inline_image_cid_transform: Optional[InlineCidTransform] = None
    signature_provider: Optional[SignatureProvider] = None
    signature_literal: str = ""

    attachments: List[str] = field(default_factory=list)

    # Optional logging callback: (log_path, message) -> None
    log_handler: Optional[Callable[[str, str], None]] = None
