"""InductionMailer Python port.

This package ports a set of VBS/Outlook COM scripts into a reusable Python library.

High-level classes:
- :class:`induction_mailer.cli.InductionMailerApp` (console workflow)
- :class:`induction_mailer.outlook.client.OutlookClient` (Outlook COM)
- :class:`induction_mailer.outlook.composer.MailComposer` (compose/deliver)
- :class:`induction_mailer.template.repository.TemplateRepository` (load templates)
- :class:`induction_mailer.name_resolution.resolver.NameResolver` (GAL/Exchange name resolution)
"""

from .core_types import (
    ActionMode,
    DeliveryMode,
    MailOptions,
    RecipientPackage,
    RecipientResolutionEntry,
    ResolutionInfo,
    Template,
)

__all__ = [
    "ActionMode",
    "DeliveryMode",
    "MailOptions",
    "RecipientPackage",
    "RecipientResolutionEntry",
    "ResolutionInfo",
    "Template",
]
