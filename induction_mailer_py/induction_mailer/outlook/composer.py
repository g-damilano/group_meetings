from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

from ..core_types import ActionMode, MailOptions
from ..logging_utils import log_event
from ..template.assets import sanitize_content_id
from ..utils import safe_text
from .client import OutlookClient


def _log(options: MailOptions, message: str) -> None:
    if options.log_handler is not None:
        options.log_handler(options.log_path, message)
    elif options.log_path:
        log_event(options.log_path, message)


PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# Outlook constants
OL_BY_VALUE = 1
OL_MSG = 3


@dataclass
class MailComposer:
    """Compose emails and deliver them through Outlook.

    This is a Python port of the VBS `ComposeAndSendMail` + inline-image support.
    """

    client: OutlookClient

    def compose(self, options: MailOptions) -> Any:
        mail = self.client.create_mail_item()

        if options.to:
            mail.To = options.to
        if options.cc:
            mail.CC = options.cc
        if options.bcc:
            mail.BCC = options.bcc
        if options.subject:
            mail.Subject = options.subject

        if options.defer_minutes and int(options.defer_minutes) > 0:
            try:
                mail.DeferredDeliveryTime = datetime.now() + timedelta(minutes=int(options.defer_minutes))
            except Exception:
                pass

        # Body
        if options.prefer_plain_text and safe_text(options.body_text).strip():
            try:
                mail.Body = safe_text(options.body_text)
            except Exception:
                # fall back to HTML
                html = self._build_html(options)
                try:
                    mail.HTMLBody = html
                except Exception:
                    mail.Body = safe_text(options.body_html or options.body_text)
        else:
            html = self._build_html(options)
            try:
                mail.HTMLBody = html
            except Exception:
                # fallback plain body
                mail.Body = safe_text(options.body_html)

        # Attach files (non-inline)
        for path in (options.attachments or []):
            p = safe_text(path).strip()
            if not p:
                continue
            try:
                mail.Attachments.Add(p)
                _log(options, f"Attachment added: {p}")
            except Exception as e:
                _log(options, f"Failed to add attachment: {p} ({e})")

        # Inline image folder
        if options.inline_image_folder:
            self._attach_inline_folder(mail, options)

        return mail

    def execute(self, mail: Any, options: MailOptions) -> None:
        action = ActionMode(options.action)
        if action == ActionMode.DISPLAY:
            _log(options, "Displaying email")
            mail.Display()
            return
        if action == ActionMode.SEND:
            _log(options, "Sending email")
            mail.Send()
            return
        if action == ActionMode.SAVE:
            save_path = safe_text(options.save_path).strip()
            if not save_path:
                save_path = str(Path.cwd() / "message.msg")
            Path(save_path).parent.mkdir(parents=True, exist_ok=True)
            _log(options, f"Saving email to: {save_path}")
            mail.SaveAs(save_path, OL_MSG)
            return

        raise ValueError(f"Unknown action: {options.action}")

    def compose_and_execute(self, options: MailOptions) -> None:
        mail = self.compose(options)
        self.execute(mail, options)

    def _build_html(self, options: MailOptions) -> str:
        body = safe_text(options.body_html)

        # body transforms
        for t in options.body_transforms or []:
            try:
                body = t(body)
            except Exception:
                continue

        signature = ""
        if options.signature_literal:
            signature = safe_text(options.signature_literal)
        elif options.signature_provider is not None:
            try:
                signature = safe_text(options.signature_provider(self.client))
            except Exception:
                signature = ""

        head_content = safe_text(options.head_content)
        body_attr = safe_text(options.body_attributes)

        if "<html" in body.lower():
            # assume it's already a full document
            return body

        return (
            "<!doctype html>\n"
            "<html>\n"
            "<head>\n"
            + head_content
            + "\n</head>\n"
            + f"<body {body_attr}>\n"
            + body
            + ("\n" + signature if signature else "")
            + "\n</body>\n"
            "</html>"
        )

    def _attach_inline_folder(self, mail: Any, options: MailOptions) -> None:
        folder = Path(options.inline_image_folder)
        if not folder.exists() or not folder.is_dir():
            return

        for f in sorted(folder.iterdir()):
            if not f.is_file():
                continue
            try:
                cid = sanitize_content_id(f.name)
                if options.inline_image_cid_transform is not None:
                    cid = safe_text(options.inline_image_cid_transform(cid)) or cid

                att = mail.Attachments.Add(str(f), OL_BY_VALUE)
                try:
                    att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                except Exception:
                    pass
                try:
                    att.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)
                except Exception:
                    pass

                _log(options, f"Inline image attached: {f.name} (cid:{cid})")
            except Exception as e:
                _log(options, f"Failed inline image attach: {f} ({e})")
