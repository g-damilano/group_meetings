from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Optional

from .com_support import require_win32com


OL_MAIL_ITEM = 0


@dataclass
class OutlookClient:
    """Thin Outlook COM client.

    Mirrors VBS GetOutlookApp/GetNamespace usage.
    """

    app: Any

    @classmethod
    def connect(cls) -> "OutlookClient":
        win32, pythoncom = require_win32com()
        # Try active Outlook instance first (GetObject in VBS)
        app = None
        try:
            app = win32.GetActiveObject("Outlook.Application")
        except Exception:
            app = None
        if app is None:
            # Create a new instance
            app = win32.gencache.EnsureDispatch("Outlook.Application")
        return cls(app=app)

    @property
    def namespace(self) -> Any:
        return self.app.GetNamespace("MAPI")

    def get_default_folder(self, folder_id: int) -> Any:
        """Return an Outlook default folder by numeric ID (e.g., 5=Sent Items)."""
        return self.namespace.GetDefaultFolder(int(folder_id))

    def create_mail_item(self) -> Any:
        return self.app.CreateItem(OL_MAIL_ITEM)

    def create_recipient(self, address: str) -> Any:
        return self.namespace.CreateRecipient(address)

    def resolve_recipient(self, address: str) -> Optional[Any]:
        """Resolve a recipient string; returns a Recipient COM object or None."""
        recip = self.create_recipient(address)
        try:
            ok = bool(recip.Resolve())
        except Exception:
            ok = False
        return recip if ok else None

    def get_address_list(self, name: str) -> Optional[Any]:
        try:
            return self.namespace.AddressLists(name)
        except Exception:
            return None

    def current_user_display_name(self) -> str:
        """Best-effort display name of the currently configured Outlook account."""
        try:
            return str(self.namespace.Accounts.Item(1).CurrentUser)
        except Exception:
            try:
                return str(self.namespace.CurrentUser)
            except Exception:
                return ""

    def get_exact_email(self, search_name: str) -> str:
        """Resolve a display name into a best-effort SMTP address (VBS GetExactEmail)."""
        search_name = (search_name or "").strip()
        if not search_name:
            return ""
        recip = self.resolve_recipient(search_name)
        if recip is None:
            return ""
        try:
            entry = recip.AddressEntry
        except Exception:
            entry = None
        if entry is not None:
            exch = None
            try:
                exch = entry.GetExchangeUser()
            except Exception:
                exch = None
            if exch is not None:
                try:
                    return str(exch.PrimarySmtpAddress or "").strip()
                except Exception:
                    pass
            try:
                return str(entry.Address or "").strip()
            except Exception:
                pass
        try:
            return str(recip.Address or "").strip()
        except Exception:
            return ""

    # -----------------------------
    # Mail lookup helpers
    # -----------------------------
    def mail_exists(self, subject: str, folder_id: int, check_only: bool = True) -> bool:
        """Return True if a MailItem with exact subject exists in a folder.

        Mirrors VBS EmailExists (exact subject match). Tries Items.Restrict first;
        falls back to sorting + iteration.
        """
        subject = (subject or "").strip()
        if not subject:
            return False

        folder = self.get_default_folder(folder_id)
        items = folder.Items

        # 1) Restrict (fast path)
        try:
            subj = subject.replace("'", "''")
            restricted = items.Restrict(f"[Subject] = '{subj}'")
            try:
                if int(restricted.Count) > 0:
                    return True
            except Exception:
                # some folders return non-int Count
                pass
        except Exception:
            pass

        # 2) Fallback: sort + scan
        sort_field = "ReceivedTime"
        if int(folder_id) == 4:
            sort_field = "CreationTime"
        elif int(folder_id) == 5:
            sort_field = "SentOn"

        try:
            items.Sort(sort_field, True)
        except Exception:
            pass

        for item in items:
            try:
                if str(item.Class) != "43":  # olMail = 43
                    continue
            except Exception:
                # Fallback to type name
                try:
                    if str(type(item).__name__) != "CDispatch":
                        pass
                except Exception:
                    pass
            try:
                if str(item.Subject) == subject:
                    return True
            except Exception:
                continue

        return False

    def mail_exists_in_sent_or_outbox(self, subject: str) -> bool:
        return self.mail_exists(subject, 5, True) or self.mail_exists(subject, 4, True)

    def has_sent_today(self, subject: str, today: Optional[date] = None) -> bool:
        """True if an email with exact subject was sent today (Sent Items only)."""
        subject = (subject or "").strip()
        if not subject:
            return False
        if today is None:
            today = date.today()

        folder = self.get_default_folder(5)
        items = folder.Items

        try:
            items.Sort("SentOn", True)
        except Exception:
            pass

        for mail in items:
            try:
                sent_on = mail.SentOn
            except Exception:
                continue
            try:
                sent_date = sent_on.date()
            except Exception:
                try:
                    # COM date sometimes behaves like datetime already
                    sent_date = date(sent_on.Year, sent_on.Month, sent_on.Day)
                except Exception:
                    continue

            # stop once we move into previous days (sorted descending)
            if sent_date < today:
                break
            try:
                if str(mail.Subject) == subject and sent_date == today:
                    return True
            except Exception:
                continue

        return False
