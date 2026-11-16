from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tempfile import gettempdir
from typing import Optional

from ..core_types import ActionMode, MailOptions
from ..logging_utils import WeeklyFileLogger
from ..outlook.client import OutlookClient
from ..outlook.composer import MailComposer

from .config import GroupMeetingsConfig
from .content import (
    build_lab_manager_reminder_subject,
    build_meeting_lead_reminder_subject,
    build_meeting_subject,
    build_presenter_reminder_subject,
    build_weekly_responsibility_reminder_subject,
    generate_email_alert,
    generate_lab_manager_reminder_email,
    generate_main_email_content,
    generate_meeting_lead_reminder_email,
    generate_presenter_reminder_email,
    generate_weekly_responsibility_reminder_email,
)
from .sync import rclone_copy
from .tentative_rota import build_tentative_rota_bundle
from .utils import next_fridays
from .workbook import MeetingWorkbook


@dataclass
class _RowContext:
    row: int
    meeting_date: date


class GroupMeetingsManager:
    """Python port of group_meetings_manager.vbs."""

    def __init__(self, config: Optional[GroupMeetingsConfig] = None):
        self.config = config or GroupMeetingsConfig()
        self.logger = WeeklyFileLogger(
            log_path=self.config.log_path,
            variant=self.config.log_variant,
            extended_logging=bool(self.config.extended_logging),
        )

        self.outlook: Optional[OutlookClient] = None
        self.composer: Optional[MailComposer] = None

    # -----------------------------
    # Entrypoint
    # -----------------------------
    def run(self) -> None:
        self.logger.initialize()

        # Sync workbook (best effort)
        self._sync_workbook_best_effort()

        # Outlook
        try:
            self.outlook = OutlookClient.connect()
            self.composer = MailComposer(client=self.outlook)
            self.logger.log("Outlook application created successfully.")
            self.logger.log("Obtained Outlook namespace for MAPI.")
        except Exception as e:
            self.logger.error("Failed to initialize Outlook", e)
            self._cleanup_synced_workbook()
            self.logger.log("=== Script execution completed ===")
            return

        try:
            self._coordinator()
        finally:
            self._cleanup_synced_workbook()
            self.logger.log("=== Script execution completed ===")

    # -----------------------------
    # Workflow
    # -----------------------------
    def _coordinator(self) -> None:
        assert self.outlook is not None
        self.logger.verbose("Coordinator invoked")
        self.logger.log("Coordinator routine started")

        fridays = next_fridays(count=3)
        meeting_subject = build_meeting_subject(fridays[0])

        # VBS: send regular email only if not in Sent Items
        should_send_regular = not self.outlook.mail_exists(meeting_subject, 5, True)
        if not should_send_regular:
            self.logger.log(
                "Existing meeting email detected in Sent Items. Regular message will be skipped, reminders will still be evaluated."
            )

        with MeetingWorkbook(
            path=self.config.excel_path,
            sheet_rota=self.config.sheet_rota,
            sheet_emails=self.config.sheet_emails,
        ) as wb:
            row_indexes = wb.find_rows_for_dates(fridays)
            self.logger.log(
                f"Rota rows located | This: {row_indexes[0]} | Next: {row_indexes[1]} | Next+1: {row_indexes[2]}"
            )

            self._evaluate_critical_windows(fridays, row_indexes, wb)
            self._handle_communications(fridays, row_indexes, wb, should_send_regular)

        self.logger.log("Coordinator routine completed and Excel resources released.")

    # -----------------------------
    # Critical alerts
    # -----------------------------
    def _send_critical_alert(self, *, alert_type: str, target_date: date, recipient: str, wb: MeetingWorkbook) -> None:
        assert self.outlook is not None
        subject = f"ALERT: {alert_type} for {format_date(target_date)}"
        if self.outlook.has_sent_today(subject):
            self.logger.log(f"Alert '{subject}' already dispatched today. Skipping duplicate send.")
            return

        tentative_text = ""
        attachment = ""
        if alert_type in {"MissingDate", "Unassigned"}:
            names = wb.eligible_names_for_tentative_rota()
            csv_path = Path(gettempdir()) / "TentativeRota.csv"
            bundle = build_tentative_rota_bundle(names, csv_path)
            tentative_text = bundle.body_text
            attachment = str(bundle.csv_path) if bundle.csv_path.exists() else ""

        body = generate_email_alert(
            alert_type=alert_type,
            missing_date_text=format_date(target_date),
            recipient_email=recipient,
            tentative_rota=tentative_text,
            building_admin_email=self.config.building_admin_email,
        )

        self.logger.log(
            f"Dispatching alert email | Type: {alert_type} | Date: {format_date(target_date)} | Recipient: {recipient}"
        )
        self._send_plain_email(
            to_address=recipient,
            subject=subject,
            body=body,
            attachment_path=attachment,
            action=ActionMode.SEND,
        )

        # negative feedback loop (optional)
        if not self.config.debug_mode:
            if self.outlook.mail_exists(subject, 5, True):
                self.logger.log(f"Negative feedback loop confirmed for alert subject: {subject}")
            else:
                self.logger.log(f"WARNING: Alert subject not observed in Sent Items after sending: {subject}")

    def _evaluate_critical_windows(self, fridays: list[date], row_indexes: list[int], wb: MeetingWorkbook) -> None:
        self.logger.verbose("EvaluateCriticalWindows invoked")
        for idx, dt in enumerate(fridays):
            row_num = row_indexes[idx]
            if row_num == -1:
                self._send_critical_alert(alert_type="MissingDate", target_date=dt, recipient=self.config.manager_email, wb=wb)
                continue

            entry = wb.get_rota_entry(row_num)
            self.logger.log(
                f"Row {row_num} data | Lit: {entry.literature} | Res: {entry.results} | Alt: {entry.alternate} | Room: {entry.room}"
            )

            if not entry.literature and not entry.results and not entry.alternate:
                self._send_critical_alert(
                    alert_type="Unassigned", target_date=dt, recipient=self.config.supervisor_email, wb=wb
                )
            if not entry.room:
                self._send_critical_alert(alert_type="MissingRoom", target_date=dt, recipient=self.config.manager_email, wb=wb)

    # -----------------------------
    # Communications
    # -----------------------------
    def _handle_communications(
        self,
        fridays: list[date],
        row_indexes: list[int],
        wb: MeetingWorkbook,
        should_send_regular_email: bool,
    ) -> None:
        assert self.outlook is not None
        self.logger.verbose("HandleCommunications invoked")
        current_row = row_indexes[0]
        if current_row == -1:
            self.logger.log("No rota entry found for this Friday. Regular email and reminders will not be processed.")
            return

        entry = wb.get_rota_entry(current_row)
        meeting_lead_name = self._derive_meeting_lead_name(wb, fridays[0])

        if not entry.literature and not entry.results and not entry.alternate:
            self.logger.log(
                f"No presenters or alternate assigned for {format_date(fridays[0])}. Regular meeting email suppressed."
            )
            manager_alert_subject = f"ALERT: MeetingNotSet for {format_date(fridays[0])}"
            if not self.outlook.has_sent_today(manager_alert_subject):
                self.logger.log(
                    f"Sending manager alert that the meeting is not set for {format_date(fridays[0])}"
                )
                body = generate_email_alert(
                    alert_type="MeetingNotSet",
                    missing_date_text=format_date(fridays[0]),
                    recipient_email=self.config.manager_email,
                    tentative_rota="",
                    building_admin_email=self.config.building_admin_email,
                )
                self._send_plain_email(
                    to_address=self.config.manager_email,
                    subject=manager_alert_subject,
                    body=body,
                    attachment_path="",
                    action=ActionMode.SEND,
                )
            else:
                self.logger.log(
                    f"Manager alert already exists for {format_date(fridays[0])}. Skipping duplicate notification."
                )
            return

        self._trigger_reminder_emails(
            wb=wb,
            literature_presenter=entry.literature,
            results_presenter=entry.results,
            lab_manager_name=entry.lab_manager,
            meeting_date=fridays[0],
        )

        header = build_meeting_subject(fridays[0])
        if should_send_regular_email:
            # Next friday context (row_indexes[1])
            next_entry = None
            if row_indexes[1] != -1:
                next_entry = wb.get_rota_entry(row_indexes[1])
            next_friday_date = fridays[1]
            next_lead = self._derive_meeting_lead_name(wb, next_friday_date)

            body = generate_main_email_content(
                sender_display_name=self.outlook.current_user_display_name(),
                meeting_date=fridays[0],
                room=entry.room,
                literature=entry.literature,
                results=entry.results,
                alternate=entry.alternate,
                lab_manager=entry.lab_manager,
                meeting_lead_name=meeting_lead_name,
                next_friday_date=next_friday_date,
                next_literature=(next_entry.literature if next_entry else ""),
                next_results=(next_entry.results if next_entry else ""),
                next_alternate=(next_entry.alternate if next_entry else ""),
                next_lab_manager=(next_entry.lab_manager if next_entry else ""),
                next_meeting_lead_name=next_lead,
            )
            self.logger.log(f"Dispatching regular meeting email for {format_date(fridays[0])}")
            self._send_plain_email(
                to_address=self.config.email_to,
                subject=header,
                body=body,
                attachment_path="",
                action=ActionMode.SEND,
            )
        else:
            self.logger.log(f"Regular meeting email already exists. Skipping dispatch for subject: {header}")

    def _trigger_reminder_emails(
        self,
        *,
        wb: MeetingWorkbook,
        literature_presenter: str,
        results_presenter: str,
        lab_manager_name: str,
        meeting_date: date,
    ) -> None:
        self.logger.verbose("TriggerReminderEmails invoked")
        self.logger.log(f"Triggering reminder workflow for meeting date {format_date(meeting_date)}")
        self._maybe_send_presenter_reminder(literature_presenter, "Literature", meeting_date)
        self._maybe_send_presenter_reminder(results_presenter, "Results", meeting_date)
        self._maybe_send_weekly_responsibility_reminder(lab_manager_name, meeting_date)

    # -----------------------------
    # Reminder primitives
    # -----------------------------
    def _resolve_email_or_manager(self, person_name: str) -> str:
        assert self.outlook is not None
        resolved = self.outlook.get_exact_email(person_name)
        if resolved.strip():
            self.logger.log(f"Resolved exact email for '{person_name}': {resolved}")
            return resolved
        self.logger.log(f"Falling back to manager email for '{person_name}'")
        return self.config.manager_email

    def _maybe_send_presenter_reminder(self, presenter_name: str, presentation_type: str, meeting_date: date) -> None:
        assert self.outlook is not None
        name = (presenter_name or "").strip()
        if not name:
            self.logger.log(f"Skipping {presentation_type} presenter reminder because no presenter is assigned.")
            return
        subject = build_presenter_reminder_subject(presentation_type, meeting_date)
        if self.outlook.mail_exists_in_sent_or_outbox(subject):
            self.logger.log(
                f"Skipping presenter reminder because Sent Items or Outbox already contain subject '{subject}'."
            )
            return
        to_addr = self._resolve_email_or_manager(name)
        body = generate_presenter_reminder_email(
            sender_display_name=self.outlook.current_user_display_name(),
            presenter_name=name,
            presentation_type=presentation_type,
            meeting_date=meeting_date,
        )
        self._send_plain_email(to_address=to_addr, subject=subject, body=body, attachment_path="", action=ActionMode.SEND)

    def _inventory_manager_cc_for(self, to_address: str) -> str:
        inventory_email = self.config.inventory_manager_email.strip()
        if not inventory_email:
            return ""
        if to_address.strip().casefold() == inventory_email.casefold():
            return ""
        return inventory_email

    def _maybe_send_weekly_responsibility_reminder(self, lab_manager_name: str, meeting_date: date) -> None:
        assert self.outlook is not None
        name = (lab_manager_name or "").strip()
        if not name:
            self.logger.log(
                f"WARNING: No lab manager assigned for {format_date(meeting_date)}. "
                "Weekly responsibility reminder will fall back to the manager address."
            )
        subject = build_weekly_responsibility_reminder_subject(meeting_date)
        if self.outlook.mail_exists_in_sent_or_outbox(subject):
            self.logger.log(
                f"Skipping weekly responsibility reminder because Sent Items or Outbox already contain subject '{subject}'."
            )
            return
        to_addr = self._resolve_email_or_manager(name)
        body = generate_weekly_responsibility_reminder_email(
            sender_display_name=self.outlook.current_user_display_name(),
            responsible_name=name,
            meeting_date=meeting_date,
            inventory_manager_name=self.config.inventory_manager_name,
            inventory_manager_email=self.config.inventory_manager_email,
            inventory_feedback_cc_email=self.config.manager_email,
            inventory_items=[],
        )
        self._send_plain_email(
            to_address=to_addr,
            cc_address=self._inventory_manager_cc_for(to_addr),
            subject=subject,
            body=body,
            attachment_path="",
            action=ActionMode.SEND,
        )

    def _maybe_send_meeting_lead_reminder(self, meeting_lead_name: str, meeting_date: date) -> None:
        assert self.outlook is not None
        name = (meeting_lead_name or "").strip()
        subject = build_meeting_lead_reminder_subject(meeting_date)
        if self.outlook.mail_exists_in_sent_or_outbox(subject):
            self.logger.log(
                f"Skipping meeting lead reminder because Sent Items or Outbox already contain subject '{subject}'."
            )
            return
        to_addr = self._resolve_email_or_manager(name)
        body = generate_meeting_lead_reminder_email(
            sender_display_name=self.outlook.current_user_display_name(),
            meeting_lead_name=name,
            meeting_date=meeting_date,
        )
        self._send_plain_email(to_address=to_addr, subject=subject, body=body, attachment_path="", action=ActionMode.SEND)

    def _maybe_send_lab_manager_reminder(self, lab_manager_name: str, meeting_date: date) -> None:
        assert self.outlook is not None
        name = (lab_manager_name or "").strip()
        subject = build_lab_manager_reminder_subject(meeting_date)
        if self.outlook.mail_exists_in_sent_or_outbox(subject):
            self.logger.log(
                f"Skipping lab manager reminder because Sent Items or Outbox already contain subject '{subject}'."
            )
            return
        to_addr = self._resolve_email_or_manager(name)
        body = generate_lab_manager_reminder_email(
            sender_display_name=self.outlook.current_user_display_name(),
            lab_manager_name=name,
            meeting_date=meeting_date,
        )
        self._send_plain_email(to_address=to_addr, subject=subject, body=body, attachment_path="", action=ActionMode.SEND)

    # -----------------------------
    # Helpers
    # -----------------------------
    def _derive_meeting_lead_name(self, wb: MeetingWorkbook, meeting_date: date) -> str:
        """Port of VBS DeriveMeetingLeadName.

        Uses the current Friday's Lab Responsible (col F) as chair/timekeeper.
        """
        row = wb.find_row_for_date(meeting_date)
        current_lab_manager = ""
        if row != -1:
            current_lab_manager = wb.get_rota_entry(row).lab_manager.strip()
        self.logger.log(
            f"Derived meeting lead as current lab responsible '{current_lab_manager}' for meeting date {format_date(meeting_date)}"
        )
        return current_lab_manager

    def _send_plain_email(
        self,
        *,
        to_address: str,
        cc_address: str = "",
        subject: str,
        body: str,
        attachment_path: str,
        action: ActionMode,
    ) -> None:
        assert self.composer is not None

        opts = MailOptions(
            to=to_address,
            cc=cc_address,
            subject=subject,
            body_text=body,
            prefer_plain_text=True,
            action=ActionMode.DISPLAY if self.config.debug_mode else action,
            attachments=[attachment_path] if attachment_path else [],
            log_path=str(self.config.log_path),
            log_handler=lambda path, msg: self.logger.log(msg),
        )

        try:
            mail = self.composer.compose(opts)
            if self.config.debug_mode:
                self.logger.log(f"DEBUG_MODE enabled - email not sent. Subject: {subject}")
                mail.Display()
            else:
                self.logger.log(f"Sending email to {to_address} with subject: {subject}")
                self.composer.execute(mail, opts)
        except Exception as e:
            self.logger.error(f"Failed to send email subject='{subject}'", e)

    def _sync_workbook_best_effort(self) -> None:
        self.logger.verbose("SyncMeetingWorkbook invoked")
        ok = rclone_copy(
            rclone_exe=self.config.rclone_exe,
            source=self.config.rclone_source,
            dest_dir=self.config.rclone_dest_dir,
            logger=self.logger,
        )
        if ok:
            self.logger.log("Workbook sync operation succeeded.")
        else:
            self.logger.log("Workbook sync operation failed. Please check rclone configuration.")

    def _cleanup_synced_workbook(self) -> None:
        if not self.config.delete_local_workbook_after_run:
            return
        path = Path(self.config.excel_path)
        try:
            if path.exists():
                self.logger.log(f"Deleting local Excel file after sync: {path}")
                path.unlink()
            else:
                self.logger.log(f"Local Excel file not found during cleanup: {path}")
        except Exception as e:
            self.logger.error(f"Failed to delete local workbook: {path}", e)


def format_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")
