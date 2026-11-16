from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from .core_types import ActionMode, DeliveryMode, MailOptions, ResolutionInfo, Template
from .group_meetings.config import GroupMeetingsConfig
from .group_meetings.sync import rclone_copyto
from .group_meetings.workbook import MeetingWorkbook, RotaPerson
from .template.repository import TemplateRepository
from .template.placeholders import build_context_from_meta, ensure_placeholder_values
from .template.link_map import load_link_map_context, apply_link_map
from .template.assets import prepare_inline_image_support, resolve_attachment_paths
from .utils import capitalize_name_part, safe_text
from .outlook.client import OutlookClient
from .outlook.composer import MailComposer
from .name_resolution.resolver import NameResolver
from .name_resolution.prompt import prompt_recipient_resolution_package


@dataclass
class InductionMailerApp:
    templates_root: Path
    link_map_file: str = "links_map.txt"
    log_path: Optional[str] = "induction_mailer.log"

    def run_interactive(self) -> int:
        print("\n=== InductionMailer (Python) ===\n")

        task = self._prompt_workflow_task()
        if not task:
            print("Operation cancelled.")
            return 0

        if task == "rota_people":
            client = OutlookClient.connect()
            resolver = NameResolver(client)
            package = prompt_recipient_resolution_package(resolver)
            if not package.entries:
                print("No valid people supplied.")
                return 0
            return self._add_recipients_to_meeting_rota_people(package)

        repo = TemplateRepository(self.templates_root)
        templates = repo.load_templates()
        if not templates:
            print(f"No templates found under: {self.templates_root}")
            return 2

        link_map = load_link_map_context(self.templates_root / self.link_map_file)

        # Mode
        mode = self._prompt_mode()
        action = self._prompt_action()

        # Outlook + name resolver
        client = OutlookClient.connect()
        resolver = NameResolver(client)
        package = prompt_recipient_resolution_package(resolver)

        # Template selection
        selected = self._prompt_template_selection(templates)
        if not selected:
            print("No templates selected.")
            return 0

        # Dispatch
        composer = MailComposer(client)

        for tmpl in selected:
            self._send_template(
                composer=composer,
                tmpl=tmpl,
                mode=mode,
                action=action,
                recipients=package.entries,
                link_map=link_map,
            )

        print("\nDone.")
        return 0

    def _prompt_workflow_task(self) -> str:
        print("Induction task")
        print("  1) prepare induction email(s)")
        print("  2) add person to group meeting rota people sheet")
        choice = input("Task (1-2, default 1): ").strip().lower()
        if choice in {"", "1", "email", "emails"}:
            return "emails"
        if choice in {"2", "rota", "people", "rota_people"}:
            return "rota_people"
        if choice in {"q", "quit", "cancel"}:
            return ""
        print("Invalid option.")
        return self._prompt_workflow_task()

    def _prompt_mode(self) -> DeliveryMode:
        print("Delivery mode")
        print("  1) broadcast  (one email -> many recipients)")
        print("  2) personal   (one email per recipient with [NAME]/[SURNAME])")
        choice = input("Mode (1-2, default 1): ").strip()
        return DeliveryMode.PERSONAL if choice == "2" else DeliveryMode.BROADCAST

    def _prompt_action(self) -> ActionMode:
        print("\nAction")
        print("  1) display  (open drafts)")
        print("  2) send     (send immediately)")
        print("  3) save     (save .msg to disk)")
        choice = input("Action (1-3, default 1): ").strip()
        if choice == "2":
            return ActionMode.SEND
        if choice == "3":
            return ActionMode.SAVE
        return ActionMode.DISPLAY

    def _prompt_template_selection(self, templates: List[Template]) -> List[Template]:
        print("\nTemplates")
        for i, t in enumerate(templates, start=1):
            stage = t.stage or "(no stage)"
            code = t.code or ""
            print(f"  {i:>2}) {stage:20} {code:8} {t.display_name}")

        raw = input("Select templates (e.g. '1', '1,3-5', 'all'): ").strip().lower()
        if not raw:
            return []
        if raw in {"all", "a"}:
            return templates

        indexes: set[int] = set()
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                try:
                    start = int(a)
                    end = int(b)
                except ValueError:
                    continue
                for k in range(min(start, end), max(start, end) + 1):
                    indexes.add(k)
            else:
                try:
                    indexes.add(int(part))
                except ValueError:
                    continue

        selected: List[Template] = []
        for k in sorted(indexes):
            if 1 <= k <= len(templates):
                selected.append(templates[k - 1])
        return selected

    def _send_template(
        self,
        composer: MailComposer,
        tmpl: Template,
        mode: DeliveryMode,
        action: ActionMode,
        recipients: List[ResolutionInfo],
        link_map: Dict[str, str],
    ) -> None:
        meta = tmpl.meta

        # Build base context (meta + dynamic link map)
        ctx = build_context_from_meta(meta, link_map_context=link_map)

        subject_raw = safe_text(meta.get("subject", ""))
        cc_raw = safe_text(meta.get("cc", ""))
        bcc_raw = safe_text(meta.get("bcc", ""))
        attachments_raw = safe_text(meta.get("attachments", ""))

        # Body (link-map first; placeholders later)
        body_html = apply_link_map(tmpl.body, link_map)

        # In the VBScript flow, recipient lookup populates variables such as NAME/SURNAME/EMAIL
        # which are then available for placeholder replacement. Mirror that behaviour here.
        recipient_keys = {
            "NAME", "SURNAME", "FIRSTNAME", "LASTNAME",
            "FULL_NAME", "FULLNAME",
            "EMAIL", "ALIAS", "USERNAME",
            "DIRECTORY_SOURCE", "DIRECTORY_ATTEMPTED",
        }

        # If we're broadcasting to exactly one recipient, treat it like VBS: populate recipient variables
        # so templates that include [NAME]/[SURNAME]/etc don't require prompting.
        if mode == DeliveryMode.BROADCAST and len(recipients) == 1:
            self._inject_recipient_context(ctx, recipients[0])

        # Prompt for any placeholders we *cannot* infer (shared step).
        # For PERSONAL mode, skip recipient-derived keys here; per-recipient prompting happens later.
        shared_skip = recipient_keys if mode == DeliveryMode.PERSONAL else None
        ensure_placeholder_values(
            ctx,
            [
                apply_link_map(subject_raw, link_map),
                apply_link_map(cc_raw, link_map),
                apply_link_map(bcc_raw, link_map),
                body_html,
                attachments_raw,
            ],
            prompt_fn=input,
            skip_keys=shared_skip,
        )

        from .template.placeholders import apply_context_to_text

        if mode == DeliveryMode.BROADCAST:
            # Render once
            to = "; ".join([r.email or r.preferred_address for r in recipients if (r.email or r.preferred_address)])

            subject = apply_context_to_text(apply_link_map(subject_raw, link_map), ctx)
            cc = apply_context_to_text(apply_link_map(cc_raw, link_map), ctx)
            bcc = apply_context_to_text(apply_link_map(bcc_raw, link_map), ctx)

            rendered_body = apply_context_to_text(body_html, ctx)
            rendered_body, inline_folder = prepare_inline_image_support(rendered_body, tmpl.path)
            attachments = resolve_attachment_paths(attachments_raw, tmpl.folder)

            opts = MailOptions(
                to=to,
                cc=cc,
                bcc=bcc,
                subject=subject,
                body_html=rendered_body,
                action=action.value,
                attachments=attachments,
                inline_image_folder=inline_folder or None,
                log_path=self.log_path,
            )
            composer.compose_and_execute(opts)
            return

        # PERSONAL: one email per recipient with recipient-derived placeholders
        for r in recipients:
            per_ctx = ctx.copy()
            self._inject_recipient_context(per_ctx, r)

            # Prompt only for placeholders that are still missing after recipient injection
            ensure_placeholder_values(
                per_ctx,
                [
                    apply_link_map(subject_raw, link_map),
                    apply_link_map(cc_raw, link_map),
                    apply_link_map(bcc_raw, link_map),
                    body_html,
                    attachments_raw,
                ],
                prompt_fn=input,
            )

            subj_i = apply_context_to_text(apply_link_map(subject_raw, link_map), per_ctx)
            cc_i = apply_context_to_text(apply_link_map(cc_raw, link_map), per_ctx)
            bcc_i = apply_context_to_text(apply_link_map(bcc_raw, link_map), per_ctx)

            body_i = apply_context_to_text(body_html, per_ctx)
            body_i, inline_folder = prepare_inline_image_support(body_i, tmpl.path)

            attachments = resolve_attachment_paths(attachments_raw, tmpl.folder)

            opts = MailOptions(
                to=r.email or r.preferred_address,
                cc=cc_i,
                bcc=bcc_i,
                subject=subj_i,
                body_html=body_i,
                action=action.value,
                attachments=attachments,
                inline_image_folder=inline_folder or None,
                log_path=self.log_path,
            )
            composer.compose_and_execute(opts)

    def _add_recipients_to_meeting_rota_people(self, package) -> int:
        cfg = GroupMeetingsConfig()

        print("\nDownloading the latest meeting rota workbook...")
        if not rclone_copyto(
            rclone_exe=cfg.rclone_exe,
            source=cfg.rclone_source,
            dest=cfg.excel_path,
        ):
            print("Could not download the meeting rota workbook.")
            return 2

        added_count = 0
        duplicate_count = 0
        skipped_count = 0

        try:
            with MeetingWorkbook(
                path=cfg.excel_path,
                sheet_rota=cfg.sheet_rota,
                sheet_emails=cfg.sheet_emails,
                editable=True,
            ) as wb:
                for raw, resolved in zip(package.raw_values, package.entries):
                    person = self._build_rota_person(raw, resolved)
                    if person is None:
                        skipped_count += 1
                        continue
                    if not self._confirm_rota_person(person):
                        skipped_count += 1
                        continue
                    status = wb.append_rota_person_if_missing(person)
                    if status == "added":
                        added_count += 1
                        print(f"  Added to people sheet: {person.full_name}")
                    elif status == "duplicate":
                        duplicate_count += 1
                        print(f"  Already present in the people sheet: {person.full_name}")
                    else:
                        skipped_count += 1

                if added_count > 0:
                    wb.save()
        except Exception as exc:
            print(f"Could not update the meeting rota workbook: {exc}")
            return 2

        if added_count > 0:
            print("Uploading the edited meeting rota workbook...")
            if not rclone_copyto(
                rclone_exe=cfg.rclone_exe,
                source=cfg.excel_path,
                dest=cfg.rclone_source,
            ):
                print(f"The local workbook was saved, but upload failed. Local copy kept at: {cfg.excel_path}")
                return 2

        try:
            if cfg.excel_path.exists():
                cfg.excel_path.unlink()
        except Exception:
            pass

        print(f"Added: {added_count} | Already present: {duplicate_count} | Skipped: {skipped_count}")
        return 0

    def _build_rota_person(self, raw_value: str, resolved: ResolutionInfo) -> Optional[RotaPerson]:
        ctx: Dict[str, str] = {}
        self._inject_recipient_context(ctx, resolved)

        from .name_resolution.text import extract_display_name_and_email

        display, parsed_email = extract_display_name_and_email(raw_value)
        email = safe_text(ctx.get("EMAIL") or parsed_email or resolved.email or resolved.preferred_address).strip()
        if not self._is_valid_email(email):
            label = safe_text(ctx.get("FULL_NAME") or display or raw_value).strip()
            email = input(f"    Enter email address for {label} or press Enter to skip: ").strip()
        if not self._is_valid_email(email):
            print("  Skipping rota entry because no valid email address was supplied.")
            return None

        first_name = safe_text(ctx.get("NAME")).strip()
        surname = safe_text(ctx.get("SURNAME")).strip()
        full_name = safe_text(ctx.get("FULL_NAME")).strip()
        if not self._is_usable_full_name(full_name):
            full_name = display.strip() if self._is_usable_full_name(display) else ""
        if not self._is_usable_full_name(full_name):
            full_name = f"{first_name} {surname}".strip()

        if not self._is_usable_full_name(full_name):
            first_name = input("    Enter first name or press Enter to skip: ").strip()
            surname = input("    Enter surname or press Enter to skip: ").strip()
            computed = f"{first_name} {surname}".strip()
            display_name = input("    Enter display name or press Enter to use first/surname: ").strip()
            full_name = display_name or computed

        if self._is_usable_full_name(full_name) and (not first_name or not surname):
            bits = [part for part in full_name.split() if part]
            if bits:
                first_name = first_name or bits[0]
                surname = surname or (bits[-1] if len(bits) > 1 else "")

        if not self._is_usable_full_name(full_name):
            print("  Skipping rota entry because no name was supplied.")
            return None

        return RotaPerson(
            full_name=full_name,
            first_name=first_name,
            surname=surname,
            email=email,
            presenter=1,
            lab_worker=1,
            past_member=0,
        )

    def _confirm_rota_person(self, person: RotaPerson) -> bool:
        print("Rota people-sheet row:")
        print(f"  Name:  {person.full_name}")
        print(f"  Email: {person.email}")
        print("  Presenter: 1 | Lab Worker: 1 | Past member: 0")
        response = input("Add this person to the meeting rota people sheet? (Y/n): ").strip().lower()
        return response in {"", "y", "yes"}

    @staticmethod
    def _is_usable_full_name(value: str) -> bool:
        text = safe_text(value).strip()
        return bool(text and "@" not in text)

    @staticmethod
    def _is_valid_email(value: str) -> bool:
        text = safe_text(value).strip()
        return bool(re.match(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$", text, re.IGNORECASE))

    def _inject_recipient_context(self, ctx: Dict[str, str], r: ResolutionInfo) -> None:
        """Populate placeholder variables from a resolved recipient.

        Mirrors the VBScript behaviour where recipient lookup populates multiple variables
        (e.g., [NAME], [SURNAME], [EMAIL], etc.) that templates can use without prompting.
        """
        name = safe_text(r.first).strip()
        surname = safe_text(r.last).strip()
        full = safe_text(r.full).strip()
        email = safe_text(r.email or r.preferred_address).strip()

        # Basic cleanups and capitalization (port of EnsureCapitalizedNameText)
        if name:
            name = " ".join(capitalize_name_part(p) for p in name.split())
        if surname:
            surname = " ".join(capitalize_name_part(p) for p in surname.split())
        if not full:
            full = (name + " " + surname).strip()
        if full:
            full = " ".join(capitalize_name_part(p) for p in full.split())

        # If first/last weren't provided by directory lookup, derive from full name (best-effort)
        if full and (not name or not surname):
            # Handle "Surname, Name" as well as "Name Surname"
            if "," in full:
                parts = [p.strip() for p in full.split(",", 1)]
                if len(parts) == 2:
                    surname = surname or parts[0]
                    name = name or parts[1].split(" ", 1)[0]
            else:
                bits = [b for b in full.split() if b]
                if bits:
                    name = name or bits[0]
                    if len(bits) > 1:
                        surname = surname or bits[-1]


        # Derive username/alias from email or preferred address (covers aliases like 'gd2613')
        if "@" in email:
            username = email.split("@", 1)[0].strip()
        else:
            username = email
        alias = username

        def set_if_value(key: str, value: str) -> None:
            v = safe_text(value).strip()
            if not v:
                return
            ctx[key] = v

        # Common variants used across VB templates
        set_if_value("NAME", name)
        set_if_value("FIRSTNAME", name)
        set_if_value("SURNAME", surname)
        set_if_value("LASTNAME", surname)

        set_if_value("FULL_NAME", full)
        set_if_value("FULLNAME", full)

        set_if_value("EMAIL", email)
        set_if_value("USERNAME", username)
        set_if_value("ALIAS", alias)

        set_if_value("DIRECTORY_SOURCE", safe_text(r.directory_source))
        if r.directory_attempted:
            set_if_value("DIRECTORY_ATTEMPTED", ";".join(r.directory_attempted))


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Template-driven Outlook mailer (Python port of InductionMailer.vbs)")
    p.add_argument("--templates", default="templates", help="Templates root folder")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    app = InductionMailerApp(templates_root=Path(args.templates).resolve())
    return app.run_interactive()
