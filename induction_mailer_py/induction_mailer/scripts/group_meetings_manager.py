from __future__ import annotations

import argparse
from pathlib import Path

from ..group_meetings import GroupMeetingsConfig, GroupMeetingsManager


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="group-meetings-manager",
        description="Send ABT group meeting emails & reminders (ported from group_meetings_manager.vbs).",
    )
    p.add_argument("--debug", action="store_true", help="Display emails instead of sending.")
    p.add_argument("--log", type=str, default=None, help="Log file path.")
    p.add_argument("--excel", type=str, default=None, help="Path to local meeting_rota.xlsx.")
    p.add_argument("--no-delete", action="store_true", help="Do not delete local workbook after the run.")
    p.add_argument("--email-to", type=str, default=None, help="Distribution list for the regular meeting email.")
    p.add_argument("--manager-email", type=str, default=None, help="Fallback manager email for name resolution.")
    p.add_argument("--supervisor-email", type=str, default=None, help="Supervisor email for critical alerts.")
    p.add_argument("--building-admin-email", type=str, default=None, help="Building admin email for room booking mailto.")
    p.add_argument("--rclone-exe", type=str, default=None, help="Path to rclone.exe.")
    p.add_argument("--rclone-source", type=str, default=None, help="rclone source remote path.")
    p.add_argument("--rclone-dest", type=str, default=None, help="rclone destination directory.")
    p.add_argument("--sheet-rota", type=str, default=None, help="Worksheet name for rota.")
    p.add_argument("--sheet-emails", type=str, default=None, help="Worksheet name for emails.")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    cfg = GroupMeetingsConfig()
    if args.debug:
        cfg.debug_mode = True
    if args.log:
        cfg.log_path = Path(args.log)
    if args.excel:
        cfg.excel_path = Path(args.excel)
    if args.no_delete:
        cfg.delete_local_workbook_after_run = False
    if args.email_to:
        cfg.email_to = args.email_to
    if args.manager_email:
        cfg.manager_email = args.manager_email
    if args.supervisor_email:
        cfg.supervisor_email = args.supervisor_email
    if args.building_admin_email:
        cfg.building_admin_email = args.building_admin_email
    if args.rclone_exe:
        cfg.rclone_exe = Path(args.rclone_exe)
    if args.rclone_source:
        cfg.rclone_source = args.rclone_source
    if args.rclone_dest:
        cfg.rclone_dest_dir = Path(args.rclone_dest)
    if args.sheet_rota:
        cfg.sheet_rota = args.sheet_rota
    if args.sheet_emails:
        cfg.sheet_emails = args.sheet_emails

    GroupMeetingsManager(cfg).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
