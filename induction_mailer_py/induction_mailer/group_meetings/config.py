from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class GroupMeetingsConfig:
    """Configuration ..."""

    debug_mode: bool = False
    extended_logging: bool = True

    # Emails
    manager_email: str = "gd2613@ic.ac.uk"
    building_admin_email: str = "i.gomezvilchis@imperial.ac.uk"
    supervisor_email: str = "agi@imperial.ac.uk"
    default_previous_lab_manager: str = "Agi"
    inventory_manager_name: str = "Yilyu Qi"
    inventory_manager_email: str = "y.qi24@imperial.ac.uk"

    email_to: str = "brandt-talbot-group@imperial.ac.uk"

    # Workbook
    excel_path: Path = Path(
        r"C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings\meeting_rota.xlsx"
    )
    sheet_rota: str = "Rota"
    sheet_emails: str = "Emails"
    delete_local_workbook_after_run: bool = True

    # Optional sync via rclone
    rclone_exe: Path = Path(r"C:\rclone\rclone.exe")
    rclone_source: str = "ABTsharepoint:/Group Meetings/meeting_rota.xlsx"
    rclone_dest_dir: Path = Path(r"C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings")

    # Log
    log_path: Path = Path("group_meetings_manager.log")
    log_variant: str = "[Main] "
