from __future__ import annotations

from datetime import date, timedelta
from typing import Mapping, Sequence
from urllib.parse import quote

from .utils import (
    extract_name_initials,
    first_name_from_display,
    format_date_uk,
    format_day_month,
    format_meeting_date_for_filename,
)


def build_meeting_subject(meeting_date: date) -> str:
    """VBS BuildMeetingSubject."""
    return f"ABT group meeting @ 11am ({format_date_uk(meeting_date)})"


def build_presenter_reminder_subject(presentation_type: str, meeting_date: date) -> str:
    """VBS BuildPresenterReminderSubject."""
    t = (presentation_type or "").strip()
    norm = t.lower()
    if norm == "literature":
        label = "Literature presentation"
    elif norm == "results":
        label = "Results presentation"
    elif t:
        label = f"{t} presentation"
    else:
        label = "Presentation"
    return f"{label} reminder ({format_date_uk(meeting_date)})"


def build_meeting_lead_reminder_subject(meeting_date: date) -> str:
    return f"Meeting lead reminder ({format_date_uk(meeting_date)})"


def build_lab_manager_reminder_subject(meeting_date: date) -> str:
    start = format_date_uk(meeting_date - timedelta(days=4))
    end = format_date_uk(meeting_date)
    en_dash = "\u2013"
    return f"Lab management reminder ({start} {en_dash} {end})"


def build_weekly_responsibility_reminder_subject(meeting_date: date) -> str:
    start = format_date_uk(meeting_date - timedelta(days=4))
    end = format_date_uk(meeting_date)
    en_dash = "\u2013"
    return f"Weekly lab, meeting, and inventory responsibilities ({start} {en_dash} {end})"


def is_no_meeting_event(alternate: str) -> bool:
    normalized = (alternate or "").strip().lower()
    return "college closure" in normalized or "no meeting" in normalized or "no-meeting" in normalized


def generate_main_email_content(
    *,
    sender_display_name: str,
    meeting_date: date,
    room: str,
    literature: str,
    results: str,
    alternate: str,
    lab_manager: str,
    meeting_lead_name: str,
    next_friday_date: date,
    next_literature: str,
    next_results: str,
    next_alternate: str,
    next_lab_manager: str,
    next_meeting_lead_name: str,
) -> str:
    """Port of VBS GenerateEmailContent.

    Produces a plain-text Outlook body.
    """

    sender_first = first_name_from_display(sender_display_name) or sender_display_name.strip()
    meeting_date_text = f"{meeting_date:%A} {format_date_uk(meeting_date)}"
    next_friday_text = f"{next_friday_date:%A} {format_date_uk(next_friday_date)}"
    meeting_short = format_day_month(meeting_date)
    next_short = format_day_month(next_friday_date)

    def _tbc(x: str) -> str:
        return x.strip() if x.strip() else "TBC (X)"

    current_results = _tbc(results)
    current_lit = _tbc(literature)
    current_lab_manager = _tbc(lab_manager)
    current_chair = _tbc(meeting_lead_name)

    next_results_ = _tbc(next_results)
    next_lit_ = _tbc(next_literature)
    next_lab_manager_ = _tbc(next_lab_manager)
    next_chair = _tbc(next_meeting_lead_name)

    lab_manager_range_start = format_day_month(meeting_date - timedelta(days=4))
    lab_manager_range_end = format_day_month(meeting_date)
    next_lab_manager_range_start = format_day_month(next_friday_date - timedelta(days=4))
    next_lab_manager_range_end = format_day_month(next_friday_date)

    lines: list[str] = []
    lines.append("Hi all,")
    lines.append("")

    if room.strip():
        lines.append(
            f"We will have our group meeting on {meeting_date_text} at 11:00, in person ({room.strip()}) "
            "or via MS Teams (see calendar invite for the link)."
        )
    else:
        lines.append(
            f"We will have our group meeting on {meeting_date_text} at 11:00, in person or via MS Teams "
            "(see calendar invite for the link)."
        )
    lines.append("")

    lines.append(f"This week ({meeting_short}):")
    lines.append("")
    if alternate.strip():
        lines.append(f"- {alternate.strip()}")
    else:
        lines.append(f"- Results: {current_results}")
        lines.append(f"- Literature: {current_lit}")
    lines.append(
        f"- Lab manager / chair / timekeeper: {current_lab_manager} "
        f"({lab_manager_range_start} - {lab_manager_range_end})"
    )
    lines.append("")

    lines.append(
        "Please treat this week's lab manager as the first point of contact for minor lab-practice "
        "questions and reminders. Their role is supportive oversight: helping us adopt the agreed "
        "lab practices, validating what is working, and bringing persistent issues to the Friday "
        "lab-status update constructively, with the focus on behaviours and improvements rather than blame."
    )
    lines.append("")

    lines.append(
        "We will be starting promptly at 11:00. Presenters, please arrive/join 5-10 minutes early to set up. "
        "If the room is occupied right up to 11:00, we will start at 11:05."
    )
    lines.append("")

    lines.append(f"For the following Friday ({next_friday_text}):")
    lines.append("")
    if next_alternate.strip():
        lines.append(f"- {next_alternate.strip()}")
    else:
        lines.append(f"- Results: {next_results_}")
        lines.append(f"- Literature: {next_lit_}")
    if not is_no_meeting_event(next_alternate):
        lines.append(
            f"- Lab manager / chair / timekeeper: {next_lab_manager_} "
            f"({next_lab_manager_range_start} - {next_lab_manager_range_end})"
        )
    lines.append("")

    lines.append("Upcoming weeks schedule:")
    lines.append(
        "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1"
    )
    lines.append("")
    lines.append("Best,")
    lines.append(sender_first.strip())

    return "\r\n".join(lines)


def generate_presenter_reminder_email(
    *,
    sender_display_name: str,
    presenter_name: str,
    presentation_type: str,
    meeting_date: date,
) -> str:
    """Port of VBS GeneratePresenterReminderEmail."""

    presenter_name = (presenter_name or "").strip()
    presentation_type = (presentation_type or "").strip()
    sender_first = first_name_from_display(sender_display_name) or sender_display_name.strip() or "Automated"

    trimmed_type = presentation_type
    norm = trimmed_type.lower()
    if norm == "literature":
        description = "the literature session"
        type_label = "Literature presentation"
    elif norm == "results":
        description = "the results update"
        type_label = "Results presentation"
    elif trimmed_type:
        description = f"the {trimmed_type} presentation"
        type_label = f"{trimmed_type} presentation"
    else:
        description = "your presentation"
        type_label = "Presentation"

    reminder_subject = type_label.lower()
    meeting_date_text = format_date_uk(meeting_date)
    meeting_date_code = format_meeting_date_for_filename(meeting_date)
    presenter_initials = extract_name_initials(presenter_name) or "XX"

    # greeting name: best effort first token after comma, else first word
    greet = presenter_name
    if "," in greet:
        parts = [p.strip() for p in greet.split(",", 1)]
        greet = parts[1] if len(parts) > 1 and parts[1] else parts[0]
    greet = greet.split()[0] if greet.split() else ""

    if not presenter_name:
        return "\r\n".join(
            [
                "Hi,",
                "",
                f"No presenter is currently assigned for the {reminder_subject} scheduled on {meeting_date_text}.",
                "",
                "Please update the rota or advise the group so we can keep the agenda on track.",
                "",
                "Best,",
                sender_first,
            ]
        )

    en_dash = "\u2013"
    apostrophe = "'"
    bullet = "\u2022"

    lines: list[str] = []
    lines.append(f"Hi {greet},")
    lines.append("")
    lines.append(
        f"Just a reminder that you are scheduled to present {description} at this week's group meeting on {meeting_date_text}. "
        "The slot is 15 minutes for your presentation plus 5 minutes for Q&A. "
        f"As a rough guide, aim for about 1 minute per slide (so around 10{en_dash}12 slides in total). "
        "The group meeting leader will keep time to ensure we don" + apostrophe + "t overrun, so please rehearse your talk beforehand "
        "to check that you are within the time limit."
    )
    lines.append("")
    lines.append(
        "Please save your presentation to the group SharePoint folder "
        "(https://imperiallondon.sharepoint.com/:f:/r/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/Group%20Meeting%20Presentations?csf=1&web=1&e=eK2FLV) "
        f"using the naming pattern YYMMDD_[NameInitials]_{norm}_presentation.pptx."
    )
    lines.append(
        f"For clarity: use the meeting date as YYMMDD, add your initials, and include the talk type (e.g., {description}), "
        f"ending with _presentation.pptx (e.g., {meeting_date_code}_{presenter_initials}_{norm}_presentation.pptx)."
    )
    lines.append("")

    if norm == "literature":
        lines.append(
            "Share the article you'll present on Friday by Wednesday, so anyone interested can skim it in advance and come prepared "
            f"with questions{en_dash}making the discussion more enjoyable and productive."
        )
        lines.append("")
        lines.append(
            "For literature presentations, please choose high-quality papers that may be of general interest, and focus on the higher-level "
            "implications (e.g., key scientific principles and/or relevance to work within the group)."
        )
        lines.append("")

    lines.append("A few pointers for preparing your slides:")
    lines.append("")
    lines.append(f"{bullet} Aim for a balanced, well-planned presentation, rather than trying to include everything.")
    lines.append("")
    lines.append(
        f"{bullet} Use no more than ~3 concise bullet points per slide to support (not duplicate) what you say, so the audience focuses on you "
        "rather than reading the screen."
    )
    lines.append("")
    lines.append(
        f"{bullet} Let the slides support your performance, not be the performance{en_dash}they shouldn" + apostrophe + "t steal your stage."
    )
    lines.append("")
    lines.append(
        f"{bullet} One clear, well-chosen image is often more effective than a lot of text, especially as many scientists are visual learners."
    )
    lines.append("")
    lines.append(
        f"{bullet} Assume a left-to-right, top-to-bottom reading pattern: place the most important information at the top left and less critical "
        "details toward the bottom right."
    )
    lines.append("")
    lines.append("Please ensure your materials are ready and let me know if you need any support.")
    lines.append("")
    lines.append("Best,")
    lines.append(sender_first)

    return "\r\n".join(lines)


def generate_meeting_lead_reminder_email(*, sender_display_name: str, meeting_lead_name: str, meeting_date: date) -> str:
    """Port of VBS GenerateMeetingLeadReminderEmail."""
    sender_first = first_name_from_display(sender_display_name) or sender_display_name.strip() or "Automated"

    lead = (meeting_lead_name or "").strip()
    greet = lead
    if "," in greet:
        parts = [p.strip() for p in greet.split(",", 1)]
        greet = parts[1] if len(parts) > 1 and parts[1] else parts[0]
    greet = greet.split()[0] if greet.split() else "there"

    lines = [
        f"Hi {greet},",
        "",
        f"Just a reminder that you will lead this week's group meeting on {format_date_uk(meeting_date)}.",
        "",
        "Each presenter has 15 minutes for their talk plus 5 minutes for Q&A.",
        "",
        "Please arrive about 5 minutes early to set up the room so the computer is ready right at meeting time.",
        "",
        "Bring the red/yellow timing cards (5 minutes/1 minute left) to signal time to each presenter, and keep presenters on schedule while running the inclusion activity.",
        "",
        "Best,",
        sender_first,
    ]
    return "\r\n".join(lines)


def generate_lab_manager_reminder_email(*, sender_display_name: str, lab_manager_name: str, meeting_date: date) -> str:
    """Port of VBS GenerateLabManagerReminderEmail."""
    manager = (lab_manager_name or "").strip() or "there"

    greet = manager
    if greet.lower() != "there":
        if "," in greet:
            parts = [p.strip() for p in greet.split(",", 1)]
            greet = parts[1] if len(parts) > 1 and parts[1] else parts[0]
        greet = greet.split()[0] if greet.split() else manager

    sender_first = first_name_from_display(sender_display_name) or sender_display_name.strip() or "Automated"

    start_text = format_date_uk(meeting_date)
    end_text = format_date_uk(meeting_date + timedelta(days=7))
    en_dash = "\u2013"
    em_dash = "\u2014"
    week_range = f"{start_text} {en_dash} {end_text}"

    # best-effort parcel addressee: use sender_display_name
    parcel_addressee = sender_display_name.strip() or "me"

    lines = [
        f"Hi {greet},",
        "",
        f"This is a reminder that you are the lab manager for the week of {week_range}.",
        "",
        f"Lab cleaning starts on Friday and must be completed by next Friday {em_dash} please try to finish it by Friday EOD so the lab is ready for Monday.",
        "",
        f"On friday, please pick up all parcels addressed to {parcel_addressee}; most of them contain the consumables ordered the last week that need to be re-stocked in the lab.",
        "",
        "Also, please procure the list of missing items on the whiteboard from the stores so that we can maintain supplies at all times.",
        "",
        "Thank you for keeping the lab organized.",
        "",
        "Best,",
        sender_first,
    ]
    return "\r\n".join(lines)


def generate_weekly_responsibility_reminder_email(
    *,
    sender_display_name: str,
    responsible_name: str,
    meeting_date: date,
    inventory_manager_name: str,
    inventory_manager_email: str = "",
    inventory_feedback_cc_email: str = "",
    inventory_items: Sequence[Mapping[str, str]] | None = None,
) -> str:
    """Combined lab-manager and group-meeting chair/timekeeper reminder."""

    manager = (responsible_name or "").strip() or "there"
    greet = manager
    if greet.lower() != "there":
        if "," in greet:
            parts = [p.strip() for p in greet.split(",", 1)]
            greet = parts[1] if len(parts) > 1 and parts[1] else parts[0]
        greet = greet.split()[0] if greet.split() else manager

    sender_first = first_name_from_display(sender_display_name) or sender_display_name.strip() or "Automated"
    week_range = f"{format_date_uk(meeting_date - timedelta(days=4))} \u2013 {format_date_uk(meeting_date)}"
    parcel_addressee = sender_display_name.strip() or "me"

    lines = [
        f"Hi {greet},",
        "",
        f"This is a reminder that you are the lab manager, group meeting chair/timekeeper, and weekly inventory reviewer for the week of {week_range}.",
        "",
        "Group meeting responsibilities",
        f"- Lead this week's group meeting on {format_date_uk(meeting_date)} at 11:00.",
        "- Keep each presenter to 15 minutes for the talk plus 5 minutes for Q&A.",
        "- Arrive about 5 minutes early so the room computer is ready at the start.",
        "- Bring the red/yellow timing cards and keep the inclusion activity moving.",
        "- At the end of the meeting, give a brief lab-status update: what is working well, any minor reminders, and any persistent issues that need group attention.",
        "",
        "Lab management responsibilities",
        "- Be the supportive first point of contact for minor lab-practice questions and reminders during the week.",
        "- Help people adopt the agreed lab practices by validating what is working and gently redirecting small issues before they grow. This is not an enforcement role; persistent concerns can be raised constructively during the Friday lab-status update, focusing on behaviours rather than blame.",
        "- Lab cleaning starts on Monday and must be completed by this Friday; please aim to finish by Friday EOD so the lab is ready for Monday.",
        f"- On Friday, please pick up parcels addressed to {parcel_addressee}; most contain consumables ordered the previous week that need re-stocking.",
        "- Please procure the missing items listed on the whiteboard from stores so we can maintain supplies.",
        "",
    ]

    if inventory_items:
        lines.extend(
            [
                "Inventory review",
                f"Please check the following Labcup records physically in the lab. Confirm the chemical identity, location, quantity, hazards, and storage temperature. Complete the review and send feedback to {inventory_manager_name} by Friday using the draft links below. I will be copied on the feedback emails.",
                "",
                "No.\tBarcode\tProduct\tCAS\tLocation\tQuantity\tHazards\tStorage",
            ]
        )
        for idx, item in enumerate(inventory_items, start=1):
            lines.append(
                "\t".join(
                    [
                        str(idx),
                        str(item.get("Barcode", "-") or "-"),
                        str(item.get("Product", "-") or "-"),
                        str(item.get("CAS", "-") or "-"),
                        str(item.get("Location", "-") or "-"),
                        str(item.get("Quantity", "-") or "-"),
                        str(item.get("Hazards", "-") or "-"),
                        str(item.get("StorageTemperature", "-") or "-"),
                    ]
                )
            )
        feedback_link = build_inventory_feedback_mailto(
            items=inventory_items,
            meeting_date=meeting_date,
            inventory_manager_name=inventory_manager_name,
            inventory_manager_email=inventory_manager_email,
            cc_email=inventory_feedback_cc_email,
        )
        if feedback_link:
            lines.append(f"Feedback draft for this batch: {feedback_link}")
        lines.append("")

    lines.extend(
        [
            "Thank you for keeping the meeting and lab work tidy and easy to follow.",
            "",
            "Best,",
            sender_first,
        ]
    )
    return "\r\n".join(lines)


def _inventory_table_cell(value: object) -> str:
    text = str(value or "").replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    return text or "-"


def build_inventory_feedback_mailto(
    *,
    items: Sequence[Mapping[str, str]],
    meeting_date: date,
    inventory_manager_name: str,
    inventory_manager_email: str,
    cc_email: str = "",
) -> str:
    to_email = (inventory_manager_email or "").strip()
    if not to_email or not items:
        return ""

    recipient_name = (inventory_manager_name or "").strip() or "Inventory team"
    subject = f"Labcup inventory feedback: week ending {format_date_uk(meeting_date)}"
    body_lines = [
        f"Hi {recipient_name},",
        "",
        f"I reviewed the following Labcup inventory records for the week ending {format_date_uk(meeting_date)}. Feedback is due by Friday.",
    ]

    for idx, item in enumerate(items, start=1):
        body_lines.extend(
            [
                "",
                f"Item {idx}",
                f"Barcode: {_inventory_table_cell(item.get('Barcode', ''))}",
                f"Product: {_inventory_table_cell(item.get('Product', ''))}",
                f"CAS: {_inventory_table_cell(item.get('CAS', ''))}",
                f"Brand: {_inventory_table_cell(item.get('Brand', ''))}",
                f"Product number: {_inventory_table_cell(item.get('ProductNumber', ''))}",
                f"Location in Labcup: {_inventory_table_cell(item.get('Location', ''))}",
                f"Quantity in Labcup: {_inventory_table_cell(item.get('Quantity', ''))}",
                f"Hazards in Labcup: {_inventory_table_cell(item.get('Hazards', ''))}",
                f"Storage temperature in Labcup: {_inventory_table_cell(item.get('StorageTemperature', ''))}",
                "",
                "Feedback / corrections:",
                "- Chemical identity/name:",
                "- Location:",
                "- Quantity:",
                "- Hazards/storage:",
                "- Other notes:",
            ]
        )

    body_lines.extend(["", "Best,"])
    body = "\r\n".join(body_lines)

    query = f"subject={quote(subject, safe='')}&body={quote(body, safe='')}"
    if cc_email.strip():
        query = f"cc={quote(cc_email.strip(), safe='')}&{query}"
    return f"mailto:{to_email}?{query}"


def generate_email_alert(
    *,
    alert_type: str,
    missing_date_text: str,
    recipient_email: str,
    tentative_rota: str,
    building_admin_email: str,
) -> str:
    """Port of VBS GenerateEmailAlert."""
    lines: list[str] = ["Hi,", ""]
    if alert_type == "MissingDate":
        lines.append(f"There is no entry in the rota for {missing_date_text}.")
        lines.append("Please update the rota accordingly.")
        if tentative_rota:
            lines.extend(["", tentative_rota])
    elif alert_type == "Unassigned":
        lines.append(f"The rota entry for {missing_date_text} has no assigned presenters.")
        lines.append("Please review the tentative 6-month rota below and approve or adjust:")
        if tentative_rota:
            lines.extend(["", tentative_rota])
    elif alert_type == "MissingRoom":
        lines.append(f"The rota entry for {missing_date_text} has no room reservation.")
        lines.append("Please reserve a room for the meeting.")
        lines.append("")
        lines.append(
            "Send Email to the building manager to request room for the upcoming term: "
            f"mailto:{building_admin_email}?subject=Room%20Reservation%20Request%20for%20ABT%20Group%20Meetings%20%E2%80%93%20Upcoming%20Term&body="
            "Hi%20building%20manager,%0D%0A%0D%0AI%20hope%20you%20are%20well.%0D%0A%0D%0AI%20would%20like%20to%20request%20a%20recurring%20booking%20for%20our%20ABT%20group%20meetings%20for%20the%20upcoming%20academic%20term.%20Could%20you%20please%20reserve%20G25%20(MSRH)%20for%20Fridays%20from%2011:00%20to%2012:00,%20starting%20from%20the%20next%20available%20week?%0D%0A%0D%0ACould%20you%20also%20let%20me%20know%20the%20exact%20date%20the%20reservation%20ends?%0D%0A%0D%0AThank%20you%20very%20much%20for%20your%20help.%0D%0A%0D%0ABest%20regards,%0D%0AGiacomo"
        )
    elif alert_type == "MeetingNotSet":
        lines.append(
            f"The rota entry for {missing_date_text} has no presenters or alternate assigned, so the regular meeting email was not sent to the distribution list."
        )
        lines.append("Please update the rota with presenters so the meeting can be scheduled.")

    lines.extend(["", f"Contact directly via email: {recipient_email}", "", "Best,", "Automated Meeting Scheduler"])
    return "\r\n".join(lines)
