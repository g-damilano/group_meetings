from __future__ import annotations

from datetime import date, datetime, timedelta


def next_fridays(today: date | None = None, count: int = 3) -> list[date]:
    """Port of VBS PopulateFridayDates.

    Calculates the next Friday (including today if today is Friday), then subsequent
    Fridays ..."""

    if today is None:
        today = date.today()
    friday_weekday = 4  # Monday=0
    delta = (friday_weekday - today.weekday()) % 7
    first = today + timedelta(days=delta)
    return [first + timedelta(days=7 * i) for i in range(max(0, int(count)))]


def format_date_uk(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def format_day_month(d: date) -> str:
    return d.strftime("%d/%m")


def format_meeting_date_for_filename(d: date) -> str:
    return d.strftime("%y%m%d")


def extract_name_initials(full_name: str) -> str:
    """Port of VBS ExtractNameInitials.

    Accepts either "Surname, First" or "First Surname".
    """
    cleaned = (full_name or "").strip()
    if not cleaned:
        return ""

    first = ""
    last = ""
    if "," in cleaned:
        parts = [p.strip() for p in cleaned.split(",", 1)]
        last = parts[0] if parts else ""
        first = parts[1] if len(parts) > 1 else ""
    else:
        tokens = [t for t in cleaned.split() if t.strip()]
        if tokens:
            first = tokens[0]
            last = tokens[-1]

    if first.lower() == last.lower():
        last = ""

    out = ""
    if first:
        out += first[:1].upper()
    if last:
        out += last[:1].upper()
    return out


def first_name_from_display(display: str) -> str:
    """Extract first name from Outlook display name ("Surname, First")."""
    s = (display or "").strip()
    if not s:
        return ""
    if "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
        if len(parts) > 1 and parts[1]:
            return parts[1].split()[0]
        return parts[0].split()[0]
    return s.split()[0]
