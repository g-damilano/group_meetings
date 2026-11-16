from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, List, Optional


@dataclass
class RotaEntry:
    meeting_date: date
    literature: str
    results: str
    alternate: str
    room: str
    lab_manager: str


@dataclass
class RotaPerson:
    full_name: str
    email: str
    first_name: str = ""
    surname: str = ""
    title: str = ""
    degree_pursued: str = ""
    presenter: int = 1
    lab_worker: int = 1
    past_member: int = 0


def _to_date(value) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    # openpyxl can return numbers for Excel serials if guess_types misbehaves;
    # in those cases we cannot reliably convert without workbook epoch.
    try:
        if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
            return date(int(value.year), int(value.month), int(value.day))
    except Exception:
        return None
    return None


class MeetingWorkbook:
    """Access to the meeting rota workbook.

    Port of VBS LoadMeetingWorkbook / LocateMeetingRows / sheet cell reads.
    Uses openpyxl so it does not require Excel.
    """

    def __init__(self, path: Path, sheet_rota: str = "Rota", sheet_emails: str = "Emails", editable: bool = False):
        self.path = Path(path)
        self.sheet_rota_name = sheet_rota
        self.sheet_emails_name = sheet_emails
        self.editable = editable
        self._wb = None
        self._rota = None
        self._emails = None

    def open(self) -> "MeetingWorkbook":
        from openpyxl import load_workbook

        # data_only=True => read cached formula values when present. For edits,
        # keep formulas intact so table totals survive round-trips.
        self._wb = load_workbook(self.path, data_only=not self.editable)
        self._rota = self._wb[self.sheet_rota_name]
        self._emails = self._wb[self.sheet_emails_name]
        return self

    def close(self) -> None:
        try:
            if self._wb is not None:
                self._wb.close()
        except Exception:
            pass
        self._wb = None
        self._rota = None
        self._emails = None

    def __enter__(self) -> "MeetingWorkbook":
        return self.open()

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    @property
    def rota(self):
        if self._rota is None:
            raise RuntimeError("Workbook not opened")
        return self._rota

    @property
    def emails(self):
        if self._emails is None:
            raise RuntimeError("Workbook not opened")
        return self._emails

    def last_rota_row(self) -> int:
        return int(self.rota.max_row or 0)

    def find_row_for_date(self, target: date) -> int:
        """Return the row index (1-based) for a date in column A; -1 if not found."""
        for r in range(2, self.last_rota_row() + 1):
            d = _to_date(self.rota.cell(r, 1).value)
            if d == target:
                return r
        return -1

    def find_rows_for_dates(self, targets: Iterable[date]) -> List[int]:
        return [self.find_row_for_date(d) for d in targets]

    def get_rota_entry(self, row: int) -> RotaEntry:
        d = _to_date(self.rota.cell(row, 1).value)
        if d is None:
            raise ValueError(f"Row {row} has no parsable date in column A")
        literature = str(self.rota.cell(row, 2).value or "").strip()
        results = str(self.rota.cell(row, 3).value or "").strip()
        alternate = str(self.rota.cell(row, 4).value or "").strip()
        room = str(self.rota.cell(row, 5).value or "").strip()
        lab_manager = str(self.rota.cell(row, 6).value or "").strip()
        return RotaEntry(
            meeting_date=d,
            literature=literature,
            results=results,
            alternate=alternate,
            room=room,
            lab_manager=lab_manager,
        )

    def eligible_names_for_tentative_rota(self) -> List[str]:
        """Read Emails sheet names where the Presenter flag is set."""
        out: List[str] = []
        max_row = int(self.emails.max_row or 0)
        presenter_col = self._find_header_col(self.emails, "Presenter", 5)
        for r in range(2, max_row + 1):
            eligible = self.emails.cell(r, presenter_col).value
            try:
                is_eligible = int(eligible) == 1
            except Exception:
                is_eligible = False
            if not is_eligible:
                continue
            name = str(self.emails.cell(r, 1).value or "").strip()
            if name:
                out.append(name)
        return out

    def save(self) -> None:
        if self._wb is None:
            raise RuntimeError("Workbook not opened")
        self._wb.save(self.path)

    def append_rota_person_if_missing(self, person: RotaPerson) -> str:
        """Append a person to the Emails table if name/email are not present.

        Returns: "added", "duplicate", or "skipped".
        """

        full_name = (person.full_name or "").strip()
        email = (person.email or "").strip()
        if not full_name or "@" not in email:
            return "skipped"

        duplicate_row = self.find_rota_person_duplicate_row(full_name, email)
        if duplicate_row > 0:
            return "duplicate"

        table = self._emails_table()
        if table is not None:
            insert_row = self._table_insert_row(table)
            self.emails.insert_rows(insert_row)
            self._copy_row_style(insert_row - 1, insert_row, max_col=self._table_max_col(table))
            self._set_person_values(insert_row, person)
            self._expand_people_table(table)
        else:
            insert_row = self._fallback_insert_row()
            self.emails.insert_rows(insert_row)
            self._copy_row_style(insert_row - 1, insert_row, max_col=7)
            self._set_person_values(insert_row, person)

        return "added"

    def find_rota_person_duplicate_row(self, full_name: str, email: str) -> int:
        name_col = self._find_header_col(self.emails, "Name", 1)
        email_col = self._find_header_col(self.emails, "Email", 2)
        last_data_row = self._last_people_data_row()
        full_key = (full_name or "").strip().casefold()
        email_key = (email or "").strip().casefold()

        for row in range(2, last_data_row + 1):
            existing_email = str(self.emails.cell(row, email_col).value or "").strip().casefold()
            existing_name = str(self.emails.cell(row, name_col).value or "").strip().casefold()
            if email_key and existing_email == email_key:
                return row
            if full_key and existing_name == full_key:
                return row
        return 0

    def _set_person_values(self, row: int, person: RotaPerson) -> None:
        values = {
            "Name": person.full_name.strip(),
            "First name": person.first_name.strip(),
            "Surname": person.surname.strip(),
            "Email": person.email.strip(),
            "Title": person.title,
            "Degree pursued": person.degree_pursued,
            "Presenter": int(person.presenter),
            "Lab Worker": int(person.lab_worker),
            "Past member": int(person.past_member),
        }
        defaults = {
            "Name": 1,
            "First name": 0,
            "Surname": 0,
            "Email": 2,
            "Title": 3,
            "Degree pursued": 4,
            "Presenter": 5,
            "Lab Worker": 6,
            "Past member": 7,
        }
        for header, value in values.items():
            col = self._find_header_col(self.emails, header, defaults[header])
            if col > 0:
                self.emails.cell(row, col).value = value

    def _emails_table(self):
        tables = self.emails.tables
        if "Table1" in tables:
            return tables["Table1"]
        first = next(iter(tables.keys()), None)
        return tables[first] if first else None

    def _table_insert_row(self, table) -> int:
        from openpyxl.utils.cell import range_boundaries

        _, _, _, max_row = range_boundaries(table.ref)
        totals = int(table.totalsRowCount or 0)
        return max_row if totals else max_row + 1

    def _table_max_col(self, table) -> int:
        from openpyxl.utils.cell import range_boundaries

        _, _, max_col, _ = range_boundaries(table.ref)
        return max_col

    def _expand_people_table(self, table) -> None:
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        new_max_row = max_row + 1
        table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

        totals = int(table.totalsRowCount or 0)
        last_data_row = new_max_row - totals if totals else new_max_row
        if totals and getattr(table, "autoFilter", None) is not None:
            table.autoFilter.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{last_data_row}"

        for header in ("Presenter", "Lab Worker"):
            col = self._find_header_col(self.emails, header, 0)
            if col <= 0:
                continue
            letter = get_column_letter(col)
            formula = f"SUBTOTAL(109,{letter}2:{letter}{last_data_row})"
            if totals:
                self.emails.cell(new_max_row, col).value = f"={formula}"
            for table_col in table.tableColumns:
                if str(table_col.name).casefold() == header.casefold() and table_col.totalsRowFormula is not None:
                    table_col.totalsRowFormula.attr_text = formula

    def _fallback_insert_row(self) -> int:
        last_row = self.emails.max_row or 1
        presenter_col = self._find_header_col(self.emails, "Presenter", 5)
        formula = str(self.emails.cell(last_row, presenter_col).value or "")
        return last_row if "SUBTOTAL" in formula.upper() else last_row + 1

    def _last_people_data_row(self) -> int:
        table = self._emails_table()
        if table is not None:
            from openpyxl.utils.cell import range_boundaries

            _, _, _, max_row = range_boundaries(table.ref)
            totals = int(table.totalsRowCount or 0)
            return max_row - totals

        last_row = self.emails.max_row or 1
        presenter_col = self._find_header_col(self.emails, "Presenter", 5)
        formula = str(self.emails.cell(last_row, presenter_col).value or "")
        return last_row - 1 if "SUBTOTAL" in formula.upper() else last_row

    def _copy_row_style(self, source_row: int, target_row: int, *, max_col: int) -> None:
        if source_row < 1:
            return
        for col in range(1, max_col + 1):
            source = self.emails.cell(source_row, col)
            target = self.emails.cell(target_row, col)
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.alignment:
                target.alignment = copy(source.alignment)
            if source.fill:
                target.fill = copy(source.fill)
            if source.border:
                target.border = copy(source.border)
            if source.protection:
                target.protection = copy(source.protection)

    @staticmethod
    def _find_header_col(ws, header_text: str, default_column: int) -> int:
        target = (header_text or "").strip().casefold()
        max_col = int(ws.max_column or 0)
        for col in range(1, max_col + 1):
            current = str(ws.cell(1, col).value or "").strip().casefold()
            if current == target:
                return col
        return int(default_column or 0)
