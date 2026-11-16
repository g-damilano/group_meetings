from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import List, Tuple

from .utils import format_date_uk, next_fridays


@dataclass
class TentativeRota:
    """A ..."""

    body_text: str
    csv_path: Path


def generate_tentative_rota(names: List[str], weeks: int = 26, start: date | None = None) -> Tuple[str, List[Tuple[date, str, str]]]:
    """Port of VBS GenerateTentativeRota."""
    eligible = [n.strip() for n in (names or []) if n and n.strip()]
    if not eligible:
        return "No eligible names for rota.", []

    rnd = eligible[:]
    random.shuffle(rnd)

    if start is None:
        start = next_fridays(date.today(), 1)[0]

    total_weeks = int(weeks)
    assignments: List[Tuple[date, str, str]] = []
    n = len(rnd)
    for i in range(total_weeks):
        dt = start + timedelta(days=7 * i)
        p1 = rnd[(i * 2) % n]
        p2 = rnd[(i * 2 + 1) % n]
        assignments.append((dt, p1, p2))

    lines: List[str] = []
    lines.append("Tentative 6-month rota (Fridays, 2 presenters/day):")
    lines.append("")
    lines.append("Date         | Presenter 1          | Presenter 2")
    lines.append("-" * 50)
    for dt, p1, p2 in assignments:
        lines.append(f"{format_date_uk(dt)} | {p1:<20} | {p2:<20}")

    return "\r\n".join(lines), assignments


def write_rota_csv(assignments: List[Tuple[date, str, str]], csv_path: Path) -> Path:
    csv_path = Path(csv_path)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        f.write("Date,Presenter1,Presenter2\n")
        for dt, p1, p2 in assignments:
            f.write(f"{format_date_uk(dt)},{p1},{p2}\n")
    return csv_path


def build_tentative_rota_bundle(names: List[str], csv_path: Path, weeks: int = 26, start: date | None = None) -> TentativeRota:
    body, assignments = generate_tentative_rota(names, weeks=weeks, start=start)
    if assignments:
        write_rota_csv(assignments, csv_path)
        body = body + f"\r\n\r\n[CSV: {csv_path}]"
    return TentativeRota(body_text=body, csv_path=Path(csv_path))
