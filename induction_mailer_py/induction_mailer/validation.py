from __future__ import annotations

import re
from typing import Any

from .utils import safe_text


_EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


def normalize_text(value: Any) -> str:
    return safe_text(value).strip()


def is_valid_email(value: Any) -> bool:
    text = normalize_text(value)
    if not text:
        return False
    return bool(_EMAIL_RE.match(text))
