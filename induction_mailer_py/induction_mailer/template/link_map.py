from __future__ import annotations

from pathlib import Path
from typing import Dict

from ..utils import safe_text


def load_link_map_context(map_path: str | Path) -> Dict[str, str]:
    path = Path(map_path)
    if not path.exists():
        return {}

    out: Dict[str, str] = {}
    with path.open("r", encoding="utf-8", errors="replace") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip()
            v = v.strip()
            if k:
                out[k.upper()] = v
    return out


def apply_link_map(html_body: str, link_map: Dict[str, str]) -> str:
    """Apply link map replacements.

    VBS replaces these forms:
    - `{{key}}`, `{{KEY}}`
    - `[key]`, `[KEY]`
    """
    result = safe_text(html_body)
    for k, v in (link_map or {}).items():
        val = safe_text(v)
        key = safe_text(k)
        result = result.replace("{{" + key + "}}", val)
        result = result.replace("{{" + key.upper() + "}}", val)
        result = result.replace("[" + key + "]", val)
        result = result.replace("[" + key.upper() + "]", val)
    return result
