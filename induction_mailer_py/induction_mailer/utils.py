from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, MutableMapping, Optional, Tuple


def safe_text(value: Any) -> str:
    """Best-effort string conversion mirroring VBS SafeText().

    - None/"empty" -> ""
    - bytes -> decode UTF-8 (fallback latin-1)
    - otherwise: str(value) (guarded)
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (bytes, bytearray)):
        try:
            return bytes(value).decode("utf-8")
        except Exception:
            return bytes(value).decode("latin-1", errors="replace")
    try:
        return str(value)
    except Exception:
        return ""


def safe_lower(value: Any) -> str:
    return safe_text(value).lower()


def remove_bom(text: str) -> str:
    # UTF-8 BOM
    if text.startswith("\ufeff"):
        return text.lstrip("\ufeff")
    return text


def read_text_smart(path: str | Path) -> str:
    """Read UTF-8 text, falling back similarly to the VBS ADODB.Stream approach."""
    p = Path(path)
    try:
        # utf-8-sig strips BOM
        return p.read_text(encoding="utf-8-sig", errors="strict")
    except Exception:
        try:
            return p.read_text(encoding="utf-8", errors="replace")
        except Exception:
            try:
                return p.read_text(encoding="latin-1", errors="replace")
            except Exception:
                return ""


def normalize_line_endings(text: str) -> str:
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    return text


def trim_leading_blank_lines(text: str) -> str:
    while text.startswith("\n") or text.startswith("\r"):
        text = text[1:]
    return text


def capitalize_name_part(value: Any) -> str:
    s = safe_text(value).strip()
    if not s:
        return ""
    if len(s) == 1:
        return s.upper()
    return s[0].upper() + s[1:]


_WORD_CHAR_RE = re.compile(r"[0-9A-Za-z\u00C0-\u02AF\u0370-\u1FFF]")


def is_word_character(ch: str) -> bool:
    if not ch:
        return False
    return bool(_WORD_CHAR_RE.fullmatch(ch))


def replace_whole_word_case_insensitive(text: str, find_value: str, replacement: str) -> str:
    """Port of ReplaceWholeWordCaseInsensitive() from InductionMailer.vbs."""
    source = safe_text(text)
    target = safe_text(find_value)
    if not target:
        return source

    # Manual scan to respect word boundaries.
    result = source
    start = 0
    tlen = len(target)

    while True:
        pos = result.lower().find(target.lower(), start)
        if pos < 0:
            break

        before = result[pos - 1] if pos > 0 else ""
        after = result[pos + tlen] if pos + tlen < len(result) else ""

        if (not before or not is_word_character(before)) and (not after or not is_word_character(after)):
            result = result[:pos] + replacement + result[pos + tlen :]
            start = pos + len(replacement)
        else:
            start = pos + tlen

    return result


class CaseInsensitiveDict(MutableMapping[str, Any]):
    """A small case-insensitive mapping.

    - Keys are compared case-insensitively
    - Original casing is preserved for iteration

    This mirrors the VBS `Scripting.Dictionary` with `CompareMode = TextCompare`.
    """

    def __init__(self, data: Optional[Dict[str, Any]] = None):
        self._store: Dict[str, Tuple[str, Any]] = {}
        if data:
            for k, v in data.items():
                self[k] = v

    def __setitem__(self, key: str, value: Any) -> None:
        lk = key.lower()
        self._store[lk] = (key, value)

    def __getitem__(self, key: str) -> Any:
        return self._store[key.lower()][1]

    def __delitem__(self, key: str) -> None:
        del self._store[key.lower()]

    def __iter__(self) -> Iterator[str]:
        for original, _ in self._store.values():
            yield original

    def __len__(self) -> int:
        return len(self._store)

    def __contains__(self, key: object) -> bool:
        if not isinstance(key, str):
            return False
        return key.lower() in self._store

    def get(self, key: str, default: Any = None) -> Any:
        return self._store.get(key.lower(), ("", default))[1]

    def items(self):
        return ((orig, val) for orig, val in self._store.values())

    def keys(self):
        return (orig for orig, _ in self._store.values())

    def values(self):
        return (val for _, val in self._store.values())

    def copy(self) -> "CaseInsensitiveDict":
        return CaseInsensitiveDict(dict(self.items()))
