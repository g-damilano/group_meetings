from __future__ import annotations

import re
from typing import Dict, Iterable, List, Sequence

from ..utils import CaseInsensitiveDict, safe_text


_PLACEHOLDER_RE = re.compile(r"\[([A-Za-z0-9_]+)\]")


def build_context_from_meta(meta: Dict[str, str], link_map_context: Dict[str, str] | None = None) -> CaseInsensitiveDict:
    # VBS clones linkMapContext and merges meta keys (uppercased)
    base = CaseInsensitiveDict(dict((k.upper(), v) for k, v in (link_map_context or {}).items()))
    for k, v in (meta or {}).items():
        base[k.upper()] = v
    return base


def extract_placeholders(text_items: Sequence[str]) -> List[str]:
    seen = set()
    for txt in text_items:
        for m in _PLACEHOLDER_RE.finditer(safe_text(txt)):
            seen.add(m.group(1).upper())
    return sorted(seen)


def apply_context_to_text(text: str, context: Dict[str, str]) -> str:
    result = safe_text(text)
    for k, v in context.items():
        result = re.sub(r"\[" + re.escape(k) + r"\]", safe_text(v), result, flags=re.IGNORECASE)
    return result


def merge_recipients(base_list: str, extra: Iterable[str]) -> str:
    seen = set()
    out: List[str] = []

    def add(addr: str) -> None:
        a = safe_text(addr).strip()
        if not a:
            return
        key = a.lower()
        if key in seen:
            return
        seen.add(key)
        out.append(a)

    # base_list can be comma or semicolon separated
    normalized = safe_text(base_list).replace(",", ";")
    for part in normalized.split(";"):
        add(part)

    for e in extra:
        add(e)

    return "; ".join(out)


def ensure_placeholder_values(
    context: Dict[str, str],
    text_items: Sequence[str],
    *,
    prompt_fn=input,
    skip_keys: Iterable[str] | None = None,
) -> None:
    """Prompt for any [PLACEHOLDER] tokens required by the provided texts.

    This is a direct port of the VBScript behaviour: scan subject/headers/body/
    attachment strings for placeholder tokens, and if a token is not already
    present in the working context, ask the user for a value.
    """
    skip = set(k.upper() for k in (skip_keys or []))
    needed = extract_placeholders(text_items)

    for key in needed:
        if key in skip:
            continue
        current = safe_text(context.get(key, "")).strip()
        if current:
            continue

        # prompt_fn is injected for testability; keep it compatible with input()
        try:
            entered = prompt_fn(f"Value for [{key}]: ")
        except TypeError:
            entered = prompt_fn()
        context[key] = safe_text(entered).strip()
