from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, List, Tuple

from ..utils import safe_text


def is_absolute_path(path: str) -> bool:
    p = safe_text(path)
    if len(p) < 2:
        return False
    if len(p) >= 2 and p[1] == ":":
        return True
    if p.startswith("\\\\"):
        return True
    return os.path.isabs(p)


def resolve_attachments(raw_value: str, base_folder: str | Path) -> List[str]:
    cleaned = safe_text(raw_value).strip().replace(",", ";")
    if not cleaned:
        return []

    base = Path(base_folder)
    out: List[str] = []
    for part in cleaned.split(";"):
        item = part.strip()
        if not item:
            continue

        cand = item
        p = Path(cand)
        if not p.exists():
            # handle relative like ./file or .\file
            if item.startswith("./") or item.startswith(".\\"):
                p = base / item[2:]
            elif not is_absolute_path(item):
                p = base / item

        if p.exists():
            out.append(str(p))
        # missing attachments are just ignored (VBS logs)

    return out


def sanitize_content_id(file_name: str) -> str:
    # VBS: spaces -> underscore; keep basename
    return Path(file_name).name.replace(" ", "_")


def rewrite_inline_references(html_text: str, template_base_name: str, file_name: str) -> str:
    cid = sanitize_content_id(file_name)
    encoded_base = template_base_name.replace(" ", "%20")
    encoded_file = file_name.replace(" ", "%20")
    target = f"cid:{cid}"

    # folder references
    html_text = html_text.replace(f"{template_base_name}.files/{file_name}", target)
    html_text = html_text.replace(f"{template_base_name}.files\\{file_name}", target)
    html_text = html_text.replace(f"{encoded_base}.files/{encoded_file}", target)
    html_text = html_text.replace(f"{encoded_base}.files\\{encoded_file}", target)

    # bare attribute
    html_text = html_text.replace(f'="{file_name}"', f'="{target}"')
    html_text = html_text.replace(f"='{file_name}'", f"='{target}'")
    html_text = html_text.replace(f'="{encoded_file}"', f'="{target}"')
    html_text = html_text.replace(f"='{encoded_file}'", f"='{target}'")

    return html_text


def prepare_inline_image_support(html_body: str, template_path: str | Path) -> Tuple[str, str]:
    """Return (rewritten_html, inline_folder).

    Matches VBS `PrepareInlineImageSupport`: if a folder named `<templateBase>.files` exists next to the template,
    rewrite references and return that folder for CID attachment.
    """
    path = Path(template_path)
    base_folder = path.parent
    base_name = path.stem

    inline_folder = base_folder / f"{base_name}.files"
    if not inline_folder.exists() or not inline_folder.is_dir():
        return html_body, ""

    rewritten = html_body
    for f in inline_folder.iterdir():
        if f.is_file():
            rewritten = rewrite_inline_references(rewritten, base_name, f.name)

    return rewritten, str(inline_folder)


# Back-compat with CLI / VBScript naming
def resolve_attachment_paths(raw_value: str, base_folder: str | Path) -> List[str]:
    """Resolve attachment paths declared in template metadata.

    `raw_value` may be a semicolon/comma-separated list of paths. Relative paths
    are resolved against `base_folder` (the template folder).
    """
    return resolve_attachments(raw_value, base_folder)
