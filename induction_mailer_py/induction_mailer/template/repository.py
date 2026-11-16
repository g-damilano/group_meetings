from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from ..core_types import Template
from ..utils import CaseInsensitiveDict, normalize_line_endings, read_text_smart, trim_leading_blank_lines


HTML_HEADER_PREFIX = "<!--"
HTML_HEADER_SUFFIX = "-->"
HEADER_PREFIX = "@--"
HEADER_SUFFIX = "--"


def extract_header_and_body(text: str) -> tuple[str, str]:
    header_text = ""
    body_text = text

    lower = text.lower()

    # HTML comment header
    start = lower.find(HTML_HEADER_PREFIX)
    if start >= 0:
        end = lower.find(HTML_HEADER_SUFFIX, start + len(HTML_HEADER_PREFIX))
        if end >= 0:
            header_text = text[start + len(HTML_HEADER_PREFIX) : end]
            body_text = text[end + len(HTML_HEADER_SUFFIX) :]
            return header_text, body_text

    # Plaintext header
    start = lower.find(HEADER_PREFIX)
    if start >= 0:
        end = lower.find(HEADER_SUFFIX, start + len(HEADER_PREFIX))
        if end >= 0:
            header_text = text[start + len(HEADER_PREFIX) : end]
            body_text = text[end + len(HEADER_SUFFIX) :]

    return header_text, body_text


def parse_header_block(block_text: str) -> CaseInsensitiveDict:
    d = CaseInsensitiveDict()
    block_text = block_text.strip("\n\r \t")
    if not block_text:
        return d

    lines = normalize_line_endings(block_text).split("\n")
    for line in lines:
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if key:
            d[key.lower()] = value
    return d


class TemplateRepository:
    """Recursively loads and parses templates from a folder."""

    def __init__(self, root_folder: str | Path):
        self.root = Path(root_folder)

    def load(self) -> Dict[str, Template]:
        out: Dict[str, Template] = {}
        counter = 0
        if not self.root.exists():
            return out

        for path in sorted(self.root.rglob("*")):
            if not path.is_file():
                continue
            ext = path.suffix.lower().lstrip(".")
            if ext not in {"htm", "html", "txt"}:
                continue
            tmpl = self.parse_template(path)
            if tmpl is None:
                continue
            counter += 1
            out[str(counter)] = tmpl
        return out

    # Back-compat with earlier CLI naming
    def load_templates(self) -> List[Template]:
        """Return templates as a sorted list.

        The VBScript workflow presents templates as an ordered list. The first
        Python port iteration called this `load_templates()` from the CLI.
        """
        return sort_templates(self.load().values())

    def parse_template(self, file_path: str | Path) -> Optional[Template]:
        path = Path(file_path)
        text = read_text_smart(path)
        if not text:
            return None

        header_text, body_text = extract_header_and_body(text)
        meta_text = normalize_line_endings(header_text)
        body_text = trim_leading_blank_lines(body_text)

        meta = parse_header_block(meta_text)

        display_name = path.name
        stage = str(meta.get("stage", "")).lower().strip()
        code = str(meta.get("code", "")).strip()

        return Template(
            path=path,
            body=body_text,
            meta=dict(meta.items()),
            display_name=display_name,
            stage=stage,
            code=code,
        )


def build_template_label(tmpl: Template) -> str:
    stage = (tmpl.stage or "").strip()
    code = (tmpl.code or "").strip()
    label = tmpl.file_name
    prefix = ""
    if stage:
        prefix = stage.capitalize()
    if code:
        prefix = f"{prefix} - {code}" if prefix else code
    return f"{prefix} | {label}" if prefix else label


def sort_templates(templates: Iterable[Template]) -> List[Template]:
    # Rough analogue of VBS SortTemplateList: stage, code, filename
    return sorted(
        list(templates),
        key=lambda t: (
            (t.stage or "~"),
            (t.code or ""),
            t.file_name.lower(),
        ),
    )
