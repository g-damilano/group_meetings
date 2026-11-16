from __future__ import annotations

from typing import Any, Dict, List, Optional

from ..utils import CaseInsensitiveDict


def create_array_list() -> List[Any]:
    return []


def ensure_array_list(lst: Optional[List[Any]]) -> List[Any]:
    return lst if lst is not None else []


def create_text_dictionary() -> CaseInsensitiveDict:
    return CaseInsensitiveDict()


def get_collection_count(lst: Optional[List[Any]]) -> int:
    return len(lst) if lst else 0


def get_dictionary_count(d: Optional[Dict[Any, Any]]) -> int:
    return len(d) if d else 0
