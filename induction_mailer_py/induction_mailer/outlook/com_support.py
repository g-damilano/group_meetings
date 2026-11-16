from __future__ import annotations

import platform
from typing import Any, Tuple


def is_windows() -> bool:
    return platform.system().lower() == "windows"


def require_win32com() -> Tuple[Any, Any]:
    """Lazy import for pywin32.

    Returns (win32com, pythoncom).
    """
    if not is_windows():
        raise RuntimeError("Outlook COM automation requires Windows.")
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception as e:
        raise RuntimeError(
            "pywin32 is required for Outlook COM. Install with: pip install -e '.[outlook]'"
        ) from e
    return win32com.client, pythoncom
