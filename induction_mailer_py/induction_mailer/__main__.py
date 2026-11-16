"""Package entrypoint.

Supports both:
- python -m induction_mailer
- python induction_mailer   (running the package directory directly)

The latter does not always set package context, so we fall back to an
absolute import when relative imports are unavailable.
"""

from __future__ import annotations

try:
    # Normal execution path (python -m induction_mailer)
    from .cli import main
except ImportError:  # pragma: no cover
    # Fallback path (python induction_mailer)
    from induction_mailer.cli import main


if __name__ == "__main__":
    raise SystemExit(main())
