# Legacy Mirror

`scaffold/legacy/` is a local quarantine area for hard-copy snapshots of
legacy/reference/generated repository material. It exists only to unblock
scaffold migration work when the original files must remain frozen outside the
scaffold boundary.

Rules:

- copy only; never move, rename, delete, or rewrite the original legacy files;
- do not wire runtime entrypoints to raw files under `legacy/`;
- do not treat `legacy/` as public scaffold content;
- do not publish `legacy/` unless a separate, explicit public-release cleanup
  has reviewed and sanitised every copied file;
- keep the scaffold implementation, tests, fixtures, contracts, and
  documentation outside `legacy/`.

The public release validator and manifest generator exclude `legacy/` by
default. The repository boundary audit also skips `scaffold/` entirely, so the
mirror is a private migration aid rather than a new release gate.

When a mirror is created, `legacy/manifest.json` should describe the snapshot at
a high level without exposing source content. The manifest is also private local
material because it lives under `legacy/`.

Audit the mirror without changing it:

```text
python scaffold\python\legacy_mirror_audit.py --root scaffold
python scaffold\python\legacy_mirror_audit.py --root scaffold --format json
python scaffold\python\legacy_mirror_audit.py --root scaffold --write-json legacy\reports\mirror_audit.json
```

The audit reports paths, categories, counts, and guardrail status only. It does
not print matched private text, modify copied files, modify original legacy
files, publish content, send email, or touch Outlook. Written reports must stay
under `legacy/`, which remains excluded from the public release surface.

Build a private, content-free remediation plan:

```text
python scaffold\python\legacy_sanitisation_plan.py --root scaffold
python scaffold\python\legacy_sanitisation_plan.py --root scaffold --write-json legacy\reports\sanitisation_plan.json
```

The plan groups findings by copied path and recommends generic actions such as
removing generated artifacts, reviewing binary files, or rewriting/redacting
text. It is still a private report because copied path names can be sensitive.
