# Release Notes Template

Use this template after the scaffold has passed the public release gates and a
repo-wide publication pass has been authorised.

## Summary

- VBS-first Office/O365 communication automation scaffold.
- Reusable core Office automation primitives.
- Group meetings and induction mailer dry-run workflow adapters.
- Report-first Outlook watchdog sidecar with managed-only repair guardrails.

## Safety Guarantees

- Production legacy scripts remain frozen references.
- Operational entrypoints default to dry-run or report-only behaviour.
- No real email sending is performed by validation.
- Watchdog repair is limited to scaffold-managed messages with valid managed
  keys and remains unavailable from the live report-only entrypoint.
- Unmarked user-authored mail is never modified.

## Validation Evidence

Record the results of:

```text
cscript.exe //NoLogo scaffold\vbs\tests\ValidateScaffold.vbs
cscript.exe //NoLogo scaffold\vbs\tests\PublicReleaseGateSmokeTest.vbs
cscript.exe //NoLogo scaffold\vbs\tests\ReleaseManifestSmokeTest.vbs
python scaffold\python\validate_scaffold.py --root scaffold
python scaffold\python\release_audit.py --root scaffold
python scaffold\python\release_manifest.py --root scaffold --format json
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root .
python scaffold\python\release_gate.py --root scaffold --repo-root . --format json
python scaffold\python\release_export_plan.py --root scaffold --format json
python -m unittest discover -s scaffold\python\tests
```

## Known Boundaries

- Live mailbox repair or real sending requires explicit operator promotion and
  additional safeguards.
- Repository-wide publication work is outside the scaffold boundary unless
  explicitly authorised.

## Release Contents

Attach or reference the generated release manifest for the exact scaffold file
inventory and hashes, plus the export plan for the reviewed public scaffold
surface.
