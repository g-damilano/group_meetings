# Public Release Checklist

Stage 7 public sanitisation is a release gate, not a workflow feature. Before
publishing the scaffold, confirm:

- all runtime code lives under `scaffold/`;
- the scaffold-local `LICENSE` is present and reviewed;
- legacy scripts outside `scaffold/` remain frozen references;
- no real names, email addresses, private paths, SharePoint links, internal
  logs, zip bundles, message files, or production-only attachments are included;
- `scaffold/legacy/`, if present as a hard-copy mirror, is treated as private
  local migration material and excluded from any public release;
- all public email examples use `example.invalid`;
- no generated `*.log`, `*.pyc`, `__pycache__`, or archive artifacts remain;
- VBS operational entrypoints remain dry-run/report-only unless explicitly
  promoted with guarded operator controls;
- watchdog repair stays managed-key-only and never touches unmarked mail.
- the publication boundary and release notes template have been reviewed before
  any repo-wide release action.
- the release gate is `PASS` for scaffold-local validation; parent repository
  findings outside `scaffold/` are advisory unless a separate repo-wide release
  pass has been explicitly authorised.
- the export plan reports `legacy_paths_in_files: 0`.

Validation commands:

```text
cscript.exe //NoLogo scaffold\vbs\tests\ValidateScaffold.vbs
cscript.exe //NoLogo scaffold\vbs\tests\RunAllSmokeTests.vbs
cscript.exe //NoLogo scaffold\vbs\tests\PublicReleaseGateSmokeTest.vbs
cscript.exe //NoLogo scaffold\vbs\tests\ReleaseManifestSmokeTest.vbs
python scaffold\python\validate_scaffold.py --root scaffold
python scaffold\python\legacy_intake_audit.py --repo-root . --scaffold-root scaffold
python scaffold\python\legacy_intake_audit.py --repo-root . --scaffold-root scaffold --format json
python scaffold\python\legacy_mirror_audit.py --root scaffold
python scaffold\python\legacy_mirror_audit.py --root scaffold --format json
python scaffold\python\legacy_sanitisation_plan.py --root scaffold
python scaffold\python\legacy_sanitisation_plan.py --root scaffold --format json
python scaffold\python\repo_boundary_audit.py --repo-root . --scaffold-root scaffold
python scaffold\python\repo_boundary_audit.py --repo-root . --scaffold-root scaffold --format json
python scaffold\python\release_audit.py --root scaffold
python scaffold\python\release_audit.py --root scaffold --format json
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root .
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root . --format json
python scaffold\python\release_gate.py --root scaffold --repo-root .
python scaffold\python\release_gate.py --root scaffold --repo-root . --format json
python scaffold\python\release_gate.py --root scaffold --repo-root . --fail-if-blocked
python scaffold\python\release_export_plan.py --root scaffold
python scaffold\python\release_export_plan.py --root scaffold --format json
python scaffold\python\release_export_plan.py --root scaffold --fail-if-invalid
python scaffold\python\release_manifest.py --root scaffold
python scaffold\python\release_manifest.py --root scaffold --format json
python -m unittest discover -s scaffold\python\tests
```

The checklist complements automated validation. It should be reviewed before
creating any public GitHub release.

See also:

- `docs/PUBLICATION_BOUNDARY.md`
- `docs/LEGACY_MIRROR.md`
- `docs/REPO_BOUNDARY_AUDIT.md`
- `docs/RELEASE_GATE.md`
- `docs/RELEASE_EXPORT_PLAN.md`
- `docs/RELEASE_NOTES_TEMPLATE.md`
- `docs/VALIDATION_RUNBOOK.md`
