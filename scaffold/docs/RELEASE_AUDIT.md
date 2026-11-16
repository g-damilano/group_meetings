# Release Audit

The release audit is a Stage 7 scaffold-only gate. It summarizes whether the
public scaffold is clean enough to prepare a release outside this folder.

Run:

```text
python scaffold/python/release_audit.py --root scaffold
python scaffold/python/release_audit.py --root scaffold --format json
python scaffold/python/legacy_intake_audit.py --repo-root . --scaffold-root scaffold
python scaffold/python/legacy_mirror_audit.py --root scaffold
python scaffold/python/legacy_sanitisation_plan.py --root scaffold
python scaffold/python/repo_boundary_audit.py --repo-root . --scaffold-root scaffold
python scaffold/python/release_manifest.py --root scaffold --format json
python scaffold/python/release_candidate_plan.py --root scaffold --repo-root .
python scaffold/python/release_gate.py --root scaffold --repo-root .
python scaffold/python/release_export_plan.py --root scaffold --format json
cscript.exe //NoLogo scaffold/vbs/tests/RunAllSmokeTests.vbs
cscript.exe //NoLogo scaffold/vbs/tests/PublicReleaseGateSmokeTest.vbs
cscript.exe //NoLogo scaffold/vbs/tests/ReleaseManifestSmokeTest.vbs
```

The audit checks:

- the standard scaffold validator passes;
- release-specific files are present;
- the scaffold-local license is present;
- workflow and entrypoint modules do not own Outlook COM creation, folder
  access, send, or delete behaviour;
- runtime VBS modules and production-shaped wrappers do not reference the
  private `legacy/` mirror;
- scaffold-local production-shaped alternate launchers are present and scanned
  for the same workflow safety rules;
- the legacy alternate map is present and every scaffold target it references
  exists;
- the read-only legacy intake audit can confirm mapped sources and scaffold
  alternates without exposing source content;
- the legacy mirror audit can inspect `legacy/` in report-only mode, verify
  copy-only guardrails, and summarize sanitisation work without exposing matched
  source content;
- the legacy sanitisation planner can turn mirror findings into a private
  content-free remediation plan without modifying source or copied files;
- the repo boundary audit tool is present for report-only inspection of
  outside-root advisory findings;
- any raw hard-copy mirror under `legacy/` is excluded from the public release
  manifest and scaffold sanitisation scan;
- the watchdog sidecar remains report-only;
- managed repair keeps its guarded delete path isolated to the repair module;
- the release manifest is deterministic, relative-path-only, and hash-backed;
- the release candidate planner can combine scaffold release evidence with
  repo-boundary status without publishing or writing outside `legacy/`;
- the release gate can turn that decision into an explicit `PASS` or `BLOCKED`
  result for automation without mutating files;
- the release export plan can enumerate the public scaffold release surface
  while keeping `legacy/` excluded;
- the VBS-first smoke suite can be run from a single scaffold-local runner;
- generated artifacts and public sanitisation blockers are absent.

This audit does not publish, tag, package, or modify anything outside
`scaffold/`. Public repository release actions still require explicit
authorisation because legacy/reference material outside the scaffold is frozen
by default.

The release audit registers the repo boundary audit tool but does not require
the surrounding repository to pass it. Run the boundary audit separately during
an authorised repo-wide release cleanup.
