# Publication Boundary

The scaffold is prepared inside `scaffold/`, but public repository release work
does not happen automatically from this folder.

The boundary is deliberate:

- legacy scripts outside `scaffold/` are frozen behavioural references;
- repository-level files and release configuration are outside the scaffold work
  root;
- publication may require reviewing material that scaffold development is not
  allowed to edit by default.

Scaffold-local production-shaped launchers are provided at:

- `RunGroupMeetings.vbs`
- `RunInductionMailer.vbs`
- `OutlookWatchdog.vbs`
- `ValidateWorkflow.vbs`

These are alternate wrappers for the rework only. They route to scaffold
dry-run, report-only, or validation entrypoints and may be edited inside
`scaffold/`. They are not moves or replacements of the frozen legacy scripts at
the repository root. The copy/port mapping is tracked in
`migration/legacy_alternate_map.json`.

`legacy/` is reserved for local hard-copy snapshots of frozen
legacy/reference/generated material. It is intentionally excluded from scaffold
release validation and release manifests, is ignored by `scaffold/.gitignore`,
and must not be wired into runtime entrypoints. Use it as a migration reference
only. `python/legacy_mirror_audit.py` may inspect that mirror in report-only
mode, and `python/legacy_sanitisation_plan.py` may write private remediation
plans under `legacy/`.

`python/release_candidate_plan.py` combines the scaffold release audit with the
repo boundary audit. It is a decision report only; outside-root findings are
reported as advisory for scaffold-local development unless an authorised
repo-wide release cleanup elects to treat them as blockers. The command does
not publish or package anything.

`python/release_gate.py` consumes that decision report and returns an explicit
`PASS` or `BLOCKED` publication gate for the scaffold release surface.
`python/release_export_plan.py` lists the deterministic public scaffold files
that would ship. Both commands are report-only by default and may write optional
JSON reports only under `legacy/`.

Before a public release, an authorised repo-wide pass should confirm:

- legacy/reference material is either excluded from publication or explicitly
  approved for publication;
- no private paths, real email addresses, names, logs, archives, message files,
  SharePoint links, or institutional identifiers are present in public content;
- any root-level README, license, packaging, CI, or release metadata changes are
  intentional and reviewed;
- generated artifacts are absent from the final release tree.

Scaffold-local commands such as the release audit and release manifest only
prepare evidence for that review. They do not publish, tag, package, or alter
anything outside `scaffold/`.
