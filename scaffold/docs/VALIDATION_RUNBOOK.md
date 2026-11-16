# Validation Runbook

This runbook gives operators and maintainers a repeatable scaffold-only
validation sequence. It does not perform real email sending, live mailbox
repair, or destructive Outlook actions.

Use the VBS smoke suite first when validating the operational path:

```text
cscript.exe //NoLogo .\vbs\tests\RunAllSmokeTests.vbs
```

The preferred public scaffold interface is the root command router:

```text
cscript.exe //NoLogo .\Scaffold.vbs help
cscript.exe //NoLogo .\Scaffold.vbs validate
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
cscript.exe //NoLogo .\Scaffold.vbs release-check
```

`Scaffold.vbs` is orchestration-only. It delegates to existing scaffold
entrypoints, keeps group meetings and induction in dry-run mode, keeps watchdog
diagnostics report-only, and treats parent repository boundary findings as
advisory unless `release-check --strict-repo-boundary` is explicitly requested.
The safe public command surface used by examples and validation is captured in
`fixtures/quickstart/dry_run_routes.json`.

Production-shaped scaffold alternates can also be exercised directly:

```text
cscript.exe //NoLogo .\RunGroupMeetings.vbs /rota:vbs\tests\fixtures\group_meetings_rota_fixture.csv /date:2026-05-29 /directory:vbs\tests\fixtures\group_meetings_recipients_fixture.csv
cscript.exe //NoLogo .\RunInductionMailer.vbs /templates:vbs\tests\fixtures\induction_template_manifest_fixture.csv /recipients:vbs\tests\fixtures\induction_recipients_fixture.csv /mode:all
cscript.exe //NoLogo .\OutlookWatchdog.vbs
cscript.exe //NoLogo .\ValidateWorkflow.vbs
```

These wrappers route only to scaffold dry-run, report-only, or validation
entrypoints. They do not replace or modify legacy root scripts.

Then run the development-side Python checks:

```text
python -B -m unittest discover -s .\python\tests
python -B .\python\validate_scaffold.py --root .
python -B .\python\contract_fixture_report.py --root . --format json --fail-if-invalid
python -B .\python\workflow_input_report.py --root . --format json --fail-if-invalid
python -B .\python\command_route_report.py --root . --format json --fail-if-invalid
```

Release tooling is normally reserved for release-readiness work. If it is
needed, keep it scaffold-local and report-only unless a strict release gate is
being intentionally exercised:

```text
python -B .\python\legacy_mirror_audit.py --root .
python -B .\python\legacy_sanitisation_plan.py --root .
python -B .\python\release_audit.py --root .
python -B .\python\release_candidate_plan.py --root . --repo-root .
python -B .\python\release_gate.py --root . --repo-root .
python -B .\python\release_export_plan.py --root . --fail-if-invalid
python -B .\python\release_manifest.py --root . --format json
```

Expected safety properties:

- workflow and entrypoint modules do not create Outlook or Excel sessions;
- workflow and entrypoint modules do not call `GetDefaultFolder`, `.Send`, or
  `.Delete`;
- runtime VBS modules and production-shaped wrappers do not read from
  `legacy/`;
- scaffold-local production-shaped launchers are mapped in
  `migration/legacy_alternate_map.json`;
- any hard-copy legacy mirror is report-only audited, excluded from release,
  and not used as a runtime source;
- any mirror sanitisation plan is report-only and writes private output only
  under `legacy/`;
- the release candidate planner remains report-only and treats repository
  boundary findings outside `scaffold/` as advisory for scaffold-local
  development;
- the release gate remains report-only and blocks only scaffold-local release
  failures, private public-scaffold material, accidental `legacy/` inclusion, or
  broken smoke/tests;
- the export plan remains deterministic, archive-free, and excludes `legacy/`;
- the shared contract fixture report remains report-only and confirms
  cross-runtime dispatch-state parity;
- the quickstart command fixture remains dry-run/report-only and includes
  fail-safe blocked examples for unsafe live-mode requests;
- repository boundary inspection is report-only unless explicitly run with
  `--fail-on-findings`;
- watchdog report-only entrypoint does not send or delete mail;
- managed repair tests use fake objects and never touch a live mailbox;
- unmarked mail is never modified by scaffold validation.
