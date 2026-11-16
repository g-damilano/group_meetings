# Scaffold Orchestration

`Scaffold.vbs` is the preferred user-facing command router for the scaffold.
It lives at the scaffold root and delegates to existing dry-run, report-only,
validation, and release-check paths.

## Commands

Run these from the scaffold root:

```text
cscript.exe //NoLogo .\Scaffold.vbs help
cscript.exe //NoLogo .\Scaffold.vbs validate
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input <rota.csv> [--date <yyyy-mm-dd>] [--directory <recipients.csv>]
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients <recipients.csv> [--templates <manifest.csv>]
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
cscript.exe //NoLogo .\Scaffold.vbs release-check [--strict-repo-boundary]
```

The public-safe example inputs exercise both workflow routes without Office or
Outlook side effects:

```text
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
```

`group-meetings` defaults to dry-run. If `--date` is omitted, the router uses
the first `meeting_date` value in the input CSV as the target date. Supplying a
date explicitly is clearer for operational runs.

`induction` defaults to dry-run. If `--templates` is omitted, the router uses
the public example template manifest under `examples/workbooks/`.

`watchdog` delegates to the report-only sidecar and does not request live
inspection or repair.

`release-check` runs the scaffold-local release gate. Parent repository
boundary findings are advisory by default. `--strict-repo-boundary` also runs
the existing repository-boundary audit in strict mode.

## Safety Boundary

The router may parse arguments, print help, call existing entrypoints, and relay
exit codes. It must not own low-level Office behaviour. It must not create
Outlook or Excel application objects, call Outlook folder APIs, send mail,
delete mail, inspect mailbox items directly, or duplicate workflow planning
logic.

Unsafe options such as `--send`, `--live`, `--repair`, `--delete`,
`--replace`, and `--inspect-live` are blocked with exit code `3`.

## Exit Codes

```text
0 = success
1 = validation or delegated command failure
2 = invalid arguments or unknown command
3 = unsafe requested operation blocked
4 = missing required input
5 = internal scaffold error
```

The compatibility wrappers remain available for legacy-shaped operator habits,
but new documentation and validation should prefer `Scaffold.vbs`.
