# Group Meetings Dry-Run Entrypoint

`vbs/entrypoints/run_group_meetings_dry_run.vbs` is the first scaffold
orchestration entrypoint for the group-meetings workflow. It gathers a target
date, rota source, optional recipient directory, and display names from command
line arguments, then calls the Stage 2 dry-run runner.

It is intentionally dry-run only. Requests for `dry_run=false`, `/live`, or
`/send` are blocked with `FAILED_SAFE` before rota loading or message planning.

## Example

```text
cscript.exe //NoLogo vbs\entrypoints\run_group_meetings_dry_run.vbs /rota:tests\fixtures\group_meetings_rota_fixture.csv /date:2026-05-29 /directory:tests\fixtures\group_meetings_recipients_fixture.csv
```

The script prints machine-readable `key=value` lines:

- `state`
- `reason`
- `action_required`
- `dry_run`
- `target_date`
- `row_count`
- `planned_count`
- `composed_count`
- `failed_count`
- `skipped_count`

Exit codes are:

- `0`: dry-run completed without action required;
- `1`: failed-safe boundary or composition failure;
- `2`: action required, such as incomplete rota setup;
- `64`: missing required arguments.

## Recipient Directory

The optional directory CSV has two columns:

```text
alias,email
group,group@example.invalid
manager,manager@example.invalid
```

The reserved `example.invalid` domain is used only in fixtures and public
documentation. Operational directories must be supplied locally and must not be
committed with real names or addresses.

## Layering

This entrypoint belongs to the UI and orchestration layer. It does not own
Outlook COM, Excel COM, recipient resolution internals, message composition, or
dispatch state decisions. Those remain in the workflow and core layers.
