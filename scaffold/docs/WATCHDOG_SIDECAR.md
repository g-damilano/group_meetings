# Watchdog Sidecar

Stage 6 starts with a report-only watchdog sidecar. It can inspect scaffold
managed Outbox state through core diagnostics, but it does not repair, replace,
delete, send, or otherwise modify mail.

The entrypoint is:

```text
cscript.exe //NoLogo vbs\entrypoints\run_watchdog_report_only.vbs
```

It prints key-value output for scheduler logs and operator checks:

```text
workflow_reference=outlook-watchdog
state=COMPOSED
report_only=true
total_count=0
managed_count=0
unmanaged_count=0
stale_count=0
replace_eligible_count=0
repair_actions_taken=0
```

These counters are report-only. `replace_eligible_count` means a managed stale
message would require action; the report-only sidecar still performs no repair,
replacement, deletion, or sending.

By default it does not inspect live Outlook. This makes the entrypoint safe for
validation and Task Scheduler installation checks. Operational live inspection
must be explicitly requested:

```text
cscript.exe //NoLogo vbs\entrypoints\run_watchdog_report_only.vbs /inspect_live:true
```

Repair-like arguments such as `/repair`, `/delete`, `/send`, and `/replace` are
blocked with `FAILED_SAFE`.

A disabled, report-only Task Scheduler XML example is provided at:

```text
examples\task_scheduler\watchdog_report_only_task.xml
```

The example uses the `%OFFICE_SCAFFOLD_ROOT%` environment variable placeholder,
runs `cscript.exe` against `run_watchdog_report_only.vbs`, and deliberately does
not pass `/inspect_live:true` or any repair-like argument. Operators should copy
and adapt it outside the public scaffold only after setting the environment
variable locally.

The doctrine remains:

- inspect broadly;
- report clearly;
- modify narrowly in future guarded work;
- only modify scaffold-managed messages;
- never modify unmarked user mail;
- prefer `ACTION_REQUIRED` or `FAILED_SAFE` when uncertain.

## Report Fixture

The canonical Stage 6 report fixture lives at:

```text
fixtures/watchdog/report_only_health_fixture.json
```

It models a mixed Outbox with stale scaffold-managed mail, young queued mail,
future-deferred mail, unmarked user mail, and an uncertain managed item. The
fixture is explicitly `report_only`, requires no live Outlook session, performs
zero mutations, and records unmarked or uncertain items as non-repairable.

Python validation checks the fixture counts, dispatch states, managed-key
guardrails, and mutation flags. The VBS watchdog smoke test also requires the
fixture and verifies that it contains no committed send, delete, replacement, or
repair action.

The companion managed-repair guardrail fixture lives at:

```text
fixtures/watchdog/managed_repair_guardrail_fixture.json
```

It is still fake-object-only and does not change the public sidecar entrypoint.
It documents the repair branch that VBS fake-object tests exercise: preview
mode is action-required and non-mutating, committed fake-object repair replaces
only stale scaffold-managed mail, blocked transport performs no replacement,
unmarked mail is never mutable, and uncertain age fails safe.
