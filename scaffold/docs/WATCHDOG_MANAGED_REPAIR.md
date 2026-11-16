# Watchdog Managed Repair

Stage 6 managed repair is intentionally narrower than report-only inspection.
Repair is allowed only when all of these are true:

- the Outbox item has a valid scaffold managed email key;
- the key on the item matches the expected key;
- Outlook transport diagnostics indicate a connected, non-blocked state;
- the item is not waiting for a future deferred-delivery time;
- the queued item age can be determined and exceeds the configured grace period;
- a replacement factory creates a replacement mail item carrying the same
  managed key;
- repair has been explicitly committed by the caller.

Unmarked mail, wrong-key mail, blocked transport, future-deferred mail, young
queued mail, uncertain item age, missing replacement creation, and wrong-key
replacement all resolve to no action, `ACTION_REQUIRED`, or `FAILED_SAFE`.

The public watchdog entrypoint remains report-only. The managed repair primitive
is tested with fake mail objects only, so scaffold validation never modifies a
real mailbox.

The report-only fixture at
`fixtures/watchdog/report_only_health_fixture.json` documents the repair
boundary in data form. Replacement eligibility appears only on stale,
scaffold-managed mail with a portable managed key and connected transport, and
even that item has `would_modify=false` because the fixture represents a report
run. Unmarked user mail and uncertain managed mail remain non-repairable and
resolve to `ACTION_REQUIRED` or `FAILED_SAFE`.

The managed-repair guardrail fixture at
`fixtures/watchdog/managed_repair_guardrail_fixture.json` documents the narrower
fake-object repair path used by tests. It covers preview mode, committed
fake-object replacement, blocked transport, unmarked mail, and uncertain age.
The fixture keeps `live_outlook_required=false`,
`real_mailbox_mutation_allowed=false`, `send_allowed=false`, and
`unmarked_mail_mutation_allowed=false`.

Committed repair in the public scaffold is therefore a fake-object validation
case, not an operational mailbox action. A future operational repair runner must
opt in separately, keep the same managed-key guardrails, and remain outside the
report-only Task Scheduler example until it has explicit deployment approval.
