# Induction Mailer Dry-Run Pipeline

Stage 3 introduces a scaffold-only dry-run path for induction mailer workflows:

1. Load anonymised recipient and template-manifest CSV inputs.
2. Validate active recipients and enabled template rows for required fields and
   supported option values.
3. Build a broadcast/personal message plan.
4. Convert template rows into `MessageSpec` and `DeliveryPolicy`-style
   contracts.
5. Resolve recipients, placeholders, attachments, and inline images through the
   shared core layer.
6. Run the plan through the core workflow facade in dry-run mode.
7. Report shared dispatch states and counts.

The VBS entrypoint is:

```text
cscript.exe //NoLogo vbs\entrypoints\run_induction_mailer_dry_run.vbs /templates:tests\fixtures\induction_template_manifest_fixture.csv /recipients:tests\fixtures\induction_recipients_fixture.csv /mode:all
```

The script prints `key=value` lines and exits with:

- `0`: dry-run completed without action required;
- `1`: failed-safe boundary or composition failure;
- `2`: action required;
- `64`: missing required arguments.

The entrypoint blocks `/send`, `/live`, and `dry_run=false` requests before
loading inputs or planning messages.

The induction workflow modules do not create Outlook sessions, open folders,
send mail, delete mail, or attach files directly. They build workflow contracts
and delegate those operations to reusable core primitives, which are exercised
with fixture-backed dry-run smoke tests.

Missing enabled template files become explicit `ACTION_REQUIRED` results rather
than empty composed messages. This keeps template migration failures visible
without requiring a live mailbox or destructive repair.

Malformed active recipient rows and malformed enabled template rows also stop
before planning. The adapter reports explicit issue codes such as
`ACTIVE_RECIPIENT_MISSING_EMAIL`, `TEMPLATE_MODE_INVALID`,
`DELIVERY_ACTION_INVALID`, and `DEFER_MINUTES_INVALID`.
