# Shared Scaffold Examples

Stage 4 keeps shared behaviour concrete through anonymised examples. These
examples are not production inputs; they document the workbook and template
shapes used by the VBS-first scaffold.

## Example Workbooks

The scaffold represents workbook conventions with CSV examples so they can be
read and diffed without Office automation:

- `examples/workbooks/group_meetings_rota_example.csv`
- `examples/workbooks/induction_recipients_example.csv`
- `examples/workbooks/induction_template_manifest_example.csv`

The operational VBS path can later read real workbooks through Excel-specific
adapters, but the shared contract remains the same column and field shape.

The examples are also executable dry-run inputs through the root router:

```text
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
```

`fixtures/quickstart/dry_run_routes.json` is the shared quickstart command
fixture for the public help, validation, dry-run, and report-only routes. It
also records fail-safe blocked commands for unsafe requests so VBS and Python
validation can check the command surface without making Python the runtime.

Python support tooling can check the same public examples without becoming the
runtime path:

```text
python -B .\python\workflow_input_report.py --root . --format json --fail-if-invalid
python -B .\python\command_route_report.py --root . --format json --fail-if-invalid
```

## Example Templates

Template examples live under `examples/templates/` and use the same metadata and
placeholder conventions as the test fixtures:

- metadata comments such as `<!-- Subject: Welcome [NAME] -->`;
- `[TOKEN]` and `{{TOKEN}}` placeholders;
- no private names, addresses, links, file paths, or institutional identifiers.
- generic attachment and inline-image asset folders that exercise the shared
  mail asset contract without requiring Outlook.

## Scheduled Sidecar Example

`examples/task_scheduler/watchdog_report_only_task.xml` is a disabled Windows
Task Scheduler example for the report-only watchdog sidecar. It uses an
environment-variable placeholder rather than a private path and does not request
live inspection, repair, deletion, replacement, or sending.

`fixtures/watchdog/report_only_health_fixture.json` is the matching Stage 6
watchdog report fixture. It keeps the sidecar behaviour concrete without live
Outlook by representing managed stale mail, managed young mail, deferred mail,
unmarked user mail, and uncertain managed mail as a no-mutation health report.

`fixtures/watchdog/managed_repair_guardrail_fixture.json` captures the narrower
fake-object managed-repair boundary. It documents preview mode, committed
fake-object replacement, blocked transport, unmarked mail, and uncertain age
without requiring Outlook, sending mail, or mutating a real mailbox.

## Shared Contracts

Both group meetings and induction mailer examples map to:

- `WorkflowContext`
- `MessageSpec`
- `RecipientSpec`
- `DeliveryPolicy`
- `ManagedEmailKey`
- `DispatchResult`

The concrete cross-runtime fixture lives at
`fixtures/contracts/stage4_shared_contract_fixture.json`. It intentionally
contains one group-meeting message and one induction message so validators can
check that both workflows share the same `MessageSpec`, `RecipientSpec`,
`DeliveryPolicy`, `DispatchResult`, `WorkflowContext`, and `HealthReport`
shape.

For a dependency-free development summary, run:

```text
python -B .\python\contract_fixture_report.py --root . --format json --fail-if-invalid
```

The examples intentionally use `example.invalid` addresses and generic names.
Real operational material must remain outside the public scaffold.
