# Office Communication Automation Scaffold

This folder is the development and deployment root for a reusable,
VBS-first Office/O365 communication automation scaffold.

The scaffold is designed for corporate Windows environments where operators may
have Outlook and Excel available, but may not be allowed to install Python,
Python packages, unsigned executables, or custom dependencies. VBS is therefore
the operational path. Python remains a development-side validator and test
harness, not the required runtime.

Everything outside this scaffold root is frozen legacy, private reference
material, or repository control material unless a specific run explicitly says
otherwise. The `legacy/` folder, when present, is a quarantined private mirror
for reference only and is not a runtime source.

## What This Scaffold Does

The scaffold turns recurring Office communication tasks into safe, repeatable
workflows. It separates low-level Office automation from workflow decisions so
new workflows can reuse the same contracts, safety rules, fixtures, and dry-run
validation.

Current product-facing capabilities:

- one VBS-first command router: `Scaffold.vbs`;
- group-meetings dry-run workflow from a rota-shaped CSV;
- induction mailer dry-run workflow from recipient and template-manifest CSVs;
- shared template rendering with `[TOKEN]` and `{{TOKEN}}` placeholders;
- recipient resolution from direct addresses or alias/display-name directories;
- attachment and inline-image discovery through shared mail asset primitives;
- `MessageSpec`, `DeliveryPolicy`, `ManagedEmailKey`, `DispatchResult`,
  `WorkflowContext`, and `HealthReport` contract shapes;
- managed dispatch identity for duplicate suppression and Outbox diagnostics;
- report-only watchdog sidecar with fake-object-tested repair guardrails;
- VBS smoke tests for the operational path;
- Python validators for schemas, fixtures, examples, command routes, and
  public-safe scaffold hygiene.

The scaffold does not perform real sends by default. Workflow routes are
dry-run or report-only unless explicitly and safely extended.

## Safety Model

Safety is the main design constraint.

Default behaviour:

- group meetings: dry-run only;
- induction mailer: dry-run only;
- watchdog: report-only by default;
- validation: local fixture and smoke-test checks only;
- release check: scaffold-local, report-only by default.

The root router blocks unsafe flags such as `--send`, `--live`, `--repair`,
`--delete`, `--replace`, and `--inspect-live` with exit code `3`.

Workflow modules must not:

- create `Outlook.Application`;
- create `Excel.Application`;
- call `GetDefaultFolder`;
- call `.Send`;
- call `.Delete`;
- inspect or modify mailbox items directly;
- duplicate low-level dispatch, template, attachment, or recipient logic.

Workflow modules should:

- load and validate workflow inputs;
- build message plans;
- build shared contract objects;
- delegate recipient, template, asset, composition, and dispatch work to core
  modules;
- return clear `COMPOSED`, `ACTION_REQUIRED`, or `FAILED_SAFE` style results.

## Architecture

The scaffold has three layers.

### 1. Core Office Automation Layer

Location: `vbs/core/`

Core modules hold reusable primitives:

- `contracts.vbs`: shared states and contract helpers;
- `session.vbs`: Outlook/Excel session acquisition helpers;
- `folder_access.vbs`: namespace, account, Outbox, and Sent Items access;
- `recipient_resolver.vbs`: address and directory resolution;
- `template_engine.vbs`: metadata and placeholder rendering;
- `mail_assets.vbs`: attachment and inline-image handling;
- `mail_composer.vbs`: `MessageSpec` preparation;
- `managed_dispatch.vbs`: managed keys, dry-run dispatch, and diagnostics;
- `outbox_inspector.vbs`: managed Outbox inspection;
- `health_diagnostics.vbs`: transport and health summaries;
- `workflow_facade.vbs`: shared workflow-to-core bridge;
- `bootstrap.vbs`: module loading.

Core modules can contain Office-specific primitives, but tests keep live effects
behind dry-run, fake-object, or report-only paths.

### 2. Workflow Pipeline Layer

Location: `vbs/workflows/`

Workflow modules decide what needs to happen and build message specifications.
They do not own low-level Office COM behaviour.

Current workflows:

- group meetings:
  - rota adapter;
  - message-plan builder;
  - dry-run runner;
  - compatibility documentation.
- induction mailer:
  - recipient/template input adapter;
  - broadcast and personal message planner;
  - dry-run runner;
  - compatibility documentation.

### 3. UI And Orchestration Layer

Location: scaffold root and `vbs/entrypoints/`

Orchestration scripts gather intent and call workflow runners. The preferred
front door is:

```text
cscript.exe //NoLogo .\Scaffold.vbs help
```

Compatibility wrappers remain available:

- `RunGroupMeetings.vbs`;
- `RunInductionMailer.vbs`;
- `OutlookWatchdog.vbs`;
- `ValidateWorkflow.vbs`.

They route to scaffold dry-run, report-only, or validation entrypoints and do
not modify frozen legacy scripts.

## Quick Start

From the scaffold root:

```text
cscript.exe //NoLogo .\Scaffold.vbs help
cscript.exe //NoLogo .\Scaffold.vbs validate
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
```

If launching from the parent repository, prefix paths with `scaffold\`.

The public quickstart routes are captured in:

```text
fixtures\quickstart\dry_run_routes.json
```

That fixture is used by Python validation to prove the documented command
surface stays executable and safe.

## Command Reference

### Help

```text
cscript.exe //NoLogo .\Scaffold.vbs help
```

Shows available commands, safety defaults, examples, and exit-code meanings.

### Validate

```text
cscript.exe //NoLogo .\Scaffold.vbs validate
```

Runs the scaffold VBS validation wrapper. This checks expected scaffold files,
core state constants, and deterministic managed-key generation. It does not
send mail or inspect a live mailbox.

### Group Meetings

```text
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input <rota.csv> [--date <yyyy-mm-dd>] [--directory <recipients.csv>]
```

Purpose:

- load a rota-shaped input;
- select a target meeting date;
- build reminders, same-day notices, regular distribution messages, no-meeting
  notices, or manager alerts;
- compose message specifications through the shared core in dry-run mode;
- print key-value status output.

If `--date` is omitted, the route uses the first `meeting_date` in the input.
Providing `--date` is clearer for operational use.

Example:

```text
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv --date 2026-05-29
```

Expected shape of successful output:

```text
scaffold_command=group-meetings
state=COMPOSED
dry_run=true
planned_count=6
composed_count=6
failed_count=0
```

### Induction Mailer

```text
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients <recipients.csv> [--templates <manifest.csv>]
```

Purpose:

- load recipient and template-manifest inputs;
- validate active recipients and enabled template rows;
- build broadcast and personal message plans;
- render template placeholders;
- resolve recipients, attachments, inline images, deferred-delivery metadata,
  and delivery actions through the shared core;
- compose message specifications in dry-run mode;
- print key-value status output.

If `--templates` is omitted, the public example template manifest is used. In
real local use, pass an explicit manifest path.

Example:

```text
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
```

Expected shape of successful output:

```text
scaffold_command=induction
state=COMPOSED
dry_run=true
recipient_count=3
active_recipient_count=2
planned_count=3
composed_count=3
failed_count=0
```

### Watchdog

```text
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
```

Purpose:

- run the watchdog sidecar in report-only mode;
- summarize managed/unmanaged/queued/stale/deferred counts;
- take no repair actions;
- avoid live Outlook inspection unless a guarded future path explicitly adds
  it.

Expected shape of successful output:

```text
scaffold_command=watchdog
state=COMPOSED
report_only=true
repair_actions_taken=0
```

### Release Check

```text
cscript.exe //NoLogo .\Scaffold.vbs release-check
```

Runs the existing scaffold-local release gate in report-only mode. Parent
repository boundary findings are advisory unless strict mode is explicitly
requested:

```text
cscript.exe //NoLogo .\Scaffold.vbs release-check --strict-repo-boundary
```

This command is for release readiness. It is not required for normal product
development or dry-run workflow operation.

## Exit Codes

```text
0 = success
1 = validation or delegated command failure
2 = invalid arguments or unknown command
3 = unsafe requested operation blocked
4 = missing required input
5 = internal scaffold error
```

Use exit codes in scheduled tasks, wrapper scripts, and CI jobs. A route that
returns `ACTION_REQUIRED` may still be safe; it means the input or workflow
state needs operator review.

## Inputs And Conventions

### Group-Meetings Rota

Public example:

```text
examples\workbooks\group_meetings_rota_example.csv
```

Important columns include:

- `meeting_date`;
- presenter channel columns, currently represented by
  `literature_presenter` and `results_presenter` in the compatibility CSV;
- `meeting_lead`;
- `lab_manager`;
- `room`;
- reminder and same-day flags.

The CSV is a public-safe workbook stand-in. Real workbook adapters should keep
the same normalized contract shape so workflows can continue producing the same
message plans. Implementations may map different workbook headings or future
presenter channel names onto the same presenter-reminder concept, provided the
adapter preserves the channel identity and the dry-run output remains stable.

### Induction Recipients

Public example:

```text
examples\workbooks\induction_recipients_example.csv
```

Important columns include:

- `alias`;
- `email`;
- `first_name`;
- `last_name`;
- `display_name`;
- `active`.

Only active recipients are targeted by personal messages. Directory aliases can
also be used for CC/BCC-style template metadata.

Recipient resolution deliberately follows a conservative Outlook-like local
order: direct SMTP or `Display Name <email>` text first, exact local directory
keys second, and normalized display-name fallback last. Display names are not
treated as authoritative identity; ambiguous local display matches remain
unresolved so the workflow can ask for review instead of guessing.

For personalization, `first_name` and `last_name` columns win. If they are
missing, the scaffold can derive missing `FIRST_NAME`, `LAST_NAME`, `SURNAME`,
and `FULL_NAME` values from `full_name`, `display_name`, or an explicit
`Display Name <email>` value. That parser is a best-effort display-name
ontology for template variables, not a replacement for structured recipient
data.

### Induction Template Manifest

Public example:

```text
examples\workbooks\induction_template_manifest_example.csv
```

Important columns include:

- `template_id`;
- `mode` (`broadcast` or `personal`);
- `target_key`;
- `template_path`;
- `inline_image_folder`;
- `attachment_base_folder`;
- `attachments`;
- `cc`;
- `bcc`;
- `delivery_action`;
- `defer_minutes`;
- `urgent_verify_after_submit`;
- `enabled`.

Enabled rows must reference existing templates. Missing enabled templates or
malformed active rows become `ACTION_REQUIRED` instead of silently composing
empty or partial messages.

### Templates

Public examples:

```text
examples\templates\induction_broadcast_example.html
examples\templates\induction_personal_example.html
```

Templates can use metadata comments and placeholders, for example:

```text
<!-- Subject: Welcome [FIRST_NAME] -->
Hello [FIRST_NAME],
Your setup checklist is attached.
```

Both `[TOKEN]` and `{{TOKEN}}` placeholder styles are supported by the shared
template engine.

### Attachments And Inline Images

Attachment paths are resolved from the manifest's `attachment_base_folder`.
Inline image folders are resolved through the shared mail asset layer. Workflow
code should pass metadata to the core and let `mail_assets.vbs` handle path and
content-id behaviour.

### Managed Email Keys

Managed email keys identify scaffold-owned messages. They are used for dry-run
dispatch results, duplicate suppression, Outbox diagnostics, and watchdog
guardrails.

The key concept is shared across workflows:

```text
ORG|WORKFLOW|YYYY-MM-DD|MESSAGE-TYPE|TARGET
```

Unmarked or user-authored mail must never be modified.

## Validation

Run the VBS operational smoke suite:

```text
cscript.exe //NoLogo .\vbs\tests\RunAllSmokeTests.vbs
```

Run Python support validation:

```text
python -B -m unittest discover -s .\python\tests
python -B .\python\validate_scaffold.py --root .
python -B .\python\workflow_input_report.py --root . --format json --fail-if-invalid
python -B .\python\command_route_report.py --root . --format json --fail-if-invalid
python -B .\python\contract_fixture_report.py --root . --format json --fail-if-invalid
```

Python is optional for operators but useful for maintainers. It checks schemas,
fixtures, public examples, route contracts, dispatch-state parity, sanitisation,
and no-generated-artifact rules.

## Deployment Profiles

### Locked-Down Operator Workstation

Use this when the operator has Windows Script Host, Outlook, and Excel, but no
Python.

1. Put the scaffold folder in an approved local or shared location.
2. Open a command prompt in the scaffold root.
3. Run:

```text
cscript.exe //NoLogo .\Scaffold.vbs help
cscript.exe //NoLogo .\Scaffold.vbs validate
```

4. Run workflow routes with `--dry-run` and public-safe or local operational
   input paths.
5. Review key-value output and any `ACTION_REQUIRED` state before enabling any
   future live behaviour.

### Developer Validation Workstation

Use this when Python is available for support checks.

1. Run the VBS smoke suite.
2. Run Python unit tests.
3. Run scaffold validation.
4. Run command-route and workflow-input reports.
5. Fix only scaffold-local files.

Recommended command set:

```text
cscript.exe //NoLogo .\vbs\tests\RunAllSmokeTests.vbs
python -B -m unittest discover -s .\python\tests
python -B .\python\validate_scaffold.py --root .
python -B .\python\command_route_report.py --root . --format json --fail-if-invalid
```

### Scheduled Watchdog Sidecar

Use this only for report-only diagnostics.

1. Review the disabled example:

```text
examples\task_scheduler\watchdog_report_only_task.xml
```

2. Copy and adapt it outside the public scaffold.
3. Set a local environment variable such as `OFFICE_SCAFFOLD_ROOT`.
4. Keep the task disabled until the operator has validated paths locally.
5. Do not add repair, delete, send, replace, or live-inspection flags.

The report-only route remains:

```text
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
```

## Case Study 1: Group Meetings Dry-Run Migration

Goal: migrate a recurring meeting communication workflow away from embedded
Outlook COM code and into reusable scaffold contracts.

Implementation steps:

1. Prepare a rota-shaped CSV using the public example headers.
2. Keep recipient aliases generic and resolve them through a directory CSV if
   needed.
3. Run the root route in dry-run mode:

```text
cscript.exe //NoLogo .\Scaffold.vbs group-meetings --dry-run --input .\examples\workbooks\group_meetings_rota_example.csv --date 2026-05-29
```

4. Confirm the route reports `state=COMPOSED`, `dry_run=true`, and the expected
   message count.
5. If the route reports `ACTION_REQUIRED`, review the rota row, missing
   presenters, missing manager fields, or intentionally suppressed meeting
   state.
6. Add or update fixture-backed tests before changing planner behaviour.

What this proves:

- the workflow can build reminder and distribution messages without Outlook;
- workflow logic delegates composition to the shared core;
- managed keys and dispatch states are generated consistently;
- the operator can inspect output before any future live send path exists.

## Case Study 2: Induction Broadcast And Personal Messages

Goal: run a template-driven induction mailer that sends shared broadcast content
and personalized messages from one manifest shape.

Implementation steps:

1. Prepare `induction_recipients_example.csv` style recipient data.
2. Prepare an `induction_template_manifest_example.csv` style template
   manifest.
3. Put HTML templates and generic assets under a template folder.
4. Run:

```text
cscript.exe //NoLogo .\Scaffold.vbs induction --dry-run --recipients .\examples\workbooks\induction_recipients_example.csv --templates .\examples\workbooks\induction_template_manifest_example.csv
```

5. Confirm `active_recipient_count`, `planned_count`, and `composed_count`.
6. If `ACTION_REQUIRED` appears, inspect missing active emails, invalid
   template modes, invalid delivery actions, bad defer minutes, or missing
   enabled template files.

What this proves:

- one template manifest can express broadcast and personal messages;
- recipient resolution, rendering, assets, and delivery policy are shared core
  responsibilities;
- malformed operational inputs stop before message planning;
- dry-run output is suitable for operator review and automated tests.

## Case Study 3: Report-Only Watchdog Sidecar

Goal: add scheduled diagnostics for managed Outbox state without modifying a
mailbox.

Implementation steps:

1. Validate the scaffold locally:

```text
cscript.exe //NoLogo .\Scaffold.vbs validate
```

2. Run the watchdog route:

```text
cscript.exe //NoLogo .\Scaffold.vbs watchdog --report-only
```

3. Review counters such as `managed_count`, `unmanaged_count`,
   `replace_eligible_count`, and `repair_actions_taken`.
4. Keep `repair_actions_taken=0` for report-only operation.
5. Use fake-object tests and fixtures before any future repair behaviour is
   considered.

What this proves:

- the sidecar can report scaffold-managed state without repair;
- unmarked mail remains non-mutable;
- uncertain states become `ACTION_REQUIRED` or `FAILED_SAFE`;
- scheduled diagnostics can be introduced without destructive mailbox actions.

## Case Study 4: Adding A New Workflow

Goal: add another communication workflow without duplicating Outlook code.

Implementation pattern:

1. Add a workflow input adapter under `vbs/workflows/`.
2. Add a message-plan module that builds `MessageSpec` and `DeliveryPolicy`
   dictionaries.
3. Add a dry-run runner that calls `workflow_facade.vbs`.
4. Add a small VBS entrypoint under `vbs/entrypoints/`.
5. Route it through `Scaffold.vbs`.
6. Add public-safe examples under `examples/`.
7. Add fixtures under `vbs/tests/fixtures/` or `fixtures/`.
8. Add VBS smoke tests proving no Outlook creation, folder access, send, or
   delete calls live in workflow or orchestration modules.
9. Add Python validation only for schemas, fixtures, cross-runtime parity, or
   command-route checks.
10. Document the workflow and add it to validation registration.

Do not place low-level Office COM logic in the new workflow. If a reusable
primitive is missing, add it to the core layer and test it with fake objects or
dry-run mode first.

## Troubleshooting

`exit code 2`: check the command name or argument spelling.

`exit code 3`: an unsafe live/send/repair/delete/replace/inspect request was
blocked. Use dry-run or report-only mode.

`exit code 4`: a required input path was missing.

`ACTION_REQUIRED`: the scaffold stayed safe, but an operator must review the
input or workflow state.

`FAILED_SAFE`: the scaffold refused to continue or a safe validation/composition
step failed.

If Python validation reports private material, remove real names, addresses,
paths, links, logs, and institution-specific identifiers from public scaffold
files. Use `example.invalid`, `Example User`, `Example Group`, `examples/`, and
`fixtures/` for public-safe content.

## Important Files

- `Scaffold.vbs`: preferred root command router;
- `vbs/core/`: reusable VBS primitives;
- `vbs/workflows/`: workflow planners and runners;
- `vbs/entrypoints/`: command-specific VBS entrypoints;
- `vbs/tests/`: smoke tests and fixtures;
- `examples/`: public-safe workbook/template/task examples;
- `fixtures/`: shared schemas, state fixtures, quickstart routes, watchdog
  fixtures, and sanitisation config;
- `python/validate_scaffold.py`: scaffold validator;
- `python/command_route_report.py`: root route contract report;
- `python/workflow_input_report.py`: example input report;
- `docs/`: detailed architecture, migration, workflow, watchdog, validation,
  and release-readiness documentation.

## License

The scaffold-local license is provided in `LICENSE`.
