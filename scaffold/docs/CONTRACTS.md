# Canonical Contracts

The canonical architecture lives in contracts, schemas, naming conventions, workbook conventions,
template conventions, managed-key formats, dispatch states, cross-runtime fixtures, and documented
behaviour. It must not live exclusively in VBS or Python.

## Public Concepts

- `MessageSpec`: desired message content, recipients, attachments, and metadata.
- `RecipientSpec`: a raw recipient plus resolution metadata.
- `DeliveryPolicy`: display, draft, send, deferred delivery, and urgent-verification policy.
- `ManagedEmailKey`: deterministic logical identity for a scaffold-managed message.
- `DispatchResult`: state, reason, and safety metadata returned by managed dispatch.
- `WorkflowContext`: workflow run metadata and settings.
- `HealthReport`: Outlook/O365 health and action-required diagnostics.

These concepts are represented in three places:

- VBS constructors in `vbs/core/contracts.vbs`, used by the operational path.
- JSON schema definitions in `contracts/workflow_schema.json`, used by
  development validation and examples.
- Cross-runtime fixture data in
  `fixtures/contracts/stage4_shared_contract_fixture.json`, which shows group
  meetings and induction mailer messages on the same contract surface.

## Workflow Facade Rule

Workflow modules should build mail through the VBS workflow facade rather than calling low-level COM
helpers directly. The facade builds `MessageSpec`, `DeliveryPolicy`, composer context, and dry-run
`DispatchResult` contracts while keeping Outlook folder/session work in the core layer.

`WorkflowContext` defaults to `dry_run = True`. Operational entrypoints may explicitly set it to
`False`, but tests, validation, and public examples should remain dry-run unless a user deliberately
opts into live Outlook actions.

## Recipient Resolution Rule

Recipient resolution is local and deterministic by default. The resolver first
accepts direct SMTP or `Display Name <smtp@example.invalid>` one-off recipient
text, then checks caller-supplied directory keys exactly. Only after those
structured paths fail may it try a normalized display-name fallback against the
provided directory.

Display names are convenience text, not authoritative structured identity. If a
normalized display-name fallback is ambiguous, the recipient remains unresolved
and the workflow can surface `ACTION_REQUIRED` rather than guessing.

Recipient specs may carry best-effort personalization name parts:

- `first_name`
- `last_name`
- `surname`
- `full_name`
- `name_parse_status`
- `name_parse_source`
- `name_parse_order`

These fields follow a conservative ontology. Structured workflow input wins
first. A direct `Given Family <smtp@example.invalid>` display string may fill
missing name parts as `given-family`; a `Family, Given <smtp@example.invalid>`
display string may fill missing name parts as `family-given`; a single display
token is only `partial`. Display-derived values are personalization hints, not
identity proof, and they must not override structured CSV/workbook fields.

Live Outlook, GAL, or O365 directory lookup belongs behind a separate guarded
adapter and is not required for the dry-run scaffold path.

## Managed Email Key Format

Default key format:

```text
ORG|WORKFLOW|YYYY-MM-DD|MESSAGE-TYPE|TARGET
```

The key must be deterministic, stable across reruns, unique per logical email, and independent of
Outlook `EntryID` values.

## Marker Rules

Managed messages should be stamped with both:

- an Outlook user property named `OfficeScaffoldManagedEmailKey`;
- a hidden HTML marker:

```html
<!-- OfficeScaffoldManagedEmailKey: ORG|WORKFLOW|YYYY-MM-DD|MESSAGE-TYPE|TARGET -->
```

The visible email body must not be degraded.
