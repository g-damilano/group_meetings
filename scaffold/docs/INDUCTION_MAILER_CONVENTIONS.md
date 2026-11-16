# Induction Mailer Conventions

Stage 3 uses anonymised CSV fixtures to express workbook and template
conventions without copying private operational data.

## Recipient CSV

Required columns:

- `alias`
- `email`

Optional columns:

- `first_name`
- `last_name`
- `full_name`
- `display_name`
- `start_date`
- `cohort`
- `active`

`active` accepts `true`, `yes`, `1`, `on`, `false`, `no`, `0`, and `off`.
Inactive recipients are loaded but skipped by the planner.

`first_name` and `last_name` are the preferred structured source for
personalized messages. If one or both are missing, the adapter may fill missing
`FIRST_NAME`, `LAST_NAME`, `SURNAME`, and `FULL_NAME` values from `full_name`,
`display_name`, or an explicit `Display Name <smtp@example.invalid>` email
string. This display-name parsing is best-effort only and never overrides
structured recipient columns.

Active recipient rows must include an `email` value. Missing email values are
reported as `ACTIVE_RECIPIENT_MISSING_EMAIL` and the dry-run stops with
`ACTION_REQUIRED` before planning messages. Inactive support rows may still be
loaded for directory resolution when they include an email address.

## Template Manifest CSV

Required columns:

- `template_id`
- `mode`
- `template_path`

Optional columns:

- `target_key`
- `display_name`
- `inline_image_folder`
- `attachment_base_folder`
- `attachments`
- `cc`
- `bcc`
- `delivery_action`
- `defer_minutes`
- `urgent_verify_after_submit`
- `enabled`

`mode` is `broadcast` or `personal`. A selected mode of `all` includes both.
Enabled rows with any other mode are reported as `TEMPLATE_MODE_INVALID`.

`template_path` may be absolute or relative to the manifest folder. Template
files may contain metadata comments at the top:

```text
<!-- Subject: Welcome [NAME] -->
```

The core template engine parses metadata and renders `[TOKEN]` or `{{TOKEN}}`
placeholders from the workflow context.

`attachments` is a semicolon- or comma-separated list resolved relative to
`attachment_base_folder`. Inline images are resolved by the core mail asset
helpers from `inline_image_folder`, or from the template's sibling `.files`
folder when the manifest field is blank.

`cc` and `bcc` use the same recipient resolver as `to`. Support recipients may
be listed as inactive recipient rows so they are available for directory
resolution without being included in broadcast `to` lists. Resolver precedence
is direct SMTP or `Display Name <smtp@example.invalid>` text first, then exact
directory keys, then a normalized display-name fallback only when it has a
single unambiguous match.

`delivery_action`, `defer_minutes`, and `urgent_verify_after_submit` are
converted into a core `DeliveryPolicy`. The scaffold dry-run runner still
blocks live send requests; these fields are carried only as contract evidence
until an operator explicitly promotes a guarded live path.

For enabled template rows, `template_id` and `template_path` are required,
`delivery_action` must be blank, `display`, `draft`, or `send`, and
`defer_minutes` must be blank or a non-negative integer. Invalid enabled rows
stop the dry-run with `ACTION_REQUIRED` before any messages are planned.

Enabled templates with missing template files are reported as
`TEMPLATE_MISSING` skips and produce `ACTION_REQUIRED` when no other messages
can be planned. The scaffold must not compose empty fallback mail from a missing
template.

## Recipient Placeholder Keys

The induction planner provides:

- `ALIAS`
- `EMAIL`
- `NAME`
- `SURNAME`
- `FIRST_NAME`
- `LAST_NAME`
- `FULL_NAME`
- `START_DATE`
- `COHORT`
- `RECIPIENT_COUNT`

Templates should use generic placeholder names and must not contain private
names, addresses, internal links, or production-only identifiers.
