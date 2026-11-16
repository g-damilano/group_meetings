# Canonical Contracts

These contracts define the shared language used by VBS and Python.

## MessageSpec

Represents one intended email before Outlook-specific creation.

Required fields:

- `message_key`: deterministic managed identity.
- `message_type`: stable workflow-specific type.
- `to`: semicolon-delimited recipients or a recipient list.
- `subject`: final rendered subject.
- `body_text` or `body_html`: final rendered body.
- `delivery_policy`: display, draft, send, save, or deferred send.

Optional fields:

- `cc`
- `bcc`
- `attachments`
- `inline_image_folder`
- `deferred_delivery_time`
- `urgent_verify_after_submit`
- `metadata`

## RecipientSpec

Represents an intended person, mailbox, list, or group.

Fields:

- `raw_value`
- `display_name`
- `smtp_address`
- `resolution_status`
- `directory_source`
- `fallback_used`

## DeliveryPolicy

Controls how a `MessageSpec` is handled.

Allowed actions:

- `display`
- `draft`
- `send`
- `save`

Optional behavior:

- `defer_minutes`
- `urgent_verify_after_submit`
- `managed_dispatch_required`

## ManagedEmailKey

Deterministic identity for one logical email.

Canonical format:

```text
ORG|WORKFLOW|DATE|MESSAGE_TYPE|TARGET
```

For legacy compatibility, workflow implementations may use shorter existing formats during migration, but must document them and preserve stability across reruns.

## DispatchResult

Result returned by managed dispatch.

Fields:

- `state`
- `reason`
- `managed_key`
- `folder`
- `outlook_entry_id`
- `action_required`
- `created_new_mail`
- `replaced_existing_mail`
- `safe_to_retry`

## WorkflowContext

Runtime context passed through a workflow.

Fields:

- `workflow_name`
- `run_id`
- `source_workbook`
- `template_root`
- `log_path`
- `dry_run`
- `operator`
- `settings`

## HealthReport

Diagnostic result from Outlook/O365 checks.

Fields:

- `offline`
- `exchange_connection_mode`
- `exchange_connection_mode_name`
- `blocked_or_auth_suspected`
- `outbox_count`
- `managed_outbox_count`
- `unmanaged_outbox_count`
- `action_required_messages`

