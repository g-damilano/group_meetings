# Dispatch State Machine

All dispatch-capable workflows log and report with these states.

| State | Meaning |
| --- | --- |
| `NOT_REQUESTED` | Workflow decided no email should be prepared. |
| `COMPOSED` | A message object was composed. |
| `DISPLAYED` | Message was displayed for user review. |
| `DRAFTED` | Message was saved as draft. |
| `SEND_REQUESTED` | Message was submitted to Outlook for sending. |
| `QUEUED_OUTBOX` | Message is in Outbox and counts as an active send attempt. |
| `SENT_CONFIRMED` | Message exists in Sent Items. |
| `DUPLICATE_SUPPRESSED` | Send was skipped because Sent Items or Outbox already had the logical email. |
| `WAITING_DEFERRED` | Deferred delivery time has not arrived. |
| `BLOCKED_OFFLINE` | Outlook appears offline or disconnected. |
| `BLOCKED_AUTH` | O365 authentication or submission blockage is suspected. |
| `STALE_MANAGED_OUTBOX` | Managed Outbox item is stale and may be eligible for guarded repair. |
| `REPLACED_MANAGED_OUTBOX` | A stale managed Outbox item was replaced under strict guardrails. |
| `ACTION_REQUIRED` | Human action is required; no destructive repair should happen. |
| `FAILED_SAFE` | The scaffold stopped safely because certainty was insufficient. |

## State Rules

- Sent Items wins over all other states.
- Outbox presence counts as an active send attempt.
- User-authored or unmarked Outbox mail must never be modified.
- Replacement is allowed only for stale, overdue, scaffold-managed mail while Outlook appears connected.
- Non-urgent sends do not require immediate post-submit verification.
- Urgent sends may opt into immediate verification.
- Uncertainty resolves to `ACTION_REQUIRED` or `FAILED_SAFE`, not destructive repair.

