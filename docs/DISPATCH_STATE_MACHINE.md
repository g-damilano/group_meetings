# Managed Dispatch State Machine

All dispatch-capable workflows should log and report using these states.

| State | Meaning |
| --- | --- |
| `NOT_REQUESTED` | Workflow decided no email should be prepared. |
| `COMPOSED` | A message specification or Outlook item has been composed. |
| `DISPLAYED` | Message was opened for user review. |
| `DRAFTED` | Message was saved as a draft or `.msg`. |
| `SEND_REQUESTED` | `.Send` was requested from Outlook. |
| `QUEUED_OUTBOX` | Matching managed email exists in Outbox and counts as an active send attempt. |
| `SENT_CONFIRMED` | Matching managed email exists in Sent Items. |
| `DUPLICATE_SUPPRESSED` | Send was skipped because Sent Items or Outbox already contains the managed email. |
| `WAITING_DEFERRED` | Outbox email has a future deferred delivery time. |
| `BLOCKED_OFFLINE` | Outlook appears offline, disconnected, or Work Offline is active. |
| `BLOCKED_AUTH` | Outlook transport is blocked and O365 authentication is suspected; direct auth detection may be unavailable. |
| `STALE_MANAGED_OUTBOX` | Managed email is stale, overdue, and Outlook appears connected. |
| `REPLACED_MANAGED_OUTBOX` | A stale managed Outbox item was safely replaced. |
| `ACTION_REQUIRED` | Human intervention is required. |
| `FAILED_SAFE` | Operation stopped without unsafe modification. |

## State Rules

- Sent Items wins over all other states.
- Outbox presence counts as an active send attempt.
- User-authored or unmarked Outbox mail must never be modified.
- Replacement is allowed only for stale, overdue, scaffold-managed mail while Outlook appears connected.
- Non-urgent sends do not require immediate post-submit verification; the next workflow or watchdog run will inspect state.
- Urgent sends may opt into immediate verification.

