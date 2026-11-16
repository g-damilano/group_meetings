# Group Meetings Workbook Conventions

This scaffold treats rota/workbook data as user-facing configuration. The canonical contract is the
normalised `meetingRecord` object consumed by `group_meetings_message_plan.vbs`.

The VBS adapter currently supports anonymised CSV fixtures that mirror workbook columns. Excel COM
loading can be added later behind the same contract, but workflow planning must remain independent of
Excel and Outlook COM.

## Required Columns

- `meeting_date`
- presenter channel columns. The compatibility CSV uses `literature_presenter`
  and `results_presenter` as the default channel columns.
- `alternate_text`
- `room`
- `lab_manager`

Presenter channels are the workflow concept; the exact workbook headings are
adapter concerns. Implementations may map other local headings onto presenter
channels, as long as the normalized meeting record preserves channel identity
and the message plan can still build deterministic presenter reminders.

## Optional Columns

- `meeting_lead`
- `group_recipient`
- `manager_recipient`
- `sender_name`
- `include_same_day_reminder`
- `send_presenter_reminders`
- `send_meeting_lead_reminder`
- `send_lab_manager_reminder`

## Compatibility Rules

- A row with `alternate_text` containing `no meeting`, `no-meeting`, or `college closure` becomes a
  no-meeting notice and suppresses reminders.
- A row with no configured presenter channels and no alternate text becomes
  `ACTION_REQUIRED` and produces a manager alert plan.
- If `meeting_lead` is blank, the scaffold may derive it from the previous rota row's `lab_manager`,
  matching the legacy behavioural shape without copying private defaults.
- Same-day reminders are not inferred by the adapter. Orchestration must set
  `include_same_day_reminder` after checking the current date and time.
- The adapter must not send mail, access Outlook folders, or repair Outbox items.

## Public Fixture Rules

Fixtures must use generic names, generic rooms, and `example.invalid` addresses where email addresses
are needed. They must not include private workbook paths, internal links, institution names, production
logs, or real recipients.
