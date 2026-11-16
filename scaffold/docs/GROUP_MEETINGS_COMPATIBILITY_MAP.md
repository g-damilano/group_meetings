# Group Meetings Compatibility Map

This document maps the frozen production group-meetings behaviour into anonymised scaffold concepts.
The legacy script remains read-only. Scaffold code must not copy private recipients, branded subject
prefixes, private schedule links, workbook paths, logs, or institution-specific content.

## Legacy Behaviour Observed Read-Only

The frozen script contains these workflow decisions:

- Build one regular group distribution for a configured meeting date.
- Build a no-meeting distribution when the rota alternate text indicates closure or no meeting.
- Suppress reminders when the meeting is a no-meeting event.
- Suppress the regular group distribution and alert the manager when no presenters or alternate are set.
- Build presenter reminders for the configured presenter channels. The
  compatibility fixture currently names the channels literature and results.
- Build a meeting-lead reminder.
- Build a lab-manager reminder.
- Build a same-day reminder only when orchestration timing says it is the meeting day before the meeting.
- Use managed email keys for regular and reminder email types.
- Treat Sent Items and Outbox as duplicate-prevention state through managed dispatch.

## Scaffold Message Plan

The scaffold adapter in `vbs/workflows/group_meetings_message_plan.vbs` uses anonymised message types:

- `REGULAR`
- `LITERATURE-REMINDER`
- `RESULTS-REMINDER`
- `MEETING-LEAD-REMINDER`
- `LAB-MANAGER-REMINDER`
- `SAME-DAY-REMINDER`
- `MEETING-NOT-SET-ALERT`

The regular group distribution remains the logical key type for a no-meeting notice. That means a
meeting date has only one active regular distribution, whether the body is a meeting notice or a
no-meeting notice.

## Safety Boundary

The message-plan adapter is workflow-layer logic. It may decide which `MessageSpec` requests are needed,
but it must not:

- create Outlook objects;
- access Outlook folders;
- send mail;
- delete mail;
- inspect or repair Outbox items.

Those operations remain in the core layer and are exercised by dry-run tests here.

## Public Scaffold Differences

The scaffold intentionally uses generic subjects, generic body text, generic recipients, and
`example.invalid` addresses in tests. Public examples must remain anonymised while preserving the
behavioural shape of the workflow.
