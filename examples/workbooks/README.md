# Example Workbook Convention

The public scaffold uses Excel as the operational configuration surface.

Minimum workbook sheets for a generic recurring communication workflow:

- `Settings`: workflow-level defaults such as organization code, workflow name, log path, and default delivery action.
- `People`: reusable recipients, display names, roles, aliases, and fallback email addresses.
- `Messages`: message definitions, message types, subjects, delivery policy, and template paths.
- `Schedule`: dated or recurring task rows that choose people, messages, and workflow-specific fields.

The current group-meetings workbook can be migrated into this convention by treating presenters, lab managers, meeting leads, rooms, and no-meeting events as workflow-specific `Schedule` columns.

