# Group Meetings Manager Flow

Reference for `group_meetings_manager.vbs`, outlining how meeting rota data is synced, turned into Outlook alerts, and safeguarded against duplicate sends.

| Area | Details |
| --- | --- |
| Purpose | Automates weekly rota distribution, reminders, and coordinator notifications based on the shared Excel workbook. |
| Entry point | `group_meetings_manager.vbs` (run with `cscript //nologo group_meetings_manager.vbs`). |
| Dependencies | Outlook, rclone for SharePoint sync, and the local `meeting_rota.xlsx` copy defined by `EXCEL_PATH`. |

## Globals and Configuration
- Script-level objects capture the shell, filesystem, script folder, and log path once (`InitializeFileSystem`).
- Core constants include recipient defaults (`EMAIL_TO`, `MANAGER`, `BUILDING_ADMIN_EMAIL`, `SUPERVISOR_EMAIL`), workbook locations (`EXCEL_PATH`, `SHEET_ROTA`, `SHEET_EMAILS`), and behaviour toggles (`DEBUG_MODE`, `EXTENDED_LOGGING`, `DISPLAY_DEBUG_POPUPS`).
- Default previous manager fallback (`DEFAULT_PREVIOUS_LAB_MANAGER`) and the weekly log prefix (`LOG_VARIANT`) keep message content and log context consistent.

## Logging and Rotation
- `LogEvent` appends timestamped entries to `group_meetings_manager.log` alongside a variant tag.
- `ShouldResetLogForNewWeek` inspects the log's last modified week and resets the file when a new week is detected; `InitializeLogging` announces whether a continuation or fresh log is in use.
- `LogVerbose` gates chatty diagnostic messages, while `LogError` captures VBScript error metadata before clearing the error state.

## Workbook Sync and Loading
- `SyncMeetingWorkbook` invokes rclone to copy `meeting_rota.xlsx` from SharePoint into the local working folder, logging the command and exit code and warning when non-zero results occur.
- `CleanupSyncedWorkbook` removes the synced workbook after processing to keep the local folder tidy.
- Workbook data is accessed via `LoadWorkbook` helpers (in the lower section of the script) that use Excel automation to read the rota (`SHEET_ROTA`) and email aliases (`SHEET_EMAILS`).

## Outlook Resolution (`GetExactEmail`)
- `GetExactEmail` constructs an Outlook recipient from a search string and resolves it to an Exchange user; if resolution fails or the SMTP address is absent, an empty string is returned.
- Uses a fresh Outlook Application/Namespace to avoid side effects and surfaces verbose logging for traceability when extended logging is enabled.

## Duplicate-Send Safeguards
- `CheckAndLogExecution` inspects a registry key under `HKCU\Software\GroupMeetingsManager` to detect whether the current week has already been processed; when found, it logs the prior run and aborts the main flow.
- On successful completion, the registry value is written with the current week number so re-runs can be detected before sending again.

## Email Composition Helpers
- `BuildEmailBody`, `BuildSubject`, and `FormatSessionDetails` assemble the rota content, presenter names, links, and timing information into human-readable text.
- `FillOutlookMail` centralises Outlook mail item creation: sets `To`/`Cc`, subject, body, importance, and handles the debug/dry-run shortcut when `DEBUG_MODE` is true.

## Rota-Based Alert and Reminder Orchestration
- `SendAlertEmails` drives the weekly distribution by loading both the rota row for the upcoming meeting and the `Emails` sheet for named addresses; it resolves presenter emails with `GetExactEmail` before sending.
- `SendReminderEmails` performs a similar lookup for same-day reminders and respects the debug flag to avoid accidental sends during testing.
- Both flows log detailed decisions (missing presenter emails, empty rota rows, send results) to aid troubleshooting.

## Coordinator and Main Flow
- `NotifyCoordinator` targets the lab manager and supervisor with a concise status email summarising the week's rota row and whether participant addresses were resolved.
- `Main` wires everything together: initialises filesystem/logging, syncs the workbook, starts Outlook, executes duplicate-run checks, and then calls the alert/reminder/coordinator routines in order before cleaning up the synced file.
