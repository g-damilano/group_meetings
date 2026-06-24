# Induction Mailer Console Flow

Documentation for `InductionMailer.vbs`, covering how the console workflow discovers templates, prompts for recipients, and dispatches personalised or broadcast messages.

| Area | Details |
| --- | --- |
| Purpose | Console-driven induction sender that stages or individually dispatches Outlook messages based on template metadata. |
| Entry point | `InductionMailer.vbs` (launched with `cscript //nologo InductionMailer.vbs [display|send|save]`). |
| Dependencies | Relies on Outlook, `mail_utils_mailer_dynamiclinks.vbs`, `links_map.txt`, and the helper modules in `utilities/`. |

## Bootstrap and Dependencies
- Loads `utilities/bootstrap_entry.vbs` via `BootstrapIncludeEntry`, then imports core helpers (`common_console`, `common_outlook`, `mail_collections`, `mail_template_utils`, `mail_delivery_utils`, `mail_input_validation`, `MNS_loader`).
- Attempts to load `mail_utils_mailer_dynamiclinks.vbs` from the script directory, the external `mail_utils_utf8_plus_dynamiclinks_ext` folder, or the templates directory before continuing.
- Pulls dynamic link substitutions from `links_map.txt`, merging them into the placeholder context alongside template metadata.

## Configuration and Constants
- Template discovery starts from `TEMPLATE_DIR_NAME = "templates"` next to the script.
- Logs are written to `LOG_FILE_NAME = "InductionMailer.log"` in the repository root.
- Default action is `display`; other actions are `send` and `save`.
- Delivery modes are constrained to `broadcast` and `personal`.
- Template headers are detected with either HTML (`<!-- ... -->`) or legacy `@-- ... --` markers.

## Environment Setup
- `InitializeEnvironment` creates the `FileSystemObject`, resolves the script and template folders, initialises a case-insensitive placeholder dictionary, and records startup information to the log.
- Name resolution helpers are initialised through `LoadNameResolutionModules` so Global/Offline Address List lookups are available before prompting.

## Template Discovery and Selection
- `LoadTemplates` recursively scans `templates/` and builds a keyed collection using metadata from each header block (stage, delivery mode, code ordering, attachments, etc.).
- The console offers a mode prompt: stage mode selects every template for a given stage; email mode lets the user pick one or more templates by index.
- Template ordering within a stage respects the `Code` field (for example `E1`–`E9`).

## Recipient Resolution and Review
- `PromptRecipientResolutionPackage` requests the target address list and validates entries with shared input validators.
- GAL/Outlook lookups populate `[NAME]`, `[SURNAME]`, and `[FULL_NAME]` for personalised runs; failures fall back to prompting so empty addresses never proceed silently.
- The script aborts early if no valid recipient entries remain after validation.

## Per-Template Processing
- `ProcessTemplate` orchestrates end-to-end handling for each selected template:
  - Parses the header block (`LoadTemplateMetadata`) and resolves placeholders across address fields, subject, body, and attachments.
  - Requests missing metadata (including delivery mode) interactively when the template leaves them blank.
  - For personal deliveries, loops over each recipient to inject per-person placeholder values and renders one email per target; broadcast deliveries combine all addresses into a single message.
  - Resolves attachments relative to the template location, including Outlook `.files/` directories for inline images.
  - Hands the fully-prepared parameters to the mailer library to display, send, or save according to the chosen action.

## Logging and Diagnostics
- `LogEvent` writes timestamped entries to `InductionMailer.log`, including startup details, template folder paths, and any missing module warnings.
- Placeholder resolution issues, absent templates, and library load failures are surfaced to both the console and the log to simplify troubleshooting.
- The name-resolution utilities can emit additional diagnostics when directory lookups encounter ambiguous or missing addresses.
