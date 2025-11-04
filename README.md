# Induction Mailer Workflow

The induction mailer automates creating Outlook drafts from HTML templates and applies the
shared dynamic-link substitutions maintained in `links_map.txt`. It reuses the
`mail_utils_mailer_dynamiclinks.vbs` helper so the same delivery logic is shared across the
team's scripts.

## Repository layout

- `InductionMailer.vbs` — interactive script that discovers HTML templates under the
  `Induction/` folder, prompts for a selection, and prepares the message metadata and body.
- `Induction/` — repository of HTML templates and their associated assets. Each template may
  include an optional HTML comment block at the top to specify recipients, subject, and
  attachments.
- `mail_utils_utf8_plus_dynamiclinks_ext/` — source of the shared mail utility and the example
  `links_map.txt` file that maps placeholder tokens to real URLs.

## Template format

Templates are plain HTML files (`.htm`/`.html`). To keep metadata with the template while
avoiding user-visible content, add an HTML comment at the very beginning:

```html
<!--
To: person@example.com
Cc: another@example.com
Subject: Welcome to the group
Attachments: intro.pdf; .\checklist.xlsx
-->
<html>...
```

- Omit any metadata line to leave the field blank.
- Attachments can be absolute paths or relative to the template's folder. Separate multiple
  paths with commas or semicolons.
- The comment must end before the rendered HTML starts.

## Dynamic link substitution

Hyperlink placeholders use double-brace tokens (for example `{{TEAMS_LINK}}`). When the mail
utility runs, it reads `links_map.txt`, substitutes each token in the HTML body, and then wraps
it with a UTF-8 `<meta>` tag so the mail renders correctly in Outlook.

To keep the substitutions up to date, edit `links_map.txt` (either in the script directory or in
`mail_utils_utf8_plus_dynamiclinks_ext/`). The mailer copies the file from the helper directory
if it is missing locally.

## Running the mailer

Run the script on Windows with Outlook installed:

```powershell
cscript //nologo InductionMailer.vbs [display|send|save]
```

- Choose `display` (default) to review the draft.
- Use `send` to dispatch immediately, or `save` to store the message in Drafts.
- The script logs activity to `InductionMailer.log` in the repository root.

If the script cannot find the mail utility or templates, it will show a message box and write the
error to the log file.
