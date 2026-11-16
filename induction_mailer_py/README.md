# Induction Mailer (Python port)

This repository is a **Python OOP port** of the VBS/Outlook COM automation you shared (InductionMailer + utilities). It keeps the same high-level workflow and naming conventions, but exposes them as reusable Python classes.

## What you get

- `OutlookClient`: thin wrapper around Outlook COM (pywin32)
- `MailComposer`: creates HTML emails, attaches files, embeds inline images, and executes actions (`display`, `send`, `save`)
- `TemplateRepository`: recursively loads templates (`.html`, `.htm`, `.txt`), parses the embedded metadata header, and exposes templates for selection
- `NameResolver`: Outlook/Exchange/GAL resolution (ported heuristics from the VBS `MNS_*` modules)
- `InductionMailerApp`: interactive console flow that mirrors `InductionMailer.vbs`

## Installation (Windows)

This automation uses Outlook COM, so it is intended to run on **Windows with Outlook installed**.

```bash
# in your environment
pip install -e .

# Outlook/COM support
pip install -e ".[outlook]"
```

## Quick start

Run the InductionMailer console app:

```bash
induction-mailer              # interactive (display by default)
induction-mailer send         # compose and send
induction-mailer save         # compose and save .msg files
```

At startup, the induction mailer now asks whether you want to prepare induction
emails or add a new starter to the group meeting rota people sheet. The rota
option downloads `meeting_rota.xlsx`, adds the person to the `Emails` table with
`Presenter = 1`, `Lab Worker = 1`, and `Past member = 0`, then uploads the
workbook back to the same SharePoint location.

Name-resolution / GAL lookup diagnostics (ported from `gal_address_lookup.vbs`):

```bash
gal-lookup "John Smith" "email:someone@domain.com"
```

Directory probe diagnostics (ported from `haoran_qin_directory_probe.vbs`):

```bash
directory-probe
```

Group meeting manager (ported from `group_meetings_manager.vbs`):

```bash
group-meetings-manager --debug    # display emails instead of sending
group-meetings-manager            # send emails (requires Outlook)
```

## Template format

Templates are files in a folder (default `templates/`) with a **metadata block** at the top.
Two header formats are supported (matching the VBS):

### HTML comment header

```html
<!--
To: [EMAIL]
CC: someone@domain.com
Subject: Safety induction recap
stage: onboarding
code: NMR
attachments: ./files/info.pdf
-->
<!doctype html>
<html>...</html>
```

### Plaintext header

```text
@--
To: [EMAIL]
Subject: Hello
--
Body starts here...
```

Placeholders like `[NAME]`, `[SURNAME]`, `[FULL_NAME]`, `[EMAIL]`, `[TARGETS]` are replaced from context.

## Inline images

If your template is `MyTemplate.html`, you can place images under:

```
MyTemplate.files/
  logo.png
  banner.jpg
```

References to these images will be rewritten to `cid:logo.png`, and the images will be attached with matching content IDs.

## Notes on completeness

The zip you uploaded contains InductionMailer and a large set of shared utilities (including `mail_composer.vbs` and the full `MNS_*` name-resolution stack). The VBS code references an external `mail_utils_mailer_dynamiclinks.vbs` library that is **not present in the zip**; the Python port uses the more modern split modules you *did* provide (`mail_composer.vbs` + `mail_delivery_core.vbs`) as the canonical implementation.

If you also have `mail_utils_mailer_dynamiclinks.vbs` (or templates), you can drop them in and we can port those too.

## Repository structure

- `induction_mailer/outlook/` — Outlook COM wrapper + mail composition
- `induction_mailer/template/` — template parsing, placeholder handling, link maps, attachments
- `induction_mailer/name_resolution/` — name resolution heuristics + directory helpers
- `induction_mailer/cli.py` — InductionMailer console app
- `induction_mailer/scripts/` — diagnostic command equivalents
