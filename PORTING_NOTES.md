# Porting notes (VBS -> Python)

## 1) What was ported

The provided zip contained:

- `InductionMailer.vbs`
- utilities under `utilities/`
- name-resolution modules under `utilities/name_resolution/`

All of those modules have corresponding Python implementations in the `induction_mailer/` package (see `TRACE_VBS.md` for a per-file map).

## 2) Missing dependency in the zip

`InductionMailer.vbs` references a mailer library (`mail_utils_mailer_dynamiclinks.vbs`, plus an optional `.ext` variant) that **was not included** in the uploaded zip. In the Python port, its functionality is covered by:

- `induction_mailer/outlook/composer.py`
- `induction_mailer/template/link_map.py`
- `induction_mailer/template/assets.py`

So you should not need that missing VBS file to run the Python implementation.

## 3) Outlook COM (pywin32)

- The Outlook automation parts require Windows and Outlook installed.
- Install the extra deps with:
  - `pip install .[outlook]`

## 4) Behavioural differences vs the VBS scripts

- The Python CLI focuses on the same end-user behaviour (template selection, placeholder prompting, broadcast/personal delivery), but it does not replicate the VBS bootstrap mechanism; it uses normal Python imports.
- Name resolution is implemented with the same strategy (Recipient.Resolve first, then directory list searching). Exact matches may differ slightly depending on Exchange/tenant configuration.

