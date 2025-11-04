# Brandt-Talbot Induction Agent - AGENT.md

## Overview
This repository contains the specification and assets to automate the Brandt-Talbot Group induction email pipeline.
The agent coordinates delivery of induction emails (ID card requests, training, induction scheduling, access requests, recap) by
calling the helper script `mail_utils_utf8_plus_dynamiclinks_ext.vbs` that handles Outlook-safe sending and dynamic link updates.

**Important:** This package contains templates, blueprints and developer notes. Per instructions, this package *does not* include a stub for InductionMailer.vbs — that script will be developed by the AI leveraging `mail_utils_utf8_plus_dynamiclinks_ext.vbs` and the assets here.

## Contents
- `blueprint.json` — machine-readable blueprint driving the agent's decisions
- `AGENT.md` — this developer/operator guide
- `InductionMailer_developer_notes.md` — extra notes, expectations, integration points
- `LICENSE` — MIT license
- `.gitignore`
- `Induction/` — folder with the current email templates (HTML/HTM) and placeholders
- `package_manifest.json` — manifest for the package contents and checkpoints

## Operator summary
1. Get HR notification and prepare inductee email(s).
2. Run the operator UI or the agent, providing one or more email addresses.
3. Choose one option (0, 2, 3, 4, 5) to execute. The agent will:
   - Resolve names via Outlook address book.
   - Compose/supply appropriate greeting ("Hi [Name],") for individual sends.
   - Use `mail_utils_utf8_plus_dynamiclinks_ext.vbs` for sending and hyperlink updating.
   - Attach forms for EMAIL 6 from `Induction/` where available.
   - After EMAIL 6 is sent, automatically send a Reception Notification to `chemistryreception@imperial.ac.uk`.

## Checkpoints & Deadlocks
- EMAIL 1 (ID) must be actioned by inductee or will block access later.
- EMAIL 2 (Fire Safety) is required prior to processing access form submission.
- EMAIL 6 requires signed forms to be returned — the pipeline's main deadlock moves here.
- The agent must log sent status and return error codes if sends fail.

## Contact & Supervisor Defaults
- IDC Card Information: id.card@imperial.ac.uk
- Supervisor (CC): agi@imperial.ac.uk
- Reception Notification: chemistryreception@imperial.ac.uk

