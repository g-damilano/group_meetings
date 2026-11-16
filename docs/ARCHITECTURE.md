# Layered Office Communication Automation Architecture

This repository is being reorganized around three layers.

## 1. Core Office Automation Layer

Reusable primitives for the repetitive Office/Outlook work that every workflow needs:

- Outlook and Excel session acquisition.
- Message composition.
- Managed email identity and dispatch.
- Recipient and address-book resolution.
- Template metadata and placeholder rendering.
- Outlook transport and Outbox diagnostics.
- Logging and action-required reporting.

The operational deployment target is VBS-first because corporate Windows/O365 environments often allow built-in scripting while restricting Python, pip packages, unsigned binaries, or custom dependencies.

Python remains first-class for development support, validation, tests, fixture generation, optional CLI use, and future non-COM backends.

## 2. Workflow Pipeline Layer

Workflow modules decide what work needs to happen and build message specifications for the core layer.

Reference workflows:

- Group meetings: rota-driven recurring communications, reminders, alerts, no-meeting notices, and lab-management notices.
- Induction mailer: template-driven broadcast and personal communications.

Workflow code should not own low-level Outlook COM behavior. It should build contract objects and call core primitives.

## 3. UI And Orchestration Layer

Small entry scripts gather user intent, choose a workflow, load configuration, and call the workflow layer.

VBS entrypoints are the operational default. Python entrypoints are optional for less restricted environments.

## Canonical Truth Rule

The architecture is canonical in contracts, schemas, naming conventions, workbook conventions, template conventions, managed-key formats, documented dispatch states, and cross-runtime fixtures.

It must not live exclusively in either Python or VBS.

