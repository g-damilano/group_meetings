# Layered Office Communication Automation Architecture

The scaffold has three layers.

## Core Office Automation Layer

The core layer owns reusable low-level Office and Outlook work:

- Outlook and Excel session acquisition;
- account, namespace, and folder access;
- message composition;
- managed email identity;
- managed dispatch;
- recipient and address-book resolution;
- template metadata and placeholder rendering;
- attachments and inline images;
- deferred delivery;
- Outlook transport diagnostics;
- Outbox diagnostics;
- logging and action-required reporting.

The core must not contain workflow-specific group-meeting or induction policy.

### Stage 1 Core Module Boundary

The VBS core is split into narrow modules so workflows can call reusable primitives without owning
Outlook COM details:

- `session.vbs`: Outlook, Excel, MAPI namespace, and transport-state acquisition;
- `folder_access.vbs`: safe account listing plus Sent Items and Outbox folder access;
- `mail_composer.vbs`: message creation without workflow policy;
- `managed_dispatch.vbs`: managed keys, duplicate suppression, and queued-mail decisions;
- `health_diagnostics.vbs` and `outbox_inspector.vbs`: report-only diagnostics and watchdog-safe
  inspection primitives.
- `workflow_facade.vbs`: workflow-facing helpers that build contracts and prepare dry-run dispatch
  results without exposing workflows to low-level COM details.

Non-urgent dispatch does not require immediate post-submit verification. Follow-up workflow runs and
the watchdog inspect Sent Items and Outbox state using the managed email key.

## Workflow Pipeline Layer

Workflow modules decide what work needs to happen and build contract objects for the core.

Reference workflows:

- Group meetings: rota-driven recurring communications, reminders, alerts, no-meeting notices,
  same-day notices, and lab-management notices.
- Induction mailer: template-driven broadcast and personal communications.

Workflow code must not own low-level Outlook COM behaviour.

### Workflow Adapter Boundary

Workflow modules prove that both reference workflows can build `MessageSpec` and `DeliveryPolicy`
contracts through the core facade. Stage 2 moves group-meeting planning onto fixture-backed adapters;
Stage 3 moves the induction mailer onto template-manifest and recipient-manifest adapters while keeping
Outlook COM behaviour in the core layer.

- `vbs/workflows/group_meetings_workflow.vbs`: recurring meeting communication request helpers.
- `vbs/workflows/induction_mailer_workflow.vbs`: induction/broadcast message request helpers.
- `vbs/workflows/induction_mailer_input_adapter.vbs`: Stage 3 induction template and recipient
  manifest loader.
- `vbs/workflows/induction_mailer_message_plan.vbs`: Stage 3 induction broadcast/personal
  `MessageSpec` and `DeliveryPolicy` planner.
- `vbs/workflows/induction_mailer_runner.vbs`: Stage 3 dry-run runner that delegates composition to
  the core workflow facade.
- `vbs/workflows/workflow_helpers.vbs`: shared request-shaping helpers for workflow modules.

These modules must not call Outlook COM directly. Folder access, message creation, dispatch, diagnostics,
and watchdog decisions remain in the core layer.

## UI And Orchestration Layer

Entry scripts gather user intent, choose a workflow, load configuration, and call the workflow layer.
VBS entrypoints are the operational default. Python entrypoints are optional.

The orchestration layer must not contain Outlook COM logic.

## VBS-First Deployment

The public scaffold is VBS-first because the target environment is corporate Windows/O365, where users
may not be able to install Python, pip packages, unsigned executables, or custom dependencies.

Python remains in the repository for development, validation, fixtures, documentation, optional CLI use,
and future non-COM backends.
