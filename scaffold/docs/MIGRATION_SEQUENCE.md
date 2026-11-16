# Migration Sequence

## Stage 0 - Freeze Production Legacy

Preserve current daily-use scripts as the operational baseline. Do not modify them.

## Stage 1 - Extract VBS Core Primitives

Move Outlook session, composer, dispatch, diagnostics, logging, name-resolution, and template primitives
into reusable VBS modules inside `./scaffold/`.

## Stage 2 - Move Group Meetings Onto VBS Core

Keep behaviour stable while replacing embedded COM routines with core calls.

## Stage 3 - Move Induction Mailer Onto VBS Core

Reuse the same composer, recipient resolver, template engine, and managed dispatch.

Stage 3 scaffold-side migration includes the induction compatibility map, template and recipient CSV
conventions, broadcast and personal message planning, attachment and inline-image metadata resolution,
dry-run VBS orchestration, fixture-backed smoke tests, and safety checks proving that induction workflow
modules do not create Outlook sessions, open folders, send mail, or delete mailbox items.

## Stage 4 - Add Shared Contracts And Example Workbooks

Define workbook and template conventions and anonymised examples.

Stage 4 includes shared JSON schema definitions, anonymised workbook/template
examples, dispatch-state fixtures, and a cross-runtime shared-contract fixture
that represents both group meetings and induction through the same canonical
message, recipient, delivery, dispatch, workflow-context, and health-report
concepts.

## Stage 5 - Add Python Validators And Test Harnesses

Use Python for schema validation, fixture generation, regression tests, cross-runtime checks, and optional
CLI support.

## Stage 6 - Add Watchdog Sidecar

Implement scheduled diagnostics and managed-only repair.

## Stage 7 - Public Sanitisation And GitHub Release

Remove private paths, emails, names, logs, zips, SharePoint links, institutional identifiers, and any
non-public operational content.
