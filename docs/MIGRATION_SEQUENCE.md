# Migration Sequence

## Stage 0 - Freeze Production Legacy

Preserve the current daily-use scripts as the operational baseline.

## Stage 1 - Extract VBS Core Primitives

Move Outlook session, composer, dispatch, diagnostics, logging, name-resolution, and template primitives into reusable VBS modules.

## Stage 2 - Move Group Meetings Onto VBS Core

Keep behavior stable while replacing embedded COM routines with core calls.

## Stage 3 - Move Induction Mailer Onto VBS Core

Reuse the same composer, recipient resolver, template engine, and managed dispatch.

## Stage 4 - Add Shared Contracts And Example Workbooks

Define workbook/template conventions and anonymized examples.

## Stage 5 - Add Python Validators And Test Harnesses

Use Python for schema validation, fixture generation, regression tests, and optional CLI.

## Stage 6 - Add Watchdog Sidecar

Implement scheduled diagnostics and managed-only repair.

## Stage 7 - Public Sanitisation And GitHub Release

Remove private paths, emails, logs, zips, names, and institution-specific content.

