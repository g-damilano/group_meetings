# Legacy Alternate Map

The scaffold may copy or port legacy behaviour into `scaffold/` when the result
is sanitised and remains separate from the frozen production files. Moving
legacy files is not allowed.

The machine-readable map lives at:

```text
scaffold/migration/legacy_alternate_map.json
```

It records which outside-root legacy/reference files have scaffold-local
alternates, which scaffold path is now used by the rework, and which safety
boundary applies.

Run the read-only intake audit with:

```text
python scaffold\python\legacy_intake_audit.py --repo-root . --scaffold-root scaffold
python scaffold\python\legacy_intake_audit.py --repo-root . --scaffold-root scaffold --format json
```

The audit checks that mapped legacy sources still exist, scaffold alternates and
routes exist, and any source that would require sanitisation is reported without
printing private source content. It does not copy, move, edit, publish, send, or
repair anything.

Current production-shaped alternates:

- `scaffold/RunGroupMeetings.vbs` routes to the group-meetings dry-run
  entrypoint.
- `scaffold/RunInductionMailer.vbs` routes to the induction dry-run entrypoint.
- `scaffold/OutlookWatchdog.vbs` routes to the watchdog report-only entrypoint.
- `scaffold/ValidateWorkflow.vbs` routes to scaffold validation.

These wrappers are intentionally similar to legacy launcher names so operators
and tests can exercise the reworked path without modifying production scripts.
They are not drop-in replacements for live sending. Live send, destructive
mailbox repair, moving legacy files, and raw private content imports remain
outside this scaffold-local migration path.
