# Release Candidate Plan

The release candidate planner is a Stage 7 scaffold-local decision aid. It
combines scaffold release evidence, the public release manifest, advisory
repository boundary findings, and private legacy mirror status into one
report-only decision.

Run:

```text
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root .
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root . --format json
python scaffold\python\release_candidate_plan.py --root scaffold --repo-root . --write-json legacy\reports\release_candidate_plan.json
```

The command writes nothing by default. If `--write-json` is used, the output
must stay under `legacy/`, which is excluded from the public scaffold release
surface.

Decision values:

- `READY_FOR_AUTHORISED_RELEASE`: scaffold checks pass and the public scaffold
  manifest excludes private legacy mirror content. Parent repository findings
  may still be present as advisories.
- `BLOCKED_BY_SCAFFOLD_VALIDATION`: scaffold release checks failed.
- `BLOCKED_BY_REPO_BOUNDARY`: retained for compatibility with old reports, but
  parent repository findings are advisory by default for scaffold-local
  development.

The planner does not package, publish, tag, send email, touch Outlook, modify
mailbox data, edit legacy files, or mutate copied mirror files. It reports
counts and categories rather than matched private source content.

Use `python\release_gate.py` when the same evidence should become an explicit
publication gate for automation, and `python\release_export_plan.py` when the
public scaffold file inventory should be reviewed before release.
