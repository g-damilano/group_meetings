# Repository Boundary Audit

The scaffold can be ready while the repository around it still contains frozen
legacy/reference or generated material. The boundary audit inspects that
surrounding material without modifying it.

Run:

```text
python scaffold\python\repo_boundary_audit.py --repo-root . --scaffold-root scaffold
python scaffold\python\repo_boundary_audit.py --repo-root . --scaffold-root scaffold --format json
```

The default command is report-only and exits successfully even when findings are
present. To use it as a release gate after an authorised repo-wide cleanup:

```text
python scaffold\python\repo_boundary_audit.py --repo-root . --scaffold-root scaffold --fail-on-findings
```

The audit skips `scaffold/` and `.git/`, then reports only paths, categories,
and counts. It does not print matched private source content, copy files, move
files, delete files, publish releases, send email, or touch Outlook.

If frozen material is hard-copied into `scaffold/legacy/`, that mirror remains
outside the public release surface because this audit skips all of `scaffold/`
and the scaffold release tools exclude `legacy/`. The mirror is a private
migration aid, not a sanitised release output.

Finding categories include generated artifacts, high-risk binary attachments,
generated cache directories, private text-pattern matches, and non-example email
addresses. Findings outside `scaffold/` are advisory for scaffold-local
development and release-gate decisions by default. An explicitly authorised
repo-wide cleanup may still use `--fail-on-findings` to treat them as blockers
for a wider repository publication decision.
