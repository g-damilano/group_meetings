# Release Gate

The release gate is a Stage 7 report-only publication control. It converts the
release candidate plan into an explicit `PASS` or `BLOCKED` result.

Run:

```text
python scaffold\python\release_gate.py --root scaffold --repo-root .
python scaffold\python\release_gate.py --root scaffold --repo-root . --format json
python scaffold\python\release_gate.py --root scaffold --repo-root . --fail-if-blocked
python scaffold\python\release_gate.py --root scaffold --repo-root . --write-json legacy\reports\release_gate.json
```

The default command writes nothing. If `--write-json` is used, the output must
stay under `legacy/`, which is private local migration material and excluded
from the public scaffold release surface.

Gate values:

- `PASS`: scaffold release checks are clear, the public manifest excludes
  `legacy/`, and any parent-repository findings are advisory only.
- `BLOCKED`: scaffold-local validation fails, public scaffold files contain
  private material, the public manifest includes `legacy/`, or smoke/tests are
  broken.

Parent repository findings outside `scaffold/` are advisory by default for
scaffold-local development. They are still reported, but they do not make the
gate `BLOCKED` unless scaffold-local release validation fails or the public
manifest accidentally includes private `legacy/` content.

The gate does not package, publish, tag, send email, touch Outlook, mutate
mailbox data, edit legacy files, or modify copied mirror files. With
`--fail-if-blocked`, it returns exit code `1` for automation while preserving
the same report-only behaviour.
