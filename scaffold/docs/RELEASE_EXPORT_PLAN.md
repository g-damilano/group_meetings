# Release Export Plan

The release export plan is a deterministic Stage 7 inventory for the public
scaffold release surface. It lists exactly which files under `scaffold/` would
ship, with byte counts, file kinds, and SHA-256 hashes inherited from the
release manifest.

Run:

```text
python scaffold\python\release_export_plan.py --root scaffold
python scaffold\python\release_export_plan.py --root scaffold --format json
python scaffold\python\release_export_plan.py --root scaffold --fail-if-invalid
python scaffold\python\release_export_plan.py --root scaffold --write-json legacy\reports\release_export_plan.json
```

The command writes nothing by default. If `--write-json` is used, the output is
guarded so it can only be written under `legacy/`, which is ignored and excluded
from public release manifests.

The export plan is not an archive and not a publication step. It does not copy
files, create zip bundles, tag commits, publish releases, send mail, touch
Outlook, or modify legacy/reference material.

The plan must report `legacy_paths_in_files: 0`. Any nonzero value means private
mirror material has entered the public scaffold export surface and the release
must be blocked.
