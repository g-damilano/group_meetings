# Release Manifest

The release manifest is a Stage 7 inventory aid for public review. It lists the
files under `scaffold/` using release-relative POSIX-style paths, byte counts,
file categories, and SHA-256 hashes.

Run:

```text
python scaffold/python/release_manifest.py --root scaffold
python scaffold/python/release_manifest.py --root scaffold --format json
```

The command writes nothing by default. It prints the manifest to stdout so a
reviewer can inspect or capture it deliberately without creating generated
files inside the scaffold tree.

The manifest excludes generated or release-blocked artifacts using the same
sanitisation conventions as the scaffold validator, including bytecode caches,
logs, message files, and archives. If any such artifacts are present, they are
reported as excluded rather than treated as release files.

The manifest also excludes `legacy/` completely. That directory is a private
copy-only migration mirror for raw legacy/reference/generated material and must
not leak into the public scaffold inventory.

The manifest is not a package and is not a publication step. Repository-wide
release actions remain outside the scaffold boundary and require explicit
authorisation.

Use `python/release_candidate_plan.py` when a single report should combine the
manifest with scaffold release readiness and repo-boundary status.

Use `python/release_export_plan.py` when the manifest should be presented as a
release-review export plan with kind summaries, total byte counts, and an
explicit `legacy_paths_in_files` guard.
