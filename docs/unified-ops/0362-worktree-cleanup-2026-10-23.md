# Worktree cleanup — 2026-10-23

Two inactive merged leaves were removed with `git worktree remove` without
`--force`. Their branch refs remain available for reproducible source recovery.

- `uo2-xlsb-provenance-review`: HEAD `7b3fd27e2ec9f6be3103807932828f183981b4fb`, ancestor of integration `a814baa1de`; 37,389 bytes reclaimed. It had no ignored or untracked files.
- `uo2-iwa-detection-api`: HEAD `7f3aef217eb9090199b8389a8d70bca9633e4ee3`, ancestor of integration `a814baa1de`; 341,906,569 bytes reclaimed. Eleven generated ignored files (461,243 bytes) were moved first to the quarantine directory below, with relative paths and SHA-256 records.

Both trees were clean, had no targets, protected measurement binaries, or active
process/file-descriptor users. Their branch refs were checked after removal.

The DOCX section/redaction leaves contain ignored performance captures and
portable bundles, the replay-error-audit leaf has tracked changes, and the W8
leaves contain active/reference targets or ignored artifacts; those trees were
retained. Full inventory, quarantine manifests, hashes, and before/after
worktree lists are in `.coord/worktree-cleanup-2026-10-23/` outside the
checkout.
