# Clean scaffold preflight: source-owner mismatch

This preflight ran from clean committed scaffold
`2a2ffa1cae4e6b7070082768ce84483e5d411dc8`. Clean checkout, local source
manifest, lock identity, and the bounded 23-recipe/42-lane corpus checks pass.
The owner-pinned source checks **fail**. No build, host probe, matrix capture,
or timing run was authorized or performed by this preflight.

The production closure differs from the original semantic-owner commit
`cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd` in 11 OPC/PPTX paths listed in
`preflight-summary.json`. Root traced those differences to three commits:

- `f2e68660e`: canonical namespace handling for markup compatibility.
- `83e60e8c6`: source content-type reuse for unchanged package membership.
- `aab225305`: source-relationship capture admission before allocation.

These are real source changes, not uncommitted checkout noise. A current
baseline requires explicit source review and an exact source pin separate
from the historical owner identity. The failed equality check must not be
silently disabled or relabeled as passing. This receipt does not approve that
baseline change.

`retained-files.json` records the original files' lengths and SHA-256 hashes.
All files were copied byte-for-byte, preserving diagnostic output and original
external paths. The original worktree is
`/var/tmp/litchi-pptx-ink-actions-profiler-2a2ffa1c`; the external preflight
directory may be removed only after committed-byte verification and completion
of any review reading it.
