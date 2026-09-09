# 0490 file-store variance and synchronization attribution

This bundle investigates the small file-store tail regressions retained by
0489. It uses existing 0487/0489 executables with alternating before/after
process order and independent synchronization diagnostics. The fixed
[method](methods.md) retains the data-sync contract, all oracles and explicit
uncertainty. No production change is justified until measurements are reviewed.

The [provenance](provenance.json) binds the goal, taxonomy, starting commit,
protected local files, and retained builds. All accepted ADR hashes are unchanged.
The spec-gap worktree and other sessions' active work are excluded.


The [results review](results-review.md) retains 72 formal children, 4,320 samples,
eight diagnostics, uncertainty and every adverse quantile. Synchronization is
about 79–82% of the traced median operation; the data-sync contract remains.
No new production change follows. The [next implementation](next-implementation.md)
addresses verified-cold/source-provider coverage and preserves the broader goal.
See `variance.py --help`, `sync_profiles.py --help`, and `seal.py --help` for
reproduction and verification commands; use the recorded source checkout.
