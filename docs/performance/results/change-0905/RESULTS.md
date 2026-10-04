# 0905 result and replay entry point

The completed measurements are summarized in the [0905 report](../../0905-xml-attribute-summary.md).
`README.md`, `METHOD.md`, `MICRO_METHOD.md`, and `parity-method.md` retain their
pre-execution wording and hashes; this file supplies the later execution context.

The main reduction is `analysis.json`; the separate XML guards are
`micro-analysis.json`; phase, allocator, RSS, and serial workload intervals are
in `diagnostics.json`. Both primary benefit gates and the micro rejection guard
pass. The [decision](decision.json) accepts the candidate, and the independent
[result review](result-review.md) found no blocker. `review-clarifications.json`
records wording corrections while retaining the original review versions.
`quality.json` binds the successful checks and retained failed build
attempt. `parity-review.md` qualifies the generator's unreachable XML-declaration
branch and nonstandard diagnostic hash multiplier. Raw byte equality and
SHA-256 are the authoritative parity evidence.

The retained-artifact replay avoids the cleaned source/build directories:

```sh
python3 -B docs/performance/results/change-0905/replay.py --output /tmp/0905-replay-new.json
```

Use a new output path; helpers never overwrite evidence. The eight executables
in `binaries/` are retained locally and gitignored. A fresh checkout does not
include those binaries. Exact source reconstruction uses the retained 0894
archive, four `baseline-overlay/` files, `candidate-audit.rs`, source manifests,
and the separately bound oracle/probe examples. The build receipts record the
isolated paths, locked dependencies, Rust 1.95.0 toolchain, release settings,
commands, and executable digests required to reproduce their custody.

`concurrent-integration.json` records unrelated shared-checkout changes after
all captures and the final source check. The measured comparison is against
the frozen `4f39f5a88d` inventory plus the audit candidate, not a claim about
the combined later iWork/ZIP integration. No unrelated staged or conflicted
files were resolved by this experiment.

The terminal decision, cleanup, and post-cleanup replay receipts are the final
disposition records. They preserve the primary confidence intervals, the two
micro p99 flags and small two-attribute regressions, the source scope, and all
owned-root removal checks. No worktree was created by this batch.

Cleanup removed 7,379,644,211 logical bytes and post-cleanup replay passed.
The batch is committed separately on `perf/0905-xml-attribute-summary` through
an isolated index while the shared target branch has an unrelated cherry-pick
in progress. Its source/report changes remain mirrored in the shared working
tree; target-branch integration is pending and no unrelated index entry was
staged or resolved by this batch.
