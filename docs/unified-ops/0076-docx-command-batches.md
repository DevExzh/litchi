# 0076 — DOCX command batches with atomic and best-effort execution

The preceding turn made verified progress on explicit closure transactions
(0075). This checkpoint connects the decoded command channel (0073) to that
coordinator. The full goal, run coalescing and performance acceptance remain open.

## Result and execution contract

`File::{apply, batch, batch_with_limits}` now execute the integrated DOCX
`set`, `add`, `remove`, `move` and `swap` subset. `apply` returns
`transaction::Commit<Applied>`. `BatchCommit` carries `file`, `patch`,
`diagnostics` and ordered `results`; each result is `ItemResult::{Applied,
Failed}`. These types are exported by the ops crate and facade as well as the
`batch` module. No separate batch execution error type is introduced.

`Applied` reports a per-item `changed` flag and an optional created canonical
path. That path is boxed only when present, avoiding the grammar path's large
inline buffer in every ordinary set/remove result. The returned path is
positional, not a stable identity after later structural operations.

Every command resolves against prior successful native results. Atomic mode
stops on the first failure. Best effort records ordinary command failures,
discards their complete candidate, and continues. Source/I/O, protection,
cancellation, resource and allocation/internal failures abort the complete
batch in either mode. Neither mode saves or mutates the caller's original.
Iterator unwinding likewise drops the candidate.

A private `Edit::isolated` savepoint retains the current handle, step count and
XML charge. It consumes the item's poison and restores all three on failure;
public closure mutations remain poison-on-first-failure. The original wire
index replaces the native successful-step index outside isolation. Failures
and metadata therefore never renumber subsequent error attribution.

`SemanticLimits` is shared across all successful steps. Its operation count
also bounds all input items, including failures and metadata. The coordinator
checks that count before converting an item and reserves results fallibly,
without trusting iterator size hints. Detecting overflow requests at most one
excess item; count overflow aborts the complete batch. Retained XML is charged
across the batch, never reset per best-effort item. As in 0075, this is a
logical main-XML charge rather than a total heap measurement.

Aggregate diagnostics count successful mutations. Metadata is not a mutation.
Per-item changes can all be true while the final artifact is unchanged after
cancellation. The shared coordinator preserves the exact original archive in
that case and supplies its owner-bound, source-guarded inverse patch.

## Command admission and reference differences

Reference evidence was read from the base checkout's OfficeCLI
`CommandBuilder.cs` (`ExecuteBatchItem`), `CommandBuilder.Batch.cs`,
`Core/BatchCompat.cs`, and Word mutation handlers. The implementation is
independent Rust code and follows record 0015's immutable execution contract.

- Command names match ASCII case-insensitively. `command`/`op` remain transport
  aliases, with duplicate envelope fields refused by the decoder.
- Set requires `path` and nonempty `props`; removal requires `path`.
  Mutation selectors travel through `path` and use the existing scope guard.
- Add accepts exactly one parent spelling (`parent` or legacy `path`), an
  element `type`, optional properties and one placement. Copying through
  `from` remains an explicit capability gap; simultaneous `type` and `from`
  is an argument error rather than an ignored type.
- Move requires `path`, rather than OfficeCLI's implicit root fallback. It
  accepts destination `to`, zero-based `index`, or a relative `after`/`before`
  placement. A parent plus index uses the existing typed inside placement.
  Parent plus relative anchor and any supplied move properties (even an empty
  object) remain explicit capability gaps.
- Swap requires `path` and exactly one of `path2` or legacy `to`.
- Negative indices and conflicting placement fields refuse. Supplied fields
  outside each command's mask are rejected, not silently ignored. Conflicting
  aliases refuse instead of using OfficeCLI's field precedence.
- Known read/raw/import/validation commands report `unsupported_capability` in
  this result channel; unknown commands report `unsupported_command`. Their
  standalone read APIs remain separate from mutation results.

A `meta` item must contain only `dumpVersion: 2`. Such items produce no-op
results and preserve original input indices, wherever they occur. Without any
metadata the input uses current semantics. Missing, old or future versions
refuse the complete batch, even in best-effort mode. Skipping a declaration
would silently change interpretation of other commands. Legacy v1 normalization
remains required work; it is not approximated here.

Inspection found an existing text-model mismatch: the native DOCX run setter
maps LF to a soft break, while OfficeCLI v2 uses LF for paragraph boundaries
and vertical tab for soft breaks. Batch `text` assignments containing LF, CR or
vertical tab therefore refuse with a discoverable capability gap until the
shared semantic handler represents that distinction. This applies with or
without v2 metadata. No dump replay compatibility is claimed for those inputs;
implementing the shared paragraph/soft-break semantics remains a follow-up.
The existing standalone typed handlers are unchanged by this checkpoint.

## Discovery and remaining work

DOCX `Verb::Batch` now covers the batch and single-command channel, with the
same protection and retained-source gates as `Verb::Edit`. Contextual gaps
name read/extended commands, text break semantics, old metadata, copying,
move properties and combined parent/relative placement. Encrypted-owner edit
refusal also names batch. Capability goldens and the generated OfficeCLI delta
are refreshed from this executable state.

This executor still closes one native commit per mutation, exactly as 0075.
Homogeneous family coalescing/fusion and commit-error bisection, advisory
results, the remaining command families, all other format writes, Python and
the specified scaling/performance evidence are still required. This checkpoint
does not claim M2 or full-goal completion.

## Verification and cleanup

- The all-feature suite passed 1,039 tests/doctests before the final strict
  text-admission and compact-result changes. Final targeted all-feature checks
  then passed 18 batch regressions and eight capability regressions.
- The final DOCX-only suite passed 665 tests/doctests. The minimal-feature suite
  passed 418. Both facade transaction/command tests passed.
- Final all-feature all-target Clippy and rustdoc passed with warnings denied.
- Scoped rustfmt, `git diff --check`, crate-boundary validation and the generated
  OfficeCLI delta check passed.
- Read-only review confirmed savepoint rollback, aggregate budgeting and original
  item attribution; its empty-move-properties finding was fixed and regression
  tested. No review worktree or temporary artifact was created.

Tests exercise byte equality to single-operation oracles for both owner forms,
all five commands, sequential creation and addressing, scoped fan-out, per-item
rollback, poison isolation, error indices after metadata/failures, iterator
panics and unbounded iterators, aggregate XML limits, source changes, early
protection admission, metadata failures, text-break refusal, cancellation and
inverse replay. The result-layout assertion bounds ordinary item storage.

The local rewritten base `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an
ancestor of HEAD. No rebase onto that revision was necessary. Required staged
size and binary-directory checks printed nothing. The artifact audit covered
334 changed/untracked paths with no prohibited evidence or compiled ELF files.

`git add -A` still cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. Changes remain in the checkout;
the full goal remains active.

Reproduction commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features --test docx_batches --test capabilities
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi --no-default-features --features docx --test unified_transaction
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --all-features --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
```

Cleanup removed 23,596 build files (34.8 GiB) and 12 scratch/cache paths,
including generated untracked lockfiles. No temporary worktree or external
fixture link was created by this checkpoint.
