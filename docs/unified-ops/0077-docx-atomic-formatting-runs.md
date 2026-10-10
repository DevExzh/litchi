# 0077 — Atomic DOCX formatting runs

Record 0076 connected the command channel to immutable transactions. This
checkpoint groups the first compatible family: consecutive atomic-batch
`set` commands on direct paragraph/run paths containing only formatting
properties. The complete goal and performance acceptance remain open.

## Execution contract

A run retains one native editor and applies each ordered assignment against
its projected properties. Direct paragraph/run ordinals remain stable because
these setters only change property markup. Repeated keys are not merged.
Per-item `changed` compares projected XML before and after that item. Text,
selectors, structural commands, metadata and other commands close the run.
Best-effort execution and public closure mutation calls still commit separately.

The text barrier is intentional: native formatting preserves exact markup,
whereas text-only publication may compact text markup. Combining these policies
could change the single-operation byte oracle. Existing owner/source/security
admission and native managed-operation refusals remain in force; grouping does
not expand the supported document surface.

`Diagnostics::operations()` counts successful caller mutations, including
no-ops. New `commits()` counts retained native commits. The primary native
editor and snapshot editor expose `remaining_operations()`; a run flushes
before an item's conservative assignment count exceeds that remaining ledger
capacity. An individually oversized command still receives a native limit
error. Item boundaries are never split.

Pending runs charge their source plus projected main XML once, in addition to
all already retained commits. This reduces retained snapshot accounting for
compatible runs. It remains a logical XML charge, not a total heap bound;
property plans, native operation ledgers and temporary error probes are not
represented by that number. Existing command and native operation limits bound
counts independently. A staging failure discards the entire atomic candidate,
including any earlier assignments in that failed item.

A native commit failure triggers error-path-only prefix/suffix bisection with
fresh editors. Successful prefixes become the base of later probes. Original
wire indices survive metadata and barriers. Source/I/O, cancellation, protection,
resource and internal/allocation failures are never retried. A group-only failure
is attributed to the first item of the run. Unindexed flush accounting/allocation
errors likewise receive the run's first index. Final source guards still run
when an iterator changes the physical source at EOF.

## Verification

- The DOCX feature suite passed 669 tests/doctests before the final four added
  regressions (two unit probes and two batch cases).
- The final focused batch suite passed 24 tests. It covers source/eager byte
  equality to separately committed operations, duplicate-key ordering, per-item
  no-ops, exact inverse replay, opaque property markup and uncompacted text,
  barriers, pending-run rollback, EOF source mutation, native-capacity splitting
  at 4,097 operations and aggregate retained-XML limits.
- Two deterministic fault-injection unit tests exercise commit-failure bisection,
  original indices, group-only failures and the no-retry global-failure boundary.
  They inject probe errors; they are not evidence of a naturally failing native
  formatting commit.
- After the final flush-error attribution fix, the combined unit/batch/transaction
  run passed 452 tests. Read-only independent review found no blocking regression.
- Minimal-feature tests/doctests passed 418 cases. DOCX all-target Clippy with
  warnings denied and DOCX rustdoc with warnings denied passed.

Reproduction:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --lib --test docx_batches --test docx_transactions
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features docx --no-deps
```

This checkpoint establishes semantic equivalence and reduced commit/retention
counts. It does not establish the goal's native-relative performance threshold.
Paired 1–10,000-operation benchmarks, wider compatible families, best-effort
coalescing, remaining formats and Python coverage are still required.

The rewritten base remains an ancestor of HEAD; no rebase was needed. Both
required staged evidence-size/binary-directory checks printed nothing.
`git add -A` failed because the linked worktree's Git metadata is read-only:
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` cannot be
created. No commit was created; implementation and records remain in the checkout.

Final hygiene: scoped formatting, `git diff --check`, crate-boundary validation
and the generated OfficeCLI capability-delta check passed. The artifact audit
covered 336 changed/untracked paths and found no prohibited evidence or ELF
binaries. Cleanup removed 9,955 build files (6.7 GiB), checkpoint scratch logs,
and generated untracked lock/cache files. No temporary worktree was created.
