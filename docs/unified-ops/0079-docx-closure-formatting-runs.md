# 0079 — Shared formatting runs for DOCX closure edits

Record 0078 established paired batch measurements and target reuse. This
checkpoint connects public `transaction::Edit::set` to the same native-run
coordinator. The full goal remains active.

## Result

Consecutive direct paragraph/run formatting sets inside `File::edit` or
`edit_with_limits` now share a native editor and retained commit. Each ordered
assignment still runs against projected properties. Calls count independently
in `Diagnostics::operations()`; `commits()` reports actual retained commits.
Native operation capacity and aggregate source/projected XML limits are shared
with command batches. Empty-property calls retain their existing admission and
no-op behavior through the ordinary handler.

Typed target and property inputs are converted once before planning. Closure
calls pass their parsed `TargetArg` directly; they do not stringify typed paths
to enter the command-channel cache. The last-spelling shortcut remains specific
to commands that already have a spelling. No new public API or format support
is introduced by this checkpoint. A case-insensitive `text` key routes directly
through the ordinary handler, avoiding duplicate target/property admission for
text-only or mixed text/formatting calls. The planner's native text-setter check
remains as a defensive barrier for future aliases.

Reads (`get` and `find`), selector mutations, text and structural operations
close the pending formatting run. Reads and selectors therefore observe all
prior successful assignments, and structural operations retain sequential
addressing. Successful setters stage work; a native commit failure can surface
at the next barrier or outer closure completion. It retains its originating
mutation index.

## Failure isolation

The shared poison helper retains the first mutation failure. It covers target
and property conversion, operation limits, native staging and native-run flush.
A caller that catches one of these failures cannot publish earlier candidates
by returning `Ok`. Ordinary read lookup errors after a successful flush remain
recoverable, as before. Operation limits precede target/property conversion.
The final physical-source fence still takes precedence over stale results.

Best-effort command isolation flushes a successful item's pending run before
computing its changed flag. A failed item drops its entire pending native editor,
clears that item's poison and restores the prior owner, retained-step count,
mutation count and XML charge. Thus this closure change does not accidentally
coalesce across best-effort items or publish a failed item's partial assignments.
`File::apply` keeps the same per-item behavior.

## Evidence

- The initial combined unit/batch/transaction run passed 453 tests.
- The DOCX feature suite passed 678 tests/doctests after four new closure tests.
- The final transaction suite passed 19 tests after adding a selector regression;
  the minimal-feature suite passed 418 tests/doctests.
- Tests compare eager/source results to independent single-call byte oracles,
  check exact inverse replay, typed targets, get/find barriers, structural
  readdressing, selectors observing pending formatting, caught conversion errors,
  operation-limit precedence and cancellation back to the exact original archive.
  A 100-call closure fits one snapshot pair under a 2 KiB XML charge.
- Read-only review found no correctness blocker. After the text-family early
  routing refinement, the combined batch/transaction suite passed all 44 tests,
  and final all-target Clippy passed with warnings denied. Rustdoc with warnings
  denied passed; scoped formatting, diff hygiene and the retained benchmark
  control patch's applicability were checked.
- Existing command regressions still cover original wire indices, per-item
  rollback, native-capacity splitting, source fences and changed flags.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --test docx_transactions
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features docx --no-deps
```

This is correctness and integration evidence, not a closure performance claim.
Wider Edit families, best-effort run grouping, varied-target performance,
text-model completion, remaining formats and Python remain required work.

The rewritten base remains an ancestor of HEAD. Both required staged evidence
size/binary-directory checks printed nothing, and the final artifact audit found
no prohibited evidence or compiled binaries. `git add -A` could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` because the
linked Git metadata is read-only. No commit was created. The changes remain in
the checkout, and the full goal is active.

Cleanup removed the build output and checkpoint scratch/cache files. No
temporary worktree was created, and the earlier small benchmark evidence and
intentional harness lockfile were preserved.
