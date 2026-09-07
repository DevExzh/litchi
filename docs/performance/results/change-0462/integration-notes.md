# Integration and capture notes

The user-requested fetch/rebase completed before this experiment; the branch
was already up to date. Baseline revision is `dbd2f8ece`. The only initial
untracked file was user-owned `docs/GOAL.md`; it remains untouched and unstaged.
The unchanged hashes of all accepted ADRs and their README carry forward the
prior complete review, with the index-specific constraints in `adr-refresh.json`.

Baseline release executables were built and copied before the source edit ACK.
Primary and supplementary protocols were frozen before baseline R1. Baseline R1
was invoked directly with `capture.py baseline R1`; its six immutable lane
receipts bind the protocol, executables, raw reports, resources and independent
oracle. The later `checks/baseline-r1.json` gate validates those existing six
reports and their passing receipts without rerunning measurements. Other
primary repeats use `check.py` around their capture invocation. No timing rerun
is used to seek favorable results.

The fixed index is confined to shape parsing; generic attribute callers retain
their existing state layout. Correctness requires lazy decoding, first appended
occurrence, exact namespace URI identity, error reachability, eager style-name
fallback, and original drawing-attribute harvest order. Fresh decode failure
advances without caching the failed attribute; cached invalid values remain
replayable. The source review and tests determine whether the candidate meets
those constraints.

Primary lifecycle clocks and allocator regions are distinct from supplementary
phase diagnostics. Phase clocks surround public API calls; whole-process perf
counters and maximum RSS also include setup, warmups, checks and reporting.
No cold/range/native-GUI/worker-scaling claim follows from these runs.

The initial strict owner check failed at compilation: production namespace
import, stale SVG import, unnecessary size-of qualification, and non-comparable
error assertions. That original source epoch and failed receipt remain intact.
The corrected `owner-clippy-r1` and 379 owner tests pass. The separate layout
capture reports ElementAttrs 144 bytes, ShapeAttrs 424 bytes, fixed index
280 bytes, and ResolvedAttribute 80 bytes on this x86_64 target.

The assembly driver's automatic `known_cached_scan.eliminated` field is a
flawed direct-call heuristic. Its baseline false positive is documented in
`assembly-review.md`; it is not used for the decision. Manual review follows
the actual indexed hit/miss branches and retained defensive fallback. Both
raw assembly receipts and outputs remain unchanged.

Candidate harness tests pass 387 cases with one ignored. Both measured Rust
files are restored byte-exact to the baseline hashes in `decision.json`. Final
restored-source owner-Clippy-r2, warning-denied rustdoc, scoped formatting and
crate-boundary checks pass. No production Rust change is retained.

Precleanup validates both retained executable epochs and the restored live
source. Fresh-copy portable replay passes; a resealed summary with its first
p50 changed by one nanosecond is rejected at recomputation. Cleanup inventories
and removes only four staged executables totaling 233,043,088 bytes under
`/tmp/litchi-goal-0462`; the directory is absent. Shared Cargo targets and
user-owned `docs/GOAL.md` are preserved.
