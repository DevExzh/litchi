# 0801 — no-replay checked-attribute preflight

This packet tests the deferred no-replay design from 0800, freshly rebased on
that batch's five-copy duplicate-error correctness repair. Production remains
unchanged during the experiment. Both helper legs run in one independent probe
binary against the same literal inputs; historical timing is not pooled.

The hypothesis is that carrying the first two key positions forward removes
0799's replay cost at the third item, while retaining the common one/two-item
benefit. The candidate prechecks raw keys before value parsing and seeds an
ordered map only when a successful third item has a non-whitespace tail. This
also moves map costs earlier, so every longer-tag regression remains visible.

The frozen plan retains all 39 cases, both construct/consume modes, six paired
native blocks of 30 samples with 4,096 iterations, and two separate Callgrind
repeats. Native children are pinned to CPU 12. Callgrind counters describe guest
instructions and branches, not native timing or allocation API counts. Inputs
are hot and repeated; opaque iterator construction and checksum work are part
of the direct-probe protocol. No public-workflow or resource improvement is
established by this packet.

Advancement requires semantic parity, at least 3% benefit with the ratio CI
entirely below 1 for both distinct-1 and distinct-2 consume cases, and no
protected consume regression above 5% with CI entirely above 1. The 18 protected
cases and all thresholds are identical to 0799. Other regressions remain review
triggers. Passing only authorizes fresh public-workflow/resource/cross-format
trials; it does not authorize adoption. Failing retains the production baseline.

Root owns all builds, helper tests and captures, serially. Both exact five-copy
helper implementations and their shared tests are exercised in minimal mirror
crates before the release probe is built. These are isolated helper checks,
not full production-crate verification. Independent source review and offline
analysis run separately from live measurements. Raw logs, frozen inputs,
failed attempts if any, source hashes, and cleanup identities are retained.

## Final result

Rejected: both short-tag benefit gates pass, but empty consumption is 7.542%
slower (ratio CI 1.062726–1.077057), the sole protected veto among 13 consume
regressions. Production remains unchanged. See the main 0801 report and
`summary.md` for every result and spread flag. All 1,248 reports and 28,392
samples replay successfully after removal of the owned target. The source-only
manifest status records handoff time; the final receipts record subsequent
execution. No workflow adoption or resource claim follows.
