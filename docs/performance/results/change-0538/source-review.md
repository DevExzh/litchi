# 0538 source and measurement review

The delegated read-only review found no blocking correctness issue in the
planning bracket. Root also inspected each constructor and the accumulation
path. The allocation region begins after `xlsx_update_sheet_selectors` and
before the existing planning clock. `finish` follows the clock endpoint and
precedes phase conversion, cell updates and commit. `?` retains the existing
typed error path and drops the region guard on failure. No edit/snapshot
ownership or lifetime changes accompany the observation.

The two lifecycle constructors set the new optional field to `None`; their
empty vectors remain omitted. The source-backed scalar publication constructor
sets `Some` for both managed and unmanaged modes, using the existing explicit
unavailable sample when the observer is disabled. Warmups are discarded before
summary recording, and vectors remain in acquisition order. Existing optional
commit/publication aggregation conventions are preserved.

The reviewer recommended explicit scope documentation and complete case-level
cardinality checks. The harness README now documents process-global scope and
the distinction between region peak and RSS. The retained verifier requires
aligned planning/commit/publication vectors for every one-cell, one-percent,
batch and multi-sheet case in both cache modes and corpus shapes. The focused
library test exercises unmanaged one-cell and managed multi-sheet paths with
warmups; the new integration test independently launches normal and allocator
binaries and validates status, counters, phase sums and semantic identity.

The review also suggested future fail-closed presence tracking in the generic
optional-vector accumulator. This batch keeps the existing aggregation model:
all current source-backed publication constructors supply a sample, both
lifecycle constructors explicitly omit it, and complete current-case
cardinality is enforced by the retained executable checks. Any future partial
observation path must preserve that all-or-none invariant.

The production worksheet parser remains byte-identical to the parent. This
enabler does not apply the 0537 candidate or claim a speedup. A fresh release
baseline is necessary because harness code and observation boundaries changed;
historical commit/publication vectors cannot stand in for planning metrics.
