# 0362ul89aca — PPTX slide-root proof profile repair

Status: source-ready repair chained from
`93070d72e3f261496b9d57d0d41b75bfc47b4f7d` (89ac), based on the
same immutable qualification source. Cargo/build/rustdoc and combined-gate
validation remain coordinator work.

## Repair

Opened slide capture now admits each raw slide under the package's selected
`SourceLimits` event, depth, attribute-byte, and part-byte policy before MCE.
The source admission is performed once for a changed or first-seen payload;
an exact memo hit reuses that package-profile proof without adding per-entry
state or changing the 32-byte memo footprint. Rebinding under different
source limits drops the memo so it cannot cross physical profiles.

The notes graph accepts a default slide-root proof only when the recipient's
selected semantic depth and aggregate attribute-byte bounds are at least the
proof bounds. Tighter profiles re-run the bounded root classifier. Slide XML
and node ceilings remain independent from the notes 8 MiB/100,000-node
profile; source events/depth/attribute bytes remain physical constraints.

Added a focused regression comparing reused-proof and cold-scan refusals for
tighter depth and attribute-byte policies, including exact typed kind,
observed, and maximum values. Capture tests now count the physical admission
pass: an exact memo hit performs no source scan while a cold recapture admits
each slide once; a changed source profile drops the memo before recapture and
retains the 32-byte entry contract. The selected events, depth,
attribute-byte, and part-byte refusals are exercised through the real slide
capture path with exact observed/maximum values. The deferred root-only path
after a prior name-projection failure now admits its raw source before MCE as
well.

The physical pass is intentional: MCE and the notes root scanner may discard
opaque branches, so neither proves the selected raw-source events, depth,
attribute, or part-byte policy. It is one pass for a cold/changed allocation;
`source_checked` proofs prevent a second pass during notes-graph loading only
when their raw slice is the exact current blob observation. A foreign or
changed allocation is admitted again, while exact memo hits skip the pass
entirely. The profile-rebind regression keeps a same-profile positive control
and reuses the changed package's slide allocation, so an empty memo is proven
to result from the selected-policy change rather than serialization.

## Source-only checks

Rustfmt parsing/no-delta and `git diff --check` passed on this follow-up;
Cargo/native/full-feature tests, rustdoc, the workspace gate, and ABBA remain
outstanding. No generated evidence, fixture, executable, or `3rdparty/` path
is included.

Commit: reported in the coordinator handoff; this record is included in the
repair tip.
