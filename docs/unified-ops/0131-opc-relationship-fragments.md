# 0131 — Guarded OPC relationship-fragment transport

## Scope

OPC now exposes bounded lexical relationship catalogs and source-bound fragment
edit plans. Catalogs can decode durable XML bytes but grant no package mutation
authority. Plans check each touched ID against its exact destination element,
owner and inherited XML context before materializing a replacement token. The
existing package replacement path still performs its source and publication
audits before mutation.

Removal preserves every surrounding byte. Replacement retains the destination
slot and emits the supplied exact element bytes. Insertion uses retained
neighbor IDs and exact intervening gaps, resolving batches iteratively and
rejecting missing, cyclic, overlapping or stale anchors. A terminal insertion
can prove its slot from retained left-hand gaps, so inverse restoration of a
deleted suffix precedes unrelated entries appended later. Entries, comments,
processing instructions, foreign markup and other destination changes outside
touched elements remain intact. Unprovable gaps are conservative refusals,
including multi-entry insertion into an empty root when the source inter-entry
gaps have no retained destination counterpart. Physical member absence is separate: this plan
preserves an existing empty member and materializes a missing member when it
acquires entries; exact presence restoration remains the guarded API from 0129.

Context equality includes root namespace bindings and inherited `xml:base`,
`xml:lang` and `xml:space`. Context changes are refused; the planner does not
rewrite prefixes. Duplicate IDs, repaired identity catalogs, malformed XML and
nonstandard roots cannot become editable fragments. Catalog and plan diagnostics
summarize sizes/counts without dumping retained XML.

Output byte length and relationship count are checked before the output buffer
is allocated, including per-member lower bounds on aggregate quotas. The event
upper bound is conservative because deletion can merge adjacent text events.
Materialization rechecks limits, reserves one output buffer, validates the native
graph and captures an audited XML token. Exact no-ops share the original token.
Caller-owned package-wide reservations and format dependency closure are still
required. No performance acceptance claim is made.

This is a prerequisite, not the durable PPTX metadata fix: content-type lexical
transport, per-key wire deltas, join/three-way integration and source provenance
are still pending. Native patch replay still has the limitation recorded in
0128–0130. The all-format goal remains active and incomplete.

## Verification

The full OPC run passed 1,111 tests/doctests across 46 targets, with one existing
ignore. After final anchor and no-op refinements, all 13 focused fragment tests
passed. Coverage includes publication/reopen with unrelated overlays, every
nonempty deletion subset of three entries restored in either batch order,
adjacent suffix restoration ahead of unrelated appends, simultaneous
replacements with changed gaps, aliased self-closing roots, exact no-ops,
tightened structural limits, stale entries/context/gaps, cycles, missing anchors,
wrong owners, malformed catalogs and exact/one-under output quotas.

OPC all-target Clippy and documentation passed with warnings denied. Native
PPTX with encryption compiled against the new API. Verification was offline and
serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-opc --test relationship_fragments
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc --no-deps
CARGO_NET_OFFLINE=true cargo check -p litchi-pptx --features encryption
```

The checked base remains an ancestor of HEAD, so no rebase is needed. Required
staged size/binary-path guards printed nothing, and the changed/untracked
artifact audit found no oversized evidence or compiled binaries.

Read-only review found a tightened-limit no-op bypass; the final implementation
reuses full fragment parse/preflight before returning the shared token, with a
passing structural-limit regression. Final read-only review verified the fix
and reported no remaining checkpoint blocker.

The requested commit could not be created: staging failed because Git could not
create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No permission workaround was attempted. Changes remain in
the worktree. Checkpoint-owned build output, generated root lockfile and scratch
logs were removed after verification handles completed; no temporary worktree
was created. The pre-existing working changes remain intact.
