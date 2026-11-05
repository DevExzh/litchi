# 0130 — Semantic PPTX slide-order selectors

## Public API

The native `slide::Key` type and root `Key` reexport are deleted. Canonical
`slide::Selector` carries `Name(Cow<str>)` or `Position(Position)` and accepts
borrowed/owned names and checked positions. The existing numeric convenience
conversion lowers explicitly to `Position`; no ordinary raw-ID selector is
introduced. All live native consumers now use this vocabulary. Producer-visible
name matching remains exact, and ambiguous names remain typed failures.

`opened::Edit::move_slide` accepts a semantic source selector and an explicit
checked destination position. The destination is the final slot after detachment.
`swap_slides` exchanges two semantic selections in one edit. `reorder_slides`
accepts a complete selector permutation resolved against the same staged order;
its old raw slide-ID lowering is private. Existing XML/dependency admission and
commit/signature rules remain in force. Invalid, missing, duplicate and ambiguous
selectors leave the staged order unchanged. Identity operations are exact no-ops.

Order selection uses the existing name index without cloning a slide or an owned
caller name. Positional selection checks the current collection directly. Order
vectors reserve fallibly, and the no-op identity comparison no longer allocates
a temporary vector. These are implementation facts, not measured speed claims.

Both PPTX glossary profiles now reject the removed root/slide `Key` and the
`Index` selector variant. Native tests and the performance harness's live move
routes use checked destination positions. Frozen historical evidence is retained
unchanged. Other legacy order families and raw positional read signatures still
need migration; this does not claim complete PPTX vocabulary conformance or
unified facade move/swap support.

## Verification

Three new public regressions cover composed named move/swap/mixed-selector
reordering, durable replay and inverse, original preservation, publication/reopen,
failed selector/permutation rollback and exact no-ops. Existing ambiguous-name
coverage now checks all three order verbs and successful positional recovery.
Selector conversion tests cover borrowed and owned names and checked positions.

The final native encryption run passed 982 tests across four targets (967 unit,
eight byte-order-marked publication, four signature-policy and three new order
tests), with one existing unit ignore. The PPTX/encryption facade suite passed
533 tests and doctests across 26 targets. The live performance harness compiled. Native/facade all-target Clippy and 19
native doctests passed with warnings denied. Read-only review found no
correctness or atomicity blocker. The pre-existing ID-to-slide lowering still
uses a bounded quadratic lookup for large permutations; this wave makes no
performance acceptance claim. Both PPTX glossary profiles passed: 79 banned
names in minimal and 70 in full, with zero transaction families claimed.

Verification commands (offline and serialized):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib --test slide_order_selectors --test byte_order_marked_parts --test opened_signature_policy
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/unified-ops-api/target cargo check --manifest-path tools/perf-baseline/Cargo.toml --lib
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Metadata integration finding and required next work

The 0128 replay defect remains open. Review confirmed that whole-owner metadata
images alone cannot represent composable lossless patches: independent edits can
share the content-types or relationship owner, and inverse removal must retain
the removed entry's exact XML. Completing a smaller exact-image-only patch would
not establish the required transport contract.

The required representation is a bounded per-owner, per-key metadata delta:
content-type overrides keyed by part name (defaults separately by extension),
and relationships keyed by owner and relationship ID. Each touched entry needs
semantic before/after state, exact before/after fragments and a stable lexical
insertion anchor. Relationship member presence is separate. Optional complete
owner images enable exact-source restoration through the guarded APIs from 0129;
images are provenance, not whole-owner conflict keys.

Join and three-way resolution operate on entry keys. Disjoint entries merge;
divergent replacements of the same key conflict. Inverse swaps complete entry
states, including removed fragments. A composed patch may discard its exact-image
shortcut while retaining its structural deltas. Root relationships should use
the same keyed representation instead of the current whole-root conflict rule.

The missing native primitive is a bounded OPC lexical metadata edit plan that
can retain removed fragments and install validated raw fragments at stable
anchors. Existing semantic `plan_edit` methods preserve untouched source ranges
but generate canonical additions, so they cannot alone restore a removed entry's
original attributes/prefixes/spacing on a divergent destination. Extend the native
planner to validate touched keys, namespace context and anchors, calculate output
and aggregate bounds, and preserve every untouched byte before installation.

Capture must avoid metadata scans/copies on ordinary text-only patches. Use
changed owner/provenance information to capture only affected metadata. The
current package fingerprint is semantic and omits raw metadata; package equality
already additionally compares retained publication manifests. Metadata-only
changes must therefore be represented explicitly, rather than inferred solely
from that fingerprint. Regression coverage must include disjoint same-owner
joins, destination overlays, exact inverse fragment restoration, same-key
conflicts, stale image fallback, hostile fragment/wire limits and atomic failure.

This checkpoint completes the independent required native selector/order API
migration while the metadata representation remains pending. The full goal and
all-format/performance acceptance remain open.


## Commit and cleanup

The rewritten base remains an ancestor of HEAD; no rebase replay is necessary.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. Both required staged size/binary
guards printed nothing, and the changed/untracked artifact audit found no
forbidden files.

After every verification process terminated, `/tmp/litchi-unified-0130-*`
scratch, generated root `Cargo.lock`, `target` and Python cache were removed.
No temporary fixture links or worktrees were created. Inherited changes, tracked
tool lockfiles and existing fixture directories were preserved. Delegated review
was read-only. The full goal remains active.
