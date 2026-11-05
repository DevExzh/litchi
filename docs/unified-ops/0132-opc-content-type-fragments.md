# 0132 — Guarded OPC content-type fragments

## Implementation

`OwnedContentTypes` now captures bounded lexical Default/Override catalogs and
plans guarded mapping edits. `ContentTypeFragmentKey` separates extension and
part-name identities. Lookup follows the native content-type map: XML attribute
normalization followed by ASCII case folding, without additional whitespace
collapse. Returned keys are normalized; retained element bytes preserve original
case, escaping, attribute order, quote style, paired tags and foreign markup.
Duplicate native keys and nonstandard lexical roots are refused.

Both metadata families now use one private scanner and splice planner. Their
public tokens and key types remain separate. Content-type plans return native
content-type bindings; relationship plans retain relationship-owner and physical
member-presence checks. There is no public constructor for a replacement token
from unchecked XML.

The shared planner guards exact touched elements and inherited XML context,
retains destination slots for replacements, and proves insertion positions from
neighbor identities and retained gaps. It preserves unrelated destination XML
and restores deleted suffixes ahead of unrelated appended entries. Unprovable,
cyclic, stale or overlapping anchors remain conservative refusals. Changed root
namespace bindings or inherited `xml:base`, `xml:lang` or `xml:space` require a
conflict resolution; no prefix rewriting is attempted.

Edit counts are admitted before adapter allocations. Output byte length and
mapping count are exact and checked before replacement-buffer allocation; XML
events use a conservative upper bound. Materialization rechecks limits and
validates both native semantics and source XML. Exact no-ops retain shared source
bytes while still enforcing tightened structural limits. Package-wide aggregate
reservations, graph dependency closure and final package installation audits
remain caller responsibilities. No performance acceptance claim is made.

A regression also demonstrated that the older `ContentTypesEditPlan` returned
its no-op token without enforcing tightened XML depth/attribute limits at
materialization. That path now runs raw XML preflight and native binding parsing
before returning the shared token. Existing lenient source semantics are
preserved; it does not require the stricter editable-fragment catalog.

## Verification

The final full OPC run passed 1,126 tests/doctests across 47 targets, with one
existing ignore. This includes all 13 relationship-fragment regressions after
extraction and nine new content-type tests, plus the new runnable catalog example.
Coverage includes mixed Default/Override publication and reopen, inverse
restoration of all nonempty three-entry deletion subsets in either batch order,
unrelated destination overlays/appends, simultaneous replacements, native key
normalization and duplicate identity, exact no-ops, tightened structural limits,
byte/mapping/event/entry quota boundaries, stale entries/context/anchors,
self-closing aliased roots and malformed XML.

The older no-op limit regression was first observed failing, then passed after
the fix. OPC all-target Clippy passed with warnings denied. Read-only review
found no remaining checkpoint blocker. Lookup of a caller-supplied content-type
key allocates its normalized selector fallibly; this is not a claim of
allocation-free lookup.

Verification commands (offline and serialized):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc --test relationship_fragments
CARGO_NET_OFFLINE=true cargo test -p litchi-opc --test content_type_fragments
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc --no-deps
```

Native PPTX passed 967 unit tests with encryption enabled, with one existing
ignore. OPC documentation passed with warnings denied. The base remains an
ancestor of HEAD. Required staged evidence-size and binary-path guards printed
nothing; a changed/untracked artifact audit found no oversized evidence or
compiled binaries.

## Remaining scope

This closes the content-type transport prerequisite identified in 0130–0131.
Native PPTX still needs durable per-key metadata deltas, conditional source
provenance, join/three-way composition, exact-image fast paths and physical
relationship-member presence integration. The diagnostic from 0128 still
identifies an unresolved forward-replay limitation. Public slide insertion and
the full all-format API goal remain incomplete.

## Repository checkpoint

The requested commit could not be created. Staging failed with
`Unable to create '/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock': Read-only file system`.
No permission workaround was attempted. Changes remain in the worktree.
Checkpoint-owned scratch logs/scripts, the generated root lockfile and build
output were removed after all verification handles were terminal. No temporary
worktree was created, and pre-existing working changes were preserved.
