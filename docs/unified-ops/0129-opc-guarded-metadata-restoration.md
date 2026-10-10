# 0129 — Guarded OPC metadata restoration

## Native boundary

Added `OpcPackage::try_replace_content_types_bytes_with_limits` and
`try_replace_relationships_bytes_with_limits` for exact-source restoration from
untrusted durable metadata images. Both verify the expected current publication
view before decoding replacement bytes, retain exact no-ops, and refuse signed
changes before parsing. They use native bounded XML/metadata parsers and the
existing publication audit before mutation. Replacement XML stays shared through
`Arc<Vec<u8>>`; no public constructor grants source authority to raw bytes.

Relationship replacement retains physical member presence separately from its
semantic edges. A package-root member cannot be absent. An absent part member
must have no edges and use its canonical empty view; lexical content cannot be
silently discarded by marking its member absent.

Review exposed an existing entry-limit omission in owned XML capture. The shared
`OwnedXmlPart` byte preflight now checks `EntryBytes` in addition to `PartBytes`
and the fixed owned-XML ceiling. Retained and authored metadata, source plans and
exact no-op restoration all respect this bound. Canonical metadata paths already
call that preflight before allocating their output buffer.

These are per-member replacement operations. The format owner must reserve
package-wide aggregate budgets, including relationship-member presence changes,
and validate the complete changed dependency closure. The source guard deliberately
rejects independently changed destination metadata. A durable structural patch
must use the existing OPC metadata edit plans for transport onto a divergent
destination, retaining its unrelated comments, aliases, declarations and edges.

## Relationship to the full goal

This supplies the native admission/install boundary for the exact-image side of
PPTX durable metadata replay and inverse. It does not change `LPTX0002`, fix the
0128 replay diagnostic, or expose public slide insertion. The next patch work
must encode bounded metadata edits/source bindings, use structural plans when
transporting, reserve aggregate resources, and cover composition, conflicts,
inverse, direct/replayed XML agreement and unrelated destination overlays.
The whole-member methods must never substitute for that transport logic.

The full goal and all-format/performance acceptance remain open. No performance
acceptance is claimed. The previous checkpoint was verified progress, not a wait
or a no-progress turn; this checkpoint adds a required native boundary.

## Verification

Six public integration regressions cover exact lexical publication/reopen and
restoration; refusal of stale source bindings without overwriting destination
overlays; malformed XML, duplicate attributes, invalid control characters and
metadata byte limits; explicit present/absent relationship members; signed
changes versus exact no-ops; and exact/one-under entry limits for retained and
authored metadata, source plans and raw no-op restoration.

The full OPC suite passed 1,103 tests and doctests across 45 targets, with one
existing ignore. The native PPTX encryption unit suite passed 967 tests, with
one existing ignore. OPC/PPTX all-target Clippy passed with warnings denied.
Public OPC documentation built with warnings denied. Read-only review found no remaining safety or atomicity blocker
for this per-member boundary after the entry-limit fix and aggregate-contract
clarification.

Verification commands (offline and serialized):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc -p litchi-pptx --features litchi-pptx/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc --no-deps
```


## Commit and cleanup

The rewritten base remains an ancestor of HEAD, so no rebase replay is needed.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. Both required staged size/binary
guards printed nothing, and the changed/untracked artifact audit found no
forbidden files.

After every test/build process terminated, `/tmp/litchi-unified-0129-*` scratch,
the generated root `Cargo.lock`, `target` and Python cache were removed. No
fixture links or temporary worktrees were created. Inherited modifications,
tracked tool lockfiles and existing fixture directories were preserved. Delegated
review was read-only. The full goal remains active.
