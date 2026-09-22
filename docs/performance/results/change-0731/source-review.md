# 0731 public PPT source map

Root completed this bounded source map against the exact `build.json` source
census. The separately assigned source-audit task was stopped before it supplied
an artifact; this document is not an independent source review. Independent
measurement/interpretation review is in `review.md`. No production edit follows.

The measured route is `slide_order`, not the common OLE editor or a text-box
transaction. The probe's public function opens `slide_order::Snapshot`, starts
an edit, removes position 1, commits, copies returned bytes, then drops the local
snapshot/commit owners. Oracle construction uses that same public function but
bypasses the collection wrapper. Allocator measurements likewise use the
existing public function and its original ownership region.

`crates/litchi-ppt/src/slide_order.rs` establishes these boundaries:

- `Snapshot::from_bytes_with_limits` (line 409) opens a limited package, resolves
  its presentation/directory and live Document persist identity, parses the
  document structure, compares slide identities/order, reads reviewer history,
  drops reader owners, and retains the source bytes in an `Arc`.
- `Snapshot::edit` (line 662) checks editability, shares source/working snapshots
  and creates the isolated document transaction and change collections.
- `Transaction::remove_slide` (line 1817) applies the guarded structural removal.
- `Transaction::commit` (line 1910) first commits document structure. For a real
  change it captures the original persisted slide payloads, opens an embedded
  editor, verifies the live Document still matches the working snapshot, installs
  the changed Document record, and calls the embedded writer's finish.
- After finish, commit validates unrelated streams, fully reopens a public
  snapshot, compares live slide order and persisted survivor bytes, validates
  applicable transferred pictures, and constructs the patch. Two artifact hashes
  bind the before/after structural artifacts (lines 2014–2015). `artifact_hash`
  (line 2737) uses the existing `BlobId::of` digest and hexadecimal representation.

The writer lives in
`crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs`. It bounds
projected record lengths, appends staged records, a persist directory and a
UserEdit record, updates Current User, writes via `OleWriter` using the retained
layout policy/source layout, and validates the rewritten persist mapping. Its
`validate_rewrite` returns the same rendered bytes; it does not perform the DOC
pattern of discarding a common validated render and then rendering it again.
The outer public validation serves a different semantic boundary and remains
mandatory.

A native observer experiment should keep these existing owner lifetimes and
failure paths unchanged. Useful sequential commit spans are document-structure
commit, before-payload capture, embedded open/stage, embedded finish,
unrelated-stream preservation, public reopen, after-payload capture/comparison,
and patch hashing/construction. Empty-observer versus ordinary controls must
measure any dispatch/lifetime difference before phase clocks are interpreted.
Never sum nested writer subspans with their enclosing finish span. Readback and
hashing are not nominated for removal by this source map or the simulated Ir
fractions. Exact no-ops, source-sharing, strict limits, refusal atomicity and the
existing Reuse policy remain constraints.
