# 0969 CFB writer proof handoff design

Status: design only.  This memo contains no repository source changes, build
outputs, or performance measurements.  It is the implementation contract for
the next candidate and deliberately keeps the existing source-bound overlay
route and full-capture fallback intact.

## Decision and evidence boundary

The current HEAD is `34fc85e23ccefcd8ad6ffd4d57d9e2671ef1a5a1`.  The working
tree already has three unrelated changes and they are outside this candidate:

```text
 M tools/check_iwork_leaf_public_api.py
?? crates/litchi/examples/bench_word_parse.rs
?? docs/UNIFIED_OPS_API_DESIGN.md
```

The 0966 ADR manifest was revalidated before this design.  Its SHA-256 is
`25fc1d3085b40b5851ae4a3b12d27a80988784f2997834ede8731fc977fd7c57`; the
manifest's 32 accepted ADR hashes and the current ADR README are unchanged.
The current `docs/adr/README.md` hash is
`2af0e5814378c2765c12213e495d02c2895a27f142bc618ca1f3037824bbe028`.
The current source anchors used below have these hashes:

| source | SHA-256 |
| --- | --- |
| `crates/litchi-ole-common/src/object/editor.rs` | `ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319` |
| `crates/litchi-ole-common/src/object/codec.rs` | `99dc73a358cbe5a76c6fc242e42af6524ba6d0954c2ae508cf55f84cb6a5cafd` |
| `crates/litchi-cfb/src/writer/core.rs` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |
| `crates/litchi-cfb/src/writer/layout.rs` | `480b30fb1c6c018c3182fc5e0ab1f53bb55f61169396e2f2d8bce257c7d1460c` |
| `crates/litchi-ole-common/src/object/directory/model.rs` | `63342c1da1ed1ac79de4f029b8023e6e27a3d5ec07b356583b908403a8f4a196` |
| `crates/litchi-ole-common/src/object/directory/codec.rs` | `afd82d7568246ed6dbff202b6525c96329f5061547eadcc360f9d840e2659a0a` |
| `crates/litchi-ole-common/src/object/discovery.rs` | `830fa39ca65ef5073d14dabcc60513bca6a1f221eaaa1e28e62e9644bbeffdea` |

The 0968 diagnostic inputs are retained outside the checkout.  Their hashes
are `ranking-v1.md` `b1bac0422027de4ee0ca650b7ee7a4bb78aea10792158185d311f0a050cab13e`,
`handoff-census-design-v1.md`
`8fdb78b39b9e7a9214556d1fea3e81767e22ed8e36b9abfa5f40bc15dcd1b789`,
`cfb-commit-attribution-v1.md`
`1e612750df98818e2a94e7bd4a5e9b7294a06d11ed519add06f3ba60bb1f3c7e`, and
`review/final-review-v1.json`
`831921117deee1c8be0312db78cb912813ed8fdaa5ae340cdf06151b875bf609`.
They establish the opportunity and route census only; they do not establish a
speedup claim.

The candidate is worth one controlled experiment because the successful
source-layout writer route already proves the model payloads before emission,
then emits them into a fresh `Vec<u8>`, and the editor immediately opens that
same vector and reads every stream again.  The handoff can remove the second
stream payload materialization and the subsequent `reuse_stream_allocations`
hash map/byte comparisons.  It cannot remove writer planning, plan validation,
final CFB opening, protection admission, directory decoding, budget checks, or
discovery.  The benefit is therefore limited to writer-reuse commits that pass
the plan gates; a measurement must decide whether that remaining removable
work is material.

## Current control flow and source anchors

The exact upper seam is
`crates/litchi-ole-common/src/object/editor.rs:977-1000`:

1. `commit_candidate_with_rendered` validates the edited model at line 978.
2. With `Reuse`, it first tries `Package::render_copy_through` at lines
   979-988; if that route declines, it calls `render_with_layout` with the
   original bytes.
3. It opens the rendered bytes with `OleFile::open` and runs
   `codec::open` at lines 993-994.
4. It calls `Package::capture` at line 995, then
   `reuse_stream_allocations` at line 996, before publishing the package and
   discovery result at lines 997-999.

The render seam is
`crates/litchi-ole-common/src/object/codec.rs:639-723`:

* It creates a new `OleWriter` and a new `Cursor<Vec<u8>>` for every render.
* It adopts the optional source layout at lines 644-648, fills the writer from
  the package's immutable storage/stream model, and calls `write_to` at
  lines 720-722.
* `render_copy_through` at lines 739-823 is a separate equal-length overlay
  mechanism.  It opens and compares every candidate stream itself and has no
  writer `ReusePlan` receipt.  It remains on the current full-capture path.

The current capture work to remove is
`crates/litchi-ole-common/src/object/codec.rs:214-233` and
`1157-1221`.  `reuse_stream_allocations` builds a path index and compares all
captured bytes.  `capture_container` decodes each directory object, charges the
budget, and calls `OleFile::open_stream` for every stream at line 1200.

The writer proof boundary is
`crates/litchi-cfb/src/writer/core.rs:312-350,462-509,1631-1659` and
`crates/litchi-cfb/src/writer/layout.rs:156-183,938-983,1307-1314`:

* `OleWriter` owns a private `last_sector_layout` report, initialized to
  `None` for a fresh writer.
* `write_to` plans and validates the `ReusePlan`; only after `plan.validate`,
  `plan.emit`, and `writer.flush` succeed does it set the report and return
  `Ok(())`.
* `ReusePlan::validate` reopens the planned CFB, requires conformance, and
  compares every model stream through `StreamComparer` without materializing a
  second full output artifact.
* a declined plan or from-scratch serialization records a non-reused report.

Final physical validation and metadata sources are unchanged:

* `OleFile::open`/`open_with_limits` are at
  `crates/litchi-cfb/src/file.rs:938-965`; bounded directory traversal is
  `visit_directory_entries` at line 3103, and payload materialization is
  `open_stream` at line 3382.
* `directory::decode` and link/source validation are at
  `crates/litchi-ole-common/src/object/directory/codec.rs:9-65`.
* all physical fields that must be projected are `Metadata` at
  `crates/litchi-ole-common/src/object/directory/model.rs:176-194`:
  SID, kind, CLSID, red-black links, state bits, creation/modified FILETIMEs,
  start sector, stream size, and mini-stream choice.
* the same `Limits` and package budget are used by `Package::capture` and its
  `Budget` at `codec.rs:28-80,1101-1155`.  Discovery remains
  `discovery::from_package` at `discovery.rs:46-74`.

## Receipt decision: private upper seal is sufficient for this candidate

`last_sector_layout` alone is not a proof handoff.  It is a public, `Copy`
summary with no model identity, source identity, limits, output ownership, or
directory projection.  It can also be stale on a writer that is reused after a
previous successful serialization and then fails before updating the report.
It must never be exposed as a boolean fast path or consumed by a generic
caller.

For the current seam, a new lower-level public CFB receipt is unnecessary.  The
existing `Package::render_with_layout` invariant is strong enough to seal a
private operation-local receipt:

* it constructs a fresh writer and a fresh `Cursor<Vec<u8>>` on every call;
* the writer starts with `last_sector_layout == None`;
* the report is inspected immediately after the only `write_to` call and only
  a report whose `reused_source_layout()` is true is accepted; and
* the writer is dropped after `Cursor::into_inner`, so no later serialization
  can make the report stale.

The candidate should introduce an opaque `pub(crate)` codec handoff used only
by the editor commit path.  Its fields and constructor stay private to the
codec module.  A minimal shape is:

```rust
struct RenderedCommit<'a> {
    bytes: Vec<u8>,
    reuse: Option<ValidatedReuse<'a>>,
}

struct ValidatedReuse<'a> {
    model: &'a Package,
    source: &'a [u8],
    limits: Limits,
    report: SectorLayoutReport,
}
```

The actual names may differ, but the invariants may not.  `model` and
`source` are immutable borrows held only for this commit operation; they bind
the proof to the exact package and source slice used to construct the writer.
`limits` binds the later projection and discovery to the same resource policy.
`bytes` is the fresh writer-owned output and is never exposed mutably.  The
receipt constructor is reachable only after the successful `write_to` return,
the report is present and reused, and `source.is_some()` was the input to
`adopt_source_layout`.

The private helper should be equivalent to:

```rust
fn render_with_layout_for_commit<'a>(
    &'a self,
    source: Option<&'a [u8]>,
    policy: SectorLayoutPolicy,
    limits: Limits,
) -> Result<RenderedCommit<'a>, OleError>;
```

The existing `render_with_layout` wrapper remains available for `finish`,
snapshots, and ordinary package rendering.  It can share the writer-building
body and discard the handoff, but it must not make the commit projection
observable to those callers.  `commit_candidate_with_rendered` calls the
sealed helper only on the source-layout render after copy-through declines.
Rewrite policy, no source, a source that is merely declined by
`adopt_source_layout`, plan decline, and the copy-through route produce
`reuse: None` and retain the existing capture and `reuse_stream_allocations`
sequence.  An allocation error from source adoption still returns that
existing writer error immediately, as it does today; it is never converted
into a receipt or silently retried.

If a future API permits a long-lived `OleWriter`, an arbitrary sink, or a
caller to serialize more than once, this upper seal would no longer be enough;
then the CFB owner would need an opaque lower-level
`ValidatedPlanEmission` returned directly from the successful
`write_to` operation.  That broader API is outside 0969 and should not be
added speculatively.

## Commit handoff and projection algorithm

The editor must retain the current publication order.  A concrete flow is:

1. Run `self.package.check(self.limits)` as today.
2. Select copy-through or source-layout rendering as today.  Only the latter
   calls the sealed helper and can return `ValidatedReuse`.
3. Open the owned rendered bytes with `OleFile::open(Cursor::new(...))` and
   call `codec::open` with `max_package_directory_entries` exactly as today.
   These checks happen before any package publication and before projection.
4. When a receipt exists, run a metadata-only capture over this final
   `OleFile`.  The traversal must share the existing root decode,
   `visit_directory_entries`, path construction, depth checks, `directory::decode`,
   and `Budget` charging order.  It must not call `open_stream`.
5. Decode and retain the final root metadata.  Set the new package sector size
   from the final `OleFile`, and set its root CLSID from the final root
   metadata, matching `capture_with_budget`.
6. For every final storage, preserve traversal order and path.  Look it up in
   the immutable staged model and require one exact storage match and kind.
   Construct the new storage with the **final** metadata, including physical
   SID, links, state, timestamps, and CLSID.  The staged model is used only for
   logical path identity and the proof that this entry belonged to the emitted
   model.
7. For every final stream, require one exact model stream match and kind.  The
   final directory `stream_size` must equal the staged `Arc<[u8]>` length.  Clone
   that exact staged `Arc` into the new `Stream`; do not read the output stream
   payload.  Attach final metadata so start sector, size, mini-stream choice,
   SID, links, state, timestamps, and any CLSID are authoritative from the
   final CFB.
8. Require one-to-one coverage: no final entry may be absent from the model,
   no model storage/stream may be missing from the final traversal, and no
   path may occur twice in either side.  A path collision, unexpected kind, or
   stream-size mismatch is a private handoff decline, not permission to reuse
   an unchecked `Arc`.
9. After projection, run `discovery::from_package` with the receipt's exact
   limits.  Only after it succeeds may the editor assign `self.package`,
   `self.objects`, and `self.changed`, as it does today.

The projection may use bounded `try_reserve` path indexes.  An index
allocation failure must abandon the handoff and execute the old full
`Package::capture` plus `reuse_stream_allocations` path on the same already
validated `OleFile`; it must not introduce a new user-visible allocation
failure merely because the optimization was selected.  The index must not
become an unbounded cache or survive the operation.

The final physical placement is allowed to differ from the staged metadata:
stream size changes, mini/regular placement, sector movement, and directory
SID changes are exactly why final metadata is projected.  The model/path
identity and staged byte length are the proof constraints; the final CFB owns
all physical fields.

## Invalidation, errors, limits, and atomicity

The following conditions invalidate the receipt or route it to the existing
fallback:

* `adopt_source_layout` returns false, the policy is `Rewrite`, or the writer
  records a declined/non-reused layout;
* the source-layout plan is not `Outcome::Planned`, plan validation declines or
  errors, emission fails, or flush fails;
* the render is the source-bound equal-length overlay path;
* the final directory has an extra, missing, duplicate, or wrong-kind path, or
  a final stream size does not equal the staged `Arc` length; and
* a temporary projection index cannot be allocated within the ordinary
  bounded allocation policy.

The first four render failures retain the existing writer error or fallback
behavior.  A projection-only proof mismatch uses the old capture path, so
current output/package semantics remain the oracle.  Actual final CFB open,
protection-admission, directory decode, link validation, depth, object-count,
stream-count, stream-size, or total-byte errors are returned in the same
validation order rather than being hidden by the optimization.  The shared
metadata traversal must charge a stream's declared size before any decision
about its staged `Arc`, matching `capture_container`; a size-limit failure
therefore cannot be bypassed by skipping payload reads.

The final `OleFile::open` and `codec::open` remain mandatory even on the fast
handoff.  The writer's successful `ReusePlan::validate` is the payload proof:
it reopened the planned output, required conformance, and compared every
model stream before emission.  The final open validates the actual fresh
vector's CFB structure.  The projection validates every final directory entry
and its declared stream size.  Together these checks replace the redundant
post-emission stream materialization while preserving the existing protection,
resource, and malformed-CFB boundaries.

No state may be published until projection/fallback and discovery succeed.  If
any error occurs, the consumed editor has no externally visible partially
updated package: `self.package`, `self.objects`, and `self.changed` are assigned
only at the existing final point.  The output `Vec` is returned only after the
new package has been accepted.  The borrowed receipt is moved into the
projection helper and is dead before `self.package = parsed`, which keeps the
immutable model/source proof scoped to the render operation.

## Required tests before any measurement

Tests should compare the candidate with the current full-capture path rather
than with another implementation that shares the same projection helper.
The existing writer tests can exercise the proof; the object-owner tests need
an explicit full-capture oracle.

1. **Proof sealing and route census.**  A successful source-layout
   `ReusePlan` creates a receipt only after validation, emission, and flush.
   Plan decline, from-scratch fallback, rewrite policy, no source, a source
   adoption decline, and source-bound copy-through create no receipt and still
   call full capture.  A source-adoption allocation error must retain its
   existing error.  A writer-reuse/stale-report test must show that a report
   from a previous serialization cannot be consumed: the helper's fresh-writer
   invariant, not a bare `last_sector_layout` boolean, is the gate.
2. **Positive metadata projection.**  Cover empty, mini-stream, regular-stream,
   and mixed packages; unchanged payloads; shorter, longer, and size-class
   crossing edits; metadata/CLSID edits; and a new stream/storage that forces
   a normal fallback.  Compare output bytes, discovery object bytes, package
   paths/order, and every `Metadata` field with a full `Package::capture`
   oracle.  On the receipt route, assert `Arc::ptr_eq` for every stream that is
   reused from the staged model.
3. **Directory and payload faults.**  Exercise malformed/truncated output,
   bad red-black links, invalid entry kinds/SIDs, duplicate/path-collision
   cases, missing/extra model entries, and wrong declared stream sizes.  A
   projection mismatch must fall back to the old capture result; malformed
   CFB/decode/limit errors must preserve the existing error class and order.
   Include allocation-failure injection for the temporary index and verify
   that it returns to the ordinary capture path.
4. **Resource and protection boundaries.**  Check depth, directory-entry,
   stream-count, per-stream, total-byte, malformed-declaration, and protected
   container cases.  Verify that protected admission fails before projection,
   and that skipping `open_stream` never skips the declared-size budget charge.
5. **Failure atomicity.**  Inject writer, final-open, projection, fallback,
   and discovery failures.  The editor/package/object state must remain at the
   pre-publication point, with no partially attached metadata or `Arc` set.
6. **Differential route matrix.**  For every fixture, run the unchanged
   full-capture oracle and the candidate across copy-through, successful
   source-layout reuse, source-layout decline, rewrite, and no-source routes.
   Compare exact output bytes, error types/messages where stable, target
   object bytes, metadata, stream ordering, and limit behavior.  This matrix
   must prove the candidate does not revive the rejected source-bound 0921
   path; 0921's full recapture behavior remains unchanged.

These tests are the gate for a later workload run.  No performance claim can
be made from the 0968 census or from the existence of the receipt; the next
step after code review is a controlled measurement of the eligible writer
reuse cells, with fallback counts and stream sizes reported separately.

## Explicit non-goals

0969 does not add a public CFB writer receipt API, a global cache, an output
hash-only shortcut, a broad equal-length overlay optimization, a validation
skip, or a source-bound replay.  It removes only the duplicated post-emission
stream payload readback/materialization when the existing writer has already
proved the exact operation-local model and a fresh output vector.  All other
routes retain the current full capture and allocation reuse behavior.
