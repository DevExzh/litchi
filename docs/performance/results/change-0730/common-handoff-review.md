# 0730 common OLE handoff API and retained-output review

Status: source review plus bounded common implementation follow-up,
2026-09-22. The review portion records the read-only source and policy
findings. After the root baseline build completed, the authorized follow-up
was limited to the common editor method and its focused common tests; it did
not change Cargo metadata, native probes, or the DOC owner.

## Review boundary and rules

The review follows the accepted goal and ADR hierarchy: correctness,
preservation, and safety come before facade simplicity, measured performance,
and modularity. The 0729 constraint manifest remains the input constraint
set; no GOAL or accepted-ADR hash was changed for this review. The relevant
source documents are:

- [GOAL](../../../GOAL.md) and the accepted-ADR index
  ([ADR README](../../../adr/README.md));
- [ADR0001, priorities and API layers](../../../adr/0001-priorities-and-api-layers.md),
  [ADR0003, snapshots and atomic publication](../../../adr/0003-snapshots-edits-and-patches.md),
  [ADR0005, I/O, memory, and performance](../../../adr/0005-io-memory-and-performance.md),
  [ADR0006, validation and preservation](../../../adr/0006-validation-security-and-compatibility.md),
  [ADR0008, migration and evidence gates](../../../adr/0008-migration-and-verification.md),
  [ADR0010, facade/archive ownership](../../../adr/0010-facade-archive-ownership.md),
  [ADR0011, physical package ownership](../../../adr/0011-ooxml-physical-package-ownership.md),
  [ADR0024, current topology](../../../adr/0024-current-topology.md),
  and [ADR0026, CFB directory metadata](../../../adr/0026-ole-directory-metadata-binding.md);
- the inherited [0729 source review](../change-0729/source-review.md),
  [0729 result review](../change-0729/result-review.md), and
  [0729 constraints](../change-0729/constraints.json);
- the retained-output precedent in
  [0656 PPTX candidate budgeting](../../0656-pptx-cross-copy-candidate-budget.md).

The common API must preserve the existing default behavior, return typed
errors for work that cannot be performed, keep ordinary CRUD free of raw
implementation details, and avoid a hidden runtime or cache. A returned
candidate is an already-created operation result; keeping it beyond the
immediate operation creates an owner-side retention obligation under the
ADR0005 amendment.

## Current source map

### Editor state and clone cost

[`Editor`](../../../../crates/litchi-ole-common/src/object/editor.rs) stores
the original bytes in `Arc<Vec<u8>>`, a base and current `Package`, object
discovery state, targets, limits, changed state, and the selected
`SectorLayoutPolicy` (lines 62-71). An `Editor::clone` therefore shares the
original byte allocation and stream payload `Arc`s, while cloning package,
object, target/path, and other model metadata. It is a state/model clone, not
a second full copy of every stream. The first effective batch change pays for
that clone; equal replacements should not.

`Package::put_stream` (in
[`object/codec.rs`](../../../../crates/litchi-ole-common/src/object/codec.rs),
lines 121-144) checks the per-stream limit, verifies the path, preserves the
stream directory metadata, installs the replacement, and checks the package
again. The replacement `Arc` is used by the writer. After a successful
candidate render, `reuse_stream_allocations` (lines 160-178) restores the
previous `Arc` for byte-identical streams, which is the relevant allocation
sharing guarantee for untouched data.

### Existing single and batched paths

[`put_stream_shared_with_rendered`](../../../../crates/litchi-ole-common/src/object/editor.rs#L342)
is the existing narrow handoff. An equal current value takes its existing
`self.clone().finish()` path; a different value clones a candidate, mutates
it, and calls `commit_candidate_with_rendered`. Its returned bytes are checked
by reopening the CFB, recapturing the package, and rediscovering objects.
That equal-value behavior should remain unchanged when the batch API is
added.

[`put_streams_shared`](../../../../crates/litchi-ole-common/src/object/editor.rs#L372)
already provides the required batch state machine:

1. It walks replacements in iterator order and checks the stream-count bound
   before processing each item.
2. It compares against the current candidate, if one exists, otherwise the
   original editor. Equal values are skipped, so repeated paths retain the
   current last-effective-value behavior.
3. It lazily clones one candidate on the first effective difference and
   applies every replacement to that candidate.
4. It calls `commit_candidate` exactly once after all replacements.
5. It assigns the candidate only after that call succeeds. Empty and
   all-effective-no-op batches return without cloning or rendering. Any
   error leaves the outer editor unchanged.

`commit_candidate_with_rendered` (lines 689-712) checks the candidate,
renders with the selected layout, reopens the rendered bytes, runs CFB codec
validation, recaptures the package, reuses equal stream allocations, and
rediscovers objects before returning `(candidate, rendered)`. The rendered
`Vec<u8>` is the exact already-validated output. It is not stored in the
common `Editor`; `commit_candidate` currently calls this helper and drops the
`Vec`.

The minimal addition should share this loop and helper. Ordinary batch
publication can map the new result to `()` and thereby continue to drop the
rendered allocation as it does today.

### Limits and layout

[`object/model.rs`](../../../../crates/litchi-ole-common/src/object/model.rs#L13)
has finite limits for object count, storage depth, streams, per-stream bytes,
selected-object output, and aggregate logical stream bytes. In particular,
the defaults are 65,536 streams, 128 MiB per stream, 256 MiB per selected
object, and 512 MiB of aggregate stream bytes. `max_stream_size`,
`max_total_size`, and `max_object_size` are validation and input/output
safety limits; none is a retained whole-package output budget and none should
be repurposed as one.

`set_sector_layout_policy` in `editor.rs` (lines 631-647) only assigns the
`Reuse` or `Rewrite` enum. It does not increment a generation or invalidate
previously returned bytes. `Reuse` attempts source-layout/copy-through
preservation and may fall back to a full writer; `Rewrite` deliberately
serializes from scratch. Both preserve the logical model, but physical bytes,
sector placement, and output size can differ. A handoff `Vec` is consequently
fresh only for the state and policy under which it was returned.

The writer uses shared stream payloads (`create_stream_shared`) and emits a
separate complete output `Vec`. Source-layout adoption and copy-through are
optimizations with ordinary fallback, not new refusal cases. The common
handoff must use the current policy and must not change policy to improve a
candidate.

## Minimal batched rendered handoff

The proposed common method is a sibling of the existing batch method:

```rust
pub fn put_streams_shared_with_rendered<'a>(
    &mut self,
    replacements: impl IntoIterator<Item = (&'a [String], Arc<[u8]>)>,
) -> Result<Option<Vec<u8>>, OleError>
```

The exact iterator bounds should follow the existing method signature. The
important contract is:

- Use the existing order, count check, equality check, repeated-path rule,
  lazy one-candidate clone, `Package::put_stream` checks, and one final
  `commit_candidate_with_rendered` call.
- Empty or all-effective-no-op input returns `Ok(None)`. It performs no
  candidate clone or render, preserves `changed` and stream `Arc` identity,
  and has the same state behavior as `put_streams_shared`.
- “All-no-op” is evaluated against the current candidate at each iterator
  step. Thus a repeated path that changes from A to B and then back to A has
  two effective changes and renders, even though its final logical bytes
  equal the starting bytes.
- A batch with an effective change returns `Ok(Some(rendered))` only after the
  candidate has passed package checks, rendering, CFB reopen, recapture,
  allocation reuse, and object discovery. The candidate is installed before
  returning, and `rendered` is the exact validated output for that installed
  state.
- A repeated path continues to use the last effective supplied value. A
  later duplicate equal to the current candidate is skipped, preserving the
  existing semantics rather than adding a second mutation.
- A count, missing-path, size, package check, render, reopen, recapture,
  discovery, or allocation error returns `Err` with no `Vec`. Because all
  work occurs in the local candidate, the outer editor remains unchanged in
  bytes, changed state, package model, and stream allocations.
- `put_streams_shared` can delegate to this method and use `.map(|_| ())`.
  That keeps ordinary batch behavior and failure atomicity while dropping the
  returned `Vec` exactly as `commit_candidate` does now.
- Do not change the existing single-stream method's equal-value behavior.
  The common layer should add no cache, token, generation, or retained output
  field. The caller must consume `Some(rendered)` immediately or assume the
  owner-side freshness and retention obligations.

This is deliberately a handoff of an existing allocation, not a second
serialization. The common layer cannot decide an owner’s public retention
policy from `Limits` and should not add a new refusal or budget error to this
API.

## Retained-output budget and fallback

If DOC consumes `Some(rendered)` immediately to construct its public result,
the bytes are operation-local. The existing final public validation remains
required: DOC's `Snapshot::open_bounded` (and its profiled equivalent) must
still validate the candidate independently before publication.

If a public or staged DOC value keeps the bytes between operations, its owner
needs an explicit finite policy. A suitable shape is an owner-side
`TransactionLimits` member such as
`max_retained_candidate_bytes` (or the clearer
`max_retained_output_bytes`), with:

- actual held bytes charged as `rendered.len()`;
- an accessor such as `retained_output_bytes() -> Option<usize>` or an
  equivalent observable weight;
- an explicit release operation that drops the same `Vec` allocation; and
- policy intersection by `min` whenever multiple operation policies meet.

The retained value must be the `Vec` returned by the common operation moved
into the owner. It must not be cloned, serialized to scratch, or represented
by a retained parsed snapshot. The 0656 PPTX document uses a 64 MiB default,
but that is a precedent rather than a DOC default; a DOC value requires
measured justification and must remain finite. A very small positive ceiling
can disable retention in practice, but zero must not accidentally convert a
supported edit into a typed refusal unless the owner explicitly defines zero
as “hold nothing.”

When `rendered.len()` exceeds the owner’s ceiling, silently drop the returned
`Vec`, retain nothing, and recompute with the existing `finish` path when the
owner later needs bytes. This may render twice, but it preserves result bytes,
supported edits, and existing typed failures. It follows the accepted ADR0005
rule that an over-ceiling retained allocation falls back to recomputation when
the work is available. The common method itself must still return `Some` for
the successful operation; the owner chooses whether it can retain it.

Do not carry a `Vec` across a layout-policy setter or another mutation without
an owner-side identity/freshness contract. Since the common setter has no
generation, a retained handoff would need to bind state identity, policy,
limits, and replacement set, then release it on any invalidating mutation.
The safer minimal integration is immediate one-shot consumption, with any
longer-lived retained result explicitly owned and budgeted by DOC.

## Proposed tests and file ownership

Once the root baseline build is available, production ownership should remain
limited to:

- `crates/litchi-ole-common/src/object/editor.rs`: the sibling method and the
  small shared batch/helper refactor;
- `crates/litchi-ole-common/tests/object.rs` (or the existing focused common
  test module): common API contract tests only.

Do not edit the DOC owner, CFB writer, Cargo files, or native probes for this
common change. The test set should cover:

1. A two-stream effective batch returns `Some`, reopens and validates, and
   has the same logical result as the ordinary batch route. Untouched stream
   payloads retain expected `Arc` sharing after recapture.
2. Empty and all-equal batches return `None`, do not change `changed`, and do
   not replace stream `Arc`s. An all-no-op call on an already changed editor
   still returns `None`; the caller can obtain current bytes by recomputing
   with `finish`.
3. Repeated paths retain last-effective-value semantics, including a later
   duplicate equal to the current candidate, and match the ordinary batch
   output.
4. Missing paths, stream-size and stream-count violations, package/render
   failures, and reopen/recapture/discovery failures are atomic:
   the original bytes, changed state, package model, and stream allocations
   remain unchanged, with no returned output.
5. Both `Reuse` and `Rewrite` are exercised for same-length and
   length-changing replacements. Tests assert logical inventory, metadata,
   CLSIDs, and successful validation; they should not require identical raw
   physical layouts across policies.
6. Existing single-stream handoff no-op behavior remains unchanged. Any test
   for stale retained output after a policy setter belongs to the DOC owner,
   because the common API returns an immediate handoff and does not retain it.
7. Ordinary `put_streams_shared` remains behaviorally identical after it
   delegates and maps away the rendered result. A no-op test should establish
   that no candidate render occurs where the test harness can observe it; do
   not infer allocation counts without an allocation-aware test lane.

The existing common object tests around the ordinary batch atomicity,
all-no-op behavior, and `Arc` sharing are the natural extension point. These
tests are required because the new public return value and atomicity contract
must be observable; they are not implementation-mirroring coverage.

## Disposition

The bounded common implementation is now present in
`crates/litchi-ole-common/src/object/editor.rs`, with focused contract tests
in `crates/litchi-ole-common/tests/object.rs`. The common object test command
passes all 21 tests. The implementation adds one rendered batch handoff,
preserves no-op and failure semantics, leaves retention policy with DOC, and
keeps final DOC validation independent. The report's retention design remains
an owner-side recommendation; it does not add a common retention cache or a
new refusal.
