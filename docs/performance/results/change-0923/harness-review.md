# 0923 — frozen DOCX ingress harness review

Reviewed 2026-10-06T08:49:03Z. The reviewed harness delta is the frozen external
file `/tmp/litchi-performance-0923/harness-0923.patch`, 17,046 bytes,
SHA-256 `5404cf09f7f840aa3a587d1bbdce6659d0d3a8cf35d2c44bbd42b747ddf217e0`.
This review did not build, run, or modify the harness or production source.

## Route and phase disposition

The generic `ReadAt` lane keeps the historical six phase boundaries. Source
observer creation, route setup, and sink reservation remain in preparation;
package construction starts the source-backed open clock; query, edit, commit,
and sequential publication retain their prior clocks. The owned lanes add only
route-specific setup before open and use the same downstream lifecycle and
correctness protocol.

`owned_vec` clones the corpus bytes in preparation and moves that vector through
`SourceBackedPackage::from_vec`. `owned_reader` moves a clone into the
instrumented `Cursor`; `Package::from_reader` consumes it through the bounded
reader path and delegates to `from_vec_with_limits`, so both owned lanes retain
an immutable source owner. After output reopen, the harness applies the
publication inverse through that same package's retained `SourceArtifact` and
requires the resulting complete byte vector to equal the corpus. This is the
owned route's direct immutability oracle: no external alias remains after the
move, and a changed retained source cannot satisfy the exact inverse result.
Generic input additionally compares its external source vector directly.

The route fields describe unavailable observations truthfully. Owned-vector
sequential fields currently contain zero placeholders because there is no
reader observer; owned-reader counters measure successful reader ingress, and
positional `ReadAt` vectors are empty for both owned lanes. The explicit
observation-scope string is therefore required when consuming the zero
placeholders; none may be interpreted as zero positional reads. Owned input
bytes are reported separately.

## Gate limitation requiring explicit labeling

`verify_docx_section_layout_controls` still opens its control cases only with
the generic `ReadAt` constructor. Nevertheless,
`stale_foreign_signed_noop_limits_partial_sink_verified` is serialized as true
for every selected ingress route. The normal owned lifecycle covers snapshot,
patch, inverse, output, cache, source-retention, and sink behavior, but it does
not independently run the stale/foreign/signed/no-op/limit/partial-sink negative
cases. The combined result must therefore either parameterize those controls
by ingress route or label this field `generic_route_controls_verified` (or
otherwise state that owned lanes inherit a generic control oracle). It must not
claim owned-lane negative-gate coverage from the current flag.

The expected-package and semantic reopen helpers are generic oracle paths. That
is acceptable for output comparison when the selected route's own inverse and
publication checks also run, but the evidence should retain this distinction.

## Memo rule applied to the later source review

ADR 0005's amendment, ADR 0032, and owner decision 0758 permit a small derived
value memo only when it is bound to an owner-held allocation by a strong
reference, misses recompute, failures are not memoized, all adoption points
re-project onto the current owner's allocations, and the memo's actual resident
layout is bounded and charged through the accepted reservation rule. Decision
0758 explicitly permits a facade with an empty memo slot to fill it from a
capture of its unchanged graph; later capture/publication adoption replaces the
separate facade memo with a fresh projection.

A `OnceLock` attached to an immutable owned-source handle is consistent with
that rule. The accepted retained-archive precedent creates the source and an
`Arc<OnceLock<Digest>>` together, fills the cell once from exactly that source,
shares it only with clones of the same owner, and re-derives the value in debug
and test builds. That one-time initialization is not in-place mutation of a
published memo entry. A source candidate that also needs a typed fallible cell
reservation may keep a one-entry `Vec<OnceLock<Digest>>` directly in the outer
`Arc` memo: the `Vec` allocation then is the reserved cell storage, while the
outer `Arc` control-block/object overhead is stated separately, as in the
accepted small memo records. It must not reserve an unrelated disposable token
vector or hide a mutable table inside an entry. There must be no memo inherited
by a different allocation/package and no memoized failure. A facade-level memo
must be rebuilt/projected and replaced as a whole at publication or mutation
boundaries; a source-level one-shot digest cell does not itself satisfy that
projection requirement.

The later source candidate still needs to prove its actual retained entry layout
(including table/header/Arc costs where applicable), reservation and refusal
ordering, and generic/managed-route exclusion. The harness review does not
approve a production implementation or a generic fingerprint optimization.
