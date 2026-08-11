# Change 0044: source-backed PPTX semantic publication

Date: 2026-08-11

Production base: `7c67c0f160cb8137e35ef26d4e77c39d629d797b`

Scope: one exact-source PPTX selected-slide transaction over the accepted
low-level source-backed OPC one-Part publisher. OLE2, RTF, ODF, iWork and IWA
production code are unchanged.

## Problem and change

The source-backed PPTX facade could enumerate and query lazily loaded slides,
but a semantic edit still converted the complete package into an eager
`OpcPackage`. The frozen control inflated all 229 ordinary Parts, owned a
complete eager package, then regenerated the archive after changing one text
box. On the fixed corpus that included 199 unselected slides and eight
incompressible 2 MiB PNG media Parts.

`SourceBackedPresentationEditor` now owns the source-backed package without a
`Clone` escape hatch. It exposes a selected-slide snapshot/edit/patch/commit
contract using PPTX-owned types, then consumes a commit into the existing OPC
one-Part overlay publisher. The selected raw slide XML is rewritten and
validated; every unselected physical ZIP member stays on the raw-copy path.
The mandatory presentation root and selected slide are the only two ordinary
Parts materialized by the measured operation.

The snapshot binds the exact source closure needed to interpret the slide:
the sorted package-root relationships, raw main-presentation URI/XML/content
type/relationships, selected slide reference and relationship, and raw
slide URI/XML/content type/relationships. Publication recomputes and resolves
that closure against the current source version before applying the patch.
The patch is source-specific, supports forward replay and inverse restoration,
and refuses foreign or stale input.

The path is deliberately narrow:

- exactly one shape-text operation is accepted per changed edit;
- markup-compatibility preprocessing must leave the selected raw slide bytes
  unchanged, otherwise the editor refuses before output;
- the package/presentation/slide relationship closure, URI, content type and
  topology must remain exact;
- changed signed sources retain the low-level typed zero-output refusal, while
  exact signed no-ops copy the complete source byte-for-byte;
- source versions, limits and partial sequential sinks retain their typed
  error contracts; and
- no slide add/remove/reorder, dependency transfer, multi-Part edit,
  signature rewrite, encryption, filesystem atomicity, runtime, global cache,
  unsafe code or dependency change is added.

## Matched media-rich measurement

The opt-in `pptx_source_backed_one_edit_save` case fixes 200 slides with eight
text boxes each, eight deterministic incompressible 2 MiB inert PNG Parts, 229
ordinary OPC Parts, and 445 physical ZIP members. The 17,017,139-byte source
contains 17,568,429 uncompressed logical Part bytes and has SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The timed interval opens the positional package, snapshots slide 100, changes
shape 0, commits, and publishes directly to a bounded sequential sink. Untimed
verification fully reopens the PPTX, checks all 200 slides and 1,600 text
boxes, every Part/content type/relationship, all eight exact media payloads,
package topology, deterministic output, patch replay, inverse restoration, and
stale rejection.

Release binaries were frozen independently. Their SHA-256 values are
`64bd5b3e4fe24abaeb44c8442edd06557055e5952859dc8cfeb815cb8e6fab98`
before and
`835910a8c38d7d26381cdbfe43d39437b51dd52edda2b2e538afec8612a8d9e8`
after. CPU-11 ABBA runs used 10 warmups and 100 samples per leg, yielding 200
pooled samples per state.

| Source-backed PPTX one-edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 296.590 ms | 8.545 ms | **-97.12% (34.71x)** |
| mean | 297.324 ms | 8.433 ms | **-97.16% (35.26x)** |
| p95 | 307.578 ms | 9.730 ms | **-96.84%** |
| p99 | 313.919 ms | 11.039 ms | **-96.48%** |

Both legs improve independently: before/after p50 is 299.419/8.561 ms in A
and 294.277/8.535 ms in B. A deterministic 10,000-resample independent-sample
bootstrap gives a 95% interval of `[-97.13%, -97.10%]` for the p50 delta and
`[-97.21%, -97.11%]` for the mean delta.

Both binaries produce the identical 17,017,144-byte output SHA-256
`bcfd05dc7590051137db64810db2d822ac814417c03c50d8a7fc33f429e50d61`.
Semantic Part materializations fall **229 -> 2**. Ordinary payload reads move
741 / 16,907,750 bytes -> 742 / 16,909,038 bytes; total source reads move
2,740 / 17,004,461 bytes -> 2,941 / 17,151,171 bytes because the optimized
path freshly revalidates the semantic closure and performs bounded physical
copying. Sink calls move 260 -> 1,403 while the maximum write falls 64 -> 32
KiB. Complete physical archive input/output remains, so no I/O-volume reduction
is claimed.

Raw ABBA records are [`before A`](../results/abba-pptx-source-edit-before-a.json),
[`after A`](../results/abba-pptx-source-edit-after-a.json),
[`after B`](../results/abba-pptx-source-edit-after-b.json), and
[`before B`](../results/abba-pptx-source-edit-before-b.json). The frozen binary
and evidence hashes are indexed by
[`pptx-source-edit-sha256.txt`](../results/pptx-source-edit-sha256.txt).

## CPU, allocation, and regression attribution

Matched three-repeat `perf stat` processes use ten measured iterations and
include corpus construction plus untimed reopen/verification. Task clock falls
66.76%, cycles 67.26%, instructions 67.91%, branches 68.99%, branch misses
87.99%, cache references 68.27%, and cache misses 39.87%. The removed work is
the eager logical ownership and re-Deflation of 227 unselected Parts.

Single-sample Heaptrack processes move 2,709,030 -> 1,938,513 allocation calls
(-28.44%), 462,895 -> 196,378 temporary allocations (-57.58%), and 175.13 ->
159.49 MiB peak heap (-8.93%). Heaptrack RSS is flat at 154.52 -> 154.75 MiB.
Uninstrumented ten-sample GNU Time processes are likewise flat at 146,868 ->
147,420 KiB maximum RSS (+0.38%), so no process-RSS improvement or regression
is claimed. The complete profile summary is
[`pptx-source-edit-profile.txt`](../results/pptx-source-edit-profile.txt).

The unchanged eager `pptx_semantic_one_edit_save` medium guard used 20 warmups
and 200 samples per ABBA leg. Its pooled p50 is 2.720135 -> 2.719669 ms
(-0.02%; bootstrap interval `[-0.46%, +0.44%]`), mean -0.71%, and p95 -2.30%.
This is neutral rather than a separate speedup claim. Raw guard reports use the
`abba-pptx-source-edit-guard` prefix under `results/`.

## Preservation and correctness gates

- Source-backed editor tests cover changed publication and full reopen,
  target-only logical-Part materialization, exact untouched shape and Part
  identity, exact no-op bytes, changed/exact-no-op signed sources, foreign and
  stale patch/source-version refusal, MCE refusal, strict relationship input,
  bounded limits, and typed partial sinks.
- The non-clone owning editor and consuming publication call prevent retaining
  a second mutable publication owner. The one-operation edit guard makes a
  future broader operation set an explicit design decision.
- Existing low-level OPC tests continue to prove complete untouched local and
  central-record identity, unknown-member preservation, unsupported-layout
  refusal, XML auditing, source monitoring, and bounded sequential writes.
- Complete all-feature/all-target PPTX tests, warning-denied Clippy and rustdoc
  gates pass, as do the complete standalone harness test and warning-denied
  Clippy gates.
- CI adds exact one-sample smoke and 15-sample release gates for corpus/output
  hashes, topology counts, two materializations, complete sink bytes, and
  bounded writes.

## Remaining work

- MCE-normalized slides require selectors tied to the rewritten representation;
  this path refuses them instead of publishing offsets derived from other bytes.
- The editor supports one shape-text change in one existing slide, not
  multi-operation or multi-Part CRUD, slide/resource addition/removal/reorder,
  dependency transfer, charts, themes, notes, signatures, encryption, or
  atomic filesystem replacement.
- The presentation root and selected original/changed slide still require
  bounded logical allocations, and every physical archive byte still crosses
  the positional source and sequential sink.
- Equivalent guarded publication for XLSX, plus real-producer and broader
  adversarial media/security/topology matrices, remains separate work.
