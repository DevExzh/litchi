# Change 0039: source-backed DOCX semantic publication

Date: 2026-08-11

Production base: `627e4a4fb35b271576ab4b056acf49a783905f8e`

Scope: the DOCX main-document transaction facade over the accepted low-level
source-backed OPC one-Part publisher. OLE2, RTF, ODF, iWork and IWA production
code are unchanged.

## Problem and change

The source-backed DOCX facade could query a lazily loaded main document, but a
semantic edit still required converting all Parts into an eager `OpcPackage`.
On the fixed media-rich document this inflated 17 ordinary Parts, including
eight incompressible 2 MiB images, before the general DOCX writer Deflated the
whole package again.

The facade now exposes `document_snapshot`, `edit_document`, and the consuming
`publish_document_commit_to_stream`. A transaction shares the raw main-Part
allocation with its initial snapshot, applies the commit's exact patch against
freshly version-checked source bytes, and sends only the resulting main XML to
the existing source-backed OPC overlay publisher. Every unselected local ZIP
member remains on the physical raw-copy path. Content-free cache diagnostics
make the one required semantic payload materialization observable.

The path is deliberately narrower than general package publication:

- markup-compatibility preprocessing must borrow the raw document unchanged;
  any branch selection or rewrite is refused before output, so selectors and
  published offsets always describe the same bytes;
- ordinary paragraph/run/field/content-control/table text operations and plain
  paragraph insertion/removal are confined to the main Part;
- transferred paragraphs are refused even with an empty transfer graph because
  their relationship-dependency digest requires package-level validation;
- the patch must match the freshly loaded raw main document exactly;
- changed signed sources retain the low-level typed refusal, while exact signed
  no-ops copy every source byte; and
- no topology, content-type, relationship, encryption, resigning, filesystem
  atomicity, runtime, global cache, unsafe code, or dependency change is added.

## Matched media-rich measurement

The opt-in `docx_source_backed_one_edit_save` case fixes 200 semantic
paragraphs, eight deterministic incompressible 2 MiB PNG payloads, 17 ordinary
OPC Parts, and 20 physical ZIP members. The 16,793,036-byte source contains
16,833,643 logical Part bytes and has SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
The timed interval opens the positional package, edits paragraph 100, commits,
and publishes to a bounded sequential sink. Untimed verification performs a
complete DOCX reopen, checks all 200 paragraphs, every Part/content type and
relationship, all eight exact media payloads, package topology, deterministic
output, patch replay, inverse restoration, and stale rejection.

Release binaries were frozen independently. Their SHA-256 values are
`b0c28ad2c1ddb308662aabea7602fc2d02d5fb83b841ee9ddad57ae42b74651e`
before and
`ca8aa608743425bc0c1df0648c256130520baee3775961f873092c9c1df8c3db`
after. CPU-2 ABBA runs used 10 warmups and 100 samples per leg, yielding 200
pooled samples per state.

| Source-backed DOCX one-edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 223.183 ms | 5.732 ms | **-97.43%** |
| mean | 223.977 ms | 5.797 ms | **-97.41%** |
| p95 | 230.829 ms | 6.298 ms | **-97.27%** |
| p99 | 235.578 ms | 7.164 ms | **-96.96%** |

Both legs improve independently: before/after p50 is 222.589/5.850 ms in A
and 223.700/5.676 ms in B. The approximate independent-sample 95% interval for
the mean delta is `[-97.65%, -97.17%]` of the before mean.

Both binaries produce the identical 16,793,048-byte output SHA-256
`9af99bf7f63aac1ffc13ff59a5038703c96229b906b3db1602d5af5590171795`.
Semantic Part materializations fall **17 -> 1**. Unavoidable raw-copy overlap
is unchanged at 529 reads / 16,789,483 compressed payload bytes. Total source
reads move 615 -> 596; total read bytes move 16,792,697 -> 16,797,356 because
the optimized path performs bounded raw structural checks rather than eager
logical ownership. Sink calls fall 651 -> 553 at the same 32 KiB maximum. No
physical input/output byte reduction is claimed.

Raw ABBA records are [`before A`](../results/abba-docx-source-edit-before-a.json),
[`after A`](../results/abba-docx-source-edit-after-a.json),
[`after B`](../results/abba-docx-source-edit-after-b.json), and
[`before B`](../results/abba-docx-source-edit-before-b.json). The frozen binary
and evidence hashes are indexed by
[`docx-source-edit-sha256.txt`](../results/docx-source-edit-sha256.txt).

## CPU, allocation, and regression attribution

Matched three-repeat `perf stat` processes use ten measured iterations and
include corpus construction plus untimed reopen/verification. Even under that
conservative whole-process scope, task clock falls 69.06%, cycles 68.95%,
instructions 74.91%, branches 79.75%, branch misses 90.24%, cache references
67.89%, and cache misses 40.61%. The removed work is the eager inflation,
ownership, and re-Deflation of the 16 MiB unselected media closure.

Ten-sample Heaptrack processes move 352,388 -> 337,965 allocation calls
(-4.09%) and 43,990 -> 39,751 temporary allocations (-9.64%). Peak heap is
flat at 189.42 -> 189.39 MiB because corpus construction, expected output, and
untimed complete verification coexist in the process. Heaptrack RSS is
165.12 -> 169.73 MiB (+2.79%); uninstrumented GNU Time ABBA RSS ranges overlap
at 159,740/175,996 KiB before and 159,932/157,960 KiB after, so no RSS
improvement or regression is claimed.

The unchanged eager `docx_semantic_one_edit_save` medium guard used 20 warmups
and 200 samples per ABBA leg. Its pooled p50 is 542.508 -> 543.860 microseconds
(+0.25%), mean +1.89%, and p95 +3.27%; the legacy constructor was restored to
its original implementation after an initial candidate exceeded the guard.

Counter, Heaptrack, GNU Time, and guard records are retained under `results/`
with the `docx-source-edit` prefix.

## Preservation and correctness gates

- Source-backed facade tests cover cold open, shared raw snapshot ingress,
  changed and byte-exact no-op publication, full DOCX reopen, strict and
  transitional main-document relationships, retained unknown main XML, exact
  opaque compressed payloads, stale commit/source-version refusal, MCE refusal,
  signed change/no-op policy, bounded limits, and typed partial sinks.
- The operation classifier is exhaustive inside the crate, so a future
  operation variant must make an explicit source-backed safety decision.
- Existing low-level OPC tests continue to prove complete untouched local and
  central-record identity, unknown non-Part preservation, unsupported-layout
  refusal, XML auditing, source monitoring, and bounded sequential writes.
- All 843 DOCX library tests, all integration tests, 74 passing doctests, and
  the 29-test standalone harness suite pass. Warning-denied all-feature DOCX
  library Clippy and warning-denied all-target harness Clippy pass. The broader
  DOCX all-target Clippy job remains red on pre-existing test-only lints outside
  this change.
- CI adds exact one-sample smoke and 15-sample release gates for corpus/output
  hashes, topology counts, one materialization, complete sink bytes, and bounded
  writes.

## Remaining work

- MCE-normalized documents and paragraph transfers require a package-aware
  semantic publication design; this path refuses them instead of guessing.
- The source-backed facade still supports one committed main-Part replacement,
  not arbitrary multi-Part CRUD, resource addition/removal, signatures,
  encrypted OOXML, atomic filesystem replacement, or parallel publication.
- The selected original main Part and changed main XML still require bounded
  logical allocations, and every physical archive byte must still cross the
  positional source and sequential sink.
- Equivalent guarded publishers for XLSX and PPTX, plus real-producer and
  adversarial media/security matrices, remain separate work.
