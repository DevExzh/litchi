# Performance optimization ADR-compliance matrix

## Change 0395 compliance update

The case-fold index remains private to the owning `litchi-opc`
`SourceBackedPackage` substrate. It adds no public raw type, dependency edge,
package identifier, archive/runtime handle, lock, unsafe code, executor path,
or CRUD signature. Exact `PackURI` lookup remains first; the index stores only
Part positions and uses an allocation-free comparator over existing names.
Source iteration order and canonical spelling are unchanged.

Correctness and bounded behavior remain ahead of selection. Package catalog
validation still rejects equivalent Part names; exact lookup, alias lookup,
misses, freshness-before-lookup, resource limits, cancellation, and existing
error precedence remain in force. Index construction is fallible. The
measured 2,048 threshold is not a semantic limit, and managed
`ExecutionContext` opens intentionally retain the old linear fallback to avoid
unreserved retained memory and non-interruptible sorting work. Mutable
`OpcPackage` behavior is untouched.

Focused boundary, unsorted-order, equivalent-name, managed-budget, and
cancellation tests passed, the full `litchi-opc` library passed `282/282`,
and exact formatting/diff checks plus independent implementation and test
reviews passed under explicit stable Rust/Cargo 1.98.1. The final CPU-2
`A1/B1/B2/A2` package used five warmups and 30 samples and binds the exact
revision, patch/source/binary hashes, corpus catalog, reports, and
adjudication in the [0395 record](changes/0395-opc-source-casefold-index.md)
and [evidence bundle](results/change-0395/). Its timing evidence is from the
normal, non-allocator unmanaged `SourceBackedPackage::from_read_at` binary;
allocator-enabled latency is observational only, and validation-constructor
coverage is correctness-only.

`performance_claim: scoped`; `claim_authorized: true`. The authorized timing
claim is limited to normal, non-allocator unmanaged packages opened through
`SourceBackedPackage::from_read_at`, with source-lookup p50 on the fixed
2,048- and 16,384-Part 144-query vectors. The initial all-size probe's 256-Part
regression and the large randomized eager-lookup drift are recorded as
rejection/withholding evidence. Validation-constructor coverage is
correctness-only, and allocator-enabled latency is observational only. No
claim follows for source-open latency, means/tails, eager, managed, mutable,
default/general OPC or OOXML behavior, RSS, physical I/O, decompression, cold
cache, throughput, or scaling. The exact allocator, retained-vector footprint,
and source-counter observations are evidence, not a broader performance
claim.

## Change 0393 compliance update

The selected-query state remains private to `litchi-pptx`, the semantic owner
of PPTX picture descriptors. It adds no raw OPC/ZIP type or identifier to the
public API, dependency edge, executor/runtime handle, cache, lock, unsafe code,
or parallel path. `images` retains its public return and allocation behavior.

Correctness remains ahead of selection: every picture is parsed and every
target resolved in scene order, including pictures after the selected one.
MCE and external-target policy, cancellation, limits, source-version fences,
payload deferral, out-of-bounds length, and malformed/missing-target error
precedence remain fail closed. Focused counters and adversarial tests establish
the mechanism independently of timing. Independent review and stable 1.98.1
validation accepted the implementation; the one full-suite exclusion was
reproduced unchanged at exact baseline.

The matched release package binds the source patch, binaries, deterministic
corpus, four raw reports, allocator vectors, and adjudication. The accepted
claim is limited to selected `image` p50 and exact selected-path allocator
reductions on that protocol. Broad latency, tails, means, RSS, physical I/O,
decompression, cold cache, throughput, scaling, and all-PPTX behavior remain
withheld. Evidence is recorded in [Change
0393](changes/0393-pptx-selected-image-query.md).

`performance_claim: scoped`; `claim_authorized: true`.

## Change 0390 compliance update

The indexed decoder session is opaque operation-scoped state: unmanaged OPC
materialization creates it only after managed refusal and passes it only through
cold loader/bypass reads. Stored entries bypass it, cache hits/waiters do not
use it, and no session or archive handle crosses an ordinary format/facade API.
Lookup, limits, reservation failure, Store/Deflate verification, accounting,
cancellation, cache rollback/publication, freshness, and managed reservation
boundaries remain unchanged.

The matched allocator evidence records one decoder session per unmanaged full
materialization, with exact reductions of 4/160,640, 510/20,481,600, and
6/240,960 allocation calls/bytes for the three deterministic corpora. Logical
reads, returned bytes, and Part counts remain invariant. The focused ZIP/OPC
tests and in-memory checked-comparison test pass under stable Rust 1.98.1.
Latency, operation-local peak/RSS, copied/decompressed/physical-I/O, and
broad/default OPC claims remain withheld. Evidence is recorded in [Change
0390](changes/0390-opc-materialization-decoder-session.md).
`performance_claim: none`; `claim_authorized: false`.

## Change 0387 compliance update

The change remains inside the `litchi-opc` package owner and reuses its existing
private shared-Part constructor. No archive type, raw physical identifier,
lock, runtime/executor handle, dependency edge, unsafe code, or public CRUD
surface is introduced. Immutable payload ownership is shared only for
unmanaged packages; existing Part mutation replaces the owning Arc and leaves
the source cache unchanged.

Managed `PartData` retains its hierarchical memory/object reservations and
still refuses bare-Arc escape before ordinary payload reads. Source freshness,
cancellation/execution, allocation failure, relationship copying, package
construction, signature policy, final validation, and typed errors remain in
the same fail-closed sequence. Focused pointer-identity/lifetime/COW tests,
279 OPC library tests, 97 OPC integration tests, five doctests, strict library
Clippy, and focused DOCX/PPTX owner gates passed on explicitly selected stable
Rust 1.98.1; the pinned 1.95 installation lacked Cargo and lint/doc components.

Matched operation-local allocator vectors confirm the exact removed allocation
mechanism, while latency, RSS, copied-byte, physical-I/O, decompression,
managed-package, and broad OPC/OOXML conclusions remain withheld. Evidence is
recorded in [Change 0387](changes/0387-opc-source-materialization-shared-payload.md).
`performance_claim: none`; `claim_authorized: false`.

## Change 0377 compliance update

Change 0377 keeps SpreadsheetML cell ownership in `litchi-xlsx`, reuses the
existing typed sheet vocabulary, and delegates physical worksheet rewriting
and package publication to their established private/raw and OPC boundaries.
It adds no dependency edge, public raw type, package identifier, archive or
runtime handle, lock, or unsafe code.

The public operation is explicit rather than silently broadening `Set`.
Insertion is numeric-only, validates the shared SpreadsheetML payload ceiling,
proves physical absence plus range ownership, and refuses unsupported
semantics before staging. Existing source freshness, caller limits,
calculation invalidation, clone-staged reparse, exact no-op/signature policy,
effective-change signature refusal, atomic output, and owned-state inverse
contracts remain intact.

Focused and complete serialized XLSX validation plus doctests passed with only
the four exact pre-existing row-visibility exclusions documented in the
[Change 0377
record](changes/0377-xlsx-source-backed-missing-numeric-insert.md).
Production Clippy passed with one named pre-existing allowance; crate
boundaries and independent production/API, safety, and test reviews passed.

`performance_claim: none`; `claim_authorized: false`. No latency,
allocation-volume, RSS, physical-I/O, throughput, fixed-memory, broad XLSX, or
general OOM-prevention claim follows.

## Change 0376 compliance update

Change 0376 keeps BIFF12 Single Cell Tables ownership and opened-workbook
publication in `litchi-xlsb`; shared XML Maps vocabulary remains in
`litchi-ooxml-common`, and OPC graph mutation remains behind the existing
package API. It adds no public dependency edge, raw CRUD type, package
identifier, archive/runtime handle, lock, or unsafe code.

First-owner creation uses a canonical part and internal worksheet relationship.
Final-owner removal requires an exact owning relationship, lossless canonical
source, and a package-wide inbound-reference preflight. URI-equivalent
collisions, shared/orphan/foreign/external/malformed/dangling ownership, and
opaque/FRT or noncanonical changes fail closed with typed errors. Clone-staged
publication, complete reparse, source freshness, exact no-op/signature
preservation, effective-change signature invalidation, and owned-state inverse
semantics remain intact.

Focused and complete serialized XLSB validation passed with only the two exact
pre-existing exclusions documented in the [Change 0376
record](changes/0376-xlsb-table-single-cells-lifecycle.md). Strict Clippy,
crate boundaries, and independent topology, publication-safety, and test
reviews passed.

`performance_claim: none`; `claim_authorized: false`. No latency,
allocation-volume, RSS, physical-I/O, throughput, fixed-memory, broad XLSB, or
general OOM-prevention claim follows.

## Change 0375 compliance update

litchi-pptx remains the sole owner of the selected-slide source-backed
semantic snapshot and publication boundary; the existing bounded litchi-opc
publisher remains a substrate. The retained snapshot is immutable planning
state. Publication validates execution, version, lineage, URI, limits,
complete selected-slide closure, and exact selected bytes before applying
against it. Identity mismatch retains the exact semantic recapture/refusal
path, and later raw errors cannot retry or publish. No dependency edge, public
raw type, archive/runtime handle, lock, unsafe code, or parallel path is
introduced.

Focused evidence proves SlidePart/Scene capture reuse at 1/1 before and after
one-slide publication and 2/2 before and after the multi-slide batch. The
foreign-identical-source case performs one semantic recapture, returns
StaleSource, and writes zero bytes. This evidence is independent of timing
and does not imply an allocation or RSS result.

The three existing selectors were run under one-Cargo/one-process-at-a-time
discipline with CARGO_BUILD_JOBS=1, CARGO_INCREMENTAL=0, 8 GiB build and
4 GiB run caps, and no parallel build/worktree lane. Validation passed
533/533 library tests and 21/21 source-backed edit tests, while recording
exactly these unrelated pre-existing stale expectation exclusions:
opened::tests::stale_and_unsupported_raw_xml_fail_before_publication,
pptx_malformed_presentation::malformed_presentation_children_are_reported_by_their_owner,
and pptx_table_styles::noncanonical_style_target_survives_transactional_raw_save.
They concern stricter direct sldIdLst owner validation and never enter the
0375 publication path. Production-library Clippy, the 64-package/240-
declaration boundary gate, and independent equivalence/safety/test reviews
passed.

The clean release ABBA package and exact provenance are in [the compact
result](results/pptx-selected-slide-retained-snapshot-0375-abba.json).
Because the raw reports contain no correctness/refusal booleans and the
paired direction/stability gates fail, performance_claim is none. Output
hash equality is limited to generated-output byte identity. No
allocation/RSS/heap, reads, decompression/materialization, physical-I/O,
cold-cache, throughput/scaling, fixed-memory, general-OOM, all-PPTX,
real-producer, topology/media/notes/theme/chart, parallel, semantic reopen,
preservation, refusal, or reversibility claim is accepted.

## Change 0374 compliance update

Change 0374 keeps story-hyperlink capture and publication in the owning
`litchi-docx` layer and reuses the existing OPC source/overlay boundary. The
validated `Snapshot` is retained by `ForwardOnlyPatch`; publication verifies
execution, source version, lineage, complete artifact fingerprints,
post-fingerprint version, and equality with the snapshot's stored artifact
fingerprint before output. No public CRUD
signature, dependency edge, raw package type, archive/runtime handle, lock,
unsafe code, or parallel execution path is introduced.

Correctness and source safety take precedence over speed. Exact no-op bytes,
redaction story/XML and relationship locality, source immutability,
determinism, freshness, signature, cancellation, and failure atomicity remain
publication fences. The focused `publication_reuses_the_planned_story_snapshot`
test keeps `capture_source = 1` and `load_story = 1` from planning through
publication; this is independent counter evidence and is not an allocation,
RSS, or timing inference.

Serial validation used one Cargo process at a time, one build job, disabled
incremental/debug compilation, one test thread, an 8 GiB build cap, and a
10 GiB available-memory admission gate. `litchi-docx` passed `935/935`, the
story-hyperlink integration target passed `23/23`, the boundary gate passed,
and production-library Clippy passed with five named pre-existing allowances.
The exact unrelated pre-existing source-change test exclusion and the two
additional all-test Clippy debt lints are recorded in [Change 0374](changes/0374-docx-story-hyperlink-retained-snapshot.md).

The clean release ABBA used CPU affinity 2, one logical CPU, one worker, 20
warmups, and 500 samples per case. It accepts only the named no-op p50/mean
and redaction p50/mean/p95/p99 metrics on the fixed seven-story corpus. No-op
tails are withheld for control drift. The [ABBA result](results/docx-story-hyperlink-publication-0374-abba.json)
records all identities, raw hashes, output hashes, gates, and adjudication.

`performance_claim: scoped`; no reads, decompression, materialization,
allocation, RSS, physical-I/O, cold-cache, throughput, fixed-memory, general
OOM, all-DOCX, unmeasured-selector, or parallel-execution claim follows.

## Change 0373 compliance update

Change 0373 keeps generic archive allocation mechanics in `soapberry-zip`,
manifest/encryption/plaintext policy in `litchi-odf-common`, and family
`content.xml` policy in the ODP/ODS owners through the doc-hidden shared
validator. It adds no public API, dependency edge, raw CRUD type, package
identifier, archive handle, runtime handle, lock, or unsafe code.

Correctness and bounded hostile-input handling take precedence over speed.
ZIP and encrypted-Deflate materializers use checked `size + 1` arithmetic,
fallible platform conversion, fallible capacity reservation, and bounded
reads while retaining exact size and CRC verification. Encrypted ODF package
reads reject alias, compression, password, missing-size, and plaintext-limit
errors before payload reads, with final decryption checks retained.

ODP full-source and ODS full/selective owners enforce the shared 256 MiB
family limit using metadata-only materialized size before content reads;
encrypted entries use manifest plaintext size. Source freshness remains an
outer publication fence, and ODP explicitly reconciles it before exposing
secondary parse failures. No unsupported input is silently truncated or
approximated.

Focused and broader locked/offline release validation passed across all four
touched crates, subject only to two exact pre-existing writer-test skips.
Scoped Clippy passed with six named pre-existing allowances, crate boundaries
passed `64/240` with 14 existing debt entries, and independent static reviews
accepted the batch. See [Change 0373](changes/0373-odf-source-allocation-preflight.md).

`performance_claim: none`; `claim_authorized: false`. No latency,
allocation-volume, RSS, physical-I/O, cold-cache, throughput, fixed-memory,
or general OOM-prevention claim follows.

## Change 0372 compliance update

The fused catalog parser remains within the owning `litchi-odp` package layer
and consumes the doc-hidden validator owned by `litchi-odf-common`. It exposes
no raw XML type through CRUD signatures and adds no public API, dependency
edge, package identifier, archive handle, runtime handle, lock, or unsafe
code. Validator-first event dispatch and deferred catalog errors retain the
former sequential error precedence through verified EOF.

Correctness and bounded hostile-input handling take precedence over speed.
All input-dependent shared tracker growth is fallible; content exceeding the
256 MiB materialized limit is rejected before allocation, including encrypted
plaintext metadata. The 4,096-depth, 256-declarations-per-element,
65,536-page, 1 MiB name, source freshness, ZIP, MIME, publication, and media
locality fences remain in force.

Focused locked/offline release tests passed for every executed common, ODP,
and ODT target, subject to two exact pre-existing writer-test exclusions.
Scoped Clippy passed with one pre-existing `large_enum_variant` allowance;
the crate-boundary gate and independent static reviews passed.

Clean CPU-2 ABBA on the fixed corpus used control
`32290f7ce`/`7291b2...`, candidate `922eb5e2c`/`9a26cd...`, one worker,
30 warmups, and 500 samples per leg. Fresh-open p50 improved 15.627% and
17.327%, with both p50 same-side drifts below 5%. The [ABBA
result](results/odp-source-catalog-0372-abba.json) records the evidence.

The authorized claim is limited to fresh
`SourceBackedPresentationCatalog::from_read_at` p50 latency on that corpus.
Mean, tails, list, and query are withheld or rejected. No broad open, all-ODP,
RSS, allocation, fixed-memory, physical-I/O, cold-cache, throughput, or OOM
claim follows. See [Change 0372](changes/0372-odp-source-catalog-fused-parse.md).

## Change 0371 compliance update

Change 0371 places ODF content validation and namespace tracking in the owning
`litchi-odf-common` layer. Its doc-hidden private surface is consumed by the
ODT owner and does not expose raw XML types through ordinary CRUD signatures.
The migration removes duplicate ODT machinery and adds no package identifier,
archive handle, runtime handle, lock, or unsafe code.

The plain-`Reader` tracker uses checked `u32` depth, limits namespace
declarations to 256 per element, and preserves reserved binding, unbinding,
scope restoration, and empty-element deferred-pop behavior. The shared
validator retains the 256 MiB input ceiling and rejects nesting beyond 4,096
with a typed invalid-format error, avoiding quick-xml's `u16` namespace-depth
path. Correctness and bounded hostile-input processing take precedence over
performance.

Exact formatting, the crate-boundary gate, and the focused locked/offline
release tests passed, subject to two exact pre-existing writer-test failures
in unmodified code. Scoped Clippy passed with only the pre-existing
`large_enum_variant` lint in unmodified `package/model.rs` allowed.
`performance_claim: none`; `claim_authorized: false`; no latency, allocation,
RSS, physical-I/O, fixed-memory, or OOM-prevention claim follows. See [Change
0371](changes/0371-odf-shared-content-validation.md).

## Change 0370 compliance update

Change 0370 is harness-only and adds no production API, dependency edge,
archive handle, raw type, runtime handle, lock, or unsafe storage. The three
opt-in ODP selectors exercise the existing source-backed catalog through its
owning format boundary. The selectable registry is **407** and the default
remains **36 cases / 198 rows**.

The fixed corpus has 12 slides, 13 archive members, and eight deterministic
2 MiB `Pictures/*` members; the archive is 16,785,912 bytes with SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.
Open, list, and query preparation remains outside the corresponding timed
scopes: fresh catalog construction, `catalog()`, and the selected slide at
index 6. Semantic, topology, source-replay, and media-locality checks remain
untimed evidence gates.

The [control report](results/odp-source-catalog-0370-control.json) uses CPU 2,
30 warmups, and 500 samples. It is a dirty descriptive control from revision
`f35486fb7085bb128eb89a4d2e9edd3ad1065f02` with binary SHA-256
`08594839ede39d7f2ed0c143d818e41de0b7cdb77bc92fbcdd2a96083ca9966a`.
`performance_claim: none`; `claim_authorized: false`; no clean A/B, latency,
RSS, allocation, physical-I/O, or OOM-prevention claim follows. Strict Clippy
exposed only 23 preexisting unrelated diagnostics and the scoped allow-list
run passed; focused selector and enumeration tests passed `1/1` each, with no
full suite run.

## Change 0369 compliance update

The fused ODT catalog pass remains inside the owning `litchi-odt` source-backed
boundary. A private handler combines the existing `content.xml` validation and
text-block-kind observation without exposing a raw parser, package identifier,
archive handle, runtime handle, lock, unsafe storage, dependency edge, or new
public API. The one borrowing XML pass preserves the former validation-before-
scan error precedence. Source freshness, ZIP verification, cancellation, and
the 256 MiB content limit remain outer publication fences; the 1,000,000-block
and 4,096-depth execution ceilings are unchanged. Styles, media, and semantic
payloads remain cold in the measured open path.

The clean CPU-2 ABBA used 30 warmups and 500 samples per leg, control
`bf1cb55c6`/`a7991b...`, candidate `b712aafbf20e`/`1a75eb...`, and corpus SHA-256
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.
Confidence intervals did not overlap and same-side drifts stayed below 15%.
Open reductions were A1/B1 p50/mean/p95/p99 `53.560%`/`53.116%`/`51.008%`/
`49.508%`, and A2/B2 `56.320%`/`56.078%`/`54.542%`/`54.304%`. The [ABBA
result](results/odt-source-catalog-0369-abba.json) is the retained evidence.

Exact rustfmt, `cargo test -p litchi-odt --lib --tests` (557 library tests and
all integration targets; 926 total), scoped Clippy with `-D warnings`, and
independent code/resource review all passed or accepted. The only accepted
claim is deterministic large media-rich ODT corpus latency for a fresh
`SourceBackedDocumentCatalog::from_read_at` open. List is excluded as a
tens-of-nanoseconds unstable result; query is excluded as a 2.9%-3.4%
below-materiality result. No all-ODT, RSS, allocation, fixed-memory,
physical-I/O, cold-cache, throughput, or OOM claim follows.

## Change 0368 compliance update

Change 0368 is harness-only and adds no production API, dependency edge,
archive handle, raw type, runtime handle, lock, or unsafe storage. The
selectors remain in the existing ODT performance harness and exercise the
existing source-backed document catalog through its current owner boundary.
Open, list, and selected-block query preparation remain explicitly separated
from their timed scopes; semantic, source-freshness, and media-locality checks
remain untimed evidence gates.

The retained [control report](results/odt-source-catalog-0368-control.json)
uses the 10,008-entry, 13-member corpus with eight deterministic 2 MiB
`Pictures/*` members, CPU 2, 30 warmups, and 500 samples. Binary SHA-256 is
`cc6f5a148f0788210814254f521c681238cf77ce9eba29ff3b29b5486d6c6ae8`; source
revision is `14884ced9d8b29b7d2155134025986e9315ac771`; `dirty: true`. The
artifact is descriptive current-control evidence only, not a clean A/B result.

The focused catalog and count tests passed `1/1` each. The initial standalone
harness context was `233 passed, 7 failed, 1 ignored`; the count assertion was
corrected and its focused test re-passed, leaving six unrelated failures. The
selectable registry is **404** and the default remains **36 cases / 198 rows**.
`performance_claim: none`; no latency, throughput, physical-I/O, allocation,
decompression, cold-cache, RSS, fixed-memory, or OOM-prevention claim follows.

## Change 0367 compliance update

Change 0367 keeps direct active `mergeCells` handling in the owning
`litchi-xlsx::raw` selected-worksheet scanner and uses the existing verified
worksheet/dependency/ZIP/source/execution fences. Exact count, nonempty `ref`,
reference-grid, singleton, placement/direct-child, and overlap validation are
performed before the canonical transient `merge::Index` is used. A fenced
single-cell non-anchor returns `Covered`; anchors retain `Stored`/`Missing`.

Range `cells` and `visit` continue to expose sparse physical records,
including merge followers, without synthetic covered cells; `stored_extent`
is unchanged. The retained merge range cap is 16,384 with `try_reserve`; a
16,385th range drains through verified EOF and then requires eager fallback.
Unknown merge attributes, children, or payload likewise fall back after
draining; malformed structure is a hard typed error. Eligible cold paths
publish no `Store`, `PartData`, or semantic caches.

The transient index's internal `BTreeMap` and heap allocations are bounded by
the cap but not individually fallible, so no fixed-memory, RSS, or OOM claim
is made. No public API, dependency edge, or `stored_extent` contract changes.
Focused validation passed `14/14`, full `litchi-xlsx` library validation passed
`906/906`, and scoped Clippy passed with only the unrelated
`clippy::useless-asref` issue allowed. `performance_claim: none`; no latency
claim follows. See [Change 0367](changes/0367-xlsx-selected-merge-streaming.md).

## Change 0366 compliance update

Change 0366 remains in the owning `litchi-xlsx::raw` selected-worksheet layer:
the canonical bounded general-reference decoder adds no dependency edge or
public API. Predefined references remain eligible in formula, value, and inline
scalar payloads; decimal/hex numeric references are limited to ASCII formula and
value tokens of at most 12 bytes. Numeric inline, overlong/non-ASCII, or
XML-1.0-`Char`-invalid numeric scalar cases become `NotEligible` and use the
verified eager fallback, while malformed, custom, and out-of-range references
remain MCE/typed errors. The scanner still drains XML/MCE/x14ac and the OPC
reader to verified EOF before publication or callback, preserving failure and
verification ordering. Eligible cold `cell`, `cells`, and `visit_cells` paths
retain no `Store`, `PartData`, or semantic caches.

No public accepted-input behavior or API surface changes; the pre-existing
eager/shared-string XML-legality residual is out of scope. Focused validation
passed `9/9`, full `litchi-xlsx` library validation passed `892/892`, and scoped
Clippy passed with `-D warnings`, with only the unrelated
`clippy::useless-asref` issue allowed. This is correctness/boundary evidence
only: `performance_claim: none`, with no latency, RSS, fixed-memory, or OOM
claim.

## Change 0365 compliance update

Change 0365 keeps source-worksheet range selection in the owning `litchi-xlsx`
semantic path and uses the existing verified-reader boundary. Eligible cold
`SourceWorksheet::cells(area)` and staged `visit_cells(area)` use a sparse raw
scan whose dependency scans reach XML/MCE/x14ac EOF before ZIP CRC/size
verification and source/execution fences permit publication or callbacks. A
multi-index SST stream and direct style-count stream avoid publishing a
worksheet `Store`, `PartData`, or semantic dependency cache on that cold path;
warm `Store` behavior remains unchanged. No public raw reader, package handle,
or dependency edge is exposed, and `stored_extent` is unchanged.

Missing coordinates are omitted while explicit empty cells are preserved.
`NotEligible` requires eager fallback only after verified-reader completion,
so merges, shared/array/data-table formulas, row/column styles, rich, phonetic,
extension, foreign, and general-reference cases retain eager semantics.
`visit_cells` stages an owned `Vec` proportional to selected physical output;
this is not a fixed-memory or OOM claim. Focused validation passed `27/27`,
full `litchi-xlsx` library validation passed `883/883`, and package Clippy
passed with `-D warnings` apart from the unrelated `clippy::useless-asref`
issue. No latency or RSS claim follows; `performance_claim: none`. See
[Change 0365](changes/0365-xlsx-source-worksheet-range-streaming.md).

## Change 0364 compliance update

Change 0364 keeps selected-cell dependency resolution in the owning
`litchi-xlsx` semantic path and reuses the existing OPC verified-reader
boundary. The scan tracks maximum shared-string and direct cell-style
references across all cells plus the target SST index. Cold plain selected SST
and direct `c@s` values stream canonical `sharedStrings` then `styles`
sequentially without publishing `Store`, worksheet `PartData`, a full text
`Vec`, a style `Catalog`, or semantic dependency-cache state. Warm semantic
caches no longer rematerialize evicted `PartData`; no public signature
changes.

Each dependency reader reaches XML EOF and CRC, size, source, and cancellation
fences before a value or fallback is returned. Invalid, missing, or
out-of-range references and unsupported or oversize parts use established
eager diagnostics only after readers close. Rich, phonetic, extension, and
foreign SST entries, row or column styles, merges, shared, array, and
data-table formulas remain eager fallbacks. The final cell source and
cancellation fence runs even when parsing returns an error, preserving error
precedence and freshness boundaries.

This introduces no public archive handle, dependency edge, or low-level
reader surface. Focused validation passed `28/28`, library validation passed
`856/856`, and scoped Clippy passed apart from the known unrelated pre-existing
`hyperlinks` `useless_asref` issue. Quick-XML and current-item allocations are
bounded only by documented limits; no latency, RSS, OOM, or fixed-memory claim
follows. See [Change 0364](changes/0364-xlsx-selected-cell-dependency-streaming.md);
`performance_claim: none`.

## Change 0363 compliance update

Change 0363 keeps selected-cell routing in the owning XLSX semantic path and
uses the existing OPC `PartView::with_verified_decoded_reader` boundary plus
the raw selected-worksheet scanner. Eligible cold `SourceWorksheet::cell`
queries do not publish full worksheet `PartData`, `Store`, or cache state and
rescan on later cold queries; warm `Store` queries retain the existing fast
path. No public signature, `cells`, `visit`, or `stored_extent` contract
changes, and no dependency streaming is added.

Every `NotEligible` result enters the eager store only after verified-reader
return and CRC/size/source/context checks complete. Merges, shared strings,
styles, shared formulas, and rich inline values therefore retain eager-parser
semantics. Source/cancellation/ZIP errors remain primary, and final outer
fences run before a value is returned. Zero, unrepresentable, and
greater-than-2-GiB declared parts bypass the scanner and retain existing eager
behavior, avoiding a new lower part-size limit. This preserves the existing
typed error, freshness, and lossless-fallback boundaries without exposing a
reader, cache, or physical package handle through ordinary CRUD APIs.

Validation passed focused `7/7`, source `16/16`, and library `828/828` gates;
scoped Clippy passed apart from the known unrelated pre-existing `hyperlinks`
`useless_asref` issue. The single-job capped validation protocol observed no
OOM; that is protocol evidence only. No latency, RSS, fixed-memory, or
OOM-safety claim follows; `performance_claim: none`. See [Change 0363](changes/0363-xlsx-source-worksheet-selected-cell-scan.md).

## XLSB source-ingress hard-probe boundary (change 0354)

Change 0354 keeps XLSB semantics in `litchi-xlsb`, source/path admission in
the existing owner/facade boundary, and fallback storage bounded by the
caller's exact input limit. Recoverable non-ZIP/no-match/missing-manifest
probes retain compatibility fallback; hard ZIP/OPC/classifier failures are a
typed `OpcError` and do not trigger an eager `Workbook::from_bytes` retry.
The catalog is dropped before fallback and retained `Bytes` is moved without a
clone, preserving pointer/capacity ownership. Known non-XLSB variants return
`NotOfficeFile` without pathname reopen. No archive type, physical identifier,
runtime handle, lock, unsafe storage, or dependency edge is added; explicit
eager APIs, public smart detection, and the positive non-ZIP fallback remain
unchanged. Serial evidence is private filter `7/7` within XLSB lib `51/51`,
facade `23/23`, and successful `xlsx`/`xlsx,xlsb` checks under one job/thread,
disabled incremental/debug compilation, one disk target, and an 8 GiB limit.
The 564 MiB target and approximately 14 GiB post-run available memory with
saturated swap are not resource bounds. `performance_claim: none`; no
latency, RSS, OOM, constant-memory, allocation, physical-I/O, or broad XLSB
claim is made.

## XLSB source-backed fallback admission boundary (change 0353)

Change 0353 keeps XLSB ownership in `litchi-xlsb` and facade admission/routing
in `litchi`; no eager archive type, physical identifier, runtime handle, lock,
unsafe storage, or dependency edge is added. Post-admission source-owner
`UnsupportedFeature` no longer enters an eager full-workbook fallback. The eager
adapter reader/caches/state and private detector-side duplicate source/limits
state are removed, while typed errors, freshness fences, eager cache behavior,
explicit `open_xlsb_workbook*`, and `DetectedFormat::Xlsb` remain. Recognized
nonworksheet tabs are filtered by worksheet position; direct nonworksheet and
sparkline/pivot/slicer/timeline selections retain typed refusals. Recoverable
pre-admission probes may use the existing `Workbook::from_bytes` fallback, and
smart detection and explicit eager APIs remain eager. A selected worksheet and
required dependencies still materialize. The `23/23` and `40/40` serial tests
ran under one job/thread, disabled incremental/debug compilation, one disk
target, and an 8 GiB limit; 647 MiB target and approximately 15 GiB available
memory with saturated swap are not resource bounds. `performance_claim: none`;
no latency, RSS, OOM, constant-memory, allocation, physical-I/O, or broad XLSB
claim is made. This supersedes 0304's fallback wording only after source-owner
admission and does not alter the public eager contracts.

## DOCX selected-story source-backed lifecycle (change 0352)

The selected-story lifecycle remains in the DOCX owner and uses the existing
`SourceBackedPackage` source/overlay boundary; it adds no facade archive type,
dependency edge, lock/runtime handle, or unsafe storage. `Main`,
`Header(index)`, and `Footer(index)` snapshot, text-streaming, direct-paragraph
edit, source-bound patch/inverse, and same-topology one-part publication keep
source lineage, freshness, fingerprints, signature, cancellation, and failure
atomicity checks. Exact no-op/trailing-byte copying and canonical
relationship/content-type/external/shared-target plus namespace/MCE validation
preserve unsupported content by typed refusal. Managed edits and
footnote/endnote/comment/glossary stories remain outside this slice. Evidence
is `11/11` for the new story-text target and `16/16` for existing
`source_backed`, serialized under the 8 GiB limit with one job/thread; the 347
MiB target and approximately 15 GiB post-run available memory with saturated
swap are not performance evidence. `performance_claim: none`; no speed, RSS,
allocation, physical-I/O, benchmark, or whole-GOAL claim.

## Indexed-stream validation (change 0351)

The low-level indexed stream proof remains within the ZIP/OPC ownership
boundary: strict sink validation preflights encryption, method, single-disk
ZIP64 provenance, complete local/central metadata and descriptor agreement,
all physical spans, and overlap/central-intrusion refusal. The locator checks
classic/ZIP64 counts, offsets, record length/adjacency, short buffers, ZIP64
resolution, fallible retention bounds, retryable immutable-source single-flight,
and `ReaderAt` byte stability. Store uses an exact range; Deflate uses exact
`total_in`; strict CRC-zero applies only to bounded sink paths, while ordinary
owned and borrowed fallback compatibility remains. One 16 KiB scratch buffer
per active member is a local bound excluding source/index, sink/output, cache,
process memory, and concurrency. The rejected compressor/zlib `~65%` premise
has no artifact support. Final evidence is `315/315` soapberry-zip and `13/13`
litchi-opc operation accounting under serialized constrained execution.
`performance_claim: none`; no latency, throughput, RSS, allocation, syscall,
physical-I/O, decompression, concurrency, selector, artifact, or whole-GOAL
claim.

The final successful package/scenario-scoped commands, not workspace-wide, are
recorded in [Change 0351](changes/0351-indexed-stream-validation.md): `cargo
fmt --package soapberry-zip -- --check`; `cargo test -p soapberry-zip --lib --
--test-threads=1` => `315/315`; and `cargo test -p litchi-opc --test
operation_accounting -- --test-threads=1` => `13/13`, with the record's exact
`ulimit`, target, serialized job/thread, and debug/incremental environment.

## Verified-streaming hardening (change 0350)

The low-level verified streaming boundary now validates overreported reads
across `ReaderAt` loops, ZIP verification, streaming, and the OPC
`BorrowedReaderAt` adapter, checking offsets and counters before use. Bounded
sink `read_to*` / `read_entry_to*` paths require strict CRC equality; ordinary
owned reads preserve zero-CRC compatibility, borrowed nonempty zero-CRC returns
`None` for owned fallback, and Deflate extra output retains `InvalidSize`
precedence. This is one fixed-size scratch buffer per active member, not a
total process-memory bound. Final serialized evidence was fmt success,
`287/287` soapberry-zip, and `4/4` litchi-opc with `261` filtered, using one
job/thread, disabled incremental/debug compilation, one disk target, and an
8 GiB limit. `performance_claim: none`; no latency, throughput, RSS,
allocation, syscall, decompression, or concurrency claim.

## PhysPkgReader stored-Part borrowed consumer (change 0349)

Crate-scoped formatting evidence: `cargo fmt --package soapberry-zip --package litchi-opc -- --check` passed after formatting.

Change 0349 keeps the stored-Part borrow in the low-level immutable-slice
`litchi-opc::PhysPkgReader` boundary. It returns a validated source-backed
`&[u8]` only for the eligible Store case, after limits, CRC, and local/central
layout checks; the destination `Vec` allocation/memcpy and materialization
budget/cache charge are thereby avoided. Encrypted Store/Deflate inputs retain
typed errors before owned fallback, and nonempty CRC-zero inputs return `None`.
Deflate, ZIP64 EOCD, generic `ReadAt`, file, remote, `SourceBackedPackage`,
content-type, and signature paths are unchanged. The `8/8`, `10/10`, and
`281/281` serialized validation evidence is correctness/ownership only, with
`performance_claim: none`; no timing, RSS, throughput, physical-I/O,
decompression, or allocator claim follows.

## 2026-08-25: change 0280 leaves architecture unchanged

- Binary identity failed before replication smoke or collection.
- The operation-scoped session remains absent; existing per-read freshness boundaries and ADR interpretation remain unchanged.

## 2026-08-25: change 0279 rejected without an ADR change

- The experimental closure-scoped CFB session received architecture and freshness-precedence review, but its strict performance evidence failed the keep gate.
- Reversion restores the existing per-read freshness boundary; no public low-level session API, high-level CRUD leak, ownership change, or ADR amendment remains.

## XLS source attribution boundary (change 0278)

Change 0278 adds only an opt-in performance runner. CFB structure remains in
`litchi-cfb`, BIFF semantics remain in `litchi-xls`, and facade path routing
remains in `litchi`; no raw record, CFB ID, runtime handle, or mutable cursor
enters an ordinary CRUD signature. The runner stages and re-verifies an
immutable input, keeps oracles outside timing, and explicitly withholds
source/eager equivalence and physical-I/O claims. Its FileSource result selects
freshness sessions for separate design and correctness review but changes no
freshness contract itself. `performance_claim: none`; the matrix remains
**398 names** and **36 cases / 198 default records**.

## CFB monotonic cursor and XLS integration (change 0277)

Change 0277 keeps stream topology and cursor state in low-level `litchi-cfb`
and BIFF semantics in `litchi-xls`; no CFB cursor or physical ID crosses an
ordinary CRUD signature. Construction and skip inspect immutable validated
catalog state without publishing bytes, while every exact read retains
before/after freshness fences and `SourceChanged` precedence. FAT/MiniFAT
tails, limits, cancellation, FILEPASS, malformed input, unknown-record
locality, and duplicate-last behavior remain typed and tested. Strict central
evidence supports retention as an enabler, but exact-neutral I/O and rejected
tails keep `performance_claim: none`. The matrix remains **398 names** and
**36 cases / 198 default records**.

## XLS source-global coalescing (change 0276)

Change 0276 keeps CFB structure and source identity in `litchi-cfb`, BIFF
semantics in `litchi-xls`, and path routing in the facade. The explicit raw
handoff derives the exact source from an immutable catalog; ordinary CRUD
signatures expose no CFB IDs or raw records. Header-only preflight preserves
FILEPASS, EOF, limits, cancellation, stale-source, and zero-worksheet-overread
behavior before one exact global-span read. CFB payload errors now receive the
documented post-read freshness fence. Clean decision evidence favors keeping
the change, but strict latency/resource/physical-I/O claims remain withheld.
The matrix stays **398 names** and **36 cases / 198 default records**.

## Source-backed XLS selective-read closure (change 0275)

Change 0275 keeps CFB ownership in `litchi-cfb`, BIFF semantics in
`litchi-xls`, and path routing in the `litchi` facade. Public results are typed
sheet/cell values rather than CFB IDs or raw records. The owner is immutable,
version-fenced, finitely bounded, cancellation-aware, and read-only; FILEPASS
and malformed data remain typed refusals. Five opt-in selectors bring the
matrix to **398 names** while the default stays **36 cases / 198 records**.
Logical locality is accepted as correctness evidence, but the dirty smoke is
slower than eager and has `performance_claim: none`; request coalescing and
global/SST deferral remain open.

## Rejected DOC owner-public-phases hypothesis (change 0274)

Change 0274 is not a retained production optimization. The clean ABBA package
disproved representative end-to-end benefit from removing the public-reader
`Vec` clone: only large lifecycle p50 passed, tiny p50 was adverse, and
payload-heavy directions plus mean/tail statistics were rejected. The
candidate was reverted under the keep/revert rule. No claim-registry entry,
selector, or default count changed; the current matrix remains 393 names and
36 cases / 198 records.

## Latest DOCX section-layout closure (change 0273)

Change 0273 adds one opt-in typed existing-main-story section-layout selector,
bringing the current selectable matrix to **393 names** while preserving the
default **36 cases / 198 records**. PageSize/Margins/Start/Columns snapshot,
edit, durable patch/inverse, sequential publication, source-lineage, strict
namespace, signed/stale, and sink/refusal gates are correctness/CRUD evidence
only. The dirty five-sample profile is descriptive; clean retained performance,
physical-I/O, allocator/RSS, and broad DOCX evidence remain open.

## Current count and production gap (change 0272)

Change 0272 adds three opt-in benchmark-only source-overlay multi-part
selectors and 27 records, taking the selectable matrix to **393 names** while
leaving the default at **36 cases / 198 records**. It makes no production
optimization and declares `performance_claim: none`. Recompression remains
required; any future parallel or compression-policy change needs an explicit
execution context and scaling evidence.

## Current evidence-boundary additions (changes 0270-0271)

Change 0270 preserves the existing OPC ownership and public API boundaries. Its
timer covers only the production relationship-open operation, fences the
returned package with `black_box`, and leaves correctness oracles after timing;
it makes no performance claim.

Change 0271 is harness/result documentation over the existing XLSX repeated-
store selectors and allocator instrumentation. It adds no production
dependency, public type, executor, lock, unsafe code, cache, or CRUD surface.
The manifest binds control/candidate revisions, binaries, configuration,
corpora, and all tracked evidence files. Allocation figures are exploratory,
RSS is descriptive only, and latency, operation-local peak/RSS, physical-I/O,
decompression, copy, and broad XLSX claims are withheld. The default remains
36 cases / 198 records and claim-0269 remains latency-only.

Status: working design gate

Every production optimization must update this matrix before implementation
and link its correctness and performance evidence after verification. `Yes`
means the design can comply; it is not evidence that the code or benchmark is
complete.

| Optimization | Owner / dependency direction | Snapshot, edit and patch contract | I/O, memory and execution contract | Preservation, validation and security | Evidence gate | Status |
|---|---|---|---|---|---|---|
| CFB monotonic stream cursor for source-backed XLS (change 0277) | `litchi-cfb` owns the public low-level cursor; `litchi-xls` uses it privately; no reverse dependency, facade edge, or ordinary CRUD raw type | Read-only immutable XLS snapshots retain no cursor; construction/skip publish no bytes; exact reads fence before/after and commit cursor state only after the final fence | Forward-only FAT/MiniFAT chain state, contiguous physical-run reads, stack BIFF headers, and no MiniStream cache materialization; no runtime, executor, global cache, unsafe code, or logical-I/O reduction | Valid partial root tails, FILEPASS, EOF, STRING/CONTINUE, malformed records, limits, cancellation, stale/error precedence, unknown payload skip, duplicate-last, and selected-only locality are tested | [Change 0277](changes/0277-cfb-monotonic-cursor-abba.md), 282 CFB tests, 29 XLS source tests, strict checks/lints/docs/boundaries, four reviews, and clean 500-sample A1/B1/B2/A2 evidence | Retained measured enabler. Source p50/mean and open p95 pass, but logical work is exact-neutral and tails are rejected; `performance_claim: none`, with no selector-wide, tail, FileSource, physical-I/O, resource, or broad XLS/CFB claim |
| XLS source-global coalescing (change 0276) | Existing `litchi-cfb` catalog/source -> explicit `litchi-xls::raw` handoff -> facade; no reverse dependency or ordinary CRUD raw type | Immutable source/version identity and read-only sheet/cell semantics remain; facade classification, owner, and metadata share the exact catalog | Header-only bounded global preflight followed by one exact global span changes `2G -> G + 1`; no cache, runtime, executor, lock, unsafe code, or worksheet prefetch is added | FILEPASS, malformed EOF/header, input/global/SST/scan limits, cancellation, host precedence, stale source, selected-only locality, and post-error CFB freshness remain typed and tested | [Change 0276](changes/0276-xls-source-global-coalescing.md), 256 CFB tests, 28 XLS source tests, facade feature tests, clean 30-sample A1/B1/B2/A2 decision package, and two final reviews | Production optimization retained; `performance_claim: none` because strict 500-sample, uninstrumented FileSource, physical-I/O, allocation/RSS, producer, and broad XLS evidence remain open |
| XLSX repeated-store cache publication reuse (change 0269) | Existing source-backed `litchi-xlsx` worksheet store/cache path and harness-only ABBA package; no dependency, facade, or public API edge is added | The existing immutable source-backed snapshot and semantic query contract is unchanged; only repeated queries over the retained worksheet store are compared across revisions | The timer is exactly `semantic_query_only; explicit PartData reacquisition excluded`; four queries repeat eight times in fresh warm children with one worker. No executor, lock, unsafe code, global state, or physical-I/O assumption is added; no resource guardrail is claimed | Pinned medium/oversized corpus, selected worksheet/target, semantic projection, cache/read/Budget counters, query order/count, child identity, and absent source/sink/output channels are strict evidence identities; structural reacquisition controls remain outside the claim | [Change 0269](changes/0269-xlsx-repeated-store-cache-abba.md), checked-in strict ABBA package, 8 accepted/0 adverse latency cells, registry/package hash checks, formatting/diff checks, and independent review | Landed latency-only claim. No allocation/RSS, physical-I/O, cold-cache, save/publication, producer, resource, or broad XLSX claim |
| XLS owned-source numeric publication (change 0268) | Existing `litchi-xls`/`litchi-ole-common`/`litchi-cfb` owned-source publication path plus harness-only ABBA package; no new dependency edge or public raw type | Existing immutable source/version, same-length numeric edit, patch, exact output, and publication contracts remain; the change removes only the named redundant owned-source work | CPU-2 A1/B1/B2/A2 release legs use one worker, 20 warmups, and 500 samples; no executor, lock, unsafe code, global state, or physical-I/O assumption is added and no resource guardrail is claimed | Pinned opaque-heavy and RK/MulRK XLS/CFB corpora, semantic/output identity, clean binaries, and source/preservation correctness remain bound by the package; unsupported formats, producers, cold states, and broader CRUD are excluded | [Change 0268](changes/0268-xls-owned-source-publication-abba.md), checked-in strict ABBA package, 8 accepted/0 adverse latency cells, registry/package hash checks, formatting/diff checks, and independent review | Landed latency-only claim. No allocation/RSS, physical-I/O, cold-cache, producer, resource, or broad XLS claim |
| XLSX repeated-store strict schema and harness (change 0267) | Harness-only `tools/perf-baseline` filesystem selectors plus `tools/perf_abba_summary.py`; no production dependency, format-owner, or public API change | Four opt-in repeated-query `Case` entries use the existing source-backed XLSX snapshot/query path; primary selectors are reserved for same-selector ABBA comparison, while reacquisition controls are structural and do not alter the production edit/patch contract | Fresh child per sample, warm-only; eight repetitions of four semantic queries report the exact `semantic_query_only; explicit PartData reacquisition excluded` interval, managed Budget/cache/read diagnostics, and typed semantic projection. Explicit control `PartData` reads are outside the timed query scope; no executor, lock, unsafe code, or global state is added | Pinned medium/oversized corpus manifests, source/full/projection identities, query order/count, cache-limit and counter arithmetic, fresh-child IDs, allocator schema, source immutability, and primary/structural claim scopes fail closed; controls prove eviction or oversized bypass but are excluded from candidate elapsed comparison | [Change 0267](changes/0267-xlsx-repeated-store-strict-harness.md), 48 strict-summary tests, focused XLSX filesystem tests, selector/count gates, formatting/diff checks, and independent review | Neutral correctness/evidence-boundary infrastructure only. No latency, allocation, RSS, physical-I/O, throughput, decompression, producer, or production-performance claim |
| Fail-closed historical REPORT claim classification (change 0266) | `tools/check_report_claim_classification.py` and its JSON sidecar; no production dependency, format-owner, or public API change | Existing historical REPORT prose and measurements remain unchanged in meaning; only the two audited tables receive explicit `historical`/`descriptive`/`withheld`/`strict_claim` dispositions, with no strict links in the current sidecar | Bounded Markdown/JSON parsing binds exact headings, preambles, headers, row order, labels, and digests; malformed, reordered, duplicate/non-finite, escaped-path, or symlink-rebound inputs fail closed; no runtime, cache, lock, unsafe code, or library path is added | The sidecar covers 167 rows (145 historical, 14 descriptive, 8 withheld, 0 strict_claim) and validates the canonical strict claim registry; surrounding report prose remains explicitly outside the audited scope | [Change 0266](changes/0266-report-claim-classification.md), classification checker, 31 focused tests, CI integration, formatting/diff checks, and independent review | Report-integrity and claim-boundary evidence only. No latency, resource, production, or performance-improvement claim |
| PPTX slide-boundary publication selectors (change 0265) | Harness-only `tools/perf-baseline` use of the existing `litchi_pptx::opened::Snapshot`, `Transaction`, typed removal plans, patches, and OPC writer; no production dependency, format-owner, or public API change | Production snapshot/transaction and serialized forward/inverse patch contracts are exercised for dependency-free whole-slide removal and move; at change 0265's landing, the historical selectable matrix was 385 while the default remained 36 cases / 198 records | Deterministic in-memory four-slide corpus and bounded sequential sinks; plan, commit, publication, and reopen phase vectors are reported, with setup and oracles outside timing; no executor, lock, unsafe code, global state, or physical-I/O assumption | First/middle/last removal and final-slide refusal, both move boundaries, `from == to` no-op, semantic reopen, exact source immutability, twice-built determinism, stale/foreign/dependency/unknown-member/MCE/signed/limits refusal, partial/zero-sink behavior, durable patch replay/inversion, and raw local plus normalized-central untouched-member identity remain gated; moves require strict `[Content_Types].xml` identity | [Change 0265](changes/0265-pptx-slide-boundary-publication.md), focused boundary harness tests, deterministic corpus inspection, semantic/ZIP/durable/refusal gates, formatting/diff checks, and independent review | Correctness and phase evidence only. No latency, allocation, RSS, decompression, throughput, physical-I/O, or broad PPTX claim |
| Real-producer security corpus (change 0264) | Harness-only `tools/perf-baseline` test over existing facade, DOCX/XLSX/PPTX, DOC, CFB, OPC, signing, and VBA APIs; no production dependency, owner, or public API change | Read-only validation, exact no-op, password-authenticated semantic readback, and existing typed refusal contracts; no signature/protection/encryption/macro authoring is introduced | Ignored locked test with bounded input/output limits and a managed `Budget`; output is rejected before any accepted bytes and `Memory`/`Objects`/`OutputBytes` release is checked after RAII drops; no timer, executor, lock, unsafe code, or physical-I/O assumption | Eight source-hashed real-producer fixtures check valid signatures, signed/protected zero-output refusals, password error distinctions and semantic digests, inert VBA CFB-stream identity, external inventory without fetch, one-under input, and release-to-zero. Producer breadth and active code/link behavior remain excluded | [Change 0264](changes/0264-real-producer-security-corpus.md), the locked ignored corpus test, source-policy CI test, focused formatting/diff checks, and independent review | Correctness-only security coverage accepted. No selector/default-count, latency, allocation/RSS, physical-I/O, or resource-performance claim |
| XLSX vendor-extension preservation shape (change 0262) | Harness-only `tools/perf-baseline` use of the existing `litchi` facade, typed `litchi-xlsx` owner, and OPC/property oracle; no production dependency, format-owner, or public API change | Existing immutable XLSX cell-value snapshot/edit/patch/publication contracts are exercised; the CLI-only shape is excluded from `XlsxCellCrudShape::ALL` and adds no `Case`, leaving the then-current 381 selectable names and the default 36 cases / 198 records unchanged | Deterministic in-memory corpus and bounded sequential sinks; source-backed and managed controls retain existing cache/Budget contracts, typed positive-prefix/zero-output refusals, and no executor, lock, unsafe code, global state, or physical-I/O assumption | Exact orphan XML/BIN payloads, content types, XML-local relationship, no root/workbook/selected-worksheet relationship leak, semantic reopen, exact no-op/lifecycle, stale/foreign source, and raw local plus central-directory identity for untouched members remain gated. Signed/protected/formula/MCE/macro/unknown-owner refusals remain inherited production contracts | [Change 0262](changes/0262-xlsx-vendor-extension-preservation.md), focused vendor-shape test, one-sample corpus inspection, formatting/diff checks, and independent review | Correctness-only corpus accepted. No latency, allocation, RSS, decompression, physical-I/O, producer, or broad XLSX claim |
| Strict claim canonical recomputation (change 0261) | Private `tools/check_perf_claims.py` / `tools/perf_abba_summary.py` verifier; no Rust dependency, format-owner, or public API change | Existing claim registry, ABBA summaries, raw reports, and performance-result meaning remain unchanged; strict mode only determines whether declared evidence is admissible | Sequentially validates raw samples and projects one compressed report at a time; recomputes bounded elapsed statistics and identity projections without retaining elapsed sample values, discards raw report/sample payload before the next leg, and incrementally hashes canonical JSON; fails closed at 512 MiB per member, 2 GiB total decompressed input, and 64 MiB summary; no runtime, cache, lock, unsafe code, or library execution path is added | Raw elapsed samples, report profile, source/sink identity, and parsed resource A1/B1/B2/A2 values are authoritative; `time`/`heaptrack` run/status/artifact/parser identities fail closed; exact resource variant/revision/binary/harness-tool/profile binding is required; raw projection-marker fields are ignored; public projection helpers are not exposed; the public verifier path creates the module-private `_ValidatedProjection` trust carrier only after raw validation, and plain mappings or mutations are rejected before summarization | [Change 0261](changes/0261-strict-claim-canonical-recomputation.md), strict four-claim result, 141 relevant Python tests, and external `/usr/bin/time -v` verifier profile | Integrity and bounded-input gate accepted. The 1,114,076 KiB verifier RSS observation is descriptive; no latency, resource improvement, speedup, or library-memory claim |
| XLSX fresh-child filesystem roots (change 0260) | Harness-only use of the existing `litchi` facade, typed `litchi-xlsx` owner, and OPC/property oracle; no production dependency or public API change | Read-only open/query selectors; immutable workbook, edit, patch, publication, and exact-source contracts are unchanged | Every sample uses a fresh child. Timed workbook/projection stay live through elapsed/allocation/procfs/cold snapshots; warm and admitted page-aligned cold sources are hashed. Logical reads are explicitly not applicable on the uninstrumented facade path; no executor, cache, lock, unsafe code, or global pool is added | Fixed corpus generator/shape/size/hash, exact measured-object semantic projection, independent metadata, aligned-source identity, page-cache admission, cleanup, and explicit cold-ineligible states are gated | [Change 0260](changes/0260-xlsx-fresh-child-filesystem-roots.md), focused integration tests, warm/cold debug smoke, strict claim-registry check, and independent review | Fresh-child/cache-state correctness accepted. Debug dirty-worktree smoke is not claim evidence; latency, allocation/RSS, physical-I/O, storage-media, remote/range, producer, edit/save, and broad XLSX claims remain withheld |
| Shared lazy OPC structural members (change 0259) | Private `litchi-opc` archive adapter -> existing `soapberry-zip` lazy reader; no dependency, format-owner, facade, or public API change | Read-only structural catalog parsing only; immutable snapshots, edit/patch, exact-source authorization, and publication are untouched | Stored members remain borrowed; lazy deflated members retain the cache's `Arc<Vec<u8>>`; indexed/positional sources retain the owned fallback. No new cache, executor, lock, unsafe code, or global state | Existing lazy `read_shared` keeps CRC, declared/materialized size, limits, cache/single-flight, and error behavior; session/cancellation paths are untouched. Pointer-identity tests bind shared, borrowed, and owned paths | [Change 0259](changes/0259-opc-shared-structural-members.md), 227 OPC library tests, all OPC integration tests, 257 ZIP tests, strict Clippy/rustdoc, and independent review | Exact private ownership-copy removal accepted. No latency, allocation-count, copied-byte, peak-memory/RSS, decompression, physical-I/O, cold/remote, semantic-format, or broad OOXML claim before relationship-heavy controlled evidence |
| Unified RTF byte-native ingress (change 0258) | Existing `litchi` facade -> existing `litchi-rtf` parser/model; feature-gated detector integration only, with no reverse dependency or raw public type | Read-only facade construction retains the native immutable source; edit, patch, and publication contracts are unchanged | Owned bytes move into the native parser without an intermediate UTF-8 gate; the reader probe reads 12 bytes and restores position; no cache, executor, lock, unsafe code, or global state is added | Literal CP-1252, LZFu, stored MELA, malformed framing, exact native source, semantic parity, and ZIP/OLE2 precedence are gated. Facade/native parity uses the same parser/model and is not an independent parser oracle | [Change 0258](changes/0258-rtf-byte-native-facade.md), focused facade/detector tests, two retained strict ABBA packages, focused summary/package unit tests, and independent production/evidence review | Correctness accepted. Back-to-back matched binary/workload captures with differing runtime toolchain metadata do not reproduce the accepted latency set, so no latency, allocation/RSS, I/O/decompression, compressed-input, real-producer, rich-RTF, or broad speedup claim is registered |
| High-level ODT source-backed filesystem ingress | Unified `litchi::Document` facade -> existing `litchi-odt` source owner -> `litchi-odf-common` package; no reverse dependency or archive type crosses the facade | Read-only source owner retains eager semantic projection, immutable source/version fences, and byte-backed `from_bytes`; edit/patch/publication APIs are unchanged | One positional source and one ODF package/index are retained; ordinary ODT avoids complete input ownership and a second container scan. No runtime/cache/executor/global state is introduced | OOXML retains precedence from the same physical source; malformed catalog, MIME, I/O and source-change errors remain visible; ODS/ODP non-claim, semantic/Markdown parity, media identity, limits and mutation tests remain | [Change 0191](changes/0191-odt-unified-source-ingress.md), [summary](results/odt-unified-ingress-0199-summary.json), focused feature matrices, ODT/harness gates and two independent reviews | Generated-corpus correctness and logical-range evidence is accepted: the untimed replay reads 29,080 logical source bytes with zero picture overlap. Open-only latency remains withheld for drift; open-plus-full-text p50/mean/p95/p99 reductions of 30.02% to 35.36% pass the paired and drift gates. Allocation/RSS, physical-I/O, cold-cache, producer, edit/save and broad ODF claims remain withheld |
| Shared source-backed OPC overlay ownership | Additive low-level `litchi-opc::SourceBackedPackage` methods; format owners call downward without exposing archive types or changing dependencies | Changed same-topology snapshots move existing immutable payload ownership; exact no-ops publish an empty overlay plan; topology-changing relationship publishers and public semantic patch contracts are unchanged | Removes one complete selected-Part `Arc -> Vec -> Arc` copy; bounded 64-Part planning, selected source materialization/compare, XML validation, compression, cancellation, budgets and sequential sink behavior remain | Managed `PartData` bare-Arc escape remains refused; exact source, signatures, stale/foreign source, limits, duplicate Parts, raw untouched members, reopen and partial sink gates remain | [Change 0185](changes/0185-opc-shared-source-overlay.md), [summary](results/opc-shared-overlay-0185-summary.json), OPC/DOCX/PPTX/XLSX focused/full gates and independent review | Deterministic ownership-copy removal accepted. Named medium cell and large row-batch statistics pass; unstable/directionally inconsistent scenarios plus allocation/RSS, physical-I/O, decompression, cold-cache, topology-changing, producer and broad OOXML claims are withheld |
| Standalone process benchmark and deterministic synthetic corpus | `tools/perf-baseline`; no production dependency edge | Read-only workloads; mutations use public transactions | Explicit samples and truthful available source/sink counters, deterministic range simulation, filesystem process isolation and local-pool scaling; external profilers optional | Fixtures are deterministic and content-safe; RTF adds content-addressed producer coverage, ODT/ODS/ODP and source-backed DOCX/PPTX/XLSX add exact media-rich publication validation, common OLE2 adds four exact 4 MiB opaque streams, and source-backed OPC/PPTX/XLSX record semantic materializations plus complete physical I/O where separately accepted; the ordinary-root DOCX selectors add correctness/logical compressed-range evidence only; malformed cases remain bounded | Tool check plus reproducible JSON on a named host | The current harness has 393 selectable cases; 200 was the count before the opt-in ODF repair-plan selector was added. Four opt-in XLSB lifecycle selectors cover deterministic tiny/medium/large/sparse BIFF12 open, list, one-cell, and prepared full-cell scans with exact canonical hashes; the default 36 cases / 198 records are unchanged. Changes 0122-0125 add the ODP/ODS correctness, logical-range, and 4095-byte MiniFAT request-amplification selectors, change 0126 adds eight ordinary-root DOCX source-path controls, change 0127 adds matched ODS repeated-cell source-read controls, change 0134 adds matched ordered ODS cell-batch/source-version evidence, change 0135 adds matched native XLS fixed-width Number/RK/MulRK publication evidence, change 0137 adds matched native XLS plan-only numeric publication evidence, change 0139 adds matched ODP repeated-text cache correctness/source-replay evidence, change 0144 adds six configured simulated-range CFB selectors, change 0145 adds two deterministic PPTX cross-presentation slide-copy phase/counter selectors, change 0146 adds twelve CFB `open_stream` correctness/counter selectors, [change 0148](changes/0148-cfb-same-target-repeat-policy.md) adds different-SID A-B-A, public-bulk A-B-A, and overlapping same-target correctness/source-event selectors, [change 0153](changes/0153-rtf-tail-publication-plan-evidence.md) adds matched RTF Commit/PublicationPlan append and exact no-op evidence, [change 0154](changes/0154-odf-content-cow-publication-evidence.md) adds six matched ODT/ODS/ODP owned/source publication selectors, change 0159 adds one source-backed PPTX cross-copy selector, change 0160 adds one native DOC phase-attribution selector, [change 0162](changes/0162-rtf-picture-crud-evidence.md) adds two bounded RTF picture CRUD selectors, [change 0163](changes/0163-xlsx-cell-clear-remove-evidence.md) adds four bounded XLSX scalar-cell clear/remove selectors, [change 0164](changes/0164-rtf-paragraph-split-merge-evidence.md) adds two bounded RTF ordinary paragraph split/merge selectors, and [change 0166](changes/0166-xlsx-row-visibility-evidence.md) adds four XLSX existing-row visibility selectors without changing the historical default tranche of 36 cases/198 records. [Change 0174](changes/0174-ods-source-backed-cell-edit-publication.md) adds four matched ODS existing-cell publication selectors with correctness-only timing coverage. [Change 0188](changes/0188-ooxml-root-lifecycle-evidence.md) adds eight matched high-level DOCX/PPTX open-plus-query lifecycle selectors; its final-source warm ABBA accepts no latency statistic because drift gates fail. [Change 0191](changes/0191-odt-unified-source-ingress.md) adds four matched high-level ODT open/open-plus-full-text selectors with typed logical-range evidence. [Change 0189](changes/0189-xlsx-edit-composition-evidence.md) adds four correctness/phase-only XLSX join and three-way selectors with explicit conflict resolution and no latency claim. [Change 0167](changes/0167-xlsx-row-visibility-provenance-reuse.md) changes production publication only and adds no selector; [change 0168](changes/0168-xls-numeric-validation-fusion.md) changes production validation only and also adds no selector. [Change 0140](changes/0140-odp-repeated-text-cache-release-abba.md) accepts only the exact four-call prepared ODP selector-pair latency and whole-process allocation-call reductions; peak heap/RSS and broader I/O/ODF claims remain open. [Change 0141](changes/0141-xlsx-source-provenance-negative-result.md) records a clean seven-case XLSX experiment whose generic provenance fast path was fully reverted after neutral-to-worse latency and sub-3% allocation reductions; no architectural optimization from that experiment remains. [Change 0144](changes/0144-cfb-simulated-range-source-evidence.md) accepts only the named configured simulator's MiniFAT request/byte/service-floor and p50/p95 result; real cold/network/device and native semantic evidence remain open. [Change 0145](changes/0145-pptx-cross-slide-copy-evidence.md) records correctness and sink-counter evidence only. [Change 0146](changes/0146-cfb-open-stream-evidence.md) adds the shared runner; [change 0147](changes/0147-cfb-open-stream-release-abba.md) accepts its scoped one-shot configured-simulator result while retaining the repeat cost. Failure/retry, ineligible-root, FAT, native semantic, resource, and performance acceptance remain open. Broader format CRUD and acceptance-grade evidence for newer cases remain |
| Current-HEAD external resource profile | `tools/perf_resource_profile.py`; standard-library tooling only, no production dependency edge | Read-only evidence around existing harness selectors; no snapshot/edit/patch semantics are changed | Three harness samples plus one-sample `/usr/bin/time`, heaptrack, perf, and strace probes; unavailable tools and counters remain null/unsupported; explicit 1/2/4/8/available widths reuse the existing `ExecutionContext` selectors | Synthetic corpora are identified by archive/target hashes; no source bytes or raw external traces are committed; syscall totals are labeled whole-process and not logical/decompressed I/O | [Change 0115](changes/0115-resource-profile-current-head.md), [compact JSON](results/resource-profile-current-head-0115.json), and twelve parser tests | Recorded current-HEAD evidence only at revision `be500459961471659f65c180de0e5fe98bc14e3a`. The locked build succeeded; binary SHA-256 `1cbb2340eae13f4ed49d5baa27532e1f9b31d5781036bb2a302837bcd2210f5c` / 36,646,512 bytes is exact, but provenance is `build_succeeded_source_snapshot_only` because only a post-build dirty snapshot was retained. A clean pre/post run with untracked-content hashing is required for stronger source identity. Managed XLSX heaptrack totals are 6,130,956 calls / 1,026,348,498 bytes; CFB save reports 1,825 logical reads / 84,838,500 bytes; both scaling selectors classify `nonideal_or_measurement_noise`, with invalid Amdahl fractions null/flagged. No before/after, cold-cache, remote-range, physical-I/O, allocation-attribution, or optimization claim |
| Filesystem cache-state evidence | Harness-only filesystem module; production OPC/CFB APIs remain unchanged | Fresh child per sample; deterministic source/output hashes; atomic saves use isolated temporary destinations | Warm and cold-requested modes record logical reads, process I/O, materializations, changed spans and output hashes | Schema 1 smoke requires ten result records/five evidence records; eager/source OPC saves are byte-identical and CFB reports one span | [Debug correctness/counter record](changes/0087-filesystem-cache-state-evidence.md) and [compact summary](results/filesystem-smoke-0096-summary.json) | One-sample debug smoke only; no latency, memory, throughput or reproducible cold-cache claim |
| Strict Linux cold-verified filesystem evidence | Harness-only verifier; no production API, dependency, or unsafe code | 64-bit Linux only; opened read-write source FD; numeric `fstatfs` magic allowlist; regular/non-empty/page-aligned source; fsync and `posix_fadvise(DONTNEED)`; canonical hashed/versioned external fincore JSON proof immediately before source-touching timer; positive `/proc/self/io` `read_bytes` | Explicit ineligible statuses retain host/proof failures; prepared query controls are excluded; `cold_verified_status`/samples sit alongside warm and cold-requested states; aligned source SHA/size, filesystem magic, fincore basename/hash/version, stderr digests/lengths, and method/fallback are recorded without arbitrary path text; page-cache/read-bytes claim is separated from physical-media claims | No timed result is emitted unless residency, dirty/writeback, and process-I/O gates all pass; ZIP verifier copy uses EOCD comment alignment and CFB uses declared-chain trailing padding; padded DOCX compares derived semantic archive identity | [Change 0236](changes/0236-cold-verified-filesystem-evidence.md) and harness parser/state/adversarial tests | Opt-in capability only; 32-bit Linux is conservatively rejected; no captured run, cold-cache speedup, physical-media, device, allocation, or production-performance claim; fincore/getconf/procfs and supported filesystem availability remain host dependencies |
| Filesystem repeated release evidence | Harness-only filesystem module; production OPC/CFB APIs remain unchanged | Fresh child per sample; deterministic source/output hashes; atomic saves use isolated temporary destinations | CPU-pinned release run records warm/cold-requested latency distributions, logical/process I/O, materializations, changed spans and output hashes | Same schema-1 correctness gates; tmpfs and accepted advisory `posix_fadvise(DONTNEED)` remain explicit limitations | [Repeated release record](changes/0089-filesystem-release-repeated-evidence.md) and raw/summary artifacts | Descriptive 300-sample tmpfs evidence only; no physical cold-cache, comparator, allocation, peak-memory, or production-performance claim |
| CFB atomic-save phase attribution | Harness-only filesystem and metrics modules; `litchi-cfb` and format dependencies unchanged | Read-only counters around the existing exact operation; no snapshot, overlay plan or publication semantics changed | Three sequential intervals record elapsed and logical `ReadAt` deltas; exact sum checks fail closed; fresh children and CPU affinity remain explicit | Existing source/output hashes, complete CFB reopen, untouched-stream checks, one-span report and atomic destination replacement remain | [Change 0142](changes/0142-cfb-atomic-save-phase-attribution.md), [compact 400-sample record](results/cfb-save-phase-current-0142-summary.json), [compressed raw record](results/cfb-save-phase-current-0142.json.zst), focused harness tests and strict Clippy | Accepted as current-revision attribution only: open 135,680 B, plan 33,962,596 B, publication 50,740,224 B; no speedup, physical-I/O, allocation, peak-RSS or semantic format claim |
| CFB fingerprint request coalescing | Private `litchi-cfb::ValidatedOverlayPlan`; no public facade, dependency or format-owner change | Two planning fingerprint brackets and direct/atomic publication fingerprint stages remain; source version/stable-token, target hash, complete reopen, typed partial output and atomic rename contracts are unchanged | Complete fingerprints use a fallible right-sized window capped at 1 MiB; comparison/emission stay at 64 KiB and are dropped before fingerprint allocation; no runtime, cache, lock, unsafe code or concurrency added | Focused malicious-source/mutation/output/atomic tests, full CFB suite, strict CFB and direct-consumer Clippy, rustdoc, exact output/span/hash gates and two independent reviews remain | [Change 0143](changes/0143-cfb-fingerprint-read-coalescing.md), [clean ABBA summary](results/cfb-fingerprint-abba-0143-summary.json), [compressed 1,600-sample raw record](results/cfb-fingerprint-abba-0143.json.zst) | Accepted for the named CFB save: calls 1,825 -> 857 at equal 84,838,500 B; both ABBA directions improve warm and advisory-cold p50/p95/mean. Code-local window +983,040 B; whole-process RSS shows no increase, but physical-I/O, operation-only allocation/peak memory, guaranteed cold and broad DOC/XLS/PPT claims remain open |
| Bounded forward-only XLSX and RTF creation | Additive production authoring APIs; shared `litchi-core` hierarchical budget accounting; XLSX-private exact-output text batching; no dependency inversion or format facade change | Finite row/cell/string or RTF structure limits; deterministic forward-only output; exact hierarchical charge/rollback/commit and XML character/entity semantics | Bounded sequential sinks avoid seek requirements; RTF ASCII spans have a hard 32-byte sink-request ceiling and no new retained buffer; XLSX retains a 4 KiB row window while cumulative charges avoid transient ancestor vectors, reservations retain four nodes inline before spilling, and ordinary UTF-8 is appended by run | Production correctness suites retain limits, exact wire output, reservation, cancellation and sink-failure validation; deep-spill, partial rollback, concurrency, scalar-reference escaping, invalid XML, retry, and reopen tests are explicit | Commits `8245da20d`, `5918be8ec`, `d38cd455d`, `b79fd0480`, `f2279b121`; [change 0097](changes/0097-rtf-bounded-ascii-streaming.md), [change 0169](changes/0169-xlsx-streaming-budget-charge.md), [change 0170](changes/0170-xlsx-streaming-escape-runs.md), [RTF ABBA summary](results/xlsx-rtf-abba-0108-summary.json), and [XLSX escape summary](results/xlsx-stream-escape-0170-summary.json) | RTF fresh creation accepts the recorded geomean result. XLSX change 0169 accepts hierarchical-budget latency/allocation boundaries; change 0170 accepts large all-statistic, medium through-p95, and tiny p50 latency reductions with exact output. Withheld tails, regressed branch misses, total-memory/physical-I/O/cold-cache/richer-XLSX/producer claims remain open |
| XLSX source-provenance publication reuse | Private source-backed scalar-cell snapshot -> publisher handoff; no public facade or dependency change | Tri-state matched/mismatched/unavailable provenance preserves stale/foreign refusal and the prior complete semantic fallback | O(1) lineage/version proof removes only the redundant publication-time worksheet reload/reparse; publisher I/O, compression and sink bounds are unchanged | Exact output, topology, relationships, raw unselected-member identity, semantic reopen and >8 MiB fallback/refusal gates remain | Commit `85ec86106`; [change 0096](changes/0096-xlsx-source-provenance-publication.md) and [ABBA summary](results/xlsx-rtf-abba-0108-summary.json) | Accepted at source-backed p50 geomean -21.66%/-22.65% and p95 -21.38%/-22.70%; physical read/materialization counters unchanged, with allocation/RSS/cold-I/O evidence still open |
| XLSX row-visibility publication provenance reuse | Private row-visibility patch -> existing cell-values source publisher; no public API, archive type, or dependency edge | Matched/mismatched/unavailable lineage and source-version handling preserves exact no-op, stale/foreign refusal, immutable snapshots, patches/inverse, cancellation and failure atomicity | Removes one redundant semantic worksheet reload, cell parse and row-tag scan; mandatory OPC overlay validation/read, compression and bounded sequential sink remain | >8 MiB exact-one-read regression, raw opaque-member identity, semantic reopen, signed/protected/formula/MCE/macro/relationship and resource/refusal suites remain | [Change 0167](changes/0167-xlsx-row-visibility-provenance-reuse.md), [summary](results/xlsx-row-visibility-provenance-0167-summary.json), 765 XLSX library tests, focused row/cell suites and independent review | Production work elimination retained. Descriptive publication reductions are 50.42%-68.23% in both paired directions, but same-implementation drift exceeds 5%; no accepted end-to-end latency, tail, allocation/RSS, physical-I/O, cold-cache, or real-producer claim |
| Native XLS semantic validation fusion | `litchi-xls` owner callback -> additive `litchi-ole-common` forwarding seam -> `litchi-cfb` composed-view/fingerprint owner; no reverse dependency or BIFF type enters CFB | Immutable same-length plan, selected-range source preconditions, exact no-op callback skip, native typed owner errors, and source/target version/fingerprint contracts remain | Reuses the existing `Arc`-backed positional composed view and removes two complete source scans; no complete target-artifact retention/materialization at commit, runtime, thread pool, cache, or publication change. Semantic validation may allocate a candidate Workbook model; there is no total-memory claim | CFB reopen/range checks precede BIFF validation; worksheet coverage, protection, macro and numeric readback remain inside the final complete fingerprint fence; signed/encrypted/DRM ingress and publication preflight remain | [Change 0168](changes/0168-xls-numeric-validation-fusion.md), [summary](results/xls-numeric-validation-fusion-0168-summary.json), full CFB/OLE-common suites, 1,015 XLS library tests, strict deprecation/lint and two independent reviews | Production work elimination retained. Code-derived per-sample scan reduction is 33,991,680 B/34 calls for Number and 405,504 B/two calls for RK/MulRK. Descriptive complete-workflow and semantic-commit directions agree, but same-implementation drift exceeds 5%; no accepted latency/tail, allocation/RSS, physical-I/O, cold-cache, or producer claim |
| Immutable CFB numeric-plan direct publication | Native XLS plan-only snapshot -> additive owned-byte `litchi-ole-common` ingress -> private `litchi-cfb` provenance; no BIFF/archive type or reverse edge crosses a facade | Only an owned `Arc<[u8]>` can establish the proof; generic `ReadAt`, source/version/range preconditions, exact no-op, semantic readback, fingerprints and partial-output typing remain | Removes direct `write_to`'s two outer complete scans while retaining bounded 64 KiB emission and source/target hashes; checked composed views retain their complete preflight and owned atomic save retains emission hashes/durability while omitting only redundant outer scans; no runtime/cache/thread pool or artifact-sized allocation is added | Protection/macro/signature/encryption refusals are identical; generic mutation-during-publication defenses and owned atomic durability remain | [Change 0172](changes/0172-cfb-owned-numeric-publication.md), [summary](results/cfb-owned-numeric-publication-0172-summary.json), 21 CFB overlay/3 OLE-common/1,015 XLS library tests, strict deprecation/lint/boundary checks and two independent reviews | Accepted warm in-memory direct-publication result: Number and RK/MulRK complete-workflow all-statistic plus publication through-p95 improvements pass clean ABBA/drift gates. RK/MulRK publication p99, allocation/RSS, physical-I/O, cold-cache and producer claims remain open |
| Native XLS comment validation and owned publication fusion | Existing XLS comment snapshot -> additive `litchi-ole-common` owner callback and sealed owned-byte CFB ingress; no reverse dependency or container type enters the format facade | Immutable same-length NOTE/TXO splices, exact no-op callback skip, native typed owner errors, source lineage/fingerprints and partial-output typing remain | Removes one post-plan composed-view scan plus direct `write_to`'s two outer scans; 64 KiB emission hashes remain and no runtime/cache/thread pool or artifact-sized candidate is added | Initial CFB fingerprints and reopen/range checks precede BIFF comment inventory/readback; sealed bytes may omit only the redundant final planning scan while generic sources retain it. Width/length, protection/signature/encryption, atomic durability and generic-source mutation defenses remain | [Change 0173](changes/0173-cfb-comment-publication-fusion.md), [summary](results/cfb-comment-fusion-0173-summary.json), 1,015 XLS library/11 focused integration tests, strict lint/deprecation and independent review | Accepted warm in-memory scalar total/semantic/publication and bounded-batch semantic results. Scalar p95, batch total/publication, allocation/RSS, physical-I/O, cold-cache and producer claims remain open |
| Immutable CFB final-planning fingerprint elision | Private `litchi-cfb` owned-byte provenance consumed through existing `litchi-ole-common`/XLS paths; no new dependency or public marker | Initial fingerprints, candidate reopen, owner semantic validation, exact no-op behavior, retained plan identity, checked composed views and publication contracts remain | Removes one complete final logical scan and one digest pair only after reading a sealed composed view; generic `ReadAt` sources retain the final fence; no new allocation, cache, runtime, lock or thread pool | Dishonest stable-token generic sources still fail the final fingerprint; protected containers, expected ranges, emission hashes, partial progress and atomic durability remain | [Change 0178](changes/0178-cfb-owned-planning-fingerprint.md), [summary](results/cfb-owned-planning-0178-summary.json), exact generic/owned counter tests, CFB/XLS focused suites, strict lint/deprecation and two independent reviews | Deterministic work reduction accepted: 16,995,840 B/17 reads for comments/Number and 202,752 B/one read for RK/MulRK per effective plan. All measured paired p50 directions are lower, but every workload fails at least one stability gate; latency, resource, physical-I/O, cold-cache and producer claims are withheld |
| Source-backed PPTX retained edit catalog | Private `litchi-pptx::presentation::source` editor state over the existing immutable OPC owner; no dependency or public API change | Open-time root/relationship/slide validation is retained; selected-Part parse, exact lineage/version/closure patch checks, consuming publication, inverse and raw-member preservation remain | Retains bounded catalog metadata and removes repeated presentation-reference parsing plus full 200-slide binding reconstruction; no executor, lock, unsafe code or payload-cache escape is added | Managed `PartData`, MCE, signatures, stale/foreign sources, cancellation, limits and partial-output checks remain; cross-copy keeps its independent graph recapture | [Change 0179](changes/0179-pptx-source-catalog-reuse.md), [summary](results/pptx-catalog-reuse-0179-summary.json), exact lifecycle build counter, full/focused PPTX tests, strict lint/deprecation, rustdoc and two independent reviews | Deterministic work accepted: one/same-slide `3 -> 1` catalogs and 400 slide nodes removed; eight-slide `10 -> 1` and 1,800 nodes removed. Materializations/source reads are unchanged; paired latency directions disagree and stability gates fail, so latency/resource/I/O/producer claims are withheld |
| Source-backed PPTX validation catalog/graph fusion | Private `litchi-pptx::validation` traversal over the existing immutable OPC catalog; no public API, dependency, runtime or cache change | Read-only validation report and source-version fences remain; no snapshot, edit or patch contract changes | Package relationship-list passes change `2 -> 1` and each Part relationship-list passes `4 -> 1`; graph lookups, XML parsing and logical source reads remain | Package-first and node-before-relationship graph order, node ceilings, missing/invalid targets, external/signature/macro presence, MCE and inert security policy remain | [Change 0182](changes/0182-pptx-validation-catalog-graph-fusion.md), [summary](results/pptx-validation-fusion-0182-summary.json), 12 focused validation tests, full PPTX suite, strict lint/deprecation/rustdoc/boundary gates and two independent reviews | Deterministic traversal reduction accepted. Large generated semantic-corpus complete validation p50 is 7.08%-11.50% lower with all distribution/stability gates passing; tiny/medium latency and resource, physical-I/O, cold-cache, scaling, producer and broad PPTX claims remain withheld |
| Legacy CFB owner-validation fusion | Existing DOC/PPT/XLS format owners -> `litchi-ole-common`/`litchi-cfb` owner callback; no reverse dependency, archive type, runtime or cache crosses a facade | Immutable same-length plans, exact source/version/range preconditions, typed native owner validation, exact no-op behavior, patch/inverse and source/target fingerprints remain | Reuses the callback's positional composed view and removes one complete source scan plus one source/target digest pair per effective transaction; later publication/save scans remain | CFB reopen/range checks precede DOC/PPT/XLS validation, which remains before the final complete fingerprint fence; macro/protection/encryption and partial-output boundaries are unchanged | [Change 0171](changes/0171-cfb-owner-validation-fusion.md), [summary](results/cfb-owner-fusion-0171-summary.json), focused 15/26/12 owner suites, strict production lint/deprecation checks and two independent reviews | Production work elimination retained. The measured 2,135,552-byte XLS corpus removes three logical reads; 64-worksheet total and scalar/batch plan p50/mean/p95 improve in clean paired ABBA. Scalar total, p99, publication, resource, physical-I/O, cold-cache, producer and DOC/PPT latency claims remain open |
| Managed XLSX source-backed editor tranche | Private `litchi-xlsx::SourcePayload` plus existing `litchi-opc::PartData`/`SourceBackedPackage`; no physical archive type or dependency edge crosses the facade | Eleven focused editors retain exact source lineage, typed edits, source-bound patches, no-op identity, stale/foreign checks and failure atomicity; managed package constructors check the caller-owned execution context before/after snapshot and publication | `Managed(PartData)` retains the cache/Budget reservation; ordinary packages use `Owned(Arc<Vec<u8>>)`; managed-to-owned Arc escape returns typed `ManagedPartDataArcEscape`; direct publication materializes only proven selected Part(s) and raw-copies the rest, with tab-state's bounded workbook/worksheet overlay exception | Exact signed no-op/changed-signature, MCE/protection/relationship/unknown-owner, cancellation, sink-failure, source-version and one-byte-under `Resource::Memory` refusal/release tests remain; parsed stores, staging, rewritten candidates and output buffers are outside the memory accounting | [Change 0151](changes/0151-xlsx-managed-source-editors.md), focused source-editor checks, and the reported 765 XLSX unit/integration/doc tests | Production correctness/resource-accounting freeze only; no latency, allocation, RSS/peak-memory, copy, decompression, cold-I/O, total-memory, hardware, or real-producer claim |
| Bounded semantic validation and ODF repair | Additive DOCX/PPTX/RTF/XLS validation reports plus the generic ODF typed repair boundary; no benchmark dependency edge | Reports are read-only; ODF repair is explicit, source-checked, reversible and failure-atomic | Finite validation/repair input, preflight candidate, output and sink limits; no runtime/cache/global state is introduced | Typed reports retain format/security boundaries; ODF repair removes only one recognized local-header extra from a first stored `mimetype`, and refuses structural, encrypted, signed, macro and semantic repairs | Focused validation suites, ODF repair plan tests, and [change 0099 selectable correctness evidence](changes/0099-odf-mimetype-repair-selector.md) | Correctness/counter evidence only; zero retained sink output and bounded write requests are not a total-memory bound, and no latency/allocation/RSS/I/O claim is made |
| RTF logical-tail append evidence | Harness-only use of the existing `litchi-rtf` transaction API; no production dependency edge | Existing-document append and empty no-op retain exact source identity, patch/inverse, stale/foreign checks and complete reopen | Fixed 16 KiB non-seek hashing window caps accepted bytes per write and retains zero output; candidate snapshot remains outside that window bound | Exact sequential bytes, durable JSON replay/inverse, semantic projection and source conflict refusal are untimed gates | [RTF logical-tail record](changes/0090-rtf-logical-tail-append-evidence.md) and focused harness tests | Correctness/coverage only; no release ABBA, allocation, RSS, or speedup claim |
| RTF ordinary paragraph split/merge evidence (change 0164) | Harness-only use of existing `litchi-rtf` ordinary-body split/adjacent-merge APIs; no production dependency or facade change | One checked source-bound operation; exact source bytes, candidate semantic state, reversible volatile/durable patch behavior, forged result-artifact refusal, and exact no-op identity remain | Separate open/stage/commit/publication/lifecycle vectors; fixed 16-KiB windowed hashing sink retains zero output; no source-backed path or total-memory claim | Selector closure is literal ASCII/root-level ordinary body; API/focused tests refuse compressed/non-ASCII/opaque/nested/control/rich/protected/external/security-sensitive inputs, while focused native tests retain exact boundary-byte and forged-boundary refusal | [Change 0164](changes/0164-rtf-paragraph-split-merge-evidence.md), focused selector gate `rtf_paragraph_split_merge_selectors_are_opt_in_bounded_and_gate_complete`, and existing RTF paragraph split/merge tests | Correctness/phase/sequential-sink evidence only; no latency, speedup, allocation/RSS, transaction-memory, physical-I/O, cold-cache, source-backed, real-producer or rich-RTF claim |
| Logical output write-size distribution | Harness-only `tools/perf-baseline` sink instrumentation; no production dependency edge | Additive report evidence only; existing snapshot/edit/patch and case identities are unchanged | Every summarized sink records fixed logical `Write::write` buckets at the accepted-write point; rejected writes do not increment; no runtime, cache, lock, syscall or disk behavior is inferred | Zero-length and inclusive 512/4,096/16,384/65,536 boundaries are tested; bucket sum equals `write_calls`; no archive bytes or semantic output is changed | [Change 0107](changes/0107-output-write-size-evidence.md), focused sink tests, comparator compatibility test, and CI histogram invariants | Correctness/schema evidence only; buckets describe logical calls, not syscalls, disk I/O, copies, compression, latency, allocation, RSS or performance |
| RTF plain paragraph split/merge | Private `litchi-rtf` ordinary-body source map and existing edit/patch/writer owners; no dependency or facade change | One bounded source-bound split or adjacent merge; exact paragraph text, source boundary, candidate readback, durable forward/inverse and foreign/stale checks remain | One exact `\\par` insertion/removal splice; finite operation/source/allocation limits; existing sequential publication, no runtime/cache/lock/unsafe path | Literal-ASCII root-level ordinary body only; compressed, unknown/opaque, nested/control, external/active, table/field/drawing/object, review, protected and malformed surfaces refuse; signed-document verification is outside the proof | [Change 0106](changes/0106-rtf-plain-paragraph-split-merge.md), six focused integration tests, crate library/integration gates and strict diff review | Correctness-only CRUD coverage; no latency, I/O-range, allocation/RSS, cold/high-latency, stream-window, producer-breadth or rich-RTF claim |
| ODP matched cross-slide text-box publication evidence | Existing `litchi-odp` owned snapshot and source-bound scalar/batch model transactions; harness-only integration and no dependency edge | Eight fixed existing names/pages, no renames, one transaction/commit; complete semantic reopen, volatile/durable patch/inverse and stale-source checks | Scalar repeats candidate staging; batch submits one bounded set; each copies one complete output to a pre-reserved sink. Owned ODP exposes no positional-source or logical-Part materialization counters, so those fields are omitted | Exact media bytes/types, raw `mimetype`/styles/meta/all-media identity, batch raw-manifest identity, scalar regenerated-manifest disclosure, and existing limit/opaque/protection/signature/encryption gates | [Matched selectable record](changes/0084-odp-cross-slide-text-box-batch-evidence.md), harness regression, focused ODP batch suite and CI hash/sink gates | Evidence added; identical semantic projection with case-specific physical digests. No latency, instruction, allocation, memory, or materialization claim before frozen CPU-pinned balanced ABBA |
| ODT matched embedded-resource publication evidence | Existing `litchi-odt` owned snapshot and scalar/bounded embedded-resource transactions; harness-only integration and no dependency edge | Sixty-four fixed existing image owners and corresponding fixed same-length source/target paths, no owner structural changes, one transaction/commit; complete semantic reopen plus volatile/durable forward/inverse and stale-source checks | Scalar repeats existing-image publication; batch submits one 64-change base-snapshot set; each copies one complete output to a pre-reserved sink. Owned ODT exposes no positional-source or logical-Part materialization counters, so those fields are omitted | Exact frame/path/media-type and source/target payload-digest state, all retained media, raw untouched-member identity, deterministic case-specific hashes, and existing bounded/atomic/envelope refusals remain | [Matched selectable record](changes/0085-odt-embedded-resource-batch-evidence.md), harness regression, focused ODT batch suite and CI hash/sink gates | Evidence added; identical semantic projection is required without requiring byte identity. No latency, instruction, allocation, memory, or materialization claim before frozen CPU-pinned balanced ABBA |
| ODS content-validation catalog transactions | Existing `litchi-ods` content owner and unified document transaction; no new dependency edge | Clone-staged add/set/update/same-name replace/remove/clear/rollback, exact no-op, reversible exact-source patch and failure atomicity; referenced removals/clear refuse and dangling references must be repaired before changed commit | Bounded input/operations/attributes/output; package publication changes `content.xml` only and preserves untouched members; no runtime/cache/global state | Duplicate names, unsafe rename, opaque owners, MCE/DTD, output overflow and changed signed packages refuse; complete typed reparse and package reopen remain | Production tests for catalog CRUD, binding closure, limits, security, member preservation and patch/inverse | Correctness-covered and unmeasured; no latency, allocation, memory, I/O, or materialization claim |
| Validated same-length OLE2 stream overlay substrate | `litchi-ole-common` protected-container wrapper -> `litchi-cfb` validated positional source and physical publisher; guarded native XLS/PPT facades consume it without exposing physical CFB types | Existing logical streams only, identical lengths, duplicate/overlap refusal, exact source/target fingerprints and reusable immutable plan | Bounded FAT/MiniFAT span derivation, 64 KiB direct sequential output with typed partial progress, plus synced sibling-temp/atomic-rename path; no fallback to topology-changing render | Complete composed artifact reopens before output; selected streams read back; source version/fingerprint are rechecked before and during output; common wrapper retains signing/encryption/DRM refusals | Focused FAT/MiniFAT/fragmentation/limit/source/sink/atomic-path/protection suites plus narrow XLS/PPT facade gates; broader format integration and ABBA remain | Substrate adopted without a broad end-to-end speed claim; generic CFB coverage alone does not certify DOC/XLS/PPT semantic CRUD |
| CFB stream-chain validation scratch | Private `litchi-cfb::file` helper and existing `OleFile::open`; no facade, dependency, source, writer, or public type change | Read-only open remains one atomic validation; no snapshot/edit/patch/publication behavior changes | One fallible MiniFAT and one fallible FAT chain vector/visited map are reused within an open instead of allocating twice per nonempty stream; no cache, lock, executor, unsafe code, or global state | Root/general collectors, exact chain error order/messages, allocation labels/order, MiniFAT/FAT ownership, overlap/cycle/marker/bounds checks and final physical-layout reconciliation remain; focused malformed and legacy corpus gates pass | [Change 0190](changes/0190-cfb-stream-chain-scratch.md), [summary](results/cfb-chain-scratch-0190-summary.json), [manifest](results/cfb-chain-scratch-0190-manifest.json), focused scratch/allocation-validation tests, strict lint/rustdoc, and independent review | Accepted only for the exact two-shape process profile: allocation calls -48.44%, Heaptrack temporary allocations -98.94%, displayed peak heap flat; many-small p95/p99 and wide-root p50/mean/p95 improve in both A/B/B/A pairs within their drift gates. Many-small p50/mean, wide-root p99, operation-local bytes, RSS, cold/physical I/O, concurrent contention and native DOC/XLS/PPT claims remain open |
| CFB selective exact-range read evidence | Public `litchi-cfb::SharedOleFile::read_stream_range`; no format-semantic facade or dependency edge | Read-only caller-owned destination; source-version checks and failure discard rules preserve the existing reader state; no snapshot/edit/patch vocabulary is changed | Traverses only the requested logical FAT/MiniFAT sectors, keeps MiniFAT's lazy root-stream cache untouched, and records exact positional source calls/bytes; no runtime, global cache, or unsafe code | Complete validated CFB index, exact target hashes/lengths, fragmented-chain ordering, malformed-chain refusal and source-change checks remain; generic substrate evidence does not certify DOC/XLS/PPT semantic CRUD | [Change 0094](changes/0094-cfb-selective-read-evidence.md), [exact-range summary](results/cfb-selective-range-abba-0106-summary.json), [Change 0144](changes/0144-cfb-simulated-range-source-evidence.md), [simulated-range summary](results/cfb-simulated-range-0144-summary.json), focused CFB tests and deterministic harness gates | Exact-range acceptance covers MiniFAT source-byte reduction (261,184 -> 36 and 2,096,192 -> 36), read-stage p50/p95 and modest total-p50 direction. The configured simulator separately covers 36- and 4095-byte MiniFAT request/byte/service-floor reductions plus total p50/p95 agreement in both ABBA directions; its exact-work 4 MiB FAT control stays near neutral. Real cold/network/device, p99, allocation, peak-RSS and native semantic evidence remain open |
| CFB MiniFAT `open_stream` evidence | Public `litchi-cfb::SharedOleFile::open_stream`; no format-semantic facade or dependency edge | Read-only immutable owner; exact output hashes, source-version checks, and `StreamNotFound` refusal remain outside timing | One-shot, repeat-3, and sequential repeat-8 selectors record ordered positional requested/returned events for 36-byte and 4,095-byte targets; target-aware same-SID repeats stay on exact direct ranges, while different-SID and multi-MiniFAT bulk retain cache takeover. In the pre-0152 overlap design, same-target callers could also take over the root cache; change 0152 supersedes that candidate path with the direct handoff documented in the next row. Cache state is inferred from source counters because the private cache is not exposed | Candidate tests bind target-aware direct-repeat formulas. [Change 0148](changes/0148-cfb-same-target-repeat-policy.md) adds different-SID, public-bulk, and pre-0152 overlapping same-target correctness/source-event coverage. Failure/retry, ineligible-root, FAT, native semantic, and complete resource accounting remain outside this slice | [Change 0146](changes/0146-cfb-open-stream-evidence.md), [one-shot release result](changes/0147-cfb-open-stream-release-abba.md), [Change 0148](changes/0148-cfb-same-target-repeat-policy.md), [repeat release result](changes/0149-cfb-same-target-repeat-release-abba.md), [one-shot summary](results/cfb-open-stream-abba-0147-summary.json), [repeat summary](results/cfb-repeat-abba-0149-summary.json), focused/full harness and strict lint gates | Exact one-shot source-work reduction and the configured simulator's roughly 62-64% total one-shot result remain accepted for the older comparison. The target-aware follow-up changes same-target work from `[D,C,0...]` to `[D,D,...]` and accepts only aggregate repeat-3/repeat-8 configured-simulator totals (roughly 56-64% improvement in both ABBA directions). Later per-invocation direct reads and noisy local bulk/concurrent tails are explicit tradeoffs; no generic/local wall-clock, per-invocation, allocation/RSS, physical-I/O, cold/network/device, native-format, or iWork claim |
| CFB same-target MiniFAT single-flight (change 0152) | Private `litchi-cfb::SharedOleFile` state; no facade or dependency edge | Read-only overlapping same-target calls retain exact output hashes/lengths, source-version stability, typed missing-stream refusal, and payload isolation | Clean release ABBA control `e486e4b1` versus candidate `f46381c6` (introduced by `c270c8f3b`) on CPU 2: 20 warmups, 500 samples, 24 records per leg, 48,000 retained samples; candidate concurrent logical source calls 6,473 versus control 8,000 (19.09% fewer) | Root MiniStream cache and resource-accounting boundaries remain separate; all correctness/source-event invariants pass; no runtime selector was added at that revision (291 then; 295 after change 0153; 301 after change 0154; 302 after change 0159; 303 after change 0160; current 305 after change 0162; only `cfg(test)` source-event acceptance and tests changed in 0152) | [Change 0152](changes/0152-cfb-same-target-singleflight-release-abba.md), [summary](results/cfb-singleflight-abba-0152-summary.json), focused/full harness checks | Accepted only for the named source-event/correctness scope. Local/generic latency, allocation/RSS/peak memory, physical I/O/syscalls, cold-cache/device/network, decompression, native semantic, OOXML/ODF/RTF/iWork, and broader performance claims remain withheld |
| RTF tail publication-plan evidence (change 0153) | Harness-only matched use of existing `litchi-rtf::TailAppendCommit` and public `TailAppendPublicationPlan`; no production dependency edge | Four opt-in selectors cover changed append and exact no-op on tiny/medium/large plain uncompressed RTF; exact bytes/digest/semantic paragraphs/no-op identity and matched Commit durable patch apply/inverse/stale/foreign checks remain outside timing | Inputs are pre-staged; `elapsed_ns` is exactly the pre-staged publication-call interval around the respective public write call to a fixed 16 KiB non-seek sink. Planning/publication vectors have per-sample cardinality; reopen/lifecycle vectors are one-element preflight-only gates run once outside the sample loop. Explicit source-retained, complete-candidate-retained, and publication-window bytes are emitted; the plan retains no complete candidate artifact | Cancellation, sink failure/partial progress, publication/output limits, source-version/fingerprint gates and semantic reopen checks remain untimed; the Commit control supplies the reversible durable-patch contract, while the plan separately proves bounded source publication with intentionally asymmetric validation/work | [Change 0153](changes/0153-rtf-tail-publication-plan-evidence.md), focused harness tests, full harness tests and strict lint gates | Correctness/publication-boundary evidence only; no end-to-end, rich-format, allocation/RSS, physical-I/O, release ABBA, speedup, or generic performance claim |
| RTF standalone-picture CRUD evidence (change 0162) | Harness-only use of existing public `litchi-rtf` picture replacement/removal transactions and `commit.snapshot().write_to`; no production dependency edge | Two opt-in selectors replace 1/7/63 same-length PNG/JPEG payloads or remove 1/4/32 exact groups; independent raw splices, visible text, unselected pictures, no-op identity, volatile/durable forward/inverse and stale/foreign gates remain | Open, one batch stage, commit and fixed-memory hashing-sink publication are reported separately; timed sinks retain zero output and accept exactly the independently known candidate length, but transaction/snapshot/total memory is not bounded by that sink | Generated ASCII uncompressed root-level mixed-case/whitespace hex only; nested/compatibility pictures, wrong size/range, 65-operation batches and partial/zero sinks refuse in the harness gate. Compressed/binary/rich drawing producers remain outside the closure | [Change 0162](changes/0162-rtf-picture-crud-evidence.md), focused selector test, six-record debug smoke and strict all-target harness Clippy | Correctness/phase/sequential-sink evidence only; no latency, speedup, allocation/RSS, physical-I/O, real-producer, image-rendering or broad rich-media claim |
| XLSX scalar-cell clear/remove evidence (change 0163) | Harness-only matched use of the existing eager `WorksheetEdit` and positional source-backed cell-values editor; no production dependency edge | Four opt-in selectors cover one existing numeric `Sheet1!A1` owner on medium and dense/sparse four-worksheet corpora; clear retains an empty `<c>`, remove deletes the owner, with semantic/package/no-op/volatile source-patch/stale/foreign checks | Open, planning/staging, commit, sequential publication, and lifecycle phases are separate; a fixed 64-KiB windowed hashing sink retains zero output bytes, while generic logical source/materialization counters are recorded and eager counters are not applicable | Source-backed publication proves raw preservation of unselected ZIP members/media; the closure is ordinary numeric owners only. Wider protected, signed, formula, MCE, metadata, relationship, malformed-input, limit and failure-atomic refusals remain in focused production tests; `cell_values` has no durable source-patch wire contract | [Change 0163](changes/0163-xlsx-cell-clear-remove-evidence.md), focused selector test and debug smoke, full/strict harness gates | Correctness/phase/counter evidence only; no latency, allocation/RSS, physical-I/O, cold-cache, decompression, durable-source-patch, real-producer, or broad XLSX deletion claim |
| ODF content-COW publication evidence (change 0154) | Harness-only matched use of the committed family-neutral owned rebuild and `SourceBackedPackage` positional publisher; no production dependency edge | Six opt-in ODT/ODS/ODP selectors use real semantic edits and verify exact content, semantic reopen, inventory, positional untouched-member raw identity plus physical/central order, no-op, limits, cancellation, source immutability, and bounded sink output | Clean CPU-2 release A/B/B/A, 20 warmups and 100 samples per record; both pair directions accept 96.35%-96.63% p50 improvement at the prepared publication boundary, with agreeing p95/p99/mean and at most 1.441% absolute same-implementation p50 drift | Logical `ReadAt` replay is untimed and is not physical-I/O/decompression evidence; edit construction, archive open/indexing, reopen and refusal gates are excluded. No end-to-end, allocation/RSS, cold-cache/filesystem, real-producer, broad ODF CRUD, or iWork claim | [Change 0154](changes/0154-odf-content-cow-publication-evidence.md), [summary](results/odf-content-cow-abba-0154-summary.json), compressed raw A/B/B/A reports, focused/full/strict harness gates | Accepted only for the named prepared in-memory ODT/ODS/ODP publication boundary |
| ODS source-backed existing-cell publication | Additive `litchi-ods` transaction over the existing ODF positional owner and common source publisher; no reverse dependency or physical archive type crosses the facade | Sparse clone-staged edits target at most 4,096 unique existing ordinary cells; exact no-op, source lineage, failure atomicity, source-bound semantic patch/inverse and complete candidate readback remain | Retains only source/index/XML projections plus touched sheets and one bounded `content.xml` replacement; sequential publication raw-copies untouched ZIP members and uses the common cancellation, budget, output-limit and truthful sink-progress contract | Repeated rows, formulas, merges, style retargeting, protected owners, unknown values in rewritten rows, lossy row markup, changed signatures, encryption and unsupported ZIP layouts refuse; standard untouched table metadata remains raw | [Change 0174](changes/0174-ods-source-backed-cell-edit-publication.md), [release evidence 0177](changes/0177-ods-source-cell-release-evidence.md), [1% evidence 0183](changes/0183-ods-one-percent-release-evidence.md), full ODS all-target suite, strict focused Clippy, matched harness selectors and independent reviews | Fixed one-cell complete-lifecycle p50 accepted at -75.03%/-74.27%; a clean current-HEAD rerun accepts the fixed 21-cell 1% lifecycle at -72.07%/-72.61% p50 with all distribution/stability gates passing. No allocation/RSS, physical-I/O, cold-cache, real-producer, durable-ZIP-patch or atomic-save claim |
| PPTX additive-topology publication evidence (change 0158) | Private owned-source ZIP/OPC preservation substrate consumed by the existing bounded PPTX cross-presentation slide-copy path; no new dependency edge or selector | Unchanged physical members are raw-copied while generated Parts are appended; inverse removal admits only an exact physical suffix. Semantic/topology/dependency/durable-patch/source-immutability/stale/foreign/refusal gates pass; adversarial production tests bind raw member/comment/data-descriptor preservation and typed refusal | Clean CPU-2 release A/B/B/A control `e8a67b19e` versus candidate `d900ae633`, 20 warmups and 200 samples per case. Plain/media-rich total p50 improves 29.643%/26.196% and 43.294%/43.604%; media-rich publication p50 improves 49.321%/49.680%. Process-wide task-clock/cycles/instructions agree; RSS and peak heap remain neutral | Prepared canonical generated owned-source slide copy only. Plain publication tails are withheld after drift triggers. Complete-source copy, eager ordinary OPC, source-backed/physical/cold I/O, decompression/recompression bytes, real producers, broad OPC/PPTX, other formats and iWork remain outside the boundary | [Change 0158](changes/0158-pptx-additive-topology-release-abba.md), [summary](results/pptx-additive-topology-abba-0158-summary.json), compressed raw A/B/B/A reports and process-wide time/perf/Heaptrack sidecars | Accepted only for the named prepared owned-source PPTX slide-copy boundary |
| CFB MiniFAT physical-run evidence | Same public `SharedOleFile::read_stream_range`; harness-only target expansion, with no semantic facade or dependency edge | Read-only exact 4095-byte destination; existing source-version/failure-discard contract remains unchanged | A deterministic largest-below-cutoff MiniFAT target records exact source calls/bytes/ranges and tests one physical positional request over 64 logical 64-byte mini-sectors; no runtime/global state or unsafe code | Manifest/archive/target hashes, exact payload length, legacy amplification and positional range shape are checked; this generic evidence does not certify DOC/XLS/PPT CRUD | [Change 0125](changes/0125-cfb-minifat-physical-run-evidence.md), focused harness test and six-selector smoke; release ABBA/resource attribution remain pending | Correctness/request-amplification evidence only; no latency, p99, cold/high-latency, allocation, RSS, physical-I/O, or native semantic claim |
| CFB atomic-save duplicate-scan removal | Private `litchi-cfb::ValidatedOverlayPlan`; no facade, dependency, or archive-type change | Existing same-length immutable overlay plan, source/target fingerprints, exact no-op, stale-source checks, inverse/readback and failure-atomic save contract remain | Atomic `save` removes only the duplicate post-emission complete scan (`4N -> 3N`); output-time source/target hashing and final pre-rename preflight remain; direct `write_to` is unchanged; no runtime, cache, lock, unsafe code, or parallelism | Candidate reopen, changed-span/target fingerprint, flush/fsync, sibling-temp cleanup, destination preservation, late mutation and protected/encrypted refusal gates remain | [Change 0103](changes/0103-cfb-atomic-save-scan-evidence.md), four raw release reports and [compact summary](results/cfb-save-atomic-scan-0112-summary.json), focused scan/mutation/atomic-path tests | Accepted only for exact logical source-read reduction in one CFB atomic-save case: 101,751,908 -> 84,838,500 bytes and 2,084 -> 1,825 calls, with identical output. ABBA p50 directions disagree; no latency, allocation, RSS, peak-memory, physical-cold, high-latency, or semantic CRUD claim |
| Native OLE2 semantic baseline | Harness-only DOC/XLS/PPT public facades and exact-source transaction owners; no production edge or API change | No-op and one-edit cases verify deterministic bytes, exact forward application, inverse restoration and full reopen | Already-open snapshot edit boundary, owned output materialization, release distributions, Heaptrack/RSS and available hardware counters | Generated content-safe writer artifacts; complete semantic verification after every timing; payload-heavy and real/security corpora explicitly excluded | [Native baseline record](changes/0015-native-ole2-semantic-baseline.md), [36-record JSON](results/ole2-semantic-baseline-a57506d23-2026-08-11.json), 23 harness tests and CI smoke/release matrices | Baseline accepted; no production speedup claim; large XLS one-edit/save (1.722 ms p50) and full reopen are the next measured candidate |
| XLS comment/visibility source-splice publication | Existing source-backed `litchi-xls` semantic owners -> protected `litchi-ole-common` wrapper -> validated CFB same-length splice plan; no physical type leaks through ordinary CRUD APIs | Existing NOTE/TXO and worksheet-visibility owners only; one/256-comment and one/64-visibility edits retain exact source lineage, no-op/refusal and eager patch/inverse semantics | Exact source-relative ranges replace the prior complete Workbook replacement: 109/27,904 vs 80,946 bytes and 1/64 vs 18,166; complete candidate/readback remains and output writes stay bounded | Complete worksheet/comment/catalog/opaque-stream readback, source/target fingerprints, width/length checks, limits and signed/encrypted/protected refusals remain | [Change 0095](changes/0095-xls-semantic-splice-publication.md), [ABBA summary](results/xls-semantic-splice-abba-0107-summary.json), focused common/XLS/harness tests | Structural replacement-byte reduction accepted; balanced ABBA accepts no speedup or material regression. Allocation, RSS, peak-memory and physical/source-I/O remain open |
| XLS fixed-width numeric publication evidence | Existing `litchi-xls::cell_values` Number/RK/MulRK transaction owners plus harness-only deterministic native corpora; no new production dependency edge | Same-family `set_number`/`set_numeric`, exact source lineage, semantic patch/inverse for ordinary source-backed commits, stale/no-op identity and atomic unsupported/security refusals remain; plan-only is explicitly forward-only without artifact patch/inverse | Six opt-in selectors separate edit/set/commit/publication vectors and publish complete target bytes through the same preallocated bounded sink; source-backed uses `SourceBackedCommit::write_to` and retains a complete target snapshot, while plan-only uses `SourceBackedPlanCommit::write_to` and retains no target artifact at commit; all paths report exact splices, replacements, spans and fingerprints where available | Number `Untouched!E21` 42 -> 43, one RK plus one two-cell MulRK transaction, full Snapshot/Workbook reopen, family/value readback, untouched CFB topology/member bytes, equal Workbook lengths, sink bytes/write counts/digests, signed/macro/protected/unsupported refusal, and untimed 54016.xls forward producer gates remain | [Change 0135](changes/0135-xls-numeric-source-publication.md), [pinned current-revision baseline](changes/0136-xls-numeric-current-revision-baseline.md), [plan-only record](changes/0137-xls-numeric-plan-only-publication.md), [balanced release record](changes/0138-xls-numeric-plan-only-release-abba.md), raw schema-1 JSON, focused/full harness suites and strict tool checks | Accepted for complete-operation latency in the two measured fixed-width families: both A1->B1 and B2->A2 directions agree at p50/p95/p99/mean; Number process VmHWM also agrees in both matched directions. RK/MulRK RSS directions disagree, valid heaptrack A/B allocation totals are descriptive whole-process profiles with identical peak heaps, and no operation-only allocation, bounded-artifact-memory, physical-I/O, cold-cache, or broad-producer claim is made |
| PPT one-shape source-splice publication | Additive `litchi-ppt::text_edit` immutable source owner -> protected `litchi-ole-common` wrapper -> validated CFB same-length splice plan | One existing `TextBytesAtom`/`TextCharsAtom`, identical encoded length, exact source/persist/offset/expected bytes, source-bound forward/inverse and no-op identity | Stages only the selected atom replacement; the selector range-reads the Current User prefix, bounded persist chain, live Document metadata and one selected Slide instead of complete metadata streams; output writes stay bounded | Complete source fingerprint and composed-artifact reopen remain; selected-shape readback, persist-map/header/topology bounds, duplicate/trailing OfficeArt, cross-topology streams, unsupported dependencies, macros, embedded storages, encryption/signatures and stale/foreign sources refuse | [Change 0100](changes/0100-ppt-source-backed-shape-text-splice.md), [selector-range follow-up](changes/0102-ppt-source-backed-selector-range-reads.md), focused/full PPT suites, strict Clippy/rustdoc and adversarial review | Correctness/selector-counter coverage only; no end-to-end latency, allocation, RSS, total-memory, cold-I/O or real-producer claim |
| Reused OPC publication plan | Private `litchi-opc` writer state | Does not change snapshot/edit semantics | Builds bounded generated bytes and deterministic order once before output | Same XML audit, Part validation and incomplete-sink behavior; no repair/normalization | [Before/after record](changes/0001-opc-publication-plan.md), machine JSON, all-feature writer tests | Accepted: -37.0% allocation calls and -5.49% mean latency on the 2,048-Part compressible save; mixed smaller latency effects disclosed |
| Exact owned-source OPC no-op publication | Private `litchi-opc`; formats continue to see semantic/Part APIs | Clones share immutable source allocation; every mutable entry revokes locally | Bounded 64 KiB sequential writes; current eager Part memory remains | Preserves the complete validated ZIP byte-for-byte; borrowed ingress and mutations fall back safely | [Before/after record](changes/0004-opc-exact-owned-source.md), EOCD comment, clone/revocation/partial-sink and public DOCX/PPTX/XLSX tests | Accepted; large incompressible save -98.4%, with +22.6% profiled peak heap disclosed |
| Opaque source-backed lazy OPC package | `litchi-opc`; additive DOCX/XLSX/PPTX facades see semantic source-backed views only | Immutable `Send + Sync` source state; source-version checked; active handles pin payloads | `litchi_core::ReadAt`; mandatory-only open; finite weighted cache; managed opens charge exact physical `InputBytes`, exact accepted direct-sink `OutputBytes`, cumulative declared cold-load `Work`, retained catalog/flight/payload `Objects`, and retained/in-flight payload `Memory` to caller-owned hierarchical `Budget`; compatibility opens retain the finite unmanaged `SourceCacheLimits` path; no hidden runtime | Unknown/untouched Parts remain preservable; limits and cancellation are checked before/during reads, decompression and managed direct publication; unpinned clean entries alone are evicted | [EOCD ABBA JSON](results/abba-eocd-before-a.json), [positional XLSX JSON](results/xlsx-source-positional.json), [budget implementation record](changes/0086-opc-source-cache-budget-management.md), [deterministic contention evidence](changes/0088-opc-source-cache-contention-evidence.md), source/cache/version/budget/facade tests | Implemented selective-open and bounded publisher stages; focused correctness tests cover managed resource charging, retained-resource releases, pinning, eviction, single-flight, cancellation, sibling competition and contention invariants. Managed direct source-backed sinks additionally cover exact/no-op and changed overlay OutputBytes refusal/partial accounting. Release ABBA covers structural/distribution counters but accepts no speedup. Allocation/peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence remain incomplete, as does broad semantic CRUD; `OpcPackage` atomic saves, `to_bytes`, and unmanaged compatibility sinks remain outside this accounting |
| Source-backed same-topology bounded-Part publication | `litchi-opc` low-level package -> existing private Soapberry preservation boundary; no format facade or dependency change | Consumes one immutable source snapshot; selected URIs, content types, relationships and topology remain fixed; exact no-op identity and source version are enforced | Materializes only up to 64 selected original Parts, raw-copies every other member, caps sink writes at 64 KiB, charges exact accepted `OutputBytes` only for managed direct sinks, and adds no runtime/cache/lock | Per-Part/aggregate limits, duplicate names, original CRC/framing and changed XML are checked before output; signed real changes and unsupported/prefixed/ZIP64 layouts refuse; unknown members retain exact raw framing; partial output remains typed | [one-Part record](changes/0037-opc-source-backed-one-part-publication.md), [multi-Part facade record](changes/0077-pptx-source-backed-multi-slide-batch-publication.md), raw-name/order, unknown-member, duplicate/signature/XML/limit/source/sink tests and CI hash/counter gates | Accepted for up to 64 existing ordinary Parts; one-Part p50 -73.12% with materializations 4 -> 1, and the eight-slide PPTX consumer p50 -95.78% with materializations 229 -> 9; OutputBytes accounting is correctness-only and does not extend to `OpcPackage` atomic saves, `to_bytes`, or unmanaged compatibility sinks; broader semantic edits/topology/signature policies remain |
| Source-backed DOCX main-document publication | Additive `litchi-docx::source_backed` facade -> accepted `litchi-opc` one-Part publisher; no physical archive type or new dependency crosses the facade | Exact raw snapshot, isolated edit, reversible patch, stale-source checks and no-op identity remain; exhaustive operation gate refuses dependency transfers | Shares the selected raw main-Part Arc, materializes one logical Part and raw-copies every other member to a sequential sink; no runtime, global cache, unsafe code or parallelism | MCE-rewritten sources and changed signed packages refuse before output; complete reopen, unknown XML, strict/transitional, opaque payload, limits, source-version and partial-sink tests remain | [DOCX ABBA/counter/memory record](changes/0039-docx-source-backed-semantic-publication.md), focused facade/OPC tests, complete DOCX/harness suites and CI hash/materialization gates | Accepted for guarded main-Part transactions: media-rich p50 -97.43%, instructions -74.91%, materializations 17 -> 1; eager DOCX guard p50 +0.25%; MCE/transfers/PPTX/XLSX remain open |
| Source-backed PPTX selected/multi-slide publication | Additive `litchi-pptx` editor -> accepted `litchi-opc` bounded multi-Part publisher; no physical archive type or new dependency crosses the facade | Non-clone editor, exact package/presentation/selected-slide closures, one operation per selected slide, reversible source-specific batch patch, stale/foreign checks and no-op identity; up to 32 slides and 256 unique nonoverlapping selectors per slide | Materializes the mandatory presentation root plus selected slides and raw-copies every other member to a bounded sequential sink; one plan regenerates only changed selected slides; no runtime, global cache, unsafe code or parallelism | Duplicate slides/selectors, aggregate/output limits, MCE-rewritten slides and changed signed packages refuse before output; complete reopen, raw untouched-member/media/Part identity, limits, source-version and partial-sink tests remain | [one-shape record](changes/0044-pptx-source-backed-semantic-publication.md), [same-slide record](changes/0063-pptx-atomic-source-backed-shape-text-batch.md), [multi-slide record](changes/0077-pptx-source-backed-multi-slide-batch-publication.md), focused facade/OPC tests, complete PPTX/harness suites and CI gates | Accepted for bounded existing-slide shape-text batches: eight-slide media-rich p50 -95.78%, allocations -32.54%, peak heap -8.94%, materializations 229 -> 9; topology/relationship/signature policy remains open |
| Source-backed XLSX calculation-metadata publication | Additive `litchi-xlsx::calculation_properties` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or new dependency crosses the facade | Non-clone editor, exact workbook owner/content-type/URI/XML closure, isolated typed edit, reversible source-specific patch, stale/foreign checks and no-op identity | Materializes only `xl/workbook.xml` and raw-copies 11 other logical Parts to a bounded sequential sink; no runtime, global cache, unsafe code or parallelism | MCE projection and changed signed sources refuse before output; complete reopen, typed semantic/media/Part identity, limits, source-version and partial-sink tests remain; cells/formulas/chains/topology are excluded | [XLSX ABBA/counter/memory record](changes/0046-xlsx-source-backed-calculation-metadata-publication.md), focused calculation/OPC tests, complete XLSX/harness suites and CI hash/materialization gates | Accepted for calculation properties/features in one existing workbook Part: media-rich p50 -99.2519%, instructions -77.78%, materializations 12 -> 1; peak heap flat; general XLSX cell/formula publication remains open |
| Source-backed XLSX defined-name publication | Additive `litchi-xlsx::defined_names` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Exact workbook owner/content-type/URI/XML and ordered-sheet closure, complete typed catalog replacement, reversible patch, stale/foreign checks and no-op identity | Materializes only `xl/workbook.xml` and raw-copies 11 other logical Parts to a bounded sequential sink; no runtime/global cache/unsafe code | Protection, MCE or unknown catalog children, invalid local scopes and changed signed sources refuse; complete reopen, name/media/Part identity, limits, version and partial-sink checks remain | [XLSX defined-name ABBA/counter/memory record](changes/0076-xlsx-source-backed-defined-names-publication.md), focused tests, complete XLSX/harness suites and CI gates | Accepted for the direct defined-name catalog in one existing workbook Part: media-rich p50 -97.84%, mean -97.81%, instructions -78.45%, materializations 12 -> 1; peak heap/RSS flat; general workbook/sheet topology remains open |
| Source-backed XLSX page-break publication | Additive `litchi-xlsx::page_breaks` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Exact workbook/relationship/worksheet closure, typed page-break edit, reversible patch, stale/foreign checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts to a bounded sequential sink; no runtime/global cache/unsafe code | MCE projection, retargeted relationships, non-worksheets and changed signed sources refuse; complete reopen, page-break/media/Part identity, limits, source-version and partial-sink checks remain | [XLSX page-break ABBA/counter/memory record](changes/0061-xlsx-source-backed-page-break-publication.md), focused source editor tests, complete XLSX/harness suites and CI gates | Accepted for page breaks in one existing normal worksheet: media-rich p50/mean -97.86%, materializations 12 -> 2; peak heap/RSS flat; general cells/formulas/topology remain open |
| Source-backed XLSX page-margin publication | Additive `litchi-xlsx::page_margins` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Exact workbook/relationship/worksheet closure, typed set/remove edit, reversible patch, stale/foreign checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts to a bounded sequential sink; no runtime/global cache/unsafe code | MCE projection, retargeted relationships, non-worksheets and changed signed sources refuse; complete reopen, page-margin/media/Part identity, limits, source-version and partial-sink checks remain | [XLSX page-margin balanced/counter/memory record](changes/0067-xlsx-source-backed-page-margin-publication.md), focused source editor tests, complete XLSX/harness suites and CI gates | Accepted for direct page margins in one existing normal worksheet: media-rich p50/mean -97.93%, materializations 12 -> 2; peak heap/RSS flat; general cells/formulas/topology remain open |
| Source-backed XLSX print-options publication | Additive `litchi-xlsx::print_options` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Exact workbook/relationship/worksheet closure, typed set/remove edit, reversible patch, stale/foreign checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts to a bounded sequential sink; no runtime/global cache/unsafe code | MCE projection, retargeted relationships, non-worksheets and changed signed sources refuse; complete reopen, print-options/media/Part identity, limits, source-version and partial-sink checks remain | [XLSX print-options ABBA/counter/memory record](changes/0070-xlsx-source-backed-print-options-publication.md), focused source editor tests, complete XLSX/harness suites and CI gates | Accepted for direct print options in one existing normal worksheet: media-rich p50 -97.87%, mean -97.88%, materializations 12 -> 2; peak heap/RSS flat; general cells/formulas/topology remain open |
| Source-backed XLSX page-setup publication | Additive `litchi-xlsx::page_setup` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Retained workbook/worksheet closure plus complete worksheet relationship identity, typed set/remove edit, reversible patch, stale checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts; no runtime/global cache/unsafe code | Printer references, MCE projection, relationship mutation, non-worksheets and changed signed sources refuse; complete reopen, media/Part identity, limits, version and partial-sink checks remain | [XLSX page-setup ABBA/counter/memory record](changes/0073-xlsx-source-backed-page-setup-publication.md), focused tests, complete XLSX/harness suites and CI gates | Accepted for relationship-free settings in one normal worksheet: media-rich p50 -97.78%, mean -97.79%, materializations 12 -> 2; peak heap/RSS flat; printer settings/general cells/topology remain open |
| Source-backed XLSX sheet-protection publication | Additive `litchi-xlsx::sheet_protection` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Retained exact workbook/worksheet and complete worksheet-relationship closure, atomic complete-metadata staging, reversible patch, stale/foreign checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts; no runtime/global cache/unsafe code | MCE-selected protection, relationship mutation, non-worksheets and changed signed sources refuse; complete typed core/x14 readback, media/Part identity, limits, version and partial-sink checks remain | [XLSX sheet-protection ABBA/counter/memory record](changes/0078-xlsx-source-backed-sheet-protection-publication.md), focused source/codec tests, complete XLSX/harness suites and CI gates | Accepted for complete direct protection metadata in one normal worksheet: media-rich p50/mean -97.75%, instructions -77.87%, materializations 12 -> 2; allocation/RSS regressions remain below 5%; password verification/general cells/topology remain open |
| Source-backed XLSX data-validation publication | Additive `litchi-xlsx::data_validation` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Retained exact workbook/worksheet and complete worksheet-relationship closure, atomic complete-collection staging, shared immutable snapshots, reversible patch, stale/foreign checks and no-op identity | Materializes workbook plus selected worksheet and raw-copies ten other logical Parts; redundant post-write/source-identity semantic reparses are removed without weakening typed readback; no runtime/global cache/unsafe code | MCE-selected collections, relationship mutation, non-worksheets and changed signed sources refuse; complete typed core/x14 readback, media/Part identity, limits, version and partial-sink checks remain | [XLSX data-validation ABBA/counter/memory record](changes/0079-xlsx-source-backed-data-validation-publication.md), focused source/codec tests, complete XLSX/harness suites and CI gates | Accepted for complete direct validation collections in one normal worksheet: media-rich p50/mean -97.75%, instructions -73.43%, materializations 12 -> 2; allocation calls +4.92%, peak heap/RSS flat; cells/formulas/topology remain open |
| Source-backed XLSX auto-filter publication | Additive `litchi-xlsx::auto_filter` editor -> accepted `litchi-opc` one-Part publisher; no physical archive type or dependency edge | Retained exact workbook/worksheet, complete worksheet relationships and styles/DXF closure; isolated add/replace/clear/no-op edit, shared immutable snapshots, reversible patch, stale/foreign checks and no-op identity | Materializes workbook, selected worksheet and styles Part, then raw-copies nine other logical Parts; no runtime/global cache/unsafe code | MCE-selected or protected filters, invalid DXF references, relationship mutation, non-worksheets and changed signed sources refuse; strict/transitional typed readback, media/Part identity, limits, version and partial-sink checks remain | [XLSX auto-filter ABBA/counter/memory record](changes/0080-xlsx-source-backed-auto-filter-publication.md), focused source/codec tests, complete XLSX/harness suites and CI gates | Accepted for one direct worksheet auto-filter/sort state: media-rich p50/mean -97.75%, instructions -73.57%, materializations 12 -> 3; allocation calls -1.94%, peak heap flat and RSS -1.35%; cells/tables/formulas/topology remain open |
| Source-backed XLSX conditional-formatting evidence | Existing additive `litchi-xlsx::conditional_formatting` editor -> accepted `litchi-opc` one-Part publisher; harness-only integration, no physical archive type or dependency edge | Exact workbook/worksheet, complete worksheet relationships and styles/DXF closure; one complete ordered core-collection replacement, immutable snapshots, exact reversible patch, stale/foreign checks and no-op identity | Matched eager/source paths use identical typed target values and the same worksheet rewriter; source-backed materializes workbook, selected worksheet and styles, then raw-copies nine logical Parts; no runtime/global cache/unsafe code | MCE/x14-selected or protected owners, invalid DXF references, relationship mutation, non-worksheets and changed signed sources retain existing refusals; complete typed reopen, exact patch/inverse, media/Part/raw-member identity, source/sink bounds and hashes are checked outside timing | [Selectable deterministic evidence](changes/0082-xlsx-conditional-formatting-performance-evidence.md), focused production source/codec tests, harness regression and CI hash/materialization gate | Evidence added: byte-identical output and materializations 12 -> 3; acceptance and all latency/instruction/allocation claims await balanced ABBA and resource-counter evidence |
| Raw unchanged ZIP entry passthrough, shared changed-Part payload and moved generated local span | Private soapberry physical ZIP boundary -> private OPC publication provenance; additive low-level shared-payload constructor only | Exact whole-source authorization remains clone-local; same-topology dirty closure is compared semantically; immutable changed payload is shared only during publication | Bounded 32 KiB raw copy and 64 KiB OPC sink adapter; eager source retention remains measured; the 4.19 MiB logical-payload and 4.20 MiB post-validation local-span copies are removed | Preserves comments/extras/descriptors/order/non-parts; topology and unsupported layouts fall back before output; complete generated archive validation and generated XML ownership remain | [Raw passthrough record](changes/0008-targeted-opc-preservation.md), [shared-payload record](changes/0021-opc-shared-regenerated-payload.md), [local-span record](changes/0022-zip-generated-local-span-move.md), raw framing/closure/fallback/partial-sink tests | Accepted for owned same-topology OPC mutation: original p50 geomean -84.98%; payload follow-up peak heap -3.42%; local-span follow-up peak heap -3.20% and few-large p50 -4.09%/-2.70%; uninstrumented RSS flat |
| Explicit bounded CPU execution | Neutral `litchi-core` context -> soapberry local session -> explicit OPC session | Ordinary document APIs expose no executor/runtime type | Worker/task/byte/threshold/affinity/cancellation and hierarchical budget policy; no global Rayon pool | Cancellation and resource failures remain typed and atomic | [1/2/4/8/12 scaling record](changes/0009-range-source-and-scaling.md), focused core/container/OPC tests | Large tasks reach 4.52x OPC / 5.93x CFB p50 at 12 CPUs; small tasks regress and require thresholds; legacy/default paths remain serial |
| Weighted Part cache and per-entry single-flight | Private source-backed OPC state | Unpinned clean values evictable; active handles and flights retain reservations; dirty edit state never evicted | Finite `SourceCacheLimits`; optional hierarchical `Budget` charging exact physical `InputBytes`, cumulative declared cold-load `Work`, retained catalog/flight/payload `Objects`, and retained/in-flight payload `Memory`; content-free diagnostics; no duplicate same-Part miss work | Reservations occur before payload I/O, cancellation/source changes fail typed, failed loads are not retained, and budget releases on eviction/drop | [Budget-management record](changes/0086-opc-source-cache-budget-management.md), [deterministic boundary/contention record](changes/0088-opc-source-cache-contention-evidence.md), eviction/pinning/single-flight/hierarchy/sibling/cancellation/version tests and harness counter gates | Implemented and correctness-tested across all managed resource dimensions; compatibility opens remain unmanaged and finite. Release ABBA adds structural/distribution evidence but accepts no speedup. Allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance acceptance remain pending |
| Move fresh XLS/PPT buffers into `OleWriter` | Concrete format writer -> `litchi-cfb` | Fresh authoring only; no snapshot or patch change | Removes one owned-buffer copy; current CFB sink contract unchanged | Identical CFB bytes and reopen semantics; bounded buffers retained | [Before/after record](changes/0003-legacy-owned-stream-handoff.md), writer suites, hashes | Accepted for XLS/PPT; DOC regressed 58% and was reverted |
| Validated cached CFB sibling-tree lookup | Private `litchi-cfb` reader | No public identity or snapshot change | Uses retained validation keys for ordered descent; no cache lock/global state | Exact MS-CFB name comparison, malformed-cache/tree checks and legacy corpus retained | [Before/after record](changes/0002-cfb-lookup-and-sector-buffers.md), 1,500-sample ABBA, corpus tests | Accepted: 56-66% faster at 256 siblings and about 94% at 2,048 |
| Direct MiniFAT and reusable sector buffers | Private `litchi-cfb` reader | No semantic state change | Parse into final bounded vectors and reuse one sector buffer | Preserves allocation/cycle/truncation checks and error order | Same CFB record; adversarial truncated/noncontiguous tests | Accepted: open allocation calls -6% to -9%, temporary allocations -21% to -28% |
| XLS validated editor reuse | Private native `litchi-xls` cell-value snapshot handoff | Exact source owner, patch lineage and immutable snapshot semantics unchanged | Reuses the already rendered/reopened CFB editor and removes one discarded BIFF owner parse | Final offset-bearing owner parse, complete public Workbook reopen, typed readback, resource checks and protection policy remain | [XLS ABBA/counter/memory record](changes/0016-xls-commit-editor-reuse.md), complete XLS and harness suites | Accepted: large one-cell edit/save p50 -7.72%; allocation calls -1.19%; peak heap and uninstrumented RSS flat |
| XLS fixed-width numeric inventory carry-forward | Private native `litchi-xls` BIFF owner inventory; no public type or dependency edge | Exact source lineage, immutable snapshots, same-family storage, semantic patches, stale checks and inverse bytes remain | Exact value-field certification carries unchanged offsets/resources forward, shares untouched worksheet inventories and clones only edited sheets; no cache/runtime/lock/global state | Candidate Workbook length and every byte outside disjoint Number/RK/MulRK fields must match; the common CFB reopen, complete public Workbook validation, all-sheet gate, changed-cell public readback and all nonnumeric/structural/resource fallbacks remain | [XLS latency/profile/counter/memory record](changes/0059-xls-fixed-numeric-inventory-carry.md), focused proof/sharing tests, complete XLS and harness suites | Accepted: large one-cell edit/save p50 -7.83%, mean -7.37%, p95 -7.20%; peak heap -5.54%, RSS flat; read-only nanosecond trigger disclosed |
| XLS terminal validated-render handoff | Rejected cross-crate OLE2 editor -> XLS snapshot prototype | Equality, failure atomicity, exact no-op, patch and complete owner/readback invariants passed in the prototype | Removed one deterministic CFB rerender without retaining a cache | Prototype fully reverted; no production or API change remains | [Rejected profile/ABBA/memory record](changes/0028-xls-terminal-render-handoff-rejected.md) | Rejected: large changed save neutral (-0.39% p50); repeated large exact no-op +22.00% p50/+16.69% mean |
| Common OLE2 publication attribution and rejected handoffs | Rejected private CFB writer/common-editor prototypes; no format or dependency edge remains | Candidate render/reopen/capture, exact output, allocation reuse, object rediscovery, no-op and public owner readback remained in every prototype | Shared writer payload removed staging copies; the alternate editor retained one already validated output allocation; inline recapture reused staged allocations only after exact comparison; none added a runtime, lock or global cache | All prototypes and tests were fully reverted; retained open/publication/finish/end-to-end cases verify four untouched 4 MiB streams, one replacement and complete CFB reopen | [Rejected common publication record](changes/0033-ole-common-publication-handoffs-rejected.md), [stage-attribution record](changes/0036-ole-common-stage-attribution.md), primary and DOC/XLS guard ABBA | Current open/publication/finish/end-to-end p50 **1.382/7.979/5.473/26.086 ms** and non-additive; rejected: shared payload +32.02%, validated render target -34.06% with DOC guards regressed, inline recapture end-to-end -2.61% p50/-2.30% mean with p95 +0.54% |
| DOC batched stream publication | Private native DOC revision owner -> shared OLE2 object editor | One isolated candidate; repeated paths are deterministic; exact no-op, patch and inverse identity unchanged | Applies WordDocument/table/Data replacements before one bounded CFB render/reopen instead of publishing each stream separately | Every stream is bounded; final strict revision-owner and independent public-document reopens remain mandatory | [DOC ABBA/counter/memory record](changes/0017-doc-batched-stream-publication.md), differential common-layer and complete DOC tests | Accepted: large one-paragraph edit/save p50 -10.52%; duplicated publication-site allocations nearly halved; peak heap and uninstrumented RSS flat |
| Native DOC owner/public-reader phase attribution | Feature-gated `litchi-doc` content-free event seam -> standalone harness-owned clock; ordinary format APIs and dependency direction are unchanged | Exact writer bytes and the complete open/edit lifecycle retain strict owner plus independent public-reader validation, exact no-op, patch/inverse/stale and typed-refusal gates | A bounded preallocated harness recorder timestamps ordered synchronous events; edit construction, staging, outer operation, output materialization and checked unattributed intervals remain explicit; successful event validation is outside named phase/outer intervals but inside checked lifecycle time; the library owns no clock, global recorder or ambient state | Source/candidate/output hashes, semantic reopen and untouched CFB streams are verified outside timing; successful harness event order/cardinality and separate format-side balanced error tests are explicit; `Finish` is labeled in-memory owner rendering, not publication | [Change 0160](changes/0160-doc-owner-public-phase-attribution.md), [clean release distribution](results/doc-owner-public-phases-0160-summary.json), focused format/harness tests and three-shape debug schema smoke | Exact attribution accepted: four CPU-2 processes/shape and 2,400 retained samples identify complete public-reader validation as the largest grouped named large/payload-heavy phase; disclosed subphase spread triggers do not change rank. No speedup, optimization, physical-I/O, allocation/RSS, cold-cache, filesystem, or real-producer claim |
| Native DOC lazy/fused fingerprint proof (change 0165) | Private `litchi-doc::body_text::Snapshot` `OnceLock` diagnostic cache plus same-allocation `Patch::apply` fast path; no public dependency, facade, runtime, global cache, lock, or unsafe edge | Exact source bytes remain authoritative. Same-lineage apply uses allocation identity; an independently reopened identical source uses the lazy fingerprint and then exact-byte comparison. Existing no-op, inverse, stale/foreign, malformed/refusal, reopen, and patch contracts remain; expected source/target fingerprints are independently computed by the harness | The established `measured_total_ns` lifecycle boundary is unchanged. Same-lineage apply and the first source/target fingerprint demand are explicit post-lifecycle vectors (`same_lineage_apply_ns`, `deferred_fingerprint_ns`, `workflow_no_diagnostic_ns`, `workflow_with_fingerprint_demand_ns`); one cached `u64` is per immutable snapshot, with no physical-I/O or cold-cache claim | Complete source/candidate/output hashes, semantic reopen, untouched CFB streams, reopened-source replay, independent FNV-1a values, exact source/target vectors, and four workflow/identity gates remain mandatory. The former `const` fingerprint accessors are now non-`const` because demand may initialize the cache; this is a capability change, not a deprecation, and is disclosed in the change record | [Change 0165](changes/0165-doc-lazy-fingerprint.md), [summary](results/doc-lazy-fingerprint-0165-summary.json), [release manifest](results/doc-lazy-fingerprint-0165-manifest.json), focused DOC/harness tests and strict lint gates | Accepted correctness/attribution evidence only for the named private lazy/fused same-lineage boundary, with no speedup claim: clean post-rebase CPU-2 control `d6818e290` versus candidate `5dd813b1e`, A1/B1/B2/A2, 20 warmups and 500 samples per shape (6,000 primary samples) plus 24,000 guard samples, with descriptive lifecycle p50 positive-faster deltas +33.77%/+33.21% tiny, +12.28%/+13.81% large, +17.33%/+17.82% payload-heavy and immediate fingerprint-demand p50 +14.56%/+13.89%, +4.50%/+5.83%, +6.55%/+7.08%; same-lineage apply/patch p50/mean/p95 deltas span ~99.6%-99.99%. Final DOC guard p50 is noop +78.84%/+79.89% tiny and +71.08%/+70.40% large, one-edit +37.23%/+40.81% and +20.45%/+19.79%, and open -3.52%/+0.13% and +0.55%/-1.80%. Neighboring XLS one-edit/open is mostly neutral or improved; its tiny no-op nanosecond cell is directionally noisy. The three-sample, preflight-inclusive whole-process Heaptrack probe reports 50,677 allocation calls and 128.28 MiB peak heap on both sides; it is not operation-scoped, and RSS is descriptive only. Broader shared physical/parsed substrate, public-reader fusion, physical-I/O, cold-cache, real-producer, generic total-memory, and broad DOC CRUD claims remain open |
| Native DOC borrowed public-validation input | Rejected private `litchi-doc` two-call prototype; no public type, dependency edge, cache, runtime or production change remains | Both strict-owner and independent complete public-reader validations, error order, final readback, exact no-op, patch/refusal and untouched streams remained | Replaced two complete `Vec` clones used as cursor input with borrowed slices; temporary commit and branch were removed | Identical source/output hashes and all gates passed, but large end-to-end latency and tails regressed | [Rejected change 0161](changes/0161-doc-public-validation-borrow-rejected.md), [A/B/B/A summary](results/doc-public-borrow-0161-summary.json) | Rejected: candidate was slower on large by 3.06%/7.31% p50 and 37.52%/14.49% p95; payload-heavy p50 direction disagreed; no allocation/I/O/cache claim |
| Native DOC PieceTable physical index | Private `litchi-doc` CLX representation; no public type or dependency edge | Indexed/scalar differential coverage preserves exact sorted CP intersections for ANSI/UTF-16, discontiguous, overlapping and saturating FC intervals | One immutable FC-ordered index and prefix maximum replaces repeated full-piece scans; no runtime, cache, lock or global state | Complete CLX, PAPX, CHPX, FKP and public DOC readback remain; fast-save overlap is handled explicitly rather than assumed absent | [DOC latency/profile/memory record](changes/0050-doc-piece-table-physical-index.md), 1,024 adversarial differential queries, complete DOC/harness suites | Accepted: large open p50 -55.91%, mean -55.78%; changed edit/save p50 -31.08%; range-scan self cycles 36.89% -> 4.17%; peak heap/RSS flat |
| Native DOC adjacent style-baseline cache | Private `litchi-doc` PAPX parser; no public type or dependency edge | Fresh/cached differential coverage preserves resolved properties, direct `grpprl`, initial style identity, piece modifiers, direct style switches and cache rekeying | One parse-local `(style index, resolved baseline)` pair replaces repeated inheritance reconstruction/validation; constant memory, no retained snapshot state, runtime, lock or global cache | Huge/Data-indirected PAPX, direct style switches/permutations, table/revision state, complete FKP parse and strict/public DOC readback remain | [DOC latency/profile/memory record](changes/0051-doc-adjacent-style-baseline-cache.md), focused base/derived/switch/rekey test and complete DOC/harness suites | Accepted: large open p50 -11.44%, mean -11.87%; edit/save p50 -4.01%; allocation calls -18.61%; peak heap/RSS flat |
| Native DOC CHPX range index | Private `litchi-doc` normalized character-run table; no public type or dependency edge | Scalar differential and reference-identity tests preserve half-open overlap behavior, output order, empty/reversed queries and numeric boundaries | Binary-searches monotonic ends and scans only the matching start slice; no storage, allocation, runtime, lock or cache | Complete CHPX/PieceTable parsing, formatting cascade, fields/pictures/comments/glossary consumers, exact patches and strict/public readbacks remain | [DOC ABBA/profile/memory record](changes/0053-doc-chpx-range-index.md), focused boundary/identity tests and complete DOC/harness suites | Accepted: large paragraph-list p50 -21.07%, mean -20.93%, p95 -20.00%; `extract_runs` self cycles 7.56% -> 1.23%; allocations and peak heap/RSS flat |
| Native DOC PAPX containment index | Private `litchi-doc` exact-source paragraph resolver; no public type, stored index or dependency edge | Scalar/indexed reference-identity tests preserve empty/gap, half-open and numeric-boundary containment | Two predecessor binary searches reuse parser-normalized piece/PAPX ordering; no allocation, runtime, lock, cache or retained state | Table filtering, CP/FC overflow and missing-run errors, strict SPRMs, exact patch/inverse, final owner and independent public readback remain | [DOC ABBA/profile/counter/memory record](changes/0056-doc-papx-containment-index.md), focused differential tests and complete DOC/harness suites | Accepted: already-open 512-paragraph list p50 -18.63%, mean -19.04%; one-edit/save p50 -8.01%; instructions -26.13%; allocations and peak heap flat |
| Source-backed Word97+ DOC paragraph splice | Additive `litchi-doc::body_text::source` positional owner -> validated `litchi-cfb` same-length `WordDocument` splice; no owned-body or archive-type dependency change | Immutable source snapshot, zero-based main-story paragraph selector, exact same-UTF-16-width replacement, exact no-op identity, source-checked commit/patch/inverse and candidate readback | Selector reads Unicode pieces in finite chunks and retains only the selected paragraph; complete artifact fingerprinting, CFB validation, candidate validation and publication scans remain mandatory; no runtime, global cache, lock or unsafe code | One uncompressed Unicode piece only; compressed/cross-piece, structural, field, drawing, revision, fast-save, PRM, encrypted/protected, macro/signed, ambiguous and malformed sources refuse; partial sinks remain typed | [Change 0105](changes/0105-doc-source-backed-paragraph-splice.md), 15 focused source-owner tests, crate-wide strict test/Clippy/rustdoc gates | Accepted for correctness and bounded selector coverage only; no end-to-end latency, physical-I/O/range, allocation/RSS, cold/high-latency, real-producer or broad DOC CRUD claim |
| PPT root snapshot CFB reuse | Private native `litchi-ppt` slide-order owner -> existing object-editor inspection | Immutable exact source and snapshot ownership unchanged | Reuses the package's already validated `OleFile`, removing one duplicate header/FAT/directory/MiniFAT/topology open; no cache, runtime, lock or global state | Independent stream resolution, Current User/live-persist validation, directory identity, document parse/round-trip, review history and public-reader verification remain | [PPT root-open ABBA/counter/memory record](changes/0024-ppt-slide-order-open-reuse.md), byte-ingress differential and complete PPT/harness suites | Accepted: repeated large root-open p50 -8.78%, mean -10.58%; allocation calls -5.01%; peak heap and uninstrumented RSS flat; cache-miss increase disclosed |
| PPT text-edit resolver reuse | Private native `litchi-ppt` semantic selector -> existing persisted-record editor | Exact-source transaction, patch/inverse and durable operation semantics unchanged | Holds selector result until the complete editor preflight succeeds, then reads the persisted record from that editor instead of reopening CFB; no cache/runtime/lock | Protected-source error precedence, fresh commit editor/source comparison, complete candidate reopen and independent readback remain | [Direct PPT ABBA/counter/memory record](changes/0026-ppt-text-edit-resolver-reuse.md), resolver differential/protection tests and complete PPT/harness suites | Accepted: large direct edit/save p50 -14.12%, mean -15.39%; allocation calls -3.53%; peak heap/RSS flat; minor-fault increase disclosed |
| PPT root text-publication adoption | Private native `litchi-ppt` text commit -> slide-order snapshot; no public type or dependency edge | Exact source, selected slide persist identity, immutable root state, patch/inverse and durable operations remain | Moves the validated output `Arc` into the root and carries only unchanged document/order/review facts; no cache/runtime/lock/global state | Text owner still performs fresh-editor source comparison, sole slide-record replacement, generic reopen and selected-shape readback; custom limits and structural paths retain complete root reopen | [Root adoption ABBA/counter/memory record](changes/0062-ppt-root-text-publication-adoption.md), full-reopen equivalence, lineage/persist rejection, custom-limit and composition tests | Accepted: large root one-shape edit/save p50 -18.59%, mean -17.83%; allocation calls -6.54%; peak heap and RSS flat |
| XLSX direct action-plan flattening | Rejected private `litchi-xlsx` worksheet-writer prototype | Deterministic effects, conflicts, inverse and final readback were unchanged in the prototype | Consumed address order through one reusable row vector instead of nested-map reconstruction | Exact untouched spans and all publication validation were retained; prototype fully reverted | [Rejected ABBA/allocation record](changes/0030-xlsx-action-plan-flattening-rejected.md) | Rejected: best formal p50 -1.61%; dense commit -0.27% p50 and mean interval crossed zero; allocation calls -0.0623%, peak heap flat |
| XLSX row-range index | Private immutable worksheet store | Cheap-to-share immutable state; selector behavior unchanged | Bounded compact row offsets over existing sorted store | No synthetic cells, normalization or changed malformed-input policy | [ABBA samples](results/abba-xlsx-range-before-a.json), parser/transaction tests | Accepted: p50 geomean -80.499%, mean geomean -79.962%; full scan +0.03% mean, first cell -1.31%, +17 allocations and +0.25% RSS |
| Bounded XLSX validated-store handoff | Private worksheet commit -> immutable target `OnceLock`; no dependency or API edge | Exact source/target style and shared-string lineages, worksheet URI/kind and published byte identity are required | Retains only stores at or below 4,096 cells and 1 MiB XML; no runtime, lock or global cache | Existing commit parse/style/change validation and later package integrity checks complete before adoption; later part rewrites refuse by `Arc` identity | [ABBA/counter/memory record](changes/0025-xlsx-validated-store-handoff.md), exact-identity, untouched-sheet and oversize tests | Accepted for bounded changed worksheets: medium commit + first read p50 -23.23%, allocation calls -21.01%, peak heap +4.29%; unrestricted dense-wide candidate rejected at +8.99% peak heap |
| XLSX row-visibility scalar-store reuse | Private existing-row visibility rewriter -> source-backed scalar snapshot; no public API or dependency edge | A private-field token borrows and identity-checks the exact source slice; the immutable `Arc<Store>` is reused only for direct `hidden`-attribute rewrites | Removes one complete scalar-cell parse per effective changed commit; retains bounded XML validation and a fresh row scan; no cache, runtime, lock or global state | Generic cell edits still fully parse; foreign-source tokens refuse; MCE, formula, macro, signature, relationship, protection, stale-source, exact no-op, patch/inverse and complete publication gates remain | [Change 0184](changes/0184-xlsx-row-visibility-store-reuse.md), focused sharing/parity/foreign-source tests, 768 XLSX library tests, strict Clippy and independent review | Accepted for the named generated existing-row workloads: all large commit statistics, large-batch total statistics and selected medium-batch commit statistics pass; unstable medium totals/hide-one plus allocation/RSS, physical-I/O, cold-cache, producer, structural-row and broad XLSX claims remain withheld |
| XLSX no-extension worksheet scan | Private raw worksheet parser; no API or dependency edge | Immutable worksheet/store and transaction/patch semantics unchanged | One bounded byte-token probe skips an otherwise empty extension traversal; no retained state, cache, runtime or lock | Any downstream rejection reruns the original x14ac collector before returning, preserving its error precedence; complete MCE, semantic parse, compact/readback, package integrity and security gates remain | [ABBA/memory/counter record](changes/0032-xlsx-no-extension-scan.md), direct x14ac/MCE/malformed tests and complete XLSX/harness suites | Accepted for `dyDescent`-free success paths: medium changed commit/save about -19% to -21%, dense 1% commit p50 -19.62%, allocation calls -25.24%, peak heap flat |
| DOCX direct paragraph selection and PPTX selected-scene reuse | Private format semantic scanners; no dependency edge | Immutable source lineage, selector ambiguity, patch resources and complete candidate readback retained | DOCX keeps the full bounded validation scan but avoids the collection; PPTX removes one scene parse per shape edit; DOCX final stream is forward-only | MCE/unknown XML, malformed trailing XML, rollback, dependency closure and full reopen tests unchanged | [Semantic ABBA/allocation record](changes/0010-docx-pptx-semantic-queries-and-edits.md), complete DOCX/PPTX library suites | Accepted: DOCX selector -4.72% p50; PPTX 1% edit/save -9.37%; PPTX one-edit neutral; PPTX allocation calls -11.67% |
| Coalesced DOCX direct-body paragraph replacements | Additive `litchi-docx` semantic transaction; no archive type or dependency edge | Strictly ordered unique selectors; ordinary durable operations, source checks, exact no-op sharing, inverse and failure atomicity retained | One bounded forward XML emission and one candidate parse replace one of each per selected paragraph; no cache, runtime, lock or unsafe code | Existing run/unknown XML rewrite rules plus complete selected-paragraph readback and package reopen | [DOCX batch ABBA/allocation record](changes/0012-docx-coalesced-paragraph-edits.md), exact scalar/durable/inverse/refusal tests, complete DOCX and harness suites | Accepted: large 100-edit save p50 -94.99% (19.97x), allocation calls -94.11%; scalar one-edit neutral; peak heap and uninstrumented RSS flat |
| ODS snapshot package reuse | Private `litchi-ods` snapshot/facade handoff; no public package type or dependency edge | Exact `Arc<[u8]>` source, immutable edit owner and source-checked patch semantics unchanged | Removes one package-sized byte clone and duplicate package parse; no cache, runtime or global state | Package/resource bounds run before the same complete facade readback; no-op bytes and changed-output reopen remain | [ODF ABBA/memory record](changes/0011-odf-semantic-baseline-and-ods-snapshot.md), complete all-feature ODS and harness suites | Accepted: large no-op p50 -11.78%, large one-cell edit/save -2.06%; peak heap flat and allocation-count increase disclosed |
| ODS row-local same-topology publication | Private `litchi-ods` worksheet XML owner; no public package or worksheet type change | Source rows remain exact spans; staged sheet identity, source-checked patching, no-op sharing and inverse semantics unchanged | Borrows changed sheets, serializes only changed logical rows, and copies untouched XML spans; no cache, runtime, lock or unsafe code | Structural cases fall back; changed opaque rows refuse; compactness, package reopen, snapshot parse and complete typed-sheet readback remain | [ODS ABBA/counter/memory record](changes/0018-ods-row-local-publication.md), opaque-row/fallback and complete ODS suites | Accepted: large one-cell edit/save p50 -9.54%, medium -7.22%; allocation calls -5.85%, peak heap -27.18%, all major counters improved |
| ODS unchanged-media publication and comparison | Additive family-neutral `litchi-odf-common::package` primitives -> private ODS effect validation; no archive type crosses the boundary | Exact source lineage, deterministic effects, immutable snapshot, patch/inverse and no-op sharing unchanged | Regenerates only compact `content.xml`; bounded raw-copy publication and exact local/central member comparison avoid unchanged payload compression/decompression; no cache/runtime/lock/global state | Encryption, signatures, manifest size metadata, unsupported ZIP layouts and unproved members fall back; compactness, complete reopen, media type/payload checks and logical recompression equality remain | [ODS media ABBA/memory/counter record](changes/0031-ods-unchanged-media-preservation.md), raw framing/no-op/signed/logical-fallback tests and complete common/ODS/harness suites | Accepted: media-rich one-cell edit/save p50 -4.73%, mean -5.73%, p95 -7.65%; peak heap -8.78%; no-media guard p50 -0.77% |
| Generic ODF verified raw content publication | Additive opt-in `litchi-odf-common::package` helper -> generic packaged ODF owner in `litchi-odt`; no physical archive type crosses the facade | Exact semantic no-ops return original bytes; changed content retains complete package reopen and the former logical writer's all-payload verification | Regenerates only `content.xml` and raw-copies eligible unchanged frames after opt-in payload verification; the ordinary shared ODT/ODS/ODP helper stays lazy, with no cache/runtime/lock/global state | Canonical MIME/root identity, local/central framing and descriptors are checked; aliases, signatures, encryption, stale sizes and unsupported layouts fall back or refuse | [Correctness/preservation record](changes/0101-generic-odf-verified-raw-publication.md), raw-frame/order, lazy-vs-verified, payload/descriptor corruption, no-op and fallback tests | Correctness-only integration; no latency, allocation, peak-memory/RSS or source-I/O claim before matched release evidence |
| ODS shared durable-patch blobs | Additive low-level `litchi-core::BlobBundle` ownership entry point -> private ODS patch builder; no archive type or dependency crosses the boundary | BlobId/wire identity, exact source lineage, limit precedence, reversible direction/order, stale checks and no-op behavior unchanged | Retains the patch's existing immutable source/target package Arcs and reuses their content addresses, removing two archive copies and two package SHA passes; no cache/runtime/lock/global state | Same per-blob/count/total bounds and target-before-source order; ZIP publication, compact audit, final package reopen and complete media/semantic verification remain | [ODS ABBA/profile/memory record](changes/0054-ods-shared-durable-patch-blobs.md), Arc/wire/boundary tests and complete Core/ODS/harness suites | Accepted: media-rich one-cell edit/save p50 -8.80%, mean -9.07%, p95 -13.85%; 33.58 MB copy site absent; peak heap -1.92%, RSS flat |
| ODS provenance-bearing row-splice publication | Additive low-level ODF common checked-splice publisher -> private ODS worksheet/package handoff; no archive type or dependency crosses the ODS public boundary | Exact package provenance, checked source ranges, compact fragments, typed worksheet identity, patch/inverse and no-op behavior remain | Carries existing row-range proofs through raw ZIP publication instead of flattening them and falling back to recompression of unchanged media; no cache/runtime/lock/global state | Foreign provenance and unexpected assembled content refuse; signatures, encryption-sensitive and unsupported ZIP layouts retain the established rebuild/signature-stripping fallback; compact audit, complete reopen and media verification remain | [ODS ABBA/profile/counter/memory record](changes/0057-ods-row-splice-raw-publication.md), raw-member/provenance/signed-fallback tests and complete common/ODS/harness suites | Accepted: media-rich one-cell edit/save p50 -74.16%, mean -74.17%, p95 -74.11%; instructions -69.04%; peak heap and RSS flat |
| ODS shared worksheet archive handoff | Private unified edit -> worksheet snapshot/patch -> ODF package owner; no public type or dependency edge | Exact source bytes, typed sheets, patch/inverse, failure atomicity and durable unified lineage remain | Shares/moves one `Arc<Vec<u8>>` allocation across nested parsing, commit readback and candidate validation; unexpected sharing falls back to an exact clone; no cache/runtime/lock/global state | Full package/worksheet parse, row-splice provenance, compact audit, effects, security policy, durable patch construction, final reopen and media readback remain | [ODS balanced latency/profile/counter/memory record](changes/0068-ods-shared-worksheet-archive-handoff.md), allocation-identity/failure-rollback tests, complete ODS/harness suites | Accepted: media-rich one-cell edit/save p50 -21.32%, mean -21.30%, p95 -21.15%; peak heap -22.03%, RSS -20.57%; large ordinary guards within 1.6% |
| ODP content-only unchanged-media publication | Existing family-neutral `litchi-odf-common::package` primitive -> private ODP content owner; no archive type or dependency crosses the boundary | Exact source lineage, deterministic effects, immutable snapshot, patch/inverse, stale-source refusal and no-op sharing unchanged | Regenerates only a checked single-splice `content.xml` and raw-copies eligible unchanged members; resource additions retain logical rebuild; no cache/runtime/lock/global state | Encryption, signatures, manifest size metadata and unsupported ZIP layouts fall back; compactness, complete presentation/rich-content reopen and exact media type/payload checks remain | [ODP media ABBA/memory/counter record](changes/0034-odp-unchanged-media-preservation.md), raw-member integration/harness proofs and complete ODP/common fallback suites | Accepted: media-rich text-box edit/save p50 -94.44%, mean -94.43%, p95 -94.29%; allocation calls +0.52%; peak heap/RSS flat; ordinary no-op/open/edit guards within thresholds |
| ODP indexed slide selector | Private compile-time-specialized `litchi-odp` parser projection; no public type or dependency edge | Immutable query result, slide ordering/index and full-list semantics unchanged | Retains text and completed shapes only for the target page; no cache, runtime, lock, global state or early XML termination | All styles and content scan through EOF with existing namespace, attribute, shape, media/link, animation and global-limit validation; absent indices return only after successful validation | [ODP selector latency/allocation record](changes/0049-odp-indexed-slide-selector.md), parser/public differential and late-failure tests, complete ODP/harness suites | Accepted: large middle-slide p50 -4.09%, mean -4.20%, p95 -5.18%; allocation calls -3.86%; peak heap/RSS flat; tiny/list/save guards neutral |
| ODP snapshot slide-projection reuse | Private `litchi-odp` snapshot -> mutable staging handoff; no public type or dependency edge | Snapshot bytes and slides remain one immutable source-bound identity; staging deep-clones mutable and comparison values; exact no-op, patch/inverse and deterministic changed commit semantics remain | Removes one duplicate complete slide traversal; no retained cache, runtime, lock, executor, global state or unsafe code | Package/security reopen, settings/declarations/page metadata and raw page-coverage validation remain; changed publication, compact audit and independent complete reopen/readback are unchanged | [ODP latency/profile/counter/memory record](changes/0060-odp-snapshot-slide-projection-reuse.md), focused isolation/coverage tests and complete ODP/harness suites | Accepted: large exact-no-op p50 -59.96%, changed edit/save p50 -20.78%; allocations -20.13%; peak heap/RSS flat; read-only central guards neutral and noisy p99 disclosed |
| ODP final slide-snapshot handoff | Private `litchi-odp` changed slide candidate -> final snapshot; no public type or dependency edge | The candidate and final bytes are the same immutable source-bound allocation only for an exact slide-only commit; patch/inverse, no-op and compound-domain semantics remain | Moves the already required parsed projection instead of parsing it again; no retained cache, runtime, lock, executor, global state or unsafe code | The independent final package reopen, raw/compact XML audits, staged-media verification and every auxiliary-domain readback remain before adoption; compound commits retain the ordinary final parse | [ODP latency/profile/counter/memory record](changes/0065-odp-final-snapshot-handoff.md), compact/media/compound transaction tests and complete ODP/harness suites | Accepted: large one-slide edit/save p50 -32.35%, mean -32.92%; allocations -16.71%; peak heap/RSS flat; ineligible no-op/media guards improve and read-only central guards remain within 3% |
| ODS adaptive logical-cell locator | Private `litchi-ods` immutable facade; no public model or dependency edge | Borrowed `CellView` and original `Cell` identity retained; every facade replacement starts cold | Atomic threshold plus `OnceLock`; compact direct-run descriptors, repeated-run `u32` endpoints, checked 4 MiB cap and fallible permanent linear fallback | Builds only after complete open; no XML or package access; missing/repeated/boundary behavior is differential-tested and malformed-input order is unchanged | [ODS lookup ABBA/profile/memory record](changes/0027-ods-adaptive-cell-locator.md), repeat/budget/concurrency/replacement tests and complete ODS/harness suites | Accepted: large sweep p50 -81.74%, full text -52.65%; dense retained index 3,216 B; peak heap/RSS flat |
| ODT existing-document snapshot byte sharing | Private `litchi-odt` document/snapshot handoff; no public package type or dependency edge | Immutable exact source, edit owner, reversible/durable patches and exact no-op semantics unchanged | Clones one private `Arc` instead of allocating/copying the package; no cache, runtime, lock or global state | Same transaction size bound; changed candidate audit/reopen and signed/encrypted refusal unchanged; private pointer-identity regression | [ODT ABBA/allocation record](changes/0014-odt-shared-snapshot-bytes.md), complete all-feature ODT and harness suites | Accepted: medium/large no-op p50 -27.05% / -18.51%; exactly two allocations removed per snapshot; open and changed-save guards within 3%, peak heap/RSS flat |
| ODT direct snapshot byte sharing | Private `litchi-odt` transaction/document handoff; no public package type or dependency edge | Consumed exact bytes become the immutable snapshot source; deterministic edit, patch/inverse, stale-base and no-op semantics unchanged | Adopts validation's private package Arc and clones it for staging rehydration, removing two archive copies; no cache, runtime, lock or global state | Same size bound and complete ODT parses; compact audit, signed/encrypted refusal, final reopen/readback and raw-media publication remain; input-buffer and Arc pointer identity are tested | [Direct snapshot ABBA/memory/counter record](changes/0038-odt-direct-snapshot-sharing.md), complete all-feature ODT suite | Accepted: 16 MiB media edit/save p50 -75.84%, mean -73.84%; eight allocation calls removed per full harness iteration; peak heap/RSS flat |
| ODT compact-audit package sharing | Private `litchi-odt` transaction/audit handoff; no public package type or dependency edge | Exact source lineage, deterministic edit, patch/inverse, stale-base and no-op semantics unchanged | Clones the validated predecessor package Arc and borrows the validated candidate package, removing three archive-sized copies; no cache, runtime, lock or global state | Complete archive/manifest and compact XML/splice audit, signed/encrypted refusal, final materialization, reopen/readback and raw-media publication remain; predecessor Arc identity is tested | [Compact-audit ABBA/memory/counter record](changes/0041-odt-compact-audit-package-sharing.md), complete all-feature ODT/common suites | Accepted: 16 MiB media edit/save p50 -30.44%, mean -31.36%, p95 -32.41%; allocations -0.57%, peak heap/RSS flat; exact no-op +39 ns p50 disclosed |
| ODT envelope-classification package sharing | Private `litchi-odt` snapshot/security-classifier handoff; no public package type or dependency edge | Exact source lineage, deterministic edit, patch/inverse, stale-base and no-op semantics unchanged | Clones the already validated snapshot package Arc, removing one archive-sized copy and temporary owner; no cache, runtime, lock or global state | Package size and ZIP validation, manifest encryption metadata, signature-Part checks, compact audit, final materialization and reopen/readback remain; package Arc identity is tested | [Envelope-sharing ABBA/memory/counter record](changes/0042-odt-envelope-package-sharing.md), complete all-feature ODT/common suites | Accepted: 16 MiB media edit/save p50 -11.40%, mean -11.95%, p95 -12.19%; two allocations/commit removed, peak heap/RSS flat; large exact no-op +152 ns p50 disclosed |
| ODT final changed-result byte handoff | Private `litchi-odt` validated-document/snapshot handoff; no public package type or dependency edge | Exact no-op source sharing, changed lineage, deterministic effects, patch/inverse and stale-base semantics unchanged | Clones the final validated document's private package Arc into a byte-only snapshot, removing one archive copy and redundant parse; no retained graph, cache, runtime, lock or global state | Same 64 MiB bound, per-operation compact audit, signed/encrypted refusal and one fresh complete final reopen remain; changed Arc identity and independent semantic readback are tested | [Final-byte ABBA/profile/memory record](changes/0052-odt-final-result-byte-handoff.md), complete all-feature ODT/harness suites | Accepted: 16 MiB media edit/save p50 -22.74%, mean -22.56%, p95 -21.48%; allocation calls -3.46%; peak heap/RSS flat; medium one-paragraph within 3% p50/mean |
| ODT consuming full-text block ownership | Private `litchi-odt` parser/element codec mode; public structured block API unchanged | Immutable query only; no document, transaction, patch or publication state | Moves each parser-created validated string through the owned full-text iterator, removing two intermediate strings per block; no cache, runtime, lock or global state | Same namespace/depth/text limits, exclusions, order and newline semantics; malformed input still completes the same bounded scan | [ODT ABBA/allocation record](changes/0023-odt-full-text-owned-blocks.md), focused ownership/order tests, complete all-feature ODT and harness suites | Accepted: repeated large full-text p50 -3.25%, mean -4.81%; allocation calls -15.48%; peak heap/RSS flat; unchanged open +3.94% p50/+4.17% mean and +10.95% p99 disclosed |
| ODT repeated source-backed text cache | Private `litchi-odt::SourceBackedDocument`; no public type, method, dependency edge, executor or global state | Immutable source owner and owned `String` return contract remain; every parse, publication and hit retains source-version checks and error precedence | Saturating atomic call threshold plus `OnceLock<Option<String>>`; the first invocation below threshold stays uncached, the first later successful parse may fallibly retain one projection capped at 16 MiB, and later hits clone it; oversize/allocation refusal becomes terminal non-retaining state while parse errors retry | Existing complete text parser, namespace/depth/text bounds, malformed behavior, media/package validation and source-change refusal remain; concurrent construction is safe but may duplicate parsing before publication, and returned strings are independent | [Change 0180](changes/0180-odt-source-text-cache.md), [summary](results/odt-text-cache-0180-summary.json), focused cache/freshness/failure/concurrency tests, complete ODT/harness gates and independent reviews | Accepted for the exact prepared four-call generated workload: complete projection phases `4 -> 2`, p50 -47.01% to -50.95% and mean -46.83% to -51.29% across two balanced cycles. p95/p99, allocation/RSS, physical I/O, open/single-call, producer, generic ODF and broad CRUD claims are withheld |
| ODT content-only unchanged-media publication | Existing family-neutral `litchi-odf-common::package` primitive -> private ODT transaction owner; no public archive type or dependency edge | Exact source lineage, deterministic effects, immutable snapshot, patch/inverse, stale-source refusal and exact no-op sharing unchanged | Regenerates only compact `content.xml` and raw-copies eligible unchanged members; ineligible/mixed edits and XML above the common 16 MiB optimization bound use the established rebuild; no cache/runtime/lock/global state | Encryption, signatures, manifest size metadata and unsupported ZIP layouts fall back; compactness, complete document reopen, paragraph/media readback and raw unchanged-member identity remain | [ODT media ABBA/memory/counter record](changes/0035-odt-content-only-paragraph-publication.md), raw-member and over-limit fallback regressions, complete ODT/common/harness suites | Accepted: media-rich paragraph edit/save p50 -95.58%, mean -95.63%, p95 -95.43%; allocation calls -6.71%; peak heap flat, RSS -0.59%; ordinary ODT guards improve |
| ODT content-only line-break publication | Existing ODT line-break transaction -> accepted family-neutral content-only publisher; no public type, archive abstraction or dependency edge | One ordered durable operation and result, exact source lineage, deterministic output, immutable snapshot, patch/inverse and stale-source refusal unchanged | Regenerates compact `content.xml` and raw-copies eligible unchanged core/media members instead of rebuilding all resources; no cache/runtime/lock/global state | Complete reopen/text/media checks, exact output digest and raw-member identity remain; other/mixed/oversized, signed/encrypted and unsupported layouts retain established behavior | [ODT line-break ABBA/memory/counter record](changes/0071-odt-content-only-line-break-publication.md), packaged transaction and harness raw-member regressions, complete ODT/common/harness suites | Accepted: media-rich line-break edit/save p50 -98.17% (54.59x), mean -98.16%, p95 -98.08%; instructions -78.34%; allocation calls -6.90%; peak heap/RSS flat |
| ODT content-only inline-run publication | Existing ODT run transaction -> accepted family-neutral content-only publisher; private no-op/changed dispatch boundary; no public type, archive abstraction or dependency edge | Styled/unstyled ordered operation and result, exact no-op source sharing, source lineage, deterministic output, immutable snapshot, patch/inverse and stale-source refusal unchanged | Regenerates compact `content.xml` and raw-copies eligible unchanged core/media members; exact no-op avoids the changed-path stack frame; no cache/runtime/lock/global state | Complete reopen/text/style/media checks, exact output digest and raw-member identity remain; changed commits retain the previous validation body, while other/mixed/oversized, signed/encrypted and unsupported layouts retain established behavior | [ODT inline-run ABBA/memory/counter record](changes/0072-odt-content-only-run-publication.md), packaged transaction and harness raw-member regressions, complete ODT/common/harness suites | Accepted: media-rich append-run edit/save p50 -98.39% (62.01x), mean -98.38%, p95 -98.27%; instructions -78.48%; allocation calls -7.00%; peak heap/RSS flat; exact no-op guards improve |
| ODT content-only hyperlink publication | Existing ODT hyperlink transaction -> accepted family-neutral content-only publisher; no public type, archive abstraction or dependency edge | Ordered inert hyperlink text/URL, exact source lineage, deterministic output, immutable snapshot, patch/inverse and stale-source refusal unchanged | Regenerates compact `content.xml` and raw-copies eligible unchanged core/media members; no relationship or external fetch is introduced | Complete reopen/text/URL/media checks, exact output digest and raw-member identity remain; ineligible/mixed/oversized, signed/encrypted and unsupported layouts retain established behavior | [ODT hyperlink ABBA/memory/counter record](changes/0074-odt-content-only-hyperlink-publication.md), packaged transaction and harness regressions, complete ODT/common/harness suites | Accepted: media-rich append-hyperlink edit/save p50 -98.20% (55.52x), mean -98.18%, p95 -98.07%; instructions -78.34%; allocation calls -6.99%; peak heap/RSS flat |
| ODT content-only structural paragraph publication | Existing ODT insert/remove paragraph transactions -> accepted family-neutral content-only publisher; no public type, archive abstraction or dependency edge | Ordered bounded plain-text insertion and exact target removal, source lineage, deterministic output, immutable snapshot, patch/inverse and stale-source refusal unchanged | Regenerates compact `content.xml` and raw-copies eligible unchanged members; removal performs no resource GC; no cache/runtime/lock/global state | Complete paragraph order/media/manifest reopen, exact digests and raw-member identity remain; ineligible/mixed/oversized, signed/encrypted and unsupported layouts retain established behavior | [ODT structural ABBA/memory/counter record](changes/0075-odt-structural-paragraph-publication.md), packaged transaction and harness regressions, complete ODT/common/harness suites | Accepted: media-rich insert/remove p50 -98.20%/-98.27% (55.55x/57.86x); instructions -82.14%; allocation calls -8.47%; peak heap/RSS flat |
| Coalesced ODT paragraph publication | Private packaged-transaction dispatch; no API, patch vocabulary, archive type or dependency edge | Ordered model-backed insert/replace/remove/run/hyperlink operations retain one result each, immutable failure atomicity, durable replay/inverse and exact source checks; inline append before later plain/topology work remains a reopen boundary | One mutable candidate, content publication, reopen and compact audit replace one of each per eligible contiguous run; XML-only, move and other package-domain operations keep their prior path; no cache/runtime/lock/unsafe code | Complete compact audit and package reopen remain; media/core members stay raw exact; over-16-MiB content uses the existing rebuild; signed/encrypted envelope refusal is unchanged | [ODT batch ABBA/profile record](changes/0045-odt-coalesced-paragraph-publication.md), scalar/mixed ordering, byte-identity, durable/refusal/media/fallback tests and complete ODT/harness suites | Accepted measurements remain scoped to consecutive replacements: large 100-edit/save p50 -98.28% (58.05x), allocations -96.13%; medium two-edit -27.62%; scalar guard neutral; peak heap and uninstrumented RSS flat; tool-inclusive RSS +9.93% disclosed. Mixed-run performance evidence remains pending |
| ODT mixed model-content publication evidence | Existing private packaged-transaction dispatch and model-backed operation result contract; no public API, patch vocabulary, archive type or dependency edge | Medium/large deterministic shapes retain 80/320 logical results, exact source lineage and the candidate's one staged publication; scalar and batch output/logical hashes agree per shape | Timed comparison isolates 49/193 repeated scalar publications versus one staged candidate publication; no cache/runtime/lock/unsafe state is introduced by the evidence case | Raw reports retain per-leg hashes and counters; source preparation, reopen/lifecycle/security/limits, I/O, serialization, allocation/RSS, and physical cold behavior are outside timing | [ODT mixed publication record](changes/0104-odt-mixed-model-publication-evidence.md), four raw reports, compact summary, and harness semantic/hash gates | Matched release A/B p50 reduction is 96.8685%–96.8695% medium and 99.2289%–99.2381% large; accepted only as repeated-publication versus one-transaction evidence, not general ODT speed or resource evidence |
| RTF retained text length, chunked ASCII emission, text-only edit, parser-state specialization, transport batching, byte-delimiter scanning and retained ordinary-body range | Private `litchi-rtf` parser/model/writer/edit state; no dependency edge or facade type change | Lazy cached full text, immutable source identity, checked edits, revision ranges, durable operations and exact no-op sharing retained | One final text allocation/pass; contiguous ASCII sink writes; text-only commits skip unused property vectors/scans; ordinary text flushes copy only required fields; ASCII source tokens enter parser buffers in one extension; ordinary source text finds five ASCII delimiters in one byte pass; direct ASCII commits reuse a range proven during initial structural preflight instead of cloning and lexing the source again | Escape spelling, UTF-8 boundaries/spans, byte-valued fallback, code pages, revision/deletion metadata, opaque refusal, limits, candidate parse/readback, inverse/stale checks and partial-sink errors retained; empty/ambiguous/binary/non-ASCII/LZFu inputs keep the established locator/refusal; CP-1252/LZFu/watermark exact no-op and native `relsize` readback are gated | [RTF text-path record](changes/0013-rtf-semantic-baseline-and-text-paths.md), [parser-state record](changes/0019-rtf-parser-state-specialization.md), [transport ABBA/counter record](changes/0020-rtf-ascii-transport-batching.md), [variant/native coverage record](changes/0029-rtf-transport-and-producer-coverage.md), [delimiter ABBA/counter record](changes/0040-rtf-byte-delimiter-scanning.md), [retained body-range record](changes/0048-rtf-retained-body-source-span.md), native facade/writer/transaction/fuzz suites and CI matrix | Accepted: large full-text -27.08%; parser-state open -20.09% and edit/save -11.54%; transport open -26.67% and edit/save -6.26%; byte delimiter open -17.23% and edit/save -14.65%; retained body range edit/save -10.72% p50/-10.11% mean; prepared LZFu no-op exception disclosed |
| RTF bounded body style-block reservation | Private `litchi-rtf` parser capacity hint; no public model, archive type or dependency edge | Block values/order, formatting, paragraph state, revision ranges, exact source/no-op identity, edits and writer bytes unchanged | Existing structural pass counts root text; first retained block performs one fallible exact reserve under source/token/16 MiB bounds; no cache/runtime/lock/global state | Sources below 64 KiB and root table/deletion controls keep lazy growth; failure retries one required block; complete lexer/parser limits, candidate parse/readback, variants, transactions and fuzz gates remain | [RTF reservation latency/profile/memory record](changes/0055-rtf-body-block-reservation.md), focused bound/table/deletion/failure tests and complete RTF/harness suites | Accepted: large open p50 -21.17%, mean -21.00%, p95 -21.04%; edit/save p50 -1.46%; vector allocations 264 -> 22; peak heap -29.73%; medium plain/CP-1252 +0.49%/+2.84% p50 disclosed |
| RTF retained story-length handoff | Private `litchi-rtf` model scalar -> borrowed public story internals; no method, archive type or dependency edge | Exact parser-derived UTF-8 length, block/boundary order, formatting, immutable snapshots, edits, patches and source bytes unchanged | Carries the already retained byte count into `Story`; `len`, `is_empty`, paragraph and inline endpoint setup become constant-time; no allocation/cache/runtime/lock/global state | Complete parser/boundary/UTF-8/limit validation remains; full queries still traverse their semantic ranges; fragmented/edit, CP-1252, LZFu and watermark proofs plus open/full-text/save/no-op guards remain | [RTF query latency/profile/memory record](changes/0064-rtf-retained-story-length.md), focused length/edit/transport tests and complete RTF suites | Accepted: large paragraph-list p50 -15.04%/mean -13.71%; middle paragraph p50 -27.19%/mean -25.23%; allocations, peak heap and RSS flat; central guards within 5% |
| RTF sparse paragraph selection | Private `Paragraphs` iterator method override; no facade, model, archive type or dependency edge | Paragraph identity/order, inline ranges, formatting, fused exhaustion, resumed traversal, immutable snapshots, edits, patches and source bytes unchanged | Scans retained boundary descriptors and constructs only the selected paragraph instead of invoking `next` for every discarded prefix item; linear, allocation-free, no index/cache/runtime/lock/global state | Existing parser/boundary/UTF-8/limit validation and selected-view location checks remain; differential structural/formatting tests, fuzz traversal, CP-1252/LZFu/watermark proofs and open/list/full-text/save/edit guards remain | [RTF sparse-query latency/profile/memory record](changes/0066-rtf-sparse-paragraph-nth.md), focused iterator equivalence, fuzz target, complete RTF and harness suites | Accepted: large middle-paragraph p50 -47.87%/mean -47.95%/p95 -49.42%; allocations and peak heap exact-flat, RSS flat; all guards within policy |
| RTF decoded-body ownership handoff | Rejected private parser/model ownership prototypes; no production or dependency edge remains | Text, UTF-8 offsets, revision ranges, exact source and final owned model were unchanged in the prototypes | Broad handoff removed one arena and one final copy for decoder-owned text but moved borrowed ASCII allocation into the parser loop; owned-only refinements added a discriminant path | Every prototype was fully reverted; a retained malformed Shift-JIS tail test proves lossy text and exact immutable output | [Rejected ABBA/profile/memory record](changes/0043-rtf-decoded-body-ownership-rejected.md), raw JSON and digest manifest | Rejected: broad CP-1252 open -3.08% p50 but plain large open +25.53%; owned-only refinements -1.41% and +1.02% p50 |
| Immutable-owned CFB atomic save | Private CFB plan provenance produced only by `SharedOleFile::open_owned(Arc<[u8]>)`; no dependency inversion or facade archive type | Existing immutable plan, exact source/target fingerprints, candidate reopen and report semantics unchanged; generic sources keep both mutation fences | Owned save omits only two redundant complete fingerprints; the full 64 KiB emission source/target hashes, exact progress, flush, fsync, sibling rename and parent sync remain | Protected-container refusal, bounds, hostile generic `ReadAt`, partial output, destination preservation, semantic reopen and untouched-stream tests remain | [Change 0175](changes/0175-cfb-owned-atomic-save.md), focused CFB/OLE-common tests, strict checks, clean raw A/B/B/A reports and independent review | Deterministic reduction accepted: 33,826,816 logical bytes and 34 fingerprint reads on the 16.9 MiB corpus. Latency withheld because control drift exceeds 5%; no physical-I/O, allocation/RSS, cold-device or broad semantic claim |
| Rejected ODS/XLSX retained-readback handoffs | Private format/common prototypes only; both exact production diffs reverted | Existing snapshots, edits, patches, no-op and source lineage restored exactly | ODS avoided one payload read but added proof hashing; XLSX avoided one duplicate parse; neither met measured usefulness | Security/compatibility reviews and exact output gates passed before reversion; no residual production path remains | [Change 0176](changes/0176-rejected-odf-xlsx-reuse.md) and clean raw A/B/B/A reports | Rejected and reverted: ODS p50 regressed 1.63%-2.83%; XLSX paired directions disagreed (-4.81%/+1.99%) |
| ODS source-backed existing-cell release evidence | Existing additive `litchi-ods` source transaction and common source publisher; harness-only instrumentation adds no dependency or public type | Existing immutable snapshot, bounded cell selection, deterministic effects, patch/inverse, exact source lineage and failure atomicity remain | Timed path stays uninstrumented; aligned lifecycle/phase vectors and an untimed logical `ReadAt` replay expose the established sparse path without changing production | Semantic/media/hash/raw-member/sink gates remain; source-only fields are nullable, while stale/security/formula/repeated-row/transaction-bound contracts remain production-test evidence outside the selector | [Change 0177](changes/0177-ods-source-cell-release-evidence.md), clean raw A/B/B/A reports, strict harness checks and independent evidence review | Accepted only for fixed one-cell complete lifecycle: p50 -75.03%/-74.27% with p50/mean/p95/p99 stability passing. The 21-cell 1% latency result is withheld; no physical-I/O, allocation/RSS, cold-cache or broad ODS claim |
| XLS immutable source policy reuse | Private `litchi-xls::cell_values::Snapshot` facts and existing plan-only numeric API; no public type, dependency edge, executor, lock, global state, or unsafe code | Immutable source lineage, exact fixed-width selectors, no-op refusal, distinct workbook/worksheet protection diagnostics, source/target fingerprints and forward-only plan contract remain | One source `Workbook` policy reopen becomes three compact facts captured from the complete snapshot validation; independent target semantic reopen and every CFB/publication pass remain | Signed/encrypted/DRM, CFB VBA storage, workbook/worksheet protection, macro, stale/foreign, topology/CLSID, partial-sink and numeric readback gates remain | [Change 0181](changes/0181-xls-source-policy-reuse.md), [summary](results/xls-source-policy-0181-summary.json), focused/full XLS gates and two independent reviews | Accepted for exact Number total and commit p50/mean/p95/p99; RK/MulRK latency and publication/I/O/allocation/RSS/cold/atomic-save/broad-XLS claims withheld |
| SIMD byte scanning | ADR-permitted low-level owner only | No public semantic effect | Runtime detection outside loop; scalar fallback; minimum size threshold | Exact malformed-input differential behavior; no weakened unsafe policy | Assembly inspection, x86_64/aarch64 differential tests, end-to-end materiality | Deferred pending profile |

| OPC source-backed reader ingress | Public litchi-opc SourceBackedPackage reader ingress; no facade API or iWork edge | Bounded logical input is consumed once; typed maximum-limit rejection with actual = maximum + 1 asserted, and existing source-backed ownership semantics remain | Reduces compressed-plus-all-decompressed eager retention to one compressed buffer plus indexed metadata and deferred selected payloads; unmanaged, with no duplicate compressed archive copy | ReadLimits and try_reserve_exact bound logical input/local admission work, not total RSS or aggregate concurrent opens; exact-limit typed failure, zero cold payload loads at open, one selected cold/successful load, and arbitrary blocking Read cancellation limitation are explicit | [Change 0345](changes/0345-opc-source-backed-reader-ingress.md), 4/4 focused tests including reader_ingress_retries_one_interrupted_read and reader_ingress_rejects_invalid_read_count_without_panicking, four owner lib checks, one Cargo process/job on a dedicated disk target | Accepted only as structural evidence; performance_claim: none; no RSS or before/after latency claim; callers needing tighter host memory must lower max_input_bytes, serialize opens, and account aggregate process memory externally; no facade/iWork claim |

## Hard review checks

For each implementation, reviewers must be able to answer all of these with
linked evidence:

1. Which named CRUD scenarios and corpus items exercise the change?
2. Which work, allocation, I/O, decompression, copy, serialization, validation,
   or synchronization is removed?
3. What source/version, budget, cancellation, and failure-atomicity checks
   remain?
4. What proves untouched unknown content, signatures, protection/encryption
   boundaries, deterministic output, and exact no-op behavior remain correct?
5. Does any public facade gain an archive type, physical identifier, lock,
   executor, runtime, source generic, or unsafe storage?
6. Does any new dependency violate ADR 0002 or ADR 0024?
7. What are the individual before/after distributions, uncertainty, allocation
   and RSS changes, copied/decompressed/recompressed bytes, regressions, and
   remaining limitations?
8. If the result is not useful, was the speculative complexity removed?

## Change 0346 update

Change 0346 adds no production code, dependency edge, public type, source
fence, freshness policy, unsafe code or iWork path. The six control reports and
the standalone `std` versus `parking_lot` probe are harness/evidence-only. The
probe's 0.36-0.40% modeled whole-operation gain is below the usefulness
threshold, so no lock substitution candidate was retained and
`performance_claim: none`. The next XLS performance batch must target a
different design; the unchanged 0279 freshness candidate remains rejected.

## Change 0347 update

Change 0347 adds no production code, dependency edge, public type, or iWork
path. It records a bounded XLSX cell-values harness repair: eager/source typed
semantic equality, exact source lineage, no-op bytes, deterministic output,
namespace-aware URI matching, finite XML normalization bounds, duplicate ZIP
name refusal, exact changed-member sets, and pinned 17-member numeric corpus
identities remain explicit. Only direct `calcPr` invalidation and the exact
optional calc-chain relationship/part/content-type removal closure are
accepted; other raw members remain identity-checked. Formula/date workloads
are excluded. `performance_claim: none`; direct medium one-edit and the
24-row serialized ABBA v1 smoke have zero failure rows and complete rows, but
timing gates and claim authorization are false. No production optimization,
RSS, allocation, physical-I/O, or broad XLSX claim follows.

| Stored ZIP borrowed validation | Private `soapberry-zip` immutable-slice Store validation; no generic `ReadAt` or remote/file borrowing edge | Complete local/central metadata, descriptor CRC/size forms, local ZIP64-extra provenance, encryption/overlap/duplicate safety, and strict nonempty zero-CRC refusal; pointer identity and existing fallbacks remain | Validated Store payloads borrow without cache/materialization charge; ZIP64 EOCD, Deflate, and positional sources remain owned/streaming; concurrency unchanged | `focused borrowed 10/10; full soapberry-zip lib 280/280` under `CARGO_BUILD_JOBS=1`, `test-threads=1`, and an 8 GiB process ceiling; downstream `litchi-opc borrowed 12/12`, not the full suite; package format check passed; raw `get_entry_borrowed` remains unverified and requires a verifier | [Change 0348](changes/0348-stored-zip-borrow-validation.md) | Correctness only; `performance_claim: none`; no latency/RSS/copy claim; stored OOXML representativeness is weak |
| PPTX source-probe error and fallback admission | Private `litchi` PPTX bytes probe and `FileSource` path fallback; no public detector/type or dependency edge | Typed `OpcError` plus terminal `OtherOoxml`/`DisabledOtherOoxml`; exact `SourceVersion`, caller `max_input_bytes`, part limits, freshness, and cancellation semantics remain | Genuine non-ZIP/short/missing `[Content_Types].xml` fallback reclaims the original `Vec`; same-source bounded `Bytes` replaces pathname re-open/unbounded `fs::read`; hard ZIP/OPC/classifier outcomes do not eager/retry PPTX or ODP | Input/part-limit, malformed-ZIP, missing-manifest ownership, wrong-family/polyglot precedence, extensionless bounded path, reserved-namespace, freshness/cancellation regressions; public `DetectedFormat`/eager and ordinary ODP native-owner handoff remain | [Change 0355](changes/0355-pptx-source-probe-fallback.md), `litchi` `pptx` check, combined `pptx,odp` lib tests `48/48`, formatter under a constrained one-target run | Correctness/ownership only; `performance_claim: none`; no speed/RSS/OOM-prevention claim; DOCX/non-Unix/ODT/ODP/public eager/selected-part residuals remain |
## Change 0356 compliance update

Change 0356 remains within the accepted topology and public-layer boundaries:
DOCX source ownership stays in `litchi`, OPC error/resource ownership stays in
`litchi-opc`/`soapberry-zip`, and no public archive type, physical identifier,
runtime handle, lock, executor, source generic, or `DetectedFormat` change is
introduced. The single-source path and final length fence preserve ADR 0003
freshness and failure atomicity; typed limits, allocation, I/O, cancellation,
and execution precedence preserve the panic-free and bounded-ingress rules.
Terminal wrong-family outcomes and lossless `OtherOoxml` preservation follow
the lossless/unsupported-content policy, while `UnsupportedPreservation` is the
only overlay-unavailable translation. No dependency edge, unsafe high-level
code, or ADR exception/supersession is required. Validation and residuals are
recorded in [Change 0356](changes/0356-docx-source-path-and-opc-errors.md);
`performance_claim: none`. The caller-sized physical result buffer uses typed
fallible reservation and releases the part reservation on admission failure;
this is correctness/resource safety only, with no performance or OOM claim.

## Change 0357 compliance update

Change 0357 remains within the accepted topology and public-layer boundaries:
Workbook and Presentation continue to own their format-specific source
arbitration, while ODF catalog budgeting remains in `litchi-odf-common`; no
public archive type, physical identifier, runtime handle, lock, executor,
source generic, or `DetectedFormat` surface is introduced. The one-source
PPTX/ODP/PPT path and freshness fence preserve ADR 0003 source identity and
failure atomicity. Caller-limited OOXML/uncertain candidates are capped by a
finite neutral 2 GiB ceiling, while ordinary ODP/ODS and lower-family/generic
fallbacks use the same checked neutral budget. Terminal wrong-family OOXML
and lossless ODF ownership preserve the unsupported-content policy. The ODF
catalog helper keeps input, compressed, entry, and total ceilings explicit;
there is no dependency edge, unsafe high-level code, or ADR exception.

Validation and residuals are recorded in [Change 0357](changes/0357-workbook-presentation-two-ceiling-policy.md);
`performance_claim: none`. The serial constrained evidence includes ODF
detection `15/15` with `260` filtered, catalog arbitration `6/6`, facade
`pptx,odp,ppt` `82/82`, facade `ods,xlsx` `84/84`, and passing quiet
`pptx`, `odp,ppt`, `odp`, `ppt`, and `xls,xlsx` checks. Two initial
`Arc<FileSource>` to `Arc<dyn ReadAt>` coercion errors were corrected before
final checks. One 8 GiB process ceiling, one Cargo job, one disk target, and
one test thread were used; the target's final/peak observed footprint was 1.3
GiB, host availability approximately 14 GiB with 133 GiB disk free, and swap
was exhausted. No parallel build or OOM occurred, and no speed, RSS,
allocation, constant-memory, or OOM-prevention claim follows. The eager
public detector, materializing neutral fallback, flat ODF MIME decode,
infallible Presentation aggregation, portable identity, native PPT mutation
coverage, Current User/Workbook classifier mismatch, prepared ODP reparse,
OPC case lookup, and selected-Part materialization remain residual.

## Change 0358 compliance update

Change 0358 was rejected and reverted, so it introduces no retained production
API, dependency edge, public-layer change, or performance claim. The candidate
used a private stateless XLS worksheet-span driver with 64 KiB/1,024-item
bounds and no CFB API; its source-backed correctness checks and all 24 ABBA
groups passed their child, schema, oracle, source, and identity gates. The
unqualified predeclared p99 stability gate failed exactly five comparisons,
so no gate narrowing or rerun was used and the production/tests were reverted.
This preserves the ADR 0003 freshness and failure-atomicity boundary and keeps
the XLS source ownership topology unchanged.

The retained evidence and resource observations are recorded in [Change 0358](changes/0358-xls-worksheet-span-batching-rejected.md);
`performance_claim: none`. Collection used one child at a time, CPU 2, a 2
GiB child cap, no retries, one Cargo build lane, and an on-disk target. The
target peak/final observed footprint was 1.9 GiB with approximately 14 GiB
host availability, 132 GiB disk free, and exhausted swap. No parallel build,
latency, RSS, allocation, physical-I/O, or OOM-prevention claim follows. XLS
freshness optimization remains an open hotspot, while the bounded serial ABBA
driver is retained as evidence infrastructure.

## Change 0359 compliance update

Change 0359 adds callback-scoped verified decoded readers at the owning ZIP and
OPC layers. The HRTB reader cannot escape its callback; Store and Deflate paths
drain and verify exact size, CRC, and compressed consumption before returning
success. OPC applies source, cancellation, part, work, and managed-memory
fences without materializing `PartData` or admitting the payload to its cache.
Typed callback errors remain available when a higher-priority source or
archive error wins. This preserves the low-level ownership boundary and the
panic-free public API rules.

The fixed 16 KiB statement covers decoder scratch, not the pre-existing
archive-wide strict-layout proof. Validation was serial and crate-scoped:
ZIP `4/4` focused and `319/319` library; OPC `6/6` focused, `277/277` library,
`13/13` accounting integration, and `6/6` source-reader integration. The
on-disk target was 381 MiB with approximately 14 GiB host availability, 134
GiB disk free, and exhausted swap. See [Change 0359](changes/0359-callback-scoped-verified-decoded-readers.md).
No speed, RSS, allocation, constant-memory, or OOM-prevention claim follows;
`performance_claim: none`.

## Change 0360 compliance update

Change 0360 adds a callback-scoped streaming MCE event processor in the owning
`litchi-ooxml-common` crate while leaving the legacy byte-buffer API unchanged.
HRTB raw and active event views cannot escape their callbacks. The processor
retains namespace, directive, AlternateContent, inactive-branch, and opaque
extension semantics, continues to EOF after callback errors, and retains typed
secondary errors. The new stream contract also makes end-name, root/tail,
declaration, hidden-reference, and AlternateContent CDATA validation explicit.

Finite event, raw-event-byte, attribute, context, name, depth, choice, and
directive ceilings support bounded streaming, but quick-XML internal state,
decoded allocations, container overhead, and callback allocations preclude a
fixed-memory or OOM-safety claim. Serial validation passed `11/11` focused,
`223/223` library, and `1/1` existing integration tests. The single on-disk
target was 267 MiB with approximately 14 GiB host availability, 134 GiB disk
free, and exhausted swap. See [Change 0360](changes/0360-bounded-streaming-mce-events.md);
`performance_claim: none`.

## Change 0361 compliance update

Change 0361 extends the owning `litchi-xlsx` consumer with bounded streaming
x14ac raw and active observers while preserving the MCE ownership boundary.
The raw observer sees ordinary and alias duplicates before generic duplicate
validation. Raw-only one-pass recovery after semantic `NonConformant` or
`MustUnderstand` retains the typed prior semantic error when a later XML,
input, or limit failure becomes primary. The MCE and `AlternateContent`
x14ac byte-compatibility branch streams, while the plain fast path is
unchanged.

The fixed 8 KiB `InterruptedRetryReader` and eight-retry ceiling complement
the existing bounded stream limits, but do not turn the public path into a
fixed-memory or OOM-safe contract: quick-XML and observer allocations remain
outside the input buffer, and `capture_rows=true` may retain a `BTreeMap` up to
configured `ROWS`. Validation passed MCE recovery `7/7`, raw attributes `4/4`,
x14ac focused `12/12`, worksheet `35/35`, `litchi-ooxml-common` library
`234/234`, and `litchi-xlsx` library `813/813`. Selected-cell and full-
worksheet streaming, latency, RSS, and OOM evidence remain open. See [Change 0361](changes/0361-bounded-streaming-x14ac-observers.md);
`performance_claim: none`.

## Change 0362 compliance update

Change 0362 publishes the narrow selected-worksheet capability in the owning
`litchi-xlsx::raw` layer through
`selected_worksheet::{scan, ScanOutcome, SelectedCell, NotEligibleReason,
StreamResult}`. One pass performs active MCE+x14ac selection through XML EOF
for an eligible single-cell subset, distinguishes `Missing` from explicit
`Empty`, validates strict row/cell order and scalar lexical forms, and lets
x14ac `ValidateOnly` parse descent without a row `BTreeMap`.

Unsupported merges, styles, shared strings, shared or array formulas, rich
inline values, and unknown valid structures become typed `NotEligible` only
after XML/MCE/raw EOF. The result is not worksheet semantic validity, so the
caller MUST fall back to the eager parser. This keeps unsupported content and
invalid input on typed boundaries without leaking a package reader, source
owner, or low-level verification handle into the ordinary semantic API.

Focused validation passed `8/8`, worksheet module `43/43`, and
`litchi-xlsx` library `821/821`. quick-XML, observer, and conversion
allocations are outside the accounting boundary; no source-worksheet routing,
OPC verified reader, CRC/size/source fence, full-worksheet streaming,
latency, RSS, or OOM claim follows. See [Change 0362](changes/0362-xlsx-selected-worksheet-scan.md);
`performance_claim: none`.
