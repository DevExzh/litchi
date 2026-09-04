# ZIP, OPC, and CFB substrate baseline

## Change 0408: attributable materialization baseline

[0408](changes/0408-opc-materialization-evidence.md) prepares verifier expectations
once, adds explicit serial materialization ZIP accounting, and captures debug-1
release binaries with frame pointers. Source `eba1f302eb1e04519925d1791a2f9d299e908d89`
uses CPU 2, 20 warmups and 500 samples on the same four-4-MiB-Part synthetic
corpus as 0406. Normal p50/p95/p99 are 2.184/2.213/2.222 ms. A separate allocator
run records 55 allocations and 16,861,253 allocated bytes per operation; the
accounted selector records 16,777,216 decoded bytes and zero output bytes.

Caller reports now attribute 71.82% of whole-process samples to verification
and 16.34% to materialization. They establish usable phase attribution, with
remaining unknown frames documented in the capture. LLC events are unsupported;
all-zero L1 results remain unvalidated. These descriptive runs authorize no
speedup, operation-peak-memory, cold-filesystem, or scaling claim. The
[artifact bundle](results/change-0408/) retains samples, catalogs, CPU data,
flamegraph, exact commands, source/binary hashes and validation logs.

## Change 0402: unmanaged OPC overlay validation decoder reuse

Candidate `51964019db3f6b0787645e3a56c2ecb83bdca65c` follows control
`46ef44966d5be16f153b1f3375ac14401b7139ac`. The private
`SourceBackedPackage::write_part_overlays_to_stream` validation path now uses
one indexed-read session for unmanaged packages and reuses one Deflate decoder
across the selected source-Part reads. Stored members bypass the decoder and
cache hits remain cache-only. Managed packages retain one-shot reads to avoid
unreserved decoder workspace. Existing limits, CRC/size/framing and XML
validation, source freshness, cancellation, signatures, managed budgets,
partial sinks, and raw preservation remain unchanged.

The opt-in `opc_source_overlay_multi_part_noop` selector uses a fixed
three-shape × three-overlay-count matrix: `overlay-small` (compressible 1 KiB
entries, 7,451-byte archive), `overlay-large` (incompressible 64 KiB entries,
2,103,195-byte archive), and `overlay-media-incompressible` (incompressible
256 KiB entries, 8,396,580-byte archive), each at counts 2, 8, and 32. The
non-empty replacement plan is an equal-payload semantic no-op; each output
reopens with the expected eager semantics and preserves raw member order and
untouched ZIP records.

Normal evidence is stable Rust/Cargo/Rustdoc 1.98.1, CPU 2, one worker,
20 warmups, and 500 retained in-process A1/B1/B2/A2 samples per leg. Only
`source.opc_source_overlay.publication_ns` is summarized. Top-level elapsed is
validated only as the preparation/open/planning/publication phase sum. The
global `["warm", "cold-requested"]` cache setting is a configuration envelope,
not cold evidence; fresh-child/process-isolated semantics are not claimed.
Accepted publication cells are small/8 and small/32 at p50/mean/p95/p99,
large/32 at p50 only, and media-incompressible/2 at p50/mean/p95/p99. The
remaining cells are withheld, so no overall matrix improvement is recorded.

Allocator-only candidate-minus-control deltas are exact per-sample vectors:
count 2 `-2/-2/0/0/-80320/-80320`, count 8
`-14/-14/0/0/-562240/-562240`, and count 32
`-62/-62/0/0/-2489920/-2489920`, in allocation calls, deallocation calls,
reallocations, failed calls, allocated bytes, and deallocated bytes order.
Allocator elapsed, live/peak/RSS, physical I/O, cache, and total-memory claims
are excluded. `litchi-opc` correctness passed 289 library tests / 386 total
test items; publication and allocator validators passed 10/10 and 22/22.
`performance_claim: none`; `claim_authorized: false`. The [0402 change
record](changes/0402-opc-overlay-decoder-reuse.md) and retained [evidence
bundle](results/change-0402/) retain the detailed matrix, raw reports, and
validator bindings.

## Change 0401: XLSX selected numeric ownership elision

Production commit `87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb` follows control
`0859063be5a67bd2aafb3531f2126020b2b5000d`. It adds a private borrowed
`Number::validate_lexical(&str)` check and elides owned numeric lexical values
only for unselected, non-formula, non-inline numeric/untyped cells. Selected
numbers retain their exact lexical spelling; formulas retain cached-number
ownership; unselected inline strings still undergo their normal validation.

The measured selector is `xlsx_file_selected_cell` on the fixed
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` medium corpus:
four 48×48 worksheets (9,216 numeric cells), 17 ZIP members, 4,226,429 archive
bytes, and source SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`. The
fixed query prepares `bEnCh01` for canonical `Bench01` (position 1) and reads
`M29` (row 28, column 12), whose stored Number lexical value is `1028012`.
The selected sheet has 2,303 unselected numeric cells for this oracle.

Normal release ABBA used CPU 2, one worker, warm fresh children, 20 warmups,
and 500 samples per leg in A1/B1/B2/A2 order. Only mean, p95, and p99 are
accepted: reductions are `+0.099577940251% / +0.625379111895% /
+1.170167332729%` (mean/p95/p99) for A1→B1 and
`+0.026562239637% / +0.198122423529% / +0.045344544337%` for A2→B2. The
p50 readings are adverse and rejected (`-0.012690677428%` and
`-0.035254218167%`); no median speedup claim is made.

The separate warm allocator ABBA used three warmups and 30 samples per leg.
Candidate-minus-control deltas are exactly **−2,303 allocation calls**,
**−2,303 deallocation calls**, **−16,121 allocated bytes**, and
**−16,121 deallocated bytes**; reallocations and failed allocations are
unchanged. The 7-byte lexical allocation underlying this fixed delta is an
oracle-specific Rust 1.98.1 observation, not a general byte-size claim.
Allocator elapsed time, live/peak/RSS, physical I/O, cold behavior,
throughput, other cells/queries/ranges/corpora, and general XLSX behavior are
excluded. See the [0401 change record](changes/0401-xlsx-selected-numeric-elision.md)
and [evidence bundle](results/change-0401/).

## Change 0400: dimension-bearing XLSX selected-cell streaming and numeric scratch

The cumulative candidate (`f159c0aed`) keeps valid SpreadsheetML
`<dimension>` metadata on the selected-cell streaming path and reuses a
private, bounded numeric-value scratch buffer. The dimension is validated but
does not bound returned results. It is compared with control `2e47ccebf`; the
measured effect is cumulative across dimension-bearing streaming and scratch
reuse, not statically attributable to scratch alone.

The fixed medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` corpus contains four
worksheets × 48 × 48 = 9,216 numeric cells, 17 ZIP members, and 4,226,429
archive bytes (source SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`). Each
fresh child prepares mixed-case `bEnCh01` for canonical `Bench01` (zero-based
sheet position `1`) and reads `M29` (zero-based row 28, column 12); the typed
oracle is Number lexical `1028012`, with selected-cell evidence digest
`36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`.
`litchi::Workbook::open(path)` and query preparation are outside the timer;
only case-insensitive sheet selection and the exact cell read are timed, with
both selected handles retained through the snapshots.

Normal release-binary CPU-2 A1/B1/B2/A2 ABBA used one worker, 20 warmups, and
500 retained warm samples per leg under Rust/Cargo/Rustdoc 1.98.1. Positive
values below mean the candidate is faster; paired reductions are listed as
`p50 / mean / p95 / p99`: A1→B1 `+27.775881% / +27.728990% / +27.657228% /
+28.150563%`, and A2→B2 `+27.711459% / +27.691341% / +27.705070% /
+27.835790%`. Same-implementation drift stayed within the harness ceilings
(maximum `0.72%`).

Separate warm allocator ABBA used three warmups and 30 samples per leg. The
candidate-minus-control exact reductions were allocation calls `-16,771`
(`-16.6063%`), deallocation calls `-14,436` (`-14.6347%`), reallocations
`-26` (`-68.4211%`), allocated bytes `-3,218,512` (`-23.1131%`), and
deallocated bytes `-2,706,847` (`-20.1822%`). Allocator elapsed time and the
global live/peak snapshots are not claim metrics.

Stable validation passed the focused selected-path tests (9/9),
source/fallback tests (3/3), the complete `0400` filter (12/12), full
`litchi-xlsx --lib` tests (918/918), unified-facade tests (4/4), the
`cargo check -p litchi-xlsx` gate, and exact changed-file rustfmt.
Warning-denied all-target
Clippy remains non-clean at 26 control and 28 candidate diagnostics; the two
candidate additions are test-only `result_large_err` findings in the 0400 test
helpers. Preliminary zero-effect captures from before dimension support and
contaminated captures were rejected and excluded. No cold/cache, RSS,
peak-memory, physical-I/O, throughput, or broad XLSX claim follows. See the
[0400 change record](changes/0400-xlsx-selected-dimension-streaming.md) and
[evidence bundle](results/change-0400/).

## Change 0399: unified XLSX selected-cell filesystem baseline

The opt-in `xlsx_file_selected_cell` selector raises the selectable registry
from **419** to **420** and leaves the default at **36 cases / 198 rows**. It
uses the fixed medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` corpus: four
worksheets × 48 × 48 = 9,216 deterministic numeric cells, 17 ZIP members,
4,226,429 archive bytes, and SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.

The query targets canonical `Bench01` (zero-based sheet position `1`) through
prepared mixed-case `bEnCh01`, then reads `M29` (zero-based row 28, column 12).
The independent typed oracle expects a stored number with lexical value
`1028012`; its dedicated selected-cell evidence digest is
`36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`.

Every sample uses a fresh child. `litchi::Workbook::open(path)` and query
preparation are outside the timer; only case-insensitive sheet selection and
the exact cell read are timed, while both selected handles remain live through
elapsed, allocation, and process snapshots. Semantic/source hashes and the
independent eager typed oracle are checked after timing. Logical source
counters are `not_applicable_filesystem_xlsx`; no physical-I/O or locality
claim is made. Cold-requested samples are advisory, and prepared query work
makes `cold-verified` explicitly ineligible. Stable 1.98.1 validation passed
the harness check, focused registry/scope tests, all four XLSX filesystem
integration tests, the **77/77** strict Python schema suite, normal warm and
cold-requested CLI oracles, explicit cold-ineligibility, and an allocator warm
oracle. `performance_claim: none` and `claim_authorized: false` with no
candidate/control speedup claim. See the [0399 change record](changes/0399-unified-xlsx-selected-cell-baseline.md).

## Change 0398: unified XLSX worksheet and cell selectors

The xlsx-gated unified `litchi::sheet::Workbook::sheet(name_or_zero_based
position)` returns `Result<Option<SelectedWorksheet>>`. Case-insensitive names
and zero-based positions select a lifetime-free `Clone + Send + Sync` handle
with `name()` and `position()`. A missing name or out-of-bounds position is
`Ok(None)`.

The selected handle's `cell` accepts A1 (one-based lexical) or raw zero-based
coordinates and returns owned `SelectedCellView::Missing`, `Covered`, or
`Stored(Cell)` without erasing stored `Empty`, `Formula`, or `Unknown` states.
Its `cells` accepts inclusive A1 ranges or zero-based half-open bounds and
returns an owned sparse `Vec<SelectedCell>`. Source selection is catalog-only;
the XLSX scanner/fallback/cache/freshness/limits/cancellation and typed source
changes remain XLSX-owned, with eager bridge parity. Charts/non-grid sheets
return typed XLSX `NotWorksheet`; non-XLSX runtimes return core `Unsupported`.
Legacy 1-based dynamic traits remain unchanged.

No selector is added: the registry remains **419** and the default remains
**36 cases / 198 rows**. Focused validation is four XLSX public tests, five
with XLSB, one owned/source bridge, one non-XLSX runtime check, and feature
checks. This is correctness/API evidence only under stable rustc/Cargo/Rustdoc
1.98.1 because pinned 1.95 lacks Cargo. `performance_claim: none`;
`claim_authorized: false`; no latency, allocation, RSS, physical-I/O, cache,
throughput, or generalization claim follows. See the [0398 change record](changes/0398-unified-xlsx-sheet-selectors.md).

## Change 0397: accepted OPC owned-open validation/index deduplication

Production commit `f275d4566` is the measured candidate
`f20d3f417edc3f3da07bf515676b8e71285ad76f`, compared with control
`6e98db9ece29c1e50241cf3e84c9410ce71dd748`. The owning
`OpcPackage::from_vec(owned)` path removes one redundant eager ZIP
validation/index pass before the real `PhysPkgReader`; `authorize_owned_source`
still performs preservation work, and public `OwnedPhysPkgReader` remains
eager-validating. Path, `from_reader`, and session paths are correctness-only,
not separately timed. Limits, error ordering, session charging, and exact
mixed-storage byte preservation remain covered.

The opt-in `opc_casefold_owned_open` selector raises the selectable registry
from **418** to **419** while the default remains **36 cases / 198 rows**.
Fixed stored corpora contain 256, 2,047, 2,048, and 16,384 ordinary Parts with
32-byte payloads. CPU-2 A1/B1/B2/A2 ABBA used one worker, five warmups, and 30
samples under rustc/Cargo/Rustdoc 1.98.1 because pinned 1.95 lacks Cargo. Normal,
non-allocator release-binary p50 speedups (A1→B1 / A2→B2; positive is faster)
are, in corpus order, `+8.617829% / +8.204676%`,
`+8.298670% / +8.719476%`, `+8.945417% / +8.268274%`, and
`+4.648655% / +4.348226%`; pooled p50 is
`+8.452941% / +8.356702% / +8.490980% / +4.645459%`.

Allocator elapsed time is observational only. On each ABBA leg, exact
allocation/deallocation call reductions are `-1,038 / -8,202 / -8,206 /
-65,550`, and allocated/deallocated-byte reductions are
`-152,024 / -1,212,620 / -1,212,888 / -9,699,800`, in the same corpus order.
Per-sample net-live after-before bytes and reallocations are exactly unchanged;
raw global live-before/after baselines are not cross-run metrics. The accepted
claim is p50 only; no p99 claim follows. See the [0397 change record](changes/0397-opc-owned-open-validation-index.md)
and [evidence bundle](results/change-0397/) for source hashes and rejected
invalid captures. No RSS, peak operation-memory, physical-I/O, cold/cache,
throughput, format/facade, or generalized constructor claim follows.

## Change 0396: rejected OPC exact-lookup experiments and benchmark coverage

The existing four opt-in case-fold selectors are expanded to seven: this
change adds three class-isolated source selectors for exact, ASCII-case-alias,
and genuine-miss lookup. It also adds a 2,047-Part corpus to the prior 256-,
2,048-, and 16,384-Part fixed stored OPC corpora. The seven available
selectors cover eager/source open and combined lookup alongside those three
source-backed lookup classes.
The combined and class vectors each exercise 144 lookups. The selectable
registry rises from **415** to **418**; the default remains **36 cases / 198
rows**. The 2,047 corpus is the explicit control immediately below the 2,048
production threshold. No production implementation is claimed by this batch.

The latency deltas reported below are derived only from normal,
non-allocator release-binary p50 evidence. Source-open measurements time normal unmanaged
`SourceBackedPackage::from_read_at`; lookup measurements time fixed pre-open
unmanaged packages. Values are ordered as **2,048 / 16,384 Parts**.
Allocator-enabled latency is observational only; allocator runs provide
lifecycle/allocation vectors. Validation-constructor coverage is
correctness-only.
The mapless exact candidate regressed approximately `+2,750% / +3,500%`.
The preliminary scalar-`Vec` exact probe was `+14.85% / +20.96%`; the full
inlined linear-probe ABBA was `+20.22% / +12.42%` with `N` source-open
allocator allocation-call savings. `std` prehashed exact was
`+13.42% / +15.88%`; and direct `HashTable` exact was
`+14.66% / +13.36%`, with a high-sample follow-up still approximately
`+14.7%`–`+15.6%`. The final pooled `Arc<str>` candidate regressed exact
lookup `+6.09% / +6.96%`, source-open `+3.38% / +4.30%`, and mixed lookup
`-0.59% / -0.50%`; its allocator source-open vectors added three allocation
calls and approximately `N` deallocation calls while net-live bytes fell by
65,536 / 524,288. This is exact allocator and net-live footprint
evidence, not an RSS, total-memory, or system-footprint claim. Every
production candidate was rejected for latency or lifecycle regressions.

Control is `c0ca6cb5f22ddc68d827b743018855f6b9dc89bd`; the final pooled
`Arc<str>` candidate is `8f7714ee011b170d938f2532fdd385fb2b61cd32`. The [0396
change record](changes/0396-opc-exact-lookup-index-experiments.md) and
[evidence bundle](results/change-0396/) retain the exact lookup/alias/miss
oracles, normal reports, allocator observations, and rejected adjudication.
`performance_claim: none`; `claim_authorized: false`. No RSS, total-memory,
eager/managed/mutable/default/general OPC, validation-constructor latency,
physical-I/O, decompression, cold-cache, throughput, or scaling claim follows.

## Change 0395: unmanaged OPC source-backed case-fold lookup index

`SourceBackedPackage` retains its exact `PackURI` hash lookup and now uses a
private, allocation-free case-fold order index for unmanaged normal and
validation opens with at least 2,048 ordinary Parts. Case-insensitive misses
use binary search over Part positions; no folded-name strings are retained.
Below the measured 2,048 tuning threshold and on managed
`ExecutionContext` opens, the bounded linear fallback remains. Source order,
canonical spelling, freshness checks, mutable `OpcPackage`, and public APIs
are unchanged. The index is fallibly reserved, costs one `usize` per admitted
ordinary Part, and the threshold is not a semantic part-count limit.

The fixed stored-OPC corpus has 256, 2,048, and 16,384 ordinary Parts with
32-byte payloads and a 144-query vector (nine query classes repeated 16
times). Final CPU-2 `A1/B1/B2/A2` release runs used five warmups and 30
samples under explicit Rust/Cargo 1.98.1. The following latency evidence is
from the normal, non-allocator unmanaged `SourceBackedPackage::from_read_at`
binary; allocator-enabled latency is observational only, and validation
constructor coverage is correctness-only. Source-lookup p50 changed by
`-74.23%`/`-74.37%` at 2,048 and `-96.60%`/`-96.45%` at 16,384; 256-Part
lookup was effectively neutral at `-0.09%`/`-1.06%`. Source-open p50 deltas
were `+0.87%`/`+4.50%`, `+3.41%`/`+3.83%`, and `+1.40%`/`+1.50%` for the
three sizes. The initial all-size index probe rejected 256 Parts
(`+31.31%`/`+33.45%` lookup p50 in the normal non-allocator binary), which
established the measured 2,048 boundary.

Allocator evidence is exact retained-vector footprint evidence: open remains
`5,977` calls/`779,619` bytes at
256 Parts, then adds one call and `16,384` or `131,072` bytes at 2,048 or
16,384 Parts; lookup remains 48 calls/1,536 bytes. Source lookup replays have
zero reads/bytes/payload bytes and 144 version calls. The [0395 change record](changes/0395-opc-source-casefold-index.md)
and [evidence bundle](results/change-0395/) bind the reports, catalogs,
hashes, environment, tests, and claim adjudication.

`performance_claim: scoped`; `claim_authorized: true`. Only normal,
non-allocator unmanaged packages opened through
`SourceBackedPackage::from_read_at`, with source-lookup p50 on the fixed
2,048- and 16,384-Part vectors, are authorized. Allocator-enabled latency is
observational only; validation-constructor coverage is correctness-only.
Tails, means, source-open latency, eager/managed/mutable/default/general
behavior, RSS, I/O, decompression, cold cache, throughput, and scaling remain
withheld.

## Change 0393: selected PPTX image metadata avoids full descriptor retention

`SourceSlide::image` and `read_image` now retain only the selected picture
descriptor while still parsing every picture, resolving every target, and
preserving final cancellation/freshness and error-precedence fences. The
full-inventory `images` behavior is unchanged.

Matched stable 1.98.1 release `A1/B1/B2/A2` evidence used one worker, five
warmups, and 30 samples on an eight-picture deterministic PPTX. The selected
`image` p50 improved by 12.434% and 12.404% in the two matched directions.
Both `image` and `read_image` remove exactly eight allocation calls, eight
deallocation calls, and 1,831 allocated/deallocated bytes, with reallocation
counts and logical source vectors invariant. The [0393 evidence
bundle](results/change-0393/) binds exact revision, patch, binaries, corpus,
raw reports, and adjudication.

`performance_claim: scoped`; `claim_authorized: true`. The accepted latency
claim is only selected `image` p50 on the named corpus and protocol. Full-
inventory and `read_image` latency, all tails and means, RSS, physical I/O,
decompression, cold-cache behavior, throughput, scaling, and broad PPTX
behavior are withheld. See [Change
0393](changes/0393-pptx-selected-image-query.md).

## Change 0390: OPC materialization decoder session

Unmanaged full materialization now reuses one operation-scoped Deflate decoder
session across bounded entries; Store entries bypass it. Managed refusal,
cache-hit/waiter behavior, source accounting, and all existing read/error,
freshness, and ownership boundaries remain in force.

Stable Rust 1.98.1 operation-allocator reports (three warmups, 15 samples)
measure exact avoided-decoder reductions of 4/160,640 calls/bytes for tiny,
510/20,481,600 for many-small, and 6/240,960 for few-large. Logical read
calls/returned bytes and materialized Part counts are invariant. The
[0390 summary](results/opc-source-materialization-decoder-session-0390-summary.json),
[raw reports/catalogs](results/change-0390/), and [checked comparison](results/opc-source-materialization-decoder-session-0390-comparison.json)
bind baseline `2c0fd89c7…`, candidate patch `d8254264…`, compiler, binaries,
corpora, and vectors. Latency, operation-local peak/RSS, copied/decompressed/
physical-I/O, and broad/default OPC behavior remain withheld.
`performance_claim: none`; `claim_authorized: false`.

## Change 0387: source-backed OPC materialization shares payload allocations

Unmanaged source-backed-to-owning conversion now adopts each cached
`Arc<Vec<u8>>` through `PartFactory::load_shared` instead of copying the
payload into a new `Vec` and then allocating a new Arc owner. Managed packages
remain refused before ordinary payload reads because their cache handles retain
hierarchical reservations.

Matched release allocator reports used three warmups and 15 retained samples
on stable Rust 1.98.1. The allocation-call and allocated-byte vectors were
constant across the samples: tiny compressible removed 6 calls and 1,656
allocated bytes; many-small incompressible removed 512 calls and 272,384
bytes; few-large incompressible removed 8 calls and 16,777,376 bytes. In every
case this is exactly two calls per Part and the uncompressed payload volume
plus 40 bytes per Part. Part counts, logical source calls, and returned-byte
vectors were unchanged.

The [0387 summary](results/opc-source-materialization-shared-0387-summary.json)
and [raw reports/catalogs](results/change-0387/) bind the exact source blobs,
patch, binaries, environment, corpus hashes, and operation vectors. The
repository-pinned 1.95 installation lacked Cargo, so the installed stable
toolchain was selected explicitly. Because the run was not balanced ABBA or
single-core pinned, latency is withheld. Peak RSS is process-lifetime rather
than operation-local; copied/decompressed/recompressed bytes and physical I/O
remain unavailable. `performance_claim: none`; `claim_authorized: false`.

## Change 0382: PPTX source-backed cross-slide image batch

Change 0382 extends the bounded source-backed PPTX cross-slide copy closure
from one direct embedded picture to a nonempty caller-bounded set of direct
`p:pic` leaves under exactly one direct `p:spTree`. Each selected picture has
exactly one direct `p:blipFill/a:blip r:embed` reference to an internal,
relationship-free `/ppt/media/` `image/*` leaf. Distinct media targets are
copied once, with deterministic destination media URIs and selected image
relationship-ID allocation/rewrite without XML normalization.

Semantic picture parsing preserves bounded foreign, non-MCE,
non-relationship `a:blip` attributes opaquely and accepts one valid
unqualified `cstate` token. Namespace-safe copy rewrites only the full-slide
resolved relationship-namespace `r:embed`. A full-slide unbound lexical
`r:embed` returns `UnsupportedRelationship`; `r:link`, unknown relationship
attributes, MCE, and duplicate or ambiguous resolved embeds refuse. Wrong-type
non-selected and non-anchor slide bindings are rejected at open, while
planning revalidates every binding defense-in-depth; no malformed-object
planner test is claimed.

The destination anchor may preserve other valid existing relationships while
anchoring exactly one dialect-correct internal `slideLayout`. The full-slide
namespace-aware `SourceSlide::images` inventory is an inventory fence; source
catalog relationship reconstruction is fallible, and physical ZIP media
deduplication is asserted. Strict XML end-name and unresolved-prefix fences
remain active for the bounded copy.

Unselected XML and package members remain preserved, as do the existing
freshness, signature, cancellation, partial-sink, and resource fences. Layout,
non-selected, and unsupported relationship-ID collisions still refuse.
Broader dependency graphs, MCE, malformed/ambiguous/misplaced blips,
external/linked/missing/mistyped media, outbound media relationships,
unreferenced image relationships, and unsupported topology fail closed. Image
decoding, conversion, rendering, durable inverse, and broad media-rich copy
remain outside the closure.

The focused cross-copy suite passed `41/41`. The default-feature library
passed `531` with the exact pre-existing
`stale_and_unsupported_raw_xml_fail_before_publication` exclusion. The
all-features library passed `533/533`, and all integration binaries passed
with the exact exclusions
`stale_and_unsupported_raw_xml_fail_before_publication`,
`malformed_presentation_children_are_reported_by_their_owner`, and
`noncanonical_style_target_survives_transactional_raw_save`. PPTX doctests
passed `6` with `2` ignored. Strict Clippy passed with warnings denied and
only the existing `clippy::nonminimal_bool`, `clippy::clone_on_copy`, and
`clippy::needless_lifetimes` allowances. The crate-boundary gate passed for
64 workspace packages and 240 internal dependency declarations with 14
existing debt entries.

Validation used one Cargo invocation at a time with `CARGO_BUILD_JOBS=1`,
incremental and debug build state disabled, one dedicated target, serial test
threads, a 6 GiB per-process virtual-memory cap, and a `>=10 GiB`
available-memory launch threshold. These are OOM-mitigating,
resource-capped controls, not evidence of OOM prevention. No benchmark or
memory measurement changes the baseline.

`performance_claim: none`; `claim_authorized: false`. No latency, hotspot,
allocation, RSS, physical-I/O, cold-cache, throughput, fixed-memory,
image-processing, broad media-rich, real-producer, or system-level
OOM-prevention claim follows.

## Change 0381: DOCX source-backed glossary entry batch

Change 0381 extends the bounded source-backed glossary text owner from one
entry to a general nonempty caller-bounded batch within one glossary. One
topology resolution/materialization and one inventory pass resolve canonical
source-order semantic selectors. Alias duplicates, overlapping selections,
and duplicate paragraph intents fail closed. Only selected paragraphs are
staged, with aggregate selector, entry, replacement, and output limits.

Staging measures every replacement size before materialization, uses one
temporary wrapper at a time, and remains cancellation-atomic. Publication
changes exactly one glossary Part, preserves root, sibling, and opaque XML,
and retains exact no-op, source-bound inverse, and source-freshness behavior.
Glossary create/delete/rename/reorder and metadata edits, cross-part or
general-story batching, managed editing, durable patch wire, and broad DOCX
remain outside this tranche.

The focused glossary-batch suite passed `18/18`; the existing story-text suite
passed `11/11`; the default-feature library passed `926/926`; the all-features
library passed `935/935` and all integration binaries passed; DOCX doctests
passed `74` with `31` ignored. Strict Clippy passed with `-D warnings`. The
crate-boundary gate passed for 64 workspace packages and 240 internal
dependency declarations with 14 existing debt entries.

Validation used one Cargo process and one test run at a time with
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. This is an OOM-mitigating,
resource-capped procedure, not proof of OOM prevention. No benchmark or
memory measurement changes the baseline.

`performance_claim: none`; `claim_authorized: false`. No latency, allocation,
RSS, I/O, throughput, fixed-memory, broad-DOCX, or system-level
OOM-prevention claim follows.

## Change 0380: PPTX source-backed cross-slide image copy

Change 0380 extends the bounded source-backed PPTX cross-slide copy closure
from a dependency-free slide to exactly one direct embedded image leaf. The
planner requires one direct `p:spTree`, one direct `p:pic`, and one direct
`p:blipFill/a:blip r:embed` reference to an internal, relationship-free
`/ppt/media/` `image/*` part. It copies the image bytes and content type as an
inert payload and allocates a deterministic collision-free destination media
URI. It does not decode, convert, or render the image.

The focused cross-copy suite passed `22/22`. The default-feature library
passed `531` tests with
`stale_and_unsupported_raw_xml_fail_before_publication` excluded. The
all-features library passed `533/533`, and all integration binaries passed
with the exact audited exclusions
`stale_and_unsupported_raw_xml_fail_before_publication`,
`malformed_presentation_children_are_reported_by_their_owner`, and
`noncanonical_style_target_survives_transactional_raw_save`. PPTX doctests
passed `6` with `2` ignored. Strict Clippy passed with the three unrelated
pre-existing allowances `clippy::nonminimal_bool`, `clippy::clone_on_copy`,
and `clippy::needless_lifetimes`. The 64-package/240-declaration crate-boundary
gate passed with 14 existing debt entries, and independent topology/API,
resource/freshness, and test/compile reviews accepted the bounded closure.

Validation used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. These are OOM-mitigating,
resource-capped controls, not evidence of OOM prevention. See
[Change 0380](changes/0380-pptx-source-backed-cross-slide-image-copy.md).

`performance_claim: none`; `claim_authorized: false`. No latency, allocation,
RSS, physical-I/O, cold-cache, throughput, fixed-memory, image-processing,
broad media-rich, real-producer, or system-level OOM-prevention claim follows.

## Change 0379: DOCX source-backed glossary entry text

Change 0379 extends the bounded source-backed story-text lifecycle to one
existing DOCX glossary entry selected by unique Unicode-caseless name,
canonical ID, combined name and ID, or checked source-order index. Selection
validates the exact glossary part, entry body span, package and XML dialect,
relationship/content-type ownership, inbound closure, resource limits, and
source fingerprint before text projection or editing. Publication applies the
bound span without semantic re-resolution and preserves glossary properties,
sibling entries, opaque XML, unrelated parts, and package topology.

The focused glossary suite passed `12/12`; the existing source-backed
story-text suite passed `11/11`; the default-feature library passed `926/926`;
the final all-features library passed `935/935` and all integration binaries
passed. DOCX doctests passed `74` with `31` ignored. Strict Clippy passed with
`-D warnings`, the 64-package/240-declaration crate-boundary gate passed with
14 existing debt entries, and independent API/compile and safety/test reviews
accepted the bounded surface.

Validation used one Cargo process at a time, `CARGO_BUILD_JOBS=1`, disabled
incremental/debug build state, one dedicated target, serial test threads, a
6 GiB per-process virtual-memory cap, and a `>=10 GiB` available-memory launch
threshold. This is an OOM-mitigating, resource-capped procedure, not proof of
OOM prevention. See [Change 0379](changes/0379-docx-source-backed-glossary-entry-text.md).

`performance_claim: none`; `claim_authorized: false`. No benchmark, latency,
allocation-volume, RSS, physical-I/O, cold-cache, throughput, scaling,
fixed-memory, broad-DOCX, or system-level OOM-prevention claim follows.

## Change 0378: DOCX source-backed secondary-story text

Change 0378 adds a bounded source-backed text lifecycle for selected DOCX
footnote, endnote, and comment entries. Selectors validate exact entry
ownership, relationship/content-type topology, package dialect, and source
freshness before streaming text or staging an exact text replacement. Forward
publication preserves unrelated package bytes and the existing source,
no-op, inverse, and signature fences; ambiguous, missing, managed,
unsupported, or unsafe secondary-story structures fail closed. Glossary entry
text is deliberately deferred to a later tranche.

The focused secondary-story suite passed `25/25`; the existing story-text
suite passed `11/11`; the default-feature library passed `926/926`; the
all-features library passed `935/935` and all integration binaries passed;
DOCX doctests passed `74` with `31` ignored. The crate-boundary policy passed.
Strict Clippy passed with `-D warnings`.

Validation used one Cargo process and one test run at a time with
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. This is an OOM-mitigating,
resource-capped procedure, not proof of OOM prevention. No benchmark or
memory measurement changes the baseline.

`performance_claim: none`; `claim_authorized: false`. No latency, allocation,
RSS, I/O, throughput, fixed-memory, or system-level OOM-prevention claim
follows.

## Change 0377: XLSX source-backed missing numeric insert

Change 0377 closes the missing create verb in the guarded source-backed XLSX
cell-value owner. An explicit numeric-only `Insert` remains distinct from
existing-owner `Set`, `Clear`, and `Remove`. It creates a physical
cell in an existing row or a new ordered row, expands an existing worksheet
dimension, invalidates calculation state, and retains source-bound exact
inverse behavior. Existing owners, formula-owned ranges, and unsupported
worksheet surfaces fail closed.

The focused source-backed cell-value target passed `59/59`. The broader
locked XLSX gate passed 58 suites and `1213/1213` executed tests with four
exact audited pre-existing row-visibility exclusions. Doctests passed
`2/2`. Production Clippy passed with one named pre-existing
`clippy::useless_asref` allowance. Crate boundaries passed for 64 packages
and 240 internal declarations with 14 existing debt entries. Independent
production, API, safety, and test reviews accepted the bounded change. See
[Change 0377](changes/0377-xlsx-source-backed-missing-numeric-insert.md).

`performance_claim: none`; `claim_authorized: false`. No benchmark, latency,
allocation-volume, RSS, physical-I/O, throughput, fixed-memory, or general
OOM-prevention claim follows.

## Change 0376: XLSB Single Cell Tables lifecycle

Change 0376 closes the opened-document lifecycle gap for canonical XLSB
`tableSingleCells` owners. Adding the first binding to a worksheet without
that owner now creates one collision-free canonical part, content-type entry,
and internal worksheet relationship. Deleting the final binding removes that
exact relationship and part only after lossless ownership and inbound-reference
validation. Existing empty owners remain exact no-ops, and ordinary-table
part lifecycle plus noncanonical or unsafe topology remain refused.

The focused XML Maps integration target passed `30/30` executed tests with
one exact pre-existing ordinary-table expectation excluded. The broader
locked XLSB library/test gate passed 15 suites and `726/726` executed tests
with two exact pre-existing exclusions. Strict production-library Clippy
passed without allowances; crate boundaries passed for 64 packages and 240
internal declarations with 14 existing debt entries. Independent topology,
publication-safety, and test reviews accepted the bounded change. See
[Change 0376](changes/0376-xlsb-table-single-cells-lifecycle.md).

`performance_claim: none`; `claim_authorized: false`. No benchmark, latency,
allocation-volume, RSS, physical-I/O, throughput, fixed-memory, or general
OOM-prevention claim follows.

## Change 0375: PPTX selected-slide retained snapshot

Change 0375 retains validated source-backed PPTX selected-slide semantic
snapshots from planning through publication. Publication verifies the current
raw selected bytes against execution, version, lineage, URI, limits, and the
complete retained selected-slide closure before applying the edit without
semantic recapture. Focused tests report SlidePart/Scene counters of 1/1
before and after one-slide publication and 2/2 before and after the
multi-slide batch; the foreign-identical-source test performs one required
recapture, returns StaleSource, and emits zero bytes.

The three existing selectors were run under serial OOM-mitigating/resource-capped
protocol constraints: one Cargo process and one run process at a time,
CARGO_BUILD_JOBS=1, CARGO_INCREMENTAL=0, 8 GiB build and 4 GiB run virtual
caps, with no parallel build or worktree lane. The litchi-pptx library suite
passed 533/533, source_backed_edit passed 21/21, and the remaining integration
targets passed with exactly three unrelated pre-existing stale expectation
exclusions:
opened::tests::stale_and_unsupported_raw_xml_fail_before_publication,
pptx_malformed_presentation::malformed_presentation_children_are_reported_by_their_owner,
and pptx_table_styles::noncanonical_style_target_survives_transactional_raw_save.
They concern stricter direct sldIdLst owner validation and do not enter this
publication path. Production-library Clippy, the 64-package/240-declaration
boundary gate, and independent equivalence/safety/test reviews passed.

The clean release ABBA package contains exact output, counter, timing, and
provenance data in [the compact result](results/pptx-selected-slide-retained-snapshot-0375-abba.json).
The source counters and output hashes are stable across all legs, but the raw
reports contain no correctness booleans. Direction and drift gates fail, so
performance_claim is none; the exact reuse proof is counter/equivalence
evidence, not a timing or memory inference. No allocation/RSS/heap, reads,
decompression/materialization, physical-I/O, cold-cache, throughput/scaling,
fixed-memory, general-OOM, all-PPTX, real-producer, topology/media, or
parallel claim follows. See [Change 0375](changes/0375-pptx-selected-slide-retained-snapshot.md).

## Change 0374: DOCX story-hyperlink retained snapshot

Change 0374 retains the validated story-hyperlink `Snapshot` inside
`ForwardOnlyPatch` and reuses it during source-backed publication. Execution,
source-version, lineage, complete artifact-fingerprint, post-fingerprint, and
stored-fingerprint equality checks remain before output; semantic story
capture is not repeated. Existing no-op identity, redaction locality, deterministic output,
source immutability, freshness, signature, cancellation, and atomicity gates
remain in force.

The focused `publication_reuses_the_planned_story_snapshot` proof reports
`capture_source = 1` and `load_story = 1` before and after publication. Serial
validation passed `litchi-docx` `935/935` and integration target
`source_backed_story_hyperlinks` `23/23`, subject only to the exact unrelated
pre-existing `replacing_the_path_reports_source_changed_without_retargeting`
exclusion. Production-library Clippy passed with five named pre-existing
allowances; the all-test run exposed two additional unrelated pre-existing
debt lints. The boundary gate passed.

Clean CPU-affinity-2 release ABBA used one logical CPU, one worker, 20
warmups, and 500 samples per case over the seven-story, 15-Part, 24-member
corpus (9,900 archive bytes; SHA-256
`457421e8f86ec8eb52fbe181cebe7d0821ce1e794a08142ff01a4c4e03df0cac`). The
[compact result](results/docx-story-hyperlink-publication-0374-abba.json)
records the clean control/candidate identities, raw-report hashes, output
hashes, all 14 gates across eight rows, and full reductions/drifts. The
authorized claim is no-op p50/mean and redaction p50/mean/p95/p99 on this
benchmark/corpus/protocol only. No-op tails are withheld for control drift.

`performance_claim: scoped`; no reads, decompression, materialization,
allocation, RSS, physical-I/O, cold-cache, throughput, fixed-memory, general
OOM, all-DOCX, unmeasured-selector, or parallel-execution claim follows.
See [Change 0374](changes/0374-docx-story-hyperlink-retained-snapshot.md).

## Change 0373: ODF source allocation preflight

Change 0373 closes known allocation-order gaps across ZIP materialization,
ODF decryption, and source-backed ODP/ODS opening. ZIP materializers and ODF
Deflate decryption now compute the overrun sentinel with checked arithmetic
and fallibly reserve the complete `size + 1` capacity before `read_to_end`.
Encrypted package reads validate manifest ambiguity, Store compression,
password, plaintext-size presence, and the 512 MiB plaintext ceiling before
reading payload bytes.

ODP full-source and ODS full/selective owners now enforce the shared 256 MiB
`content.xml` limit through metadata-only materialized-size inspection before
payload materialization. Encrypted members use manifest plaintext size. ODP
also reconciles freshness before exposing secondary parse errors. CRC, size,
ZIP, MIME, family validation, and publication behavior remain in force.

Focused regressions and broader locked/offline release suites passed:
`soapberry-zip` `320/320`, ODF-common `284/284` plus integrations, ODP
`163/163` plus integrations, and ODS `199/199` plus integrations. Two exact
pre-existing writer tests were skipped. Scoped Clippy passed with six named
pre-existing allowances, crate boundaries passed `64/240` with 14 existing
debt entries, and independent static reviews accepted the batch. See [Change
0373](changes/0373-odf-source-allocation-preflight.md).

`performance_claim: none`; `claim_authorized: false`. The specific checked-
allocation and preflight-before-read invariants are established, but no
latency, allocation-volume, RSS, physical-I/O, cold-cache, throughput, fixed-
memory, or general OOM-prevention claim follows.

## Change 0372: ODP source-backed catalog fused parse

Change 0372 fuses ODP `content.xml` validation and catalog scanning into one
borrowing XML pass while preserving validation-before-catalog error
precedence and all publication fences. It also makes input-dependent shared
namespace-tracker allocation fallible and adds a 256 MiB materialized-size
preflight before content allocation. Existing depth, declaration, page,
name, source-freshness, ZIP, MIME, and media-locality bounds remain in force.

Clean CPU-2 ABBA used one worker, 30 warmups, and 500 samples per leg over the
16,785,912-byte deterministic media-rich ODP corpus, SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.
Control `32290f7ce`/`7291b2...` and candidate
`922eb5e2c`/`9a26cd...` produced open p50 reductions of 15.627% and 17.327%;
both p50 same-side drifts stayed below 5%. See [Change 0372](changes/0372-odp-source-catalog-fused-parse.md)
and the [clean ABBA result](results/odp-source-catalog-0372-abba.json).

The accepted claim is fresh `SourceBackedPresentationCatalog::from_read_at`
p50 latency on this corpus only. Open mean and tails are withheld for drift
or inconsistent direction; list is unstable at tens of nanoseconds; query
regressed in both pairings. No broad open, list, query, all-ODP, RSS,
allocation, fixed-memory, physical-I/O, cold-cache, throughput, or
OOM-prevention claim follows. Focused release tests, scoped Clippy with one
pre-existing allowed lint, crate boundaries, and independent static reviews
passed subject to two exact pre-existing writer-test exclusions.

## Change 0371: shared ODF content validation

Change 0371 extracts the checked plain-`Reader` namespace tracker and content
validator into `litchi-odf-common`, then migrates `litchi-odt` off its
duplicate implementation. The doc-hidden private substrate preserves ODF
namespace binding, rebinding, unbinding, empty-element deferred-pop, content
root/body/version, XML-reference, and tokenizer behavior. It accepts at most
256 namespace declarations per element, retains the 256 MiB input bound, and
rejects nesting beyond 4,096 with a typed invalid-format error rather than
using quick-xml's `u16` namespace-depth path.

Locked/offline release tests passed for all focused `litchi-odf-common` and
`litchi-odt` cases except two exact pre-existing failures in unmodified writer
tests. Scoped Clippy passed with only the pre-existing
`large_enum_variant` diagnostic in `package/model.rs` allowed, and the crate
boundary gate passed. See [Change 0371](changes/0371-odf-shared-content-validation.md).

`performance_claim: none`; `claim_authorized: false`. No timing, allocation,
RSS, physical-I/O, cold-cache, fixed-memory, throughput, or OOM-prevention
claim follows.

## Change 0370: ODP source-backed catalog selectors

Change 0370 adds three opt-in ODP source-backed catalog selectors to the
performance harness: `odp_source_backed_catalog_open`,
`odp_source_backed_catalog_list`, and `odp_source_backed_catalog_query`. This
is a benchmark-coverage change only; production code, CRUD APIs, and the
default matrix are unchanged. The selectable registry is **407** and the
default remains **36 cases / 198 rows**.

The fixed media-rich ODP corpus contains 12 slides, 13 archive members, and
eight deterministic 2 MiB `Pictures/*` members. Its archive is 16,785,912
bytes with SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.
The open selector times fresh source-backed catalog construction; list times
`catalog()` after owner preparation; and query times the selected slide at
index 6 after owner/index preparation. Semantic, topology, source-replay, and
media-locality checks are outside the timed scopes.

The retained [control report](results/odp-source-catalog-0370-control.json)
uses CPU 2, 30 warmups, and 500 samples. It is a dirty descriptive control
from revision `f35486fb7085bb128eb89a4d2e9edd3ad1065f02` with binary SHA-256
`08594839ede39d7f2ed0c143d818e41de0b7cdb77bc92fbcdd2a96083ca9966a`.
Selector p50/mean/p95/p99 timings are open `57,538/61,057.616/76,884/88,020`
ns, list `31/63.854/161/200` ns, and query
`60,062/64,323.154/83,354/101,659` ns.

`performance_claim: none`; `claim_authorized: false`. This is not clean A/B
evidence and makes no latency, RSS, allocation, physical-I/O, decompression,
cold-cache, fixed-memory, throughput, or OOM-prevention claim. The focused
selector and enumeration tests passed `1/1` each; no full suite was run.

## Change 0369: ODT source-backed catalog fused parse

Change 0369 replaces sequential `content.xml` validation and text-block-kind
classification in source-backed ODT catalog opening with one borrowing XML
pass. The private handler preserves the former error precedence: XML/tokenizer
and validation/tokenizer finish errors take precedence over a deferred kind
scan error. Source freshness, ZIP verification, cancellation, and limits
fences still complete before publication. The 256 MiB content limit,
1,000,000-block ceiling, and 4,096 nesting-depth ceiling are unchanged.
Styles, media, and semantic payloads remain cold; no public API, dependency
edge, archive handle, or storage contract changes.

Clean CPU-2 ABBA used control `bf1cb55c6`/`a7991b...` and candidate
`b712aafbf20e`/`1a75eb...`, with 30 warmups and 500 samples per leg over the
corpus SHA-256
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.
Confidence intervals did not overlap and same-side drifts stayed below 15%.
Open reductions were 53.560% p50, 53.116% mean, 51.008% p95, and 49.508%
p99 for A1/B1; and 56.320% p50, 56.078% mean, 54.542% p95, and 54.304% p99
for A2/B2. See [Change 0369](changes/0369-odt-source-catalog-fused-parse.md)
and the [clean ABBA result](results/odt-source-catalog-0369-abba.json).

The accepted claim is deterministic large media-rich ODT corpus latency only
for fresh `SourceBackedDocumentCatalog::from_read_at` open. List is excluded
because its tens-of-nanoseconds result was unstable, and query is excluded
because its 2.9%-3.4% result is below materiality. All-ODT, RSS, allocation,
fixed-memory, physical-I/O, cold-cache, throughput, and OOM claims are
withheld. Correctness evidence is exact rustfmt, `cargo test -p litchi-odt
--lib --tests` (557 library tests and all integration targets; 926 total),
scoped Clippy `-D warnings`, and independent code/resource review acceptance.

## Change 0368: ODT source-backed document catalog evidence

Change 0368 adds three opt-in source-backed ODT catalog selectors:
`odt_source_backed_catalog_open`, `odt_source_backed_catalog_list`, and
`odt_source_backed_catalog_query`. They reuse the fixed 10,008-entry,
13-member corpus with eight deterministic 2 MiB `Pictures/*` members. The
archive is 16,811,815 bytes with SHA-256
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.

The selectors time fresh catalog construction, `catalog()` after owner
preparation, and `block_at(5000)` after owner/index preparation respectively.
Semantic digests and instrumented-source replay gates are outside timing. The
retained [control report](results/odt-source-catalog-0368-control.json) uses
30 warmups and 500 samples on CPU 2, but comes from dirty revision
`14884ced9d8b29b7d2155134025986e9315ac771` and is descriptive only. It is not
a clean A/B comparison or speedup baseline.

The focused catalog oracle and selectable-count tests passed `1/1` each; the
initial standalone harness context was `233 passed, 7 failed, 1 ignored`,
with the count assertion corrected afterward and six unrelated failures
remaining. The selectable registry is **404** and the default remains **36
cases / 198 rows**. `performance_claim: none`; no latency, throughput,
physical-I/O, allocation, decompression, cold-cache, RSS, fixed-memory, or
OOM-prevention claim follows. See [Change 0368](changes/0368-odt-source-backed-catalog-selectors.md).

## Change 0367: XLSX selected merge streaming

Change 0367 extends the selected worksheet scanner with valid direct active
`mergeCells` support globally through verified worksheet EOF. It validates the
exact merge count, nonempty `ref`, reference grid, singleton rejection,
placement/direct-child rules, and overlap rules, then builds the canonical
transient `merge::Index`. After all worksheet, dependency, ZIP, source, and
execution fences complete, a selected single-cell non-anchor returns
`Covered`; anchors retain `Stored`/`Missing`.

Range `cells` and `visit` remain sparse physical records, including merge
followers, with no synthetic covered cells; `stored_extent` is unchanged. The
range path retains at most 16,384 merges with `try_reserve`; 16,385+ drains to
verified EOF and then takes the mandatory eager fallback. Unknown merge
attributes, children, or payload fall back after the drain, while malformed
structure is a hard typed error. Eligible cold paths publish no `Store`,
`PartData`, or semantic caches.

The transient index's internal `BTreeMap` and heap allocations are bounded by
the cap but are not individually fallible, so no fixed-memory, RSS, or OOM
claim follows. Focused validation passed `14/14`, full `litchi-xlsx` library
validation passed `906/906`, and scoped Clippy passed with only the unrelated
`clippy::useless-asref` issue allowed. `performance_claim: none`; no latency
claim follows. See [Change 0367](changes/0367-xlsx-selected-merge-streaming.md).

## Change 0366: XLSX selected general-reference decoding

The selected worksheet single/range scanner now decodes bounded XML general
references through the canonical decode helper. Predefined `amp`, `lt`, `gt`,
`quot`, and `apos` references are eligible in formula, value, and inline
payloads. Decimal and hexadecimal numeric references are eligible in formula
and value payloads only when ASCII and the full token is at most 12 bytes.
Numeric inline references, overlong or non-ASCII numeric spellings, and numeric
scalars outside the XML 1.0 `Char` production return `NotEligible` and use the
verified eager fallback; malformed, custom, and out-of-range references remain
MCE/typed errors. XML/MCE/x14ac and the OPC reader still drain to verified EOF
before publication or callbacks. Eligible cold `cell`, `cells`, and
`visit_cells` paths retain no `Store`, `PartData`, or semantic caches. There is
no API or public accepted-input change; the pre-existing eager/shared-string
XML-legality residual is out of scope. Validation passed focused `9/9`, full
`litchi-xlsx` library `892/892`, and scoped Clippy with `-D warnings` with only
the unrelated `clippy::useless-asref` issue allowed. `performance_claim: none`:
no latency, RSS, fixed-memory, or OOM claim.

## Latest XLSX source-worksheet range streaming (change 0365)

Change 0365 extends cold `SourceWorksheet::cells(area)` and staged
`visit_cells(area)` with a verified sparse raw range scan for eligible
worksheets. Dependency scans reach XML/MCE/x14ac EOF, then ZIP CRC/size
verification and source/execution fences complete before publication or
callbacks. Output is sparse physical output only: missing coordinates are
omitted and explicit empty cells are retained. A multi-index SST stream and a
direct style-count stream avoid worksheet `Store`, `PartData`, and semantic
dependency-cache publication on the eligible cold path; warm `Store` remains
fast. `visit_cells` stages an owned `Vec` sized by selected physical output,
not a fixed-memory guarantee.

`NotEligible` falls back eagerly only after the verified reader returns.
Merges, shared/array/data-table formulas, row/column styles, rich, phonetic,
extension, foreign, and general-reference cases remain eager, and
`stored_extent` is unchanged. Focused validation passed `27/27`, full
`litchi-xlsx` library validation passed `883/883`, and package Clippy passed
with `-D warnings` apart from the unrelated `clippy::useless-asref` issue.
No latency, RSS, fixed-memory, or OOM claim follows; `performance_claim: none`.
See [Change 0365](changes/0365-xlsx-source-worksheet-range-streaming.md).

## Latest XLSX selected-cell dependency streaming (change 0364)

Change 0364 extends the eligible cold selected-cell path with sequential,
verified dependency reads. The scan tracks maximum shared-string and direct
cell-style references across all worksheet cells plus the target SST index,
then reads canonical `sharedStrings` followed by `styles`. Plain selected SST
and direct `c@s` values resolve without `Store`, worksheet `PartData`, a full
text `Vec`, a style `Catalog`, or semantic dependency-cache publication.
Warm semantic caches no longer rematerialize evicted `PartData`; public
signatures remain unchanged.

Dependency readers reach XML EOF and CRC, size, source, and cancellation
fences before values or fallback are returned. Invalid, missing, or
out-of-range references and unsupported or oversize parts retain established
eager diagnostics after readers close. Rich, phonetic, extension, and foreign
SST entries, row or column styles, merges, shared, array, and data-table
formulas remain eager fallbacks, and the final cell source/cancellation fence
runs even on parser errors.

Focused validation passed `28/28`, library validation passed `856/856`, and
scoped Clippy passed apart from the known unrelated pre-existing `hyperlinks`
`useless_asref` issue. Quick-XML and current-item allocations are bounded only
by documented limits. No latency, RSS, OOM, or fixed-memory claim follows;
`performance_claim: none`. See [Change 0364](changes/0364-xlsx-selected-cell-dependency-streaming.md).

## Latest XLSX source-worksheet selected-cell routing (change 0363)

Change 0363 routes cold `SourceWorksheet::cell` through
`PartView::with_verified_decoded_reader` and the raw selected-worksheet
scanner for eligible simple scalar worksheets. Eligible cold queries do not
publish full worksheet `PartData`, `Store`, or cache state; repeated cold
queries rescan, while warm `Store` queries retain their existing fast path.
Every `NotEligible` result falls back to the eager store only after the
verified reader returns and CRC/size/source/context checks complete. Merges,
shared strings, styles, shared formulas, and rich inline values preserve
semantics through that fallback. Source/cancellation/ZIP errors remain
primary, and final outer fences run before a value. Zero, unrepresentable, and
greater-than-2-GiB declared parts bypass the scanner and retain existing eager
behavior. Public signatures, `cells`, `visit`, and `stored_extent` are
unchanged.

Focused/source/library evidence is `7/7`, `16/16`, and `828/828`; scoped Clippy
passed apart from the known unrelated pre-existing `hyperlinks` `useless_asref`
issue. Single-job capped validation observed no OOM as a protocol fact only.
No latency, RSS, fixed-memory, OOM-safety, or dependency-streaming claim
follows; `performance_claim: none`. See [Change 0363](changes/0363-xlsx-source-worksheet-selected-cell-scan.md).

## Latest XLSX selected-worksheet raw scan (change 0362)

Change 0362 adds the public narrow raw path
`litchi_xlsx::raw::selected_worksheet::{scan, ScanOutcome, SelectedCell,
NotEligibleReason, StreamResult}`. It performs one-pass MCE+x14ac active
selection through XML EOF for an eligible single-cell subset, distinguishes
`Missing` from explicit `Empty`, validates strict row/cell order and scalar
lexical forms, and lets x14ac `ValidateOnly` parse descent without a row
`BTreeMap`. Merges, styles, shared strings, shared or array formulas, rich
inline values, and unknown valid structures produce typed `NotEligible` only
after XML/MCE/raw EOF; callers MUST fall back to the eager parser because
`NotEligible` is not worksheet semantic validity.

Focused validation passed `8/8`, worksheet module `43/43`, and
`litchi-xlsx` library `821/821`. No source-worksheet routing, OPC verified
reader, CRC/size/source fence, or full-worksheet streaming was added.
quick-XML, observer, and conversion allocations remain outside the accounting
boundary. This is correctness evidence only: no latency, RSS, OOM, or
performance claim changes the baseline. See [Change 0362](changes/0362-xlsx-selected-worksheet-scan.md);
`performance_claim: none`.

## Latest DOCX source-path ingress and OPC error result (change 0356)

Change 0356 makes Unix and Windows `Document` path ingress single-source:
one `FileSource` and one `SourceVersion` cover ODT MIME/catalog arbitration,
DOCX source ownership, and the bounded `Bytes` fallback. Pathname reopen and
unbounded `fs::read` are removed; portable fallback uses checked file length,
one reserve, fixed `read_exact`, and a final handle-length freshness check.
Caller DOCX `ReadLimits` apply to ZIP/known-OOXML candidates, while the neutral
generic fallback has a finite 2 GiB ceiling and still materializes its bytes.
DOCX terminally preserves `OtherOoxml`/`DisabledOtherOoxml`, genuine
missing-manifest/no-match reclaims the original allocation, and a valid DOCX
wins over an ODT MIME hint with a missing or malformed ODF manifest. Ordinary
ODT retains its separate native-owner policy.

The OPC prerequisite types allocation, all six `LimitExceeded` resources, and
raw I/O errors. Archive-index, normal/validation catalogs (including validation
phase), selected streams, and preservation-index paths retain cancellation,
execution, and source-freshness precedence; only `UnsupportedPreservation`
becomes overlay-unavailable. Public APIs and `DetectedFormat` are unchanged.
Evidence is exact-limit/source-sink-validation-I/O, polyglot, extensionless,
wrong-family, neutral-fallback, and freshness regression coverage, plus
`litchi-opc` `source_backed_reader` `6/6`, OPC lib `271/271`, combined DOCX/ODT
facade `90/90`, and serial DOCX/ODT/PPTX/XLSX/XLS feature checks under the
constrained one-target/one-job environment. The target reached 1009 MiB with
approximately 14-15 GiB host availability and exhausted swap; no OOM occurred.
This is correctness/ownership evidence only (`performance_claim: none`), with
no speed, RSS, or OOM-prevention claim. The caller-sized physical result buffer
uses typed fallible reservation and releases the part reservation on admission
failure; this is correctness/resource safety evidence only. Public eager smart
detection, neutral
2 GiB materialization, non-Unix ODT policy differences, lower-family probe
limits, ordinary ODT limits, `parts_by_name` casing, and selected-Part
materialization remain open. See [Change 0356](changes/0356-docx-source-path-and-opc-errors.md).

## Latest workbook and presentation two-ceiling path result (change 0357)

Change 0357 gives Workbook and Presentation filesystem paths a two-ceiling
policy. OOXML and uncertain/polyglot candidates use caller `ReadLimits` capped
by a neutral 2 GiB absolute ceiling; ordinary canonical ODP/ODS, content-
derived renamed ODP, OLE, and generic non-ZIP fallback use the neutral 2 GiB
policy. Unknown or missing-content-types ZIPs and ODF inputs with an uncertain
OOXML catalog remain caller-limited. PPTX, ODP, native PPT, and bounded `Bytes`
arbitration share one `FileSource` and `SourceVersion`, including freshness,
without pathname reopen. Wrong-family `OtherOoxml`/`DisabledOtherOoxml` is
terminal. The ODF catalog neutral-budget helper applies the input,
compressed-byte, entry, and total ceilings together.

Evidence is `15/15` focused ODF detection tests with `260` filtered, `6/6`
catalog arbitration tests, `82/82` `litchi` `pptx,odp,ppt` library tests, and
`84/84` `litchi` `ods,xlsx` library tests, plus passing quiet checks for
`pptx`, `odp,ppt`, `odp`, `ppt`, and `xls,xlsx`. The first constrained compile
found two `Arc<FileSource>` to `Arc<dyn ReadAt>` coercion errors; they were
fixed before final checks. The serial run used one 8 GiB process ceiling, one
Cargo job, disabled incremental/debug compilation, one disk target, and one
test thread. The target's final/peak observed footprint was 1.3 GiB; host
availability was approximately 14 GiB with 133 GiB disk free and exhausted
swap. No parallel build or OOM occurred. This is correctness/resource
evidence only (`performance_claim: none`), not speed, RSS, allocation,
constant-memory, or OOM-prevention evidence. See [Change 0357](changes/0357-workbook-presentation-two-ceiling-policy.md).

The eager public `DetectedFormat`, full neutral fallback materialization, flat
ODF MIME decode before strict bounding, infallible Presentation aggregate
`Vec`/`join`, portable same-size identity, native PPT mutation coverage,
`Current User` plus `Workbook` OLE classifier inconsistency, applicable
prepared ODP reparse, OPC case lookup, and selected-Part materialization
remain open.

## Latest rejected XLS worksheet span result (change 0358)

Change 0358 tested a stateless worksheet payload-span candidate bounded by 64
KiB and 1,024 consecutive payloads, with no CFB API. Correctness passed while
the candidate was present: Python driver `15/15`, span `9/9`, `source_backed`
`46/46`, `litchi-xls` library `1021/1021`, CFB cursor `7/7`, and fragmentation
`9/9`. A serial six-selector A1/B1/B2/A2 ABBA retained 12,000 samples across
24 groups with 20 warmups per fresh child, one child at a time, CPU 2, a 2 GiB
child cap, no retries, one Cargo build lane, and an on-disk target.

One-cell counters changed by `+316` read bytes, `-79` reads, and `-158`
version calls. Claim-bearing p50/mean deltas were
`+4.984845886382849%`/`+4.771328093073383%` and
`+5.785582423178705%`/`+5.78027327071032%` in the two A-to-B legs. Exactly
five p99 gates failed: FileSource/list A1 -> A2 `+6.59542478684531%`,
FileSource/list A2 -> B2 `-6.916640348285569%`, FileSource/one-cell B1 -> B2
`+8.748517200474495%`, AtomicFile/one-cell A1 -> A2 `-5.699947129465672%`,
and AtomicFile/one-cell B1 -> B2 `+6.439283716879541%`. The candidate and
tests were reverted without narrowing or rerunning the gate. Evidence is
retained in [Change 0358](changes/0358-xls-worksheet-span-batching-rejected.md);
`performance_claim: none`.

The target peak/final observed footprint was 1.9 GiB with approximately 14 GiB
host availability, 132 GiB disk free, and exhausted swap. This is rejected
correctness/resource evidence only, not latency, RSS, allocation, physical-I/O,
or OOM-prevention evidence. XLS freshness optimization remains open; the
serial ABBA driver is reusable evidence infrastructure.

## Latest verified decoded-reader foundation (change 0359)

Change 0359 adds callback-scoped verified Store/Deflate readers to
`soapberry-zip` and `litchi-opc`. The callback reader uses fixed 16 KiB decoder
scratch, cannot escape its HRTB scope, and is drained through exact size, CRC,
and compressed-consumption verification after an ordinary callback return.
The OPC wrapper preserves source/cancellation/error precedence, charges part
work and managed scratch, and bypasses `PartData` and payload-cache admission.
The pre-existing archive-wide strict-layout proof is explicitly outside the
fixed-scratch claim.

Serial correctness evidence is ZIP `4/4` focused and `319/319` library; OPC
`6/6` focused, `277/277` library, `13/13` accounting integration, and `6/6`
source-reader integration. The single on-disk target was 381 MiB with
approximately 14 GiB host availability, 134 GiB disk free, and exhausted swap.
No performance baseline changed. See [Change 0359](changes/0359-callback-scoped-verified-decoded-readers.md);
`performance_claim: none`.

## Latest bounded streaming x14ac observer foundation (change 0361)

Change 0361 adds the bounded streaming x14ac raw and active observer
foundation to `litchi-xlsx`. The MCE raw observer sees ordinary and alias
duplicates before generic duplicate validation. After semantic
`NonConformant` or `MustUnderstand`, raw-only one-pass recovery retains a
typed prior semantic error when a later XML, input, or limit failure becomes
primary. The MCE and `AlternateContent` x14ac byte-compatibility branch now
streams while the plain fast path remains unchanged.

Input uses a fixed 8 KiB `InterruptedRetryReader` with at most eight
interrupted-read retries and the existing bounded stream limits. Validation
passed MCE recovery `7/7`, raw attributes `4/4`, x14ac focused `12/12`,
worksheet `35/35`, `litchi-ooxml-common` library `234/234`, and `litchi-xlsx`
library `813/813`. No performance baseline changed. x14ac `capture_rows=true`
can retain a `BTreeMap` up to configured `ROWS`; quick-XML and observer
allocations are outside the fixed input-buffer claim. This is not selected-cell
or full-worksheet streaming and makes no latency, RSS, or OOM-safety claim.
See [Change 0361](changes/0361-bounded-streaming-x14ac-observers.md);
`performance_claim: none`.

## Latest bounded streaming MCE foundation (change 0360)

Change 0360 adds independent raw and selected-semantic callback observers to
`litchi-ooxml-common::mce`. The parser validates every branch to EOF and applies
finite event, raw-event-byte, attribute, context, name, and existing MCE
structural ceilings without creating a normalized document buffer. Split UTF-8
BOM handling, finite interrupted retries, typed deferred consume errors, and
callback-secondary error retention are part of the contract.

Serial validation passed `11/11` focused streaming tests, `223/223` library
tests, and `1/1` existing markup-compatibility integration test. The one
on-disk target was 267 MiB with approximately 14 GiB host availability, 134 GiB
disk free, and exhausted swap. This does not change a performance baseline and
does not establish fixed-memory or OOM safety. See [Change 0360](changes/0360-bounded-streaming-mce-events.md);
`performance_claim: none`.

## Latest XLSB source-ingress hard-probe result (change 0354)

Change 0354 separates recoverable private XLSB source probes from hard
ZIP/OPC/classifier errors. Non-ZIP, no-match, and missing-manifest inputs may
retain the compatibility fallback; hard failures return typed `OpcError`, and
`Workbook::from_bytes` does not eagerly retry them. Path `FileSource` preflights
the caller's exact `max_input_bytes`, passes that exact limit to fallback
reading, drops the catalog before fallback, and moves retained `Bytes` into
the owned source without a clone, preserving pointer/capacity ownership.
Known non-XLSB detector variants return `NotOfficeFile` without reopening the
pathname. Explicit eager APIs, public smart detection, and the positive
non-ZIP compatibility fallback remain unchanged. Final serial evidence is
the XLSB-only private filter `7/7` (included in the `litchi-xlsb` lib `51/51`),
XLSB facade `23/23`, and successful `xlsx` and `xlsx,xlsb` checks under one
job/thread, disabled incremental/debug compilation, one disk target, and an
8 GiB limit. The target reached 564 MiB; post-run available memory was
approximately 14 GiB with swap saturated. This is correctness/resource-run
evidence only (`performance_claim: none`), not an OOM, RSS, latency, or
constant-memory claim. PPTX hard-probe fallback, public eager/portable
fallback separation, finer constructor ZIP mapping, and full selected-
worksheet materialization remain open.

## Latest XLSB source-backed fallback admission result (change 0353)

Change 0353 removes the post-admission XLSB source-text retry through an eager
full-workbook fallback, along with the eager adapter reader/caches/state and
private detector-side duplicate source/limits state. Source typed errors,
freshness fences, and the eager workbook cache remain; explicit
`open_xlsb_workbook*` APIs and `DetectedFormat::Xlsb` are unchanged. Recognized
nonworksheet tabs are skipped through filtered worksheet positions, while
direct nonworksheet and sparkline/pivot/slicer/timeline selections remain typed
refusals. The 0304 fallback wording is superseded only after source-owner
admission. Final evidence is `23/23` for `litchi` `xlsb_facade` and `40/40` for
`litchi-xlsb` `source_backed`, serialized with one job/thread, disabled
incremental/debug compilation, one disk target, and an 8 GiB limit. The target
reached 647 MiB; post-run available memory was approximately 15 GiB with swap
saturated. This is correctness/admission evidence only (`performance_claim:
none`); no latency, RSS, OOM, constant-memory, or broad XLSB claim follows.
Smart `DetectedFormat`, explicit eager APIs, unsupported platforms,
pre-admission recoverable probes, `Workbook::from_bytes` fallback, and selected
worksheet/dependency materialization remain in scope.

## Latest DOCX source-backed selected-story text result (change 0352)

Change 0352 adds a source-backed DOCX lifecycle for one selected `Main`,
`Header(index)`, or `Footer(index)` story: bounded snapshot/text streaming,
direct-paragraph edits, reversible source-bound patches/inverses, and a
same-topology one-part overlay. Exact no-op/trailing-byte copying and
canonical relationship/content-type/external/shared-target, namespace, and
markup-compatibility validation remain; ambiguous or unsupported XML, stale or
foreign sources, signature changes, cancellation, and non-atomic failures are
typed boundaries. Managed edits are refused. Final evidence is the new
`source_backed_story_text` target `11/11` and existing `source_backed` `16/16`,
serialized with one job/thread, disabled incremental/debug compilation, one
disk target, and an 8 GiB limit. The target reached 347 MiB; post-run
available memory was approximately 15 GiB with swap saturated. This is
correctness/CRUD evidence only (`performance_claim: none`); no speed, RSS,
allocation, physical-I/O, benchmark, or broader story-family claim follows.
The focused cases include strict duplicate/end-tag validation, inverse
hostile-writer refusal, decoded namespaces, and actual emitted-byte bounds.

## Latest indexed-stream validation result (change 0351)

Change 0351 is correctness/resource hardening only (`performance_claim: none`).
The initial compressor/zlib premise was rejected: no artifact establishes a
`~65%` residual-zlib result, and existing `read_to*` already uses a fixed-buffer
verifier. Strict sink validation preflights encryption, method, single-disk
ZIP64 provenance, complete local/central raw metadata and descriptor agreement,
then rejects overlaps or central-directory intrusion while allowing prefixes
and gaps. Locator counts, offsets, adjacency, short buffers, ZIP64 metadata,
physical-entry bounds, fallible allocations, layout single-flight retry, and
`ReaderAt` byte stability are validated. Store uses its exact range; Deflate
uses exact decoder `total_in`; strict CRC-zero applies only to bounded sink
paths, with owned/borrowed fallback compatibility retained. The only resource
statement is one 16 KiB scratch buffer for one active member, excluding source,
index, sink/output, cache, process memory, and concurrency. Final evidence is
`315/315` soapberry-zip and `13/13` litchi-opc operation accounting under
serialized jobs/threads, incremental/debug disabled, one disk target, and an
8 GiB limit. No performance claim, selector, artifact, or GOAL completion is
made.

The final successful package/scenario-scoped commands, not workspace-wide, are
recorded in [Change 0351](changes/0351-indexed-stream-validation.md): `cargo
fmt --package soapberry-zip -- --check`; `cargo test -p soapberry-zip --lib --
--test-threads=1` => `315/315`; and `cargo test -p litchi-opc --test
operation_accounting -- --test-threads=1` => `13/13`, with the record's exact
`ulimit`, target, serialized job/thread, and debug/incremental environment.

## Latest verified-streaming hardening result (change 0350)

Change 0350 is correctness/resource hardening only (`performance_claim: none`).
It adds shared overreported-read validation across `ReaderAt` loops, ZIP
verification, streaming, and the OPC `BorrowedReaderAt` boundary, with checked
offsets/counters. Strict CRC equality is limited to bounded sink `read_to*` /
`read_entry_to*`; ordinary owned reads retain zero-CRC compatibility, while a
borrowed nonempty zero-CRC member returns `None` for owned fallback. Deflate
extra output reports `InvalidSize` before accounting overflow. The bounded
statement is one fixed-size scratch buffer for one active member, excluding
caller source/archive/index, sink/output, cache, and aggregate process memory.
Final evidence is fmt success, `soapberry-zip` `287/287`, and filtered
`litchi-opc` `4/4` with `261` filtered, serialized with one job/thread,
incremental/debug disabled, one disk target, and an 8 GiB limit. No latency,
throughput, RSS, allocation, syscall, decompression, or concurrency claim is
made.

## Latest PhysPkgReader stored-Part borrow result (change 0349)

Crate-scoped formatting evidence: `cargo fmt --package soapberry-zip --package litchi-opc -- --check` passed after formatting.

Change 0349 records a validated immutable-slice `PhysPkgReader` consumer for
eligible stored Parts. The source-backed `&[u8]` handoff removes the legacy
destination `Vec` allocation/memcpy and Part materialization budget/cache
charge, while CRC and ZIP layout validation remain. Encrypted Store/Deflate
inputs return typed errors before owned fallback, and nonempty CRC-zero inputs
return `None` for the legacy owned path. The bounded evidence is
`litchi-opc` integration `8/8`, lower borrowed filter `10/10`, and full
`soapberry-zip` `281/281`, serialized with one job/test thread and an 8 GiB
ceiling. `performance_claim: none`; no timing, RSS, throughput,
physical-I/O, decompression, or allocator claim follows. The mixed stored
corpora remain too weak for timing evidence.

## 2026-08-25: change 0280 replication aborted before smoke

- The frozen replication required exact 0279 binary identities; control matched, candidate did not, so collection stopped with zero samples and performance_claim:none.
- No 0280 latency, counter, semantic, tail, or pooled evidence exists. The unchanged freshness-session candidate remains rejected.
- Evidence: changes/0280-freshness-session-replication-aborted.md and results/0280-freshness-session-replication-aborted-20260825/.

## 2026-08-25: change 0279 freshness session rejected

- A 12,000-sample fresh-child ABBA run failed the unqualified same-side drift limit in four p95/p99 cells, so the candidate was reverted and performance_claim:none.
- Descriptively, version calls fell from 1266 to 26 for open/list and 1802 to 34 for one-cell; FileSource direct comparisons improved 48.44-56.31% with exact logical-read and semantic parity.
- Evidence: changes/0279-cfb-operation-freshness-session-rejected.md and results/0279-cfb-operation-freshness-session-rejected-20260825/.

## Latest XLS FileSource attribution (change 0278)

[Change 0278](changes/0278-xls-source-attribution.md) adds an isolated
owned/atomic/tracked/FileSource/eager/facade attribution runner. On a clean
CPU-2 release build, 17 reports with 20 warmups and 100 retained warm samples
show FileSource p50/mean gaps of 180-267 microseconds over the matched atomic
file source; extra measured `version()` time closely explains those gaps and
uses 46.83%-49.68% of FileSource mean elapsed time. This selects a private
freshness-session candidate for a separate A/B batch. The single-revision,
shared-child, warm-cache package has `performance_claim: none`; no cross-family,
tail, physical-I/O, cold-cache, allocation/RSS, or broad XLS/CFB claim follows.
The matrix remains **398 names** and **36 cases / 198 default records**.

## Latest CFB monotonic cursor XLS evidence (change 0277)

[Change 0277](changes/0277-cfb-monotonic-cursor-abba.md) adds a forward-only
FAT/MiniFAT cursor in `litchi-cfb` and uses it privately for XLS global and
selected-sheet scans. A clean CPU-2, 20-warmup/500-sample A1/B1/B2/A2 run
accepts all three source-path p50/mean cells and open p95, with central gains
of 0.36%-2.03%. Logical calls, bytes, version checks, locality, and semantic
outputs are exact-neutral. List/one-cell tails are rejected, including
31.64% and 5.31% adverse p99 review triggers, so `performance_claim: none`;
there is no selector-wide, tail, FileSource, physical-I/O, allocation/RSS, or
broad XLS/CFB claim. The matrix remains **398 names** and **36 cases / 198
default records**.

## Latest XLS source-global coalescing (change 0276)

[Change 0276](changes/0276-xls-source-global-coalescing.md) retains bounded
two-pass global parsing, one-catalog facade reuse, and a CFB range-error
freshness fix. On the pinned opaque-heavy corpus, global logical ranges fall
`136 -> 69`, open/list total ranges `401 -> 334`, and all 24 source-backed
p50/mean/p95/p99 cells improve in the clean CPU-2 30-sample A1/B1/B2/A2
decision smoke. The instrumented source remains roughly 11-12x eager and the
run is below the 500-sample strict protocol, so `performance_claim: none`;
no physical-I/O, FileSource, allocation/RSS, or broad XLS claim follows.
The matrix remains **398 names** and **36 cases / 198 default records**.

## Latest source-backed XLS selective-read closure (change 0275)

[Change 0275](changes/0275-xls-source-backed-selective-read.md) adds a BIFF8
source owner, facade filesystem routing, and five matched opt-in lifecycle
selectors. Open/list read CFB metadata plus Workbook globals and zero worksheet
or opaque payload bytes; one-cell additionally reads only the selected sheet.
The matrix is now **398 names** and the default remains **36 cases / 198
records**. A dirty five-sample release smoke is prioritization evidence only:
fine-grained source reads are materially slower than eager open, so
`performance_claim: none`; no physical-I/O, allocation/RSS, or broad XLS claim
follows.

## Latest rejected DOC owner-public-phases hypothesis (change 0274)

[Change 0274](changes/0274-doc-owner-public-phases-abba.md) records a clean
CPU-2 ABBA test of public-reader `Vec`-clone removal across three DOC shapes.
Only large lifecycle p50 passed both paired directions; tiny p50 was adverse,
payload-heavy directions disagreed, and means/tails were rejected. The
candidate was reverted under the keep/revert rule, so `performance_claim:
none` and no production optimization remains. The current matrix is unchanged
at 393 names and 36 cases / 198 default records.

## Latest DOCX section-layout closure (change 0273)

[Change 0273](changes/0273-docx-source-backed-section-layout.md) adds one
opt-in typed existing-main-story section layout selector covering
PageSize/Margins/Start/Columns snapshot/edit/commit, durable patch/inverse, and
sequential publication. The current selectable matrix is **393 names** and
the default remains **36 cases / 198 records**. This is correctness/CRUD
closure with `performance_claim: none`; dirty five-sample timing/profile and
whole-process counters are descriptive only. The remaining evidence gap is a
clean retained performance package; no physical-I/O, allocator/RSS, or broad
DOCX claim follows.

## Current matrix and gap (change 0272)

[Change 0272](changes/0272-opc-source-overlay-multi-part-matrix.md) adds three
opt-in source-overlay multi-part selectors and 27 benchmark records: changed,
equal-payload no-op, and mixed, each across sizes 2/8/32 and
small/large/media-incompressible payloads. The selectable matrix is **393
names**; the default remains **36 cases / 198 records**. This is
benchmark-only correctness/evidence coverage with `performance_claim: none`.
The dirty five-sample profile is only prioritization evidence; the current gap
is clean retained evidence plus explicit-context/scaling evidence for any
future recompression, parallel, or compression-policy work.

## Latest OPC timing-boundary correction and XLSX allocator probe (changes 0270-0271)

[Change 0270](changes/0270-opc-relationship-open-timing.md) corrects the
`opc_relationship_open` boundary to production open only, fences the returned
package with `black_box`, and runs relationship/package oracles after timing.
It is evidence-boundary correction only; no performance claim follows.

[Change 0271](changes/0271-xlsx-repeated-store-allocator-probe.md) records a
tracked A1/B1/B2/A2, three-warmup/30-sample, warm fresh-child probe for the two
opt-in repeated-store selectors. Operation-scoped allocator observations are
medium `568 -> 560` calls and `225206 -> 81224` bytes, and oversized
`816 -> 560` calls and `271112552 -> 81224` bytes; both pairings are identical
across 20 compared metrics with zero regressions. Latency is excluded, RSS is
descriptive only, and no operation-local peak/RSS, physical-I/O, decompression,
copy, or broad XLSX claim follows. The default remains 36 cases / 198 records;
claim-0269 remains latency-only. The checked-in artifacts and SHA-256/size
manifest are under
[`results/0271-xlsx-repeated-store-allocator-probe-20260824/`](results/0271-xlsx-repeated-store-allocator-probe-20260824/).

Date: 2026-08-10
Production revision: `2665d572b78f0b3efd9ecfc4bd1fda09f8786ae3`
Branch: `feat/office-format-completeness`

This is the first measured baseline in the performance program. It covers the
shared ZIP/OPC substrate and an initial CFB/OLE2 slice. It does not stand in for
the still-required DOCX, PPTX, XLSX, DOC/XLS/PPT semantic, ODF, iWork,
encrypted, malformed, cold-file, and edit/patch scenario matrices.

The complete raw samples and corpus manifests are in
[`results/baseline-opc-2665d572b-2026-08-10.json`](results/baseline-opc-2665d572b-2026-08-10.json).
The full-process resource result is in
[`results/baseline-opc-2665d572b-2026-08-10.time.txt`](results/baseline-opc-2665d572b-2026-08-10.time.txt).

The additive schema-2 catalog and migration rules are documented in
[`CORPUS_MANIFEST_V2.md`](CORPUS_MANIFEST_V2.md).  The checked default catalog
is [`results/perf-corpus-manifest-v2.json`](results/perf-corpus-manifest-v2.json).
The schema-1 corpus objects and the comparator's case/corpus identity digest are
unchanged; V2 fields that are not represented by the historical report remain
explicitly unknown rather than being inferred.

## Latest XLSX repeated-store cache ABBA (change 0269)

[Change 0269](changes/0269-xlsx-repeated-store-cache-abba.md) is the direct
same-selector release ABBA result for the two primary repeated-store cases.
Control revision `18633404d27bc4c442c09915972e7655cdae813b` is compared with
landed candidate `8a0ca40b1a9d77a9494c74c0cdca38dd61ee68b1` on the pinned
medium and oversized four-sheet XLSX/OPC/ZIP corpora. Each sample repeats
`cell`, `cells`, `visit`, and `stored_extent` eight times in the exact
`semantic_query_only; explicit PartData reacquisition excluded` interval,
using fresh warm children, CPU 2, 20 warmups, and 500 samples. All eight
p50/mean/p95/p99 cells are accepted and zero are adverse-both. This is a
latency-only claim; no resource guardrail, allocation/RSS, physical-I/O,
cold-cache, publication/save, producer, or broad XLSX claim is retained.
See [change 0269](changes/0269-xlsx-repeated-store-cache-abba.md).

## Latest XLS owned-source publication ABBA (change 0268)

[Change 0268](changes/0268-xls-owned-source-publication-abba.md) is the direct
same-selector release ABBA result for the landed owned-source XLS numeric
publication path. Control revision
`1dabd40976d94abdd30ad03bbad6cae0b1a24bf5` is compared with candidate
`6a93ded5dbc14e4b823555bd453740643ce6af10` on the pinned opaque-heavy
comment and RK/MulRK XLS/CFB corpora. CPU-2 A1/B1/B2/A2 legs use 20 warmups
and 500 samples; all eight p50/mean/p95/p99 cells are accepted and zero are
adverse-both. This is a latency-only claim; no resource guardrail,
allocation/RSS, physical-I/O, cold-cache, producer, or broad XLS claim is
retained. See [change 0268](changes/0268-xls-owned-source-publication-abba.md).

## Latest XLSX repeated-store strict schema and harness (change 0267)

The four opt-in selectors from [change 0267](changes/0267-xlsx-repeated-store-strict-harness.md)
raise the current selectable matrix to **389 names** while the default remains
**36 cases / 198 records**:
`xlsx_source_repeated_store_medium`,
`xlsx_source_repeated_store_oversized`, and their
`*_reacquisition_control` structural controls. They use the pinned
`litchi-xlsx-source-repeated-store-corpus-v1` generator over a selected
`xl/worksheets/sheet1.xml` member. Each fresh warm child runs `cell`, `cells`,
`visit`, and `stored_extent` eight times under the explicit timing scope
`semantic_query_only; explicit PartData reacquisition excluded`.

Primary selectors are reserved for future same-selector ABBA comparisons. The
medium and oversized reacquisition controls prove eviction and oversized
bypass through explicit `PartData` reads, but their elapsed/query vectors are
structural-only and are excluded from candidate latency comparison. The strict
summary path pins corpus, query, semantic, cache/read/Budget, child-process,
allocator, and result-channel schemas. This is correctness and evidence-boundary
coverage only: no latency, allocation, RSS, physical-I/O, throughput,
decompression, producer, or production-performance claim is retained. See
[change 0267](changes/0267-xlsx-repeated-store-strict-harness.md).

## Latest fail-closed historical REPORT classification (change 0266)

[Change 0266](changes/0266-report-claim-classification.md) adds the
`report-claim-classification-v1.json` sidecar and its fail-closed checker for
the two audited historical REPORT tables. It binds their headings, headers,
row order, labels, and digests, classifying 167 rows as 145 `historical`, 14
`descriptive`, 8 `withheld`, and 0 `strict_claim`; no strict-claim link is
present. CI runs the checker and focused tests. This is report-integrity
evidence only and makes no latency, resource, or production claim.
## Non-iWork release gate

The workspace-wide CI still owns the complete all-features and iWork gates.
The separate non-iWork gate derives its package and facade-feature closure from
Cargo metadata, excludes the 17 iWork packages and the `litchi-py` facade
consumer (18 packages total), and checks the remaining 45 workspace packages
plus the 35 safe `litchi` features. It rejects any selected dependency tree that
contains an excluded package or a `prost`/`protobuf` package. Bulk commands use
`--all-features`; the facade is compiled with the exact safe-feature union, and
the verifier checks that union in addition to each individual safe feature.
The gate uses argv lists rather than shell-expanded package or feature strings.
It intentionally does not pass Cargo `--locked`: the root `Cargo.lock` is
ignored by this repository, so clean checkouts resolve the workspace normally.

From a clean checkout, run the fast planner and tree checks with an isolated
target directory:

```sh
export CARGO_TARGET_DIR=target/non-iwork-gate
export CARGO_BUILD_JOBS=1
export CARGO_INCREMENTAL=0
export CARGO_PROFILE_DEV_DEBUG=0
export CARGO_PROFILE_TEST_DEBUG=0
python3 -m unittest tools.test_non_iwork_gate
python3 tools/non_iwork_gate.py print
python3 tools/non_iwork_gate.py verify
```

The exact serialized CI modes are:

```sh
python3 tools/non_iwork_gate.py check
python3 tools/non_iwork_gate.py clippy
python3 tools/non_iwork_gate.py doc
python3 tools/non_iwork_gate.py lib-tests
python3 tools/non_iwork_gate.py doc-tests
python3 tools/non_iwork_gate.py deprecated
```

The command derives exclusions and safe facade features on every invocation;
do not replace them with `--all-features`, which intentionally includes iWork.
The six CI modes reuse this one isolated target sequentially with incremental
compilation disabled and debug/test symbol retention reduced. These settings
reduce duplicate and debug artifact size but do not impose a hard filesystem
cap; monitor target size and available disk while running the complete suite.
Clippy intentionally mirrors the repository lint alias: library targets only,
`--no-deps`, and `-D warnings`.
The `lib-tests` mode serializes all 45 bulk package test commands with
`--all-features --lib --tests`, runs `cargo clean -p` for each completed bulk
root, and then tests the exact safe facade union with both library and
integration tests. This bounds retained package test artifacts while leaving
shared dependencies reusable; a failed test stops before its cleanup command.

Measured on 2026-08-24 with the debug/test settings above, one aggregate
non-iWork `lib-tests` invocation failed while linking the `litchi-rtf`
integration tests (`ld` bus error): exit 101 after 8:48.51, maximum RSS
1,659,568 KiB, and approximately 6.3 GiB in the target directory. This is
why the release-equivalent job uses the serialized package commands; the
settings are measured mitigations, not a hard disk-space guarantee.

## Latest real-producer security correctness coverage (change 0264)

The ignored locked `security_corpus` test covers eight pinned public fixtures:
signed POI DOCX/XLSX/PPTX, a read-only protected DOCX, POI CryptoAPI and
binary-RC4 encrypted DOC files, POI `SimpleMacro.xls`, and an OOXML startup
external-link XLSX. Exact source SHA-256 values, paths, semantic decryption
digests, and the external inventory digest are recorded in
[change 0264](changes/0264-real-producer-security-corpus.md) and in the Rust
source manifest.

The gate validates signature status, exact no-op bytes, zero-output signed and
protected mutation refusals, typed password errors and correct-password
semantic readback, inert VBA CFB-stream identity and typed source-backed edit
refusal, external inventory without target resolution, a one-under input
limit, and a zero-output managed budget with post-drop `Memory`, `Objects`,
and `OutputBytes` release. It remains correctness-only: no selector or
default-count change and no latency, allocation, RSS, physical-I/O, or
resource-performance claim.

## Latest PPTX slide-boundary publication coverage (change 0265)

The opt-in `pptx_slide_remove_boundary_save` and
`pptx_slide_move_boundary_save` selectors use a deterministic four-slide,
dependency-free PPTX corpus with 45 ZIP members, 32,396 source archive bytes,
and source SHA-256
`685a1805ad291e8f9852d3ccd584320f20847bd0ac8fdf29857f96efe1109477`.
Removal covers first, middle, and last positions and a final-only refusal;
move covers `0 -> 3`, `3 -> 0`, and the exact `from == to` no-op. At change
0265's landing, the historical selectable matrix was 385 names while the
default remained 36 cases / 198 records.

The production opened-presentation `Snapshot`/`Transaction` paths are used.
Phase vectors cover plan, commit, sequential OPC publication, and semantic
reopen; setup and independent oracles are untimed. Twice-built determinism,
semantic reopen, source immutability, serialized durable forward/inverse
patches, stale/foreign, dependency, unknown-member, markup-compatibility,
signed-package, limits, partial-sink, and zero-sink gates are required. Raw
local and local-offset-normalized central records for untouched members must
match; move requires strict `[Content_Types].xml` identity. This is
correctness and phase evidence only: no latency, allocation, RSS,
decompression, throughput, physical-I/O, or broad PPTX claim is retained. See
[change 0265](changes/0265-pptx-slide-boundary-publication.md).

## Latest DOCX story-hyperlink publication coverage (change 0263)

The opt-in `docx_story_hyperlink_noop_save` and
`docx_story_hyperlink_redaction_save` selectors use a pinned seven-story DOCX
corpus: main, header, footer, footnotes, endnotes, comments, and glossary;
15 OPC Parts; 24 ZIP members; 9,900 archive bytes; and source SHA-256
`457421e8f86ec8eb52fbe181cebe7d0821ce1e794a08142ff01a4c4e03df0cac`.
At change 0263's landing, the selectable matrix was 383 names and the default
remained 36 cases / 198 records.

Source and sink preparation plus the independent story-XML and `.rels`
namespace/membership/target/type/mode oracles are outside the timed interval.
The timer covers open, strict target planning, commit, and sequential
publication. Exact no-op bytes, exact redaction member locality, raw local and
offset-normalized central ZIP identity for untouched members, deterministic
output, source immutability, and stale/foreign/signed/unknown-owner/
partial/zero-sink refusals are required. This is correctness and phase
evidence only; no speedup, allocation, RSS, physical-I/O, or broad DOCX claim
is retained. See [change 0263](changes/0263-docx-story-hyperlink-publication.md).

## Latest strict claim-evidence integrity gate (change 0261)

The strict checker now validates all four compressed ABBA reports and
recomputes the elapsed p50/mean/p95/p99 cells, accepted/adverse sets, resource
A1/B1/B2/A2 values, paired ratios/deltas, and complete canonical summary from
raw sources. Summary labels and declared resource deltas are no longer trusted.
Legacy-v1 and current-v1 report profiles reproduce their direct summaries
exactly; mixed profiles refuse. `time`/`heaptrack` run, status, artifact, and
parser identities fail closed, and every resource leg binds exact variant,
revision, binary, harness tool, and profile metadata. Raw projection-marker
fields are ignored. Public projection helpers are not exposed; the public verifier
path creates the module-private `_ValidatedProjection` trust carrier only after
raw validation, while plain mappings and mutations are rejected before
summarization.

For each leg, `_project_report` sequentially validates raw samples, recomputes
bounded elapsed statistics and identity projections without retaining elapsed
sample values, and discards the raw report/sample payload before the next leg.
The fail-closed ceilings are 512 MiB per member, 2 GiB total decompressed input,
and 64 MiB for the summary. The strict run validates four performance claims,
and the relevant Python verifier suite records 141 passing tests.
`/usr/bin/time -v` records 1,114,076 KiB maximum RSS for that evidence-package
run. These figures describe the verifier and its input boundary, not the Rust
library or a performance improvement. See
[change 0261](changes/0261-strict-claim-canonical-recomputation.md).

## Latest retained XLSX filesystem input-mode coverage (change 0260)

The opt-in `xlsx_file_open` and `xlsx_file_open_lifecycle` selectors now run
each warmup and retained sample in a fresh child over one pinned medium XLSX.
The former selector times path open/root construction; the latter adds
worksheet names, count, and full text. The exact measured workbook and saved
lifecycle projection remain live through operation-only evidence snapshots and
are validated afterward against typed XLSX and OPC/property oracles.

Warm, cold-requested, and admitted cold-verified cache states are explicit.
Cold verification proves initial page-cache state plus positive process
`read_bytes` on a page-aligned, independently hashed source; it does not prove
physical device I/O. This is harness correctness and reproducibility coverage,
not retained before/after evidence. See
[change 0260](changes/0260-xlsx-fresh-child-filesystem-roots.md).

## Latest retained XLSX vendor-extension correctness shape (change 0262)

The opt-in `vendor-extension` shape extends the existing four-sheet, 48-by-48,
eight-media cell-values corpus with orphan XML and binary vendor Parts plus one
XML-local internal relationship. It is selected through
`--xlsx-cell-crud-shape vendor-extension`, is excluded from
`XlsxCellCrudShape::ALL`, and adds no `Case`; at change 0262's landing, the
selectable count was 381 and the default matrix remained 36 cases / 198
records.

The generated source archive has exactly 20 ZIP members, 4,227,295 bytes, and
4,231,356 logical uncompressed payload bytes. Its source SHA-256 is
`b031d236b0f48b45ab357126ff238a40e2a2e147c1471cf82e429ab6f6d250fb`.
Independent OPC/topology/content-type/relationship checks reject a package
root, workbook, or selected-worksheet vendor relationship. Source-backed and
managed one-cell edits retain all 19 untouched raw ZIP local and
central-directory records after local-offset normalization, while exact
no-op, lifecycle, stale/foreign, semantic reopen, managed Budget, and typed
sink refusal gates remain active. Existing signed/protected/unknown/MCE and
other unsupported-owner refusals are inherited production contracts. This is
correctness-only evidence: no latency, allocation, RSS, decompression,
physical-I/O, or producer claim follows.

## Latest retained OPC structural ownership result (change 0259)

Private OPC structural parsing now reuses the lazy ZIP reader's validated
shared decompression allocation for deflated content-types and relationship
manifests. Stored members retain their direct source borrow, while
`IndexedArchive` and other positional sources keep the existing owned fallback.
Focused `Arc::ptr_eq` and enum-variant tests bind all three paths; CRC, size and
resource limits, cache/single-flight behavior, cancellation, and error
propagation remain on their existing implementations.

This deterministically removes one complete post-decompression `Vec` clone per
eligible structural member, but it is not a measured allocation, memory, RSS,
latency, I/O, or end-to-end result. See
[change 0259](changes/0259-opc-shared-structural-members.md); a fixed
relationship-heavy corpus and controlled resource/ABBA evidence remain open.
At change 0259's landing, the opt-in `opc_relationship_open` harness selector
raised the then-current selectable count to 381; it is unmeasured
correctness/timing-boundary coverage only and makes no performance claim.

## Latest retained unified RTF byte-ingress evidence (change 0258)

`litchi::Document::from_bytes` now hands owned RTF bytes directly to the native
RTF parser. The facade therefore accepts literal CP-1252 plus native LZFu and
stored MELA transports without an intermediate UTF-8 requirement, while the
byte, reader, and smart detectors preserve ZIP and OLE2 precedence. Exact
native-source round trip,
semantic parity, malformed-frame refusal, and cursor restoration are focused
correctness gates.

Two clean CPU-2 release A1/B1/B2/A2 captures reuse identical binaries over the
two new opt-in facade selectors, tiny/medium/large plain generated RTF, 20
warmups, and 500 samples per leg. The pinned-metadata capture accepts five
in-run cells, while the immediately preceding matched capture accepts none.
The accepted set is not reproducible, so no latency statistic or speedup claim
is retained. Both compact packages and the complete claim boundary are in
[change 0258](changes/0258-rtf-byte-native-facade.md).

## Latest retained high-level ODT source-ingress result (change 0191)

`litchi::Document::open(Path)` now retains validated ODT files through one
positional ODF package and source-backed semantic owner. Eager `from_bytes`,
OOXML-before-ODF precedence, and ODS/ODP ownership are unchanged.

CPU-2 A1/B1/B2/A2 release runs used 30 warmups and 500 samples over one
16.8 MB package with 10,000 paragraphs and eight 2 MiB pictures. Open-only
statistics remain withheld because same-implementation drift fails each tier.
Open-plus-full-text p50/mean/p95/p99 reductions are
31.41%/31.35%/35.36%/30.02% and 31.74%/32.44%/32.77%/32.50% in the paired
directions; all four pass their predeclared drift ceilings.

An untimed typed-source replay reads 29,080 logical bytes and zero picture
range bytes. This is warm in-process and logical-range evidence, not a
physical-I/O, cold-cache, allocation/RSS, producer, edit/save, or broad ODF
claim. See [change 0191](changes/0191-odt-unified-source-ingress.md) and its
[summary](results/odt-unified-ingress-0199-summary.json).

## Latest retained high-level XLSX source-ingress result (change 0187)

`litchi::Workbook::open(Path)` now hands validated XLSX files from one
positional filesystem source into the existing source-backed OPC/workbook
owner instead of retaining the complete input and eagerly decompressing every
worksheet. Byte-backed opening and edit/save APIs are unchanged.

Clean candidate CPU-2 A1/B1/B2/A2 release runs use 20 warmups and 500 samples
on a deterministic four-sheet, 4.23 MiB media-rich corpus. Open-only
p50/mean/p95/p99 are 91.59%-93.10% lower across both paired directions. Open
plus worksheet names/count/full text is 14.35%-18.30% lower. Every named
statistic passes the 5%/5%/10%/15% same-implementation drift gates.

This is warm in-process high-level elapsed evidence, not a physical-I/O,
cold-cache, allocation/RSS, producer-breadth, edit/save, or broad OOXML claim.
See [change 0187](changes/0187-xlsx-unified-source-ingress.md) and its
[machine-readable summary](results/xlsx-unified-ingress-0195-summary.json).

## Latest retained eager OPC payload-sharing result (change 0186)

Ordinary eager OPC opening now carries the ZIP reader's immutable
`Arc<Vec<u8>>` decompression allocation through serialized-part and XML/binary
Part construction. It removes one full payload ownership copy per admitted
Part while leaving eager all-Part decompression, validation, limits,
cancellation, exact publication, and save semantics unchanged.

The four-Part, 16 MiB incompressible owned-open Heaptrack diagnostic records
whole-process peak heap changing from 71.72M to 55.02M. Clean CPU-2
A1/B1/B2/A2 release runs use 20 warmups and 500 samples for borrowed/owned
opens over many-small and few-large. Few-large p50 is 40.44%-48.10% lower in
both paired directions, but the control drift gate fails, so it is withheld.
Only few-large owned-open p99 passes every paired-direction and stability gate,
at 32.99%/43.51% lower. Many-small latency is withheld.

This is payload-ownership evidence, not selective-open laziness or a broad
OOXML result. See [change 0186](changes/0186-opc-eager-shared-payloads.md) and
its [machine-readable summary](results/opc-eager-shared-0194-summary.json).

## Latest retained OPC shared-overlay result (change 0185)

The source-backed OPC publisher now accepts caller-owned `Arc<Vec<u8>>`
replacement payloads. Existing Vec APIs remain compatible, while changed DOCX,
PPTX, and XLSX same-topology publishers avoid one complete selected-Part
`Arc -> Vec -> Arc` ownership copy. Exact no-ops use the empty-overlay exact
source path; selected-member comparison, XML validation, compression,
signatures, managed budgets, source fences, and partial-sink behavior remain.

Clean CPU-2 A/B/B/A release runs use 20 warmups and 500 samples for twelve
existing XLSX scalar-cell and row-visibility records. Medium 1%, medium
exact-256, and large row-batch complete p50/mean/p95/p99 pass paired-direction
and stability gates; accepted p50 reductions are respectively 2.14%/1.98%,
2.94%/1.15%, and 0.21%/3.13%. Other named statistics are reported
individually, and unstable/directionally inconsistent dense and row cases are
withheld. Heaptrack shows no accepted peak-memory or allocation result. See
[change 0185](changes/0185-opc-shared-source-overlay.md) and its
[machine-readable summary](results/opc-shared-overlay-0185-summary.json).

## Latest retained XLSX row-visibility result (change 0184)

The existing-row visibility editor now carries a lifetime/source-bound proof
from its direct `hidden`-attribute rewriter and reuses the immutable scalar-cell
store after independently validating candidate XML and rescanning row state.
Each changed commit removes one complete scalar-cell parse; generic cell-value
edits retain their full candidate parse.

Clean CPU-2 A/B/B/A release runs use 20 warmups and 500 samples for medium and
large hide-one/unhide-256 workflows. Large commit p50 is 37.79%-43.93% lower in
the first pair and 40.88%-43.25% in the second, with all large commit
distribution/stability gates passing. Large unhide-256 complete lifecycle is
21.70%-29.98% lower across accepted statistics. Medium unhide-256 commit p50
and p99 pass; medium total latency and all medium hide-one latency are withheld
for drift. No allocation/RSS, physical-I/O, cold-cache, producer, formula,
structural-row, or broad XLSX claim follows. See
[change 0184](changes/0184-xlsx-row-visibility-store-reuse.md) and its
[machine-readable summary](results/xlsx-row-visibility-store-0184-summary.json).

## Latest retained ODS one-percent result (change 0183)

The previously withheld fixed ODS 21-existing-cell workload now has a clean
current-HEAD rerun over the same bounded source-backed lifecycle. A CPU-2
A/B/B/A with one release binary, 20 warmups, and 500 samples per fresh process
passes every predeclared stability gate. Complete open, stage, commit, and
sequential-publication p50 is 72.07%-72.61% lower than eager owned-snapshot
publication; mean, p95, and p99 are 68.20%-72.33% lower.

This is evidence closure for the existing implementation, not a new production
or harness change. The claim is limited to the fixed generated two-sheet,
2,048-cell, eight-resource corpus and its 21 existing-cell replacements.
Logical source replay is not physical I/O. Allocation/RSS, cold cache, real
producers, formulas, merges, structural rows, insert/delete, durable ZIP patch,
atomic save, and broad ODS CRUD remain open. See
[change 0183](changes/0183-ods-one-percent-release-evidence.md) and its
[machine-readable summary](results/ods-one-percent-release-0183-summary.json).

## Latest retained PPTX validation result (change 0182)

The source-backed PPTX validator now collects catalog presence facts and
relationship-graph facts in one ordered traversal. Package relationship-list
passes change `2 -> 1`; every Part relationship-list changes `4 -> 1`. Graph
target lookups, XML parsing, report topology, and logical source reads remain
unchanged.

A clean CPU-2 release A/B/B/A with 20 warmups and 500 samples per existing
tiny/medium/large `pptx_validation_report` shape accepts the deterministic
large corpus: complete validation p50 is 7.08%-11.50% lower, with mean/p95/p99
directions and all stability gates also passing. Tiny and medium latency remain
withheld because control drift and, for medium, paired mean/p95 directions fail.
No physical-I/O, allocation/RSS, cold-cache, scaling, producer, or broader
PPTX claim follows. See [change 0182](changes/0182-pptx-validation-catalog-graph-fusion.md)
and its [machine-readable summary](results/pptx-validation-fusion-0182-summary.json).

## Latest retained XLS source-policy result (change 0181)

The plan-only fixed-width numeric path now reuses the immutable snapshot's
already validated worksheet-coverage, protection-classification, and
macro-free facts. Each effective plan removes one complete source
`Workbook` policy reopen while retaining the independent composed-target
semantic reopen and every CFB/publication fence.

A clean CPU-2 20-warmup/500-sample A/B/B/A accepts the exact Number workload:
total p50 is 1.92%-5.91% lower and isolated commit p50 is 3.95%-8.27% lower;
p50/mean/p95/p99 paired directions and stability gates pass. RK/MulRK latency
is withheld because candidate and tail drift exceed policy, though the same
deterministic `1 -> 0` source reopen applies. No publication, physical-I/O,
allocation/RSS, cold-cache, atomic-save, or broad XLS claim follows. See
[change 0181](changes/0181-xls-source-policy-reuse.md) and its
[machine-readable summary](results/xls-source-policy-0181-summary.json).

## Latest retained ODT repeated-text result (change 0180)

`SourceBackedDocument::text()` now retains one fallibly allocated, at-most
16 MiB projection on the first successful parse after its two-call threshold
is reached. On the fixed
10,000-paragraph media-rich ODT, four calls perform two complete `content.xml`
projection phases instead of four while returning four distinct owned strings.
Every sample proves zero source reads after preparation and exact semantic,
archive, media, range, and freshness parity.

Two clean CPU-2 release A/B/B/A cycles accept p50 reductions of
47.01%-50.95% and mean reductions of 46.83%-51.29% across four paired
directions. p95 and p99 remain withheld because the first candidate cycle
failed their stability gates; the balanced retry is retained and disclosed.
No allocation/RSS, physical-I/O, cold-cache, single-call/open, producer,
generic ODF, or broad CRUD claim follows. See
[change 0180](changes/0180-odt-source-text-cache.md) and its
[machine-readable summary](results/odt-text-cache-0180-summary.json).

## Latest retained PPTX catalog result (change 0179)

The source-backed PPTX editor now retains its already validated presentation
catalog across slide capture and publication. On the fixed 200-slide corpus,
one-slide workflows remove two complete catalog builds and 400 slide-node
allocations (`3 -> 1`); the eight-slide batch removes nine builds and 1,800
nodes (`10 -> 1`). Payload materializations and logical source reads are
unchanged.

A clean CPU-2 release A/B/B/A over three existing selectors has identical
non-timing projections in all four legs, but paired p50 directions disagree
and required stability gates fail for every workload. Only the deterministic
metadata-work reduction is accepted. Latency, physical I/O, total allocation,
RSS, cold-cache, scaling, producer, and broader PPTX claims are withheld. See
[change 0179](changes/0179-pptx-source-catalog-reuse.md) and its
[machine-readable summary](results/pptx-catalog-reuse-0179-summary.json).

## Latest retained CFB planning result (change 0178)

Sealed immutable CFB sources now omit one redundant final complete fingerprint
after candidate reopen and optional format-owner validation. Generic `ReadAt`
sources retain the fence. On the fixed XLS corpora this removes exactly one
logical source scan per effective plan: 16,995,840 bytes/17 one-MiB reads for
comments and Number, or 202,752 bytes/one read for RK/MulRK, plus one
source/target digest pair.

A clean CPU-2 release A/B/B/A over four existing selectors records consistently
lower candidate p50 values (23.75%-36.47% across the two paired directions),
but every workload fails at least one predeclared same-implementation stability
gate. Only the deterministic work reduction is accepted; latency, physical
I/O, allocation/RSS, cold-cache, scaling, and producer claims are withheld.
See [change 0178](changes/0178-cfb-owned-planning-fingerprint.md) and its
[machine-readable summary](results/cfb-owned-planning-0178-summary.json).

## Current-HEAD resource probe (change 0115)

The standard-library orchestrator and compact machine-readable result are in
[`tools/perf_resource_profile.py`](../../tools/perf_resource_profile.py),
[`tools/test_perf_resource_profile.py`](../../tools/test_perf_resource_profile.py),
and [`results/resource-profile-current-head-0115.json`](results/resource-profile-current-head-0115.json).
This is current-HEAD evidence, not a before/after comparison or an accepted
optimization result.  It intentionally excludes iWork.

The frozen revision is `be500459961471659f65c180de0e5fe98bc14e3a`; the release
harness SHA-256 is
`1cbb2340eae13f4ed49d5baa27532e1f9b31d5781036bb2a302837bcd2210f5c`.
The aggregate was produced with three timed samples and one warm-up per
workload.  External `/usr/bin/time`, perf, strace, and heaptrack probes use one
sample and include process start-up and profiler overhead.  The worktree was
dirty from unrelated concurrent edits.  The locked release build completed
successfully, so the exact binary hash/size and successful build are recorded.
The original run retained only a post-build dirty source snapshot; it did not
capture the pre-build identity or bounded untracked-file contents, so the
result is `build_succeeded_source_snapshot_only`, not a complete or
cryptographic source-to-binary binding.  The recorded HEAD tree is
`739ba8e610208d2528d580595106a88787143098`, with status-z SHA-256
`94b0a8c2fdd8f508e18cbb3278b21abea36a535c270cf748e7a81a7fe1cc08ed` and
head-to-worktree diff SHA-256
`58a78363d20bd4db858f01a96f33735ac418ea0199a010367242780ad90a6f00` over
49,538 bytes.  A clean rerun with pre/post snapshots and untracked-content
hashing is required before claiming source-to-binary binding.

| Workload / corpus | Harness p50 (ns) | Harness p95 (ns) | `/usr/bin/time` max RSS (KiB) | Heaptrack calls / allocated bytes / peak heap / peak RSS |
|---|---:|---:|---:|---:|
| OPC source one-Part / few-large incompressible | 59,684,605 | 59,822,185 | 118,176 | 1,576 / 306,633,284 / 132,791,664 / 126,573,608 |
| Managed XLSX batch / cell-values medium | 33,260,724 | 33,895,459 | 66,132 | 6,130,956 / 1,026,348,498 / 63,239,618 / 75,801,559 |
| RTF streaming / medium | 10,016,573 | 10,114,007 | 30,080 | 450,852 / 66,379,667 / 26,025,656 / 35,232,153 |
| CFB selective MiniFAT / 36-byte target | 140,654 | 145,330 | 30,336 | 13,589 / 148,580,902 / 23,142,072 / 27,682,406 |
| CFB selective FAT / 4 MiB target | 374,947 | 1,225,272 | 30,336 | same paired process profile as the selective run |
| CFB same-length atomic save / few-large | 156,307,917 | 157,041,972 | 110,884 | 1,722 / 460,627,078 / 115,186,073 / 122,704,363 |

The logical counters are separate from physical syscall observations.  The
OPC source case recorded 549 source reads and 16,785,201 source bytes per
sample, one ordinary payload materialization, and a 16,783,632-byte sink with
461 writes.  Managed XLSX recorded 225 source reads and 4,230,793 source bytes
per sample, six materializations, and a 4,226,645-byte sink with 163 writes.
RTF retained zero output bytes and a 37-byte authoring window; its sink accepted
630,819 bytes in 90,122 writes.  CFB selective returned 36 bytes from one
MiniFAT range and 4,194,304 bytes from one FAT range.  The CFB save samples
each reported 1,825 logical reads / 84,838,500 bytes, one changed span, and a
16,913,408-byte publication; the filesystem wrapper's parent wall time is
reported separately from the inner operation time.

The host reported Linux `6.8.0-101-generic`, AMD EPYC 9575F, 12 logical CPUs,
Rust `1.95.0`, `perf_event_paranoid=1`, heaptrack 1.5.0, perf 6.8.12, strace
6.8, and GNU `/usr/bin/time`.  All six requested perf counters were available
in the one-sample probes.  The strace distributions are whole-process
`read`/`write` syscall return sizes; they are not decompressed, recompressed,
or memory-copy byte measurements.

The explicit execution-context scaling selectors covered 1, 2, 4, 8, and the
host-capped available width (12).  On the many-small incompressible corpus,
both OPC and CFB were classified `nonideal_or_measurement_noise`: their raw p50
values showed no measured speedup and at least one derived Amdahl fraction was
outside [0,1].  Invalid fractions are null in the estimate field and retained
as raw values with validity flags.  These are descriptive calculations at the
measured widths, not a claim about a hardware limit or general parallel
behavior.

The probe does not establish cold-cache, remote-range, allocation attribution,
decompressed/recompressed bytes, memory-copy volume, or before/after change.

## Rejected XLSX publisher-provenance experiment (change 0141)

Clean release binaries at control `b5ace54a7` and candidate `eccd8de78` ran
seven media-rich source-backed XLSX edit/save cases in strict CPU-2
`A1, B1, B2, A2` order, with 20 warmups and 200 measured samples per case and
leg. The candidate skipped publication-time semantic reloads by retaining
private lineage/version metadata in each snapshot. It was 1.04% slower on the
pooled seven-case p50 geometric mean; pooled individual p50 changes ranged from
-1.52% to +3.84%, and paired directions were inconsistent.

All 5,600 observations retained identical corpus, output, sink, logical source,
and materialization evidence. Heaptrack recorded 675,330 -> 656,136
whole-process allocation calls (-2.84%) and 83,519 -> 81,745 temporary
allocations (-2.12%), with peak heap unchanged at 152.90M. One matched
`/usr/bin/time -v` direction observed 147,916 -> 146,900 KiB VmHWM (-0.69%),
which is classified as neutral. The candidate was fully reverted by
`a12387478`; see
[`change 0141`](changes/0141-xlsx-source-provenance-negative-result.md) and the
[`machine-readable summary`](results/xlsx-source-provenance-0141-summary.json).
No physical-I/O, decompression, recompression, copy-byte, or cold-cache claim is
made.

## Measurement environment

| Item | Value |
|---|---|
| OS | Linux 6.8.0-101-generic, x86_64, KVM |
| CPU visible to process | 12 logical CPUs, AMD EPYC 9575F |
| Memory | 31 GiB visible |
| Rust | `rustc 1.95.0 (59807616e 2026-04-14)` |
| Build | Cargo `release`, locked dependencies, system allocator |
| Hardware counters | Unavailable: `/proc/sys/kernel/perf_event_paranoid` is `4` |
| Samples | 3 untimed warm-ups and 15 measured iterations per matrix cell |
| Input state | Deterministically generated in memory before timing; warm-memory workload |
| Output state | Bounded forward-only counting sink; bytes are not retained |

The JSON reports `git_worktree_dirty: true` because the harness and performance
documents were uncommitted and an unrelated pre-existing documentation edit
was present. No production source file differed from the named revision when
this baseline was captured.

Command:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
/usr/bin/time -v tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 \
  --json docs/performance/results/baseline-opc-2665d572b-2026-08-10.json
```

The deterministic corpus has tiny, 256-member many-small, and four-member
few-large shapes, each with compressible and deterministic incompressible
payloads. The few-large shape contains 16 MiB of logical Part data. The JSON
records generator parameters, archive and target SHA-256 hashes, logical and
physical byte counts, raw sorted samples, p50/p95/p99, sample standard
deviation, and a two-sided Student's-t 95% interval for the mean.

## Latency and bytes

Times below are p50; p95 is included where it changes the interpretation.

| Case / corpus | Archive bytes | Logical Part bytes | p50 | p95 | Observed output |
|---|---:|---:|---:|---:|---:|
| ZIP index, 256 compressible Parts | 54,615 | 262,144 | 41.4 us | 55.9 us | n/a |
| ZIP index, 256 incompressible Parts | 302,935 | 262,144 | 27.8 us | 35.1 us | n/a |
| ZIP read one, 4 MiB compressible Part | 99,044 | 4,194,304 | 408 us | 431 us | n/a |
| ZIP read one, 4 MiB incompressible Part | 16,783,565 | 4,194,304 | 480 us | 526 us | n/a |
| OPC open, 256 compressible Parts | 54,615 | 262,144 | 622 us | 713 us | n/a |
| OPC open, 256 incompressible Parts | 302,935 | 262,144 | 737 us | 1.01 ms | n/a |
| OPC open, 16 MiB compressible Parts | 99,044 | 16,777,216 | 499 us | 1.41 ms | n/a |
| OPC open, 16 MiB incompressible Parts | 16,783,565 | 16,777,216 | 648 us | 1.08 ms | n/a |
| OPC no-op save, 256 compressible Parts | 54,615 | 262,144 | 1.57 ms | 1.84 ms | 54,615 B / 1,813 writes |
| OPC no-op save, 256 incompressible Parts | 302,935 | 262,144 | 5.73 ms | 6.09 ms | 302,935 B / 1,813 writes |
| OPC no-op save, 16 MiB compressible Parts | 99,044 | 16,777,216 | 3.38 ms | 3.56 ms | 99,044 B / 49 writes |
| OPC no-op save, 16 MiB incompressible Parts | 16,783,565 | 16,777,216 | 212.8 ms | 229.5 ms | 16,783,565 B / 557 writes |

The 16 MiB incompressible save processes about 78.9 MB/s of logical payload at
p50 and rewrites the complete 16.8 MB archive. This is the dominant measured
latency. The 256-member cases expose a different fixed cost: both save variants
perform 1,813 sink writes and regenerate metadata proportional to Part count.

The complete 24-cell matrix also includes tiny cases; those sub-100 us timings
show visibly higher relative noise and are retained as smoke/regression inputs,
not as optimization decision evidence.

## Allocation and peak-memory profile

Heaptrack was run on 100 iterations of the 256-Part incompressible cases. Its
process totals include one deterministic corpus build, one package open for the
save case, report construction, and process/runtime startup, so they are useful
for before/after comparisons with the identical command rather than exact
per-operation allocation counts.

| Workload, 100 iterations | Allocation calls | Temporary allocations | Peak heap | Peak RSS with Heaptrack |
|---|---:|---:|---:|---:|
| `opc_open` | 809,803 | 78,589 | 1.92 MB | 13.56 MB |
| `opc_noop_save` | 356,632 | 79,136 | 1.73 MB | 12.39 MB |

The save allocation stack directly identified duplicated work in
`PackageWriter`: 25,600 `ContentTypesItem::to_xml` allocation paths under
publication validation and another 25,600 under emission across the 100 save
iterations. `ContentTypesItem::from_package` showed the same two-pass shape.
This makes a reused, prevalidated publication plan the first low-risk measured
optimization. It will not remove Deflate work, so Amdahl's law predicts a much
larger relative effect for many-small packages than for the 16 MiB
incompressible case.

The uninstrumented complete matrix consumed 4.49 seconds of wall time, 4.52
seconds of user CPU, 0.08 seconds of system CPU, and 72,516 KiB maximum RSS.
Those are full-matrix process figures, not per-case peaks.

## CFB baseline

The CFB generator uses the same deterministic payload families and adds a
2,048-stream wide-root shape. Tiny and 256-stream inputs exercise MiniFAT;
four 4 MiB streams exercise regular FAT chains; the lexicographically greatest
wide-root stream makes the existing full-tree name lookup traverse its costly
successful path. Raw samples and hashes are in
[`results/baseline-cfb-2665d572b-2026-08-10.json`](results/baseline-cfb-2665d572b-2026-08-10.json).

| CFB case / corpus | p50 | p95 | Interpretation |
|---|---:|---:|---|
| open, 256 1 KiB MiniFAT streams | 139 us | 155-161 us | Eager topology/allocation validation, not all payload reads |
| open, four 4 MiB FAT streams | 139-142 us | 164-173 us | Payload size has little open effect because regular stream bytes remain lazy |
| open, 2,048 root streams | 948-957 us | 1.05-1.07 ms | Directory and allocation metadata scale with member count |
| list 2,048 stream paths | 76.9-82.5 us | 91.6-96.1 us | Materializes every path |
| read last 64 B stream among 2,048 | 7.47-7.52 us | 7.54-7.70 us | Full sibling-tree DFS dominates the tiny payload |
| read one 4 MiB FAT stream | 104-110 us | 135-149 us | Lookup is trivial; contiguous memory-backed copy dominates |
| insert borrowed prepared 4 MiB stream | 640-675 us | 717-747 us | `create_stream` allocates and copies the complete payload |
| insert owned prepared 4 MiB stream | 0.17-0.29 us | 0.31-0.45 us | Ownership transfer only; payload creation and CFB serialization excluded |

The writer comparison deliberately times only insertion of an already-prepared
payload. It proves the cost of the extra 4 MiB copy and provides a direct gate
for fresh DOC/XLS/PPT writers that already own their generated buffers; it is
not an end-to-end CFB-save speedup claim. The two payload families produce
similar CFB results because CFB does not compress these streams.

The complete CFB matrix took 0.29 seconds wall, 0.19 seconds user CPU, 0.11
seconds system CPU, and 44,468 KiB maximum RSS. These figures include corpus
generation and all 40 measurement cells, so they are not per-case peaks.

## CFB stream-chain validation scratch (change 0190)

CFB open validation now reuses one fallible chain vector and visited map for
MiniFAT streams and one pair for FAT streams. Root, directory, allocation-table,
ownership and physical-layout validation are unchanged. On the exact
many-small plus wide-root Heaptrack process (three warmups and 100 samples per
shape), allocation calls fall 988,558 -> 509,749 (-48.44%) and temporary
allocations 242,178 -> 2,567 (-98.94%); peak heap is flat at the displayed
2.72 MiB. The control attributes 237,312 calls to each removed per-stream site.

Release A/B/B/A timing used 200 warmups and 5,000 samples per shape. Accepted
many-small tail reductions are 7.31%-18.36% p95 and 16.81%-19.94% p99;
p50/mean are withheld on control drift. Accepted wide-root reductions are
1.49%-2.18% p50, 2.20%-3.90% mean, and 6.53%-12.51% p95; p99 is withheld on
candidate drift. See [change 0190](changes/0190-cfb-stream-chain-scratch.md), the
[summary](results/cfb-chain-scratch-0190-summary.json), and the [manifest](results/cfb-chain-scratch-0190-manifest.json).

## CFB selective exact-range ABBA

Change 0094 measures the public `SharedOleFile::read_stream_range` seam against
the legacy full-stream reader on the same deterministic archives. The release
run used a pinned before-A/after-A/after-B/before-B order, 30 warm-ups and 500
samples per cell. The paired values below are in ABBA order; percentages
are after versus its adjacent before control.

| Target / shape | Source bytes, legacy -> range | Read p50, legacy -> range | Read p95, legacy -> range | Total p50, legacy -> range |
|---|---:|---:|---:|---:|
| 36-byte MiniFAT / many-small | 261,184 -> 36 (one request) | 9,823/9,238 -> 481/480 ns (-95.1%/-94.8%) | 12,967/12,828 -> 731/671 ns (-94.4%/-94.8%) | 138,936/148,224 -> 127,265/127,175 ns (-8.4%/-14.2%) |
| 36-byte MiniFAT / wide-root | 2,096,192 -> 36 (one request) | 84,276/82,613 -> 671/651 ns (-99.2%/-99.2%) | 95,602/92,907 -> 1,052/821 ns (-98.9%/-99.1%) | 1,163,541/1,240,638 -> 1,086,951/1,092,570 ns (-6.6%/-11.9%) |

The FAT controls retain exactly one 4,194,304-byte request and one source read
call before and after. Their p50s are control-like rather than an accepted
FAT improvement (many-small read p50 117,416/114,287 -> 112,960/112,094 ns;
wide-root 152,310/153,601 -> 157,311/152,194 ns). Paired FAT read and total
p50 changes stay within 5% control drift; p95 and p99 FAT tail claims are not
accepted. Recorded p99 values, cold-filesystem behavior, simulated high-latency
range behavior, allocation, and peak-RSS conclusions remain withheld. This is generic CFB substrate
evidence; it does not certify DOC/XLS/PPT semantic CRUD adoption. See the
[change record](changes/0094-cfb-selective-read-evidence.md) and
[compact ABBA summary](results/cfb-selective-range-abba-0106-summary.json).

## CFB selective simulated-range ABBA (change 0144)

The follow-up clean-revision release run keeps the same deterministic
final-position MiniFAT and FAT targets, but applies a harness-only bounded
range model: 100 us fixed latency, 25 us request overhead, 50 MiB/s bandwidth,
and a 64 KiB physical-request ceiling. Four CPU-2-pinned legs ran in
`A1 legacy, B1 shared, B2 shared, A2 legacy` order, with 20 warmups and 200
samples for each of three targets and both `many-small` and `wide-root` shapes.

| Target / shape | Selective read work, legacy -> shared | Total p50 reduction, pair 1 / pair 2 | Total p95 reduction, pair 1 / pair 2 |
|---|---:|---:|---:|
| 36-byte MiniFAT / many-small | 4 requests / 261,184 B -> 1 / 36 B | 40.12% / 39.99% | 40.64% / 39.08% |
| 4095-byte MiniFAT / many-small | 5 requests / 265,216 B -> 1 / 4,095 B | 40.09% / 39.82% | 40.26% / 39.75% |
| 36-byte MiniFAT / wide-root | 32 requests / 2,096,192 B -> 1 / 36 B | 41.96% / 41.83% | 42.23% / 41.58% |
| 4095-byte MiniFAT / wide-root | 33 requests / 2,100,224 B -> 1 / 4,095 B | 42.00% / 41.84% | 41.96% / 41.70% |

The 4 MiB FAT controls retain exactly 64 requests, 4,194,304 returned bytes,
and an 88,000,000 ns modeled read-service floor for both implementations.
Their paired p50 changes are between -0.09% and +0.08%; they are classified as
matched-work near-neutral controls, not improvements. The accepted result is
only for this configured simulator. It is not cold-filesystem, physical-device,
ambient-network, production scheduling, allocation, RSS, or native DOC/XLS/PPT
evidence. See the [compact summary](results/cfb-simulated-range-0144-summary.json)
and [change record](changes/0144-cfb-simulated-range-source-evidence.md).

## PPTX cross-presentation slide-copy evidence (change 0145)

Two opt-in selectors exercise deterministic plain and media-rich cross-
presentation slide-copy plans. Each reports plan, commit, and sequential OPC
publication phases separately; reopen is retained as a non-publication
diagnostic. Complete semantic/package topology, dependency-closure,
collision-remap, source-immutability, durable-patch, stale/foreign, and refusal
checks remain outside timing. This is correctness and sink-counter evidence
only at the 0145 revision. [Change 0158](changes/0158-pptx-additive-topology-release-abba.md)
now accepts a clean release comparison for the later owned-source additive-
topology publisher; allocation attribution and physical-I/O remain open. See
the [original selector record](changes/0145-pptx-cross-slide-copy-evidence.md).

## PPTX additive-topology release ABBA (change 0158)

Clean control `e8a67b19e` and candidate `d900ae633` release binaries used the
byte-identical harness and lockfile in strict CPU-2 `A1, B1, B2, A2` order.
Each leg retained 200 samples per plain and media-rich selector after 20
warmups, for 1,600 total observations. All semantic, topology, dependency,
durable-patch, immutability, stale/foreign, and refusal gates passed.

| Corpus | Total p50 improvement, pair 1 / pair 2 | Publication p50 improvement, pair 1 / pair 2 |
|---|---:|---:|
| Plain | 29.643% / 26.196% | 82.798% / 82.304% |
| Media-rich | 43.294% / 43.604% | 49.321% / 49.680% |

Plain total and media-rich total/publication p95, p99, and mean agree in both
directions. Plain publication tails are withheld because candidate same-
implementation drift crossed the p95/p99 thresholds. Matched process-wide
profiles agree with the media-rich total direction: task-clock falls
42.399%/43.122%, cycles 42.583%/43.116%, and instructions
46.686%/46.775%; maximum RSS is 0.486%/0.480% higher and peak heap is
effectively unchanged. This accepts only canonical generated owned-source
prepared slide copy. It is not end-to-end file save, source-backed/cold-I/O,
decompression, generic OPC/PPTX, real-producer, or iWork evidence. See the
[record](changes/0158-pptx-additive-topology-release-abba.md) and
[summary](results/pptx-additive-topology-abba-0158-summary.json).

## CFB MiniFAT `open_stream` evidence (change 0146)

Twelve opt-in selectors now call `SharedOleFile::open_stream` directly for
36-byte and 4,095-byte MiniFAT targets across the deterministic 256- and
2,048-sibling shapes. One-shot, repeat-3, and sequential repeat-8 operations
record exact output hashes, per-invocation positional source events, root Mini
Stream identity, source-version checks, and matched deterministic-range-model
evidence. Current-candidate tests bind the direct-then-root-cache counter shape;
the same runner also permits the clean parent revision's initial root
materialization. This is harness/correctness evidence only. Release ABBA,
allocation, RSS, physical-I/O, cold/network/device, native DOC/XLS/PPT, and
cross-format claims remain open. See the
[change record](changes/0146-cfb-open-stream-evidence.md).

## CFB MiniFAT `open_stream` release ABBA (change 0147)

Four clean CPU-2 release processes ran in `A1 control, B1 candidate, B2
candidate, A2 control` order with 20 warmups and 200 samples for each of 24
records. Under the configured 100 us fixed latency + 25 us/request, 50 MiB/s,
4 KiB-range model, every 36- and 4,095-byte one-shot cell improves total
p50/p95/p99/mean by about 62-64% in both directions; the isolated
`open_stream` interval improves about 98.4-99.9%. Exact positional work falls
from the complete 261,184/265,216/2,096,192/2,100,224-byte root Mini Stream to
one exact 36- or 4,095-byte range.

The result is not generalized to repeats. Candidate repeat work is
`[L,R,0...]` rather than the control's `[R,0...]`; several many-small modeled
p50/mean cells regress about 0.3-1.2%, with consistent tails up to about 2.8%.
One 9.5% p99 leg reverses direction and carries same-implementation tail drift.
No generic local wall-clock, allocation/RSS, physical-I/O, cold/network/device,
native-format, or cross-format claim is accepted. See the
[release record](changes/0147-cfb-open-stream-release-abba.md) and
[compact summary](results/cfb-open-stream-abba-0147-summary.json).

## CFB target-aware repeat-policy harness (change 0148)

The 0148-era 291-name harness added six production-only selectors for different-
SID A-B-A, public bulk A-B-A, and overlapping same-target calls at 36-byte and
4095-byte MiniFAT targets. Their correctness/source-event records retain
ordered workload names, output hashes and lengths, exact positional ranges,
source-version stability, and typed missing-stream refusal. The runner accepts
the control root-only vector, the prior direct-then-root vector, and the
target-aware same-SID repeat vector; concurrent overlap uses only a harness-side
entry gate, and bulk calls the public `bulk_read` API.

This is correctness/source-event evidence only. Failure/retry, ineligible-root,
FAT, native semantic, resource, and performance acceptance for those extended
selectors remain open; no release, latency, allocation, RSS, physical-I/O, or
generic CRUD claim is made by change 0148 itself.
See [change 0148](changes/0148-cfb-same-target-repeat-policy.md).

## CFB same-target repeat release ABBA (change 0149)

Four clean CPU-2 release processes compared the current target-aware policy
with the immediate pre-change production policy in strict
`A1 control, B1 candidate, B2 candidate, A2 control` order. The matrix uses 20
warmups and 200 samples for each of 36 records per leg, retaining 28,800
samples. Both revisions use the exact same harness and deterministic 36-/4095-
byte `many-small` / `wide-root` corpora.

Sequential same-target source work changes from control `[L,R,0...]` to
candidate `[L,L,...]`: the candidate avoids root Mini Stream materialization,
but later calls are exact target reads rather than zero-source cache hits.
Different-SID remains `[D,C,0]`, public multi-MiniFAT bulk changes from
control `{D,C}` to candidate `{C}`, and overlap changes from control `{D,C}`
to bounded `{D,D}` or `{D,C}` candidate outcomes. Output hashes,
source versions, returned lengths, and typed refusal remain exact.

Under the harness-only 100 us fixed latency + 25 us/request, 50 MiB/s, 4 KiB-
range model, aggregate total improvements agree in both ABBA directions:

| Operation | many / 36 p50 | many / 4,095 p50 | wide / 36 p50 | wide / 4,095 p50 |
|---|---:|---:|---:|---:|
| repeat-3 | 61.47% / 61.55% | 60.70% / 60.70% | 64.09% / 64.01% | 63.85% / 63.69% |
| repeat-8 | 58.19% / 58.15% | 55.92% / 55.86% | 63.67% / 63.57% | 63.16% / 63.16% |

P95, p99, and mean agree at roughly the same 56-64% aggregate-total scale.
Configured-simulator one-shot totals remain near neutral. The local
in-memory, per-invocation, bulk, and concurrent distributions are not accepted:
later cache-hit positions deliberately regress, and local special-workload
tails include reversing >5% review triggers with substantial same-
implementation drift. No allocation/RSS, bounded-memory, physical-I/O,
cold/network/device, native-format, or generic performance claim is made. See
the [release record](changes/0149-cfb-same-target-repeat-release-abba.md),
[summary](results/cfb-repeat-abba-0149-summary.json), and retained compressed
raw legs.

## CFB same-target MiniFAT single-flight release ABBA (change 0152)

The final same-target MiniFAT single-flight revision `f46381c6f` (introduced by
`c270c8f3b`) was compared with clean control `e486e4b1` in strict CPU-2
`A1 control, B1 candidate, B2 candidate, A2 control` order. Each leg used 20
warmups and 500 samples across 24 records, retaining 48,000 samples. All
correctness and logical source-event invariants passed. In the existing
concurrent scenarios, the candidate recorded 6,473 logical source calls versus
8,000 for control, a 19.09% reduction.

This accepts only the named source-event/correctness result. At the 0152
revision the 291-name selector matrix was unchanged: no runtime
selector was added; only `cfg(test)` source-event acceptance and tests changed.
Change 0153 adds four RTF selectors measured at the pre-staged
publication-call interval, making that matrix 295. Change 0154 adds six ODF
content-COW publication selectors, making that matrix 301; change 0159 later
made it 302, change 0160 made it 303, change 0162 made it 305, change 0163
made it 309, change 0164 made it 311, change 0166 made it 315, change 0174
made it 319, and change 0175 made the then-current matrix 320. Local or generic latency, allocation/RSS/peak memory, physical
I/O/syscalls, cold-cache/device/network behavior, decompression, native
semantic, OOXML, ODF, RTF, and iWork
claims are withheld. The root MiniStream cache and resource-accounting
boundaries remain, as do broader performance gaps. See the
[change record](changes/0152-cfb-same-target-singleflight-release-abba.md) and
[machine-readable summary](results/cfb-singleflight-abba-0152-summary.json).

## CFB MiniFAT physical-run boundary evidence (change 0125)

The current harness adds a matched 4095-byte MiniFAT boundary pair over the
same 256- and 2,048-sibling shapes. This target is distinct from the accepted
36-byte control: it occupies 64 logical 64-byte mini-sectors (eight regular
512-byte sectors) and therefore exercises
physical root-sector run coalescing. The legacy case materializes the complete
root mini-stream; the positional case records exact source ranges while
filling a 4095-byte caller buffer. Each sample keeps separate open/read/total
timing arrays, source call/byte/range vectors, returned length, and payload
hash. The focused test requires legacy source bytes to exceed 4095 and the
positional source bytes to equal 4095 in one exact request.

This is correctness and request-amplification evidence only. No latency,
tail, physical-I/O, allocation, RSS, cold-cache, high-latency-source, or
semantic native Office claim is accepted until release ABBA and resource
attribution are available. See [change 0125](changes/0125-cfb-minifat-physical-run-evidence.md).

## CFB atomic-save scan evidence

Change 0103 measures the same-length `cfb_file_same_length_overlay_atomic_save`
case across a pinned release before-A/after-A/after-B/before-B run (five
warm-ups and 30 fresh-child samples per leg, CPU 2, warm ext2/ext3). The
atomic `save` path removes only the duplicate post-emission fingerprint scan:
its complete source-scan shape is mechanically `4N -> 3N`. Direct
`write_to` retains its post-emission scan and is unchanged.

| Leg | Revision | Logical reads | p50 | Output |
|---|---|---:|---:|---|
| before-A | `32e5a9f8` | 2,084 calls / 101,751,908 B | 143,425,701 ns | 16,913,408 B, SHA `7994759e...` |
| after-A | `4ededfa2` | 1,825 calls / 84,838,500 B | 148,870,583 ns | same |
| after-B | `4ededfa2` | 1,825 calls / 84,838,500 B | 148,368,923 ns | same |
| before-B | `32e5a9f8` | 2,084 calls / 101,751,908 B | 164,880,142 ns | same |

The exact logical reduction is 16,913,408 bytes (16.6222%) and 259 calls
(12.4280%), with identical output bytes and SHA-256 on every leg. The latency
directions disagree: after-A is +3.7963% versus before-A, while after-B is
-10.0141% versus before-B. This is therefore logical `ReadAt` work and
correctness evidence only; no latency, allocation, RSS, peak-memory,
physical-cold, high-latency, or general semantic CRUD claim is accepted.
Parent-wall and warm process `read_bytes` fields remain descriptive counters,
not speed or storage-device evidence. See the [change record](changes/0103-cfb-atomic-save-scan-evidence.md)
and [compact summary](results/cfb-save-atomic-scan-0112-summary.json).

### Current CFB save phase attribution

[Change 0142](changes/0142-cfb-atomic-save-phase-attribution.md) divides the
selector into open, plan/validation and atomic-publication intervals; the last
retains the three scans in `ValidatedOverlayPlan::save`. No production code
changed. A clean CPU-2 release capture used
20 warm-ups and 200 fresh-child samples in both warm and advisory-cold states.
All 400 samples retained the exact 1,825 calls / 84,838,500 logical bytes and
the same 16,913,408-byte output.

| Phase | Calls | Logical bytes | Warm p50 | Cold-requested p50 |
|---|---:|---:|---:|---:|
| open | 264 | 135,680 | 311,740 ns | 1,418,851 ns |
| plan and candidate validation | 784 | 33,962,596 | 33,442,779 ns | 46,936,548 ns |
| atomic publication | 777 | 50,740,224 | 103,842,832 ns | 86,794,070 ns |
| operation | 1,825 | 84,838,500 | 138,153,550 ns | 135,319,622 ns |

Phase percentiles are independent and do not sum. This is current-revision
attribution, not a speedup result. It identifies fingerprint request
coalescing—not removal of another required scan—as the next bounded A/B
hypothesis. See the [compact record](results/cfb-save-phase-current-0142-summary.json).
The [compressed full capture](results/cfb-save-phase-current-0142.json.zst)
retains the raw aligned filesystem evidence.

### Accepted CFB fingerprint-request coalescing

[Change 0143](changes/0143-cfb-fingerprint-read-coalescing.md) implements the
bounded hypothesis from Change 0142. Complete fingerprint scans use a
right-sized window capped at 1 MiB, while comparison and publication remain at
64 KiB and the buffers never overlap. No fingerprint pass or source-mutation,
candidate-reopen, typed-output or atomic-rename check was removed.

A clean CPU-2 `A1 control, B1 candidate, B2 candidate, A2 control` release run
used 20 warm-ups and 200 fresh-child samples per warm and advisory-cold state in
every leg. All 1,600 samples retained the same 84,838,500 logical bytes, one
changed span and exact 16,913,408-byte output. Logical requests fell from 1,825
to 857 (53.0411%): plan 784 -> 300 and atomic publication 777 -> 293, while
open remained 264.

| Direction / state | p50 improvement | p95 improvement | Mean improvement |
|---|---:|---:|---:|
| A1 -> B1 warm | 3.3327% | 3.0259% | 3.5940% |
| B2 -> A2 warm | 1.3163% | 1.6195% | 1.1008% |
| A1 -> B1 cold-requested | 10.7679% | 13.9112% | 18.3154% |
| B2 -> A2 cold-requested | 9.4641% | 9.0335% | 9.1743% |

The code-local fingerprint window is at most 983,040 bytes larger. A matched
whole-process `/usr/bin/time -v` boundary found no candidate RSS increase
(control 111,640/111,508 KiB; candidate 111,508/111,508 KiB), but this is not an
operation-only allocation or peak-memory measurement. `cold-requested` remains
advisory, and logical `ReadAt` calls are not physical device I/O. See the
[compact summary](results/cfb-fingerprint-abba-0143-summary.json) and
[compressed raw capture](results/cfb-fingerprint-abba-0143.json.zst).

## Parallel scaling observation

This historical `opc_open` experiment used `RAYON_NUM_THREADS` in separate
processes. Current production bulk execution uses caller-sized local pools and
has no hidden global Rayon path; the figures below remain historical rather
than current-HEAD scaling evidence. Each cell used 10 warm-ups and 50 samples.
Raw reports are the
[`results/baseline-opc-open-workers-*.json`](results/) files.

| Corpus | 1 worker p50 | 2 workers | 4 workers | 8 workers | 12 workers | Best observed speedup |
|---|---:|---:|---:|---:|---:|---:|
| 256 small, compressible | 630 us | 539 us | 497 us | 525 us | 549 us | 1.27x at 4 |
| 256 small, incompressible | 590 us | 511 us | 485 us | 505 us | 507 us | 1.22x at 4 |
| four 4 MiB, compressible | 5.42 ms | 2.64 ms | 434 us | 428 us | 697 us | 12.7x at 8 |
| four 4 MiB, incompressible | 6.28 ms | 3.02 ms | 664 us | 671 us | 662 us | 9.5x at 12, effectively flat from 4 |

Four workers match the four large payload Parts and are the practical knee on
this host. More workers do not improve the many-small case and increase its
tail/median latency. This is evidence for a bounded explicit execution context
with task-size thresholds; it is not evidence for retaining an implicit global
pool.

## CPU, syscalls, locks, and unavailable counters

Linux `perf stat` and sampled `perf record` are denied by the host policy, so no
cycles, instructions, cache-miss, or branch-miss claim is made. A Valgrind
Callgrind fallback on five many-small saves recorded 1.624 billion interpreted
instructions for the whole process; optimized/inlined Rust symbols made the
fine-grained CPU attribution insufficient for an optimization claim. The
allocation profile is the useful attribution for the first change.

The measured input and output are memory-backed. No filesystem I/O occurs in a
timed operation. A process-level `strace -f -c` is preserved in
[`results/baseline-opc-many-small.strace.txt`](results/baseline-opc-many-small.strace.txt),
but it includes Git/Rust environment probes, JSON publication, and global
Rayon initialization. Its 65 `futex` calls cannot be attributed to the timed
save loop. The ordinary OPC-open path bypasses the lazy ZIP cache, and the save
path has no Part cache, so hit/miss, eviction, duplicate-flight, and cache-lock
metrics are not applicable to these cases. Dedicated lazy-reader concurrency
and source-backed range-I/O scenarios remain required.

## Ranked result and next gate

1. Implement and measure one pre-output `PublicationPlan`. It removes the
   duplicated sort/content-type/relationship serialization proven by
   Heaptrack while preserving all validation and sequential-sink behavior.
2. Design source-backed lazy OPC and raw-copy unchanged ZIP entries. The full
   16.8 MB no-op rewrite shows their potential, but both require a larger
   preservation/security change and must not be folded into the small plan.
3. Refresh current-HEAD explicit local-pool scaling with task-size thresholds;
   the historical knee was four large-entry tasks on this host.
4. Add format-owner and CFB matrices before choosing XLSX/DOCX/PPTX/legacy
   semantic optimizations.

An optimization is accepted only after the same hashes, sink byte/write
summary, correctness suite, and before/after measurement protocol pass. A
latency-only movement inside overlapping uncertainty is not sufficient.

## Implemented follow-up results

Four measured change records now extend this original baseline:

The aggregate outcome, verification gates, disclosed regressions, and
remaining program scope are summarized in [`REPORT.md`](REPORT.md).

1. [`changes/0001-opc-publication-plan.md`](changes/0001-opc-publication-plan.md)
   removes duplicated OPC publication planning: -37.0% allocation calls and
   -5.49% mean latency on the intended 2,048-Part compressible save.
2. [`changes/0002-cfb-lookup-and-sector-buffers.md`](changes/0002-cfb-lookup-and-sector-buffers.md)
   uses cached validated name keys and bounded reusable sector buffers:
   successful final-stream lookup is 56-66% faster at 256 siblings and about
   94% faster at 2,048, with 6-9% fewer open-process allocations.
3. [`changes/0003-legacy-owned-stream-handoff.md`](changes/0003-legacy-owned-stream-handoff.md)
   retains PPT (-20.2% p50, -12.4% peak heap) and XLS (-9.5% peak heap)
   ownership transfers. The DOC variant regressed 58.4% and was removed.
4. [`changes/0004-opc-exact-owned-source.md`](changes/0004-opc-exact-owned-source.md)
   makes unchanged owned OPC output byte-exact and avoids complete
   recompression: the 16.78 MB case falls from 211.5 ms to 3.44 ms. Retaining
   that compressed source increases the large profile's peak heap by 22.6%,
   so lazy Part materialization remains the next architectural dependency.

The original stage-1 harness had 14 cases and 97 default result records. In
addition to the original matrix it measured owned OPC open, one-Part mutated
save, and public
DOC/XLS/PPT writer packaging with tiny, moderate, and 4-5 MiB stream-heavy
shapes. Scheduled CI records the deterministic full matrix without applying
machine-noisy latency thresholds.

## Current stable tranche update

The stage-1 records above are retained unchanged. The current harness has
**393 selectable cases**; 200 was the count before the opt-in ODF `mimetype`
repair-plan selector and later opt-in selectors were added. The
historical 36-default-case/198-default-record tranche remains measured as
documented below; newer selectable cases do not inherit those measurements.

Change 0258 adds the opt-in `rtf_file_open` and
`rtf_file_open_lifecycle` selectors over plain, CP-1252, LZFu, and MELA
variants. The non-plain variants are correctness-only because the historical
facade control cannot open them. Back-to-back matched plain
20-warmup/500-sample ABBA captures fail cross-run reproducibility, so the
retained evidence supports no latency claim and does not change the historical
default tranche.

Four opt-in XLSB lifecycle selectors cover fresh open, worksheet listing, one
cell, and a prepared full `worksheet.cells()` scan over deterministic tiny,
medium, large, and sparse BIFF12 corpora. The scan timer excludes archive
cloning and workbook/worksheet preparation, while exact canonical cell hashes
and counts are verified outside timing. The default 36-case/198-record tranche
is unchanged.

Change 0188 adds eight opt-in DOCX/PPTX fresh-open-plus-query lifecycle
selectors. A CPU-2 release A1-eager/B1-source/B2-source/A2-eager run with 30
retained warm fresh-child samples per case records lower directional values but
accepts no latency statistic: source-backed PPTX and paragraph-count p50/mean
drift plus eager full-text p50/mean drift miss the predeclared gates, and every
tail is conservatively withheld. This is correctness/attribution evidence, not
physical-I/O, cold-cache, allocation/RSS, edit/save, producer, or broad OOXML
evidence.

Change 0189 adds four opt-in XLSX edit-composition selectors for disjoint join,
recoverable overlap, disjoint three-way planning, and explicit conflict
resolution. Its two-shape debug smoke is correctness and phase evidence only;
no latency, allocation, memory, I/O, source-backed, or filesystem-save result
is accepted.

Change 0160 adds one opt-in native DOC owner/public-reader attribution case.
For each retained sample it records strict-owner, complete public-reader,
exact-source retention, edit construction, replacement staging, in-memory
owner rendering, final owner/public validation, patch construction, outer
operation, output-materialization, and checked unattributed intervals. It
reuses the exact deterministic tiny, large, and payload-heavy writer bytes.
All semantic, no-op, patch/inverse/stale, malformed/typed-refusal, hash, and
untouched-stream checks are outside timing. Successful event-order/cardinality
validation follows the named outer interval but remains inside the complete
lifecycle timer and therefore its checked unattributed remainder. Separate
format tests bind balanced error events. A clean release run at exact revision
`ab333008d3`, pinned to CPU 2 on the named AMD EPYC 9575F host, used four fresh
processes per shape, 20 warmups and 200 retained samples per process. Lifecycle
p50 was 0.081 ms tiny, 1.157 ms large, and 44.227 ms payload-heavy. The grouped
initial-plus-final complete public-reader validation p50 was 0.016, 0.598, and
20.721 ms respectively; patch p50 was 0.026, 0.165, and 8.413 ms. Every untimed
case-level gate passed in all 12 reports, and all 2,400 timed samples passed
arithmetic, event, and output checks. Lifecycle p50/mean spread across
processes remained below 3.0%/3.8%; two tiny subphase means crossed the 5%
review trigger without changing rank. This accepts only
the exact phase distribution, not an optimization or speedup. Physical-I/O,
allocation/RSS, cold-cache, and real-producer results remain open. See
[`0160`](changes/0160-doc-owner-public-phase-attribution.md), the
[summary](results/doc-owner-public-phases-0160-summary.json), and the
[raw-artifact manifest](results/doc-owner-public-phases-0160.sha256).

Change 0161 tested the smallest direct follow-up: borrow the DOC bytes during
initial/final public-reader validation instead of cloning them. Clean release
`A1 control, B1 candidate, B2 candidate, A2 control` processes on CPU 2 used
20 warmups and 500 samples for all three shapes. Tiny lifecycle p50 improved
3.20%/3.24%, but large regressed 3.06%/7.31% and payload-heavy directions were
-0.18%/+2.52%. Large p95 regressed 37.52%/14.49%. The candidate was rejected
and fully removed; production remains the control. See [change 0161](changes/0161-doc-public-validation-borrow-rejected.md)
and its [summary](results/doc-public-borrow-0161-summary.json).

Change 0162 adds two opt-in RTF standalone-picture CRUD selectors over a
dedicated generated ASCII/uncompressed corpus with 2/8/64 alternating PNG and
JPEG groups. Replacement changes 1/7/63 same-length payloads while leaving one
group unselected; removal deletes 1/4/32 alternating groups. Independent raw
splices preserve mixed-case hexadecimal digit slots, whitespace, surrounding
source and every unselected group. Open, bounded batch staging, commit,
fixed-memory hashing-sink publication and complete lifecycle are reported
separately. A focused test and six-record debug smoke pass semantic reopen,
no-op, volatile/durable forward/inverse, stale/foreign, refusal, partial/zero
sink and digest gates. This raises the selectable matrix to 305 without
changing the default 36/198 tranche. No debug latency, allocation/RSS,
physical-I/O, real-producer or broad RTF media claim is accepted. See
[`0162`](changes/0162-rtf-picture-crud-evidence.md).

Change 0163 adds four opt-in XLSX scalar-cell lifecycle selectors: eager and
positional source-backed clear, plus eager and positional source-backed remove.
They reuse the existing deterministic medium and dense/sparse four-worksheet
numeric corpus and target one existing `Sheet1!A1` owner. Clear retains an
empty `<c>` owner; remove deletes that owner. Open, planning/staging, commit,
sequential publication and lifecycle vectors are separate. A fixed 64-KiB
windowed hashing sink retains zero output bytes; generic logical source and
materialization counters are recorded, with eager counters explicitly
not-applicable. Semantic, package, exact no-op, volatile source-patch,
stale/foreign, and source-backed raw-unselected-member gates remain outside
timing. The four selectors raise the matrix from 305 to 309 without changing
the default 36/198 tranche. This is correctness/phase/counter evidence only:
no latency, allocation/RSS, physical-I/O, cold-cache, decompression,
durable-source-patch, or real-producer claim is accepted. See
[`0163`](changes/0163-xlsx-cell-clear-remove-evidence.md).

Change 0164 adds two opt-in RTF ordinary-paragraph structure selectors:
`rtf_semantic_split_paragraph_save` and
`rtf_semantic_merge_paragraph_save`. Both reuse the exact generated plain
lifecycle corpus at tiny/medium/large sizes (24/200/10,000 paragraphs). Split
inserts one canonical five-byte `\\par ` boundary at a checked interior
source position; merge removes only the authenticated adjacent boundary, so
their independent raw expected outputs are respectively five bytes larger and
smaller. The selectors report separate open, stage, commit, publication and
lifecycle vectors and publish through a fixed 16-KiB windowed hashing sink
that retains zero output bytes. Untimed gates cover semantic reopen, exact raw
splice and unchanged surrounding bytes, exact no-op/source identity, volatile
and deterministic durable patch forward/inverse, stale/foreign refusal,
bounded refusal cases, partial/zero sinks and source/output hashes. The
selector gate also verifies `forged_result_artifact_refusal_verified`; the
existing focused RTF tests remain the authority for exact boundary-byte
restoration and forged-boundary precondition refusal. The two selectors raise
the matrix from 309 to 311 without changing the default 36/198 tranche. This is
correctness,
phase and sequential-sink evidence only: no latency, speedup, transaction
memory, allocation/RSS, physical-I/O, cold-cache, source-backed,
real-producer, or general rich-RTF claim is accepted. See
[`0164`](changes/0164-rtf-paragraph-split-merge-evidence.md).

Change 0165 records the DOC lazy-fingerprint and same-lineage patch-replay
implementation plus a bounded descriptive comparison on the exact deterministic
tiny, large, and payload-heavy native DOC owner/public-reader lifecycle. `Snapshot` keeps its FNV-1a diagnostic value
in an inline lazy `OnceLock`; patch construction no longer scans complete
before/after artifacts, and immutable `Arc` identity plus length lets
same-lineage no-op/apply paths return retained snapshots. Independently
reopened sources still perform the lazy fingerprint check followed by exact
byte comparison, so the fingerprint is not an authorization boundary. The
`source_fingerprint` and `target_fingerprint` accessors are intentionally
non-`const` because their first call may initialize the cache.

The final clean control revision is
`d6818e290aa77fd7666b7b16ee6908319d0f332b`; the candidate is
`5dd813b1e108e253457ccb6c504c125c2becc1c6`. Their release binaries are
identified by SHA-256 `344c0504c254109ee6b4361e375599d187f8a12333abb44f207d837af259ef8c`
and `c95e6c6004cbd725c789597566a81c0897ab6915ecd7c274deab222d134b3fd3`,
respectively. Both builds were clean exact-revision builds.

The original `measured_total_ns` lifecycle boundary is unchanged. Same-lineage
apply and the first source/target fingerprint demand are explicit workflow
extensions. Clean CPU-2 release `A1 control, B1 candidate, B2 candidate, A2
control` runs used 20 warmups and 500 retained samples per shape and leg, for
6,000 lifecycle samples. Descriptive lifecycle p50/mean/p95 positive-faster deltas were
`+33.77/+35.19/+38.94` and `+33.21/+34.76/+39.67` tiny,
`+12.28/+12.59/+17.53` and `+13.81/+13.55/+11.68` large, and
`+17.33/+17.09/+16.58` and `+17.82/+17.75/+16.25` payload-heavy. With
immediate fingerprint demand included, workflow p50/mean/p95 positive-faster deltas are
`+14.56/+16.34/+22.24` and `+13.89/+15.80/+21.90` tiny,
`+4.50/+4.82/+10.24` and `+5.83/+5.64/+4.26` large, and
`+6.55/+6.41/+6.26` and `+7.08/+7.08/+6.33` payload-heavy.

The isolated edit-patch/same-lineage-apply extension is approximately
99.6-99.99% across the reported p50/mean/p95 deltas versus the eager-fingerprint
control. The deferred first
fingerprint demand is explicit rather than hidden: roughly 20-170 ns in the
control boundary versus 25.7 us, 164 us, and 8.37-8.39 ms for the candidate's
tiny, large, and payload-heavy source-plus-target scans. Same-implementation
lifecycle drift is disclosed in the change record; paired directions remain
positive but are not generalized beyond the named host and corpora.

Mandatory DOC no-op, one-edit, and open guards remain within the declared
policy: p50 no-op is `+78.84%/+79.89%` tiny and `+71.08%/+70.40%` large;
one-edit is `+37.23%/+40.81%` and `+20.45%/+19.79%`; open is
`-3.52%/+0.13%` tiny and `+0.55%/-1.80%` large. Neighboring XLS one-edit
and open guards are mostly neutral or improved, while XLS no-op remains
directionally noisy. A representative final payload heaptrack probe records
50,677 allocation calls and 128.28M peak heap for both revisions, with
profiler RSS 145.14M versus 142.81M; a 30-sample `/usr/bin/time` boundary
records `138160/138024/138028/138032 KiB` in A1/B1/B2/A2 order. These are
descriptive whole-process probes, not operation-only allocation or total-memory
claims. No speedup, physical-I/O, cold-cache, real-producer, generic-DOC, or
CRUD-completeness claim is accepted. See
[`0165`](changes/0165-doc-lazy-fingerprint.md), the
[summary](results/doc-lazy-fingerprint-0165-summary.json), and the
[release manifest](results/doc-lazy-fingerprint-0165-manifest.json).

Change 0167 removes one redundant semantic worksheet reload, cell-store parse,
and row-tag scan from matched source-backed XLSX row-visibility publication by
reusing the existing cell-values lineage/version proof. The mandatory OPC
overlay validation and selected-member read remain. Clean CPU-2 release
`A1/B1/B2/A2` runs used 20 warmups and 500 retained samples for medium/large
hide-one and unhide-256. Descriptive publication p50/mean/p95/p99 reductions
span 50.42%-68.23% and agree in both paired directions, while logical source
reads remain exactly 204/209 calls with one/six selected-worksheet overlaps.
The 5% stability gate fails: maximum absolute drift is 34.80% for control
large/unhide publication p99 and 10.23% for candidate medium/hide complete-
workflow p50; first-pair medium hide/unhide complete-workflow p99 regresses
6.95%/2.69%. Therefore no acceptance-grade end-to-end latency, tail, allocation,
RSS, or physical-I/O claim is made. See
[`0167`](changes/0167-xlsx-row-visibility-provenance-reuse.md), the
[summary](results/xlsx-row-visibility-provenance-0167-summary.json), and the
[manifest](results/xlsx-row-visibility-provenance-0167-manifest.json).

Change 0168 removes two redundant complete source scans from native XLS
fixed-width Number/RK/MulRK plan-only commit. BIFF semantic owner validation now
runs on the exact composed view after CFB reopen/range checks and before CFB's
final source/target fingerprint fence. Number therefore avoids 33,991,680
logical source bytes and 34 one-MiB reads per effective sample; RK/MulRK avoids
405,504 bytes and two reads. These are code-derived in-memory scan counts, not
physical-I/O measurements. Clean CPU-2 release A/B/B/A runs used 20 warmups and
500 samples per family. Complete-workflow p50/mean/p95/p99 values are
descriptively 19.22%-28.16% lower and semantic-commit values 37.58%-48.04%
lower in both paired directions, but same-implementation drift reaches 10.56%
for control and 9.81% for candidate. The 5% gate fails, so no acceptance-grade
latency, tail, allocation/RSS, peak-memory, physical-I/O, cold-cache, or
producer claim is made. See
[`0168`](changes/0168-xls-numeric-validation-fusion.md), the
[summary](results/xls-numeric-validation-fusion-0168-summary.json), and the
[manifest](results/xls-numeric-validation-fusion-0168-manifest.json).

Change 0169 removes transient owned-node-vector construction from cumulative
hierarchical budget charges and retains up to four releasable reservation nodes
inline. The existing one-sheet `xlsx_streaming_create` selector supplied the
measured scale path; no selector or schema changed. Clean CPU-2 release A/B/B/A
runs used 20 warmups and 200 samples per tiny/medium/large shape. Medium and
large p50/mean/p95/p99 improve in both paired directions by 1.05%-9.76%; tiny
p50/mean/p95 also improve, while tiny p99 regresses 1.81%/2.75% and is withheld.
Same-implementation drift stays inside the predeclared 5%/10%/15% tiers.
Matched whole-process Heaptrack captures record 48.81% fewer allocation calls
and 69.77% fewer temporary allocations with unchanged 225.45M peak heap; RSS
directions disagree. Exact archive/worksheet hashes, rows/cells, logical sink
counters, zero retained output, and the 4 KiB authoring window remain fixed.
This is warm in-memory synthetic one-sheet creation evidence, not a total-memory,
physical-I/O, cold-cache, multi-sheet, producer, or every-`Budget` claim. See
[`0169`](changes/0169-xlsx-streaming-budget-charge.md), the
[summary](results/xlsx-stream-budget-charge-0169-summary.json), and the
[manifest](results/xlsx-stream-budget-charge-0169-manifest.json).

Change 0170 batches ordinary XLSX streaming text as contiguous UTF-8 runs
between XML entities, skips scalar counting when the byte count already proves
the 32,767-character bound, and formats each row number once. The existing
selector and corpus are unchanged. Clean CPU-2 release A/B/B/A runs used 20
warmups and 300 samples per tiny/medium/large shape. Large p50/mean/p95/p99
improve in both directions by 5.02%-6.99%; medium p50/mean/p95 by 4.45%-5.52%;
tiny p50 by 5.03%/7.74%. Tiny mean/p95/p99 and medium p99 are withheld because
paired directions disagree. Exact worksheet/archive hashes, rows/cells, sink
counters, zero retained output, and the 4 KiB row window remain fixed. Matched
whole-process instructions and branches fall, but branch misses regress and no
allocation/RSS/total-memory/I/O claim is made. See
[`0170`](changes/0170-xlsx-streaming-escape-runs.md), the
[summary](results/xlsx-stream-escape-0170-summary.json), and the
[manifest](results/xlsx-stream-escape-0170-manifest.json).

Change 0171 moves source-backed DOC paragraph, PPT shape-text, and XLS
worksheet-visibility semantic readback onto the exact composed CFB view already
created by the common planner's owner callback. Each effective transaction
therefore removes one complete artifact scan, `ceil(artifact_bytes / 1 MiB)`
logical reads, and one source/target SHA-256 pair while retaining CFB reopen,
format-owner validation, the final complete fingerprint fence, publication,
and atomic-save checks. On the measured 2,135,552-byte XLS corpus this is one
scan and three logical reads. Clean CPU-2 release A/B/B/A runs used 20 warmups
and 300 samples. The 64-worksheet source-backed complete workflow improves
p50/mean/p95 by 12.51%-15.38% in both directions; scalar and batch semantic
staging/plan p50/mean/p95 improve by 31.44%-33.16%. Scalar total, p99,
publication, DOC/PPT latency, allocation/RSS, physical-I/O, cold-cache, and
producer claims are withheld. See
[`0171`](changes/0171-cfb-owner-validation-fusion.md), the
[summary](results/cfb-owner-fusion-0171-summary.json), and the
[manifest](results/cfb-owner-fusion-0171-manifest.json).

Change 0172 carries the immutable `Arc<[u8]>` proof held by native XLS
plan-only numeric snapshots into the CFB owner. Only direct sequential
`write_to` uses that private provenance: it removes the complete pre-emission
and post-emission fingerprint scans while retaining the 64 KiB emission pass,
source and target SHA-256, exact progress/partial-output handling, and flush.
Generic positional sources, composed views, and atomic saves retain their
existing fences. The code-derived reduction is 33,991,680 logical bytes/34
one-MiB reads for Number and 405,504 bytes/two reads for RK/MulRK.

Clean CPU-2 release A/B/B/A runs used 20 warmups and 500 samples. Number
complete-workflow p50/mean/p95/p99 improves by 37.54%-39.00% and direct
publication by 64.44%-65.63%; RK/MulRK complete workflow improves by
36.63%-38.96% and publication p50/mean/p95 by 65.54%-66.76%. Every accepted
statistic agrees in both directions and passes the 5% same-implementation
drift gate. RK/MulRK publication p99 is withheld because control drift is
5.28%. Allocation/RSS, physical-I/O, cold-cache, producer, compression and
atomic-save claims are withheld. See
[`0172`](changes/0172-cfb-owned-numeric-publication.md), the
[summary](results/cfb-owned-numeric-publication-0172-summary.json), and the
[manifest](results/cfb-owned-numeric-publication-0172-manifest.json).

Change 0173 applies both proven CFB seams to native XLS existing-comment
publication. Semantic readback now consumes the composed view inside the
planner's final fingerprint bracket, and the immutable snapshot enters through
the sealed owned-byte path. Each effective scalar or 256-comment transaction
therefore removes three complete 16,995,840-byte scans, 51 one-MiB logical
reads, and three source/target digest pairs while retaining 64 KiB emission
hashing and every atomic-save fence.

Clean CPU-2 release A/B/B/A used 20 warmups and 500 samples. The scalar
complete workflow p50/mean/p99 is 45.54%-47.19% lower, semantic staging/plan is
30.78%-32.42% lower, and direct publication is 59.15%-61.03% lower. The
256-comment semantic phase is 30.53%-32.57% lower. Scalar complete p95 and
batch complete/publication are withheld by the predeclared 5% drift/guard
policy. Allocation/RSS, physical-I/O, cold-cache, producer, compression and
atomic-save claims remain open. See
[`0173`](changes/0173-cfb-comment-publication-fusion.md), the
[summary](results/cfb-comment-fusion-0173-summary.json), and the
[manifest](results/cfb-comment-fusion-0173-manifest.json).

Change 0117 adds eight opt-in native PPT `Pictures` selectors and two pinned,
balanced release attempts. The matched corpus has eight slides and 32
deterministic 256 KiB PNG records. Source-backed timed samples use
uninstrumented `OwnedSource`; separate untimed replays prove that open overlaps
zero `Pictures` bytes, the cold query reads the complete stream once, and a
cached query reads nothing further. Both latency attempts were rejected:
same-implementation p50 or p95 drift exceeded the predeclared 5%/10% limits in
every phase, including the directly timed fresh open-plus-all-images case.
The raw distributions and whole-process RSS observations are retained, but no
latency, allocation, RSS-attribution, cold-cache, or save-path result is
accepted. See
[`0117`](changes/0117-ppt-pictures-release-evidence.md) and the
[raw report](results/ppt-pictures-release-0117.json).

Change 0119 adds three opt-in native PPT selected-shape controls: a positional
query-only pair against the existing eager case and a fresh-open-plus-query
eager/source-backed pair. Independent untimed replays retain deterministic
logical source-read counters and semantic hashes while timed source-backed
samples remain uninstrumented. No latency or resource claim is accepted
without a frozen release ABBA run.

Change 0120 adds eight opt-in PPTX ordinary-root filesystem selectors over the
same fixed 200-slide/eight-text-box/eight-2 MiB-media corpus: matched eager and
source-path open, full owned slide listing, slide-count, and selector-first
slide-100 queries. The source candidate calls the unified
`litchi::Presentation::open(path)` path; the eager open control times
`fs::read` plus `Presentation::from_bytes`, while query roots are prepared
before timing. Every sample runs in a fresh warm/cold-requested child and
checks source hash, full eager/source semantic parity, metadata, slide size,
slide names, text hashes, and corpus length outside timing. Each measured source
sample also receives one separate untimed `SourceBackedPresentation` replay with
exact compressed ZIP range classification: open/count have zero slide/media
overlap, selected slide 100 has no unselected-slide/media overlap, and listing
reads all slide payloads without media. Eager controls explicitly have no
`ReadAt` replay; the generic filesystem counter scope marks them not applicable.
This change enables correctness and logical-read evidence only. It makes no
latency, tail, allocation, RSS, decompression, physical-I/O, or cold-cache
claim before a frozen release ABBA run. See
[`0120`](changes/0120-pptx-root-source-path-evidence.md).

Change 0121 adds two opt-in native PPT repeated selected-shape controls,
bringing the matrix to 229 names at that point (before changes 0122-0124) while
preserving the default 36-case / 198-record tranche. Each matched eager/source-backed control keeps
one prepared owner and issues eight identical selected-shape queries; source
timing uses an uninstrumented source and separate replays record exact logical
calls, bytes, prior-covered bytes per later logical read, and a canonical
semantic digest. The
production regression binds 74 calls / 8,310 bytes for legacy
two-query CFB reconstruction and 66 calls / 3,190 bytes with a retained parsed
CFB index. These are logical-I/O and correctness figures only, not latency or
resource claims.

Changes 0122, 0123, and 0124 add four ODP media-rich, four ODP unified-root
filesystem, and six ODS unified-root/source selectors respectively. They move
the selectable matrix from 229 to 233, 237, and finally 243 names while
preserving the default 36-case / 198-record tranche. These are correctness and
logical compressed-range evidence only: corpus/file publication and owner
preparation stay outside each named timer, and complete semantic, metadata,
member, media, and hash parity remains untimed. No latency, physical-I/O,
decompression, allocation, RSS, or release claim is accepted. See
[`0122`](changes/0122-odp-media-source-read-evidence.md),
[`0123`](changes/0123-odp-unified-root-filesystem-evidence.md), and
[`0124`](changes/0124-ods-unified-root-filesystem-evidence.md).

Change 0125 adds two matched 4095-byte MiniFAT boundary selectors, bringing
the current selectable matrix to 245 names while preserving the default
36-case / 198-record tranche. Their focused evidence records exact source
calls, bytes, physical range sizes, payload hash, and separate open/read/total
timing; it makes no release speed claim.

Change 0126 adds eight ordinary-root DOCX filesystem selectors, bringing the
selectable matrix from 245 to 253 names while preserving the default 36-case /
198-record tranche. The fixed corpus is the existing 200-paragraph,
eight-incompressible-2 MiB-media source-edit corpus and its bytes/hash are
unchanged. Eager open times `fs::read` plus `Document::from_bytes` while source
open times `Document::open(path)`; query roots are prepared outside timing and
only the named paragraph-count, paragraph-list, or full-text query is timed.
An independent untimed typed DOCX source replay records calls, bytes, request
sizes, compressed-range coverage, and materializations: open has zero
main/media/unselected/core payload overlap; for query selectors, preparation
completely covers the compressed main-document range; and the query has zero
subsequent main/media/unselected/core overlap. Untimed parity covers paragraphs, full
text, tables, elements, and metadata; exact source SHA plus logical OPC
part/relationship/content-type/blob-hash gates cover package preservation,
including media hashes and source immutability. This is
correctness/logical-range evidence only; it makes no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, ABBA, broad-security, or Markdown
performance claim. See
[`0126`](changes/0126-docx-root-source-path-evidence.md).

Change 0127 adds two matched ODS repeated-cell sweep selectors, bringing the
selectable matrix to 255 names while preserving the default 36-case /
198-record tranche. The fixed corpus is the existing two-sheet 32 by 32
media-rich ODS archive. Each owner is opened before timing; the timer covers
four identical row-major sweeps, including the adaptive locator threshold
transition. An independent instrumented source replay per measured sample
resets counters after preparation and requires zero reads during the sweep.
Digest/count, source SHA, member topology, semantic grid, manifest media type,
and retained-media payload checks remain outside timing. This is
correctness/logical-read evidence only; it makes no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, ABBA, or release claim. See
[`0127`](changes/0127-ods-source-cell-sweep-evidence.md).

Change 0134 adds matched eager/source-backed ODS ordered cell-batch sweep
selectors over that same corpus. Owners and 2,048 borrowed selectors are
prepared before timing; each timed sample contains four bounded `cell_batch`
calls and 8,192 black-boxed result slots. Independent source replay records
exactly eight version observations and zero post-preparation payload reads per
four-call sweep, while ordered digest/count and source/member/media identity
gates remain untimed. The additions bring the current selectable matrix to
257 names without changing the default 36 cases / 198 records. This is
correctness/logical-read evidence only; no release speed or resource claim is
made without ABBA. See
[`0134`](changes/0134-ods-source-cell-batch-sweep-evidence.md).

The four 0135 selectors bring the current selectable matrix to 261 names while
the default 36 cases / 198 records remain unchanged. Change 0137 adds two
additional opt-in plan-only selectors, bringing the selectable matrix to 263
names without changing that default.

Change 0135 adds four matched eager/source-backed native XLS fixed-width
numeric selectors. The Number pair reuses `Untouched!E21` (`42` -> `43`) from
the deterministic comments corpus; the RK/MulRK pair uses one standalone RK
and one two-cell MulRK record in a deterministic native corpus and edits all
three values in one transaction. The timer separates transaction creation,
`set_number`/`set_numeric`, eager/source-backed commit, and complete publication
to the same preallocated bounded sink. Complete target materialization is
reported on both paths because source-backed commits retain a reopened target
snapshot. Source ingress, no-op/fingerprint, patch/inverse/stale,
security/unsupported refusal, complete Snapshot/Workbook reopen, untouched
CFB topology/member bytes, and the untimed 54016.xls real-producer gate remain
outside timing. This is correctness/coverage evidence only: no positional-I/O,
allocation/RSS, bounded-artifact-memory, speedup, or broad-producer claim is
made. See [`0135`](changes/0135-xls-numeric-source-publication.md).

Change 0136 binds those four selectors to a clean-revision, CPU-2-pinned
release baseline at `9577cd16f` with 20 warmups and 200 samples per case:

| XLS fixed-width numeric selector | p50 | p95 | p99 | mean | commit p50 | publication p50 | complete target retained |
|---|---:|---:|---:|---:|---:|---:|---:|
| eager Number | 31.492 ms | 34.116 ms | 35.916 ms | 31.763 ms | 30.741 ms | 0.729 ms | 16,995,840 B |
| source-backed Number | 146.410 ms | 149.108 ms | 150.693 ms | 146.642 ms | 101.618 ms | 44.783 ms | 16,995,840 B |
| eager RK/MulRK | 0.100 ms | 0.120 ms | 0.127 ms | 0.103 ms | 0.097 ms | 0.003 ms | 202,752 B |
| source-backed RK/MulRK | 1.627 ms | 1.659 ms | 1.690 ms | 1.630 ms | 1.117 ms | 0.509 ms | 202,752 B |

The source-backed/eager p50 ratios are 4.65x and 16.25x respectively, with
byte-identical output within each family. This is a descriptive before
baseline, not an optimization or regression classification: all four paths
retain a complete target, source ingress and verification are untimed, and the
single-process run has no allocation, peak-memory/RSS, hardware-counter,
physical-I/O, cold-cache, or fresh-process evidence. The raw schema-1 artifact,
exact binary/result hashes, environment, commands, and interpretation are in
[`0136`](changes/0136-xls-numeric-current-revision-baseline.md).

Change 0137 adds matched plan-only Number and RK/MulRK selectors over the same
corpora and edits. Their commit timer includes validated overlay-plan
construction and composed semantic validation, while publication remains a
separate complete `write_to` interval. The plan retains only the immutable
source, checked overlay plan and bounded numeric splices; it retains and
materializes no complete target artifact at commit. Evidence records zero
`complete_target_materialized_bytes`, explicit false target-retention and
target-materialization flags, and complete published sink bytes. Because this
forward-only API does not expose the ordinary artifact patch, its evidence
marks patch/inverse support false; exact source/target fingerprint preflights,
forward reopen, topology, security, no-op, partial-sink and 54016.xls producer
gates remain required. Composed semantic validation may allocate/read a
candidate Workbook model, so zero retained target-artifact bytes is not a
bounded total-memory claim. This
is correctness/descriptive evidence only and does not claim a latency,
allocation, RSS, I/O, or memory improvement before balanced release ABBA.
See [`0137`](changes/0137-xls-numeric-plan-only-publication.md).

Change 0138 records the balanced CPU-2 release comparison for those two
plan-only selectors. Each family ran one process per leg in strict `A1, B1,
B2, A2` order with 20 warmups and 200 samples; A is ordinary source-backed
publication and B is plan-only. Number total p50 improves 27.57% and 28.58%
in the two paired directions; RK/MulRK improves 24.90% and 24.56%. P95,
p99 and mean move in the same direction for both families. The commit phase
also agrees; publication is near-neutral and is not claimed independently.
Matched three-warmup/30-sample `/usr/bin/time -v` legs show process VmHWM
falling 10.73% and 10.66% for Number in both directions, while RK/MulRK
directions disagree. Valid heaptrack 1.5.0 profiles report whole-process
allocation/temporary totals and identical 205.56/154.93 MiB peak heaps for
the Number/RK families' A/B pairs; no operation-only allocation or peak-heap
improvement is accepted. The plan-only latency result is accepted only for
these two deterministic fixed-width families and this release configuration;
no bounded-total-memory, physical-I/O or cold-cache claim is made. See
[`0138`](changes/0138-xls-numeric-plan-only-release-abba.md) and its schema-1
raw artifacts.

Change 0139 adds two opt-in source-backed ODP repeated-text selectors, bringing
the selectable matrix to 265 names while preserving the default 36-case /
198-record tranche. Both selectors use the same deterministic 12-slide,
eight-picture, 16 MiB-uncompressed-media corpus and prepare the
`SourceBackedPresentation` owner plus four output slots outside timing. The
control reconstructs the pre-cache public sequence (`slides()` plus filtered
`Slide::all_text()` joined with exact `\n\n`, followed by the trailing source
check); the candidate calls `SourceBackedPresentation::text()` four times.
The timer contains only those projections. Untimed instrumented replays record
preparation and post-preparation source evidence; the four-call replay is
required to have zero reads, bytes, compressed-range overlap, and `Pictures`
payload reads, with freshness vectors `[3,3,3,3]` for control and
`[3,5,2,2]` for candidate (12 observations total each). Archive topology,
media/text parity, and digest gates remain outside timing. Preparation
compressed-range overlap is recorded separately and is not interpreted as
media materialization. This is correctness/logical-replay evidence only: no
latency, physical-I/O, decompression, allocation, RSS, cold-cache, ABBA, or
release claim is made until a frozen measured ABBA run. See
[`0139`](changes/0139-odp-repeated-text-cache-evidence.md).

Change 0140 records a clean-revision CPU-2 `A1, B1, B2, A2` release run for
those two selectors with 20 warmups and 200 samples per fresh process. Cached
four-call p50 improves 45.80%/46.32%, p95 improves 45.25%/45.83%, p99 improves
39.91%/45.41%, and mean improves 45.74%/46.33% in the paired directions.
Matched Heaptrack 1.5.0 profiles (three warmups/30 samples) record deterministic
whole-process allocation-call reductions of 14.31% and temporary-allocation
reductions of 17.25%, with unchanged 89.22M peak heap. Matched process VmHWM
is neutral (0.00%/0.16%), so no peak-heap or RSS reduction is accepted. Exact
archive/text/media hashes, zero post-preparation reads, and freshness vectors
remain green on every raw record. This accepted result is limited to the exact
four-call prepared source-backed projection shape; it makes no single-call,
open, physical-I/O, decompression, cold-cache, operation-local allocated-byte,
or generic ODF claim. See
[`0140`](changes/0140-odp-repeated-text-cache-release-abba.md) and its linked
schema-1 raw artifacts.

The earlier five-case filesystem smoke exercises eager/source-backed OPC open,
eager/source-backed one-Part atomic save, and same-length CFB atomic overlay
save. A
one-sample debug correctness/counter smoke covers warm and cold-requested modes
(10 result records and five evidence records). Source OPC open makes 13 logical
reads totaling 1,008 bytes and materializes no Parts; eager open materializes
four Parts. Both OPC saves produce SHA-256
`f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009`;
CFB reports one changed span and SHA-256
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.
Cold-requested records contain nonzero process `read_bytes`, but do not prove
a reproducible cold cache. The debug, dirty-worktree, one-sample artifact makes
no latency, allocation, memory, throughput, warm/cold comparison, or
production-performance claim. See
[`0087`](changes/0087-filesystem-cache-state-evidence.md) and the
[compact counter summary](results/filesystem-smoke-0096-summary.json).

Change 0236 adds an opt-in `cold-verified` state without changing the
`cold-requested` default. It is Linux-only and admits only regular, non-empty,
page-aligned, read-write-open sources on an allowlisted block-backed
filesystem identified from the opened FD's numeric `statfs` magic (`0xef53`
ext2/3/4, `0x58465342` XFS, `0x9123683e` Btrfs, `0xf2f52010` F2FS, or
`0x2fc12fc1` ZFS). It records the canonical fincore basename, executable
SHA-256/version, stderr digests/lengths, method/fallback evidence, and a
strict external `fincore` JSON proof of zero resident, dirty, and writeback
bytes. The timed operation must produce a positive process `/proc/self/io`
`read_bytes` delta. Prepared query controls are excluded. Ineligible
host/proof conditions are explicit statuses and emit no timed result. This
proves page-cache and process-I/O conditions only; it makes no physical-media
or device-cache claim and has no captured performance result. On 32-bit Linux
the state is conservatively rejected as `ineligible_linux_non64_bit`. The
`filesystem_evidence` state records `cold_verified_status` and optional
`cold_verified_samples` alongside the existing `warm` and `cold-requested`
states; sample provenance contains aligned source SHA-256/size and fincore
method/fallback fields, while stderr contents and absolute paths are omitted.
See
[`0236`](changes/0236-cold-verified-filesystem-evidence.md).

Source-backed OPC payload retention is now optionally charged to a caller's
hierarchical `Budget`. The managed cache preserves pinned handles, reserves
active single-flight loads, evicts only unpinned clean entries, and reports
content-free budget diagnostics. Three opt-in harness selectors now cover the
exact/one-under managed Budget boundary and matched finite-control/managed
same-Part plus fixed-work disjoint-Part contention across `1/2x`, `1x`, and
`2x` capacities. They enforce exact gate, cache, pinning and Budget-release
counters and classify Amdahl results only where request count remains fixed.
The committed managed source-backed OPC change (`f8d417ac3`) charges exact
physical `InputBytes`, cumulative declared cold-load `Work`, retained
catalog/flight/payload `Objects`, and retained/in-flight payload `Memory` to the
caller-owned hierarchical `Budget`; compatibility opens remain on the finite
unmanaged `SourceCacheLimits` path. Focused correctness tests cover these
resource charges, retained-resource releases, pinning, eviction, single-flight,
cancellation, sibling competition, and contention invariants. The fixed-delay
harness is a coordination instrument, not production latency. Its controlled
release ABBA provides structural and distribution evidence only; no
managed-versus-control speedup is accepted. Allocation, peak-memory/RSS,
hardware-counter, copied/decompressed-byte, CPU-utilization, and
production-performance evidence remain absent. See
[`0086`](changes/0086-opc-source-cache-budget-management.md) and
[`0088`](changes/0088-opc-source-cache-contention-evidence.md).

The current managed direct `SourceBackedPackage` sequential sinks also charge
`Resource::OutputBytes` per write and commit only the exact bytes accepted by
the caller sink. Exact/no-op copies and changed overlays retain typed
zero-output refusals, partial/cancelled/source-changed `IncompleteOutput`, and
content-free refusal diagnostics. This accounting is correctness evidence
only: it does not extend to `OpcPackage` atomic saves, `to_bytes`, or the
unmanaged compatibility path, and makes no performance claim.

The five filesystem cases also have a repeated release capture: 30 fresh-child
samples in each of `warm` and `cold-requested` state on a CPU-pinned tmpfs
process (300 samples total). It records logical and process I/O counters,
materializations, changed spans, output hashes, and descriptive latency
distributions. `cold-requested` remains only an accepted advisory
`posix_fadvise(DONTNEED)` request; tmpfs process `read_bytes == 0` is a
process-I/O observation and does not establish physical cold-cache behavior or
a storage-device claim. No comparator, allocation, or peak-memory acceptance
follows from this run. See
[`0089`](changes/0089-filesystem-release-repeated-evidence.md).

Bounded forward-only one-sheet XLSX creation and RTF authoring exist in
production (`8245da20d` and `5918be8ec`). RTF's accepted ASCII batching result
is recorded in change 0097. XLSX change 0169 accepts the precise warm in-memory
one-sheet latency directions and descriptively records the matched whole-process
allocation-call reductions described above. Change 0170 additionally accepts
large p50/mean/p95/p99, medium p50/mean/p95, and tiny p50 improvements from
batched XML-safe UTF-8 runs; the other tiny statistics and medium p99 are
withheld. RSS, total-memory/peak-memory attribution, physical/cold I/O,
multi-sheet/richer authoring, and producer evidence remain open.

Bounded semantic validation reports are now implemented for DOCX, PPTX, RTF,
and XLS, alongside the existing CFB, OPC, and ODF validation reports. They
retain finite limits, typed failure attribution, and format-specific
preservation/security checks, but are correctness APIs rather than measured
performance cases. The opt-in ODF repair selector now exercises the deliberately
narrow typed non-destructive plan that removes one recognized local-header extra
from a first, stored `mimetype` member after source/provenance and full reopen
checks. Its retained-output-free sink and bounded write requests do not imply a
total memory bound because planning performs a bounded full-candidate preflight;
structural, encrypted, signed, macro, and semantic repairs remain unsupported.

Existing-document RTF logical-tail append now has two opt-in harness selectors
over tiny, medium, and large plain corpora. They verify bounded sequential
publication, complete reopen, patch/inverse and foreign-source refusal. The
16 KiB sink write window caps accepted bytes per write and retains zero output;
it does not bound the transaction's validated candidate snapshot. No release
latency, allocation, RSS, or speedup claim is made. See
[`0090`](changes/0090-rtf-logical-tail-append-evidence.md).

Change 0153 adds four matched RTF tail selectors: Commit versus
PublicationPlan for changed append and exact no-op. Their `elapsed_ns` is the
pre-staged publication-call interval around the respective public write call,
using the same fixed 16 KiB non-seek sink; separate planning, publication,
reopen, and lifecycle vectors report their scopes. `planning_ns` and
`publication_ns` have one entry per retained sample, while `reopen_ns` and
`lifecycle_ns` are one-element preflight-only vectors because the expensive
correctness gates run once outside the sample loop rather than repeating for
each sample. The results explicitly distinguish retained source, complete
candidate, and publication-window bytes.
Commit and PublicationPlan intentionally perform asymmetric validation and
publication work. Exact output/digest/semantic/no-op, durable
apply/inverse/stale/foreign, cancellation, sink failure/partial progress,
limits, and source-version gates remain untimed correctness checks. No
end-to-end, rich-format, allocation/RSS, physical-I/O, or ABBA latency claim is
made. See [`0153`](changes/0153-rtf-tail-publication-plan-evidence.md).

Change 0154 adds six matched ODT/ODS/ODP owned-rebuild and source-positional
`content.xml` publication selectors. A clean CPU-2 release ABBA used 20
warmups and 100 samples per record in strict A/B/B/A order. Both pair
directions accept prepared-publication p50 improvements of 96.35%-96.63%; p95,
p99, and mean agree, and maximum absolute same-implementation p50 drift is
1.441%. Exact content, family reopen, inventory, positional raw untouched
identity plus physical/central order, no-op, limits, cancellation, source
immutability, bounded sink, and logical `ReadAt` replay remain untimed gates.
This is an in-memory prepared-publication result only: no end-to-end,
allocation/RSS, physical-I/O, decompression, cold-cache, filesystem,
real-producer, or iWork claim is made. See
[`0154`](changes/0154-odf-content-cow-publication-evidence.md) and the
[summary](results/odf-content-cow-abba-0154-summary.json).

The source-backed XLS worksheet-visibility overlay landed in committed
production change `bac279116`. Committed change `0091` adds four opt-in eager
and source-backed scalar/batch selectors over one-owner and bounded 64-owner
visibility edits. They verify complete worksheet/catalog/opaque-stream
readback, exact overlay bytes, patch/inverse, source fingerprints/spans, and
cap/protection refusals. This is correctness/coverage evidence only: it makes
no release ABBA, speedup, allocation, RSS, peak-memory, or physical-I/O claim.
The source-backed path retains its complete candidate snapshot; its 64 KiB
publication sink bound limits writes, and retained output is only for digest
and reopen assertions, not a candidate-memory bound. See
[`0091`](changes/0091-xls-visibility-source-overlay-evidence.md).

Change 0095 replaces the complete source-backed `Workbook` handoff for both
existing-comment and worksheet-visibility publication with checked logical
range splices. One/256-comment plans now submit 109/27,904 replacement bytes
instead of 80,946, while one/64-visibility plans submit 1/64 instead of
18,166. A CPU-pinned 10-warm-up/100-sample balanced ABBA run kept every
source-backed p50 direction inside 1.5%; for each matched workload, the largest
absolute source-backed delta was below the largest absolute eager-control delta,
so no latency improvement is accepted. Allocation, RSS and physical-I/O
claims remain open; full semantic readback, exact fingerprints and every
preservation/refusal gate remain. See [change 0095](changes/0095-xls-semantic-splice-publication.md)
and its [compact result](results/xls-semantic-splice-abba-0107-summary.json).

The previously measured tranche includes six opt-in simulated-range cases,
two opt-in execution-scaling cases, one opt-in XLSX
commit/read attribution case, four opt-in opaque-heavy common OLE2 publication
stage/control cases, one opt-in source-backed OPC one-Part publication case,
one opt-in source-backed DOCX semantic publication case, one opt-in media-rich
PPTX semantic publication case, four opt-in matched same-slide/multi-slide
PPTX batch cases, two opt-in matched cross-slide ODP text-box cases, six opt-in media-rich ODT paragraph,
line-break, inline-run, hyperlink, insertion, and removal publication cases,
20 opt-in matched XLSX calculation-metadata/defined-name/page-break/page-margin/print-options/page-setup/sheet-protection/data-validation/auto-filter/conditional-formatting
publication cases, 16 opt-in DOCX/PPTX semantic
cases, 13 opt-in RTF semantic case names across four capability-bounded
variants (39 tiny / 70 tiny-plus-large rows), 24 shape-selected ODT/ODS/ODP
semantic cases, twelve fixed media-rich ODF cases, and 21 opt-in native
DOC/XLS/PPT semantic cases. It remains an
incomplete program and CRUD matrix.

- The XLSX row-start index is accepted for the narrow-range case: ABBA p50
  geometric mean **-80.499%**, mean geometric mean **-79.962%**; full scan
  mean **+0.03%**, first-cell mean **-1.31%**, heap allocations **+17**, and
  RSS **+0.25%**. Raw samples: [`before A`](results/abba-xlsx-range-before-a.json),
  [`after A`](results/abba-xlsx-range-after-a.json),
  [`before B`](results/abba-xlsx-range-before-b.json),
  [`after B`](results/abba-xlsx-range-after-b.json).
- Positional `SharedOleFile`, bounded CFB bulk reads, one-index positional ZIP,
  opaque ZIP `EntryId`, local `ParallelReadSession`, and the runtime-neutral
  `ExecutionContext`/OPC `OpenSession` are implemented. Default/legacy opens
  are serial; hidden global Rayon scheduling is removed. Current evidence is
  correctness and boundedness, not a new aggregate latency claim. Change 0094
  adds pinned ABBA evidence for exact CFB range reads: MiniFAT source bytes fall
  from 261,184 to 36 (many-small) and from 2,096,192 to 36 (wide-root), with
  stable read-stage p50/p95 reductions and only modest total p50 movement.
  FAT remains one 4 MiB request/call with no accepted tail claim; the result is
  substrate-only and does not adopt a DOC/XLS/PPT semantic speedup.
- Source-backed OPC now has source versions, finite weighted-LRU/single-flight
  cache diagnostics, and additive DOCX/XLSX/PPTX facades. EOCD terminal-probe
  samples show structural-open bytes down **73.6% to 98.5%** and payload overlap
  at zero. Latency is intentionally not compared because later EntryId/cache
  diagnostics changes confound the ABBA pair and some cells exceed 5% variance.
  See [`EOCD A`](results/abba-eocd-before-a.json),
  [`EOCD B`](results/abba-eocd-before-b.json), and
  [`source versus eager`](results/stage3-source-vs-eager-many-small.json).
- The low-level source-backed package can now consume one existing ordinary
  Part replacement without changing URI, content type, relationships or
  topology. It validates/materializes only the target and raw-copies every
  other ZIP member. On the fixed four-Part 16.78 MiB corpus, pooled p50 falls
  from 223.602 to 60.112 ms (-73.12%), semantic materializations fall from four
  to one, and output remains byte-identical to the eager baseline. Signed real
  changes and unsupported layouts refuse before output. See
  [`0037`](changes/0037-opc-source-backed-one-part-publication.md).
- The guarded DOCX facade now carries an exact raw main-document transaction
  through that publisher. On the fixed 17-Part media-rich corpus, pooled p50
  falls from 223.183 to 5.732 ms (-97.43%), instructions fall 74.91%, and
  semantic materializations fall 17 -> 1 with identical deterministic output.
  MCE rewrites, dependency transfers and signed real changes refuse before
  output; the unchanged eager DOCX guard is neutral. See
  [`0039`](changes/0039-docx-source-backed-semantic-publication.md).
- The guarded PPTX facade now carries an exact raw selected-slide transaction
  through the same consuming publisher. On the fixed 229-Part, 200-slide,
  eight-media corpus, pooled p50 falls from 296.590 to 8.545 ms (-97.12%,
  34.71x), instructions fall 67.91%, semantic materializations fall 229 -> 2,
  and output remains byte-identical. Its bounded atomic same-slide extension
  replaces eight unique shape texts in one scan/emission: matched p50/mean fall
  97.45%, allocation calls fall 39.80%, and materializations remain 229 -> 2.
  MCE-normalized slides, duplicate/overlapping batch selectors, topology
  changes and changed signed sources refuse before output; the unchanged eager
  PPTX guard is neutral. See
  [`0044`](changes/0044-pptx-source-backed-semantic-publication.md) and
  [`0063`](changes/0063-pptx-atomic-source-backed-shape-text-batch.md).
- The bounded multi-slide PPTX extension publishes eight selected slide Parts
  through one source-backed OPC preservation plan. Against the same 229-Part
  media-rich archive, pooled p50 falls from 331.362 to 13.997 ms (-95.78%,
  23.67x), allocation calls fall 32.54%, and semantic materializations fall
  229 -> 9 with byte-identical output. Duplicate, stale, foreign, signed,
  topology-changing and MCE-projected batches refuse before output. See
  [`0077`](changes/0077-pptx-source-backed-multi-slide-batch-publication.md).
- The guarded XLSX calculation-metadata editor now carries exact raw
  `xl/workbook.xml` transactions through the one-Part publisher. On the fixed
  12-Part, eight-media corpus, pooled p50 falls from 215.457 to 1.612 ms
  (-99.2519%, 133.67x), instructions fall 77.78%, and semantic
  materializations fall 12 -> 1 with byte-identical output. MCE projection,
  changed signed sources, stale/foreign closures and topology changes refuse
  before output. Cells, formulas, cached results and calculation-chain
  ownership are deliberately outside this capability. See
  [`0046`](changes/0046-xlsx-source-backed-calculation-metadata-publication.md).
- The guarded XLSX defined-name editor replaces or clears only the direct
  workbook catalog. On the same 12-Part media-rich archive, pooled p50 falls
  from 220.101 to 4.752 ms (-97.84%, 46.32x), instructions fall 78.45%, and
  semantic materializations fall 12 -> 1 with byte-identical output.
  Protection, MCE/unknown catalog children, invalid local scope, changed
  signatures and topology changes refuse. See
  [`0076`](changes/0076-xlsx-source-backed-defined-names-publication.md).
- The guarded XLSX page-break editor applies the same publisher to one selected
  normal worksheet after exact workbook-relationship closure checks. On that
  media-rich corpus, pooled p50 falls from 216.789 to 4.647 ms (-97.86%,
  46.65x), and semantic materializations fall 12 -> 2 with byte-identical
  output. MCE projection, relationship retargeting, changed signed sources,
  and topology changes refuse before output. See
  [`0061`](changes/0061-xlsx-source-backed-page-break-publication.md).
- The guarded XLSX page-margin editor binds the same exact selected-worksheet
  closure and exposes direct typed set/remove only. On the same media-rich
  archive, pooled p50 falls from 216.799 to 4.492 ms (-97.93%, 48.26x), and
  semantic materializations fall 12 -> 2 with byte-identical output.
  Chartsheets, MCE projection, relationship retargeting, changed signed
  sources and topology changes refuse before output. See
  [`0067`](changes/0067-xlsx-source-backed-page-margin-publication.md).
- The guarded XLSX print-options editor binds the same exact selected-worksheet
  closure and publishes only its direct five-flag element. On the fixed 16 MiB
  media corpus, p50 falls from 219.294 to 4.668 ms (-97.87%, 46.98x), while
  semantic materializations fall from 12 to 2 and output remains byte-identical
  across eager/source controls. See
  [`0070`](changes/0070-xlsx-source-backed-print-options-publication.md).
- The guarded XLSX page-setup editor additionally retains the selected
  worksheet's complete outbound relationship set and accepts only
  relationship-free settings. It refuses printer references rather than
  silently widening a one-Part edit to a printer-resource graph. The matched
  media-rich pair records 12 versus two semantic materializations and exact
  byte-identical output; see
  [`0073`](changes/0073-xlsx-source-backed-page-setup-publication.md).
- The guarded XLSX sheet-protection editor retains that complete workbook,
  worksheet and outbound-relationship closure while replacing the full typed
  core/Office 2010 protection state. On the same media-rich archive, formal
  p50 falls from 221.877 to 4.982 ms (-97.75%, 44.54x), instructions fall
  77.87%, and semantic materializations fall 12 -> 2 with byte-identical
  output. MCE-selected protection, stale/foreign or relationship-mutated
  closures, chartsheets and changed signed sources refuse before output. See
  [`0078`](changes/0078-xlsx-source-backed-sheet-protection-publication.md).
- The guarded XLSX data-validation editor binds the same complete worksheet
  closure and replaces typed core plus Office 2010 validation collections. Its
  media-rich p50 falls from 222.945 to 5.009 ms (-97.75%, 44.51x),
  instructions fall 73.43%, and materializations fall 12 -> 2 with
  byte-identical output. Allocation calls remain within policy and peak
  heap/RSS are flat; see
  [`0079`](changes/0079-xlsx-source-backed-data-validation-publication.md).
- The guarded XLSX auto-filter editor binds the workbook, selected worksheet,
  complete outbound worksheet relationships, and the styles Part plus DXF
  count when present. It replaces or clears the direct typed filter/sort state,
  while MCE-selected, protected, stale, foreign, relationship-mutated and
  changed signed sources refuse. On the media-rich control, p50 falls from
  219.615 to 4.946 ms (-97.75%), instructions fall 73.57%, and semantic
  materializations fall 12 -> 3 with byte-identical output; see
  [`0080`](changes/0080-xlsx-source-backed-auto-filter-publication.md).
- The guarded XLSX conditional-formatting editor now has selectable matched
  eager/source-backed publication evidence over the same 12-Part, eight-media
  corpus shape. Both paths replace the same complete three-owner typed core
  collection through the same worksheet rewriter and produce byte-identical
  output. The source-backed path materializes workbook, selected worksheet and
  styles (12 -> 3); exact patch/inverse, complete reopen, all unselected Part
  and media payloads, raw ZIP members, hashes, source reads and sink bounds are
  checked outside timing. No latency claim is made before balanced ABBA
  evidence is retained; see
  [`0082`](changes/0082-xlsx-conditional-formatting-performance-evidence.md).
- Native RTF middle-paragraph removal and first-to-final reordering now have
  independently selectable cases over the same deterministic plain corpus.
  Their intervals include edit/stage/commit, a constant-size diagnostics
  assertion, one shared snapshot-handle clone and bounded serialization.
  Full projection after reopen, volatile and durable forward/inverse replay,
  stale conflict, exact equal-position move no-op, bounded sink counters and
  output hashes are untimed gates. Changed CP-1252, LZFu, watermark and
  opaque/formatted inputs remain fail-closed. This adds coverage only and
  makes no latency or materialization claim; see
  [`0083`](changes/0083-rtf-paragraph-lifecycle-performance-evidence.md).
- Existing ODP text-box scalar and bounded-batch APIs now have matched
  selectable successful-path evidence over eight fixed-name owners distributed
  across a 12-slide, eight-media corpus. Both paths retain names and produce
  the same complete slide/full-text and rich-content projection. The batch
  raw-preserves the manifest; repeated scalar staging regenerates it, so the
  physically distinct outputs retain case-specific digests. Complete reopen,
  volatile/durable forward and inverse replay, stale refusal, auxiliary/media
  raw identity and real one-write sink counters are untimed gates. ODP exposes
  no source/materialization diagnostics, so none are invented. No latency,
  allocation, memory, or materialization claim is made before frozen CPU-pinned
  balanced ABBA evidence; see
  [`0084`](changes/0084-odp-cross-slide-text-box-batch-evidence.md).
- Consecutive packaged ODT plain-text replacements now share one mutable
  candidate, content publication, reopen and compact audit while retaining
  ordinary scalar durable operations. The large 100-edit/save p50 falls from
  906.439 to 15.615 ms (-98.28%, 58.05x), allocation calls fall 96.13%, and
  scalar one-edit guards remain neutral. See
  [`0045`](changes/0045-odt-coalesced-paragraph-publication.md).
- A matched release ABBA now covers the mixed model-content ODT workload as
  well. On the medium 80-operation shape, scalar A/B p50 values of 25.640 ms
  and 25.052 ms compare with batch values of 0.803 ms and 0.785 ms
  (31.9435x/31.9334x; 96.8695%/96.8685% reduction). On the large
  320-operation shape, scalar A/B p50 values of 2.759 s and 2.756 s compare
  with 21.276 ms and 20.998 ms (129.6876x/131.2449x;
  99.2289%/99.2381% reduction). This is only repeated-publication versus
  one-transaction evidence: source preparation, reopen/lifecycle/security/
  limits, I/O, serialization, allocation/RSS, and physical cold behavior are
  outside the timed claim. See
  [`0104`](changes/0104-odt-mixed-model-publication-evidence.md).
- Public ODT middle-paragraph lookup now validates the complete XML while
  retaining only the requested paragraph. Large-corpus p50 falls from 3.202 to
  1.647 ms (-48.56%), allocation calls fall 27.05%, peak heap falls 24.74%,
  and uninstrumented RSS falls 10.93%. The unchanged paragraph-list p50 moves
  +0.38%; a shared-mode parser prototype that regressed listing was removed.
  See [`0047`](changes/0047-odt-indexed-paragraph-selector.md).
- Public ODP middle-slide lookup now uses a compile-time-specialized full-EOF
  parser projection that retains semantic text and completed shapes only for
  the requested slide. Across 10,000 large samples, p50 falls from 1.019 to
  0.977 ms (-4.09%), mean falls 4.20%, p95 falls 5.18%, and whole-process
  allocation calls fall 3.86%. Tiny is neutral, medium improves 1.55% p50,
  and the unchanged list/save guards remain within thresholds. See
  [`0049`](changes/0049-odp-indexed-slide-selector.md).
- ODP transaction staging now reuses the complete slide projection already
  validated and retained by its immutable editing snapshot. Large exact-no-op
  edit/save p50 falls from 1.728 to 0.692 ms (-59.96%, 2.50x), while large
  changed edit/save falls 20.78%. Process allocation calls fall 20.13%, peak
  heap and uninstrumented RSS remain flat, and the complete package/security,
  raw-page-coverage, publication and independent readback boundaries remain.
  See [`0060`](changes/0060-odp-snapshot-slide-projection-reuse.md).
- Native DOC now indexes CLX pieces by physical FC with prefix maximum ends,
  so repeated PAPX/CHPX FKP range mapping skips non-overlapping pieces without
  assuming fast-save intervals are disjoint. Large public open p50 falls from
  790.727 to 348.679 us (-55.91%), changed one-paragraph edit/save falls
  31.08%, and the former 36.89% self-cycle range-scan frame falls to 4.17%.
  Peak heap and uninstrumented RSS remain flat. See
  [`0050`](changes/0050-doc-piece-table-physical-index.md).
- Native DOC PAPX reconstruction now retains one resolved paragraph-style
  baseline and reuses it when the next source run starts from the same style.
  Every run still applies and validates its own direct PAPX, piece modifier,
  and any direct style switch. Large public open p50 falls from 343.503 to
  304.199 us (-11.44%), mean falls 11.87%, and large changed edit/save p50
  falls 4.01%. Allocation calls fall 18.61%, while peak heap and
  uninstrumented RSS remain flat. See
  [`0051`](changes/0051-doc-adjacent-style-baseline-cache.md).
- Native DOC CHPX range queries now binary-search the first possible overlap
  and stop after the matching slice instead of filtering every character run
  for every paragraph. Large paragraph-list p50 falls from 454.100 to 358.414
  us (-21.07%), mean falls 20.93%, and p95 falls 20.00%. The attributed
  `extract_runs` self-cycle frame falls from 7.56% to 1.23%; allocation counts,
  peak heap, and uninstrumented RSS remain flat. See
  [`0053`](changes/0053-doc-chpx-range-index.md).
- Native DOC exact-source paragraph enumeration now resolves its ordered CLX
  piece and PAPX containment tables with predecessor binary searches instead
  of two fresh linear scans per paragraph terminator. The already-open
  512-paragraph snapshot list falls from 206.644 to 168.142 us p50 (-18.63%);
  one-edit/save falls from 888.602 to 817.424 us (-8.01%). Instructions fall
  26.13%, while allocation calls and peak heap remain flat. See
  [`0056`](changes/0056-doc-papx-containment-index.md).
- ODS durable-patch construction now retains its already owned immutable
  source and target package allocations in the semantic blob bundles and
  reuses their content addresses for operation preconditions. On the fixed
  16 MiB-media one-cell edit/save case, p50 falls from 326.694 to 297.958 ms
  (-8.80%), mean falls 9.07%, and p95 falls 13.85%. The former 33.58 MB
  `BlobBundle::insert` payload-copy site disappears; matched peak heap falls
  1.92%, while uninstrumented RSS is flat. See
  [`0054`](changes/0054-ods-shared-durable-patch-blobs.md).
- Eligible ODS row-local worksheet transactions now retain their exact checked
  source ranges through package publication. The prior flattened result could
  not be rediscovered as one conservative maximal diff, so the package layer
  rebuilt the archive and recompressed eight unchanged 2 MiB media members.
  On that fixed media-rich case, p50 falls from 287.766 to 74.365 ms (-74.16%),
  mean and p95 fall 74.17%/74.11%, instructions fall 69.04%, and matched peak
  heap/RSS remain flat. Foreign provenance refuses; signatures,
  encryption-sensitive/unsupported ZIP layouts and structural edits retain
  the established fallback. See
  [`0057`](changes/0057-ods-row-splice-raw-publication.md).
- The unified ODS worksheet handoff now moves its current `Vec` into the nested
  worksheet's exact `Arc<Vec<u8>>` owner, shares that owner with the private ODF
  package, and moves the validated target back out. The same media-rich
  edit/save falls from 76.440 to 60.140 ms p50 (-21.32%); peak heap and
  uninstrumented RSS fall 22.03%/20.57%. Exact failure rollback, patch/inverse,
  final reopen and security/layout fallbacks remain. See
  [`0068`](changes/0068-ods-shared-worksheet-archive-handoff.md).
- Exact unified ODS worksheet no-ops now stop at the nested worksheet handoff
  and construct their empty durable patch without reopening and diffing the
  same package again. Large exact-no-op p50 falls 23.26%, instructions fall
  10.54%, and peak heap remains flat. Changed commits retain every former
  audit and publication gate. See
  [`0058`](changes/0058-ods-exact-noop-handoff.md).
- Same-family fixed-width native XLS numeric commits now certify exact changed
  value ranges and carry the private BIFF cell-offset inventory forward while
  retaining the complete public Workbook validation/readback. On the large
  8,192-cell one-edit/save case, p50 falls 7.83%, mean falls 7.37%, peak heap
  falls 5.54%, and uninstrumented RSS is flat. See
  [`0059`](changes/0059-xls-fixed-numeric-inventory-carry.md).
- Large plain RTF parsing now derives an exact root-text block count during the
  existing structural preflight and performs one bounded lazy style-block
  reservation. Across 6,000 samples/state, open p50 improves 21.17%, mean
  21.00%, and p95 21.04%; one-edit/save improves 1.46% p50 and 1.75% mean.
  The block vector moves from 264 geometric allocations to 22 exact reserves
  over 22 parses, and peak heap falls 29.73%. Medium plain/CP-1252 centers move
  +0.49%/+2.84% p50 and are disclosed. See
  [`0055`](changes/0055-rtf-body-block-reservation.md).
- The positional XLSX source record reports p50 opens of 33.881 us (tiny),
  56.493 us (medium), and 139.897 us (dense); list-after-open has zero timed
  source reads. First-cell and narrow-range operations physically overlap only
  the selected worksheet member, with zero unselected worksheet read calls.
  These are overlap counts, not materialization counts. See
  [`xlsx-source-positional.json`](results/xlsx-source-positional.json).
- Targeted same-topology OPC publication now raw-copies unchanged ZIP members.
  Four-cell pooled ABBA p50 improves **58.24% to 96.41%** (geometric mean
  **84.98%**); few-large/incompressible falls from 216.299 to 61.206 ms. The
  same process profile cuts cycles **69.21%**, but retained source/provenance
  raises peak heap **37.18%** and one-shot RSS **22.26%**. See
  [`0008`](changes/0008-targeted-opc-preservation.md).
- The deterministic high-latency source records logical and physical request
  distributions, proves zero timed XLSX list requests, and proves zero
  unselected-sheet overlap. Explicit local-pool scaling reaches 4.52x p50 for
  six large OPC tasks and 5.93x for four large CFB streams at 12 visible CPUs;
  sub-kilobyte many-task cases are overhead dominated. See
  [`0009`](changes/0009-range-source-and-scaling.md).
- Generated 10,000-paragraph DOCX and 10,000-text-box PPTX corpora now cover
  semantic list/one/full-text/create/no-op/one-edit/1%-edit paths. Direct DOCX
  selection improves 4.72% p50; reusing PPTX's selected scene improves the
  1% edit/save case 9.37% p50/mean and cuts process allocation calls 11.67%.
  The PPTX one-edit guardrail is neutral. See
  [`0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md).
- Deterministic ODT/ODS/ODP corpora now cover public open, list, one-object,
  full-text, small-create, no-op and one-edit/save paths. Reusing the already
  validated ODS package during snapshot construction improves pooled p50 by
  **7.45% / 11.78%** for medium/large no-op edit-save and **3.57% / 2.06%**
  for one-cell edit-save. Full-process allocated bytes fall 1.46% in the
  medium no-op profile; peak heap is flat. See
  [`0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md).
- The DOCX one-percent transaction now coalesces canonical direct-body
  paragraph replacements into one bounded XML emission and candidate parse.
  Pooled large-corpus p50 falls from 487.542 to 24.418 ms (**-94.99%,
  19.97x**) and whole-process allocation calls fall **94.11%**, with flat peak
  heap and RSS. See
  [`0012`](changes/0012-docx-coalesced-paragraph-edits.md).
- Deterministic native RTF corpora now cover public open, lazy paragraph
  listing/selection, first full text, exact stream/no-op save and one paragraph
  edit/save. Retained text length removes the temporary fragment vector,
  ordinary ASCII emits in chunks, and text-only edits skip unused property
  scans. Large full-text p50 falls **27.08%** and large one-edit/save p50 falls
  **25.79%**; open moves +3.41%, peak heap is flat and RSS +0.32% (flat). See
  [`0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md).
- Existing ODT documents now hand their private immutable package allocation to
  transaction snapshots by shared handle instead of copying the archive.
  Medium/large no-op edit-save p50 falls **27.05% / 18.51%**; exactly two
  allocations and one package copy disappear per snapshot, while open and
  changed edit-save guardrails remain within 3% and peak heap/RSS stay flat.
  See [`0014`](changes/0014-odt-shared-snapshot-bytes.md).
- The same deterministic native DOC/XLS/PPT writer artifacts now have public
  open/list/one/full/no-op/one-edit semantic baselines. On the large shapes,
  one-edit/save p50 is 1.416 ms for DOC, 1.722 ms for XLS, and 0.357 ms for
  PPT; XLS open is 1.383 ms. See
  [`0015`](changes/0015-native-ole2-semantic-baseline.md) and the
  [`raw report`](results/ole2-semantic-baseline-a57506d23-2026-08-11.json).
- Reusing the already rendered/reopened CFB editor in native XLS commit removes
  one discarded BIFF owner parse and redundant package capture. Large one-cell
  edit/save p50 improves 7.72%, with peak heap and uninstrumented RSS flat.
  See [`0016`](changes/0016-xls-commit-editor-reuse.md).
- Native DOC paragraph commit now applies its ordinary WordDocument and table
  stream replacements to one isolated candidate and publishes the CFB once.
  Large one-edit/save p50 improves 10.52%; the final strict revision-owner and
  independent public-document reopens remain. See
  [`0017`](changes/0017-doc-batched-stream-publication.md).
- Eligible same-topology ODS worksheet commits now serialize only changed
  modeled rows and copy untouched XML spans exactly. Large/medium one-cell
  edit-save p50 improves 9.54% / 7.22%, allocation calls fall 5.85%, and peak
  heap falls 27.18%. Structural edits retain full-table fallback and changed
  opaque rows refuse publication. See
  [`0018`](changes/0018-ods-row-local-publication.md).
- Ordinary RTF body-text flushes now borrow the parser state and copy only the
  encoding plus block properties; the complete state is cloned only for
  insertion/deletion metadata. Large open p50 improves 20.09% and large
  one-edit/save p50 improves 11.54%. The former 8.53% exclusive clone frame is
  absent after the change; process allocations, peak heap and RSS are flat.
  See [`0019`](changes/0019-rtf-parser-state-specialization.md). An ODS
  target-package adoption candidate measured only -0.44% p50 with +0.30% p95
  and was fully reverted.
- RTF transport-byte accumulation now extends each all-ASCII source token in
  one batch instead of invoking the generic `SmallVec::extend` path once per
  character. Large open p50 improves 26.67% and large one-edit/save improves
  6.26%; instructions fall 18.40%, while allocation count, peak heap and RSS
  remain flat. The checked byte-valued non-ASCII and invalid-Unicode paths are
  unchanged. See [`0020`](changes/0020-rtf-ascii-transport-batching.md). An
  ODT final-document adoption candidate was reverted because its medium
  one-paragraph read guard regressed 6.33% mean and 17.64% p95.
- RTF ordinary-text lexing now finds the next structural or physical-line
  delimiter in one byte pass instead of decoding each UTF-8 scalar twice.
  Large open p50 improves 17.23% and one-edit/save improves 14.65%; plain,
  raw CP-1252 and LZFu opens improve at medium and large. Instructions fall
  21.27%, while peak heap and RSS remain flat. A prepared LZFu no-op segment
  moves +0.290 us/+6.41% p50 after parsing; the changed large LZFu open
  improves 19.39%. See
  [`0040`](changes/0040-rtf-byte-delimiter-scanning.md).
- Direct RTF decoded-body ownership was measured and fully reverted. The broad
  prototype improved large raw CP-1252 open 3.08% p50 and removed 20.15% of
  process allocation calls, but regressed ordinary plain large open 25.53%
  p50/22.45% mean. Owned-only refinements were compiler-layout sensitive at
  -1.41% and +1.02% p50. Only a malformed multibyte-tail exact-preservation
  regression remains. See
  [`0043`](changes/0043-rtf-decoded-body-ownership-rejected.md).
- Ordinary changed RTF body commits now retain a compact source range proven
  during the initial parser preflight instead of cloning and lexing the source
  again to rediscover it. Large one-edit/save p50 improves **10.72%**, mean
  **10.11%**, instructions fall **10.64%**, and the before-only locator subtree's
  588 allocations over 20 edits disappear. Candidate parse/readback and every
  conservative fallback/refusal remain. See
  [`0048`](changes/0048-rtf-retained-body-source-span.md).
- The RTF parser now retains exact visible body paragraph cardinality while it
  admits the already bounded root-body paragraph boundaries. A cold public
  count on the generated 10,000-paragraph story falls from 28.898 us to 20 ns
  p50 (-99.93%); full validation, transport variants, enumeration and save/edit
  paths remain. See
  [`0069`](changes/0069-rtf-retained-paragraph-count.md).
- ODT changed-operation compactness audits now share the already validated
  predecessor and candidate packages instead of allocating and copying three
  complete archives. The fixed 16 MiB-media paragraph edit/save improves
  30.44% p50, 31.36% mean and 32.41% p95; allocation calls fall 0.57% and peak
  heap/RSS remain flat. A dedicated exact no-op segment, which returns before
  the changed path, moves +39 ns p50 and is explicitly disclosed. See
  [`0041`](changes/0041-odt-compact-audit-package-sharing.md).
- ODT changed-commit envelope classification now shares the immutable snapshot
  package instead of allocating/copying one complete archive into a temporary
  owner. Across two balanced ABBA cycles, the fixed 16 MiB-media edit/save
  improves 11.40% p50, 11.95% mean and 12.19% p95; Heaptrack removes exactly
  two allocation calls per commit and peak heap/RSS remain flat. Archive,
  manifest, encryption and signature checks remain. See
  [`0042`](changes/0042-odt-envelope-package-sharing.md).
- ODT changed-result finalization now transfers the already validated
  document's immutable package bytes into a byte-only snapshot instead of
  copying 16.79 MB and parsing that copy. One independent final reopen remains.
  Across two balanced cycles, media-rich edit/save improves 22.74% p50,
  22.56% mean and 21.48% p95; the attributed allocation disappears, allocation
  calls fall 3.46%, and peak heap/RSS remain flat. See
  [`0052`](changes/0052-odt-final-result-byte-handoff.md).
- Targeted OPC changed-Part publication now shares the Part's existing
  immutable payload with the ZIP regeneration layer. Heaptrack removes one
  4.19 MiB allocation and peak heap falls 3.42%. Few-large compressible save
  improves 20.73% p50 and 18.49% mean; incompressible and many-small latency is
  within 3% p50/p95 except a +3.00% many-small p95, and uninstrumented RSS is
  flat (+0.22%). See
  [`0021`](changes/0021-opc-shared-regenerated-payload.md).
- The shared ZIP writer now moves each validated generated local span instead
  of cloning it after archive inspection. Heaptrack removes the remaining
  4.20 MiB local-span allocation and peak heap falls 3.20%. Few-large
  compressible/incompressible p50 improves 4.09%/2.70%; repeated small and
  exact-no-op guardrails remain within 5% on p50 and mean, and uninstrumented
  RSS is flat (-0.10%). See
  [`0022`](changes/0022-zip-generated-local-span-move.md).
- ODT full-text extraction now consumes parser-created block strings on its
  private path instead of cloning each string twice. Repeated large-corpus p50
  improves 3.25% and mean 4.81%; process allocation calls fall 15.48% and
  temporary allocations 45.52%, with peak heap and uninstrumented RSS flat.
  Structured queries remain near neutral. The unchanged open guard moves
  +3.94% p50/+4.17% mean and its +10.95% p99 trigger is disclosed. See
  [`0023`](changes/0023-odt-full-text-owned-blocks.md).
- Native PPT root slide-order capture now reuses the validated `OleFile`
  already owned by its package instead of rebuilding the CFB index. Four-cycle
  large-corpus ABBA improves p50 8.78% and mean 10.58%; allocation calls fall
  5.01%, temporary allocations 12.22%, and peak heap/RSS remain flat. All
  independent live-document, slide-order, review-history and public-reader
  checks remain. See
  [`0024`](changes/0024-ppt-slide-order-open-reuse.md).
- Eligible changed XLSX worksheets now hand their exact commit-validated store
  to the published snapshot under a 4,096-cell / 1 MiB XML bound. Medium commit
  plus first read improves 23.23% p50 and allocation calls fall 21.01%; the
  unrestricted dense-wide candidate was rejected at +8.99% peak heap. See
  [`0025`](changes/0025-xlsx-validated-store-handoff.md).
- Direct PPT text-edit setup now uses its complete editor preflight for live
  persisted-record resolution instead of reopening and recapturing the CFB.
  Large direct edit/save improves 14.12% p50 and 15.39% mean; allocation calls
  fall 3.53%, peak heap/RSS remain flat, and the minor-fault increase is
  disclosed. See
  [`0026`](changes/0026-ppt-text-edit-resolver-reuse.md).
- The PPT root transaction now accepts a private text publication only after
  exact working-source, selected-slide persist-ID, and non-document-record
  checks. Large root one-shape edit/save improves 18.59% p50 and 17.83% mean;
  allocation calls fall 6.54%, peak heap/RSS remain flat, and custom limits
  retain the original complete root reopen. See
  [`0062`](changes/0062-ppt-root-text-publication-adoption.md).
- Repeated public ODS cell lookup now builds a private bounded locator only on
  the 64th successful query. Large cell-sweep p50 improves 81.74% and full-cell
  text p50 improves 52.65%; the dense locator requests 3,216 bytes, while peak
  heap and RSS remain flat. See
  [`0027`](changes/0027-ods-adaptive-cell-locator.md).
- An XLS-only handoff of the first validated terminal CFB rendering was
  measured and fully reverted. Tiny changed-save p50 improved 7.55%, but large
  changed-save p50 improved only 0.39% and four repeated large exact-no-op
  cycles regressed 22.00% p50 / 16.69% mean. Peak heap stayed flat and
  allocation calls fell 0.33%; the regression remains the rejection gate. See
  [`0028`](changes/0028-xls-terminal-render-handoff-rejected.md).
- Direct XLSX action-plan flattening was measured and fully reverted. Formal
  medium 1% commit/save p50 improved only 1.54%/1.61%; dense-wide improved
  0.27%/0.68%, process allocation calls fell 0.0623%, and peak heap was flat.
  The writer's larger scan/emission/parse/readback boundary still dominates.
  See [`0030`](changes/0030-xlsx-action-plan-flattening-rejected.md).
- A new media-rich ODS case attributes unchanged package-member work with
  eight 2 MiB incompressible resources. Eligible compact `content.xml` edits
  now raw-copy other validated members, and exact physical comparison skips
  their six former semantic-diff inflations only while the manifest is exact.
  Media-rich one-cell edit/save improves 4.73% p50, 5.73% mean and 7.65% p95;
  peak heap falls 8.78%, while the existing medium no-media p50 improves 0.77%.
  Unsupported layouts and every unproved member retain established logical
  fallback. See
  [`0031`](changes/0031-ods-unchanged-media-preservation.md).
- Successful XLSX worksheet reads now skip the narrow x14ac collector when the
  raw XML contains no `dyDescent` token; rejected inputs rerun the collector so
  its historical error precedence remains exact. Medium commit and commit/save
  cells improve about 19-21% p50/mean, cold reads improve about 35%, dense-wide
  1% commit improves 19.62% p50, allocation calls fall 25.24%, and peak heap is
  flat. See [`0032`](changes/0032-xlsx-no-extension-scan.md).
- A deterministic common OLE2 publication case now edits one tiny MiniFAT
  stream while preserving four exact 4 MiB regular streams. A shared-payload
  writer prototype regressed the end-to-end p50 32.02%. Retaining the first
  fully validated render improved the heavy path 34.06%, but regressed large
  DOC open 21.64% and DOC one-edit/save 9.08%; both production prototypes were
  fully reverted. See
  [`0033`](changes/0033-ole-common-publication-handoffs-rejected.md).
- A media-rich ODP source-backed text-box case attributes the complete logical
  rebuild of eight unchanged 2 MiB members. Content-only rich-object edits now
  reuse the accepted common checked-splice/raw-copy path, while resource
  additions and unsupported/security-sensitive layouts retain the established
  rebuild. Pooled edit/save p50 improves 94.44% and p95 94.29%; allocation
  calls move +0.52%, and peak heap/RSS remain flat. See
  [`0034`](changes/0034-odp-unchanged-media-preservation.md).
- A fixed media-rich ODT case replaces one of 200 paragraphs while preserving
  eight exact 2 MiB incompressible resources. Content-only paragraph
  publication now uses the common checked-splice/raw-copy path, while XML over
  its 16 MiB optimization limit returns to the established ODT rebuild. Pooled
  edit/save p50 improves 95.58%, mean 95.63%, and p95 95.43%; allocation calls
  fall 6.71%, peak heap is flat, and RSS improves 0.59%. The ordinary ODT
  open/no-op/one-edit guards all improve. See
  [`0035`](changes/0035-odt-content-only-paragraph-publication.md).
- A matched case appends one line break to the middle paragraph through that
  same accepted content-only boundary instead of rebuilding and recompressing
  the eight unchanged 2 MiB resources. Pooled p50 falls from 217.532 to 3.985
  ms (-98.17%, 54.59x), mean falls 98.16%, instructions fall 78.34%, and
  allocation calls fall 6.90% with flat peak heap/RSS. Only `content.xml`
  changes at the raw ZIP-member level; patch replay, exact inverse, stale
  refusal, complete media readback and deterministic output remain checked.
  See [`0071`](changes/0071-odt-content-only-line-break-publication.md).
- A second matched case appends one unstyled inline run to the same middle
  paragraph through the accepted content-only boundary. Pooled p50 falls from
  225.431 to 3.635 ms (-98.39%, 62.01x), mean falls 98.38%, instructions fall
  78.48%, and allocation calls fall 7.00% with flat peak heap/RSS. Styled and
  unstyled regressions prove raw identity of every untouched member. Exact
  no-op dispatch also avoids the changed-path frame while all changed commits
  retain their prior validation body. See
  [`0072`](changes/0072-odt-content-only-run-publication.md).
- A third matched case appends one inert hyperlink through the same checked
  boundary. Pooled p50 falls from 221.443 to 3.988 ms (-98.20%, 55.52x), with
  exact URL/text reopen and raw preservation of every untouched member. See
  [`0074`](changes/0074-odt-content-only-hyperlink-publication.md).
- Two structural cases insert or remove the middle paragraph while changing
  only `content.xml`. Pooled p50 falls 98.20% (55.55x) for insertion and 98.27%
  (57.86x) for removal; instructions fall 82.14% in the combined profile,
  allocation calls fall 8.47%, and peak heap/RSS remain flat. See
  [`0075`](changes/0075-odt-structural-paragraph-publication.md).
- The opaque-heavy common OLE2 case now separates editor open, candidate
  `put_stream` publication, changed `finish` rendering, and the end-to-end
  control. Current p50 values are 1.382, 7.979, 5.473, and 26.086 ms; the
  isolated stages are explicitly non-additive. An inline exact recapture
  allocation-reuse prototype improved candidate publication 6.49% p50 but the
  end-to-end control only 2.61%, with p95 +0.54%, so it was fully reverted.
  See [`0036`](changes/0036-ole-common-stage-attribution.md).

See change records [`0005`](changes/0005-xlsx-row-start-index.md),
[`0006`](changes/0006-positional-containers-and-explicit-execution.md), and
[`0007`](changes/0007-source-backed-opc-and-facades.md), and
[`0094`](changes/0094-cfb-selective-read-evidence.md). Managed source-backed
OPC caches now charge exact physical `InputBytes`, cumulative declared
cold-load `Work`, retained catalog/flight/payload `Objects`, and
retained/in-flight payload `Memory` to a hierarchical `Budget`; compatibility
opens retain the finite unmanaged `SourceCacheLimits` path. Focused correctness
tests cover these resource charges, retained-resource releases, pin pressure, eviction,
single-flight, cancellation, sibling competition, and release accounting. The
committed release ABBA provides structural/distribution evidence only with no
accepted speedup. Allocation/peak-memory/RSS, hardware, copied/decompressed-
byte, CPU-utilization, and production-latency evidence remain pending. The
release filesystem evidence is likewise descriptive tmpfs data, not physical
cold-cache proof.

## XLSX provenance and RTF streaming ABBA (2026-08-14)

A CPU-2 release `before-A / after-A / after-B / before-B` run used 10 warm-ups
and 100 samples for six matched XLSX scalar-cell pairs and three RTF streaming
shapes. XLSX source-backed p50 geomean improves **21.66%/22.65%** and p95
**21.38%/22.70%** after eliminating a redundant publication-time semantic
worksheet reload. Physical read/materialization counters stay unchanged, so
this is not an I/O claim. RTF streaming p50 geomean improves
**76.41%/76.47%** and p95 **75.23%/75.76%** after batching escape-free ASCII
into at most 32-byte sink requests; the large case drops from 7,208,970 to
1,441,802 writes. Exact output hashes match every leg.

The medium eager XLSX exact-256 after-A control outlier (+30.59% p50,
+105.28% p95) moved opposite the paired source improvement and normalized in
after-B (+1.63%/+4.25%); no eager-path claim is accepted. Allocation, peak
heap/RSS, physical cold I/O and compression-byte conclusions remain pending.
See [change 0096](changes/0096-xlsx-source-provenance-publication.md),
[change 0097](changes/0097-rtf-bounded-ascii-streaming.md), and the
[compact summary](results/xlsx-rtf-abba-0108-summary.json).

Consolidated changed-crate tests, formatter checks, warning-denied production
Clippy and rustdoc gates passed. The current ODS all-target Clippy gate retains
the unrelated pre-existing test-only findings recorded in change 0027. The ODT
tranche compiled the ODF fuzz target offline; the PPT and ODS tranches have no
dedicated fuzz target in the current tree. A workspace all-target/all-feature
gate was not rerun because iWork was explicitly excluded while its crates are
changing independently.

## Immutable-owned CFB atomic save and rejected reuse experiments (changes 0175-0176)

The opt-in owned CFB filesystem selector raises the matrix to 320 names while
leaving the default 36 cases / 198 records unchanged. For its fixed
16,913,408-byte corpus, sealed ownership removes exactly 33,826,816 logical
source bytes, 34 one-MiB fingerprint reads, and two source/target digest pairs
per atomic save. Generic sources retain both complete fences; owned emission
still hashes every source and target byte and preserves flush/fsync/rename/
parent-sync durability. Clean CPU-2 A/B/B/A totals are lower in both warm and
advisory-cold paired directions, but 11.29%-14.16% control drift exceeds the
5% gate, so latency is descriptive only. See
[`0175`](changes/0175-cfb-owned-atomic-save.md).

Authenticated ODS `content.xml` reuse and XLSX conditional-formatting parsed
readback reuse were both fully reverted. ODS regressed source-backed p50 by
1.63%-2.83% in both directions; XLSX moved -4.81%/+1.99% across paired
directions. Exact output hashes and correctness gates passed, but neither
experiment met the usefulness/repeatability gate. See
[`0176`](changes/0176-rejected-odf-xlsx-reuse.md).

## ODS source-backed existing-cell release evidence (change 0177)

The four existing ODS selectors now retain aligned lifecycle and phase vectors
plus a separately untimed logical `ReadAt` replay. Clean CPU-2 A/B/B/A uses one
release binary, 20 warmups and 500 samples per workload/leg over the fixed
16.01 MiB media-rich corpus. For one existing cell, source-backed complete-
lifecycle p50 is 75.03%/74.27% lower in the two paired directions; mean, p95
and p99 also improve, and eager/source drift passes the predeclared
5%/5%/10%/15% thresholds. That one-cell latency result is accepted.

The 21-cell deterministic 1% path is correctness/phase evidence only. Its p50
is 73.59%/73.16% lower, but candidate mean drift is 5.86%, p95 drift reaches
14.06%, and p99 drift reaches 18.41%. No 1% latency claim is accepted. The
617-call/16,801,025-byte replay is logical `ReadAt` evidence, not physical I/O
or decompression. Allocation/RSS, cache-temperature, real-producer, durable
ZIP patch, atomic-save and broader ODS CRUD claims remain open. See
[`0177`](changes/0177-ods-source-cell-release-evidence.md).

## Change 0345: OPC source-backed reader ingress

[Change 0345](changes/0345-opc-source-backed-reader-ingress.md) records structural evidence only (performance_claim: none): the public litchi-opc reader input was consumed once, an exact maximum-limit overrun returned the typed error with actual = maximum + 1 asserted, open performed zero cold ordinary-payload loads, and one selected part produced one cold successful load. Relative to compressed-plus-all-decompressed eager retention, this path retains one compressed buffer plus indexed metadata and deferred selected payloads. ReadLimits and try_reserve_exact bound logical input/local admission work, not total RSS or aggregate concurrent opens. Validation was 4/4 focused tests, including reader_ingress_retries_one_interrupted_read and reader_ingress_rejects_invalid_read_count_without_panicking, plus four owner-library checks, run with one Cargo process/job and a dedicated on-disk target. No RSS, allocation, or before/after latency claim is made; callers needing tighter host memory must provide a lower max_input_bytes, serialize opens, and account aggregate process memory externally. Arbitrary blocking Read cancellation remains limited, and the change adds no facade API or iWork path.

## Change 0346: FileSource lock attribution

Change 0346 retains a control-only six-selector XLS source attribution smoke
at HEAD `3a2926f8a` (three warmups and 50 samples, CPU 2, one serialized
worker) plus a standalone 40-block mutex/fingerprint probe. The control reports
preserve exact logical reads, source-version counts, stable semantics and the
fixed `ConditionalFormattingSamples.xls` identity. The probe reports 155.47
ns/call mean for `std` and 154.03 ns/call for `parking_lot`; the modeled
whole-operation gain is only 0.36-0.40%. No production candidate was applied,
no XLS/CFB fence changed, and `performance_claim: none`. The evidence is
diagnostic only; the next XLS line must target a different design. See
[`0346`](changes/0346-file-source-lock-candidate-rejected.md) and
[`results/change-0346`](results/change-0346/).

## Change 0347: XLSX cell-values harness calculation-closure repair

Change 0347 is harness-only evidence (`performance_claim: none`). The old
clean control failed on the intentional `calcPr` invalidation as a stale raw
unselected `xl/workbook.xml` oracle; the repaired oracle passes the direct
medium one-edit eager/source smoke with three warmups and 30 samples, plus a
24-row ABBA v1 smoke with zero warmups and one sample, recording zero failure
rows and complete rows. Timing gates and claim authorization remain false.
The numeric medium and dense-sparse identities are pinned; formula/date cases
are excluded. Source planning/commit dominate the retained source phases, so a
shared publication-copy design remains unproven and deferred. See
[`0347`](changes/0347-xlsx-cell-values-harness-calculation-closure.md) and
[`results/change-0347`](results/change-0347/).

### Change 0348: Stored ZIP borrowed validation

Stored ZIP entries now undergo complete local and central metadata validation
before immutable-slice borrowing. Signed and unsigned 32-bit descriptor
CRC/size forms, ZIP64 local-extra provenance, encryption/overlap/duplicate
safety, and strict nonempty zero-CRC refusal are covered; ZIP64 EOCD, Deflate,
and generic positional sources retain owned or streaming fallback. Pointer
identity is preserved without a cache or materialization charge, and
concurrency is unchanged. The serialized evidence used `CARGO_BUILD_JOBS=1`,
`test-threads=1`, and an 8 GiB process ceiling:
`focused borrowed 10/10; full soapberry-zip lib 280/280`. Downstream
`litchi-opc borrowed 12/12` passed; this is a filtered result, not the full
`litchi-opc` suite. `cargo test --locked -p litchi-opc --lib borrowed
-- --test-threads=1` and `cargo fmt --package soapberry-zip -- --check` passed
after
formatting. `performance_claim: none`: no latency, RSS, allocation, or bytes-copied claim;
stored OOXML corpus representativeness remains weak. The low-level raw
`get_entry_borrowed` accessor remains unverified and requires a verifier. See
[Change 0348](changes/0348-stored-zip-borrow-validation.md).

## Change 0355: PPTX source-probe fallback admission

Change 0355 is correctness and ownership evidence only
(`performance_claim: none`). The private PPTX bytes probe now returns typed
`OpcError` outcomes and terminal `OtherOoxml`/`DisabledOtherOoxml` classifier
outcomes. Only genuine non-ZIP, short-input, or missing `[Content_Types].xml`
inputs enter the compatibility fallback and reclaim the original `Vec`
allocation; hard ZIP, OPC, and classifier errors do not eagerly retry PPTX or
ODP. The public `DetectedFormat` and eager path are unchanged, as is the
ordinary proven ODP native-owner handoff/reparse.

Path `FileSource` captures `SourceVersion`, preflights the caller's exact
`max_input_bytes`, and uses a same-source bounded `Bytes` fallback instead of
pathname re-open or unbounded `fs::read`. Semantic conversion failure
rechecks freshness first, and `Presentation` consumes retained bytes under
the exact input and part limits. Input/part-limit, malformed-ZIP typed-error,
missing-manifest allocation, wrong-family/polyglot precedence, extensionless
bounded-path, reserved-namespace, and freshness/cancellation regressions
remain covered.

With an 8 GiB virtual-memory ceiling, one Cargo job, disabled incremental and
debug compilation, and one disk target, `cargo check -p litchi
--no-default-features --features pptx`, the combined `pptx,odp` library test
(`48/48`), and `cargo fmt --package litchi` passed. The final target was 674
MiB with approximately 15 GiB host availability and saturated swap; no
additional pressure or OOM was observed. No speed, RSS, or OOM-prevention
claim follows. DOCX, non-Unix, ODT, ODP prepared-package, public eager-smart,
and selected-part materialization seams remain open.
## Change 0383 update

Change 0383 adds bounded source-backed cross-slide copying of zero or more
direct ordinary chart graphic frames alongside the existing 0382 picture
subset. Only self-contained internal canonical `/ppt/charts/` relationship-free
leaf parts with a dialect-correct chart content type/root are admitted; ChartEx,
workbooks, `externalData`, style/color, chart drawing/userShapes, outbound
relationships, and broader chart graphs are refused. Distinct chart parts are
copied once while separate slide bindings remain, with deterministic canonical
part/relationship allocation and exact namespace-resolved `r:id` rewrites.
Malformed/ambiguous/nested hosts, stray chart namespace content, MCE/DTD/PI,
unresolved namespaces, stale/foreign/signed/limited inputs, cancellation, and
unsupported collisions fail closed according to the source-backed contract.

Correctness evidence is focused `52/52`, isolated typed cancellation `1/1`,
default-feature library `531` passed plus one named filtered test,
all-features primary library `533` passed plus one named filtered test,
all-features integrations green with three existing exact exclusions, and
doctests `6` passed/`2` ignored. Clippy is green with inherited
`nonminimal_bool`, `clone_on_copy`, and `needless_lifetimes` allowances. The
boundary audit is `64` packages/`240` internal dependencies/`14` debt edges.
No performance measurement or performance claim is made.

The run used one Cargo process, one build job, a 6 GiB virtual-memory cap, and
a 10 GiB `MemAvailable` gate. This is resource-capped/OOM-mitigating execution
policy, not proof of OOM prevention. See [Change 0383](changes/0383-pptx-source-backed-cross-slide-chart-leaf.md).

`performance_claim: none`; `claim_authorized: false`.
