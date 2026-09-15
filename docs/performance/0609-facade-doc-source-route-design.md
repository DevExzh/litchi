# 0609: the facade's `.doc` slurp is the cheaper route — routing it to the source-backed DOC reader would cost 2-5× the cycles and admit four artifacts the eager reader refuses

Status: **retained, design only. Nothing under `crates/` changed.**
`performance_claim: none` — the admission census, value comparison, read counts,
allocator counts, instruction counts, native cycles and paired medians below are
reported as evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is the frozen design record item **CORE-1** of change
[0587](0587-remaining-opportunity-survey.md) (rank 30) asked for. 0587 ranked
CORE-1 with size **unknown** and named its own falsification test: *"falsified if
on the eight admitted fixtures the source-backed open is not lower in both
instructions and peak RSS, or widening admission requires weakening a validation
the eager path performs."* Both halves are now measured, and both fail. The
design is recorded in full so that the next attempt starts from the numbers
rather than from the hypothesis, and **no routing is implemented**, because none
exists that is value-identical.

## What was changed

Nothing. `git diff --name-only 8fe9efa55 -- crates/` is empty on this branch. The
change adds this record and the evidence packet under
`results/change-0609/`, whose `probe/` directory carries the retained scratch
driver that produced every number here.

## The two routes as they stand at `8fe9efa55`

**Route E (eager, what `litchi::Document::open(path)` does today for `.doc`).**
`detect_document_source_path_with_limits`
(`crates/litchi/src/detection_smart/detected.rs:2389`) opens a `FileSource`,
probes for ODT and for an OOXML package, and when neither owns the bytes calls
`read_path_source_bytes` (`:1198`) — `version()`, `len()`, `version()`, one
`read_exact_at` of the whole file, `version()` — under the 2 GiB
`UNIFIED_DOCUMENT_FALLBACK_MAX_INPUT_BYTES` ceiling (`:2148`). The `Vec<u8>` is
handed to `Document::from_bytes_with_limits`, classified by
`check_office_signatures`, wrapped once in `OleFile<Cursor<Vec<u8>>>`
(`:42`, `:1707`, `:1823`) and parsed eagerly by `doc::Package::from_ole_file`
plus `package.document()` plus `ole_file().get_metadata()`
(`crates/litchi/src/document/doc.rs:928-947`). The parsed `doc::Document` and
its `Metadata` are retained; the package, the `Cursor` and the slurped `Vec` are
dropped at the end of that match arm.

**Route S (source-backed, `litchi_doc::body_text::source::SourceSnapshot`).**
`SourceSnapshot::open(Arc<dyn ReadAt>)`
(`crates/litchi-doc/src/body_text/source.rs:490`) validates the CFB envelope and
the FIB, parses the CLX piece table, and captures three complete-artifact
identity fingerprints (`:526`, `:591`, `:599`) around two `SharedOleFile` index
opens (`:513`, `:532`). Change [0105](changes/0105-doc-source-backed-paragraph-splice.md)
designed it as a *narrow editor* for one ordinary Unicode main-story paragraph,
not as a reader: "the public owner is intentionally separate from the existing
owned body editor", and "it is not broad DOC CRUD coverage". Its entire read
surface is `paragraph(Position) -> Paragraph { text() }`, `len()`,
`fingerprint()`, `source_version()` and `limits()`.

The facade never references route S. That is the gap CORE-1 proposed to close.

## Frozen design

### The two candidate routes

**R-A, source-first with typed fallback.** `Document::open` builds the
`FileSource`, calls `SourceSnapshot::open`, and on `Ok` retains the snapshot as
a new `DocumentImpl::DocSource` variant; on any `Error::Refused`, `Error::Ole`,
`Error::InvalidData` or `Error::Limit` it discards the snapshot and continues
into route E on the same `Arc<dyn ReadAt>`. `Error::Io` and the overlay's
`SourceChanged` propagate rather than falling back, because they are not
admission decisions.

**R-B, capability dispatch.** `Document::open` always takes route E, and a new
selected-paragraph entry point (`Document::open_paragraph_source(path)`, or a
`paragraph_text` fast path) takes route S. The facade's general `Document` is
never source-backed for `.doc`; the source-backed owner is reachable only where
its capability matches.

### Which refusals differ between the two readers — the whole corpus

Measured over all 57 `.doc` fixtures under `test-data/`
(`results/change-0609/measurements/census.tsv`):

| | route S admits | route S refuses | total |
| --- | ---: | ---: | ---: |
| **route E admits** | **4** | 38 | 42 |
| **route E refuses** | **4** | 11 | 15 |
| total | 8 | 49 | 57 |

Route S's 49 refusals: `Drawing` 34, `Field` 4, `AmbiguousTopology` 4,
`Encrypted` 3, `Revision` 2, `Macro` 1, and one `Ole` envelope error. Route E's
15: `CorruptedFile` 7, `InvalidFormat` 7, `NotOfficeFile` 1.

The decisive cell is the bottom-left. **Four artifacts are admitted by route S
and refused by route E**, every one of them with the same typed refusal:

| fixture | bytes | route E | route S |
| --- | ---: | --- | --- |
| `ole/doc/footnote.doc` | 9,728 | `CorruptedFile("invalid stylesheet: style names and aliases must be unique")` | **`Ok`** |
| `ole/doc/lists-margins.doc` | 10,752 | same | **`Ok`** |
| `ole/doc/duplicate-style-names.doc` | 64,512 | same | **`Ok`** |
| `ole/doc/picture.doc` | 1,448,448 | same | **`Ok`** |

This is structural, not a corpus accident. Route E validates the stylesheet
during `package.document()`; route S never reads the stylesheet at all, because
a same-width paragraph splice does not need it. No ordering of the two readers
can make route S produce a refusal it has no reason to compute. Under **R-A** all
four would open successfully through `Document::open`, so the facade would
**admit four artifacts its own reader calls corrupt** — the exact outcome
ADR 0006 forbids ("readers preserve real-world quirks" is about preservation, not
about skipping a validation another reader performs) and the outcome
`docs/GOAL.md` names when it says never to trade a typed refusal for a partial
result. R-A is therefore **rejected on admission alone**, before any performance
argument.

The error *identity* also differs on eleven of the fixtures both refuse —
route E says `InvalidFormat("Word 6.0 documents (nFib 0x0065) are not supported")`
where route S says `Refused::AmbiguousTopology`, and
`InvalidFormat("DOC password required")` where route S says `Refused::Encrypted`.
Under R-A the fallback makes route E's error the one the caller sees, so error
identity is preserved there; it is only the four admissions that cannot be
reconciled.

### What the source-backed reader can answer through the facade

`Document`'s DOC arm serves `text()`, `paragraph_count()`, `paragraphs()`,
`paragraph_text(i)`, `tables()`, `styles()`, `metadata()` and the markdown list
resolver. `SourceSnapshot` serves one of them, partially. Measured
(`measurements/identity.tsv`):

| fixture | route E `text()` | route E `paragraph_count()` | route E `paragraph_text(0)` | route S `paragraph(0)` | route S positions walked before refusal |
| --- | ---: | ---: | ---: | --- | ---: |
| `noheadfoot-litchi.doc` | 134 B | 10 | 31 B, digest `3eba…1f52` | `Ok`, 31 B, digest `3eba…1f52` | **1** (31 B) |
| `documentProperties.doc` | 22 B | 1 | 21 B, digest `bc22…d87b` | `Ok`, 21 B, digest `bc22…d87b` | 1 (21 B) |
| `endingnote.doc` | 40 B | 3 | 18 B | `Refused::StructuralContent` | 0 |
| `table-merged-cells.doc` | 85 B | 85 | 0 B | `Refused::StructuralContent` | 0 |
| `duplicate-style-names.doc` | (refused) | (refused) | (refused) | `Ok`, 20 B | 15 (274 B) |
| `lists-margins.doc` | (refused) | (refused) | (refused) | `Ok`, 41 B | 4 (171 B) |
| `footnote.doc` | (refused) | (refused) | (refused) | `Refused::StructuralContent` | 0 |
| `picture.doc` | (refused) | (refused) | (refused) | `Refused::StructuralContent` | 0 |

Read down the table: of the 8 fixtures route S admits, only 4 yield a single
paragraph; of those 4, only **2** are also admitted by route E; and on those 2
the paragraph text is **byte-identical** to the facade's, which is the one
genuine value-identity result in this record. It does not extend to `text()`:
walking every position a `text()` built on route S could reach yields 31 of the
134 bytes on `noheadfoot-litchi.doc` — the walk stops at position 1 — and 1
paragraph against the facade's 10. `paragraph_count()` has no route-S primitive
at all, and neither do tables, styles or metadata.

So **R-B's capability is one paragraph on two of 57 fixtures (3.5%)**, and a
general `Document` built on route S would have to open route E as well for every
other query, paying both costs.

### `SourceChanged` across the fallback

Today route E pins bytes once. `read_path_source_bytes` observes `version()`
before the read, re-observes after, and maps a mismatch to
`litchi_core::Error::SourceChanged` through `docx_path_core_error_to_opc` and
`Document::map_source_opc_error` (`doc.rs:236`). After that, every query is
served from the parsed model: **a `.doc` `Document` can return `SourceChanged`
at open and never afterwards.**

Under R-A both halves change. A snapshot retained past `open` re-checks source
identity on every `paragraph` call (`source.rs:746`, `ensure_current`), so
`paragraph_text` would gain a `SourceChanged` outcome it does not have today —
a new failure mode on an unchanged public signature. And the fallback itself
opens a window: the snapshot open reads the artifact six times (below) and then
route E reads it a seventh, so a file rewritten between the refusal and the
fallback would be parsed from the new bytes under the old `FileSource`, or would
raise `SourceChanged` where today the single pinned read either succeeded or
failed atomically. Closing that honestly means re-fencing the fallback against
the version captured before the snapshot attempt, which relocates the typed
`SourceChanged` boundary. Under **R-B** nothing moves: the general `Document`
keeps route E's fences and the new entry point owns route S's.

### What the facade's public types expose (ADR 0005)

ADR 0005 requires that "paths, moved byte owners, mmap-like owners, remote range
sources, and borrowed byte scopes adapt to [`ReadAt`] without exposing source
generics on a document". Both routes satisfy this: `SourceSnapshot` already
erases its source behind `Arc<dyn ReadAt>`, and a `DocumentImpl::DocSource`
variant would be crate-internal exactly as `DocumentImpl::DocxSource` and
`OdtSourcePathCandidate` are. No design difficulty exists here, and none of the
blockers below is an ADR 0005 problem.

### Admission gates for any future attempt

A source-backed facade DOC route may land only when **all** of these hold, each
checked with the probe retained in this packet:

1. Route S refuses every artifact route E refuses. The stylesheet validation is
   the first counterexample; the census is the test.
2. Route S's read surface covers `text()` and `paragraph_count()` with
   byte-identical output on every fixture both admit, not just
   `paragraph_text(0)` on two.
3. Route S's native cycles per open are below route E's on the admitted
   population, on both the small and the large end.
4. The fallback probe's cost on the refused population is below the noise floor,
   or the route does not probe (R-B).
5. No `SourceChanged` outcome changes on any existing change-under-read test.

Gates 1, 3 and 4 fail today by the margins measured below; gate 2 fails
structurally.

## Measured

Base `8fe9efa55`, one tree (no before/after legs: the comparison is between two
routes over the same source). AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws,
rustc 1.95.0, `--release`, every measured process pinned with `taskset -c 29`,
eight measurement agents active on the host throughout. Probe:
`results/change-0609/probe/`, built against this worktree with
`features = ["doc", "docx"]`, sha256
`8d956d66b72e88536cddc5406654cf220b91622fd56b6862923ca82f36f539aa`.

### Route S reads the complete artifact six times per open (**measured**)

`identity_fingerprint` (`source.rs:2089`) is
`plan_same_length_stream_splices(Vec::new(), …)`, which reaches
`finish_overlay_plan_with_owner` (`crates/litchi-cfb/src/overlay.rs:1022`). That
function runs `fingerprints` once, reopens the composed CFB, and — because a
generic `Arc<dyn ReadAt>` is not `source_is_owned_immutable` — runs `fingerprints`
a second time as the closing half of the read-twice-compare bracket 0589's
inventory classifies as P2. So **each identity pass is two complete artifact
reads**, and `SourceSnapshot::open` takes three of them: **six complete reads**.
Change 0589 halved the *hashers* driven over those reads (twelve SHA-256 passes
to six); it did not remove a read.

A counting `ReadAt` adapter confirms the arithmetic
(`measurements/readat.txt`):

| fixture | bytes | `read_at` calls | bytes read | × file | `version()` | `len()` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 9,728 | 30 | 72,669 | 7.47 | 116 | 35 |
| `noheadfoot-litchi.doc` | 10,240 | 30 | 75,741 | 7.40 | 116 | 35 |
| `lists-margins.doc` | 10,752 | 30 | 78,813 | 7.33 | 116 | 35 |
| `table-merged-cells.doc` | 17,408 | 30 | 118,749 | 6.82 | 116 | 35 |
| `duplicate-style-names.doc` | 64,512 | 31 | 401,149 | 6.22 | 116 | 35 |
| `picture.doc` | 1,448,448 | 46 | 8,761,309 | 6.05 | 140 | 35 |
| refused — `saved-by-table.doc` | 65,024 | 14 | 137,728 | 2.12 | 46 | 15 |
| refused — `FloatingPictures.doc` | 335,360 | 29 | 689,152 | 2.05 | 56 | 15 |
| refused — `ca.kwsymphony…doc` | 1,619,457 | 18 | 3,285,078 | 2.03 | 52 | 15 |

Six passes over `picture.doc` is 8,690,688 bytes of the 8,761,309 measured
(99.2%); the remainder is the CFB index, FIB and CLX. A **refused** open still
costs two complete reads, because the opening identity pass runs before
`finish_open` reaches the refusal.

### Route E reads the artifact once (**measured**)

`strace` isolation pairs, 1 against 11 operations, difference divided by 10
(`measurements/strace-summary.txt`):

| fixture | route | `openat` | `pread64` | bytes read | `statx` |
| --- | --- | ---: | ---: | ---: | ---: |
| `documentProperties.doc` (9,728 B) | E | 1 | **2** | 9,732 | **11** |
| | S | 1 | 30 | 72,669 | 153 |
| `noheadfoot-litchi.doc` (10,240 B) | E | 1 | **2** | 10,244 | **11** |
| | S | 1 | 30 | 75,741 | 153 |
| `table-merged-cells.doc` (17,408 B) | E | 1 | **2** | 17,412 | **11** |
| | S | 1 | 30 | 118,749 | 153 |
| `FloatingPictures.doc` (335,360 B) | E | 1 | **2** | 335,364 | **11** |
| `ca.kwsymphony…doc` (1,619,457 B) | E | 1 | **2** | 1,619,461 | **11** |

Route E's two reads are the 4-byte signature probe and one `read_exact_at` of
the file. On every fixture measured, route S issues **15× the reads, about 7×
the bytes and 14× the `statx`**. This is the opposite of what a source-backed
route is normally adopted for, and it is the reason the rest of the numbers go
the way they do.

### Allocator peak and retained bytes (**measured**)

A counting global allocator, armed for the duration of one operation
(`measurements/alloc.txt`). "Retained" is live bytes with the handle still alive.

| fixture | bytes | E peak | E retained | S peak | S retained |
| --- | ---: | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 9,728 | 47,245 | 22,075 | **24,605** | **4,396** |
| `noheadfoot-litchi.doc` | 10,240 | 53,081 | 25,636 | **25,129** | **4,400** |
| `endingnote.doc` | 9,728 | 50,949 | 23,992 | **24,605** | **4,396** |
| `table-merged-cells.doc` | 17,408 | 198,439 | 159,178 | **32,213** | **4,372** |
| `lists-margins.doc` | 10,752 | (refused) | — | 25,653 | 4,404 |
| `duplicate-style-names.doc` | 64,512 | (refused) | — | 78,929 | 4,392 |
| `picture.doc` | 1,448,448 | (refused) | — | 1,097,209 | 15,648 |
| `saved-by-table.doc` | 65,024 | 1,126,154 | 1,028,876 | (refused) | — |
| `FloatingPictures.doc` | 335,360 | 1,093,359 | 711,859 | (refused) | — |
| `ca.kwsymphony…doc` | 1,619,457 | 3,640,417 | 1,090,299 | (refused) | — |

**This is the one axis on which route S wins**, and it wins decisively: peak
falls 47-84% and retained bytes fall 80-97% on the four fixtures both admit.
0587's falsification test required route S to be lower in instructions *and*
peak; it is lower in peak only.

Two corrections to 0587 fall out of the same table. First, CORE-1 modelled "the
whole-file read is a further full copy **retained for the document's lifetime**";
it is not. `ca.kwsymphony…doc` retains 1,090,299 bytes for a 1,619,457-byte file,
because the slurped `Vec` is dropped with the package at the end of the match
arm. The retained figure tracks the parsed model, not the file. Second, route E's
peak does include the slurp (2.25× the file on that fixture), so the peak
argument survives even though the retention argument does not.

### Instructions, callgrind isolation pairs (**measured**, secondary)

Profile N and N+M operations, difference the `summary:` totals, divide by M
(`measurements/callgrind.tsv`). Route E is `facade-*`, route S is `snap-*`.

| fixture | E open | E open+`text()` | E open+`paragraph_count()` | E open+`paragraph_text(0)` | S open | S open+`paragraph(0)` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `documentProperties.doc` | 220,758 | 220,842 | 221,568 | 237,611 | **3,464,966** | **7,930,140** |
| `endingnote.doc` | 260,864 | — | — | — | **3,465,105** | — |
| `noheadfoot-litchi.doc` | 283,696 | 283,758 | 287,316 | 392,858 | **3,630,814** | **8,318,112** |
| `table-merged-cells.doc` | 744,360 | 744,556 | 747,616 | 1,819,774 | **5,900,971** | **9,763,935** |
| `picture.doc` | (refused) | — | — | — | **460,073,873** | **766,573,040** |
| `FloatingPictures.doc` | 3,863,545 | — | — | — | (refused) | — |
| `saved-by-table.doc` | 4,614,121 | — | — | — | (refused) | — |
| `ca.kwsymphony…doc` | 5,884,101 | 5,882,859 | — | — | (refused) | — |

Route S costs **7.9× to 15.7×** route E's instructions on the four fixtures both
admit (`endingnote.doc` 13.3×). Callgrind runs SHA-256 in software — 87.95% and
87.55% of route S's total on the two smallest fixtures is
`sha2::sha256::soft::unroll::compress`
(`measurements/callgrind-annotate.txt`) — so these ratios overstate latency and
are reported only to rank the work. `picture.doc` at 460,073,873 Ir per open is
close to half of the 912,551,293 Ir change 0587 measured before 0589 landed,
which is the independent confirmation that 0589's halving is in this base.

### Native cycles, `perf stat -r 3` isolation pairs (**measured**, primary)

200 against 1,200 operations for the small fixtures, 20 against 120 for the
large, difference divided by the operation delta (`measurements/perfstat.tsv`).

| fixture | operation | cycles/op | instructions/op | IPC |
| --- | --- | ---: | ---: | ---: |
| `documentProperties.doc` | E open | **67,643** | 221,339 | 3.272 |
| | E open+`text()` | 68,040 | 221,527 | 3.256 |
| | E open+`paragraph_count()` | 68,185 | 222,483 | 3.263 |
| | E open+`paragraph_text(0)` | 72,793 | 238,211 | 3.272 |
| | S open | **342,797** | 721,100 | 2.104 |
| | S open+`paragraph(0)` | 728,601 | 1,479,492 | 2.031 |
| `endingnote.doc` | E open | **78,591** | 258,749 | 3.292 |
| | E open+`text()` | 79,883 | 258,453 | 3.235 |
| | E open+`paragraph_count()` | 79,734 | 260,653 | 3.269 |
| | E open+`paragraph_text(0)` | 99,782 | 318,759 | 3.195 |
| | S open | **343,717** | 721,735 | 2.100 |
| `noheadfoot-litchi.doc` | E open | **84,029** | 279,405 | 3.325 |
| | E open+`text()` | 84,412 | 279,247 | 3.308 |
| | E open+`paragraph_count()` | 84,765 | 283,374 | 3.343 |
| | E open+`paragraph_text(0)` | 117,505 | 383,512 | 3.264 |
| | S open | **348,065** | 732,485 | 2.104 |
| | S open+`paragraph(0)` | 746,505 | 1,506,888 | 2.019 |
| `table-merged-cells.doc` | E open | **205,021** | 689,263 | 3.362 |
| | E open+`paragraph_text(0)` | 520,752 | 1,702,927 | 3.270 |
| | S open | **435,287** | 832,057 | 1.912 |
| | S open+`paragraph(0)` | 689,908 | 1,284,853 | 1.862 |
| `picture.doc` | S open | 18,906,418 | 24,265,363 | 1.283 |
| | S open+`paragraph(0)` | 31,471,715 | 40,220,852 | 1.278 |
| `FloatingPictures.doc` | E open | 826,274 | 2,739,162 | 3.315 |
| `saved-by-table.doc` | E open | 1,479,317 | 4,777,525 | 3.230 |
| `ca.kwsymphony…doc` | E open | 2,894,998 | 6,764,463 | 2.337 |

Natively, with SHA-NI available, route S costs **2.12× to 5.07×** route E's
cycles on the four fixtures both admit (`endingnote.doc` 4.37×). The gap narrows
from the callgrind ratio exactly as expected and does not close.

**The two routes scale on different variables**, and that is the design's core
finding. Route S's cost fits `217,274 + 12.90 × file_bytes` cycles over two
orders of magnitude — fitted on `documentProperties.doc` (9,728 B) and
`picture.doc` (1,448,448 B), it predicts `table-merged-cells.doc` at 441,895
against 435,287 measured (+1.5%) and `noheadfoot-litchi.doc` at 349,404 against
348,065 (+0.4%). Route E's cost is not proportional to bytes at all: 65,024-byte
`saved-by-table.doc` costs 1,479,317 cycles while 335,360-byte
`FloatingPictures.doc` costs 826,274, because the eager open's terms are text
units and formatting entries (change [0596](0596-doc-eager-open-terms.md)), not
file length. So route S is *worst* precisely where a positional reader is
supposed to win — a large artifact from which the caller wants one paragraph —
because six complete-artifact passes are levied on the artifact's size before any
paragraph is resolved. `picture.doc` is the demonstration: 18.9 M cycles to open
1.45 MB, 6.5× what route E spends opening the *larger* 1.62 MB
`ca.kwsymphony…doc`, after which `paragraph(0)` refuses with
`StructuralContent` anyway.

### Paired timing on the one query both routes answer identically (**measured**)

`Document::paragraph_text(0)` against `SourceSnapshot::paragraph(Position::new(0))`,
on the two fixtures where both routes admit the artifact and return the identical
string. 32 samples per leg, each sample 40 operations, 50 warm-up operations,
order A1 B1 B2 A2, plus two further route-E legs as the A/A control in the same
window (`measurements/bench/`). Nanoseconds per operation.

| fixture | leg | p50 | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| `noheadfoot-litchi.doc` | A (route E) pooled | **26,160** | 26,209 | 26,656 | 26,844 |
| | B (route S) pooled | **164,779** | 164,857 | 165,686 | 166,060 |
| `documentProperties.doc` | A (route E) pooled | **16,337** | 16,313 | 16,663 | 16,681 |
| | B (route S) pooled | **161,768** | 161,788 | 162,772 | 163,351 |

Both directions: route S against route E is **+529.9%** and **+890.2%** at p50
(6.30× and 9.90×); route E against route S is **−84.1%** and **−89.9%**.

**A/A floor, same window:** p50 0.9% and 0.0%, p99 0.5% and 1.9%. The floor is
far below this repository's usual p50 4% / p99 14% because each sample amortizes
40 operations; the deltas above exceed it by two orders of magnitude.

### The fallback penalty R-A would pay on the refused population (**measured**)

The cost of a `SourceSnapshot::open` that is *expected to be refused* — what R-A
adds to route E on 49 of 57 fixtures — measured natively with the refusal
swallowed (`snap-open-try`):

| fixture | bytes | route S probe (wasted) | route E open | R-A total | penalty |
| --- | ---: | ---: | ---: | ---: | ---: |
| `saved-by-table.doc` | 65,024 | 380,936 | 1,479,317 | 1,860,253 | **+25.8%** |
| `FloatingPictures.doc` | 335,360 | 1,592,261 | 826,274 | 2,418,535 | **+192.7%** |
| `ca.kwsymphony…doc` | 1,619,457 | 7,103,653 | 2,894,998 | 9,998,651 | **+245.4%** |

The penalty grows with file size because the wasted probe is two complete
artifact passes. On the largest fixture in the repository, R-A would make
`Document::open` **3.45× slower** and would have changed nothing about the
result.

## Why no routing was implemented

Three independent blockers, in the order they disqualify:

1. **Admission (ADR 0006, `docs/GOAL.md`).** Route S admits four artifacts route
   E refuses with `CorruptedFile`. R-A would silently widen the facade's admitted
   population by skipping the stylesheet validation the eager reader performs.
   0587's own falsification clause — "or widening admission requires weakening a
   validation the eager path performs" — is satisfied.
2. **Capability.** Route S answers one of the DOC arm's eight queries, on 2 of 57
   fixtures, and cannot answer `text()` or `paragraph_count()` at all. Closing
   that means new reader primitives inside `litchi-doc`, which this change's
   scope excludes and which would be a much larger design than a facade route.
3. **Cost.** On the population where both routes work, route S is 2.12-5.07× the
   native cycles, 7.9-15.7× the instructions, 15× the reads, 7× the bytes and
   6.30-9.90× the p50 latency of route E, against an A/A floor below 2%. Its
   only advantage is a 47-84% lower allocator peak.

Blocker 3 alone would not close the item — a memory-for-CPU trade can be right —
but blockers 1 and 2 are contract and capability problems, and the brief's
standing instruction for this case is explicit: stop at a frozen design,
implement only what is value-identical and semantics-preserving. Nothing here is.

**R-B is not implemented either**, although it clears blocker 1 by construction.
It would add a public entry point whose measured benefit is a 47-84% lower
allocator peak on 2 of 57 fixtures, at 6.3-9.9× the latency, for a single
paragraph. That is speculative complexity under `docs/GOAL.md`'s decision rules,
and it is recorded here rather than built so the next reader does not rediscover
it.

## Predicted saving for the admitted population

Stated as the brief requires, for the 4 fixtures route S admits and route E also
admits (7.0% of the corpus), per open:

| metric | direction | size | tier |
| --- | --- | --- | --- |
| allocator peak bytes | **saving** | −47.9% to −83.8% | measured |
| allocator retained bytes | **saving** | −80.1% to −97.3% | measured |
| native cycles | **cost** | +112% to +407% | measured |
| instructions | **cost** | +692% to +1,470% | measured |
| `read_at` calls | **cost** | +1,400% (2 → 30, three fixtures traced) | measured |
| bytes read | **cost** | +582% to +647% (three fixtures traced) | measured |
| `statx` | **cost** | +1,291% (11 → 153, three fixtures traced) | measured |
| p50 latency, one paragraph | **cost** | +530% to +890% | measured |
| queries served identically | — | 1 of 8, on 2 of 57 fixtures | measured |

If a future change removed route S's second identity pass and duplicate index
parse — the two items change [0589](0589-ole2-snapshot-fingerprint-passes.md)
designed and deliberately did not implement — the six complete reads would fall
to four, and the fitted per-byte term would fall from 12.90 to about 8.6 cycles
per file byte (**modelled**, by proportion; not measured). On
`noheadfoot-litchi.doc` that predicts about 305,300 cycles against route E's
84,029: still 3.6×. Route S does not become competitive by removing fingerprint
passes; it becomes competitive only if the per-byte term disappears entirely,
which means a different identity design than the complete-artifact fingerprint
ADR 0006 and record 0105 established.

## Correctness evidence

No production code changed, so there is no behaviour to test. What this record
asserts about behaviour was measured, on the whole corpus rather than on a
sample:

- **Admission census**, all 57 `.doc` fixtures under `test-data/`, both routes,
  with the error kind for every refusal: `measurements/census.tsv`. The crosstab
  above is a direct count of that file.
- **Value comparison**, all 57 fixtures: `measurements/identity.tsv`. For each,
  route E's `text()` length and FNV-1a digest, `paragraph_count()`,
  `paragraph_text(0)` length and digest; route S's `paragraph(0)` outcome, length
  and digest, and the length and digest of the concatenation of every position it
  serves before refusing. The two paragraph digests that this record calls
  byte-identical are equal in that file; no other pair of values is equal.
- **Route S's six complete reads** are counted, not inferred, by a counting
  `ReadAt` adapter, and cross-checked against the whole-process `strace` deltas
  and against the byte arithmetic (99.2% of `picture.doc`'s measured bytes).

### Gates

Run in `/home/zhuhe/code/litchi-worktrees/0609`; tails in
`results/change-0609/gates.txt`.

| gate | result |
| --- | --- |
| `git diff --name-only 8fe9efa55 -- crates/` | empty — no crate file changed |
| `cargo fmt --all --check` | pass |
| `cargo fmt --check` (probe) | pass |
| `cargo clippy --release --all-targets` (probe) | pass, no warnings |
| probe build against this worktree's `litchi`, `litchi-doc`, `litchi-cfb`, `litchi-core` | compiles clean (one pre-existing `dead_code` warning in `litchi`, reproduced on the untouched before checkout) |

Per-crate `clippy`, `test` and `doc` gates are not run because no crate was
touched; the probe's compilation against this tree is the evidence that the tree
is intact. The `dead_code` warning on
`missing_ooxml_catalog_part_error` (`crates/litchi/src/detection_smart/detected.rs:612`)
appears only in the `doc,docx`-without-`pptx`/`xlsx` feature combination the probe
selects, is present identically on the before checkout, and is **pre-existing**.

## Validation preserved

Everything, by construction: no reader, no fence, no limit and no refusal moved.
The record's purpose is to state precisely why one of them would have had to.

## Limitations — what is not claimed

- **No speedup or regression is claimed**, and no claim-registry entry is made.
  The numbers rank two existing routes against each other; neither was changed.
- **No cold-cache, physical-device, range-source, cross-platform or
  concurrency-scaling result.** Every measurement is warm, local, single-threaded
  and on one host with SHA-NI. A host without SHA-NI would widen route S's cycle
  disadvantage; a network filesystem would widen its `statx` and read
  disadvantage. Neither was measured.
- **RSS was not measured**; the peak and retained figures are counting-allocator
  bytes, which exclude the page-cache copy, stack and any allocation the global
  allocator does not see. They are the right metric for comparing two routes'
  heap behaviour and the wrong one for an absolute RSS claim.
- **The corpus ceiling is 1.62 MB**, with no DIFAT-scale artifact and no
  4,096-byte-sector CFB. Route S's per-byte term means its disadvantage grows
  with size, so the corpus understates rather than overstates it — but the fit is
  extrapolated beyond 1.45 MB and is labelled modelled where it is used.
- **Feature combination.** The facade was measured with `features = ["doc",
  "docx"]`. The `docx`-disabled arm reaches the same `read_path_source_bytes`
  slurp through `DocumentSourcePathDetection::Bytes` (`detected.rs:2167`,
  `:2488`) rather than `DocxSourcePathDetection::Bytes`, so the ingress is the
  same; the ODT/OOXML probes that precede it differ and were not measured
  separately.
- **R-B was designed, not prototyped.** Its cost figures are route S's measured
  figures; no patch exists in this packet, because the capability blocker makes
  one premature.
- **`Position` semantics.** Route S's `Position` is a zero-based main-story
  paragraph position and route E's `paragraph_text(i)` index is a position in the
  full paragraph projection. They coincide at 0 on the two fixtures compared;
  this record does not claim they coincide elsewhere, and the walk column shows
  they cannot be compared beyond position 0 on this corpus.

## Retained evidence

[`results/change-0609/README.md`](results/change-0609/README.md) — the contents
table, provenance, the probe source and every raw output cited above.
