# 0429: PPTX provider/native baselines and ZIP short reads

The ZIP central-directory iterator now assembles a fixed header from buffered
and unread bytes, parses a header already fully buffered, and returns `Eof`
for a truncated trailing record. This corrects valid capped-read failures and
silent truncation acceptance. Four regressions fail before and pass after the
bounded fix; six expanded cases include ZIP32/ZIP64 extra fields and scratch
boundaries. It is a correctness enabler, with no measured speedup claim.

The separate performance harness adds bytes, warm positional files, and
explicit capped/delayed range adapters to the matched synthetic plain/media-rich
PPTX cross-copy lifecycle. Two unmodified POI image fixtures cover selected-image
metadata/read ownership separately; native cross-copy remains outside this
positive evidence. Native shape metadata, returned descriptor identity, and
payload hashes are checked independently, including payload survival after
package-owner drops.

Thirty-two fresh release processes retain 960 samples and 8,160 phase points,
with three warmups and 30 samples per process. All final releasable managed
Memory/Objects/Depth gauges are zero. Non-RSS numeric phase observations match
exactly within processes and across repeats, including logical read histograms.
All 27 repeat review flags are RSS points in media-rich bytes/file/short-range
lanes; those lanes already differ at baseline. No API-duration/rate median
crosses the 5% review threshold. Comparative RSS conclusions are withheld.

The following values are median sums of immediately timed public API calls,
in milliseconds. Setup, source construction/copies, sink reservation, corpus
gates, checks, observations, and drops are excluded. These are baselines from
one implementation, not a causal before/after comparison. Full p95/p99, rates,
bootstrap median intervals, and all raw points are in the bundle.

| Operation/input | Provider | R1 API sum ms | R2 API sum ms |
| --- | --- | ---: | ---: |
| Cross-copy/plain | bytes | 2.228 | 2.223 |
| Cross-copy/plain | file | 2.533 | 2.530 |
| Cross-copy/plain | range 4096 / 0 µs | 2.223 | 2.234 |
| Cross-copy/plain | range 65536 / 200 µs | 122.310 | 122.324 |
| Cross-copy/media-rich | bytes | 257.709 | 252.880 |
| Cross-copy/media-rich | file | 255.108 | 255.056 |
| Cross-copy/media-rich | range 4096 / 0 µs | 258.566 | 254.129 |
| Cross-copy/media-rich | range 65536 / 200 µs | 645.529 | 645.286 |
| Selected image/POI slide | bytes | 0.709 | 0.711 |
| Selected image/POI slide | file | 0.770 | 0.769 |
| Selected image/POI slide | range 4096 / 0 µs | 0.713 | 0.710 |
| Selected image/POI slide | range 65536 / 200 µs | 27.328 | 27.321 |
| Selected image/POI video poster | bytes | 0.359 | 0.363 |
| Selected image/POI video poster | file | 0.380 | 0.381 |
| Selected image/POI video poster | range 4096 / 0 µs | 0.360 | 0.361 |
| Selected image/POI video poster | range 65536 / 200 µs | 11.547 | 11.548 |

Files are recently written or hash-read; these are warm observations. Range
caps limit returned chunks, and the 18-bin histogram counts caller-requested
buffer sizes. Fixed delay is a caller-side sleep including scheduler overhead,
not a calibrated bandwidth or remote-service model. Managed gauges, cache
retention, source-read counters and process RSS remain different scopes.

The [bundle](../results/change-0429/README.md) retains machine/build/protocol
identities, raw reports, exact oracles, all development failures, an independent
validator, mutation controls and lossless artifact custody. The
[ADR review](../results/change-0429/adr-review.md),
[ZIP review](../results/change-0429/zip-review.md), and
[native source audit](../results/change-0429/native-review.md) record the
ownership and refusal boundaries. The inspected original/LibreOffice shapes
fixtures remain unchanged typed-refusal controls.

ZIP tests pass 439 tests with two ignored, followed by the final six expanded
regressions. OPC/PPTX tests pass 1,290 with three ignored. ZIP strict Clippy and
warning-denied docs pass, as do the scoped harness tests/docs/formatting,
boundaries, registered claims and CRUD index. The 16-lane CLI preflight rejects
416 report mutations and preserves existing output files. Harness strict
Clippy retains the same 29 old findings in 17 groups and has no findings in
modified harness modules; the failed command is not called passing. The
existing ZIP fuzz target passes 1,000 deterministic AddressSanitizer runs from
13 pinned seeds in an isolated copy, which is removed after custody checks.

A documented validator amendment corrects the omitted 100-sample supplementary
profile policy. The original build/verifier/policy and 131 unchanged capture
artifacts are retained and hashed. All 32 baseline reports are revalidated;
none is recaptured. The completed bytes CPU recording is reused, and its
unretained original start timestamp is explicitly unavailable. The failed
profile wrapper remains in evidence.

The global non-iWork goal remains open: broader native/size and CRUD coverage,
true cold I/O, physical-copy/allocator attribution, semantic streaming and
explicit scaling remain incomplete. No full-lifecycle speedup, general leak,
physical-copy or scaling claim follows from this batch.

The two supplementary CPU profiles retain 17,217 and 16,635 samples. The
[profile review](../results/change-0429/profile-review.md) separates the frozen
namespace-filter result from an additional analysis of unqualified DWARF names.
Only 2,396/2,074 samples have resolved iteration ancestry; the large unassigned
population prevents complete API attribution or a new optimization ranking.
Raw perf data, decoded stacks and both analyses remain independently inspectable.

Portable replay passes before and after copied-executable cleanup, validating
all 32 formal reports and their 832 mutations, both profiles, and the explicit
amendment chain. An output change with updated receipt/amendment hashes still
fails the cross-repeat identity gate. The supplementary symbol analysis also
passes from an exported bundle after cleanup. The original executable and both
existing target directories are preserved; the user goal file is unchanged.
