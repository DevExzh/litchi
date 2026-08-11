# XLSX plain worksheets skip the x14ac capture pass

Date: 2026-08-11

Production base: `9454f49fa573f063ba40eaa341b26faec2e205c7`

Scope: private XLSX worksheet semantic reads and changed-sheet commit readback.
OLE2, RTF and ODF production code are unchanged, and iWork/IWA crates were
explicitly excluded.

## Hypothesis and profile evidence

Every selected worksheet parse first ran the narrow x14ac collector, then MCE
processing and the complete semantic worksheet parser. The generated benchmark
worksheets contain neither an x14ac `dyDescent` attribute nor MCE markup, so
the first namespace-aware traversal could not produce a semantic value. A
changed-sheet commit performs that work twice: once while resolving the source
store and once while validating the final compacted worksheet.

The retained profile from change 0025 attributes 4.46% exclusive time to
`x14ac::capture_inner` in the medium commit-plus-first-read case. Shared
namespace-reader work makes the inclusive opportunity larger. The rejected
direct action-plan flattening in change 0030 reached only 1.61% p50 in its best
formal cell, so removing a complete input traversal was the stronger next
measured hypothesis.

## Change and error-order preservation

The private worksheet parser now uses `memchr::memmem` to look for the
`dyDescent` local-name token. When it is present anywhere, including a comment
or text node, the established x14ac collector runs unchanged. When it is
absent, the parser starts with empty extension values and proceeds through the
same complete MCE and semantic parsers.

On any downstream rejection in the no-token path, the original x14ac collector
runs before the error is returned. If it rejects, its original typed error wins;
otherwise the downstream error is returned. This preserves the former
extension-first malformed-input precedence without charging successful plain
worksheets for the extra traversal. Direct x14ac attributes, AlternateContent,
namespace resolution, depth and row-number validation are unchanged. The
focused `parse_defaults` path is unchanged.

No public API, dependency edge, cache, retained state, limit, runtime, lock,
unsafe-code boundary, transaction, patch, publication, signature/encryption
policy or save behavior changes.

## Corpus and protocol

The existing deterministic medium XLSX corpus has four 32-by-32 worksheets,
4,096 cells, 41 distributed 1% updates, nine archive members, a 15,254-byte
archive and SHA-256
`9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6`.
The dense-wide corpus has two 256-by-256 worksheets, 131,072 cells and 1,311
distributed updates. Both generated worksheet forms lack the `dyDescent`
token.

Medium primary and guard ABBA runs used 50 warmups and 500 samples per leg on
CPU 2. Dense-wide used five warmups and 50 samples per leg. Results below pool
the two same-state legs: 1,000 medium or 100 dense-wide samples per state.
Mean-delta intervals use an independent-sample normal approximation over the
pooled distributions; every interval excludes zero.

Before executable SHA-256:
`38536012fac010c736a547ce1bed5b79e34570a0fdddbca1ac89855373b9ff68`.
After executable SHA-256:
`8f28547d8a4ccb6876d1d666a2cbed151e5e025f77dfc72ffd3639a21317ae03`.
Their `.text` section hashes are
`dbff56d4765d750a44aa6eaafd5f286b34522aa85344e2d34a020fe1c8edbd41`
and
`20f80d7fe7a3df575457a8c353f2a7b00ad3b03c4518a318b97361b79beea117`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The harness source and case definitions are identical
between states; only the production XLSX parser differs.

## Medium latency

| Case | Before p50 | After p50 | p50 | p95 | Mean | 95% interval for mean delta |
|---|---:|---:|---:|---:|---:|---:|
| One-cell commit | 3.432 ms | 2.723 ms | **-20.66%** | **-17.34%** | **-20.46%** | `[-21.07%, -19.84%]` |
| One-cell commit + first read | 3.440 ms | 2.762 ms | **-19.71%** | **-20.06%** | **-20.10%** | `[-20.54%, -19.67%]` |
| 1% commit | 13.845 ms | 10.974 ms | **-20.74%** | **-19.72%** | **-20.52%** | `[-20.78%, -20.27%]` |
| One-cell commit + save | 3.879 ms | 3.130 ms | **-19.31%** | **-18.47%** | **-19.16%** | `[-19.57%, -18.75%]` |
| 1% commit + save | 15.485 ms | 12.372 ms | **-20.10%** | **-19.90%** | **-20.06%** | `[-20.31%, -19.81%]` |

Both after legs have lower p50 and mean than both before legs in every primary
case.

## Dense-wide scale

| Case | Before p50 | After p50 | p50 | p95 | Mean | 95% interval for mean delta |
|---|---:|---:|---:|---:|---:|---:|
| One-cell commit | 215.985 ms | 173.821 ms | **-19.52%** | **-18.23%** | **-19.16%** | `[-19.75%, -18.57%]` |
| One-cell commit + first read | 281.956 ms | 215.541 ms | **-23.55%** | **-23.77%** | **-23.74%** | `[-24.27%, -23.21%]` |
| 1% commit | 436.605 ms | 350.938 ms | **-19.62%** | **-21.05%** | **-20.03%** | `[-20.56%, -19.49%]` |
| 1% commit + save | 521.802 ms | 427.209 ms | **-18.13%** | **-16.98%** | **-17.74%** | `[-18.17%, -17.31%]` |

The dense validated-store retention limit still takes its established cold
fallback; this result removes parse work and does not retain the 131,072-cell
store.

## Read and no-op guardrails

Cold medium `xlsx_first_cell` improves 36.41% p50 and 35.57% mean. Cold
`xlsx_full_cell_scan` improves 34.81% p50 and 34.86% mean. The preloaded
narrow-column traversal stays flat at 390 ns p50 and improves 0.48% mean.
Lazy owned open, which does not parse worksheet payloads, changes by +0.89%
p50 and +0.81% mean; its 122.7/140.0 us p99 shift (+14.11%) is disclosed.

No-op commit remains below the reliable timing floor: 120/220 ns p50 and
198/390 ns mean. No-op commit plus exact save is 571/421 ns p50. The fast path
is unreachable inside either timed operation because no worksheet is parsed;
the sub-microsecond absolute movements are not treated as a regression or a
speed claim.

## Allocations, memory and counters

Matched Heaptrack processes used two warmups and ten dense-wide 1% commits plus
one complete post-timing verifier:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 210,837,749 | 157,627,475 | **-25.24%** |
| Temporary allocations | 33,547,137 | 33,533,200 | -0.04% |
| Peak heap | 91.79 MiB | 91.79 MiB | unchanged |
| Heaptrack RSS | 123.32 MiB | 123.37 MiB | +0.04% |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

Uninstrumented GNU Time processes with five warmups and 20 dense-wide commits
report 115,452/116,092 KiB maximum RSS (+0.55%).

Matched medium 1% `perf stat` processes used 50 warmups and 500 samples:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| Task clock | 7,926.93 ms | 6,303.81 ms | **-20.48%** |
| Cycles | 38,944,820,496 | 30,983,506,349 | **-20.44%** |
| Instructions | 158,223,258,750 | 125,815,787,592 | **-20.48%** |
| Branches | 32,069,027,224 | 25,473,924,280 | **-20.57%** |
| Branch misses | 39,326,369 | 36,760,474 | -6.52% |
| Cache references | 743,193,698 | 698,948,047 | -5.95% |
| Cache misses | 58,796,223 | 57,836,456 | -1.63% |
| Page faults | 9,866 | 9,849 | -0.17% |
| Context switches | 89 | 68 | -23.60% |
| CPU migrations | 0 | 0 | unchanged |

The instruction, cycle and branch reductions match removal of two complete
namespace-aware traversals per changed-sheet commit.

## Correctness and review gates

- New focused tests prove plain-token detection, conservative false-positive
  fallback, and the historical extension-first error for malformed plain XML.
- Existing direct x14ac, active MCE branch, depth, invalid descent, row-order,
  defaults, rewrite, namespace-injection and malformed worksheet tests pass.
- All-feature/all-target XLSX tests pass: 732 library tests plus every
  integration and example target, including exact patch/inverse, stale-source,
  MCE, encryption, signatures, protection, durable JSON, source preservation
  and bounded validated-store cases.
- Warning-denied production-library Clippy passes. Warning-denied all-target
  Clippy still reports only the unchanged `xml_maps_api_compat` needless
  question mark and two existing test-module inception lints.
- All 24 harness tests and warning-denied all-target harness Clippy pass. The
  two known ODF GenericArray deprecation warnings remain in harness dependency
  builds.
- Formatting, `git diff --check`, raw JSON digests, binary identities and
  staged-scope checks are commit gates.

Raw ABBA, guard, Heaptrack, GNU Time and `perf stat` harness reports are under
`docs/performance/results/`; their digests are in
`xlsx-x14ac-scan-sha256.txt`.

## Remaining non-iWork work

1. OLE2: add a media/opaque-heavy common editor case before implementing the
   shared CFB `Arc<[u8]>` writer handoff.
2. ODF: add a media-rich ODP source-backed text-box publication case before
   reusing the common raw `content.xml` preservation path.
3. RTF: add a table-text open variant before changing the remaining
   table-state clones.
4. OOXML: broader planning/emission fusion and source-backed editable
   publication remain separate architecture and measurement problems. Do not
   revive the rejected direct action-plan flattening.

iWork remains deferred while the `iwa-*` crates are modified independently.
