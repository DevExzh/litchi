# PPT slide-order open reuses its validated CFB

Date: 2026-08-11

Production base: `4a2020d5c61bb14eac6646b6b50c7f1fa46c8df5`

Scope: native binary PPT root slide-order snapshot capture only. OOXML, RTF
and ODF production code are unchanged, and iWork/IWA crates were explicitly
excluded.

## Hypothesis

`slide_order::Snapshot::from_bytes` first opened the PPT through `Package`,
parsed the presentation, and retained that package as the snapshot owner. It
then passed the same source bytes to `Editor::inspect_live_document`, which
opened a second `OleFile` before independently resolving `PowerPoint Document`
and `Current User`. The second open repeated CFB header, FAT, directory,
MiniFAT, and allocation-topology work even though the first validated
`OleFile` remained available inside `Package`.

Passing the already-open `OleFile` into the live-document inspection should
remove only that duplicate CFB index construction. The independent stream
selection, current-user validation, live persist mapping, document slicing,
slide-directory agreement, document parse/round-trip, and review-history
checks must remain.

## Change

The PPT object editor now has a crate-private generic
`inspect_live_document_from_ole` helper. The existing byte-ingress helper
still opens an `OleFile` and delegates to it, so its behavior and callers are
unchanged. Root slide-order snapshot capture calls the private helper with the
validated `OleFile` already owned by `Package`.

The helper still lists the complete stream directory, resolves the two
required stream names independently, rereads both streams, validates Current
User, rebuilds the live persist mapping, and returns the selected document
bytes. The snapshot then retains all existing persist-ID agreement,
slide-order, document-structure, review-history, package-ownership, resource,
and public-reader checks.

This changes no public API, dependency edge, source identity, patch or edit
semantics, publication path, limit, security policy, runtime, lock, cache,
durability, or unsafe-code boundary.

## Dedicated public benchmark

The opt-in `ppt_slide_order_snapshot_open` case captures the public
`litchi_ppt::slide_order::Snapshot` from owned exact-source bytes. Corpus
cloning occurs before timing. The timed region includes the complete root
snapshot validation; a full generic public-PPT semantic verification runs once
outside timing after all samples. The case is included in native OLE2 smoke
and scheduled release matrices, taking that family to 19 selectable cases and
38 tiny/large records.

The baseline executable contains the exact same new harness but production
code from the named base. Its SHA-256 is
`8b137f1a166d49e006ff876915c3dda849e5fea96f6fce390bb169bac58edaec`.
The final after executable SHA-256 is
`1c85d778da1bb54770a3d5ceda2930f157ef703b3c8ed4dc81ff875e44dcfbb0`.
Its `.text` section matches the measured after executable and has SHA-256
`058af970f211669bad857e2d5e1857812d79d8001245cdb8f225955b76424cb7`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large PPT contains 144 authored
shapes, four CFB streams, a 40,960-byte archive, and a 37,385-byte
`PowerPoint Document` stream. Archive SHA-256 is
`229052cd918c0e5b7ef44070bafe20833531eee119b5943b18499503e225ff52`;
the document-stream SHA-256 is
`bef446ada643821b87531c06be7564b7ff8ca5539bb6a39766fbd28c11f65523`.

## Matched latency measurement

Four short ABBA cycles each used 100 warmups and 2,000 samples per leg.
Pooling 16,000 samples per state gives:

| Large PPT root snapshot open | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 37.522 us | 34.227 us | **-8.78%** |
| p95 | 54.676 us | 43.430 us | **-20.57%** |
| p99 | 59.712 us | 52.121 us | **-12.71%** |
| mean | 39.677 us | 35.477 us | **-10.58%** |

The approximate independent-sample 95% interval for the mean delta is
`[-10.90%, -10.27%]` of the before mean. Every matched after leg has a lower
p50 than its before leg. Same-state cycle p50 drift is about 1.2% or less.

Raw primary files are
`abba-ppt-slide-order-root-repeat-{1,2,3,4}-{before-a,after-a,after-b,before-b}.json`.
Their SHA-256 digests are:

| Cycle | Before A | After A | After B | Before B |
|---:|---|---|---|---|
| 1 | `f3d6dde16ffc80f92cf55e95684f84ac71d5882f0cbee055d020d08c061c67be` | `140df871bc7a3ed6b1576f7f663a5e86d3034718c992da04aa4ad434bfe2d3ea` | `22157ca0598fdcae98b85eb21263a6e65c0368b2d37123d0df36ab0a8aef5e19` | `5d706cdb69b289926ac4a588f876138774fbe9ea9d3f93f44fa4e2df06ffe7a3` |
| 2 | `b919b06a7ab19f40e6e82764eccd5f2a287e76c3ef23f6f4133dffee59c7da47` | `c7b9b9fbf63e8c298221dc01387847882f31405578eef6781b235ec8624782bb` | `3c8bf7bdfbe23c7ab754d29086dd703d7513052ca8e97b537d9f01fcced65a78` | `b6bdd050523e97ba60c8a33534f41f366139f5a27f2e2563d29dcfe1007be968` |
| 3 | `557b79601d24a0fe058cc7743feccab19e005ad7ac30e27380db67f50a7ff4a1` | `dce8eb60e4ef75830147cdce6da46d99fc645afb310c3d1f5bb83ae93040e759` | `e9004dadf300ca245d1963e51171589ba59691719efd40c18f59882fa0282eab` | `671d4905d5cb62464e3361c6d3993fe2db78dfb44fdb0cb7f31d20e533d4af56` |
| 4 | `c31151f35dabc9c9867f298b5ac2d028ba19adba883f546299a51b48ab5e89ae` | `9334bc6e807b1579bc7dee8bdbf98100dedfcad20c7459deecdbdcf22f35299f` | `b67ba0f7f43afbd0ace732427a354d802ebc1e35cbabdd6c8b84d6206b4663c6` | `d256e05ec6657e56650bf7c49b132a9e2392c164f0b65c9db6995dff431693c3` |

## End-to-end and reader guardrails

The ordinary large PPT one-shape edit/save ABBA run used 50 warmups and 500
samples per leg. The reader guards used 100 warmups and 2,000 samples per leg,
except exact no-op at 10,000 and the tiny root case at 4,000 samples per leg.

| Guardrail | p50 delta | p95 delta | p99 delta | Mean delta |
|---|---:|---:|---:|---:|
| Large one-shape edit/save | -1.78% | -1.54% | +3.91% | -1.68% |
| Ordinary open | +0.60% | +0.08% | -8.76% | -0.55% |
| List slides | -0.24% | -0.74% | -5.41% | -0.57% |
| One selected shape, repeated | +0.32% | -0.06% | -0.05% | +0.33% |
| Full text | -4.85% | +0.80% | -1.57% | -4.19% |
| Exact no-op edit/save | +0.54% | +2.82% | +2.66% | +0.85% |
| Root snapshot open, tiny | -13.67% | -11.84% | -16.99% | -12.89% |

The first selected-shape guard had a noisy after-B tail, moving p95 +21.79%
and p99 +12.30%. It cannot execute the changed root-snapshot path. The trigger
was retained as `abba-ppt-slide-order-one-*.json` and repeated in four full
ABBA cycles. The repeated 16,000-sample/state result above is neutral at every
reported percentile and mean; raw repeats are
`abba-ppt-slide-order-one-repeat-*.json`.

The one-edit path includes more work than root capture and improves 1.78% p50,
so it is a supporting end-to-end observation rather than the headline. All
raw guard and edit reports are retained beside the primary files.

## Allocations, RSS, counters, and CPU attribution

Matched Heaptrack processes used 1,000 root-snapshot samples and one complete
post-timing verifier:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 897,632 | 852,632 | **-5.01%** |
| Temporary allocations | 65,472 | 57,472 | **-12.22%** |
| Peak heap | 808.53 KiB | 808.53 KiB | unchanged |
| Heaptrack RSS | 11.74 MiB | 11.43 MiB | -2.64% |
| Leaked bytes | 544 B | 544 B | unchanged |

Exactly 45 allocation calls disappear per timed root capture. Sampled profiles
also move CFB metadata frames in the expected direction:
`directory_name_data` falls from 6.49% to 5.06% exclusive share,
`OleFile::load_directory` from 0.90% to 0.59%, and `OleFile::load_fat` from
0.76% to 0.33%. These relative shares support the mechanism but are not added
or interpreted as removed wall time.

GNU Time ABBA processes used 100 warmups and 20,000 samples per leg. Maximum
RSS is 30,976/30,848 KiB before and 30,976/30,976 KiB after; it is flat at the
128 KiB measurement granularity.

Matched `perf stat` ABBA processes at the same sample count give:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 1,797.493 ms | 1,679.638 ms | -6.56% |
| cycles | 8,659,344,673 | 8,065,724,599 | -6.85% |
| instructions | 28,647,281,256 | 25,907,249,293 | -9.57% |
| branches | 5,670,258,767 | 5,086,524,664 | -10.29% |
| branch misses | 50,566,974 | 47,768,516 | -5.53% |
| cache references | 833,138,574 | 788,852,835 | -5.32% |
| cache misses | 7,698,874 | 8,853,681 | +15.00% |
| page faults | 17,232 | 17,233 | effectively flat |
| context switches | 69 | 76 | +7 events |
| CPU migrations | 2 | 1 | -1 event |

The cache-miss increase exceeds the 5% review threshold and is retained. The
absolute miss ratio remains about 0.9% to 1.1% of cache references, while
direct latency, task time, cycles, instructions, branches, allocation calls,
peak heap, and RSS all improve or remain stable. No cache-locality improvement
is claimed.

Raw evidence is in `ppt-slide-order-{before,after}-heaptrack.txt`,
`ppt-slide-order-{before,after}-perf-report.txt`,
`ppt-slide-order-perf-stat-*.csv`, and `ppt-slide-order-time-*.txt`.

## Correctness verification

- a focused differential test proves that already-open and byte-ingress
  live-document inspection return identical persist identity and bytes;
- complete all-feature `litchi-ppt` unit, integration, security,
  real-producer, encryption/protection, signed-package and doctest suites pass;
- warning-denied all-target/all-feature Clippy and warning-denied crate rustdoc
  pass, including small pre-existing link corrections in the touched crate;
- the benchmark case verifies exact corpus hashes, slide counts and a complete
  generic public-reader semantic pass outside timing;
- the complete harness test suite and warning-denied Clippy pass; and
- formatting, all 60 JSON reports, `git diff --check`, and final staged-scope
  checks are commit gates.

No dedicated `litchi-ppt` fuzz target exists in the current tree. A
workspace-wide gate was not run because iWork was explicitly excluded while
its crates are being modified independently.

## Next non-iWork audits

1. ODF: profile folding the ODP transition-style prepass into the main page
   parser; every current slide query scans `content.xml` twice, but this is a
   broader parser refactor requiring differential malformed-input guards.
2. OOXML: benchmark XLSX commit plus the first public read of the touched
   worksheet; `SourceBackedPackage::into_opc_package` has no production caller
   today, so clone optimization there is deferred.
3. RTF: add editable byte-1252, read-only LZFu, LibreOffice watermark, and
   relative-font-size fixtures before another parser specialization.
4. OLE2: attribute remaining DOC/XLS final publication and independent public
   readback separately; do not remove either correctness boundary.

iWork remains deferred while the `iwa-*` crates are modified independently.
