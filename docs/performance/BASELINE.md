# ZIP, OPC, and CFB substrate baseline

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

## Parallel scaling observation

`opc_open` currently uses Rayon through the global pool. Separate processes set
`RAYON_NUM_THREADS` to 1, 2, 4, 8, and 12; each cell used 10 warm-ups and 50
samples. Raw reports are the
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
3. Replace hidden global scheduling only with an explicit bounded execution
   contract. The current scaling knee is four large-entry tasks on this host.
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

The harness now has 14 cases and 97 default result records. In addition to the
original matrix it measures owned OPC open, one-Part mutated save, and public
DOC/XLS/PPT writer packaging with tiny, moderate, and 4-5 MiB stream-heavy
shapes. Scheduled CI records the deterministic full matrix without applying
machine-noisy latency thresholds.

## Current stable tranche update

The stage-1 records above are retained unchanged. The current harness has **147
selectable cases**: 36 default cases and 198 default records, plus six opt-in
simulated-range cases, two opt-in execution-scaling cases, one opt-in XLSX
commit/read attribution case, four opt-in opaque-heavy common OLE2 publication
stage/control cases, one opt-in source-backed OPC one-Part publication case,
one opt-in source-backed DOCX semantic publication case, one opt-in media-rich
PPTX semantic publication case, four opt-in matched same-slide/multi-slide
PPTX batch cases, six opt-in media-rich ODT paragraph,
line-break, inline-run, hyperlink, insertion, and removal publication cases,
14 opt-in matched XLSX calculation-metadata/defined-name/page-break/page-margin/print-options/page-setup/sheet-protection
publication cases, 16 opt-in DOCX/PPTX semantic
cases, nine opt-in RTF semantic case names across four capability-bounded
variants (33 tiny / 58 tiny-plus-large rows), 23 shape-selected ODT/ODS/ODP
semantic cases, eight fixed media-rich ODF cases, and 21 opt-in native
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
  correctness and boundedness, not a new aggregate latency claim.
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
- Consecutive packaged ODT plain-text replacements now share one mutable
  candidate, content publication, reopen and compact audit while retaining
  ordinary scalar durable operations. The large 100-edit/save p50 falls from
  906.439 to 15.615 ms (-98.28%, 58.05x), allocation calls fall 96.13%, and
  scalar one-edit guards remain neutral. See
  [`0045`](changes/0045-odt-coalesced-paragraph-publication.md).
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
[`0007`](changes/0007-source-backed-opc-and-facades.md). Source cache bytes are
bounded by `SourceCacheLimits`, but are not yet charged to the hierarchical
`Budget`.

Consolidated changed-crate tests, formatter checks, warning-denied production
Clippy and rustdoc gates passed. The current ODS all-target Clippy gate retains
the unrelated pre-existing test-only findings recorded in change 0027. The ODT
tranche compiled the ODF fuzz target offline; the PPT and ODS tranches have no
dedicated fuzz target in the current tree. A workspace all-target/all-feature
gate was not rerun because iWork was explicitly excluded while its crates are
changing independently.
