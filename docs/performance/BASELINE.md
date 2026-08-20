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
**341 selectable cases**; 200 was the count before the opt-in ODF `mimetype`
repair-plan selector and later opt-in selectors were added. The
historical 36-default-case/198-default-record tranche remains measured as
documented below; newer selectable cases do not inherit those measurements.

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
page-aligned sources on an allowlisted block-backed filesystem after source
`fsync`, accepted `posix_fadvise(DONTNEED)`, and a strict external `fincore`
JSON proof of zero resident, dirty, and writeback bytes. The timed operation
must produce a positive process `/proc/self/io` `read_bytes` delta. Prepared
query controls are excluded. Ineligible host/proof conditions are explicit
statuses and emit no timed result. This proves page-cache and process-I/O
conditions only; it makes no physical-media or device-cache claim and has no
captured performance result. See
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
