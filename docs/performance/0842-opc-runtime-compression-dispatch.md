# 0842 — restore runtime compression dispatch

Adopt explicit `flate2/runtime_detection` in the workspace and independent
soapberry-zip declaration, with minimum flate2 1.1.10. The measured generic OPC
large read p50 improves 14.46–21.32%; the XML/repeated changed-Part cases improve
33.81–38.36%. Exact ZIP bytes remain identical across both legs. This restores
a feature lost when flate2 1.1.10 moved zlib-rs runtime detection behind a
separate opt-in feature with defaults disabled. It is not a new compressor,
compression policy, or semantic Office benchmark.

## Change and provenance

Base: `45ca1ca9fe9ff04f1ba14ded41cfe620629ffaa1`. Both measured binaries use
flate2 1.1.10 and zlib-rs 0.6.7, the same probe, compiler and release settings.
Resolved feature additions are exactly flate2/runtime_detection and zlib-rs/std.
The existing dependency supplies CPU detection and scalar fallback; no unsafe
library code, API, resource policy, compression level, framing, durability,
preservation rule or execution policy changes.

Historical 0836 used flate2 1.1.9 with detection already enabled, so it is not
the before timing for this experiment. Root Cargo.lock stays byte-identical.
The live perf-baseline lock moves 1.1.9→1.1.10 and removes unused miniz/adler
packages. Perf-execution gains the crc32fast feature edge. Native-resave also
refreshes stale local crypto/DOC/ODT/XLS edges and adds same-file/winapi-util;
the identical drift occurs with original compression declarations. Compared
with that normalized baseline, only flate2’s crc32fast edge changes. No existing
registry package version/checksum changes in those last two locks. Native-resave
retains its existing zlib-rs 0.6.8; its locked check is integration evidence,
not part of the measured 0.6.7 comparison.

The final two tool locks are not inputs to the probe or root workspace builds.
[Source-final](results/change-0842/source-final.json),
[lock control](results/change-0842/native-resave-lock-drift.json), and
[candidate patch](results/change-0842/candidate.patch) bind that distinction.

## Protocol and results

AMD EPYC 9R45, x86_64 Linux, rustc 1.95.0 / LLVM 22.1.2; native children pinned
to CPU 12, default allocator, release opt-level 3 / LTO / one codegen unit.
Full host/environment records are retained in the packet. Four opaque OPC parts
have Office-like names but do not form semantic DOCX/XLSX/PPTX documents. Tiny
has four 64-byte parts; other shapes have three 1 MiB XML parts and one 1 MiB
opaque part. Only random’s opaque member is pseudorandom; the others repeat a
deterministic pattern. Content types and package relationships make six ZIP
members.

Write times `PackageWriter::to_bytes` on a prepared package. Read times ZIP
reader construction and all six member reads. Changed times source-backed
open, one opaque-Part replacement plan, and publication, with input/replacement
clones prepared before the clock. All outputs are verified outside the clock.
Write/changed decode output twice during untimed validation, so cache, thermal
and frequency effects are coupled to validation. These results describe that
fixed warm protocol; they do not isolate compressor-only cycle shares.

The fixed native schedule ran six alternating blocks × twelve cases × two
legs, with 30 samples after three warmups. Every one of the 144 processes ran
once. Ratios below are medians of six matched block ratios, not ratios of the
displayed medians. Intervals use 10,000 paired bootstrap draws, seed 842084.
A p50/RSS ratio above 1.05 with lower endpoint above 1 triggers review.

| Shape | Operation | Before p50, ms | After p50, ms | Paired ratio | 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | write | 0.036261 | 0.036435 | 1.00302 | [0.99754, 1.01326] |
| tiny | read | 0.007290 | 0.007305 | 1.00274 | [0.99939, 1.00761] |
| tiny | changed | 0.029050 | 0.029140 | 0.99861 | [0.99036, 1.01356] |
| xml | write | 16.279178 | 15.970156 | 0.98101 | [0.97436, 0.98377] |
| xml | read | 1.306696 | 1.042560 | 0.79818 | [0.79403, 0.79952] |
| xml | changed | 0.480142 | 0.349847 | 0.66192 | [0.62075, 0.74884] |
| random | write | 26.284749 | 25.756556 | 0.97964 | [0.97770, 0.98105] |
| random | read | 1.519176 | 1.299931 | 0.85536 | [0.85203, 0.86368] |
| random | changed | 14.243754 | 13.876592 | 0.97455 | [0.97123, 0.97797] |
| repeated | write | 12.328622 | 12.030170 | 0.97441 | [0.97160, 0.97930] |
| repeated | read | 1.053849 | 0.828569 | 0.78680 | [0.78271, 0.79038] |
| repeated | changed | 0.513132 | 0.318976 | 0.61644 | [0.55718, 0.73162] |

No original p50/RSS guard or block-spread flags. All 84 metric rows, including
p95/p99, mean latency, whole-child RSS and faults, appear in
[native.md](results/change-0842/native.md) and [CSV](results/change-0842/native.csv).
Original RSS paired point ratios range from 0.98122 to 1.00834. RSS includes
fixtures, validation, and retained reports; it is not an operation heap peak.
No allocation, copied-byte, hardware-counter, cold-cache, remote, filesystem
durability, native-Office, or scaling improvement is claimed.

## Tail review and separate confirmation

Two original p99 points exceeded 5%: tiny write 1.29372 [1.02437, 1.71656] and
tiny read 1.39455 [0.94048, 3.38554]. The original p99 is the maximum of only 30
samples. Both observations remain in the main results and adoption record;
they were not deleted or merged with later results.

A separately [frozen tail plan](results/change-0842/tail-plan.json) then ran
six alternating blocks × two tiny cases × two legs, each with 1,000 samples
after ten warmups, using the same frozen binaries and oracles. All 24 processes
ran once. The independent bootstrap units remain **six block pairs**: 1,000
observations improve each block’s percentile estimate without creating 1,000
independent before/after pairs. Confirmation results:

| Tiny case | Metric | Before | After | Paired ratio | 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| write | p50_ns | 35755.0 | 35810.0 | 1.00222 | [0.99595, 1.00590] |
| write | p95_ns | 37340.0 | 37105.0 | 0.99316 | [0.97754, 1.01604] |
| write | p99_ns | 44025.0 | 43780.0 | 0.99442 | [0.97607, 1.03398] |
| write | rss_kib | 5278.0 | 5346.0 | 1.01182 | [1.00000, 1.02158] |
| read | p50_ns | 7240.0 | 7255.0 | 1.00276 | [0.99587, 1.00764] |
| read | p95_ns | 7455.0 | 7455.0 | 0.99934 | [0.99264, 1.01080] |
| read | p99_ns | 8125.0 | 8080.5 | 0.99761 | [0.89017, 1.26600] |
| read | rss_kib | 4990.0 | 5150.0 | 1.03705 | [0.97663, 1.05401] |

Latency units in this confirmation table are ns; RSS is KiB. No confirmation
p50/p95/p99/RSS flag fired. Original tiny-tail increases were not reproduced.
Read p99 remains uncertain, with upper endpoint 1.266, and read RSS’s 1.03705
point ratio has upper endpoint 1.05401. This supports adoption for the named
larger workloads, not a universal tail or memory improvement claim. No
unresolved reproduced common-workload damage was found.

## Correctness, audit and limits

Twenty-four OPC qualifications independently reconstruct payloads, verify CRC,
logical digests, exact whole archives and Deflate streams, and prove unchanged
local ZIP spans/central metadata apart from offsets in changed publication.
Eighteen auxiliary codec archives cover nine named route outputs per leg,
including a duplicate borrowed/reusable control. All 72 codec members have
explicit local/central Deflate method checks and independent expected payloads.
All timed samples retain expected-payload and exact archive identity checks.

The final evidence contains 192 valid timing reports / 28,344 samples: 168 /
4,344 original qualification/native plus 24 / 24,000 tail confirmation. Two
auxiliary codec processes are separate. One earlier failed qualification and
all build/reader/lock preflight failures are preserved; no formal native or
tail run was replaced. Nineteen reader checks pass. Independent arithmetic
recomputes all 84 native rows / 72 defined intervals and eight confirmation
rows/intervals. Undefined zero-denominator fault ratios remain undefined.

Quality gates: format; workspace libraries/tests check; workspace library
Clippy with warnings denied; all-target checks, warning-denied docs and full
all-feature tests for ZIP, OPC, DOCX, XLSX and PPTX; crate boundaries; and locked
checks/metadata for all three live tools. Affected-owner tests total **6,931
passed, zero failed, 38 ignored**. The before ZIP/OPC gate had 1,650 passed, zero
failed, three ignored. Both final probe builds pass tests and warning-denied
Clippy. Full workspace all-target and rustdoc gates remain blocked by baseline-
confirmed iWork defects: a Numbers chart example return type and two Numbers
wire documentation links. iWork is explicitly excluded; its source is untouched.

Independent review required full lock-record comparisons (including duplicate
package names), tool-check binding, and original-manifest evidence for the
documentation control. Those fixes and the native-resave baseline drift passed
re-review. See [selection/ADR matrix](results/change-0842/selection.md),
[review](results/change-0842/final-review.md), and the packet receipts.

The broad OLE2/OOXML performance goal remains active; ODF is deferred and iWork
excluded. Runtime detection is restored on this measured host, but semantic
Office workflows, cold/storage costs, other architectures, and existing larger
CRUD bottlenecks still need their own evidence.

## Reproduction and retained evidence

[Packet README](results/change-0842/README.md) describes the frozen inputs,
replay commands and original failures. [Decision](results/change-0842/decision.json)
separates artifact validity from adoption. `close_v2.py` verifies both datasets
and invokes the original closure/seal checks. Only marker-owned temporary roots
are removed after evidence verification; the packet retains source/locks,
commands, raw reports, exact qualification archives, readers, decisions and hashes.
