# ODT envelope-classification package sharing

Date: 2026-08-11

Production base: `1194fbc7f`

Scope: private ODT transaction envelope-classification ownership only. OLE2,
OOXML, RTF, other ODF families, iWork and IWA production code are unchanged.

## Disposition

Accepted. A changed ODT commit now gives envelope classification a shared
handle to the snapshot's already validated immutable archive instead of
allocating and copying the complete package into a temporary owner. On the
fixed media-rich corpus this removes one 16,786,287-byte copy and two
allocation calls per changed commit.

Across two balanced ABBA cycles and 2,000 samples per state, media-rich
paragraph edit/save improves **11.40% at p50**, **11.95% at mean**, **12.19%
at p95**, and **12.55% at p99**. Heaptrack reports exactly two fewer allocation
calls per iteration, with unchanged peak heap and profiler RSS.

Archive validation, package/manifest parsing, encryption and signature
classification, compact-write auditing, final output materialization,
publication, reopen/readback, patch/inverse and stale-source checks all remain.
This changes no public API, dependency, cache, runtime, lock, global state or
unsafe-code boundary.

## Ownership change and retained security boundary

`Snapshot` already owns the exact package as a private `Arc<Vec<u8>>` after
complete ODT validation. `envelope_package` preserves the established package
size check and passes a clone of that handle to
`OwnedPackage::from_shared_bytes`. That constructor still runs the ZIP archive
validator. `package()` then performs the same MIME/manifest work before
classification checks encryption metadata and document/macro signature Parts.

The previous path first allocated and copied the archive into a `Vec`, then
placed that vector into another temporary `Arc` allocation. Neither allocation
was semantic or security state. A focused regression proves pointer identity
between the snapshot and envelope package, then performs a real plain-envelope
classification.

The final changed-result `copy_bytes` call is deliberately retained: it belongs
to transaction result materialization and independent final-document
readback, not envelope inspection. Change 0041 separately removed the three
compact-audit copies without weakening its validation.

## Primary matched latency

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, glibc/system allocator, CPU 11 pinned with `taskset`, and the
identical explicit allocator policy
`MALLOC_MMAP_THRESHOLD_=33554432 MALLOC_TRIM_THRESHOLD_=-1` for both states.
The corpus and allocator policy are unchanged from change 0041.

The fixed corpus has 200 paragraphs and eight deterministic incompressible
2 MiB media members. Its archive is 16,786,287 bytes and has SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The before executable is the frozen change-0041 final binary with SHA-256
`7759ef7aa66ef1328215fbf047e8b778cbe9d174fae67968951ada7b04ae7c8d`.
The final candidate executable SHA-256 is
`671c9dcab035382fe23c8383a6c2cc50019d8204842069cac9ce6b0cfbb4335f`.

An exploratory 20-warmup/200-sample four-leg run was discarded because one
paired comparison disagreed while the first before leg drifted about 13% from
the last. It is not committed or included in any result below. The accepted
record comprises two subsequent before-A, after-A, after-B, before-B cycles,
each with 100 warmups and 500 samples per leg. Pooling all eight legs yields
2,000 raw samples per state.

| Media-rich paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 5.555 ms | 4.921 ms | **-11.40%** |
| mean | 5.596 ms | 4.927 ms | **-11.95%** |
| p95 | 6.187 ms | 5.432 ms | **-12.19%** |
| p99 | 6.804 ms | 5.950 ms | **-12.55%** |

The four paired p50 changes are -11.73%, -12.22%, -14.91%, and -10.47%.
Process drift remains visible inside individual states, but the balanced order
and every independent pair agree on a material improvement. Raw reports are
the four `rerun` and four `final2` JSON files under `results/`; their complete
digests are in
[`odt-envelope-sharing-sha256.txt`](../results/odt-envelope-sharing-sha256.txt).

## Regression guards

The ordinary ODT matrix used ten warmups and 100 samples per leg. Changed
edit/save p50/mean remains neutral: medium improves 1.89%/1.70%, while large
moves +1.01%/+0.27%. Large edited p95 improves 0.25%; medium edited p95 improves
2.84%. The changed branch's ordinary corpus is small, so the removed copy is
correspondingly small.

A dedicated open/no-op run used 100 warmups and 1,000 samples per leg,
yielding 2,000 samples per state:

| Untouched-path guard | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| Medium open | -0.00% | -2.23% | -7.29% | -12.69% |
| Medium exact no-op | -6.98% | -12.51% | -16.10% | -30.11% |
| Large open | +0.04% | +0.76% | +3.46% | +20.31% |
| Large exact no-op | +7.75% | +7.98% | +2.98% | +0.29% |

The changed function is not reached by open or exact no-op. Large-open paired
p50 changes are +0.97% and -1.07%, and paired means are both within 1%; its
pooled p99 reflects one after-B tail rather than p50/mean movement. The large
exact no-op increase is 152 ns at p50 and 165 ns at mean. Adding open and no-op
times yields a complete warm open-plus-no-op p50 movement of about +0.06%.
This code-layout/microtimer exception is accepted and explicitly not presented
as a no-op improvement.

Raw evidence is in the four `guards` and four `read-guard` JSON files under
`results/`.

## Profile, counters and memory

Matched whole-process `perf stat` ABBA runs used ten warmups and 100 primary
samples per leg. The process includes corpus creation, complete verification,
patch/inverse checks and report construction outside the timed interval.

| Counter, A+B | Before | After | Delta |
|---|---:|---:|---:|
| Task clock | 9,672.32 ms | 9,434.27 ms | -2.46% |
| Context switches | 116 | 109 | -6.03% |
| CPU migrations | 0 | 0 | unchanged |
| Page faults | 68,486 | 68,477 | -0.01% |
| Cycles | 47,121,136,001 | 46,324,768,497 | -1.69% |
| Instructions | 97,635,160,364 | 97,439,219,095 | -0.20% |
| Branches | 10,106,341,370 | 10,089,575,259 | -0.17% |
| Branch misses | 32,876,261 | 32,778,498 | -0.30% |
| Cache references | 4,436,568,581 | 4,141,907,957 | -6.64% |
| Cache misses | 268,442,765 | 237,209,907 | **-11.63%** |

Sampled profiling over three warmups and 30 primary samples removes the
resolved `envelope_kind` memmove caller, which held 1.44% exclusive share in
the before graph. The remaining total memmove share moves from 13.81% to
15.00% as independent result materialization, verification and raw ZIP
preservation dominate the sampled denominator. Approximate event count falls
0.58%; kernel symbols were restricted and no samples were lost.

Heaptrack over two warmups and 20 primary samples reports 274,619/274,575
allocation calls: **44 fewer, exactly two per changed commit**. Temporary
allocations remain 48,789, peak heap remains 106.03 MiB, profiler RSS remains
120.60 MiB, and retained leak accounting remains 1.78 KiB.

Uninstrumented GNU Time ABBA maximum RSS averages 112,448/112,576 KiB
(+0.11%, flat). Mean user time falls 3.36% and wall time falls 3.21%. One major
fault occurred in before-A; the other three legs had zero, and minor faults are
flat.

Raw profile, counter, Heaptrack and GNU Time summaries are under `results/`
with the `odt-envelope-sharing` prefix.

## Preservation, safety and verification

The change retains exact no-op bytes and sharing, deterministic operation
order, source-checked patch and inverse semantics, stale-source refusal, ODF
package/resource limits, manifest and XML validation, compact-write audit,
signed/encrypted envelope policy, raw unchanged-member publication, final
document reopen and semantic/media readback.

Final gates:

- complete all-feature `litchi-odt` and `litchi-odf-common` suites, including
  integration and doctests;
- warning-denied all-target/all-feature Clippy and warning-denied rustdoc for
  both crates;
- all benchmark-harness tests and warning-denied all-target Clippy;
- focused envelope/package pointer identity plus existing signed/encrypted
  refusal and packaged-transaction coverage;
- JSON syntax, exact-file formatting, evidence digests and repository diff
  hygiene.

No ODT fuzz manifest exists. No test, validation, limit, security refusal or
CI threshold was weakened or removed.

## Next bounded work

1. Remove ordinary RTF decoded-body double copies while retaining the existing
   arena path for insertion/deletion revision text.
2. Build the source-backed PPTX owning one-slide editor and media-rich control
   before attaching the existing OPC overlay publisher.
3. Add native XLS owner-stage attribution before another OLE2 production
   change.
4. Continue broader ODF source-backed reads, repeated ODT/ODP scans,
   resource-adding/structural publication and real-producer media coverage.

iWork/IWA remains explicitly deferred while other agents modify `iwa-*`.
