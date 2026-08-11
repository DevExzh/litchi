# ODT compact-audit package sharing

Date: 2026-08-11

Production base: `f6db7e2c59afd4a4fb2647bb5367e86d16ec8ee0`

Scope: private ODT transaction compact-audit ownership plus ancillary
warning cleanup in the shared ODF encryption and datatype documentation.
OLE2, OOXML, RTF, iWork and IWA production code are unchanged.

## Disposition

Accepted. Each changed ODT operation now gives the compactness audit shared
references to the already validated predecessor and candidate packages. The
previous path copied the complete predecessor once before dispatch, copied it
again while constructing the audit package, and copied the complete candidate
while constructing the second audit package. On the fixed 16,786,287-byte
archive this removes 50,358,861 bytes of transient package copying per
operation.

The media-rich paragraph edit/save improves **30.44% at p50**, **31.36% at
mean**, and **32.41% at p95**. Allocation calls fall 0.57% across the complete
instrumented process; peak heap and RSS remain flat. Compact XML/splice
validation itself, archive and manifest parsing, envelope classification,
final result materialization, complete reopen/readback, patch/inverse and
stale-source behavior all remain.

This changes no public API, dependency, cache, runtime, lock, global state or
unsafe-code boundary.

## Ownership change and retained checks

`Edit::commit` already owns a validated private `OwnedPackage` for the current
transaction state. Before each operation it now clones that package's private
immutable `Arc` instead of allocating a byte vector. After dispatch, the audit
accepts the predecessor and candidate packages directly rather than rebuilding
both from fresh archive-sized copies.

The audit still opens both packages, enumerates candidate members, reads the
manifest media types, checks changed XML for compactness and validates the
eligible single-splice relation. Sharing cannot bypass package construction:
both inputs reached the audit only after their existing complete validation.
A focused pointer-identity regression proves the predecessor clone shares the
same immutable allocation before and after a real changed-candidate audit.

The final output copy, result snapshot/document reopen and the separate
envelope-classification copy are intentionally retained. Removing them would
cross distinct publication, independent-readback or security boundaries and
requires separate attribution.

## Primary matched latency

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, glibc/system allocator, CPU 11 pinned with `taskset`, and the
identical explicit allocator policy
`MALLOC_MMAP_THRESHOLD_=33554432 MALLOC_TRIM_THRESHOLD_=-1` for both states.
The policy keeps repeated 16--17 MiB benchmark allocations in process-local
heap reuse after warmup. Default-policy exploratory runs were discarded:
their medians shifted between about 5 and 28 ms as glibc changed large-allocation
and page-reuse behavior on a host with fully occupied swap, even though no
active swap I/O was observed. No result from those unstable runs is reported
as evidence.

The fixed corpus has 200 paragraphs and eight deterministic incompressible
2 MiB media members. Its archive is 16,786,287 bytes and has SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The before executable SHA-256 is
`ba1b6041e7ca0ebf721700c25bad90884cb6bf2ada665989e004b4d0bf7168d9`;
the final after executable SHA-256 is
`7759ef7aa66ef1328215fbf047e8b778cbe9d174fae67968951ada7b04ae7c8d`.

The final ABBA run used 20 warmups and 200 samples per leg in before-A,
after-A, after-B, before-B order, yielding 400 raw samples per state.

| Media-rich paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 7.773 ms | 5.407 ms | **-30.44%** |
| mean | 7.879 ms | 5.408 ms | **-31.36%** |
| p95 | 8.815 ms | 5.958 ms | **-32.41%** |
| p99 | 9.363 ms | 6.569 ms | -29.84% |

Individual p50s are 8.210/5.265/5.477/7.473 ms in ABBA order, so both paired
before/after comparisons materially improve. Raw reports are the four
[`final primary`](../results/abba-odt-compact-audit-final-before-a.json) ABBA
JSON files; their complete digests are in
[`odt-compact-audit-final-sha256.txt`](../results/odt-compact-audit-final-sha256.txt).

## Regression guards

A four-leg ordinary ODT matrix used ten warmups and 100 samples per leg:

| Guard | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| Medium open | -1.24% | -1.63% | -4.10% | -7.52% |
| Medium one edit/save | +0.76% | +0.64% | +7.18% | -5.68% |
| Large open | -1.69% | -2.59% | -6.31% | -14.51% |
| Large exact no-op | -16.96% | -13.76% | -17.36% | -7.02% |

The mixed matrix's medium one-edit p95 is noisy while its p50/mean are neutral.
A dedicated large one-edit guard with 50 warmups and 500 samples per leg
(1,000/state) resolves to **-2.26% p50 / -2.09% mean / -2.35% p95**; p99 is
+1.21%.

A dedicated medium exact-no-op run with 200 warmups and 5,000 samples per leg
(10,000/state) moves from 261 to 300 ns at p50: **+39 ns / +14.94%**. Mean is
+26.85 ns/+9.46%, p95 +20 ns/+5.26%, and p99 +39 ns/+8.65%. The measured
no-op returns before the changed audit path executes, and its byte identity and
source sharing are unchanged. This sub-microsecond code-layout/timer movement
is accepted and explicitly not presented as a no-op improvement.

Raw evidence is in the four `final-guards`, four `final-edit-guard` and four
`final-noop-guard` JSON files under `results/`.

## Profile, counters and memory

Matched whole-process `perf stat` ABBA runs used ten warmups and 100 primary
samples per leg. The process includes deterministic corpus creation, complete
verification, patch/inverse checks and report construction outside the timed
interval.

| Counter, A+B | Before | After | Delta |
|---|---:|---:|---:|
| Task clock | 10,234.570 ms | 9,636.380 ms | -5.85% |
| Cycles | 50,276,470,352 | 47,207,058,428 | -6.11% |
| Instructions | 98,173,203,208 | 97,596,497,979 | -0.59% |
| Branches | 10,153,407,015 | 10,102,707,457 | -0.50% |
| Branch misses | 33,119,343 | 32,858,284 | -0.79% |
| Cache references | 5,220,529,443 | 4,421,795,138 | -15.30% |
| Cache misses | 358,221,458 | 280,049,968 | **-21.82%** |
| Page faults | 68,291 | 68,472 | +0.27% (flat) |

No CPU migrations or major page faults occurred. Sampled profiling over three
warmups and 30 samples lowers the approximate event count from 8.509 to 8.174
billion. `memmove` falls from 18.48% to 14.96% exclusive share (19.05% lower
relative); the before graph attributes 2.72% to compact-audit construction,
which is absent from the final path. Kernel symbols were restricted and no
samples were lost.

Heaptrack over two warmups and 20 primary samples reports 276,182/274,620
allocation calls (**-1,562 / -0.57%**), identical 48,790 temporary allocations,
identical 106.03 MiB peak heap, and 120.29/120.71 MiB Heaptrack RSS (+0.35%,
flat). Uninstrumented GNU Time ABBA maximum RSS averages 112,192/112,512 KiB
(+0.29%, flat); mean user time falls 5.27% and wall time falls 5.20%.

Raw profile, counter, Heaptrack and GNU Time summaries are under `results/`
with the same `odt-compact-audit-final` prefix.

## Deprecation and documentation warnings

The shared ODF Blowfish CFB8 encrypt/decrypt loops no longer call the
deprecated generic-array `clone_from_slice` constructor. Each loop now creates
the fixed-size block with `Default` and copies the eight-byte feedback value
into it. This is a warning-only compatibility cleanup with no performance
claim; the complete ODF encryption vectors and all-feature suite pass.

Two public datatype rustdoc links that targeted private modules are now plain
code-formatted module names, making warning-denied ODF-common rustdoc clean.
Neither ancillary cleanup changes a public API or serialized behavior.

## Preservation, safety and verification

The change retains exact no-op bytes, deterministic operation order,
source-checked patch and inverse semantics, stale-source refusal, complete ODF
package/resource limits, manifest and XML validation, compact-write audit,
signed/encrypted envelope policy, raw unchanged-member publication, final
document reopen and semantic/media readback.

Final gates:

- complete all-feature `litchi-odt` and `litchi-odf-common` suites, including
  integration and doctests;
- warning-denied all-target/all-feature Clippy and warning-denied rustdoc for
  both crates;
- all 29 benchmark-harness tests and its warning-denied all-target Clippy;
- focused immutable package pointer-identity and Blowfish CFB8 vectors;
- JSON syntax, exact-file formatting and repository diff hygiene.

No ODT fuzz manifest exists. No test, validation, limit, security refusal or
CI threshold was weakened or removed.

## Next bounded work

1. Build a source-backed PPTX owning editor and a media-rich control before a
   one-slide overlay publisher; keep consuming publication out of the
   cloneable read facade.
2. Attribute native XLS commit stages before another OLE2 code change.
3. Continue broader ODF source-backed reads, repeated ODT/ODP semantic scans,
   resource-adding/structural publication and real-producer media coverage.

iWork/IWA remains explicitly deferred while other agents modify `iwa-*`.
