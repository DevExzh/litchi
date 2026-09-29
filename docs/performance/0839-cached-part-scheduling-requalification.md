# 0839 — cached ordered-Part reads avoid private worker creation

**Adopted after a fresh paired trial.** When every requested source-backed OPC
Part is already cached, the private-worker route now performs the ordinary
prepared reads on the calling thread. Existing resource admission, source
checks, cache authority, wave/error ordering, and caller-supplied workers remain
in force. Sixteen eligible cached cases improve paired p50 by **96.68–98.99%**.
No frozen native latency, raw RSS, phase-residency, or operation-allocation
regression gate is triggered.

This is a new decision on current source, not a reversal of 0787's failed
experiment. Its small/primed/width-four RSS rejection remains authoritative for
that capture. The 0788–0790 accounting and launcher investigations informed the
new protocol; no historical samples or RSS corrections enter these results.

## Change and semantic boundaries

Base is `6b33c24b53`. The three affected files still matched the old candidate's
base exactly, so its implementation and seven focused tests were requalified
against the current tree. `PartCache::all_entries_ready` checks completed
entries under the existing cache lock, excluding flights and provisional
publication. The hint neither pins/clones payloads nor updates cache statistics,
LRU state, or reservations. Every requested Part still goes through the
source-checked authoritative prepared-read path; eviction after the hint can
therefore cause an ordinary cold read.

The hint is consulted after normal preparation, output/scheduler reservations,
CPU-task charging, and the parallel-entry fence. An explicit caller worker
facility keeps its callback route. Private cached reads retain wave boundaries,
per-wave fences, panic translation, draining of a failed wave, lowest-ordinal
error selection, and stopping before later waves. Only unnecessary private
thread/channel construction and its internal failure opportunities disappear.

No public API, dependency, global executor, cache-retention policy, compression
policy, publication rule, or durable wire changes. ADR 0005/0031 execution and
budget contracts remain intact; the existing cache remains the sole payload
authority. Source and normative input hashes are retained. iWork is excluded.

## Fresh protocol

The native matrix contains 60 cases: 32 Parts of 4 KiB, 256 KiB, or 31 large
plus one small member; fresh versus primed packages; zero or 64 KiB task floors;
and requested widths 1/2/4/8/32. Input is immutable memory. Fresh is a package
state, not physical filesystem cold. Timed ordered reads exclude corpus/setup,
preloading, byte verification and returned-batch drop. Each process records
30 samples after three warmups; six blocks alternate source-leg and case order.
All CPU affinity, finite execution budgets, toolchain, build flags, and corpus
identities are in the packet.

Separate populations record source work, acknowledged residency, and operation
allocation. Residency uses the ordinary fork launcher from 0790, both all-CPU
and one-CPU affinities, and one-sample/no-warmup plus 30-sample/three-warmup
protocols. Only the first/last measured samples have phase markers. Each phase
uses its maximum observed summed smaps RSS within that child; the separate
maximum-observed metric includes all acknowledged setup/report markers. These
are whole-process point observations, not physical transient peaks. Raw smaps,
rollup, status, maps, stat, PID/start-time/parent identity and direct wait4 usage
remain retained. Formal capture verifies phase order and launcher/executable
identity before each acknowledgement.

A separate allocator executable uses the existing System-allocator observer.
Its region brackets only `read_parts_ordered`, including the retained returned
batch and worker callbacks but excluding setup, preload, verification and drop.
It reports calls, requested bytes, observer-ordered peak above entry and retained
live-byte delta. It measures no latency. These metrics do not represent physical
RSS or allocator-internal transient overlap.

The frozen native gate rejects a paired p50 or raw GNU-time RSS median ratio
above 1.05 when its bootstrap lower endpoint exceeds one. The same independent
rule applies to phase residency and allocation metrics. Zero/nonpositive
allocation baselines use a conservative absolute-increase check. Benefit needs
an eligible cached p50 improvement of at least 3% with upper endpoint below one.
The plan amendment makes reducers explicit before formal paired capture; it
changes neither thresholds nor cases. Ten thousand bootstrap draws use seed
839083. Preflight, observer, traced and instrumented timings are never pooled.

## Results

All 1,414 reports / 28,450 samples validate: 720 native reports, 240 source
observers, 192 residency reports, 108 allocation reports, 120 qualification
reports and 34 separate preflights. Six additional one-sample syscall reports
verify the thread mechanism. Complete individual rows and uncertainty are in
[paired.md](results/change-0839/paired.md),
[paired.csv](results/change-0839/paired.csv), and
[analysis.json](results/change-0839/analysis.json).

Large, primed, zero-floor control; absolute columns are medians of six process
p50 values, while ratios are paired-block medians:

| Requested width | Before p50 µs | After p50 µs | After/before ratio | 95% interval |
| ---: | ---: | ---: | ---: | --- |
| 1 | 2.900 | 3.045 | 1.037538 | [0.992294, 1.129004] |
| 2 | 226.341 | 6.975 | 0.030086 | [0.028024, 0.033974] |
| 4 | 236.746 | 7.530 | 0.031659 | [0.029989, 0.034066] |
| 8 | 297.476 | 8.095 | 0.027236 | [0.026829, 0.028132] |
| 32 | 577.277 | 7.185 | 0.012489 | [0.012097, 0.013035] |

Geometric means of the individually paired p50 ratios, with equal case weights,
are 0.023151 for the sixteen eligible cached cases, 0.999836 for the thirty
fresh cases, and 0.999518 for fourteen cached serial/ineligible controls.
These are named condition groups, not whole-Office workflow speedups. The
largest native p50 ratio is 1.037538 and largest raw RSS ratio is 1.016563;
no native p50 or RSS point estimate exceeds 1.05.

The historically rejecting small/primed/zero-floor/width-four case now has a
raw RSS ratio of **0.991253 [0.969069, 1.017966]**, with absolute medians
4,118 → 4,090 KiB. The maximum phase-residency ratio is 1.019577
[0.986393, 1.038337]. All 2,592 formal acknowledged snapshots have equal summed
smaps, rollup, and status RSS; this does not establish equality with exit
high-water accounting or continuous peak coverage.

On that small cached width-four case, medians of process allocation medians
change as follows:

| Operation metric | Before | After |
| --- | ---: | ---: |
| Allocation calls | 46 | 2 |
| Requested bytes | 9,224 | 1,792 |
| Region peak above entry, bytes | 9,128 | 1,792 |
| Retained live-byte delta | 768 | 768 |

The width-32 small cached case falls from 100 to two calls and from 9,000 to
1,792 requested bytes, with the same 768-byte retained delta. All independent
allocation guards pass. Source observers retain 64 logical reads for fresh
operations and zero for primed operations; CPU-task charges and permit release
match across legs. Fresh width-four traces create four threads on either leg.
Primed width-four creates eight before/four after, and primed width-32 creates
64 before/32 after: preload remains parallel, while the cached measured read
avoids new workers.

## Tail and evidence limitations

Two p95 and three p99 paired point estimates exceed 1.05. Every corresponding
interval includes one. The largest is large/primed/zero-floor/width-one p99:
ratio 1.298694 [0.777163, 1.684886], with absolute medians 3.240 → 4.130 µs.
Other flagged points are large/primed/64-KiB-floor/width-one p99,
mixed/fresh/64-KiB-floor/width-eight p99, and
mixed/primed/64-KiB-floor/width-one p95; the largest case also has a p95 flag.
These controls do not enter the new cached-private-worker branch, but that
alone does not prove tail neutrality or explain variation. The frozen p50/RSS
adoption gates pass; all tail observations remain reviewable, and no general
tail improvement is claimed.

This is category-15 low-level read evidence on three synthetic shapes and one
host. It does not establish a whole-document edit/save improvement, remote or
physical-cold behavior, cross-session contention, producer compatibility,
hardware-counter gains, or general parallel scaling. Compression/inverse work
remains separate, and the broader OLE2/OOXML goal stays active.

## Validation and closure

Fresh OPC all-feature tests pass 988 baseline and 995 candidate tests, with one
ignored on each leg. The seven added tests cover completed/flight/pending cache
state, eviction after the hint, single/multiwave reads, caller facilities,
lowest-ordinal panic handling and source change after the fence. Formatting,
all-target compilation, warning-denied library Clippy, rustdoc, and boundaries
pass. Both standalone native probes pass five tests; both final allocator
observers pass 36 tests. Fourteen reader mutations and four statistical checks
pass. All source legs, executable identities, quality receipts, raw artifacts,
and policy/reader versions are bound in the
[packet](results/change-0839/README.md).

The initial allocator test composition mixed real global-allocation callbacks
with synthetic exact-count tests; its failure is retained. The corrected unit
test binary invokes the wrapper explicitly without installing it globally;
normal observer binaries retain the global wrapper and are independently
qualified. A reader-test driver initially caught the wrong exception class;
that failure and the corrected passing run are also retained. Neither changes
production or measured native code. Cleanup removes only marker-owned roots,
with executable identities verified before removal. Unrelated work is preserved.
