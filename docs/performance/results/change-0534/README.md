# 0534: CFB physical role/FAT reconciliation

The previous turn made progress by accepting 0533's checked claim inlining.
This campaign starts from that committed source and tests a separate physical
reconciliation loop. It pairs the role map with the FAT's common prefix,
retains every unclaimed-marker check in order, and reports a short FAT only
after all earlier prefix checks succeed. Extra FAT padding remains ignored.
No validation, mutation, ownership, allocation or public API is removed.

The preceding measured XLS-owned constructor spends about 17.60% of exclusive
instructions in physical reconciliation, and CFB few-large about 19.31%.
The sealed prior assembly shows twelve 275-byte variants with a FAT-bound
comparison inside the per-sector loop. These observations justify a fresh
comparison; they do not establish a speedup for the proposed loop.

The frozen matrix contains all nine XLS and all three CFB scenarios. Native
capture uses 20 warmups and 1,000 samples per case, with two paired repeats in
baseline–candidate–candidate–baseline order. Each of four primary XLS p50
values must improve by at least 3% in both repeats. XLS-owned constructor Ir
and physical-reconciliation exclusive Ir for XLS-owned and CFB few-large must
also decrease in both repeats. Every adverse matched or absolute same-build
variation above 5% is retained for individual review.

The allocator executable supplies separate operation-scoped vectors with
three warmups and 30 samples per case, two repeats per stage. Its elapsed
times never enter native latency evidence. Each stage has eight constructor
profile children, 40 timed dumps and six separate CFB setup dumps; positive
incoming edges distinguish setup from timed calls. Hardware counters cover
the whole child, including fixture and correctness work, and do not establish
operation-local cycles or IPC.

Every captured child binds the frozen plan, source manifest, script, binary,
command, raw artifacts, corpus identity and runtime oracles. Compiler-process
observations do not establish a quiescent host. The standalone harness has
its own Cargo workspace and no explicit release LTO profile. The candidate
must be measured with its final runtime and tests; a changed final build
requires a fresh full ABBA. Failed admission reverts production while useful
independent contract tests may remain after restored-source quality gates.

The coordinator runs all Rust builds, quality gates and captures serially.
Owned temporary storage is `/home/zhuhe/litchi-goal-0534-target`; the
`/tmp/litchi-goal-0534` alias points to retained stage binaries. Owned `TMPDIR`
and `--vgdb=no` are configured from the first relevant child, following the
previous confirmed `/tmp` quota failure.

OLE2 and OOXML remain active. ODF is deferred until that optimization goal
completes; iWork is excluded. These warm synthetic in-memory measurements
cannot complete cold/range, real-producer breadth, concurrency scaling, broad
CRUD coverage or the full performance program. The rejected visited-bit
fusion and freshness-session proposals remain separate and rejected.

## Result: rejected runtime candidate

Every primary native p50 is slower in both paired repeats:

| Primary workflow | Repeat 1 change | Repeat 2 change |
| --- | ---: | ---: |
| XLS source-backed open | +10.6386% | +1.0596% |
| XLS source-backed open and one cell | +8.3510% | +2.5009% |
| XLS owned-source open | +7.1373% | +4.9282% |
| XLS owned-source open and one cell | +6.6620% | +6.0609% |

All eight rows fail the frozen requirement for at least 3% lower p50.
The runtime patch has been reverted; only the two independent physical-layout
contract tests remain. No runtime speedup is retained or claimed.

The instruction hypothesis partially holds: XLS-owned constructor Ir falls
2.8968%/2.9363%, and physical-reconciliation self Ir falls about 16.6648% in
both selected workloads and repeats. All twelve physical symbols shrink
from 275 to 220 bytes, with stack reservations decreasing from 80 to 48 bytes.
The candidate executable nevertheless grows by 2,448 bytes. These observations
do not explain the native slowdown or override admission.

The allocation-call, reallocation-call, allocated-byte and incremental-region-
peak vectors are identical in all 24 paired rows. The raw comparison retains
86 matched adverse flags and 67 same-build variation flags for individual
review. Direct CFB few-large p50 improves 3.2095%/2.9344%; that guard workload
cannot substitute for the failed primary workflows.

Both candidate and restored final source passed all 14 quality gates and
4,382 test executions per source state (8,764 executions total). All 153 variation flags were individually reviewed. Cleanup and full
post-cleanup verification passed, covering 118 successful serial receipts;
the evidence inventory is sealed. The final source manifest is
`9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75`.
