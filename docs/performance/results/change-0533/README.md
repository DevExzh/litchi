# 0533: CFB sector-claim error layout experiment

The previous turn made progress: 0532 established a fresh CFB/OLE2 native,
allocation, constructor-profile and generated-code baseline. This campaign
starts from its committed revision and tests private cold error helpers plus
ordinary `claim_sector` inlining. It preserves checked conversion, lookup,
ownership conflict detection, error text, mutation order and physical
reconciliation. A table of exact success, error and state expectations supplies the independent behavioral oracle.

The frozen protocol compares baseline–candidate–candidate–baseline native
captures across all nine XLS workflows and all three direct CFB shapes. Four
primary XLS p50 values must improve by at least 3% in each repeat. All rows,
mean/tails, whole-child RSS, repeat drift, allocator vectors and constructor
instruction counts remain visible. Every adverse change above 5% requires
review. Instrumented elapsed times never supply native latency evidence.

Each stage retains its exact source manifest, patch, normal and allocator
binary hashes, child commands, raw vectors, corpus identities, runtime oracles,
and accessible compiler-process observations. These observations do not imply
a quiescent host. Profiles use parent-constructor toggles with five timed calls
per child, preserve separate CFB setup dumps, and classify calls through
positive incoming edges. Hardware counters cover the whole child.

The standalone harness has its own Cargo workspace and no explicit release
LTO profile. Root-workspace LTO must not be inferred. Only the final candidate
source and binary can support adoption; a changed rebuild requires a fresh
full native comparison. A rejected production change will be reverted while
independent behavioral tests can remain after restored-source quality checks.

Owned temporary storage is `/home/zhuhe/litchi-goal-0533-target`, with the
`/tmp/litchi-goal-0533` alias pointing to its retained binaries. Builds and
quality checks use owned temporary directories from their first child, and
Callgrind uses `--vgdb=no`, following 0532's confirmed `/tmp` quota failure.
All Rust builds, quality gates and captures run serially under the coordinator.

OLE2 and OOXML remain the optimization priority. ODF is deferred until that
goal completes; iWork is excluded. Synthetic warm in-memory measurements do
not establish cold I/O, remote providers, native-producer breadth, concurrency
scaling, or completion of the full performance program.

## Completed measurement

The full comparison retains 48,000 native samples, 1,440 separate allocation
samples, 80 timed constructor dumps and 12 separate setup dumps. All 14 fresh
quality gates passed, including 4,378 test executions. The four primary XLS
p50 improvements span 20.5382–24.5798%; all 24 paired allocation rows preserve
the four measured vectors exactly. Peak RSS changes range from -0.7800% to
+0.5981%. The complete 158-receipt serial timeline includes the planned native
ABBA order. [Mechanism review](mechanism-review.md) connects the source and
caller assembly to the independently measured constructor instruction totals.

All raw maximum-sample and variability flags remain in the comparison. The
[adverse review](adverse-review.json), [decision](decision.json),
[post-cleanup verification](verification.json) and [seal](SHA256SUMS) record
the final disposition; no aggregate statistic substitutes for those gates.

The [decision](decision.json) is **accepted**. [Cleanup](cleanup.json) removed
both owned temporary paths, and full post-cleanup verification passed. The
retained source is the exact measured candidate, including its two contract
tests; no changed final binary is substituted for the measured executable.
