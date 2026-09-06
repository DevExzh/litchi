# 0444: Source-backed OPC Part-addition lifecycle baseline

The previous turn committed 0443. The latest ODP profile still contains
namespace comparison work, but the goal audit explicitly lacks dedicated Part
addition evidence. This batch adds that missing low-level package scenario,
not another speculative ODP predicate optimization.

Exercise SourceBackedPackage open, construction of a one-Part/one-relationship
SourceTopologyPlan, and consuming sequential publication. Use 64/1024/4096
existing Parts with deterministic mixed-entropy 1 KiB decoded payloads, and
one fixed 64 KiB added typed leaf plus a root internal relationship. This is
OPC package topology, not a claim about semantic Office owner creation.

Caller-owned immutable input/source wrapper, prepared added payload and bounded
hashing-discard sink exist before timing and allocation entry. The measured
interval includes catalog opening, plan construction and sequential topology
publication. Publication consumes/drops its catalog before endpoint snapshots.
Input and prepared payload remain alive; no output Vec is retained by the timed
sink. Hash finalization, source counters, complete artifact/semantic/raw-member
oracles and report assembly remain outside the timer. Counts describe observed
source reads, not physical remote requests or decoded-payload materializations.

Export actual source/output fixtures via an explicit harness-only example and
verify them independently with Python zipfile/XML parsing, deterministic payload
regeneration and local/central raw-record comparison. Check exact empty-plan
copy, duplicate and missing-target preflight refusal, stale source, short reads,
partial sink and output limits. Existing OPC topology tests retain broader
ZIP64/signature/security coverage; no new native or broad remote claim follows.

Freeze the final capture/oracle contract after implementation/fixture verification
and before any retained baseline measurements: 12 reports/360 samples, normal
and allocator modes, three sizes, two reversed repeats, CPU 2, one worker,
30 samples/3 warmups, then one stat and one record profile. This is a baseline
addition with no before/after speedup claim. Retain every attempt and review all
>5% repeat flags. Serialize root CPU jobs and bind exact source/build identities.

Accepted ADR tree c950b6c8be822561b498d7bbe87c460873dcbf49 is unchanged from
the prior complete read. No production API, dependencies, unsafe code, executor
or ambient runtime I/O changes. The full non-iWork goal remains active.
