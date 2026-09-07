# Next work after the 0463 validation-proof candidate

The 0463 candidate is retained after its frozen latency gate, final harness
checks and sealed evidence replay all pass.
0462 was rejected on the large-shape gate, with unchanged allocation metrics
and 280 bytes of extra per-shape state. The next batch should move to a
program-level end-to-end requirement rather than another ODP attribute or XML
micro-optimization. The full non-iWork goal remains open.

## Priority 1: make source-backed OPC CRUD a measured semantic lifecycle

The highest broader-ROI gap is adoption of the already-tested source-backed
OPC substrate by one checked, end-to-end semantic workflow. This is a program
priority, not a claim that it is the globally largest latency endpoint. The
ordinary ODP control is a different owned `Snapshot`/`Commit`/`Patch` contract;
the historical 0458 large lifecycle is about 152 ms p50 and its commit phase
is 45.6% of that phase envelope. The PPTX media-rich range lifecycle is roughly 1.77--2.57 s,
but includes a simulated provider and a specialized publication result. Those
numbers cannot be ranked directly.

The best concrete candidate is the bounded PPTX cross-slide copy in
`crates/litchi-pptx/src/presentation/source_cross_copy.rs`: measure
`SourceBackedPresentationEditor::plan_cross_slide_copy` and
`publish_cross_slide_copy_to_stream` (around lines 379--498), with OPC's
`SourceTopologyPlan`/`write_topology_to_stream` as the physical owner. The
existing plan already validates source/destination lineage and revisions,
rechecks the plan before output, preserves untouched ZIP records, and returns
a specialized `SourceBackedCrossSlideCopySnapshot`.

This choice is supported by, but not closed by, 0452/0453. The media-rich
simulated-range p50 falls from 2572.729/2572.933 ms to 1766.853/1774.087 ms
after retained capture, and shared decoded ownership removes 16,777,408
planning bytes. The rows remain correctness-only/generated-fixture evidence;
they do not establish a general PPTX speedup, ordinary Patch equivalence, or
physical network behavior.

The next capture should bind a checked-catalog source and destination, then
measure separate bytes and caller-`ReadAt`/range lanes with a sequential sink:
open, plan, publication, accepted source/destination reads, sink writes,
allocation/peak/RSS and managed reservations. Independent checks must reopen
the output, compare the complete dependency closure, verify untouched member
bytes and source immutability, and exercise stale-source, cancellation, short
write and budget-refusal paths. A compatible distinct-package pair is required
for a successful cross-package claim; self-pairs remain useful controls only.

This result must keep the specialized one-way publication contract. It must not
be reported as an ordinary `Snapshot`/`Commit`/`Patch` or as evidence for lazy
ordinary ODP snapshots. If no compatible distinct pair can be admitted, record
the typed refusal and retain the self-pair as a scoped control; do not bypass
the graph-compatibility boundary.

## Priority 2: close the input/output and producer matrix

The current native evidence is a self-pair preservation inventory: 0454 has
four published self-pair cases and 185 unchanged outcomes, with no Office
application launch. The six distinct native PPTX probes recorded in 0456
refuse incompatible shared layout/master/theme graphs. This leaves two
concrete requirements:

1. Run a pinned producer/application roundtrip with an independently hashed
   output oracle, or record the environment prerequisite that blocks it.
2. Admit and measure a distinct source/destination pair only when the selected
   dependency graph is proven compatible; otherwise preserve the refusal matrix
   as the result.

After the semantic lifecycle is bound, repeat it over cold filesystem input,
caller-provided positional/range input and explicit bounded worker counts.
The 0452/0455 pacing and transfer results are useful service simulations, not
physical cold-I/O or network evidence. Any worker comparison needs an explicit
`ExecutionContext` and must report source reads, output writes, memory limits,
cancellation and scaling efficiency.

## Priority 3: promote correctness-only CRUD rows to measured checked-catalog cases

The coverage index remains 15 categories, 33 representative mappings, 10
measured mappings and 23 correctness-only mappings; the registry remains 439
selectors / 36 defaults. In particular, generated ODP append, PPTX cross-copy,
XLSX structural edits, deletion, merge/split, and join/three-way patch rows do
not close the timing requirement. Candidate follow-on rows include
`xlsx_join_disjoint_commit_save`, `xlsx_three_way_disjoint_commit_save`, and
the existing structural/delete rows, using their source-backed/eager owners
only after a pinned corpus and complete commit/save scope are fixed.

The required evidence for each promoted row is p50/p95/p99 with uncertainty,
allocation and peak/RSS data, source/read/write counters, exact semantic and
raw-member oracles, and explicit no-op, inverse/conflict and failure-atomicity
checks. Correctness tests alone are useful boundaries but do not establish a
performance baseline.

0463 retains a scoped ODP proof-reuse improvement after its frozen gate and
all validation checks pass. The source-backed
semantic lifecycle, native/distinct-package evidence, cold/range input,
explicit scaling, and correctness-only CRUD rows remain open; no full-goal
completion claim follows from the ODP candidate.
