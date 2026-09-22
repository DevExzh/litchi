# 0733 — PPT finish instruction partition and payload handoff candidate

The retained profiles identify a concrete copy to test next: PPT finish passes
borrowed stream bytes to `OleWriter::create_stream`, which allocates and copies
them even though finish owns the buffers and CFB already offers
`create_stream_owned`. On the qualified `45543.ppt` edit, the five output stream
payloads total 384,906 bytes. This is a source-derived copy volume, not a newly
measured allocation reduction or speedup. Production is unchanged in this batch.

## Evidence and scope

This is a retrospective analysis of the three sealed 0731 Callgrind profiles,
informed by 0732's native phase observations. It does not run a new PPT timing
matrix. Both ancestor packets pass their exact artifact censuses and hashes.
All 7,206 current workspace Rust/manifest files match the 0732 build census; 7,204
also match 0731. The two differences are the diagnostic feature declaration
and slide-order diagnostic method/tests. Archived source proves ordinary commit
is byte-identical, and the finish/CFB implementation files are unchanged. This
source bridge supports inspection of the retained implementation; it does not
claim identical compiler output across the two builds.

0732 measures embedded finish at 16.73–16.87% of the observed native whole,
with three clock-observer median flags. 0731 masks native SHA acceleration.
Neither fact permits multiplying an instruction fraction below by the native
phase fraction to predict ordinary latency savings. No whole-workflow native
ranking is derived from Callgrind's software-SHA total.

## Conserved instruction partitions

For every selected node, the parser requires exactly one positive-cost caller
and proves incoming inclusive cost equals self cost plus outgoing edge costs.
It also reproduces all 0731 self and edge costs, the owner total, and the file
summary. Shared descendants such as generic stream readers and `memcpy` are
not assigned globally to this subtree. The table uses direct edges with named
denominators; nested rows must not be summed across levels.

| Parent | Direct child | Collected Ir | Fraction of parent |
| --- | --- | ---: | ---: |
| Embedded finish | Package writing | 1,970,376 | 63.8158–63.8167% |
| Embedded finish | Rewrite validation | 769,284 | 24.9153–24.9156% |
| Embedded finish | Checked appends | 312,235 | 10.1125–10.1127% |
| Package writing | CFB write | 1,434,432 | 72.7999% |
| Package writing | Copying stream ingress | 383,349 | 19.4556% |
| Package writing | Source-layout adoption | 147,056 | 7.4633% |
| CFB write | Reuse plan validation | 731,935 | 51.0261% |
| CFB write | Reuse plan emission | 520,891 | 36.3134% |
| CFB write | Sector-layout planning | 180,562 | 12.5877% |
| Rewrite validation | Read document/current-user streams | 671,509 | 87.2901% |

Embedded finish totals 3,087,556–3,087,601 collected instructions across the
three profiles. Stream ingress accounts for 12.4158–12.4159% of that total;
377,604 of its 383,349 instructions are on its direct `memcpy` edge. The
remaining costs and self instructions are retained in machine-readable
partitions. These are observed instruction counts under Valgrind, not native
cycles, time, allocation counts, or available savings.

## Call-count interpretation control

The raw `calls=` fields report three calls on several finish edges, despite
one collected public owner. A new standalone safe-Rust witness calls `work`
once before, once inside, and once after the collection owner. Across three
fixed Callgrind runs, the owner-to-work edge reports one call and 5,009 Ir;
work-to-leaf reports three calls but only 5,008 Ir. The entire collected profile
is 5,010 Ir. Native and instrumented executions return the same three values.

Thus, in this collection setup, call metadata includes uncollected invocations
while Ir is restricted to the owner window. It cannot establish three finishes
inside one measured PPT lifecycle, nor can it be used to divide collected Ir
into a per-call cost. The source shows one ordinary finish in this workflow.
The witness establishes this installed Callgrind 3.26.0 behavior; it makes no
PPT performance claim.

## Next optimization pilot

Move the selected stream buffers into the existing owned CFB writer ingress
while consuming the private PPT editor. Preserve source layout adoption before
staging, insertion order, the selected appended document/current-user bytes,
sector-layout policy, checked lengths, finite output bounds, typed refusals,
Reuse plan validation, final rewrite validation, public reopen, payload checks,
and both artifact digests. Do not retain a new cache or borrow payloads from a
shorter-lived owner. The source review must account for destruction boundaries
and peak memory, because moving buffers changes where their lifetimes end.

Before retaining the candidate:

1. Capture fresh ordinary-path before/after native and allocation results with
   the same binary configuration, fixed interleaved order, all samples and
   tails, and the exact sealed output/preservation oracle. Do not compare a new
   candidate directly with 0731 timing or estimate benefit from this partition.
2. Qualify another representative PPT fixture before claiming broader corpus
   benefit. Record unsupported/refused cases explicitly rather than weakening
   their edit gates.
3. Test changed and unchanged finish, Reuse/Rewrite behavior, output limits,
   source immutability, patch/inverse, and typed failure parity. Retain full
   affected-owner checks and boundary checks for the final source.
4. Require an explicit latency/memory decision: flag >5% latency or peak
   regressions, and show whether removed copies yield a practical allocation
   or end-to-end benefit. Revert complexity that does not earn its cost.

The larger validation costs are required work with different obligations:
Reuse validates the planned CFB layout before touching the sink; PPT rewrite
validation checks the emitted package's live persist mapping; the outer commit
checks semantic publication and unrelated streams. Their existence alone is
not evidence that one can be removed.

## Verification

The actual analyzer accepts the complete evidence and rejects ten corrupted
parser, context-attribution, ancestor, and witness controls. The independent
review checks the source path, raw partition arithmetic, and interpretation.
The only compiled code in this batch is the standalone counter witness;
production Rust gates are not rerun for unchanged source. Prior quality logs
remain sealed and are identified by the source bridge, not presented as fresh
runs. Owned witness binaries are removed after identity verification.

[Evidence and replay](results/change-0733/README.md). The non-iWork performance
goal remains active; this batch nominates a measured ownership pilot and makes
no optimization claim.
