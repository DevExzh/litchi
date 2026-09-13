# 0546 OLE2 next opportunity

The highest-impact unresolved measured OLE2 owner is the successful chain walk in
`SectorChainScratch::collect_exact`, reached from
`OleFile::validate_stream_allocations`. The next bounded step should be a fresh
sub-operation attribution of this loop on the restored baseline. It should not
start with another code candidate: the prior candidates changed layout without
showing that a required operation was removed. OLE2 and OOXML remain active;
ODF is deferred until that optimization goal is complete and iWork is out of
scope.

This review uses the current [performance hotspot record](../../HOTSPOTS.md),
[performance report](../../REPORT.md), the [0536 decision](../change-0536/decision.json),
the [0536 profile comparison](../change-0536/profile-comparison.json), the
[0536 instruction comparison](../change-0536/instruction-analysis-comparison.json),
the retained [baseline assembly](../change-0536/baseline/assembly-2.stdout),
the [candidate assembly](../change-0536/candidate/assembly-2.stdout), and the
[0535 collector attribution review](../change-0535/collector-attribution-review.md).
All profile values below are sums over five positive timed dumps per profile;
CFB setup dumps are excluded.

## Measured priority and attribution boundaries

`collect_exact` is the largest named exclusive owner in the two large OLE2
shapes. Each percentage below uses that workload's own constructor-inclusive
Ir as its denominator; rows from different workloads are not combined.

| Timed workflow | Constructor inclusive Ir | `collect_exact` self Ir | Self / constructor | Collector inclusive Ir |
| --- | ---: | ---: | ---: | ---: |
| XLS-owned, repeat 1 | 11,316,236 | 5,601,140 | 49.4965% | 5,816,614 |
| XLS-owned, repeat 2 | 11,318,721 | 5,601,140 | 49.4856% | 5,816,500 |
| CFB few-large, repeat 1 | 10,264,094 | 5,571,945 | 54.2858% | 5,680,135 |
| CFB few-large, repeat 2 | 10,264,094 | 5,571,945 | 54.2858% | 5,680,135 |
| CFB many-small, either repeat | 13,979,277 | 764,485 | 5.4687% | 830,050 |
| CFB tiny, either repeat | 244,607 | 5,200 | 2.1259% | 6,240 |

The selected XLS constructor is
`SourceBackedWorkbook::from_read_at_with_limits`, reached through its positive
`from_read_at` edge. The CFB constructor is `OleFile::open`, reached through
the positive benchmark-runner edge. The XLS collector direct Ir is 215,474 in
repeat 1 and 215,360 in repeat 2; its inclusive Ir is therefore 5,816,614 and
5,816,500. CFB few-large direct Ir is 108,190 and inclusive Ir is 5,680,135.
These are separate self, direct-child, and inclusive boundaries, not additive
work totals. The aggregate profile metadata reports 2,730 collector edge calls,
but collection-off call metadata must not be treated as an operation-local
dynamic call count.

The ordinary valid loop accounts for 5,597,360 / 5,601,140 = 99.9325% of XLS
collector self Ir and 5,570,400 / 5,571,945 = 99.9723% of CFB few-large
collector self Ir in the baseline range analysis. These are separate workload
attributions. They establish where the cost is, not which instruction can be
removed or what native latency will change.

## What prior work closed

The 0190 reusable scratch change already removed the repeated per-stream
temporary allocations. Its retained allocation evidence shows a large drop in
allocation calls and temporary allocations, while the current unresolved cost
is the successful validation loop itself. A new allocation claim needs fresh
operation-local evidence rather than the old aggregate allocation result.

The 0524 visited-bit fusion reduced collector self Ir but produced only about a
1.17–1.19% constructor reduction and failed the native gate; it remains closed.
The 0533 claim-sector cold-error and inline-success change is already present in
the current source, so `claim_sector` is not the next target. The 0534 paired
physical-role/FAT-prefix loop reduced physical-reconciliation self Ir but made
all eight primary XLS p50 rows slower by 1.0596–10.6386%; it must not be
revived under the same shape.

The 0536 cold-error-layout candidate moved eight invalid-input error paths into
helpers and changed the valid loop layout. It was rejected: only four of eight
primary XLS p50 rows met the 3% threshold, XLS constructor inclusive Ir rose
from 11,316,236 to 11,316,310 in repeat 1 and from 11,318,721 to 11,321,634
in repeat 2, and the candidate was not retained. On the required large lanes,
XLS collector self Ir changed by only -260 (-0.004642%) and CFB few-large by
-100 (-0.001795%) in each repeat. The helpers had no valid-input profile Ir.
Those results close another cold-helper extraction; they do not explain the
remaining valid-loop work.

## Current source boundary and required invariants

The relevant call sites are in
[`validate_stream_allocations`](../../../../crates/litchi-cfb/src/file.rs#L1079)
and the target is
[`SectorChainScratch::collect_exact`](../../../../crates/litchi-cfb/src/file.rs#L2799).
The root mini-stream uses the separate owned-result
`collect_sector_chain_exact` helper. The target is the reusable `mini_scratch`
and `regular_scratch` path: MiniFAT streams call `mini_scratch.collect_exact`,
regular streams call `regular_scratch.collect_exact`, and the completed result
is then checked and claimed in its respective namespace. Mini-sector ownership
uses a separate `CheckedBitSet`; regular chains use `claim_chain`. Collection
must finish before ownership publication, and MiniFAT and FAT scratch state
must remain separate.

Inside `collect_exact`, the current source performs, in order:

- reset of the sector vector and visited logical length;
- empty-chain, start-marker, and allocation-table-length checks;
- exact fallible reservation of the sector vector;
- visited-map capacity preparation and zero fill;
- checked index conversion and bounds validation for every sector;
- visited `contains` and checked `insert` for cycle detection;
- exact-order sector-vector append;
- FAT or MiniFAT lookup and final/intermediate marker checks; and
- reset after every error before returning it.

The same error precedence, fallible resource labels, allocation boundaries,
reset behavior, chain order, physical bounds, and post-collection claims are
part of the measurement contract. The later
[`validate_physical_sector_layout`](../../../../crates/litchi-cfb/src/file.rs#L1178)
pass remains a required owner and is a separate investigation.

## What the mapped code shows

The 0536 baseline collector symbol is 1,436 bytes with 313 mapped
instructions; the candidate is 1,238 bytes with 277 mapped instructions. The
baseline valid loop occupies approximately `0x2f29480..0x2f2952e`, while the
candidate loop occupies approximately `0x2f29530..0x2f295ba`. Both retain the
checked visited-word test and inline word update in the caller. The separate
`CheckedBitSet::insert` body is 223 bytes, and no positive
`collect_exact -> CheckedBitSet::insert` edge appears in the timed dumps.
The sibling insertion edges observed under the parent are
`claimed_mini_sectors.insert` after collection, not `self.visited.insert` in
the target loop.

The candidate removed two baseline loop-carried stack stores: the current
sector store at `0x2f294b8` and the next-marker store at `0x2f29500`, with the
next sector carried in a register. It retained a stack spill at `0x2f29573`
for cold/error state and changed the table-pointer and visited-address
recomputation layout. The corresponding candidate loop still has a direct
table load at `0x2f295a2`. The mapped change is therefore a register/control
layout change with one remaining spill, not evidence that a required check,
lookup, bitset update, or append disappeared. The almost unchanged positive
self Ir confirms that a smaller text body did not yield a material measured
work reduction. No static instruction-latency claim follows from these
addresses or byte counts.

The visible direct children in the earlier attribution were `memset` and
`finish_grow`. The `memset` is consistent with `visited.words.fill(0)` in
[`prepare_visited`](../../../../crates/litchi-cfb/src/file.rs#L2783); the
`finish_grow` edge is a lower-level fallible-growth path and does not identify
which vector reserved it. Collection-off call labels can retain setup or other
context, so neither edge count is an operation-local call or allocation count.
`CheckedBitSet::contains` is inline, while the absent collector edge to
`CheckedBitSet::insert` is compatible with the successful update being inlined
or folded into the caller. A stronger claim requires the bound binary and
source mapping together.

## Minimal next measurement

Run one read-only, baseline-only attribution pass against the restored current
binary before designing a candidate. The smallest useful first pass covers the
`xls_owned_source_open` or equivalent source-backed constructor lane and the
CFB few-large open; CFB many-small and tiny remain shape guards. Keep positive
timed constructor ancestry, separate CFB setup from timed parts, and retain the
five timed dumps per profile. A single repeat is sufficient for this diagnostic
ranking. Any candidate that survives the ranking must use the established two
repeat ABBA protocol before an admission decision.

Bind the source manifest, binary receipt, plan, and host record, then map the
exact `collect_exact` symbol in the actual baseline assembly. Attribute these
non-overlapping regions separately:

1. entry reset, exact reservation, visited preparation, and `memset`/growth;
2. checked index conversion, visited-word load, test, and update;
3. sector-vector append and its capacity path;
4. FAT/MiniFAT table load, marker checks, and loop-carried state; and
5. cold error tails and failure reset.

For each region retain collector self Ir, visible direct-child Ir, and collector
inclusive Ir. Also retain the constructor inclusive denominator and the
positive incoming owner edge. Do not infer per-stream calls, allocation counts,
or operation-local work from Callgrind call labels. The purpose is to locate a
single repeated operation with a source-preserving elimination or a proven
cheaper equivalent. If no region has such evidence, record that result and
close the collector as an attribution-only hotspot rather than staging a
speculative rewrite.

If a later candidate is justified, it must preserve all source invariants and
pass the full existing admission boundary: all four primary XLS p50 rows must
improve by at least 3% in both repeats; XLS constructor inclusive Ir and the
XLS/CFB few-large collector self rows must move in the required directions;
allocation calls, bytes, and peak memory must show no material growth;
correctness, malformed-input, error-order, reset, physical-reconciliation, and
quality checks must pass; and the actual candidate and restored binaries must
be inspected. Ir reduction alone is not a speedup claim.

This audit made no Rust edits, builds, tests, captures, production changes,
existing-evidence rewrites, or commits.

## Design note: terminal proof and deferred cycle replay

A bounded work-elimination hypothesis is worth recording before any new
attribution: for an immutable, deterministic allocation table, let the visited
sectors be s0 .. s(N-1). If two positions repeat, si == sj for i < j, the
table gives the same successor from both positions. The suffix from si is
therefore periodic and cannot reach ENDOFCHAIN at the first later terminal
step. Consequently, after all N sector indexes and table entries have passed
their existing checks, observing ENDOFCHAIN from s(N-1) proves that no sector
repeated. That proof applies only to the exact-N successful path. It does not
remove the start-marker, table-length, index, bounds, marker, or terminal
checks, and it depends on the table remaining unchanged during the call.

This makes a distinct design possible in principle: retain the already exact
reserved sectors vector, walk the successful path without preparing or touching
visited, and use the terminal proof to return success. If a walk would return a
nonterminal or invalid failure, inspect the recorded prefix in order before
returning the failure. At minimum, a final next != ENDOFCHAIN must be replayed:
a duplicate would have produced the current cycle error before the current code
reached its late chain-length error. A replay that finds the first duplicate
must return the same sector-specific cycle error; otherwise it must return the
original early, invalid-marker, invalid-index, or late-length error at the
original boundary. The safe design question is whether every malformed exit
should replay, or whether a deterministic proof can narrow replay to the final
nonterminal case without changing error precedence. This is a proof
obligation, not an implementation result.

Several current contracts block treating this hypothesis as a free deletion:

- The [current collector](../../../../crates/litchi-cfb/src/file.rs#L2799)
  reserves sectors first, then prepare_visited reserves the missing words under
  the "sector-chain map" resource and fills them before the first chain step.
  Deferring or removing that map allocation changes which fallible allocation
  can fail and when it fails. Allocating a map only on a malformed path could
  mask the original parse error with a late allocation error; removing it
  entirely would change allocation vectors and the established resource
  boundary. The existing source comment explicitly preserves this order.
- reset clears the sector length and visited logical length after every
  failure, while retaining capacities for reuse. reset_visited does not clear
  the backing words. If the map remains allocated but zeroing is omitted, stale
  bits from a prior chain can report a false cycle. Avoiding that requires a
  generation or touched-word scheme, which adds state and its own reset and
  allocation proof. Replaying the sector vector avoids that map, but incurs a
  potentially quadratic scan on a hostile long chain that fails near N.
- The current cycle check occurs after checked index conversion and bounds
  validation but before the vector append, table lookup, and early/late marker
  checks. A fast path must preserve that order on every malformed result, not
  only the common final-terminal case. The vector must be replayed before the
  current error is emitted and only then be reset, or an equivalent argument
  must prove that an earlier cycle cannot coexist with the observed failure.
- Existing tests assert exact error text and scratch cleanup, buffer reuse,
  retained visited storage, and visited.bit_len after success and failure; see
  the reusable scratch checks in
  [file.rs](../../../../crates/litchi-cfb/src/file.rs#L3163) and the
  differential checks at
  [file.rs](../../../../crates/litchi-cfb/src/file.rs#L3286). Those are
  observable implementation contracts for this private helper and would need
  a deliberate, separately reviewed update.
- Collection still has to precede MiniFAT or regular-FAT ownership publication,
  and the two scratch namespaces and physical reconciliation remain separate.
  A replay cannot move claim_chain, mini-sector overlap checks, or the later
  [physical-layout validation](../../../../crates/litchi-cfb/src/file.rs#L1178)
  pass into the collector.

The broad visited-work direction already has a rejected precedent: 0524's
[test_and_set candidate](../change-0524/candidate.patch) fused the existing
contains and insert operations, and its [design review](../change-0524/design-review.md)
and [decision](../change-0524/decision.json) document the native-gate failure.
The exact terminal-proof/deferred-replay algorithm was not measured there, so
it is not directly disproved by that result. It does, however, reopen the same
valid-loop visited cost and must not be described as an already accepted or
low-risk optimization. The 0524 failure is a reason to demand a new proof and
fresh sub-operation attribution before considering it, not a reason to infer a
benefit from the theorem.

The minimal evidence for this hypothesis is therefore a baseline attribution of
the current prepare_visited zero fill, per-step visited-word test/update,
vector append, table load, and marker branches on XLS-owned and CFB few-large
positive timed workflows. If that attribution shows a credible target, a
separate design-only differential must cover an exact-N terminal chain, early
terminal, invalid marker, invalid index, cycle before N, cycle at the final
position, overlong declaration, empty chain, and repeated scratch reuse. It
must also exercise the existing fallible "sector-chain entries" and
"sector-chain map" boundaries or explicitly prove why their observable
behavior is unchanged. Only after those proofs and guards should a candidate
be built and measured under the full two-repeat native and allocation gates.
No speedup or allocation claim follows from the terminal proof.

The appended design note made no Rust edits, builds, tests, captures, production
changes, existing-evidence rewrites, or commits.

### Coordinator clarification: smallest terminal-proof experiment

The replay proposal need not reconstruct cycles with repeated linear searches
of the output vector. Re-running the existing authoritative bitset collector
once on failure is linear in the bounded chain/table work. Quadratic replay is
a risk of an alternative vector-search implementation, not a requirement of
this proposal. Likewise, retaining existing reservation and zero-fill ordering
initially would avoid stale bits and isolate removal of per-step membership
tests/updates. Deferring zero-fill to fallback is a separate, later possibility
requiring its own storage-state proof; neither generations nor touched-word
tracking is inherently required for one authoritative replay.

Private scratch-state tests identify invariants to review, not additional
public requirements by themselves. Any changed internal postcondition must be
justified against its consumers and the actual preservation/resource contract.
Keep fallible allocation ordering in the first proposed experiment to avoid
combining that question with cycle-check elimination. Malformed-input latency
still needs explicit bounds and guards: a short cycle could consume the declared
N steps before fallback. No candidate or speedup is admitted by these design
observations; fresh attribution and proof remain next.
