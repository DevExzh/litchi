# 0797 — first-attribute allocation-avoidance preflight

**Do not advance this prototype to workflow trials.** It passes semantic checks
and improves one-attribute consumption, but 20 consumption rows trigger the
frozen preflight regression flag. Production remains unchanged.

This batch tests a new, smaller iterator prototype without changing production.
The 0794 candidate remains rejected. A direct-helper result cannot establish
public-workflow benefit or authorize production adoption.

## Candidate and protocol

The [packet](results/change-0797/README.md) begins at `08a406b319`. The prototype
reads its first attribute with quick-xml's duplicate check disabled. A successful
first item whose remaining bytes are whitespace ends the iterator immediately.
If another item may follow, the next call rebuilds quick-xml's checked iterator,
replays the first attribute to seed duplicate state, and returns to the existing
checked path. The 32-name handoff and ordered-map bound remain in place.

This avoids building a duplicate-name list when only one attribute is consumed,
but adds a replay and dispatch for tags with more attributes. Unlike 0794, it
preserves quick-xml's early duplicate rejection through the first 32 names.
The archival patch covers the five canonical helper copies and their shared
tests. Production source and the 35 inherited architecture inputs remain exact.

Thirty-three literal inputs independently freeze the same error/boundary matrix
as 0796 plus distinct-name counts two and three. Both source legs compile into
one standalone release binary. Construction exposes the iterator reference to
`black_box` then drops it; consumption includes construction, first-error/end
iteration, checksum work, and destruction. No oracle or JSON work is timed.

Root runs all builds, tests, and captures serially on recorded CPU 12. Six
alternating native block pairs use 30 samples after three warmups, each with
4,096 repetitions of the same hot input. Process p50 is the fifteenth sorted
sample; the median of six paired ratios has a 10,000-resample bootstrap with
seed 797079 and ranks 250/9749. Two Callgrind repeats reverse the leg order and
use one sample/iteration without warmup. The exact non-inlined owners delimit
Ir, Bc, Bcm, Bi, and Bim, with one positive dump and one empty termination dump.

The frozen preflight rule advances only if semantics match, distinct-1
consumption improves at least 3% with interval upper bound below one, and no
consumption row exceeds ratio 1.05 with interval lower bound above one.
Construction flags remain diagnostic. Passing would justify fresh paired
public-workflow, resource, and cross-format trials; failing archives this
prototype without advancing it. This is a bounded experiment-selection rule,
not a substitute for the production adoption policy.

## Correctness and measurement scope

Minimal isolated workspaces compile the exact baseline and candidate helper
copies with quick-xml 0.41.0 and the shared helper tests. Thin wrappers supply
crate layout for the canonical-copy test; they do not compile full production
crates or establish facade compatibility. The separate direct probe checks
borrowed key/value sequences, first-error variants and byte positions, clone
transitions, and repeated terminal `None` against a fail-fast quick-xml adapter.
The literal input/checksum audit and scalar raw-counter audit are independent
of the full analyzer.

Opaque construction forces materialization and can differ from inlined callers.
Repeated hot inputs, timer/function-pointer and checksum costs, compiler layout,
and host variation limit native micro ratios. Counter runs have a different
iteration count; guest instructions and simulated branch misses are not native
hardware events or allocation API counts. Historical timing is not pooled.
No baseline, CRUD, real-producer, cold/range, or concurrency coverage is promoted.

## Results and disposition

The final helper gates pass with 65 baseline and 85 candidate test executions
(13 and 17 shared tests compiled in each of five copies), zero ignored or failed,
and Clippy with warnings denied. The direct probe passes formatting, locked
release build/check/Clippy, literal catalog comparison, and semantic preflight.
A dedicated sole empty-value input is not included in the timing matrix;
the shared differential tests cover empty-value parsing. This preflight does
not claim exhaustive parser verification. The first helper attempt's test-only E0631 compilation failure is retained with
its exact source mirror and the one-line correction; it was not a semantic
failure or a discarded measurement.

All 1,056 capture processes terminate successfully, covering 24,024 measured
samples. Independent readers reconstruct every input, accepted sequence, first
error and result checksum, and conserve all five counters across 528 raw dumps.
All 264 exact owners qualify with one incoming call and conserved self-plus-child
partitions. All 66 independently recomputed native rows match the full analyzer.
Both iterators measure 120 bytes in every report. There are no retries or omitted samples. The table selects the important
short-tag and early-error boundaries. Ratios are candidate/baseline; guest Ir
counts use one iteration and are identical in both repeats for these rows.

| Case | Mode | Native ratio | Bootstrap interval | Before Ir | Candidate Ir |
|---|---|---:|---:|---:|---:|
| distinct-0 | construct | 0.923207 | 0.922428–1.074916 | 70 | 69 |
| distinct-0 | consume | 1.050633 | 1.050351–1.348677 | 170 | 181 |
| distinct-1 | consume | 0.642317 | 0.621102–0.643662 | 574 | 428 |
| distinct-2 | consume | 1.429871 | 1.426508–1.453885 | 910 | 1,312 |
| distinct-3 | consume | 1.333872 | 1.310475–1.353962 | 1,286 | 1,711 |
| distinct-4 | consume | 1.270328 | 1.249569–1.291744 | 1,702 | 2,150 |
| distinct-5 | consume | 1.165349 | 1.154222–1.182471 | 2,615 | 3,086 |
| distinct-8 | consume | 1.095853 | 1.086595–1.106744 | 4,416 | 4,956 |
| distinct-16 | consume | 1.055608 | 1.047777–1.071422 | 9,294 | 10,018 |
| distinct-32 | consume | 1.037729 | 1.029105–1.045308 | 25,699 | 26,789 |
| distinct-64 | consume | 1.007392 | 0.998386–1.019682 | 81,984 | 83,714 |
| duplicate-valid-after-1 | consume | 1.389820 | 1.356750–1.412975 | 754 | 1,133 |
| duplicate-long-quoted-after-1 | consume | 1.366083 | 1.355620–1.391794 | 754 | 1,133 |
| duplicate-long-unterminated-after-1 | consume | 1.367566 | 1.356111–1.400956 | 754 | 1,133 |
| syntax-unique-tail-after-0 | consume | 0.800047 | 0.795135–0.824381 | 448 | 307 |

Between-process p50 spread exceeds 5% in 22 of 132 case/mode/leg groups.
All spreads and raw tails remain in the packet. The six-block bootstrap describes
these observed pairs, without removing host or code-generation uncertainty.

Distinct-one consumption improves 35.768%, with ratio 0.642317 and interval
0.621102–0.643662. This clears the one-attribute benefit requirement. However,
distinct-two regresses 42.987% and distinct-three 33.387%, with their entire
intervals above one. Twenty of 33 consumption rows fail the diagnostic rule.
The replay tradeoff is therefore not advanced to public-workflow testing.
All construction ratios lie between 0.922428 and 0.925676, with no construction
regression flag; these opaque, very short operations remain diagnostic only.

The candidate avoids 0794's long-duplicate-value scan through the early checked
path: after one accepted attribute, quoted and unterminated duplicate regions
both execute 1,133 guest instructions versus baseline 754. Their approximately
1.37 native ratios still regress because of added work, but the 4,096-byte value
is not scanned before duplicate rejection. This source/counter agreement does
not assign a native cycle fraction to replay or establish an allocation count.

The next investigation should quantify the attribute-count and caller-consumption
distribution in the representative document workflows before selecting another
iterator specialization. A replacement must avoid first-attribute replay on
common multi-attribute tags while preserving early duplicate refusal and the
existing hostile-input bound. No optimization is retained from this packet.

After all captures and audits completed, the owned target was removed
(176,597,444 logical bytes), retaining its exact binary identity in the cleanup
witness. Unrelated working-tree files and existing worktrees remain intact.
The failed helper attempt, final tests, raw captures, analysis, independent
audits, and reviews are retained and sealed.
