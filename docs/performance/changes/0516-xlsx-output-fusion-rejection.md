# 0516: XLSX emitted-output parser evaluation

The candidate fails the frozen latency admission gate. Its changed-edit pilot
medians rise across all 12 main rows and all six warm changed-edit guard rows.
This batch makes no speedup claim and does not run the conditional formal
ABBA campaign. The independent no-op test-fixture repair is retained.

## Mechanism and correctness

0515 attributed about 6.9% of scoped commit instructions to the duplicate
output XML reader and namespace-reader maintenance. This candidate feeds the
existing semantic parser from normalized events emitted by the compactor,
only for effective changed worksheets requiring Store verification. It retains
exact compacted bytes, preservation and error ordering, web/style/readback and
publication checks, and the existing bounded Store handoff. Failed admission
uses the authoritative exact-output parser. Source-layout planning and exact
no-op detection keep their existing paths.

The safety proof excludes MCE/x14ac, shared-string dependencies, and shared
formula expansion; it checks decoded forms, namespace scopes, and emitted
bytes. Resource accounting includes four retained payload copies throughout
parsing, including inline decoding, plus record, row-index, and merge-index
allowances. The 128 MiB threshold concerns compact output capacity and new
parser/proof state; it is not a process-memory or whole-worksheet limit.

The final candidate passed all 1,284 all-features tests, with zero failures,
ignored tests, or exclusions, and passed Clippy, formatting, rustdoc, owner,
and boundary checks. The preceding snapshot passed 970 default-feature unit
tests. The [semantic review](../results/change-0516/final-implementation-review.md),
[resource review](../results/change-0516/final-resource-review.md), and
[ADR matrix](../results/change-0516/adr-matrix.md) record the proof and scope.

## Pilot admission

The frozen protocol uses CPU 2 and serial captures. Each pilot row contains
20 samples after two warmups; the three shapes are 8×8, 32×32, and 256×256,
with two sheets. These are admission observations, not a formal confidence
claim. Every individual row and raw sample is retained in the evidence.

| Native pilot family | Median change, candidate minus control | Rows above 5% |
| --- | ---: | ---: |
| Main commit and commit/save, all 12 rows | +4.46% to +8.30% | 10/12 |
| Plain guard, six warm changed-edit rows | +5.82% to +8.64% | 6/6 |
| x14ac fallback, six warm changed-edit rows | −0.44% to +2.50% | 0/6 |

All 15 cold-read and same-value guard medians are at or below +3.89%; apparent
pilot improvements are not promoted to speedup claims. Protecting these paths
does not compensate for the consistently slower eligible changed-output path.

Whole-child native pilot RSS is 137,196→137,604 KiB for the main family
(+0.30%), 165,104→170,600 KiB for the plain guard (+3.33%), and
246,048→261,800 KiB for the fallback guard (+6.40%). The last observation is
an additional review flag, not a per-operation memory result. Allocator-run
elapsed time and RSS are excluded from native comparisons.

## Independent baseline fixture repair

Four managed exact-no-op tests also failed on unchanged production. Their
bounds omitted the OPC owner's existing 64 KiB exact-source copy window. Two
test files now allow that scratch explicitly and use deterministic
incompressible untouched members, with assertions that both the member and
entire archive exceed the available memory budget. Byte identity and zero
budget after drop remain checked. The repaired unchanged-production suite
passed all 1,263 tests without exclusions. See the
[fixture review](../results/change-0516/baseline-test-review.md).

Earlier compile/test failures, the target-cache quota failure, corrected
source snapshots, and their exact patches remain archived. They are not
silently relabeled as final passing checks.

## Allocation and instruction diagnostics

Two separate allocator repeats agree. On the four dense main rows, allocation
calls increase 5.20–5.52% and allocated bytes increase 2.94–3.12%; the dense
warm changed-edit guard increases calls 7.04–7.10% and bytes 4.70–4.78%.
Incremental peak live demand is essentially unchanged (differences below
0.002%). Fallback and no-op allocation observations retain the same counts.
These are operation-region observations, not document memory or RSS.

The three-call dense commit/save diagnostic increases from 14,458,444,767 to
14,928,683,870 simulated instruction references (+3.25%). Its commit edge
increases from 11,911,869,450 to 12,382,107,894 (+3.95%); the writer edge is
essentially unchanged at 2,546,574,684→2,546,575,343. Raw, inclusive, and
exclusive selected-edge attribution agree and preserve exactly three calls.
Both profile runs reported the Valgrind `brk segment overflow` warning and
exited successfully. This diagnostic is not hardware-counter or timing
evidence, and does not separately quantify every new safety-check cost.

See all individual results in
[`metrics-summary.json`](../results/change-0516/metrics-summary.json),
the [`profile-summary.json`](../results/change-0516/profile-summary.json), and
the [independent admission review](../results/change-0516/pilot-admission-review.md).

## Completion record

Five tracked candidate files are restored to the control revision and two
candidate-only test files are removed. The two independent fixture fixes
remain. The live source exactly matches the separately tested `baseline-tests`
manifest; the complete nine-file candidate is retained as an authenticated
patch. See [`restoration.json`](../results/change-0516/restoration.json).

The evidence retains 26 captures, eight historical failed commands, all source
epochs and reviews, and the frozen protocol. Replay its custody, correctness,
metrics, and decision with:

```sh
python3 -B docs/performance/results/change-0516/verify.py --post-cleanup
```

OLE2 and OOXML remain the priority until their full optimization goal is
complete; ODF is deferred and iWork is excluded. The next measured
investigation is DOCX publication CPU attribution, with the 0499/0500 tail,
K=1, and RSS review flags still open. Completed CFB FAT batching is not an
unstarted task. Another worksheet-reader fusion should not proceed without a
measured design that avoids the additional work recorded here.
