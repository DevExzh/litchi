# 0362h — Return W8 File performance qualification for completion

- Date: 2026-10-21.
- Reviewed branch tip: `8f84379e31`, `uo2/w8-file`.
- Status: final performance qualification required; no production merge or
  performance exception is accepted by this review.

The historical Part A measurements in record 0362 remain useful observations
of the unified API before and after its naming migration. They do not close
the performance requirements of [0351 §2.6](0351-api-contract.md#26-consumers-fidelity-performance)
and the [goal's performance bar](../UNIFIED_OPS_GOAL.md). The immutable source
blobs and their hashes are recorded in the
[source-review receipt](results/0362h/source-review.json).

`file_w8_probe` times only unified `File` calls. It has no equivalent leaf
legs, output route or repeated-save comparison. Its `black_box` calls keep
successful operation results observable to the compiler, but the write
results are dropped without checking the changed state or publishing their
output. The open/get route observes the selected path's string length.
These observations cannot establish deferred-output fidelity or the output
cost that the contract specifically requires.

The JSON retains sample counts, medians and minima, without raw samples,
operation counts or elapsed totals. The shell driver records load averages
but has no process-overlap guard. A failed probe is followed by `echo FAIL`,
so a later successful command can leave the driver with a zero exit status.
The old measurements are retained, including the unfavorable DOC merge run;
this review does not substitute a favorable rerun for it.

Record 0362's general 1.05 before/after threshold is not the goal's paired
leaf comparison. Its PPTX code-placement explanation is an inference from
an unchanged leaf path, not an instruction, allocation or placement
attribution in that receipt. Neither establishes acceptance under the
current coordinator rule.

The repair must use new harness routes and preserve the historical sources
and results. It must measure the final integrated source with matched
dependencies and features, equivalent leaf operations, calibrated timed
spans, raw operation counts/totals and guarded ABBA legs. Validate changed
state and published output outside the edit timer where appropriate; measure
publication and repeated output separately. Preserve exact no-op and inverse
checks where the route supports them. Record source, lock, fixture and binary
hashes, propagate failed child exits, retain every attempt, and include
allocation evidence for changed ownership paths. Register the resulting
routes in the benchmark matrix.

This is a source review, with no new timing or correctness-test claim.
W4/W4b/W5/W8 integration, the separate PPT repair, the remaining raw/shared
API work and the final workspace gate remain required. Memo item 22 is the
only deferral, and its diagnostic redaction rules continue to apply.
