# 0555 metrics-analysis contract

`analyze_metrics.py` is the read-only numerical consumer for the matched
0555 OLE2 physical-marker experiment. It consumes only the retained `plan.json`,
`frozen-inputs.json`, stage manifests, build/capture receipts, native reports,
allocator reports, corpus catalogs, and GNU-time RSS sidecars. It never builds,
runs a capture, edits source, removes the owned target, or invents a missing
measurement. The CLI writes its report with exclusive-create/identical-replay
semantics; importing the module and calling `analyze()` is read-only.

The stage matrix is reconstructed from the frozen plan. Each stage has two
native and two allocator jobs per repeat, in the frozen ABBA order. Baseline
repeat two is measured with the retained baseline binary while its receipt is
bound to the candidate execution manifest. Every receipt is checked for its
actual output stage, execution stage, source and execution-manifest hashes,
workspace-lock hashes, binary hash, driver and plan hashes, environment,
artifact inventory, timestamps, and non-overlap. The four binary descriptors
are retained as explicit records with `present`, `sha256`, byte count, build
receipt, source-manifest, and workspace-lock bindings. A missing retained
binary is therefore observable after cleanup without replacing its descriptor
hash with a synthetic value.

Native elapsed vectors come only from the non-instrumented executable. The
allocator executable's elapsed vector is validated but excluded from latency
comparisons. All allocation vectors are retained, checked for non-negative
values and live-byte balance, and summarized independently. The native process
RSS sidecar is a whole-child high-water observation and is compared separately
from per-case elapsed samples. Corpus/catalog binding and semantic/source
identity are checked before any numerical comparison.

The report schema is `ole2_physical_marker_metrics_0555_v1`. It retains every
native and allocation comparison, every matched adverse row above the 5%
threshold, and every same-stage repeat-drift row whose absolute change exceeds
5%. Rows are never filtered to make a gate pass. Percent changes use
`candidate / baseline - 1`; a zero baseline is passable only when the paired
value is also zero. The report includes the complete stage rows, sample vectors,
identities, RSS comparisons, allocation summaries, and the frozen plan/input
hashes needed for deterministic replay.

The mandatory main gates are:

- the four frozen primary XLS workflows improve in p50 by at least 3% in both
  repeats;
- the same four primary XLS workflows have mean change no greater than +5% in
  both repeats;
- every other XLS workflow has p50 and mean change no greater than +5% in both
  repeats;
- every CFB shape, including `many-small`, has p50 and mean change no greater
  than +5% in both repeats;
- each retained native process-RSS comparison is no greater than +5%; and
- all 72 allocation rows (12 workloads, two repeats, and allocation calls,
  allocated bytes, and incremental region peak) are no greater than +5%.

The 0554 candidate-specific positive `many-small` requirement is deliberately
not carried forward. The 0555 plan records this before any 0555 capture and
treats `many-small` as a required CFB control. This removes a candidate-specific
positive target while preserving the inherited four-primary-XLS p50 requirement;
it does not use observed 0555 results to change a gate.

`main_gates.all_frozen_main_gates_pass` is the conjunction of those groups.
`disposition` remains descriptive: a failed group reports rejection of the
metrics admission, while a passing numerical matrix still requires the
independent profile, correctness, review, and quality gates before adoption.
No hardware-counter, cold-cache, remote-provider, scaling, ODF, iWork, or
production-adoption claim follows from this consumer.
