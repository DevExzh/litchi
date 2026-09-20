# 0717 DOCX phase process-counter diagnostic

This packet adds an explicit `ordinary-save-process-metrics` harness feature
following 0716's whole-process fault association. No production crate changes.
The prior 0715 candidate remains rejected.

The frozen matrix contains sixteen blocks, two corpora and two executable
lanes: native and opt-in procfs. Each of the 64 fresh processes uses CPU 12,
100 warmups and 200 in-process counting-publication samples. Four cyclic
orderings repeat four times; every treatment occupies each position four times.
All children, samples and descriptive flags are retained.

The procfs lane snapshots same-process counters around the phase interval,
before owner drop and digest/readback verification. It retains 32 adjacent empty
snapshot controls before warmup and raw acquisition-order sample deltas. The
analyzer validates their alignment against every elapsed-sorted metric vector.
Controls measure counter deltas only: they have no durations and are never
subtracted. Instrumented elapsed is diagnostic and is not native latency.
Counters are not owner-exclusive; RSS deltas are nonnegative endpoint changes
and peak RSS is a process-lifetime high-water mark.

The parent separately retains whole-child `wait4` and before/after system
context. Those include setup, warmup, untimed opens/edits, verification and
reporting. System context can also include unrelated activity. Rust hash seeds,
allocator/ASLR layout and unlisted environment variables remain uncontrolled.
No fault association establishes a causal explanation. The generic report
filesystem isolation flags do not describe this route: each child contains
200 in-process samples, as recorded in the ordinary-save phase evidence.

Final-source builds, fresh standalone harness checks and six repository gates
are recorded. Unchanged production verification is explicitly reused from 0713:
4,995 tests, 92 passed/46 ignored doctests, with seven baseline-proven PPTX/XLSB
test Clippy exceptions. Unused initial builds precede scope wording and Clippy
cleanup; their logs and removal witnesses remain under `build-attempts/01`.

Two analyzer attempts are retained under `analysis-attempts`: an overbroad
artifact glob and a live-file/cleanup-witness output-label mismatch.
`analysis-reconciliation.json` verifies that the final correction changes only
validation labels and analyzer identity, with all statistics and capture bindings
unchanged. No raw capture was replaced.

Raw artifacts, source/build/fixture identities, frozen protocol, analysis,
corruption checks, review and cleanup records are covered by the final artifact
manifest. Replay needs no retained executable.

```sh
python3 -B docs/performance/results/change-0717/audit.py
python3 -B docs/performance/results/change-0717/artifact-seal.py --check
```
