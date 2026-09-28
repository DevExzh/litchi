# 0820 independent repair results review

Status: **repair quality evidence and pre-cleanup packet validation pass**.
This review covers the test-only repair after the original quality stop. It
admits no durability timing, release build, export, qualification, or workload
result. Cleanup and seal remain root-owned final checks.

## Fresh repair gates

`repair/quality.json` records `status: pass` and six gates. The focused
`docx_replayable_tail_append` receipt exits zero; its log reports all five
allocator wrapper tests passed under two test threads, including the repaired
`global_allocator_records_successful_alloc_and_dealloc_with_process_live_accounting`
test, with no failure or poisoned lock.

The full repair test summary reports 28 result summaries, 641 passed, zero
failed, and one ignored. The raw test log contains no `test result: FAILED.`
The six terminal gate receipts all have exit code zero:

| Gate | Command class | Result |
| --- | --- | --- |
| 1 | formatting check | pass |
| 2 | all-feature, all-target check | pass |
| 3 | all-feature tests, two test threads | pass |
| 4 | warnings-denied Clippy | pass |
| 5 | warnings-denied rustdoc | pass |
| 6 | crate-boundary validation | pass |

The boundary log independently reports 65 workspace packages, 244 internal
dependency declarations, and 11 existing explicit debt entries. These are
quality and structural checks; they do not measure the planned save workload.

## Repair scope and retained failure

The fresh repair source witness matches all 9,197 production files and all 87
tool files except
`tools/perf-baseline/src/bin/support/counting_allocator.rs`. Its bytes before
`#[cfg(test)]` match the archived pre-repair file; the fragile net-live-byte
assertion is absent and the signed conservation helper is present. The repair
origin records `production_changed: false`, `runtime_harness_changed: false`,
and the test-only allowlist. The quality run therefore verifies the intended
test boundary without changing allocator runtime behavior.

The original frozen attempt remains separate evidence. Its quality receipt
still records gate 3 exit 101, with the process-global `live_bytes` assertion
failure followed by the poisoned test-lock failure. That failed run is not
relabelled as a passing quality result, and the planned 216 reports / 4,488
samples remain unexecuted counts.

## Validator replay and correction

The first read-only packet-validator run, made after the terminal boundary
receipt and `repair/quality.json` appeared, exited at `original frozen-input
keys changed`. The archived
`quality-0/frozen-inputs.json` has these top-level keys:

```text
architecture, corpus, drivers, host, locks, packet, root_inputs, schema, unrelated
```

The first validator required an additional `provenance` key. The frozen-input
writer in `quality.py` also writes the nine-key shape, while `provenance.json`
remains a member of the frozen packet hash set. This was a validator/evidence
schema mismatch, not a failed repair gate. The bounded correction aligns
`validate.py:frozen_inputs()` with the already frozen nine-key witness and
leaves provenance covered by the packet hash set. `validate.py` parses cleanly,
and the read-only replay now returns `status: accepted`, `focused: true`, and
`quality_gates: 6`; cleanup and seal are correctly still unchecked at this
stage.

No Cargo, build, binary, workload, or timing command was run for this review.
