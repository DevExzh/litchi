# 0360adcaa — XLSX raw protocol review returns

Date: 2026-10-24. Parent: [0360adca](0360adca-xlsx-raw-native-protocol.md).

The initial protocol was returned before merge for incomplete source identity
and workload validation. The accepted preparation packet is
`5e6555d8406709382e26ecbcdd6cb18ded449e31`; independent review reports READY.

Preparation now requires real detached Git worktrees, including the original
`0befb112d96d2409e6f40c9abfeb4813665905ad` baseline. It binds both actual
HEADs and trees, a clean candidate, and only the explicitly allowed baseline
probe adapter. All eight consumed probe, manifest, allocator and lock inputs
have fixed reviewed hashes. A pair of equally modified inputs cannot satisfy
that contract. The receipt lives outside the fresh run output directory so
preparation does not make the build preflight reject its own output.

The report validator requires the exact sixteen-row matrix, read phases,
selected workbook part, decode counters, relationship IDs, part names and
content types. Allocator and native reports must cover the same lifecycle
and fidelity checks. Partial or substituted reports fail validation. Tests
cover lock and allocator mutations as well as the preparation path rules.

The probe header's historical candidate name describes the retained probe
source; the protocol binds the actual final candidate HEAD independently.
Changing that comment alone would change the reviewed source hash, so the
preparation packet preserves its byte identity.

Coordinator validation on the merged source passes all 37 preparation,
validator and shared-runner tests in 2.862 seconds:

```sh
python3 -m unittest tools.test_prepare_xlsx_raw_inputs tools.test_xlsx_raw_native_qualification tools.test_native_qualification_runner
```

This is an executable protocol preparation result. No release build, native
measurement, performance ratio or final acceptance is claimed. Candidate
source freezing, real execution and coordinator performance review remain
required. Nothing is pushed.
