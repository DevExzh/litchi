# 0362ul135c — paired current-native v2 execution contract

Date: 2026-10-25. This source/Python-only follow-up preserves 135, 135a and
135b. It does not run Cargo, rustdoc, a native process, timing, allocator or
performance workload.

The v2 consumer now builds and executes both fixed source endpoints. The
`before` plan is bound to the baseline source, target and archive, and the
`after` plan is bound to the candidate source, target and archive. Their
working directories, Cargo target trees, lock guards, adapter features and
binary identities are distinct. The legacy adapter uses
`rtf-legacy-api`/`rtf-current-api`; the PPT auxiliary adapter uses the reviewed
`--all-features` profile at both endpoints. Every route receipt records the
actual source cwd, argv and selected binary identity, and validation compares
those fields against the fixed plan rather than trusting side labels.

The historical RTF contract is retained: each route has five independent
ABBA blocks, each block uses `before, after, after, before`, and each native
child retains its 31-sample report. PPT keeps one four-leg ABBA sequence. The
execution and qualification receipts carry the lane block count and complete
per-side binary provenance; the validator rejects missing, reused or
mis-bound block/leg streams.

Route output admission is typed and fail-closed. Legacy streams must contain
the exact scenario, workload, positive median, outcome, fidelity and fixed
refusal fields. PPT suite/probe/validation outputs must contain their reviewed
route, status or correctness receipt fields. Untagged `stdout` and arbitrary
`output_bytes` are rejected; the PPT final artifacts come from the actual
side-specific child outputs (or the built binary), not a synthesized route
metadata object. Execution rejects qualification/source/archive/target/state
overlaps before invoking a child or writing a receipt.

Focused source-only verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
Ran 25 tests
OK

PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2 tools.test_no_3rdparty_references
Ran 29 tests
OK
```

The tests cover distinct baseline/candidate cwd, target, feature, build argv
and binary identities; five-block legacy execution; strict typed route output;
qualification overlap refusal before executor entry; and receipt tampering.
