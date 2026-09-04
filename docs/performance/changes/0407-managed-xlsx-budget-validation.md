# Change 0407: validate managed XLSX budget evidence

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

Review of 0405 found that configuration comparison distinguished the new
planning allowance, but did not independently validate its type or its
relationship to the managed source evidence. Commit `2de94ea3f` validates that
contract before comparison.

When managed source evidence contains either new memory field, both fields
and the configuration allowance are required. Values must be unsigned 64-bit
integers (booleans are rejected), the reported allowance must match the
configuration, and the checked payload-plus-allowance sum must equal the
resulting memory limit. Partial tuples, malformed values, and overflow fail
closed. Historical reports without either new field retain old-to-old
comparability. The allowance describes the benchmark's bounded planning
context; it is not a total allocator or RSS bound.

All 81 focused comparator tests pass, including valid new evidence, legacy
comparability, missing fields, invalid scalar types, mismatched allowances,
and sum overflow. A fresh three-sample managed-XLSX CLI smoke reports a
64,321-byte payload budget, 65,536-byte planning allowance, and 129,857-byte
resulting limit. Its budget schema and parallel-metrics validations pass.
The [validation bundle](../results/change-0407/validation.json) retains the
source hashes, exact commands, logs, and complete diagnostic report. These
checks authorize no latency or memory-improvement claim.
