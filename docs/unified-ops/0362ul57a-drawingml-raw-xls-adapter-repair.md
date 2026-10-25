# 0362ul57a — combined DrawingML historical consumer adapter

Date: 2026-10-25. This source-only follow-up preserves the original 0362ul57
failure and repair record. It extends the role-aware preparation adapter for
the reviewed XLS consumer rename in `2ffacb16e9`; it does not run Cargo, a
native probe, or timing.

## Composed source fence

Integration `94c7bc105e55e82911fa61680774884165b24275` still contains the
pre-XLS-composition `tools/perf-baseline/src/lib.rs` with hash:

~~~text
pre-composition source: ff0ca3a924852cfc5d8497f812fbf412c4eb71b5f289f1c3b3136dea92c7ed9f
~~~

The XLS leaf changes exactly one perf consumer spelling:

~~~text
litchi_xls::validation::XlsValidationLimits::default()
  -> litchi_xls::validation::Limits::default()
~~~

The coordinator composes that one source change before creating a fresh raw
preparation receipt. The resulting candidate `lib.rs` identity is:

~~~text
composed candidate: c28d436cdbd6c1746fc0be31b7b1ca28d46526143d6186ceb2c7bf5dbcdecf02
historical baseline: 280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
~~~

The historical hash remains exactly `280d61…`: baseline preparation replaces
the two current RTF spellings and the one current XLS spelling. The candidate
is never rewritten. The adapter records three replacements and a 29-byte
historical byte delta; all other adapter files and all DrawingML timed code
remain byte-identical.

## Role-aware receipt and checks

The preparer now names the combined policy as the consumer API adapter rather
than an RTF-only adapter. New receipts use `role_consumer_adapter` and bind:

~~~text
path: tools/perf-baseline/src/lib.rs
replacement_count: 3
changed_bytes: 29
candidate_sha256: c28d436cdbd6c1746fc0be31b7b1ca28d46526143d6186ceb2c7bf5dbcdecf02
baseline_sha256: 280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
~~~

The source-bound unit test constructs the composed candidate bytes in memory,
proves the baseline bytes are exactly those three substitutions, checks both
role hashes and rejects an already-transformed/wrong-hash input. Existing
0362ul57 receipts and records are retained; this follow-up requires a fresh
preparation output before any native build or measurement.
