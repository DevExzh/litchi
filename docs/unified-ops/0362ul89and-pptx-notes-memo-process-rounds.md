# 0362ul89and — PPTX notes-memo repeated process-round closure

This source-only follow-up closes the remaining process-boundary findings from
89anc. It changes no public API, production dependency, fixture, build, or
performance result.

`run_pptx_notes_memo_abba.py` now consumes a strict version-2 manifest with an
explicit one-based `rounds` array. Every round contains fresh commands in the
fixed `A1, B1, B2, A2` order; four rounds therefore launch 16 distinct child
processes before any acceptance summary is written. The receipt records each
round and its validator summary separately from the probe's in-process
`samples_per_round`/`rounds` statistics. The synthetic packet executes all 16
processes and proves unique `(pid, start_time_ticks)` and report identities.

Each leg is bound in the manifest to its source revision and canonical
absolute target path, SHA-256, byte length, and permission bits. The runner
checks command execute bits, observes the expected target in the same PID
(direct execution or a bounded same-PID exec wrapper), rechecks the target
after exit, and binds the report's source and binary identity to the leg. The
receipt includes every command/proc/report identity and the exact manifest
byte length and SHA-256. Build provenance is intentionally not inferred here;
the coordinator's later approved build receipt must bind the source/build
inputs.

Both validator and runner use strict JSON parsing: duplicate object members
and `NaN`/`Infinity` constants are rejected. The direct four-report validator
requires an explicit `--descriptive-only` marker; process acceptance is only
through the repeated-process runner. Direct report validation also rejects an
unknown mode before inspecting report fields. Negative tests cover the strict
JSON, unknown-mode, execute-bit, target-identity, duplicate-report, and
direct-CLI bypass cases.

Validation performed:

```text
PYTHONWARNINGS=error python3 -m unittest tools/test_pptx_notes_memo_abba.py tools/test_run_pptx_notes_memo_abba.py -v  # 23 passed; 16 child processes in the positive runner case
python3 -m py_compile tools/pptx_notes_memo_abba.py tools/run_pptx_notes_memo_abba.py tools/test_pptx_notes_memo_abba.py tools/test_run_pptx_notes_memo_abba.py
rustfmt --edition 2024 --config-path rustfmt.toml --check tools/perf-baseline/src/pptx_notes_memo.rs
```

No Cargo, native execution, benchmark, Callgrind, quiet-window, or
performance qualification was run.
