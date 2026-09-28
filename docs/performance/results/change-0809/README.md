# 0809 — repair the baseline PPTX refusal-test lint gate

This batch replaces exactly three `.err().expect(message)` expressions with
`.expect_err(message)` in `crates/litchi-pptx/src/opened/tests.rs`. Existing
fixtures, calls, messages, and typed error assertions remain intact. Both
success types already support `Debug`; no production type or lint policy changes.

The previous batch reproduced these three Clippy errors on the unchanged
baseline as well as the archived direct-event candidate. This repair is a
necessary quality enabler, not an optimization or a performance result.
The next event-handling experiment must build fresh before/after binaries.

`origin.json` records the base revision, the one-file scope, and hash-bound
architecture/source references. `quality.py` runs six serial package gates in
an initially absent `/home/zhuhe/code/litchi-target-0809`: format, all-targets
check, tests, warning-denied all-targets Clippy, warning-denied rustdoc, and the
repository crate-boundary check. `plan.json`, `checks.json`, and the six raw logs
record commands, source identity, environment, status, and test totals.

`validate.py` independently checks the complete 9,196-file source census and
reconstructs the exact three replacements from the base revision. It replays
receipt/log identities, commands and test totals. After root removes the owned
target, replay with:

```text
python3 -B docs/performance/results/change-0809/validate.py --require-cleanup
```

No benchmark, profiler, new fuzz campaign, native Office test, cross-platform
run, or iWork work is included. The broad performance program remains open.

`additional-inputs.json` and `inputs/` retain the exact ignored root Cargo
lockfile and tracked rustfmt configuration. They were captured after the Cargo
gates; both original modification times precede the first gate, all dependency
commands used `--locked`, and rustfmt configuration matches the base commit.
Use these retained inputs when reproducing the recorded commands. This is a
supplemental input record, not a claim that they were in the original source
census. No root dependency or configuration file was changed.

All six gates pass: 1,238 tests passed, zero failed, and three were ignored.
The repository boundary check accepts 65 packages and 244 internal dependency
declarations with the existing 11 explicit debt items. The owned target was
removed (7,120 files; 1,947,648,722 logical bytes), and offline validation passes
after removal. Sealing includes this packet, six report/index documents, and the
single test-source change; staged and committed blob inventories are checked
separately with `seal.py`.
