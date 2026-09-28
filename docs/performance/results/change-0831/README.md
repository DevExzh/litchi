# 0831 — avoid the unused XLSX column-action map

This packet tests one production change against the exact preceding source:
return from `validate_column_actions` when the action map is empty, after the
protected-sheet guard. The 0830 allocation profile identified the unnecessary
512 KiB map; this packet makes a new matched before/after measurement.

`plan.json` fixes nine cases, six counterbalanced native blocks, two observer
blocks, CPU 12, warmups, sample counts and the adoption gate before execution.
Native real-file and no-op reports have 500 samples; synthetic scale guards
have 30. Allocation-observer reports have three samples and do not contribute
latency claims. Qualification adds one sample for each source/binary/case.
The planned total is 180 reports and 20,304 measured samples.

The `driver.py` stages are root-owned and serial. Agents prepare candidate and
reader source and perform read-only reviews. `inputs.json` freezes production,
tools, normative inputs, corpus, driver, candidate archives and the pinned
0830 correctness probe. Every child command has start/terminal receipts and a
log. The only permitted production difference between legs is the archived
three-line guard in `validation.rs`.

The command sequence, from the recorded `/home/zhuhe/code/litchi` checkout, is:

```sh
python3 -B docs/performance/results/change-0831/driver.py prepare
python3 -B docs/performance/results/change-0831/driver.py before
python3 -B docs/performance/results/change-0831/recovery.py freeze
python3 -B docs/performance/results/change-0831/recovery.py before
python3 -B docs/performance/results/change-0831/driver.py install
python3 -B docs/performance/results/change-0831/recovery.py after
python3 -B docs/performance/results/change-0831/run_reader.py qualify.py
python3 -B docs/performance/results/change-0831/recovery.py capture
python3 -B docs/performance/results/change-0831/run_reader.py analyze.py
python3 -B docs/performance/results/change-0831/run_reader.py audit.py
```

Commands refuse to overwrite evidence. Repeating the experiment requires a
new packet, owned target/scratch paths, base revision and input freeze. Existing
analysis can be replayed with `analyze.py --check`; no executable is needed for
offline numerical replay. Reader sources and every attempt are retained under
`reader-attempts/`, including failures.

Both source legs run fresh XLSX formatting, all-feature/all-target checking,
full tests, Clippy, rustdoc, crate boundaries, the independent pinned 0830 byte
oracle and harness library tests. The root and standalone harness locks have
different existing versions; `lock-parity.json` records this, and each lock
stays fixed across legs. Measurements use the harness lock. The initial
preparation failure requiring cross-lock equality is retained; it occurred
before the input freeze or any build/workload execution.

The real edit benchmark times the public edit and records its admitted outcome;
it does not serialize every timed edit. The separately run pinned probe checks
the complete reference output on both source legs. Real lifecycle samples
include opening, editing and atomic saving, and check published hashes.
Synthetic commit/save samples compare deterministic expected bytes and reopen
the result; setup/edit planning is outside those timers. The no-op control
does not reach the changed validator. These scopes are not interchangeable.

Allocation metrics count requested bytes, calls, region peak and signed net
live differences under the existing serialized allocator observer. Process
RSS is separate. A requested-byte reduction alone is not a peak/RSS reduction.
Every per-case result and regression flag must remain visible in the report.

The original `before` invocation stops at an incorrect observer identity
assertion after the observer process succeeds. Its selected process-metrics
feature emits the procfs-decorated allocator identity. The additive
`recovery.py` requires that exact identity and freezes its source plus all
preexisting evidence before resuming. Complete baseline reports are validated
in place; none are rerun or overwritten. Workload parameters and adoption
rules remain fixed. The original driver and failure are retained.

The completed result adopts the guard: real-edit requested bytes fall by
524,288 (18.09%), with paired native p50 ratio 0.980416 and bootstrap interval
0.975937–0.983917. No frozen regression threshold is triggered. The lifecycle
interval includes 1.0; real peak above entry is unchanged. The complete
nine-case results, 20 native spread flags and 10 tail diagnostics appear in
[the report](../../0831-xlsx-empty-column-actions.md).

`analysis.json` and independent `audit.json` agree on all 180 reports / 20,304
samples. `decision.json` records root adoption and replays baseline-adjusted
peak diagnostics. The initial reader preflight failures (wrong allocation
argument envelope and the native no-op's omitted allocation envelope) remain
under `reader-attempts/`; no workload was repeated to repair them.

Offline replay after cleanup uses:

```sh
python3 -B docs/performance/results/change-0831/analyze.py --check
python3 -B docs/performance/results/change-0831/audit.py --check
python3 -B docs/performance/results/change-0831/decision.py
```

`cleanup.json` retains all four removed executable identities and the owned
target/scratch inventory. Reader attempts 09–11 replay the analysis, independent
audit and root decision successfully after cleanup. The immutable seal covers
this packet, the production guard, the report and the five performance indexes.
