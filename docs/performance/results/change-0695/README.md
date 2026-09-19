# Change 0695 — weight remaining PPTX MCE work

performance_claim: none

This is an attribution experiment on the unchanged production revision
`df44030d6`, not a production optimization or an end-to-end performance claim.
It prices the exact input members and call order in the retained 0693 real-deck
capture trace using the current default `process_ooxml` implementation.
The 0694 change altered name ownership, not the capture call topology.
`prepare.py` verifies that historical topology's source and corpus bindings.
`topology-binding.json` additionally records identical complete PPTX Git trees
from `git rev-parse <revision>:crates/litchi-pptx` for commits 0693 and 0694.

Run from a fresh disposable checkout containing this packet and the recorded
production revision. Preserve this packet's raw results elsewhere before
reproduction; drivers write new receipts. The standalone Cargo.lock pins the
probe dependencies. The first offline build needs those dependencies already
in Cargo's cache; otherwise fetch the locked dependencies explicitly first.
CPU 12, sibling scratch paths and the recorded toolchain are host-specific.
Adapt them when necessary and record the changed setup.

```sh
python3 docs/performance/results/change-0695/prepare.py
python3 docs/performance/results/change-0695/build.py
python3 docs/performance/results/change-0695/measure.py
python3 docs/performance/results/change-0695/profile.py
python3 docs/performance/results/change-0695/summarize-profile.py
python3 docs/performance/results/change-0695/report-metrics.py
python3 docs/performance/results/change-0695/count-elements.py
python3 docs/performance/results/change-0695/validate.py
python3 docs/performance/results/change-0695/audit.py
python3 docs/performance/results/change-0695/cleanup.py
python3 docs/performance/results/change-0695/seal.py
```

The probe loads each distinct sequence path once and reuses that allocation
for repeated entries. Native timers include default capability construction,
MCE processing and output destruction. Input loading, output identity hashes,
printing and timing-vector allocation are outside. The 200 samples in each
of four process legs contain four sequence executions each, following ten
warmup batches; samples are divided by four before summarization. Consequently
p95/p99 describe four-execution batch averages, not individual call tails.
Odd legs reverse case order. Every raw sample and per-leg statistic is retained;
there is no population-confidence or isolated-host claim.

The hardware-counter experiment separately runs three repeats of 10 and 210
samples with ten sequence executions per sample and no timed warmups. A
startup-subtracted per-sequence slope divides the counter difference by 2,000.
This is an estimate affected by process initialization and shared-host noise;
raw values, event availability and counter running percentages are checked.
The slopes still include sample-vector handling and one formatted output row
per ten executions; subtraction does not remove work proportional to samples.
Perf sampling uses a separate invocation. Neither hardware-counter invocation
wall time nor perf-instrumented samples are native timing evidence.

All measurements are warm isolated MCE processing. They exclude package open,
other semantic parsers, edits, publication, serialization and I/O. The full
sequence preserves MCE call order but omits intervening consumers and their
allocation lifetimes. Its times cannot be substituted for capture latency.
Ratios within this experiment rank kernel work; they do not prove an achievable
end-to-end speedup. No allocation, RSS, cold-cache, concurrency or new native
Office interoperability claim follows from this batch.

The initial probe build failed because the pinned SHA-256 result type did not
implement hexadecimal formatting. `build-initial.*` retains that failure;
the successful frozen probe formats digest bytes explicitly, outside timers.
Production sources were never modified. `build.json` binds every tracked
crate Rust source and Cargo manifest, probe inputs, binary and build log.
The cleanup receipt records removal of only this batch's target/profile scratch
and preservation of the workspace lock.

The bounded [next investigation](next-investigation.md) records the source seam
and required A/B proof; its namespace-owner savings remain a hypothesis.
