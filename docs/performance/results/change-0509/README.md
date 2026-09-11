# 0509 evidence: bounded ODT sink paragraph buffer reuse

The control is commit `0455d9abe479bef8980a403f0349ae2a95b254e3` and its
binary is byte-identical to the final 0508 control. `before/` retains its
source/build bindings and profiles. Root source/build artifacts bind the
candidate, while `after/` retains candidate measurements. Both use the
unchanged full baseline harness and deterministic ODT semantic generator.

`admission.json` records the fresh heaptrack allocation evidence before the
production edit. `profile-scope.md`, `profile-summary.json`, raw heaptrack zstd
streams, filtered stack reports, and raw Callgrind/annotations distinguish
allocation count and instruction references from native elapsed time.

Run `python3 -B docs/performance/results/change-0509/verify.py` from the repo
root to replay sample statistics, source/corpus/output/sink identities,
build/gate receipts, profiles, paired deltas, and repeat drift after temporary
binaries have been cleaned. `SHA256SUMS` binds retained artifacts separately.
The verifier requires the recorded final Rust/TOML/lock sources; it does not
pretend historical binaries can be rebuilt from a later source tree.

Reproduction requires checking out the control and applying `candidate.patch`
for the candidate, using the pinned Rust toolchain and environment in build
receipts. `run.py build` uses an exclusive owned target under
`/tmp/litchi-goal-0509`; preserve each resulting binary in `before/` or `after/`
beneath that scratch root. `capture.py timing` runs 500 samples per each of
three shapes in A1/B1/B2/A2 order, with ten warmups and CPU affinity 2. Paths
and preexisting exclusive artifacts must be relocated for a new capture.
`capture.py heaptrack` and `capture.py callgrind` reproduce candidate profiles;
control commands are retained in the raw profile headers/logs. Instrumented
measurements never enter the native latency comparison.

`gates.py` runs crate formatting, ODT tests/Clippy/rustdoc and full harness
library tests/Clippy serially. Root separately runs the boundary and unchanged
strict claim registry checks. Read `source-review.md` and `adr-review.md` for
semantic/resource review. `claims.log` retains the initial missing-evidence-root
invocation error; `claims-final.log` is the corrected passing strict check.

The initial large p99 flag is retained in `initial-summary.json`. The post-hoc
`capture.py tail-followup` run uses 2,000 samples per large-only child, adding
8,000 native samples; its admission and raw reports are retained. `summary.json`
includes both sequences and all adverse/drift flags. See the change record for
the acceptance limits, including the absence of a tail-latency improvement.

`plan.json` retains the initial profiling plan; `acceptance.json` records the
final scoped disposition. `cleanup.json` records removal of the owned target;
`replay-after-cleanup.json` matches the final summary without those binaries.
