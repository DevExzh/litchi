# Value evaluator diagnostic and performance harness

This candidate-only corpus exercises the new ODS value evaluator and the
production worksheet resolver. It contains 99 value cases and 22 worksheet
cases. Sizes from 1 through 4,096 expose repeated range scans, array growth,
resource refusals, and lazy branch costs. Semantic preflight checks run before
timing; the lazy aggregate cases require one range scan across selected output
cells.

Build from the isolated candidate workspace after synchronizing its sources,
including this harness. Use the existing shared target only when no other
build owns it:

```sh
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-array-target \
TMPDIR=/home/zhuhe/code/litchi-array-tmp \
cargo build --locked --offline --release --manifest-path \
  docs/report/spec-gap-validation-evidence/ods-formula-array-reference-evaluation/performance/value-harness/Cargo.toml
```

Run the resulting binary through `run.py`, using a fresh output directory:

```sh
python3 run.py \
  --binary /home/zhuhe/code/litchi-array-target/release/ods-formula-array-reference-evaluation-value-profile \
  --output /home/zhuhe/code/litchi-array-tmp/value-release
```

The default captures setup, parse, evaluate, and parse/evaluate separately.
Use `--adapter worksheet` for setup, index construction, evaluation with a
reused index, and parse/evaluate. Repeat that capture with `--instrumented`
in a separate output directory to collect worksheet resolver counters.
The fixture resolver in the value corpus always includes instrumentation;
worksheet timing is uninstrumented by default. Counter overhead makes these
different measurement lanes.

`--phase evaluate --warmups 0 --iterations 1` with a debug binary is a quick
semantic diagnostic, not performance evidence. Prior diagnostics exposed
quadratic lazy aggregate reads and work-limit failures at larger sizes.
Adding this harness does not establish that the candidate passes these cases.

The runner pins children to CPU 6 and records raw output, process status,
commands, timing, and executable identity. Retain the compiler version, build
flags, lockfile and complete candidate source manifest with any published
capture; executable identity alone does not identify its build inputs. Avoid
concurrent profiling and record host contention. The existing scalar harness
and its retained baseline executable remain the separate compatibility and
performance comparison for the established scalar API.

Use `--source-identity` and `--source-manifest` to attach the build's production
source identity and manifest hash. These are caller-supplied provenance, not
proof that the executable was built from those files. The runner checks that
the executable and its four harness inputs remain unchanged during capture,
validates complete result fields and success/refusal counts, and exits nonzero
for failed children or invalid captures.

`memory_retained_used_*` samples execution-budget reservations while the result
is retained; it is not a transient peak. Allocator peak-live measurements and
process maximum RSS are separate metrics. `adapter_index_reserved_bytes_*`
reports the retained worksheet index. Borrow and copy counters describe the
instrumented provider, not every copy inside the evaluator. Percentiles from
a single diagnostic iteration must not be presented as tail-latency evidence.
