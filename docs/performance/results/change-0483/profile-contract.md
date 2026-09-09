# 0483 excluded process-profile contract

`profile.py` is a custody driver for diagnostic CPU and syscall profiles. It
must run after the accepted normal binary has been built and copied into
`builds/binaries-accepted.json`, and before `capture.py --freeze` writes the
immutable formal protocol. It does not build the binary, modify the Rust
harness, freeze the protocol, or interpret profiler output.

`profile.py` is included in the helper lists in `common.py` and `analyze.py`;
the frozen protocol binds its hash with the other evidence drivers.

Run the driver once with a fresh attempt token. The outer lock must be distinct
from the lock used by `gate.py`, because the driver invokes the gate once per
profiler command:

```sh
flock /home/zhuhe/.cache/litchi-goal-0483/profile-driver.lock \
  python3 -B docs/performance/results/change-0483/profile.py --attempt accepted
```

`gate.py` continues to own `/home/zhuhe/.cache/litchi-goal-0483/cpu.lock` for
each child command. The profile driver deliberately does not acquire that lock
around the loop, which avoids a nested flock deadlock.

## Fixed workload

The driver reads the normal executable and source manifest identity from
`builds/binaries-accepted.json`, verifies the copied executable's byte count
and SHA-256, and uses that same absolute executable for both routes:

```text
/home/.../docx_bounded_tail_append_compare \
  --route materialized|bounded \
  --counts 131072 --samples 30 --warmups 3 --json REPORT
```

The route labels are `materialized` and `bounded`, mapping to the harness's
`materialized_paragraph_copy` and `bounded_plain_text_tail_append` report names.
Each route is a fresh process and is profiled independently. There are three
workload invocations per route (`perf stat`, `perf record`, and `strace`), so
the plan is six runs and at most 180 retained harness samples. The profile
manifest records planned runs/samples separately from the successful workload
reports' actual run/sample totals; unavailable or failed invocations add no
actual sample count.

The timed values remain the harness's route-specific `elapsed_ns` samples, and
those instrumented timings are excluded from the formal evidence. Profiler
counters and stacks cover the complete process, including corpus construction,
untimed independent output/member/semantic oracles, warmups, all 30 measured
harness lifecycles, report JSON serialization, and teardown. They must not be
presented as transaction-only timing or as an operation PMU delta.

## Commands and retained custody

For each route, `profile.py` invokes these commands through separate gate
receipts, in this order:

```text
/usr/bin/taskset -c 2 /usr/bin/perf stat -x, -o profiles/accepted/<route>/perf-stat.csv \
  -e cycles,instructions,branches,branch-misses,cache-misses,page-faults -- WORKLOAD
/usr/bin/taskset -c 2 /usr/bin/perf record -e cycles --call-graph fp -F 99 \
  -o /home/.../.cache/litchi-goal-0483/accepted/profiles/<route>/perf-record.data -- WORKLOAD
/usr/bin/taskset -c 2 /usr/bin/strace -c -o profiles/accepted/<route>/strace.txt -- WORKLOAD
/usr/bin/perf report --stdio --no-children --percent-limit 0 -i RAW_DATA
/usr/bin/perf script -i RAW_DATA
```

`WORKLOAD` is the exact route command above. The report and script commands
are postprocessors over the same route's `perf-record.data`; they are not
claimed as workload samples. Every invocation is sent to `gate.py` with labels
of the form `profile-<route>-<kind>-accepted`. The gate receipt retains the
exact child argv, environment, source-before/source-after manifests,
source-unchanged result, and hashes for its stdout and stderr. The profile
manifest repeats the command argv and binds each receipt path and digest to the
normal binary and source-manifest SHA-256.

The driver checks that `/usr/bin/perf` or `/usr/bin/strace` is an executable
before sending a corresponding command to `gate.py`. An unavailable tool is
recorded as `unsupported` with no gate receipt, no workload sample count, and
no invented metric; this also prevents a missing profiler from leaving an
orphan gate `.started.json` receipt. The temporary `perf-record.data` is
compressed with deterministic gzip after the report and script postprocessors
finish. The compressed artifact is kept
as `profiles/accepted/<route>/perf-record.data.gz`; decompression is checked
against the original byte count and SHA-256 before the temporary file is
removed. `perf-stat.csv`, the harness report for each profiler workload, and
`strace.txt` remain raw route artifacts. The report/script text is retained in
their gate stdout files.

The top-level manifest is `profiles/<attempt>/profile.json` with schema
`docx-tail-append-process-profiles-v1`. It records the attempt, timestamps,
normal binary/build/source identities, helper/gate/common hashes, exact
configuration, route records, output metadata, and a `pass` or `partial`
status. An unavailable profiler, nonzero profiler exit, missing output, or
missing postprocessor data produces `partial`; the driver retains the failed
gate receipt when available and omits any unavailable metric. A repeated
attempt or existing gate/output path is refused rather than overwritten; use a
new attempt token for a new diagnostic capture.

The existing 0479 profile analyzer is not reused automatically. Its parser
expects a different report shape and lifecycle marker; 0483 needs route-aware
postprocessing bound to this CLI and protocol. A later analyzer may consume
these raw artifacts after the route report schema and accepted protocol are
frozen.
