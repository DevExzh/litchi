# Execution budget benchmark

`litchi-perf-execution` is a standalone process benchmark for the three
explicit read routes governed by ADR 0031:

* `--route opc` times `OpenSession::from_bytes` over a borrowed OPC archive;
* `--route cfb` times `SharedOleBulkRead::read_streams` over an owned immutable
  positional source;
* `--route parts` times `SourceBackedPackage::read_parts_ordered` over an owned
  immutable positional source.

The command line is:

```text
litchi-perf-execution \
  --route opc|cfb|parts \
  --shape small|large|mixed \
  --workers N \
  --task-floor BYTES \
  --state fresh|primed \
  --samples N \
  --warmup N \
  [--source-max-read-bytes N] \
  [--source-delay-us N] \
  --output PATH
```

`--source-max-read-bytes` and `--source-delay-us` apply only to the CFB and
source-backed Parts routes.  They default to `0`, meaning an uncapped source
and no delay.  A nonzero maximum read is an in-memory short-read simulation and
may be at most 1 MiB; a nonzero delay is applied in microseconds after each
successful nonempty read and may be at most 100,000 microseconds.  Supplying
either option nonzero for `--route opc` is rejected because
`OpenSession::from_bytes` bypasses `ReadAt`.

The generated corpus always has 32 distinct members. `small` uses 4 KiB per
member, `large` uses 256 KiB per member, and `mixed` uses 31 large members plus
one 4 KiB member. OPC members are Deflate-compressed through
`StreamingArchiveWriter`; CFB streams are emitted by `OleWriter`. Corpus
construction, hashing, and output destruction are outside every timed region.

`fresh` constructs the session and package metadata before timing and measures
the first public read/open. `primed` performs and verifies one preload outside
the timed region, drops that result, and measures the next operation on the same
session. For `parts`, the primed result is explicitly a cache-hit control: the
source-backed payload cache is warm. The other primed routes primarily exercise
reused session scheduling state.

Every sample owns a fresh finite hierarchical `Budget`. Workers, positional I/O,
CPU task count, memory, input, output, object, depth, and work limits are finite;
the `ExecutionLimits` task floor is the caller's `--task-floor`, while the
aggregate parallel threshold is fixed at 64 KiB. Resource snapshots are taken
before and after the timed operation and after dropping the output and session.
The harness requires worker and I/O permits to be released at the final drop and
checks that cumulative CPU task use never exceeds its finite limit.

The normal binary has no source-counter atomics in its `ReadAt` implementation.
Build with `--features source-metrics` for a separate diagnostic observer that
records logical calls, requested/returned bytes, short reads, a request-size
histogram, and active/max simultaneous reads. Observer metrics include only the
timed operation after the source counters are reset; they do not establish disk,
network, cold-cache, or physical-read claims. OPC uses `from_bytes`, so external
`ReadAt` metrics are recorded as not applicable for that route.

When source simulation is enabled, every successful nonempty `ReadAt` call
returns no more than the configured maximum (unless the remaining source is
smaller), then sleeps for the configured delay while the observer's active-read
window is still open. Empty, EOF, and error calls do not sleep. The delay models
caller-visible latency in this in-memory harness only; it does not represent a
physical network, filesystem, remote-range, or cold-cache measurement.

CPU time is read with safe `rustix` `ProcessCPUTime` support on Linux. It is
reported as unavailable on targets where that clock is not compiled in. The
process CPU interval is labelled slightly wider than the wall interval because
the two clock reads bracket the operation with separate calls.

The final output path is created only after all samples and byte checks pass;
write failure removes the file created by that run. With the default source
settings its schema is `litchi.execution-baseline.v1`; when either source
setting is nonzero it is `litchi.execution-range-baseline.v1`. Both schemas
contain the exact per-member SHA-256 values, corpus hashes, raw wall/CPU
durations, verification status, observer availability, the exact source
configuration, and resource snapshots. The range schema identifies the
simulated in-memory source scope described above.
