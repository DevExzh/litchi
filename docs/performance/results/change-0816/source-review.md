# 0816 source review

**Disposition: pass after the formatted handoff.** This is a static
review of the uncommitted 0816 harness handoff at `c8ff2b9f65`. It reviewed
`tools/perf-execution/src/main.rs` and its README only. No Cargo command,
compiler, binary, benchmark, workload, profiler, or network operation was
run.

The implementation has the intended shape: new source settings are parsed,
the same provider is passed to the CFB and ordered-Parts public routes, the
provider caps each returned prefix and sleeps inside `ReadAt::read_at`, and
the existing finite execution context, task floor, output verification, and
permit-release receipt remain in the sample path. The source observer and its
atomics are behind `#[cfg(feature = "source-metrics")]`; the ordinary binary
path retains no observer counters. Priming resets the observer after its
preload and the timed operation is observed before teardown. These points
match the 0816 protocol.

## EOF and overflow review

`InMemoryReadAt::read_at` returns zero for a representable offset at or beyond
the source length, and the capped path never returns bytes beyond that range.
On a narrow host, converting an unrepresentable `u64` offset to `usize`
retains the existing explicit `InvalidInput` error. The test now treats the
`u64::MAX` case portably: it expects EOF where the value is representable and
the conversion error where it is not. Neither path sleeps. This preserves the
harness's pre-existing overflow behavior without changing production source
semantics, while still checking ordinary EOF and empty output.

## Invariants confirmed after those corrections

The following source decisions are consistent with the frozen protocol:

* `--source-max-read-bytes` defaults to zero and accepts `0..=1,048,576`;
  `--source-delay-us` defaults to zero and accepts `0..=100,000`. Duplicate or
  unknown options fail, and nonzero source settings are rejected for OPC,
  whose `from_bytes` route does not use the external `ReadAt`.
* The count is bounded by the caller buffer, remaining bytes, and the positive
  cap. A zero cap leaves the normal range uncapped. A nonempty successful read
  sleeps for the configured duration while still inside `read_at`; zero-length,
  EOF, and error calls do not sleep. `len()` and `version()` are unaffected.
* The provider copies only the valid prefix, returns the prefix length, and
  leaves subsequent calls able to reconstruct the complete payload. It does
  not fabricate bytes or a zero-progress non-EOF result.
* CFB and Parts pass the configured provider for every fresh sample and for
  both the priming and timed operations. Source construction, metadata setup,
  priming, verification, observer snapshot, and drops remain outside the wall
  interval; the delay itself is inside the timed `ReadAt` call. Observer reset
  occurs after priming and before timing.
* Feature-gated observer counters account for logical calls, requested bytes,
  returned bytes, short reads, active reads, maximum simultaneous reads, and
  request-size buckets. The active count is decremented on both success and
  error paths. The normal build reports unavailable metrics without compiling
  those counters into the provider.
* The route-level byte/order/SHA-256 oracle remains mandatory. Worker and I/O
  permits are checked after the result, session, package, and context drop;
  cumulative CPU-task usage remains checked against its finite ceiling. The
  source arm does not alter budget limits or bypass public admission.
* Report schema selection is intentionally conditional: local `0/0` uses
  `litchi.execution-baseline.v1`; a nonzero cap or delay uses
  `litchi.execution-range-baseline.v1`. `ReportConfig` records both source
  fields in all arms. The packet reader must check this pairing explicitly;
  the two schemas are not assumed byte-identical merely because their payload
  fields overlap.

The formatted handoff and portable EOF assertion satisfy the static source
review. Root may proceed to its separately owned quality/build sequence; this
review itself performed no build or execution.
