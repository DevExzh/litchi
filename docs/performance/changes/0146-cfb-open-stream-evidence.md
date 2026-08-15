# Change 0146: CFB MiniFAT `open_stream` evidence

Date: 2026-08-16

Status: harness and correctness evidence added; no performance claim.

## Scope

This tranche adds 12 opt-in selectors to `tools/perf-baseline` for the
`SharedOleFile::open_stream` path introduced by production optimization
`3375729f4`. The selectors are shared-only so the same harness can be built
against the clean parent revision and the candidate revision for a later
ABBA comparison. They cover both `many-small` (256 root siblings) and
`wide-root` (2,048 root siblings), with MiniFAT targets of exactly 36 and
4,095 bytes. Each target has one-shot, repeat-3, and sequential repeat-8
operations, plus a matched deterministic-delay simulator selector.

The selector names are:

```text
cfb_open_stream_{mini,mini_4095}_shared_{one_shot,repeat,repeat8}
cfb_open_stream_simulated_{mini,mini_4095}_shared_{one_shot,repeat,repeat8}
```

Repeat-8 means eight sequential `SharedOleFile::open_stream` calls against
one immutable owner. It does not call `bulk_read`, use concurrent workers, or
enumerate all root streams. Ineligible-root, FAT, bulk, and
concurrency controls remain follow-up work; this tranche makes no claim about
those workloads.

This is generic CFB/OLE2 substrate evidence only. It does not establish native
DOC/XLS/PPT semantic CRUD or performance, and it makes no OOXML, ODF, RTF, or
iWork claim.

## Evidence contract

The production selectors record, for every measured invocation and aggregate
sample:

- target/root/sector/shape identity, including the declared
  `root_ministream_bytes` and MiniFAT target start sector;
- exact output byte counts and SHA-256 hashes for every returned output;
- public `SharedOleFile::source_version` observations before and after the
  operation, plus typed refusal of a missing stream path;
- logical positional source events as raw
  `(offset, requested_len, returned_len)` triples in call order;
- open, operation, and checked `open + operation` timings, with source
  snapshots, refusal checks, output verification, and hash generation outside
  the measured duration;
- per-invocation timings, calls, bytes, request sizes, ranges, and observed
  repeated source bytes.

The positional `read_bytes` fields count returned bytes; raw events retain
requested and returned lengths separately. Simulator request-size vectors
record requested lengths while `physical_request_bytes` records returned
bytes.

The simulator records separate open and per-operation logical phases, raw
physical range events, configured latency/bandwidth parameters, request-size
buckets, and a deterministic service-floor calculation for each phase. Cache state is private in the
production API, so `cache_state_diagnostic` is explicitly an inference from
source counters rather than a private-state assertion. `root_ministream_bytes`
comes from the public root directory entry. `expected_direct_physical_range`
is populated only when the first measured operation issued exactly one
target-sized source event; it is `None` for a materializing baseline shape.
Empty repeat-3/repeat-8 cache-hit phases are represented by all-zero counters,
empty event vectors, and a zero service floor; non-empty phases retain strict
validation.

Current-tree evidence has the following expected identity values for the
candidate direct path (the clean parent is allowed to report the full-root
range instead):

| shape | target | root Mini Stream `R` | direct physical start `P` |
| --- | ---: | ---: | ---: |
| many-small | 36 | 261,184 | 261,632 |
| many-small | 4,095 | 265,216 | 261,632 |
| wide-root | 36 | 2,096,192 | 2,096,640 |
| wide-root | 4,095 | 2,100,224 | 2,096,640 |

For a current candidate repeat-3 sample, per-invocation logical bytes are
`[L, R, 0]` and calls are `[1, 1, 0]`, where `L` is the target length. The
parent materializing baseline is accepted by the same runner and may report
`[R, 0, ...]`; these are attribution controls, not a speedup assertion.

## Verification boundary

No latency, throughput, allocation, RSS, physical-I/O, decompression, cold
cache, or release-performance result is claimed here. A performance statement
requires the identical harness on clean release builds of the parent and
candidate revisions, CPU-pinned A1/B1/B2/A2 ordering, and the existing release
ABBA reporting rules. The JSON evidence is intended to make that later
comparison auditable and to prevent a direct MiniFAT read from being confused
with a bulk or concurrent workload.

Focused validation for this change:

```text
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline --no-default-features cfb_open_stream
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline --no-default-features
RUSTFLAGS='-D warnings -D deprecated' cargo clippy --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline --no-default-features -- -D warnings
```
