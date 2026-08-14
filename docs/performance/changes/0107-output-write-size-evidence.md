# Change 0107: logical output write-size evidence

Date: 2026-08-14

Status: correctness and schema evidence only; no performance claim

## Scope

The standalone `tools/perf-baseline` harness now serializes a fixed
`sink.write_size_buckets` object alongside the existing `accepted_bytes`,
`write_calls`, and `largest_write` fields. The buckets are:

| Field | Inclusive logical `Write::write` length |
|---|---:|
| `bytes_0` | 0 |
| `bytes_1_to_512` | 1–512 |
| `bytes_513_to_4096` | 513–4,096 |
| `bytes_4097_to_16384` | 4,097–16,384 |
| `bytes_16385_to_65536` | 16,385–65,536 |
| `bytes_over_65536` | greater than 65,536 |

The distribution is updated at the same accepted-write point as
`write_calls` and `largest_write` for the bounded memory, hashing/discard,
windowed hashing, and seekable counting sinks. A write rejected by a sink
limit or conversion check increments neither `write_calls` nor a bucket.
Accepted zero-length writes are represented explicitly in `bytes_0`.

## Compatibility and invariants

This is an additive schema-v1 field. `SCHEMA_VERSION` remains `1`, the default
case set remains 36 cases and 198 records, and the regression policy and case /
corpus manifest are unchanged. The comparator compatibility suite proves that
an unclassified additive sink histogram in a current report does not change
the six existing compared metrics or fail comparison against an older report.

The fixed Rust test exercises every boundary value, checks that all six bucket
counts sum exactly to `write_calls`, verifies JSON serialization, and proves a
rejected write leaves the counters unchanged. CI repeats the bucket-key and
sum invariant for the default, native OLE2, semantic, ODF, and RTF
smoke/release matrices; the two existing exact one-write sink assertions now
include the expected `bytes_over_65536: 1` distribution.

## Claim boundary

These counters describe logical bytes accepted by the harness sink's
`Write::write` calls only. They are not syscall counts, disk-I/O sizes, memory
copy sizes, compressed or decompressed sizes, latency, throughput, allocation,
RSS, cache, or physical-storage evidence. No production crate or public CRUD
API is changed, and no speedup or resource reduction follows from this report.

## Verification

The focused comparator suite (`python3 -m unittest tools.test_perf_compare`)
passes, including the additive-field test. `cargo fmt --manifest-path
tools/perf-baseline/Cargo.toml -- --check` passes. The standalone harness
focused sink tests and the strict harness/deprecation checks remain required
CI gates alongside the existing non-iWork workspace validation.
