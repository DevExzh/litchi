# Change 0346: FileSource lock attribution and candidate rejection

Date: 2026-08-31

Status: diagnostic control smoke retained; no production candidate applied

Performance claim: none

## Decision

No production candidate was applied in this batch. The current-head control
reports and a standalone lock/fingerprint microbenchmark were retained to
bound the hypothesis that replacing the `std::sync::Mutex` used by
`FileSource` would materially improve the XLS source-backed path. The
microbenchmark does not justify a production lock replacement: the projected
whole-operation gain is only 0.36-0.40%, below the investigation threshold,
and no XLS/CFB source fence, freshness boundary, or public API changed.

The batch is diagnostic only. It does not reopen the rejected operation-scoped
freshness session from changes 0279-0280, and it makes no XLS, CFB, latency,
allocation, RSS, physical-I/O, or iWork claim.

## Control protocol

The control is HEAD `3a2926f8a`, using the release
`xls_source_attribution` binary with 8,565,704 bytes and SHA-256
`e3a6744a0ebe720dbbd583ddd5f5d82fb2ca007f4e73c4dcf9667b82cc8b0fe4`.
Collection used Rust 1.95, CPU 2, one serialized worker, warm cache, three
warmups, and 50 retained samples per selector. Each selector ran in its own
process; warmups and retained samples shared that process.

The selectors, in order, were:

- `file-source/open`
- `file-source/list`
- `file-source/one-cell`
- `atomic-file/open`
- `atomic-file/list`
- `atomic-file/one-cell`

The fixed corpus is
`test-data/ole/xls/ConditionalFormattingSamples.xls`, 1,402,368 bytes,
SHA-256 `d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7`.
Its Workbook stream is 1,314,225 bytes with SHA-256
`99305abd97f40bfc2fa4c052701bbebc971c1feb12278e8b76ecfbaca777676f`.
The selected coordinate is worksheet index 1, row 1, column 0.

| Mode | Operation | Elapsed p50 (ns) | Elapsed mean (ns) | Version calls | Version ns/call (mean) |
| --- | --- | ---: | ---: | ---: | ---: |
| `file-source` | open | 487742 | 494581.66 | 1266 | 174.802733 |
| `file-source` | list | 488428 | 505194.64 | 1266 | 179.829795 |
| `file-source` | one-cell | 651062 | 657285.02 | 1813 | 174.532620 |
| `atomic-file` | open | 293089 | 300644.34 | 1266 | 26.050047 |
| `atomic-file` | list | 293304 | 302483.28 | 1266 | 25.881216 |
| `atomic-file` | one-cell | 381511 | 388565.04 | 1813 | 26.780331 |

Open/list each performed exactly 655 reads and 567,685 logical bytes; the
one-cell selector performed 921 reads and 569,398 logical bytes. Every report
has one `len` call, zero seeks, and a stable source version. The semantic
projection contains 16 worksheet names and `Products1!A2` is
`string:4:Date`.

Raw reports and the machine-readable summary/manifest are retained under
[`results/change-0346/`](../results/change-0346/).

## Lock/fingerprint probe

The standalone probe in
[`probes/0346-file-source-lock/`](../probes/0346-file-source-lock/) uses
`parking_lot 0.12.5` and compares the existing standard mutex with a
`parking_lot` mutex. Each measured block performs 200,000 iterations of a
mutex-protected file metadata fingerprint check. It records 40 blocks, with
orders 0 through 3 repeated and the first implementation alternated. The
result is 8,000,000 measured calls per implementation after one unrecorded
warm-up block for each implementation.

| Implementation | Mean (ns/call) | P50 (ns/call) | P95 (ns/call) |
| --- | ---: | ---: | ---: |
| `std` | 155.47 | 155 | 161 |
| `parking_lot` | 154.03 | 153 | 158 |

The raw TSV is
[`file-source-lock-probe.tsv`](../results/change-0346/file-source-lock-probe.tsv).
The result is a lock/fingerprint diagnostic, not a whole-operation
measurement. The projected whole-operation gain is 0.36-0.40%, so no
production lock substitution is retained.

Reproduction is serialized and CPU-pinned:

```sh
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0346/probe-target \
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 taskset --cpu-list 2 \
cargo run --release --locked \
  --manifest-path docs/performance/probes/0346-file-source-lock/Cargo.toml -- \
  test-data/ole/xls/ConditionalFormattingSamples.xls
```

## Gate and next step

This record intentionally has no candidate side, no A/B/B/A comparison, and
`performance_claim: none`. It cannot support an accepted latency, tail,
memory, allocation, or physical-I/O result. No source freshness fence or
CFB/XLS semantic behavior was changed. iWork is out of scope. The next XLS
performance batch must target a different measured design rather than revive
this lock substitution or the unchanged 0279 candidate.
