# Filesystem release repeated evidence

Date: 2026-08-13

Harness and snapshot revision: `5b70f19e01abdcc660a37918742bfbf3a214a02a`

Status: controlled release evidence; no production-performance or comparator claim

## Scope and reproducibility

This run uses the exact five filesystem cases added by the harness:

- `opc_file_eager_open`;
- `opc_file_source_open`;
- `opc_file_eager_one_part_atomic_save`;
- `opc_file_source_one_part_atomic_save`; and
- `cfb_file_same_length_overlay_atomic_save`.

The harness was built from a plain archive of commit `5b70f19e0` into an
isolated `/dev/shm` Cargo target. The binary is release-profile, uses the Rust
system allocator, and is pinned to CPU 2. Each case has three untimed warm-ups
and 30 measured fresh-child samples in each of `warm` and `cold-requested`
states. The source and same-filesystem sibling save destinations are on
`tmpfs`; host evidence reports Linux 6.8.0-101-generic, AMD EPYC 9575F, and
one logical CPU available to each pinned process.

The complete schema-1 report is
[`filesystem-release-0100-5b70f19e.json`](../results/filesystem-release-0100-5b70f19e.json)
and its compact extraction is
[`filesystem-release-0100-summary.json`](../results/filesystem-release-0100-summary.json).
The raw report SHA-256 is
`e62fca57282fbfc8b5bad7f32d04f9caceccff32f6b37c2a237991ceb286b1e7` and the
release binary SHA-256 is
`571d7d1b776dda11779758a60a176aa923879ce0499b45a769403fc42d07e407`.
The raw schema records the release profile in `tool.profile`; the binary hash
was captured by the external runner after the build and is retained in the
compact summary rather than duplicated in the raw report.

The reproducible command shape was:

```text
taskset -c 2 litchi-perf-baseline --warmup 3 --samples 30 \
  --filesystem-cache warm,cold-requested \
  --filesystem-root /dev/shm/<caller-selected-parent> \
  --case opc_file_eager_open,opc_file_source_open,\
opc_file_eager_one_part_atomic_save,opc_file_source_one_part_atomic_save,\
cfb_file_same_length_overlay_atomic_save --json <raw-report>
```

The harness removes its per-run child directory after completion. Absolute
paths are not retained in the report.

## Correctness and physical evidence

All 300 measured samples completed: 150 warm and 150 cold-requested. Every
cold advice request was accepted, every warm sample recorded
`not_requested`, and every parent source re-hash and semantic reopen check
passed. The two OPC save cases emitted the same 16,783,632-byte output on all
samples with SHA-256
`f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009`.
The CFB publisher emitted 16,913,408 bytes with one changed span and SHA-256
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.

The source-backed OPC open performed 13 positional logical reads totaling
1,008 bytes and materialized zero Parts; eager open materialized four Parts.
Source-backed OPC save performed 549 logical reads totaling 16,785,201 bytes
and materialized zero ordinary Parts. CFB overlay performed 2,084 logical
reads totaling 101,751,908 bytes. Eager cases use owned file reads, so their
`ReadAt` counters are zero by design. Process `syscr`/`syscw`, `rchar`/`wchar`,
read/write bytes, current RSS, and peak RSS are retained per sample in the raw
report. On this tmpfs run, process `read_bytes` and `write_bytes` were zero;
that is a filesystem observation, not evidence of zero work on a storage
device.

Child-operation p50 values (nanoseconds) are retained with p95/p99, means,
sample standard deviations, and two-sided Student's-t 95% mean intervals in
the summary. For orientation only, warm versus cold-requested p50 pairs are:

| Case | Warm | Cold-requested |
| --- | ---: | ---: |
| eager OPC open | 18,791,047 | 18,902,756 |
| source-backed OPC open | 218,087 | 218,157 |
| eager OPC one-Part save | 237,448,214 | 235,475,465 |
| source-backed OPC one-Part save | 64,461,139 | 64,486,494 |
| CFB same-length overlay save | 99,061,497 | 99,237,288 |

These are descriptive distributions on this named tmpfs host. They are not a
warm/cold speedup claim.

## Claim boundary and next gate

`cold-requested` means that Linux `posix_fadvise(DONTNEED)` accepted an
advisory request immediately before the timed child operation. It does not
prove that the kernel or device supplied a cold cache, and this run's tmpfs
`read_bytes == 0` confirms that no storage-device conclusion is available.
The data therefore does not support a physical cold-cache claim, a production
latency claim, an allocation claim, or comparator-baseline approval. It also
does not compare hardware or storage devices. Any future cache-state claim
needs a controlled block-backed filesystem and independently verified cache
temperature; any optimization acceptance still needs allocation/peak-memory
instrumentation and a balanced comparator run.
