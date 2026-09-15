# Evidence: change 0615, execution-context completeness and the coexistence counts

Change record:
[`0615-execution-context-completeness-design.md`](../../0615-execution-context-completeness-design.md).
Proposed ADR: [`docs/adr/0031-execution-context-budgets.md`](../../../adr/0031-execution-context-budgets.md)
— **proposed, awaiting human review, not accepted, not in the accepted table.**

Disposition: **design only, plus a proposed ADR.**
`performance_claim: none`. **No file under `crates/` was modified.** No timing
was measured and none is claimed; every number here is a deterministic count of
threads or of concurrent positional reads.

## Contents

| Path | What it is |
| --- | --- |
| `probe/main.rs`, `probe/Cargo.toml` | The scratch probe. A separate Cargo project (`[workspace]` stanza detaches it) with path dependencies on this worktree's `litchi-core`, `litchi-cfb`, `litchi-opc` and `soapberry-zip`. It builds a four-member in-memory OPC archive and a four-stream CFB file, one MiB each, drives all three parallel sessions from **one** hierarchical `Budget` root, and counts live OS threads from `/proc/self/task` plus simultaneous `ReadAt::read_at` calls. Usage: `probe0615 <per-read delay µs> <workers>`. |
| `probe-w1-delay500.txt` | Raw output, `workers = 1`, 500 µs per-read instrument. The one width at which the three sessions compose — because none of them goes parallel. |
| `probe-w2-delay500.txt` | Raw output, `workers = 2`. |
| `probe-w4-delay500.txt` | Raw output, `workers = 4`. |
| `probe-w8-delay500.txt` | Raw output, `workers = 8`. The OPC scoped wave caps at four (the part count), which is why the thread total is 20 and not 24. |
| `probe-repeats.txt` | Five repeats at each of widths 1, 2, 4 and 8, one row per run, with the instrument. Every thread count is identical in all five repeats at every width; only the process-wide read-overlap figure varies. |
| `probe-w4-nodelay.txt` | Raw output, `workers = 4`, **no** instrument. Identical thread counts; per-session read concurrency reports 1, because a `read_at` over an in-memory `Vec` is a `memcpy` and four workers on one pinned core do not overlap. This is the control that establishes the thread table, not the read table, as the load-bearing evidence. |
| `gates.txt` | The tail of every gate run. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree area. |

## The instrument

`probe-w*-delay500.txt` apply a fixed 500 µs delay inside `read_at`. It is a
**deterministic test instrument** that widens the window in which concurrent
reads can be observed; it models no production storage, scheduler, allocator or
decompression latency. This is the same role change 0088's fixed 10 ms source
delay plays, and the same disclaimer applies. The thread counts are identical
with and without it (`probe-w4-delay500.txt` against `probe-w4-nodelay.txt`).

## What the outputs say

Session-owned worker threads at peak, from one budget root and one
`ExecutionLimits { workers: W }`:

| `workers` | after `OpenSession::new` | after CFB bulk read | peak during OPC part batch | peak, all three live | session-owned at peak |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 1 | 1 | 1 | 4 | 0 |
| 2 | 3 | 5 | 7 | 10 | 6 |
| 4 | 5 | 9 | 13 | 16 | 12 |
| 8 | 9 | 17 | 21 | 24 | 20 |

Baseline is one thread. The "all three live" column is sampled from inside a
worker's own `read_at` and includes the probe's main thread and three driver
threads; the last column subtracts those four and equals `2W + min(W, tasks)`.

Concurrent `read_at` calls, with the instrument:

| `workers` | CFB bulk read | OPC part batch | process-wide, both sources at once |
| ---: | ---: | ---: | ---: |
| 1 | 1 | 1 | 2 |
| 2 | 2 | 2 | 4 |
| 4 | 4 | 4 | 7–8 |
| 8 | 4 | 4 | 7–8 |

The per-session columns are `min(W, independent work)` in all five repeats. The
process-wide column is a **lower bound**: it is the largest overlap the two
independently scheduled sessions' read windows happened to achieve in a run, not
a proof of the maximum reachable; it hits the arithmetic sum of 8 in three of
the five repeats at `W ≥ 4`. The separately retained single run
`probe-w4-delay500.txt` recorded 3 rather than 4 for the CFB session, which is
what a lower-bound observation does; it is retained rather than discarded.

`threads_after_250ms_settle` is 1 at every width: the pools do not leak, they
are held for their session's lifetime and reaped asynchronously on drop.

## Provenance

- Base commit: `1e4198321` (change 0606,
  `perf(ppt): span the retained stream instead of copying every record payload`).
- Branch: `perf/0615-execution-context-completeness-adr-draft`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Eight agents
  were building concurrently throughout; every probe run was pinned with
  `taskset -c 10`. Pinning does not move a thread count; it is why the
  no-instrument leg reports read concurrency of 1, and that leg is retained for
  exactly that reason.
- Toolchain: `rustc 1.95.0 (59807616e 2026-04-14)`.
- Probe binary, `cargo build --release` (`debug = 1`), built with
  `CARGO_TARGET_DIR` outside the repository:
  `6e59079909acef25f0739b8e003c4e9f90991afeb38c9d93975dbe6836b58a6c`.
  Every retained output in this directory was produced by that one binary.
- The probe was built from a working copy at
  `/home/zhuhe/code/litchi-worktrees/0615-probe` (`src/main.rs` plus
  `Cargo.toml`); `probe/main.rs` here is that file byte-for-byte. The working
  copy and its target directory were deleted after the runs — see
  `cleanup.json`.

## Replay

```sh
mkdir -p /tmp/probe0615/src
cp probe/Cargo.toml /tmp/probe0615/Cargo.toml     # edit the four path dependencies
cp probe/main.rs    /tmp/probe0615/src/main.rs
cd /tmp/probe0615
CARGO_TARGET_DIR=<somewhere on disk> cargo build --release
for w in 1 2 4 8; do taskset -c 10 <target>/release/probe0615 500 "$w"; done
taskset -c 10 <target>/release/probe0615 0 4
# probe-repeats.txt is five runs of the first line, tabulated
```

The path dependencies in `probe/Cargo.toml` point at
`/home/zhuhe/code/litchi-worktrees/0615/crates/...`, which was this change's
worktree; a replay must repoint them at a checkout of `1e4198321`.

## What is not here

No timing leg, no A/B or A/A pair, no allocation or RSS profile, no callgrind
run, no harness capture. None was taken: this batch's scope is documentation
plus the probe, and its record registers no claim. The scaling figures quoted in
the record's §4 are **retained** from changes 0009, 0498 and 0499 and are not
re-measured here.
