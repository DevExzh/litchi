# 0012 — open-route, allocation and clone measurements

Evidence for record [0012](../../0012-opc-shared-package.md) §6. This is a
quick release-mode loop on a shared host, not the S1/S10 protocol of record
0003, and it supports no speedup claim.

## Files

- `bench/`: the measurement program, a standalone crate (its own
  `[workspace]`) with path dependencies on `litchi-opc`, `litchi-core`,
  `litchi-docx` and `litchi-pptx`. It was built from a copy in the session
  scratchpad whose manifest named the same crates by absolute path; it is
  otherwise identical to this file. Its routes are listed at the top of
  `bench/src/main.rs`.
- `abba.sh`: the driver, as run. `S` is the scratchpad copy holding the two
  binaries; `T` is the worktree's `test-data/`. Arguments: CPU, rounds, output.
- `abba_summary.py`: the per-run summary (median and range of per-process
  p50s per leg).
- `abba-run1.jsonl` (5 rounds, CPU 26) and `abba-run2.jsonl` (10 rounds,
  CPU 12): one line per process with round, leg, one-minute load, route,
  fixture path, sample count, checksum, and p10/p50/p90/mean in nanoseconds.
- `valgrind-before.txt`, `valgrind-after.txt`: `valgrind --tool=memcheck
  --leak-check=no <binary> <route> <fixture> 1 <n>` heap summaries for n = 10
  and n = 20; the clone-only route with `CLONES=0` and `CLONES=1000`; and
  `opc-clone-ns-managed <fixture> 0 <b>` with one and two batches of 1,000
  clones.
- `clone-ns.txt`: `opc-clone-ns` and `opc-clone-ns-managed` lines.

## Binaries (not committed)

Built with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0 cargo build
--release --offline` in the scratchpad copy, rustc 1.95.0, with the
workspace `Cargo.lock` copied in to pin third-party versions.

| Leg | Source tree | `bench/src/main.rs` | SHA-256 |
|---|---|---|---|
| before | `67dfe7a171` (unmodified `uo/m0a`) | without the `opc-clone`, `opc-clone-ns` and `opc-clone-ns-managed` routes, and without the `litchi-core` dependency | `415b5eb948a88494d049e350913e73550481bd957ece3cfe6a4d048a34d235cb` |
| after | this change | with `opc-clone` and `opc-clone-ns` | `5cf46a913dbc191061ec6ddea77bbd74991c69edbe8dee2befaebbcaef1ddc58` |
| after-clone | this change | this file (adds `opc-clone-ns-managed`) | `9b6bf7d7f22740a4b71260e73d9224c7ffa9ed36a207d82a728f14cddf470a5f` |

The timed open and full-text routes are the same code in all three files; the
added routes are matched after them and never run in the timed processes.
`after-clone` was used only for `clone-ns.txt` and the managed clone count.

## Commands

```sh
abba.sh 26 5 abba-run1.jsonl && python3 abba_summary.py abba-run1.jsonl
abba.sh 12 10 abba-run2.jsonl && python3 abba_summary.py abba-run2.jsonl
for n in 10 20; do valgrind --tool=memcheck --leak-check=no bench-<leg> <route> <fixture> 1 $n; done
for c in 0 1000; do CLONES=$c valgrind --tool=memcheck --leak-check=no bench-after opc-clone <fixture> 1 1; done
for b in 1 2; do valgrind --tool=memcheck --leak-check=no bench-after-clone opc-clone-ns-managed <fixture> 0 $b; done
taskset -c 12 bench-after-clone opc-clone-ns <fixture> 100 2000
taskset -c 12 bench-after-clone opc-clone-ns-managed <fixture> 100 2000
```

Host: AMD EPYC 9R45, 32 vCPUs, Linux 7.0.0-1013-aws, shared with other agents
(one-minute load 1.2–3.9 during the timed runs).
