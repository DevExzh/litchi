# 0362ul55 — composed tooling row

Date: 2026-10-25. The complete current tooling row passes on clean integration
`a10687e3b29583d117381710ecbc4e10685336b0`: 1,687 tests in 110.974 seconds.
The supervising row takes 114.067702 seconds and exits zero. This is the
single selected row 48 of the 53-step gate, not a full-gate result.

This source includes the reviewed required-signature checker, XLSX drawing
consumer migration, Ink final-cap check, tooling environment and lifecycle
repairs, and PPT qualification guards. The normal gate has 53 steps; four
PPT qualification steps are explicit auxiliary selections. The correctness
companion checks changed output, forward replay equal to the committed
target, and inverse replay equal to the original bytes. These checks do not
alter the historical timed workload.

The tooling row explicitly lists 35 modules. W8 lifecycle tests are registered
directly instead of being discovered through an unused imported test class;
the PPT gate-wiring tests are also registered. Before this full run, the
composed lifecycle/chunked-runner suites pass 43 tests, and the projection,
lifecycle and PPT wiring suites pass 44 tests.

## Actual run

```text
env -u PYTHONPATH -u CARGO_TARGET_DIR python3 docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py --root /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration --target (deleted-coord-artifact) --state (deleted-coord-artifact) --only tooling-unit
```

Source HEAD and tracked status remain unchanged. The three observed Cargo
locks remain unchanged; the root lock is approved `4100b1fb…bfaea1`. No Cargo
or native qualification is invoked by this row. Its Cargo trace is empty,
and pre/live/post target size remains zero. Owned processes terminate; the
runner PID 3854373 is gone. Concurrent, separately owned compiler processes
are recorded in the preflight and are not attributed to this row.

All artifacts below remain under external root
`(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `tooling-row-a10687e3b2-20261025T-source-freeze.txt` | 874 | `440eed3c1ab632f1d30ba0bb19f4a5d233c46101d849eaeeff4b64ce519fa83a` |
| `tooling-row-a10687e3b2-20261025T2-wrapper.log` | 4,000 | `6f2bd886a5dfe597003a06f5f4af70cee443fc562978b1cd18a2ab13eff36f90` |
| `tooling-row-a10687e3b2-20261025T2-state/plan.json` | 24,652 | `c8cc371b262d818b343fc40b1dcee9122ee8601a2e32870a40111462f2680ed8` |
| `tooling-row-a10687e3b2-20261025T2-state/state.json` | 12,636 | `1c773dfa373aa6106f27299f980946e60b9b71ca42a863d1a9337e55f4149392` |
| `tooling-row-a10687e3b2-20261025T2-state/48-tooling-unit-attempt-1792886887488743497-3854373.log` | 635,289 | `5459e2ecc88b064dd8bbd1623c26835eea26e1013d6df15693836e6120aa999f` |

Root independently verifies these digests and the terminal receipt. An
earlier setup invocation rejected the target-directory shape before tests;
its immutable `tooling-row-a10687e3b2-20261025T-wrapper.log` is 139 bytes,
SHA-256 `00bb4dc4a524172b2da45e1f6633ae9f3d419d4b73b532e9f8a61e46f8479b19`.
It is not counted as a passed attempt.

## Legacy gate synchronization

The external `(deleted-coord-artifact)` still had an eight-module tooling command.
Only that command is synchronized to the same 35 modules. Shell syntax and
exact argv comparison pass. Its former file is retained externally:
5,024 bytes, SHA-256
`fff9a78f21a2e6f98a05fd0132c75fc7ab435bb6331569eb3e32862596a6f379`.
The updated file is 6,115 bytes, SHA-256
`1ad5e255780c85a7bdc468414d64d4a288f2d484c80dee8fab5872862b732876`.

This result supersedes the historical 1,679-test tooling receipt only for
the composed source above. Raw-codec review returns, final candidate capture,
the remaining full gate, the signed-batch contract decision, and actual
performance qualification remain required. Nothing is pushed.
