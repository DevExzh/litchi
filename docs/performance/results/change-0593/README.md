# Evidence packet: change 0593

Reuse the open-time relationship proof in the OPC publication plan, decide the
content-types manifest from provenance before building it, settle identical
payloads on the pointer before `memcmp`, and buffer the atomic temporary file.

Record: [`../../0593-opc-publication-pristine-members.md`](../../0593-opc-publication-pristine-members.md).
Implements SAVE-1 and SAVE-2 of [0587](../../0587-remaining-opportunity-survey.md).

## Provenance

| field | value |
| --- | --- |
| base commit (before leg) | `08d968f8ec7db27cf1187d01911fd08b9d014d91` |
| branch | `perf/0593-opc-publication-pristine-members` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0 |
| build | `cargo build --release --locked` (workspace `[profile.release]`: `lto = true`, `panic = "abort"`) |
| pinning | every measured process `taskset -c 13` |
| host state | seven other agents were building and measuring concurrently; run-window load average 35–44 (`timing/window.txt`) |

Binary SHA-256:

| binary | sha256 |
| --- | --- |
| before `examples/tabs` | `d261b572d743a374a5e0aab32061e65306f3ce635b1d5166b8b0238866028e61` |
| after `examples/tabs` | `e657a1f4723c5afb5e84d9750e8839d4af8926064999bc1b0df52d663135b8de` |
| before `examples/edit_cells` | `0f0c21b7e0c1aa3a6d921547a41ecb36b1646476007567ef9400842d34d37c6f` |
| after `examples/edit_cells` | `11c708c8113e397ac5a6ccebc02323a2cd4af1cc98ec9163b78e165bf8a6a6b2` |
| before `opc-save-probe` | `4b26334814ff33d0bd31991d239ce66b08b75ec1ce94a13fee0a031d827d6ab4` |
| after `opc-save-probe` | `ee4cb04a2d7d19a5f4bc1aa66c7e3d7d839c1f599d35e473cd737f8a8cad811c` |

## Contents

| path | what it is |
| --- | --- |
| `probe/src/main.rs` | the scratch publication probe; `tools/perf-baseline` has no ordinary-save (Path A) selector, so this drives `OpcPackage::open` → mutate → `PackageWriter` directly |
| `probe/Cargo.toml.example` | the probe manifest; point the `litchi-opc` path dependency at the leg being measured and give each leg its own `CARGO_TARGET_DIR` |
| `corpus-before.txt`, `corpus-after.txt` | 336 OOXML fixtures × 4 scenarios = 1,344 rows, each the published length and SHA-256 or the typed error |
| `corpus-diff.txt` | empty: the two corpus runs are byte-identical |
| `editor-oracle.txt` | `tabs` and `edit_cells` over three fixtures × four operations on both legs; digests and error messages |
| `syscalls/strace-before-tabs-hide.txt` | `strace -s 0` of the real `save(path)` route, before leg (531 writes to the temporary file) |
| `syscalls/strace-after-tabs-hide.txt` | the same, after leg (14 writes, same bytes, same digest) |
| `counts/cg-*-tabs-hide-incl.txt`, `counts/cg-*-edit-cells-incl.txt` | `callgrind_annotate --inclusive=yes` for the two real editor saves |
| `counts/cg-*-tabs-hide-self.txt`, `counts/cg-*-edit-cells-self.txt` | the self-cost annotations behind the per-function rows in the record |
| `counts/isolation-summary.txt` | computed per-publish Ir, cycle and instruction deltas for every isolation pair |
| `counts/perf-isolation-raw.txt` | raw `perf stat -x, -r 5` cycles and instructions for the M = 20 and M = 220 legs |
| `counts/callcount.py` | the callgrind call-count extractor used for the audit and reserialization counts |
| `timing/publish-*.txt` | per-sample publish nanoseconds, A1 B1 B2 A2, 2,000 samples per leg-run |
| `timing/savepath-*.txt` | per-sample `PackageWriter::write(path)` nanoseconds, A1 B1 B2 A2, 100 samples per leg-run |
| `timing/window.txt` | wall-clock start and end of the timing window with the host load average |
| `timing/stats.py` | the percentile and A/A-floor calculator |
| `run-gates.sh` | the gate runner that produced `gates.txt` |
| `gates.txt` | the tail of every gate |
| `decision.json` | the machine-readable decision, accepted evidence, accepted costs and known gaps |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE |

## Replaying

The before leg is a detached checkout of the base commit; the after leg is the
branch worktree. Both were built with `cargo build --release --locked`, each
with its own `CARGO_TARGET_DIR`.

```sh
# counts
strace -s 0 -e trace=write,fsync target/release/examples/tabs \
  test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xlsx OUT.xlsx Home hide
valgrind --tool=callgrind --cache-sim=no --branch-sim=no --callgrind-out-file=cg.out \
  opc-save-probe savebench <fixture> pkgrels 12      # pair with 2, difference, divide by 10
taskset -c 13 perf stat -r 5 -e cycles,instructions opc-save-probe savebench <fixture> pkgrels 220

# oracles
opc-save-probe corpus test-data > corpus-<leg>.txt   # then diff the two legs

# timing
taskset -c 13 opc-save-probe time <fixture> pkgrels 200 2000        # publish only
taskset -c 13 opc-save-probe timesave <fixture> pkgrels OUT 10 100  # atomic save(path)
```

`timing/stats.py` reads `timing/publish-*.txt` and prints p50, mean, p95, p99 per
leg, the A/A floor from the A1/A2 pair, and the pooled paired deltas in both
directions.

## What this packet does not contain

No cold-cache run, no peak-RSS measurement, no real-device fsync distribution,
no DOCX or PPTX semantic-editor save through Path A (no such example exists —
see 0587 §4), and no `tools/perf-baseline` selector. The p95 and p99 rows in
`timing/` are retained but not interpreted: the host's A/A floor at p99 in this
window was −38.54% to +99.40%.
