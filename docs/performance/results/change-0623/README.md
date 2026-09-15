# Evidence packet — change 0623

A read-side local-span accessor in `soapberry-zip`, and change 0577's candidate
(c) — the run-coalesced structural prefetch — in `litchi-opc`. Record:
[`docs/performance/0623-zip-structural-span-accessor-and-prefetch.md`](../../0623-zip-structural-span-accessor-and-prefetch.md).

## Provenance

| | |
| --- | --- |
| base commit | `3156bff3bd37d7d282891e8257944bb967566ff7` (`feat/office-format-completeness`, change 0611) |
| branch | `perf/0623-zip-structural-span-accessor-and-prefetch` |
| commit | see `decision.json` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-3156bff3b`, a read-only detached checkout of the base |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0 |
| build | `cargo build --release --locked`; workspace `[profile.release] lto = true, panic = "abort"` |
| pinning | `taskset -c 17` for every measured process |
| host load | seven other agents building and measuring throughout; every timing window carries its own A/A floor |

Binary identities are recorded inside each timing JSON
(`binary_identity.binary_sha256`), in `oracle/binaries.sha256` for the
differential and in `workbook-timing/binaries.sha256` for the delayed-transport
probe.

## Contents

| path | what it is |
| --- | --- |
| `probe/src/main.rs` | change 0587's request/byte/allocation/observation probe as change 0611 left it, plus one `classify` clause: a read that begins at a member's local header and covers a second member's whole fixed header is a `run-span` |
| `probe/Cargo.toml.template` | the probe manifest with its path dependencies parameterised by tree |
| `probe/run-counts.sh` | builds the probe against both legs and runs it on the three fixtures |
| `counts/counts-<fixture>-<leg>.txt` | the probe's output, six files, cited in the record's counts table. The before leg reproduces change 0611's retained after-counts exactly |
| `oracle/src/main.rs` | the open differential: for every ZIP container under `test-data`, the open's verdict, its request/byte/observation cost, every package and part relationship, every admitted Part with its content type, every non-part member, and the decoded length and CRC of every Part |
| `oracle/Cargo.toml.template`, `oracle/run.sh` | the oracle's manifest template and its two-leg runner, which ends in `cmp` |
| `oracle/report-diff-classes.txt` | every line on which the two reports differ, classified |
| `oracle/corpus-open-costs.tsv` | per container, the open's requests, bytes and source observations on both legs |
| `oracle/summary.txt` | the corpus totals the record cites |
| `oracle/read-set-identity.txt` | every read of both legs logged on five packages, with the byte sets compared |
| `oracle/binaries.sha256` | the two oracle binaries |
| `geometry/run_geometry.py` | the run model: reimplements `local_span_hint` and `structural_runs` from package bytes alone, with no help from the library |
| `geometry/run-geometry.tsv` | its per-container output: structural members, runs, members covered, retained bytes, longest run |
| `geometry/run-geometry-summary.txt` | the corpus extremes, which are what the two named ceilings are set against |
| `geometry/zip-containers.txt` | the 533 containers both the oracle and the model were run over |
| `geometry/model-vs-measurement.txt` | the model's predicted request saving against the oracle's measured one, per container |
| `timing/run-timing.sh` | change 0611's paired ABBA runner: A1 B1 B2 A2 plus an A/A floor A3 A4, both kinds |
| `timing/summarise.py` | change 0611's summariser, unchanged |
| `timing/range-*.json`, `timing/local-*.json` | the twelve legs verbatim, each carrying its own binary sha256, environment and per-sample counters |
| `timing/range-summary.txt`, `timing/local-summary.txt` | `summarise.py` output, cited in the record |
| `timing/range-request-counts.txt` | the simulator's physical request and byte counts per case per leg |
| `timing/local-code-size-control.txt` | the controlled local experiment: before, after, and a third binary built from the after tree with `is_structural_member_name` forced to `false`, so the mechanism is compiled and linked but never fires. It separates the cost of `litchi-opc` growing from the cost of the prefetch running |
| `workbook-timing/src/main.rs` | a probe that opens one real package through change 0572's simulated transport — 1 ms per physical request, 100 MiB/s, 64 KiB maximum range — and reports the median with the request count that produced it |
| `workbook-timing/Cargo.toml.template`, `workbook-timing/run.sh` | its manifest template and ABBA runner |
| `workbook-timing/*.json` | the six legs, four fixtures each |
| `workbook-timing/summary.txt` | the paired deltas and the A/A floor |
| `gates.sh` | the gate runner |
| `gates.txt` | the tail of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs the coordinator merges |

## How to reproduce

```sh
# deterministic request and byte counts, both legs, three fixtures
docs/performance/results/change-0623/probe/run-counts.sh <work-dir>

# the open differential over every ZIP container under test-data, both legs
docs/performance/results/change-0623/oracle/run.sh <work-dir> \
    docs/performance/results/change-0623/geometry/zip-containers.txt

# the run model, from package bytes alone
python3 docs/performance/results/change-0623/geometry/run_geometry.py \
    docs/performance/results/change-0623/geometry/zip-containers.txt

# paired timing, both legs, plus the A/A floor
docs/performance/results/change-0623/timing/run-timing.sh <work-dir>
python3 docs/performance/results/change-0623/timing/summarise.py <work-dir> range
python3 docs/performance/results/change-0623/timing/summarise.py <work-dir> local

# the delayed-transport open on four real packages, both legs
docs/performance/results/change-0623/workbook-timing/run.sh <work-dir>

# every gate
docs/performance/results/change-0623/gates.sh
```

The local code-size control is not scripted: it is `timing/run-timing.sh`'s local
half run against a third `litchi-perf-baseline`, built from this worktree with
`is_structural_member_name` returning `false` before its first line, and timed
A N B B N A with an A/A floor in the same window.

Every runner names the before checkout (`litchi-worktrees/before-3156bff3b`) and
this worktree explicitly; point them at any two trees of the same shape to
re-run the comparison elsewhere.
