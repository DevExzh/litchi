# Evidence packet — change 0632

One bounded read of the head of the central directory serves both the ZIP
locator's first-central-record probe and the index's central-directory scan, and
the buffer it lands in is the scan's buffer. Record:
[`docs/performance/0632-zip-directory-prefill-locate.md`](../../0632-zip-directory-prefill-locate.md).

## Provenance

| | |
| --- | --- |
| base commit | `c7326f680` (`feat/office-format-completeness`, change 0630) |
| branch | `perf/0632-zip-tail-window-locate` |
| commit | see `decision.json` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680`, a read-only detached checkout of the base |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0 |
| build | `cargo build --release --locked`; workspace `[profile.release] lto = true, panic = "abort"`; both legs resolved from the same `Cargo.lock` |
| pinning | `taskset -c 8` for every measured process; every build pinned off CPU 8 |
| host load | seven other agents building and measuring throughout, load average 26–42; every timing window carries its own A/A floor and a B/B floor |

Binary identities: `timing/binaries.sha256`, `workbook-timing/binaries.sha256`,
`differential/binaries.sha256`, `oracle/binaries.sha256`, and inside each
range/local leg JSON (`binary_identity.binary_sha256`).

## Contents

| path | what it is |
| --- | --- |
| `design.md` | the frozen design, written before any production line changed, with its falsification criterion |
| `census/cd_census.py` | a census of the ZIP tail geometry — central-directory size and offset, comment length, ZIP64 sentinels, whether the EOCD sits at `len - 22`, whether the first central record sits at the declared offset — computed from the bytes alone, with no help from the library |
| `census/cd-census.tsv` | its output for all 533 containers, which is what the 64 KiB window is set against |
| `census/zip-containers.txt` | the 533 containers, the same list changes 0582 and 0623 use |
| `probe/src/main.rs` | change 0587's request/byte/allocation/observation probe as change 0623 left it, verbatim: change 0632 removes the `cd-probe(46)` request that probe already classifies, so no new class is needed |
| `probe/Cargo.toml.template`, `probe/run-counts.sh` | the probe manifest with its path dependencies parameterised by tree, and its two-leg runner |
| `counts/counts-<fixture>-<leg>.txt` | the probe's output, six files, cited in the record's counts table. The before leg reproduces change 0623's retained after-counts exactly |
| `oracle/src/main.rs` | change 0623's open differential, verbatim: for every ZIP container under `test-data`, the open's verdict, its request/byte/observation cost, every package and part relationship, every admitted Part with its content type, every non-part member, and the decoded length and CRC of every Part |
| `oracle/Cargo.toml.template`, `oracle/run.sh` | its manifest template and two-leg runner, which ends in `cmp` |
| `oracle/report-diff-classes.txt` | every line class on which the two reports differ: one class, `open-cost`, 533 lines each side |
| `oracle/corpus-open-costs.tsv` | per container, the open's requests, bytes and source observations on both legs |
| `oracle/summary.txt` | the corpus totals the record cites |
| `oracle/reports.sha256`, `oracle/binaries.sha256` | the two reports and the two oracle binaries |
| `read_grammar_differential.rs` | change 0611's extended copy of change 0582's differential harness: change 0582's seven strict-layout APIs plus `I.read_entry` (the ordinary indexed read path) and `R.read` (the slice-backed control), both limit profiles, both read directions |
| `classify.py` | change 0582's classifier with the two added APIs registered |
| `run-differential.sh` | the runner: two `git archive` trees, the after tree overlaid with change 0632's three production files, change 0582's corpus regenerated from its fixed seed, both builds, both runs, classify |
| `differential/classification-summary.json` | its output: 22,875 inputs, 2,886,786 member verdicts, zero divergences of every class, zero oracle failures, zero panics |
| `differential/divergences.tsv` | empty, because there are none |
| `differential/report-identity.txt`, `differential/reports.sha256` | the two 3,801,502-line reports are **byte-identical**, with their shared SHA-256 |
| `differential/binaries.sha256` | the two harness binaries |
| `timing/run-timing.sh` | the paired ABBA runner: A1 B1 B2 A2 plus an A/A floor A3 A4, both kinds, with both binaries copied out of their Cargo target directories first (change 0627) |
| `timing/summarise.py` | change 0611's summariser plus a B/B floor line |
| `timing/range-*.json` | the six range-source legs verbatim, each carrying its own binary sha256, environment and the simulator's per-sample counters |
| `timing/local-*.json` | the six local legs, trimmed to identity, configuration and every `elapsed_ns` figure the record cites; the harness's per-sample filesystem and process diagnostics are dropped, as change 0611 dropped them |
| `timing/range-summary.txt`, `timing/local-summary.txt` | `summarise.py` output, cited in the record |
| `timing/range-request-counts.txt` | the simulator's physical request and byte counts per case per leg, deterministic across all 30 samples of every leg |
| `timing/binaries.sha256` | the two staged `litchi-perf-baseline` binaries |
| `workbook-timing/src/main.rs` | change 0623's probe, verbatim: opens one real package through change 0572's simulated transport — 1 ms per physical request, 100 MiB/s, 64 KiB maximum range — and reports the median with the request count that produced it |
| `workbook-timing/Cargo.toml.template`, `workbook-timing/run.sh` | its manifest template and ABBA runner, at 60 samples and 10 warmups |
| `workbook-timing/*.json` | the six legs, four fixtures each |
| `workbook-timing/summarise.py`, `workbook-timing/summary.txt` | the paired deltas, the A/A floor and the B/B floor |
| `workbook-timing/binaries.sha256` | the two probe binaries |
| `gates.sh` | the gate runner |
| `gates.txt` | the tail of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs the coordinator merges |

## How to reproduce

```sh
# the corpus census the 64 KiB window is set against
python3 docs/performance/results/change-0632/census/cd_census.py \
    docs/performance/results/change-0632/census/zip-containers.txt .

# deterministic request, byte and allocation counts, both legs, three fixtures
docs/performance/results/change-0632/probe/run-counts.sh <work-dir>

# the open differential over every ZIP container under test-data, both legs
docs/performance/results/change-0632/oracle/run.sh <work-dir> \
    docs/performance/results/change-0632/census/zip-containers.txt

# the read-grammar differential, both builds, full corpus (needs ~2 GB on disk)
docs/performance/results/change-0632/run-differential.sh <work-dir>

# paired timing, both legs, plus the A/A and B/B floors
docs/performance/results/change-0632/timing/run-timing.sh <work-dir>
python3 docs/performance/results/change-0632/timing/summarise.py <work-dir> range
python3 docs/performance/results/change-0632/timing/summarise.py <work-dir> local

# the delayed-transport open on four real packages, both legs
docs/performance/results/change-0632/workbook-timing/run.sh <work-dir>
python3 docs/performance/results/change-0632/workbook-timing/summarise.py <work-dir>

# every gate
docs/performance/results/change-0632/gates.sh
```

Every runner names the before checkout (`litchi-worktrees/before-c7326f680`) and
this change's worktree explicitly; point them at any two trees of the same shape
to re-run the comparison elsewhere.

## What this packet does not contain

No cold-cache, real-device, peak-RSS, instruction-count, syscall,
concurrency-scaling or cross-platform measurement, and no registered performance
claim.
