# Evidence packet — change 0611

One bounded positional read per ZIP member first-read. Record:
[`docs/performance/0611-zip-single-read-per-member.md`](../../0611-zip-single-read-per-member.md).

## Provenance

| | |
| --- | --- |
| base commit | `2d6fbeaed2083de104bbb0b52b2990ce69ac7274` (`feat/office-format-completeness`) |
| branch | `perf/0611-zip-single-read-per-member` |
| commit | see `decision.json` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0 |
| build | `cargo build --release --locked`; workspace `[profile.release] lto = true, panic = "abort"` |
| pinning | `taskset -c 31` for every measured process |
| host load | seven other agents building and measuring throughout; the A/A floor was taken in the same window |

Binary identities are recorded inside each timing JSON
(`binary_identity.binary_sha256`) and for the differential in
`differential/binaries.sha256`.

## Contents

| path | what it is |
| --- | --- |
| `design.md` | the frozen design, written before any code, with the amendment the differential forced |
| `probe/src/main.rs` | change 0587's request/byte/allocation/observation probe, verbatim except that `classify` names a read starting at a member's local header and longer than 640 bytes `member-span` |
| `probe/Cargo.toml.template` | the probe manifest with its path dependencies parameterised by tree |
| `probe/run-counts.sh` | builds the probe against both legs and runs it on the three fixtures |
| `counts/counts-<fixture>-<leg>.txt` | the probe's output, six files, cited in the record's counts tables |
| `timing/run-timing.sh` | the paired ABBA runner: A1 B1 B2 A2 plus an A/A floor A3 A4, both kinds |
| `timing/summarise.py` | turns the leg JSONs into the record's tables |
| `timing/range-*.json` | the six range-source legs verbatim, each carrying its own binary sha256, environment and the simulator's per-sample request counters |
| `timing/local-*.json` | the six local legs, trimmed to identity, configuration and every `elapsed_ns` figure the record cites; the harness's per-sample filesystem and process diagnostics are dropped |
| `timing/range-summary.txt`, `timing/local-summary.txt` | `summarise.py` output, cited in the record |
| `timing/range-request-counts.txt` | the simulator's physical request and byte counts per case per leg |
| `read_grammar_differential.rs` | change 0582's differential harness plus two per-member verdict APIs, `I.read_entry` (the changed path) and `R.read` (the slice-backed control) |
| `classify.py` | change 0582's classifier with the two added APIs registered |
| `run.sh` | the runner: two `git archive` trees, the after tree overlaid with the two changed files, corpus, both builds, both runs, classify |
| `differential/classification-summary.json` | the clean re-run after the amendment |
| `differential/classification-summary-first-run.json` | the first run, which found the two class-E divergences |
| `differential/line-divergences-first-run.txt` | those two verdict lines in full |
| `differential/binaries.sha256` | the two example binaries |
| `gates.txt` | the tail of every gate |
| `decision.json` | the decision record |
| `log-sections.md` | the four log paragraphs the coordinator merges |

## How to reproduce

```sh
# deterministic counts, both legs
docs/performance/results/change-0611/probe/run-counts.sh <work-dir>

# paired timing, both legs, plus the A/A floor
docs/performance/results/change-0611/timing/run-timing.sh <work-dir>
python3 docs/performance/results/change-0611/timing/summarise.py <work-dir> range
python3 docs/performance/results/change-0611/timing/summarise.py <work-dir> local

# the differential, both builds, full corpus
docs/performance/results/change-0611/run.sh <work-dir>   # needs ~2.2 GB on disk
```

`run-counts.sh` and `run-timing.sh` name the before checkout
(`litchi-worktrees/before-2d6fbeaed`) and this worktree explicitly; point them
at any two trees of the same shape to re-run the comparison elsewhere. The
differential's corpus regenerates byte for byte from change 0582's
`build_corpus.py` and its fixed `RNG_SEED`.

## What this packet does not contain

No cold-cache, real-device, peak-RSS, allocation-profile, cross-platform or
concurrency-scaling measurement, and no registered performance claim.
