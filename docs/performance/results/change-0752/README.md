# Evidence packet for change 0752

Record: [0752](../../0752-streaming-writer-small-write-batching.md) — a
one-pass CRC-32 stage for owned ZIP entries, handle-free and borrowed budget
reservations, and run-wise DOCX escaping. Base `6d989cad63`, candidate
`1167398a3e`. Every measured process was pinned to CPU 28 on an AMD EPYC 9R45.

| path | what it holds |
| --- | --- |
| `binaries.txt` | build commands (identical for both legs) and SHA-256 of every measured binary and probe |
| `gates.txt` | every gate command, exit code and test count |
| `timing/abba-raw.tar.gz` | the 144 raw harness reports of the ABBA campaign (4 rounds × 9 cases × before, after, after, before), with `progress.log` and `stderr.log` |
| `timing/summary.json`, `timing/flags.json` | per-case medians, paired changes, bootstrap intervals, output digests, per-process rows; every paired p50/p95/mean comparison beyond 5% |
| `counters/` | per-timed-iteration user-space instructions and cycles (low/high-sample differencing, two ABBA rounds), raw and summary |
| `alloc/` | allocation-build reports (one process per leg and case) and the per-case medians |
| `layout-control/` | the A/A campaign (the base built at a second path against the before binary), the CFB control's front-end counters (directory `frontend/` inside the tarball) and its top-symbol profiles for the before, layout and after builds |
| `differential/` | the base-versus-candidate probe: source, manifest template (`ROOT` is each tree), shared lock file (`Cargo.lock.pinned`), run summaries, transcript digests, outcome categories and the transcript's first line |
| `probe/` | the CRC-32 and Deflate chunking probe: source, manifest, lock (`Cargo.lock.pinned`: the harness lock's crc32fast 1.5.0, flate2 1.1.9, zlib-rs 0.6.7), output, and the comparison of its model stream with each leg's real `word/document.xml` member |
| `profiles/` | flat `perf report` summaries (whole process) of the large DOCX and XLSX cases, before and after |
| `experiments/scoped-reservation-upper-bound/` | the throwaway experiment that priced the reservation's reference counting (diff, raw reports, summary); not committed code |
| `scripts/` | every runner and summarizer used |
| `cleanup.json` | what was removed and what was kept |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Replaying

Build each leg with the commands in `binaries.txt`, then:

```sh
scripts/run_abba.sh OUT 4              # timing campaign
python3 scripts/summarize_abba.py OUT
scripts/run_counters.sh OUT            # instruction and cycle differencing
python3 scripts/summarize_counters.py OUT
scripts/run_alloc.sh OUT               # allocation build
python3 scripts/summarize_alloc.py OUT
scripts/run_aa_layout.sh OUT 4         # A/A layout control
```

The scripts carry this host's absolute paths; edit the binary variables at
their tops.

The differential probe is built twice from `differential/diffprobe_main.rs`
(as `src/main.rs`), each with `Cargo.toml.template` whose `ROOT` names one
tree and with `Cargo.lock.pinned` copied in as `Cargo.lock` (lock files are
gitignored, hence the name); `cargo build --release --offline`, then
run both binaries and compare stdout byte for byte. Its stderr gives the
line count and the SHA-256 of the transcript without its final newline
(`95bc9a12…`); `transcripts.sha256` hashes the files as written (`d15ea98b…`).
Setting `DOCX_LARGE_OUT=path` instead writes the harness's large streaming
DOCX corpus to `path`.

The chunking probe (`probe/crc_deflate_probe/`, with `Cargo.lock.pinned`
copied in as `Cargo.lock`) takes an optional paragraph
count (default 131,072) and `MODEL_STREAM_OUT=path` to save its model of the
writer's `word/document.xml` Deflate stream.
