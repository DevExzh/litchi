# Evidence packet for change 0762

Record: [0762](../../0762-streaming-batch-compression.md) — owned (authored)
ZIP Deflate entries hand the codec fixed 16 KiB chunks cut at absolute member
offsets, end without a pre-finish sync flush, and compress at zlib level 5.
Base `1d1044e3ac`, candidate `bc18e8abdd`. Every measured process was pinned
to CPU 12 on an AMD EPYC 9R45.

| path | what it holds |
| --- | --- |
| `binaries.txt` | build commands (identical for both legs) and SHA-256 of every measured binary; the probes |
| `gates.txt` | every gate command, exit code and per-crate test count |
| `timing/abba-raw.tar.gz` | the 128 raw harness reports of the ABBA campaign (4 rounds × 8 cases × before, after, after, before), with `progress.log` and `stderr.log` |
| `timing/summary.json`, `timing/flags.json` | per-case medians, paired changes, bootstrap intervals, per-leg output digests, per-process rows; every paired p50/p95/mean comparison beyond 5% |
| `counters/` | per-timed-iteration user-space instructions and cycles (low/high-sample differencing, two ABBA rounds), raw and summary |
| `alloc/` | allocation-build reports (one process per leg and case) and the per-case medians |
| `probe/level_probe/` | the compression study: source and manifest template (`ROOT` is the branch tree) |
| `probe/level-{docx,xlsx,pptx}.tsv` | levels 1–6, chunk sizes 4 KiB–one call, with and without the sync flush, on the harness corpora (sizes exact, times the minimum and median of repeated passes) |
| `probe/level-real.tsv` | levels 1–6 and the base protocol over every member of the workspace's real DOCX, XLSX and PPTX fixtures |
| `probe/member_probe/` | the member differential, split-determinism and preservation probe: source and manifest template (`ROOT` is each tree) |
| `probe/member-digest-summary.txt` | per corpus: archive digests and sizes on both legs, the digest of all member lines (names, lengths, uncompressed SHA-256s) on each leg, and the tiny corpora's per-member compressed sizes |
| `probe/split-{before,after}.txt` | archive digests of the large corpora written in five processes, the DOCX runs' text split five ways |
| `probe/preserve-{before,after}.txt` | preservation-writer output digests (or refusals) for the 62 DOCX fixtures after a one-paragraph semantic edit |
| `probe/chunk_dependence/` | a 40-line demonstration that zlib-rs's stream depends on its input splitting (levels 5 and 6) |
| `profiles/` | flat `perf report` summaries (whole process, untimed corpus work included) of the three large streaming cases, before and after |
| `scripts/` | every runner and summarizer used |
| `cleanup.json` | what was removed and what was kept |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Replaying

Build each leg with the commands in `binaries.txt`, copy the two binaries to
`bin/b/` and `bin/a/` of equal-length paths, then:

```sh
scripts/run_abba.sh OUT 4              # timing campaign
python3 scripts/summarize_abba.py OUT
scripts/run_counters.sh OUT            # instruction and cycle differencing
python3 scripts/summarize_counters.py OUT
scripts/run_alloc.sh OUT               # allocation build
python3 scripts/summarize_alloc.py OUT
```

The scripts carry this host's absolute paths; edit the variables at their
tops. The probes build with `cargo build --release --offline` after replacing
`ROOT` in their manifest template with a tree and copying the workspace
`Cargo.lock` in (lock files are gitignored). `level_probe` takes a corpus
prefix (`docx`, `xlsx`, `pptx`) or `real DIR`; `member_probe` takes `digest`,
`split SEED` or `preserve DOCX...`.
