# Evidence packet — change 0638

Change 0587's evidence gap 5 had two halves — no selector opened a `.doc` or a
`.ppt` through the `litchi` facade, and no selector measured the ordinary
documented OOXML save path. Thirty opt-in selectors close both, and the first
baseline says the documented save costs 8 to 62 times more to publish than to
serialize. Record:
[`docs/performance/0638-facade-and-ordinary-save-selectors.md`](../../0638-facade-and-ordinary-save-selectors.md).
`performance_claim: none`. **No file under `crates/` was modified.**

## Provenance

| field | value |
| --- | --- |
| base commit | `c7326f680` (`feat/office-format-completeness`) |
| branch | `perf/0638-harness-facade-and-ordinary-save-selectors` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0638` (removed after commit) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21) |
| build | `cargo build --release --locked` from `tools/perf-baseline` |
| CPU affinity | every measured process pinned to CPU 14 with `taskset` |
| destination device | the host's ext4 root filesystem, via `--filesystem-root`; never `/tmp`'s tmpfs |
| host state | eight agents building and testing concurrently throughout |

There is no before leg: nothing under `crates/` changed, and every selector is
new, so the whole capture is one descriptive baseline. The binary was staged
outside the Cargo target directory before the run, for the reason change 0627
gives (a concurrent build relinked one mid-run):

| binary | SHA-256 |
| --- | --- |
| `litchi-perf-baseline` | `6abee148b50bb6beff8f1bbfe5ea61a28ee095fbbf0664fe7c5a14d1111df4bb` |

## Contents

| path | what it is |
| --- | --- |
| `scripts/baseline.sh` | the capture: 20 warm-ups and 50 retained samples per case, pinned, four invocations per repeat (two facade fixture pairs, one refusing DOC fixture, and the twenty-four ordinary-save selectors over change 0593's three OOXML fixtures) |
| `scripts/summarize.py` | the summariser. It re-asserts every per-case gate over the retained reports and exits non-zero on any failure, so its output is evidence rather than narration: no retained sample may differ from its frozen facade observation, no publication from the corpus reference, no edit outcome from the frozen one, and every corpus must have passed the 0625/0631 repeated-cycle and repeated-save proofs |
| `raw/R{1,2,3}-facade-small.json` | the six facade selectors on `documentProperties.doc` (9,728 B) and `ppt_with_png.ppt` (39,424 B) |
| `raw/R{1,2,3}-facade-large.json` | the same six on `FloatingPictures.doc` (335,360 B) and `SampleShow.ppt` (125,440 B) |
| `raw/R{1,2,3}-facade-refusal.json` | the three DOC selectors on `duplicate-style-names.doc`, one of the four fixtures change 0609's census found the facade refuses and the source-backed snapshot admits |
| `raw/R{1,2,3}-ordinary-save.json` | all twenty-four ordinary-save selectors: three generated corpora and change 0593's `alt-chunk-header.docx`, `ConditionalFormattingSamples.xlsx` and `slide-section-test.pptx` |
| `baseline-summary.txt` | `summarize.py`'s output over all three repeats: the per-row table, the byte split, the phase attribution, the A/A floor and the eight rows whose floor exceeds 5% |
| `gates.txt` | the tail and exit status of every gate, the three pre-existing Python failures reproduced on the untouched base checkout, and the baseline's own asserted gates |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge |

The corpora are caller-named files and existing harness corpora, so no fixture
is added to the repository and no probe crate is retained: the selectors *are*
the probe, and they are checked in under `tools/perf-baseline/src/`.

## Headline numbers

| measurement | value | tier |
| --- | --- | --- |
| selectable registry | 471 → **501** names; `Case::DEFAULT` unchanged at 41 | measured |
| checked default catalog SHA-256 | unmoved (`test_corpus_manifest_v2`, `test_crud_coverage_index` pass) | measured |
| A/A floor, all 39 rows over three repeats | **p50 2.09%**, max 60.36%; 8 rows above 5%, all named | measured |
| facade DOC open, 9,728 B → 335,360 B | 15.09 µs → 178.63 µs | measured |
| facade PPT open, 39,424 B → 125,440 B | 25.55 µs → 25.95 µs (+1.6% for 3.2× the bytes) | measured |
| facade DOC one paragraph vs. full text, 335 KB fixture | 538.31 µs vs. 275.32 µs — **1.96×** | measured |
| facade DOC refusal (`duplicate-style-names.doc`) | 14.37 µs, a complete open's cost | measured |
| atomic publication ÷ counting publication | **7.7× to 61.7×** over six corpora | measured, medians device-bound |
| DOCX generated lifecycle = edit + atomic publish | 5.74 ms = 0.38 + 5.16 ms (96.4%) | measured |
| XLSX generated lifecycle | 17.37 ms = 4.21 + 11.58 ms (90.9%) | measured |
| PPTX real-file one shape-text edit | **133.61 ms**, 95.7% of its lifecycle, 490× its serialization | measured |
| one-element edit, payload bytes regenerated | 0.45% (XLSX real), 1.9% (PPTX real), 13.1% (DOCX generated) | measured |
| ZIP framing on the 131-member XLSX real file | 19,658 B, 3.0% of the output | measured |
| DOCX `document_mut()` on `alt-chunk-header.docx` | typed refusal, stable over 600 retained samples | measured |
| retained samples with no oracle, publication or edit-outcome deviation | 39 × 50 × 3 = **5,850** | measured |
| 0625/0631 determinism proofs | 6 corpora × 2, all passing | measured |

Eight rows' A/A floor exceeds 5%; five of them carry the atomic publication,
whose two `fsync` calls and `rename` are served by a device shared with seven
other agents. Those medians are an order of magnitude, not a number. The
*ratios* above are robust: all three repeats agree on their ordering and
magnitude.

## Reproducing

```sh
# build and stage the binary outside the Cargo target directory
cd tools/perf-baseline && cargo build --release --locked && cd ../..
install -D tools/perf-baseline/target/release/litchi-perf-baseline /some/disk/bin/

# one repeat (about six minutes)
bash docs/performance/results/change-0638/scripts/baseline.sh \
  "$PWD" /some/disk/bin/litchi-perf-baseline /some/disk/raw 14 R1

# the summary, once at least two repeats exist
python3 docs/performance/results/change-0638/scripts/summarize.py /some/disk/raw
```

`baseline.sh` passes `--filesystem-root "$OUT/scratch"`, so point `$OUT` at a
disk-backed directory: the save family publishes real files, and this host's
`/tmp` is a RAM-backed tmpfs whose timings would not be the filesystem's.

A single selector, outside the script:

```sh
litchi-perf-baseline --warmup 20 --samples 50 \
  --case doc_facade_file_open,ppt_facade_file_open \
  --ole2-file test-data/ole/doc/documentProperties.doc \
  --ole2-file test-data/ole/ppt/SampleShow.ppt --json facade.json

litchi-perf-baseline --warmup 20 --samples 50 --filesystem-root /some/disk/path \
  --case xlsx_ordinary_save_lifecycle,xlsx_ordinary_save_atomic_publish \
  --json save.json
```

## What this packet does not establish

No speedup, regression, allocation, RSS, instruction, cycle, cold-cache,
physical-device, network or cross-platform result, and no `fsync` attribution
inside the atomic interval. Six corpora, five real fixtures, one edit shape per
format, one host, one toolchain. The byte split is derived from the two
archives, not from production counters, so
`payload_bytes_identical_to_source` is an upper bound on what a copy-through
publisher could have avoided, not an observation that this writer copied
anything. The record's Limitations section is the complete list.
