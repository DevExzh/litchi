# Retained evidence — change 0654

Record: [`../../0654-opc-original-bytes-audit-loosened.md`](../../0654-opc-original-bytes-audit-loosened.md).
Authority: change [0652](../../0652-owner-decisions-for-the-third-wave.md)
decision 2. Design: [0602](../../0602-xlsx-real-producer-admission-design.md)'s
D0. Price: [0613](../../0613-opc-original-audit-memo.md).

## Provenance

| | |
| --- | --- |
| base commit | `70d7768cc` |
| branch | `perf/0654-opc-original-bytes-audit-loosened` |
| before leg | the shared read-only checkout `/home/zhuhe/code/litchi-worktrees/before-70d7768cc` at the same commit, built with its own `CARGO_TARGET_DIR` |
| host | AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, valgrind 3.26.0 |
| CPU pin | `taskset -c 9` for every measured process |
| build | `cargo build --release --locked` for both harness legs; binaries staged outside any Cargo target directory before any leg ran |
| `litchi-perf-baseline` before | `44cc4e2875bb20f6fd5c054a967a052ee95b6b6a0ca5063492ffecd2a603da25` |
| `litchi-perf-baseline` after | `48ecaf8088f60cb89ee1f8660d4204fe31f2ab372969658f279012752eca0aad` |

`lto = true` makes release binaries non-reproducible byte for byte (change
0635), so the two SHA-256s above identify the binaries that were actually
timed. They are also in [`binaries.sha256`](binaries.sha256).

The corpus probe is a separate Cargo project with path dependencies on each
leg's crates; its source is [`probe/src/main.rs`](probe/src/main.rs) and its
manifest template [`probe/Cargo.toml.example`](probe/Cargo.toml.example). The
before leg builds it with default features; the after leg adds
`--features source-policy`, which is the only difference between the two probe
binaries and is what selects `verify_source` for the original-bytes column.

## Contents

| path | what it is |
| --- | --- |
| [`census/corpus-95.txt`](census/corpus-95.txt) | change 0602's 95-package corpus: every `.xlsx` under `test-data/ooxml/xlsx` and `test-data/office-interop` |
| [`census/corpus-all.txt`](census/corpus-all.txt) | the 321-package OOXML fixture corpus under `test-data` |
| [`census/corpus-xlsx180.txt`](census/corpus-xlsx180.txt) | the 180 `.xlsx` fixtures, change 0610's `xlsx-hide` population |
| [`census/census-95.jsonl.gz`](census/census-95.jsonl.gz) | one row per XML member of the 95-package corpus (1,403 rows): file, member, length, and the authored and original-bytes verdicts on both legs |
| [`census/census-all.jsonl.gz`](census/census-all.jsonl.gz) | the same over the 321-package corpus (6,981 rows) |
| [`census/summary-95.txt`](census/summary-95.txt) | the 95-package verdict tables, transitions and per-package roll-up |
| [`census/summary-all.txt`](census/summary-all.txt) | the same for 321 packages, including the 17 members that stay refused, named |
| [`publication/publish-before.tsv`](publication/publish-before.tsv) | 321 publications through `write_part_overlay_to_stream` on the before leg: fixture, replaced Part, outcome, error text, output SHA-256 |
| [`publication/publish-after.tsv`](publication/publish-after.tsv) | the same on the after leg |
| [`publication/diff-summary.txt`](publication/diff-summary.txt) | the outcome transitions, the kept-refusal text comparison, the identical-digest check, and the member-by-member byte comparison of every untouched member of every published package |
| [`publication/hide-before.tsv`](publication/hide-before.tsv) | change 0610's `xlsx-hide` route over 180 `.xlsx` fixtures, before leg |
| [`publication/hide-after.tsv`](publication/hide-after.tsv) | the same after: identical outcome, identical error text, identical digests |
| [`witness/witness-before.txt`](witness/witness-before.txt) | the base witness: a synthetic package whose original bytes are compact, clean and with trailing bytes appended |
| [`witness/witness-after.txt`](witness/witness-after.txt) | the same on the after leg — identical digest, identical refusal |
| [`witness/eager-writer-blob-provenance.txt`](witness/eager-writer-blob-provenance.txt) | the refused Part of `dataValidity.xlsx` on the eager route: its blob is byte-identical to the source ZIP member |
| [`counts/callgrind-isolation-pairs.txt`](counts/callgrind-isolation-pairs.txt) | the four profiles' audit call counts per call site, the `summary:` totals, and the inclusive Ir of publication and of each audit helper |
| [`timing/A1.json`](timing/A1.json) … [`timing/A4.json`](timing/A4.json), [`timing/B1.json`](timing/B1.json), [`timing/B2.json`](timing/B2.json) | the six raw harness reports, 30 samples per leg on three selectors |
| [`timing/summary.json`](timing/summary.json) | per-leg p50, mean, p95, p99, min and max for all five scenarios, both paired deltas, the A/A floor, and the output SHA-256 set |
| [`timing/driver.txt`](timing/driver.txt) | the leg order as it ran |
| [`probe/`](probe/src/main.rs) | the corpus probe: `audit`, `publish`, `xlsx-hide`, `witness` and `blob-provenance` |
| [`scripts/census.py`](scripts/census.py) | pipes every XML member of a corpus through the probe on both legs and joins the verdicts |
| [`scripts/summarize.py`](scripts/summarize.py) | the census summary tables |
| [`scripts/diff_publish.py`](scripts/diff_publish.py) | the publication differential and the untouched-member byte comparison |
| [`scripts/callgrind.sh`](scripts/callgrind.sh) | the four isolation-pair profiles |
| [`scripts/timing.sh`](scripts/timing.sh) | the A1 B1 B2 A2 A3 A4 sequence |
| [`scripts/audit-call-counts.py`](scripts/audit-call-counts.py) | change 0613's extractor, reused unchanged |
| [`scripts/gates.sh`](scripts/gates.sh) | the gate sequence |
| [`gates.txt`](gates.txt) | the tail of every gate that was run, with its command and exit status |
| [`decision.json`](decision.json) | the decision, its reason codes, accepted evidence, accepted costs, known gaps and provenance |
| [`log-sections.md`](log-sections.md) | the four log paragraphs for the coordinator to merge |
| [`cleanup.json`](cleanup.json) | what was removed and what was kept |

## Replay

```sh
BASE=70d7768cc
# before leg: the shared read-only checkout at $BASE; after leg: this branch.
# 1. harness binaries, both legs, --release --locked, staged outside target/
# 2. probe: one Cargo project per leg (probe/Cargo.toml.example with the leg's
#    checkout substituted for <CHECKOUT>); the after leg adds
#    --features source-policy
# 3. scripts/census.py corpus-95.txt census-95.jsonl   (and corpus-all.txt)
# 4. probe publish <outdir> corpus-all.txt, both legs; scripts/diff_publish.py
# 5. probe xlsx-hide corpus-xlsx180.txt, both legs
# 6. probe witness, both legs
# 7. scripts/callgrind.sh   (deterministic counts first)
# 8. scripts/timing.sh      (timing last, CPU 9, one window)
# 9. scripts/gates.sh
```

## What the numbers are

Every count in the record is deterministic and reproduces exactly. The
instruction figures are callgrind isolation pairs — profile `--samples 1` and
`--samples 3`, difference the inclusive totals and halve, because one harness
iteration of `xlsx_source_backed_cell_values_one_edit_save` publishes two corpus
shapes. The timing figures are wall-clock medians with the A/A floor of the same
window beside them; they establish "no practical difference", not "no
difference". Instruction counts rank work, not latency.
