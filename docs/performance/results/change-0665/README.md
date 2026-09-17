# Evidence: change 0665, the eager publication audit stops asserting compactness

Change record:
[`0665-opc-eager-writer-publication-audit.md`](../../0665-opc-eager-writer-publication-audit.md).
Authority: change [0652](../../0652-owner-decisions-for-the-third-wave.md)
decision 2. The decision speaks about original part bytes; this record applies
the same `verify_source` profile to eager planned payloads, including
source-spliced replacements, as an explicit implementation interpretation
informed by 0654 and the source-backed replacement precedent in 0657. It is
not an additional owner decision. Predecessors:
[0654](../../0654-opc-original-bytes-audit-loosened.md) (the same movement on
the source-backed route's original half),
[0660](../../0660-docx-compaction-policy.md) (the gate this unblocks),
[0610](../../0610-opc-lazy-part-decode-design.md) (the `xlsx-hide` route and its
59 refusals), [0593](../../0593-opc-publication-pristine-members.md) and
[0647](../../0647-opc-get-or-add-noop-reuse-design.md) (the provenance proof
shape).

Disposition: retained. `performance_claim: none`. Everything the record cites is
here; nothing here is registered as a claim.

## Provenance

| | |
| --- | --- |
| base commit | `5af158123` (the branch after change 0660) |
| branch | `perf/0665-opc-eager-writer-source-provenance` |
| before leg | the read-only detached checkout `/home/zhuhe/code/litchi-worktrees/before-5af158123` at the base commit, built with its own `CARGO_TARGET_DIR` |
| host | AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, valgrind 3.26.0; seven other agents building and measuring throughout |
| CPU pin | `taskset -c 20` for every measured process |
| build | `cargo build --release --locked` for both harness legs; both binaries staged outside any Cargo target directory before any leg ran |

`lto = true` makes release binaries non-reproducible byte for byte (change
0635), so [`binaries.sha256`](binaries.sha256) identifies the binaries that were
actually profiled and timed.

## Contents

| path | what it is |
| --- | --- |
| [`probe/src/main.rs`](probe/src/main.rs), [`probe/Cargo.toml.template`](probe/Cargo.toml.template) | the corpus probe, built once per leg with the leg's checkout substituted. Five modes: `eager-noop`, `eager-edit`, `eager-edit-noncompact`, `xlsx-hide`, and `members` |
| [`docx-probe/src/main.rs`](docx-probe/src/main.rs), [`docx-probe/Cargo.toml.template`](docx-probe/Cargo.toml.template) | change 0660's census probe, reused unchanged except for one added aspect, `gate_source`, which reports the verdict the flipped gate reaches. Built with `--features policy` on both legs, because 0660 is in the base |
| [`census/corpus-all.txt`](census/corpus-all.txt), [`census/corpus-xlsx180.txt`](census/corpus-xlsx180.txt) | the 321-package OOXML fixture corpus and change 0610's 180 `.xlsx` population, as change 0654 enumerated them |
| [`publication/noop-{before,after}.tsv`](publication/noop-after.tsv) | the eager exact no-op over 321 fixtures, per leg: outcome, error text, output SHA-256 |
| [`publication/edit-{before,after}.tsv`](publication/edit-after.tsv) | the eager open-edit-save with a **compact** authored payload, same shape |
| [`publication/editnc-{before,after}.tsv`](publication/editnc-after.tsv) | the same with a deliberately **non-compact** authored payload |
| [`publication/hide-{before,after}.tsv`](publication/hide-after.tsv) | change 0610's `xlsx-hide` route over 180 `.xlsx` fixtures, per leg |
| [`publication/members-{before,after}.tsv`](publication/members-after.tsv) | for every fixture the `xlsx-hide` route publishes, each XML part of the artifact compared by digest with the same part of the source: identical count, changed names, added, removed |
| [`publication/route-summary.txt`](publication/route-summary.txt) | the four differentials: outcome transitions, identical-digest counts, newly published counts, kept-refusal text comparison and residual `NotCompact` |
| [`census/docx-census-{before,after}.txt`](census/docx-census-after.txt) | 63 DOCX fixtures × eleven aspects, per leg |
| [`census/docx-census-diff.txt`](census/docx-census-diff.txt) | `diff` of the two |
| [`census/docx-census-summary.txt`](census/docx-census-summary.txt) | per-aspect unchanged/moved counts, the gate transition table, the fixture that still refuses, and the `one_edit_span` byte totals |
| [`counts/counts.sh`](counts/counts.sh) | the capture script: `--warmup 0 --samples N` under callgrind at N=1 and N=3 for three selectors, per leg, pinned to CPU 20 |
| `counts/incl-<leg>-<case>-<N>.txt` | `callgrind_annotate --inclusive=yes --threshold=100`, top 220 rows, for each of the twelve runs |
| `counts/cg-<leg>-<case>-<N>.txt` | valgrind's own tail for each run, including the harness's verification output |
| [`counts/compare.py`](counts/compare.py) | change 0660's isolation-pair extractor, re-pointed at this change's symbols and selectors |
| [`counts/counts-summary.json`](counts/counts-summary.json), [`counts/counts-table.txt`](counts/counts-table.txt) | its output |
| [`timing/timing.sh`](timing/timing.sh) | the A1 B1 B2 A2 A3 A4 sequence, `--warmup 5 --samples 40` per run, CPU 20, one window |
| `timing/{A1,B1,B2,A2,A3,A4}.json` | the six harness reports with every per-sample `elapsed_ns`, the corpus manifests and the binary identity of the leg that produced them |
| [`timing/summarize.py`](timing/summarize.py), [`timing/summary.json`](timing/summary.json), [`timing/summary.txt`](timing/summary.txt) | pooled p50/mean/p95/p99/min/max per leg, both paired deltas, and the A/A and B/B floors of the same window |
| [`scripts/diff_routes.py`](scripts/diff_routes.py) | the route differential |
| [`scripts/docx_census_summary.py`](scripts/docx_census_summary.py) | the census summary |
| [`scripts/gates.sh`](scripts/gates.sh) | the gate sequence |
| [`gates.txt`](gates.txt) | the tail of every gate that was run, with its command and exit status |
| [`decision.json`](decision.json) | the decision, its reason codes, accepted evidence, accepted costs, known gaps and provenance |
| [`log-sections.md`](log-sections.md) | the four log paragraphs for the coordinator to merge |
| [`cleanup.json`](cleanup.json) | what was removed and what was kept |

## Replaying it

```sh
BASE=5af158123
# before leg: a detached checkout at $BASE; after leg: this branch.
# 1. harness binaries, both legs, --release --locked, staged outside target/
# 2. probes: one Cargo project per leg per probe, with <CHECKOUT> substituted;
#    the DOCX probe takes --features policy on both legs.
# 3. probe eager-noop / eager-edit / eager-edit-noncompact corpus-all.txt
#    and xlsx-hide / members corpus-xlsx180.txt, both legs
#    -> scripts/diff_routes.py
# 4. docx-probe census <repo>/test-data, both legs
#    -> scripts/docx_census_summary.py
# 5. counts/counts.sh <leg-binary> <before|after> <outdir> <rawdir>   (counts first)
#    -> counts/compare.py <outdir> <rawdir>
# 6. timing/timing.sh <outdir> <stage-dir>                            (timing last)
#    -> timing/summarize.py <outdir>
# 7. scripts/gates.sh <worktree> <outdir>
```

## What the numbers are

Every count is deterministic and reproduces exactly: the route outcomes, the
digests, the member comparisons, the census aspects and the callgrind isolation
pairs. The timing figures are wall-clock medians with the A/A and B/B floors of
the same window beside them; they establish "no practical difference", not "no
difference", and the one scenario that exceeds its floor is named in the record.
Instruction counts rank work, not latency.
