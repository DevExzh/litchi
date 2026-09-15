# Evidence: change 0596, the eager DOC open's three size-independent terms

Change record: [`0596-doc-eager-open-terms.md`](../../0596-doc-eager-open-terms.md).

Disposition: retained. `performance_claim: none`. The counts in this packet are
deterministic; the timings are two-binary comparisons on a shared host and are
reported with their A/A floors and with one scenario that moved the wrong way
before the candidate was corrected.

## Contents

| Path | What it is |
| --- | --- |
| `counts/counts-summary.txt` | The per-open instruction totals, the inclusive shares of the touched terms, and the call counts, derived from the annotations below. This is the table the record quotes. |
| `counts/an-{before,after}-doc{small,mid,big}.txt` | `callgrind_annotate` isolation-pair analyses of the eager `Package::document()` open (`docsmall` = `saved-by-table.doc`, `docmid` = `FloatingPictures.doc`, `docbig` = the 1.6 MB kwsymphony form). Each is the difference of a 110-sample and a 10-sample profile, divided by 100: inclusive per-op deltas, self per-op deltas, and call counts per op. |
| `counts/an-{before,after}-readcount.txt` | The same analysis for `Document::paragraph_count()` alone, with the document opened once outside the loop (s=1100 against s=100, divided by 1000). This is what shows the query does +0.03% instructions. |
| `counts/an-{before,after}-readtext.txt` | The same for `Document::text()` alone. |
| `counts/an-{before,after}-ranges.txt` | The same for `FileInformationBlock::get_all_subdoc_ranges()` alone. |
| `counts/perf-stat-opens.csv` | Native `perf stat` for the eager open on all three fixtures, three runs per leg, 20 warmups plus 2,000 timed opens each: instructions, cycles, task-clock. Callgrind counts `rep movsb` per byte, so the copy removals are priced here too. |
| `counts/perf-stat-readcount.txt` | Native `perf stat` for `paragraph_count()` on `saved-by-table.doc`, three runs per leg, 50 warmups plus 2,000 calls each. |
| `timing/window{1,2,3,4}-summary.txt` | Pooled percentiles and paired deltas per scenario, both directions, with the A/A floor of the same window. **Windows 1 and 2 measure a superseded candidate build** (before `get_all_subdoc_ranges` was corrected) and are retained because they are where the `doc_semantic_paragraph_count` regression was found. **Windows 3 and 4 measure the committed candidate.** Each summary names the two binary sha256s it used. |
| `timing/window{N}-{a1,a2,b1,b2,a3,a4}.json` | The raw harness output for every leg, with every sample. Legs ran in the order a1 b1 b2 a2 a3 a4; A = a1+a2 pooled, B = b1+b2 pooled, A/A floor = a3 against a4. |
| `differential/digest-{before,after}.tsv` | One line per `.doc` fixture under `test-data/` (57 files): open outcome or the exact `Debug` text of the refusal, `Document::text()` length and hash, paragraph count and a hash over every paragraph's `Debug` form (which carries its resolved properties, its runs and each run's character properties and revision marks), `paragraph_count()` from the separate public query, the section count and a hash of the section table, a hash of `get_all_subdoc_ranges()`, and the length and hash of `FileInformationBlock::raw_data()`. |
| `differential/digest-diff.txt` | Empty: the two digests are byte-identical. |
| `probe/main.rs`, `probe/Cargo.toml` | The measurement driver. It is change 0587's retained `docppt-survey` driver plus three additions: `digest-doc` (the differential oracle above), `profile-read` (open once, then loop one query, for isolating `count`, `text` and `ranges`), and nothing else. The `Cargo.toml` names the repository crates; each leg was built from a copy whose path dependencies pointed at that leg's checkout. |
| `probe/cg.sh`, `probe/cg-read.sh` | Callgrind capture for the opens and for the isolated reads. Both pin CPU 16 and set `RAYON_NUM_THREADS=1`. |
| `probe/analyze.py` | Change 0587's isolation-pair analyzer, unchanged except that it takes its profile directory from `$CGDIR`. |
| `probe/timing.sh`, `probe/summarize_timing.py` | The paired-timing driver and its pooling analysis. |
| `gates.txt` | Tails of `cargo fmt --all --check`, `cargo clippy -p litchi-doc --all-targets`, `cargo test -p litchi-doc` and `cargo doc -p litchi-doc --no-deps`, run in the candidate worktree. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge into `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`. |
| `decision.json` | The `litchi-perf-change-decision` record. |

## Provenance

- Base commit: `08d968f8ec7db27cf1187d01911fd08b9d014d91`
  (`docs(perf): survey what remains across the OLE2 and OOXML path (0587)`).
- Candidate branch: `perf/0596-doc-eager-open-terms`.
- Control leg built from the shared read-only checkout at the base commit
  (`litchi-worktrees/before-08d968f8e`) with its own `CARGO_TARGET_DIR`;
  candidate leg built from the candidate worktree. Both `--release`, and
  `--locked` for the harness.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws;
  rustc 1.95.0 (59807616e 2026-04-14); valgrind 3.26.0; `perf` available.
  Every measured process pinned to CPU 16 with `taskset`. Seven other agents
  were building and measuring on the same host throughout.

Binary sha256:

| binary | sha256 |
| --- | --- |
| `docppt_survey`, control | `69855a7f5ebe8fb4973aeeb3dcb4760bd01c26ff2ec63e372a5a17be3c96ca54` |
| `docppt_survey`, candidate | `62caf1790f370f83e5765a9c0f3f9deef0c080ef26d2b4dbd3e2a7ab4ec0b04c` |
| `litchi-perf-baseline`, control | `dab64c26373072ab9d88883c64cce710df29aa8a36255d751dbcf3d72f2521a1` |
| `litchi-perf-baseline`, candidate (windows 1 and 2, superseded) | `aec50c61f77b4657c73b3e71b48b091984a72e96aacb42e2eefa27f98c6768bb` |
| `litchi-perf-baseline`, candidate (windows 3 and 4, committed code) | `26c709e688d293d3ee4b52e49a70118b6cfddff14623a44969fe0d0cd1946af0` |

Fixtures, all already in the tree:

| fixture | bytes |
| --- | ---: |
| `test-data/poi/test-data/document/saved-by-table.doc` | 65,024 |
| `test-data/ole/doc/FloatingPictures.doc` | 335,360 |
| `test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc` | 1,619,457 |
| `test-data/ole/doc/picture.doc` | 1,448,448 (refuses the eager open on both legs; no open to profile) |

## What this packet does not establish

No cold-page-cache, peak-RSS, allocation-profile, encrypted-document,
attached-glossary, non-Linux or non-x86-64 result. No claim-registry entry. The
harness corpora are generated by the DOC writer and are not the fixtures the
callgrind legs use; the record says which evidence comes from which.

The timing windows are two differently linked binaries, and this packet contains
a direct demonstration of how far that alone can move a scenario: a 90-instruction
edit to `get_all_subdoc_ranges` moved `doc_semantic_paragraph_count` from +15% to
-2.6% at p50 while the query's instruction count stayed within 0.03% of the
control in both builds. Percentile deltas here rank candidates; they do not
measure them.
