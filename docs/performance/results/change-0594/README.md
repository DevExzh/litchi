# Evidence: change 0594, one Deflate decoder per OOXML open

Change record: [`0594-zip-session-reuse-per-open.md`](../../0594-zip-session-reuse-per-open.md).

Disposition: retained, after a first implementation was held for a regression on
`docx_file_source_open`. `performance_claim: none`. The counts are exact and
deterministic; the instruction figures are callgrind; the cycle figures are
`perf stat` isolation pairs; the timings are paired, reported beside the A/A
floor measured in the same window, and every scenario that ran is reported,
including the one that is still adverse.

## Contents

| Path | What it is |
| --- | --- |
| `probe/` | Change 0587's `zip-opc-read` probe — an in-memory `ReadAt` logging every `read_at`/`version()` call, a counting global allocator, a central-directory classifier — extended here with a `read_parts_ordered` phase and a `PROBE_PHASES=time` open-timing phase. Each leg was built from a copy whose two path dependencies point at that leg's checkout, with its own `CARGO_TARGET_DIR`. |
| `counts/` | Full probe output for three fixtures on both legs: open, first cold part read, second (cached) read, `DeflateDecoder::new` measured alone, all parts in both physical orders, `stream_to`, and the serial `read_parts_ordered` wave, each with allocations, allocated bytes, `version()` calls and the positional requests classified against the fixture's own central directory. |
| `callgrind/` | `callgrind_annotate` self-cost tables for open plus one part on the 132-member workbook, one run per leg. |
| `attribution/` | Everything that explains the held regression. `perf-isolation-pairs.txt` and the `perf-*.csv` behind it: `perf stat -r 5` over 100 and 200 opens of the same package, differenced, for the DOCX corpus (below the threshold) and the PPTX corpus (above it), on the retained implementation (`perf-docx-*`, `perf-pptx-*`) and on the held one (`perf-held-docx-*`). `callgrind-docx-open-*`: a 20/40 callgrind isolation pair of the DOCX open on both legs of the held implementation. `interleaved-*`: the selector's own `--filesystem-child` run 150 (DOCX) and 120 (PPTX) times per leg, **alternating leg by sample** so drift hits both equally, with its `elapsed_ns` and `/proc` deltas; one set for the held implementation and one for the retained. `openprobe/`: the isolation-pair probe source. The drivers and summarizers are beside them. |
| `timing/` | `abba-r4-*` (docx, pptx and xlsx selectors, 40 samples per leg) and `abba-r5-*` (`xlsx_source_open` alone, 200 samples per leg) with `summary-r4.txt`, `summary-r5.txt` and their drivers; `probe-open-abba.txt` and `probe-time-*` for the in-process open timing on two real fixtures, 40 timed opens per leg in the same A1 B1 B2 A2 order. |
| `held-first-implementation/` | `threshold.patch`, the patch that turns the held commit `c8b23c5f4` into the retained implementation — reverse it to reproduce the held one and re-measure the +8 minor page faults. `timing/` holds the held implementation's harness evidence: `summary-r{1,2,3}.txt`, r3's four raw reports, and `faults.txt` (minor faults for one fresh-process open, 15 runs per leg). |
| `gates.txt` | The tail of every gate run on the retained implementation in the worktree. |
| `binary-sha256.txt` | SHA-256 of the six measured binaries. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `cleanup.json` | What was removed from the session scratchpad and from this packet, and why. |

## Provenance

* Base commit `08d968f8ec7db27cf1187d01911fd08b9d014d91` (branch
  `feat/office-format-completeness`); candidate branch
  `perf/0594-zip-session-reuse-per-open`; held first implementation
  `c8b23c5f429ea240ed3c58dda1fa1746d7412e5b` on the same branch.
* Before leg: the shared read-only checkout `litchi-worktrees/before-08d968f8e`
  at the base commit, with its own `CARGO_TARGET_DIR`. After leg: the candidate
  worktree. Both legs `cargo build --release --locked`, same rustc, same flags.
* Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0;
  valgrind 3.26.0; `perf` with `perf_event_paranoid=1`. Every measured process
  pinned with `taskset -c 14`.
* Repository fixtures: `test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx`
  (654,688 B, 132 members, 41 relationship members),
  `test-data/ooxml/pptx/shapes.pptx` (68,822 B, 48 members, 20 relationship
  members), `test-data/ooxml/docx/comment.docx` (4,926 B, 10 members, 2
  relationship members).
* Harness corpora, generated deterministically by the selectors and identified
  by SHA-256 in every timing report and in `decision.json`:
  `docx_file_source_open` (`a4a2e492…`, 16,793,036 B, 20 members, 3 relationship
  members) and `pptx_file_source_open` (`61b2b990…`, 17,017,139 B, 445 members,
  223 relationship members). They were extracted from a harness run for the
  attribution and deleted afterwards; `cleanup.json` records that.

## What this packet does not establish

The counts are exact for these fixtures on an in-memory positional source with a
warm page cache; they are not a latency, RSS, peak-memory, cold-cache,
physical-device, range-source or concurrency result. The callgrind figures are
one run per leg and charge `rep stosb`/`rep movsb` per byte — this packet
measures how far that can diverge from native cost, −7.38% Ir against −0.81%
native instructions and flat cycles on the same operation. `perf stat` counters
carry their own 0.19–0.72% run error, stated where a delta is inside it. Every
timing row is reported beside the A/A floor measured in the same window; rows
that do not clear it are reported as inside the floor, and the one row that
remains adverse is reported as adverse. The threshold value is a scoped
measurement on one host with one allocator. Nothing here is registered in the
claim registry.
