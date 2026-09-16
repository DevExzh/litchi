# Evidence: change 0643, the borrowing DOCX sink parser and change 0592's attributed regression

Change record:
[`0643-docx-paragraph-count-and-sink-text.md`](../../0643-docx-paragraph-count-and-sink-text.md).

Disposition: retained. `performance_claim: none`. One file under
`crates/litchi-docx/src/` and one test file changed; the counts and paired
medians here are reported as evidence and are not registered as claims. Part (a)
of the change — the attribution of change 0592's open
`docx_file_eager_paragraph_count` regression — changed no source at all.

## Contents

| Path | What it is |
| --- | --- |
| `probe/src/main.rs` | The isolation probe, derived from change 0592's. Its fixture is either a synthetic *N*-paragraph package (0592's, verbatim) or a real `.docx` opened through the same `detect_format_smart_with_limits` + `Package::from_opc_package` pair `PreparedDocx::eager` uses. Three `*_paragraph_count` operations differ **only** in the untimed preparation, which is what part (a) turns on. `reps` may be 0, so a 0 → 1 isolation pair prices the *first* call in a fresh process — the term 0592's 4 → 20 pairs differenced away. A `timed` mode prints one nanosecond figure per line. A counting global allocator reports allocations and bytes. |
| `probe/src/corpus.rs` | The differential corpus checker. Change 0592's signature — open outcome, `text()`/`extract_text()`, `paragraph_count()`, `paragraphs()`, every `paragraph(i)` to one past the end, `paragraph_text(i)`, `tables()`, `elements()`, `blocks()`, both facades — plus `write_text_to` under five sink configurations (default; a 64-byte output ceiling; a one-object ceiling; a `\|` separator with empty objects excluded; a sink that refuses its third write), signing the exact accepted bytes, the write-call count, and either the `TextOutputReport` or the complete `TextOutputError` with its retained partial progress. |
| `probe/Cargo.toml.template` | The probe manifest; `__ROOT__` is replaced by each leg's tree, so all four legs are the same probe source against different crates. |
| `probe/revert-0592.patch` | The reverse of change 0592's two source files, applied to `c7326f680` to build the **rev0592** leg. |
| `probe/layout-control.patch` | The **layoutctl** leg (0623's method): the `OnceLock` field and the `get_or_init` initializer are linked exactly as on base, but `from_part` fills the cell, so the lazy path is never taken. `C − A` is therefore the field and the code layout alone, and `B − C` the placement of the scan alone. |
| `probe/repro0592.sh` | The part (a) harness reproduction: `docx_file_eager_paragraph_count`, `docx_file_eager_full_text` and `docx_file_source_full_text`, 60 samples, order `A1 B1 B2 A2`, CPU 18. |
| `probe/firstcall-ns.sh` | The first timed call in a fresh process, timed in-process exactly as the harness child times it, one process per sample, three preparations, `A1 B1 B2 A2`. |
| `probe/layoutctl.sh` | The same measurement with the layout control interleaved: `A1 C1 B1 B2 C2 A2`. |
| `probe/firstcall-perf.sh` | `perf stat` (cycles, instructions, branch misses, dTLB load misses, page faults) over the reps = 0 / reps = 1 pair, 40 processes per leg per value. |
| `probe/callgrind-firstcall.sh` | The callgrind isolation of the same pair, three legs, plus the per-symbol profiles the scan count is read from. |
| `probe/alloc.sh`, `probe/callgrind.sh`, `probe/timing-sink.sh` | The part (b) allocation, instruction and paired-timing captures. |
| `probe/analyze.py`, `probe/nsstats.py`, `probe/layoutstats.py`, `probe/perfstats.py`, `probe/csvstats.py` | The five analysis scripts, exactly as run. |
| `counts/allocations.txt` | Allocations and allocated bytes per operation at 24, 200 and 10,000 paragraphs, both legs, from the reps 4 → 20 isolation pairs. Deterministic. |
| `counts/instructions-sink.txt` | Instructions per operation at the same three shapes, both legs, from callgrind isolation pairs. |
| `counts/instructions-firstcall.txt` | Instructions for the selector's timed region — the reps 0 → 1 pair — on all three part (a) legs, with each leg's untimed preparation total beside it. |
| `counts/first-call-symbols.txt` | The per-symbol extract: inclusive Ir for `scan_word_element_ranges::<ParagraphIndex::from_xml::{closure#0}>` at reps 0 and 1 on each leg. This is the measurement that says the scan runs twice per harness sample on the before leg and once on the after leg. |
| `counts/callgrind-totals.txt` | The 74 raw callgrind `summary:` totals every instruction figure is differenced from, one line per run. |
| `counts/first-call-perf.csv` | The raw `perf stat` rows: 5 counters × 3 operations × 2 legs × 2 reps values × 40 repeats. |
| `counts/first-call-perf-summary.txt` | Difference of medians over those rows. Instructions and page faults are the usable columns; cycles are retained because they are cited as **inconclusive** (the enclosing process is 78–81 M cycles with an IQR of 3.2–4.8 M against a sub-1 M signal). |
| `counts/first-call-harness.csv` | A first, discarded attempt at the same counters that took the *median of per-pair differences* instead of the difference of medians. Retained because the record's "cycles are inconclusive at this scope" statement is read from how badly it behaved: A/A floors of 28% to 139%. |
| `corpus/testdata-{base,after}.txt` | The 62 `.docx` fixtures under `test-data/`, signed as above. Byte-identical between legs. |
| `corpus/fuzz-{base,after}.txt` | The 270 `.docx` inputs retained under `docs/performance/results/` by the change-0483 and change-0495 fuzz campaigns. Byte-identical. |
| `corpus/harness-{base,after}.txt` | The harness's own generated `docx_file_eager_paragraph_count` corpus file — 20 members, 16.79 MB, a 200-paragraph 29,027-byte main document. Byte-identical. |
| `timing/fs-{A1,B1,B2,A2}.json` | The part (a) harness reproduction, one JSON per leg, 60 samples per selector. |
| `timing/repro-summary.txt` | Its per-leg p50/mean/p95/p99 and the paired deltas in both orders, with the A/A floor. |
| `timing/firstcall-summary.txt` | The three-preparation first-call measurement. |
| `timing/layoutctl-summary.txt` | The layout control, both fixtures, all three preparations. |
| `timing/sink-summary.txt` | The part (b) paired timing over seven sink cases and four `text()` controls. |
| `timing/firstcall-raw.tar.gz` | Every raw per-sample nanosecond file behind those three summaries (one integer per line, 40 or 60 lines per leg per case). |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `gates.txt` | The tail of every gate. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree trees. |

## Provenance

- Base: `c7326f680d4e6a22e0dc9e14bcd0e5e70a92dc5e` (change 0630), branch
  `feat/office-format-completeness`.
- Branch: `perf/0643-docx-paragraph-count-and-sink-text`, which carries this
  packet in its single commit.
- Four legs, all built `--release --locked` from the same lockfile:
  - **rev0592** — `/home/zhuhe/code/litchi-worktrees/0643-rev0592`, the base with
    `probe/revert-0592.patch` applied (change 0592's two source files reversed).
  - **layoutctl** — `/home/zhuhe/code/litchi-worktrees/0643-layoutctl`, the base
    with `probe/layout-control.patch` applied.
  - **base** — the shared read-only detached checkout
    `/home/zhuhe/code/litchi-worktrees/before-c7326f680`, never modified.
  - **after** — `/home/zhuhe/code/litchi-worktrees/0643`, this branch.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, ext4 root;
  rustc 1.95.0, valgrind 3.26.0, quick-xml 0.41.0.
- Every timed process was pinned to **CPU 18** with `taskset`; the callgrind runs
  used CPU 19; the deterministic allocation counts were not pinned, because the
  probe's counting allocator is not timing dependent and pinning them would have
  contended with the timed legs. Seven other agents were building and measuring
  on the other cores throughout; the A/A and C/C floors in the timing summaries
  are the only statement this packet makes about that.
- Binary sha256s are in `decision.json`. Every timed binary was copied out of its
  Cargo target directory before use (change 0627: a concurrent build relinked one
  mid-run).

## What this packet does not establish

The synthetic probe fixture is marker-free — it carries no `mc:` markup, so
`visible_document_xml` takes its cheap presence-scan path and the MCE term is a
lower bound for real producer files (change 0587, XML-1); the harness corpus file
and the 333 corpus documents are the counterweight. The instruction counts are
callgrind's, which rank work rather than latency, and callgrind counts
`rep movsb`/`rep stosb` per byte (change 0604 measured a 35× overstatement) — the
`__memcpy_avx_unaligned_erms` line is 60% of every part (a) profile for that
reason, which is why the part (a) result is read from the per-symbol scan count
and the timed-region difference rather than from the totals. The paired timing is
warm, in-memory or warm-page-cache, on one host, with other agents active. Native
cycles for the single timed call of part (a) are explicitly inconclusive. No
cold-cache, physical-device, range-source, peak-RSS, concurrency-scaling or
cross-platform result is taken or claimed.
