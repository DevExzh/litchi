# Evidence: change 0592, the lazy DOCX paragraph index

Change record:
[`0592-docx-lazy-paragraph-index.md`](../../0592-docx-lazy-paragraph-index.md).

Disposition: retained. `performance_claim: none`. Two files under
`crates/litchi-docx/` changed; the counts and paired medians here are reported as
evidence and are not registered as claims.

## Contents

| Path | What it is |
| --- | --- |
| `probe/src/main.rs` | The isolation probe. Builds one DOCX in memory with *N* paragraphs carrying the perf-baseline corpus payload string, then repeats one of thirteen main-document operations *R* times. A counting global allocator reports allocations and bytes. Two invocations at different *R* are differenced to price one operation. |
| `probe/src/corpus.rs` | The differential corpus checker. For every `.docx` under a root it prints one line holding a deterministic signature of the eager and the source-backed facade: open outcome, `text()`/`extract_text()` digest and length, `paragraph_count()`, `paragraphs()` digest and length, every `paragraph(i)` up to one past the end, `paragraph_text(i)` (source-backed), `tables()`, `elements()` and `blocks()`. Errors are captured as their `Display` text, so a moved refusal changes the signature. |
| `probe/Cargo.toml.template` | The probe manifest; `__ROOT__` is replaced by the before checkout or by this branch's worktree, so both legs are the same source against different crates. |
| `probe/alloc.sh`, `probe/callgrind.sh`, `probe/cycles.sh`, `probe/coldfirst.sh`, `probe/timing.sh`, `probe/analyze.py` | The six capture and analysis scripts, exactly as run. |
| `counts/instructions.txt` | Instructions per operation, both legs, both shapes, from the callgrind isolation pairs. |
| `counts/callgrind-totals.txt` | The 88 raw callgrind `summary:` totals the table above is differenced from, one line per run. |
| `counts/allocations.txt` | Allocations and allocated bytes per operation, both legs, both shapes, from the same isolation pairs without valgrind. Deterministic: identical across two independent runs. |
| `counts/cycles-200.csv` | `perf stat -e cycles` isolation pairs for four operations on the 200-paragraph document, 15 per leg per order, `before after after before`. 240 rows. |
| `counts/first-call-faults.csv` | Page faults attributable to the first measured call in a fresh process, with and without a harness-style untimed `text()` preparation, 18 pairs per leg. This is the measurement that refutes the first-touch explanation of the `docx_file_eager_paragraph_count` regression. |
| `counts/cold-first-call.csv` | The same first-call experiment in cycles. Retained because it is cited as *inconclusive*: its own A/A floor came out at ±26%. |
| `counts/fixture-sizes.txt` | The probe fixture's archive sizes at both shapes. |
| `corpus/testdata-before.txt`, `corpus/testdata-after.txt` | The 62 `.docx` fixtures under `test-data/`, signed as above. Byte-identical between legs. |
| `corpus/fuzz-before.txt`, `corpus/fuzz-after.txt` | The 270 `.docx` inputs retained under `docs/performance/results/` by the change-0483 and change-0495 fuzz campaigns — seeds, accepted corpora and crash inputs — signed as above. Byte-identical between legs. |
| `timing/` | The paired timing runs in `A1 B1 B2 A2` order, one JSON per leg per selector family, plus `summary.txt` from `analyze.py`. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `gates.txt` | The tail of every gate. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree trees. |

## Provenance

- Base: `08d968f8ec7db27cf1187d01911fd08b9d014d91` (change 0587), branch
  `feat/office-format-completeness`.
- Branch: `perf/0592-docx-lazy-paragraph-index`; the commit hash is in
  `decision.json`.
- Before leg: the shared detached checkout
  `/home/zhuhe/code/litchi-worktrees/before-08d968f8e` at the base commit, never
  modified. After leg: `/home/zhuhe/code/litchi-worktrees/0592`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, ext4 root;
  rustc 1.95.0, valgrind 3.26.0.
- Every measured process was pinned to **CPU 12** with `taskset`. Seven other
  agents were building and measuring on the other 31 cores throughout; the
  A/A floor in `timing/summary.txt` is the only statement this packet makes
  about that.
- Binary sha256s are in `decision.json`; the harness JSONs carry their own
  `binary_identity` block. The instruction and allocation counts were re-taken
  with the exact probe binaries whose hashes `decision.json` records, after one
  operation was added to the probe; the allocation table came out identical and
  the instruction table moved in its fifth significant digit, which is the
  run-to-run spread of callgrind on this host and is smaller than every share
  the record quotes.

## What this packet does not establish

The probe's fixture is synthetic and marker-free: it carries no `mc:` markup, so
`visible_document_xml` takes its cheap presence-scan path and the MCE term in
these figures is a lower bound for real producer files (change 0587, XML-1). The
instruction counts are callgrind's, which rank work rather than latency. The
paired timing is warm, in-memory or warm-page-cache, on one host, with other
agents active. No cold-cache, physical-device, range-source, peak-RSS,
concurrency-scaling, real-producer or cross-platform result is taken or claimed.
