# Evidence: change 0660, the DOCX compaction policy and the shared main-part snapshot

Change record: [`0660-docx-compaction-policy.md`](../../0660-docx-compaction-policy.md).

Disposition: retained. `performance_claim: none`. A committed DOCX edit stops
re-serializing the paragraphs it did not touch, under a public policy whose
default is preservation, and `Package::document_snapshot` stops copying the main
part. Everything the record cites is here; nothing here is registered as a
claim.

## Contents

| Path | What it is |
| --- | --- |
| `probe/Cargo.toml.template`, `probe/src/main.rs` | The scratch probe, built once per leg with `REPLACE_WITH_CHECKOUT` substituted for that leg's checkout and `--features policy` on the change leg only. Two modes: `census <root>` walks every DOCX-family fixture under a corpus root and prints one self-describing line per fixture per aspect; `commit <paragraphs> <compact\|noncompact> <measured>` builds a constant four snapshots of one generated document and runs `edit()`, one `replace_paragraph_text` and `commit()` on `measured` of them, for an isolation pair. |
| `census/census-before.txt`, `census/census-after.txt` | The census on each leg, 630 lines each: 63 fixtures × ten aspects. |
| `census/census-diff.txt` | `diff` of the two, 220 lines, every one of them the `one_edit_optin` aspect the base leg cannot run. Filtering that aspect out makes the two files identical, which is the record's "every published byte and every refusal is unchanged" statement. |
| `census/probe-binaries.txt` | `sha256sum` of the two probe binaries that produced the census. |
| `counts/counts.sh` | The capture script: `litchi-perf-baseline --warmup 0 --samples N --case <case>` under callgrind at N=1 and N=3 for the three DOCX semantic selectors, on each leg, pinned to CPU 15. |
| `counts/incl-<leg>-<case>-<N>.txt` | The top 200 rows of `callgrind_annotate --inclusive=yes --threshold=100` for each of the twelve runs; every symbol the record cites is inside them. |
| `counts/cg-<leg>-<case>-<N>.txt` | Valgrind's own tail for each run, including the harness's verification output. |
| `counts/compare.py` | The isolation-pair extractor and comparator: inclusive Ir per symbol from the annotations, call counts from the raw `callgrind.out` `cfn=`/`calls=` pairs, both differenced between N=3 and N=1 and halved, for both legs side by side. |
| `counts/counts-summary.json` | Its output: every deterministic count the record's first table states. |
| `region/region.sh` | The commit-region capture: the probe's `commit` mode under callgrind at `measured = 1` and `3`, for `compact` and `noncompact` × 24, 200 and 10,000 paragraphs, on each leg, pinned to CPU 15 — 24 runs. |
| `region/region-<leg>-<shape>-<paragraphs>-<N>.txt` | Valgrind's tail for each of those runs. |
| `region/summarize.py`, `region/region-summary.json` | `(Ir(3) − Ir(1)) / 2` per cell, which is the record's second table. |
| `timing/timing.sh` | The paired-timing script: order A1 B1 B2 A2, `--warmup 5 --samples 50` per run, the three semantic selectors and then the four `docx_ordinary_save_*` control selectors per run, pinned to CPU 15. |
| `timing/w1/`, `timing/w2/`, `timing/w3/` | The three windows, all retained. `w1` timed the change binary before a late no-op refactor (two identical match arms merged, three private functions moved); `w2` and `w3` timed the committed binary. `w2` was heavily contended — its `B/B` floor reached −27.05% on one scenario and −60.57% on an ordinary-save control — and the record quotes `w3`. |
| `timing/<window>/A1.json`, `B1.json`, `B2.json`, `A2.json` | The harness reports for the semantic selectors as written, with every per-sample `elapsed_ns` value, the corpus manifests and the binary identity of the leg that produced them. |
| `timing/<window>/ord-A1.json`, `ord-B1.json`, `ord-B2.json`, `ord-A2.json` | The same four runs for the `docx_ordinary_save_*` control selectors, which go through `document_mut()` and never reach the changed code. |
| `timing/summarize-windows.py`, `timing/paired-summary.json` | Pools the two runs of each leg (100 samples per leg) per window and produces the record's third table: p50, mean, p95, p99 per leg, the delta in both directions, and the A/A and B/B floors observed in that window. `timing/analyze.py` is the single-window form. |
| `timing/binaries.txt` | `sha256sum` of the staged harness and probe binaries, with the `w1` exception named. |
| `gates.txt` | The tail of every gate run in the change worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; this batch did not edit those files. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and from disk after this packet was assembled. |

## The census aspects

Ten lines per fixture, all through documented entry points:

| Aspect | What it runs |
| --- | --- |
| `open` | `Package::open(path)` |
| `source` | length and SHA-256 of `/word/document.xml` as opened |
| `gate` | `xml_minifier::audit::verify_authored` on that part with the package writer's own limits — the verdict `CompactionPolicy::PreserveUnmodified` consults |
| `noop_save` | open, `to_stream` with no edit; SHA-256 of the artifact |
| `exact_noop` | open, `edit_document`, `commit` with no operation, `publish_document_edit`, `to_stream`; SHA-256 |
| `managed_edit` | open, `edit_document`, `insert_paragraph(0, ..)`, publish, `to_stream`; SHA-256 (change 0650's shape) |
| `one_edit` | open, `edit_document`, `replace_paragraph_text` on the first paragraph with text, publish, `to_stream`; SHA-256 — the default policy on the change leg |
| `one_edit_optin` | the same edit under `CompactionPolicy::WholeDocument`; `unavailable` on the base leg, which has no policy |
| `one_edit_span` | source versus published `/word/document.xml`: both lengths, the common prefix and the common suffix, so the bytes the edit moved are visible |
| `one_edit_reopen` | reopen the published artifact and report the paragraph, table and block-control counts and how many paragraphs carry the edit |

Totals on the change leg: 63 fixtures, 55 `open ok` and 8 refused at open (all
eight package-level refusals in `litchi-opc`/`litchi-ooxml-common`, unchanged
from change 0650's census); `gate` 1 compact and 54 noncompact; `noop_save` 55
ok; `exact_noop` 54 ok and 1 refused; `managed_edit` 54 ok and 1 refused;
`one_edit` 26 ok and 29 refused. Of the 54 noncompact verdicts, 51 are
`FormattingWhitespace at byte 55`, 2 are `FormattingWhitespace at byte 38` — in
both cases a line break between the XML declaration and the root element — and 1
is `invalid XML declaration boundary`.

## Provenance

- Base: `70d7768cc6dada420ede063f72c88dc99ad30383` (branch
  `feat/office-format-completeness`, carrying change 0652).
- Change branch: `perf/0660-docx-compaction-policy`, worktree
  `/home/zhuhe/code/litchi-worktrees/0660`.
- Before leg: the shared read-only checkout
  `/home/zhuhe/code/litchi-worktrees/before-70d7768cc` at the base commit, built
  into its own `CARGO_TARGET_DIR`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0
  (the workspace pin, used for both harness legs and both probe legs);
  valgrind 3.26.0. Seven other agents were building and measuring on the same
  host throughout; the A/A and B/B floors in `timing/paired-summary.json` are
  what that contention amounted to in each window.
- Builds: `cargo build --release --locked` on both legs, identical flags. Both
  harness binaries were staged outside their Cargo target directories before any
  leg ran.

`lto = true` makes release binaries non-reproducible byte for byte (change
0635), so each timed and profiled binary's digest is recorded.

| binary | sha256 |
| --- | --- |
| `litchi-perf-baseline` (before) | `be066f8268e5dafde992fb4c8a6a55cdb67c941b2a1be9d760f9ebadddda5eb5` |
| `litchi-perf-baseline` (after) | `8a8325b47d58092133dbc81deda484019c7c44055ffa5930ec392e93ccb588b0` |
| `litchi-perf-baseline` (after, timing window `w1` only) | `963b899b949b636646bda11c4b21b0fdaa0f4f3b37ce5064fb1da2bad803bbfd` |
| `docx-compaction-policy-probe` (before) | `e0565740b5c6e75699add514d0043cedbbbff632d6a0c812220fc20232399c6c` |
| `docx-compaction-policy-probe` (after) | `80ee99e095083740ccd9d47dc652e2511e2e45dc60579e7fe69383a98bef013c` |

## Replaying it

```sh
# the probe, once per leg
sed 's|REPLACE_WITH_CHECKOUT|<leg checkout>|' probe/Cargo.toml.template > <dir>/Cargo.toml
cp probe/src/main.rs <dir>/src/main.rs
cargo build --release [--features policy]   # the feature only on the change leg

# the corpus census
<probe> census <repo>/test-data > census-<leg>.txt

# deterministic counts
counts/counts.sh <leg-binary> <before|after> <outdir> <raw-callgrind-dir>
python3 counts/compare.py <outdir> <raw-callgrind-dir>

# the commit region alone
region/region.sh <probe> <before|after> <outdir> <raw-callgrind-dir>
python3 region/summarize.py          # run from <outdir>

# paired timing, once per window
timing/timing.sh <outdir>/<window> <stage-dir> <filesystem-root>
python3 timing/summarize-windows.py <outdir>
```

## What this packet does not establish

Thirteen timed scenarios in three windows on one synthetic corpus and one
fixture corpus, on one host, with two builds. No cold cache, no physical device, no peak RSS, no
allocation profile from the allocator (the harness's allocation binary reports
no metrics for the `docx_semantic_*` family), no concurrency scaling, no
cross-platform result, and no claim about DOCX operations other than a
direct-body paragraph text rewrite, a paragraph insertion, an exact no-op and an
ordinary save. Instruction counts rank work and are not latency; the paired
medians are reported beside the floor measured in the same window and are not
registered as a speedup. The raw `callgrind.out` files were deleted after the
annotations and call counts were extracted; `cleanup.json` records them.

Above all: the default policy's preservation is reachable on **one** of the 55
openable fixtures today, because the OPC writer's authored-XML compactness
contract refuses the rest. That contract is change 0652's decision 2, a separate
record in this wave. This packet is the before-picture for it.
