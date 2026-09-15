# Evidence: change 0602, XLSX real-producer admission

Change record: [`0602-xlsx-real-producer-admission-design.md`](../../0602-xlsx-real-producer-admission-design.md).

Disposition: design, retained. `performance_claim: none`. **No production code
changed.** This packet holds the sizing measurements that were taken to write
the design, and the scripts and probe that produced them. Nothing in it is a
before/after comparison of a candidate, because there is no candidate.

## Contents

| Path | What it is |
| --- | --- |
| `probe/` | The scratch probe, retained in full: `Cargo.toml` with path dependencies on `litchi-xlsx`, `litchi-sheet` and `xml-minifier`, and `src/main.rs`. Four modes. `census` opens each package through the source-backed read door and both value-editor doors and reports admission or the typed refusal, then runs a complete one-cell plan, commit and publication cycle. `compact` runs `xml_minifier::audit::verify_authored` over raw part bytes, the same call `litchi-opc` applies to the original and the replacement of every replaced part. `edit` runs N plan-and-commit cycles against one retained editor, for callgrind isolation pairs; `--insert` selects the row-creating insert that disables change 0525's reduced readback. `bench` emits one duration per line for the timing legs. |
| `census/census.py` | Structural census of a corpus: shared-string part, `t="s"` in the first sheet, first-sheet and any-sheet worksheet relationships, MCE and x14ac markers, sheet byte size. |
| `census/census.tsv` | That census over the 95-fixture corpus. Reproduces change 0587's counts: 95 files, 77 with a `sharedStrings` part. |
| `census/ladder.py` | Structural reimplementation of the value editor's first five admission gates (G1 package relationships, G2 workbook relationships, G3 worksheet relationships, G4 the shared-string planning parse, G5 `stored_entry_is_supported`), with the source line of each in its docstring. |
| `census/ladder.tsv` | The ladder over the 95-fixture corpus: which gate refuses each file first, and which gates would refuse it independently. All 95 stop at G1; G5 refuses none. |
| `census/admission.tsv` | The editor's own verdict on all 95 fixtures, captured verbatim from `probe census`. 93 refuse at G1 with the typed message; 2 fail earlier in the read door. |
| `census/synthetic-twins.tsv` | `probe census` over the seven synthetic gate twins. Each clause of `stored_entry_is_supported` is refused upstream; `numeric` and `inline` run a complete cycle. |
| `census/derived-fixtures.tsv` | `probe census` over the derived fixtures in their `plain`, `rich` and `cm` variants. `rich` refuses on element `r`, `cm` on attribute `cm`, `plain` plans and commits. |
| `census/compact.tsv` | `xml_minifier::audit::verify_authored` over the `xl/workbook.xml` and first worksheet of all 95 fixtures. 94 of each are non-compact; the one compact package is the one litchi wrote. |
| `census/geometry.tsv` | Cell, row, shared-string-cell and byte counts of the edited worksheet part, in the fixture as shipped and in its projection. |
| `fixtures/degate.py` | Derives an admissible twin of a real `.xlsx` by opening G1, G2 and G3 only: drops package-root relationships outside the allow-list and the parts they reach, drops workbook relationships outside the allow-list, folds `sharedStrings` back into the worksheets as inline strings so no cell content is lost, drops worksheet relationship parts and the `r:id`-bearing worksheet children that reach them. `plain` and `rich` variants. |
| `fixtures/project.py` | Projects `workbook.xml` and each worksheet onto the value-only element and attribute allow-lists verbatim, leaving `<sheetData>` untouched, so the producer's row and cell geometry survives. `plain`, `rich` and `cm` variants. |
| `fixtures/synth.py` | Builds the seven minimal synthetic twins, each inside every allow-list except for the one feature its variant adds, so the refusal it produces names exactly one gate. |
| `cg/run.sh` | The callgrind driver: isolation pairs at N=1 and N=4 samples, `--separate-callers=1`, for both the `set` and the `insert` kind on four fixtures. |
| `cg/diff.py` | Differences a pair into inclusive Ir per operation. |
| `cg/attribution.txt` | The differenced output for all eight pairs. This is the table the record's instruction attribution is read from. |
| `cg/inclusive-tops.txt` | Per-profile `PROGRAM TOTALS` and the `litchi_xlsx` and `litchi_ooxml_common::mce` inclusive lines from each of the sixteen annotations, with absolute paths normalized. The full 3.3 MB annotations are not retained; these are the lines the record cites. |
| `bench/measure.sh` | The timing driver: A1/B1/B2/A2 (set, insert, insert, set) plus S1..S4 (four identical `set` legs, for the A/A floor in the same window), pinned to CPU 18. |
| `bench/stats.py` | Per-leg quantiles and paired deltas in both directions. |
| `bench/legs/` | The raw per-sample durations, 32 files, one per fixture and leg. |
| `bench/summary.txt` | The computed quantiles, paired deltas and A/A floor. |
| `binary.sha256` | SHA-256 of the probe binary every measurement in this packet was taken with. |
| `gates.txt` | The tail of each documentation gate run in the worktree. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

Base commit: `f22f9393598da39272f2185d2dac3d36c4b61d23`
(`perf(docx): build the paragraph index on the first paragraph query (0592)`),
branch `perf/0602-xlsx-real-producer-admission-design`. Nothing under `crates/`
was modified; `git diff f22f93935 -- crates/` is empty.

Probe binary SHA-256
`84de9731098472ca0ae3e6acfc558b5c2bdb71e0323234b94d837b57c854a448`, built
`--release` with `debug = true` from `probe/` against the worktree's crates,
into a `CARGO_TARGET_DIR` outside the worktree. Every census, audit, callgrind
pair and timing leg in this packet was taken with that one binary.

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0;
valgrind 3.26.0. Every measured process was pinned to CPU 18 with `taskset`.
Seven other agents were building and measuring on the host throughout, which is
what the A/A floor in `bench/summary.txt` measures.

Corpus: the 95 `.xlsx` files under `test-data/ooxml/xlsx` and
`test-data/office-interop`, which is the scope change 0587 counted, plus four
fixtures derived from it and from `test-data/poi/test-data/spreadsheet`.

## How to reproduce

```
cd docs/performance/results/change-0602/probe
CARGO_TARGET_DIR=<somewhere outside the worktree> cargo build --release
B=<that target dir>/release/xlsx-admission-probe

# Section 1: the editor's verdict on the 95 real fixtures.
$B census $(ls <repo>/test-data/ooxml/xlsx/*.xlsx <repo>/test-data/office-interop/*/*.xlsx)

# Section 2: the structural ladder.
python3 ../census/ladder.py <repo>/test-data/ooxml/xlsx <repo>/test-data/office-interop

# Section 3: the gate twins.
python3 ../fixtures/synth.py /tmp/synth 256 24 numeric inline rich sst sstfree cm vm
$B census /tmp/synth/*.xlsx

# Section 4: derive, then measure.
python3 ../fixtures/degate.py <in>.xlsx /tmp/d-plain.xlsx plain
python3 ../fixtures/project.py /tmp/d-plain.xlsx /tmp/d-proj-plain.xlsx plain
$B edit /tmp/d-proj-plain.xlsx "<sheet>" "<a1>" 4            # and 1, for the pair
$B edit /tmp/d-proj-plain.xlsx "<sheet>" "<absent-row>" 4 --insert

# Section 5: the publication compactness audit.
$B compact <extracted part bytes>
```

## What this packet does not establish

* **No real producer file was measured as shipped.** Every timing and
  instruction figure is from a derived fixture whose package envelope was
  stripped to the admission surface. The cell geometry is the producer's; the
  package is not.
* The `insert` leg is a proxy for a disabled reduced readback, not a patch, so
  its latency delta is an upper bound on the readback's own share.
* The ladder's G2 to G5 counts are structural; no file reaches them through the
  editor, so only G1 is validated against the editor's own verdict.
* `no_drawing_patriarch`'s A/A floor reached 4.03% at p50 in one pair. Its
  latency numbers are reported, not relied on.
* No claim about publication cost, allocation, peak RSS, cold cache, range
  sources, concurrency, or any host or platform other than this one. No claim is
  registered.
* Nothing here authorizes a change. The design in the record names its own
  prerequisites, admission gates and falsification conditions.
