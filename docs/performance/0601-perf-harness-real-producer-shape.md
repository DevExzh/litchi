# 0601: a real-producer shape in the harness, and the first baseline taken on it

Status: retained, harness and corpus work. `performance_claim: none` — this
record carries a deterministic corpus census, a typed-refusal census and a
descriptive paired baseline with its own A/A floor, not a claim-registry entry.
**No file under `crates/` was modified.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is the first prerequisite named by change
[0587](0587-remaining-opportunity-survey.md): finding 1 ("the measured path is
not the path real files take") and gap 2 of its evidence table. Every planning,
selected-cell and commit record from 0362 to 0553 measured generated corpora,
and change [0032](changes/0032-xlsx-x14ac-descent.md) recorded that those
worksheets contain neither `dyDescent` nor MCE markup. Items **XML-1** (rank 1),
**XML-2** (rank 6), **XML-3** (rank 13) and **XLSX-2** (rank 12) of the survey's
queue all name a producer-shaped corpus as the thing they need before any
design record. This change adds one, proves it is deterministic, proves it
trips the same library gates a real Excel file trips, and takes the first
baseline on it.

## What was changed

Four files under `tools/`, two documents. Nothing in any production crate.

### A new generator: `tools/perf-baseline/src/producer_shape.rs`

One module, built on one idea: **rewrite a package the production writers
themselves produced**, and author only the parts whose markup is under test.
The surrounding package is therefore exactly as valid as the corpora that
already exist, and the generator is a pure function of its shape — no clock, no
PRNG, no ambient state.

**XLSX.** The base is `build_xlsx_workbook`, the existing integer-grid writer.
Every worksheet is replaced with one authored here, carrying what Excel writes
by default:

- a root with `mc:Ignorable="x14ac xr xr2 xr3"`, the matching `xmlns:mc`,
  `xmlns:x14ac`, `xmlns:xr`, `xmlns:xr2`, `xmlns:xr3` declarations, and an
  `xr:uid`;
- `x14ac:dyDescent` on `sheetFormatPr` and on every `<row>`, with `spans`;
- a `<cols>` block before `sheetData`;
- `pageMargins`, and `pageSetup r:id` on the relationship-bearing worksheet.

Around them: `xl/sharedStrings.xml` with a `count`/`uniqueCount` pair and 40% of
one worksheet's cells stored as `t="s"` references; an
`xl/worksheets/_rels/sheetN.xml.rels` pointing at a fixed 1,024-byte
`xl/printerSettings/printerSettings1.bin`; and an `mc:Ignorable="x15 xr xr6 xr10
xr2"` workbook root with its five namespace declarations.

Two sizes. `producer-medium` is 4 worksheets of 32 × 32, the `XlsxShape::Medium`
size. `producer-dense` is 3 worksheets of 128 × 128 — the dense size the XLSX
planning guard already uses, not `XlsxShape::DenseWide`'s 256 × 256, for the
reason given under *Limitations*.

**DOCX.** The base is `semantic_docx_bytes(Medium)`. `word/document.xml` gets
the Word 2013 root namespace set — 17 declarations, `wpc` through `wps` — with
`mc:Ignorable="w14 w15 wp14"`, and every `<w:p>` gets the `w14:paraId`,
`w14:textId` and `w:rsidR` attributes Word writes.

**PPTX.** The base is `semantic_pptx_bytes(Medium)`. Every slide gets one
`mc:AlternateContent` wrapper of the kind PowerPoint emits for a chart or
3D-text shape: an `mc:Choice xmlns:a14="…/drawing/2010/main" Requires="a14"`
branch and an `mc:Fallback` branch, both cloned from a `<p:sp>` the production
writer authored and re-identified so the shape catalog stays unique.

### Three XLSX variants, because one archive cannot serve both scenarios

The work uncovered a library fact the survey did not have. The source-backed
**value-only editor refuses five separate parts of the producer signature**, and
each of the five is universal in Excel output. Its reach on real Excel files is
therefore not *reduced*, as survey item XLSX-2 assumed — it is *zero*, and it is
zero for reasons that have nothing to do with shared strings alone.

So the XLSX family is built in three variants per shape:

| Variant | Carries | Selectors |
| --- | --- | --- |
| `read` | the complete producer signature | open, selected cell |
| `edit` | the namespace declarations and `<cols>` — the largest subset the editor admits | planning, one-cell edit/save |
| `control` | none of it: the marker-free counterpart of the same grid | the control selectors |

The `edit` variant is still producer-shaped where it matters for the ranked
candidates: the namespace **declarations alone** already defeat
`source_stream_eligible` (`crates/litchi-xlsx/src/raw/worksheet/mod.rs:60`,
which keys on the namespace string occurring anywhere in the part) and already
send the whole part through the MCE codec's rewrite (whose presence scan keys on
the same string), and `<cols>` alone already makes the selected-cell stream
ineligible. What it drops is exactly what the editor refuses, and each refusal
is proven rather than asserted.

### Sixteen opt-in selectors, and two options

`Case` grows from 441 to 457 selectable names. None is in `Case::DEFAULT`; no
existing selector, corpus identity or default-matrix row changes, and the
checked default catalog SHA-256 does not move.

| Selector | Scenario |
| --- | --- |
| `xlsx_producer_{medium,dense}_source_open` | `SourceBackedWorkbook::from_read_at` |
| `xlsx_producer_{medium,dense}_source_selected_cell` | one middle cell of the shared-string worksheet; the owner and the worksheet are resolved outside timing |
| `xlsx_producer_{medium,dense}_source_planning` | `SourceBackedEditor::edit_sheets`, exactly the interval `xlsx_planning_guard` measures |
| `xlsx_producer_{medium,dense}_source_one_edit_save` | plan, set one cell, commit, publish to a bounded counting sink |
| `xlsx_producer_{medium,dense}_control_selected_cell` | the same selected cell over the marker-free grid |
| `xlsx_producer_{medium,dense}_control_planning` | the same planning interval over the marker-free grid |
| `docx_producer_source_selected_paragraph` | one middle paragraph through `document().paragraph(index)` |
| `pptx_producer_source_selected_slide` | `text_and_name()` of one middle slide |
| `xlsx_real_file_source_{open,selected_cell}` | the same two read flows over a caller-named file |

`--real-file PATH` is the only input in this harness whose bytes come from
outside the process. It is bounded to 32 MiB, both selectors that use it are
opt-in and absent from the default matrix, and the file's path, size and
SHA-256 are bound into the corpus identity. The target cell and its expected
value are derived from the file, not assumed.

`--producer-evidence PATH` writes the marker and refusal census for every
producer corpus a run builds, under the schema
`litchi.perf-baseline.producer-shape-evidence.v1`. This is the first-class
census the survey's evidence gap 3 asks for: instead of each record
re-deriving "41 of 60 sheets declare `mc`" by hand, a run states, per part,
which producer markers it carries and which library gate each one trips.

## Why it is sound

**Nothing in a production crate moved.** `git diff --stat` touches
`tools/perf-baseline/{README.md,src/lib.rs,src/producer_shape.rs}`,
`tools/test_perf_baseline_source_policy.py` and two documents. No refusal
moves, no output byte changes, no fence or defence is weakened, no typed limit
is relocated, because no production code ran differently.

**The default matrix is frozen.** The schema-2 default catalog is generated
from `results/perf-regression-default-manifest-v1.json`, which covers exactly
`Case::DEFAULT`; a case that is not in `Case::DEFAULT` never reaches it. A
source-policy test asserts that none of the sixteen new variants appears in the
`DEFAULT` array, and the checked `catalog_sha256`
`f03c9f56…854e41a` and `content_set_sha256` `8e629fef…3452a5` are unchanged.

**No new unsafe, no ambient I/O beyond the declared opt-in.** The module is
inside the crate's `#![forbid(unsafe_code)]`; a source-policy test asserts it
contains no `unsafe`, no `#[global_allocator]`, no `std::env` and no `Command`,
and that it performs exactly one whole-file read — the bounded `--real-file`
one.

**ADR reading.** ADR 0005's mandatory validation, ADR 0006's preservation
contract and ADR 0003's readback requirement are untouched: this change adds
inputs and observers, not code paths. The refusals the corpus proves are the
library's existing typed refusals, recorded verbatim; nothing here argues that
any of them should change, and the survey's ADR table already says XLSX-2 and
XML-1/2/3 each need a frozen design record before code.

## Measured

### Deterministic counts, first

**Generation is deterministic.** Two independent release-mode processes built
every corpus and emitted the census sidecar; `runs/determinism.diff` and
`runs/corpus-identity.diff` are both empty. Archive identities:

| Corpus | Members | Archive bytes | Part bytes | Archive SHA-256 |
| --- | ---: | ---: | ---: | --- |
| `xlsx-producer-medium-read` | 12 | 22,739 | 127,499 | `55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8` |
| `xlsx-producer-medium-edit` | 9 | 20,472 | 123,378 | `7d94095cccf1ccddfb9adcd19f25894fd0a78b126a395271564e5acbbdbbfec5` |
| `xlsx-producer-medium-control` | 9 | 19,787 | 121,550 | `d4db82db66d8c25cd3c667b78396d0c5772213c6fee9491af15d9a6daf5a4e60` |
| `xlsx-producer-dense-read` | 11 | 204,756 | 1,471,545 | `3734cf101df8e215207c9f121efbd9cfde76180532436f6eb0fa8afd5d9b70b8` |
| `xlsx-producer-dense-edit` | 8 | 197,640 | 1,456,543 | `b81f85ac44ad021824b2a2dc81f1680347b7ff8a2373c7778f979df83f73c6f0` |
| `xlsx-producer-dense-control` | 8 | 197,093 | 1,455,169 | `5c5fd4520e809c0b2a68b829115db8b6ad58db37477ad5dc3c555ffc04c4065d` |
| `docx-producer-medium` | 12 | 10,439 | 40,147 | `259f9511046c5ed383b54fc078ea33aec8d80fed6cb03400c9e1dbcb9af7087b` |
| `pptx-producer-medium` | 61 | 42,531 | 4,546 | `c4cd7ca569bf47cfbff77d8b1457b110a99d6913698911a297266ddee8c9f9cb` |

**The generated shape trips the same gates a real Excel file trips.** The census
is produced by one code path for generated and real input, so the two columns
are comparable by construction:

| Fact | `xlsx-producer-medium-read` `sheet2.xml` | `Excel_file_with_trash_item.xlsx` `sheet1.xml` |
| --- | --- | --- |
| bytes | 32,728 | 209,931 |
| `mc:Ignorable` | `x14ac xr xr2 xr3` | `x14ac` |
| `x14ac` namespace | yes | yes |
| `dyDescent` occurrences | 33 | 681 |
| `<cols>` | yes | yes |
| shared-string cells | 410 of 1,024 (40%) | 3,403 |
| shared-string part | `xl/sharedStrings.xml` | `xl/sharedStrings.xml` |
| workbook `mc:Ignorable` | `x15 xr xr6 xr10 xr2` | `x15` |
| `source_stream_eligible` | **false** | **false** |
| ineligibility reasons | markup-compatibility namespace, `x14ac` namespace, `dyDescent` | markup-compatibility namespace, `x14ac` namespace, `dyDescent` |
| selected-cell stream ineligible on `<cols>` | yes | yes |

The real fixture is the same 209,931-byte worksheet the 0587 survey's XML
section profiled, so the two records are talking about the same file.

**The typed-refusal census.** Each row is the admitted package plus exactly one
producer fact, because the gates fire in package-then-part order and an archive
carrying several can only witness the first. All six are reproduced verbatim in
`runs/evidence-1.json`:

| Producer fact | Typed refusal from the value-only editor |
| --- | --- |
| `mc:Ignorable` / `x14ac:dyDescent` on the worksheet | `value-only edits refuse attribute 'mc:Ignorable' on 'worksheet'` |
| `pageMargins` | `value-only edits refuse dependency-bearing or unknown element 'pageMargins'` |
| the shared-string part | `value-only edits refuse workbook relationship '…/relationships/sharedStrings'` |
| a worksheet relationship | `value-only edits refuse worksheet relationships` |
| `mc:Ignorable` on the workbook | `value-only edits refuse attribute 'mc:Ignorable' on 'workbook'` |
| the complete signature | `value-only edits refuse attribute 'mc:Ignorable' on 'workbook'` (the first gate that fires) |

The sites are `crates/litchi-xlsx/src/cell_values/validation.rs:420` (the
element whitelist), `:478` (every qualified attribute except `r:id` on a
workbook `sheet` and `xml:space` on `t`), and
`crates/litchi-xlsx/src/cell_values/snapshot.rs:388,444` (worksheet
relationships) and `:1729` (workbook relationships). This is reported, not
fixed: each is a deliberate narrowing under 0362's "deliberately narrow" scope,
and moving any of them is a frozen-design-record question.

### Paired timing

20 warm-ups, 100 samples per selector, `--release --locked`, pinned to CPU 23,
two identical legs (A and B) in the same window. Seven other measurement agents
were active on the host; the A/A leg is what states the floor, and it is
reported per case in `runs/summary.json`.

| selector | corpus | read part bytes | A p50 ms | A p95 ms | A p99 ms | B p50 ms | A/A p50 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xlsx_producer_medium_source_open` | `xlsx-producer-medium-read` | — | 0.106 | 0.113 | 0.114 | 0.105 | 1.70% |
| `xlsx_producer_medium_source_selected_cell` | `xlsx-producer-medium-read` | 32,728 | 13.311 | 13.587 | 13.651 | 13.430 | 0.89% |
| `xlsx_producer_medium_source_planning` | `xlsx-producer-medium-edit` | 29,013 | 11.538 | 11.630 | 11.662 | 11.517 | 0.18% |
| `xlsx_producer_medium_source_one_edit_save` | `xlsx-producer-medium-edit` | 29,013 | 13.759 | 14.117 | 14.174 | 13.463 | 2.15% |
| `xlsx_producer_medium_control_selected_cell` | `xlsx-producer-medium-control` | 30,998 | 1.890 | 1.900 | 1.916 | 1.869 | 1.09% |
| `xlsx_producer_medium_control_planning` | `xlsx-producer-medium-control` | 28,556 | 0.620 | 0.626 | 0.628 | 0.602 | 2.93% |
| `xlsx_producer_dense_source_open` | `xlsx-producer-dense-read` | — | 0.095 | 0.103 | 0.105 | 0.092 | 3.00% |
| `xlsx_producer_dense_source_selected_cell` | `xlsx-producer-dense-read` | 504,407 | 195.537 | 197.681 | 199.655 | 194.496 | 0.53% |
| `xlsx_producer_dense_source_planning` | `xlsx-producer-dense-edit` | 465,059 | 168.972 | 173.070 | 173.900 | 166.451 | 1.49% |
| `xlsx_producer_dense_source_one_edit_save` | `xlsx-producer-dense-edit` | 465,059 | 189.209 | 193.382 | 194.781 | 189.608 | 0.21% |
| `xlsx_producer_dense_control_selected_cell` | `xlsx-producer-dense-control` | 495,284 | 28.800 | 29.390 | 29.433 | 28.599 | 0.70% |
| `xlsx_producer_dense_control_planning` | `xlsx-producer-dense-control` | 464,601 | 8.881 | 8.930 | 8.951 | 8.817 | 0.73% |
| `docx_producer_source_selected_paragraph` | `docx-producer-medium` | 40,147 | 0.487 | 0.493 | 0.497 | 0.496 | 1.80% |
| `pptx_producer_source_selected_slide` | `pptx-producer-medium` | 4,546 | 1.383 | 1.393 | 1.400 | 1.385 | 0.18% |
| `xlsx_real_file_source_open` | `xlsx-real-file` | — | 0.123 | 0.130 | 0.131 | — | — |
| `xlsx_real_file_source_selected_cell` | `xlsx-real-file` | 209,931 | 53.634 | 55.413 | 55.586 | — | — |

### What the baseline says

The A/A floor over the fourteen generated selectors is **3.0% at p50 at
worst and 1.09% median**, well inside the 5% review threshold this host
usually shows, so the ratios below are not floor artefacts. Every leg is
reported; nothing got worse, because nothing changed.

**The producer signature costs between 6.8× and 19× on the two scenarios that
have a byte-comparable control.** The control corpus is the same grid with the
whole producer signature removed. The two rows are not measuring the same
subset: the selected-cell rows compare the *complete* signature against
nothing, while the planning rows compare only the namespace declarations and
the `<cols>` block against nothing, because the value-only editor refuses
everything else.

| scenario | producer p50 | marker-free control p50 | ratio |
| --- | ---: | ---: | ---: |
| medium selected cell | 13.31 ms | 1.89 ms | **7.04×** |
| dense selected cell | 195.54 ms | 28.80 ms | **6.79×** |
| medium planning | 11.54 ms | 0.62 ms | **18.61×** |
| dense planning | 168.97 ms | 8.88 ms | **19.03×** |

Two things are worth stating plainly about the planning rows, which are the
larger ratio. First, their producer leg carries **no `mc:Ignorable`, no
`x14ac:dyDescent`, no `pageMargins`, no shared strings and no relationships** —
the value-only editor refuses all five. The ~19× is therefore what the
*namespace declarations and the `<cols>` block alone* cost, because those two
facts are enough to send the whole part through the MCE codec's rewrite and to
take 0546's fused validate-and-parse traversal off the table. Adding the rest
of the signature could only make it worse; it cannot be measured on this path
because the editor will not accept it. Second,
the ratio is stable across a 16× change in worksheet size (18.6× at 29 KB,
19.0× at 465 KB), and both legs scale linearly in part bytes, so this is a
per-byte constant rather than a fixed overhead.

**The open scenario is unaffected**, as expected: a source-backed open reads the
catalog and defers every worksheet, so it stays near 100 µs on every variant and
on the real file. That is the control that says the cost above is worksheet
work, not package work.

**A real Excel file agrees.** One selected cell of
`Excel_file_with_trash_item.xlsx` — the same 209,931-byte worksheet the 0587
survey profiled — costs **53.63 ms** through the same public API, against
122.9 µs to open the workbook. Per worksheet byte that is 255 ns, against
407 ns for the generated medium shape and 388 ns for the generated dense shape:
the generated corpus is about 1.5× more expensive per byte, because its
`mc:Ignorable` names the full `x14ac xr xr2 xr3` set where this file's names
only `x14ac`, and because its grid is denser in `<c>` elements per byte. The
generated shape therefore *overstates* the producer cost slightly rather than
understating it, which is the safe direction for a corpus whose purpose is to
make a candidate falsifiable.

**DOCX and PPTX are baselines without controls.** One paragraph of a Word 2013
`document.xml` (40,147 bytes, 17 root namespace declarations) is 0.487 ms, and
one `mc:AlternateContent`-bearing slide (4,546 bytes) is 1.383 ms. No
marker-free control variant was built for either format, so these are numbers
to price changes 0597 and 0603 against, not ratios.

## Correctness evidence

**Ten new unit tests** in `producer_shape::tests`, all passing:

- generation is byte-identical across two in-process builds, for all three
  XLSX variants of the medium shape and for the DOCX and PPTX shapes (the dense
  shape is proven across two independent processes in the packet instead, so
  the debug-profile gate stays seconds rather than minutes);
- the `read` and `edit` variants are distinct corpora with distinct identities;
- every producer worksheet carries the namespace declarations and `<cols>`, is
  not `source_stream_eligible`, and trips the documented ineligibility reasons;
  the `read` variant additionally carries `mc:Ignorable` and `dyDescent`;
- the control variant carries none of it and **is** `source_stream_eligible`;
- the shared-string share is exactly 40% of 1,024 cells, and the worksheet
  relationship resolves to the `printerSettings` target;
- the `read` variant proves all six typed refusals in order;
- the `edit` and `control` variants are admitted by the value-only editor;
- the DOCX document carries the `w14` and `w15` namespaces and `w14:paraId`;
- the PPTX slide carries exactly one `mc:AlternateContent`;
- a missing `--real-file` path is refused with a message naming the option.

**Three new source-policy tests** in
`tools/test_perf_baseline_source_policy.py`: the module owns no unsafe or
ambient surface; the `--real-file` selectors are bounded, self-identifying and
have exactly one whole-file read; and none of the sixteen new cases appears in
`Case::DEFAULT`.

**Per-sample oracles inside the timed selectors.** The planning selector checks
that the transaction selects one worksheet, that the empty commit is an exact
no-op, and that the source bytes are unchanged. The edit/save selector checks
that the commit reports a change, that every retained sample produces the same
output digest, and that the source bytes are unchanged. Every read selector
compares the value it read against the value the corpus derived at
construction.

**Gates.** Run in the worktree; tails in `gates.txt`.

| Gate | Result |
| --- | --- |
| `cargo fmt --all --check --manifest-path tools/perf-baseline/Cargo.toml` | exit 0 |
| `cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets` | exit 0, no warning |
| `cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml` | 501 passed, 7 failed, 1 ignored, 1,948 s — **all seven pre-existing**, see below |
| `cargo doc --locked --manifest-path tools/perf-baseline/Cargo.toml --no-deps` | exit 0, no rustdoc warning |
| every `tools/test_*.py` (26 modules) | 22 OK, 4 failed — **all four pre-existing** |

No crate under `crates/` was touched, so no production crate's `clippy`, `test`
or `doc` gate is in scope for this change.

**The seven Rust failures are one pre-existing failure and six cascades from
it.** `tests::xls_source_backed_lifecycle_selectors_are_matched_and_local`
asserts `open_reads_zero_worksheet_payload == [true]` and gets `[false]`. It
fails identically, in isolation, on the untouched `before-f8cf7d2a1` checkout
built with an external `CARGO_TARGET_DIR` — the transcript is in `gates.txt`.
Its panic poisons the shared allocation-metrics mutex, and the six
`PoisonError` failures are every later test that unwraps it; all six pass when
run alone on this branch, which `gates.txt` also records. Nothing in this change
touches XLS, allocation metrics or that mutex.

**The four Python failures** — `test_check_crate_boundaries`,
`test_native_odf_resave`, `test_perf_claims` and `test_perf_compare` — produce
byte-identical output on the untouched before checkout. Both sweeps are in
`gates.txt`.

**`test_corpus_manifest_v2` and `test_crud_coverage_index` pass**, which is the
mechanical statement that the checked default catalog SHA-256 did not move.

## Validation preserved

Nothing about validation changed, because no production code changed. The one
thing worth stating explicitly: the generated worksheets are **compact** — no
whitespace between elements and no newline after the XML declaration — although
real Excel writes `\r\n` there. That is deliberate. Defect 1 of change 0587
records that `validate_authored_xml`
(`crates/litchi-opc/src/pkgwriter.rs:988`) refuses that `\r\n` for the whole
save, so a corpus carrying it could not exercise the edit/save scenario at all.
The omission is named here rather than left for a reader to discover, and it is
the one place where the shape deliberately differs from Excel's bytes in a way
the library can observe.

## Limitations

- **Nothing here makes anything faster.** It makes a slow path measurable. No
  speedup, regression, allocation, peak-RSS, cold-cache, physical-I/O,
  range-source, concurrency or cross-platform result is claimed.
- **The dense shape is 128 × 128, not `XlsxShape::DenseWide`'s 256 × 256.** A
  256 × 256 producer worksheet costs seconds per sample with the MCE codec on
  the path, so a 100-sample baseline would take about half an hour per selector
  on a host shared by eight agents. 128 × 128 is the dense size the XLSX
  planning guard already uses and is 2.4× the real fixture the survey profiled.
- **The producer shape is one shape, not a population.** It is shown to trip the
  same gates as one real fixture. It is not shown to be representative of the
  real-producer population; the survey's own fixture census (41 of 60 sheets
  declare `mc`, 27 of 60 carry `<cols>`, 77 of 95 have a shared-string part, 57
  of 95 have worksheet relationships) remains the breadth evidence.
- **The `edit` and `control` variants deliberately omit parts of the producer
  signature.** Each omission is exactly one typed refusal, and each refusal is
  in the census. A future change that admits more of the signature should widen
  `ArchiveOptions::ADMITTED` and watch the corresponding census row disappear.
- **The new generator identities are not in the schema-2 family map.** That map
  is the source-audited contract for the *default* catalog, and every entry in
  it is asserted to appear there. An unmapped identifier retains null
  `algorithm_id` and `seed_spec`, which is the correct record for an opt-in
  corpus and the only honest record for a caller-named real file. The
  identities are documented in `CORPUS_MANIFEST_V2.md` instead.
- **No selector was added** for the DOCX or PPTX edit and save scenarios, for
  the eager XLSX paths, for file-backed or range sources, or for XLSB on a
  producer shape. The survey's gap 5 stays open.
- **No callgrind pair, no hardware counters, no allocation profile** were taken
  for this batch. The instruction-level attribution XML-1's design record will
  need is the next measurement, and it should be taken on these corpora.
- **Seven Rust tests and four Python tests fail on this host and every one is
  pre-existing**, reproduced on the untouched `before-f8cf7d2a1` checkout and
  recorded in `gates.txt`. The Rust root cause,
  `xls_source_backed_lifecycle_selectors_are_matched_and_local`, is a real
  standing failure in the harness's XLS lifecycle assertion that this change
  did not introduce and does not fix; it is worth a look from whoever owns
  change 0595's area.

## Retained evidence

[`results/change-0601/`](results/change-0601/README.md) — the contents table,
provenance with commit and binary hashes, the determinism diffs, the marker and
refusal census sidecars, both timing legs, the summary, the gate tails and the
four log paragraphs.
