# 0638: the facade's `.doc` and `.ppt` routes and the ordinary documented OOXML save get selectors, and the save turns out to cost 8 to 62 times more to publish than to serialize

Status: retained, harness-only. Thirty opt-in selectors and a first descriptive
baseline of the two routes change 0587 named as its evidence gap 5.
`performance_claim: none` — this record carries paired medians over three
repeats, deterministic byte accounting and per-corpus determinism proofs, not a
claim-registry entry. **No file under `crates/` was modified.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0587](0587-remaining-opportunity-survey.md)'s harness-gap table states
gap 5 in one line: *"no selectors for the paths this record ranks highest …
missing: DOC and PPT facade opens, OLE2 length-changing saves, **the ordinary
OOXML save**, XLS full text and all cells, file-backed cold part reads, PPTX
capture/commit/apply timers, eager DOCX phases"*. Its core-area blocker section
is blunter: *"No harness selector opens `.doc` or `.ppt` through the facade, so
CORE-1 cannot be A/B-measured today."* Change
[0630](0630-queue-refresh-after-the-first-wave.md) confirmed both halves were
still open after the first wave — *"DOC and PPT facade selectors and an
ordinary-save selector"* — and made its refreshed queue row 10 depend on the
first: *"until then the facade's `.doc` route stays eager by measurement"*.

Three records had to build throwaway probes because of this. Change
[0593](0593-opc-publication-pristine-members.md) priced "Path A" with an
`opc-save-probe` crate outside the workspace and wrote down what a selector
would cost: it *"changes `tools/perf-baseline`'s checked catalog SHA-256, its
selector registry and its coverage-index minimum, and belongs in its own
record"*. Change [0607](0607-pptx-authored-slide-regeneration-design.md) built
`probe0607` for the authored PPTX save. Change
[0609](0609-facade-doc-source-route-design.md) built `facade_doc_route` for the
facade DOC route, over 57 real `.doc` fixtures, and its numbers have been
unreproducible from the repository ever since.

This record registers all three routes. Of 0593's stated price, only one third
is real: the registry grows, the coverage-index minimum is a *minimum* and is
already met, and the checked catalog SHA-256 does **not** move, because none of
the thirty selectors is in `Case::DEFAULT`.

## What was changed

Nine files under `tools/`, two of them new, plus one document under
`docs/performance/`:

* `tools/perf-baseline/src/facade_ole2.rs` (new, 795 lines including its four
  unit tests): the DOC and PPT facade corpora, the six selectors' runner, the
  frozen-oracle gate and the evidence block.
* `tools/perf-baseline/src/ordinary_save.rs` (new, 1,391 lines including its
  four unit tests): the DOCX, XLSX and PPTX corpora and private workspaces, the
  twenty-four selectors' runner, the two determinism proofs, the byte
  accounting and the evidence block.
* `tools/perf-baseline/src/lib.rs`: two module declarations, thirty `Case`
  variants with their `name`/`parse_case` mappings, the
  `is_facade_ole2`/`facade_ole2_scenario` and
  `is_ordinary_save`/`ordinary_save_plan`/`ordinary_save_case` classifiers, the
  exclusion from the generic corpus loop, two explicit refusals in
  `run_case_with_config`, two family blocks in `run()`, the `--ooxml-file`
  option, two `SourceSummary` fields, the usage text and the registry count.
* `tools/perf-baseline/src/ole2_range_source.rs`: a `Format::Doc` arm so
  `--ole2-file` stays the single classification authority for a caller-named
  OLE2 fixture, a shared `classify_inputs` both families call, and `pub(crate)`
  on the three bounded helpers the facade module reuses rather than duplicates.
* `tools/perf-baseline/src/filesystem.rs`: `filesystem_root` factored into a
  `scratch_root(requested_root, label)` the save family calls, so `std::env`
  stays in the one module that already owns it.
* `tools/perf-baseline/Cargo.toml` and `Cargo.lock`: the `litchi` facade
  dependency gains its `doc` and `ppt` features. Both leaf crates were already
  direct dependencies of this harness, so the lock gains two dependency lines
  and no package.
* `tools/test_perf_baseline_source_policy.py`, `tools/perf-baseline/README.md`
  and `docs/performance/CORPUS_MANIFEST_V2.md`.

The tracked diff is +829 / −35 across the eight changed files, plus the two new
modules and this record's packet.

### The six facade selectors

```text
doc_facade_file_open            ppt_facade_file_open
doc_facade_file_full_text       ppt_facade_file_full_text
doc_facade_file_one_paragraph   ppt_facade_file_one_slide_text
```

They call the documented entry points and nothing else: `litchi::Document::open`
and `litchi::Presentation::open`, then `text()`, `paragraph_text(index)` or
`slide(index)` followed by `Slide::text()`. The facade takes a **path**, so —
uniquely in this harness — these run against a real file on the filesystem;
that is the route's defining property, and 0609 established what it buys
(`Document::open` reads the artifact once; the source-backed snapshot reads it
six times).

The corpus is the caller-named file supplied with `--ole2-file PATH`, change
0627's flag, which now recognizes a `WordDocument` stream as well. The two
families share one classifier, so a run may name a DOC, an XLS and a PPT fixture
and each goes to the family that reads it. Bounds, provenance and default-matrix
exclusion are 0627's: 32 MiB, path/size/SHA-256 plus CFB stream count, sector
size and target stream in the corpus identity, and absent from `Case::DEFAULT`.

`facade-open` stops its clock before reading the projection that names the
sample, so its timed region is the documented open alone. The other two include
the fresh open, as 0627's scenarios do, because the facade owns no cache across
calls and a caller pays for *reaching* the text.

A facade refusal is frozen as the scenario's outcome rather than dropped. 0609's
census found four `.doc` fixtures the facade refuses and the source-backed
snapshot admits; one of them, `duplicate-style-names.doc`, is measured here.

**The facade exposes no per-shape accessor on either legacy route.** The PPT
scenario is therefore one *slide's* text, not one shape's, and is named
`ppt_facade_file_one_slide_text` so that it is not read as change 0627's
`ppt_range_source_open_one_shape_text`.

### The twenty-four ordinary-save selectors

Three formats × two corpus origins × four phases:

```text
docx_ordinary_save_lifecycle          docx_real_file_ordinary_save_lifecycle
docx_ordinary_save_edit               docx_real_file_ordinary_save_edit
docx_ordinary_save_atomic_publish     docx_real_file_ordinary_save_atomic_publish
docx_ordinary_save_counting_publish   docx_real_file_ordinary_save_counting_publish
```

and the same eight names with `xlsx_` and `pptx_`. The route is the documented
one, the one the format crates' own examples use:

| format | open | edit | save |
| --- | --- | --- | --- |
| DOCX | `Package::open(path)` | `document_mut().add_paragraph_with_text(..)` | `Package::save(path)` |
| XLSX | `Workbook::open(path)` | `edit().sheet(..).set(..)`, `commit()` | `Workbook::save(path)` |
| PPTX | `Package::open(path)` | `opened_presentation_transaction().set_shape_text(..)`, `apply_opened_presentation_commit(..)` | `Package::save(path)` |

The generated origin reuses three existing harness corpora —
`build_semantic_docx_corpus(Medium)`, `build_xlsx_cell_crud_corpus(Medium)` and
`build_semantic_pptx_corpus(Medium)` — so this change introduces no generated
corpus of its own. The real-file origin takes `--ooxml-file PATH`, repeatable,
bounded to 32 MiB and classified by the package's own main part
(`word/document.xml`, `xl/workbook.xml`, `ppt/presentation.xml`), never by
extension. The baseline names change 0593's three fixtures, so its numbers and
this record's are about the same bytes.

`lifecycle` times open, edit and save together. `edit` times the semantic edit
alone. `atomic_publish` times the save alone, **reported separately from the
edit** exactly as change [0497](results/change-0497/README.md) reports its
atomic arm: the interval is one `litchi_opc::atomic::replace_with`, so it covers
the destination permission probe, sibling temporary creation in the
destination's own directory, the publication write, permission preservation,
`sync_all` on the temporary, the `rename` that replaces the destination, and the
parent-directory sync. Like 0497's arm, the private workspace is prepared before
the clock starts and the readback, the digest and the cleanup happen after it
stops. `counting_publish` times the documented sequential serialization into a
bounded counting sink.

The private workspace lives under the caller's `--filesystem-root` when one is
supplied, so the destination device is the caller's choice rather than an
ambient default, and it is removed when the corpus drops.

## Why it is sound

**Nothing under `crates/` changed**, so no invariant, error identity, limit,
defence or output byte could move. Every call is a documented public entry
point; the selectors add no `unsafe`, no global allocator, no `std::env` and no
`Command` (all four asserted by
`tools/test_perf_baseline_source_policy.py`), and no ambient I/O: the two
caller-named inputs are bounded and self-identifying, and the one directory the
save family writes into is the caller's `--filesystem-root` with this harness's
long-standing temporary fallback, reached through the single module that already
owns `std::env`.

**A typed refusal is an outcome, not a failure.** Both families freeze the
refusal verbatim and require every retained sample to reproduce it. This follows
change 0627, which froze the XLS full-text and PPT one-shape-text refusals for
the same reason: a caller pays for what the route read before refusing. A
refusing edit leaves the owner exactly as the open produced it, so the save
phases still measure the documented save of an unedited package — change 0593's
`noop` scenario.

**The default matrix and the checked catalog are untouched.** `Case::DEFAULT`
stays 41 cases; the selectable registry grows from 471 to 501 names, which the
enum-scanning assertion in `lib.rs` enforces and which the coverage index's
`minimum_selectable_cases: 439` already admits. `tools.test_corpus_manifest_v2`
and `tools.test_crud_coverage_index` both pass unchanged, which is the
mechanical statement that `catalog_sha256` and `content_set_sha256` did not
move.

**ADR reading.** ADR 0006 requires deterministic serialization absent a `Clock`,
actor identity or cryptographic RNG. Changes
[0625](0625-cfb-writer-deterministic-storage-order.md) and
[0631](0631-ooxml-relationship-order-verdict-sites.md) made that true of the
OLE2 writer's storage order and of three OOXML relationship-order verdict sites.
This record's save family asserts the consequence on the route a caller actually
uses: before any sample runs, each corpus proves that **two fresh
open/edit/save cycles publish the same digest** and that **saving one edited
owner twice publishes the same digest**, and every retained sample then has to
reproduce that digest. Six corpora × two proofs, all six passing, is the first
statement of 0625/0631's invariant on `Package::save`/`Workbook::save` rather
than on a probe.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
cargo 1.95.0. Every measured process pinned with `taskset -c 14`. Binary staged
outside the Cargo target directory (0627's lesson), SHA-256
`6abee148b50bb6beff8f1bbfe5ea61a28ee095fbbf0664fe7c5a14d1111df4bb`. 20 warm-ups
and 50 retained samples per case, three complete repeats (R1, R2, R3). Eight
agents were building and testing concurrently throughout. Destination directory
on the ext4 root filesystem, not on `/tmp`'s tmpfs.

### The A/A floor, and where it fails

The three repeats are the same binary over the same bytes, so the widest p50
disagreement between them is this window's A/A floor per row.

| population | p50 spread | worst row | rows |
| --- | ---: | ---: | ---: |
| facade | 1.55% | 47.01% | 15 |
| ordinary save | 3.02% | 60.36% | 24 |
| all | **2.09%** | 60.36% | 39 |

Eight rows exceed 5%, and this record reports them rather than burying them in
the median:

| row | spread |
| --- | ---: |
| `pptx_real_file_ordinary_save_atomic_publish` | 60.36% |
| `doc_facade_file_full_text` (FloatingPictures.doc) | 47.01% |
| `docx_real_file_ordinary_save_atomic_publish` | 23.59% |
| `pptx_real_file_ordinary_save_lifecycle` | 12.31% |
| `xlsx_ordinary_save_lifecycle` | 10.77% |
| `docx_real_file_ordinary_save_lifecycle` | 10.62% |
| `doc_facade_file_one_paragraph` (FloatingPictures.doc) | 6.93% |
| `pptx_real_file_ordinary_save_counting_publish` | 6.66% |

Five of the eight carry the atomic publication, whose two `fsync` calls and
`rename` are served by a device shared with seven other agents: the PPTX
real-file atomic median was 6.89, 11.04 and 9.33 ms on the three repeats, and
its p95 ranged from 11.52 to 27.79 ms. **Every atomic-publication median below
should be read as an order of magnitude, not as a number.** Two of the other
three are the two largest DOC facade scenarios on the same 335 KB fixture,
where R1 and R2 agreed within 0.7% (275.3 and 277.2 µs) and R3 ran 32% faster
(188.6 µs) — a host-state change, not sample noise. The last is the PPTX real
file's counting publication at 6.66%, the only sub-millisecond row in the
eight.

### The facade routes

R1 medians; the spread column is the three-repeat A/A floor for that row.

| selector | fixture | file bytes | p50 | p95 | spread | outcome |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `doc_facade_file_open` | `documentProperties.doc` | 9,728 | 15.09 µs | 15.83 µs | 1.39% | 1 paragraph |
| `doc_facade_file_full_text` | same | 9,728 | 15.74 µs | 16.63 µs | 1.97% | 22 chars |
| `doc_facade_file_one_paragraph` | same | 9,728 | 16.94 µs | 18.05 µs | 0.62% | 21 chars |
| `doc_facade_file_open` | `FloatingPictures.doc` | 335,360 | 178.63 µs | 186.89 µs | 1.29% | 210 paragraphs |
| `doc_facade_file_full_text` | same | 335,360 | 275.32 µs | 285.58 µs | 47.01% | 9,017 chars |
| `doc_facade_file_one_paragraph` | same | 335,360 | **538.31 µs** | 561.56 µs | 6.93% | 55 chars |
| `doc_facade_file_open` | `duplicate-style-names.doc` | 64,512 | 14.37 µs | 15.60 µs | 2.09% | typed refusal |
| `doc_facade_file_full_text` | same | 64,512 | 14.30 µs | 14.76 µs | 1.63% | typed refusal |
| `doc_facade_file_one_paragraph` | same | 64,512 | 13.99 µs | 15.23 µs | 1.12% | typed refusal |
| `ppt_facade_file_open` | `ppt_with_png.ppt` | 39,424 | 25.55 µs | 26.68 µs | 1.15% | 1 slide |
| `ppt_facade_file_full_text` | same | 39,424 | 29.62 µs | 33.52 µs | 1.99% | 20 chars |
| `ppt_facade_file_one_slide_text` | same | 39,424 | 32.04 µs | 32.71 µs | 0.06% | 20 chars |
| `ppt_facade_file_open` | `SampleShow.ppt` | 125,440 | 25.95 µs | 26.56 µs | 0.62% | 2 slides |
| `ppt_facade_file_full_text` | same | 125,440 | 31.93 µs | 33.52 µs | 1.55% | 236 chars |
| `ppt_facade_file_one_slide_text` | same | 125,440 | 32.38 µs | 33.41 µs | 2.06% | 92 chars |

Three things the queue can use.

**One paragraph costs 1.96× the whole text.** On `FloatingPictures.doc`,
`paragraph_text(55 chars)` is 538.31 µs against `text()`'s 275.32 µs — the
selected-paragraph route is nearly twice the cost of the complete extraction it
is meant to avoid. Both start from the same 178.63 µs open, so the marginal
query is 359.69 µs against 96.70 µs, a 3.7× ratio for reading 0.6% of the text.
The same pair on the 9,728-byte fixture is 1.85 µs against 0.65 µs. Nothing here
says why; it says the question is worth an instruction profile.

**The PPT facade is flat in file size and the DOC facade is not.**
`ppt_facade_file_open` costs 25.55 µs on a 39 KB deck and 25.95 µs on a 125 KB
one (+1.6% for 3.2× the bytes), while `doc_facade_file_open` costs 15.09 µs on
9.7 KB and 178.63 µs on 335 KB (11.8× for 34× the bytes).

**A refusal costs a full open.** `duplicate-style-names.doc` is refused at
14.37 µs, 4.8% *below* a successful open of a 9,728-byte document, on a file
6.6× that document's size. The refusal is `Corrupted file: invalid stylesheet: style
names and aliases must be unique` — the exact message 0609's census recorded for
the four fixtures route E refuses and route S admits, now reproducible from a
registered selector.

### The ordinary documented save

R1 medians. `edit+atomic` is the share of the lifecycle median that the edit and
the atomic publication medians account for.

| format / origin | lifecycle | edit | atomic publish | counting publish | edit+atomic |
| --- | ---: | ---: | ---: | ---: | ---: |
| DOCX / generated | 5.74 ms | 0.38 ms | 5.16 ms | 0.32 ms | 96.4% |
| DOCX / real file | 6.60 ms | 0.51 ms | 5.53 ms | 0.09 ms | 91.6% |
| XLSX / generated | 17.37 ms | 4.21 ms | 11.58 ms | 1.50 ms | 90.9% |
| XLSX / real file | 12.96 ms | 3.74 ms | 7.65 ms | 0.30 ms | 87.8% |
| PPTX / generated | 7.66 ms | 1.95 ms | 5.25 ms | 0.09 ms | 93.9% |
| PPTX / real file | **139.62 ms** | **133.61 ms** | 6.89 ms | 0.27 ms | 100.6% |

The phase decomposition is the point of the family, and it says two things.

**The documented save is publication-bound, not serialization-bound.** On five
of the six corpora the atomic publication is 59% to 90% of the lifecycle while
the identical serialization into a counting sink is 0.2% to 8.7%; on the sixth,
PPTX real file, the publication is only 4.9% because the edit is 133.61 ms. The
ratio of atomic publication to counting publication is 16.0× (DOCX generated),
59.1× (DOCX real), 7.7× (XLSX generated), 25.6× (XLSX real), 61.7× (PPTX
generated) and 25.2× (PPTX real) — the serialization is never the cost. Change
0593 observed this once — its atomic `PackageWriter::write` leg moved −2.33%,
"inside floor; route is fsync-bound" — and could not act on it; this is the same
fact on registered selectors, over six corpora and three formats. It also says
what a compression or layout optimization is worth on this route: at most the
counting-publish share, which is 0.2% to 8.7% of a lifecycle.

**One shape-text edit on a real 108 KB deck costs 133.61 ms.** That is 95.7% of
the PPTX real-file lifecycle and 490× the same deck's complete serialization. It
is 68.7× the same edit on the generated 61-member corpus (1.95 ms), for a deck
with 103 members instead of 61. The A/A spread on that row is 3.02%, so the
figure is solid even though its lifecycle sibling's is 12.31%. Nothing in this
record explains it; `opened_presentation_transaction()` captures a snapshot of
the whole opened presentation, and this is the first measurement that prices
that capture on a real deck.

### The byte split

Derived from the published archive and the source archive, one sample per
corpus. `deflate`/`stored` is a compression-method split of the published
members' stored payload; `identical`/`regen` is a provenance split against the
same-named source member's compressed bytes; `regen in` is the uncompressed
payload of the regenerated Deflate members — the part the save certainly fed to
the compressor.

| format / origin | source | output | deflate | stored | identical | regen | regen in | framing | members identical |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| DOCX / generated | 9,051 | 9,070 | 7,510 | 0 | 6,524 | 986 | 21,601 | 1,560 | 11/12 |
| DOCX / real file | 77,621 | 77,370 | 54,993 | 15,161 | 69,883 | 271 | 602 | 7,216 | 36/37 |
| XLSX / generated | 4,226,429 | 4,226,568 | 4,224,110 | 0 | 4,216,929 | 7,181 | 63,894 | 2,458 | 15/17 |
| XLSX / real file | 654,688 | 652,357 | 109,438 | 523,261 | 629,871 | 2,828 | 19,414 | 19,658 | 127/131 |
| PPTX / generated | 40,788 | 40,802 | 31,460 | 0 | 30,926 | 534 | 3,469 | 9,342 | 60/61 |
| PPTX / real file | 108,164 | 108,148 | 91,986 | 0 | 90,279 | 1,707 | 16,936 | 16,162 | 102/103 |

On every corpus a one-element edit regenerates one to two members and leaves the
rest byte-identical: **0.45% of the XLSX real file's payload bytes, 1.9% of the
PPTX real file's, 13.1% of the generated DOCX's**. Change 0593's pristine-member
preservation is doing what it was landed to do, and this is the first
byte-denominated statement of it. The counterpart is the framing: 19,658 bytes
of local headers, central directory and end record on the 131-member XLSX real
file, 3.0% of its output, which no compression change can touch.

Two rows are worth reading twice. The DOCX real file's edit was **refused** (see
below), so its 271 regenerated bytes over 602 uncompressed bytes are a near-pure
passthrough — the closest thing in this table to 0593's `noop`. And the XLSX
real file stores 523,261 of its 632,699 payload bytes, so the compressor sees
only 17% of that package at all.

### An admission fact: the documented DOCX editor refuses a fixture its reader opens

`litchi_docx::Package::open` accepts
`test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx`
— change 0593's own DOCX fixture — and `Package::save` republishes it. But
`Package::document_mut()`, the documented way to edit an opened DOCX, refuses
it:

```text
refused:invalid DOCX XML: syntax error: tag not closed: `>` not found before end of input
```

`document_mut()` re-parses `/word/document.xml` into a `MutableDocument` with
`MutableDocument::from_xml`, and that parse fails on markup the reader admits.
The refusal is stable across all 600 retained samples of all four DOCX real-file
phases in all three repeats. This record does not fix it — no file under
`crates/` may change here — and does not characterize its extent: one fixture is
one fixture. It is reported because change 0593's Limitations section named
exactly this hole (*"no DOCX or PPTX semantic-editor save through Path A exists
— no example opens a real `.docx`, edits through the model and saves"*), and the
first selector that tries finds the editor declining on the first real file
asked of it.

## Correctness evidence

**Nine new unit tests**, all passing.
`facade_ole2::tests`: `routes_and_scenarios_describe_themselves`;
`every_facade_case_is_opt_in_and_round_trips` (all six round-trip through
`Case::name`/`parse_case`, all six answer `is_facade_ole2`, none is in
`Case::DEFAULT`); `doc_corpus_derives_its_target_from_the_file`;
`doc_facade_phases_reproduce_their_frozen_oracles`;
`ppt_corpus_selects_a_slide_and_freezes_its_outcome`.
`ordinary_save::tests`: `formats_origins_and_phases_describe_themselves`;
`every_ordinary_save_case_is_opt_in_and_round_trips` (all twenty-four, by
enumerating the format × origin × phase product rather than restating the
names); `generated_docx_save_is_deterministic_and_reports_its_byte_split` (both
determinism invariants, the byte split's two internal identities, and all four
phases); `generated_xlsx_and_pptx_saves_are_deterministic`.

**Five new source-policy tests** in `tools/test_perf_baseline_source_policy.py`,
following the three change 0627 added: neither module declares `unsafe`,
`#[global_allocator]`, `std::env` or `Command`, and the save family reaches the
filesystem only through `filesystem::scratch_root`; the facade family adds no
caller-supplied input of its own and performs no whole-file read (it reuses
0627's bounded reader); `--ooxml-file` is bounded, self-identifying, and reads a
caller-named path exactly once; the atomic steps, the byte split and the
determinism invariants are structural fields rather than prose; and none of the
thirty variants appears in the `DEFAULT` array.

**Per-case gates, on every retained sample of every repeat.** Facade: the
sample's projection must equal the oracle derived from the file before the run.
Ordinary save: the published artifact must equal the corpus's reference
publication, and the edit outcome must equal the frozen one. Per corpus, before
any sample: two fresh open/edit/save cycles must publish the same digest, and
saving one edited owner twice must publish the same digest. The summariser
re-asserts all five over the retained reports and exits non-zero on any of them;
it exited 0 on all 39 case identities of all three repeats, so none fired. That
is 39 × 50 × 3 = 5,850 retained samples with no oracle, publication or edit-
outcome deviation, and 6 × 2 determinism proofs.

**Gates**, all in the worktree; only `tools/perf-baseline` has a Cargo package,
so only it is gated.

| Gate | Result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy --release --locked --all-targets` | clean, 0 warnings (workspace lints are deny) |
| `cargo test --release --locked` | 522 passed, 0 failed, 1 ignored in the library; 12 + 2 + 4 passed across the binary and integration targets; 0 failed anywhere |
| `cargo doc --release --locked --no-deps` | clean (rustdoc lints are deny) |
| `python3 -m unittest` on all 27 `tools/test_*.py` modules | 24 OK, 3 pre-existing failures |
| `tools/check_perf_claims.py --mode strict` | OK, 10 claims validated |
| `tools/check_report_claim_classification.py` | OK, 167 REPORT rows across 2 tables |
| `tools/validate_crud_coverage_index.py` | OK, 15 categories, 33 mapped selectors |
| `tools/non_iwork_gate.py verify` | OK, 45 bulk tree roots, 35 facade safe trees, 1 combined tree |

The three failing Python modules were reproduced on the untouched base checkout
`/home/zhuhe/code/litchi-worktrees/before-c7326f680`, so they are pre-existing
and unrelated: `test_check_crate_boundaries` (an iWork crate-boundary edge
count), `test_native_odf_resave` (a LibreOffice filter registry not installed on
this host) and `test_perf_claims` (`test_seed_registry_is_structurally_valid`).
They are the same three change 0627 recorded. `test_corpus_manifest_v2` and
`test_crud_coverage_index` both pass, which is the mechanical statement that the
checked default catalog SHA-256 did not move.

No crate under `crates/` changed, so no consumer suite is reachable from this
change. The feature-bearing `cargo test -p litchi --features docx,xlsx,pptx,xls`
gate change 0629 identified is not reachable either, but this change does newly
enable `litchi/doc` and `litchi/ppt` in the harness, and the harness's own
`cargo build`, `clippy --all-targets`, `test` and `doc` all compile that feature
set.

The gate tails are in [`results/change-0638/gates.txt`](results/change-0638/gates.txt).

## Validation preserved

No validation path was changed, relaxed or bypassed. Every scenario calls a
public entry point, and every validation those entry points perform —
`Document::open`'s format detection and read limits, `Presentation::open`'s,
`Package::open`'s OPC validation, `Workbook::from_package`'s complete
workbook/relationship validation, `opened_presentation`'s topology capture, and
the destination symlink refusal and permission preservation inside
`litchi_opc::atomic::replace_with` — is inside the measured region of the phase
that invokes it. The two typed refusals the baseline met are reported
verbatim, not suppressed. No limit was weakened: the `--ooxml-file` ceiling is a
new limit, not a relaxed one, and it matches the two that already exist.

## Limitations

**What is not claimed.**

* **No speedup, regression or production result.** Nothing under `crates/`
  changed, so there is nothing to speed up or slow down. Every number is a
  first descriptive baseline.
* **Five of thirty-nine rows' medians are device-bound and shared.** The atomic
  publication rows' A/A spread reaches 60.36% and their p95s exceed their p50s
  by up to 4×. The `edit+atomic ÷ lifecycle` shares and the atomic-to-counting
  ratios are robust to that (all three repeats agree on the ordering and on the
  order of magnitude), but the individual atomic medians are not numbers to
  optimize against.
* **No cold-cache, physical-device, page-cache or `fsync`-attribution result.**
  The atomic interval is one opaque `replace_with`; this record does not
  separate its sibling creation from its two syncs and its rename, because
  doing so needs either production instrumentation or `strace`, and 0497's
  precedent is the single interval. The destination lived on this host's ext4
  root filesystem with seven other agents writing to it.
* **One edit per format, one shape per corpus, one address per workbook.** The
  edit is a fixed marker: one appended paragraph, one cell set to a string, one
  shape's text replaced. Nothing is claimed about a batch edit, a
  length-changing structural edit, or an edit that adds or removes a part.
* **Six corpora and five real fixtures.** Three generated, three
  caller-named OOXML files (change 0593's), three `.doc` and two `.ppt`
  fixtures. The DOCX editor refusal is one fixture; the "one paragraph costs
  1.96× the whole text" observation is one fixture. Neither is a census, and
  0609's 57-fixture census is not reproduced here.
* **The byte split is derived, not instrumented.** `litchi-opc`'s
  `OpcOperationAccounting` deliberately excludes `PartWriter` and the topology
  publishers, which is the path a documented save takes, so
  `payload_bytes_identical_to_source` is an **upper bound on what a
  copy-through publisher could have avoided re-deflating**, measured by
  comparing compressed member payloads, and not an observation that this writer
  took a copy path.
* **`counting_publish` is not `save` minus the filesystem.** For PPTX it times
  `Package::to_bytes`, which has no sequential-sink alternative, so that row
  additionally pays a whole-archive allocation the DOCX and XLSX rows do not.
* **No allocation, RSS, instruction or cycle counts.** None was taken. The
  phase attribution is wall-clock only.
* **`ppt_facade_file_one_slide_text` is one slide, not one shape.** The facade
  exposes no per-shape accessor on the legacy PPT route.

**Deliberately not done.** The stale sentence at
`tools/perf-baseline/README.md:10` ("The registry has 441 selectors") predates
change 0601 and is left alone, for the reason change 0627 left it alone: the
authoritative count is the test-enforced `selectable_count` assertion, now 501.
The DOCX `document_mut()` refusal is reported and not fixed; it is a
`crates/litchi-docx` question and this is a harness record.

## Retained evidence

[`results/change-0638/README.md`](results/change-0638/README.md) — the twelve
raw schema-1 reports (three repeats × four invocations), the two scripts, the
summariser's output, the gate tails, the decision record and the log paragraphs.
