# 0299 — Pages to parity in the unified layer (M6)

- Status: implemented on `uo2/iwork-pages`, cut from `feat/unified-ops-api`
  at `3dccbac8e8`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 24 (Pages) and step 28 (partial, §7). Builds on
  [0296](0296-iwork-joins-file.md).
- Date: 2026-10-15
- Owner decisions applied: **D1** and **D2** (2026-10-15): the iWork crates
  may be edited; `File::Pages` holds the leaf `Package`.
- Owned: `crates/litchi-pages` (ADR 0034 W1 and the new leaf APIs), the
  Pages file of the ops iWork adapter
  (`crates/litchi-ops/src/format/iwork/pages.rs`), small separated hunks in
  the shared ops files, the Pages consumers in `litchi`, `litchi-py` and the
  legacy migration host `litchi-iwa`, one build-script marker in
  `litchi-iwa-protos`, and the Pages audits of
  `tools/check_crate_boundaries.py`.
- `performance_claim`: paired leaf/unified medians in §8 (release build,
  one pinned core of a shared host, ABBA). No frozen-harness qualification.

Record 0296 made `File::Pages` read paragraphs and sections and write
paragraph and section text; creation, structure, tables, headers, footers,
footnotes, metadata and package directories were typed gaps, 16 of the 32
Pages corpus files refused the set-then-revert probe ("the Pages text edit
intersects dependent content"), and `litchi-pages` still carried the
pre-ADR-0034 vocabulary. This record finishes ADR 0034 wave W1 for
`litchi-pages`, adds the leaf APIs the Pages requests asked for, and wires
them into the unified layer.

## 1. ADR 0034 W1 for `litchi-pages`

Commit `0636b1a95a` (one commit: the rename spans the leaf and every
consumer); the linear paragraph staging of §8 is `379abad4db`; the unified
layer, tests, goldens and this record are the last commit.

- **Focused transaction families.** Every family is a module with the six
  canonical names, the nominal types themselves renamed (the old prefixed
  names are not aliased: the glossary checker flags a nominal name behind
  an alias):

  | Old crate-root names | Module |
  |---|---|
  | `SectionText*` | `section::text` |
  | `SectionName*` | `section::name` |
  | `SectionPagination*` | `section::pagination::transaction` |
  | `PageLayout*` | `page_layout::transaction` |
  | `HeaderFooterText*` | `header_footer::text` |
  | `FootnoteText*` | `footnote::text` |
  | `BodyFootnote*` (+ `Snapshot`) | `footnote::body::transaction` |
  | `BodyChartArrangement*`, `BodyChartData*`, `BodyChartMetadata{Error,LimitKind}` | `chart::{arrangement, data, metadata}` (new `chart`) |
  | `BodyDrawableOrder*` | `drawable::order` (new `drawable`) |
  | `BodyImageAdjustments*` | `image::adjustments` |
  | `BodyTableAppearance*`, `…HiddenAxes*`, `…Name*`, `…Sort*` (+ `Path`) | `table::{appearance, hidden_axes, name, sort}::transaction` |
  | `BodyTableDimension*`, `BodyTableHeaderSettings*` (+ `InvalidReason`), `BodyTableMerges*` | `table::{dimension, headers, merge}::transaction` |
  | `BodyTableTitle*`, `BodyTableLock*` | `table::{title, lock}` |
  | `BodyTableDeletion*` | `table::deletion` |
  | `BodyTableCatalog` → `Catalog`, `BodyTableSnapshot` → `Snapshot`, `BodyTableCatalog{Error,LimitKind}` | `table::catalog` |
  | `BodyTableCells{Error,LimitKind}` | `table::cells` |
  | root aliases `BodyTableHeaderSettings`, `BodyTableLockState`, `BodyTableName`, `BodyTableNameValueError` | deleted (`table::{headers::Settings, lock::State, name::Name, name::Error}`) |

  `section::settings`, `section::background` and `document_settings`
  already had the shape. New families: `section::paragraphs` and
  `metadata` (§2). The owner modules are `pub(crate)` in
  `package.rs`; the crate root keeps only the ingress/egress types.
- **Limits and options (ADR 0034 §5).** `SourceLimits` (input bytes, members,
  entry bytes, aggregate bytes, component bytes; default 512 MiB input,
  every other dimension at the shared iWork archive ceiling) and the
  existing `SemanticLimits` (sections, text bytes) are built with consuming
  builders whose `max_<noun>` setters return
  `PackageError::InvalidLimit { kind, value, maximum }`;
  `LimitKind::{Source, Semantic}` names the dimension. `OpenOptions`
  carries both. Deleted: `DocumentSourceLimits`,
  `DocumentSourceLimitKind`, `DocumentSourceLimitsError`,
  `SemanticLimitsError`, `SemanticLimits::new`, `DocumentReadOptions`, the
  root `Limits` re-export of the archive profile, and
  `Package::{open_with_limits, from_bytes_with_limits}`. A hidden
  `OpenOptions::__archive_limits` keeps the iWork crates' own tests and
  the migration host able to tune archive-internal dimensions.
- **Owner and projection (ADR 0034 §2).** `Package::{open, open_with_options,
  from_bytes, from_bytes_with_options, from_shared_bytes,
  from_shared_bytes_with_options, new, edit_x/apply_x, write_to, to_bytes,
  save}`; `Package::document()` (owned, one reference-count increment) and
  `document_ref()` replace `semantic_document()`; `Package::snapshot()` is
  deleted (`Clone`). `Document` keeps no ingress verb: its
  `open*`/`from_bytes*`/`from_shared_bytes*` and `snapshot`, the
  archive-free `ReadError`/`ReadLimitKind`/`IoKind`, and the raw-`usize`
  `section(index)` are deleted; `section_at` takes a `Position`.
- **Errors (ADR 0034 §6).** `PackageError::NotPages` → `UnexpectedFormat { detected:
  FileFormat }`; public ingress classifies the package first, so a
  Keynote/Numbers package is `UnexpectedFormat` and a ZIP without an iWork
  root `NotOfficeFile`. New `Limit { kind, observed, maximum }` (directory
  capture) and `MissingIdentity`. The metadata sidecars are now bounded on
  every `Package` ingress (`PayloadLimit` above 1 MiB, the bound the
  archive-free reader enforced).
- **Selectors (ADR 0034 §8).** `ImageSelector`/`BodyChartSelector` variants
  `Index(Position)` → `Position(Position)`; the footnote selector
  `At(Position)`/`Index(usize)` → `Anchor(Position)`/`Position(Position)`.
- **Creation (ADR 0034 §4, §9; request 1).** `Package::new()` returns the
  deterministic minimal document; `builder::{Builder, Identity}` builds one
  with an explicit identity (`Identity::generate()` or
  `Identity::from_bytes`) and initial body text; `Builder::build()` without
  an identity is `PackageError::MissingIdentity`; two builds with one fixed
  identity are byte-identical (`builder` unit tests). The built-in template
  (`crates/litchi-pages/assets/blank.pages`, 5 384 bytes, sha256
  `4083042088532964e7b986a584d32e4001e2526a45369b33bfeb1505cd2b1949`) was
  authored once by the legacy `litchi_iwa::pages::PagesDocumentBuilder` at
  `3dccbac8e8` (`PagesDocumentBuilder::new().build()?.to_bytes()`, a
  throwaway example); its document identity is that run's. A builder
  rewrites `Metadata/DocumentIdentifier` and the UUID keys of
  `Metadata/Properties.plist` (`documentUUID`, `stableDocumentUUID`,
  `shareUUID`, `privateUUID`); the version UUID and the object UUIDs of
  `Index/Metadata.iwa` stay the template's (as for a copied document).
- **Glossary.** New `pages` profile in `tools/api_glossary.json`: 164 banned
  names (every old prefixed name, the deleted aliases, the old limit/option
  types, `NotPages`, the removed `Package`/`Document` verbs) and 24 family
  modules; rows flipped in `docs/API_GLOSSARY.md`.
- **Consumers migrated in the same commit:** `litchi-iwa` (the legacy
  migration host: imports, `OpenOptions` for its focused table packages,
  footnote selectors), the facade (`litchi::pages::semantic` now re-exports
  the projection values, the limit profiles, `OpenOptions` and
  `PackageError`, and its `open`/`from_bytes` read through `Package`), the
  facade tests and fuzz targets, `litchi-py` (`iwork::open_pages` reads
  through `Package`, classifying `PackageError`), the ops adapter, and the
  `litchi-iwa-protos` build-script provenance marker for
  `footnote_text.rs` (rustfmt joined the line the marker quoted). No
  deprecated legacy migration-host field is Pages-specific in
  `litchi-pages`; the deprecated raw-ID drawable-comment methods of the
  `litchi-iwa` Pages editor stay (its tests still use them, and no focused
  owner replaces them yet).

## 2. Leaf APIs (`litchi-pages`)

- **Paragraphs (request 12)** — `section::paragraphs::{Edit, Patch, Commit,
  Diagnostics, Error, LimitKind}` and `Package::{section_paragraphs,
  edit_section_paragraphs, apply_section_paragraphs}`: `set`, `insert`,
  `remove`, `relocate`, `swap` by `Position`. The commit diffs the staged
  section text against the source (common prefix and suffix) and publishes
  **one** section-text splice when it consumes and inserts no object
  anchor, however many operations were staged; otherwise one minimal
  splice per operation, each over the package the previous one
  published. Setting an anchored paragraph's text keeps its anchors
  (leading or trailing; interleaved anchors are `InterleavedAnchors`);
  removing, relocating or swapping a paragraph that holds an anchor is
  `AnchoredParagraph`; a section's only paragraph is `LastParagraph`.
- **Section-text dependency boundary.** The ownership census of the rooted
  body storage now accepts one aggregate back-edge from an object the body
  anchors (an attachment, or the drawable it holds — a table's
  `DrawableArchive.parent`): it holds no text position, and the rewrite
  preserves it. This is why the 3 source-built table files refused every
  splice; the other 13 of 0296's 16 refusals were the unified layer
  replacing a whole line including its anchor, which the paragraph family
  now avoids.
- **Metadata (request 14)** — `metadata::{Edit, Patch, Commit, Diagnostics,
  Error, LimitKind}` and `Package::{edit_metadata, apply_metadata}` write
  title, author, keywords and description into `Metadata/Properties.plist`
  (`Title`, `Author`, `Keywords`, `Comments`; clearing also drops the
  legacy fallback keys the reader honours), rewriting only that member
  through the archive's bounded reassembly, then reopen and compare the
  read-back metadata. The Pages app does not display these keys; no native
  acceptance is claimed.
- **Package directories (request 16)** — `Package::open` captures an
  app-authored directory once (no symbolic links or special files, depth
  ≤ 16, member/entry/input bounds of `SourceLimits`, each file's length
  re-checked after reading) and converts it to the single-file form: the
  `Index.zip` members, then every other file as a stored member in byte
  order of its name. The result is an ordinary exact-source package; it
  saves as a single file. A loose `Index/` directory is refused.
- `table::read::Value` is re-exported (the unified cell types).

Leaf tests: `tests/section_paragraphs.rs` (every Pages corpus section
accepts a paragraph set plus insert and its inverse restores the input
bytes, including the 16 files 0296 refused), `tests/package_creation.rs`
(creation, identity, metadata round trip and inverse, directories), the
rewritten `tests/document_reader.rs` (Package ingress and projection,
directory capture limits inclusive at max and max − 1, the properties
bound), and the unit tests of `limits`, `builder` and
`section_paragraphs`.

## 3. Unified coverage (`File::Pages`)

| Element | Paths | Read | Write |
|---|---|---|---|
| `document` | `/` | `sectionCount`, `paragraphCount`, `tableCount` | `create` (`Package::new`) |
| `body` | `/body` | `paragraphCount` | — |
| `p` | `/body/p[N]` | `text` (without anchor characters), `section`, `objects` (anchored objects) | `set text`, `add` (Start/End/Position/Before/After, or into a section), `remove`, `move` (within or across sections), `swap` |
| `section` | `/section[N]` | `name`, `text`, `paragraphCount` | `set name`, `set text` |
| `tbl` / `tr` / `tc` | `/body/tbl[N]/tr[R]/tc[C]` | `name`, `rows`, `columns`; `cellCount`; `text`, `type` | `remove tbl` (leaf `table::deletion`) |
| `header` / `footer` | `/section[N]/header[M]` | `text`, `template` (first/even/odd), `slot` | `set text` |
| `footnotes` / `footnote` | `/footnotes/footnote[N]` | `count`; `text`, `mark` | `add` (anchored at the end of the parent paragraph, or of the last one), `set text`, `remove` |
| `/metadata` | set-only | `File::metadata()` | `title`, `author` (`creator`), `keywords`, `description`; `null` clears |

- **Batches.** Consecutive paragraph `text` sets in one `File::edit` /
  `batch` coalesce into one leaf paragraph transaction per section (one
  splice when anchor-free); a non-paragraph operation flushes the run first,
  so the order of mixed batches holds (`pages_unified.rs`
  `a_batch_mixing_structure_and_sets_keeps_its_order`). One `set` over a
  query that matches many paragraphs coalesces the same way.
- **Package directories** open through `litchi::open`/`litchi_ops::open`
  (`File::Pages`, no origin: `save_as` writes the single-file package);
  Keynote and Numbers directories are still refused.
- **Typed gaps** (`capabilities()`, delta table): `r` (runs and formatting,
  request 13), section `add`/`remove` (request 11), `tc` `set` (no Pages
  cell-write transaction), `tbl` `add`, removing or moving a paragraph that
  anchors an object, `/metadata` reads (set-only path), and `create` with a
  fresh identity (a leaf `Builder` input).
- Errors stay typed: seven new `LeafError` variants (`PagesParagraphs`,
  `PagesName`, `PagesHeaderFooter`, `PagesFootnote`, `PagesFootnotes`,
  `PagesMetadata`, `PagesTableDeletion`) with their code mappings.

## 4. Tests, goldens and checklist

- `crates/litchi-ops/tests/pages_unified.rs` (12 tests): paragraph add
  (position, end, into a section), remove, move (within and across
  sections), swap; an anchored paragraph's set keeps its table and its
  removal is refused; tables, rows and cells read, cell set refused, table
  removal; section names; headers and footers (9 + 9 per section on the
  corpus, footer set round trip); footnotes added at a paragraph end, set
  and removed; `/metadata` writes, clears and refusals; `create(Pages)`
  (deterministic) and the capability verbs; directory open and `save_as`;
  a 300-set `File::edit` publishing **one** leaf commit and reverting
  byte-exactly; a mixed structure/set batch keeping its order.
- `iwork_unified.rs`: the Pages gap test now asserts creation works and
  section `add` stays a typed gap. `litchi-cli` `tests/cli.rs`: `create
  blank.pages` succeeds and a set round-trips (Numbers takes over the
  "creation is a typed gap" checks). Facade `unified_open_facade.rs`: a
  Pages package directory opens as `File::Pages` with the single-file
  form's text. Python: `create("pages")` joins the creatable formats, a
  Pages directory opens through `litchi_py.open` (Keynote/Numbers
  directories still refuse); README, `.pyi` and the `open` docstring say so.
- **Query corpus** (`tests/golden/query_corpus/digests.json`): the readable
  keys were recorded before the write path was exercised: 7 new Pages
  selectors (`p[objects>0]` 34, `tbl` 22, `tbl > tr` 100, `tc[type=text]`
  48, `header` 288, `footer[template=odd]` 96, `footnote` 0 matches) and 5
  re-blessed Pages digests (`p`, `body > p`, `p:contains(e)`,
  `p:not(:empty)`, `p[last()]`): a paragraph's `text` no longer carries the
  object-replacement character and `objects` is a new property; 20
  anchor-only paragraphs are now empty. Every other format's digest is
  unchanged.
- **Phase G checklist** (`docs/unified-ops/results/0289/crud-checklist.*`,
  re-blessed after review; `body-table-read-native.pages` added to the
  Pages fixtures): Pages 11 elements, 85 claimed cells, **85 passed**, 0
  failed, 0 unreached, 36 not applicable; `move`/`swap` ok on 2 element
  rows, refused (typed) on 14. No `KNOWN_FAILURES` entry for Pages; no
  recipe was needed (the generated `create() + add paragraph + add
  footnote` seeds reach every footnote cell).
- **Capability and registry goldens** (`pages.json`), identical in the
  all-features set and the all-but-`encryption` set; the OfficeCLI delta
  table regenerated (Pages section only).

## 5. Corpus fidelity

`docs/unified-ops/results/0289/fidelity.{md,tsv}` (re-blessed): all 32
Pages corpus files **open → empty edit → save byte-identically** (as in
0296). The set-then-revert probe now reaches `section.name` first (it
became settable) — 29 `byte-exact`, 3 `value-exact` (a rewritten
`Index/Document.iwa` is re-encoded, so the reverted ZIP differs from offset
14, the entry CRC) — instead of 0296's 16 `byte-exact` and 16 `refused`.
The paragraph-text evidence for those 16 files is the leaf corpus test
(§2): every section of every Pages corpus file takes a paragraph set plus
an insert and the inverse restores the exact input bytes. New structural
round trip (record 0293's harness, extended): a paragraph added at the
start of the body and removed again reads back with the same first
paragraph on all 32 files — 29 `byte-exact`, 3 `value-exact`.

## 6. `litchi::iwork::Document` (plan step 28)

Kept, for Keynote and Numbers. `File::Pages` now covers every read the
facade reader gave Pages: package directories (§2), body tables with their
names and cells, headers, footers and footnotes (§3). The remaining reasons
are Keynote's and Numbers' (their directories and table/text roles,
records 0297 and 0298). No in-tree code outside the facade's own tests and
fuzz targets calls `litchi::iwork::Document`; the Python family classes
read Pages through `litchi::pages::semantic`, which now opens a `Package`.
Delete the facade reader when 0297/0298 close their gaps.

## 7. Verification

All with `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`; exit status 0
unless stated.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-ops -p litchi -p litchi-py -p litchi-cli -p
  litchi-pages --all-targets --all-features -- -D warnings`; `cargo clippy
  -p litchi-iwa --lib -- -D warnings` (its test targets keep the
  pre-existing Keynote `deprecated` lints, and its examples do not compile
  at the base either: `litchi_iwa::charts::ChartData::new` returns a
  common-crate error, unrelated to Pages); `cargo clippy -p litchi-ops
  --no-default-features [--features F] --all-targets -- -D warnings` for no
  features, `pages`, `keynote`, `numbers`, `pages,keynote`,
  `pages,numbers`, `keynote,numbers`, `docx`.
- Tests: `cargo test -p litchi-pages --all-features` (682 passed),
  `cargo test -p litchi-ops --all-features` (2 240 passed), `cargo test
  -p litchi --all-features` (260), `cargo test -p litchi-cli
  --all-features` (32), `crates/pyo3-litchi/tests/run_tests.sh` (270).
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p
  litchi-pages -p litchi-ops -p litchi -p litchi-iwa`.
- `python3 tools/check_crate_boundaries.py` (67 packages, 282 internal
  dependency declarations, 11 debt items); `python3 -m unittest
  tools.test_check_crate_boundaries` 1139 tests, 1 failure that predates
  this record (`test_iwa_package_edge_is_archive_owned_and_host_edge_cannot_return`:
  281 policy edges against an expected 270 in the unmodified
  `tools/crate_boundaries.json`); `python3 tools/non_iwork_gate.py verify`
  (46 bulk roots, 42 safe facade trees, combined tree).
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile pages`
  (164 bans, 24 families) and `--profile ops-iwork`; `python3 -m unittest
  tools.test_check_api_glossary`.
- Goldens: `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities
  --test registry` with `--all-features` and with every feature but
  `encryption` (identical); `crud_matrix` and `corpus_fidelity` blessed
  with `LITCHI_BLESS=1`, `query_corpus_differential` with
  `LITCHI_BLESS_GOLDEN=1`, then re-run green; `python3
  tools/officecli_capability_delta.py --officecli …/OfficeCLI`.
- `cd tools/perf-baseline && cargo check --locked --all-targets` (no
  lockfile change).
- The OfficeCLI corpus has no iWork input; unaffected.
- `boundaries`: the Pages audits of `tools/check_crate_boundaries.py`
  enforce the new layout: one `pub(crate) mod` owner per family, the short
  names defined there and re-exported once from the public module, and no
  old prefixed name, alias, glob or second path anywhere in the crate.

## 8. Performance

`crates/litchi-ops/examples/pages_bench.rs` (new): leaf vs unified over
the same bytes in one process, release build, pinned to core 29 of a shared
32-core host (load 5–7.5), ABBA × 8 rounds × 30 calls (7 for the 1k/10k
batches), three runs. Inputs are opened outside the timed region; the
unified legs include building the object tree on their first read. Per-run
median ratio unified/leaf (summary:
`docs/unified-ops/results/0299/pages-bench-summary.jsonl`; raw JSONL kept
outside the repository, 108 KB, sha256 `f668c54d…`, `32f6611b…`,
`b1536a2e…`; binary sha256 `59bdb941…`; regenerate with the command in the
example header):

| Scenario | leaf (median) | unified (median) | ratio runs 1 / 2 / 3 |
|---|---|---|---|
| open (`basic.pages`) | 913–919 µs | 912–919 µs | 0.996 / 0.996 / 1.001 |
| open + get | 922–978 µs | 936–989 µs | 1.014 / 1.017 / 1.014 |
| set | 2.02–2.04 ms | 2.04–2.06 ms | 1.008 / 1.011 / 1.009 |
| add (paragraph) | 2.02–2.05 ms | 2.05–2.09 ms | 1.022 / 1.017 / 1.018 |
| remove | 2.02–2.04 ms | 2.04–2.06 ms | 1.010 / 1.014 / 1.012 |
| move | 2.02–2.04 ms | 2.05 ms | 1.009 / 1.012 / 1.010 |
| batch 1 (10 000-paragraph document) | 2.13–2.14 ms | 3.53–3.56 ms | 1.663 / 1.653 / 1.666 |
| batch 10 | 2.13–2.14 ms | 3.54–3.55 ms | 1.663 / 1.655 / 1.666 |
| batch 100 | 2.14 ms | 3.58 ms | 1.675 / 1.671 / 1.678 |
| **batch 1 000** | 2.22–2.23 ms | 3.96–3.97 ms | **1.784 / 1.781 / 1.789** |
| batch 10 000 | 2.94–2.95 ms | 7.75–7.78 ms | 2.638 / 2.635 / 2.636 |

- **Inside the budgets:** open (≤ 5 % or 20 µs: −0.4 % to +0.1 %), open +
  get (+1.4 % to +1.7 %, 12–15 µs), single `set` (≤ 10 %: +0.8 % to
  +1.1 %), the structural single ops (+1.0 % to +2.2 %), and the 1k-op
  batch (≤ 2× a hand-written leaf edit doing the same work: 1.78×). The
  batch scales linearly: 0.48 µs per coalesced operation above a
  constant.
- **Missed:** the 10k-op batch is 2.64× (the 2× target is set at 1k; 10k
  only has to scale linearly, which it does). The constant ~1.4 ms of the
  unified legs on the 10 000-paragraph document is the object tree built on
  the first read (10 000 paragraph items, each with its own text copy, path
  and property values) and the snapshot-size count of the published
  package; the leaf's coalesced commit is one section-text splice, so the
  constant dominates small batches (1.65× at 1–100 sets on that
  document). Follow-up: share each item's text with its `text` property
  and build paragraph paths on demand.
- **Found and fixed while measuring:** (1) the leaf paragraph edit
  re-applied every staged operation on each `stage` (quadratic: 245 ms for
  1 000 sets, 2.9 s for 10 000); it now keeps the running paragraphs (2.2 ms
  and 2.9 ms); (2) the ops tree builder counted a parent's children of a
  kind on every push, and `Nav::child` counted siblings for a positional
  step (both quadratic in a large parent; now an ordinal table and a
  direct slot probe, which also benefits Keynote and Numbers trees);
  (3) reading the section headers and footers costs ~280 µs per package
  (the leaf discovers every template slot), so the Pages tree now adds them
  the first time a read reaches a section's children, a header or footer
  path, or a whole-tree view — open + get went from +31 % to +1.5 %.
- No callgrind run: after the three fixes every per-op cost is linear and
  the remaining constant is attributed above.

## 9. Follow-ups

- **Request 11** — section add/remove: a section break plus a new native
  section object (and its template references) in the leaf.
- **Request 13** — text runs and character formatting (`r`), Pages body
  text styles.
- **Pages table cells** — a cell-write transaction (the Numbers TST writer
  is not shared with Pages) and body-table creation.
- **Directory save-back** to the bundle form (today a directory opens as
  its single-file form and saves as one file); Keynote and Numbers
  directories (0297/0298).
- **`litchi::iwork::Document`** — delete when 0297/0298 close the Keynote
  and Numbers reasons (§6).
- **Unified open options** — pass `litchi_ops::OpenOptions` input limits
  through to `litchi_pages::SourceLimits` (the unified open still uses the
  leaf defaults through the prepared source).
- **Metadata** — native acceptance in the Pages app (the written keys are
  read back by this library only); `Keywords` typing of the shared
  `Metadata.keywords` (2026-10-13 request) is unchanged.
- **Builder identity** — a fresh identity also for the version UUID and the
  object UUIDs of `Index/Metadata.iwa` (needs a package-metadata rewrite).
- **Tree build cost** (§8) — share item text with the `text` property.
- **Tools** — the pre-existing `test_check_crate_boundaries` edge-count
  failure and the non-compiling `litchi-iwa` chart examples belong to their
  owners.
