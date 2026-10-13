# 0268 — ODF W3: package split, atomic save and limits

- Status: implemented on `uo2/s22-odf-w3`, cut from `feat/unified-ops-api`
  at `4d1afe92bc`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 22; ADR 0034 wave W3.
- Date: 2026-10-13
- Scope: `crates/litchi-odf-common`, `crates/litchi-odt`, `crates/litchi-ods`,
  `crates/litchi-odp`; publication and the `SourceLimits` rename in
  `litchi-odc` and `litchi-odf-formula`; one test in `litchi-oth`. Consumers
  migrated in the same change: `litchi-ops` (`error.rs`, `options.rs`,
  `format/od{t,s,p}*`, `detect/detected.rs`, `open/path.rs`,
  `view/presentation.rs`), the `litchi` facade examples and tests,
  `litchi-odf` examples, fuzz target, test and README,
  `tools/perf-baseline`, `tools/native-resave`, and the glossary gate. No
  iWork crate and none of the shared dispatch files
  (`read`, `write`, `transaction`, `save`, `create`, `capabilities`, `file`)
  changed. `litchi-py` built and tested unchanged.
- `performance_claim`: paired before/after measurements in §5 (release, LTO,
  one pinned CPU, ABBA ×3). No frozen-harness qualification.

## 1. Owner split (ADR 0034 §2–§3)

Each of ODT, ODS and ODP now has one immutable `Package` that owns every
state verb, and a read-only projection that owns none.

| Format | Owner | Projection | Primary family |
|---|---|---|---|
| ODT | `litchi_odt::Package` (was `transaction::Snapshot`) | `Document` (`Arc` inner) | `litchi_odt::transaction::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` |
| ODS | `litchi_ods::Package` (was `document::Snapshot`; the physical `package::Package` is now crate-private `Archive`) | `Workbook` (was `Spreadsheet`) | `litchi_ods::document::{…}` |
| ODP | `litchi_odp::Package` (was `edit::Snapshot`) | `Presentation` | `litchi_odp::edit::{…}` (`Transaction` renamed `Edit`) |

- **State verbs on `Package`:** `new`, `open`, `open_with_options`,
  `open_with_password(_and_options)`, `from_bytes(&[u8])`,
  `from_bytes_with_options`, `from_bytes_with_password(_and_options)`,
  `from_reader(_with_options)`, `from_prepared_package`, `edit`, `apply`,
  `save`, `save_with_durability`, `to_bytes`, `write_to`, `as_bytes`, and the
  projection accessor (`document()` / `workbook()` / `presentation()`), plus
  `tolerance_report`. `from_bytes` borrows and copies once after the input
  check.
- **Projections are read-only.** Every `open*`, `from_*`, `save`,
  `to_bytes`, `into_bytes`, `edit`, `snapshot`, `document_snapshot`,
  `apply_document_patch` and direct mutator on `Document`, `Workbook` and
  `Presentation` is deleted (ODP layout/master/handout/annotation setters
  and ODT annotation/form/embedded/RDF/script mutators became crate-private
  or moved into `Edit`). Projections are cheap clones of an `Arc` inner.
- **Commit shape.** `Commit::{package, patch, diagnostics, into_package,
  into_parts}`; `Patch::is_noop`; `Diagnostics::changed` (ODT adds
  `operation_count`, ODP `domains`). `Patch::apply` is crate-private;
  `Package::apply(&patch)` is the public route. ODS `Edit::rollback` is
  deleted (drop the `Edit`). ODT annotation operations and ODP handout
  operations are now `Edit` operations.
- **History/merge** types keep their verbs and now return `Package`
  (`History::current/undo/redo/record`, ODP `three_way*`, `join_snapshot`).
- **Laziness.** Opened packages parse their projection at open (as before).
  A freshly built ODT/ODS/ODP package indexes and validates its archive at
  `build()` and decodes the projection on first access; ODP decodes its
  bounded slide projection on first `slides()`/`edit()`.
- **No-op identity.** An empty commit returns the source byte-for-byte
  (`package_owner.rs` in each crate; existing exact no-op suites unchanged).
  Signature and encryption policies are unchanged: changed writes on signed
  or encrypted ODF packages are still refused at `edit()`/`commit()` with
  the existing errors.

## 2. Atomic publication

`litchi_odf_common::publish::save_bytes(path, bytes, Durability)` stages a
sibling temporary file and renames it over the destination through
`litchi_core::atomic` (record 0226). Every whole-package save of ODT, ODS
and ODP `Package`, the ODT generic and ODC owners, the ODT/ODS mutable
writers and the ODF formula package uses it; `save()` is
`Durability::Full`, and `save_with_durability` lets a caller choose. No
`std::fs::write` of a whole package remains in `litchi-od*` sources
(only test seeding).

CRUD evidence (`crates/litchi-od{t,s,p}/tests/package_owner.rs`):
an existing destination is replaced with exactly `as_bytes()` and no
staging file is left; a save through a symlink is refused and the link
target keeps its bytes; a save into a read-only parent directory fails and
the existing destination keeps its bytes (asserted when the process cannot
bypass directory permissions).

## 3. Limits (ADR 0034 §5)

- `litchi_odf_common::SourceLimits` replaces `SourcePackageLimits` (deleted;
  banned). It is a checked builder (`SourceLimits::builder().max_<dim>(n)?
  .build()?`) with `DEFAULT` and `MAXIMUM`; dimensions are
  `SourceLimitKind::{InputBytes, Members, ArchiveMemberNameBytes, ArchiveMetadataBytes, ArchiveCompressedBytes, EntryBytes, AggregateBytes, MimetypeBytes, ManifestBytes}`.
- `SemanticLimits` bounds `content.xml` bytes, `styles.xml` bytes and units
  (ODT body children, ODS sheets, ODP slides), checked against declared
  member sizes before inflation.
- `OpenOptions { source_limits, semantic_limits }` is the one envelope on
  every public ingress of the three owners and of
  `SourceBacked{Document,Spreadsheet,Presentation}` and the catalog/source
  facades (`from_read_at_with_options`, `open_with_options`, …). The
  `*_with_limits*` combinations and the ODP/ODT `ReadLimits` aliases are
  deleted.
- `litchi_odf_common::Error { InvalidLimit, Limit { kind: LimitKind, observed,
  maximum }, Core }` is the family `Error`; `LimitKind = Source(…) |
  Semantic(…)`. Core `ResourceLimit`s with an ODF source scope convert into
  the typed variant and back losslessly. `litchi-ops` classifies `Limit` as
  `limit_exceeded` and `InvalidLimit` as `invalid_argument`; every other ODF
  failure keeps its pre-W3 `Source::Core` payload.

## 4. Builders, determinism and the oth DTD test

- `Builder::build()` returns `Package` for ODT, ODS and ODP;
  `Builder::save` is deleted; `Builder::set_metadata` stays.
- Fresh creation is deterministic: `Package::new()` and identical builder
  input produce identical bytes, and generated `meta.xml` has no
  `meta:creation-date`/`dc:date` (asserted in each `package_owner.rs`). No
  clock or randomness is consulted on these paths.
- `litchi-oth semantic_api::raw_zip_negatives_are_rejected_by_the_open_boundary`
  asserted that a prologue `<!DOCTYPE>` is refused. That DOCTYPE is on the
  shared ingress path, which deliberately ignores it without processing it
  (leniency F9.11, as LibreOffice's SAX import does). The test now asserts
  the tolerated template opens with the same body, and that a DTD
  declaring an entity referenced from the body is still refused.
- `litchi-odf-common` family probing accepts an aliased `office` prefix
  (`contains_aliased_body` fallback, with a unit test); a builder
  authoring a non-default prefix failed adoption before it.

## 5. Measurements

Harness: `results/0268/harness/` (two standalone crates with the same
workload; `before` builds trunk `4d1afe92bc`, `after` this branch; release,
`lto = true`, `panic = "abort"`). Workloads: ODT 2,000 paragraphs, ODS
500×10 text cells, ODP 200 titled slides, all built by each side's own
builder (identical bytes on both sides: FNV digests `2b254dd9b0700b63`,
`a6719c4f1444fa48`, `966ebe1a60086c2f`). 25 timed samples per
operation per run after 3 warm-ups; 12 runs in ABBA order on CPU 7
(AMD EPYC 9R45, shared host); medians of the per-run medians. Raw TSVs in
`results/0268/runs-tmpfs/` (save to tmpfs) and `results/0268/runs-disk/`
(save to ext4).

| Operation | before (µs) | after (µs) | delta |
|---|---:|---:|---:|
| odt build | 812.0 | 825.6 | +1.7% |
| odt open (`from_bytes`) | 130.3 | 135.2 | +3.8% |
| odt open + paragraphs | 1101.4 | 994.3 | −9.7% |
| odt edit + commit | 2580.7 | 2300.9 | −10.8% |
| ods build | 91134.3 | 93868.8 | +3.0% |
| ods open (`from_bytes`) | 5589.1 | 5315.6 | −4.9% |
| ods open + workbook | 11190.3 | 5333.8 | −52.3% |
| ods edit + commit | 30096.8 | 29515.5 | −1.9% |
| odp build | 474.5 | 481.9 | +1.6% |
| odp open + slides | 832.8 | 855.1 | +2.7% |
| odp open + presentation | 898.3 | 67.2 | −92.5% |
| odp edit + commit | 2214.4 | 2229.7 | +0.7% |

Save, `before` = `std::fs::write`, `after` = atomic `save_with_durability(NoSync)`:

| Destination | odt | ods | odp |
|---|---|---|---|
| tmpfs | 2.6 → 8.0 µs | 3.8 → 9.4 µs | 1.6 → 6.8 µs |
| ext4 | 912.6 → 872.9 µs (−4.3%) | 976.9 → 899.2 µs (−8.0%) | 896.6 → 831.2 µs (−7.3%) |

Default `save()` (`Durability::Full`, file and directory `fsync`) is
4.9–5.0 ms on ext4 and 8.5–10 µs on tmpfs; trunk had no durable save.

Reading the results:

- **Two regressions found and fixed in this change.** The first candidate
  was +8.8% on ODT open: the new unit counter in `OpenParse` computed
  `local_name()` for every start tag. It now names only the root/body/text
  chain (+3.8% after). ODP build was +18.8% because adoption decoded
  the projection eagerly; it is now lazy (+1.6%). Routing ODP `from_bytes`
  through the prepared detector cost about 25 µs; it now indexes directly as
  trunk's `Snapshot::from_bytes` did.
- **The save rows over 5% are the intended behavior change.** On tmpfs the
  atomic save costs about 5 µs more per call (staging file + rename against
  one write). On ext4 it is faster than trunk's truncate-in-place write.
- **`to_bytes` deltas (+11–25%) are 10–20 ns on 0.1 µs copies,** below
  timer resolution.
- **The "open + projection" gains come from a different operation shape,
  not from faster parsing.** Trunk needed a second parse to get a
  `Spreadsheet`/`Presentation`/`Document` from its snapshot; the owner
  now keeps the projection it parsed. The `from_bytes` row is the like-for-like
  comparison.
- The ODS build cost (about 90 ms for 5,000 cells) is a pre-existing
  builder cost and is the same on both sides.

Regenerate: export trunk with `git archive 4d1afe92bc Cargo.toml .cargo
crates rust-toolchain.toml`, point `harness/before/Cargo.toml` at it, build
both crates with `cargo build --release`, then run
`BENCH=<dir> TARGET=<target> [O=<save dir>] harness/abba.sh` and
`python3 harness/summarize.py <runs dir>`.

## 6. Glossary gate

`tools/api_glossary.json` adds profiles `odf-common` (1 ban), `odt`
(47 bans; family `litchi_odt::transaction`), `ods` (51 bans; family
`litchi_ods::document`), `odp` (58 bans; family `litchi_odp::edit`) and
`ops-odf` (features `odt`, `ods`, `odp`). The bans cover the deleted names:
`Snapshot`, `Spreadsheet`, `Transaction`, `SourcePackageLimits`,
`ReadLimits`, `Builder::save`, `Patch::apply`, and the projection state
verbs. In `docs/API_GLOSSARY.md` these rows flip for ODT/ODS/ODP:
immutable owner, spreadsheet and deck projections, `Builder::build()`,
`Package::edit()`, the `Edit` type, commit accessor, diagnostics, split
commit, patch application, no-op introspection, `from_bytes(&[u8])`,
configured open, atomic save, `SourceLimits`, `SemanticLimits` and
ingress coverage. `litchi-odt`'s `drawing::resources` glob re-exports
became named module re-exports so the gate can resolve them.

## 7. Gates (all exit 0)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-odf -p litchi-odf-formula -p litchi-odc -p litchi-oth -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test --no-fail-fast --all-features -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-odf -p litchi-odf-formula -p litchi-odc -p litchi-oth` | 0 (4,031 passed) |
| `cargo test --no-fail-fast -p litchi-ops --all-features` | 0 (1,878 passed; capability goldens unchanged) |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 (260 passed) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-odf -p litchi-odf-formula -p litchi-odc -p litchi-oth -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (26 profiles, 1,201 bans, 36 families) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (246 tests) |
| `cargo check --all-targets` in `tools/perf-baseline` and `tools/native-resave` | 0 |

## 8. Gaps and follow-ups

- **G1 other ODF leaves.** `litchi-odb`, `litchi-odg`, `litchi-odi`,
  `litchi-odm`, `litchi-oth`, `litchi-odc` and `litchi-odf-formula` got only
  the `SourceLimits` rename and, where they save a whole package (ODC,
  formula), atomic publication. Their owner split, `from_bytes(&[u8])`,
  `OpenOptions` ingress and `Builder::build() -> Package` remain open.
- **G2 mutable writers.** `MutableDocument` (ODT), `MutableSpreadsheet` (ODS)
  and `MutablePresentation` (ODP) are still public authoring writers with
  their own atomic `save`. `MutableSpreadsheet` now starts only from a
  `Workbook` (`from_workbook`) and ends in `into_package`. The ADR 0034
  "no mutable writer on a package" row needs them folded into `Builder` or
  `Edit`.
- **G3 scoped families.** ODS scoped families (`worksheet`, `charts`,
  `data_pilot`, `annotations`, …) and the ODT plain-paragraph/joined
  families keep their own `Snapshot`/`Commit` vocabulary. Only the primary
  families are registered.
- **G4 source-backed names.** `SourceBacked{Document,Spreadsheet,Presentation}`
  take `OpenOptions` but keep their names (ODS still says `Spreadsheet`
  there), and the ops path-open route flattens their typed limits to a core
  `ResourceLimit` (still classified `limit_exceeded`).
- **G5 security vocabulary.** ODF signature and encryption refusals keep the
  existing errors, not ADR 0034's `Error::Signed`/`EncryptionPolicy`. There
  is no `save_plain`/`save_reencrypted` on ODF owners. An ODP projection
  opened with a password retains the credential (the owner's physical
  package does not).
- **G6 limits policy.** `SourceLimits::DEFAULT` input is 512 MiB while the
  smart detector admits up to 2 GiB (`MAXIMUM`). Callers that need more
  pass `OpenOptions`. `SemanticLimits` defaults equal their maxima.
- **G7 durability.** ODT annotation operations are not encodable into
  durable patches (refused by `Patch::durable` and transfers, as
  `RestoreSnapshot` already was).
- **G8 diagnostics.** ODT `PatchReplay` diagnostics are no longer public
  after the commit reshaping.
