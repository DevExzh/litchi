# 0253 — Shared OOXML metadata

- Status: implemented on `uo2/p6-ooxml-meta`, cut from `feat/unified-ops-api`
  at `908f9cccdd` and merged with trunk through records 0249, 0250 and
  0254–0256. Plan step 6 of record
  [0248](0248-coordinator-restart-plan.md) §6 Phase B, a P0 item.
- Date: 2026-10-13
- Scope:
  - `crates/litchi-core` metadata (`Metadata`, `Timestamp`, `Field`,
    `Keywords`);
  - `crates/litchi-ooxml-common/src/properties` (core codec, application
    properties, the shared transaction) and the small `custom` hooks it
    needs;
  - one `metadata` module each in `litchi-pptx`, `litchi-xlsx`,
    `litchi-xlsb`;
  - one `metadata` module per OOXML format in `litchi-ops` plus the shared
    `format/metadata.rs`, with small hunks in the shared dispatch files;
  - mechanical compile fixes for the merged type in `litchi-docx`,
    `litchi-ole-common`, `litchi-odf-common`, `litchi-odt`, `litchi-odp`,
    `litchi-ods`, `litchi-oth`, `litchi-doc` (test) and the facade.
- `performance_claim`: one paired probe (§7). It shows no frozen-harness
  qualification.

Path prefixes:
- `P/` = `crates/litchi-ooxml-common/src/properties/`;
- `O/` = `crates/litchi-ops/src/format/`.

## 1. One `Metadata` type (ADR 0034 §7, glossary §6)

`ooxml_common::Props` is deleted; `litchi_core::Metadata` is the one type.

- **Author.** `creator` is gone; `author` is the only spelling. `dc:creator`
  stays inside the core codec (`P/read.rs`, `P/write.rs`).
- **Real time types.** The pairs `created`/`created_local`,
  `modified`/`modified_local`, `last_printed_time`/`last_printed_local` and
  `last_backup_local` become one `Option<Timestamp>` each: `created`,
  `modified`, `last_printed`, `last_backup`. `Timestamp` is
  `Utc(DateTime<Utc>)` or `Unzoned(NaiveDateTime)`, so a zone-less producer
  time is never reinterpreted as UTC. It parses RFC 3339 and zone-less ISO
  8601 and serializes untagged. W3CDTF and `xsd:dateTime` lexical forms stay
  in the codec (`P/time.rs`).
- **OOXML-only fields.** `category`, `content_status`, `version` and
  `last_printed` are ordinary fields.
- **`Field`.** `metadata::Field` names every field; errors and diagnostics
  name fields with it, and `Metadata::changed_fields` drives the routing.
- **Keywords (pending).** `Keywords` (with `keywords::{Lang, Value, Item}`)
  moved to `litchi_core::metadata`, and the OOXML codec keeps
  `cp:keywords` mixed content losslessly. The `Metadata.keywords` field
  stays `Option<String>`: three iWork crates assign an `Option<String>` to
  it and may not be edited by this program. The one-line
  `.map(Keywords::from)` change is requested in
  [IWORK_API_REQUESTS.md](IWORK_API_REQUESTS.md) (2026-10-13), and the
  glossary row is marked pending on it.
- **Merged trunk users.** Record 0254's `SummaryProperties::merge_metadata`
  reads `Timestamp` through its one `utc_time` helper; unzoned times and
  `last_backup` stay refused there.

## 2. Shared core/app/custom properties (`P/`)

Ported from the stale reference `4dba739fac` (`document_properties`,
`extended`, `package_graph`), adapted to the trunk:

- **Core codec** (`properties.rs`, `read.rs`, `write.rs`). The record type
  is private (`Core`); public reads return `Option<Metadata>`. A write
  merges a `Metadata` onto the record it came from (`Core::merged`): an
  unchanged field keeps its exact lexical form (keyword mixed content, a
  W3CDTF offset, the Strict dialect). `write(package, &Metadata)` and the
  docx `Slot` refuse a field core properties cannot store, instead of
  dropping it.
- **Application properties** (`app/`, from `extended`). The 23 scalar
  `CT_Properties` elements; writes splice only changed elements, so
  `HeadingPairs`, `TitlesOfParts` and extension content keep their bytes.
  `app::METADATA_FIELDS` maps the 11 fields `Metadata` also carries.
- **The transaction** (`transaction.rs`, from `document_properties`):
  `Snapshot`, `Edit`, `Commit`, `Patch`, `Change`, `Diagnostics`, `Error`,
  `LimitKind`, `Family`. It is exact and reversible (`Patch::inverse`,
  `Patch::check`), creates missing parts with free names and relationship
  IDs, and refuses a stale patch (`Error::PatchConflict`).
  - Each changed field lands in the family that stores it.
  - A changed field no OOXML part stores (`content_type`, `last_backup`,
    `codepage`) is `Error::Refused(Refusal)` with `UnsupportedProperty`, never
    dropped.
  - An exact no-op commits an empty patch; a signed package keeps its
    bytes. A changed patch of a signed package is `Error::Signed`.
  - No value is invented: nothing stamps `modified`.
- **Custom.** The reference's `decode_for`, `encode`, `same_values` and the
  shared graph inspection were applied to `custom/`.

## 3. Leaf entries

| Crate | Owner | Read | Mutation | Family |
|---|---|---|---|---|
| pptx | `Package` | `metadata()`, `app_props()` | `edit_metadata()`, `apply_metadata(&Patch)` | `metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` |
| xlsb | `Package` | same | same | same |
| xlsx | `Package` (after 0250, §6) | same | same | same |

`Edit` stages `set`/`update` of the whole `Metadata`, `set_app_props` and
`set_custom_props`; `Commit::{package, into_package, patch, diagnostics,
into_parts}`. `Patch`, `Diagnostics`, `Error` and `LimitKind` are the shared
engine's. The gate registers all four families
(`litchi_{pptx,xlsx,xlsb}::metadata`, `litchi_ooxml_common::properties`) and
bans `Props` at the ooxml-common root and in `properties`.

## 4. Unified layer (`O/`)

- **Set path.** `set("/metadata", props)` for XLSX, PPTX and XLSB, aligned
  with record 0254: the same 22 OfficeCLI root-metadata keys (ASCII
  case-insensitive), `custom.<name>` (bool, integer, number or text), RFC
  3339 text for `created`/`modified`/`lastPrinted`, and `null` to remove.
  An unknown key is `unsupported_property` with a suggestion; a repeated
  key (including `company`/`extended.company`) is `invalid_argument`;
  `contentType` is refused by the leaf as `unsupported_property` naming
  `content_type`. The commit runs inside `File::edit` transactions too
  (pptx `Proof::Metadata`, xlsb `Patch::Metadata`, xlsx native patch slot).
- **Set-only.** As in 0254, `/metadata` is set-only: `get` gaps name
  `File::metadata()` and the root node. No generic readable `/metadata`
  node is added here, so 0254's set-only gap stays open for every format
  (follow-up F4).
- **Reads.** `File::metadata()` and the root node now include the
  application properties. The root gains OfficeCLI's read keys:
  `category`, `created`, `modified` (pptx; .NET round-trip form, closing two
  pptx gaps) and `extended.{template, manager, company, application, pages,
  words, characters, totalTime}` (xlsx, xlsb, pptx). One node-view golden
  gains the `extended.*` values OfficeCLI also prints.
- **Create.** `create_with(xlsx|pptx|xlsb, CreateOptions::metadata(..))`
  commits a metadata edit on the fresh package instead of refusing. The
  creation gap is narrowed to the three unstorable fields.
- **DOCX.** Only compile changes; its `props*` and create path stay with
  uo2/p7a (create still refuses non-core fields).
- Capability goldens were re-blessed (pptx, xlsb, xlsx, xlsx-encryption)
  and the OfficeCLI delta regenerated (xlsx 89, xlsb 62, pptx 161
  supported rows).

## 5. Tests

- core: `Timestamp` parse/serde, `changed_fields`, keyword model.
- ooxml-common: `transaction_tests.rs` (routing, exact inverse, vector
  bytes kept, signed no-op and refusal, unsupported fields, creation from
  empty, stale conflict), app codec tests, merge tests, the `Slot`.
- leaf: one test module per host crate (replay, inverse, conflict,
  unsupported).
- ops: `tests/ooxml_metadata.rs` (7 tests: set/reopen for three formats,
  refusals, no-op identity, transaction replay/inverse, create-time
  metadata, custom properties and timestamps, capabilities) and the
  admission unit test.
- Python: one expectation gains the application properties.

## 6. XLSX owner (record 0250)

The leaf module (`crates/litchi-xlsx/src/metadata.rs`) was written against
the pre-0250 owner `Workbook`, as a separate module so the W2 rename stayed
mechanical. After merging 0250 it hangs on `xlsx::Package`: `apply_metadata`
applies the patch to a copy of `Package::opc()` and adopts it through the
projection (`Workbook::adopt_published_package`, `Package::from_workbook`),
keeping encryption provenance and repair policy. The ops XLSX adapter
commits on its `Package` owner; the metadata step keeps an empty
`workbook::edit::Patch` in the transaction slot, and history compares the
retained endpoints' content, which covers the `docProps` parts. The xlsx
gate profiles register both families (`workbook::edit`, `metadata`).

## 6a. Reviewed golden changes

- Capability goldens (pptx, xlsb, xlsx, xlsx-encryption): the set-only
  `/metadata` gap, the narrowed create gaps and the new root read keys.
- `crates/litchi/tests/golden/markdown_corpus.tsv`: 131 rows, only the
  `metadata+html` column (Markdown with YAML front matter), whose metadata
  serialization changed by design: PPTX front matter gains the application
  properties (78 rows), and the merged time fields serialize as `created`,
  `modified`, `last_printed` instead of `created_local`,
  `last_printed_time` and the like (RTF 43, DOC 5, PPT 3, DOCX 2 rows). The
  other five columns are unchanged.
- One XLSX node-view golden gains OfficeCLI's `extended.*` values; one
  Python expectation gains `application` and `security`.

## 7. Measurements

`crates/litchi-ops/examples/ooxml_metadata_bench.rs`, release, `taskset -c 7`,
ABBA, 5 warm-up + 41 rounds, medians of each run; three runs at the final
commit. Raw output:
[results/0253/ooxml-metadata-bench.txt](results/0253/ooxml-metadata-bench.txt).

| Scenario | Leaf | Unified | Ratio (3 runs) |
|---|---|---|---|
| XLSX `/metadata` title, eager open | 8.8 µs | 9.2 µs | 1.04–1.06 |
| XLSX, source-backed open | 8.7 µs | 9.2 µs | 1.06 |
| PPTX, eager open | 71.9 µs | 72.6 µs | 1.01 |
| PPTX, source-backed open | 78.8 µs | 197.6 µs | **2.45–2.53** (miss) |
| XLSB, eager open | 32.7 µs | 36.0 µs | 1.09–1.11 (at the 10% bar) |
| XLSB, source-backed open | 32.7 µs | 45.4 µs | **1.37–1.39** (miss) |
| XLSX source open + `app.xml` read | 71.6 µs | +8.8 µs (12%) | within the 20 µs open budget |

- The leaf leg is the eager leaf package in every row, so the
  source-backed rows include the unified owner's promotion: PPTX adopts the
  source graph into an eager package (the leaf source owner has no
  part-topology metadata edit), XLSB materializes its workbook as every XLSB
  set does. The second and later sets on the returned handle are eager.
- Fixes made for the budget: the set reuses the staged metadata instead of
  re-reading the parts; XLSB keeps the leaf's validated projection and
  serializes transaction-history images lazily; PPTX adopts an unrevised
  source graph without re-serializing it.

## 8. Gaps and follow-ups

- **F1 — typed keywords.** Flip `Metadata.keywords` to `Option<Keywords>`
  after the iWork request lands.
- **F2 — PPTX/XLSB source-owner promotion** (§7 misses): a leaf
  source-backed metadata edit would remove the materialization.
- **F3 — DOCX.** `props*` → metadata `Edit`, `document_mut`, create-time
  application fields (uo2/p7a).
- **F4 — readable `/metadata` node.** A generic node kind shared by the
  workbook vocabulary, for all six formats.
- **F5 — ooxml-common residue from record 0249:** `custom_xml::Transaction`,
  `Transaction::is_changed`, `custom_xml::Commit::snapshot` (needs a
  family reshape of the in-place `&mut OpcPackage` owner and docx call
  sites) and `custom_data::Limits` (re-exported by xlsx/xlsb). Not cheap,
  not done.
- **F6 — OLE reads (from 0254).** `content_status`, `content_type`,
  `language`, `version`, `character_count_with_spaces` and
  `editing_time_minutes` are written by 0254 but not read by the OLE
  `metadata()` projection.
- **F7 — `/metadata` key table.** OOXML and OLE each keep a 22-key table
  with the same spellings; one shared table would remove the duplication.
- Pre-existing, unrelated: test targets that `include_str!` evidence purged
  on 2026-10-05 (`litchi-ods` formula oracles, `litchi-xlsx`
  form-control tests, `soapberry-zip`, two `litchi-opc` tests) do not
  compile; `litchi-iwa` chart examples have a type error; one `litchi-oth`
  DTD test fails.

## 9. Gates (final commit)

| Gate | Command | Exit |
|---|---|---|
| format | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy --keep-going -p litchi-core -p litchi-ooxml-common -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-docx -p litchi-ole-common -p litchi-odf-common -p litchi-odt -p litchi-odp -p litchi-ods -p litchi-oth -p litchi-doc -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 101, only from the 26 pre-existing test targets that `include_str!` purged evidence (23 litchi-ods, 3 litchi-xlsx); no lint |
| tests | `cargo test --no-fail-fast --all-features -p litchi-core -p litchi-ooxml-common -p litchi-pptx -p litchi-xlsb -p litchi-docx -p litchi-ole-common -p litchi-odf-common -p litchi-odt -p litchi-odp -p litchi-oth -p litchi-doc -p litchi-ops -p litchi` | 101: one pre-existing `litchi-oth` DTD test (`raw_zip_negatives_are_rejected_by_the_open_boundary`); everything else passes |
| tests (xlsx, ods) | `cargo test --all-features -p litchi-xlsx -p litchi-ods --lib` and `--doc` | 0 |
| ops | `cargo test -p litchi-ops --all-features` | 0 |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-core -p litchi-ooxml-common -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-ops -p litchi-ole-common` | 0 |
| crate boundaries | `python3 tools/check_crate_boundaries.py` | 0 |
| non-iWork | `python3 tools/non_iwork_gate.py verify` | 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 843 bans, 22 families) | 0 |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` (242 tests) | 0 |
| CLI | `cargo test -p litchi-cli --all-features` | 0 |
| goldens | `LITCHI_BLESS=1` capability tests (all features; `xlsx,xlsb,pptx` without encryption) and `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` | 0 |
