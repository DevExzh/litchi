# 0254 — Legacy OLE metadata for DOC, XLS and PPT

- Status: implemented on `uo2/p10-ole-meta`, cut from `feat/unified-ops-api`
  at `908f9cccdd`. Plan step 10 of record
  [0248](0248-coordinator-restart-plan.md) §6 Phase C.
- Date: 2026-10-13
- Scope:
  - `crates/litchi-ole-common` (the shared model, its `litchi_core::Metadata`
    conversion, one editor change);
  - the builder and `package::metadata` modules of `litchi-doc`, `litchi-xls`
    and `litchi-ppt`;
  - the DOC/XLS/PPT format modules of `litchi-ops`, with small hunks in the
    shared `error.rs`, `transaction.rs`, `create.rs`, `capabilities.rs` and
    `format/mod.rs`.
- `performance_claim`: one paired probe (§7). It shows no frozen-harness
  qualification.

Path prefixes:
- `C/` = `crates/litchi-ole-common/src/property_set/`;
- `O/` = `crates/litchi-ops/src/`.

## 1. Port of the stale merge

The source is commit `1f7d5edbbc`, merged on `uo/m2` as `b31129653f`; its
record is `docs/unified-ops/l-legacy-creators.md` on that branch.

- **Shared model.** `C/summary_properties/` was cherry-picked unchanged:
  `SummaryProperties` with its codec, name table and 13 tests.
  - The ole-common `property_set` tree on trunk was identical to the source's
    parent, so the patch applied cleanly.
- **Also ported:**
  - `Section::is_empty`;
  - the editor fix that gives a `DocumentSummaryInformation` stream it
    creates a code page;
  - the leaf transaction changes:
    - DOC and PPT typed edits create a missing stream and replace a blank
      section;
    - XLS reads a blank section as absent.
- **Re-expressed for the trunk.** The trunk replaced the doc/xls/ppt
  `Writer`s with ADR 0034 `Builder`s that `build()` a `Package`, so the
  writer wiring was rewritten rather than merged:
  - DOC and XLS write the two streams inside `build`;
  - PPT replaces its fixed blank streams only when properties are set.
- **Builder API (ADR 0034 §7).** The source's `Writer::set_property(name,
  text)` is not ported. The builders expose:
  - `Builder::set_metadata(Metadata) -> Result`, which replaces all
    properties;
  - `summary_properties()` and `summary_properties_mut()`, for custom and
    typed properties.

  The trunk `set_property` was deleted and is now banned on all three
  builders. On DOC it was a refusing stub; on PPT it silently dropped its
  argument (record 0002 finding 7).

## 2. Behaviour kept from the original

| Property | Where | Evidence |
|---|---|---|
| Output without properties is byte-identical (DOC/XLS: no property stream; PPT: the fixed blank streams) | builders | `output_without_metadata_*` in `crates/litchi-{doc,xls,ppt}/tests/*_metadata.rs`; ops `creation_authors_initial_metadata` (`create_with(.., Metadata::default())` = `create`) |
| Properties combined with CryptoAPI RC4 are refused | DOC `WriteError::EncryptedMetadata`, XLS `Error::EncryptedMetadata`, PPT `WriteError::EncryptedMetadata` (all PPT password-to-open profiles) | `cryptoapi_encryption_with_metadata_is_refused`, `encryption_policies`, `encryption_with_metadata_is_refused` |
| XOR and binary RC4 write the properties in plaintext | DOC, XLS | the same tests read them back after decryption |
| Names that are reserved (`_…`) or not text (counts, `Security`, …) are refused | `SummaryProperties::{set, remove}` | ported tests plus `remove_is_the_inverse_of_set` |
| A missing-stream transaction creates the stream, and the created section has a code page | editor fix, leaf transactions | ported unit tests; `editor_bridge_stages_only_changed_sections`; `metadata_transaction_*` (code page asserted) |
| `create` goes through the builder route, which keeps the root CLSID | `O/format/{doc,xls,ppt}/metadata.rs::create`/`author` | `builder_metadata_reads_back_through_every_reader` (DOC CLSID equals the plain build's) |

## 3. New in this step

### 3.1 One conversion from `litchi_core::Metadata`

`C/summary_properties/metadata.rs` holds the only reads of
`litchi_core::Metadata` fields in this work, through `from_metadata` and
`merge_metadata`.

- **Exhaustive.** The destructuring is exhaustive, so a new unified field
  fails to compile here until someone maps or refuses it.
- **Time fields.** The time fields go through one helper,
  `utc_time(field, utc, unzoned)`. The `Timestamp { Utc, Unzoned }`
  reshaping announced by `uo2/p6-ooxml-meta` is therefore a local edit of
  that helper and its three call sites.
- **Mapping.** The mapping table is in the rustdoc of `merge_metadata`.
- **Refused fields.** These fields are refused with
  `MetadataError::Unsupported { field }`, never dropped:
  - `identifier`;
  - `hyperlink_base`;
  - the unzoned timestamps, because FILETIME is UTC.
- **Refused values.** These are `MetadataError::Invalid { field, source }`.
  Errors name the field, never the value.

### 3.2 Opened-package metadata transactions (ADR 0034 §7)

Each leaf exposes:
- `Package::{metadata, edit_metadata, apply_metadata}`;
- `package::metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`.

`package::metadata` is registered as a family in the doc, xls and ppt
glossary profiles.

**How an edit works:**
- It reads the three sections into `SummaryProperties`
  (`SummaryProperties::from_editor`).
- It edits them through `Edit::{merge_metadata, properties_mut}`.
- It stages back only the sections that changed (`stage_changes`), through
  the existing format property-set transaction. That transaction keeps host
  validation and the DOC protection policy.

**Guarantees:**
- An exact no-op returns the source package.
- `Patch` is exact-source, and `inverse()` reverts it.
- `apply_metadata` refuses another source with `PatchConflict`.
- Encrypted packages are refused (DOC `Encrypted`, PPT `Encrypted`, XLS
  `Protected`). Signed packages are refused too: the shared editor rejects
  protected containers.

`LimitKind` is an empty, non-exhaustive enum: no caller-selected dimension
exists yet. Text is bounded at 1 MiB per value, and the republished package
is admitted under the source's limits.

### 3.3 Root class identifiers (record finding F-1, partly fixed)

`property_set::Editor::finish` now adopts the source layout
(`OleWriter::adopt_source_layout`). Every property-set transaction therefore
keeps the root and storage class identifiers; Word's
`00020906-0000-0000-C000-000000000046` was lost before. Tests:
- `editor_bridge_stages_only_changed_sections`;
- `word_documents_keep_their_other_streams`, `FancyFoot.doc`;
- the XLS and PPT fixture equivalents.

Still lost on a changed commit: state bits, timestamps and empty storages.

### 3.4 Unified layer

**Creation.**
- `create_with(Doc|Xls|Ppt, CreateOptions::metadata(..))` now authors the
  metadata through `Builder::set_metadata`.
- An unrepresentable field returns `unsupported_capability`, with `property`
  set to the field.
- The capability gap `initial-metadata` became
  `unrepresentable-initial-metadata`.

**The `/metadata` set path** (`O/format/ole_metadata.rs`).
- `set("/metadata", props)` stages through `edit_metadata` and republishes
  the handle.
- Keys come from a static sorted registry of 22 entries:
  - the root-metadata spellings `title`, `author`, `subject`, `keywords`,
    `description`, `category`, `lastModifiedBy`, `revisionNumber`, `created`,
    `modified`;
  - `company`, `manager`, `contentStatus`, `contentType`, `language`,
    `version`, `template`, `application`, `lastPrinted`;
  - the OfficeCLI `extended.{company,manager,template}` aliases.
- `custom.<name>` sets a user-defined property: text, bool, `VT_I4` or
  `VT_R8`.
- `null` removes a property.
- Errors:
  - an unknown key is `unsupported_property`;
  - a bad value is `invalid_value`.
- XLS: a call that carries a sheet property (`name`, `hidden`, `visibility`)
  still addresses a worksheet named `metadata`.

**Transaction history.**
- DOC/XLS/PPT steps retain `StepPatch::{Primary, Metadata}`.
- Metadata and content edits therefore mix in one `edit(..)`/batch history,
  which inverts exactly (`doc_metadata_and_text_edits_share_one_history`).

**Error mapping.**
- Three typed `Source::{Doc,Xls,Ppt}Metadata` variants with `Classify`
  impls; nothing is flattened to a string.
- The DOC `WriteError::UnsupportedProperty` variant is gone.

## 4. Decisions to confirm

1. **The `/metadata` path is set-only, by interception.**
   - Making it a navigable node needs a `metadata` vocabulary entry. In
     the workbook family that also means a new shared `workbook::Kind`
     variant, which changes the XLSX/XLSB specs and goldens.
     `uo2/p6-ooxml-meta` is expected to settle that vocabulary.
   - Until then, reads go through `File::metadata()`.
   - The capability report lists `get metadata path-read` as a gap rather
     than overclaim.
   - The coordinator should reconcile this with p6's OOXML spelling, root
     `/` or `/metadata`, when merging.
2. **`Builder::set_metadata` replaces all properties**, including custom
   ones; `Edit::merge_metadata` merges. The builder is fresh authoring
   (ADR 0034 §7), so replacing matches the ODF `set_metadata`.
3. **The read projection is unchanged.**
   - These fields are now written but not projected by
     `Package::metadata()` / `File::metadata()`: `content_status`,
     `content_type`, `language`, `version`, `character_count_with_spaces`
     and `editing_time_minutes`. They are readable through the typed
     sections.
   - Projecting them changes the recorded facade Markdown front matter
     (`crates/litchi/tests/markdown_corpus_golden.rs`, 102 rows), and p6 is
     reworking the same `From` impl, so the change is left to that merge.

## 5. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| `/metadata` is not a navigable node (`get`/`query`) | add after the shared metadata vocabulary (p6) lands |
| Read projection omits six written fields (§4.3) | extend `C/codec/package.rs` + re-record the Markdown corpus golden |
| Encrypted summary stream ([MS-OFFCRYPTO] 2.3.5.4) not written | L-OLE-W3 (source record F-2) |
| Property-set rebuild still drops state bits, timestamps, empty storages | F-1 remainder: carry storage metadata through `Editor::finish` |
| DOC/PPT `edit_user_defined_properties` start a section without a code page | source record F-3; the metadata family always supplies one |
| PPT fixed blank streams are non-conformant (no code page, `SystemIdentifier` 0) | kept byte-for-byte (F-4) |
| `LimitKind` has no dimension | add one if a caller-selected bound is needed |

## 6. Tests

| Crate | New |
|---|---|
| `litchi-ole-common` | 13 ported + 6 new (`summary_properties/tests.rs`: conversion round trip, refusals, merge, remove, editor bridge, blank sections) + 1 ported editor test |
| `litchi-doc` | 7 (`tests/doc_metadata.rs`) + 1 ported unit + updated lifecycle tests |
| `litchi-xls` | 6 (`tests/xls_metadata.rs`) + 1 ported unit |
| `litchi-ppt` | 6 (`tests/ppt_metadata.rs`, one gated on `encryption`) + 1 ported unit + updated writer test |
| `litchi-ops` | 4 (`tests/ole_metadata.rs`) + 2 unit (`format/ole_metadata.rs`); `xls_create`/`ppt_create` updated to the narrower refusal; capability goldens re-blessed (doc, xls, ppt) and the OfficeCLI delta regenerated |

## 7. Measurement

Probe:
- `crates/litchi-ops/examples/ole_metadata_bench.rs`, release build.
- Pinned with `taskset -c 5`; 21 ABBA samples × 50 repeats, median.
- Native leg: leaf `edit_metadata` → `set("title")` → `commit` → `to_bytes`.
- Unified leg: `File::set("/metadata", {"title"})` → `to_bytes`.
- Raw lines: [`results/0254/ole-metadata-bench.jsonl`](results/0254/ole-metadata-bench.jsonl).
- Three runs at `908f9cccdd` plus this change.

| Case | Bytes | Native µs | Unified µs | Overhead |
|---|---|---|---|---|
| DOC `FancyFoot.doc` | 27,136 | 83.8–85.2 | 85.9–86.3 | 1.4–2.4% |
| XLS `Simple.xls` | 13,824 | 57.2–57.5 | 57.7–58.0 | 0.7–1.3% |
| PPT `SampleShow.ppt` | 125,440 | 171.0–200.4 | 181.9–210.5 | 5.1–6.3% |

All three are within the single-op budget of ≤ 10%. The PPT difference is
the unified handle's republication: it reprojects the presentation and its
metadata.

Not measured:
- `Editor::finish` with layout adoption against the previous from-scratch
  rebuild (no frozen route covers property-set transactions);
- builder output with properties against the source writer.

## 8. Not claimed

- No OOXML metadata change.
- No navigable `/metadata` node.
- No frozen-harness performance qualification.
- No owner decision on §4.
