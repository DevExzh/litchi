# 0003 — M0a migration plan: flat `File`, single `litchi::open()`

- Status: plan (no code changed by this record)
- Date: 2026-10-04
- Scope: design §15 M0a and §16; goal `docs/UNIFIED_OPS_GOAL.md` (including the
  owner override and the 2026-10-04 tier priority)
- Base: `uo/0003-m0a-migration-plan` at `cc3c793e5d`. Every `file:line` below
  was read in that tree. "Inferred" marks a statement read from code but not
  executed.

Paths are relative to the repository root. `L/` means `crates/litchi/src/`.

## 0. Summary of decisions taken here (provisional, listed for owner confirmation)

| # | Decision | Section |
|---|---|---|
| D1 | `File` is **flat per format** (`File::Docx`, `File::Xlsx`, …), plus `File::Locked` and `File::Signed`. Families become accessors and borrowed views, not enum layers. | §2.1 |
| D2 | Each variant is one `Arc`-backed handle. `size_of::<File>() == 16`, `File: Clone + Send + Sync`, checked at compile time. | §2.2 |
| D3 | M0a variants: tier 1 `Docx/Xlsx/Pptx/Xlsb` and tier 2 `Doc/Xls/Ppt` retain an editable owner. Tier 3 formats the facade reads today (`Rtf/Odt/Ods/Odp`) carry over as **read-only** variants. `Pages/Keynote/Numbers` stay on `litchi::iwork::Document` (§16.4). `Odg/Odc/Odi/Odb/Oth/Odm/OdfFormula` are added in M6. The enum is `#[non_exhaustive]`, so adding a variant needs no reshaping. | §2.1, §2.3 |
| D4 | `litchi-ops` has **no iWork dependency in M0a**. `tools/non_iwork_gate.py` rejects any edge from a bulk package to an iWork crate, optional or not. The iWork classification of an unknown ZIP stays in the facade, as a fallback that only runs on the error path. | §3.3 |
| D5 | One `OpenOptions` struct with the common knobs plus optional per-format sub-options. It carries no credentials; unlocking is the consuming `Locked::unlock`. | §2.5 |
| D6 | There is one `litchi_ops::Error`, which keeps the typed leaf `source`, carries one `Code` from the closed list, and has an optional suggestion. `map_ooxml_error` and the sheet family's `Box<dyn Error>` are deleted. | §2.6 |
| D7 | Wrong-family refusals keep their semantics through the family views (`file.document()?`, `file.presentation()?`, `file.workbook()?`). They return `unsupported_format` and carry the detected format. | §2.4, §4.4 |
| D8 | Detection (`detection_smart`, without its iWork arms) moves into `litchi_ops::detect`. `open()` makes one container pass and hands the prepared owner (OPC package, OLE file, ODF index) to the constructor. | §2.7 |

---

## 1. Inventory of what exists today

### 1.1 Line counts

| Area | Files (lines) | Total |
|---|---|---|
| `L/document/` | `doc.rs` 4351 (tests from :1966), `table.rs` 876, `run.rs` 233, `types.rs` 264, `paragraph.rs` 165, `element.rs` 162, `cell_merge.rs` 82, `mod.rs` 61 | **6194** |
| `L/presentation/` | `prs.rs` 2838 (non-test code is 1–115 and 543–1616), `slide.rs` 268, `types.rs` 245, `mod.rs` 46 | **3397** |
| `L/sheet/` without `text/` | `workbook.rs` 2647, `adapters/xlsx.rs` 1069, `adapters/xlsb.rs` 601, `functions.rs` 477, `workbook_types.rs` 213, `selection.rs` 153, `mod.rs` 119, `adapters/mod.rs` 13 | **5292** |
| `L/sheet/text/` (stays, §16.4) | `workbook.rs` 319, `parser.rs` 286, `formats/{dif 262, fixed_width 248, sylk 235, delimited 154, tests 253, mod 15}`, `tests.rs` 161, `worksheet.rs` 154, `iterators.rs` 106, `cell.rs` 69, `mod.rs` 56 | 2318 |
| `L/iwork/` (stays, §16.4) | `mod.rs` 953, `model.rs` 769, `error.rs` 307, `limits.rs` 299 | 2328 |
| `L/detection_smart/` | `detected.rs` 4455, `functions.rs` 1006, `ooxml.rs` 313, `mod.rs` 92, `ole2.rs` 76 | **5942** |
| `L/lib.rs` | 846 (164 `//!` lines, 17 of which name `Document`/`Presentation`/`Workbook`) | 846 |
| Dependents in the facade | `markdown/` (`writer.rs` 2161, `docx.rs` 517, …), `tolerance.rs` 285 | — |

In-source unit tests that move with the code (`#[test]` counts): `doc.rs` 88, `prs.rs` 55, `workbook.rs` 62, `detected.rs` 35, `table.rs` 14, `types.rs` (document) 11, `adapters/xlsx.rs` 11, `functions.rs` 10, `presentation/types.rs` 10, `workbook_types.rs` 10, `element.rs` 6, `slide.rs` 6, `paragraph.rs` 5, `run.rs` 5, `cell_merge.rs` 4, `adapters/xlsb.rs` 1. That is **333** tests.

### 1.2 Public wrapper types and `*Impl` enums

**`Document`** is defined at `L/document/doc.rs:215` as `{ inner: DocumentImpl }`. It has no derives, no `Clone`/`Debug`, and no Send/Sync or size assertion. Its enum, `DocumentImpl` (`L/document/types.rs:32`, with `#[allow(clippy::large_enum_variant)]`), has these variants:

| Variant | Line | Payload | cfg | Owner kept |
|---|---|---|---|---|
| `Doc` | 35 | `(doc::Document, Metadata)` | `doc` | Reader only. `doc::Package` and the OLE file are dropped after open (`doc.rs:962-973`, `:1070-1081`). |
| `Docx` | 41–45 | `(Box<docx::Package>, Metadata, Box<ToleranceReport>)` | `docx` | Eager, editable. Reached only through `from_detected` (`doc.rs:987`). |
| `DocxSource` | 48 | `(docx::source_backed::Package, DocxSourceCache)` | `docx` | Source-backed, editable (`edit_document` at `litchi-docx/src/source_backed.rs:1142`). This is the default route. |
| `Pages` | 51 | `litchi_pages::Package` | `pages` | Eager. |
| `Rtf` | 54 | `RtfDocument<'static>` | `rtf` | Eager. |
| `Odt` | 57 | `litchi_odt::Document` | `odt` | Eager, built from a `PreparedPackage` (`doc.rs:1037`). |
| `OdtSource` | 60 | `litchi_odt::SourceBackedDocument` | `odt` + `any(unix, windows)` | Source-backed. |

`DocxSourceCache` is `OnceLock<Box<docx::Package>>` when `docx+markdown` are on (`types.rs:13-14`) and a ZST otherwise (`types.rs:19-21`).

**`Presentation`** is defined at `L/presentation/prs.rs:102` as `{ inner: PresentationImpl, cached_metadata: OnceLock<Option<Metadata>> }` (`:104`, `:108`). It has no derives. Its enum, `PresentationImpl` (`L/presentation/types.rs:35`), has these variants:

| Variant | Line | Payload | cfg | Owner kept |
|---|---|---|---|---|
| `Ppt` | 38 | `ppt::Presentation` | `ppt` | Reader only. The byte path drops `ppt::Package` (`prs.rs:892-911`). The path route drops `ppt::SourceBackedPackage` (`prs.rs:785-796`), although the leaf still keeps an `Arc<SharedOleFile>` for lazy pictures (`litchi-ppt/src/presentation/model.rs:36-37`). |
| `Pptx` | 41 | `Box<pptx::Package>` | `pptx` | Eager, editable (`opened_presentation_transaction`, `litchi-pptx/src/package/model.rs:296`). Reached only through `from_detected`. |
| `PptxSource` | 44 | `pptx::SourceBackedPresentation` | `pptx` | **Read-only.** The type says so itself (`litchi-pptx/src/presentation/source.rs:408`); the editable owner is the separate `SourceBackedPresentationEditor` (`:672`). |
| `Keynote` | 47 | `litchi_keynote::Package` | `keynote` | Eager. |
| `Odp` | 50 | `litchi_odp::Presentation` | `odp` | Eager. |
| `OdpSource` | 57 | `litchi_odp::SourceBackedPresentation` | `odp` + `any(unix, windows)` | Source-backed. |

**`sheet::Workbook`** is defined at `L/sheet/workbook.rs:265-276` as `{ inner: WorkbookImpl, cached_metadata: Metadata, metadata_repairs: MetadataRepairs }`. Its enum, `WorkbookImpl` (`L/sheet/workbook_types.rs:16-55`), has these variants:

| Variant | Line | Payload | cfg | Note |
|---|---|---|---|---|
| `Numbers` | 18 | `litchi_numbers::Package` | `numbers` | |
| `Xlsx` | 22 | `adapters::Workbook`, which is `WorkbookModel::{Owned(xlsx::Workbook), Source(xlsx::SourceBackedWorkbook)}` (`adapters/xlsx.rs:81-89`) | `xlsx` | This is the two-state pattern §16.3.7 calls for. `Owned` is editable (`xlsx::Workbook::edit`, `litchi-xlsx/src/workbook/model.rs:1313`). `Source` is read-only (`workbook/source.rs:117`). |
| `Xlsb` | 24 | `xlsb::Workbook` | `xlsb` | Eager, editable (`edit_cell_values`, `litchi-xlsb/src/workbook/model.rs:380`). |
| `XlsbSource` | 26 | `adapters::XlsbWorkbook` holding `xlsb::SourceBackedWorkbook` (`adapters/xlsb.rs:37`) | `xlsb` | Read-only. The leaf has no source-backed edit API. |
| `XlsFile` | 34 | `xls::Workbook<BufReader<File>>` | `xls` | **Never constructed.** It appears only in match arms. |
| `XlsMem` | 36 | `xls::Workbook<Cursor<Vec<u8>>>` | `xls` | |
| `Ods` | 41 | `RefCell<litchi_ods::Spreadsheet>` | `ods` | **`RefCell` makes `Workbook` `!Sync`.** |
| `OdsSource` | 45 | `litchi_ods::SourceBackedSpreadsheet` | `ods` + `any(unix, windows)` | |
| `XlsSource` | 49 | `XlsSource { SourceBackedWorkbook, Arc<SharedOleFile>, OnceLock<Metadata> }` (`workbook.rs:118`) | `xls` + `any(unix, windows)` | |
| `Other` | 54 | unit | `any(xls, xlsx, xlsb)` | **Never constructed.** |

The facade matches on these enums in three files only: `doc.rs` 127 times, `prs.rs` 84 times and `workbook.rs` 74 times. No other facade file matches an `*Impl` enum.

**`iwork::Document`** (`L/iwork/model.rs:19`) is `{ state: Arc<State{format, data: StructuredData}> }`, a snapshot only with no editable owner. Its openers are at `L/iwork/mod.rs:79-144` and go through `decode` at `:179`. It stays put in M0a (§16.4; goal tier 3).

### 1.3 Public open functions (all to be replaced)

| Owner | Functions (file:line) | Options type |
|---|---|---|
| `Document` | `open` doc.rs:686, `open_with_limits` :745, `from_bytes` :841, `open_with_doc_options` :861, `from_bytes_with_doc_options` :881, `from_bytes_with_limits` :891, `from_prepared` :934 | `docx::ReadLimits`, `doc::OpenOptions`, `opc::PreparedOoxmlSource` |
| `Presentation` | `open` prs.rs:569, `open_with_limits` :668, `from_bytes` :758, `from_bytes_with_limits` :801, `from_prepared` :860 | `pptx::ReadLimits`, `PreparedOoxmlSource` |
| `sheet::Workbook` | `from_prepared` workbook.rs:292, `open` :340, `from_bytes` :437 | none |
| `sheet::functions` (all return `Box<dyn WorkbookTrait>` or a leaf type under `Box<dyn Error>`) | xlsx `open_workbook*` functions.rs:84/94/124/130; xls `open_xls_workbook*` :148/161/174/183; xlsb `open_xlsb_workbook*` :195–314; text `open_{csv,tsv,prn,text,sylk,dif,fixed_width}_workbook*` :342–468 | `xlsx/xlsb::ReadLimits`, `text::TextConfig` |
| Detection | `detect_file_format` functions.rs:50, `detect_and_prepare` :98, `detect_and_prepare_with_limits` :112, `detect_file_format_from_bytes` :147, `detect_format_from_reader` :253, `detect_format_smart` detected.rs:175, `detect_format_smart_with_limits` :1807 | `opc::ReadLimits` |

There is **no top-level `open`** in `L/lib.rs`. The root re-exports are `Error, Result` (:728), `Document, DocumentElement` (:737), `Presentation` (:745), `SlideDescriptor` (:748), `Workbook` (:757), `ToleranceReport` (:775), `FileFormat, …` (:778), `detect_file_format*` (:796) and `PreparedDetection, detect_and_prepare*` (:802).

### 1.4 Detection, TOCTOU guards and polyglot precedence

- **Byte precedence.** OLE2 is tried first. For a ZIP, OOXML is tried next unless the input is a validated plain ODF package, then ODF, then iWork. Compressed RTF comes after the container formats (`L/detection_smart/functions.rs:160-210`).
- **Per-family path detectors.**
  - Document: `detect_document_source_path_with_limits` (detected.rs:2411) and the non-docx variant (:2495).
  - Presentation: `detect_pptx_source_path_with_limits` (:2812) and `detect_presentation_source_path` (:3094/:3183).
  - Workbook: `detect_workbook_source_path` (`WorkbookSourcePathDetection` :753).
  - The byte detectors are `detect_docx_source_bytes` (:449), `detect_presentation_source_bytes` (:652), `detect_workbook_source_bytes`, `detect_prepared_odp[_with_limits]` (:1666/:1653) and `detect_prepared_ods[_with_limits]`.
- **Prepared-source reuse.**
  - `detect_and_prepare_path_with_limits` (detected.rs:3416) builds one `SourceBackedPackage` (:3481) and a `PreparedOoxmlSource` (:3504). The resulting `prepared` is `None` for OLE2, ODF, RTF and iWork (:3318-3333).
  - `detect_format_smart` hands the parsed `OpcPackage` and `OleFile` back through `DetectedFormat` (:29-69), and ODT gets its `PreparedPackage` back the same way.
- **TOCTOU guards.**
  - `finish_source_docx_result` (doc.rs:293, 15 call sites) and `finish_source_odt_result` (:306, 9 call sites).
  - `check_source_freshness_after_markdown` (:647).
  - The presentation side uses up-front `check_source()` calls instead (prs.rs:1024, 1075, 1126/1131/1136, 1212, 1271, 1463).
- **Polyglot arbitration.**
  - detected.rs:2437-2458 decides between ODT and DOCX.
  - `odt_source_candidate_has_ooxml_owner` (:2642) handles the case where an ODT candidate is also claimed by an OOXML owner.
  - `PptxSourcePathDetection::DisabledOtherOoxml` (:2088) and the precedence rule documented at detected.rs:641-643 cover presentations.
- **Fallback ceilings.** These are `UNIFIED_DOCUMENT_FALLBACK_MAX_INPUT_BYTES` = 2 GiB (detected.rs:2173) and `UNIFIED_PRESENTATION_FALLBACK_MAX_INPUT_BYTES` (:2130).

### 1.5 Error mapping

- `map_ooxml_error` (`L/lib.rs:427-433`) maps every error to `Error::InvalidFormat(error.to_string())`. It has **73 call sites**: 35 in `document/` (17 in doc.rs, 11 in table.rs, 5 in run.rs, 2 in paragraph.rs) and 38 in `prs.rs`.
- **Duplicated helpers.** `map_source_opc_error` appears at doc.rs:236 and prs.rs:610, and `map_opc_read_resource` at doc.rs:257 and prs.rs:631. Other helpers are `map_source_docx_error` (doc.rs:224), `map_pptx_catalog_error` (prs.rs:49) and the identity function `map_odp_error` (prs.rs:112).
- **Inconsistency.**
  - In `text`, `slide_count`, `slides`, `slide`, `metadata` and `tolerance_report`, PPTX `check_source()` failures go through `map_ooxml_error`. A source change therefore becomes `InvalidFormat(String)`.
  - The openers and `slide_catalog` keep `Error::SourceChanged`.
  - The test at prs.rs:2484-2504 asserts only a message substring.
- **String flattening outside OOXML.** ODT, Pages and RTF errors become `Error::ParseError(format!(..))` in 19 places in doc.rs. Keynote and eager ODP do the same in 10 places in prs.rs.
- **Sheet family.** `sheet::types::Result` is `Box<dyn Error + Send + Sync>` (`litchi-core/src/sheet/types.rs:4`). The literal `Box<dyn std::error::Error` appears 37 times in `sheet/`.
- **Other error types.**
  - `litchi_core::Error` (`litchi-core/src/error/types.rs:14`, `#[non_exhaustive]`, 18 variants) has **no encrypted, password or signed variant**.
  - `iwork::Error` is a separate structured type (`L/iwork/error.rs:108`).

### 1.6 Sidecar state

- `Doc`: eager `Metadata` read from OLE property sets (doc.rs:966-970).
- `Docx`: eager `Metadata` (:1007-1011) and `Box<ToleranceReport>` (:1004).
- `DocxSource`: the markdown `OnceLock` cache.
- `Presentation`: `cached_metadata: OnceLock<Option<Metadata>>`. It is eager for PPT, Keynote and eager PPTX. For PPTX source it is filled lazily (prs.rs:1484-1489). For ODP and ODP source it is always `Some(Metadata::default())` (:587, 697, 765, 974).
- `Workbook`: `cached_metadata` and `metadata_repairs` (workbook.rs:268, 275), plus `XlsSource.metadata: OnceLock<Metadata>`.
- The return shapes differ. `Document::metadata` returns `Result<Metadata>` (doc.rs:1872), while `Presentation::metadata` returns `Result<Option<Metadata>>` (prs.rs:1459).
- `L/tolerance.rs:35-75` holds a per-format `ToleranceReport` enum that the facades construct.

### 1.7 cfg gates

- **Format features.** `doc`, `docx`, `ppt`, `pptx`, `xls`, `xlsx`, `xlsb`, `rtf`, `odt`, `ods`, `odp`, `pages`, `keynote` and `numbers` (`crates/litchi/Cargo.toml` `[features]`). `iwork` turns on all three iWork formats and `internal-iwork-source`.
- **The facade has no feature for odg, odc, odi, odb, oth, odm or ODF formula,** and does not depend on those leaf crates. They are detected (`litchi-odf-common/src/detect.rs:235-251`) but fall through to `UnexpectedFormat` (for example workbook.rs:621-628).
- **Target gates.**
  - `OdtSource`, `OdpSource`, `OdsSource` and `XlsSource` require `any(unix, windows)`.
  - `PreparedDetection` and `detect_and_prepare*` require OOXML plus `any(unix, windows)` (`detection_smart/mod.rs:79-91`).
  - The open paths fork on target, for example doc.rs:794 versus :752.
- **Possible compile gap (not verified).** `Presentation::open`'s `OdpSource` arm (prs.rs:582-588) is gated only on `odp`, but the variant also needs `any(unix, windows)` (presentation/types.rs:52).
- **The `Box<dyn WorkbookTrait>` text openers** (functions.rs:342-477) are not feature-gated.

### 1.8 Format-specific leak methods

These become per-variant handle methods:

- `Document::docx_section_inventory[_with_limits]` (doc.rs:1193/1207)
- `Cell::v_merge` (table.rs:387)
- `Presentation::slide_catalog` (prs.rs:1111) with `SlideDescriptor` (:24)
- `slide_width`/`slide_height`, which only work for PPTX (:1362/:1407)
- `Workbook::sheet` (workbook.rs:645-658, XLSX only) and its selection types (`L/sheet/selection.rs:12,39-42`)

Two read semantics differ by format today and **are kept unchanged in M0a**:

- Sheet names match case-insensitively in XLSX (`litchi-xlsx/src/sheet.rs:178,186`) but exactly in XLSB (`litchi-xlsb/src/workbook/source.rs:930`) and in text formats (`L/sheet/text/workbook.rs:289-295`).
- Pages supports `paragraph_text()` (doc.rs:1275) but not `paragraph()`, which returns `Unsupported` (:1372).

### 1.9 Element wrapper enums

These enums take about 1.6k lines. Design §16.1 deletes them in M1, not M0a. All derive `Debug, Clone`.

- `DocumentElement` (element.rs:29)
- `Paragraph` (paragraph.rs:17; variants `Doc`, `Docx`, `Pages(String)`, `Rtf`, `Odt`)
- `Run` (run.rs:16)
- `Table`, `Row` and `Cell` (table.rs:17/122/255)
- `CellMerge` (cell_merge.rs:188)
- `Slide` (slide.rs:11; no derives)
- `SlideData` and `LegacySlideData` (presentation/types.rs:13/20)

Markdown matches on these per-format variants 49 times (writer.rs 44, docx.rs 3, document.rs 2).

---

## 2. Proposed Rust shape (`litchi-ops`)

### 2.1 Flat per-format `File` (D1)

```rust
/// One opened input. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
#[non_exhaustive]
pub enum File {
    // Tier 1 — retain an editable owner (M0a)
    #[cfg(feature = "docx")] Docx(Docx),
    #[cfg(feature = "xlsx")] Xlsx(Xlsx),
    #[cfg(feature = "pptx")] Pptx(Pptx),
    #[cfg(feature = "xlsb")] Xlsb(Xlsb),
    // Tier 2 — retain an editable owner (M0a)
    #[cfg(feature = "doc")]  Doc(Doc),
    #[cfg(feature = "xls")]  Xls(Xls),
    #[cfg(feature = "ppt")]  Ppt(Ppt),
    // Tier 3 — read-only in M0a; write handlers arrive in M6
    #[cfg(feature = "rtf")]  Rtf(Rtf),
    #[cfg(feature = "odt")]  Odt(Odt),
    #[cfg(feature = "ods")]  Ods(Ods),
    #[cfg(feature = "odp")]  Odp(Odp),
    // M6 adds Odg, Odc, Odi, Odb, Oth, Odm, OdfFormula, Pages, Keynote, Numbers
    // (non_exhaustive: purely additive)
    /// Encrypted input: not readable or editable until `Locked::unlock`.
    Locked(Locked),
    /// Signed input: readable; edit verbs refuse until explicitly unsigned/resigned.
    Signed(Signed),
}
```

The alternative was `File::{Document(DocumentKind), Presentation(..), Workbook(..)}` with an inner enum per family. Flat per format is chosen for five reasons:

1. **The owner override needs it.** Every claimed format gets a variant that keeps its leaf owner. Seven of the 21 claimed formats (odg, odc, odi, odb, oth, odm, ODF formula; `litchi-core/src/detection/types.rs:36-48`) fit none of the three families. Grouped, they would need new groups (drawing, chart, image, database, template, master, formula), which is the reshaping the coordinator ruled out. Flat, M6 adds variants to a `#[non_exhaustive]` enum.
2. **ADR 0007 asks for this.** It wants "autodetection returns a flat data-bearing `File` enum" (docs/adr/0007-office-object-models.md:6-7). Design §4 calls the enum "flat" and says "the variant names are this design's choice", so `{Document, Presentation, Workbook}` is a naming choice, not a constraint from another ADR. The design text says "flat" but lists three family names; this record treats that as a **finding** (F2) for the owner to confirm.
3. **Dispatch is per format anyway.** The property registry (§6) is keyed by element × format. Every write handler is per format (ADR 0003 Edit types differ per leaf). A grouped enum would match twice on every operation.
4. **ADR 0004 holds either way.** Both shapes use static dispatch with no boxed trait objects. Flat gives one `match` with no nested discriminant.
5. **Families stay available,** as data rather than structure: `file.family() -> Family` (`Document | Presentation | Workbook | Other`, non-exhaustive) and borrowed family views (§2.4).

### 2.2 Handles: size, sharing, Send + Sync (D2)

Every per-format payload is a single `Arc` newtype, for example:

```rust
#[derive(Clone)] pub struct Docx(Arc<DocxState>);
struct DocxState {
    owner: DocxOwner,                         // two-state, §2.3
    metadata: OnceLock<Option<Metadata>>,     // sidecar, §1.6
    tolerance: OnceLock<ToleranceReport>,
    #[cfg(feature = "markdown")] markdown: OnceLock<Box<docx::Package>>,
}
enum DocxOwner { Source(docx::source_backed::Package), Eager(docx::Package) }
```

- **Size.** `File` is then a tag plus one pointer. The compile-time assertions are `const _: () = assert!(size_of::<File>() == 16);` and `const fn _assert<T: Send + Sync + Clone>() {}` applied to `File`, every handle, `Locked`, `Signed` and `Error`.
- **Clone.** Cloning is one atomic increment, which meets the "cheap to clone" requirement in the goal's handle-cost budget.
- **What blocks `Sync` today.**
  - `WorkbookImpl::Ods(RefCell<..>)` (workbook_types.rs:41). This becomes `OnceLock`/owned state or `Mutex`, decided in step S5e by what `litchi_ods::Spreadsheet` read APIs need.
  - `XlsMem(xls::Workbook<Cursor<..>>)`. `xls::Workbook` derives only `Debug` and owns `OleFile<R>`. The byte route instead normalizes onto `xls::SourceBackedWorkbook::from_read_at(OwnedSource)` (`litchi-xls/src/workbook/source.rs:869`), so both routes produce the same owner.
- **Unknown Send/Sync status.** `pptx::SourceBackedPresentation`, `xlsx/xlsb::SourceBackedWorkbook`, `ppt::SourceBackedPackage` and the doc/ppt snapshots have no Send/Sync assertion today. The only ones that exist are `docx::source_backed::Package` (`litchi-docx/tests/source_backed_semantic.rs:566-571`), `xlsx::Workbook` (`litchi-xlsx/src/workbook/tests.rs:171-196`) and `ppt::slide_order::Snapshot` (`slide_order.rs:561-567`). S6's assertions are the first compile-time proof for the rest (risk R3).
- **Commit shapes.** Write verbs return a new `File`, so the handle never needs `&mut`. Several leaf publishers consume `self`:
  - `docx::source_backed::Package::publish_document_commit_to_stream(self, …)` (`source_backed.rs:1500`)
  - `SourceBackedPresentationEditor::publish_slide_commit_to_stream(self, …)` (`source.rs:1727`)

  M2 must therefore get the owner out of the `Arc` (`Arc::try_unwrap` with a cheap re-open fallback, or a `&self` publish API added to the leaf). M0a only records this (R4).

### 2.3 What each variant retains (eager/source duality normalized)

The rule follows the pattern of the xlsx adapter (`adapters/xlsx.rs:81-89`). There is one variant per format, holding an owner enum with two states, `Source(..)` and `Eager(..)`. Today's routing decides which state is built: source-backed on `any(unix, windows)` path and byte routes, eager on fallback and other targets.

"Editable owner" in M0a means: the variant holds the leaf owner, **plus the shared source handle that every leaf edit entry can be built from without another detection pass or another `ReadAt` open**. Where building the leaf editor would still re-index the container, a leaf API request is filed. These requests may be implemented in the leaf because none of these crates is iWork (goal rule 7).

| Variant | Source state | Eager state | Edit entry reachable | Leaf gap / request |
|---|---|---|---|---|
| `Docx` | `docx::source_backed::Package` | `docx::Package` (+ metadata, tolerance) | `edit_document` (`source_backed.rs:1142`; eager `package/package/document.rs:36`) | none |
| `Pptx` | `pptx::SourceBackedPresentation` (read) | `pptx::Package` | eager `opened_presentation_transaction` (`package/model.rs:296`). From source: `SourceBackedPresentationEditor` (`source.rs:672`) | **L-PPTX-1:** construct the editor from the already-indexed `SourceBackedPackage` behind a `SourceBackedPresentation`. Fallback: build the editor lazily at the first edit from the retained `Arc<dyn ReadAt>`, which re-indexes once at first edit (not at open), and measure it. Design §16.2 says `SourceBackedPresentation` is the mutable owner; that is incorrect (F1). |
| `Xlsx` | `xlsx::SourceBackedWorkbook` (read) | `xlsx::Workbook` (functional `edit`, model.rs:1313) | eager `edit()`. Source edits go through per-feature `SourceBackedEditor`s (`workbook/source_merge.rs:39`, `cell_values/source.rs:190`, …) | **L-XLSX-1:** expose the `SourceBackedPackage` (or an editor factory) from `SourceBackedWorkbook` without re-indexing. |
| `Xlsb` | `xlsb::SourceBackedWorkbook` (read; there is no source edit API) | `xlsb::Workbook` (`edit_cell_values` :380) | eager only | **L-XLSB-1:** `SourceBackedWorkbook` → eager `Workbook` from the same shared bytes or source (materialize on first edit), or source-backed editors. |
| `Doc` | retained `Arc<dyn ReadAt>` + `body_text::source::SourceSnapshot` (`body_text/source.rs:495`, Clone) | `body_text::Snapshot` over `Arc<[u8]>` (`body_text.rs:609`) | `SourceSnapshot::edit` (:890); `Snapshot::edit` (:930) | **L-DOC-1:** build `SourceSnapshot` from the `SharedOleFile` that detection already parsed. The facade also keeps the reader `doc::Document` for M0a reads; today it drops the package (doc.rs:962-973). |
| `Xls` | `xls::SourceBackedWorkbook` + `Arc<SharedOleFile>` (already kept, workbook.rs:118) | same type over `OwnedSource` (replaces `XlsMem`) | `cell_values::Snapshot` (`cell_values/mod.rs:550`) and `SourceBackedOverlayPublisher` over the shared CFB (`litchi-ole-common/src/source_backed_overlay.rs:22`) | **L-XLS-1:** an editor constructor from `Arc<SharedOleFile>`. Today `cell_values::Snapshot::from_bytes_with_limits_and_cfb` (:584) wants owned bytes. |
| `Ppt` | retain `ppt::SourceBackedPackage` (Clone, `Arc<SharedOleFile>`, `package/source.rs:27`) instead of dropping it (prs.rs:785-796) | retain input `Arc<[u8]>` | `text_edit::SourceSnapshot` (`text_edit.rs:349`), `slide_order::Snapshot` (`slide_order.rs:572`) | **L-PPT-1:** build `text_edit::SourceSnapshot` from the shared CFB. |
| `Rtf`, `Odt`, `Ods`, `Odp` | carried over as is (`OdtSource`/`OdsSource`/`OdpSource` where present) | as is | none in M0a (read-only) | M6 |

**Open must not get slower.** No editor is built at open. Each variant holds (a) the read owner today's route already builds and (b) a handle (`Arc`) to the source that route already created. The editable form is built at the **first edit verb** (M2), from (b), through the L-* constructors:

- XLSX: the lazy `SourceBackedWorkbook` gets a per-feature `SourceBackedEditor` over the same `SourceBackedPackage`.
- XLSB: the lazy source is materialized into an eager `Workbook`.
- PPTX: a `SourceBackedPresentationEditor` is built.
- DOC/XLS/PPT: the edit snapshots are built over the retained `SharedOleFile` or bytes.

Record 0002 (`feat/unified-ops-api` `cfe7681627`, §"Facade holders today") found that the facade drops or never holds the editable object for DOC, PPT and RTF, and holds only the non-editable lazy form for XLSX and XLSB. This lazy edit-time construction is the M0a answer to that finding for every tier 1 and tier 2 variant. RTF is tier 3 and read-only in M0a. Its eager `RtfDocument` already has `Document::edit()` (0002 table), so M6 only needs a handler. S1/S10 open timing is the gate that retention adds nothing measurable to open.

**Memory cost of retention.**
- Source-backed retention costs one `Arc` of a handle the leaf already holds.
- The eager byte routes for `Doc` and `Ppt` keep the input bytes alive, which the facade released before. That is a memory regression of up to the input size on those routes. It is measured in S10, and if the owner prefers, it can be put behind `OpenOptions::retain_source(false)` as an opt-out. This is the trade-off the owner override (2026-10-04) accepted when it replaced the "legacy package retention" decision point.

### 2.4 Read surface in M0a (interim, replaced in M1)

M0a must not regress any read a consumer uses today; the Python read API in particular stays source-compatible (§7). The surface is:

```rust
impl File {
    pub fn format(&self) -> FileFormat;            // Signed reports the inner format
    pub fn family(&self) -> Family;
    pub fn text(&self) -> Result<String>;
    pub fn metadata(&self) -> Result<Option<Metadata>>;   // unifies §1.6 shapes
    pub fn tolerance_report(&self) -> Result<ToleranceReport>;
    pub fn document(&self) -> Result<DocumentView<'_>>;     // paragraphs, tables, elements, …
    pub fn presentation(&self) -> Result<PresentationView<'_>>; // slides, slide_count, …
    pub fn workbook(&self) -> Result<WorkbookView<'_>>;     // worksheet_names, … ; impl WorkbookTrait
}
```

- **Family views refuse the wrong family.** Each view returns `Code::UnsupportedFormat` with `detected: Some(file.format())` when the file is not of that family. This carries over the wrong-format refusal semantics (§4.4).
- **`WorkbookView` implements `WorkbookTrait`** statically. `litchi_eval::FormulaEvaluator<'a, W: WorkbookTrait + ?Sized>` (`litchi-eval/src/lib.rs:160`) then works without `Box<dyn>`. This replaces `open_workbook(..)?.as_ref()` in the 8 `sheet_eval_*` examples and the 3 `eval_xlsx_integration` tests.
- **The element wrapper enums move unchanged** into `litchi_ops::compat::{word, slide}`. Markdown (49 matches) and the Python `Paragraph`/`Run`/`Table` classes need them. They are scheduled for deletion in M1, when the node view and ADR 0007 vocabulary replace them (§16.1). Every compat item's rustdoc says so.
- **Leak methods become methods of the variant's handle:** `Docx::section_inventory`, `Pptx::slide_catalog`, `Pptx::slide_size`, `Xlsx::sheet(selector)`. They are not part of the generic surface.

### 2.5 `OpenOptions` (D5)

```rust
#[derive(Debug, Clone, Default)]
#[non_exhaustive]
pub struct OpenOptions {
    max_input_bytes: Option<u64>,                 // replaces the 2 GiB UNIFIED_*_FALLBACK ceilings
    execution: Option<litchi_core::ExecutionContext>,   // ADR 0031; Clone (execution.rs:257)
    source: SourceMode,                           // Auto (today's routing) | Eager
    retain_source: bool,                          // default true (§2.3)
    // per-format sub-options; None = leaf default
    ooxml: Option<litchi_opc::ReadLimits>,        // docx/pptx/xlsx/xlsb all re-export it
    ooxml_cache: Option<litchi_opc::SourceCacheLimits>,
    doc: Option<DocOptions>,                      // leniency; NO password (litchi_doc::OpenOptions is !Clone and holds one)
    ppt: Option<litchi_ppt::RecordLimits>,
    odf: Option<litchi_odf_common::SourcePackageLimits>, // the odt/ods/odp ReadLimits alias
}
```

- **Entry points.** `open(path)`, `open_bytes(Vec<u8>)` and `open_with(path, &OpenOptions)` / `open_bytes_with(..)`. Builders use short names (`.ooxml_limits(..)`, `.execution(..)`).
- **No credentials.** Credentials must be non-Clone (ADR 0006:133-135), so they never enter `OpenOptions`.
- **Signed input** always surfaces as `Signed`. There is no option to strip it silently.
- **M6 additions** are the fields `iwork: Option<iwork::Options>` and `odf_minor: …`, added in a non-breaking way thanks to `#[non_exhaustive]` plus private fields.

### 2.6 `Locked`, `Signed`, and the error enum (D6)

**`Locked`** is `pub struct Locked(Arc<LockedState>)`. It holds the detected inner format and the encryption kind, and keeps the shared source.

- **Producing it.**
  - OOXML: a CFB that has the `EncryptionInfo` and `EncryptedPackage` streams but none of the stream names `detection_smart/ole2.rs:38-66` accepts. Today this falls through to `NotOfficeFile` (inferred; F8).
  - Legacy: a leaf open that reports `PasswordRequired` (`litchi-doc/src/package/model.rs:464`, `litchi-xls/src/error.rs:42`, `litchi-ppt/src/package/model.rs:181`).
  - The kind (agile, standard, RC4, …) comes from `litchi_crypto::ooxml::inspect` (`litchi-crypto/src/ooxml/mod.rs:505`) when `encryption` is on, and from stream presence otherwise.
- **`Locked::unlock(self, Password) -> Result<File>`** (cfg `encryption`). It uses the leaf password openers: docx `Package::from_reader_with_password` (`package/codec.rs:441`), pptx `from_vec_with_password` (`package/encryption.rs:92`), xlsx `Workbook::from_bytes_with_password`, and the legacy `OpenOptions::with_password`.
  - The result is the eager owner, which already keeps its encryption `Mode` and rejects a plain `save` (ADR 0008:2630-2638). That is the ADR 0006 `Sensitive` state, carried inside the handle.
  - XLSB has no decrypt API, so unlocking it returns `unsupported_capability`.
- **What M0a does not claim.** It does not implement a separate `Sensitive<T>` type. It does not claim ADR 0006's final type-state facade (the ADR 0008:2630-2638 caveat still stands).

**`Signed`** is `pub struct Signed(Arc<SignedState { file: File, origin: SignatureOrigin }>)`.

- **Detection is metadata-only:**
  - OPC uses `SourceBackedPackage::has_signature_infrastructure()` (`litchi-opc/src/source_backed.rs:11531`) or `OpcPackage::is_signed()` (`package.rs:2067`).
  - CFB checks whether `_signatures`/`_xmlsignatures` streams exist, via `OleFile::exists` (`litchi-cfb/src/file.rs:3545`).
- **Nothing is verified at open.** Verification is an explicit call later (P1).
- **Behaviour.** Reads go to the inner `File` (`Signed::file(&self) -> &File`, and `File::content()` unwraps it). Edit verbs, which arrive in M2, return `document_protected` with a suggestion naming `unsign`/`resign` (ADR 0006:136-144).

**Explicit signature policy.**

- **The problem.** Record 0002 (§finding 4) found leaf commits that **silently strip signatures**, contradicting ADR 0006:
  - pptx `opened::Transaction::commit` calls `working.unsign()` (`crates/litchi-pptx/src/opened/transaction.rs:1255-1270`).
  - Several xlsb families call `unsign()` (e.g. `crates/litchi-xlsb/src/cell_values/workbook.rs:167`).
  - odf-formula commits drop signature members (`crates/litchi-odf-common/src/core/writer.rs:1618-1622`).
  - Other leaves refuse instead: docx until `unsign()`, xlsx with `Error::Signed`, and the doc/xls/ppt incremental editors.
- **`Signed` lives in `litchi-ops`, not in the leaves.** It is the single place where the unified API decides. `Signed::policy` exposes the consuming transitions:
  - `Signed::unsign(self) -> File`, which strips the signature explicitly and records it in diagnostics.
  - `Signed::resign(self, Signer)`, for P1 (cfg `sign`).
- **Signed input never reaches a leaf commit.** Every edit verb on `File::Signed` refuses with `document_protected`, so a stripping leaf commit cannot run on it through the unified API.
- **What M0a delivers** is S6b (§4): an audit test that every leaf commit reachable from a `File` variant is either refusing or not reachable on signed input. The leaf fixes that turn silent strips into refusals are filed as (b)-class requests against `litchi-pptx`, `litchi-xlsb` and `litchi-odf-common`, matching 0002's "(b) fixes".

**`Error`** is defined as:

```rust
pub struct Error(Box<Inner>);              // 8 bytes; Result<File> stays 16–24 bytes
struct Inner { code: Code, message: Cow<'static, str>, suggestion: Option<Cow<'static, str>>,
               format: Option<FileFormat>, detected: Option<FileFormat>, source: Source }
#[non_exhaustive] pub enum Code { NotFound, PathNotFound, …, InternalError, UnsupportedCapability } // §16 closed list, 37 + 1
#[non_exhaustive] pub enum Source {
    None, Io(std::io::Error), Core(litchi_core::Error), Opc(litchi_opc::OpcError),
    Docx(litchi_docx::Error), Pptx(litchi_pptx::Error), Xlsx(litchi_xlsx::Error),
    Xlsb(litchi_xlsb::PackageError), XlsbRaw(litchi_xlsb::Error),
    Doc(litchi_doc::Error), Xls(litchi_xls::Error), XlsSource(litchi_xls::SourceBackedError),
    Ppt(litchi_ppt::Error), Rtf(litchi_rtf::Error), Crypto(litchi_crypto::ooxml::Error), Cfb(litchi_cfb::OleError),
}
```

- Each variant is behind its feature gate and has a `#[from]`.
- `std::error::Error::source()` returns the leaf error, so transaction diagnostics are not lost (§16.3.2).
- **The enum anticipates M2's commit adapter.** Record 0002 found that commit shapes differ:
  - docx returns (snapshot, `Patch`, `Diagnostics`); xls `cell_values` returns the same.
  - xlsx returns (`Workbook`, `Patch`) with no `Diagnostics`.
  - xlsb returns (Snapshot, Patch) or applies in place through `&mut`.
  - doc and ppt return (snapshot, `Patch`).

  `Source` therefore reserves `#[non_exhaustive]` room for each leaf's transaction error (docx `TransactionError`, `document/transaction.rs:455`; doc `body_text::Error`, `body_text.rs:477`; ppt/xls transaction errors). `Inner` reserves an optional `diagnostics` slot that is empty in M0a. M2 then adds variants without reshaping `Error`. M0a defines no commit adapter.
- **Leaf error locations:** docx `error.rs:10`, pptx `error.rs:247`, xlsx `error.rs:36`, xlsb `host/error.rs:28` and `raw/error.rs:41`, doc `package/model.rs:448`, xls `error.rs:21`, ppt `package/model.rs:161`, rtf `codec/error.rs:63`.
- **Exhaustiveness.** `xls::Error` and `ppt::Error` are not `#[non_exhaustive]`, so their mappings are exhaustive. Every other leaf mapping needs a wildcard arm that maps to `corrupt_file`, and each such arm is listed in a unit test.

How errors map to codes in M0a (the open and read paths):

| Source condition | Code |
|---|---|
| `io::ErrorKind::NotFound` | `file_not_found` |
| other `io::Error`; `Error::SourceChanged`; OPC `SourceChanged` | `io_error` (`SourceChanged` keeps its own `message`; F10 fixes the PPTX flattening) |
| no format recognized (today `NotOfficeFile`) | `unsupported_format`, `detected: None` |
| wrong family for a view (today `UnexpectedFormat{detected}`) | `unsupported_format`, `detected: Some(..)` |
| format detected but its feature is off (`FeatureDisabled`); iWork input in M0a | `unsupported_format` + suggestion (for iWork: "use `litchi::iwork::Document::open`") |
| `ParseError`, `InvalidFormat`, `CorruptedFile`, `XmlError`, `ZipError`, `InvalidContentType`, malformed leaf payload | `corrupt_file` |
| `ResourceLimit` / OPC `ReadLimit` / leaf limit errors | `invalid_input`, with resource, observed and limit in the context (**open question Q1**: the closed list has no limit code) |
| `Unsupported` (per-format capability, e.g. `Xlsx::sheet` on another format) | `unsupported_capability` |
| `Locked` read/edit attempt; `Signed` edit attempt; `ProtectionDenied` | `document_protected` + suggestion |
| `Allocation`, `Committed`, `Other`, `Cancelled` | `internal_error` (**Q1** also covers cancellation) |

### 2.7 Where detection lives (D8)

- **Move.** `L/detection_smart/` moves to `litchi_ops::detect`. This follows ADR 0010, which lets the root coordinator compose leaf probes: ODF via `litchi_odf_common::detect` (ADR 0009), OOXML via `litchi_opc` content types, OLE via `litchi_cfb` plus stream-name classification (`ole2.rs`), and RTF via `litchi_core::detection::rtf`.
- **The iWork arms stay in the facade** (functions.rs:199, :341, :393; detected.rs:1774, :1936), because of D4 and §3.3.
  - Moving them out does not change precedence, because iWork is already the last ZIP candidate (functions.rs:193-200).
  - The facade's `litchi::detect_file_format*` and `litchi::open*` call the ops function first. They run `litchi_iwa_detect::{path, bytes}` only when ops reports an unclassified ZIP, i.e. on the failure path.
- **One container index per open.** `litchi_ops::open` replaces the three per-family path detectors (§1.4) with one combined path detector (a merge of `detect_*_source_path*`) that builds a container once:
  1. ZIP is checked before OOXML: an ordinary ODF package goes to ODF `PreparedPackage` (`litchi-odf-common/src/detect.rs:759/770`) and then to the ODT/ODS/ODP `from_prepared*` functions.
  2. Otherwise one `SourceBackedPackage` → `PreparedOoxmlSource` (detected.rs:3416-3504) → `from_source_backed_package` for docx, pptx, xlsx or xlsb.
  3. A CFB is parsed once as `OleFile`/`SharedOleFile` and handed to `from_ole_file`/`from_shared_if_powerpoint` (`litchi-doc/src/package/codec.rs:134`, `litchi-ppt/src/package/source.rs:105`, `litchi-xls/src/workbook/package.rs:103`). Encrypted and signed classification read the same handle.
- **Kept unchanged:** every polyglot rule in §1.4 and every TOCTOU guard. The guards become methods on the source-state owners, which removes the duplicated helpers in doc.rs and prs.rs.
- **Prepared-source entry.** `from_prepared(PreparedOoxmlSource)` survives as `litchi_ops::open_prepared(PreparedDetection)` for callers that ran `detect_and_prepare` themselves.

---

## 3. Crate placement and dependency edges

### 3.1 Edges

```
litchi-core, litchi-opc, litchi-cfb, litchi-ole-common, litchi-crypto?, litchi-sign?,
litchi-ooxml-common, litchi-odf-common, litchi-sheet,
litchi-{docx,pptx,xlsx,xlsb,doc,xls,ppt,rtf,odt,ods,odp}   (all optional, feature-gated)
        ↑
    litchi-ops         (no iWork crate; no `litchi`)
        ↑ optional, feature `dep:litchi-ops`, enabled by every format feature, reachable from `all`
    litchi             (re-exports `open*`, `File`, `OpenOptions`, `Error`, `Result`, `Locked`, `Signed`, `Family`, detection;
                        keeps leaf re-export modules, `sheet::text`, `iwork`, `markdown`, `tolerance`)
        ↑
    litchi-py (pyo3-litchi), perf-baseline, examples, tests
```

- **`litchi` may not depend on `litchi-ops` unconditionally.** The facade audit allows only `litchi-core` as a non-optional normal dependency (`tools/check_crate_boundaries.py:99`, `:19612-19617`), and every optional dependency must be owned by a `dep:` feature and reachable from `all` (:19665-19691). So each format feature in `crates/litchi/Cargo.toml` adds `"dep:litchi-ops", "litchi-ops/<fmt>"`.
- **No cycles.** `litchi-ops` must not depend on `litchi` (cycle check :19504-19506).
- **pyo3-litchi keeps depending on `litchi`.** Design §7 wants it switched to `litchi-ops`, but it still needs `litchi::iwork` for the Python `Document`/`Presentation`/`Workbook` on Pages, Keynote and Numbers (§5.1). Switching directly to `litchi-ops` moves to M6 (finding F5).

### 3.2 `tools/crate_boundaries.json`

- Add a `"litchi-ops": [sorted deps]` entry to `packages`, between `"litchi-opc"` (line 230) and `"litchi-oth"` (236). Optional dependencies count (:19373-19383).
- Add `"litchi-ops"` to the `litchi` list (lines 4-37).
- Add `litchi-ops` to `runtime_neutral` (486-533). It must never depend on rayon or tokio.
- **iWork edges.** No rule in the boundary checker forbids a non-facade crate from depending on an iWork crate. The only iWork-related ban is on the migration host `litchi-iwa` (:19230-19236). So `crate_boundaries.json` alone **would allow** `litchi-ops → litchi-{pages,keynote,numbers,iwa-detect,iwa-structured}`.
- The workspace uses `members = ["crates/*"]` (Cargo.toml:3), so no member edit is needed. Add the `[workspace.dependencies]` row between :141 (`litchi-opc`) and :142 (`litchi-oth`). `cargo sort --workspace --check` runs in CI (`rust-ci.yml:282-283`).

### 3.3 `tools/non_iwork_gate.py` — the binding constraint

- **The rule.** Any package that is not iWork and not an unsafe-facade dependent is "bulk" (:1240-1250). `_require_bulk_dependency_boundary` (:1137-1160) rejects **any** direct edge from a bulk package to an excluded package, whether or not the edge is optional. The bulk `--all-features` tree may not contain iWork packages or `litchi` (:1936-1951).
- **M0a passes under D4.** Changes needed:
  - `EXPECTED_BULK_PACKAGES` 46 → 47 (:71).
  - Add `litchi-ops` to `EXPECTED_BULK_PACKAGE_NAMES` (:119-168) and `EXPECTED_WORKSPACE_PACKAGE_NAMES` (:169-240; 65 → 66).
  - Update the fixtures in `tools/test_non_iwork_gate.py` (counts at :382-389, :616, :712-713; name sets at :21-115).
  - `EXPECTED_EXCLUDED_PACKAGES` (18) is unchanged.
- **M6 must redesign the gate before iWork variants join `File`** (finding F3). Two options:
  - (a) Make `litchi-ops` a second gated facade with a safe/unsafe feature partition, as `litchi` has (:1190-1210).
  - (b) Keep `litchi-ops` iWork-free and add the iWork variants in a separate excluded crate, for example `litchi-ops-iwork`, that `litchi` composes. Option (b) cannot add variants to a `File` defined in `litchi-ops`, so in practice it means (a) or moving the `File` definition.
- This is why design §16.5's "46 → 47" holds only while `litchi-ops` stays iWork-free.

### 3.4 ADR 0024 inventory and other gates

- **ADR 0024** has no package table; it is prose plus dated count amendments. The latest is "2026-09-24 amendment" at `docs/adr/0024-current-topology.md:4697`, with counts at :4709-4712 ("65 workspace packages, 244 internal dependency declarations, 233 canonical edges, 12 development-only edges, 11 ordered migration debts …"). Registration adds a new `## <date> amendment: litchi-ops registration` with the new package, its edges and the recounted totals (66 packages; declarations and edges recounted from `cargo metadata`).
- **`tools/check_iwork_public_api.py`** is unaffected because `litchi::iwork` does not move. It rejects any `litchi_ops` type reachable from `litchi::iwork` (:48), so the iwork module must not start naming ops types. No CI workflow runs this script today.
- **`-D deprecated`** (`rust-ci.yml:288-291`, `perf-baseline.yml:189`): old APIs are deleted, never deprecated.
- These registrations are record **0004**'s scaffold work. This plan assumes 0004 lands first, as step S0.

---

## 4. Step sequence (ADR 0008: every step builds and is green)

Steps land on an M0a integration branch. The merge into `feat/unified-ops-api` is **one merge**, so trunk sees one switch (§15, §16.5) while each internal step stays buildable. The old types live alongside the new ones until S9 deletes them. They are never deprecated (`-D deprecated`).

| Step | Content | Depends on | Parallel with | Owner size |
|---|---|---|---|---|
| **S0** | Record 0004: `litchi-ops` scaffold, features (`doc docx ppt pptx xls xlsx xlsb rtf odt ods odp encryption sign markdown`; no iWork), gate registration (§3.2–3.4), the facade's optional edge. | — | M0b | (0004) |
| **S1** | **Baseline perf capture** on the pre-M0a commit for every route in §5.3. Store machine-readable results under `docs/unified-ops/results/0003/before/`. | — | S2–S4 | 1 |
| **S2** | `litchi_ops::error`: `Code` (closed list + `unsupported_capability`), `Error`, `Source` + `#[from]`s, per-leaf classifiers, `Display`, suggestions. Unit tests: one per row of the §2.6 table, plus one per wildcard arm, plus a size assertion. | S0 | S1, S3, S4 | 1 |
| **S3** | `OpenOptions`, `SourceMode`, `Family`, `Locked`/`Signed` shells (types, `format()`, `Debug` redaction; no unlock yet). | S0 | S1, S2, S4 | 1 |
| **S4** | Move `detection_smart` into `litchi_ops::detect` without the iWork arms. The facade re-exports the detection API and adds the iWork fallback wrapper. Facade internals call ops detection through a `#[doc(hidden)] pub mod __migration`, which S9 deletes. Carries the 45 in-source detection tests and `tests/catalog_*`. Adds a test that iWork is still classified last. | S0 | S1–S3 | 1 |
| **S5a** | `Docx` handle: two-state owner, sidecars, `finish_source` guard as an owner method, `section_inventory`, typed errors (deletes 35 `map_ooxml_error` uses on the document side as code moves). | S2, S4 | S5b–e | 1 |
| **S5b** | `Pptx` handle: two-state owner, lazy metadata, `slide_catalog`/`slide_size`, typed `SourceChanged` everywhere (F10), L-PPTX-1 in `litchi-pptx`. | S2, S4 | S5a, c–e | 1 |
| **S5c** | `Xlsx` + `Xlsb` handles: relocate `adapters/`, `selection.rs` (case-fold semantics kept verbatim), typed errors in place of `Box<dyn Error>`, L-XLSX-1 and L-XLSB-1. | S2, S4 | S5a, b, d, e | 1 |
| **S5d** | `Doc`, `Xls`, `Ppt` handles with retained owners (§2.3); collapse `XlsFile`/`XlsMem`/`Other`; L-DOC-1, L-XLS-1, L-PPT-1. | S2, S4 | S5a–c, e | 1 |
| **S5e** | Read-only `Rtf`, `Odt`, `Ods`, `Odp` handles: carry the existing payloads over, replace `ParseError(format!)` flattening with typed `Source::Core`, make `Ods` `Sync` (drop the `RefCell`). | S2, S4 | S5a–d | 1 |
| **S6** | `File` enum, family views, `compat::{word, slide}` relocation, single-pass `open/open_bytes/open_with/open_prepared`, `Locked`/`Signed` classification and `unlock`, size/Send/Sync assertions, `impl WorkbookTrait for WorkbookView`. The in-source tests from doc.rs, prs.rs and workbook.rs are ported here as `File` tests. | S3, S5* | — | 1 (2 if `Locked`/`Signed` are split off) |
| **S6b** | Signature policy (§2.6): `Signed::unsign`, edit-verb refusal on `Signed`, and the audit test over every leaf commit reachable from a `File` variant. File the (b) leaf fixes for the silent strippers (pptx `opened::Transaction`, xlsb families, odf-formula). | S6 | S7 | 1 |
| **S7** | Facade: `litchi::{open, open_bytes, open_with, File, OpenOptions, Error, Result, Locked, Signed, Family}` re-exports; iWork-aware `open` wrapper; `impl ToMarkdown for File` (markdown switches to `compat`); `tolerance.rs` fed from handles. The old types still exist. | S6 | — | 1 |
| **S8a** | `crates/pyo3-litchi` (§5.1). | S7 | S8b–d | 1 |
| **S8b** | `tools/perf-baseline`: add unified routes (§5.3), keep old routes until S9, update Python policy tests. | S7 | S8a, c, d | 1 |
| **S8c** | `crates/litchi/tests/*` and `fuzz` (§4.4 table). | S7 | S8a, b, d | 1 |
| **S8d** | Examples, `README.md`, `docs/OFFICE_API_GUIDE.md`, `L/lib.rs` crate docs, `litchi-markdown` doctests. | S7 | S8a–c | 1 |
| **S9** | Delete `L/document/`, `L/presentation/`, `L/sheet/{workbook.rs, workbook_types.rs, adapters/, selection.rs}`, the `Box<dyn WorkbookTrait>` openers in `L/sheet/functions.rs` (all 42), `map_ooxml_error` (`L/lib.rs:427-433`), `__migration`, and the old perf routes from the harness (their recorded results stay). Keep `sheet::text` (concrete `TextWorkbook`/`formats::read_*`) and `iwork`. Run the full gate list from the goal. | S8* | — | 1 |
| **S10** | After-measurement: ABBA runs alternating the pre-M0a and post-S9 binaries on the same corpus and machine. Results go to `docs/unified-ops/results/0003/after/`, with a record. Any regression above 5% blocks the merge. | S9 | — | 1 |

Ordering under the 2026-10-04 tier priority: S5a–c (tier 1) and S5d (tier 2) are the critical path. S5e (tier 3) is a mechanical carry-over and may land last. If it slips, it can temporarily keep the tier-3 formats on the old `Document`/`Presentation`/`Workbook` until S9, but S9 cannot run until S5e lands, because the old types are deleted there.

Critical path: S0 → S4 → S5 (longest is S5c or S5d) → S6 → S6b → S7 → S8 → S9 → S10. Up to 9 implementers can work at once: S1–S4, then S5a–e, then S8a–d.

### 4.4 Tests whose semantics carry over verbatim

| Category | Test (file:line) | New form |
|---|---|---|
| Wrong-family refusals | `unexpected_format.rs`: `an_xlsx_given_to_the_document_opener_reports_xlsx` :44, `a_pptx_…document…` :53, `a_legacy_ole2_presentation_…document…_reports_ppt` :62, `an_odf_presentation_…document…` :71, `a_non_office_file_given_to_the_document_opener…` :80, `a_docx_given_to_the_presentation_opener_reports_docx` :91, `a_legacy_ole2_document_…presentation…` :100, `an_odf_document_…presentation…` :109, `a_non_office_file_…presentation…` :118, `a_docx_given_to_the_workbook_opener_reports_docx` :129, `a_pptx_…workbook…` :138, `a_legacy_ole2_document_…workbook…` :147, `an_odf_document_…workbook…` :156, `an_odf_presentation_…workbook…` :165, `a_non_office_file_…workbook…` :174, `the_workbook_opener_separates_every_wrong_format_it_can_name` :196; `catalog_detection_arbitration.rs::wrong_family_ooxml_path_reports_the_detected_family` :385 | `open(p)?.document()` (or `.presentation()`, `.workbook()`) → `Code::UnsupportedFormat`, `detected == Some(<fmt>)`. Non-office input: `open(p)` → `UnsupportedFormat`, `detected == None`. |
| Read limits | `read_limits_facade.rs`: `docx_limits_are_nameable` :6, `document_limits_bound_path_and_byte_ingress` :15, `document_limits_reject_invalid_suffix_and_unknown_extension_ooxml` :35, `pptx_limits_are_nameable` :63, `presentation_limits_bound_path_and_byte_ingress` :72, `presentation_limits_reject_invalid_suffix_and_unknown_extension_ooxml` :92, `xlsx_limits_and_bounded_sheet_helpers_are_nameable` :121, `xlsb_limits_and_bounded_sheet_helpers_are_nameable` :129 | `OpenOptions::default().ooxml_limits(..)` with `open_with`/`open_bytes_with`. The same limit, observed value and resource are asserted through `Error::source()`. |
| Selector case-insensitivity | `unified_sheet_selectors.rs`: `public_selector_preserves_exact_owned_cell_semantics` :63 (`.sheet("dAtA")` at :69), `selected_access_defers_unselected_malformed_worksheet` :139, `non_grid_sheet_returns_typed_not_worksheet_error` :157, `non_xlsx_workbook_returns_explicit_unsupported_error` :169, `selected_handle_preserves_source_change_error` :185 | `File::Xlsx(x) => x.sheet("dAtA")`. On a non-XLSX workbook → `unsupported_capability`. A source change → `io_error` with typed `Source`. |
| Detection / arbitration | `catalog_first_detection.rs` :256, :310, :390, :421; `catalog_detection_arbitration.rs` :93, :163, :198, :230, :290, :309, :326, :353 | Unchanged through the facade's detection re-exports. Opens go through `open*`. |
| Tolerance reports | `tolerance_report_facade.rs` :132, :155, :195, :227, :252 (Pages is untracked, so it moves to `iwork`), :263, :276, :323, :352, :380, :440, :481 | `file.tolerance_report()` |
| Markdown | `markdown_docx_lists.rs` (2), `markdown_docx_semantics.rs` (10), `markdown_real_fixtures.rs` (1) | `open(p)?.to_markdown()` |
| xlsb facade | `xlsb_facade.rs` (23 tests; 22 calls to `sheet::open_xlsb_*`, 49 `WorkbookTrait` mentions) | `open(p)?.workbook()?` (static `WorkbookTrait`), or the leaf `litchi::xlsb` where the test is about the leaf |
| Pages via `Document` | `pages_facade.rs:52` | `litchi::iwork::Document::open`, plus a new assertion that `litchi::open` on a `.pages` file returns `unsupported_format`, `detected == Some(Pages)`, with a suggestion |
| New in M0a | — | encrypted OOXML → `File::Locked` (and `unlock` round trip under `encryption`); signed OPC/CFB → `File::Signed`, with reads delegated; `size_of::<File>() == 16`; Send/Sync/Clone; one-container-index-per-open counter test (extends `record_opc_probe`, functions.rs:186) |

Tests that do not change: `encryption_facade.rs`, `vba_inspection_facade.rs`, `semantic_*`, `numbers_*`, `keynote_facade.rs`, every `iwork_*` test, and the fuzz targets (only `fuzz_targets/parse_iwork.rs:6,19` uses the facade, through `litchi::iwork`).

---

## 5. Breakage per consumer, and perf routes

### 5.1 `crates/pyo3-litchi` (48 `litchi::` references; 0 Python tests exist)

- **References by file:**
  - `common.rs`: 36. These are `litchi::Error` ×5 (:18-23, `to_py_err`), `FileFormat` ×23 (:81-104), `detect_file_format` :275, `detect_file_format_from_bytes` :290; the rest are `RGBColor`/`Length`.
  - `document.rs`: 7, namely `Document` :44/:65 and `document::{Paragraph, Run, Table, Row, Cell}` :113/156/213/249/277.
  - `presentation.rs`: 3 (:43, :64, :109).
  - `sheet.rs`: 2 (:39, :60).
- **New error shape.** Python `Document`, `Presentation` and `Workbook` each wrap a pyo3-internal enum `{ File(litchi::File), Iwork(litchi::iwork::Document) }`, so `.pages`, `.key` and `.numbers` keep opening (R6). `to_py_err` switches to `Code`. `boxed_err_to_py_err` is deleted.
- **Fix existing drift in the same step.** The `From` impl sends `Odb` to `FileFormat::Doc` (common.rs:104-109), and the `.pyi` stub lists only 14 of the 21 variants (`python/litchi_py/__init__.pyi:17-30`).
- **A minimal pytest smoke suite is added,** because none exists. It covers open per format, text, and a wrong-format error.

### 5.2 Other consumers

| Consumer | Count | Sites |
|---|---|---|
| `tools/perf-baseline/src/lib.rs` | 43 lines (Document 16, Presentation 10, Workbook 10, detect 6, `document::` 1) | e.g. `Workbook::from_bytes` :22433/22506/42070/42199/43234; `Document` :26011–:38971; `Presentation` :38276–:38528; detection :21481, :45937–:45970, :61134/:61141 |
| `tools/perf-baseline/src/filesystem.rs` | 30 (+2 `detection_smart`, :5944/:5950) | `Presentation` :2844–:7368; `Document` :5939–:7156; `Workbook` :6182–:8565; `sheet::Selected*` :6199–:6788 |
| `tools/perf-baseline/src/facade_ole2.rs` | 24 | `Document::open` :313/:632/:643/:652; `Presentation::open` :439/:663/:674/:683; route names :85/:86 (unit test :714-715) |
| perf-baseline bins | 7 | `xlsb_crud.rs` :9, :10, :911, :916, :1378; `xls_source_attribution.rs:1298`; `docx_source_selected_paragraph.rs:1482` |
| perf-baseline docs and policy | README 13 lines; `tools/test_perf_baseline_source_policy.py:394-404` (and :46, :344-353) | route-selector assertions |
| `crates/litchi/tests` | 18 files, about 200 call lines | see §4.4 |
| `crates/litchi/examples` | 15 files: `to_markdown.rs` (:39, 290, 295, 308, 323, 344, 354), `profile_all_formats.rs` (5), `profile_test1.rs`, `xlsb_reader.rs`, 8 × `sheet_eval_*`, `read_xlsb.rs`/`read_shapes_all_formats.rs` (`WorkbookTrait` only); 6 ODF examples return `litchi::Result` (they keep compiling through `From<litchi_core::Error> for Error`) | — |
| `README.md` | about 16 sites | :59, 62, 65, 69, 118, 126, 133, 134, 141, 144, 229–230 |
| `docs/OFFICE_API_GUIDE.md` | "Unified API" section | :11, :637–:672 (:642, 645, 649, 652, 667, 670) |
| `crates/litchi-markdown/src/traits.rs` | 3 doctests (`litchi::Error`) | :18, :45, :71 |
| `L/lib.rs` | 17 doc lines plus the root re-exports | §1.3 |

### 5.3 Perf-baseline routes to re-measure (S1 before, S10 after)

The new selectors use the `unified_` infix, for example `docx_file_unified_full_text`. Frozen records are never rewritten.

| Scenario | Existing selectors (lib.rs definition) | Runner |
|---|---|---|
| Open | `docx_file_source_open` and `…_eager_open` (:1749-1766); `pptx_file_{eager,source}_open` (:1729-1748); `xlsx_file_open[_lifecycle]` (:2084-2091); `xlsx_bytes_open[_lifecycle]` (:2095-2096); `doc_facade_file_open`, `ppt_facade_file_open` (:2041-2046, opt-in); `rtf_file_open[_lifecycle]` (:2215-2216); `odt_file_{eager,source}_open` (:1767-1770); `odp_file_*_open` (:2353-2356); `ods_file_*_open` (:2362-2367) | `run_docx_operation` filesystem.rs:6009, `run_pptx_operation` :5856, `run_xlsx_operation` :6281, `run_xlsx_bytes_root_access` lib.rs:43189, `facade_ole2::run_case` lib.rs:12537, `run_rtf_root_access` :26040, `run_odt_file_root` :38854, `run_odp_root_file_access` :38361, `run_ods_root_file_access` :42054 |
| List | `docx_file_*_list_paragraphs`, `…_paragraph_count[_lifecycle]`; `pptx_file_*_list_slides`, `…_slide_count[_lifecycle]` | same |
| Read one | `docx_source_selected_paragraph` (bin); `pptx_file_*_selected_slide[_lifecycle]`; `xlsx_file_selected_cell`; `doc_facade_file_one_paragraph`; `ppt_facade_file_one_slide_text`; `odp_file_*_selected_slide`; `ods_file_*_selected_{cell,media}` | same |
| Full text | `docx_file_*_full_text[_lifecycle]`; `doc_facade_file_full_text`; `ppt_facade_file_full_text`; `odt_file_*_open_full_text_lifecycle` | same |
| Detection | `detect_{odt,ods,odp}_{bytes,reader,polyglot,catalog_alias}` (:2283-2286, :2324-2327, :2349-2352) | `run_detection_case` :45927 |

- **Leaf leg.** The eager docx path already measures the leaf directly through `detection_smart` + `litchi_docx` (filesystem.rs:6009). S8b adds a matching leaf leg wherever a route has none, for the goal's paired leaf-versus-unified rule.
- **Merge rule.** Any regression above 5% on these scenarios blocks the merge (goal "PERFORMANCE BAR"). Open overhead must stay within 5% or 20 µs, whichever is larger.
- **Expected change.** The single combined path detector replaces the per-family detectors, so the main expected change is on cross-family inputs. Same-family opens should do the same container work as today. This is a hypothesis to measure, not a claim.

---

## 6. Risks, findings and open questions

### Risks

- **R1 (gate).** iWork cannot join `File` without redesigning `tools/non_iwork_gate.py` (§3.3). Plan this in M6, not M0a.
- **R2 (leaf requests).** L-PPTX-1, L-XLSX-1, L-XLSB-1, L-DOC-1, L-XLS-1 and L-PPT-1 are leaf work inside S5. If one is rejected or slips, its fallback (build the editor lazily from the retained source at the first edit) still meets "retain the editable owner", but costs one re-index at the first edit. That must be measured and recorded.
- **R3 (Send/Sync).** There is no existing proof for `pptx::SourceBackedPresentation`, `xlsx/xlsb::SourceBackedWorkbook`, `ppt::SourceBackedPackage` or the doc/ppt snapshots. S6's assertions may show a `!Sync` field, which would then need fixing in the leaf. The facade `Workbook` is already `!Sync` under `ods`.
- **R4 (consuming publishers).** Several leaf publishers take `self` (§2.2). `Arc` sharing forces M2 to use `try_unwrap`, with a fallback, or `&self` publishers.
- **R5 (memory).** Keeping the legacy byte-route input (`Doc`, `Ppt`) raises peak RSS by up to the input size on those routes. S10 measures RSS as well as time.
- **R6 (Python parity on iWork).** Python `Document`/`Presentation`/`Workbook` on iWork inputs move from the facade's `Pages`/`Keynote`/`Numbers` payloads to `litchi::iwork::Document`. `paragraphs()` on Pages previously gave `Paragraph::Pages(String)` per paragraph; mapping that from `iwork` sections needs a parity check. Python has no tests today, which raises this risk.
- **R7 (behaviour changes that are not semantic regressions).**
  - Encrypted OOXML now opens as `Locked` instead of failing as `NotOfficeFile`.
  - Signed packages now open as `Signed`, so code matching `File::Docx` must call `content()`.
  - PPTX source changes now report `io_error` with a typed source instead of `InvalidFormat(String)`.

  Each change is listed in the S10 record.
- **R8 (cfg matrix).** There are 14 format features × target gates. The possible `odp` without `pptx`, non-unix/windows compile gap (§1.7) must be checked with `cargo check -p litchi-ops --no-default-features --features odp --target <non-unix>`, or recorded as unsupported.
- **R9 (size of S4/S6).** `detected.rs` is 4455 lines. S4 is a pure move with the iWork arms split out. Keeping it pure makes review tractable.

- **R10 (M2 prerequisite, noted only).** DOC, XLS and PPT save is bytes-only today. Atomic sibling-temp save exists only on narrow source-backed commits (0002 §finding 6). ODF packaged saves use `fs::write`. M2 therefore needs a shared facade atomic writer in `litchi-ops`: sibling temp, fsync per ADR 0005 durability, then rename, over `publish_*_to_stream` / `bytes()`. M0a does not build it.

### Findings (disagreements with the design, recorded per the goal)

- **F1.** §16.2 says `PptxSource` and the xlsx `Source` state are already mutable owners. Both are read-only (`litchi-pptx/src/presentation/source.rs:408`, `litchi-xlsx/src/workbook/source.rs:117`). XLSB has no source-backed edit API.
- **F2.** §4 lists `File::{Document, Presentation, Workbook}`. This record recommends a flat per-format enum (§2.1).
- **F3.** §16.5 "`EXPECTED_BULK_PACKAGES` 46 → 47" holds only while `litchi-ops` has no iWork dependency.
- **F5.** §7 "litchi-py switches its dependency to `litchi-ops`" is deferred to M6, because Python still needs `litchi::iwork`.
- **F8.** No facade test covers encrypted input. Per the code, encrypted OOXML surfaces as `NotOfficeFile` (inferred, not run).
- **F10.** PPTX source-change errors are flattened on most read paths (prs.rs; §1.5).
- **F12.** `WorkbookImpl::XlsFile` and `Other` are dead code.
- **F14.** The Python `FileFormat` has drifted from the core enum (§5.1).
- **F15.** Sheet-name matching is case-insensitive for XLSX and exact for XLSB and text. M0a keeps this as is; unifying it is an M1 or registry decision.
- **Size note on §16.5.** "~30 shallow delegation sites" in pyo3: the actual count is 48 `litchi::` references, about 45 of them on facade types.

### Open questions (provisional defaults adopted; listed for owner confirmation)

- **Q1.** The closed code list has no resource-limit or cancellation code. Default: `invalid_input` and `internal_error`, with structured context. Recommendation: add `limit_exceeded` and `cancelled` as Litchi extensions, alongside `unsupported_capability`.
- **Q2.** Should iWork inputs to `litchi::open` return `unsupported_format` with a suggestion (default), or should M0a add read-only iWork variants despite R1? The default follows the coordinator's direction and §16.4.
- **Q3.** Should `retain_source` default to `true` (default, per the owner override) given R5?
- **Q4.** Should `Signed` wrap every signed input (default; ADR 0006 "Signed input is `Signed<T>`"), or only those whose signature covers the parts an edit would touch? The default is the simple and safe reading.
- **Q5.** Merge granularity: one merge commit of the step series (default) versus squash.
