# litchi-ole — workspace-split parity audit

Branch: `refactor/workspace-split` vs `main`.
Crate: `crates/litchi-ole/`. Old location: `src/ole/`.

Three commits touched the crate:

| Commit | Subject |
|--------|---------|
| `145026a` | P4c: carve out litchi-ole crate (legacy OLE2 binary formats) |
| `3c1d141` | Audit stale comments and fix doctest paths post-workspace-split |
| `71158bd` | Add per-crate README files |

## 1. Scope

Files reviewed under the new tree (`crates/litchi-ole/`):

- `Cargo.toml`, `src/lib.rs`, `src/consts.rs`, `src/sprm.rs`,
  `src/sprm_operations.rs`, `src/plcf.rs`, `src/mtef_extractor.rs`,
  `src/extractor.rs`
- `src/escher/{mod,container,parser,properties,record,shape,shape_factory,
  text,types,writer}.rs`
- `src/doc/{mod,document,package,paragraph,header_footer,footnote,hyperlink,
  image,shapes,table}.rs`
- `src/doc/parts/{mod,fib,piece_table,chp,chp_bin_table,fkp,pap,tap,
  tap_parser,headers,footnotes,hyperlinks,images,fields,numbering,text,
  paragraph_extractor}.rs`
- `src/doc/writer/{mod,core,fib,fkp,bin_table,dop,font_table,footnotes,
  headers,hyperlinks,images,numbering,ole_metadata,piece_table,section,
  sprm,stylesheet,tap}.rs`
- `src/xls/{mod,error,utils,records,workbook,worksheet,cell,autofilter,
  comments,hyperlinks,merged_cells,pivot_table,protection,shapes,writer/*}.rs`
- `src/ppt/{mod,package,presentation,current_user,escher_textbox,text_prop,
  text_run}.rs`
- `src/ppt/{animation,escher,parsers,persist,records,shapes,slide,text,
  transition,writer}/**`

Compared against the corresponding `main:src/ole/**` paths.

## 2. Method

- `git show 145026a --stat | grep litchi-ole` to enumerate touched files
  and the magnitude (`{0,2,4,6,8,10,...}` line deltas) of each rename.
- For every file with a non-zero delta in the stat, `diff` of the
  CRLF-stripped old-vs-new contents (most files on `main` had CRLF
  terminators; the carve-out re-saved them with LF, which inflates raw
  diff counts).
- `git show 3c1d141 -- crates/litchi-ole/` for the comment + doctest
  cleanup.
- Spot-check of:
  - `consts.rs` (CFB/PPT/Escher constants)
  - `sprm.rs` / `sprm_operations.rs` / `plcf.rs` (FIB/PIECE/SPRM core)
  - `doc/document.rs::elements()` (the only logic seam where a
    crate-internal type was renamed)
  - `doc/parts/{fib,piece_table}.rs` (FIB + PIECE-table walking)
  - `doc/image.rs` (PICF / BLIP extraction, full content diff)
  - `doc/writer/core.rs` (FIB/FKP/PAPX/CHPX writer pipeline)
  - `xls/{records,workbook,worksheet,cell,utils}.rs` (BIFF dispatch +
    cell-value decoding + UTF-16 string parsing)
  - `ppt/{presentation,parsers/parser,records/record,persist/mapping,
    shapes/{shape,picture}}.rs` (record dispatch, persist directory,
    Escher-derived shape types)
  - `ppt/escher/mod.rs` (re-export surface to PPT)
  - `extractor.rs` (image extraction, relocated from `src/images/`)

## 3. Findings

### Top-level (`consts.rs`, `sprm*.rs`, `plcf.rs`, `mtef_extractor.rs`, `lib.rs`, `extractor.rs`)

✅ **No logic change detected.**

- `consts.rs`: CFB primitives (`MAGIC`, `MINIMAL_OLEFILE_SIZE`,
  `DIRENTRY_SIZE`, `SECTOR_SIZE_V{3,4}`, `MAXREGSECT`, `DIFSECT`,
  `FATSECT`, `ENDOFCHAIN`, `FREESECT`, `MAXREGSID`, `NOSTREAM`,
  `STGTY_*`, `UNKNOWN_SIZE`, `VT_*`) were moved to `litchi-cfb` and
  re-exported via `pub use litchi_cfb::consts::*`. Verified against
  `crates/litchi-cfb/src/consts.rs` — all values preserved verbatim.
  PPT-specific (`PptRecordType`, `EscherRecordType`, `EscherShapeType`,
  `WORD_CLSID`, `ESCHER_*`) constants stay in `litchi-ole` byte-for-byte
  identical.
- `sprm.rs`, `sprm_operations.rs`, `plcf.rs`, `mtef_extractor.rs`:
  R100 renames; binary-identical content.
- `lib.rs`: only adds a `pub mod extractor` block (relocation of
  `src/images/extractor.rs`); the rest is the same module declarations
  the old `src/ole/mod.rs` carried.
- `extractor.rs`: relocated from `src/images/extractor.rs`. Diff is
  three import-path swaps:
  `crate::common::error::Result` → `litchi_core::error::Result`,
  `crate::images::{Blip,BlipStore,BlipStoreEntry}` →
  `litchi_imgconv::{Blip,BlipStore,BlipStoreEntry}`,
  `crate::ole::{escher,ppt::escher}::*` → `crate::{escher,ppt::escher}::*`.
  All function bodies identical.

### `doc/`

✅ **No logic change detected** for `doc/{document,package,paragraph,
header_footer,footnote,hyperlink,image,shapes,table}.rs` and the entire
`doc/parts/*` and `doc/writer/*` trees.

Notable items I scrutinised:

- `doc/mod.rs` adds a crate-local enum at the bottom:

  ```rust
  pub enum DocElement {
      Paragraph(Box<Paragraph>),
      Table(Box<Table>),
  }
  ```

  This replaces the umbrella `crate::document::DocumentElement` that
  `Document::elements()` previously instantiated directly. The umbrella
  now performs the wrap at the seam in
  `crates/litchi/src/document/doc.rs` (lines ~465-475). The payload
  (`Paragraph`, `Table`, `rows`, `para.clone()`) is identical on both
  sides; only the variant constructor name differs. **No behavioural
  change.**

- `doc/document.rs` `elements()` body — same control flow, same
  `extract_rows_from_table_paragraphs(&table_paras, 1)?` call, same
  `if !rows.is_empty()` and `else if !props.in_table` branches.
  Only `DocumentElement::Table(Box::new(crate::document::Table::Doc(
  Table::new(rows))))` is now `DocElement::Table(Box::new(Table::new(rows)))`,
  matching the rename above.

- `doc/image.rs` (largest non-trivial file): real content diff is exactly
  three lines —
  `ExtractImageFailed(crate::Error)` → `litchi_core::Error`,
  `crate::images::ExtractedImage` → `crate::extractor::ExtractedImage`,
  and a use-statement reorder. All struct layouts (`PictureFields`
  `#[repr(C, packed)]`, `BlockType` `#[repr(u8)]`) preserved.
  `BLOCK_TYPE_OFFSET = 0xE` and `MM_MODE_TYPE_OFFSET = 0x6` unchanged;
  the discriminant table for `BlockType::TryFrom<u8>` is identical.

- `doc/writer/core.rs`: every `crate::ole::doc::writer::*::*` call site
  becomes `crate::doc::writer::*::*`. No reordering of writer phases,
  no removed validation, no changed offsets.

- `doc/parts/{fib,piece_table,chp_bin_table,fkp}.rs`: R100 (verbatim)
  except for the trivial `crate::ole::plcf::PlcfParser` →
  `crate::plcf::PlcfParser` swap.

- `doc/package.rs`: gained an `impl From<DocError> for litchi_core::Error`
  block (previously in `src/common/error/conversions.rs`). Match arms
  preserved 1-for-1: `Io→Io`, `Ole→from(ole_err)`, `InvalidFormat`,
  `StreamNotFound→ComponentNotFound`, `Corrupted→CorruptedFile`.

### `xls/`

✅ **No logic change detected** across all 14 files plus `writer/`.

- `mod.rs`, `error.rs`, `records.rs`, `workbook.rs`, `worksheet.rs`,
  `cell.rs`, `pivot_table.rs`, `autofilter.rs`, `comments.rs`,
  `hyperlinks.rs`, `merged_cells.rs`, `protection.rs`, `shapes.rs`:
  diffs are exclusively `use crate::ole::xls::…` → `use crate::xls::…`
  and `use crate::sheet::…` → `use litchi_core::sheet::…` (the latter
  was promoted to `litchi-core` in P2b).
- `utils.rs`: only changes are `is_multiple_of(2)` → `% 2 != 0` at two
  call sites (UTF-16 length validation in `parse_unicode_*`). Semantically
  identical for `len ≥ 0`. Documented in commit `145026a` as MSRV
  workaround (`is_multiple_of` is stable from 1.87; workspace MSRV is
  1.85).
- BIFF record dispatch in `records.rs::CellRecord::parse_record`,
  `XlsEncoding`, `FormulaValue`: untouched (R100). The `WorkbookTrait`
  impl on `XlsWorkbook` simply re-binds to the new
  `litchi_core::sheet::WorkbookTrait` trait path; method bodies
  identical.

### `ppt/`

✅ **No logic change detected** across `presentation.rs`,
`parsers/parser.rs`, `persist/mapping.rs`, `persist/ptr_holder.rs`,
`records/{record,slide_atoms_set,document_info,slide_info}.rs`,
`shapes/{shape,picture,placeholder,autoshape,escher,geometry,textbox,
shape_enum,mod}.rs`, `slide/{factory,types,mod}.rs`, `text/extractor.rs`,
`animation/{parser,types,writer,…}.rs`, `escher/mod.rs`,
`escher_textbox.rs`, `current_user.rs`, `package.rs`, `mod.rs`.

Notable items I scrutinised:

- `ppt/presentation.rs`: visibility upgrade
  `pub(crate) fn extract_text_fast` → `pub fn extract_text_fast`. Called
  by the umbrella `crates/litchi/src/presentation/prs.rs` across the
  crate seam. **Allowed by the audit policy.** Body identical.
  Imports adjusted: `crate::images::{BlipStore,ExtractedImage,
  ImageExtractor}` is split — `BlipStore` now from `litchi_imgconv`
  (gated on `imgconv` feature), `{ExtractedImage,ImageExtractor}` from
  `crate::extractor`.
- `ppt/shapes/shape.rs` (similarity 77 % in carve commit): the entire
  delta is the per-arm `crate::ole::consts::EscherShapeType::X` →
  `crate::consts::EscherShapeType::X` rewrite for ~50 match arms in
  `From<EscherShapeType> for ShapeType`. Mapping preserved exactly.
- `ppt/shapes/picture.rs`: import path swaps + a single 3-line
  `EscherProperty::parse_properties(child.data, prop_count)` reflow that
  rustfmt now keeps on one line. No semantic change.
- `ppt/escher/mod.rs`: just `crate::ole::escher::*` → `crate::escher::*`
  in the re-export list.
- `ppt/package.rs`: gained `impl From<PptError> for litchi_core::Error`
  with the same five-arm match
  (`Io,Ole,InvalidFormat,StreamNotFound→ComponentNotFound,Corrupted→
  CorruptedFile`) as the umbrella's previous
  `src/common/error/conversions.rs`. Verified arm-by-arm.

### `escher/` and `writer` paths

✅ **No logic change detected.**

- `escher/{container,parser,properties,record,shape,shape_factory,types,
  writer}.rs`: R100, byte-identical except for line endings.
- `escher/text.rs`: only `is_multiple_of(2)` → `% 2 != 0` at two
  decoding-length checks (UTF-16 input length). Semantically equivalent.
- `doc/writer/*` — already covered above; just path rewrites, no
  algorithm changes, no removed bounds checks, no `#[repr(C)]` field
  reorderings.

Cargo metadata (`Cargo.toml`) declares the dependencies the source
actually uses (`litchi-cfb`, `litchi-core`, `bumpalo`, `bytes`,
`bitflags`, `chrono`, `once_cell`, `smallvec`, `thiserror`, `zerocopy`,
`zerocopy-derive`, optional `litchi-formula` and `litchi-imgconv`).
Default features = `[]`; the umbrella threads the `formula` and
`imgconv` features through `litchi-ole?/{formula,imgconv}`, preserving
the same gating that existed pre-split.

## 4. Summary

**Verdict: PURE MECHANICAL.**

All deltas observed in `145026a` are one of:

1. Path rewrites (`crate::ole::X` → `crate::X`,
   `crate::common::Y` → `litchi_core::Y`,
   `crate::images::Z` → `litchi_imgconv::Z` or `crate::extractor::Z`,
   `crate::sheet::T` → `litchi_core::sheet::T`).
2. CRLF → LF line-ending normalisation on rename.
3. The `DocElement` rename (umbrella↔crate seam glue, conversion happens
   on the umbrella side with identical payload).
4. Visibility upgrade `pub(crate) → pub` on `Presentation::extract_text_fast`
   (explicitly allowed).
5. Two `From<DocError>/From<PptError> for litchi_core::Error` impls
   relocated from `src/common/error/conversions.rs` into the per-package
   modules — match arms verified arm-by-arm.
6. `is_multiple_of(N)` → `% N != 0` (MSRV workaround; semantically
   equivalent for the integer types in use).

`3c1d141` is documentation-only (header breadcrumb cleanup + doctest
import-path corrections). `71158bd` adds a per-crate README; no source
changes.

I found **no algorithm changes, no removed validation/bounds checks, no
struct field reorderings, no changed default values, no removed
panic/assert/debug_assert, no changed error semantics, and no new
TODO/FIXME/unimplemented!()** in `litchi-ole`.

**Concerns: 0.**
