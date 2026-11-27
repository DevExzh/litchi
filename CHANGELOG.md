# Changelog

All notable changes to Litchi are documented in this file.

The format is based on [Keep a Changelog], and this project adheres to
[Semantic Versioning].

## [Unreleased]

Litchi now reads and writes every supported format through one API, a command
line tool and Python bindings, and is split into a workspace of format crates.
This section covers all changes since 0.0.1, including work that reached
`main` after that release.

### Added

- **Unified API.** `litchi::open`, `open_bytes`, `open_with` and
  `open_bytes_with` detect the format from the content and return one `File`
  for 21 formats: DOC, DOCX, RTF, ODT, OTH and ODM documents; PPT, PPTX and ODP
  presentations; XLS, XLSX, XLSB and ODS workbooks; ODG, ODC, ODI, ODB and ODF
  formula objects; and Pages, Keynote and Numbers files. `litchi::create`
  starts a new file. The operations live in the new `litchi-ops` crate.
- **Reading.** `text()` and `metadata()` work on every format, and the
  `document()`, `presentation()` and `workbook()` views give family-specific
  reads. `get` and `query` address elements by path (`/body/p[1]`,
  `/Sheet1/B2`) or selector (`paragraph[style=Heading1]`); text, outline,
  annotated, statistics and JSON views summarize a file; `capabilities()`
  lists the verbs and properties it supports; `tolerance_report()` lists the
  repairs made while reading damaged input.
- **Editing.** `set`, `add`, `remove`, `move_to` and `swap` return a new `File`
  and leave the input unchanged. `File::edit` groups edits into a transaction
  whose patch re-applies or inverts exactly, and `batch` runs a list of
  operations atomically or best-effort. A request a format cannot honor fails
  with `unsupported_capability` instead of writing part of the change.
- **Saving.** `save`, `save_as`, `save_with`, `to_bytes` and `write_to`. Saving
  to a path replaces the file atomically and can refuse to overwrite.
- **Document engines.** Find and replace; template merge from JSON or CSV
  (`{{key}}` placeholders and Word merge fields); refresh of fields, tables of
  contents, slide numbers, dates and formula caches; CSV and TSV import into
  worksheets; raw part access with XPath edits; and package validation.
- **Options and errors.** `OpenOptions` sets input and per-format limits, lazy
  or eager reading, and an execution context whose budgets and cancellation
  stop long operations. Every failure is a `litchi::Error` with a stable
  `Code`, a message, an optional suggestion, the detected format and the
  format crate's typed error as its source.
- **Command line.** The `litchi` binary (crate `litchi-cli`) runs the unified
  operations on files: `create`, `get`, `query`, `view`, `set`, `add`,
  `remove`, `move`, `swap`, `batch`, `capabilities`, `schema`, `import`,
  `evaluate`, `recalculate`, `merge`, `refresh`, `dump`, `raw`, `raw-set`,
  `add-part` and `validate`. It prints text or OfficeCLI-compatible JSON,
  saves edits atomically in place unless `--out` is given, and returns
  distinct exit codes for usage, path, property, unsupported, I/O,
  corrupt-input, protected and partial failures.
- **Python.** `litchi_py.open`, `open_bytes` and `create` return an immutable
  `File` with the same paths, queries, edits, transactions, batches, patches,
  saves, import, merge, refresh and raw-part operations. Exceptions carry the
  error code, message, suggestion and failing batch item, and subclass
  `OSError`, `ValueError` or `LookupError` where that fits. Type stubs ship
  with the package.
- **Word.** `litchi-doc` reads and writes Word 97-2003 documents, including
  every story (main text, footnotes, endnotes, headers, comments and text
  boxes), fields, bookmarks, tables, sections, revisions, hyperlinks, pictures
  and embedded objects. `litchi-docx` reads, edits and builds DOCX, DOCM and
  DOTX files: styles, numbering, settings, themes, comments, notes, headers and
  footers, glossary documents, content controls, fields, custom XML,
  bibliography sources, charts, SmartArt, text boxes, watermarks and OLE
  objects.
- **RTF.** `litchi-rtf` parses RTF 1.9.1 and writes it with a streaming
  writer. Edits to paragraphs, runs, table cells, headers and footers become
  reversible patches that keep unknown control groups. Compressed RTF
  ([MS-OXRTFCP]) can be read and written.
- **Excel.** `litchi-xls` reads and writes BIFF8 workbooks: cells, formulas,
  names, styles, merged cells, comments, validation, conditional formatting,
  filters, tables, protection, pivots, charts, drawings and external links.
  `litchi-xlsx` edits XLSX and XLSM packages without losing unknown content,
  covering sheets, cells, styles, tables, filters, pivots, charts, drawings,
  threaded comments, XML maps and the calculation chain. `litchi-xlsb` reads,
  writes and creates XLSB workbooks.
- **Text workbooks.** With `sheet-text`, `litchi-sheet-text` reads and writes
  CSV, TSV, PRN, SYLK, DIF and fixed-width text, with byte-order-mark
  detection and size limits.
- **PowerPoint.** `litchi-ppt` reads and writes PowerPoint 97-2003
  presentations: slides, text, shapes, notes, comments, media, hyperlinks,
  animations, charts and embedded fonts. `litchi-pptx` adds, copies, reorders
  and removes slides in PPTX and PPTM files and edits layouts, masters, notes,
  classic and modern comments, tables and table styles, charts, media,
  transitions, animations, ink and embedded fonts. A slide can be copied into
  another PPTX file that has a matching layout, master and theme.
- **OpenDocument.** ODT, ODS and ODP files and their OTT and OTS templates are
  read, edited and saved through immutable packages and transactions. ODS adds
  an OpenFormula parser and an evaluator for common function groups. New
  crates cover ODG drawings, ODC charts, ODI images, ODB database front ends,
  ODM master documents, OTH HTML templates and ODF formula documents (MathML).
  Flat single-file XML variants such as `.fodt`, `.fods` and `.fodp` open
  through the same readers, with edits where the format supports them.
  Password-protected packages open with a password, and ODT packages can also
  be encrypted and signed. `litchi-odf` now detects the format and re-exports
  the family crates behind features.
- **Apple iWork.** `litchi-pages`, `litchi-numbers` and `litchi-keynote` read
  single-file packages and app-authored package directories: Pages text,
  tables, footnotes, headers, layout, media and comments; Numbers sheets,
  tables and cells with formulas, cached values and formats; Keynote slides,
  notes, builds, transitions, charts, media and tables. Focused edits, such as
  Pages section text and tables, Numbers cells, formulas and sheets, and
  Keynote slides, text and transitions, return exact reversible patches and
  leave every other record byte-identical. New Pages, Numbers and Keynote
  files can be created.
- **Formula evaluation.** With `eval`, `FormulaEvaluator` computes cells and
  sheets of any workbook view, including defined names, structured table
  references and circular-reference detection, across aggregate, criteria,
  database, date and time, engineering, financial, information, logical,
  lookup, math, statistical and text functions. `File::evaluate`,
  `File::recalculate` and the CLI use it. `web-functions` adds `ENCODEURL`,
  `FILTERXML` and `WEBSERVICE`, which fetches through a function the caller
  supplies.
- **Equations.** With `equation`, `litchi-equation` converts Office Math
  (OMML), MathType and Equation Editor (MTEF) and MathML equations to LaTeX,
  converts LaTeX to OMML and to MTEF, and can render PNG previews. The DOC,
  DOCX, PPT, PPTX, XLS and XLSX crates read and edit the equations they
  contain.
- **Markdown.** `litchi-markdown` reads CommonMark and GitHub Flavored Markdown
  into bounded snapshots that support reversible edits, JSON patches and
  three-way merges.
- **Fonts.** With `fonts`, DOCX, PPT and PPTX files can embed fonts, including
  OOXML font obfuscation and the PowerPoint embedded-font format.
  `font-subset` subsets glyphs, `font-discovery` finds system fonts, and
  `automatic-fonts` enables both.
- **Encryption.** With `encryption`, password-protected DOC, XLS, PPT, DOCX,
  XLSX and PPTX files open as `File::Locked` and unlock with
  `Locked::unlock`. Legacy files use RC4 or CryptoAPI encryption, and XLS
  also XOR obfuscation; OOXML packages use Standard or Agile encryption, Agile
  with integrity checks. An unlocked
  file stays encrypted when saved unless the caller asks otherwise, and the
  format crates can encrypt and re-key files. `litchi-crypto` provides the
  [MS-OFFCRYPTO] primitives.
- **Digital signatures.** Signed OOXML and legacy binary files open as
  `File::Signed`. With `sign`, `litchi-sign` verifies and creates XML
  signatures (RSA with SHA-2 and ECDSA P-256) for OOXML packages and for the
  signature storage of DOC, PPT and XLS files. Verification accepts RSA keys
  of 1024 to 8192 bits, and signing needs at least 2048. Certificate trust and
  revocation are left to the caller.
- **Macros and active content.** VBA projects can be inspected (the
  `vba-inspection` feature for the OOXML formats and PPT) and are preserved
  when a file is edited. `litchi-vba` reads and writes [MS-OVBA] projects
  without compiling or running them.
- **Shared drawing models.** `litchi-drawingml` models DrawingML charts,
  SmartArt, fills, geometry and ink; `litchi-odraw` OfficeArt drawing records;
  `litchi-ograph` charts in legacy files ([MS-OGRAPH]); and
  `litchi-spreadsheet-drawing` SpreadsheetDrawing anchors, shapes and VML.
- **Building blocks.** `litchi-cfb` reads and writes compound files, including
  version 4 files larger than 2 GiB. `litchi-opc` reads and
  writes bounded OPC packages and decodes parts on first use. `litchi-biff`
  frames BIFF records losslessly, `litchi-codepage` decodes and encodes legacy
  code pages, and `litchi-xldm` inspects Spreadsheet Data Model storage.
  `litchi-ole-common` and `litchi-ooxml-common` hold shared OLE and OOXML
  structures, and `litchi-word`, `litchi-sheet` and `litchi-slide` define
  format-neutral vocabularies.
- **Documentation.** An Office API guide, the unified operations design, an API
  glossary, architecture decision records, a feature checklist for each format
  crate, and examples.
- **Test corpus and tooling.** Fixture documents from Apache POI, LibreOffice,
  the Open XML SDK, OfficeCLI, odfpy, odfdo and native Office files, each with
  its source and license recorded; fuzz targets for the main parsers; and
  scripts that check crate dependency boundaries, API naming and fixture
  integrity.

### Changed

- **Breaking:** The workspace is split into one crate per format plus shared
  foundation crates. The `litchi` crate is a facade that re-exports them as
  modules such as `litchi::doc`, `litchi::docx`, `litchi::xlsx`, `litchi::odt`
  and `litchi::pages`, each behind its feature.
- **Breaking:** The `litchi` crate enables no formats by default. Choose
  format features (`doc`, `docx`, `ppt`, `pptx`, `xls`, `xlsx`, `xlsb`, `rtf`,
  `odt`, `ods`, `odp`, `oth`, `odm`, `odg`, `odc`, `odi`, `odb`, `odf-formula`,
  `pages`, `keynote` and `numbers`), bundles (`legacy`, `ooxml`, `odf`,
  `iwork`, `word`, `slides`, `spreadsheets`, `office` and `all-formats`), or
  `all` to add every optional capability. Markdown conversion now needs the
  `markdown` feature.
- **Breaking:** Rust 1.95 is the minimum supported version, and the toolchain
  is pinned in `rust-toolchain.toml`.
- **Breaking:** Format crates expose immutable packages. A package returns
  read-only document, workbook or presentation projections; changes go through
  `edit()` and `commit()`, which produce a new package and a patch that can be
  re-applied or inverted. This replaces the mutable documents and builders of
  the former OpenDocument, RTF and OOXML modules.
- **Breaking:** Readers take explicit limits. Compound files open with
  `OleFileLimits`, OPC packages with `SourceLimits`, and each format crate has
  source and semantic limits built through checked builders. A limit error
  names the limit, the observed value and the maximum.
- **Breaking:** The Python package `litchi-py` requires Python 3.9 or later.
  Its `Document`, `Presentation` and `Workbook` classes remain for reading;
  edits and path queries go through `File`.
- Markdown conversion works on any document or presentation opened with
  `litchi::open` and refuses content Markdown cannot represent instead of
  dropping it.
- Image conversion (`images`, formerly `imgconv`) prefers vector output: EMF,
  EMF+, WMF and PICT pictures become SVG unless only a bitmap is possible, and
  PNG, JPEG and WebP remain available. Each conversion reports diagnostics,
  strict helpers refuse approximations, and no conversion substitutes a
  placeholder image.
- Equation conversion moved to `litchi-equation`, which uses the `mathema`
  converter.
- Web functions in the formula engine no longer make network requests
  themselves. They use a fetcher supplied by the caller and bound response
  sizes.
- The `fonts` feature no longer always pulls in glyph subsetting and system
  font discovery. Enable `font-subset` or `font-discovery` for them.
- `soapberry-zip` reads archives within explicit `ArchiveLimits`, and reads
  members in parallel only through an explicit `ParallelReadSession` instead
  of the global Rayon pool.
- `xml-minifier` adds an auditor that checks compact XML without rewriting
  it. The `minified_xml!` macros moved to the new `xml-minifier-macros` crate,
  remain re-exported by `xml-minifier`, and reject line breaks and tabs used
  only for indentation.
- Dependencies are updated, among them PyO3 0.30, parking_lot 0.13,
  serde-saphyr 2, zip 9 and rstest 0.29. The formula engine no longer depends
  on `reqwest`, and equation conversion no longer uses `rowan`. Allsorts is
  vendored with bounded allocation budgets, together with pinned copies of
  flate2 and zlib-rs.
- CI checks formatting, manifest sorting, lints, documentation, deprecated
  APIs, crate dependency boundaries and feature combinations, and runs the
  workspace tests on Linux and the foundation and Microsoft Office crate tests
  on macOS and Windows. Release gates build, lint, document and test each Rust
  package except the Python extension on its own with all its features, in two
  lanes: one for the non-iWork packages, which must build without `protoc`,
  and one for the iWork packages, `litchi-ops` and `litchi-cli`. The facade
  runs in both. Checkouts keep every file byte-identical on all three
  platforms.

### Deprecated

- `soapberry-zip`'s `read_many_parallel`, `read_many_parallel_results`,
  `read_many_parallel_cached` and `read_all_parallel` now read serially. Use a
  `ParallelReadSession` with `read_many_with_session` or
  `read_all_with_session`.

### Removed

- **Breaking:** The Rust `Document`, `Presentation` and `Workbook` openers and
  the `sheet::open_*_workbook` helpers. Call `litchi::open` and use the
  `document()`, `presentation()` or `workbook()` view, or a format module.
- **Breaking:** The `litchi::document`, `litchi::presentation`,
  `litchi::ooxml`, `litchi::odf`, `litchi::iwa` and `litchi::formula`
  modules. Use the format modules such as `litchi::docx`, `litchi::xlsx`,
  `litchi::odt`, `litchi::pages` and `litchi::equation`.
- **Breaking:** The `litchi-ole`, `litchi-ooxml`, `litchi-iwa` and
  `litchi-formula` crates. Their code now lives in the format and foundation
  crates; `litchi::ole` keeps only shared OLE structures.
- **Breaking:** The old OpenDocument root types (`Document`, `Spreadsheet`,
  `Presentation`, `DocumentBuilder`, `PresentationBuilder`,
  `MutableDocument` and `MutablePresentation`), the RTF `RtfDocument` and
  `RtfWriter`, and `litchi_iwa::Document`. Use each format crate's `Package`.
- **Breaking:** The old Cargo feature names. Use `iwork` (or `pages`,
  `keynote` and `numbers`) for `iwa`, `equation` for `formula`, `images` for
  `imgconv`, `encryption` for `ooxml_encryption`, `eval` for `eval_engine`,
  `web-functions` for `eval_engine_web_functions`, and `all` for `full`. The
  `ooxml` and `odf` features now select the formats they name, and `ole` now
  enables only the shared OLE structures; use `legacy` for DOC, PPT and XLS.

### Fixed

- Legacy Office compound files open when the final sector is shorter than
  the declared sector size, including when the allocation table occupies that
  sector.
- Extracting pictures from binary Word documents returns each stored image
  once when several references point to it.

### Security

- Passwords never pass through `OpenOptions`: an encrypted Office file stays
  `File::Locked` until `Locked::unlock` receives one. Signed content cannot
  change without an explicit `unsign()`.
- Every reader is bounded. Configurable limits cap input size, archive
  expansion, ZIP members, OOXML relationship traversal, nesting depth and
  embedded image decoding, and exceeding one fails with `limit_exceeded`.
- Macros, ActiveX controls, OLE objects, DDE and external links, scripts and
  database connections are never executed or fetched; they are kept or
  exposed as inert data.
- iWork package directories reject symbolic links and special files, and
  sources that change while being read are refused.
- XML attribute parsing avoids quadratic duplicate checks on hostile tags.

## [0.0.1] - 2025-10-19

### Added

- First release: reading DOC, DOCX, XLS, XLSX, XLSB, PPT, PPTX, ODT, ODS, ODP,
  Pages, Numbers and Keynote files with automatic format detection, converting
  MathType and Office Math equations to LaTeX, converting EMF, WMF and PICT
  pictures to raster images, and converting documents and presentations to
  Markdown.

[Keep a Changelog]: https://keepachangelog.com/en/1.1.0/
[Semantic Versioning]: https://semver.org/spec/v2.0.0.html
[Unreleased]: https://github.com/DevExzh/litchi/compare/928bf5e6c7accab381b4f13ebabb6e7155a5dd05...HEAD
[0.0.1]: https://github.com/DevExzh/litchi/tree/928bf5e6c7accab381b4f13ebabb6e7155a5dd05
[MS-OFFCRYPTO]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
[MS-OGRAPH]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ograph/
[MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
[MS-OXRTFCP]: https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp/
