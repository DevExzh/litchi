# Leniency audit: over-strict validation across the format parsers

Date: 2026-10-01

## Purpose and method

This document compiles a nine-scope parallel review of the format parsers' error paths,
looking for validation stricter than the reference applications (Word, Excel, PowerPoint,
LibreOffice) and reference libraries (Apache POI, Open-XML-SDK). Scopes: litchi-doc,
litchi-docx, litchi-xls + litchi-biff + litchi-xlsb, litchi-xlsx, litchi-ppt,
litchi-pptx, litchi-cfb + litchi-ole-common + litchi-crypto, litchi-opc +
litchi-ooxml-common, and litchi-rtf + the ODF family (litchi-odt/ods/odp/odf-common).
Each strict check was weighed against normative spec text ([MS-DOC], [MS-XLS], [MS-XLSB],
[MS-PPT], [MS-CFB], [MS-OLEPS], ECMA-376, ODF 1.2, RTF 1.9.1) and the checked-in reference
sources under `3rdparty/` (libreoffice-core, poi, specs mirrors).

Every finding kept below was verified against the working tree on 2026-10-01: each cited
file exists and each quoted check or error string was located at (or within a few lines
of) the cited position. Line numbers here are the verified ones; corrections applied are
listed in Appendix A, and areas reviewers found already tolerant in Appendix B.

**108 findings kept, 0 dropped: 8 P0, 67 P1, 33 P2.** Confidence is high unless the
finding title says otherwise.

## The tolerance contract

Findings align with the repair model established in `crates/litchi-doc/src/leniency.rs`
and mirrored by `XlsLeniency` in `crates/litchi-xls/src/leniency.rs`:

1. **Word-compatible repairs are applied by default and recorded.** Under every
   `Leniency` variant, including default `Leniency::Strict`, the reader repairs bytes the
   spec leaves undefined but real producers emit anyway — bytes Word and LibreOffice read
   past without complaint. Each repair substitutes the value documented on the defect
   variant and is *recorded* in a bounded `ToleranceReport` (`MAX_RECORDED_DEFECTS =
   1024`; overflow is counted in `unrecorded()`). Nothing is swallowed silently: an empty
   report means the document conformed. The established example is
   `SprmDefect::CellBorderSideFlags` — undefined `bordersToApply` side bits on
   `sprmTSetBrc80`/`sprmTSetBrc` are masked to the defined set and still applied.
2. **Opt-in repairs trade strictness for reach behind explicit `Leniency` variants.**
   Repairs that involve a judgment call stay gated: `StylesheetDefect::DuplicateStyleName`
   only under `Leniency::TolerateStylesheetDefects`; litchi-xls's five `FormattingDefect`
   classes (`FontFamily`, `FontNameEmpty`, `AlignmentJustifyLastLine`,
   `ExtendedFormatCountMismatch`, `FormatStringOverrun`) only under
   `Leniency::TolerateFormattingDefects`.
3. **Structural corruption and security limits stay hard errors everywhere.** Malformed
   framing (FIB, piece table, record/sector grammar, XML well-formedness), broken stream
   grammar, unsupported/absent decryption material, and configurable resource/DoS limits
   are rejected under every variant.

Two plumbing gaps frame the roadmap: only litchi-doc and litchi-xls have a
`ToleranceReport` carrier at all (docx, xlsx, ppt, pptx, cfb, ole-common, opc,
ooxml-common, rtf, odf-common have none), and litchi-xls's leniency is not threaded into
the source-backed open, which hardcodes `Leniency::Strict`
(`crates/litchi-xls/src/workbook/source.rs:2114`). Introducing the carrier everywhere is
the prerequisite step for nearly every fix below.

## Summary by scope

| Scope | Kept | P0 | P1 | P2 | Top theme |
|---|---:|---:|---:|---:|---|
| litchi-doc | 12 | 2 | 7 | 3 | "MUST be ignored" reserved bits validated; ~24 optional tables `?`-propagate to open |
| litchi-docx | 12 | 1 | 7 | 4 | Single-pass whole-part codecs turn one schema deviation into a part-wide `Err` |
| litchi-xls + biff + xlsb | 12 | 1 | 10 | 1 | Cross-reference integrity (XF/format/font) and reserved bits treated as fatal |
| litchi-xlsx | 12 | 0 | 8 | 4 | Application authoring limits enforced as format rules at load time |
| litchi-ppt | 12 | 2 | 6 | 4 | Slide-directory / persist-chain cross-reference checks are file-fatal |
| litchi-pptx | 12 | 1 | 5 | 6 | Value-domain strictness, sometimes contradicting the schema's own types |
| litchi-cfb + ole-common + crypto | 12 | 0 | 8 | 4 | Producer-MUST fields treated as parser-fatal at compound-file ingress |
| litchi-opc + ooxml-common | 12 | 0 | 9 | 3 | Metadata parts (core/custom properties) veto the whole document open |
| litchi-rtf + ODF family | 12 | 1 | 7 | 4 | Container conformance and degenerate scalars as read-time fatality |

Three cross-cutting themes repeat in nearly every scope and are the cheapest places to
build shared machinery: (a) **reserved-bit masking** — the spec's "MUST be zero and MUST
be ignored" wording appears in doc F1.2-F1.5/F1.9, xls F3.4/F3.6/F3.7, xlsb F3.11/F3.12,
ppt F5.2 — the class `SprmDefect::CellBorderSideFlags` already repairs; (b) **per-item
degradation** — one bad cell/row/property/directory entry/animation/comment kills an
entire part or file where reference apps degrade per-item; (c) **lossy text decode** —
strict UTF-16/MBCS rejects files where every reference substitutes U+FFFD or a default
charset (xls F3.8, xlsb F3.12, rtf F9.6, ole-common F7.8, xlsx F4.12; codepage fallbacks
in xls F3.2, ole-common F7.9, rtf F9.12).

## Top findings (P0)

Eight items drive the iteration goal; each has a verified real-world producer class, a
total-failure blast radius out of proportion to the defect, and a repair the reference
applications already perform. Ranked by expected impact:

1. **litchi-doc F1.1** — **[FIXED 2026-10-01]** ~24 optional metadata tables `?`-propagate at document open; any
   deviation in any auxiliary FIB table fails the whole file. Largest single blast radius
   in the audit; the crate already proves the tolerant pattern in the same function.
2. **litchi-docx F2.1** — **[FIXED 2026-10-02]** one OMML equation fails `Document::write_text_to` for the whole
   document; the sibling scanner in the same file already skips foreign subtrees.
   Equations are ubiquitous in technical documents.
3. **litchi-xls F3.1** — **[FIXED 2026-10-01]** per-cell XF cross-reference validation kills the worksheet;
   out-of-range `ixfe` is one of the most common non-Microsoft-producer defects, and the
   repair (default XF) touches styling only, never values.
4. **litchi-ppt F5.1** — **[FIXED 2026-10-02]** `StyleTextPropAtom` exact-coverage framing; miscounted coverage
   is among the most common .ppt deviations, and the crate already ships the lenient
   parser the read path never uses.
5. **litchi-ppt F5.2** — **[FIXED 2026-10-02]** reserved BulletFlags/PFWrapFlags bits fatal; verbatim "MUST be
   ignored" spec text, and one stray bit kills a whole slide's shapes via the missing
   per-shape isolation (F5.11).
6. **litchi-pptx F6.1** — **[FIXED 2026-10-02]** animation stale-shape-target and duplicate-shape-ID checks kill
   the timing tree; PowerPoint routinely leaves orphaned timing entries when shapes are
   deleted, so edited decks are very likely carriers.
7. **litchi-doc F1.2** — **[FIXED 2026-10-01]** stylesheet Stshif reserved-bit/fixed-count checks; the
   stylesheet is mandatory for every Word 97+ file, so one stray bit costs the whole
   document — the same class as the already-fixed cell-border side flags.
8. **litchi-rtf/ODF F9.1** — **[FIXED 2026-10-02]** deflated or re-zipped ODF `mimetype` member rejected at open
   and at detection (classifies as "unsupported format"); LibreOffice inflates it through
   ordinary ZIP APIs. The crate already has an opt-in repair plan for the deviation class.

## 1. litchi-doc (legacy .doc)

Infrastructure: `Leniency`/`ToleranceReport` exist (`src/leniency.rs`) — the template for
all other scopes. Framing layers (FIB pointer-array degradation, lossy text extraction,
`Option`-based FKP/PLCF, ignored unknown SPRM opcodes) are already tolerant; strictness
concentrates in reserved-bit validation, fatal `?` propagation, and enum-domain checks.
### 1.1 Optional metadata-table parse errors abort the entire document open — P0
> **Status: FIXED (2026-10-01)** — every optional FIB-pointed table now parses fallibly
> (absent/default on error) and records the new `AuxiliaryTableDefect` (23 variants,
> context = FIB pointer index), enumerable via `Document::tolerance_report()`.
- **Location**: `crates/litchi-doc/src/document/package.rs:170-207` — `?` on every
  FIB-pointed auxiliary table (comments :178, bookmarks :187, rsids :189, ~20 more).
- **Check**: any `Corrupted` from any one of these tables fails the whole
  `Package::document()` call.
- **Spec says**: no spec basis — [MS-DOC] 2.5 makes these fc/lcb pairs independently
  optional metadata; a malformed `SttbfRMark` or `PlcfAsumy` invalidates nothing else.
- **Real-world behavior**: Word/LibreOffice open and ignore the broken sub-structure; the
  crate itself uses `.ok()`/deferred `Result` for headers/footnotes/endnotes/list tables
  (:173-177, :253, :310-358). Third-party writers routinely emit slightly-off PLCFs.
- **Suggested posture**: default-repair — parse each table fallibly, on error store
  `None`/deferred `Result` and record `Defect::AuxiliaryTable(index)`; hard failure only
  for FIB, piece table, encryption, stylesheet framing.
### 1.2 Stylesheet Stshif reserved-bit and fixed-count checks reject the whole stylesheet — P0
> **Status: FIXED (2026-10-01)** — both checks now record `StshifDefect::{StyleNameFlags,
> FixedStyleCount}` and parse on; enumerable via `Document::tolerance_report()`.
- **Location**: `crates/litchi-doc/src/parts/styles/codec/parser.rs:82` ("Stshif
  style-name flags are invalid"), :84-87 (`istdMaxFixedWhenSaved != 15` → error).
- **Check**: the Stshif flag word must equal exactly 1 and the count must equal 15, else
  `PackageError::Corrupted`; the stylesheet is mandatory for every Word 97+ file.
- **Spec says**: [MS-DOC] 2.9.274: "fStdStylenamesWritten (1 bit): MUST be 1 **and MUST
  be ignored**. fReserved (15 bits): MUST be 0 **and MUST be ignored**";
  `istdMaxFixedWhenSaved` is never consulted for layout.
- **Real-world behavior**: LibreOffice masks bit 0 and reads on (`a16Bit & 0x0001`,
  ww8scan.cxx:6905) and never range-checks the count. Same class as the fixed cell-border
  case.
- **Suggested posture**: default-repair — test bit 0 only, accept any count, record.
### 1.3 GRFSTD / StdfPost2000 reserved bits and iPriority range validated instead of ignored — P1
> **Status: FIXED (2026-10-02)** — reserved GRFSTD/StdfPost2000 bits masked, iPriority > 99 clamped; recorded as `StyleDefect::{ReservedStyleFlags,
> ReservedPost2000Flags, Priority}` (context = istd); raw STD bytes kept for round-trip, writer still rejects priority > 99.
- **Location**: `crates/litchi-doc/src/parts/styles/codec/parser.rs:166` ("GRFSTD
  contains reserved flags"), :183-189 (StdfPost2000), :191-193 (`priority > 99`).
- **Check**: any set reserved/unused bit or a style priority outside 0-99 aborts the
  stylesheet and hence the document.
- **Spec says**: [MS-DOC] 2.9.112/2.9.261: the reserved bits "MUST be 0 **and MUST be
  ignored**"; `iPriority` is pure UI-ordering metadata.
- **Real-world behavior**: LibreOffice reads `fAutoRedef`/`fHidden` and explicitly skips
  the rest (ww8scan.cxx:7007-7013); localized/third-party producers leave garbage in
  reserved style bits.
- **Suggested posture**: default-repair — mask defined bits, clamp/ignore priority,
  record. Same shape as `SprmDefect::CellBorderSideFlags`.
### 1.4 Comment descriptor (ATRDPre10) reserved fields, author bounds, bookmark-tag domain — P1
> **Status: FIXED (2026-10-02)** — non-zero ATRDPre10 reserved fields ignored, out-of-range author → empty author, lTagBkmk < -1 → point comment;
> recorded as `CommentDefect::{ReservedFields, AuthorIndex, BookmarkTag}` (context = comment index).
- **Location**: `crates/litchi-doc/src/parts/comments.rs:48-54` ("ATRDPre10 reserved
  fields must be zero"), :43-47 (author index), :62-66 (`lTagBkmk < -1`); fatal via
  package.rs:178.
- **Check**: one non-zero reserved field, out-of-range author index, or tag below -1 in
  one comment kills the document.
- **Spec says**: [MS-DOC] 2.9.7: "bitsNotUsed (2 bytes): MUST be 0, **and MUST be
  ignored**. grfNotused (2 bytes): MUST be 0, **and MUST be ignored**." The `ibst` bound
  is a cross-reference, not framing.
- **Real-world behavior**: Word renders the comment with a generic author; POI never
  validates these fields. Comments are extremely common and round-tripped through tools
  that leave junk in unused ATRD bytes.
- **Suggested posture**: default-repair — ignore reserved fields, degrade author to a
  placeholder, clamp negative tags to `None`; record.
### 1.5 Bookmark BKC fPub reserved bit, ibkl uniqueness, CP-consistency checks — P1
> **Status: FIXED (2026-10-02)** — BKC.fPub masked; bad/duplicate ibkl or reversed range drops only that bookmark; empty column range demotes to a text bookmark;
> recorded as `BookmarkDefect::{ReservedPublicFlag, EndIndex, ReversedRange, ColumnRange}` (context = bookmark index); table framing stays fatal.
- **Location**: `crates/litchi-doc/src/parts/bookmarks.rs:91` ("bookmark BKC fPub must be
  zero"), :83-87 (non-unique `ibkl`), :98-101 (start > end), :106-110 (empty column
  range); fatal via package.rs:187.
- **Check**: one bookmark with the reserved bit, a duplicate end-index, or reversed CPs
  rejects the document.
- **Spec says**: [MS-DOC] 2.9.8 (BKC): "A - fPub (1 bit): MUST be 0, **and MUST be
  ignored**." Uniqueness/CP ordering are internal consistency rules, no reader
  obligation.
- **Real-world behavior**: POI reads `fPub` as an inert bitfield accessor
  (BKFAbstractType.java:34); Word opens documents with broken bookmarks.
- **Suggested posture**: default-repair the fPub bit (mask + record); demote bad
  individual bookmarks rather than failing the table; F1.1 makes wholesale table failure
  non-fatal anyway.
### 1.6 Table border operands: unknown brcType / Ico index are fatal — P1
> **Status: FIXED (2026-10-02)** — undefined brcType → `BorderType::None` (LibreOffice default), Brc80 Ico > 16 → automatic; recorded as
> `SprmDefect::{BorderType, BorderColor}` (context = SPRM opcode).
- **Location**: `crates/litchi-doc/src/parts/tap_parser/codec/borders.rs:42` ("Brc80
  contains invalid color index {ico}"), :89 ("DOC border contains invalid border type").
- **Check**: any border-type byte outside the enumerated set or a color index above 16
  raises `Corrupted`, propagating via `PapBinTable::parse` (pap_bin_table.rs:201) to
  document open failure (package.rs:290).
- **Spec says**: [MS-DOC] 2.9.17 (Brc80): "brcType … MUST NOT be 0x1A or 0x1B" — a
  writer constraint; sibling bits in the same byte are "MUST be ignored".
- **Real-world behavior**: POI's `BorderCode` stores brcType raw with zero validation
  (BorderCode.java:34); LibreOffice maps unknown types to default/none. Same operand
  family as the already-fixed `bordersToApply` side flags.
- **Suggested posture**: default-repair — unknown brcType → `None` (or `Single`),
  out-of-range Ico → `Automatic`; record. Closest analog of
  `SprmDefect::CellBorderSideFlags`.
### 1.7 TC80 cell-descriptor unused enum values and boundary checks are fatal — P1
> **Status: FIXED (2026-10-02)** — unused TC80 textFlow/vertMerge/vertAlign/ftsWidth values and negative wWidth read as neutral values, out-of-range rgdxaCenter clamped;
> recorded as `TableDefinitionDefect` (context = cell/boundary index); 63-column limit, truncation, decreasing boundaries stay fatal.
- **Location**: `crates/litchi-doc/src/parts/tap_parser/codec/definitions.rs:121` ("TC80
  contains invalid textFlow value"; vertMerge = 2 :131-135, vertAlign = 3 :144-148,
  ftsWidth > 3 :179-183, negative wWidth :159-163, XAS range :42-54).
- **Check**: any unused bit-pattern in a 20-byte `TC` cell descriptor — e.g. vertAlign
  value 3, which the format leaves undefined — kills the table property parse, the PAPX
  bin table, and the document.
- **Spec says**: [MS-DOC] 2.9.313 (TC80) defines the used values; unused encodings carry
  no "MUST reject" reader rule.
- **Real-world behavior**: Word clamps/ignores unknown cell-flag values and renders the
  table; LibreOffice reads TC80 flags bit-by-bit without rejecting unknown encodings.
- **Suggested posture**: default-repair — unknown enum → neutral member (textFlow→LrTb,
  vertMerge→None, vertAlign→Top, ftsWidth→None), record. Keep >63-column and truncation
  checks (framing).
### 1.8 One malformed PAPX run aborts all paragraph properties and the document — P1
> **Status: FIXED (2026-10-02)** — malformed/unreachable PapxFkp page skipped, unreadable run gets default paragraph properties, producer pad trim on by default;
> recorded as `PapxDefect::{Page, Run, AlignmentPadding}` (context = PlcfBtePapx page index).
- **Location**: `crates/litchi-doc/src/parts/pap_bin_table.rs:172-174` (one malformed FKP
  entry → Err), :193 (`parse_properties_with_direct_cached` in the per-run loop), :312;
  propagates at package.rs:227/290.
- **Check**: `PapBinTable::parse` returns `Err` on the first bad FKP page or the first
  `grpprl` the strict SPRM walker cannot consume exactly.
- **Spec says**: no spec basis for all-or-nothing — [MS-DOC] 2.4.3 defines each
  `PapxFkp` page independently.
- **Real-world behavior**: Word/LibreOffice render affected paragraphs with default
  formatting. The crate acknowledges a "repeated producer witness" of an odd SPRM prefix
  plus a stray zero byte (comment at pap_bin_table.rs:175-180) but gates the trim behind
  off-by-default `allow_papx_alignment_padding` (:130,181).
- **Suggested posture**: default-repair — skip the offending run/page with default
  `ParagraphProperties` (mirroring the CHPX per-run `unwrap_or_default()` at
  chp_bin_table.rs:186), record; make the proven one-byte pad trim unconditional like the
  cell-border mask.
### 1.9 Section SEPX: Brc80/SPgbProp reserved bytes and enum ranges are fatal — P1
> **Status: FIXED (2026-10-02)** — Brc80/SPgbProp reserved bits ignored, undefined page-border type/Ico → none/automatic, undefined section enums keep defaults, Boolean8 ∉ {0,1} → true;
> recorded as `SectionDefect` (context = SPRM opcode); SEPX framing and numeric range checks stay fatal.
- **Location**: `crates/litchi-doc/src/section/borders/codec.rs:74` ("{name} contains
  reserved Brc80 effect bits"), :98 ("sprmSPgbProp reserved byte must be zero"), :64-68
  (Ico range); enum/Boolean8 guards in `parts/sections/model.rs` (e.g. :491); fatal via
  package.rs:230.
- **Check**: the top reserved bit of a section border's effect byte, a non-zero SPgbProp
  reserved byte, an out-of-range color, or any undefined enum value in any SEPX rejects
  the document.
- **Spec says**: [MS-DOC] 2.9.17 (Brc80): "C - reserved (1 bit): MUST be zero, **and MUST
  be ignored**." The enum fields are defined-value lists with no reader-side rejection
  requirement.
- **Real-world behavior**: Word ignores the reserved border bit and uses defaults for
  undefined enum bytes. The walker already silently skips unknown SPRM opcodes
  (sections/model.rs:374) — only known opcodes with unexpected payloads are fatal, which
  is precisely backwards.
- **Suggested posture**: default-repair reserved bits (mask + record); out-of-range
  enums/Boolean8 → defaults + record; keep genuine framing errors (truncated SEPX,
  negative grpprl size).
### 1.10 PLRSID header constants validated with exact equality — P2
- **Location**: `crates/litchi-doc/src/parts/rsids.rs:79-87` ("PLRSID cbRsidInFile is not
  4" :80; also `cbHeadExtraInFile != 8`, `reserved1 != 229`); fatal via package.rs:189.
- **Check**: any deviation in the three fixed header words of the revision-save-ID table
  — including the reserved word that must equal 229 — fails the open.
- **Spec says**: [MS-DOC] 2.9.203: `cbRsidInFile` "MUST be 4", `reserved1` "MUST be
  229" — writer constants for an inert lookup table; no reader rule.
- **Real-world behavior**: Word ignores an odd RSID table (RSIDs only drive
  per-session-change UI).
- **Suggested posture**: default-repair — ignore `reserved1`; treat `cbRsidInFile != 4`
  as "table unreadable → None" and record.
### 1.11 CHP operand strictness: Boolean8 domain, Hres mode, revision-reason codes — P2 (confidence: medium-high)
- **Location**: `crates/litchi-doc/src/parts/chp/codec.rs:680` (`strict_bool8` errors on
  operand ∉ {0,1}), :670-678 (`revision_flag` same), :24-36 (hyphenation mode ∉ 1-6),
  :744-748 (`revision_reason` above max), :45 ("normal sprmCHresi requires ChHres 0x00").
- **Check**: a Boolean8 SPRM carrying any byte other than 0/1 (e.g. the toggle values
  0x80/0x81 the format itself uses elsewhere), an undefined hyphenation mode, or an
  undefined revision-reason code raises `Corrupted`. The CHPX bin-table path already
  degrades per-run (`unwrap_or_default`, chp_bin_table.rs:186), but the stylesheet-revision
  (styles/codec/parser.rs:347,353) and paragraph-extractor
  (paragraph_extractor.rs:642-669) paths fail wholesale on the same byte.
- **Spec says**: [MS-DOC] 2.9.27 (Boolean8): "MUST be 0 or 1"; POI's `getCHPFlag`
  semantics (0/1/0x80/0x81) show the ecosystem treats the byte as a toggle encoding.
- **Real-world behavior**: Word interprets any non-zero Boolean8 via its toggle rule;
  POI degrades per-property.
- **Suggested posture**: default-repair — apply the existing `get_toggle_value` semantics
  (chp/codec.rs:916-924) instead of erroring; map unknown enum codes to "unset" + record,
  consistently across bin-table, style, and extractor paths.
### 1.12 One unresolvable Pcd.Prm discards the entire piece table (silent total formatting loss) — P2
> **Status: FIXED (2026-10-02)** — an unresolvable Pcd.prm gives only that piece an empty modifier (text and FKP formatting kept) instead of silently dropping both bin tables;
> recorded as `PieceDefect::PropertyModifier` (context = piece index).
- **Location**: `crates/litchi-doc/src/parts/piece_table.rs:231`
  (`Self::resolve_property_modifier(prm, &property_groups)?` in the per-piece loop), :295
  (returns `None` on an out-of-range `Prm1` index or unlisted `Prm0.isprm`).
- **Check**: a single piece whose prm cannot be resolved makes `PieceTable::parse` return
  `None` — which at document/codec.rs:66-71 quietly yields *no* CHP/PAP bin tables at
  all: the document opens with every character and paragraph at default formatting, no
  error and no record.
- **Spec says**: [MS-DOC] 2.9.215 (Prm) defines isprm values; an unknown value carries no
  reader-rejection rule, and the prm modifies one piece only.
- **Real-world behavior**: Word ignores an uninterpretable prm and keeps the piece's text
  and FKP formatting. Silent total formatting loss is worse than an error.
- **Suggested posture**: default-repair — per piece, substitute an empty modifier when
  the prm cannot be resolved, record the piece index; never fail the whole table.

## 2. litchi-docx (WordprocessingML .docx)

Infrastructure: no tolerance-report/leniency mechanism yet — the litchi-doc
`ToleranceReport` pattern is the natural template. Container and structural layers are
tolerant (lazy open — package/codec.rs:373 checks only the main-part content type;
`OpaqueBlock`/`RunContent::Unknown`/`Inline::Unknown`; lossy style attribute reads), but
typed semantic codecs are single-pass, whole-part parsers that convert any schema
deviation into a part-wide or document-wide `Err`. The reviewer reproduced these findings
against the crate's public API.
### 2.1 Semantic text extraction rejects any foreign `*:t` element — OMML equations kill `write_text_to` — P0
> **Status: FIXED (2026-10-02)** — litchi-docx gained its `ToleranceReport` carrier (`Document`/`source_backed::Package::tolerance_report()`); text extraction skips each foreign `*:t`
> subtree (OMML `m:t`, DrawingML `a:t`) including entities/CDATA and records `TextDefect::SkippedForeignMarkup` (context = start-tag offset). Ill-formed XML, unbound prefixes, truncation and limits stay fatal.
- **Location**: `crates/litchi-docx/src/paragraph/codec/text.rs:925` (Start), :978
  (Empty), :1043 (End) — "foreign semantic DOCX text element is not w:t".
- **Check**: every element whose local name is `t` but is not bound to the
  WordprocessingML namespace raises `Error::InvalidFormat`; the driver routes *all*
  events through this check regardless of namespace.
- **Spec says**: no spec basis — internal strictness. OMML (`m:oMath/m:r/m:t`, ECMA-376
  Part 1 §22) legitimately lives inside `w:p`; DrawingML `a:t` appears in inline `wps:`
  shapes.
- **Real-world behavior**: Word/LibreOffice/POI all extract text from documents
  containing equations. Reproduced: one `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>`
  fails `Document::write_text_to`, while the older `extract_word_text` scanner
  (text.rs:42) tolerates it. Any document with an equation loses the primary streaming
  text API.
- **Suggested posture**: default-repair — treat non-w: namespaced subtrees as opaque
  (skip, as `extract_word_text` already does), record `SkippedForeignMarkup`. Same shape
  as litchi-pptx F6.9.
### 2.2 Unknown `w:compat` flag name hard-fails all settings reads — P1
- **Location**: `crates/litchi-docx/src/settings/document/codec.rs:352` ("invalid
  compatibility flag '{raw}'"; strict-conformance gate at :354-358).
- **Check**: any child of `w:compat` that is neither `compatSetting` nor one of the 66
  closed `CompatFlag` tokens raises `Error::InvalidFormat`, failing the whole settings
  part and every settings accessor.
- **Spec says**: ECMA-376 CT_Compat is a closed choice group (schema-backed), but nothing
  sanctions refusing the whole settings part.
- **Real-world behavior**: LibreOffice handles compat flags by token switch and silently
  ignores unknown children (SettingsTable.cxx:256-374); Word preserves/opens them. Newer
  Word versions and WPS introduce flags over time — a forward-compatibility trap.
  Reproduced: `<w:someFutureFlag/>` fails `pkg.attached_template()`.
- **Suggested posture**: default-repair — skip unknown flag elements (record
  `UnknownCompatFlag`); transitional flags in Strict docs → tolerated deviation.
### 2.3 Duplicate singleton settings elements hard-fail settings parsing — P1
- **Location**: `crates/litchi-docx/src/settings/document/codec.rs` — "duplicate …
  setting" at :359-362 (compat flag), :469-472 (documentProtection), :533-536
  (attachedTemplate), :563-564 (view), :592-594 (defaultTabStop), :623-625
  (clrSchemeMapping), :393-434 (note properties), ~15 sites total.
- **Check**: a second occurrence of any singleton setting raises `Error::InvalidFormat`.
- **Spec says**: schema maxOccurs=1 — schema-backed, but no basis for part-wide refusal.
- **Real-world behavior**: Word/LibreOffice process occurrences in order, last-wins.
  Reproduced: two `<w:view>` elements fail the whole settings parse. Duplicates arise
  from WPS, merge tools, and hand-edited settings.xml.
- **Suggested posture**: default-repair — last-wins (matches LO dmapper overwrite),
  record `DuplicateSetting`.
### 2.4 Numbering model is all-or-nothing: one bad value in numbering.xml kills `list_items()` — P1
- **Location**: `crates/litchi-docx/src/numbering/codec.rs:526` ("invalid numbering
  format"), :543 (suff other than tab/space/nothing), :551 (lvlRestart > 9), :907
  ("invalid on/off value"), :587/:608/:641/:657/:669/:677/:688/:699 (duplicate IDs and
  levels), :628 ("override level indices do not match").
- **Check**: any of these raises `Error::Invalid` from single-pass `parse_numbering`;
  `Document::list_items()` and `numbering_snapshot()` fail wholesale — even for lists
  unrelated to the offending abstractNum.
- **Spec says**: ST_NumberFormat/ST_LvlSuffix are closed enums (schema-backed); duplicate
  abstractNumId is an ID-uniqueness rule Word does not enforce on read.
- **Real-world behavior**: reproduced all four variants (`numFmt="hangulChosung"`,
  duplicate `abstractNumId=0`, `isLgl w:val="yes"`, `suff w:val="colon"`). LibreOffice
  maps unknown numFmt to a fallback and treats any non-{true,True,1,on,On} on/off value
  as false (OOXMLPropertySet.cxx:322-329); POI never validates; Word substitutes
  decimal/tab defaults.
- **Suggested posture**: default-repair per-node: unknown numFmt → record + decimal (or
  `Custom` with preserved token); unknown suff → tab; lvlRestart > 9 → clamp/ignore;
  unknown on/off → false (LO semantics); duplicate IDs → first-wins; all recorded. Fail
  only on structurally unparseable XML.
### 2.5 Tracked revision without `w:author` kills the entire revisions listing — P1
- **Location**: `crates/litchi-docx/src/revision.rs:771` ("tracked revision requires a
  Word author attribute"); same batch hard-fails on invalid `w:date` (:679-684), non-UTC
  `w16du:dateUtc` (:626-646), non-integer `w:id` (:665-677).
- **Check**: `Error::InvalidFormat` from `revision_from_element`, failing
  `Document::revisions()` as one batch.
- **Spec says**: ECMA-376 CT_TrackChange declares `author` and `date` **optional** — only
  `id` is required.
- **Real-world behavior**: reproduced — `<w:ins w:id="1" w:date="…">` without author
  fails the whole listing. Privacy-scrubbed documents and third-party producers omit
  author; Word shows the change with an empty/"unknown" author.
- **Suggested posture**: default-repair — model author as `Option<String>` (or empty),
  record `MissingRevisionAuthor`; degrade invalid dates to preserved-raw/None with
  record.
### 2.6 Paragraph `w:numPr` with only `w:ilvl` errors — and takes `style_id()` down with it — P1
- **Location**: `crates/litchi-docx/src/paragraph/codec/paragraph_properties.rs:181`
  ("paragraph numPr is missing numId"; empty-`numPr` fast path at :130-133).
- **Check**: a `numPr` lacking `numId` raises `Error::InvalidFormat`; because
  `list_properties()` returns `(numbering, style_id)` from one scan,
  `Paragraph::style_id()` on the same paragraph also fails.
- **Spec says**: ECMA-376 CT_NumPr declares both `ilvl` and `numId` with minOccurs=0;
  `ilvl` alone is a level override over style-inherited numbering.
- **Real-world behavior**: reproduced — both `numbering()` and `style_id()` fail on
  `<w:numPr><w:ilvl w:val="2"/></w:numPr>`. Word/LibreOffice resolve numbering from the
  paragraph style with the ilvl override; LO and RTF converters emit ilvl-only
  overrides.
- **Suggested posture**: default-repair — return `numbering: None` when numId is absent;
  split the scan results so one property's failure cannot poison the other; record
  `NumPrMissingNumId`.
### 2.7 Section column without `w:w` (and column-model validation) kills section layout parsing — P1
- **Location**: `crates/litchi-docx/src/section/codec.rs:1128` ("section column omits
  required width"); validation at :161-177 (count must be 1..=45; unequal columns require
  exactly one `w:col` per `w:num`).
- **Check**: `parse_columns` requires every `w:col` to carry `w:w`, then
  `validate_columns` enforces count range and cardinality; any violation fails the
  whole-document section model.
- **Spec says**: ECMA-376 CT_Col declares both `w` and `space` **optional**;
  ST_DecimalNumber for `w:num` has no 1..=45 bound. Schema contradicts the check.
- **Real-world behavior**: reproduced — `<w:col w:space="720"/>` fails the whole section
  snapshot. Word/LO compute equal or default widths when omitted; WPS and template
  generators emit sparse `w:col` elements.
- **Suggested posture**: default-repair — default missing width to 0/equal-share and
  record; clamp or accept out-of-range `w:num` with record.
### 2.8 Content-control inventory: one malformed `w:sdtPr` kills the whole `content_controls()` listing — P1 (confidence: high on code; not separately reproduced)
- **Location**: `crates/litchi-docx/src/content_control.rs:1418-1432` ("invalid
  content-control lock value" on unknown/missing `w:lock@val`), :1461-1469 (dataBinding
  requires `xpath` + `storeItemID`), :1509-1511, :1517-1523, :1546, :1560-1567
  (cardinality, duplicate/missing properties, invalid on/off); single-pass driver at
  :772-900.
- **Check**: the inventory scans the entire document in one pass; any property error
  aborts the whole inventory (`doc.content_controls()`).
- **Spec says**: CT_SdtLock `val` is optional; CT_SdtDataBinding `xpath`/`storeItemID`
  are optional. Required-attribute checks have no schema basis.
- **Real-world behavior**: SDTs are ubiquitous (placeholders, TOCs, bibliographies);
  Word/LO render controls with missing/garbage lock values as plain content; POI reads
  them without validation.
- **Suggested posture**: default-repair per control — degrade the offending control to an
  untyped occurrence (keep range and content), record the reason; never abort the
  inventory; unknown lock values → unlocked with record.
### 2.9 `ilvl` value > 8 hard-errors paragraph and style numbering — P2 (confidence: high on behavior; medium on frequency)
- **Location**: `crates/litchi-docx/src/paragraph/codec/paragraph_properties.rs:108` and
  :149 ("invalid paragraph ilvl '{raw}'"); styles at `styles.rs:424-427` ("invalid style
  outlineLvl '{raw}'").
- **Check**: `w:ilvl w:val` rejected if > 8; style-level `outlineLvl` > 9 fails the whole
  style-list load used by `resolved_numbering`.
- **Spec says**: ilvl `w:val` is ST_DecimalNumber — unbounded in schema; 0-8 is a Word-UI
  convention.
- **Real-world behavior**: reproduced — `ilvl="9"` fails `numbering()` and `style_id()`.
  Word clamps/ignores out-of-range levels; buggy generators produce these.
- **Suggested posture**: default-repair — clamp to 8 (or treat as no-direct-level) and
  record `IlvlOutOfRange`; for styles, skip the offending property rather than failing
  the style table load.
### 2.10 Drawing shape inventory rejects unknown `a:prstGeom@prst` values — P2 (confidence: high on behavior; medium on frequency)
- **Location**: `crates/litchi-docx/src/drawing/codec.rs:357` ("invalid DrawingML shape
  preset '{value}'"); rejection asserted as desired behavior in drawing/tests.rs:12-15.
- **Check**: `prst` missing or not in the closed `Preset` enum raises
  `Error::InvalidFormat`, failing `Paragraph::drawing_objects()` for that paragraph — one
  odd shape suppresses all drawings in the paragraph (per-paragraph scope caps the blast
  radius).
- **Spec says**: ST_ShapeType is a closed list (schema-backed), but consumers are
  expected to ignore unrecognized geometry; `custGeom` interop means non-standard tokens
  occur.
- **Real-world behavior**: WPS and third-party producers emit non-standard tokens (the
  test itself uses `prst="customShape"`); LibreOffice falls back to a default shape; Word
  renders best-effort.
- **Suggested posture**: default-repair — model `Preset::Unknown(String)` (preserve
  token) or degrade `preset` to `None` keeping the rest of the object; record.
### 2.11 `w:fldSimple` missing `w:instr` kills the entire field listing — P2
- **Location**: `crates/litchi-docx/src/field/codec.rs:1366` ("DOCX simple field is
  missing w:instr").
- **Check**: `Error::Invalid` from the single-pass extractor; `doc.fields()` returns no
  fields at all, including well-formed ones earlier in the document.
- **Spec says**: CT_SimpleField `instr` is use=required (schema-backed), but dropping
  every field in the document is disproportionate.
- **Real-world behavior**: reproduced — one bad `fldSimple` after a good one yields zero
  fields. Word renders nothing for the broken field and keeps the rest; POI returns null
  instr without failing. (Balanced design nearby: unterminated quoted instructions kept
  raw, unknown field types get opaque models.)
- **Suggested posture**: default-repair — skip the offending field (or model it with an
  empty instruction), record `FldSimpleMissingInstr`.
### 2.12 Section header/footer reference rejects non-XML-ID relationship IDs — P2 (confidence: high on code; medium-low on frequency)
- **Location**: `crates/litchi-docx/src/section/codec.rs:1148` (`is_xml_id` check on
  `headerReference`/`footerReference` `r:id`; helper at :1159).
- **Check**: `r:id` must be non-empty and start with a letter/underscore; otherwise
  `Error::InvalidFormat`, failing section parsing.
- **Spec says**: ST_RelationshipId is xsd:ID (schema-backed), but OPC relationship lookup
  is by string equality — the IDNESS constraints are never needed to resolve the
  reference.
- **Real-world behavior**: sloppy producers use numeric ids (`"0"`, `"1"`); Word and LO
  resolve them fine since .rels lookup is literal.
- **Suggested posture**: default-repair — accept any non-empty string as the relationship
  id; record `NonConformingRelationshipId`.

## 3. litchi-xls + litchi-biff + litchi-xlsb (legacy Excel)

Infrastructure: `XlsLeniency`/`FormattingDefect` exist but reach only five cosmetic
defects behind opt-in `Leniency::TolerateFormattingDefects`, and the source-backed open
hardcodes `Leniency::Strict` (`workbook/source.rs:2114`). litchi-biff is clean (pure
framing layer; unknown records preserved). litchi-xls is tolerant in the right places
(unknown records skipped/preserved, ExtSST lazy errors, malformed
Dimensions/MergeCells/AutoFilter ignored, duplicate cells tolerated, out-of-range
`LabelSst` → error cell, infallible RK decoding) — the findings are the remaining hard
errors.
### 3.1 Per-cell XF cross-reference validation kills the whole worksheet — P0
> **Status: FIXED (2026-10-01)** — `validate_cell_xf` replaced by a never-fatal check that
> substitutes XF index 0x0F and records `CellFormatDefect::XfReference` (new always-applied
> enum, not the opt-in `FormattingDefect` as suggested below); leniency also threaded into
> the source-backed open via `SourceBackedOptions` + `SourceBackedWorkbook::tolerance_report()`.
- **Location**: `crates/litchi-xls/src/number_format/codec.rs:27-51`
  (`Formatting::validate_cell_xf` — "cell references reserved style-XF slot" :37); called
  per cell at `workbook/codec/semantic/worksheet.rs:691`, :847, :869, per column/row at
  `layout/mod.rs:98,111`, and six `workbook/source.rs` sites (:2956 … :4325).
- **Check**: a cell record's `ixfe` must be ≥ 15, name a *cell* XF, and be within the
  parsed XF table — otherwise `Error::InvalidRecord` aborts the sheet (and workbook).
- **Spec says**: [MS-XLS] 2.5.115 CellIXF "specifies a cell XF" — a producer obligation;
  no reader-side refusal defined.
- **Real-world behavior**: Excel renders with default formatting; POI does not validate
  `ixfe` at read time; LibreOffice clamps to the default XF. Out-of-range `ixfe` is one
  of the most common non-Microsoft-producer defects (Gnumeric, Lotus, stream-cut files).
- **Suggested posture**: default-repair — substitute XF index 0x0F (the mandatory first
  cell XF), record as a new `FormattingDefect::CellXfReference`. Values untouched; only
  styling falls back, exactly as in Excel.
### 3.2 Unsupported CODEPAGE record value fails the entire workbook — P1
- **Location**: `crates/litchi-xls/src/workbook/source.rs:1997-2006` and
  `workbook/codec/semantic/globals.rs:199-204`, via `records.rs:307-313`
  (`Encoding::from_codepage` → `Mbcs::require`; error text at
  `crates/litchi-codepage/src/lib.rs:41`; pinned by test records.rs:2216-2219).
- **Check**: any `CODEPAGE` value without an exact codec in `Mbcs` (437, 850, 861,
  10029, 20127, 28591–28599, 1201, 12000/12001) raises `Error::Encoding` at open.
- **Spec says**: [MS-XLS] 2.4.52 — `cv` "MUST be one of the code page values in [CODEPG]
  or 1200". 437/850/20127 *are* valid [CODEPG] values; litchi-codepage has no codec.
- **Real-world behavior**: Excel falls back to the system ANSI page (BIFF8 text is
  Unicode anyway); POI stores the value with no codec requirement. DOS-era converters,
  Lotus 1-2-3 exports, and some Far East tools still emit 437/850/10029.
- **Suggested posture**: default-repair — keep the raw value, fall back to cp1252,
  record. `Encoding::Codepage` is nearly unused on the read path (records.rs:323-352), so
  the strictness buys almost nothing.
### 3.3 Worksheet-name rules enforced on read (BoundSheet8) — P1 (confidence: medium-high)
- **Location**: `crates/litchi-xls/src/records.rs:219-224` + `sheet_name_fault` at
  :269-287; duplicate case-insensitive names rejected at `workbook/package.rs:173` and
  `workbook/source.rs:2103-2113`.
- **Check**: read-time rejection of empty names, names > 31 UTF-16 units, names
  containing `: \ * ? / [ ]` or NUL/U+0003, leading/trailing apostrophes, and duplicate
  case-insensitive names; any aborts the open.
- **Spec says**: [MS-XLS] 2.4.28 BoundSheet8 `stName` — writer-side "MUST NOT"
  constraints; no reader behavior specified (records.rs:264-268 shows deliberate
  writer/reader symmetry, not a spec mandate).
- **Real-world behavior**: POI's `BoundSheetRecord` performs no name validation on read
  (BoundSheetRecord.java:75-81); LibreOffice opens and auto-repairs. WPS, iWork-exported
  `.xls`, and third-party converters produce over-long or illegal-character names and
  case-only duplicates.
- **Suggested posture**: default-repair for length/forbidden-char/apostrophe (keep the
  name verbatim, record). Duplicates: keep both (sheets addressable by index), record;
  opt-in at most.
### 3.4 Palette record: reserved byte, exact count, duplicate rejection — P1
- **Location**: `crates/litchi-xls/src/palette.rs:155-196` (exact length :156-161,
  `color_count != 56` :162-167, reserved byte :172-177, duplicate :186-195).
- **Check**: the `Palette` record must be exactly 226 bytes, declare exactly 56 colors,
  have zero in every `LongRGB` reserved byte, and appear at most once — any violation
  aborts the open.
- **Spec says**: [MS-XLS] 2.5.177 LongRGB — "reserved (1 byte): MUST be zero, and MUST be
  ignored." 2.4.188 — `ccv` "MUST be 56" (producer obligation). Same directive as the
  cell-border repair.
- **Real-world behavior**: POI reads `numcolors` as-is and ignores the unused byte
  (`PaletteRecord.java:176` — `in.readByte(); // unused`); Excel/LibreOffice tolerate
  duplicates (later wins). Older producers emit short palettes; RGBA tools leave nonzero
  alpha bytes.
- **Suggested posture**: default-repair — mask/ignore the reserved byte (record); accept
  `2 + ccv*4`-sized payloads, fill remainder from the default palette; last-record-wins
  for duplicates with a recorded defect.
### 3.5 Font record scalar-field validation far beyond the tolerated defects — P1
- **Location**: `crates/litchi-xls/src/font.rs:196-227` (height, Icv, weight, escapement,
  underline :221-227), charset :250-275, name length :287-291, NUL-in-name :322-324,
  `validate_font_table` ordering/count :367-384.
- **Check**: `dyHeight ∉ {0, 20..=8191}`, invalid Icv, `weight ∉ {0, 100..=1000}`,
  escapement > 2, underline byte outside {0,1,2,0x21,0x22}, charset byte outside the
  19-value enumeration, `cch > 31`, NUL in the name, fewer than 4 fonts, or fonts out of
  order — all hard errors *even under `Leniency::TolerateFormattingDefects`* (covers only
  `bFamily` and `cch == 0`).
- **Spec says**: [MS-XLS] 2.4.122 Font — each field has a "MUST be" enumeration; all
  producer obligations.
- **Real-world behavior**: POI `FontRecord` reads underline/family/charset as raw bytes
  with zero validation (`FontRecord.java:98`); Excel substitutes defaults. The leniency
  header itself documents such files as "widespread real-world workbooks"
  (`crates/litchi-xls/src/leniency.rs:1-10`).
- **Suggested posture**: extend `FormattingDefect` per field: clamp height/weight (or
  zero); unknown escapement → Normal, underline → Single, charset → Default, Icv →
  0x7FFF; keep name bytes verbatim (drop the NUL check on read); index the font table
  as-read; record each.
### 3.6 BoundSheet8 reserved string bits, visibility value 3, unknown sheet type — P1
- **Location**: `crates/litchi-xls/src/records.rs:198-209` (reserved option bits + exact
  name-length equality), :136-138 (`hsState == 3` rejected), :162-164 (`dt ∉ {0,1,2,6}`
  rejected).
- **Check**: `stName.grbit & 0xFE != 0` → `Error::InvalidRecord`; payload length must
  equal `8 + cch*width` exactly; visibility 3 and unknown sheet type are errors.
- **Spec says**: [MS-XLS] 2.5.294 — "reserved (7 bits): MUST be zero, and MUST be
  ignored"; 2.4.28 — "unused (6 bits): Undefined and MUST be ignored". The crate's own
  `Format` and SST parsers already ignore the identical bits — BoundSheet8 is an
  inconsistent outlier.
- **Real-world behavior**: POI reads the option byte and only tests bit 0
  (BoundSheetRecord.java:75-81); LibreOffice treats unknown `hsState` as hidden and
  unknown `dt` as worksheet; trailing bytes tolerated by both.
- **Suggested posture**: default-repair — mask reserved option bits (record, same class
  as Format/SST); clamp `hsState` → `VeryHidden`, unknown `dt` → `WorkSheet` (recorded);
  ignore trailing payload bytes.
### 3.7 XF alignment / border / fill enumeration strictness — P1
- **Location**: `crates/litchi-xls/src/alignment.rs:140-176` (valign 5–7, rotation
  181–254, reserved indent bit 0x20, "XF reading order {value} is invalid" :175);
  `crates/litchi-xls/src/border_fill.rs:44` ("reserved BIFF8 border style {value}" > 13),
  :174 (fill pattern > 18), :226-230 (diagonal contradiction), :239-253 (color index >
  0x41).
- **Check**: any of the above aborts the whole open in *both* leniency modes; only the
  `fJustLast` contradiction was given a repair (alignment.rs:149-158).
- **Spec says**: [MS-XLS] 2.4.353 XF — enumerated "MUST be" tables; the indentation
  reserved bit is "MUST be zero, and MUST be ignored" — erroring on it contradicts the
  spec's reader directive outright.
- **Real-world behavior**: LibreOffice maps unknown enums to defaults (bottom alignment,
  no border, no fill); POI stores raw bytes/shorts. Diagonal-flag contradictions and
  bogus Icv values appear in WPS and older third-party files.
- **Suggested posture**: default-repair — mask the reserved indent bit (record); valign
  5–7 → Bottom; rotation 181–254 → None; reading order 3 → Context; border style > 13 →
  None; fill pattern > 18 → None; diagonal contradiction → keep style, derive direction;
  out-of-range Icv → 0x7F. Each recorded.
### 3.8 One unpaired UTF-16 surrogate (SST, Label, font name) fails the whole workbook — P1
- **Location**: `crates/litchi-xls/src/records.rs:949-951` (SST via
  `String::from_utf16`), `utils.rs:99` (Label), `font.rs:309-318` (font name), ExtRst
  :1091-1092; companion refusals: formatting runs "not strictly increasing" :1010-1014,
  ExtRst run validation :1062-1066, :1099-1129; error text pinned by corpus test
  :1474-1475.
- **Check**: one lone surrogate in any of potentially millions of shared strings aborts
  the entire SST scan — the whole open; a single out-of-order formatting run has the same
  effect.
- **Spec says**: no spec basis — [MS-XLS] 2.5.293/294 store UTF-16 code units without
  defining reader behavior for malformed sequences.
- **Real-world behavior**: Java/POI decode via `new String(bytes, UTF_16LE)` (substitutes
  U+FFFD); LibreOffice the same; Excel displays garbage and opens. Buggy exporters that
  truncate UTF-16 mid-surrogate-pair are a known wild occurrence.
- **Suggested posture**: default-repair — decode with replacement (U+FFFD), record the
  string index; drop offending formatting runs (or sort-and-clip to `cch`) and record.
  Every *well-formed* string and the cell grid survive.
### 3.9 Number-format ID cross-reference and XFCRC reserved-field strictness — P1 (XFCRC/ifnt parts P2; confidence: medium-high)
- **Location**: `crates/litchi-xls/src/number_format/codec.rs:180-199` (reserved format
  ID :184, "references missing custom number format" :193), :231-250 (DXF), :299-304
  (`Format` record's own `ifmt`), :412-414 (style XF `f123Prefix`), :462-486 (XFCRC
  header/reserved/count), `font.rs:360-364` (XF `ifnt == 4` or > 1022).
- **Check**: all hard errors in both leniency modes — even though the same function
  tolerates an `XFCRC.cxfs` disagreement under leniency (codec.rs:150-168), so a workbook
  whose XFCRC reserved bytes are nonzero fails while one with a wrong count passes.
- **Spec says**: [MS-XLS] 2.4.126/2.4.353/2.4.354 — producer-side "MUST" tables; XFCRC
  reserved fields follow the standard "MUST be zero, and MUST be ignored" pattern.
- **Real-world behavior**: Excel renders cells whose `ifmt` has no `Format` record with
  General; POI's style lookup is lazy, never validating referential integrity at open.
  Stream-cut files and files carrying `ifnt == 4` open fine in Excel.
- **Suggested posture**: default-repair — unresolved/reserved `ifmt` → General (id 0),
  record; `f123Prefix` on style XF → mask + record; XFCRC reserved/header/count-range →
  ignore + record; `ifnt == 4`/out-of-range → substitute font 0 + record.
### 3.10 litchi-xlsb: cached numeric values rejected unless finite, normal, non-negative-zero — P1
- **Location**: `crates/litchi-xlsb/src/cell_values/worksheet.rs:1974-1983`
  (`validate_xnum`), called for `BrtCellReal` at :1723 and the `BrtFmlaNum` cache at
  :1742; same policy at `sparkline/model.rs:1357-1372` and
  `volatile_dependencies/model.rs:465-477`.
- **Check**: NaN, ±Inf, subnormal, or negative-zero `Xnum` in any single cell aborts the
  entire worksheet snapshot with `Error::UnsupportedFeature`.
- **Spec says**: [MS-XLSB] 2.5.153 Xnum — the value "MUST NOT" be
  infinity/NaN/denormalized/negative-zero; a producer obligation with no reader refusal.
- **Real-world behavior**: Excel opens such files (shows the stale cache or
  recalculates); classic producers are add-ins/RTD feeds caching NaN and third-party
  writers serializing f64 verbatim. One bad float kills every other cell on the sheet.
- **Suggested posture**: per-cell degradation (that cell → error/blank, keep the sheet)
  with the bit pattern recorded; opt-in leniency at minimum since a NaN cache is genuinely
  ambiguous, but default-repair to per-cell degradation is defensible given Excel opens
  the file.
### 3.11 litchi-xlsb: systemic reserved-bit / strict-layout / trailing-byte refusals — P1
> **Status: FIXED (2026-10-02)** — litchi-xlsb gained its `ToleranceReport` carrier (`Workbook`/`Package`/`Worksheet`/`SourceBackedWorkbook::tolerance_report()`); reserved bits on Cell,
> BrtRowHdr, BrtColInfo, GrbitFmla, BrtBorder, Blxf, BrtCalcProp, BrtName are masked and trailing bytes ignored, recorded as `ReservedFieldDefect`/`LayoutDefect`; truncation and limits stay fatal.
- **Location**: `crates/litchi-xlsb/src/cell_values/worksheet.rs:1671-1677` ("nonzero
  reserved Cell flag bits" on *every* cell record); `styles/codec.rs:213-218`
  (`BrtBorder` flags), :305-312 (`Blxf` style/reserved), :188-193 (`BrtColor`
  `fValidRGB = 0`), :243-262 (`BrtFmt`); `calc/model.rs:352-359` (`BrtCalcPr` reserved
  bits); `raw/error.rs:119-128` + the `cursor.finish()?` pattern (e.g.
  `named_ranges/codec.rs:127-129`, `workbook/codec/records.rs:617-618`) erroring on
  trailing payload bytes.
- **Check**: all abort the containing part (styles, workbook, or the whole sheet) on
  out-of-spec reserved bits or extra bytes.
- **Spec says**: [MS-XLSB] 2.5.10 Cell — "reserved (7 bits): MUST be 0 and MUST be
  ignored"; 2.5.5 Blxf — "reserved (1 byte): MUST be 0 and MUST be ignored". More
  generally 2.1.4/2.1.6 build the format around the future-record architecture:
  applications "can ignore but preserve" records/fields they do not support.
- **Real-world behavior**: Excel reads BIFF12 with ignore-and-advance semantics per record
  length; files written by newer Excel versions or third-party BIFF12 writers trip these
  checks. The `Cell` flag check is the most exposed: it runs on every cell of every
  sheet.
- **Suggested posture**: default-repair — mask reserved bits (record each site class once
  with counts); `fValidRGB = 0` → "no color"; replace `cursor.finish()?` hard errors with
  "ignore remaining bytes + record" on the read path (keep on the write path); unknown
  border style → None.
### 3.12 litchi-xlsb: strict Boolean bytes, unknown BErr codes, strict UTF-16, DVal reserved fields — P2 (confidence: medium)
- **Location**: `crates/litchi-xlsb/src/cell_values/worksheet.rs:2789-2797` (`parse_bool`),
  :280-294 (`CellError::from_code`); `raw/cursor.rs:164-173` (`InvalidBool`), :181-209
  (`InvalidUtf16` at :203); `data_validation/codec.rs:33-44`, :150-161 (`BrtDVal`/
  `BrtDVal14` unknown enums/reserved bits), reached eagerly at
  `host/cells_reader/codec.rs:560`, :571.
- **Check**: a boolean byte other than 0/1, an error code outside the nine defined
  values, an unpaired surrogate in any wide string, or one malformed data-validation
  record each aborts the containing read.
- **Spec says**: `BrtCellBool.bBoolean` and `BErr` are enumerated producer obligations;
  `BrtDVal`/`BrtDVal14` reserved bits carry "MUST be zero and MUST be ignored"; no basis
  for refusing malformed UTF-16.
- **Real-world behavior**: Excel treats any nonzero byte as TRUE, renders unknown error
  codes as `#NULL!`, decodes strings with replacement; an unrecognized validation rule
  simply never fires. Rare individually, but each is a one-record-kills-the-part failure.
- **Suggested posture**: default-repair — nonzero boolean → TRUE (record); unknown `BErr`
  → `#NULL!` (0x00) with raw byte recorded; wide strings → U+FFFD lossy decode (record);
  DVal unknown enums → default (AnyValue/Stop), reserved bits → mask (recorded); degrade a
  single bad validation rather than failing the sheet.

## 4. litchi-xlsx (SpreadsheetML .xlsx)

Infrastructure: **no** Leniency/ToleranceReport mechanism at all — every finding presumes
introducing one modeled on `crates/litchi-doc/src/leniency.rs`. Graceful postures already
present and worth using as the model: unknown cell types → `Cell::Unknown`
(semantic.rs:41-51), unknown formula types → `RawFormulaKind::Unknown`, unknown
visibility → `Visibility::Unknown`, unknown error values → `ErrorValue::Unknown`,
`dimension` stored but never cross-checked, MCE/Strict namespaces accepted, foreign
elements ignored.
### 4.1 Row/column geometry domain clamping fails the whole sheet — P1
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/codec.rs:928` (row `ht` via
  `row::Height::new`), `row.rs:18-25` (0..=409 domain), `column.rs:17-23` (width
  0..=255), `raw/worksheet/validation.rs:33-48` (`sheetFormatPr` defaults).
- **Check**: `ht` outside 0..=409 points, column `width` outside 0..=255, or a
  negative/non-finite `defaultRowHeight`/`defaultColWidth` aborts the entire worksheet
  parse.
- **Spec says**: ECMA-376 §18.3.1.73 (`row`) — `ht` is a plain `xsd:double`; the schema
  imposes no cap. 409 is an application authoring limit (Excel's UI max is actually
  409.5, which this check also rejects).
- **Real-world behavior**: LibreOffice reads `ht` as a raw double with no validation
  (sheetdatacontext.cxx:303) and its own limit is ~100 cm, so LO-produced files
  legitimately contain `ht` > 409; Excel clamps; POI stores the raw double. Tall rows
  (pasted wrapped text, converted ODS) are occasional but real.
- **Suggested posture**: default-repair — clamp to the domain edge or drop the explicit
  height and record; strict rejection only for non-finite values under an opt-in strict
  mode.
### 4.2 Style-index cross-reference validation kills the whole sheet — P1
- **Location**: `crates/litchi-xlsx/src/workbook/source.rs:1436-1481` (`validate_styles`
  — "worksheet cell {} references shared style {}, but the workbook contains {len} cell
  formats" :1456; rows :1461-1470, columns :1472-1481), pre-check at
  `raw/worksheet/codec.rs:1163` (`s > MAX_CELL_STYLE` → Err).
- **Check**: any cell/row/column whose style index is ≥ the `cellXfs` count (or >
  65,430) fails the entire sheet.
- **Spec says**: ECMA-376 §18.3.1.4 (`c/@s`) requires the index to reference `cellXfs`;
  no basis for rejecting the document when the reference is stale.
- **Real-world behavior**: stale style indices are a classic artifact of third-party
  editors (WPS, report generators, row-deletion tools that don't prune `cellXfs`). Excel
  repairs silently substituting the default format; LibreOffice clamps; POI defers the
  lookup to access time.
- **Suggested posture**: default-repair — substitute style 0 (the mandatory default xf)
  and record one entry per affected range; keep `MAX_CELL_STYLE` as an allocation guard
  only.
### 4.3 Malformed cell-value lexicals (number / boolean / date) kill the whole sheet — P1
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/semantic.rs:127` (`Number::new`),
  `cell.rs:402-417` (f64 parse + finite check at :413-414), semantic.rs:151-154
  (boolean), `cell.rs:244-256` (`Date::new` ISO 8601 only).
- **Check**: a `<v>` that doesn't parse as a finite f64 (`1,234.56`, `1E+999`, `INF`,
  locale-formatted), a boolean not in {`0`,`1`,`true`,`false`} (e.g. `TRUE`, `2`), or a
  `t="d"` date outside the three accepted layouts aborts the whole sheet.
- **Spec says**: ECMA-376 §18.3.1.4/§18.18 — lexical domains are defined, but nothing
  mandates failing the document on a bad instance.
- **Real-world behavior**: LibreOffice converts with lossy `toDouble()`/`toInt32()` and
  explicitly tolerates boolean text — "Some generators may write true or false instead
  of 1 or 0" (sheetdatacontext.cxx:204-233, comment at :211); POI retains the raw string
  and only throws if the caller requests the numeric value.
- **Suggested posture**: default-repair per cell — demote an unparseable `n` to
  `Cell::Unknown`/error value (preserving the lexical), parse booleans
  case-insensitively plus non-zero-numeric-is-true, demote bad `t="d"` to text; record
  each demotion.
### 4.4 Shared-string index out of range or missing table kills the whole sheet — P1
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/semantic.rs:139-148` ("worksheet
  uses shared strings but the workbook has no shared-string part" :141; "shared-string
  index {index} exceeds table length {}" :143-146).
- **Check**: `t="s"` with an index ≥ table length, or any `t="s"` cell when the workbook
  has no sharedStrings part, returns `Err` for the entire sheet.
- **Spec says**: ECMA-376 §18.3.1.4 — the index must reference the table; no basis for
  document-level failure on a dangling reference.
- **Real-world behavior**: dangling references occur from string-table-pruning tools and
  some WPS/third-party writers. Excel repairs and substitutes blank/`#REF!`; LibreOffice
  resolves out-of-range lookups to empty (sheetdatacontext.cxx:228); POI surfaces the
  problem only on access.
- **Suggested posture**: default-repair — substitute an empty string (Excel-repair
  semantics), record the cell address; record (don't fail) the missing-table case per
  cell.
### 4.5 Malformed `sst` count/uniqueCount hints are fatal in the production parser — P1
- **Location**: `crates/litchi-xlsx/src/raw/strings.rs:646-654` (`reserve_hint`
  propagates), :773-788 (`parse_count_hint_value` rejects non-numeric, negative, >
  2³¹−1); contrast `raw/shared_strings.rs:219-235` (tolerant twin, tests at :927-940).
- **Check**: the streaming shared-strings parser hard-fails on `uniqueCount="NaN"`,
  `count="-1"`, etc. The crate's *other* shared-strings parser explicitly documents the
  opposite: "a missing, unparseable, or contradictory hint is ignored rather than failing
  the workbook" (shared_strings.rs:219-225).
- **Spec says**: ECMA-376 §18.4.9 (`sst`) — `count`/`uniqueCount` are optional,
  informational attributes.
- **Real-world behavior**: producers occasionally emit non-numeric or stale hints; Excel
  and LibreOffice ignore them entirely. This is an internal self-contradiction.
- **Suggested posture**: default-repair — treat unparseable hints as absent (as
  shared_strings.rs already does); at most record a note.
### 4.6 Broken shared-formula groups kill the whole sheet — P1 (confidence: medium)
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/semantic.rs:201-260`
  (`resolve_shared_formulas`: master missing `ref` :203, no expression :209, not first
  in range :217, duplicate master :235, follower with no master :243, member outside
  master range :249, follower carrying master data :256).
- **Check**: any structural inconsistency in a shared-formula group aborts the whole
  sheet parse.
- **Spec says**: ECMA-376 §18.3.1.40 (`f`, `t="shared"`) defines the master/slave
  protocol; nothing prescribes failing the document over a malformed group.
- **Real-world behavior**: orphan/partial groups are common after third-party row/column
  deletions. LibreOffice falls back — invalid shared reference → import as plain/empty
  formula (sheetdatacontext.cxx:161-165); Excel repairs; POI resolves lazily.
- **Suggested posture**: default-repair — keep master text verbatim; orphan followers →
  empty/`Unknown` formula instead of translating; record per group. Only "member outside
  master range" (translation would invent content) might stay strict.
### 4.7 `activeTab` / `localSheetId` out of range kills the whole workbook — P1
- **Location**: `crates/litchi-xlsx/src/raw/catalog.rs:202-209` ("workbook activeTab
  {active_tab} exceeds the {} available sheets" :207), :211-219 (definedName
  `localSheetId`).
- **Check**: `workbookView/@activeTab` ≥ sheet count, or any `definedName/@localSheetId`
  ≥ sheet count, fails the entire workbook catalog at open.
- **Spec says**: ECMA-376 §18.2.20 — `activeTab` is an unconstrained `unsignedInt`;
  §18.2.5 — `localSheetId` indexes sheets "in order of appearance"; no failure semantics.
- **Real-world behavior**: both go stale routinely when sheets are deleted by non-Excel
  tools. Excel clamps/ignores; LibreOffice ignores invalid local scopes; POI reads both
  without validation.
- **Suggested posture**: default-repair — clamp `activeTab` to the last sheet (or 0);
  demote out-of-range `localSheetId` names to workbook-global scope or drop just that
  name; record each.
### 4.8 Sheet-name authoring rules enforced at load time — P1 (confidence: medium)
- **Location**: `crates/litchi-xlsx/src/raw/catalog.rs:591-592` (invokes `validate_str`),
  `sheet.rs:190-216` (`validate_str`).
- **Check**: sheet names > 31 chars, containing `* ? / \ : [ ]`, NUL/U+0003, or with
  leading/trailing apostrophes are rejected while *parsing* `workbook.xml` — one
  non-conforming name fails the whole workbook open.
- **Spec says**: ECMA-376 §18.2.18 (`sheet/@name`) is `ST_Xstring` with no character or
  length constraints; the rules are Excel *application* constraints, not package-format
  rules.
- **Real-world behavior**: BI/reporting tools and hand-rolled generators emit > 31-char
  names; Excel repairs by renaming; LibreOffice truncates/sanitizes on import; POI
  accepts any name.
- **Suggested posture**: opt-in leaning default-repair — retain the name verbatim for
  reading (sheets are already keyed by position), record; enforce `validate_str` only at
  authoring/rename time.
### 4.9 Formula with leading `=` or empty expression kills the whole sheet — P2 (confidence: high on code; medium on frequency)
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/codec.rs:1388` ("worksheet formula
  must omit the leading '='"), :1384-1396 (empty expression error).
- **Check**: a scalar/array `<f>` whose text starts with `=` or is empty aborts the
  sheet.
- **Spec says**: ECMA-376 §18.3.1.40 — formula text excludes `=`; an empty formula is a
  schema-valid (if meaningless) string; no failure semantics mandated.
- **Real-world behavior**: naive generators periodically write `<f>=SUM(A1:A2)</f>`;
  LibreOffice takes the string as-is and its compiler accepts the leading `=`
  (sheetdatacontext.cxx:120,144-146 show empty formulas are skipped, not fatal); Excel
  repairs.
- **Suggested posture**: default-repair — strip one leading `=` (or tolerate and
  preserve); empty formula → `Kind::Unknown`/no-formula with a recorded note.
### 4.10 Worksheet sequencing and advisory-metadata strictness — P2 (confidence: high for dimension/counts; medium for out-of-order rows)
- **Location**: `crates/litchi-xlsx/src/raw/worksheet/codec.rs:771-789` ("worksheet
  dimension appears after column or cell data" :775, "duplicate dimension elements" :779,
  `ref` required), :915-926 (duplicate/out-of-order rows → Err), :992-996/:1012-1014
  (`sheetFormatPr`/`cols` after `sheetData` → Err), :818-828 (`mergeCells` placement),
  :1331-1348 (empty `<cols/>`/`<mergeCells/>`, merge-count mismatch → Err).
- **Check**: schema sequence violations and advisory count mismatches abort the whole
  sheet — including `dimension` (purely informational), a merge count that disagrees with
  the records, and descending/duplicated `row r` values.
- **Spec says**: ECMA-376 §18.3.1.35 — `dimension` is optional, "for informational
  purposes", applications recalculate; §18.3.1.55 — merge `count` is informational; the
  CT_Worksheet sequence defines canonical order but prescribes no failure semantics.
- **Real-world behavior**: bogus/re-ordered `dimension` is endemic (many generators write
  `A1` unconditionally); Excel recalculates and ignores it; LibreOffice ignores it; POI
  stores rows in a sorted map, silently tolerating out-of-order and duplicate rows.
- **Suggested posture**: default-repair — accept `dimension` in any position, ignore
  unparseable `ref` (recompute), collect `cols`/`mergeCells` position-independently, drop
  count-mismatch errors (records authoritative, mirroring the crate's own sst-hint
  posture), sort-or-last-wins duplicate rows; record each repair.
### 4.11 Theme `clrScheme` strictness: positional slots, exactly 12, mandatory `lastClr`, strict hex — P2
- **Location**: `crates/litchi-xlsx/src/style/theme.rs:241-249` (slot must be the exact
  expected element in order — "a:clrScheme slot {} must be <a:{}>" :245), :276-296
  (`srgbClr/@val`, `sysClr/@lastClr` required, hex must be exactly 6 digits), :329-351
  (fewer than 12 slots or unterminated scheme → Err).
- **Check**: any ordering deviation, missing slot, missing `lastClr`, or malformed hex
  fails the entire theme parse.
- **Spec says**: ECMA-376 §20.1.6.9 (CT_ColorScheme) does define a fixed required
  sequence, but §20.1.2.3.32 (CT_SystemColor) makes `lastClr` **optional** — requiring it
  has no spec basis.
- **Real-world behavior**: LibreOffice dispatches color-scheme children by element name,
  order-independently, leaving defaults for missing slots (clrschemecontext.cxx:72-87).
  Impact currently limited: `Theme::parse` is a public API but not on the workbook-open
  path.
- **Suggested posture**: default-repair — match slots by name, default missing slots,
  missing `lastClr` → resolve-by-name-or-black, bad hex → slot-default; record each.
### 4.12 Shared-string rich-text micro-structure and `_xHHHH_` escape failures kill the whole table — P2 (confidence: medium-high)
- **Location**: `crates/litchi-xlsx/src/raw/shared_strings.rs:300`/:475 (run without
  `<t>` → Err), :328-331/:360-366 (mixed simple+rich text, "duplicate text in
  shared-string item" :366), :347 ("duplicate shared-string run properties")/:847-853,
  :420-452 (bad `sz`, unknown `u` value, malformed RGB); production twins at
  `raw/strings.rs:220-234`/:669-704; escape failures at `raw/strings.rs:809-832` and
  :452-500 ("unpaired high/low surrogate in SpreadsheetML escape").
- **Check**: a single structurally imperfect `<si>` — a run missing its text, a
  duplicated `<b/>`, an unknown underline token, a bad color, or an unpaired `_xD800_`
  escape — aborts the entire shared-strings parse, failing every string cell in the
  workbook.
- **Spec says**: ECMA-376 §18.4.8/§18.4.3 define the content models, but no failure
  semantics; §18.3 `ST_Xstring` escapes are producer-generated and reference decoders are
  lossy.
- **Real-world behavior**: LibreOffice ignores unknown/duplicate run properties and
  decodes escapes best-effort; Excel repairs individual items. Hand-assembled rich text
  from string-template generators hits these corners.
- **Suggested posture**: default-repair at item granularity — drop the offending run or
  property (keep the text), unknown underline → "underline present", drop bad colors,
  unpaired escapes → U+FFFD; record per item.

## 5. litchi-ppt (legacy .ppt)

Infrastructure: no `PptLeniency`/`ToleranceReport` yet — every repair implies introducing
one on the litchi-doc two-tier model. The framing layer is already tolerant (record loop
and container parsing resynchronize byte-by-byte on garbage — `offset += 1` at
parsers/parser.rs:110-116; unknown record types → `RecordType::Unknown` with raw payload
preserved; truncated containers degrade; text atoms decode lossily; notes index degrades
instead of failing). The semantic layer (slide directory, CurrentUser/UserEdit chain,
text-run/style parsing) is strict-to-fatal.
### 5.1 `StyleTextPropAtom` strict framing rejects files whose style coverage ≠ text length + 1 — P0
> **Status: FIXED (2026-10-02)** — the text-box read path uses the clamping POI/LibreOffice StyleTextPropAtom parser and records each clamp/early stop/dropped tail as
> `TextPropDefect::StyleTextPropCoverage` (run index + persist/shape ID); `Slide`/`SpeakerNotes::tolerance_report`, `Presentation::collect_tolerance_report`. Writer/reparse paths stay strict.
- **Location**: `crates/litchi-ppt/src/text_prop.rs:376` ("TextPFRun has invalid
  character coverage"), :408 (TextCFRun), :433 ("StyleTextPropAtom has trailing bytes"),
  via `parse_style_text_prop_atom_strict` (:358), called at `text_run/codec.rs:95`.
- **Check**: run `charactersCovered` must tile exactly `text_length + 1`, with no
  zero-length runs and no trailing bytes, else `Error::Corrupted` propagates into every
  text-box parse.
- **Spec says**: [MS-PPT] 2.9.40–2.9.41 — coverage must equal the character count plus
  one; nothing prescribes rejecting the document when a producer miscounts.
- **Real-world behavior**: miscounted coverage is one of the most common .ppt deviations.
  POI only *logs* and truncates (StyleTextPropAtom.java:254,286,299); LibreOffice clamps
  (svdfppt.cxx:5071,5110). The crate ships a clamping POI-style parser
  (`parse_style_text_prop_atom`, text_prop.rs:262) the read path never uses. One
  miscounted text box kills the slide's whole shape list (F5.11).
- **Suggested posture**: default-repair — route the read path through the lenient parser
  (clamp coverage, break on garbage, drop trailing bytes), record each clamp; keep the
  strict variant for the writer/reparse-verification path.
### 5.2 Reserved BulletFlags / PFWrapFlags bits are fatal instead of ignored — P0
> **Status: FIXED (2026-10-02)** — reserved BulletFlags/PFWrapFlags bits are masked to 0x000F/0x0007 and recorded as `TextPropDefect::{ReservedBulletFlagBits, ReservedWrapFlagBits}`
> (context = raw flags, plus persist/shape ID); defined bits still apply and raw words round-trip.
- **Location**: `crates/litchi-ppt/src/text_run/codec.rs:301` ("TextPFRun has reserved
  BulletFlags bits set"), :386 ("TextPFRun has reserved PFWrapFlags bits set").
- **Check**: any reserved bit set in the 2-byte `bulletFlags` or `wrapFlags` fields
  raises `Error::Corrupted`.
- **Spec says**: [MS-PPT] 2.9.22 BulletFlags — "reserved (12 bits): MUST be zero and MUST
  be ignored"; 2.9.25 PFWrapFlags — "reserved (13 bits): MUST be zero and MUST be
  ignored". Exactly the cell-border pattern.
- **Real-world behavior**: PowerPoint, POI and LibreOffice read only the defined bits;
  third-party producers (WPS, older converters) routinely leak garbage into the high
  bits. One such paragraph kills the whole slide's shapes.
- **Suggested posture**: default-repair — mask reserved bits (keep raw only for
  round-trip) and record, mirroring the litchi-doc `SprmDefect` model.
### 5.3 Text enum / ColorIndexStruct / indent-level range checks are fatal — P1 (confidence: high on propagation; medium on frequency)
- **Location**: `crates/litchi-ppt/src/text_run/codec.rs:328` ("invalid
  TextAlignmentEnum value"), :339, :350, :368 (font-alignment/direction/tab-type enums),
  :274 ("indent level exceeds the 0..=4 range"), :186/:311 (ColorIndexStruct indexes).
- **Check**: out-of-enumeration values and ColorIndexStruct indexes other than
  0x00–0x07/0xFE/0xFF raise `Error::Corrupted`.
- **Spec says**: [MS-PPT] 2.13 enumerations and 2.2.13 IndentLevel ("MUST be ≤ 0x0004")
  define the ranges; no recovery is prescribed, and the reference apps degrade.
- **Real-world behavior**: unknown alignment/tab values appear from newer or
  non-Microsoft producers; PowerPoint opens with default formatting; LO reads values and
  branches only on ones it knows (svdfppt.cxx:5124ff); POI leaves unknowns
  uninterpreted.
- **Suggested posture**: default-repair — unknown enum values → `None`/default, record;
  clamp indent level to 4 with a recorded repair.
### 5.4 Per-entry `SlidePersistAtom` validation fails the entire file — P1 (null/duplicate/length variants P2 on frequency)
- **Location**: `crates/litchi-ppt/src/slide/directory.rs:211` ("invalid SlidePersistAtom
  header or length"), :221 (null persistIdRef/slideId), :228 ("SlidePersistAtom cTexts
  exceeds 8"), :231/:236 (duplicate slideId / persistIdRef).
- **Check**: any deviation in *one* slide-list entry aborts
  `SlideDirectory::build_with_limits`, which runs eagerly in `Package::presentation()` —
  the whole document becomes unreadable.
- **Spec says**: [MS-PPT] 2.4.14.5 — recVer/recInstance/recLen MUST be 0/0/0x14, cTexts
  MUST be ≤ 0x08; slideId uniqueness implied by 2.2.26. All content-metadata checks, not
  framing.
- **Real-world behavior**: POI's `SlidePersistAtom` reads five u32s by offset with no
  validation (SlidePersistAtom.java:89); LibreOffice reads fields unconditionally and
  later skips entries whose persist ref is unusable (svdfppt.cxx:201,2466). Orphaned
  slide references after incremental saves are real.
- **Suggested posture**: default-repair — skip the offending entry (and its following
  text records) and record; duplicates keep first occurrence. Directory build should
  degrade per-entry, never fail.
### 5.5 Slide persist ID without a directory entry, or not pointing at a SlideContainer, is file-fatal — P1 (confidence: high on LO; medium-high on frequency)
- **Location**: `crates/litchi-ppt/src/slide/directory.rs:243` ("slide persist ID
  {persist_id} has no directory entry"), :251-258 ("persist ID … does not resolve to a
  SlideContainer").
- **Check**: every presentation-slide entry must resolve through the persist mapping to a
  well-formed `SlideContainer` header (type 1006, version 0x0f, instance 0), else
  `Error::Corrupted` for the whole file.
- **Spec says**: no basis beyond the general persist-directory grammar ([MS-PPT]
  2.3.4/2.5.1) — internal strictness.
- **Real-world behavior**: LibreOffice skips such pages silently
  (`if (nPersist > 0 && nPersist < m_nPersistPtrCnt)`, svdfppt.cxx:1576,2466-2468).
  Incrementally-saved files carry stale references to objects no longer in the live
  persist directory.
- **Suggested posture**: default-repair — drop the entry, record `StaleSlidePersistRef`;
  the presentation opens with the remaining slides, matching LO.
### 5.6 Unusable "Current User" stream or edit offset has no stream-scan fallback — P1
- **Location**: `crates/litchi-ppt/src/presentation/package.rs:109`
  (`Error::StreamNotFound("Current User")`), `slide/directory.rs:118` ("CurrentUser does
  not reference a valid UserEditAtom"), `embedded/object/editor/mapping/chain.rs:23`
  ("invalid UserEditAtom").
- **Check**: missing "Current User" stream, an edit offset of 0, or an offset not landing
  on a `UserEditAtom` (type 0x0FF5, length 28|32) aborts the open.
- **Spec says**: [MS-PPT] 2.3.2/2.3.3 define the header values as MUSTs; the recovery
  strategy is unspecified.
- **Real-world behavior**: LibreOffice treats the Current User stream as optional
  (pptin.cxx:129-133) and, when `offsetToCurrentEdit` is unusable, *scans the whole
  PowerPoint Document stream backwards for the last UserEditAtom*
  (svdfppt.cxx:1319-1337). POI throws like litchi, so files exercising this path are ones
  only LO/PowerPoint rescue today — typically stripped-down exports.
- **Suggested posture**: default-repair — if the stream is absent or the head offset
  doesn't validate, walk top-level records for the last `RT_UserEditAtom` (LO's exact
  algorithm) and record the recovery; relax the recLen check to ≥ 28 (POI pads short
  reads without a 28|32 gate).
### 5.7 Persist-directory decoding: wrong record type, zero-count run, or duplicate ID is fatal / drops the holder — P1
- **Location**: `crates/litchi-ppt/src/embedded/object/editor/mapping/directory.rs:15`
  ("invalid PersistDirectoryAtom"), :25-27 ("zero persist run");
  `persist/ptr_holder.rs:72` ("PersistDirectoryEntry has a zero entry count"), :90
  ("duplicate persist identifier"), :97 (alignment); the latter surfaces silently via
  `filter_map(|r| PersistPtrHolder::parse(r).ok())` at `persist/mapping.rs:53,84`,
  discarding *all* entries of that holder.
- **Check**: an `offsetPersistDirectory` pointing at any other record type, an entry
  header with count 0, or a repeated persist ID kills the UserEdit-chain read (open path)
  or voids the entire holder record (public `PersistMapping` path).
- **Spec says**: [MS-PPT] 2.3.5 implies count ≥ 1 by construction; 2.3.4 requires
  uniqueness within a directory — no rejection policy prescribed.
- **Real-world behavior**: LibreOffice only processes the directory when the target
  record *is* `PPT_PST_PersistPtrIncrementalBlock` and skips otherwise
  (svdfppt.cxx:1367); zero counts simply don't enter the read loop; duplicates are
  first-wins (svdfppt.cxx:1381). POI advances 4 bytes per header with no checks.
- **Suggested posture**: default-repair — skip non-matching directory records, skip
  zero-count headers, first-wins on duplicates (matching the chain-merge `or_insert`
  semantics already at directory.rs:31), record each; stop dropping the whole holder on
  one bad entry.
### 5.8 `CurrentUserAtom` fixed-field validation far exceeds POI and LibreOffice — P2 (username length P1)
- **Location**: `crates/litchi-ppt/src/current_user.rs:103` (recVer/recInstance/recType),
  :110 (fixed size != 20), :117-125 ("Invalid CurrentUser header token"), :128-132
  ("Invalid CurrentUser username length"); reached from the open path at
  `slide/directory.rs:107`.
- **Check**: any mismatch raises `Error::InvalidFormat` — file-fatal.
- **Spec says**: [MS-PPT] 2.3.2 — all four fields are MUSTs (token MUST be 0xE391C05F or
  0xF3D1C4DF; lenUserName MUST be ≤ 255), so spec basis exists.
- **Real-world behavior**: POI's `CurrentUserAtom.init()` validates *none* of these — for
  a garbage username length it logs "Invalid username length … treating as if there was
  no username set" (CurrentUserAtom.java:170-178); LibreOffice checks only `nRecType`
  (svdfppt.cxx:178). Only `offsetToCurrentEdit` and the token are load-bearing.
- **Suggested posture**: default-repair — ignore recVer/recInstance/size mismatches,
  clamp lenUserName to available bytes with lossy decode (POI posture), treat an unknown
  token as unencrypted *only* when no DocumentEncryptionAtom exists (record either way).
  Keep rejecting when the token claims encryption and decryption support is compiled out.
### 5.9 Cyclic `UserEdit` chain hard-errors instead of breaking or repairing — P2 (confidence: high; low wild frequency)
- **Location**: `crates/litchi-ppt/src/embedded/object/editor/mapping/chain.rs:17`
  ("cyclic or excessive UserEdit chain").
- **Check**: a repeated edit offset in the chain fails the whole open (the 4096-entry cap
  is a legitimate DoS bound and is *not* flagged).
- **Spec says**: [MS-PPT] 2.3.3 requires `offsetLastEdit` to be strictly decreasing, so a
  cycle is corruption — but no spec basis for the error's totality.
- **Real-world behavior**: LibreOffice terminates the walk on non-decreasing offsets and
  keeps what it has (svdfppt.cxx:1401); POI actively *repairs* ("Repairing invalid user
  edit atom", HSLFSlideShowImpl.java:419-433).
- **Suggested posture**: opt-in leniency (or default-repair with recording): break the
  chain at the repeated offset and keep the mappings collected so far — merging is
  newest-first, so live objects are already covered in practice.
### 5.10 Duplicate/versioned `SlideListWithTextContainer` and text-record cross-checks are file-fatal — P2 (confidence: high on code; low-medium frequency)
- **Location**: `crates/litchi-ppt/src/slide/directory.rs:168` (container version !=
  0x0f), :175 ("duplicate presentation SlideListWithTextContainer"), :275 ("text records
  reference unknown slideId"), :280-281 (interaction parse errors propagate out of
  directory build).
- **Check**: a second instance-0 slide list, a non-container version nibble, a text atom
  grouping whose slideId isn't in the directory, or any strictness failure inside
  `SlideAtomsSet::text_interactions()` aborts the whole presentation open.
- **Spec says**: [MS-PPT] 2.4.14.3 defines one presentation slide list; duplicates and
  dangling text records are beyond spec — internal strictness.
- **Real-world behavior**: LibreOffice iterates *multiple* SlideListWithText headers in
  sequence, scanning each for persist atoms (svdfppt.cxx:1507-1541), and ignores
  slide-list text it can't map. Incrementally-saved files can accrete duplicate lists.
- **Suggested posture**: default-repair — keep the last (live) instance-0 list like LO's
  sequential scan, drop the recVer gate, skip unmappable text records, isolate
  interaction parsing per slide (record failures, yield empty interactions).
### 5.11 No per-shape error isolation: one bad text box kills the slide's whole shape list — P1
- **Location**: `crates/litchi-ppt/src/slide/types/codec.rs:71`
  (`Self::convert_odraw_to_shape_enum(escher_shape)?` in the per-shape loop), with strict
  child parsers at `escher_textbox.rs:80` ("ClientTextbox contains multiple
  TextRulerAtom records") and `odraw/codec.rs:110-127` (duplicate
  `RoundTripHFPlaceholder12Atom`; placeholder ID not in 7–10 → error).
- **Check**: any error converting one OfficeArt shape (bad style runs from F5.1-F5.3, a
  duplicated `TextRulerAtom`, or a malformed round-trip placeholder atom) propagates and
  voids every shape on the slide.
- **Spec says**: 2.9.43 TextRulerAtom and 2.8.x round-trip atoms define singleton
  grammar; no rejection policy is prescribed.
- **Real-world behavior**: PowerPoint and LibreOffice render the remaining shapes; POI's
  `EscherTextboxWrapper` tolerates multiple rulers (last wins) and never validates
  round-trip placeholder IDs. Duplicated round-trip atoms are exactly what round-tripping
  producers emit.
- **Suggested posture**: default-repair at two levels — (a) per-shape isolation in
  `parse_shape_enums`: skip the failing shape (or fall back to a raw/minimal shape) and
  record; (b) relax the specific checks: keep first TextRulerAtom, drop duplicate
  round-trip atoms, unknown placeholder IDs → `None`.
### 5.12 `UserEditAtom` payload length restricted to exactly 28 or 32 — P2 (confidence: high; low frequency)
- **Location**: `crates/litchi-ppt/src/slide/directory.rs:123` ("UserEditAtom has invalid
  length").
- **Check**: the head-of-chain UserEditAtom must have recLen 0x1C or 0x20 exactly.
- **Spec says**: [MS-PPT] 2.3.3 — "rh.recLen MUST be 0x0000001C or 0x00000020" — spec
  basis exists.
- **Real-world behavior**: POI pads short reads (`if(len < 34) len = 34`,
  UserEditAtom.java:89) and never checks the declared length; LibreOffice reads fixed
  fields without inspecting recLen (svdfppt.cxx:466). All fields litchi consumes are
  present in any record ≥ 28 bytes.
- **Suggested posture**: default-repair — accept `data_len >= 28`, read the three
  offsets, treat `encryptSessionPersistIdRef` as present only when `data_len >= 32`,
  record. (Mid-chain atoms in mapping/chain.rs:21 already accept `record.len() >= 36` —
  the head check is inconsistent with that.)

## 6. litchi-pptx (PresentationML .pptx)

Infrastructure: no tolerance-report or leniency mechanism at all. Errors are scoped to
per-feature accessors (`slide.shapes()`, `slide.text()`, transition/sequence/chart
reads), so "whole-file failure" means failure of that feature for the slide/part — but
`shapes()`/`text()` are the primary read surface. Graceful postures already present:
unknown transition effects preserved as `Kind::Raw`, shape-like extension elements index
as `Kind::Unknown`, unknown chart plot types → `Type::Unknown`, `extLst` payloads opaque,
MCE baseline processing everywhere, `dur` strips an "ms" suffix, `on`/`off` booleans
accepted. The crate preserves unknown *structure* generously but refuses unknown *values*
absolutely.
### 6.1 Animations: stale shape-target and duplicate-shape-ID cross-reference checks kill the whole slide timing tree — P0
> **Status: FIXED (2026-10-02)** — stale `spTgt`/build targets are skipped per entry and duplicate `cNvPr@id`s resolve first-wins, recorded as `TimingDefect::{OrphanedTarget,
> DuplicateShapeId}` (slide part + spid) in `Sequence::tolerance_report()` via `Slide::animations()`; `Sequence::parse_slide_xml_strict` keeps the old check.
- **Location**: `crates/litchi-pptx/src/animations/codec/xml/parser/semantic.rs:1117-1141`
  ("animation target {} is not a shape on the current slide"; "duplicate shape ID in
  slide" at :154-158), default-on via `animations/codec/semantic.rs:82-85`
  (`parse_slide_xml` enables target validation).
- **Check**: every animation/build target `spid` must resolve to a `cNvPr@id` present on
  the slide, and no two `cNvPr` elements may share an id — any violation returns
  `Error::Invalid` for the entire `Sequence`.
- **Spec says**: ECMA-376-1 §19.5 (`p:spTgt`, `p:bldP`) defines no cross-reference
  validity requirement; orphaned timing entries are representable. ST_DrawingElementId
  uniqueness is a schema expectation, not something consumers enforce.
- **Real-world behavior**: PowerPoint routinely leaves orphaned `p:bldP`/timing entries
  behind when shapes are deleted, and duplicate shape IDs survive copy/paste and
  third-party round-trips; PowerPoint and LibreOffice ignore the dead entries and play
  the rest. Decks that have been edited at all are very likely carriers.
- **Suggested posture**: default-repair — drop the unresolvable animation/build entries
  (keep the rest of the sequence), tolerate duplicate IDs first-wins, record both.
  Target-existence checks that only matter for lossless rewrite can stay strict on the
  write path.
### 6.2 Shape scene: one malformed `a:xfrm` (missing `x`/`y`, non-numeric, negative `cx`/`cy`) fails every shape on the slide — P1 (confidence: high on code; medium on frequency)
- **Location**: `crates/litchi-pptx/src/shape/reader.rs:1379-1413` (`parse_i64`,
  `parse_nonnegative` at :1399), invoked at :734,743-744.
- **Check**: `<a:off>` without `x`/`y`, any non-`i64` coordinate, or any negative
  `a:ext@cx/cy` returns `Error::Invalid`, aborting `Scene::read` for the entire
  slide/layout/master.
- **Spec says**: ECMA-376-1 §20.1.7.3/§20.1.7.4 type these as `CT_Point2D` (required
  `x`,`y`) and `CT_PositiveSize2D` — negative extents are schema-invalid, but nothing
  mandates refusing the whole subtree.
- **Real-world behavior**: third-party converters and WPS emit negative `cx`/`cy` for
  flipped shapes; PowerPoint/LibreOffice open such files (clamping or normalizing). One
  bad transform currently makes `slide.shapes()` return `Err` for all well-formed
  sibling shapes. LibreOffice's `oox` helpers default missing/invalid numeric attributes.
- **Suggested posture**: default-repair — substitute 0 for missing/unparseable
  coordinates, clamp negative extents to 0 (or record raw), dropping only the offending
  shape's `Bounds`, never the scene; record each repair.
### 6.3 Processing instructions are hard-rejected crate-wide (108 `Event::PI` rejection sites) — P2 (confidence: high; low frequency, near-zero fix cost)
- **Location**: `crates/litchi-pptx/src/shape/reader.rs:551-554` (representative; also
  `parts/slide.rs:146-148,418-422,915-919`, `font/codec.rs:296-298`,
  `view_properties/codec.rs:178-180`, `media_parts/codec.rs:545`,
  `laser/package.rs:102-105`, `shape/zoom/codec.rs:339`,
  `master_layout/package.rs:580`, `model3d/codec.rs:223`, …).
- **Check**: any `<?pi …?>` anywhere in a part returns `Error::Invalid` ("DOCTYPE and
  processing instructions are forbidden/rejected"), failing the feature reading that
  part.
- **Spec says**: no spec basis — PIs are legal XML 1.0 (§2.6) and OOXML imposes no PI
  prohibition on part content. (DOCTYPE rejection is a defensible security posture and is
  *not* reported here.)
- **Real-world behavior**: PowerPoint and LibreOffice skip PIs via ordinary XML parsers.
  PIs get injected by XSLT-based publishing pipelines, CMS re-serializers, and archival
  tools — rare but legal. One stray PI kills shapes/text/viewProps/fonts/media reads.
- **Suggested posture**: default-repair — ignore `Event::PI` (as
  `presentation/transition.rs:1014` and `source.rs:4517` already do), record once per
  part. Keep DOCTYPE rejection.
### 6.4 Placeholder `p:ext` owner must have exactly one attribute, exactly one child, no character data — P1
- **Location**: `crates/litchi-pptx/src/shape/reader.rs:1422-1448`
  (`validate_extension_attributes`, invoked at :849), :850-854,866-872,950-958,993-998
  (child-count), :586-593 (character data).
- **Check**: a `p:ext` under `p:ph/p:extLst` with more than one non-namespace attribute,
  zero or 2+ children, or direct non-whitespace text returns `Error::Invalid` for the
  whole slide scene — for *any* extension URI, not just the p232 one (the URI gate runs
  only after `validate_extension_attributes`).
- **Spec says**: ECMA-376-1 pml.xsd `CT_Extension` is `<xsd:any processContents="lax"
  minOccurs="0" maxOccurs="unbounded"/>` with a required `uri` token — zero or *many*
  children are explicitly legal. The Markup Compatibility model requires unrecognized
  extension content to be ignorable.
- **Real-world behavior**: PowerPoint ignores extensions it doesn't understand; it never
  fails a slide over extension arity. Third-party producers that piggyback `p:ext` with
  richer payloads would kill every shape on the slide here.
- **Suggested posture**: default-repair — treat non-p232 `p:ext` subtrees as opaque: skip
  without inspection (mirror the existing `Kind::Unknown`/`Raw` postures). Reserve the
  strict grammar for the `urn:litchi:pptx:p232:phTypeExt` URI only, and even there
  degrade to "no type extension" + record.
### 6.5 Wheel/wheelReverse transition `spokes` restricted to {1,2,3,4,8} — P1
- **Location**: `crates/litchi-pptx/src/transition/reader.rs:927` ("wheel transition
  spoke count '{value}' is not supported by PowerPoint").
- **Check**: `spokes="6"` (or any value outside 1/2/3/4/8) returns `Error::Invalid`,
  failing the entire transition read.
- **Spec says**: ECMA-376-1 pml.xsd `CT_WheelTransition` declares `spokes` as plain
  `xsd:unsignedInt` default 4 — every unsigned value is schema-valid; the {1,2,3,4,8} set
  is only what PowerPoint's *UI* offers.
- **Real-world behavior**: LibreOffice imports with `rAttribs.getUnsigned(XML_spokes, 4)`
  (slidetransitioncontext.cxx:117) — any value, default 4, never an error — and maps
  unknown counts to the closest effect.
- **Suggested posture**: default-repair — parse any `unsignedInt`; model non-{1,2,3,4,8}
  counts as the nearest supported variant or `Kind::Raw` (already preserved inertly);
  record the substitution.
### 6.6 Transition attribute token domains (`spd`, `dir`, morph `option`, `advTm`/`dur` > i32::MAX) hard-fail — P1 (enum part P2)
- **Location**: `crates/litchi-pptx/src/transition/reader.rs:679-688` (`spd`), :735-936
  (direction enums), :906-915 (morph), :705-713 + `transition/model.rs:15` (`MAX_MS =
  i32::MAX`; "transition time {}ms exceeds the PowerPoint maximum" at :80).
- **Check**: any unrecognized enum token, or `advTm`/`dur` in (2³¹, 2³²), returns
  `Error::Invalid` for the whole transition; regression tests enshrine this
  (`transition/mod.rs:375-394,418-429`).
- **Spec says**: `spd`/`dir`/morph `option` are schema enumerations (invalid tokens are
  non-conformant), but `advTm` is `xsd:unsignedInt` — values up to 4,294,967,295 are
  schema-valid yet rejected above 2,147,483,647. The cap flatly contradicts the schema's
  own type.
- **Real-world behavior**: LibreOffice defaults unknown tokens (`getToken(XML_val,
  XML_col)` style, typegroupcontext.cxx:117); PowerPoint ignores/repairs bad transition
  attributes. The crate already demonstrates the right degradation one level down:
  unknown *elements* become `Kind::Raw` (reader.rs:456-461), but an unknown *attribute
  token* on a known element is fatal.
- **Suggested posture**: default-repair — unknown enum tokens degrade the effect to
  `Kind::Raw` (bytes already capturable); out-of-range times clamp/saturate to `MAX_MS`;
  each recorded.
### 6.7 Presentation slide list: uniqueness of slide IDs / relationship IDs / targets and child hygiene are fatal — P1
- **Location**: `crates/litchi-pptx/src/parts/presentation.rs:447-476` ("presentation
  slide reference ids must be unique" :447; relationship ids :456; targets :470),
  :188-192,248-252 (non-`sldId` child), :402-413 (any attribute on `sldIdLst`).
- **Check**: `slide_references()` returns `Error::Invalid` if two `p:sldId` entries share
  an id, an `r:id`, or a target part; if `sldIdLst` contains any non-`sldId` element or
  any attribute — failing slide enumeration for the whole presentation.
- **Spec says**: `CT_SlideIdList` permits only `sldId` children and `ST_SlideId` is
  256..2³¹ (the crate's range check matches the spec — fine), but the schema states no
  uniqueness constraint; duplicate-id handling is left to applications.
- **Real-world behavior**: LibreOffice reads `sldId` children and overwrites
  `maSlideIdToIndexMap[nSldId]` on duplicates (presentationfragmenthandler.cxx:822-832) —
  no rejection. PowerPoint opens decks with duplicate slide IDs, repairing silently.
  Duplicate IDs occur in files assembled by script generators.
- **Suggested posture**: default-repair — keep first occurrence on duplicates (record
  each), ignore unknown children/attributes of `sldIdLst`. These checks remain appropriate
  for the lossless *edit* planners.
### 6.8 Shape metadata: unparseable `cNvPr@id` / `p:ph@idx` fails the entire slide scene — P2
- **Location**: `crates/litchi-pptx/src/shape/reader.rs:674-679` ("invalid non-visual
  shape ID '{value}'"), :703-707 ("invalid placeholder index '{value}'").
- **Check**: a shape whose non-visual id is non-numeric/negative, or a placeholder `idx`
  that doesn't parse, returns `Error::Invalid` for the whole scene.
- **Spec says**: `ST_DrawingElementId` and `p:ph@idx` are `xsd:unsignedInt`, so the input
  is non-conformant — but the values are non-structural metadata the reader already treats
  as `Option`.
- **Real-world behavior**: PowerPoint re-assigns broken shape IDs on load; LibreOffice
  never parses these eagerly. A producer emitting `id=""` or a 64-bit id currently poisons
  `slide.shapes()` for every sibling shape.
- **Suggested posture**: default-repair — treat unparseable id/idx as `None` (both fields
  are already `Option` in `Record`), record the drop. Semantics-preserving.
### 6.9 Semantic slide text: foreign `t` elements, CDATA/entity references outside `a:t` are fatal — P2
- **Location**: `crates/litchi-pptx/src/parts/slide.rs:683` ("foreign text element is not
  a DrawingML a:t element", also :715, :747), :906-911 (CDATA / general entity reference
  outside DrawingML text), raw-scan equivalents :415-422,445-447,466-469.
- **Check**: any element with local name `t` in a non-DrawingML namespace (`p:t`, an
  extension payload's `x:t`), or a CDATA section / `&entity;` reference anywhere outside
  an open `a:t`, returns `Error::Invalid` for the whole `slide.text()` extraction.
- **Spec says**: no spec basis — internal strictness. `CT_ExtensionList` explicitly
  allows arbitrary foreign content; character-data kinds are ordinary XML.
- **Real-world behavior**: PowerPoint's text extraction doesn't look at foreign elements.
  Extension content containing an element coincidentally named `t` kills the slide's
  text, even though `shape/text.rs:extract` on the same fragment ignores foreign
  lookalikes gracefully (its test at shape/text.rs:763-766 asserts exactly that).
- **Suggested posture**: default-repair — ignore non-DrawingML `t` elements and pass
  through/skip CDATA and entity events outside `a:t` (the sibling extractor already proves
  this is safe); record the encounter. Same shape as litchi-docx F2.1.
### 6.10 Chart info: `barDir` value, duplicate structural elements, required `chart`/`plotArea` are fatal — P2
- **Location**: `crates/litchi-pptx/src/chart/codec.rs:334-347` ("barDir is missing its
  val attribute" :337; "invalid barDir value" :343), :217-247,272-302 (duplicate
  chart/title/legend/plotArea), :153-157 (missing chart or plotArea).
- **Check**: `chart_info()` returns `Error::Invalid` on an unknown `barDir` token, a
  duplicated structural child, or a chart part lacking `chart`/`plotArea`.
- **Spec says**: `ST_BarDir` is an enum and `plotArea` is required by `CT_Chart`, so the
  input is non-conformant; nothing requires the consumer to fail the whole inspection.
- **Real-world behavior**: LibreOffice defaults an unknown `barDir` to `col`
  (typegroupcontext.cxx:117) and tolerates duplicates first-wins; PowerPoint opens such
  charts after quiet repair. Mis-generated script charts currently yield zero information.
- **Suggested posture**: default-repair — unknown `barDir` → `Column` (LO's default);
  duplicates → first wins; missing `chart`/`plotArea` → `Type::Unknown` with empty info.
  Record each.
### 6.11 View properties: allow-listed attributes/children plus strict child ordering — P2
- **Location**: `crates/litchi-pptx/src/view_properties/codec.rs:871-878` (`only()` — any
  non-allow-listed attribute is `Error::Invalid`; call sites at :265, :311, :354, :361,
  :376), :278-283 (unknown `viewPr` child / out-of-order children), :322-325,363-368
  (exact child-sequence assertions).
- **Check**: any attribute or child element the codec doesn't model — including
  attributes added by newer PowerPoint versions — fails parsing of the entire
  `viewProps.xml`.
- **Spec says**: the sequence order and attribute sets do follow ECMA-376
  `CT_ViewProperties` & co., but OOXML consumers are expected to ignore what they don't
  understand; `viewProps.xml` is non-authoring metadata (window positions, zoom).
- **Real-world behavior**: PowerPoint round-trips unknown view-property attributes and
  ignores order glitches; a new Office release adding one attribute to `normalViewPr`
  would make this crate unable to read view properties from every file that version
  saves. (`extLst` *is* preserved opaquely — the strictness applies to first-class
  elements.)
- **Suggested posture**: default-repair — ignore unknown attributes/children, accept any
  child order (assign by name), preserve unknown subtrees opaquely as for `extLst`; record
  the skips.
### 6.12 Comments: `idx == 0` and duplicate `(authorId, idx)` keys fail the whole comment part — P2
- **Location**: `crates/litchi-pptx/src/comments/codec.rs:124-131` ("duplicate
  author/index comment key in slide" :130; `idx` parse at :120), mirrored at :210-218
  (validate).
- **Check**: a comment with `idx="0"`, or two comments sharing an `(authorId, idx)` pair,
  returns `Error::Invalid`, failing the entire slide comment list.
- **Spec says**: `CT_Comment@idx` is `ST_Index`, plain `xsd:unsignedInt` — zero is
  schema-valid, so "comment index must be at least 1" is internal strictness; the schema
  states no uniqueness constraint on `(authorId, idx)`.
- **Real-world behavior**: duplicate comment indexes show up after co-authoring merges
  and in producer round-trips; PowerPoint deduplicates/renumbers on load rather than
  discarding all comments. One bad entry currently hides every well-formed comment on the
  slide.
- **Suggested posture**: default-repair — accept `idx=0`; on duplicate keys keep the
  first comment and record the dropped duplicates (or renumber), never fail the list.

## 7. litchi-cfb + litchi-ole-common + litchi-crypto (compound-file substrate)

This is the substrate under every legacy format: each fix multiplies across litchi-doc,
litchi-xls, and litchi-ppt. Infrastructure: no `ToleranceReport` plumbing in any of the
three crates, and `XlsLeniency` never reaches this layer. All three verify-everything at
ingress; every deviation is a hard `CorruptedFile`/`InvalidFormat` that aborts
`OleFile::open`. Already-tolerant behavior worth using as the template: header DIFAT tail
garbage ignored (file.rs:1192-1197), FAT padding past EOF marked `ENDOFCHAIN` accepted
with a comment citing Word/POI (file.rs:1651-1657), short final sector reads as zeroes
(file.rs:1026-1030), v3 stream-size high DWORD masked per MS-CFB 2.6.1 (file.rs:1875-1880),
classic-Mac `00 52` root name canonicalized (file.rs:1842-1854), unbalanced red-black
color metadata tolerated (file.rs:1736-1742), non-DWORD property offsets accepted, unknown
top-level VT types preserved as `Value::Unknown`, nonzero `VARIANT_BOOL` → true, ≤ 3
opaque tail bytes on string/empty properties. litchi-crypto was largely out of scope
(encryption failures excluded); its DataSpaces parser repeats the reserved-zero pattern
("UNICODE-LP-P4 padding is nonzero", crates/litchi-crypto/src/spaces.rs:974) — worth
revisiting only if IRM/encrypted files from non-Microsoft producers become a target.
### 7.1 FAT chain longer than the declared stream size is fatal — P1
> **Status: FIXED (2026-10-02)** — litchi-cfb/ole-common gained `ToleranceReport` carriers (`OleFile`/`SharedOleFile::tolerance_report()`); a chain longer than the declared size is cut at
> the size (POI/LO) and unclaimed sectors recorded; cycles, out-of-range sectors, short chains and shared sectors stay fatal.
- **Location**: `crates/litchi-cfb/src/file.rs:4019` ("{table_name} chain exceeds its
  declared length"; scratch variant :4153; unclaimed-sector reconciliation "unclaimed
  physical sector {sector} has FAT marker" at :1667).
- **Check**: `collect_sector_chain_exact` requires the chain to hit `ENDOFCHAIN` exactly
  at `ceil(size/sector_size)` entries; `validate_physical_sector_layout` additionally
  rejects any allocated sector no claimed chain covers. Any surplus link or orphaned
  sector fails `OleFile::open`.
- **Spec says**: MS-CFB 2.6.1 defines Stream Size as authoritative; nothing requires the
  FAT chain to terminate at exactly that length, and no parser behavior is mandated for
  unreferenced allocated sectors.
- **Real-world behavior**: POI reads bounded by the declared size and never walks surplus
  links (`DocumentInputStream`; `POIFSFileSystem.java:441`); LibreOffice `StgDataStrm`
  likewise; Word/ole32 opens such files. Surplus links and orphaned sectors are exactly
  what interrupted fast-saves and incremental writers leave behind.
- **Suggested posture**: default-repair — collect only the first `expected_count`
  sectors, stop validating there, and drop the unclaimed-sector-vs-FREESECT
  reconciliation to a recorded warning.
### 7.2 Zero-size stream whose start sector is not ENDOFCHAIN is fatal — P1
> **Status: FIXED (2026-10-02)** — a zero-size stream with a stale start sector reads as empty (entry reset to ENDOFCHAIN in memory) and is recorded per SID.
- **Location**: `crates/litchi-cfb/src/file.rs:3978` ("Empty {table_name} chain must
  start with ENDOFCHAIN"; scratch path :4098), applied to every stream at open via
  `validate_stream_allocations` (file.rs:1545, 1590, 1634).
- **Check**: one empty stream with a stale start sector (0, or a dangling sector id)
  kills the whole compound file.
- **Spec says**: MS-CFB 2.6.1 does not prescribe a starting-sector value for zero-length
  streams, and a zero-byte read is unaffected by the field.
- **Real-world behavior**: POI and LibreOffice both return an empty stream for size 0
  regardless of the start field (LibreOffice `StgEntry::Load` only rejects
  negative/oversized sizes, stgelem.cxx:418-445). Several third-party CFB writers
  zero-fill or leave `FREESECT`/stale values here.
- **Suggested posture**: default-repair — when size is 0, ignore the start marker
  entirely (chain = empty) and record. Cannot produce wrong content.
### 7.3 Red-black sibling-tree name-ordering violation fails the whole file — P1 (confidence: high on references; medium on prevalence)
> **Status: FIXED (2026-10-02)** — sibling-tree name-ordering violations are recorded per SID and name lookups fall back to a tree scan; ordered files keep the binary-search path.
- **Location**: `crates/litchi-cfb/src/file.rs:1769` ("CFB sibling tree violates name
  ordering at SID {sid}"; repeated-SID/ownership check at :1748).
- **Check**: full binary-search-tree ordering per MS-CFB 2.6.4 is validated across every
  sibling tree; any violation → `CorruptedFile` and the file cannot be opened at all.
- **Spec says**: MS-CFB 2.6.4 constrains writers ("The left sibling MUST always be less
  than the right sibling"); no parser behavior is specified for violation.
- **Real-world behavior**: LibreOffice never validates ordering — `StgDirStrm::SetupEntry`
  (stgdir.cxx:728) just walks links and on a circular link discards it and continues
  (fdo#41642); POI walks the same links and looks children up by linear name scan
  (PropertyTable.java:234), so a misordered tree is fully transparent. Buggy writers
  (locale-sensitive comparators instead of the 2.6.4 simple-uppercase rule) produce such
  trees occasionally.
- **Suggested posture**: default-repair — keep bounds/acyclicity/unique-ownership checks
  (safety-relevant), but on ordering violation fall back to a linear name index for
  lookups instead of failing; record.
### 7.4 Directory entries unreachable from the root are fatal — P1
> **Status: FIXED (2026-10-02)** — well-formed directory entries unreachable from the root are treated as unused and recorded per SID (with their unclaimed sectors); malformed ones stay fatal.
- **Location**: `crates/litchi-cfb/src/file.rs:1801` ("CFB directory SID {} is not owned
  by a storage").
- **Check**: after tree traversal, every non-empty entry must be owned by a storage; a
  single orphaned entry fails the open.
- **Spec says**: no spec basis — MS-CFB 2.6 never requires all entries be reachable and
  says nothing about parser behavior for orphans.
- **Real-world behavior**: POI only materializes entries reachable from entry 0
  (PropertyTable.java:110-118); LibreOffice recurses only from the root
  (stgdir.cxx:720-741). Orphans are the classic residue of crash-interrupted saves.
- **Suggested posture**: default-repair — skip unreachable entries (treat as
  `STGTY_EMPTY`), record count.
### 7.5 MiniFAT/DIFAT start-sector vs count cross-check is fatal — P2 (confidence: medium)
- **Location**: `crates/litchi-cfb/src/file.rs:1033` ("MiniFAT start sector and sector
  count disagree"), :1038 (DIFAT same).
- **Check**: `(count == 0) != (first == ENDOFCHAIN)` → `InvalidFormat`, rejecting e.g.
  count=0 with start=0 (zero-filled instead of ENDOFCHAIN).
- **Spec says**: MS-CFB 2.2 defines the two fields independently; it never states they
  must be mutually consistent.
- **Real-world behavior**: POI reads the MiniFAT chain only when `_sbat_count > 0` and
  ignores the start field otherwise (HeaderBlock.java:171-180); LibreOffice accepts
  `FREESECT`/`ENDOFCHAIN` starts with any count (`isKnownSpecial`, stgelem.cxx:197-202).
  Writers that zero-init the header (some C/Ruby/Python CFB libraries) emit
  start=0,count=0.
- **Suggested posture**: default-repair — the count field is authoritative: count=0 ⇒
  ignore start; count>0 with a non-sector start ⇒ treat the table as empty (mini streams
  fall back like a missing MiniFAT); record.
### 7.6 CFB header reserved bytes must be zero — P2 (confidence: high that references ignore; medium-low prevalence)
- **Location**: `crates/litchi-cfb/src/file.rs:977` ("Reserved CFB header bytes must be
  zero").
- **Check**: `header[0x22..0x28] != [0; 6]` → `InvalidFormat`, aborting open.
- **Spec says**: MS-CFB 2.2 "Reserved (6 bytes): MUST be set to all zeroes" — a producer
  requirement; the bytes carry no semantics.
- **Real-world behavior**: POI never reads these bytes (HeaderBlock.java:124-180);
  LibreOffice skips them entirely (`StgHeader::Load`, stgelem.cxx:137-139); Word/ole32
  ignores them.
- **Suggested posture**: default-repair — ignore, record. Same logic already applied to
  the header DIFAT tail at file.rs:1192.
### 7.7 Directory-entry field byte-exactness bundle — P2 (confidence: high on references; medium-low prevalence per sub-check)
- **Location**: `crates/litchi-cfb/src/file.rs:272` ("invalid CFB directory node color"),
  :1893-1901 (storage start/size nonzero → fatal), :1902-1906 (stream `sid_child !=
  NOSTREAM` → fatal), :1709 (root must be named exactly "Root Entry" with
  `sid_left == sid_right == NOSTREAM`), :1817-1823 (empty entry with nonzero name length
  → fatal), :1863 + `directory_name.rs:88` ("CFB directory entry name contains forbidden
  character" — `/ \ : !` and >31-unit names rejected on read).
- **Check**: per-entry MUST-fields are validated and any deviation aborts the whole open.
- **Spec says**: MS-CFB 2.6 marks all of these MUST — all producer constraints with no
  parser behavior defined.
- **Real-world behavior**: LibreOffice ignores the color byte entirely ("0 or 1 (tree
  balance?)", stgelem.cxx:398), never checks storage start/size or stream child, forces
  size=0 on empty entries (tdf#112399: "opens fine in MSOffice 2013 despite a massive
  m_nSize field", stgelem.cxx:424-434), truncates over-long names instead of failing
  (stgelem.cxx:453-457), and forces entry 0 to type root without checking its name
  (stgdir.cxx:749). POI maps name_length<1 to "" and enforces no character rules on read.
- **Suggested posture**: default-repair with recording — mask color to parity (already
  ignored semantically), ignore storage start/size and stream child on read, accept any
  root name (canonicalize for lookup only), treat forbidden-char names as opaque lookup
  keys, treat garbage in empty entries as empty. Keep SID/type/bounds checks.
### 7.8 Property-set strings: strict MBCS and UTF-16 decode, no lossy fallback — P1
- **Location**: `crates/litchi-ole-common/src/property_set/codec/binary/wire.rs:456`
  (ANSI via `Mbcs::require`/strict decode) and :419-438 (`decode_utf16` with hard error
  on lone surrogates).
- **Check**: one malformed byte sequence (truncated DBCS pair, invalid trail byte) or one
  unpaired surrogate in any LPSTR/LPWSTR/BSTR/dictionary string fails the property, which
  fails the section, which fails the whole property-set stream.
- **Spec says**: no spec basis for total failure — MS-OLEPS 2.5/2.15 define the encodings
  but no parser recovery. `litchi-codepage` already ships `decode_lossy`
  (crates/litchi-codepage/src/lib.rs:135) for exactly this situation.
- **Real-world behavior**: Windows `MultiByteToWideChar` (used by ole32's property
  reader) substitutes a default character; POI decodes via Java `Charset` in replacing
  mode. Truncated-at-buffer-boundary Shift-JIS/GBK titles are a classic defect in East
  Asian legacy files.
- **Suggested posture**: default-repair — decode via `Mbcs::recover`/`decode_lossy` and
  `String::from_utf16_lossy`, record (mirroring the VBA precedent for which `recover`
  exists).
### 7.9 Property-set PID 1 codepage outside the supported table kills the stream — P1 (confidence: high on references; medium on prevalence)
- **Location**: `crates/litchi-ole-common/src/property_set/codec/binary/validation.rs:589`
  (`CodePage::try_from`; error text "unsupported Property Set code page {page}" at
  `crates/litchi-ole-common/src/property_set/model.rs:247`).
- **Check**: if PID 1 names a codepage not in the ~34-entry table
  (crates/litchi-codepage/src/lib.rs:546-581 — 28591 ISO-8859-1, 20127 US-ASCII, 437/850,
  65000, 50220 ISO-2022-JP are all absent), the entire section — and therefore the whole
  SummaryInformation stream — fails, even if every string is ASCII-compatible.
- **Spec says**: MS-OLEPS 2.18 permits any valid Windows code page identifier in PID 1;
  the spec's list is illustrative, not exhaustive.
- **Real-world behavior**: POI maps the identifier through Java's charset registry (far
  broader) and falls back to the default when unmappable; Windows resolves any installed
  code page. 28591 in PID 1 is a known artifact of Java-/Linux-based producers writing
  "ISO-8859-1" instead of 1252.
- **Suggested posture**: default-repair — unknown codepage ⇒ fall back to WINDOWS_1252
  (the already-used default when PID 1 is absent, validation.rs:594) and record; bytes
  0x00-0x7F decode identically either way.
### 7.10 Property-set strings must be NUL-terminated exactly as sized — P1 (confidence: medium-high)
- **Location**: `crates/litchi-ole-common/src/property_set/codec/binary/wire.rs:381`
  ("{description} is not NUL-terminated" — ANSI CodePageString; UTF-16 variant :405),
  :369, and validation.rs:680/688 (dictionary entries).
- **Check**: a `CodePageString`/`UnicodeString` whose byte buffer does not end with the
  terminator (or whose terminator is missing within the declared size) fails the property
  and the stream.
- **Spec says**: MS-OLEPS 2.5 says the string "MUST be terminated by a null character" —
  producer constraint; the size field makes the terminator redundant for a length-bounded
  reader.
- **Real-world behavior**: POI reads the declared size and strips a trailing terminator
  only if present; Word/ole32 reads length-bounded. Producers that disagree on whether
  the size includes the terminator (a famous interop ambiguity) hit this.
- **Suggested posture**: default-repair — treat the terminator as optional: decode `size`
  bytes, strip a trailing NUL if present, record when absent. Lossless given the explicit
  size field.
### 7.11 One malformed property fails the entire property-set stream; nested reserved-zero/padding checks — P1 (confidence: high on achievability; medium on frequency)
- **Location**: `crates/litchi-ole-common/src/property_set/codec/binary/validation.rs:613-627`
  (per-property parse error wrapped and propagated), semantic.rs:584/:653 (variant
  vector/array element reserved word must be zero), composite.rs:83/:187 (HeadingPairs
  reserved words must be zero), semantic.rs:549 (unknown *nested* variant type → error),
  wire.rs:256 (`finish_zero_padding` — trailing bytes must be exactly zero; only ≤3
  bytes for 5 VT types excused at semantic.rs:347-355).
- **Check**: any single property that trips these aborts `parse_section`, so e.g. one
  HeadingPairs entry with a nonzero padding word loses *all* document metadata.
- **Spec says**: MS-OLEPS 2.15 Padding "MUST be 0x0000" and MS-OSHARED composite layouts
  are producer constraints; no parser behavior specified. The crate already ignores the
  same reserved word at the top level (semantic.rs:330-334) — the nested checks are
  inconsistent with its own posture.
- **Real-world behavior**: POI resolves property values lazily per PID in common access
  patterns (`Section.getProperty`), so a malformed HeadingPairs doesn't destroy access to
  Title/Author; Word's property pages display readable properties from partially damaged
  sets. Fast-saved documents are a known source of garbage padding between values.
- **Suggested posture**: default-repair — (a) ignore nested reserved words and trailing
  padding uniformly (mask, don't compare), and (b) on a per-property parse failure degrade
  that property to `Value::Unknown { variant_type: 0, data: raw_range }` (the range is
  known from the offset table, validation.rs:564-576) and record, instead of failing the
  section.
### 7.12 Property-set stream header exactness (version 0/1, section count 1..=2, duplicate descriptors, section size equality) — P2 (confidence: medium)
- **Location**: `crates/litchi-ole-common/src/property_set/codec/binary/validation.rs:46`
  (version), :55-57 ("Property Set section count {section_count} is outside 1..=2"), :84
  (duplicate FMTID/offset), :99 (overlap), :507 (declared section size must equal its
  computed range exactly).
- **Check**: version ∉ {0,1}, section count ∉ {1,2}, two descriptors sharing an FMTID or
  offset, overlapping section ranges, or a section-size field that disagrees with the
  descriptor-implied range each fail the whole stream.
- **Spec says**: MS-OLEPS 2.20 constrains writers; the section size check duplicates
  information the offsets already provide.
- **Real-world behavior**: POI reads the version and iterates sections by count without
  these cross-checks; unusual producers (padding-aware writers that round the section
  size up) open there. Rare overall.
- **Suggested posture**: opt-in leniency for version/section-count (a v2+ stream might
  genuinely need new semantics — skip unknown extra sections); default-repair for the
  size-equality check (trust the offset/descriptor, clamp to stream end) with recording.

## 8. litchi-opc + litchi-ooxml-common (OOXML packaging)

This scope is the packaging substrate shared by docx/xlsx/pptx, so each finding surfaces
through three format crates. Neither crate has a `ToleranceReport` mechanism.
Container-boundary ingress is in good shape (unmappable ZIP members become reported junk;
missing `.rels` parses as empty; dangling targets stay visible; content-type mappings fold
ASCII case; duplicate ZIP names last-write-wins; query/fragment stripped from internal
targets; part lookup tolerates case-mismatched references; digital signatures never
validated at open — verification is opt-in, package.rs:2038-2052). The strictness
concentrates in metadata parts vetoing the document, relationship-manifest conformance,
and manifest grammar purism.
### 8.1 Core-properties XML body: producer-conformance rules (M4.2–M4.5) plus extra strictness fail the whole open — P1 (some triggers arguably P0)
- **Location**: `crates/litchi-ooxml-common/src/properties/read.rs:573` ("OPC M4.5
  requires xsi:type on created and modified"; related :545, :550, :560), :539 ("OPC M4.4
  forbids xml:lang in core properties"), :566 (any other attribute forbidden), :591
  ("duplicate core property"), :361 (unknown element forbidden), :300 (unsupported
  namespace), :88/:125 (XML-vs-relationship dialect mismatch), :83 (non-UTF-8 fatal);
  propagated with `?` into open paths at `crates/litchi/src/sheet/workbook.rs:270`,
  :522, `crates/litchi/src/presentation/prs.rs:916`,
  `crates/litchi-docx/src/package/codec.rs:382`,
  `crates/litchi/src/document/doc.rs:989`,
  `crates/litchi/src/detection_smart/detected.rs:984`.
- **Check**: any deviation — `<dcterms:created>` without `xsi:type="dcterms:W3CDTF"`,
  `xml:lang` on any property, any non-xmlns attribute other than the sanctioned xsi:type,
  a repeated property, an unknown element, a strict-namespace XML behind a transitional
  relationship type, or a Latin-1-encoded core.xml — returns `Error::Invalid`, aborting
  `Workbook::open`, `Presentation::open`, and `docx::Package::from_opc_package` entirely.
- **Spec says**: ECMA-376 Part 2 §8.3.1, §8.3.4.1 (non-repeatable), §8.3.4.2 (no
  xsi:type/xml:lang on dc elements), §8.3.4.3 (xsi:type required on created/modified) —
  all *producer* conformance statements; no consumer recovery behavior defined.
- **Real-world behavior**: Word and LibreOffice read core.xml leniently (per-property
  lookup; unknown elements/attributes ignored; dates parse with or without xsi:type).
  POI enforces M4.2–M4.5 but made it configurable (`OPCComplianceFlags.disableAll()`,
  PackagePropertiesUnmarshaller.java:256-293). Files missing xsi:type, or carrying
  xml:lang or duplicated Dublin Core fields, are regular outputs of third-party
  generators — metadata is non-structural, so one bad attribute killing the file is the
  worst possible ratio.
- **Suggested posture**: default-repair — parse known properties leniently (ignore unknown
  elements/attributes, accept missing/extra xsi:type, keep first or last of duplicates,
  drop the dialect cross-check to a recorded note), record every deviation. Lossy-decode
  non-UTF-8 core.xml. Keep M4.2 as the only candidate for hard error if desired.
### 8.2 Core-properties graph checks: orphaned part, dangling relationship, duplicate typed parts fail the open — P1
- **Location**: `crates/litchi-ooxml-common/src/properties/graph.rs:96` (two parts with
  the core-properties content type), :104 (relationship owned by a part), :132
  ("core-properties part '{}' is orphaned"), :173 (relationship target part missing →
  `Error::Missing`); same propagation as 8.1.
- **Check**: any of these four graph anomalies errors the metadata read, which aborts
  document/workbook/presentation open.
- **Spec says**: ECMA-376 Part 2 §8.2/§8.3 require a package-level core-properties
  relationship and one part; no consumer rule covers an orphaned or dangling case.
- **Real-world behavior**: Word, LibreOffice, and POI locate core.xml *through* the
  relationship; an orphan part is ignored, a dangling relationship yields no properties,
  a second typed part is dead weight. Broken resave tools that delete
  `docProps/core.xml` but keep the `/_rels/.rels` entry (or vice versa) produce exactly
  these shapes and open fine everywhere else.
- **Suggested posture**: default-repair — orphaned part → absent (or adopt), dangling
  relationship → `Ok(None)`, duplicates → pick the relationship-selected part; record
  each. Degraded-but-correct: no content is invented.
### 8.3 Custom-properties codec: unknown `vt:*` types, missing pid/fmtid, PIs, duplicate PIDs, strict value parsing — fatal on the DOCX open path — P1 (P0 for the DOCX blast radius)
- **Location**: `crates/litchi-ooxml-common/src/custom/codec.rs:543-547` ("unsupported
  custom-property value type" — only 9 `vt` kinds supported, custom/model.rs:124),
  :486-495 (missing name/pid/fmtid), :489 (pid < 2), :656-679 (fmtid must be braced
  GUID; summary fmtids forbidden :646-654), :99-102 (duplicate PID), :72 ("processing
  instructions are forbidden in custom-properties XML"), :89 (unknown element under
  root), :237/:693-711 (entity refs only inside values; only 5 predefined entities),
  :594-606 (strict bool / RFC3339 filetime), :579/:589 (non-finite floats). Called on
  the open path at `crates/litchi-docx/src/package/codec.rs:381`
  (`CustomProps::read_for(&opc, CustomPropsHost::Word)?`); accessor-only for xlsx/pptx
  (`crates/litchi-xlsx/src/package.rs:568`, `crates/litchi-pptx/src/package/model.rs:55`).
- **Check**: any one bad `<property>` — a `vt:date`, `vt:cy`, `vt:ui4`, `vt:bstr`,
  `vt:decimal`, `vt:vector`, or `vt:clsid` value; a missing or sub-2 `pid`; a non-GUID
  `fmtid`; two properties sharing a PID; a stray `<?xml-stylesheet?>` PI; an unknown
  entity — fails `decode`, and for DOCX this aborts the entire open.
- **Spec says**: ECMA-376 Part 2 Annex C permits the full OLE property-variant vocabulary
  ([MS-OLEPS]); pid/fmtid presence is a producer rule. The rejected variant types are
  schema-valid but unimplemented.
- **Real-world behavior**: Word shows what it understands and ignores the rest of
  custom.xml — it never refuses a document over a custom property. Enterprise files
  carrying `vt:date`/`vt:cy`/typed properties from DMS systems, SharePoint, or add-ins
  are a steady trickle.
- **Suggested posture**: default-repair per property — skip an unparseable/unsupported
  `<property>` (or read it as opaque text) and record it; never let the property set veto
  the document. Keep collection-level well-formedness errors (unclosed root) fatal.
### 8.4 Reserved sensitivity-label (MS-OI29500 3.11) custom properties are host- and format-validated on open — P1 (confidence: medium-high)
- **Location**: `crates/litchi-ooxml-common/src/custom/reserved.rs:107` (property not
  valid for this host), :126 (`Sensitivity` must be a hyphenated GUID), :48-67
  (font/shape-id/location formats; `Sensitivity` name match at :48); reached from
  `custom/package.rs:41-43` (`read_for` → `validate_for`), on the DOCX open path at
  `crates/litchi-docx/src/package/codec.rs:381`.
- **Check**: a Word document whose custom.xml contains a PowerPoint-scoped
  `ClassificationWatermark*` property, a `Sensitivity` value that isn't a 36-char GUID,
  or a malformed `ClassificationContentMarking*FontProps`/shape-id list makes the whole
  DOCX fail to open.
- **Spec says**: MS-OI29500 3.11 enumerates the reserved names; the module doc itself
  admits the spec "does not define a property-name grammar, value type, or date format"
  for the MIP samples (reserved.rs:1-6) — the enforced formats are inferred, not
  specified.
- **Real-world behavior**: these properties are written by Microsoft Information
  Protection SDK and third-party classification/DLP tools; cross-host leftovers and
  malformed label values occur in labeled enterprise corpora. Word renders what it can
  and ignores the rest.
- **Suggested posture**: default-repair — treat reserved-property validation failures as
  per-property skips recorded in the tolerance report; never open-fatal. Host-scope
  checks should be advisory at read time (keep them for write-time validation).
### 8.5 Duplicate relationship `Id` rejects the entire package — P1 (confidence: high on check/references; medium on frequency)
- **Location**: `crates/litchi-opc/src/pkgreader.rs:2200,2205`
  (`OpcError::DuplicateRelationshipId`); fired while parsing any reachable `.rels`
  (pkgreader.rs:1149-1176 → `load_rels_lazy`), so one duplicated rId anywhere in the
  reachable graph kills the open.
- **Check**: two `<Relationship>` elements sharing an `Id` in one manifest → error.
- **Spec says**: ECMA-376 Part 2 §6.5.3.4: "The value of the Id attribute shall be unique
  within the Relationships part" — producer rule; no consumer directive.
- **Real-world behavior**: POI inserts into a map with no duplicate check — last write
  wins (PackageRelationshipCollection.java:187-193); LibreOffice appends to a sequence
  with no uniqueness check (ofopxmlhelper.cxx:326-363); Word opens such files (one of the
  duplicates wins). Duplicate-rId files from buggy generators are a recurring theme in
  POI/OOXML bug trackers.
- **Suggested posture**: default-repair — keep the first occurrence, drop subsequent
  duplicates, record the dropped rIds (first-wins matches what most producers' r:id
  references were emitted against).
### 8.6 One unresolvable internal relationship target fails the whole package — P1
- **Location**: `crates/litchi-opc/src/pkgreader.rs:1879-1913` (`enqueue_target`
  propagates `target_partname()` errors with `?`), :736-756,
  `crates/litchi-opc/src/packuri.rs:120` (`from_rel_ref` rejects backslashes and absolute
  URIs), :74-81 (literal spaces rejected), :355-363 (root-escaping `..` rejected).
- **Check**: during the open-time graph walk, every internal target is resolved to a
  `PackURI`. A target with a literal space (`media/image 1.png`), a backslash, an
  absolute URI written as an internal target, or enough `../` to escape the package root
  turns the whole open into `OpcError::InvalidPackUri`.
- **Spec says**: ECMA-376 Part 2 §9.1.1.1 part-name grammar and §6.5.3.4 (Internal ⇒
  relative reference) — producer rules. (Query/fragment are already tolerated —
  `relationship_target_components` strips them, pkgreader.rs:742 + rel.rs:799-811.)
- **Real-world behavior**: POI catches `URISyntaxException`, logs, and substitutes a
  dummy URI "thus enable processing of the package"
  (PackageRelationshipCollection.java:349-359); LibreOffice percent-encodes on
  resolution; Word opens the file with the one relationship dead. Literal-space targets
  are a known wild case from SoftMaker/WPS-era and hand-rolled generators.
- **Suggested posture**: default-repair — treat the unresolvable relationship as dangling
  (keep it visible on its source, never enqueue), record the raw target. A light
  normalization pass (percent-encode spaces, `\` → `/`) with recording would recover most
  outright.
### 8.7 `TargetMode` attribute parsed case-sensitively — P1 (confidence: high on code/POI; medium on frequency)
- **Location**: `crates/litchi-opc/src/rel.rs:23-28` (only exact `"Internal"`/`"External"`;
  other values → `OpcError::InvalidRelationshipTargetMode` at :27), invoked per attribute
  at `crates/litchi-opc/src/pkgreader.rs:2176`.
- **Check**: `TargetMode="external"` (any non-exact casing or any other value) is fatal
  to the package.
- **Spec says**: ECMA-376 Part 2 §6.5.3.4 ST_TargetMode — producer-side value domain; no
  consumer directive.
- **Real-world behavior**: POI lowercases the value and treats *anything* other than
  "internal" as `TargetMode.EXTERNAL` (PackageRelationshipCollection.java:340-346) —
  garbage degrades to external, never fatal; LibreOffice doesn't validate the value at
  all (ofopxmlhelper.cxx:356-360). (System.IO.Packaging is strict — a mixed-reference
  case, but the project's compatibility pole is Word/LibreOffice.)
- **Suggested posture**: default-repair — case-insensitive match; any other value →
  `External` (POI semantics — safer than Internal since it prevents bogus part
  resolution); record the raw value.
### 8.8 Missing or empty `Target` (or `Type`) attribute is fatal per relationship — P1 (confidence: high on code/LO/POI; medium on frequency)
- **Location**: `crates/litchi-opc/src/pkgreader.rs:2186-2188` (call sites) + :2233-2239
  (`required_relationship_attribute` rejects missing *and* empty values).
- **Check**: `<Relationship Id="rId5" Type=".../hyperlink" TargetMode="External"/>` or
  `Target=""` → `InvalidRelationshipsManifest`, fatal to the package.
- **Spec says**: CT_Relationship marks Target/Type required — producer/schema rule; no
  consumer directive.
- **Real-world behavior**: LibreOffice requires only `Id` (ofopxmlhelper.cxx:333-360 —
  Type/Target/TargetMode appended only if non-empty); POI reads an empty string and
  parses it as an empty relative URI; Word keeps the relationship as a dead link.
  Empty-Target external hyperlink relationships are emitted by several enterprise
  exporters.
- **Suggested posture**: default-repair — accept the relationship with an empty target
  (it resolves to nothing, exactly like the dangling targets this crate already
  tolerates, pkgreader.rs:1831-1837), record it; a missing `Type` can be kept as an
  opaque empty-type relationship.
### 8.9 Duplicate content-type Default/Override mappings reject the package — P1 (confidence: high on code/POI/LO; medium on frequency)
- **Location**: `crates/litchi-opc/src/content_type.rs:207`
  (`DuplicateContentTypeDefault`), :219 (`DuplicateContentTypeOverride`,
  case-insensitive both).
- **Check**: two `<Default Extension="xml">` (even differing only in case or declared
  content type) or two `<Override>` for equivalent part names → hard error; the whole
  package fails.
- **Spec says**: ECMA-376 Part 2 §9.1.2.2 — one Default per extension / one Override per
  part name; producer rule, no consumer directive.
- **Real-world behavior**: POI's DOM parse simply overwrites map entries
  (ContentTypeManager.java:398-403); LibreOffice appends to a sequence and later lookups
  take the first match (ofopxmlhelper.cxx:117-126); Word opens such manifests. Duplicated
  mappings from buggy or hand-merged generators are recurring and usually benign.
- **Suggested posture**: default-repair — keep first mapping, drop later duplicates,
  record; identical-mapping duplicates are trivially safe, conflicting ones degrade to
  first-wins exactly like LibreOffice.
### 8.10 A `.rels` member declared with a non-relationships content type rejects the package — P2 (confidence: high on code/references; low-medium on frequency)
- **Location**: `crates/litchi-opc/src/pkgreader.rs:1922-1937`
  (`relationship_part_content_type`: a manifest mapping a `_rels/*.rels` name to anything
  other than the fixed constant → `OpcError::InvalidContentType`).
- **Check**: cross-checks the declared content type of every relationship part. The crate
  synthesizes the constant when the mapping is *missing* (:1934) — the wrong-type case is
  stricter than the missing case, inverted relative to reference behavior.
- **Spec says**: ECMA-376 Part 2 §9.2 fixes the content type (producer rule).
- **Real-world behavior**: LibreOffice and POI identify relationship parts purely by name
  pattern (`_rels/.rels` paths) and never cross-check the declared content type. Packages
  declaring `<Default Extension="rels" ContentType="application/xml"/>` open everywhere
  else.
- **Suggested posture**: default-repair — override the declared type with the fixed
  constant (the part *is* parsed as relationships regardless), record the mismatch.
### 8.11 Manifest shape strictness: unknown elements, unexpected attributes, text, CDATA in `[Content_Types].xml` and `.rels` — P2 (confidence: high on code/POI/LO; medium on rarity)
- **Location**: `[Content_Types].xml` — `crates/litchi-opc/src/content_type.rs:136-146`
  (non-whitespace text / CDATA), :327 ("unexpected element or namespace in content types
  manifest"), :348-350 (children other than Default/Override), :393 ("unexpected
  attribute on {element_name}"). `.rels` — `crates/litchi-opc/src/pkgreader.rs:1184-1192`
  (text / DocType), :2133-2137 (only direct Relationship children), :2179-2183
  ("unexpected Relationship attribute").
- **Check**: any extension element, foreign-namespace child, producer-specific attribute,
  or stray character data in either OPC manifest → error → whole package fails.
- **Spec says**: ECMA-376 Part 2 Annex C closed content models — producer rules; no
  consumer directive.
- **Real-world behavior**: POI queries only the known elements/attributes out of the DOM
  and ignores everything else (ContentTypeManager.java:396-407,
  PackageRelationshipCollection.java:319-360); LibreOffice throws on unknown *elements*
  but ignores unknown attributes and all character data (ofopxmlhelper.cxx:364-365,
  445-446) — litchi is stricter than both on attributes and text.
- **Suggested posture**: default-repair for unknown *attributes* and character data
  (ignore + record); opt-in leniency for unknown elements (skip + record), since
  LibreOffice shares the element-level strictness.
### 8.12 MIME grammar validation of every ContentType value — P2 (confidence: high on code/POI parity; low-medium on frequency)
- **Location**: `crates/litchi-opc/src/content_type.rs:232-266`
  (`validate_content_type`) + :268 (`is_token`); :406-418 (`InvalidContentTypeExtension`
  for extensions with spaces/non-ASCII).
- **Check**: full RFC-2616-token grammar enforcement on every declared content type and
  extension at manifest parse time — `application/xml; charset=utf-8` (space after `;`),
  a parameter without `=`, or a content type missing `/` → `OpcError::InvalidContentType`
  → package fails.
- **Spec says**: ECMA-376 Part 2 §9.1.2.2 requires MIME types per RFC 2616 grammar —
  producer rule.
- **Real-world behavior**: LibreOffice treats content type strings opaquely (stores and
  compares verbatim); POI also validates grammar — references are mixed; Word is the
  tolerant pole. Wild occurrences (stray whitespace, empty parameters from hand-edited
  manifests) are uncommon.
- **Suggested posture**: opt-in leniency — store the value verbatim, compare with
  whitespace-collapsed matching, record. Grammar failures never change which part is
  which.

## 9. litchi-rtf + ODF family (odt/ods/odp/odf-common)

Infrastructure: no `Leniency`/tolerance-report mechanism in either family. litchi-rtf's
headline paths are tolerant (unknown control *words*/destinations preserved as opaque
nodes, color-table components clamped, body text decodes lossily), but its secondary
layers are systematically strict. The ODF family is strictest exactly at the container
boundary (mimetype grammar, manifest cross-checks, DOCTYPE ban, canonical manifest-path
validation) while the XML content layers are comparatively tolerant (unknown manifest
elements skipped, unknown cell value-types preserved, styles.xml/meta.xml optional). Two
deliberate non-findings from the reviewer: the missing-manifest hard error *matches*
LibreOffice's default (ZipPackage.cxx:469-471), and DOCTYPE rejection in OOXML is a
defensible security posture. A verified additional strictness note folded into F9.2's
theme: duplicate manifest entries are rejected ("Duplicate manifest file entry",
crates/litchi-odf-common/src/package/codec.rs:347-352) where LibreOffice's
duplicate-check-free `ManifestImport` tolerates them.
### 9.1 Deflated or re-zipped `mimetype` member rejected at package open — P0
> **Status: FIXED (2026-10-02)** — owned, source-backed, prepared and detection paths read a deflated / extra-field / not-first `mimetype` member by name and record
> `PackageDefect::{MimetypeDeflated, MimetypeExtraField, MimetypeNotFirst}` on odt/ods/odp `tolerance_report()`; verified through `litchi::Document::{open, from_bytes}`.
- **Location**: `crates/litchi-odf-common/src/core/package.rs:709` ("ODF mimetype member
  must use ZIP Store compression"); companion detector grammar at `detect.rs:958`
  (`packaged_mime_layout` rejects `compression != 0`, any local extra field, any flag
  other than UTF-8-name, anything but a fixed offset-0 header; detect.rs:368-373 also
  rejects a deflated mimetype after central-directory indexing).
- **Check**: the source-backed package owner refuses any ODF whose `mimetype` ZIP entry
  is not Stored; the facade detector (`packaged_mime`, used throughout
  crates/litchi/src/detection_smart/detected.rs:239 etc.) refuses to even classify such a
  file, so it falls through to "unsupported format".
- **Spec says**: ODF 1.2 Part 3 §3.3 — the stream *shall* be first, uncompressed, without
  extra field. The checks have spec basis, but nothing requires a *reader* to reject
  violating packages.
- **Real-world behavior**: LibreOffice reads the mimetype stream through normal ZIP APIs
  (`ZipPackage::parseManifest()`, ZipPackage.cxx:473-533 — `getInputStream()` inflates
  any compression method; no stored/first/extra-field check on the read path). Re-zipped
  ODFs (`zip -r`, PHP ZipArchive, Info-ZIP with UT extra fields) are a classic real-world
  case. The crate already recognizes the class — `repair.rs:47` implements
  `odf.repair.mimetype_local_extra` — but only as an opt-in file-rewriting repair plan.
- **Suggested posture**: default-repair — on read, accept any compression method/extra
  fields/position for the mimetype entry (read it by name from the central directory) and
  record (`mimetype-deflated`, `mimetype-extra-field`, `mimetype-not-first`). Keep the
  strict grammar only for the offset-0 sniff fast path, with the tolerant path as fallback
  rather than `None`.
### 9.2 Missing (or empty-media-type) manifest "/" root entry is a fatal MIME "mismatch" — P1
- **Location**: `crates/litchi-odf-common/src/core/package.rs:763` ("ODF manifest root
  MIME type '{}' does not match mimetype '{}'").
- **Check**: package open hard-fails when the manifest has no root entry — the empty
  string from `unwrap_or_default()` is compared for equality with the authoritative
  mimetype member, so "absent" and "conflicting" collapse into the same fatal comparison.
- **Spec says**: ODF 1.2 Part 3 §4.2 — the manifest *should* contain a root entry; it is
  not stated that a package without one is unreadable.
- **Real-world behavior**: LibreOffice treats this as a media-type *fallback*, not an
  error: when the parsed manifest has no root media type it adopts the mimetype stream
  value (`m_bMediaTypeFallbackUsed = true`, ZipPackage.cxx:496-515), and throws only when
  both sources are non-empty and genuinely conflicting (:517-529). Minimal third-party
  generators (small Go/Python/Java writers) produce rootless manifests occasionally.
- **Suggested posture**: default-repair — when the manifest root media type is empty,
  adopt the mimetype stream value and record `manifest-root-missing`; keep the hard error
  only for genuinely conflicting non-empty values (matches LibreOffice exactly).
### 9.3 RTF: no EOF recovery — unclosed groups and trailing content are fatal — P1 (confidence: medium-high)
> **Status: FIXED (2026-10-02)** — litchi-rtf gained its `ToleranceReport` carrier (`Document::tolerance_report()`, `read::Options`); content after the root group's close is ignored and
> recorded by default (Word/LO stop there); unclosed groups at EOF close under opt-in `Leniency::TolerateUnclosedGroups`, one record each.
- **Location**: `crates/litchi-rtf/src/codec/parser/codec/content.rs:445` (any group left
  open at EOF → `RtfError::UnexpectedEof`, aborting the whole document);
  `codec/parser/codec/api.rs:555` ("RTF document contains trailing non-whitespace
  tokens").
- **Check**: a missing final `}` (or N unclosed braces at EOF) discards the entire parsed
  document; junk after the root `}` (a second appended group, a stray `}`, producer
  trailer text) likewise fails the whole file.
- **Spec says**: RTF 1.9.1 defines the document as one root group, but the reader guidance
  is to "read and ignore" unrecognized material; no basis for discarding successfully
  parsed content.
- **Real-world behavior**: Word auto-closes open groups at EOF and ignores trailing
  garbage — truncated/missing-final-brace RTF from buggy generators (Crystal Reports,
  hand-rolled mail-merge writers) is a well-known corpus Word opens. LibreOffice is
  stricter on unmatched `{` (throws WrongFormatException, rtfdocumentimpl.cxx:930-937)
  but tolerates characters after the last `}` (`RTFError::CHAR_OVER` → log and keep,
  rtfdocumentimpl.cxx:942-944). Litchi's trailing-token rejection diverges from both.
- **Suggested posture**: opt-in leniency for unclosed groups (auto-close at EOF, record
  `unclosed-groups-at-eof`); default-repair for trailing content (ignore tokens after
  root close, record `trailing-content-after-root`), which LO models as safe.
### 9.4 RTF: unknown control symbols are fatal — P1
> **Status: FIXED (2026-10-02)** — undefined control symbols are kept as inert unknown controls and recorded as `SyntaxDefect` at the backslash; spec-defined `\|`/`\:` are not recorded.
- **Location**: `crates/litchi-rtf/src/codec/lexer/codec/scanner/control.rs:82`
  (`RtfError::InvalidControlWord`).
- **Check**: a backslash followed by a non-alphabetic character that is not one of the 11
  recognized control symbols (`\|`, `\;`, `\:` — legacy Word index-entry and formula
  symbols — `\!`, `\?`, or any producer-specific symbol byte) kills the file at lexing.
- **Spec says**: RTF 1.9.1 "Control Symbol" — the reader model treats unrecognized
  control words/symbols as skippable, and `\*` destinations must be ignorable.
- **Real-world behavior**: LibreOffice dispatches *any* single-character symbol through
  `dispatchKeyword` and silently ignores unknown ones (rtftokenizer.cxx:2089-2093 and the
  unknown-keyword fallthrough at :2170-2176, which arms an `RTFSkipDestination` and
  returns OK); Word likewise ignores them. Old Word-era symbols do appear in legacy
  documents.
- **Suggested posture**: default-repair — tokenize an unrecognized control symbol as
  `ControlWord::Unknown` (or a dedicated token) and feed it into the existing
  `preserve_unknown_control` path (content.rs:25-33), recording `unknown-control-symbol`.
### 9.5 RTF: strict 0/1 toggle, parameterless-flag, and required-parameter validation — P1
> **Status: FIXED (2026-10-02)** — malformed parameter shapes get Word's interpretation (nonzero toggle = on, stray flag/destination parameter ignored, missing value → LibreOffice default),
> recorded as `ParameterDefect` at the control word; `\rtf`, `\u`, `\htmltag`, `\xmlns`, `\xmlattrns`, `\shpinst` and range checks stay strict.
- **Location**: representative sites — `crates/litchi-rtf/src/codec/parser/codec/mod.rs:136`
  ("RTF {name} accepts only 0 or 1", also :163), :1265-1273 (`associated_toggle`), :985
  ("RTF {name} does not accept a parameter"), :155-166 (`required_paragraph_bool`), and
  hard range/enum checks at :1120-1128, :1131-1153, :1160-1190 (`cgrid`, `animtext`,
  `fittext`, paper bins, `\fs` font size 1..=65535).
- **Check**: a toggle control word with parameter ≠ 0/1 (`\b2`), a flag with an
  accidental parameter (`\page2`), or a value control missing its parameter raises
  `MalformedDocument` and kills the whole file; out-of-range numeric metadata is likewise
  fatal.
- **Spec says**: RTF 1.9.1 defines toggles as 0/1 and flags as parameterless, but no spec
  basis forces a reader to abort; the robustness guidance is to ignore what cannot be
  interpreted.
- **Real-world behavior**: LibreOffice explicitly "flags ignore any parameter by
  definition" (rtftokenizer.cxx:2182-2183), applies value defaults when a parameter is
  absent (`if (!bParam) nParam = pSymbol->defValue`, :2205-2206), and passes toggle
  parameters straight through (nonzero → on). Word behaves the same. Sloppy generators
  emitting `\b2`, `\page1`, or oversized `\fs` values are not rare.
- **Suggested posture**: default-repair — toggles: `param.unwrap_or(1) != 0` (Word
  semantics), record `toggle-nonbinary-parameter`; flags: ignore the parameter, record
  `flag-unexpected-parameter`; value controls missing parameters: apply the spec default,
  record `value-missing-parameter`; range overflows: clamp and record — the mask-and-record
  pattern from litchi-doc's leniency.rs.
### 9.6 RTF: lone-surrogate `\uN` sequences are fatal — P1 (confidence: high on references; medium on frequency)
> **Status: FIXED (2026-10-02)** — every `\uN` path substitutes U+FFFD for an unpaired surrogate and records `UnicodeDefect`; Word's split non-BMP pairs now decode cleanly.
- **Location**: `crates/litchi-rtf/src/codec/parser/codec/unicode.rs:75`
  (`String::from_utf16` → `RtfError::InvalidUnicode` for the whole document); same
  strictness in styles.rs:995-999, markup.rs:104, markup.rs:700, info.rs:81,
  navigation.rs:609, protection.rs:348.
- **Check**: any unpaired UTF-16 surrogate in a `\uN` run — a truncated emoji pair, a
  producer that emits lone surrogates, or a surrogate split across `\ucN` fallback
  boundaries — aborts the entire parse.
- **Spec says**: RTF 1.9.1 defines `\uN` as a signed 16-bit Unicode value; nothing
  prescribes aborting on unpaired surrogates.
- **Real-world behavior**: Word and LibreOffice substitute U+FFFD per Unicode
  lossy-conversion rules and continue. Files with lone surrogates occur from
  Qt/macOS-adjacent producers and generators that split surrogate pairs incorrectly.
- **Suggested posture**: default-repair — decode with
  `char::decode_utf16(...).map(|r| r.unwrap_or(char::REPLACEMENT_CHARACTER))`, record
  `unpaired-surrogate-substituted`. Output matches what Word shows.
### 9.7 RTF: invalid or incomplete `\'hh` hex escapes are fatal — P1
> **Status: FIXED (2026-10-02)** — under opt-in `Leniency::TolerateMalformedHexEscapes`, a bad `\'hh` keeps its valid hex digits (LibreOffice) and is recorded; the default still rejects.
- **Location**: `crates/litchi-rtf/src/codec/lexer/codec/scanner/text.rs:16-23`
  ("Incomplete hex escape" / "Invalid hex escape" → `RtfError::InvalidUnicode`).
- **Check**: one malformed hex escape anywhere (`\'g1`, `\'f` at EOF) discards the whole
  file.
- **Spec says**: RTF 1.9.1 specifies `\'hh` as exactly two hex digits; deviations have no
  spec-mandated reader behavior.
- **Real-world behavior**: LibreOffice is explicitly tolerant: "`\'<number><junk>` means
  `\'0<number>`" — a non-hex second nibble is consumed harmlessly and the escape completes
  with one nibble (rtftokenizer.cxx:2044-2063); an escape truncated by EOF ends the parse
  loop without error. Word similarly recovers.
- **Suggested posture**: opt-in leniency (the deviation hides real text corruption, so
  recording matters): parse one nibble when the second is junk (LO semantics), substitute
  U+FFFD or skip when no nibble is valid, record `malformed-hex-escape`.
### 9.8 ODS: one malformed numeric/boolean cell value kills the whole spreadsheet — P1
- **Location**: `crates/litchi-ods/src/worksheet/codec.rs:2125-2143` (`parse_float`),
  invoked per cell at :224-242; boolean cells; date/time cells lacking the value
  attribute error at :243-248 ("date cells require office:date-value"); `positive()` for
  repeat/span counts at :2116-2123.
- **Check**: a float/currency/percentage cell whose `office:value` does not parse as a
  finite f64 (locale-comma decimals from EU ERP generators, `"NaN"`, empty string), a
  boolean cell with a value other than true/false/1/0, or a date cell missing its value
  attribute fails the *entire document open* (worksheet parsing is pass 4 of the fused
  open, crates/litchi-ods/src/open_parse.rs:83-99).
- **Spec says**: ODF 1.2 Part 1 §19.385/§19.386 define the value datatypes; no
  reader-side rejection is mandated.
- **Real-world behavior**: LibreOffice's `sax::Converter::convertDouble` failure leaves
  the cell value at 0/empty and import continues — a bad cell degrades, never aborts.
  Litchi already models unknown `office:value-type` tolerantly (`CellValue::Unknown`,
  codec.rs:249-252), so the strictness is inconsistent within the same match.
- **Suggested posture**: default-repair — on value-parse failure degrade the cell to
  `CellValue::Unknown { kind, value: Some(raw) }` (or Empty with the raw string
  preserved) and record `cell-value-parse-failure` with coordinates; missing date/time
  value → Empty, record.
### 9.9 ODS: unknown or foreign-namespace child of `table:calculation-settings` is fatal at open — P2 (confidence: high on LO posture; medium on frequency)
- **Location**: `crates/litchi-odf-common/src/calculation/codec.rs:507` and :549
  ("unsupported calculation settings child" — any child other than
  `table:null-date`/`table:iteration`, including elements in *any other* namespace);
  :512-515 (duplicate or out-of-order `null-date`/`iteration`).
- **Check**: a producer extension element inside `table:calculation-settings` (a
  `loext:`/vendor-namespaced setting) or out-of-schema-order children aborts ODS open
  (pass 2a of the fused open, open_parse.rs:85-86).
- **Spec says**: ODF 1.2 Part 1 §3.3 — elements in namespaces not defined by the
  specification *shall be ignorable* by consumers; the schema's extension points rely on
  this.
- **Real-world behavior**: LibreOffice's Calc XML import skips unknown children of
  settings elements. Files carrying vendor-extended settings (LibreOffice extensions,
  WPS, Collabora) occur occasionally.
- **Suggested posture**: default-repair — skip (depth-balanced) any unknown child; record
  `unknown-calculation-settings-child`. Accept `null-date`/`iteration` in any order,
  record `calculation-settings-reordered`.
### 9.10 ODP: unknown animation element kills the whole presentation — P2 (confidence: high on LO posture; medium on frequency)
- **Location**: `crates/litchi-odp/src/codec/parser/codec/xml/codec.rs:1203-1208`
  (`Element::UnknownAnimation if in_slide` → hard error); same posture recursively at
  :52 ("anim:{} contains a non-animation element") and :63-68 (disallowed child kind).
- **Check**: any animation-namespace element outside the closed `Kind` table — newer SMIL
  constructs, vendor animation extensions, or future ODF versions — aborts the entire
  deck's parse; one slide's animation poisons every slide.
- **Spec says**: same ODF foreign-element ignorability rule (ODF 1.2 Part 1 §3.3); the
  SMIL profile is explicitly extensible.
- **Real-world behavior**: LibreOffice's Impress import skips animation nodes it cannot
  map and keeps the slide. The same file already ignores unknown animation elements in
  *notes* (`_ if in_notes => {}`, codec.rs:1202), so the slide path is an inconsistent
  strictness pocket.
- **Suggested posture**: default-repair — skip the unknown animation subtree
  (depth-balanced), drop that one animation from the slide's list, record
  `unknown-animation-element` with slide index.
### 9.11 ODF: DOCTYPE declaration in content.xml is fatal — P2 (confidence: high on LO tolerance; medium on frequency)
- **Location**: `crates/litchi-odf-common/src/core/content_validation.rs:222`
  ("{family_name} content.xml must not contain a doctype"; driven from
  crates/litchi-odt/src/document/open_parse.rs:130 and the ODS/ODP equivalents); repeated
  in scoped codecs, e.g. crates/litchi-odt/src/font_face/codec.rs:196, :249-251.
- **Check**: any `<!DOCTYPE …>` in a document part fails the file, even when the DTD is
  never referenced — general entity references are *already* independently validated at
  content_validation.rs:225-229, and quick-xml does not fetch external DTDs, so no
  XXE/expansion hole is being defended here.
- **Spec says**: no spec basis — ODF 1.2 documents are not DTD-validated, and XML's own
  rules let non-validating processors skip the DOCTYPE.
- **Real-world behavior**: LibreOffice's SAX import ignores DTD declarations. Old
  generator pipelines (DocBook→ODF stylesheets, flat-ODF converters, hand-assembled
  `.fodt` files) emit DOCTYPE prologues; these open fine in LO.
- **Suggested posture**: default-repair — skip the DocType event, record
  `doctype-present`. Entity safety is unaffected (undeclared `GeneralRef`s remain
  rejected).
### 9.12 RTF: unsupported `\fcharset` / `\cpg` values in the font table are fatal — P2 (confidence: high on LO fallback; medium on frequency)
- **Location**: `crates/litchi-rtf/src/codec/parser/codec/resources.rs:408` ("unsupported
  RTF font charset {charset_id}"), :442 ("unsupported RTF font code page {page}");
  companion: :518-523 (known charset without an exact page mapping + non-ASCII font name
  → error, text at :520).
- **Check**: one font-table entry with an out-of-table charset byte (e.g. `\fcharset5`,
  seen from legacy converters) or a code page the codec lacks fails the whole document.
- **Spec says**: RTF 1.9.1 enumerates charset values but prescribes no reader abort;
  unknown values fall outside spec.
- **Real-world behavior**: LibreOffice looks the charset up in a fixed table
  (rtfcharsets.cxx:19-52) and simply uses the document default encoding when no row
  matches — no error path exists; Word likewise falls back to the default code page.
- **Suggested posture**: default-repair for unknown charset/cpg: treat as
  `FontCharset::Default` / header encoding, record `unsupported-font-charset` /
  `unsupported-font-codepage`. Keep the charset-without-exact-page + non-ASCII-name case
  as opt-in at most (guessing there could produce genuinely wrong font-name text) —
  though falling back to the header encoding exactly like LO would still be defensible.

## Iteration roadmap

Phase 0 (prerequisite, blocks everything else): introduce a `ToleranceReport` carrier in
every crate that lacks one — docx, xlsx, ppt, pptx, cfb, ole-common, opc, ooxml-common,
rtf, odf-common — modeled on `crates/litchi-doc/src/leniency.rs` (closed defect enums,
bounded recording with `MAX_RECORDED_DEFECTS`, `unrecorded()` overflow count, merge
support). Thread litchi-xls's existing `XlsLeniency` into the source-backed open, which
today hardcodes `Leniency::Strict` (`crates/litchi-xls/src/workbook/source.rs:2114`).

**Phase 1 — P0 items, default-repair + tolerance-record.** Eight items; each has a
verified real-world producer class and a reference-app repair. In order: litchi-doc F1.1
(auxiliary-table fatal propagation), litchi-docx F2.1 (foreign `*:t` on OMML), litchi-xls
F3.1 (per-cell XF validation), litchi-ppt F5.1 (StyleTextPropAtom coverage — the lenient
parser already exists in-crate), litchi-ppt F5.2 (BulletFlags/PFWrapFlags reserved bits),
litchi-pptx F6.1 (animation targets/duplicate IDs), litchi-doc F1.2 (Stshif reserved
bits), ODF F9.1 (mimetype member grammar). The reserved-bit items land as new variants of
the existing `SprmDefect`-style closed enums — the same shape as
`SprmDefect::CellBorderSideFlags`.

**Phase 2 — P1 items: default-repair where clearly Word-compatible, else opt-in.**
Default-repair class (reference apps provably tolerate; repair cannot invent content):
reserved-bit masking bundles (doc F1.3-F1.5, F1.9; xls F3.4, F3.6, F3.7; xlsb F3.11),
per-item degradation bundles (docx F2.2-F2.8; xlsx F4.1-F4.7; pptx F6.2, F6.4-F6.7; cfb
F7.1-F7.4; opc F8.1-F8.9; rtf F9.4-F9.6; ODS F9.8), lossy-decode items (xls F3.2, F3.8;
ole-common F7.8-F7.10; rtf F9.6), and the substrate cross-reference items (cfb F7.1-F7.4;
opc F8.5-F8.9; ppt F5.4-F5.7, F5.11). Opt-in class (repair involves a judgment call or
ambiguous ground truth): xls F3.3 (duplicate sheet names), xlsx F4.8 (sheet-name rules),
xlsb F3.10 (NaN Xnum cache), rtf F9.3 (unclosed groups at EOF), rtf F9.7 (malformed hex
escapes).

**Phase 3 — P2 items: opt-in only, or defer.** The remaining 33 items are lower-frequency
or mixed-reference strictness (opc F8.10-F8.12 where POI shares the strictness; cfb
F7.5-F7.7, F7.12; ppt F5.8-F5.10, F5.12; pptx F6.3, F6.8-F6.12; docx F2.9-F2.12; doc
F1.10-F1.12; xls F3.12; xlsx F4.9-F4.12; ODF F9.9-F9.12). Sequence them behind the Phase
1-2 plumbing; several (doc F1.12 silent piece-table loss, pptx F6.11 view-props forward
compatibility) are worth pulling forward once the tolerance report exists, because their
repair is trivial once recording is possible.

Acceptance pattern per fix (all four required):

1. The file class that previously failed now opens **by default** (or under the named
   opt-in `Leniency` variant where the finding says opt-in).
2. The repair is enumerable: the crate's tolerance report lists the defect class and
   location context; an unmodified file produces an empty report. Nothing is swallowed
   silently.
3. Regression tests: a fixture per defect class (hand-built minimal file carrying the
   deviation) asserting open + repair + recorded defect, and a test that the same file
   without the deviation repairs nothing.
4. The crate's existing test suite stays green; structural-corruption and security-limit
   rejections (framing, encryption, resource bounds) are untouched and still fatal.

## Appendix A. Dropped and corrected findings

**Dropped findings: none.** All 108 findings across the nine scopes were verified against
the working tree on 2026-10-01: every cited file exists, and every quoted check or error
string was found at or within a few lines of the cited position. No finding duplicated
the already-fixed cell-border side-flags case (`SprmDefect::CellBorderSideFlags`) or the
already-documented `StylesheetDefect::DuplicateStyleName`; both are cited as the
established repair pattern instead (see "The tolerance contract").

**Corrections applied during verification** (content present, citation imprecise):

- Line numbers were commonly off by 1-4 lines across all nine reports (agents cited block
  starts or slightly stale snapshots). All line numbers in this document are the verified
  ones. Examples: docx `text.rs` foreign-`t` sites cited as :923/:976/:1041, actual
  :925/:978/:1043; ODF mimetype check cited package.rs:707, actual :709; ole-common PID-1
  error text cited model.rs:246, actual :247; ppt `text_prop.rs` coverage errors cited
  :374/:406, actual :376/:408.
- xls F3.2: the quoted "unsupported code page {page}" error text lives in
  `crates/litchi-codepage/src/lib.rs:41`, not in the cited call-site file; the call sites
  (`workbook/source.rs:1997-2006`, `codec/semantic/globals.rs:199-204`) are as cited.
- pptx F6.1: the target-validation flag is read in
  `animations/codec/xml/parser/semantic.rs` (checks at :1117-1141); the cited
  `animations/codec/semantic.rs:82-85` is where `parse_slide_xml` enables it. Both halves
  verified.
- ODF scope: the duplicate-manifest-entry check was cited as "package/codec.rs:347-350";
  the file is `crates/litchi-odf-common/src/package/codec.rs` (check at :347-352,
  verified, with a pinning test at `package/tests.rs:48-50`).
- litchi-crypto: the scope summary's quoted "UNICODE-LP-P4 padding is nonzero" is at
  `crates/litchi-crypto/src/spaces.rs:974` as cited (verified).

## Appendix B. Areas reviewers found already appropriately tolerant

- **litchi-doc**: FIB degrades gracefully on truncated pointer arrays (fib.rs:287-326);
  text extraction lossy with CLX fallback (text.rs:99-146); FKP/PLCF `Option`-based;
  unknown SPRM opcodes ignored in CHP/SEP walkers; several auxiliary tables `.ok()` or
  deferred `Result` (package.rs:173-177, 310-358); malformed MTEF → placeholder strings
  (document/codec.rs:136-148).
- **litchi-docx**: lazy open (only main-part content type checked); unknown body children
  → `OpaqueBlock`; unknown run contents/inlines → `RunContent::Unknown`/`Inline::Unknown`;
  unknown field types get opaque models; unknown settings/numbering children skipped;
  style attribute parsing first-wins and lossy; `Paragraph::text()` already skips foreign
  subtrees.
- **litchi-biff / litchi-xls**: unknown records skipped and preserved
  (worksheet.rs:925-927); ExtSST malformation lazy non-fatal; malformed
  Dimensions/MergeCells/AutoFilter ignored; duplicate cells tolerated; out-of-range
  `LabelSst` → error cell (cell.rs:135-141); RK decoding infallible; SST walk ignores
  reserved string flags and phonetic markers per spec. litchi-biff is a pure framing
  layer whose only refusals are truncation and configurable security limits.
- **litchi-xlsx**: unknown cell types → `Cell::Unknown` (semantic.rs:41-51); unknown
  formula types → `RawFormulaKind::Unknown`; unknown visibility → `Visibility::Unknown`
  (catalog.rs:613); unknown error values → `ErrorValue::Unknown`; `dimension` stored but
  never cross-checked; MCE/Strict namespaces accepted; foreign-namespace elements
  ignored; the alternate shared-strings parser documents and tests the tolerant hint
  posture (raw/shared_strings.rs:219-235).
- **litchi-ppt**: record loop and container parsing resynchronize byte-by-byte on garbage
  (parsers/parser.rs:110-116); unknown record types → `RecordType::Unknown` with raw
  payload preserved (consts.rs:615); truncated containers degrade; text atoms decode
  lossy (text/extractor.rs:32); notes index degrades instead of failing
  (slide/notes.rs:53-59); a clamping POI-style StyleTextPropAtom parser exists (unused —
  finding 5.1).
- **litchi-pptx**: unknown transition effects preserved inertly as `Kind::Raw`
  (transition/reader.rs:456-461); shape-like extension elements → `Kind::Unknown`
  (shape/reader.rs:643-649); unknown chart plot types → `Type::Unknown`; `extLst`
  payloads preserved opaquely; MCE baseline processing in every reader; `dur` strips an
  "ms" suffix; `on`/`off` booleans accepted; two readers already ignore PIs.
- **litchi-cfb / ole-common**: header DIFAT tail garbage ignored (file.rs:1192-1197);
  FAT padding past EOF marked `ENDOFCHAIN` accepted with a comment citing Word/POI
  (file.rs:1651-1657); short final sector reads as zeroes; v3 stream-size high DWORD
  masked per MS-CFB 2.6.1; classic-Mac `00 52` root name canonicalized; unbalanced
  red-black color metadata tolerated; unknown top-level VT types → `Value::Unknown`;
  nonzero `VARIANT_BOOL` → true. `litchi-codepage` already ships the lossy decode
  machinery (`decode_lossy`, `Mbcs::recover`).
- **litchi-opc**: unmappable ZIP members become reported junk (members.rs:25-43,
  pkgreader.rs:1554-1578); missing `.rels` parses as empty; dangling relationship targets
  stay visible (pkgreader.rs:1831-1837); content-type mappings fold ASCII case; duplicate
  ZIP names last-write-wins; query/fragment stripped from internal targets; part lookup
  tolerates case-mismatched references; digital signatures never validated at open —
  verification is opt-in (package.rs:2038-2052).
- **litchi-rtf**: unknown control words and destinations preserved as opaque nodes
  (content.rs:25-33, resources.rs:915-923); color-table components clamped and
  unterminated entries accepted (resources.rs:891-945); body text decodes lossily.
- **ODF family**: unknown manifest elements skipped; unknown cell value-types preserved
  (`CellValue::Unknown`); styles.xml/meta.xml optional; an opt-in repair-plan module
  already exists for the mimetype-extra-field deviation class
  (odf-common/src/repair.rs). The missing-manifest hard error matches LibreOffice's
  default and was deliberately not reported as a finding.
