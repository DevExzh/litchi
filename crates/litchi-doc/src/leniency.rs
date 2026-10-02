//! Repair reporting for non-structural defects in binary documents.
//!
//! Two kinds of repair can happen while a document is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the two closed enumerations changes: piece tables, FIB structure, stream
//! grammar, and encryption remain hard errors under every variant.
//!
//! **Word-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover bytes MS-DOC leaves
//! undefined but real producers emit anyway — bytes Word and LibreOffice read
//! past without complaint, so rejecting them would cost the caller the entire
//! document over metadata that cannot change what a conforming file means.
//! Each substitutes the value documented on the corresponding variant of a
//! [`Defect`] class other than [`Defect::Stylesheet`]. The established
//! examples are masking undefined cell-border side flags out of
//! `sprmTSetBrc80`/`sprmTSetBrc`, ignoring the Stshif reserved bits MS-DOC
//! 2.9.274 tells readers to ignore, and degrading a malformed optional
//! FIB-pointed metadata table to its absent value instead of failing the
//! whole document open.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. MS-DOC 2.9 states that a style's name and each of
//! its aliases "MUST NOT be empty and MUST be unique within all names in the
//! stylesheet". Word enforces that when it writes, but files produced by
//! other tools — and by localized Word builds — do reach users with a
//! duplicated name, and Word opens them without complaint. [`Leniency::Strict`]
//! rejects such a file; [`Leniency::TolerateStylesheetDefects`] repairs it
//! with the substitution documented on the corresponding [`StylesheetDefect`]
//! variant.
//!
//! This mirrors the contract `XlsLeniency` provides for legacy workbooks so the
//! two legacy readers behave the same way.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous stylesheet must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural document defects.
///
/// Structural defects — malformed FIB or piece-table framing, a broken stream
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant. The Word-compatible repairs listed by every [`Defect`]
/// class other than [`Defect::Stylesheet`] are applied and recorded under
/// every variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: the [`StylesheetDefect`] classes are hard errors,
    /// while the Word-compatible repairs (every other [`Defect`] class) are
    /// still applied and recorded in the [`ToleranceReport`].
    #[default]
    Strict,
    /// Repair and record the defects listed by [`StylesheetDefect`].
    TolerateStylesheetDefects,
}

impl Leniency {
    /// Whether stylesheet defects are repaired rather than rejected.
    #[inline]
    #[must_use]
    pub fn tolerates_stylesheet_defects(self) -> bool {
        matches!(self, Self::TolerateStylesheetDefects)
    }
}

/// The class of stylesheet defect a lenient read is allowed to repair.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// [`Leniency::TolerateStylesheetDefects`] may change about a document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StylesheetDefect {
    /// A style name or alias repeats one already used in the stylesheet.
    ///
    /// MS-DOC 2.9 requires uniqueness. The duplicate is kept as-is on the style
    /// that carries it — no name is rewritten or discarded — because the name
    /// only labels the style; lookup by index, which is how every stored
    /// reference works, is unaffected.
    DuplicateStyleName,
}

impl StylesheetDefect {
    /// A short, stable description of what a lenient read did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::DuplicateStyleName => "kept the duplicated name; styles are resolved by index",
        }
    }
}

/// The class of table SPRM defect the reader repairs the way Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each one matches what Word and LibreOffice do with
/// bytes MS-DOC leaves undefined, so applying it never changes what a
/// conforming document means.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SprmDefect {
    /// `sprmTSetBrc80` (0xD620) or `sprmTSetBrc` (0xD62F) carried
    /// `bordersToApply` side bits outside the set its operand variant defines.
    ///
    /// MS-DOC 2.9.304 defines only bits 0x01–0x08 on the legacy 7-byte
    /// `TableBrc80Operand`, and MS-DOC 2.9.305 defines bits 0x01–0x20 on the
    /// 11-byte `TableBrcOperand`, but producers set the diagonal bits on the
    /// legacy operand — and stray bits on either — anyway. Word and
    /// LibreOffice test the defined bits and ignore the rest, so the reader
    /// masks `bordersToApply` down to the defined set (0x0F on the legacy
    /// operand, 0x3F on the full one) and still applies the border to every
    /// side whose bit survives.
    CellBorderSideFlags,
    /// A table or cell border (`Brc80`/`Brc`, in any table SPRM or `TC80`)
    /// carried a `brcType` outside the defined set.
    ///
    /// MS-DOC 2.9.17's restriction is a writer constraint. LibreOffice maps
    /// an unknown type to no line (`ConvertBorderStyleFromWord`'s default)
    /// and POI stores it raw, so the reader decodes it as
    /// `BorderType::None`, keeping the border's other fields.
    BorderType,
    /// A legacy table or cell border (`Brc80`) carried an `ico` color index
    /// above 16.
    ///
    /// No color is defined for the index, so the border reads as automatic
    /// color — the same value as `ico` 0 — and is otherwise applied.
    BorderColor,
}

impl SprmDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::CellBorderSideFlags => {
                "masked bordersToApply to the side bits the operand defines; surviving sides were still applied"
            },
            Self::BorderType => "read the undefined border type as no line",
            Self::BorderColor => "read the undefined border color index as automatic color",
        }
    }
}

/// The class of stylesheet-header (`Stshif`) defect the reader repairs the
/// way Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// Stshif repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. The stylesheet is mandatory in every Word 97+ file,
/// so rejecting it over bytes MS-DOC tells readers to ignore would cost the
/// whole document; LibreOffice masks the flag word and reads on
/// (`ww8scan.cxx:6905`). The recorded [`ToleratedDefect::context`] is the
/// offending 16-bit word as stored in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StshifDefect {
    /// The 16-bit style-name flag word was not exactly `0x0001`.
    ///
    /// MS-DOC 2.9.274: `fStdStylenamesWritten` (bit 0) "MUST be 1 and MUST be
    /// ignored" and `fReserved` (bits 1-15) "MUST be 0 and MUST be ignored".
    /// The word carries no information the reader acts on, so only the
    /// deviation is recorded.
    StyleNameFlags,
    /// `Stshif.istdMaxFixedWhenSaved` was not 15.
    ///
    /// Fixed style slots are defined by the format, not by the count, and the
    /// count is never consulted for layout, so the stored value is kept
    /// verbatim in `StyleSheetHeader::fixed_style_count`.
    FixedStyleCount,
}

impl StshifDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::StyleNameFlags => {
                "ignored every style-name flag bit except fStdStylenamesWritten, as MS-DOC 2.9.274 requires"
            },
            Self::FixedStyleCount => {
                "kept the stored count; fixed style slots are defined by the format, so the value is never consulted"
            },
        }
    }
}

/// The optional FIB-pointed auxiliary table whose parse failure was degraded
/// instead of aborting the document open.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// which metadata tables a read may drop. MS-DOC 2.5 makes each of these
/// fc/lcb pairs an independently optional metadata table; a malformed one
/// invalidates nothing else, and Word and LibreOffice open the document and
/// ignore the broken sub-structure. The degraded table reports the same
/// empty/absent value as a document that simply does not carry the table, so
/// no accessor ever panics or invents content. The recorded
/// [`ToleratedDefect::context`] is the table's FIB `rgFcLcb` pointer index
/// (for multi-table groups, the index of the group's primary table), except
/// for [`Self::RevisionAuthorResolution`], whose context is the index of the
/// style whose author could not be resolved.
///
/// Tables whose bytes feed structural framing — the FIB itself, the piece
/// table/CLX, encryption material, and stylesheet framing — are not here:
/// they remain hard errors under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum AuxiliaryTableDefect {
    /// The story field PLCFs (`PlcffldMom` and its siblings, FIB index 16
    /// and onward) failed to parse. The document reports no fields and no
    /// field-derived hyperlinks.
    Fields,
    /// The comment tables (`PlcfandRef`/`PlcfandTxt`, primary FIB index 4)
    /// failed to parse. The document reports no comments.
    Comments,
    /// The standard bookmark tables (`SttbfBkmk`/`PlcfBkf`/`PlcfBkl`, primary
    /// FIB index 21) failed to parse. The document reports no bookmarks.
    Bookmarks,
    /// The smart-tag bookmark and recognizer tables (primary FIB index 114)
    /// failed to parse. The document reports no smart tags.
    SmartTags,
    /// The revision-save identifier table (`PLRSID`, FIB index 113) failed
    /// to parse. The document reports no rsids.
    Rsids,
    /// The e-mail review threading table (`RmdThreading`, FIB index 94)
    /// failed to parse. The document reports no threading data.
    RmdThreading,
    /// The embedded font description table (`SttbTtmbd`, FIB index 61)
    /// failed to parse. The document reports no embedded fonts.
    EmbeddedFonts,
    /// The `MsoEnvelopeCLSID` range (FIB index 97) failed to parse. The
    /// document reports no envelope.
    Envelope,
    /// The `AutoSummary` table (`PlcfAsumy`, FIB index 89) failed to parse.
    /// The document reports no auto-summary ranges.
    AutoSummary,
    /// The range-level protection tables (primary FIB index 141) failed to
    /// parse. The document reports no protected ranges.
    ProtectedRanges,
    /// The format consistency-checker bookmark tables (primary FIB index
    /// 120) failed to parse. The document reports no consistency marks.
    FormatConsistency,
    /// The structured document tag bookmark tables (primary FIB index 137)
    /// failed to parse. The document reports no structured tags.
    StructuredTags,
    /// The XML schema reference table (`Hplxsdr`, FIB index 136) failed to
    /// parse. The document reports no XML schema references.
    XmlSchemas,
    /// The custom XML save transform path (`fcCustomXForm`, FIB index 140)
    /// failed to parse. The document reports no transform path.
    CustomXmlTransform,
    /// The OLE control table (`RgxOcxInfo`, FIB index 85) failed to parse.
    /// The document reports no OLE controls.
    OleControls,
    /// The mail-merge state (`Pms` and the ODSO property set, primary FIB
    /// index 44) failed to parse. The document reports no mail-merge state.
    MailMerge,
    /// The master-document subdocument directory (`PlcfWKB`/`SttbFnm`,
    /// primary FIB index 54) failed to parse. The document reports no
    /// subdocuments.
    Subdocuments,
    /// The revision-mark author table (`SttbfRMark`, FIB index 51) failed to
    /// parse. The document reports no revision authors; revision marks that
    /// name an author stay unresolved rather than aborting the read.
    RevisionAuthors,
    /// The fixed associated-string table (`SttbfAssoc`, FIB index 32) failed
    /// to parse. The document reports no associated strings.
    AssociatedStrings,
    /// The list-name table (`SttbListNames`, FIB index 91) failed to parse.
    /// The document reports no list names.
    ListNames,
    /// The list-level template table (`SttbRgtplc`, FIB index 96) failed to
    /// parse. The document reports no list templates.
    ListTemplates,
    /// The section table (`PlcfSed`, FIB index 6) failed to parse. The
    /// document reports no section layout; text and property extraction are
    /// unaffected because section ranges feed only layout queries.
    Sections,
    /// A revision-marked style named an author index outside `SttbfRMark`.
    /// The style keeps its numeric author index and its author name is left
    /// unresolved.
    RevisionAuthorResolution,
}

impl AuxiliaryTableDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Fields => "discarded the malformed field tables; the document reports no fields",
            Self::Comments => {
                "discarded the malformed comment tables; the document reports no comments"
            },
            Self::Bookmarks => {
                "discarded the malformed bookmark tables; the document reports no bookmarks"
            },
            Self::SmartTags => {
                "discarded the malformed smart-tag tables; the document reports no smart tags"
            },
            Self::Rsids => "discarded the malformed PLRSID table; the document reports no rsids",
            Self::RmdThreading => {
                "discarded the malformed RmdThreading table; the document reports no threading data"
            },
            Self::EmbeddedFonts => {
                "discarded the malformed SttbTtmbd table; the document reports no embedded fonts"
            },
            Self::Envelope => {
                "discarded the malformed envelope range; the document reports no envelope"
            },
            Self::AutoSummary => {
                "discarded the malformed PlcfAsumy table; the document reports no auto-summary ranges"
            },
            Self::ProtectedRanges => {
                "discarded the malformed protection tables; the document reports no protected ranges"
            },
            Self::FormatConsistency => {
                "discarded the malformed consistency-mark tables; the document reports no marks"
            },
            Self::StructuredTags => {
                "discarded the malformed structured-tag tables; the document reports no structured tags"
            },
            Self::XmlSchemas => {
                "discarded the malformed Hplxsdr table; the document reports no XML schema references"
            },
            Self::CustomXmlTransform => {
                "discarded the malformed transform path; the document reports no custom XML transform"
            },
            Self::OleControls => {
                "discarded the malformed RgxOcxInfo table; the document reports no OLE controls"
            },
            Self::MailMerge => {
                "discarded the malformed mail-merge state; the document reports no mail merge"
            },
            Self::Subdocuments => {
                "discarded the malformed subdocument directory; the document reports no subdocuments"
            },
            Self::RevisionAuthors => {
                "discarded the malformed SttbfRMark table; the document reports no revision authors"
            },
            Self::AssociatedStrings => {
                "discarded the malformed SttbfAssoc table; the document reports no associated strings"
            },
            Self::ListNames => {
                "discarded the malformed SttbListNames table; the document reports no list names"
            },
            Self::ListTemplates => {
                "discarded the malformed SttbRgtplc table; the document reports no list templates"
            },
            Self::Sections => {
                "discarded the malformed PlcfSed table; the document reports no section layout"
            },
            Self::RevisionAuthorResolution => {
                "left the revision author unresolved; the style keeps its numeric author index"
            },
        }
    }

    /// The FIB `rgFcLcb` pointer index recorded as the defect context.
    ///
    /// [`Self::RevisionAuthorResolution`] is recorded with the index of the
    /// style whose author could not be resolved; the index returned here is
    /// the `SttbfRMark` table the resolution references.
    pub(crate) const fn fib_index(self) -> u16 {
        match self {
            Self::Fields => 16,
            Self::Comments => 4,
            Self::Bookmarks => 21,
            Self::SmartTags => 114,
            Self::Rsids => 113,
            Self::RmdThreading => 94,
            Self::EmbeddedFonts => 61,
            Self::Envelope => 97,
            Self::AutoSummary => 89,
            Self::ProtectedRanges => 141,
            Self::FormatConsistency => 120,
            Self::StructuredTags => 137,
            Self::XmlSchemas => 136,
            Self::CustomXmlTransform => 140,
            Self::OleControls => 85,
            Self::MailMerge => 44,
            Self::Subdocuments => 54,
            Self::RevisionAuthors | Self::RevisionAuthorResolution => 51,
            Self::AssociatedStrings => 32,
            Self::ListNames => 91,
            Self::ListTemplates => 96,
            Self::Sections => 6,
        }
    }
}

/// The class of style-record (`STD`) defect the reader repairs the way Word
/// repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// per-style repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. The stylesheet is mandatory, so rejecting
/// it over metadata bits MS-DOC tells readers to ignore would cost the whole
/// document; LibreOffice reads `fAutoRedef`/`fHidden` and skips the rest of
/// the word (`ww8scan.cxx:7007-7013`). The raw `STD` bytes are preserved
/// verbatim for round-trip, so a repair only affects the decoded model. The
/// recorded [`ToleratedDefect::context`] is the style index (`istd`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StyleDefect {
    /// `StdfBase.grfstd` set a reserved bit (`0x0080` or `0xE000`).
    ///
    /// MS-DOC 2.9.112: the reserved bits "MUST be 0 and MUST be ignored".
    /// The decoded `StyleFlags` read only the defined bits.
    ReservedStyleFlags,
    /// `StdfPost2000` set a reserved bit (`info1` bits 13-15 or the unused
    /// `info3` bit 3).
    ///
    /// MS-DOC 2.9.261: the reserved bits "MUST be 0 and MUST be ignored".
    /// The decoded `StylePost2000` reads only the defined bits.
    ReservedPost2000Flags,
    /// `StdfPost2000.iPriority` was above 99.
    ///
    /// The priority only orders styles in the UI, so the decoded value is
    /// clamped to 99 — the lowest priority, which Word assigns to styles with
    /// no explicit one.
    Priority,
}

impl StyleDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedStyleFlags => {
                "ignored the reserved GRFSTD bits; only the defined style flags were decoded"
            },
            Self::ReservedPost2000Flags => {
                "ignored the reserved StdfPost2000 bits; only the defined fields were decoded"
            },
            Self::Priority => "clamped the out-of-range style priority to 99",
        }
    }
}

/// The class of comment-descriptor (`ATRDPre10`) defect the reader repairs
/// the way Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// per-comment repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. Each repair is confined to the one comment
/// that carries it; the rest of the comment table is read unchanged. The
/// recorded [`ToleratedDefect::context`] is the comment's index in
/// `PlcfandRef` (saturating at `u16::MAX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CommentDefect {
    /// `ATRDPre10.bitsNotUsed` or `grfNotused` was non-zero.
    ///
    /// MS-DOC 2.9.7: both "MUST be 0, and MUST be ignored". They carry no
    /// information the reader acts on, so only the deviation is recorded.
    ReservedFields,
    /// `ATRDPre10.ibst` named an author outside the comment-owner array.
    ///
    /// The index is a cross-reference, not framing. The stored index is
    /// kept in `CommentDescriptor::author_index` and the comment's resolved
    /// author is the empty string — no author is invented — mirroring Word,
    /// which shows such a comment with a generic author.
    AuthorIndex,
    /// `ATRDPre10.lTagBkmk` was below -1.
    ///
    /// Such a tag cannot name an annotation bookmark, so the comment is read
    /// as a point comment with no commented range, as if the tag were -1.
    BookmarkTag,
}

impl CommentDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedFields => "ignored the non-zero ATRDPre10 reserved fields",
            Self::AuthorIndex => {
                "kept the out-of-range author index; the comment's author is left empty"
            },
            Self::BookmarkTag => {
                "read the comment as a point comment; the tag below -1 names no bookmark"
            },
        }
    }
}

/// The class of standard-bookmark defect the reader repairs the way Word
/// repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// per-bookmark repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. Each repair is confined to the one
/// bookmark that carries it; the framing of `SttbfBkmk`/`PlcfBkf`/`PlcfBkl`
/// (presence, sizes, counts, CP ordering) stays a hard error for the table.
/// The recorded [`ToleratedDefect::context`] is the bookmark's index in
/// `SttbfBkmk`/`PlcfBkf` (saturating at `u16::MAX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BookmarkDefect {
    /// The `BKC.fPub` bit was set.
    ///
    /// MS-DOC 2.9.8: fPub "MUST be 0, and MUST be ignored". The bookmark is
    /// kept and decoded from its other bits.
    ReservedPublicFlag,
    /// `FBKF.ibkl` was out of range or already claimed by an earlier
    /// bookmark. The bookmark has no end of its own, so it is dropped; the
    /// first bookmark to claim an end entry keeps it.
    EndIndex,
    /// The bookmark's start CP lay after its end CP. The bookmark cannot
    /// denote a range, so it is dropped.
    ReversedRange,
    /// A column bookmark (`BKC.fCol`) had an empty or reversed column range.
    /// The bookmark's text range is still valid, so it is kept as an
    /// ordinary text bookmark with no column restriction.
    ColumnRange,
}

impl BookmarkDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedPublicFlag => "ignored the reserved BKC fPub bit",
            Self::EndIndex => {
                "dropped the bookmark whose end index was out of range or already claimed"
            },
            Self::ReversedRange => "dropped the bookmark whose start followed its end",
            Self::ColumnRange => {
                "kept the bookmark as a text bookmark; its column range was empty or reversed"
            },
        }
    }
}

/// The class of table-definition (`sprmTDefTable`) defect the reader
/// repairs the way Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies to the cell boundaries (`rgdxaCenter`) and cell
/// descriptors (`TC80`, MS-DOC 2.9.313) of a table row under every
/// [`Leniency`] variant, including [`Leniency::Strict`]. MS-DOC defines the
/// used encodings of each field but places no reader-side rejection on the
/// unused ones; Word clamps or ignores them and LibreOffice reads the flags
/// bit by bit, so each repair substitutes the field's neutral member and
/// leaves the rest of the row intact. Operand framing — the 63-column limit,
/// truncated boundaries or descriptors, and decreasing boundaries — stays a
/// hard error. The recorded [`ToleratedDefect::context`] is the zero-based
/// cell index (`itc`) within the row, or the boundary index for
/// [`Self::BoundaryRange`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TableDefinitionDefect {
    /// A cell boundary lay outside the XAS range (-31680 to 31680 twips).
    /// It is clamped to the nearest end of the range.
    BoundaryRange,
    /// `TC80.textFlow` used an undefined value (2, 6, or 7). The cell reads
    /// as left-to-right, top-to-bottom text (`grpfTFlrtb`).
    TextFlow,
    /// `TC80.vertMerge` used the undefined value 2. The cell reads as not
    /// vertically merged.
    VerticalMerge,
    /// `TC80.vertAlign` used the undefined value 3. The cell reads as
    /// top-aligned.
    VerticalAlignment,
    /// `TC80.ftsWidth` used an undefined unit (4 through 7). The cell reads
    /// as having no preferred width.
    WidthUnit,
    /// `TC80.wWidth` was negative. The cell reads as having no preferred
    /// width.
    NegativeWidth,
}

impl TableDefinitionDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::BoundaryRange => "clamped the cell boundary to the XAS range",
            Self::TextFlow => "read the undefined text flow as left-to-right, top-to-bottom",
            Self::VerticalMerge => "read the undefined vertical merge state as not merged",
            Self::VerticalAlignment => "read the undefined vertical alignment as top",
            Self::WidthUnit => "read the undefined width unit as no preferred width",
            Self::NegativeWidth => "read the negative preferred width as no preferred width",
        }
    }
}

/// The class of paragraph-property (`PapxFkp`) defect the reader repairs the
/// way Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// paragraph bin-table repairs a read applies under every [`Leniency`]
/// variant, including [`Leniency::Strict`]. MS-DOC 2.4.3 defines each
/// `PapxFkp` page independently, so one bad page or run never costs the
/// others; Word and LibreOffice render the affected paragraphs with default
/// formatting. The `PlcfBtePapx` page index itself is framing. The recorded
/// [`ToleratedDefect::context`] is the page's index in `PlcfBtePapx`
/// (saturating at `u16::MAX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PapxDefect {
    /// A `PapxFkp` page lay outside the `WordDocument` stream or was
    /// malformed. The page is skipped; its paragraphs read with default
    /// paragraph properties.
    Page,
    /// One run's `PapxInFkp`/`GrpPrlAndIstd` could not be read — a
    /// malformed entry, a SPRM sequence the parser cannot consume exactly,
    /// or an unresolvable data indirection. The run reads with default
    /// paragraph properties; the other runs on the page are unaffected.
    Run,
    /// A run's `GrpPrlAndIstd` ended in the one zero pad byte a repeated
    /// producer emits after a complete odd-length SPRM sequence. The byte —
    /// accepted only when the strict SPRM parser proves it is exactly one
    /// incomplete trailing opcode byte — is dropped and the run is read.
    AlignmentPadding,
}

impl PapxDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Page => {
                "skipped the malformed PAPX FKP page; its paragraphs use default properties"
            },
            Self::Run => "read the malformed PAPX run with default paragraph properties",
            Self::AlignmentPadding => {
                "dropped the producer alignment pad byte after the PAPX SPRMs"
            },
        }
    }
}

/// The class of section-property (`SEPX`) defect the reader repairs the way
/// Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// per-SPRM section repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. MS-DOC's section enumerations are
/// defined-value lists with no reader-side rejection rule, and Word reads
/// past undefined bytes with defaults. SEPX framing — a truncated SEPX, a
/// negative grpprl size, an operand of the wrong size, or a SPRM sequence
/// that does not fill the SEPX — stays a hard error, as do numeric range
/// checks. The recorded [`ToleratedDefect::context`] is the SPRM opcode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionDefect {
    /// A page border's reserved `Brc80` effect bit was set, or the
    /// `sprmSPgbProp` reserved byte was non-zero.
    ///
    /// MS-DOC 2.9.17: the Brc80 reserved bit "MUST be zero, and MUST be
    /// ignored". The bits carry no information; the rest of the operand is
    /// decoded unchanged.
    ReservedBits,
    /// A page border (`sprmSBrc*80`) carried a `brcType` that names neither a
    /// line style nor a border art. The side reads as having no border, as
    /// LibreOffice reads it.
    BorderType,
    /// A page border carried an `ico` above 16. The border reads with
    /// automatic color.
    BorderColor,
    /// A section SPRM (or one `sprmSPgbProp` placement field) carried a value
    /// outside its defined enumeration. The property keeps its default (or
    /// previously applied) value, as Word does.
    EnumValue,
    /// A section Boolean8 operand was neither 0 nor 1. It reads as true, the
    /// way Word and LibreOffice test the byte against zero.
    Boolean,
}

impl SectionDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedBits => "ignored the reserved page-border bits",
            Self::BorderType => "read the undefined page-border type as no border",
            Self::BorderColor => "read the undefined page-border color index as automatic color",
            Self::EnumValue => "kept the property's default for the undefined section value",
            Self::Boolean => "read the non-0/1 Boolean8 section operand as true",
        }
    }
}

/// The class of piece-table (`Pcd`) defect the reader repairs the way Word
/// repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// per-piece repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. The CLX framing itself stays structural:
/// a malformed `RgPrc`/`Pcdt`/`PlcPcd` still yields no piece table. The
/// recorded [`ToleratedDefect::context`] is the piece's index in `PlcPcd`
/// (saturating at `u16::MAX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PieceDefect {
    /// A piece's `Pcd.prm` could not be resolved: a `Prm1` index beyond the
    /// `RgPrc` array or a `Prm0.isprm` MS-DOC 2.9.215 does not list.
    ///
    /// The modifier applies to its own piece only, so that piece reads with
    /// no modifier — its text and FKP formatting are kept, as in Word — and
    /// every other piece is unaffected. Before this repair one such piece
    /// silently discarded the whole piece table and with it all character
    /// and paragraph formatting.
    PropertyModifier,
}

impl PieceDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::PropertyModifier => {
                "ignored the unresolvable piece property modifier; the piece keeps its FKP formatting"
            },
        }
    }
}

/// One class of defect repaired while a document was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a document. [`Defect::Stylesheet`] classes require
/// [`Leniency::TolerateStylesheetDefects`]; every other class is repaired
/// under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A stylesheet defect repaired by an opt-in [`Leniency`] variant.
    Stylesheet(StylesheetDefect),
    /// A table SPRM defect repaired the way Word repairs it.
    Sprm(SprmDefect),
    /// A stylesheet-header defect repaired the way Word repairs it.
    Stshif(StshifDefect),
    /// An optional FIB-pointed auxiliary table whose parse failure was
    /// degraded to the table's absent value.
    AuxiliaryTable(AuxiliaryTableDefect),
    /// A style-record (`STD`) defect repaired the way Word repairs it.
    Style(StyleDefect),
    /// A comment-descriptor (`ATRDPre10`) defect repaired the way Word
    /// repairs it.
    Comment(CommentDefect),
    /// A standard-bookmark defect repaired the way Word repairs it.
    Bookmark(BookmarkDefect),
    /// A table-definition (`sprmTDefTable`) defect repaired the way Word
    /// repairs it.
    TableDefinition(TableDefinitionDefect),
    /// A paragraph-property (`PapxFkp`) defect repaired the way Word
    /// repairs it.
    Papx(PapxDefect),
    /// A section-property (`SEPX`) defect repaired the way Word repairs it.
    Section(SectionDefect),
    /// A piece-table (`Pcd`) defect repaired the way Word repairs it.
    Piece(PieceDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Stylesheet(defect) => defect.repair(),
            Self::Sprm(defect) => defect.repair(),
            Self::Stshif(defect) => defect.repair(),
            Self::AuxiliaryTable(defect) => defect.repair(),
            Self::Style(defect) => defect.repair(),
            Self::Comment(defect) => defect.repair(),
            Self::Bookmark(defect) => defect.repair(),
            Self::TableDefinition(defect) => defect.repair(),
            Self::Papx(defect) => defect.repair(),
            Self::Section(defect) => defect.repair(),
            Self::Piece(defect) => defect.repair(),
        }
    }
}

impl From<StylesheetDefect> for Defect {
    fn from(defect: StylesheetDefect) -> Self {
        Self::Stylesheet(defect)
    }
}

impl From<SprmDefect> for Defect {
    fn from(defect: SprmDefect) -> Self {
        Self::Sprm(defect)
    }
}

impl From<StshifDefect> for Defect {
    fn from(defect: StshifDefect) -> Self {
        Self::Stshif(defect)
    }
}

impl From<AuxiliaryTableDefect> for Defect {
    fn from(defect: AuxiliaryTableDefect) -> Self {
        Self::AuxiliaryTable(defect)
    }
}

impl From<StyleDefect> for Defect {
    fn from(defect: StyleDefect) -> Self {
        Self::Style(defect)
    }
}

impl From<CommentDefect> for Defect {
    fn from(defect: CommentDefect) -> Self {
        Self::Comment(defect)
    }
}

impl From<BookmarkDefect> for Defect {
    fn from(defect: BookmarkDefect) -> Self {
        Self::Bookmark(defect)
    }
}

impl From<TableDefinitionDefect> for Defect {
    fn from(defect: TableDefinitionDefect) -> Self {
        Self::TableDefinition(defect)
    }
}

impl From<PapxDefect> for Defect {
    fn from(defect: PapxDefect) -> Self {
        Self::Papx(defect)
    }
}

impl From<SectionDefect> for Defect {
    fn from(defect: SectionDefect) -> Self {
        Self::Section(defect)
    }
}

impl From<PieceDefect> for Defect {
    fn from(defect: PieceDefect) -> Self {
        Self::Piece(defect)
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found: the style index within the stylesheet for
    /// a [`Defect::Stylesheet`], the SPRM opcode for a [`Defect::Sprm`], the
    /// offending 16-bit Stshif word for a [`Defect::Stshif`], and the FIB
    /// `rgFcLcb` pointer index for a [`Defect::AuxiliaryTable`] (the style
    /// index for [`AuxiliaryTableDefect::RevisionAuthorResolution`]), and
    /// the style index (`istd`) for a [`Defect::Style`], and the comment
    /// index in `PlcfandRef` for a [`Defect::Comment`], and the bookmark
    /// index in `PlcfBkf` for a [`Defect::Bookmark`], and the cell (or
    /// boundary) index within the row for a [`Defect::TableDefinition`], and
    /// the `PlcfBtePapx` page index for a [`Defect::Papx`], and the SPRM
    /// opcode for a [`Defect::Section`], and the piece index in `PlcPcd` for
    /// a [`Defect::Piece`].
    pub context: u16,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the document conformed and no substitution was made.
/// After a [`Leniency::Strict`] read the report can still list every
/// Word-compatible repair — every class other than [`Defect::Stylesheet`],
/// which Word itself applies; it never lists [`Defect::Stylesheet`] repairs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the document parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u16) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                context,
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            self.record(tolerated.defect, tolerated.context);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
        assert!(!Leniency::default().tolerates_stylesheet_defects());
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(StylesheetDefect::DuplicateStyleName, 15);
        report.record(SprmDefect::CellBorderSideFlags, 0xD620);
        report.record(StylesheetDefect::DuplicateStyleName, 42);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Stylesheet(StylesheetDefect::DuplicateStyleName), 15),
                (Defect::Sprm(SprmDefect::CellBorderSideFlags), 0xD620),
                (Defect::Stylesheet(StylesheetDefect::DuplicateStyleName), 42),
            ]
        );
    }

    /// A hostile stylesheet must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(StylesheetDefect::DuplicateStyleName, index as u16);
        }
        for _ in 0..3 {
            report.record(SprmDefect::CellBorderSideFlags, 0xD62F);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(StylesheetDefect::DuplicateStyleName, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(SprmDefect::CellBorderSideFlags, 0xD620);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(SprmDefect::CellBorderSideFlags, 0xD62F);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Stylesheet(StylesheetDefect::DuplicateStyleName),
                context: 7,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(StylesheetDefect::DuplicateStyleName),
            Defect::from(SprmDefect::CellBorderSideFlags),
            Defect::from(SprmDefect::BorderType),
            Defect::from(SprmDefect::BorderColor),
            Defect::from(StshifDefect::StyleNameFlags),
            Defect::from(StshifDefect::FixedStyleCount),
            Defect::from(AuxiliaryTableDefect::Fields),
            Defect::from(AuxiliaryTableDefect::Comments),
            Defect::from(AuxiliaryTableDefect::Bookmarks),
            Defect::from(AuxiliaryTableDefect::SmartTags),
            Defect::from(AuxiliaryTableDefect::Rsids),
            Defect::from(AuxiliaryTableDefect::RmdThreading),
            Defect::from(AuxiliaryTableDefect::EmbeddedFonts),
            Defect::from(AuxiliaryTableDefect::Envelope),
            Defect::from(AuxiliaryTableDefect::AutoSummary),
            Defect::from(AuxiliaryTableDefect::ProtectedRanges),
            Defect::from(AuxiliaryTableDefect::FormatConsistency),
            Defect::from(AuxiliaryTableDefect::StructuredTags),
            Defect::from(AuxiliaryTableDefect::XmlSchemas),
            Defect::from(AuxiliaryTableDefect::CustomXmlTransform),
            Defect::from(AuxiliaryTableDefect::OleControls),
            Defect::from(AuxiliaryTableDefect::MailMerge),
            Defect::from(AuxiliaryTableDefect::Subdocuments),
            Defect::from(AuxiliaryTableDefect::RevisionAuthors),
            Defect::from(AuxiliaryTableDefect::AssociatedStrings),
            Defect::from(AuxiliaryTableDefect::ListNames),
            Defect::from(AuxiliaryTableDefect::ListTemplates),
            Defect::from(AuxiliaryTableDefect::Sections),
            Defect::from(AuxiliaryTableDefect::RevisionAuthorResolution),
            Defect::from(StyleDefect::ReservedStyleFlags),
            Defect::from(StyleDefect::ReservedPost2000Flags),
            Defect::from(StyleDefect::Priority),
            Defect::from(CommentDefect::ReservedFields),
            Defect::from(CommentDefect::AuthorIndex),
            Defect::from(CommentDefect::BookmarkTag),
            Defect::from(BookmarkDefect::ReservedPublicFlag),
            Defect::from(BookmarkDefect::EndIndex),
            Defect::from(BookmarkDefect::ReversedRange),
            Defect::from(BookmarkDefect::ColumnRange),
            Defect::from(TableDefinitionDefect::BoundaryRange),
            Defect::from(TableDefinitionDefect::TextFlow),
            Defect::from(TableDefinitionDefect::VerticalMerge),
            Defect::from(TableDefinitionDefect::VerticalAlignment),
            Defect::from(TableDefinitionDefect::WidthUnit),
            Defect::from(TableDefinitionDefect::NegativeWidth),
            Defect::from(PapxDefect::Page),
            Defect::from(PapxDefect::Run),
            Defect::from(PapxDefect::AlignmentPadding),
            Defect::from(SectionDefect::ReservedBits),
            Defect::from(SectionDefect::BorderType),
            Defect::from(SectionDefect::BorderColor),
            Defect::from(SectionDefect::EnumValue),
            Defect::from(SectionDefect::Boolean),
            Defect::from(PieceDefect::PropertyModifier),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Stylesheet(StylesheetDefect::DuplicateStyleName).repair(),
            StylesheetDefect::DuplicateStyleName.repair()
        );
        assert_eq!(
            Defect::Sprm(SprmDefect::CellBorderSideFlags).repair(),
            SprmDefect::CellBorderSideFlags.repair()
        );
        assert_eq!(
            Defect::Stshif(StshifDefect::StyleNameFlags).repair(),
            StshifDefect::StyleNameFlags.repair()
        );
        assert_eq!(
            Defect::AuxiliaryTable(AuxiliaryTableDefect::Sections).repair(),
            AuxiliaryTableDefect::Sections.repair()
        );
    }

    #[test]
    fn every_auxiliary_table_defect_locates_a_fib_pointer() {
        let indexes = [
            (AuxiliaryTableDefect::Fields, 16),
            (AuxiliaryTableDefect::Comments, 4),
            (AuxiliaryTableDefect::Bookmarks, 21),
            (AuxiliaryTableDefect::SmartTags, 114),
            (AuxiliaryTableDefect::Rsids, 113),
            (AuxiliaryTableDefect::RmdThreading, 94),
            (AuxiliaryTableDefect::EmbeddedFonts, 61),
            (AuxiliaryTableDefect::Envelope, 97),
            (AuxiliaryTableDefect::AutoSummary, 89),
            (AuxiliaryTableDefect::ProtectedRanges, 141),
            (AuxiliaryTableDefect::FormatConsistency, 120),
            (AuxiliaryTableDefect::StructuredTags, 137),
            (AuxiliaryTableDefect::XmlSchemas, 136),
            (AuxiliaryTableDefect::CustomXmlTransform, 140),
            (AuxiliaryTableDefect::OleControls, 85),
            (AuxiliaryTableDefect::MailMerge, 44),
            (AuxiliaryTableDefect::Subdocuments, 54),
            (AuxiliaryTableDefect::RevisionAuthors, 51),
            (AuxiliaryTableDefect::AssociatedStrings, 32),
            (AuxiliaryTableDefect::ListNames, 91),
            (AuxiliaryTableDefect::ListTemplates, 96),
            (AuxiliaryTableDefect::Sections, 6),
        ];
        for (defect, index) in indexes {
            assert_eq!(defect.fib_index(), index, "{defect:?}");
        }
    }
}
