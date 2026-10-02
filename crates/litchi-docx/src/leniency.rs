//! Repair reporting for non-structural defects in WordprocessingML documents.
//!
//! Two kinds of repair can happen while a document is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: OPC packaging, XML well-formedness, part
//! grammar, and decryption remain hard errors under every variant.
//!
//! **Word-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover markup the schema
//! constrains but real producers emit anyway — markup Word and LibreOffice
//! read past without complaint, so rejecting it would cost the caller the
//! whole part over bytes that cannot change what a conforming file means.
//! Each substitutes the behavior documented on the corresponding defect
//! variant: [`TextDefect`] (semantic text extraction), [`SettingsDefect`]
//! (`settings.xml`), [`NumberingDefect`] (`numbering.xml` and list
//! resolution), [`StyleDefect`] (`styles.xml`), [`RevisionDefect`]
//! (tracked-change metadata),
//! [`SectionDefect`] (section column models), and [`ContentControlDefect`]
//! (`w:sdtPr` properties).
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. There are no opt-in repairs yet: [`Leniency::Strict`]
//! is the only variant, so every read applies the same contract. Judgment-call
//! repairs land behind explicit variants as the roadmap findings that need
//! them are implemented.
//!
//! This mirrors the contract `litchi_doc::leniency` provides for legacy
//! binary documents so the two Word readers behave the same way.
//!
//! # Carrier status
//!
//! [`Package::document`](crate::Package::document) aggregates the report and
//! [`Document::tolerance_report`](crate::Document::tolerance_report) exposes
//! it. Because the typed codecs run lazily under the read facades, each
//! accessor call parses into its own report and merges it into the
//! document-level aggregation with `ToleranceReport::append` — the idiom
//! `litchi-doc` established at open time — so the report accumulates across
//! calls in call order. The aggregation sits behind a mutex so the facades
//! keep their `&self` receivers; `tolerance_report` returns a snapshot.
//!
//! Recording sites, by finding:
//!
//! - 2.1 — [`Document::write_text_to`](crate::Document::write_text_to) (and
//!   the source-backed package sink) records [`TextDefect`].
//! - 2.2, 2.3 — [`Document::settings`](crate::Document::settings) records
//!   [`SettingsDefect`]; the parsed
//!   [`DocumentSettings`](crate::settings::DocumentSettings) also carries its
//!   own report for the package-level settings shortcuts.
//! - 2.4, 2.6 — [`Document::numbering`](crate::Document::numbering),
//!   [`Document::numbering_snapshot`](crate::Document::numbering_snapshot),
//!   and the list resolvers ([`Document::list_items`](crate::Document::list_items)
//!   and friends) record [`NumberingDefect`]; the
//!   [`numbering::Collection`](crate::numbering::Collection) carries its own.
//! - 2.9 — the list resolvers record [`NumberingDefect::LevelOutOfRange`]
//!   for a paragraph `w:ilvl` above 8 and merge the [`StyleDefect`]s of the
//!   [`Styles`](crate::styles::Styles) they load, which carries its own.
//! - 2.5 — [`Document::revisions`](crate::Document::revisions) and the pinned
//!   source-backed `Document::revisions` record [`RevisionDefect`].
//! - 2.7 — [`Document::section_inventory`](crate::Document::section_inventory)
//!   (and the source-backed inventories) record [`SectionDefect`];
//!   [`Section`](crate::Section) and the inventory carry their own.
//! - 2.8 — [`Document::content_controls`](crate::Document::content_controls)
//!   records [`ContentControlDefect`]; the
//!   [`content_control::Inventory`](crate::content_control::Inventory) carries
//!   its own.
//!
//! - Packaging and metadata — the OPC repairs `litchi-opc` makes while the
//!   package is opened ([`litchi_opc::leniency`]) are folded in as
//!   [`Defect::Package`] by every open path: [`Package::document`](crate::Package::document)
//!   (and [`Package::tolerance_report`](crate::Package::tolerance_report)),
//!   the source-backed [`Package`](crate::source_backed::Package), and its
//!   pinned [`Document`](crate::source_backed::Document). Core- and
//!   custom-properties repairs ([`litchi_ooxml_common::leniency`]) are folded
//!   in as [`Defect::Properties`] when the metadata is read: at open for the
//!   owning package (which loads it eagerly), by
//!   `source_backed::Package::metadata` for the source-backed one. These
//!   entries carry a [`PartLocation`] instead of a byte offset.
//!
//! Fragment-level accessors without a document handle (a standalone
//! [`Paragraph`](crate::Paragraph)'s numbering or revisions, table revisions)
//! apply the same repairs; the document-level accessor over the same part
//! enumerates them.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous part must not be able to make the report grow
/// without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// Upper bound on the heap a single [`ToleranceReport`] can retain, for
/// workspace estimates that must account for every owner a parse can grow.
///
/// Codec reports never carry a [`PartLocation`]; only the open-time packaging
/// and metadata fold does, whose strings are bounded by the package read
/// limits that admitted them.
pub(crate) const MAX_REPORT_HEAP_BYTES: usize = MAX_RECORDED_DEFECTS * size_of::<ToleratedDefect>();

/// How the reader treats non-structural document defects.
///
/// Structural defects — malformed OPC packaging, ill-formed XML, broken part
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant. The Word-compatible repairs listed by the [`Defect`]
/// classes are applied and recorded under every variant, including
/// [`Leniency::Strict`].
///
/// No opt-in repairs exist yet, so `Strict` is currently the only variant and
/// the default: every read applies the same contract. Variants gate
/// judgment-call repairs as the roadmap findings that need them land.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: there are no opt-in repairs to reject, while the
    /// Word-compatible [`Defect`] repairs are still applied and recorded in
    /// the [`ToleranceReport`].
    #[default]
    Strict,
}

/// The class of semantic text-extraction defect the reader repairs the way
/// Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// text-extraction repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. Each one matches what Word and LibreOffice
/// do with markup no reader-rejection rule covers, so applying it never
/// changes what a conforming document means.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextDefect {
    /// A text-bearing subtree bound to a namespace other than WordprocessingML
    /// appeared where the semantic text codec walks the paragraph: an OMML
    /// `m:t` inside `m:oMath` (ECMA-376 Part 1 §22), or a DrawingML `a:t` in
    /// an inline shape.
    ///
    /// No spec text sanctions refusing the extraction over this: Word,
    /// LibreOffice, and POI all extract the surrounding text, and the crate's
    /// own `extract_word_text` scanner already treats foreign subtrees as
    /// opaque. The typed codec does the same — the subtree contributes no
    /// text and the walk resumes after it — instead of failing the whole
    /// document. The semantic text sink records one defect per foreign `*:t`
    /// element, with the byte offset of the element's start tag within the
    /// visible part XML as the context.
    SkippedForeignMarkup,
}

impl TextDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::SkippedForeignMarkup => {
                "skipped the foreign subtree as opaque; surrounding WordprocessingML text was still extracted"
            },
        }
    }
}

/// The class of `settings.xml` defect the reader repairs the way Word and
/// LibreOffice repair it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// settings repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each is recorded with the byte offset of the
/// offending element's start tag within the visible (markup-compatibility
/// processed) settings XML, and is carried by the parsed
/// [`DocumentSettings`](crate::settings::DocumentSettings) itself
/// ([`DocumentSettings::tolerance_report`](crate::settings::DocumentSettings::tolerance_report)),
/// so every settings accessor reports it; [`Document::settings`](crate::Document::settings)
/// also merges it into the document-level report.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SettingsDefect {
    /// A `w:compat` child is neither `w:compatSetting` nor one of the closed
    /// ECMA-376 `CT_Compat` flag names — typically a flag introduced by a newer
    /// Word release or a vendor such as WPS.
    ///
    /// LibreOffice dispatches compat flags by token and ignores unknown
    /// children (`SettingsTable.cxx`); Word preserves and opens them. The
    /// reader skips the element: it contributes no
    /// [`CompatibilityOption`](crate::settings::CompatibilityOption), and every
    /// other setting still parses.
    UnknownCompatFlag,
    /// A Transitional-only compat flag appeared in a Strict WordprocessingML
    /// settings part. Strict readers map both namespaces onto the same flag
    /// set, so the flag is kept with its declared value.
    TransitionalCompatFlagInStrict,
    /// A setting the schema allows once (`maxOccurs="1"`) appeared again:
    /// a singleton element such as `w:view`, `w:defaultTabStop`, or
    /// `w:documentProtection`; a second `w:compat`, `w:footnotePr`, or
    /// `w:endnotePr` group; or a repeated flag or note property inside one.
    ///
    /// LibreOffice's settings importer assigns each occurrence in document
    /// order, so the later value wins; WPS, merge tools, and hand edits emit
    /// such duplicates. The reader does the same — a repeated singleton
    /// replaces the earlier element, and a repeated group's children apply on
    /// top of the first group's — and records one defect per repeated
    /// element.
    DuplicateSetting,
}

impl SettingsDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnknownCompatFlag => {
                "skipped the unknown compatibility flag; all other settings were still read"
            },
            Self::TransitionalCompatFlagInStrict => {
                "kept the Transitional-only compatibility flag found in a Strict settings part"
            },
            Self::DuplicateSetting => {
                "applied the repeated setting over the earlier occurrence (later value wins)"
            },
        }
    }
}

/// The class of `numbering.xml` defect the reader repairs per item, the way
/// Word and LibreOffice degrade it.
///
/// This enumeration is closed by design. Each repair degrades only the
/// offending property, level, override, or definition — never the whole
/// numbering model — so one bad value cannot take down
/// [`Document::list_items`](crate::Document::list_items) for unrelated lists.
/// Each is recorded with the byte offset of the offending element's start
/// tag within `numbering.xml` and is carried by the decoded
/// [`numbering::Collection`](crate::numbering::Collection)
/// ([`tolerance_report`](crate::numbering::Collection::tolerance_report));
/// [`Document::numbering`](crate::Document::numbering) and
/// [`Document::numbering_snapshot`](crate::Document::numbering_snapshot)
/// also merge it into the document-level report. Ill-formed XML, the
/// resource limits, required identity attributes (`w:abstractNumId`,
/// `w:numId`, `w:ilvl`), and nesting violations stay hard errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum NumberingDefect {
    /// A level's `w:numFmt` value is missing or outside `ST_NumberFormat`
    /// (for example `hangulChosung`). Word substitutes decimal numbering and
    /// LibreOffice maps unknown tokens to a fallback; the level uses
    /// [`Format::Decimal`](crate::numbering::Format::Decimal), keeping any
    /// `w:format` custom pattern.
    UnknownNumberFormat,
    /// A level's `w:suff` value is missing or outside `ST_LevelSuffix`; the
    /// schema default, a tab, is used.
    UnknownLevelSuffix,
    /// A level's `w:lvlRestart` value is not an integer in `0..=9`; the
    /// property is ignored and the default restart behavior applies.
    InvalidLevelRestart,
    /// An `ST_OnOff` value (such as `w:isLgl`) is outside the lexical domain.
    /// Following LibreOffice, `True` and `On` read as true and anything else
    /// (for example `yes`) as false.
    InvalidOnOff,
    /// An abstract definition's `w:multiLevelType` value is missing or
    /// outside `ST_MultiLevelType`; the definition is kept without a type.
    UnknownMultiLevelType,
    /// A numeric or string property value (`w:start`, `w:startOverride`,
    /// `w:lvlPicBulletId`, `w:abstractNumId`, `w:pStyle`, `w:styleLink`,
    /// `w:numStyleLink`) is missing or unparseable; that property is ignored
    /// and its default kept.
    InvalidValue,
    /// An ID or a singleton appeared twice: a repeated `w:abstractNumId`,
    /// `w:numId`, or `w:numPicBulletId`; a repeated level within one
    /// definition; a repeated `w:lvlOverride`; or a repeated singleton
    /// property. LibreOffice resolves IDs by first match
    /// (`ListsManager::GetAbstractList`/`GetList`), so the first occurrence
    /// is kept and the later one dropped.
    DuplicateDefinition,
    /// A `w:num` instance has no `w:abstractNumId` or references an abstract
    /// definition that does not exist. It numbers nothing in Word or
    /// LibreOffice, so it is dropped; paragraphs referencing it are
    /// unnumbered.
    MissingAbstractDefinition,
    /// A replacement `w:lvl` inside `w:lvlOverride` names a different level
    /// index than its override; the replacement is dropped and the override
    /// keeps any `w:startOverride`.
    MismatchedOverrideLevel,
    /// A paragraph's numbering (direct, style-inherited, or through a
    /// numbering style link) names a `w:numId` with no instance in the
    /// numbering model — never declared, or dropped by one of the repairs
    /// above. Word and LibreOffice show the paragraph unnumbered, so list
    /// resolution skips it instead of failing.
    ///
    /// Recorded by [`Document::list_items`](crate::Document::list_items) and
    /// the other list resolvers; the context is the zero-based index of the
    /// paragraph in the main story, not a byte offset.
    MissingNumberingInstance,
    /// A paragraph's `w:numPr` has no `w:numId` — typically only a `w:ilvl`
    /// level override, as LibreOffice and RTF converters emit. `CT_NumPr`
    /// makes both children optional; Word and LibreOffice take the
    /// numbering from the paragraph style and apply the `w:ilvl` override.
    /// [`Paragraph::numbering`](crate::Paragraph::numbering) reports no
    /// direct numbering, `Paragraph::style_id` is unaffected, and list
    /// resolution applies the override to the style-inherited numbering.
    ///
    /// Recorded by [`Document::list_items`](crate::Document::list_items) and
    /// the other list resolvers; the context is the zero-based index of the
    /// paragraph in the main story, not a byte offset.
    NumPrMissingNumId,
    /// A paragraph's `w:numPr/w:ilvl` exceeds 8. `ST_DecimalNumber` is
    /// unbounded, but Word models only list levels 0–8; the level is clamped
    /// to 8, the deepest level, so the paragraph stays numbered (LibreOffice,
    /// which models a tenth level, ignores values of 10 and above).
    /// [`Paragraph::numbering`](crate::Paragraph::numbering) reports the
    /// clamped level and `Paragraph::style_id` is unaffected.
    ///
    /// Recorded by [`Document::list_items`](crate::Document::list_items) and
    /// the other list resolvers; the context is the zero-based index of the
    /// paragraph in the main story, not a byte offset.
    LevelOutOfRange,
}

impl NumberingDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnknownNumberFormat => "used decimal numbering for the unknown number format",
            Self::UnknownLevelSuffix => "used a tab for the unknown level suffix",
            Self::InvalidLevelRestart => "ignored the out-of-range level restart",
            Self::InvalidOnOff => "read the invalid on/off value the way LibreOffice does",
            Self::UnknownMultiLevelType => "kept the definition without a multi-level type",
            Self::InvalidValue => "ignored the missing or unparseable property value",
            Self::DuplicateDefinition => "kept the first occurrence and dropped the duplicate",
            Self::MissingAbstractDefinition => {
                "dropped the numbering instance that has no abstract definition"
            },
            Self::MismatchedOverrideLevel => {
                "dropped the override level whose index does not match its override"
            },
            Self::MissingNumberingInstance => {
                "left the paragraph unnumbered because its numbering instance does not exist"
            },
            Self::NumPrMissingNumId => {
                "took the numbering from the paragraph style, applying the direct level override"
            },
            Self::LevelOutOfRange => "clamped the list level to 8, the deepest level Word models",
        }
    }
}

/// The class of `styles.xml` defect the reader repairs per property, the way
/// Word and LibreOffice read the style.
///
/// This enumeration is closed by design. Each repair affects one property of
/// one style; the rest of the style table loads normally. Each is recorded on
/// the loaded [`Styles`](crate::styles::Styles)
/// ([`tolerance_report`](crate::styles::Styles::tolerance_report)) with the
/// byte offset of the offending element's start tag within the visible
/// (markup-compatibility processed) styles XML, and the list resolvers
/// ([`Document::list_items`](crate::Document::list_items) and friends) merge
/// it into the document-level report. Ill-formed XML, a non-integer level,
/// and repeated numbering or outline properties stay hard errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StyleDefect {
    /// A style's `w:numPr/w:ilvl` exceeds 8 (`ST_DecimalNumber` is
    /// unbounded; Word models levels 0–8). The level is clamped to 8.
    LevelOutOfRange,
    /// A style's `w:outlineLvl` exceeds 9 (0–8 are heading levels, 9 is body
    /// text). Word ignores such a value, so the style has no outline level.
    OutlineLevelOutOfRange,
}

impl StyleDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::LevelOutOfRange => "clamped the style's list level to 8",
            Self::OutlineLevelOutOfRange => "ignored the out-of-range outline level",
        }
    }
}

/// The class of tracked-revision defect the reader repairs the way Word
/// repairs it.
///
/// This enumeration is closed by design. Each repair keeps the revision in
/// the listing with the offending metadata absent, instead of failing
/// [`Document::revisions`](crate::Document::revisions) as one batch. Each is
/// recorded with the byte offset of the revision element's start tag within
/// the main document XML, merged into
/// [`Document::tolerance_report`](crate::Document::tolerance_report) (and the
/// pinned source-backed [`Document`](crate::source_backed::Document)'s
/// report). Fragment-level accessors (paragraph and table revisions) apply
/// the same repairs; the document-level listing enumerates them. A missing
/// `w:id`, duplicate attributes, ill-formed XML, and the revision resource
/// limits stay hard errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RevisionDefect {
    /// A tracked change has no `w:author`. ECMA-376 `CT_TrackChange` declares
    /// the attribute optional; privacy-scrubbed documents and third-party
    /// producers omit it, and Word shows such a change with an unknown
    /// author. The revision is kept with
    /// [`author`](crate::revision::Revision::author) `None`.
    MissingAuthor,
    /// A `w:date` is not a valid `xsd:dateTime`; the revision is kept undated.
    InvalidDate,
    /// A `w16du:dateUtc` is not a bounded UTC `xsd:dateTime`; the revision is
    /// kept without it.
    InvalidDateUtc,
    /// A `w:id` is not an `xsd:integer`. The ID only keys the annotation, so
    /// the authored token is kept as the revision's ID.
    NonIntegerId,
}

impl RevisionDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::MissingAuthor => "kept the revision without an author",
            Self::InvalidDate => "kept the revision without its unreadable date",
            Self::InvalidDateUtc => "kept the revision without its unreadable UTC date",
            Self::NonIntegerId => "kept the non-integer revision ID as authored",
        }
    }
}

/// The class of section-layout (`w:sectPr`) defect the reader repairs the
/// way Word and LibreOffice lay it out.
///
/// This enumeration is closed by design. Each repair degrades only the
/// column model of one section instead of failing that section's semantic
/// decode — and with it [`Section::columns`](crate::Section::columns), the
/// section inventory, and the layout snapshot. Each is recorded on the
/// decoded section ([`Section::tolerance_report`](crate::Section::tolerance_report),
/// context: byte offset within that section's `w:sectPr` XML) and on the
/// section inventory ([`section::Inventory::tolerance_report`](crate::section::Inventory::tolerance_report),
/// context: byte offset within the visible main-document XML), which
/// [`Document::section_inventory`](crate::Document::section_inventory) merges
/// into the document-level report. Out-of-domain measurements and the
/// section resource limits stay hard errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionDefect {
    /// A `w:col` has no `w:w`. ECMA-376 `CT_Col` declares the width
    /// optional; LibreOffice's `SectionColumnHandler` starts each column at
    /// width 0 and keeps it, and Word computes the layout. The column is kept
    /// with width 0.
    ColumnWidthMissing,
    /// `w:cols/@w:num` is not an integer in `1..=45` (`ST_DecimalNumber` has
    /// no bound, but Word lays out 1 to 45 columns). The count is clamped to
    /// that range, or 1 when it is not a number.
    ColumnCountOutOfRange,
    /// Unequal-width columns whose `w:col` list does not match `w:num`. The
    /// explicit columns define the layout (their count replaces `w:num`) when
    /// there are 1 to 45 of them; otherwise the section uses equal-width
    /// columns.
    ColumnCountMismatch,
}

impl SectionDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ColumnWidthMissing => "kept the column with width 0",
            Self::ColumnCountOutOfRange => "clamped the column count to 1..=45",
            Self::ColumnCountMismatch => {
                "took the column count from the explicit columns (or used equal widths)"
            },
        }
    }
}

/// The class of content-control (`w:sdtPr`) defect the reader repairs per
/// control, the way Word and LibreOffice render such a control.
///
/// This enumeration is closed by design. Each repair affects one property of
/// one control; the control keeps its range, its content, and every other
/// property, and the rest of the inventory is read normally — one malformed
/// `w:sdtPr` never aborts [`Document::content_controls`](crate::Document::content_controls).
/// Each is recorded on the [`content_control::Inventory`](crate::content_control::Inventory)
/// ([`tolerance_report`](crate::content_control::Inventory::tolerance_report))
/// with the byte offset of the offending `w:sdtPr` child within the visible
/// (markup-compatibility processed) document XML, and
/// `Document::content_controls` merges it into the document-level report.
/// Nested `w:sdtPr`, resource limits, markup-compatibility violations, and
/// ill-formed XML stay hard errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContentControlDefect {
    /// A `w:lock` has no `w:val` or a value outside `ST_Lock`. `CT_Lock`
    /// declares the value optional; the control is read as unlocked.
    UnknownLockValue,
    /// A property failed validation: a missing or out-of-domain required
    /// value, a repeated singleton property, a second control type, or a
    /// data binding without `w:xpath`/`w:storeItemID` (both optional in
    /// `CT_DataBinding`). The property is dropped — the first occurrence of a
    /// repeated property wins — and the control is otherwise kept.
    InvalidProperty,
}

impl ContentControlDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnknownLockValue => "read the content control as unlocked",
            Self::InvalidProperty => "dropped the malformed property and kept the content control",
        }
    }
}

/// One class of defect repaired while a document was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a document. Every class listed here is repaired
/// under every [`Leniency`] variant; opt-in classes join the union together
/// with the [`Leniency`] variants that gate them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A semantic text-extraction defect repaired the way Word repairs it.
    Text(TextDefect),
    /// A `settings.xml` defect repaired the way Word and LibreOffice repair it.
    Settings(SettingsDefect),
    /// A `numbering.xml` defect degraded per item.
    Numbering(NumberingDefect),
    /// A tracked-revision metadata defect repaired the way Word repairs it.
    Revision(RevisionDefect),
    /// A `styles.xml` property defect repaired per style.
    Style(StyleDefect),
    /// A section column-model defect repaired the way Word lays it out.
    Section(SectionDefect),
    /// A content-control property defect repaired per control.
    ContentControl(ContentControlDefect),
    /// An OPC packaging defect (relationship or content-types manifest)
    /// repaired by `litchi-opc` while the package was opened; see
    /// [`litchi_opc::leniency`]. Folded into the document report at open, with
    /// its [`ToleratedDefect::location`] set.
    Package(litchi_opc::leniency::Defect),
    /// A core- or custom-properties defect repaired by `litchi-ooxml-common`
    /// while the document metadata was read; see
    /// [`litchi_ooxml_common::leniency`]. Folded in when the metadata is read,
    /// with its [`ToleratedDefect::location`] set.
    Properties(litchi_ooxml_common::leniency::Defect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Text(defect) => defect.repair(),
            Self::Settings(defect) => defect.repair(),
            Self::Numbering(defect) => defect.repair(),
            Self::Revision(defect) => defect.repair(),
            Self::Style(defect) => defect.repair(),
            Self::Section(defect) => defect.repair(),
            Self::ContentControl(defect) => defect.repair(),
            Self::Package(defect) => defect.repair(),
            Self::Properties(defect) => defect.repair(),
        }
    }
}

impl From<TextDefect> for Defect {
    fn from(defect: TextDefect) -> Self {
        Self::Text(defect)
    }
}

impl From<SettingsDefect> for Defect {
    fn from(defect: SettingsDefect) -> Self {
        Self::Settings(defect)
    }
}

impl From<NumberingDefect> for Defect {
    fn from(defect: NumberingDefect) -> Self {
        Self::Numbering(defect)
    }
}

impl From<RevisionDefect> for Defect {
    fn from(defect: RevisionDefect) -> Self {
        Self::Revision(defect)
    }
}

impl From<StyleDefect> for Defect {
    fn from(defect: StyleDefect) -> Self {
        Self::Style(defect)
    }
}

impl From<SectionDefect> for Defect {
    fn from(defect: SectionDefect) -> Self {
        Self::Section(defect)
    }
}

impl From<ContentControlDefect> for Defect {
    fn from(defect: ContentControlDefect) -> Self {
        Self::ContentControl(defect)
    }
}

impl From<litchi_opc::leniency::Defect> for Defect {
    fn from(defect: litchi_opc::leniency::Defect) -> Self {
        Self::Package(defect)
    }
}

impl From<litchi_ooxml_common::leniency::Defect> for Defect {
    fn from(defect: litchi_ooxml_common::leniency::Defect) -> Self {
        Self::Properties(defect)
    }
}

/// Where a packaging or metadata defect ([`Defect::Package`],
/// [`Defect::Properties`]) was found, as the originating crate's report
/// located it.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PartLocation {
    part: Box<str>,
    item: Option<Box<str>>,
    value: Option<Box<str>>,
}

impl PartLocation {
    fn new(part: &str, item: Option<&str>, value: Option<&str>) -> Self {
        Self {
            part: part.into(),
            item: item.map(Into::into),
            value: value.map(Into::into),
        }
    }

    /// The part name the defect belongs to (for relationship defects, the
    /// relationship source; `/` for package relationships).
    #[inline]
    #[must_use]
    pub fn part(&self) -> &str {
        &self.part
    }

    /// The item within the part — a relationship `Id`, a content-type key,
    /// or a property name — when one is known.
    #[inline]
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// The raw value the repair replaced or discarded, as documented on the
    /// originating defect variant.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found. Unless the defect variant documents a
    /// different locator, this is the byte offset of the offending element's
    /// start tag within the part's visible XML payload. Packaging and
    /// metadata defects carry `0` here and are located by
    /// [`location`](Self::location) instead.
    pub context: u64,
    /// The part/item/value locator of a [`Defect::Package`] or
    /// [`Defect::Properties`] defect; `None` for the WordprocessingML classes,
    /// which [`context`](Self::context) locates.
    pub location: Option<Box<PartLocation>>,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means every read so far conformed and no substitution was
/// made. A [`Leniency::Strict`] report lists only the Word-compatible repairs
/// Word itself applies. The semantic text sink merges each call's repairs
/// into the report, so defects accumulate across calls in call order.
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
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u64) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            context,
            location: None,
        });
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Fold the packaging repairs `litchi-opc` recorded while the package was
    /// opened, keeping their order, locations, and unrecorded count.
    pub(crate) fn append_package(&mut self, other: &litchi_opc::leniency::ToleranceReport) {
        for tolerated in other {
            self.push(ToleratedDefect {
                defect: Defect::Package(tolerated.defect()),
                context: 0,
                location: Some(Box::new(PartLocation::new(
                    tolerated.part(),
                    tolerated.item(),
                    tolerated.value(),
                ))),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded());
    }

    /// Fold the core- or custom-properties repairs `litchi-ooxml-common`
    /// recorded while the metadata was read, keeping their order, locations,
    /// and unrecorded count.
    pub(crate) fn append_properties(
        &mut self,
        other: &litchi_ooxml_common::leniency::ToleranceReport,
    ) {
        for tolerated in other {
            self.push(ToleratedDefect {
                defect: Defect::Properties(tolerated.defect()),
                context: 0,
                location: Some(Box::new(PartLocation::new(
                    tolerated.part(),
                    tolerated.item(),
                    tolerated.value(),
                ))),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded());
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        self.append_mapped(other, |context| context);
    }

    /// Merge another report whose contexts are relative to a different
    /// origin, translating each context with `map`.
    pub(crate) fn append_mapped(&mut self, other: &ToleranceReport, map: impl Fn(u64) -> u64) {
        for tolerated in other.defects() {
            self.push(ToleratedDefect {
                defect: tolerated.defect,
                context: map(tolerated.context),
                location: tolerated.location.clone(),
            });
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
        report.record(TextDefect::SkippedForeignMarkup, 1024);
        report.record(TextDefect::SkippedForeignMarkup, 4096);
        report.record(TextDefect::SkippedForeignMarkup, 8192);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Text(TextDefect::SkippedForeignMarkup), 1024),
                (Defect::Text(TextDefect::SkippedForeignMarkup), 4096),
                (Defect::Text(TextDefect::SkippedForeignMarkup), 8192),
            ]
        );
    }

    /// A hostile part must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TextDefect::SkippedForeignMarkup, index as u64);
        }
        for _ in 0..3 {
            report.record(TextDefect::SkippedForeignMarkup, 0);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TextDefect::SkippedForeignMarkup, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TextDefect::SkippedForeignMarkup, 1024);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TextDefect::SkippedForeignMarkup, 4096);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Text(TextDefect::SkippedForeignMarkup),
                context: 7,
                location: None,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(TextDefect::SkippedForeignMarkup),
            Defect::from(SettingsDefect::UnknownCompatFlag),
            Defect::from(SettingsDefect::TransitionalCompatFlagInStrict),
            Defect::from(SettingsDefect::DuplicateSetting),
            Defect::from(NumberingDefect::UnknownNumberFormat),
            Defect::from(NumberingDefect::UnknownLevelSuffix),
            Defect::from(NumberingDefect::InvalidLevelRestart),
            Defect::from(NumberingDefect::InvalidOnOff),
            Defect::from(NumberingDefect::UnknownMultiLevelType),
            Defect::from(NumberingDefect::InvalidValue),
            Defect::from(NumberingDefect::DuplicateDefinition),
            Defect::from(NumberingDefect::MissingAbstractDefinition),
            Defect::from(NumberingDefect::MismatchedOverrideLevel),
            Defect::from(NumberingDefect::MissingNumberingInstance),
            Defect::from(NumberingDefect::NumPrMissingNumId),
            Defect::from(NumberingDefect::LevelOutOfRange),
            Defect::from(RevisionDefect::MissingAuthor),
            Defect::from(RevisionDefect::InvalidDate),
            Defect::from(RevisionDefect::InvalidDateUtc),
            Defect::from(RevisionDefect::NonIntegerId),
            Defect::from(StyleDefect::LevelOutOfRange),
            Defect::from(StyleDefect::OutlineLevelOutOfRange),
            Defect::from(SectionDefect::ColumnWidthMissing),
            Defect::from(SectionDefect::ColumnCountOutOfRange),
            Defect::from(SectionDefect::ColumnCountMismatch),
            Defect::from(ContentControlDefect::UnknownLockValue),
            Defect::from(ContentControlDefect::InvalidProperty),
            Defect::from(litchi_opc::leniency::Defect::from(
                litchi_opc::leniency::RelationshipDefect::DuplicateId,
            )),
            Defect::from(litchi_ooxml_common::leniency::Defect::from(
                litchi_ooxml_common::leniency::CustomPropertiesDefect::InvalidValue,
            )),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Text(TextDefect::SkippedForeignMarkup).repair(),
            TextDefect::SkippedForeignMarkup.repair()
        );
    }

    /// The carrier is wired through the open path: opening a conformant
    /// fixture yields a document whose report exists and is empty.
    #[test]
    fn conformant_fixture_opens_with_a_clean_report() {
        let fixture = include_bytes!("../../../test-data/ooxml/docx/Hyperlink.docx");
        let package =
            crate::Package::from_reader(std::io::Cursor::new(fixture.as_slice())).unwrap();
        let document = package.document().unwrap();

        let report = document.tolerance_report();
        assert!(report.is_clean());
        assert!(report.defects().is_empty());
        assert_eq!(report.unrecorded(), 0);
        assert_eq!(report.total(), 0);
    }
}
