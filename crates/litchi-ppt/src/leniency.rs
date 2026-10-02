//! Repair reporting for non-structural defects in legacy presentations.
//!
//! Two kinds of repair can happen while a presentation is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: record framing, the UserEdit/persist
//! chain grammar, container structure, and encryption remain hard errors
//! under every variant.
//!
//! **`PowerPoint`-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover bytes
//! MS-PPT leaves undefined or reserved but real producers emit anyway —
//! bytes `PowerPoint`, POI, and LibreOffice read past without complaint, so
//! rejecting them would cost the caller the entire presentation over
//! metadata that cannot change what a conforming file means. Each
//! substitutes the value documented on the corresponding [`Defect`] class
//! variant.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. No repair currently requires opting in; the first
//! candidate on the roadmap is breaking a cyclic `UserEdit` chain (finding
//! F5.9 in `docs/LENIENCY_AUDIT.md`), and the [`Leniency`] variant and
//! open-options plumbing for it land with that finding.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` and
//! `crates/litchi-xls/src/leniency.rs` provide for the other legacy readers
//! so all three behave the same way.
//!
//! # Roadmap state
//!
//! The following findings from `docs/LENIENCY_AUDIT.md` record into this
//! report today, all `PowerPoint`-compatible (default) repairs:
//!
//! - **F5.1** ([`TextPropDefect::StyleTextPropCoverage`]) — the read path
//!   parses `StyleTextPropAtom` with the clamping POI/LibreOffice posture
//!   (coverage clamped to the text length + 1, parsing stopped at the first
//!   unreadable run, trailing bytes dropped) and records each clamp. The
//!   exact-coverage parser
//!   ([`crate::text_prop::parse_style_text_prop_atom_strict`]) and the
//!   strict [`crate::StyleTextPropAtom`] codec stay in use by the writer and
//!   the text-edit/reparse-verification path, which must not rewrite bytes
//!   it cannot frame exactly.
//! - **F5.2** ([`TextPropDefect::ReservedBulletFlagBits`],
//!   [`TextPropDefect::ReservedWrapFlagBits`]) — reserved
//!   `BulletFlags`/`PFWrapFlags` bits are masked to the defined set instead
//!   of rejecting the paragraph; the defined bits are still applied.
//! - **F5.3** ([`TextPropDefect::UnknownTextAlignment`],
//!   [`TextPropDefect::UnknownFontAlignment`],
//!   [`TextPropDefect::UnknownTextDirection`],
//!   [`TextPropDefect::UnknownTabType`],
//!   [`TextPropDefect::IndentLevelOutOfRange`],
//!   [`TextPropDefect::InvalidColorIndex`]) — undefined text enumeration
//!   values leave the property unset (a tab stop keeps its position as a
//!   left tab), indent levels above 4 clamp to 4, and an undefined
//!   `ColorIndexStruct` index reads as an unset color, instead of rejecting
//!   the text box.
//! - **F5.11** ([`ShapeDefect`]) — shape conversion is isolated per shape:
//!   a shape whose content cannot be read is dropped and recorded
//!   ([`ShapeDefect::UnreadableShape`]) while its siblings are kept, a
//!   duplicated `TextRulerAtom` or round-trip shape atom keeps its first
//!   occurrence, and an undefined round-trip placeholder ID reads as absent.
//!   A shape whose host-record framing is broken still fails the slide.
//! - **F5.4** ([`SlideListDefect`]) — the slide directory degrades per
//!   entry instead of failing the open: a nonconforming `SlidePersistAtom`
//!   header is read by field offset (POI/LibreOffice), a truncated atom or a
//!   null `persistIdRef`/`slideId` drops that entry, a repeated `slideId` or
//!   `persistIdRef` keeps the first entry, and `cTexts` above 8 clamps to 8.
//!   These are open-time repairs.
//! - **F5.5** ([`SlideListDefect::StaleSlidePersistRef`]) — a slide whose
//!   persist ID the live persist directory does not map, or that maps
//!   outside the document stream or to a record other than a
//!   `SlideContainer`, is skipped as LibreOffice skips it; the presentation
//!   opens with the remaining slides. A `SlideContainer` whose own length
//!   overruns the stream stays fatal.
//! - **F5.6** ([`PersistDefect::MissingCurrentUser`],
//!   [`PersistDefect::UnusableCurrentEditOffset`]) — when the "Current User"
//!   stream is missing or its edit offset does not land on a well-formed
//!   `UserEditAtom`, the live `UserEditAtom` is recovered by scanning the
//!   document stream's top-level records for the last one, LibreOffice's
//!   algorithm. A recovered atom that names an encryption session is
//!   rejected rather than read as plaintext.
//! - **F5.7** ([`PersistDefect::PersistDirectoryRecordType`],
//!   [`PersistDefect::ZeroCountPersistRun`],
//!   [`PersistDefect::DuplicatePersistId`]) — persist-directory decoding on
//!   the read path ignores a directory pointer to another record type, skips
//!   zero-count entry headers, and keeps the first offset of a persist ID
//!   repeated within one directory, as LibreOffice does. The public
//!   [`crate::PersistPtrHolder::parse_with_report`] and
//!   [`crate::PersistMapping::build_from_records_with_report`] apply the
//!   same repairs per entry instead of dropping a whole holder. A truncated
//!   or misaligned directory stays an error.
//! - **F5.8** (username-length part,
//!   [`PersistDefect::CurrentUserNameLength`]) — a "Current User"
//!   `lenUserName` above 255, or one the stream truncates, reads as no
//!   username (POI's posture) instead of failing the open; the edit offset
//!   and header token are still read and validated.
//!
//! Record framing, the `UserEdit`/persist chain, container structure,
//! encryption, and the configured record/interaction resource limits are not
//! touched by any of these repairs and stay fatal.
//!
//! # Where the report lives
//!
//! Text-style repairs fire during *lazy* parsing, so the report is threaded
//! explicitly (`&mut ToleranceReport`, no interior mutability in the report
//! itself) through the text-box parse and attached to the lazily cached value
//! it describes:
//!
//! - [`crate::Slide::tolerance_report`] and
//!   [`crate::SpeakerNotes::tolerance_report`] return the repairs made while
//!   that slide's or notes page's shapes were materialized. The report is
//!   cached in the same `OnceCell` as the shapes, so it is complete no matter
//!   which accessor (`shapes`, `shapes_with_report`, `tolerance_report`)
//!   triggered the parse.
//! - `shapes_with_report` appends that cached report to a caller-owned one,
//!   on every call, so one report can be threaded across many slides.
//! - [`crate::Presentation::tolerance_report`] is the open-time aggregate:
//!   the repairs made while the `UserEdit`/persist chain was located and
//!   read ([`PersistDefect`]) and while the slide directory was built
//!   ([`SlideListDefect`]). The open path threads one `&mut ToleranceReport`
//!   through each open-time parse that repairs anything.
//! - [`crate::Presentation::collect_tolerance_report`] returns the open-time
//!   aggregate plus the repairs of every slide's and notes page's shapes —
//!   the one-call answer to "what did the reader change about this file?".
//! - Lower-level readers ([`crate::EscherTextboxWrapper::new_with_report`],
//!   [`crate::TextRunExtractor::extract_from_records_with_report`], and the
//!   `Presentation::*_with_report` stream walks) record into a caller-owned
//!   report; their report-less conveniences apply the same repairs without
//!   enumerating them.
//!
//! Every recorded defect carries a [`DefectLocation`]: the persist ID of the
//! slide or notes page and the `OfficeArt` shape ID of the text box, filled
//! in by the layer that knows each one, plus the per-class `context` value.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous style atom must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural presentation defects.
///
/// Structural defects — malformed record framing or container structure, a
/// broken `UserEdit`/persist-chain grammar, or unsupported/absent decryption
/// material — are hard errors under every variant. The
/// `PowerPoint`-compatible repairs listed by [`Defect`] are applied
/// and recorded under every variant, including [`Leniency::Strict`].
///
/// The enumeration is non-exhaustive because opt-in variants arrive with the
/// roadmap findings that need them (the first candidate is F5.9); new
/// variants are never a breaking change for callers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: the `PowerPoint`-compatible [`Defect`]
    /// repairs are still applied and recorded in the [`ToleranceReport`].
    #[default]
    Strict,
}

/// The class of text-property defect the reader repairs the way `PowerPoint`
/// repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// the text-property repairs a read applies under every [`Leniency`]
/// variant, including [`Leniency::Strict`]. Each one matches what
/// `PowerPoint`, POI, and LibreOffice do with bytes MS-PPT leaves undefined
/// or reserved, so applying it never changes what a conforming presentation
/// means.
///
/// Each variant implements one finding of `docs/LENIENCY_AUDIT.md` (F5.1,
/// F5.2, F5.3) and is recorded by the text-box read path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextPropDefect {
    /// A `StyleTextPropAtom` run's `charactersCovered` did not tile exactly
    /// the text length plus one, a run was truncated, or the atom had
    /// trailing bytes (finding F5.1).
    ///
    /// MS-PPT 2.9.40–2.9.41 require coverage equal to the character count
    /// plus one, but nothing prescribes rejecting the document when a
    /// producer miscounts, and miscounted coverage is one of the most common
    /// .ppt deviations. POI only logs and truncates
    /// (`StyleTextPropAtom.java`), and LibreOffice clamps (`svdfppt.cxx`),
    /// so the reader parses with the clamping
    /// [`crate::text_prop::parse_style_text_prop_atom`] posture on the read
    /// path: coverage is clamped to what remains, parsing stops at the first
    /// unreadable run, and trailing bytes are dropped. The strict parser
    /// ([`crate::text_prop::parse_style_text_prop_atom_strict`]) is retained
    /// for the writer and reparse-verification path.
    StyleTextPropCoverage,
    /// A `TextPFRun`'s two-byte `bulletFlags` carried bits outside the set
    /// MS-PPT 2.9.22 defines (finding F5.2).
    ///
    /// MS-PPT 2.9.22 defines bits 0x0001–0x0008 and states the remaining
    /// twelve bits "MUST be zero and MUST be ignored". `PowerPoint`, POI,
    /// and LibreOffice read only the defined bits while third-party
    /// producers routinely leak garbage into the high ones, so the reader
    /// masks `bulletFlags` down to 0x000F and still applies every defined
    /// bit that survives.
    ReservedBulletFlagBits,
    /// A `TextPFRun`'s two-byte `wrapFlags` carried bits outside the set
    /// MS-PPT 2.9.25 defines (finding F5.2).
    ///
    /// MS-PPT 2.9.25 defines bits 0x0001–0x0004 and states the remaining
    /// thirteen bits "MUST be zero and MUST be ignored". As with
    /// [`TextPropDefect::ReservedBulletFlagBits`], the reader masks
    /// `wrapFlags` down to 0x0007 and still applies every defined bit that
    /// survives.
    ReservedWrapFlagBits,
    /// A `TextPFRun`'s `textAlignment` was outside `TextAlignmentEnum`
    /// (MS-PPT 2.13, values 0–6); the alignment is left unset, so the
    /// inherited/default alignment applies (finding F5.3).
    ///
    /// Newer and non-Microsoft producers emit values PowerPoint does not
    /// define; PowerPoint opens such text with default formatting,
    /// LibreOffice branches only on values it knows (`svdfppt.cxx`), and POI
    /// leaves unknown values uninterpreted. The context is the raw value.
    UnknownTextAlignment,
    /// A `TextPFRun`'s `fontAlign` was outside `TextFontAlignmentEnum`
    /// (MS-PPT 2.13, values 0–3); the property is left unset (finding
    /// F5.3). The context is the raw value.
    UnknownFontAlignment,
    /// A `TextPFRun`'s `textDirection` was outside `TextDirectionEnum`
    /// (MS-PPT 2.13, values 0–1); the property is left unset (finding
    /// F5.3). The context is the raw value.
    UnknownTextDirection,
    /// A `TabStop`'s `type` was outside `TextTabTypeEnum` (MS-PPT 2.13,
    /// values 0–3); the tab stop keeps its position with left alignment, the
    /// default tab type (finding F5.3). The context is the raw value.
    UnknownTabType,
    /// A `TextPFRun`'s `indentLevel` exceeded the 0–4 range MS-PPT 2.2.13
    /// allows; it is clamped to 4 (finding F5.3). The context is the raw
    /// level.
    IndentLevelOutOfRange,
    /// A text or bullet `ColorIndexStruct` carried an index other than the
    /// scheme slots 0x00–0x07, RGB (0xFE), or undefined (0xFF) that MS-PPT
    /// defines; the color is treated as unset and the raw value is
    /// kept for round-trip (finding F5.3). The context is the raw 32-bit
    /// `ColorIndexStruct`.
    InvalidColorIndex,
}

impl TextPropDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::StyleTextPropCoverage => {
                "clamped StyleTextPropAtom run coverage to the text length + 1, stopped at the first unreadable run, and dropped trailing bytes; surviving runs were still applied"
            },
            Self::ReservedBulletFlagBits => {
                "masked bulletFlags to the bits MS-PPT 2.9.22 defines (0x000F); defined bits were still applied"
            },
            Self::ReservedWrapFlagBits => {
                "masked wrapFlags to the bits MS-PPT 2.9.25 defines (0x0007); defined bits were still applied"
            },
            Self::UnknownTextAlignment => {
                "left an undefined TextAlignmentEnum value unset so the default alignment applies"
            },
            Self::UnknownFontAlignment => "left an undefined TextFontAlignmentEnum value unset",
            Self::UnknownTextDirection => "left an undefined TextDirectionEnum value unset",
            Self::UnknownTabType => {
                "kept a tab stop with an undefined TextTabTypeEnum value as a left tab"
            },
            Self::IndentLevelOutOfRange => "clamped a paragraph indent level above 4 to 4",
            Self::InvalidColorIndex => {
                "treated a ColorIndexStruct with an undefined index as an unset color; the raw value was kept"
            },
        }
    }
}

/// The class of `OfficeArt` shape defect the reader repairs the way
/// `PowerPoint` and LibreOffice repair it (finding F5.11).
///
/// This enumeration is closed by design. Every repair is applied under every
/// [`Leniency`] variant, including [`Leniency::Strict`], on the slide and
/// notes shape read paths. Each recorded entry is located by the persist ID
/// of the slide or notes page and the shape ID (`spid`) of the shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ShapeDefect {
    /// One shape could not be converted and was dropped; its siblings, and
    /// the rest of the slide, are still read (finding F5.11).
    ///
    /// `PowerPoint` and LibreOffice render the remaining shapes of a slide
    /// when one is unreadable, so a content-level failure while one shape is
    /// converted — malformed text styling, a malformed placeholder or
    /// round-trip atom, an unreadable property — drops that shape only. The
    /// shape's own record framing (its `ClientTextbox`/`ClientData` host
    /// records and their child-record headers), resource limits, and the
    /// shape-tree nesting bound stay fatal for the whole slide. Repairs
    /// recorded while the dropped shape was being converted are discarded
    /// with it. The context is the zero-based index of the shape among its
    /// siblings (top-level shapes of the drawing, or children of a group).
    UnreadableShape,
    /// A `ClientTextbox` carried more than one `TextRulerAtom`; the first is
    /// kept and every later one ignored (finding F5.11).
    ///
    /// MS-PPT 2.9.43 defines a single optional ruler, and POI's
    /// `EscherTextboxWrapper` tolerates several. The context is the
    /// one-based ordinal of the ignored ruler (1 for the second ruler).
    DuplicateTextRuler,
    /// A shape's `ClientData` carried a `PowerPoint` 2007 round-trip atom
    /// (`RoundTripHFPlaceholder12Atom`, `RoundTripNewPlaceholderId12Atom`,
    /// `RoundTripShapeId12Atom`, or
    /// `RoundTripShapeCheckSumForCustomLayouts12Atom`) more than once; the
    /// first occurrence is kept (finding F5.11).
    ///
    /// Duplicated round-trip atoms are exactly what round-tripping producers
    /// emit, and none of the reference readers reject them. The context is
    /// the record type of the ignored duplicate.
    DuplicateRoundTripAtom,
    /// A round-trip placeholder atom named a placeholder ID outside the set
    /// MS-PPT defines (7–10 for `RoundTripHFPlaceholder12Atom`, 25–26 for
    /// `RoundTripNewPlaceholderId12Atom`); the identity is treated as absent
    /// (finding F5.11).
    ///
    /// POI never validates these IDs. The context is
    /// `(record type << 16) | placeholder ID`.
    UnknownRoundTripPlaceholderId,
}

impl ShapeDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnreadableShape => {
                "dropped one unreadable shape and kept every other shape on the slide"
            },
            Self::DuplicateTextRuler => {
                "kept the first TextRulerAtom of the text box and ignored the duplicate"
            },
            Self::DuplicateRoundTripAtom => {
                "kept the first round-trip shape atom and ignored the duplicate"
            },
            Self::UnknownRoundTripPlaceholderId => {
                "treated a round-trip placeholder atom with an undefined placeholder ID as absent"
            },
        }
    }
}

/// The class of presentation slide-list defect the reader repairs the way
/// LibreOffice and POI repair it (findings F5.4 and F5.5).
///
/// These repairs happen while the presentation is opened, so they are
/// recorded in [`crate::Presentation::tolerance_report`]. Each entry's
/// [`DefectLocation::persist_id`] is the offending `SlidePersistAtom`'s
/// `persistIdRef` (when it could be read); the context is documented per
/// variant. Every repair is applied under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideListDefect {
    /// A presentation `SlidePersistAtom`'s record header deviated from
    /// recVer 0 / recInstance 0 / recLen 0x14, but the atom still holds the
    /// four fields the reader uses; they are read by offset, as POI
    /// (`SlidePersistAtom.java`) and LibreOffice (`svdfppt.cxx`) read them
    /// (finding F5.4). The context is the atom's payload length.
    SlidePersistAtomHeader,
    /// A presentation `SlidePersistAtom` was too short to hold
    /// `persistIdRef`, `cTexts`, and `slideId`; the entry and its text
    /// records are dropped (finding F5.4). The context is the atom's payload
    /// length.
    TruncatedSlidePersistAtom,
    /// A presentation `SlidePersistAtom` had a null `persistIdRef` or
    /// `slideId`; the entry and its text records are dropped, as LibreOffice
    /// skips entries whose persist reference is unusable (finding F5.4). The
    /// context is the `slideId`.
    NullSlideReference,
    /// A presentation `SlidePersistAtom`'s `cTexts` exceeded 8; it is clamped
    /// to 8 and the slide kept (finding F5.4). POI and LibreOffice never
    /// consult the count. The context is the raw `cTexts`.
    TextPlaceholderCount,
    /// A presentation `SlidePersistAtom` repeated an earlier entry's
    /// `slideId`; the first entry wins and the duplicate is dropped with its
    /// text records (finding F5.4). The context is the `slideId`.
    DuplicateSlideId,
    /// A presentation `SlidePersistAtom` repeated an earlier entry's
    /// `persistIdRef`; the first entry wins and the duplicate is dropped
    /// with its text records (finding F5.4). The context is the `slideId`.
    DuplicateSlidePersistRef,
    /// A presentation `SlidePersistAtom` referenced a persist ID that the
    /// live persist directory does not map, or that maps outside the
    /// document stream or to a record other than a `SlideContainer`; the
    /// slide is skipped, as LibreOffice skips it (`svdfppt.cxx`) (finding
    /// F5.5). Incrementally saved files carry such stale references. The
    /// context is the `slideId`.
    StaleSlidePersistRef,
}

impl SlideListDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::SlidePersistAtomHeader => {
                "read a SlidePersistAtom with a nonconforming record header by field offset"
            },
            Self::TruncatedSlidePersistAtom => {
                "dropped a truncated SlidePersistAtom and its slide-list text"
            },
            Self::NullSlideReference => {
                "dropped a slide-list entry with a null persistIdRef or slideId"
            },
            Self::TextPlaceholderCount => "clamped a SlidePersistAtom cTexts above 8 to 8",
            Self::DuplicateSlideId => {
                "kept the first slide-list entry for a repeated slideId and dropped the duplicate"
            },
            Self::DuplicateSlidePersistRef => {
                "kept the first slide-list entry for a repeated persistIdRef and dropped the duplicate"
            },
            Self::StaleSlidePersistRef => {
                "skipped a slide whose persist reference does not resolve to a SlideContainer"
            },
        }
    }
}

/// The class of `UserEdit`/persist-chain defect the reader repairs the way
/// LibreOffice or POI repairs it (findings F5.6, F5.7, and F5.8).
///
/// These repairs happen while the presentation is opened, so they are
/// recorded in [`crate::Presentation::tolerance_report`]; the context is
/// documented per variant. Every repair is applied under every [`Leniency`]
/// variant. The chain grammar itself — each `UserEditAtom` and persist
/// directory the chain reaches — stays strict, and an encrypted
/// presentation is never opened through a recovered chain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PersistDefect {
    /// The package has no "Current User" stream; the live `UserEditAtom`
    /// was recovered by scanning the document stream's top-level records
    /// for the last `UserEditAtom`, as LibreOffice does (`pptin.cxx` treats
    /// the stream as optional; `svdfppt.cxx` scans) (finding F5.6). The
    /// context is the recovered atom's stream offset.
    MissingCurrentUser,
    /// The "Current User" stream's `offsetToCurrentEdit` was zero, pointed
    /// outside the document stream, or did not land on a well-formed
    /// `UserEditAtom`; the live `UserEditAtom` was recovered by the same
    /// scan as [`Self::MissingCurrentUser`] (finding F5.6). The context is
    /// the recovered atom's stream offset.
    UnusableCurrentEditOffset,
    /// The "Current User" stream's `lenUserName` exceeded the 255 MS-PPT
    /// 2.3.2 allows, or the stream truncated the username and release
    /// fields; the username is treated as absent, as POI treats it
    /// (`CurrentUserAtom.java`), and LibreOffice never reads it (finding
    /// F5.8). The context is the raw `lenUserName`.
    CurrentUserNameLength,
    /// A `UserEditAtom`'s `offsetPersistDirectory` pointed at a record other
    /// than a `PersistDirectoryAtom`; that edit contributes no persist
    /// entries and the chain walk continues, as LibreOffice only processes
    /// the directory when the target record is one (`svdfppt.cxx`) (finding
    /// F5.7). The context is the record type found there.
    PersistDirectoryRecordType,
    /// A `PersistDirectoryEntry` header declared zero offsets; it is skipped
    /// (LibreOffice's read loop simply does not run for it, and POI advances
    /// past it) (finding F5.7). The context is the entry's base persist ID.
    ZeroCountPersistRun,
    /// A persist ID appeared more than once within one persist directory;
    /// the first offset wins, matching LibreOffice (`svdfppt.cxx`) and the
    /// newest-first merge of the chain (finding F5.7). The context is the
    /// persist ID.
    DuplicatePersistId,
}

impl PersistDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::MissingCurrentUser => {
                "recovered the live UserEditAtom by scanning the document stream because the Current User stream is missing"
            },
            Self::UnusableCurrentEditOffset => {
                "recovered the live UserEditAtom by scanning the document stream because the current edit offset is unusable"
            },
            Self::CurrentUserNameLength => {
                "treated a Current User username with an invalid or truncated length as absent"
            },
            Self::PersistDirectoryRecordType => {
                "ignored a persist-directory pointer to a record that is not a PersistDirectoryAtom"
            },
            Self::ZeroCountPersistRun => "skipped a persist-directory entry with a zero count",
            Self::DuplicatePersistId => {
                "kept the first offset of a persist ID repeated within one persist directory"
            },
        }
    }
}

/// One class of defect repaired while a presentation was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a read may change about a presentation. Every class listed today is
/// repaired under every [`Leniency`] variant; opt-in classes join this
/// enumeration together with the [`Leniency`] variants that gate them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A text-property defect repaired the way `PowerPoint` repairs it.
    TextProp(TextPropDefect),
    /// An `OfficeArt` shape defect repaired the way `PowerPoint` repairs it.
    Shape(ShapeDefect),
    /// A slide-list defect repaired at open time the way LibreOffice
    /// repairs it.
    SlideList(SlideListDefect),
    /// A `UserEdit`/persist-chain defect repaired at open time the way
    /// LibreOffice repairs it.
    Persist(PersistDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::TextProp(defect) => defect.repair(),
            Self::Shape(defect) => defect.repair(),
            Self::SlideList(defect) => defect.repair(),
            Self::Persist(defect) => defect.repair(),
        }
    }
}

impl From<TextPropDefect> for Defect {
    fn from(defect: TextPropDefect) -> Self {
        Self::TextProp(defect)
    }
}

impl From<PersistDefect> for Defect {
    fn from(defect: PersistDefect) -> Self {
        Self::Persist(defect)
    }
}

impl From<SlideListDefect> for Defect {
    fn from(defect: SlideListDefect) -> Self {
        Self::SlideList(defect)
    }
}

impl From<ShapeDefect> for Defect {
    fn from(defect: ShapeDefect) -> Self {
        Self::Shape(defect)
    }
}

/// Where in the presentation a repaired defect was found.
///
/// Each field is filled in by the parse layer that knows it; a field stays
/// `None` when the defect was found by a reader that has no such context
/// (for example [`crate::EscherTextboxWrapper::new_with_report`] called on a
/// detached text-box payload, which knows neither its slide nor its shape).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct DefectLocation {
    /// Persist ID of the slide or notes record whose drawing holds the text.
    pub persist_id: Option<u32>,
    /// `OfficeArt` shape ID (`spid`) of the text box holding the text.
    pub shape_id: Option<u32>,
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Class-specific detail. For
    /// [`TextPropDefect::StyleTextPropCoverage`] this is the zero-based index
    /// of the offending run within its `StyleTextPropAtom` (paragraph runs
    /// first, then character runs; a dropped tail records the index where
    /// another run would have started). The reserved-bit classes store the
    /// raw, unmasked 16-bit flags value. Every other class documents its
    /// context on its variant.
    pub context: u32,
    /// The slide/notes page and shape the defect was found in.
    pub location: DefectLocation,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the presentation conformed and no substitution was
/// made. After a [`Leniency::Strict`] read the report can still list
/// [`Defect::TextProp`] repairs, which `PowerPoint` itself applies.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

/// The extent of a [`ToleranceReport`] at one point of a parse.
#[derive(Debug, Clone, Copy)]
pub(crate) struct ReportCheckpoint {
    recorded: usize,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the presentation parsed without any repair.
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
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u32) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                context,
                location: DefectLocation::default(),
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Record one repair whose location is already known.
    pub(crate) fn record_at(
        &mut self,
        defect: impl Into<Defect>,
        context: u32,
        location: DefectLocation,
    ) {
        let start = self.defects.len();
        self.record(defect, context);
        self.locate_from(start, location);
    }

    /// Capture the report's current extent, so the repairs recorded by a
    /// parse that is later abandoned can be discarded with
    /// [`Self::rollback`].
    pub(crate) fn checkpoint(&self) -> ReportCheckpoint {
        ReportCheckpoint {
            recorded: self.defects.len(),
            unrecorded: self.unrecorded,
        }
    }

    /// Discard every repair recorded since `checkpoint` was taken.
    pub(crate) fn rollback(&mut self, checkpoint: ReportCheckpoint) {
        self.defects.truncate(checkpoint.recorded);
        self.unrecorded = checkpoint.unrecorded;
    }

    /// Merge another parse's report into this one, keeping discovery order
    /// and each entry's location.
    ///
    /// The stored-entry bound applies to the merged report as a whole;
    /// entries that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            if self.defects.len() < MAX_RECORDED_DEFECTS {
                self.defects.push(*tolerated);
            } else {
                self.unrecorded = self.unrecorded.saturating_add(1);
            }
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }

    /// Fill in location fields that are still unknown on every defect
    /// recorded at or after index `start`.
    ///
    /// Callers capture `start` as `self.defects().len()` before handing the
    /// report to a lower layer and stamp what they know afterwards. Fields a
    /// lower layer already set are kept.
    pub(crate) fn locate_from(&mut self, start: usize, location: DefectLocation) {
        for tolerated in self.defects.iter_mut().skip(start) {
            if tolerated.location.persist_id.is_none() {
                tolerated.location.persist_id = location.persist_id;
            }
            if tolerated.location.shape_id.is_none() {
                tolerated.location.shape_id = location.shape_id;
            }
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
        report.record(TextPropDefect::StyleTextPropCoverage, 3);
        report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        report.record(TextPropDefect::ReservedWrapFlagBits, 0xFFF8);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::TextProp(TextPropDefect::StyleTextPropCoverage), 3),
                (
                    Defect::TextProp(TextPropDefect::ReservedBulletFlagBits),
                    0x00F0
                ),
                (
                    Defect::TextProp(TextPropDefect::ReservedWrapFlagBits),
                    0xFFF8
                ),
            ]
        );
    }

    /// A hostile style atom must not be able to grow the report without
    /// limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TextPropDefect::StyleTextPropCoverage, index as u32);
        }
        for _ in 0..3 {
            report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TextPropDefect::StyleTextPropCoverage, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TextPropDefect::ReservedWrapFlagBits, 0x00F8);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::TextProp(TextPropDefect::StyleTextPropCoverage),
                context: 7,
                location: DefectLocation::default(),
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn rolling_back_discards_later_repairs_only() {
        let mut report = ToleranceReport::default();
        report.record(TextPropDefect::StyleTextPropCoverage, 0);
        let checkpoint = report.checkpoint();
        for _ in 0..MAX_RECORDED_DEFECTS + 2 {
            report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F1);
        }
        assert_eq!(report.unrecorded(), 3);
        report.rollback(checkpoint);
        assert_eq!(report.total(), 1);
        assert_eq!(
            report.defects()[0].defect,
            Defect::TextProp(TextPropDefect::StyleTextPropCoverage)
        );
    }

    #[test]
    fn locating_fills_only_unknown_fields_of_later_defects() {
        let mut report = ToleranceReport::default();
        report.record(TextPropDefect::StyleTextPropCoverage, 0);
        let start = report.defects().len();
        report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F1);
        report.locate_from(
            start,
            DefectLocation {
                persist_id: None,
                shape_id: Some(1025),
            },
        );
        report.locate_from(
            0,
            DefectLocation {
                persist_id: Some(256),
                shape_id: Some(9999),
            },
        );

        let locations: Vec<_> = report.defects().iter().map(|d| d.location).collect();
        assert_eq!(
            locations,
            vec![
                DefectLocation {
                    persist_id: Some(256),
                    shape_id: Some(9999),
                },
                DefectLocation {
                    persist_id: Some(256),
                    shape_id: Some(1025),
                },
            ]
        );

        // Appending keeps each entry's location.
        let mut merged = ToleranceReport::default();
        merged.append(&report);
        assert_eq!(merged, report);
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(TextPropDefect::StyleTextPropCoverage),
            Defect::from(TextPropDefect::ReservedBulletFlagBits),
            Defect::from(TextPropDefect::ReservedWrapFlagBits),
            Defect::from(TextPropDefect::UnknownTextAlignment),
            Defect::from(TextPropDefect::UnknownFontAlignment),
            Defect::from(TextPropDefect::UnknownTextDirection),
            Defect::from(TextPropDefect::UnknownTabType),
            Defect::from(TextPropDefect::IndentLevelOutOfRange),
            Defect::from(TextPropDefect::InvalidColorIndex),
            Defect::from(SlideListDefect::SlidePersistAtomHeader),
            Defect::from(SlideListDefect::TruncatedSlidePersistAtom),
            Defect::from(SlideListDefect::NullSlideReference),
            Defect::from(SlideListDefect::TextPlaceholderCount),
            Defect::from(SlideListDefect::DuplicateSlideId),
            Defect::from(SlideListDefect::DuplicateSlidePersistRef),
            Defect::from(SlideListDefect::StaleSlidePersistRef),
            Defect::from(PersistDefect::MissingCurrentUser),
            Defect::from(PersistDefect::UnusableCurrentEditOffset),
            Defect::from(PersistDefect::CurrentUserNameLength),
            Defect::from(PersistDefect::PersistDirectoryRecordType),
            Defect::from(PersistDefect::ZeroCountPersistRun),
            Defect::from(PersistDefect::DuplicatePersistId),
            Defect::from(ShapeDefect::UnreadableShape),
            Defect::from(ShapeDefect::DuplicateTextRuler),
            Defect::from(ShapeDefect::DuplicateRoundTripAtom),
            Defect::from(ShapeDefect::UnknownRoundTripPlaceholderId),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::TextProp(TextPropDefect::StyleTextPropCoverage).repair(),
            TextPropDefect::StyleTextPropCoverage.repair()
        );
    }

    /// A conformant presentation opens through the default
    /// ([`Leniency::Strict`]) path with an empty report.
    #[test]
    fn a_conformant_presentation_opens_with_a_clean_report() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt");
        let mut package = crate::Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();

        let report = presentation.tolerance_report();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());

        // Reading every slide's and notes page's text boxes repairs nothing.
        let collected = presentation.collect_tolerance_report().unwrap();
        assert!(collected.is_clean());
        for slide in presentation.slides().unwrap() {
            assert!(slide.tolerance_report().unwrap().is_clean());
        }
    }
}
