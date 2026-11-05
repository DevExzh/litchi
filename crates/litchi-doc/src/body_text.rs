//! Bounded, source-preserving edits for ordinary main-story DOC paragraphs and
//! selected inert auxiliary tables.
//!
//! Length-changing replacements append Unicode text, rebuild the CLX and CHPX
//! FKPs, shift modeled main-story PLCFs, and update the FIB story length. The
//! transaction refuses structural text, tracked ranges, non-uniform character
//! formatting, interior position boundaries, and unmodeled dependencies before
//! publication. Text, direct formatting, revision marks, and managed embedded
//! resources share one immutable multi-operation transaction.

#![deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "the ordinary immutable DOC root uses checked conversions and propagated typed errors"
)]

mod properties;

use properties::MarkChange;
pub use properties::{FieldPart, HyperlinkInsert, SectionValue};

use crate::DateTime;
use crate::package::{Error as PackageError, OpenOptions as PackageOpenOptions};
use crate::parts::dofr::{DofrArray, Patch as DofrPatch};
use crate::parts::protection::ProtectionPolicy;
use crate::parts::saved_selection::{
    Edit as SavedSelectionEdit, FIB_INDEX_WSS, Patch as SavedSelectionPatch, SavedSelection,
    SavedSelectionSpliceError,
};
use crate::tracked_revision::{
    Limits, ParagraphCopy, Revision, RevisionEditor, RevisionKind, RevisionMetadata,
};
use litchi_cfb::OleFile;
use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, CompositionError, JoinedSubEdits, PatchError, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation, SubEdit as CoreSubEdit,
};
pub use litchi_core::patch::{CompositionLimits, HistoryLimits, SubEditJoinFailure};
#[cfg(feature = "performance-diagnostics")]
pub use litchi_ole_common::object::{CfbParseEvent, CfbParseOutcome};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Seek};
use std::sync::{Arc, OnceLock};

pub(crate) fn intersect_revision_limits(left: Limits, right: Limits) -> Limits {
    Limits {
        max_objects: left.max_objects.min(right.max_objects),
        max_storage_depth: left.max_storage_depth.min(right.max_storage_depth),
        max_streams_per_object: left
            .max_streams_per_object
            .min(right.max_streams_per_object),
        max_streams: left.max_streams.min(right.max_streams),
        max_stream_size: left.max_stream_size.min(right.max_stream_size),
        max_object_size: left.max_object_size.min(right.max_object_size),
        max_total_size: left.max_total_size.min(right.max_total_size),
    }
}

/// Bounded undo/redo history over immutable DOC snapshots.
#[cfg(test)]
pub(crate) type History = litchi_core::patch::History<Snapshot>;

/// Configuration for opening an owned body-text package.
///
/// The envelope keeps the package, strict revision-owner, transaction, and
/// protected-publication policies together so every byte-ingress route can
/// apply the same caller-selected limits.  The fields remain private so the
/// set of format-specific knobs can grow without exposing an unstable struct
/// layout.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct OpenOptions {
    source_limits: crate::package::SourceLimits,
    revision_limits: Limits,
    transaction_limits: TransactionLimits,
    protection_policy: ProtectionPolicy,
}

impl OpenOptions {
    /// Selects the complete-source and DOC-entry admission limits.
    #[must_use]
    pub(crate) const fn source_limits(
        mut self,
        source_limits: crate::package::SourceLimits,
    ) -> Self {
        self.source_limits = source_limits;
        self
    }

    /// Selects limits for strict tracked-revision ownership and parsing.
    #[must_use]
    pub(crate) const fn revision_limits(mut self, revision_limits: Limits) -> Self {
        self.revision_limits = revision_limits;
        self
    }

    /// Selects staged operation and replacement-payload limits.
    #[must_use]
    pub(crate) const fn transaction_limits(
        mut self,
        transaction_limits: TransactionLimits,
    ) -> Self {
        self.transaction_limits = transaction_limits;
        self
    }

    /// Selects the explicit policy for changed publication of protected DOCs.
    #[must_use]
    pub(crate) fn protection_policy(mut self, protection_policy: ProtectionPolicy) -> Self {
        self.protection_policy = protection_policy;
        self
    }
}

/// Finite limits for one body transaction.
///
/// The default policy retains at most 8 MiB of already-rendered output
/// allocation capacity between an edit and its commit. [`Self::new`] disables
/// that retention unless [`Self::with_max_retained_render_bytes`] enables it.
/// This ceiling excludes source bytes and temporary rendering/validation
/// allocations; it is not a bound on total process memory. An output above
/// the ceiling is recomputed at commit without changing the result.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TransactionLimits {
    operations: usize,
    replacement_units: usize,
    total_units: usize,
    retained_render_bytes: usize,
}

impl TransactionLimits {
    /// Creates explicit per-transaction bounds with rendered-output retention disabled.
    #[must_use]
    pub const fn new(
        max_operations: usize,
        max_replacement_units: usize,
        max_total_replacement_units: usize,
    ) -> Self {
        Self {
            operations: max_operations,
            replacement_units: max_replacement_units,
            total_units: max_total_replacement_units,
            retained_render_bytes: 0,
        }
    }

    /// Adds a finite ceiling for retaining one already-rendered DOC output.
    ///
    /// A zero ceiling is a supported setting that keeps no rendered handoff;
    /// the transaction recomputes its output through the existing finish path
    /// when it is committed.
    #[must_use]
    pub const fn with_max_retained_render_bytes(mut self, max_bytes: usize) -> Self {
        self.retained_render_bytes = max_bytes;
        self
    }

    /// Maximum staged semantic operations.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.operations
    }

    /// Maximum UTF-16 units in one replacement.
    #[must_use]
    pub const fn max_replacement_units(self) -> usize {
        self.replacement_units
    }

    /// Maximum aggregate UTF-16 replacement units.
    #[must_use]
    pub const fn max_total_replacement_units(self) -> usize {
        self.total_units
    }

    /// Maximum allocation capacity retained for one validated rendered DOC.
    #[must_use]
    pub const fn max_retained_render_bytes(self) -> usize {
        self.retained_render_bytes
    }

    /// Intersects two finite transaction policies component by component.
    ///
    /// The operation that combines values must preserve the stricter member of
    /// every policy, including the retained-render ceiling.
    #[must_use]
    pub fn intersect(self, other: Self) -> Self {
        Self {
            operations: self.operations.min(other.operations),
            replacement_units: self.replacement_units.min(other.replacement_units),
            total_units: self.total_units.min(other.total_units),
            retained_render_bytes: self.retained_render_bytes.min(other.retained_render_bytes),
        }
    }
}

impl Default for TransactionLimits {
    fn default() -> Self {
        Self::new(256, 1024 * 1024, 4 * 1024 * 1024).with_max_retained_render_bytes(8 * 1024 * 1024)
    }
}

/// Main-story text visibility used for a review projection.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Projection {
    /// Stored text, including both insertion and deletion redline text.
    #[default]
    All,
    /// Text visible after accepting insertion and deletion revisions.
    Accepted,
    /// Text visible after rejecting insertion and deletion revisions.
    Rejected,
}

/// One of the seven text stories represented by `FibRgLw97` character counts.
///
/// The same values name a story to the readers
/// ([`crate::Document::story_blocks`]) and to the editors here, so a read
/// position maps onto the story an edit addresses.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Story {
    Main,
    Footnote,
    Header,
    Comment,
    Endnote,
    Textbox,
    HeaderTextbox,
}

impl Story {
    /// Every story, in the order the FIB lays their text out.
    pub const ALL: [Self; 7] = [
        Self::Main,
        Self::Footnote,
        Self::Header,
        Self::Comment,
        Self::Endnote,
        Self::Textbox,
        Self::HeaderTextbox,
    ];

    /// The story's `FibRgLw97` character-count slot (`ccpText` is 0); slot 3,
    /// the retired macro story, still offsets the stories after it.
    pub(crate) const fn character_count_slot(self) -> usize {
        match self {
            Self::Main => 0,
            Self::Footnote => 1,
            Self::Header => 2,
            Self::Comment => 4,
            Self::Endnote => 5,
            Self::Textbox => 6,
            Self::HeaderTextbox => 7,
        }
    }

    const fn index(self) -> usize {
        match self {
            Self::Main => 0,
            Self::Footnote => 1,
            Self::Header => 2,
            Self::Comment => 3,
            Self::Endnote => 4,
            Self::Textbox => 5,
            Self::HeaderTextbox => 6,
        }
    }

    const fn wire_name(self) -> &'static str {
        match self {
            Self::Main => "main",
            Self::Footnote => "footnote",
            Self::Header => "header",
            Self::Comment => "comment",
            Self::Endnote => "endnote",
            Self::Textbox => "textbox",
            Self::HeaderTextbox => "header-textbox",
        }
    }
}

/// A semantic text range owned by the opened DOC transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum TextTarget {
    /// Ordinary paragraph in one DOC story.
    StoryParagraph { story: Story, position: Position },
    /// Simple main-story table cell, excluding its paragraph/cell marks.
    TableCell(Position),
    /// Cached result of one simple, non-nested main-story field.
    FieldResult(Position),
    /// A non-empty UTF-16 range of one ordinary main-story paragraph's text,
    /// counted from the paragraph's first character: one character run.
    BodySpan {
        paragraph: Position,
        start: u32,
        end: u32,
    },
    /// The instruction or the cached result of the main-story field whose
    /// begin mark is at CP `start` (record 0324). The part must hold no
    /// nested field.
    Field { start: u32, part: FieldPart },
}

impl TextTarget {
    /// Convenience constructor for an ordinary main-body paragraph.
    #[must_use]
    pub const fn body_paragraph(position: Position) -> Self {
        Self::StoryParagraph {
            story: Story::Main,
            position,
        }
    }
}

/// One inert text item and its checked semantic selector.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextItem {
    target: TextTarget,
    text: String,
}

impl TextItem {
    #[must_use]
    pub const fn target(&self) -> TextTarget {
        self.target
    }

    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// Direct character property supported without style-table resource creation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum CharacterProperty {
    Bold,
    Italic,
    Underline,
}

/// Valued direct character formatting set without creating or rewriting
/// style-table or font-table resources.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CharacterValue<'a> {
    /// Font size in half-points (`sprmCHps`), from 2 through 3276.
    Size(u16),
    /// Direct RGB text color (`sprmCCv`, plus `sprmCIco` when the color is
    /// one of the sixteen Word 97 palette colors); `None` clears the direct
    /// color so the text is automatic again.
    Color(Option<[u8; 3]>),
    /// The ASCII font (`sprmCRgFtc0`), named exactly as the document's font
    /// table (`SttbfFfn`) names it. An absent name is refused, since adding a
    /// font-table entry is a resource change.
    Font(&'a str),
    /// The character style (`sprmCIstd`), by stylesheet index. The default
    /// character style (index 10) clears the direct style (record 0324).
    Style(u16),
}

impl CharacterValue<'_> {
    /// The property this value sets.
    #[must_use]
    pub const fn property(&self) -> CharacterValueProperty {
        match self {
            Self::Size(_) => CharacterValueProperty::Size,
            Self::Color(_) => CharacterValueProperty::Color,
            Self::Font(_) => CharacterValueProperty::Font,
            Self::Style(_) => CharacterValueProperty::Style,
        }
    }
}

/// The valued character property a [`CharacterValue`] sets.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[non_exhaustive]
pub enum CharacterValueProperty {
    /// Font size.
    Size,
    /// Text color.
    Color,
    /// ASCII font.
    Font,
    /// Character style.
    Style,
}

impl CharacterValueProperty {
    const fn wire_name(self) -> &'static str {
        match self {
            Self::Size => "size",
            Self::Color => "color",
            Self::Font => "font",
            Self::Style => "style",
        }
    }

    /// Every direct SPRM opcode the property owns.
    const fn opcodes(self) -> &'static [u16] {
        match self {
            Self::Size => &[SPRM_C_HPS],
            Self::Color => &[SPRM_C_ICO, SPRM_C_CV],
            Self::Font => &[SPRM_C_RG_FTC0],
            Self::Style => &[SPRM_C_ISTD_VALUE],
        }
    }
}

const SPRM_C_HPS: u16 = 0x4A43;
const SPRM_C_ICO: u16 = 0x2A42;
const SPRM_C_CV: u16 = 0x6870;
const SPRM_C_RG_FTC0: u16 = 0x4A4F;
const SPRM_C_ISTD_VALUE: u16 = 0x4A30;
/// `FibRgFcLcb97` index of the font table, `SttbfFfn`.
const FIB_INDEX_STTBF_FFN: usize = 15;

/// The sixteen Word 97 palette colors, by `sprmCIco` index.
const ICO_PALETTE: [[u8; 3]; 16] = [
    [0, 0, 0],
    [0, 0, 255],
    [0, 255, 255],
    [0, 255, 0],
    [255, 0, 255],
    [255, 0, 0],
    [255, 255, 0],
    [255, 255, 255],
    [0, 0, 128],
    [0, 128, 128],
    [0, 128, 0],
    [128, 0, 128],
    [128, 0, 0],
    [128, 128, 0],
    [128, 128, 128],
    [192, 192, 192],
];

/// Non-destructive tracked-mark disposition supported by the root transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RevisionDisposition {
    Accept,
    Reject,
}

/// Active binary dependency carried by a special DOC text character.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum DrawingDependency {
    /// `0x0001`: inline picture or an embedded object's preview PICF.
    InlinePictureOrObjectPreview,
    /// `0x0008`: floating shape/picture with `PlcfSpa` and drawing-group edges.
    FloatingOfficeArt,
    /// The selected picture has ambiguous/nested ownership, noncanonical
    /// records, PICF transforms, producer extensions, or a delayed BLIP.
    UnsupportedPictureGraph,
    /// The receiver already owns picture/shape identifiers or a BLIP store.
    PictureGraphCollision,
    /// Picture graph rewriting outside the main story is not modeled.
    AuxiliaryStoryPicture,
    /// `0xFFFC`: producer-defined object replacement character.
    ObjectReplacement,
    /// Another active control whose coupled binary owner is not modeled here.
    UnsupportedControl,
}

impl DrawingDependency {
    const fn transfer_name(self) -> &'static str {
        match self {
            Self::InlinePictureOrObjectPreview => "inline-picture/embedded-preview",
            Self::FloatingOfficeArt => "floating-OfficeArt",
            Self::UnsupportedPictureGraph => "unsupported-picture-graph",
            Self::PictureGraphCollision => "picture-graph-collision",
            Self::AuxiliaryStoryPicture => "auxiliary-story-picture",
            Self::ObjectReplacement => "object-replacement",
            Self::UnsupportedControl => "active-control",
        }
    }
}

impl CharacterProperty {
    const fn wire_name(self) -> &'static str {
        match self {
            Self::Bold => "bold",
            Self::Italic => "italic",
            Self::Underline => "underline",
        }
    }
}

/// A visible ordinary paragraph in the main document story.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    position: Position,
    text: String,
}

impl Paragraph {
    /// Zero-based paragraph position in the selected projection.
    ///
    /// Constructing a [`Position`] is infallible. Resolving it against a
    /// snapshot collection is checked by [`crate::Edit::replace_paragraph`].
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Plain inert paragraph text, without the terminating paragraph mark.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A reason why an edit is outside this intentionally small safe closure.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected [`Position`] does not exist in the source body.
    ParagraphNotFound,
    /// Retained for source compatibility; length changes are now modeled by
    /// the bounded CLX/FKP transaction.
    LengthChange { expected: usize, actual: usize },
    /// The paragraph crosses pieces, which can have distinct encodings and PRMs.
    CrossesPieceBoundary,
    /// The selected paragraph is stored in an ANSI/compressed piece.
    CompressedPiece,
    /// Fields, object markers, cell markers, or other structural controls occur.
    StructuralContent,
    /// A special character has a drawing/resource owner that cannot be
    /// rewritten safely by a whole-target text replacement.
    DrawingDependency { dependency: DrawingDependency },
    /// The paragraph intersects text-affecting tracked revisions.
    TrackedText,
    /// The requested replacement contains structural controls.
    ReplacementContainsStructuralContent,
    /// The source's review ranges overlap in a way this projection cannot prove.
    AmbiguousReviewRanges,
    /// An empty paragraph has no text run from which formatting can be copied.
    EmptyParagraph,
    /// Character formatting or another CP-bound dependency is not uniform or
    /// cannot be shifted without changing its meaning.
    FormattingDependency,
    /// A known CP-indexed table has coupled records outside the resize model.
    PositionDependency { fib_index: usize },
    /// Non-main stories currently permit exact-length replacement only.
    StoryLengthChange {
        story: Story,
        expected: usize,
        actual: usize,
    },
    /// The requested field is nested or lacks one balanced separator/result.
    ComplexField,
    /// The requested table cell has multiple paragraphs or structural content.
    ComplexTableCell,
    /// A semantic target is absent from the selected collection.
    TargetNotFound,
    /// Cross-document transfer is limited to inert text with no resource edges.
    TransferDependency { dependency: &'static str },
    /// The receiver already owns the requested semantic resource identifier.
    ResourceCollision { storage_id: u32 },
    /// No managed embedded resource owns the requested identifier.
    ResourceNotFound { storage_id: u32 },
    /// This disposition would delete text or require restoring prior formatting.
    DestructiveRevisionDisposition { kind: RevisionKind },
    /// A changed DOC publication would leave retained binary signature
    /// metadata stale.
    SignedSource,
    /// Document or range protection is active for this edit policy.
    Protected,
    /// Protection metadata is malformed or outside the supported grammar.
    UnknownProtection,
    /// Opaque auxiliary-table bytes cannot be merged by the semantic
    /// three-way body plan without silently dropping one side.
    AuxiliaryThreeWayMergeUnsupported,
    /// The configured operation count was exhausted.
    OperationLimit { observed: usize, limit: usize },
    /// One or all replacement payloads exceed the configured UTF-16 bound.
    ReplacementLimit { observed: usize, limit: usize },
    /// The document's font table does not name the requested font.
    FontNotFound,
    /// The requested font size is outside 1 through 1638 points.
    InvalidFontSize,
    /// Removing the main story's only ordinary paragraph would leave no body.
    SoleParagraph,
    /// The structural edit would join two tables or end the main story in a
    /// table.
    TableBoundary,
    /// The paragraph carries a field, bookmark, note, comment, drawing or
    /// proofing anchor, or a section boundary, that the structural edit
    /// would collapse.
    AnchoredContent,
    /// The package is encrypted or obfuscated; the body editor never
    /// rewrites protected streams.
    Encrypted,
    /// A bookmark, smart tag, comment range, field, note or proofing range
    /// recorded in the CP-indexed FIB table `fib_index` has a boundary
    /// strictly inside the replaced text; replacing the text would collapse
    /// that boundary.
    RangeBoundary { fib_index: usize },
    /// The stylesheet has no style of the required kind at that index.
    StyleNotFound,
    /// A property value is outside what the DOC structure can store.
    InvalidProperty,
    /// No main-story field begins at the requested position.
    FieldNotFound,
    /// A bookmark name is empty, too long, holds other characters than
    /// letters, digits, `_`, `.` and `-`, or is already used.
    BookmarkName,
    /// No standard bookmark has the requested name.
    BookmarkNotFound,
    /// The requested section is past the document's last section.
    SectionNotFound,
}

/// What the CP-indexed FIB table `index` positions, for refusal messages
/// ([MS-DOC] 2.5.6–2.5.10, `FibRgFcLcb97` through `FibRgFcLcb2007`).
fn cp_table_content(index: usize) -> &'static str {
    match index {
        2 | 46 => "note reference",
        4 | 42 | 43 => "comment",
        6 => "section",
        16 => "field",
        22 | 23 | 142 | 143 | 124 | 125 => "bookmark",
        40 => "drawing anchor",
        115 | 117 | 132 => "smart tag",
        121 | 122 => "format consistency mark",
        138 | 139 => "content control",
        55 | 90 | 89 | 93 => "proofing",
        _ => "CP range",
    }
}

impl Refusal {
    /// The shared refusal kind (record 0351 §2.4), by which litchi-ops
    /// classifies the refusal (record 0362). The budget variants
    /// (`OperationLimit`, `ReplacementLimit`) are limits, not refusals:
    /// litchi-ops reports them as `limit_exceeded` before consulting the
    /// kind, which is `InvalidValue` for them.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;

        match self {
            Self::ParagraphNotFound
            | Self::TargetNotFound
            | Self::ResourceNotFound { .. }
            | Self::FontNotFound
            | Self::StyleNotFound
            | Self::FieldNotFound
            | Self::BookmarkNotFound
            | Self::SectionNotFound => RefusalKind::NotFound,
            Self::SignedSource | Self::Protected | Self::UnknownProtection | Self::Encrypted => {
                RefusalKind::Protected
            },
            Self::ResourceCollision { .. }
            | Self::InvalidFontSize
            | Self::InvalidProperty
            | Self::BookmarkName
            | Self::OperationLimit { .. }
            | Self::ReplacementLimit { .. } => RefusalKind::InvalidValue,
            Self::EmptyParagraph | Self::SoleParagraph => RefusalKind::UnsupportedTarget,
            Self::LengthChange { .. }
            | Self::CrossesPieceBoundary
            | Self::CompressedPiece
            | Self::StructuralContent
            | Self::DrawingDependency { .. }
            | Self::TrackedText
            | Self::ReplacementContainsStructuralContent
            | Self::AmbiguousReviewRanges
            | Self::FormattingDependency
            | Self::PositionDependency { .. }
            | Self::StoryLengthChange { .. }
            | Self::ComplexField
            | Self::ComplexTableCell
            | Self::TransferDependency { .. }
            | Self::DestructiveRevisionDisposition { .. }
            | Self::AuxiliaryThreeWayMergeUnsupported
            | Self::TableBoundary
            | Self::AnchoredContent
            | Self::RangeBoundary { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ParagraphNotFound => {
                formatter.write_str("body paragraph position is out of range")
            },
            Self::LengthChange { expected, actual } => write!(
                formatter,
                "replacement has {actual} UTF-16 units; this source paragraph has {expected}"
            ),
            Self::CrossesPieceBoundary => {
                formatter.write_str("body paragraph crosses DOC text pieces")
            },
            Self::CompressedPiece => {
                formatter.write_str("body paragraph is stored in a compressed DOC text piece")
            },
            Self::StructuralContent => {
                formatter.write_str("body paragraph contains DOC structural content")
            },
            Self::DrawingDependency { dependency } => write!(
                formatter,
                "DOC text target contains an active {dependency:?} dependency"
            ),
            Self::TrackedText => formatter.write_str("body paragraph intersects tracked text"),
            Self::ReplacementContainsStructuralContent => {
                formatter.write_str("replacement contains DOC structural content")
            },
            Self::AmbiguousReviewRanges => {
                formatter.write_str("tracked revision ranges overlap ambiguously")
            },
            Self::EmptyParagraph => {
                formatter.write_str("empty body paragraph has no editable text formatting run")
            },
            Self::FormattingDependency => formatter.write_str(
                "body paragraph has character formatting or CP-bound dependencies that cannot be resized losslessly",
            ),
            Self::PositionDependency { fib_index } => write!(
                formatter,
                "body length change depends on unmodeled CP-indexed FIB table {fib_index}"
            ),
            Self::StoryLengthChange {
                story,
                expected,
                actual,
            } => write!(
                formatter,
                "{story:?} story replacement changes UTF-16 length from {expected} to {actual}"
            ),
            Self::ComplexField => formatter.write_str(
                "field result is nested, unbalanced, or otherwise outside the simple field model",
            ),
            Self::ComplexTableCell => formatter.write_str(
                "table cell has multiple paragraphs or structural content outside the simple cell model",
            ),
            Self::TargetNotFound => formatter.write_str("DOC semantic text target is out of range"),
            Self::TransferDependency { dependency } => write!(
                formatter,
                "DOC transfer has an unsupported {dependency} dependency"
            ),
            Self::ResourceCollision { storage_id } => write!(
                formatter,
                "DOC receiver already contains embedded storage ID {storage_id}"
            ),
            Self::ResourceNotFound { storage_id } => write!(
                formatter,
                "DOC embedded storage ID {storage_id} was not found"
            ),
            Self::DestructiveRevisionDisposition { kind } => write!(
                formatter,
                "{kind:?} revision disposition is destructive and outside reversible mark editing"
            ),
            Self::SignedSource => {
                formatter.write_str("changed DOC edits are refused for signed sources")
            },
            Self::Protected => {
                formatter.write_str("protected DOC sources require explicit edit authorization")
            },
            Self::UnknownProtection => formatter.write_str(
                "DOC protection metadata is malformed or outside the supported grammar",
            ),
            Self::AuxiliaryThreeWayMergeUnsupported => formatter.write_str(
                "three-way body plans do not merge opaque auxiliary-table changes",
            ),
            Self::OperationLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} operations; limit is {limit}"
            ),
            Self::ReplacementLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} UTF-16 replacement units; limit is {limit}"
            ),
            Self::FontNotFound => {
                formatter.write_str("the DOC font table does not name the requested font")
            },
            Self::InvalidFontSize => {
                formatter.write_str("DOC font size must be 1 through 1638 points in half points")
            },
            Self::SoleParagraph => {
                formatter.write_str("the only ordinary body paragraph cannot be removed")
            },
            Self::TableBoundary => formatter.write_str(
                "the structural edit would join two tables or end the main story in a table",
            ),
            Self::AnchoredContent => formatter.write_str(
                "the paragraph carries an anchored range or section boundary that the edit would collapse",
            ),
            Self::Encrypted => {
                formatter.write_str("encrypted or obfuscated DOC packages cannot be edited")
            },
            Self::RangeBoundary { fib_index } => write!(
                formatter,
                "the replaced text holds a {} boundary (CP-indexed FIB table {fib_index}); replacing it would collapse that range",
                cp_table_content(*fib_index)
            ),
            Self::StyleNotFound => formatter
                .write_str("the DOC stylesheet has no style of the required kind at that index"),
            Self::InvalidProperty => {
                formatter.write_str("the DOC property value is outside its stored range")
            },
            Self::FieldNotFound => {
                formatter.write_str("no DOC main-story field begins at the requested position")
            },
            Self::BookmarkName => formatter.write_str(
                "DOC bookmark names are 1 through 39 letters, digits, `_`, `.` or `-` and unique",
            ),
            Self::BookmarkNotFound => {
                formatter.write_str("no DOC bookmark has the requested name")
            },
            Self::SectionNotFound => {
                formatter.write_str("the DOC section position is out of range")
            },
        }
    }
}

/// Failure from a body-text transaction or source-checked patch.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The DOC/CFB source or its required invariants is invalid.
    Invalid(PackageError),
    /// The request is valid in general but unsafe for this preservation seam.
    Refused(Refusal),
    /// A patch or replay plan was presented with a different source snapshot.
    PatchConflict,
    /// Two three-way branches write one semantic facet differently.
    Conflict,
    /// A shared bounded composition rejected staged work.
    Composition(CompositionError),
    /// A durable semantic patch is malformed or exceeds its explicit limits.
    Durable(PatchError),
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::PatchConflict => {
                formatter.write_str("DOC body-text patch base does not match recipient")
            },
            Self::Conflict => formatter.write_str("body-text three-way composition conflict"),
            Self::Composition(error) => error.fmt(formatter),
            Self::Durable(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Invalid(error) => Some(error),
            Self::Composition(error) => Some(error),
            Self::Durable(error) => Some(error),
            Self::Refused(_) | Self::PatchConflict | Self::Conflict => None,
        }
    }
}

impl From<CompositionError> for Error {
    fn from(error: CompositionError) -> Self {
        Self::Composition(error)
    }
}

impl From<PatchError> for Error {
    fn from(error: PatchError) -> Self {
        Self::Durable(error)
    }
}

fn map_protection_error(error: PackageError) -> Error {
    match error {
        PackageError::ProtectionDenied(
            crate::parts::protection::EditProtection::Unknown
            | crate::parts::protection::EditProtection::Unrecognized,
        ) => Error::Refused(Refusal::UnknownProtection),
        PackageError::ProtectionDenied(_) => Error::Refused(Refusal::Protected),
        error => Error::Invalid(error),
    }
}

/// A deterministic phase boundary exposed by the opt-in performance
/// diagnostics seam.
///
/// The phases describe format-owned work rather than implementation objects or
/// physical identifiers. They intentionally carry no source content, byte
/// offsets, CFB stream names, or package identifiers.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticPhase {
    /// Validation by the strict DOC owner used for safe editing.
    StrictOwnerValidation,
    /// Complete validation by the public package/document reader.
    PublicReaderValidation,
    /// Ownership allocation retaining the exact source for the immutable snapshot.
    SourceRetention,
    /// In-memory owner rendering of the staged editor into candidate bytes.
    Finish,
    /// The candidate matched the source byte-for-byte.
    ExactNoOp,
    /// Construction of the reversible semantic patch.
    Patch,
}

/// Result class attached to a completed diagnostic phase.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticOutcome {
    /// The phase completed and its enclosing operation continued.
    Success,
    /// The phase returned an ordinary typed operation error.
    Error,
}

/// One ordered, content-free phase event emitted by a profiled operation.
///
/// Every started phase has one matching finished event, including phases that
/// return an error. An observer is called synchronously in operation order;
/// the library neither timestamps events nor consults ambient runtime state.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticEvent {
    /// A phase is about to run.
    Started { phase: DiagnosticPhase },
    /// A phase has returned to its caller.
    Finished {
        phase: DiagnosticPhase,
        outcome: DiagnosticOutcome,
    },
}

#[cfg(feature = "performance-diagnostics")]
fn observe_phase<T, Operation>(
    observer: &mut impl FnMut(DiagnosticEvent),
    phase: DiagnosticPhase,
    operation: Operation,
) -> Result<T>
where
    Operation: FnOnce() -> Result<T>,
{
    observer(DiagnosticEvent::Started { phase });
    let result = operation();
    observer(DiagnosticEvent::Finished {
        phase,
        outcome: if result.is_ok() {
            DiagnosticOutcome::Success
        } else {
            DiagnosticOutcome::Error
        },
    });
    result
}

/// Immutable, exact-source snapshot for the body-text transaction seam.
pub(crate) struct Snapshot {
    source: Arc<[u8]>,
    /// A private mutation-capable candidate for a source carrying only CFB
    /// signature metadata. The published owner always retains `source`; this
    /// candidate exists solely so an eager edit can validate and prove an
    /// exact no-op without admitting a protected container to the object
    /// editor.
    candidate_source: Option<Arc<[u8]>>,
    /// Whether the retained source carries any signature marker recognized by
    /// the strict DOC owner or by the CFB signature editor.
    signed: bool,
    fingerprint_cache: OnceLock<u64>,
    limits: Limits,
    source_limits: crate::package::SourceLimits,
    transaction_limits: TransactionLimits,
    protection_policy: ProtectionPolicy,
}

impl Clone for Snapshot {
    fn clone(&self) -> Self {
        let fingerprint_cache = OnceLock::new();
        if let Some(value) = self.fingerprint_cache.get() {
            if fingerprint_cache.set(*value).is_err() {
                unreachable!("a fresh fingerprint cache must accept its cloned value");
            }
        }
        Self {
            source: Arc::clone(&self.source),
            candidate_source: self.candidate_source.as_ref().map(Arc::clone),
            signed: self.signed,
            fingerprint_cache,
            limits: self.limits,
            source_limits: self.source_limits,
            transaction_limits: self.transaction_limits,
            protection_policy: self.protection_policy.clone(),
        }
    }
}

impl Snapshot {
    /// Opens an owned Word 97+ DOC source after validating its safe edit basis.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the CFB, Word 97+ FIB, selected table
    /// stream, piece table, or FKP basis cannot support safe editing.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn open(input: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::open_bounded(input, limits, TransactionLimits::default())
    }

    /// Opens an owned DOC source with explicit package and transaction bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the complete package or its safe edit
    /// basis cannot be reopened.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn open_bounded(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
    ) -> Result<Self> {
        Self::open_bounded_with_policy(
            input,
            limits,
            transaction_limits,
            ProtectionPolicy::default(),
        )
    }

    /// Opens an owned DOC source with an explicit protected-edit policy.
    ///
    /// Protected sources remain readable and support exact no-op publication
    /// under the default policy. Supplying
    /// [`ProtectionPolicy::AllowProtected`] is an explicit capability that
    /// authorizes changed publication.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn open_bounded_with_policy(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        Self::from_bytes_with_options(
            input,
            OpenOptions::default()
                .revision_limits(limits)
                .transaction_limits(transaction_limits)
                .protection_policy(protection_policy),
        )
    }

    fn open_with_options(input: impl Into<Vec<u8>>, options: OpenOptions) -> Result<Self> {
        Self::open_checked(input.into(), options, false)
    }

    /// [`Self::from_bytes_with_options`] over the exact bytes of a package.
    /// With `reader_validated`, the package's document is the strict public
    /// reader's parse of these bytes under `options`' source limits, so the
    /// reader gate below does not parse them again; the strict owner still
    /// opens them (record 0377a).
    pub(crate) fn from_package_bytes(
        input: Vec<u8>,
        options: OpenOptions,
        reader_validated: bool,
    ) -> Result<Self> {
        Self::open_checked(input, options, reader_validated)
    }

    fn open_checked(bytes: Vec<u8>, options: OpenOptions, reader_validated: bool) -> Result<Self> {
        crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(
            bytes.len(),
            options.source_limits,
        )
        .map_err(Error::Invalid)?;
        let source_limits = options.source_limits;
        let (strict_editor, mut ole, candidate_source) =
            match RevisionEditor::open_with_ole_file_with_policy_and_admission(
                bytes.clone(),
                options.revision_limits,
                options.protection_policy.clone(),
                |ole| crate::package::preflight_declared_streams(ole, source_limits).map(|_| ()),
            ) {
                Ok((editor, ole)) => (editor, ole, None),
                Err(original_error) => {
                    validate_original_source_for_fallback(&bytes, source_limits)?;
                    let Some(candidate) = unsigned_doc_source(&bytes) else {
                        return Err(Error::Invalid(original_error));
                    };
                    let candidate: Arc<[u8]> = Arc::from(candidate.into_boxed_slice());
                    let (editor, ole) =
                        RevisionEditor::open_with_ole_file_with_policy_and_admission(
                            candidate.to_vec(),
                            options.revision_limits,
                            options.protection_policy.clone(),
                            |ole| {
                                crate::package::preflight_declared_streams(ole, source_limits)
                                    .map(|_| ())
                            },
                        )
                        .map_err(Error::Invalid)?;
                    (editor, ole, Some(candidate))
                },
            };
        reject_malformed_signature_containers(&ole)?;
        // An unsigned candidate is not the bytes the package's parse covered.
        if reader_validated && candidate_source.is_none() {
            crate::raw::Reader::validate_package_ole_file(&ole, source_limits)
        } else {
            crate::raw::Reader::validate_ole_file(&mut ole, source_limits)
        }
        .map_err(Error::Invalid)?;
        let signed = candidate_source.is_some() || strict_editor.dofr_source_is_signed();
        Ok(Self {
            source: Arc::from(bytes.into_boxed_slice()),
            candidate_source,
            signed,
            fingerprint_cache: OnceLock::new(),
            limits: options.revision_limits,
            source_limits: options.source_limits,
            transaction_limits: options.transaction_limits,
            protection_policy: options.protection_policy,
        })
    }

    /// Opens an owned DOC source while reporting both semantic phases and the
    /// operation-local top-level in-memory CFB parse separately.
    ///
    /// The [`DiagnosticEvent`] stream has the same fixed phase sequence as
    /// [`Self::open_bounded_profiled`]. The separate [`CfbParseEvent`]
    /// observer receives one balanced pair around the successful strict
    /// owner's top-level `OleFile::open` when source-length preflight and
    /// strict-owner input validation permit parsing. A signed CFB fallback
    /// discards the failed strict parse events and reports the candidate parse
    /// instead. Preflight itself is intentionally silent in the high-level
    /// stream.
    ///
    /// # Panics
    ///
    /// Observer panics propagate. Event balancing is guaranteed only when
    /// both observers return normally from every callback.
    ///
    /// # Errors
    ///
    /// Returns the same error as [`Self::open_bounded`]. Failed phases and
    /// physical parses close their respective observer events before the
    /// error is returned.
    #[cfg(feature = "performance-diagnostics")]
    pub(crate) fn open_bounded_profiled_with_cfb_observer(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
        source_limits: crate::package::SourceLimits,
        protection_policy: ProtectionPolicy,
        observer: impl FnMut(DiagnosticEvent),
        cfb_observer: impl FnMut(CfbParseEvent),
    ) -> Result<Self> {
        Self::open_bounded_profiled_with_source_limits(
            input,
            limits,
            transaction_limits,
            source_limits,
            protection_policy,
            observer,
            cfb_observer,
        )
    }

    #[cfg(feature = "performance-diagnostics")]
    fn open_bounded_profiled_with_source_limits(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
        source_limits: crate::package::SourceLimits,
        protection_policy: ProtectionPolicy,
        mut observer: impl FnMut(DiagnosticEvent),
        mut cfb_observer: impl FnMut(CfbParseEvent),
    ) -> Result<Self> {
        let bytes = input.into();
        crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(bytes.len(), source_limits)
            .map_err(Error::Invalid)?;
        let mut failed_strict_events = Vec::new();
        let (mut ole, signed, candidate_source) = observe_phase(
            &mut observer,
            DiagnosticPhase::StrictOwnerValidation,
            || {
                let strict_result = RevisionEditor::open_with_ole_file_profiled_with_admission(
                    bytes.clone(),
                    limits,
                    protection_policy.clone(),
                    |event| {
                        failed_strict_events.push(event);
                    },
                    |ole| {
                        crate::package::preflight_declared_streams(ole, source_limits).map(|_| ())
                    },
                );
                match strict_result {
                    Ok((strict_editor, ole)) => {
                        for event in failed_strict_events.drain(..) {
                            cfb_observer(event);
                        }
                        Ok((ole, strict_editor.dofr_source_is_signed(), None))
                    },
                    Err(original_error) => {
                        if let Err(error) =
                            validate_original_source_for_fallback(&bytes, source_limits)
                        {
                            for event in failed_strict_events.drain(..) {
                                cfb_observer(event);
                            }
                            return Err(error);
                        }
                        let Some(candidate) = unsigned_doc_source(&bytes) else {
                            for event in failed_strict_events.drain(..) {
                                cfb_observer(event);
                            }
                            return Err(Error::Invalid(original_error));
                        };
                        let candidate: Arc<[u8]> = Arc::from(candidate.into_boxed_slice());
                        let (_editor, ole) =
                            RevisionEditor::open_with_ole_file_profiled_with_admission(
                                candidate.to_vec(),
                                limits,
                                protection_policy.clone(),
                                cfb_observer,
                                |ole| {
                                    crate::package::preflight_declared_streams(ole, source_limits)
                                        .map(|_| ())
                                },
                            )
                            .map_err(Error::Invalid)?;
                        Ok((ole, true, Some(candidate)))
                    },
                }
            },
        )?;
        reject_malformed_signature_containers(&ole)?;
        observe_phase(
            &mut observer,
            DiagnosticPhase::PublicReaderValidation,
            || {
                crate::raw::Reader::validate_ole_file(&mut ole, source_limits)
                    .map_err(Error::Invalid)
            },
        )?;
        let source = observe_phase(&mut observer, DiagnosticPhase::SourceRetention, || {
            Ok(Arc::from(bytes.into_boxed_slice()))
        })?;
        Ok(Self {
            source,
            candidate_source,
            signed,
            fingerprint_cache: OnceLock::new(),
            limits,
            source_limits,
            transaction_limits,
            protection_policy,
        })
    }

    /// Parses a borrowed DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn parse(input: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(input.to_vec(), OpenOptions::default())
    }

    /// Opens an owned DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn from_bytes(input: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_options(input, OpenOptions::default())
    }

    /// Opens owned DOC bytes under one explicit configuration envelope.
    ///
    /// The selected package and strict-owner limits are retained by the
    /// immutable snapshot and reused for every changed commit and inverse
    /// patch application.
    pub(crate) fn from_bytes_with_options(
        input: impl Into<Vec<u8>>,
        options: OpenOptions,
    ) -> Result<Self> {
        Self::open_with_options(input, options)
    }

    /// Exact CFB source bytes retained for source checks and byte-exact no-ops.
    #[must_use]
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub(crate) fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Stable first-stage fingerprint for diagnostics and stale-source checks.
    ///
    /// The FNV-1a pass is lazy. Opening, cloning, and constructing an
    /// in-memory patch do not scan the complete artifact solely to populate
    /// this diagnostic value. A clone copies an already initialized value;
    /// otherwise each immutable snapshot may initialize its own cache.
    #[must_use]
    pub(crate) fn fingerprint(&self) -> u64 {
        *self
            .fingerprint_cache
            .get_or_init(|| fingerprint(&self.source))
    }

    /// Configured finite semantic-operation and text-payload bounds.
    #[must_use]
    pub(crate) const fn transaction_limits(&self) -> TransactionLimits {
        self.transaction_limits
    }

    /// Revision-owner limits retained by this immutable body snapshot.
    pub(crate) const fn revision_limits(&self) -> Limits {
        self.limits
    }

    /// Protected-publication policy retained by this immutable body snapshot.
    pub(crate) fn protection_policy(&self) -> ProtectionPolicy {
        self.protection_policy.clone()
    }

    /// Configured outer-package and aggregate stream limits retained by this
    /// immutable source owner.
    #[must_use]
    pub(crate) const fn source_limits(&self) -> crate::package::SourceLimits {
        self.source_limits
    }

    /// Lists ordinary source-body paragraphs under the requested review projection.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an invalid source or [`Error::Refused`]
    /// when tracked text ranges overlap ambiguously for the projection.
    pub(crate) fn paragraphs(&self, projection: Projection) -> Result<Vec<Paragraph>> {
        let editor = self.editor()?;
        projected_paragraphs(&editor, projection)
    }

    /// Lists ordinary paragraphs from any DOC story. Main-story review
    /// projection remains available through [`Self::paragraphs`]; non-main
    /// stories expose their exact stored text.
    pub(crate) fn story_paragraphs(&self, story: Story) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::Story(story)).map(items_from_spans)
    }

    /// Lists simple main-story table cells that can be edited without changing
    /// table marks or PAPX/TAP structure.
    pub(crate) fn table_cells(&self) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::TableCells).map(items_from_spans)
    }

    /// Lists cached results of balanced, non-nested main-story fields.
    pub(crate) fn field_results(&self) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::FieldResults).map(items_from_spans)
    }

    /// Managed embedded fields and their passive `ObjectPool` metadata.
    pub(crate) fn embedded_objects(&self) -> Result<crate::embedded_object::Inventory> {
        crate::embedded_object::Snapshot::open(self.finish(), self.limits)
            .map_err(Error::Invalid)?
            .inventory()
            .map_err(Error::Invalid)
    }

    /// Lists tracked main-story ranges in stable CP/kind order.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn revisions(&self) -> Result<Vec<Revision>> {
        self.editor()?.revisions().map_err(Error::Invalid)
    }

    /// Reads the optional inert `Selsf` cache through the strict DOC owner.
    ///
    /// The public value exposes only the fixed record bytes, while its private
    /// owner handle keeps source-bound edits tied to this immutable snapshot.
    /// No selection is applied to document or host UI state.
    pub(crate) fn saved_selection(&self) -> Result<Option<SavedSelection>> {
        self.editor()?
            .saved_selection()
            .map(|selection| selection.map(|selection| selection.with_owner(self.bytes_shared())))
            .map_err(Error::Invalid)
    }

    /// Reads the optional bounded `RgDofr` frame-set/list records through the
    /// strict DOC owner. Frame names and paths remain inert bytes.
    pub(crate) fn dofr_records(&self) -> Result<Option<DofrArray>> {
        self.editor()?
            .dofr_records()
            .map(|records| records.map(|records| records.with_owner(self.bytes_shared())))
            .map_err(Error::Invalid)
    }

    /// Starts a staged bounded body text-and-formatting transaction.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] if the retained source no longer validates.
    pub(crate) fn edit(&self) -> Result<Edit> {
        Edit::new(self.clone())
    }

    /// Starts a staged transaction under the transaction profile of
    /// `options`. Every dimension is intersected with the source snapshot's
    /// retained policy, so a looser profile can never discard a stricter
    /// ingress or edit bound already attached to the source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] if the retained source no longer validates.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    #[allow(
        dead_code,
        reason = "the test crate exercises the engine through its public package route"
    )]
    pub(crate) fn edit_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        let mut edit = Edit::new(self.clone())?;
        edit.source.transaction_limits = edit
            .source
            .transaction_limits
            .intersect(options.transaction_limits_value());
        edit.editor
            .set_retained_render_limit(edit.source.transaction_limits.max_retained_render_bytes());
        Ok(edit)
    }

    /// Starts a bounded disjoint composition for this exact artifact.
    #[must_use]
    pub(crate) fn compose(&self, limits: CompositionLimits) -> Composition {
        Composition {
            source: self.clone(),
            joined: JoinedSubEdits::new(self.lineage(), limits),
            transaction_limits: self.transaction_limits,
        }
    }

    /// Starts explicit bounded undo/redo history at this immutable snapshot.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Prepares one independently composable paragraph replacement against
    /// this exact immutable artifact.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as [`Edit::replace_paragraph`]
    /// or a common composition-bound error.
    pub(crate) fn prepare_replace(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        replacement: impl Into<String>,
    ) -> Result<SubEdit> {
        let replacement = replacement.into();
        let mut validation = self.edit()?;
        let target = TextTarget::body_paragraph(position);
        validation.replace_text(target, &replacement)?;
        SubEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Text {
                target,
                replacement,
            },
            self.transaction_limits,
        )
    }

    /// Prepares one independently composable direct-bold change.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as
    /// [`Edit::set_paragraph_bold`] or a common composition-bound error.
    pub(crate) fn prepare_bold(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        enabled: bool,
    ) -> Result<SubEdit> {
        let mut validation = self.edit()?;
        let target = TextTarget::body_paragraph(position);
        validation.set_character_property(target, CharacterProperty::Bold, enabled)?;
        SubEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Format {
                target,
                property: CharacterProperty::Bold,
                enabled,
            },
            self.transaction_limits,
        )
    }

    /// Prepares one independently composable semantic text replacement.
    pub(crate) fn prepare_text(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        replacement: impl Into<String>,
    ) -> Result<SubEdit> {
        let replacement = replacement.into();
        let mut validation = self.edit()?;
        validation.replace_text(target, &replacement)?;
        SubEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Text {
                target,
                replacement,
            },
            self.transaction_limits,
        )
    }

    /// Prepares one independently composable direct-format mutation.
    pub(crate) fn prepare_format(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    ) -> Result<SubEdit> {
        let mut validation = self.edit()?;
        validation.set_character_property(target, property, enabled)?;
        SubEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Format {
                target,
                property,
                enabled,
            },
            self.transaction_limits,
        )
    }

    /// Applies supported durable body-text operations to this exact artifact.
    ///
    /// # Errors
    ///
    /// Returns an error for a foreign vocabulary, malformed selector,
    /// artifact or semantic precondition conflict, bound violation, or typed
    /// body-edit refusal.
    pub(crate) fn stage_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<Edit> {
        if patch.format() != "litchi-doc-body" {
            return Err(invalid_durable_patch("unsupported format"));
        }
        if patch.operations().is_empty() {
            if !patch.blobs().is_empty() {
                return Err(invalid_durable_patch(
                    "operation-free durable patch has unreferenced blobs",
                ));
            }
            return self.edit();
        }
        let expected_artifact = BlobId::of(self.bytes()).as_hex();
        let mut edit = self.edit()?;
        let mut used_blobs = BTreeSet::new();
        for operation in patch.operations() {
            if operation.preconditions.len() != 2 {
                return Err(invalid_durable_patch(
                    "body operation must have exactly two preconditions",
                ));
            }
            let artifact = operation
                .preconditions
                .get("artifact_sha256")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| invalid_durable_patch("missing artifact hash precondition"))?;
            if artifact != expected_artifact {
                return Err(Error::PatchConflict);
            }
            if operation.op == "revision-mark.set" {
                apply_durable_revision(&mut edit, operation)?;
                continue;
            }
            if operation.op == "embedded-display.set" {
                apply_durable_embedded_display(&mut edit, operation)?;
                continue;
            }
            if operation.op == "embedded-object.set" {
                apply_durable_embedded_object(
                    &mut edit,
                    operation,
                    patch.blobs(),
                    &mut used_blobs,
                )?;
                continue;
            }
            if operation.op == "picture.set" {
                apply_durable_picture(&mut edit, operation, patch.blobs(), &mut used_blobs)?;
                continue;
            }
            let target = parse_durable_target(&operation.target)?;
            match operation.op.as_str() {
                "text.set" | "body-text.set" => {
                    let expected = operation
                        .preconditions
                        .get("text")
                        .and_then(serde_json::Value::as_str)
                        .ok_or_else(|| invalid_durable_patch("missing text precondition"))?;
                    let span = resolve_target(&edit.editor, target)?;
                    if span.text != expected {
                        return Err(Error::PatchConflict);
                    }
                    let replacement = operation
                        .value
                        .as_str()
                        .ok_or_else(|| invalid_durable_patch("body text value is not a string"))?;
                    edit.replace_text(target, replacement)?;
                },
                "character-bold.set" | "body-bold.set" => {
                    apply_durable_format(&mut edit, target, CharacterProperty::Bold, operation)?;
                },
                "character-italic.set" => {
                    apply_durable_format(&mut edit, target, CharacterProperty::Italic, operation)?;
                },
                "character-underline.set" => {
                    apply_durable_format(
                        &mut edit,
                        target,
                        CharacterProperty::Underline,
                        operation,
                    )?;
                },
                _ => return Err(invalid_durable_patch("unsupported operation vocabulary")),
            }
        }
        if used_blobs.len() != patch.blobs().len() {
            return Err(invalid_durable_patch(
                "durable patch has unreferenced blobs",
            ));
        }
        Ok(edit)
    }

    /// Plans an inert text transfer while proving the receiving selector and
    /// exact target artifact. Text carries no DOC style, drawing, or resource
    /// identifiers, so its dependency closure is empty.
    pub(crate) fn plan_text_transfer_from(
        &self,
        donor: &Self,
        source: TextTarget,
        destination: TextTarget,
    ) -> Result<TransferPlan> {
        let donor_editor = donor.editor()?;
        let value = resolve_target(&donor_editor, source)?.text;
        if let Some(dependency) = drawing_dependency(&value) {
            return Err(Error::Refused(Refusal::TransferDependency {
                dependency: dependency.transfer_name(),
            }));
        }
        if has_structural_content(&value) {
            return Err(Error::Refused(Refusal::TransferDependency {
                dependency: "structural text",
            }));
        }
        let transaction_limits = self.transaction_limits;
        let mut validation = self.edit()?;
        validation.replace_text(destination, &value)?;
        Ok(TransferPlan {
            target: self.lineage(),
            destination,
            value,
            transaction_limits,
        })
    }

    /// Plans transfer of a complete inert embedded-object dependency closure.
    ///
    /// The donor's standalone object CFB and exact preview block are retained;
    /// the receiver regenerates the field text, field PLCF, CHPX picture
    /// references, Data-stream offset, and `ObjectPool` storage identity.
    pub(crate) fn plan_embedded_transfer_from(
        &self,
        donor: &Self,
        source_storage_id: u32,
        destination_storage_id: u32,
    ) -> Result<EmbeddedTransferPlan> {
        if self
            .embedded_objects()?
            .get(destination_storage_id)
            .is_some()
        {
            return Err(Error::Refused(Refusal::ResourceCollision {
                storage_id: destination_storage_id,
            }));
        }
        let donor = crate::embedded_object::Snapshot::open(donor.finish(), donor.limits)
            .map_err(Error::Invalid)?;
        if donor
            .inventory()
            .map_err(Error::Invalid)?
            .get(source_storage_id)
            .is_none()
        {
            return Err(Error::Refused(Refusal::ResourceNotFound {
                storage_id: source_storage_id,
            }));
        }
        let options = donor
            .export_for_transfer(source_storage_id, destination_storage_id)
            .map_err(Error::Invalid)?;
        let transaction_limits = self.transaction_limits;
        let mut validation = self.edit()?;
        validation.add_embedded_object(options.clone())?;
        Ok(EmbeddedTransferPlan {
            target: self.lineage(),
            options,
            transaction_limits,
        })
    }

    /// Plans transfer of one canonical inline or floating picture into a
    /// non-empty inert main-story placeholder.
    ///
    /// The bounded closure owns the marker CHPX, exact PICF/Data block, and,
    /// for floating pictures, a re-homed singleton `PlcfSpaMom` and `DggInfo`.
    /// The selected PICF/Data block is proved independently. Floating pictures
    /// additionally prove one exact SPA/top-level-shape/`pib`/`BStore` identity,
    /// so unrelated bounded main/header shapes, groups, textboxes, and shared
    /// slots may remain in the donor. Selected nested/extended graphs,
    /// delay-loaded images, auxiliary-story selection, and occupied receivers
    /// are refused.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture transfer planning is a checked binary-resource boundary"
    )]
    pub(crate) fn plan_picture_transfer_from(
        &self,
        donor: &Self,
        source: TextTarget,
        destination: TextTarget,
    ) -> Result<PictureTransferPlan> {
        if target_story(source) != Story::Main || target_story(destination) != Story::Main {
            return Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::AuxiliaryStoryPicture,
            }));
        }
        let donor_editor = donor.editor()?;
        let source_span = resolve_target(&donor_editor, source)?;
        let floating = match source_span.text.as_str() {
            "\u{0001}" => false,
            "\u{0008}" => true,
            _ => {
                return Err(Error::Refused(Refusal::DrawingDependency {
                    dependency: DrawingDependency::UnsupportedPictureGraph,
                }));
            },
        };
        let graph = donor_editor
            .picture_graph_at_cp(source_span.start_cp, floating)
            .map_err(|error| {
                let _reason = error.to_string();
                Error::Refused(Refusal::DrawingDependency {
                    dependency: DrawingDependency::UnsupportedPictureGraph,
                })
            })?;
        let transaction_limits = self.transaction_limits;
        let mut validation = self.edit()?;
        validation.install_picture(destination, graph.clone())?;
        Ok(PictureTransferPlan {
            target: self.lineage(),
            destination,
            graph,
            transaction_limits,
        })
    }

    /// Non-mutating three-way plan for two patches based on this exact source.
    ///
    /// Opaque `RgDofr` and `Selsf` auxiliary-table changes are refused because
    /// this plan merges semantic body changes and cannot combine those byte
    /// ranges without silently dropping one side.
    pub(crate) fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
        if left.before != *self || right.before != *self {
            return Err(Error::PatchConflict);
        }
        if left.has_auxiliary_changes() || right.has_auxiliary_changes() {
            return Err(Error::Refused(Refusal::AuxiliaryThreeWayMergeUnsupported));
        }
        Ok(ThreeWayPlan::new(self.clone(), left, right))
    }

    /// Exact source bytes. A snapshot has no implicit serialization step.
    #[must_use]
    pub(crate) fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    fn editor(&self) -> Result<RevisionEditor> {
        let source_limits = self.source_limits;
        crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(
            self.source.len(),
            source_limits,
        )
        .map_err(Error::Invalid)?;
        let source = self
            .candidate_source
            .as_ref()
            .map_or_else(|| self.source.as_ref().to_vec(), |source| source.to_vec());
        let (mut editor, _ole) = RevisionEditor::open_with_ole_file_with_policy_and_admission(
            source,
            self.limits,
            self.protection_policy.clone(),
            |ole| crate::package::preflight_declared_streams(ole, source_limits).map(|_| ()),
        )
        .map_err(Error::Invalid)?;
        editor.set_retained_render_limit(self.transaction_limits.max_retained_render_bytes());
        Ok(editor)
    }

    fn with_transaction_limits(&self, transaction_limits: TransactionLimits) -> Self {
        let mut snapshot = self.clone();
        snapshot.transaction_limits = transaction_limits;
        snapshot
    }

    fn with_policies(
        &self,
        limits: Limits,
        source_limits: crate::package::SourceLimits,
        transaction_limits: TransactionLimits,
        protection_policy: ProtectionPolicy,
    ) -> Self {
        let mut snapshot = self.clone();
        snapshot.limits = limits;
        snapshot.source_limits = source_limits;
        snapshot.transaction_limits = transaction_limits;
        snapshot.protection_policy = protection_policy;
        snapshot
    }

    fn lineage(&self) -> Lineage {
        Lineage(Arc::clone(&self.source))
    }
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.len())
            .field("fingerprint", &self.fingerprint())
            .field("signed", &self.signed)
            .field("limits", &self.limits)
            .field("source_limits", &self.source_limits)
            .field("transaction_limits", &self.transaction_limits)
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// Package-owned commit carrier used by the canonical `package::edit`
/// boundary.  Unlike the ordinary body [`Commit`], a changed candidate keeps
/// the already parsed CFB and semantic document projection alive so the
/// immutable package owner can adopt them without opening the container a
/// second time.  The carrier is crate-private because the body snapshot and
/// package limits remain the authority for all validation.
pub(crate) struct PackageCommit {
    pub(crate) snapshot: Snapshot,
    pub(crate) patch: Patch,
    pub(crate) validated: Option<ValidatedPackage>,
}

/// Already validated package parts transferred from one body commit.
pub(crate) struct ValidatedPackage {
    pub(crate) ole: OleFile<Cursor<Vec<u8>>>,
    pub(crate) document: crate::Document,
}

/// Clone-first staged text edit over one immutable source snapshot.
pub(crate) struct Edit {
    source: Snapshot,
    editor: RevisionEditor,
    signed: bool,
    changes: Vec<Change>,
    replacement_units: usize,
    auxiliary_changed: bool,
    auxiliary_operations: usize,
}

impl Edit {
    fn new(source: Snapshot) -> Result<Self> {
        let editor = source.editor()?;
        let signed = source.signed;
        Ok(Self {
            source,
            editor,
            signed,
            changes: Vec::new(),
            replacement_units: 0,
            auxiliary_changed: false,
            auxiliary_operations: 0,
        })
    }

    /// Reopens a complete candidate rendered by a resource owner under this
    /// transaction's protection policy and retained-render ceiling.
    fn reopen_editor(&self, bytes: Vec<u8>) -> Result<RevisionEditor> {
        let source_limits = self.source.source_limits;
        crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(bytes.len(), source_limits)
            .map_err(Error::Invalid)?;
        let (mut editor, _ole) = RevisionEditor::open_with_ole_file_with_policy_and_admission(
            bytes,
            self.source.limits,
            self.source.protection_policy.clone(),
            |ole| crate::package::preflight_declared_streams(ole, source_limits).map(|_| ()),
        )
        .map_err(Error::Invalid)?;
        editor
            .set_retained_render_limit(self.source.transaction_limits.max_retained_render_bytes());
        Ok(editor)
    }

    fn with_transaction_limits<T>(
        &mut self,
        incoming: TransactionLimits,
        operation: impl FnOnce(&mut Self) -> Result<T>,
    ) -> Result<T> {
        let previous = self.source.transaction_limits;
        let effective = previous.intersect(incoming);
        // Auxiliary-table edits count against the operation limit exactly as
        // `ensure_operation_capacity` counts them.
        let operations = self.changes.len().saturating_add(self.auxiliary_operations);
        if operations > effective.operations {
            return Err(Error::Refused(Refusal::OperationLimit {
                observed: operations,
                limit: effective.operations,
            }));
        }
        if self.replacement_units > effective.total_units {
            return Err(Error::Refused(Refusal::ReplacementLimit {
                observed: self.replacement_units,
                limit: effective.total_units,
            }));
        }
        self.source.transaction_limits = effective;
        let result = operation(self);
        if result.is_ok() {
            self.editor.set_retained_render_limit(
                self.source.transaction_limits.max_retained_render_bytes(),
            );
        } else {
            self.source.transaction_limits = previous;
        }
        result
    }

    /// Render a candidate through the lower package owner while preserving
    /// this facade's typed signed-source refusal. The lower public editor
    /// guard intentionally reports a package error for direct callers; this
    /// owner compares the exact retained bytes before publication so body
    /// callers keep [`Refusal::SignedSource`].
    fn finish_candidate(&self) -> Result<Vec<u8>> {
        let bytes = self
            .editor
            .clone()
            .finish_unchecked()
            .map_err(Error::Invalid)?;
        self.authorize_candidate(bytes)
    }

    /// Commit-time form of [`Self::finish_candidate`].
    ///
    /// A validated render that the editor retained under this transaction's
    /// ceiling is moved out instead of rendering the candidate a second time
    /// (performance change 0730); otherwise the candidate is rendered as
    /// [`Self::finish_candidate`] renders it. Both routes return the same
    /// bytes and apply the same authorization and signed-source checks.
    fn take_candidate(&mut self) -> Result<Vec<u8>> {
        let bytes = match self.editor.take_retained_render() {
            Some(rendered) => rendered,
            None => self
                .editor
                .finish_unchecked_in_place()
                .map_err(Error::Invalid)?,
        };
        self.authorize_candidate(bytes)
    }

    fn authorize_candidate(&self, bytes: Vec<u8>) -> Result<Vec<u8>> {
        self.editor
            .authorize_rendered(&bytes)
            .map_err(Error::Invalid)?;
        if self.signed {
            if self
                .source
                .candidate_source
                .as_ref()
                .is_some_and(|candidate| bytes.as_slice() == candidate.as_ref())
            {
                return Ok(self.source.bytes().to_vec());
            }
            if bytes != self.source.bytes() {
                return Err(Error::Refused(Refusal::SignedSource));
            }
        }
        Ok(bytes)
    }

    fn prepare_saved_selection_length_change(
        &self,
        start: u32,
        end: u32,
        added: usize,
    ) -> Result<()> {
        let added = u32::try_from(added).map_err(|_error| {
            Error::Invalid(PackageError::Corrupted(
                "Selsf splice length exceeds u32".to_string(),
            ))
        })?;
        match self.editor.saved_selection_splice(start, end, added) {
            Ok(Some(_)) | Ok(None) => Ok(()),
            Err(SavedSelectionSpliceError::Invalid(error)) => Err(Error::Invalid(error)),
            Err(SavedSelectionSpliceError::Ambiguous) => {
                Err(Error::Refused(Refusal::PositionDependency {
                    fib_index: FIB_INDEX_WSS,
                }))
            },
        }
    }

    /// Exact source bytes retained by this staged edit.
    #[must_use]
    pub(crate) fn source_bytes(&self) -> &[u8] {
        self.source.bytes()
    }

    /// Borrowed semantic changes staged so far, including authored text.
    #[must_use]
    pub(crate) fn changes(&self) -> impl ExactSizeIterator<Item = ChangeRef<'_>> {
        self.changes.iter().map(Change::as_ref)
    }

    /// Allocation capacity held by the already-validated rendered handoff,
    /// when the transaction's finite retention ceiling admitted it.
    #[must_use]
    pub(crate) fn retained_render_bytes(&self) -> Option<usize> {
        self.editor.retained_render_capacity()
    }

    /// Releases the retained rendered handoff without changing the staged
    /// document or any semantic transaction result. Later mutations may retain
    /// a new output under the transaction's ceiling.
    pub(crate) fn release_retained_render(&mut self) {
        self.editor.release_retained_render();
    }

    /// Applies a source-checked same-length `Selsf` patch to this staged DOC
    /// edit through the normal package owner. A changed patch must have been
    /// produced from this exact immutable snapshot; the component patch also
    /// checks the FIB-selected record before cloning the table stream.
    pub(crate) fn apply_saved_selection_patch(
        &mut self,
        patch: &SavedSelectionPatch,
    ) -> Result<bool> {
        let changed = self
            .editor
            .preflight_saved_selection_patch(patch)
            .map_err(Error::Invalid)?;
        if !changed {
            return Ok(false);
        }
        let owner = self.source.bytes_shared();
        if !patch.owner_matches(&owner) {
            return Err(Error::PatchConflict);
        }
        if self.signed {
            return Err(Error::Refused(Refusal::SignedSource));
        }
        self.ensure_operation_capacity()?;
        let changed = self
            .editor
            .apply_saved_selection_patch(patch)
            .map_err(Error::Invalid)?;
        self.auxiliary_changed |= changed;
        if changed {
            self.auxiliary_operations = self.auxiliary_operations.saturating_add(1);
        }
        Ok(changed)
    }

    /// Runs a checked same-length edit over the optional `Selsf` record.
    ///
    /// The closure receives the typed transaction and returns the package
    /// error produced by its field setters. A document without `Selsf` is
    /// reported as a missing semantic target.
    pub(crate) fn edit_saved_selection(
        &mut self,
        edit: impl FnOnce(&mut SavedSelectionEdit) -> crate::package::Result<()>,
    ) -> Result<bool> {
        let source = self
            .editor
            .saved_selection()
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::TargetNotFound))?
            .with_owner(self.source.bytes_shared());
        let mut transaction = source.edit();
        edit(&mut transaction).map_err(Error::Invalid)?;
        let commit = transaction.commit().map_err(Error::Invalid)?;
        self.apply_saved_selection_patch(commit.patch())
    }

    /// Applies a source-checked same-length `RgDofr` patch to this staged DOC
    /// edit through the normal package owner. A changed patch must have been
    /// produced from this exact immutable snapshot; the lower-level component
    /// patch methods intentionally retain their narrower `RgDofr` source
    /// checks.
    pub(crate) fn apply_dofr_patch(&mut self, patch: &DofrPatch) -> Result<bool> {
        let changed = self
            .editor
            .preflight_dofr_patch(patch)
            .map_err(Error::Invalid)?;
        if !changed {
            return Ok(false);
        }
        let owner = self.source.bytes_shared();
        if !patch.owner_matches(&owner) {
            return Err(Error::PatchConflict);
        }
        if self.signed {
            return Err(Error::Refused(Refusal::SignedSource));
        }
        self.ensure_operation_capacity()?;
        let changed = self
            .editor
            .apply_dofr_patch(patch)
            .map_err(Error::Invalid)?;
        self.auxiliary_changed |= changed;
        if changed {
            self.auxiliary_operations = self.auxiliary_operations.saturating_add(1);
        }
        Ok(changed)
    }

    /// Replaces one complete, validated `RgDofr` record in this staged edit.
    /// Record size and frame-set/list sequence invariants are checked before
    /// the candidate table stream is published.
    pub(crate) fn replace_dofr_record(&mut self, index: usize, replacement: &[u8]) -> Result<bool> {
        let source = self
            .editor
            .dofr_records()
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::TargetNotFound))?
            .with_owner(self.source.bytes_shared());
        let changed = source
            .preflight_record_replacement(index, replacement)
            .map_err(Error::Invalid)?;
        if !changed {
            return Ok(false);
        }
        if self.signed {
            return Err(Error::Refused(Refusal::SignedSource));
        }
        self.ensure_operation_capacity()?;
        let mut transaction = source.edit();
        transaction
            .replace_record_bytes(index, replacement)
            .map_err(Error::Invalid)?;
        let commit = transaction.commit().map_err(Error::Invalid)?;
        self.apply_dofr_patch(commit.patch())
    }

    /// Replaces text in one ordinary source-body paragraph.
    ///
    /// The replacement can change UTF-16 length. Publication appends a Unicode
    /// piece, rebuilds CHPX FKPs and the CLX, shifts modeled main-story PLCFs,
    /// and updates the FIB story length.
    ///
    /// `position` is a format-neutral [`Position`]; its membership in this
    /// source body is checked here and an absent paragraph is reported as
    /// [`Refusal::ParagraphNotFound`].
    ///
    /// # Errors
    ///
    /// Returns a typed [`Refusal`] for every operation outside the proven
    /// bounded dependency closure and [`Error::Invalid`] for a failed package
    /// update.
    pub(crate) fn replace_paragraph(
        &mut self,
        position: Position,
        replacement: &str,
    ) -> Result<()> {
        self.replace_text(TextTarget::body_paragraph(position), replacement)
            .map_err(|error| match error {
                Error::Refused(Refusal::TargetNotFound) => {
                    Error::Refused(Refusal::ParagraphNotFound)
                },
                other => other,
            })
    }

    /// Replaces one checked ordinary paragraph, simple table-cell value, or
    /// simple field cached result. Main-story targets may change length when
    /// every `Selsf` CP has a deterministic splice-boundary mapping; a `Selsf`
    /// CP strictly inside the replaced text is refused. Non-main story
    /// paragraphs require equal UTF-16 length.
    pub(crate) fn replace_text(&mut self, target: TextTarget, replacement: &str) -> Result<()> {
        let span = resolve_target(&self.editor, target)?;
        if let Some(dependency) = drawing_dependency(&span.text) {
            return Err(Error::Refused(Refusal::DrawingDependency { dependency }));
        }
        if has_structural_content(&span.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if has_structural_content(replacement) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        // Compare before charging the replacement. An exact semantic no-op
        // must remain admissible even when the caller's quota is tighter than
        // the already-stored value (and an empty paragraph must be able to
        // accept an empty no-op).
        if span.text == replacement {
            return Ok(());
        }
        if let Some(fib_index) = self
            .editor
            .interior_range_boundary(span.start_cp, span.end_cp)
        {
            return Err(Error::Refused(Refusal::RangeBoundary { fib_index }));
        }
        let actual = replacement.encode_utf16().count();
        self.ensure_replacement_capacity(actual)?;
        if span.start_cp == span.end_cp
            && !matches!(
                target,
                TextTarget::StoryParagraph {
                    story: Story::Main,
                    ..
                }
            )
        {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        let dependency_end = if span.start_cp == span.end_cp {
            span.start_cp.checked_add(1).ok_or_else(cp_overflow)?
        } else {
            span.end_cp
        };
        if text_revision_intersects(&self.editor, span.start_cp, dependency_end)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        let expected = span.text.encode_utf16().count();
        let story = target_story(target);
        if story != Story::Main && actual != expected {
            return Err(Error::Refused(Refusal::StoryLengthChange {
                story,
                expected,
                actual,
            }));
        }
        if story != Story::Main && !self.editor.is_unicode_range(span.start_cp, span.end_cp) {
            return Err(Error::Refused(Refusal::CompressedPiece));
        }
        if story == Story::Main && actual != expected {
            self.prepare_saved_selection_length_change(span.start_cp, span.end_cp, actual)?;
        }
        if actual != expected
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()?;
        // An empty ordinary paragraph has no text run, but its paragraph mark
        // still carries the direct character formatting that must be copied to
        // newly inserted text. Include that one mark in the formatting proof;
        // the replacement itself remains inserted before the mark.
        let formatting_end = if span.start_cp == span.end_cp {
            span.start_cp.checked_add(1).ok_or_else(cp_overflow)?
        } else {
            span.end_cp
        };
        if !self
            .editor
            .has_uniform_character_format(span.start_cp, formatting_end)
            .map_err(Error::Invalid)?
        {
            return Err(Error::Refused(Refusal::FormattingDependency));
        }
        let before = span.text;
        if story == Story::Main {
            self.editor
                .replace_plain_text(span.start_cp, span.end_cp, replacement)
                .map_err(Error::Invalid)?;
        } else {
            self.editor
                .replace_unicode_text_same_length(span.start_cp, span.end_cp, replacement)
                .map_err(Error::Invalid)?;
        }
        self.replacement_units = self.replacement_units.saturating_add(actual);
        self.changes.push(Change::Text {
            target,
            before,
            after: replacement.to_string(),
        });
        Ok(())
    }

    /// Sets a direct bold override for one ordinary body paragraph in the
    /// same transaction as text replacements.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the paragraph is absent, empty,
    /// structural, tracked, or has non-uniform direct bold semantics.
    pub(crate) fn set_paragraph_bold(&mut self, position: Position, enabled: bool) -> Result<()> {
        self.set_character_property(
            TextTarget::body_paragraph(position),
            CharacterProperty::Bold,
            enabled,
        )
    }

    /// Sets a direct bold, italic, or single-underline override on one semantic
    /// text target without creating or rewriting style-table resources.
    pub(crate) fn set_character_property(
        &mut self,
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    ) -> Result<()> {
        self.set_character_property_override(target, property, Some(enabled))
    }

    /// Sets a valued direct character property (size, color or font) on one
    /// semantic text target without creating or rewriting style-table or
    /// font-table resources. Setting the value the target already carries
    /// directly is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the target is absent, empty, structural
    /// or tracked, when its direct value is not uniform, when the size is out
    /// of range, or when the font table does not name the font.
    pub(crate) fn set_character_value(
        &mut self,
        target: TextTarget,
        value: CharacterValue<'_>,
    ) -> Result<()> {
        let fragment = self.character_value_fragment(value)?;
        self.set_character_fragment(target, value.property(), fragment)
    }

    fn character_value_fragment(&self, value: CharacterValue<'_>) -> Result<Vec<u8>> {
        let mut fragment = Vec::new();
        match value {
            CharacterValue::Size(half_points) => {
                if !(2..=3276).contains(&half_points) {
                    return Err(Error::Refused(Refusal::InvalidFontSize));
                }
                fragment.extend_from_slice(&SPRM_C_HPS.to_le_bytes());
                fragment.extend_from_slice(&half_points.to_le_bytes());
            },
            CharacterValue::Color(None) => {},
            CharacterValue::Color(Some(rgb)) => {
                if let Some(index) = ICO_PALETTE.iter().position(|color| *color == rgb) {
                    fragment.extend_from_slice(&SPRM_C_ICO.to_le_bytes());
                    fragment.push(u8::try_from(index + 1).unwrap_or(0));
                }
                fragment.extend_from_slice(&SPRM_C_CV.to_le_bytes());
                fragment.extend_from_slice(&[rgb[0], rgb[1], rgb[2], 0]);
            },
            CharacterValue::Font(name) => {
                let ftc = self.font_index(name)?;
                fragment.extend_from_slice(&SPRM_C_RG_FTC0.to_le_bytes());
                fragment.extend_from_slice(&ftc.to_le_bytes());
            },
            CharacterValue::Style(istd) => {
                let sheet = self.editor.stylesheet().map_err(Error::Invalid)?;
                if sheet
                    .get(istd)
                    .is_none_or(|style| style.kind != crate::parts::styles::StyleKind::Character)
                {
                    return Err(Error::Refused(Refusal::StyleNotFound));
                }
                if istd != 10 {
                    fragment.extend_from_slice(&SPRM_C_ISTD_VALUE.to_le_bytes());
                    fragment.extend_from_slice(&istd.to_le_bytes());
                }
            },
        }
        Ok(fragment)
    }

    /// The font-table index of the font named exactly `name`.
    fn font_index(&self, name: &str) -> Result<u16> {
        let fonts = self
            .editor
            .table_slice(FIB_INDEX_STTBF_FFN)
            .map_err(Error::Invalid)?
            .filter(|bytes| !bytes.is_empty())
            .map(crate::parts::fonts::FontTable::parse)
            .transpose()
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::FontNotFound))?;
        fonts
            .iter()
            .position(|font| font.name() == name)
            .and_then(|index| u16::try_from(index).ok())
            .ok_or(Error::Refused(Refusal::FontNotFound))
    }

    fn set_character_fragment(
        &mut self,
        target: TextTarget,
        property: CharacterValueProperty,
        fragment: Vec<u8>,
    ) -> Result<()> {
        let span = resolve_target(&self.editor, target)?;
        if span.start_cp == span.end_cp {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        if let Some(dependency) = drawing_dependency(&span.text) {
            return Err(Error::Refused(Refusal::DrawingDependency { dependency }));
        }
        if has_structural_content(&span.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if text_revision_intersects(&self.editor, span.start_cp, span.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        let before = self
            .editor
            .uniform_character_sprms(span.start_cp, span.end_cp, property.opcodes())
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::FormattingDependency))?;
        if before == fragment {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        self.editor
            .set_character_sprms(span.start_cp, span.end_cp, property.opcodes(), &fragment)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::CharacterValue {
            target,
            property,
            before,
            after: fragment,
        });
        Ok(())
    }

    /// Inserts one ordinary main-story paragraph holding `text` before the
    /// paragraph at `position`, or after the last one when `position` is the
    /// paragraph count. The new paragraph copies the paragraph formatting of
    /// that neighbour and the character formatting of its paragraph mark.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a position past the end, structural text,
    /// a tracked neighbour mark, unmodeled CP-indexed tables, or a limit.
    pub(crate) fn insert_paragraph(&mut self, position: Position, text: &str) -> Result<()> {
        if has_structural_content(text) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        let paragraphs = source_paragraphs(&self.editor)?;
        let template = paragraphs
            .get(position.get())
            .or_else(|| {
                (position.get() == paragraphs.len())
                    .then(|| paragraphs.last())
                    .flatten()
            })
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        let mut units = text.encode_utf16().collect::<Vec<_>>();
        self.ensure_replacement_capacity(units.len())?;
        self.ensure_structural_capacity()?;
        let mark = template.end_cp;
        let mark_end = mark.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, mark, mark_end)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        units.push(0x000D);
        let copy = self
            .editor
            .template_paragraph(mark, units)
            .map_err(Error::Invalid)?;
        let at = match paragraphs.get(position.get()) {
            Some(paragraph) => paragraph.start_cp,
            None => mark_end,
        };
        self.editor
            .insert_paragraph_copy(at, &copy)
            .map_err(Error::Invalid)?;
        self.replacement_units = self.replacement_units.saturating_add(copy.len());
        self.changes.push(Change::Paragraph {
            position,
            before: None,
            after: Some(copy),
        });
        Ok(())
    }

    /// Removes the ordinary main-story paragraph at `position`, its mark
    /// included.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the paragraph is absent or the only one,
    /// holds structural or tracked content or an anchored range, or when its
    /// removal would join two tables or end the story in a table.
    pub(crate) fn remove_paragraph(&mut self, position: Position) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        if paragraphs.get(position.get()).is_none() {
            return Err(Error::Refused(Refusal::ParagraphNotFound));
        }
        if paragraphs.len() < 2 {
            return Err(Error::Refused(Refusal::SoleParagraph));
        }
        self.ensure_structural_capacity()?;
        let copy = self.detach_paragraph(&paragraphs, position)?;
        self.changes.push(Change::Paragraph {
            position,
            before: Some(copy),
            after: None,
        });
        Ok(())
    }

    /// Moves the ordinary main-story paragraph at `from` so that it ends at
    /// position `to`, counted after it is detached, with its exact text,
    /// character runs and paragraph formatting.
    ///
    /// # Errors
    ///
    /// Returns the refusals of [`Self::remove_paragraph`] for the moved
    /// paragraph, and [`Refusal::ParagraphNotFound`] for a destination past
    /// the end.
    pub(crate) fn move_paragraph(&mut self, from: Position, to: Position) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        if from.get() >= paragraphs.len() || to.get() >= paragraphs.len() {
            return Err(Error::Refused(Refusal::ParagraphNotFound));
        }
        if from == to {
            return Ok(());
        }
        self.ensure_structural_capacity()?;
        let candidate = self.editor.clone();
        let copy = self.detach_paragraph(&paragraphs, from)?;
        let result = (|| {
            let remaining = source_paragraphs(&self.editor)?;
            let at = match remaining.get(to.get()) {
                Some(paragraph) => paragraph.start_cp,
                None => remaining
                    .last()
                    .ok_or(Error::Refused(Refusal::ParagraphNotFound))?
                    .end_cp
                    .checked_add(1)
                    .ok_or_else(cp_overflow)?,
            };
            self.editor
                .insert_paragraph_copy(at, &copy)
                .map_err(Error::Invalid)
        })();
        if let Err(error) = result {
            self.editor = candidate;
            return Err(error);
        }
        self.changes.push(Change::ParagraphMove { from, to });
        Ok(())
    }

    /// Capture and delete one ordinary paragraph after the structural
    /// admission checks.
    fn detach_paragraph(
        &mut self,
        paragraphs: &[SourceParagraph],
        position: Position,
    ) -> Result<ParagraphCopy> {
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        if let Some(dependency) = drawing_dependency(&paragraph.text) {
            return Err(Error::Refused(Refusal::DrawingDependency { dependency }));
        }
        if has_structural_content(&paragraph.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        let start = paragraph.start_cp;
        let end = paragraph.end_cp.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, start, end)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        if let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first() {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        if self.editor.has_cp_dependency_inside(start, end) {
            return Err(Error::Refused(Refusal::AnchoredContent));
        }
        let story_end = self.editor.main_story_cp_len();
        let after_table = match start.checked_sub(1) {
            Some(previous) => self
                .editor
                .is_in_table_at_cp(previous)
                .map_err(Error::Invalid)?,
            None => false,
        };
        let before_table =
            end < story_end && self.editor.is_in_table_at_cp(end).map_err(Error::Invalid)?;
        if after_table && (before_table || end >= story_end) {
            return Err(Error::Refused(Refusal::TableBoundary));
        }
        let copy = self
            .editor
            .capture_paragraph(start, end)
            .map_err(|_error| Error::Refused(Refusal::FormattingDependency))?;
        self.editor
            .delete_paragraph_range(start, end)
            .map_err(|_error| Error::Refused(Refusal::AnchoredContent))?;
        Ok(copy)
    }

    /// Re-insert an exact captured paragraph at `position` (replay).
    fn insert_paragraph_copy(&mut self, position: Position, copy: &ParagraphCopy) -> Result<()> {
        self.ensure_structural_capacity()?;
        let paragraphs = source_paragraphs(&self.editor)?;
        let at = match paragraphs.get(position.get()) {
            Some(paragraph) => paragraph.start_cp,
            None if position.get() == paragraphs.len() => paragraphs
                .last()
                .ok_or(Error::Refused(Refusal::ParagraphNotFound))?
                .end_cp
                .checked_add(1)
                .ok_or_else(cp_overflow)?,
            None => return Err(Error::Refused(Refusal::ParagraphNotFound)),
        };
        self.editor
            .insert_paragraph_copy(at, copy)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Paragraph {
            position,
            before: None,
            after: Some(copy.clone()),
        });
        Ok(())
    }

    fn ensure_structural_capacity(&self) -> Result<()> {
        if let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first() {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()
    }

    /// Applies a dependency-checked text transfer prepared for this exact
    /// receiving artifact.
    pub(crate) fn apply_transfer(&mut self, plan: &TransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::PatchConflict);
        }
        self.with_transaction_limits(plan.transaction_limits, |edit| {
            edit.replace_text(plan.destination, &plan.value)
        })
    }

    /// Applies a dependency-closed embedded-object transfer prepared for this
    /// exact receiving artifact.
    pub(crate) fn apply_embedded_transfer(&mut self, plan: &EmbeddedTransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::PatchConflict);
        }
        self.with_transaction_limits(plan.transaction_limits, |edit| {
            edit.add_embedded_object(plan.options.clone())
        })
    }

    /// Applies a canonical picture graph prepared for this exact receiver.
    pub(crate) fn apply_picture_transfer(&mut self, plan: &PictureTransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::PatchConflict);
        }
        self.with_transaction_limits(plan.transaction_limits, |edit| {
            edit.install_picture(plan.destination, plan.graph.clone())
        })
    }

    fn install_picture(
        &mut self,
        target: TextTarget,
        graph: crate::tracked_revision::PictureGraph,
    ) -> Result<()> {
        if target_story(target) != Story::Main {
            return Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::AuxiliaryStoryPicture,
            }));
        }
        let span = resolve_target(&self.editor, target)?;
        if drawing_dependency(&span.text).is_some() {
            return Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::PictureGraphCollision,
            }));
        }
        if span.start_cp == span.end_cp || has_structural_content(&span.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if text_revision_intersects(&self.editor, span.start_cp, span.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        if !self
            .editor
            .has_empty_picture_graph()
            .map_err(Error::Invalid)?
        {
            return Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::PictureGraphCollision,
            }));
        }
        let actual = 1;
        self.ensure_replacement_capacity(actual)?;
        if !self
            .editor
            .has_uniform_character_format(span.start_cp, span.end_cp)
            .map_err(Error::Invalid)?
        {
            return Err(Error::Refused(Refusal::FormattingDependency));
        }
        let expected = span.text.encode_utf16().count();
        if expected != actual {
            self.prepare_saved_selection_length_change(span.start_cp, span.end_cp, actual)?;
        }
        if expected != actual
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()?;
        let before = PictureSlot::Text(span.text);
        let installed = self
            .editor
            .replace_with_picture_graph(span.start_cp, span.end_cp, &graph)
            .map_err(Error::Invalid)?
            .ok_or(Error::PatchConflict)?;
        self.replacement_units = self.replacement_units.saturating_add(actual);
        self.changes.push(Change::Picture {
            target,
            before,
            after: PictureSlot::Graph(installed),
        });
        Ok(())
    }

    fn restore_picture_text(
        &mut self,
        target: TextTarget,
        graph: crate::tracked_revision::PictureGraph,
        replacement: String,
    ) -> Result<()> {
        if target_story(target) != Story::Main {
            return Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::AuxiliaryStoryPicture,
            }));
        }
        let span = resolve_target(&self.editor, target)?;
        let expected_marker = if graph.floating {
            "\u{0008}"
        } else {
            "\u{0001}"
        };
        if span.text != expected_marker {
            return Err(Error::PatchConflict);
        }
        if has_structural_content(&replacement) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        let actual = replacement.encode_utf16().count();
        self.ensure_replacement_capacity(actual)?;
        if actual != 1 {
            self.prepare_saved_selection_length_change(span.start_cp, span.end_cp, actual)?;
        }
        if actual != 1
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()?;
        self.editor
            .replace_picture_graph_with_text(span.start_cp, span.end_cp, &graph, &replacement)
            .map_err(Error::Invalid)?;
        self.replacement_units = self.replacement_units.saturating_add(actual);
        self.changes.push(Change::Picture {
            target,
            before: PictureSlot::Graph(graph),
            after: PictureSlot::Text(replacement),
        });
        Ok(())
    }

    /// Adds one inert embedded object through the dedicated field,
    /// `ObjectPool`, Data-stream, CHPX, CLX, and field-PLCF owner.
    pub(crate) fn add_embedded_object(
        &mut self,
        mut options: crate::embedded_object::WriteOptions,
    ) -> Result<()> {
        let storage_id = options.storage_id;
        options.instruction = format!(" EMBED LITCHI_OBJECT _{storage_id} ");
        let bytes = self.finish_candidate()?;
        let snapshot = crate::embedded_object::Snapshot::open_with_policy(
            bytes,
            self.source.limits,
            self.source.protection_policy.clone(),
        )
        .map_err(Error::Invalid)?;
        if snapshot
            .inventory()
            .map_err(Error::Invalid)?
            .get(storage_id)
            .is_some()
        {
            return Err(Error::Refused(Refusal::ResourceCollision { storage_id }));
        }
        self.ensure_operation_capacity()?;
        let mut transaction = snapshot.edit();
        transaction
            .add(options.clone())
            .map_err(|error| Error::Invalid(error.into()))?;
        let bytes = transaction
            .commit()
            .map_err(|error| Error::Invalid(error.into()))?
            .snapshot()
            .finish();
        self.editor = self.reopen_editor(bytes)?;
        self.changes.push(Change::EmbeddedObject {
            storage_id,
            before: None,
            after: Some(options),
        });
        Ok(())
    }

    /// Removes one managed embedded field together with its preview and owning
    /// `ObjectPool` storage. The exact dependency closure is retained in the
    /// reversible patch.
    pub(crate) fn remove_embedded_object(&mut self, storage_id: u32) -> Result<()> {
        let bytes = self.finish_candidate()?;
        let snapshot = crate::embedded_object::Snapshot::open_with_policy(
            bytes,
            self.source.limits,
            self.source.protection_policy.clone(),
        )
        .map_err(Error::Invalid)?;
        if snapshot
            .inventory()
            .map_err(Error::Invalid)?
            .get(storage_id)
            .is_none()
        {
            return Err(Error::Refused(Refusal::ResourceNotFound { storage_id }));
        }
        let before = snapshot
            .export_for_transfer(storage_id, storage_id)
            .map_err(Error::Invalid)?;
        self.ensure_operation_capacity()?;
        let mut transaction = snapshot.edit();
        transaction
            .remove(storage_id)
            .map_err(|error| Error::Invalid(error.into()))?;
        let bytes = transaction
            .commit()
            .map_err(|error| Error::Invalid(error.into()))?
            .snapshot()
            .finish();
        self.editor = self.reopen_editor(bytes)?;
        self.changes.push(Change::EmbeddedObject {
            storage_id,
            before: Some(before),
            after: None,
        });
        Ok(())
    }

    /// Accepts an insertion/move-to mark or rejects a deletion/move-from mark
    /// while retaining the stored text. Destructive dispositions and property
    /// revision restoration are explicitly refused.
    pub(crate) fn dispose_revision(
        &mut self,
        position: Position,
        disposition: RevisionDisposition,
    ) -> Result<()> {
        let revisions = self.editor.revisions().map_err(Error::Invalid)?;
        let revision = revisions
            .get(position.get())
            .cloned()
            .ok_or(Error::Refused(Refusal::TargetNotFound))?;
        let safe = matches!(
            (disposition, revision.kind),
            (
                RevisionDisposition::Accept,
                RevisionKind::Insertion | RevisionKind::MoveTo
            ) | (
                RevisionDisposition::Reject,
                RevisionKind::Deletion | RevisionKind::MoveFrom
            )
        );
        if !safe {
            return Err(Error::Refused(Refusal::DestructiveRevisionDisposition {
                kind: revision.kind,
            }));
        }
        self.ensure_operation_capacity()?;
        let before = RevisionSpec::from_revision(&revision);
        self.editor.remove(position.get()).map_err(Error::Invalid)?;
        self.changes.push(Change::RevisionMark {
            identity: before.identity(),
            before: Some(before),
            after: None,
        });
        Ok(())
    }

    /// Changes the passive display-as-icon bit of one existing managed
    /// embedded object through the dedicated `ObjectPool` owner, then reopens the
    /// complete candidate into this root transaction.
    pub(crate) fn set_embedded_display_as_icon(
        &mut self,
        storage_id: u32,
        enabled: bool,
    ) -> Result<()> {
        let bytes = self.finish_candidate()?;
        let snapshot = crate::embedded_object::Snapshot::open_with_policy(
            bytes,
            self.source.limits,
            self.source.protection_policy.clone(),
        )
        .map_err(Error::Invalid)?;
        let inventory = snapshot.inventory().map_err(Error::Invalid)?;
        let before = inventory
            .get(storage_id)
            .and_then(|entry| entry.metadata().obj_info())
            .map(|info| info.display_as_icon)
            .ok_or(Error::Refused(Refusal::TargetNotFound))?;
        if before == enabled {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        let mut transaction = snapshot.edit();
        transaction
            .update_info(storage_id, |info| info.display_as_icon = enabled)
            .map_err(|error| Error::Invalid(error.into()))?;
        let bytes = transaction
            .commit()
            .map_err(|error| Error::Invalid(error.into()))?
            .snapshot()
            .finish();
        self.editor = self.reopen_editor(bytes)?;
        self.changes.push(Change::EmbeddedDisplay {
            storage_id,
            before,
            after: enabled,
        });
        Ok(())
    }

    /// Publishes a validated snapshot and its reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the rendered candidate cannot be
    /// reopened with the original safety limits, or
    /// `Error::Refused(Refusal::SignedSource)` when a changed candidate
    /// would leave retained signature metadata stale.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the engine directly; Package is the public route"
        )
    )]
    pub(crate) fn commit(mut self) -> Result<Commit> {
        let bytes = self.take_candidate()?;
        let snapshot = if bytes == self.source.bytes() {
            self.source.clone()
        } else {
            self.editor
                .authorize_rendered(&bytes)
                .map_err(map_protection_error)?;
            Snapshot::open_with_options(
                bytes,
                OpenOptions::default()
                    .source_limits(self.source.source_limits)
                    .revision_limits(self.source.limits)
                    .transaction_limits(self.source.transaction_limits)
                    .protection_policy(self.source.protection_policy.clone()),
            )?
        };
        let patch = Patch::new(
            self.source,
            snapshot.clone(),
            self.changes,
            self.auxiliary_changed,
        );
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(snapshot, patch, diagnostics))
    }

    /// Publishes a package-owned commit while transferring the validated
    /// changed candidate's CFB and semantic projection to the package owner.
    ///
    /// The changed path performs the same ordered gates as the ordinary body
    /// commit: finish and protection authorization, strict revision-owner
    /// reopening, and complete public `Document` validation.  It retains the
    /// resulting `OleFile` and `Document` instead of reopening the candidate
    /// from bytes in `package::Package`.  Exact no-ops retain the original
    /// snapshot and carry no candidate owner.
    pub(crate) fn commit_with_package(mut self) -> Result<PackageCommit> {
        let bytes = self.take_candidate()?;
        let (snapshot, validated) = if bytes == self.source.bytes() {
            (self.source.clone(), None)
        } else {
            self.editor
                .authorize_rendered(&bytes)
                .map_err(map_protection_error)?;
            crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(
                bytes.len(),
                self.source.source_limits,
            )
            .map_err(Error::Invalid)?;
            let source_limits = self.source.source_limits;
            let (_strict_editor, mut ole) =
                RevisionEditor::open_with_ole_file_with_policy_and_admission(
                    bytes.clone(),
                    self.source.limits,
                    self.source.protection_policy.clone(),
                    |ole| {
                        crate::package::preflight_declared_streams(ole, source_limits).map(|_| ())
                    },
                )
                .map_err(Error::Invalid)?;
            let document = crate::Document::from_ole_with_options(
                &mut ole,
                PackageOpenOptions::default().source_limits(self.source.source_limits),
            )
            .map_err(Error::Invalid)?;
            let snapshot = Snapshot {
                source: Arc::from(bytes.into_boxed_slice()),
                candidate_source: None,
                signed: false,
                fingerprint_cache: OnceLock::new(),
                limits: self.source.limits,
                source_limits: self.source.source_limits,
                transaction_limits: self.source.transaction_limits,
                protection_policy: self.source.protection_policy.clone(),
            };
            (snapshot, Some(ValidatedPackage { ole, document }))
        };
        let patch = Patch::new(
            self.source,
            snapshot.clone(),
            self.changes,
            self.auxiliary_changed,
        );
        Ok(PackageCommit {
            snapshot,
            patch,
            validated,
        })
    }

    /// Publishes a profiled commit with separate semantic and in-memory CFB
    /// observers.
    ///
    /// The semantic observer receives the same fixed phase stream as
    /// [`Self::commit_profiled`]. The separate CFB observer receives the
    /// operation-local parse events emitted while a changed candidate is
    /// reopened. Exact no-op commits do not reopen a CFB and therefore emit
    /// no new CFB events during that commit call; earlier edit staging is
    /// outside this observer window.
    ///
    /// # Panics
    ///
    /// Observer panics propagate. Event balancing is guaranteed only when
    /// both observers return normally from every callback.
    ///
    /// # Errors
    ///
    /// Returns the same error as [`Self::commit`], including
    /// `Error::Refused(Refusal::SignedSource)`. Failed phases and physical
    /// parses close their respective observer events before the error is
    /// returned. Auxiliary `RgDofr` and `Selsf` changes remain outside the
    /// semantic three-way merge boundary and are refused by
    /// [`Snapshot::plan_three_way`].
    #[cfg(feature = "performance-diagnostics")]
    pub(crate) fn commit_profiled_with_cfb_observer(
        mut self,
        observer: impl FnMut(DiagnosticEvent),
        mut cfb_observer: impl FnMut(CfbParseEvent),
    ) -> Result<Commit> {
        let mut observer = observer;
        let signed = self.signed;
        let bytes = observe_phase(&mut observer, DiagnosticPhase::Finish, || {
            self.take_candidate()
        })?;
        if bytes != self.source.bytes() {
            if signed {
                return Err(Error::Refused(Refusal::SignedSource));
            }
            self.editor
                .authorize_rendered(&bytes)
                .map_err(map_protection_error)?;
        }
        let snapshot = if bytes == self.source.bytes() {
            observe_phase(&mut observer, DiagnosticPhase::ExactNoOp, || {
                Ok(self.source.clone())
            })?
        } else {
            Snapshot::open_bounded_profiled_with_source_limits(
                bytes,
                self.source.limits,
                self.source.transaction_limits,
                self.source.source_limits,
                self.source.protection_policy.clone(),
                &mut observer,
                &mut cfb_observer,
            )
            .map(|mut snapshot| {
                snapshot.protection_policy = self.source.protection_policy.clone();
                snapshot
            })?
        };
        let patch = observe_phase(&mut observer, DiagnosticPhase::Patch, || {
            Ok(Patch::new(
                self.source,
                snapshot.clone(),
                self.changes,
                self.auxiliary_changed,
            ))
        })?;
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(snapshot, patch, diagnostics))
    }

    fn ensure_operation_capacity(&self) -> Result<()> {
        let observed = self
            .changes
            .len()
            .saturating_add(self.auxiliary_operations)
            .saturating_add(1);
        let limit = self.source.transaction_limits.operations;
        if observed > limit {
            Err(Error::Refused(Refusal::OperationLimit { observed, limit }))
        } else {
            Ok(())
        }
    }

    fn ensure_replacement_capacity(&self, units: usize) -> Result<()> {
        let per_value = self.source.transaction_limits.replacement_units;
        if units > per_value {
            return Err(Error::Refused(Refusal::ReplacementLimit {
                observed: units,
                limit: per_value,
            }));
        }
        let total = self.replacement_units.saturating_add(units);
        let total_limit = self.source.transaction_limits.total_units;
        if total > total_limit {
            return Err(Error::Refused(Refusal::ReplacementLimit {
                observed: total,
                limit: total_limit,
            }));
        }
        Ok(())
    }

    fn set_character_property_override(
        &mut self,
        target: TextTarget,
        property: CharacterProperty,
        value: Option<bool>,
    ) -> Result<()> {
        let span = resolve_target(&self.editor, target)?;
        if span.start_cp == span.end_cp {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        if let Some(dependency) = drawing_dependency(&span.text) {
            return Err(Error::Refused(Refusal::DrawingDependency { dependency }));
        }
        if has_structural_content(&span.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if text_revision_intersects(&self.editor, span.start_cp, span.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        let before = uniform_property_override(&self.editor, span.start_cp, span.end_cp, property)
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::FormattingDependency))?;
        if before == value {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        set_property_override(
            &mut self.editor,
            span.start_cp,
            span.end_cp,
            property,
            value,
        )
        .map_err(Error::Invalid)?;
        self.changes.push(Change::Format {
            target,
            property,
            before,
            after: value,
        });
        Ok(())
    }
}

/// Validated commit result for one body-text edit: the published snapshot,
/// its reversible exact-source patch and diagnostics.
pub(crate) type Commit = litchi_core::edit::Commit<Snapshot, Patch>;

/// Content-free diagnostics of one body-text commit.
pub(crate) use litchi_core::edit::Diagnostics;
/// In-memory reversible replacement guarded by exact source bytes.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct Patch {
    before: Snapshot,
    after: Snapshot,
    changes: Vec<Change>,
    auxiliary_changed: bool,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before)
            .field("after", &self.after)
            .field("change_count", &self.changes.len())
            .field("auxiliary_changed", &self.auxiliary_changed)
            .finish()
    }
}

impl Patch {
    fn new(
        before: Snapshot,
        after: Snapshot,
        changes: Vec<Change>,
        auxiliary_changed: bool,
    ) -> Self {
        Self {
            before,
            after,
            changes,
            auxiliary_changed,
        }
    }

    /// Semantic text and formatting changes in application order.
    #[must_use]
    pub(crate) fn changes(&self) -> impl ExactSizeIterator<Item = ChangeRef<'_>> {
        self.changes.iter().map(Change::as_ref)
    }

    /// Whether this in-memory patch changed a supported inert auxiliary table.
    ///
    /// Auxiliary bytes remain fully covered by exact-source apply/inverse
    /// operations; the durable semantic envelope intentionally refuses them
    /// until it has a dedicated wire vocabulary.
    #[must_use]
    pub(crate) const fn has_auxiliary_changes(&self) -> bool {
        self.auxiliary_changed
    }

    /// Exact source snapshot required for application.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot produced by the transaction.
    #[must_use]
    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Fast stale-source precheck; exact bytes remain authoritative.
    ///
    /// This accessor is intentionally non-`const`: the first call may lazily
    /// populate the snapshot fingerprint cache. The former accessor
    /// was `const`, but no public constructor could produce a `Patch` in a
    /// const context.
    #[must_use]
    pub(crate) fn source_fingerprint(&self) -> u64 {
        self.before.fingerprint()
    }

    /// Target diagnostic fingerprint.
    ///
    /// See [`Self::source_fingerprint`] for the lazy-cache and const-accessor
    /// compatibility note.
    #[must_use]
    pub(crate) fn target_fingerprint(&self) -> u64 {
        self.after.fingerprint()
    }

    /// Whether this patch preserves the exact artifact.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        same_source_allocation(&self.before.source, &self.after.source)
            || self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless `source` has byte-for-byte equality
    /// with this patch's captured source snapshot.
    pub(crate) fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        let limits = intersect_revision_limits(
            intersect_revision_limits(source.limits, self.before.limits),
            self.after.limits,
        );
        let transaction_limits = source
            .transaction_limits
            .intersect(self.before.transaction_limits)
            .intersect(self.after.transaction_limits);
        let source_limits = source
            .source_limits
            .constrained_by(self.before.source_limits)
            .constrained_by(self.after.source_limits);
        // The same-allocation shortcut skips only the byte comparison; a
        // changed patch is still authorized under the source's protection
        // policy before its target is published.
        let same_allocation = same_source_allocation(&source.source, &self.before.source);
        if !same_allocation
            && (source.fingerprint() != self.before.fingerprint()
                || source.bytes() != self.before.bytes())
        {
            return Err(Error::PatchConflict);
        }
        if self.is_noop() {
            return Ok(source.with_policies(
                limits,
                source_limits,
                transaction_limits,
                source.protection_policy.clone(),
            ));
        }
        source
            .editor()?
            .authorize_rendered(self.after.bytes())
            .map_err(map_protection_error)?;
        if self.after.source_limits == source_limits
            && self.after.limits == limits
            && self.after.protection_policy == source.protection_policy
        {
            return Ok(self.after.with_policies(
                limits,
                source_limits,
                transaction_limits,
                source.protection_policy.clone(),
            ));
        }
        Snapshot::open_with_options(
            self.after.bytes().to_vec(),
            OpenOptions::default()
                .source_limits(source_limits)
                .revision_limits(limits)
                .transaction_limits(transaction_limits)
                .protection_policy(source.protection_policy.clone()),
        )
    }

    /// Exact inverse patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
            auxiliary_changed: self.auxiliary_changed,
        }
    }

    /// Converts this exact-source patch to the shared bounded deterministic
    /// semantic envelope.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError`] when the requested wire limits cannot represent
    /// every semantic operation and inverse. In-memory auxiliary-table edits
    /// are intentionally refused until the durable envelope has a matching
    /// binary-table vocabulary.
    pub(crate) fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> std::result::Result<litchi_core::patch::Patch<Reversible>, PatchError> {
        if self.auxiliary_changed {
            return Err(PatchError::InvalidText {
                field: "DOC auxiliary-table durable patch",
            });
        }
        let before_artifact = BlobId::of(self.before.bytes()).as_hex();
        let after_artifact = BlobId::of(self.after.bytes()).as_hex();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let mut operations = Vec::with_capacity(self.changes.len());
        for change in &self.changes {
            let forward = durable_operation(limits, change, &before_artifact, &mut forward_blobs)?;
            let inverse = durable_operation(
                limits,
                &change.inverse(),
                &after_artifact,
                &mut reverse_blobs,
            )?;
            operations.push(ReversibleOperation::new(forward, inverse));
        }
        litchi_core::patch::Patch::<Reversible>::new(
            limits,
            "litchi-doc-body",
            operations,
            forward_blobs,
            reverse_blobs,
        )
    }
}

/// Borrowed semantic change carried by an in-memory patch.
#[derive(Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChangeRef<'a> {
    /// One semantic text replacement.
    Text {
        target: TextTarget,
        before: &'a str,
        after: &'a str,
    },
    /// One direct character-property mutation.
    Format {
        target: TextTarget,
        property: CharacterProperty,
        before: Option<bool>,
        after: Option<bool>,
    },
    /// One reversible tracked-mark presence change.
    RevisionMark {
        kind: RevisionKind,
        start_cp: u32,
        end_cp: u32,
        before_present: bool,
        after_present: bool,
    },
    /// One passive embedded-object presentation flag mutation.
    EmbeddedDisplay {
        storage_id: u32,
        before: bool,
        after: bool,
    },
    /// One complete embedded field/preview/storage presence change.
    EmbeddedObject {
        storage_id: u32,
        before_present: bool,
        after_present: bool,
    },
    /// One text placeholder/picture-graph presence transition.
    Picture {
        target: TextTarget,
        before_present: bool,
        after_present: bool,
    },
    /// One valued direct character-property mutation, as the property's
    /// direct SPRMs before and after (empty when it had none).
    CharacterValue {
        target: TextTarget,
        property: CharacterValueProperty,
        before: &'a [u8],
        after: &'a [u8],
    },
    /// One ordinary main-story paragraph insertion (`before_units` is
    /// `None`) or removal (`after_units` is `None`), with UTF-16 lengths that
    /// include the paragraph mark.
    Paragraph {
        position: Position,
        before_units: Option<usize>,
        after_units: Option<usize>,
    },
    /// One ordinary main-story paragraph moved from `from` to `to`.
    ParagraphMove { from: Position, to: Position },
    /// One paragraph style change, as stylesheet indexes (record 0324).
    ParagraphStyle {
        position: Position,
        before: u16,
        after: u16,
    },
    /// One paragraph tab-stop change, as effective stops.
    ParagraphTabs {
        position: Position,
        before: &'a [crate::parts::pap::TabStop],
        after: &'a [crate::parts::pap::TabStop],
    },
    /// One main-story field insertion or removal, at its begin mark.
    Field {
        start_cp: u32,
        before_present: bool,
        after_present: bool,
    },
    /// One standard bookmark change, as its CP range before and after.
    Bookmark {
        name: &'a str,
        before: Option<(u32, u32)>,
        after: Option<(u32, u32)>,
    },
    /// One section SEPX change, as its SPRMs before and after.
    Section {
        section: Position,
        before: &'a [u8],
        after: &'a [u8],
    },
}

impl std::fmt::Debug for ChangeRef<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Text {
                target,
                before,
                after,
            } => formatter
                .debug_struct("Text")
                .field("target", target)
                .field("before_units", &before.encode_utf16().count())
                .field("after_units", &after.encode_utf16().count())
                .finish(),
            Self::Format {
                target,
                property,
                before,
                after,
            } => formatter
                .debug_struct("Format")
                .field("target", target)
                .field("property", property)
                .field("before", before)
                .field("after", after)
                .finish(),
            Self::RevisionMark {
                kind,
                start_cp,
                end_cp,
                before_present,
                after_present,
            } => formatter
                .debug_struct("RevisionMark")
                .field("kind", kind)
                .field("start_cp", start_cp)
                .field("end_cp", end_cp)
                .field("before_present", before_present)
                .field("after_present", after_present)
                .finish(),
            Self::EmbeddedDisplay {
                storage_id,
                before,
                after,
            } => formatter
                .debug_struct("EmbeddedDisplay")
                .field("storage_id", storage_id)
                .field("before", before)
                .field("after", after)
                .finish(),
            Self::EmbeddedObject {
                storage_id,
                before_present,
                after_present,
            } => formatter
                .debug_struct("EmbeddedObject")
                .field("storage_id", storage_id)
                .field("before_present", before_present)
                .field("after_present", after_present)
                .finish(),
            Self::Picture {
                target,
                before_present,
                after_present,
            } => formatter
                .debug_struct("Picture")
                .field("target", target)
                .field("before_present", before_present)
                .field("after_present", after_present)
                .finish(),
            Self::CharacterValue {
                target,
                property,
                before,
                after,
            } => formatter
                .debug_struct("CharacterValue")
                .field("target", target)
                .field("property", property)
                .field("before_bytes", &before.len())
                .field("after_bytes", &after.len())
                .finish(),
            Self::Paragraph {
                position,
                before_units,
                after_units,
            } => formatter
                .debug_struct("Paragraph")
                .field("position", position)
                .field("before_units", before_units)
                .field("after_units", after_units)
                .finish(),
            Self::ParagraphMove { from, to } => formatter
                .debug_struct("ParagraphMove")
                .field("from", from)
                .field("to", to)
                .finish(),
            Self::ParagraphStyle {
                position,
                before,
                after,
            } => formatter
                .debug_struct("ParagraphStyle")
                .field("position", position)
                .field("before", before)
                .field("after", after)
                .finish(),
            Self::ParagraphTabs {
                position,
                before,
                after,
            } => formatter
                .debug_struct("ParagraphTabs")
                .field("position", position)
                .field("before_stops", &before.len())
                .field("after_stops", &after.len())
                .finish(),
            Self::Field {
                start_cp,
                before_present,
                after_present,
            } => formatter
                .debug_struct("Field")
                .field("start_cp", start_cp)
                .field("before_present", before_present)
                .field("after_present", after_present)
                .finish(),
            Self::Bookmark { before, after, .. } => formatter
                .debug_struct("Bookmark")
                .field("before", before)
                .field("after", after)
                .finish_non_exhaustive(),
            Self::Section {
                section,
                before,
                after,
            } => formatter
                .debug_struct("Section")
                .field("section", section)
                .field("before_bytes", &before.len())
                .field("after_bytes", &after.len())
                .finish(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RevisionSpec {
    kind: RevisionKind,
    start_cp: u32,
    end_cp: u32,
    metadata: RevisionMetadata,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum PictureSlot {
    Text(String),
    Graph(crate::tracked_revision::PictureGraph),
}

impl PictureSlot {
    const fn is_graph(&self) -> bool {
        matches!(self, Self::Graph(_))
    }
}

impl RevisionSpec {
    fn from_revision(revision: &Revision) -> Self {
        Self {
            kind: revision.kind,
            start_cp: revision.start_cp,
            end_cp: revision.end_cp,
            metadata: RevisionMetadata {
                author: revision.author.clone(),
                timestamp: revision.timestamp,
                reason: revision.reason,
                revision_save_id: revision.revision_save_id,
            },
        }
    }

    fn identity(&self) -> String {
        format!(
            "revision/{}:{}:{}",
            revision_kind_name(self.kind),
            self.start_cp,
            self.end_cp
        )
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum Change {
    Text {
        target: TextTarget,
        before: String,
        after: String,
    },
    Format {
        target: TextTarget,
        property: CharacterProperty,
        before: Option<bool>,
        after: Option<bool>,
    },
    RevisionMark {
        identity: String,
        before: Option<RevisionSpec>,
        after: Option<RevisionSpec>,
    },
    EmbeddedDisplay {
        storage_id: u32,
        before: bool,
        after: bool,
    },
    EmbeddedObject {
        storage_id: u32,
        before: Option<crate::embedded_object::WriteOptions>,
        after: Option<crate::embedded_object::WriteOptions>,
    },
    Picture {
        target: TextTarget,
        before: PictureSlot,
        after: PictureSlot,
    },
    CharacterValue {
        target: TextTarget,
        property: CharacterValueProperty,
        before: Vec<u8>,
        after: Vec<u8>,
    },
    Paragraph {
        position: Position,
        before: Option<ParagraphCopy>,
        after: Option<ParagraphCopy>,
    },
    ParagraphMove {
        from: Position,
        to: Position,
    },
    /// A record-0324 style, tab, field, bookmark or section change.
    Mark(MarkChange),
}

impl Change {
    fn effect(&self) -> String {
        match self {
            Self::Text { target, .. } => format!("{}/text", durable_target(*target)),
            Self::Format {
                target, property, ..
            } => format!("{}/{}", durable_target(*target), property.wire_name()),
            Self::RevisionMark { identity, .. } => identity.clone(),
            Self::EmbeddedDisplay { storage_id, .. } => {
                format!("resource/embedded:{storage_id}")
            },
            Self::EmbeddedObject { storage_id, .. } => {
                format!("resource/embedded:{storage_id}")
            },
            Self::Picture { target, .. } => format!("{}/text", durable_target(*target)),
            Self::CharacterValue {
                target, property, ..
            } => format!("{}/{}", durable_target(*target), property.wire_name()),
            Self::Paragraph { .. } | Self::ParagraphMove { .. } => {
                "story/main/structure".to_string()
            },
            Self::Mark(change) => change.effect(),
        }
    }

    fn same_after(&self, other: &Self) -> bool {
        match (self, other) {
            (
                Self::Text {
                    target: left_target,
                    after: left_after,
                    ..
                },
                Self::Text {
                    target: right_target,
                    after: right_after,
                    ..
                },
            ) => left_target == right_target && left_after == right_after,
            (
                Self::Format {
                    target: left_target,
                    property: left_property,
                    after: left_after,
                    ..
                },
                Self::Format {
                    target: right_target,
                    property: right_property,
                    after: right_after,
                    ..
                },
            ) => {
                left_target == right_target
                    && left_property == right_property
                    && left_after == right_after
            },
            (
                Self::RevisionMark {
                    identity: left_identity,
                    after: left_after,
                    ..
                },
                Self::RevisionMark {
                    identity: right_identity,
                    after: right_after,
                    ..
                },
            ) => left_identity == right_identity && left_after == right_after,
            (
                Self::EmbeddedDisplay {
                    storage_id: left_id,
                    after: left_after,
                    ..
                },
                Self::EmbeddedDisplay {
                    storage_id: right_id,
                    after: right_after,
                    ..
                },
            ) => left_id == right_id && left_after == right_after,
            (
                Self::EmbeddedObject {
                    storage_id: left_id,
                    after: left_after,
                    ..
                },
                Self::EmbeddedObject {
                    storage_id: right_id,
                    after: right_after,
                    ..
                },
            ) => left_id == right_id && left_after == right_after,
            (
                Self::Picture {
                    target: left_target,
                    after: left_after,
                    ..
                },
                Self::Picture {
                    target: right_target,
                    after: right_after,
                    ..
                },
            ) => left_target == right_target && left_after == right_after,
            (
                Self::CharacterValue {
                    target: left_target,
                    property: left_property,
                    after: left_after,
                    ..
                },
                Self::CharacterValue {
                    target: right_target,
                    property: right_property,
                    after: right_after,
                    ..
                },
            ) => {
                left_target == right_target
                    && left_property == right_property
                    && left_after == right_after
            },
            (left @ (Self::Paragraph { .. } | Self::ParagraphMove { .. }), right) => left == right,
            (Self::Mark(left), Self::Mark(right)) => left == right,
            _ => false,
        }
    }

    fn as_ref(&self) -> ChangeRef<'_> {
        match self {
            Self::Text {
                target,
                before,
                after,
            } => ChangeRef::Text {
                target: *target,
                before,
                after,
            },
            Self::Format {
                target,
                property,
                before,
                after,
            } => ChangeRef::Format {
                target: *target,
                property: *property,
                before: *before,
                after: *after,
            },
            Self::RevisionMark { before, after, .. } => ChangeRef::RevisionMark {
                kind: before
                    .as_ref()
                    .or(after.as_ref())
                    .map_or(RevisionKind::Insertion, |spec| spec.kind),
                start_cp: before
                    .as_ref()
                    .or(after.as_ref())
                    .map_or(0, |spec| spec.start_cp),
                end_cp: before
                    .as_ref()
                    .or(after.as_ref())
                    .map_or(0, |spec| spec.end_cp),
                before_present: before.is_some(),
                after_present: after.is_some(),
            },
            Self::EmbeddedDisplay {
                storage_id,
                before,
                after,
            } => ChangeRef::EmbeddedDisplay {
                storage_id: *storage_id,
                before: *before,
                after: *after,
            },
            Self::EmbeddedObject {
                storage_id,
                before,
                after,
            } => ChangeRef::EmbeddedObject {
                storage_id: *storage_id,
                before_present: before.is_some(),
                after_present: after.is_some(),
            },
            Self::Picture {
                target,
                before,
                after,
            } => ChangeRef::Picture {
                target: *target,
                before_present: before.is_graph(),
                after_present: after.is_graph(),
            },
            Self::CharacterValue {
                target,
                property,
                before,
                after,
            } => ChangeRef::CharacterValue {
                target: *target,
                property: *property,
                before,
                after,
            },
            Self::Paragraph {
                position,
                before,
                after,
            } => ChangeRef::Paragraph {
                position: *position,
                before_units: before.as_ref().map(ParagraphCopy::len),
                after_units: after.as_ref().map(ParagraphCopy::len),
            },
            Self::ParagraphMove { from, to } => ChangeRef::ParagraphMove {
                from: *from,
                to: *to,
            },
            Self::Mark(change) => change.as_ref(),
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Text {
                target,
                before,
                after,
            } => Self::Text {
                target: *target,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Format {
                target,
                property,
                before,
                after,
            } => Self::Format {
                target: *target,
                property: *property,
                before: *after,
                after: *before,
            },
            Self::RevisionMark {
                identity,
                before,
                after,
            } => Self::RevisionMark {
                identity: identity.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::EmbeddedDisplay {
                storage_id,
                before,
                after,
            } => Self::EmbeddedDisplay {
                storage_id: *storage_id,
                before: *after,
                after: *before,
            },
            Self::EmbeddedObject {
                storage_id,
                before,
                after,
            } => Self::EmbeddedObject {
                storage_id: *storage_id,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Picture {
                target,
                before,
                after,
            } => Self::Picture {
                target: *target,
                before: after.clone(),
                after: before.clone(),
            },
            Self::CharacterValue {
                target,
                property,
                before,
                after,
            } => Self::CharacterValue {
                target: *target,
                property: *property,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Paragraph {
                position,
                before,
                after,
            } => Self::Paragraph {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ParagraphMove { from, to } => Self::ParagraphMove {
                from: *to,
                to: *from,
            },
            Self::Mark(change) => Self::Mark(change.inverse()),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Lineage(Arc<[u8]>);

#[derive(Clone, Debug, PartialEq, Eq)]
enum PreparedOperation {
    Text {
        target: TextTarget,
        replacement: String,
    },
    Format {
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    },
}

impl PreparedOperation {
    fn target(&self) -> TextTarget {
        match self {
            Self::Text { target, .. } | Self::Format { target, .. } => *target,
        }
    }

    fn effect(&self) -> String {
        let facet = match self {
            Self::Text { .. } => "text",
            Self::Format { property, .. } => property.wire_name(),
        };
        format!("{}/{facet}", durable_target(self.target()))
    }
}

/// One independently prepared body-text or formatting edit.
pub struct SubEdit {
    inner: CoreSubEdit<Lineage, PreparedOperation>,
    transaction_limits: TransactionLimits,
}

impl SubEdit {
    fn new(
        lineage: Lineage,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        operation: PreparedOperation,
        transaction_limits: TransactionLimits,
    ) -> Result<Self> {
        let writes = [operation.effect()];
        let inner = CoreSubEdit::new(
            lineage,
            limits,
            identifier,
            std::iter::empty(),
            writes,
            operation,
        )?;
        Ok(Self {
            inner,
            transaction_limits,
        })
    }

    /// Stable caller-selected composition identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        self.inner.id()
    }

    /// Semantic target in the immutable source collection.
    #[must_use]
    pub fn target(&self) -> TextTarget {
        self.inner.payload().target()
    }
}

impl std::fmt::Debug for SubEdit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SubEdit")
            .field("identifier", &self.identifier())
            .field("target", &self.target())
            .finish_non_exhaustive()
    }
}

/// Recoverable failure to join one independently prepared edit.
pub struct JoinError {
    failure: SubEditJoinFailure,
    rejected: SubEdit,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Recovers the rejected prepared edit.
    #[must_use]
    pub fn into_rejected(self) -> SubEdit {
        self.rejected
    }
}

impl std::fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected", &self.rejected)
            .finish()
    }
}

/// Bounded deterministic composition of provably disjoint body edits.
pub(crate) struct Composition {
    source: Snapshot,
    joined: JoinedSubEdits<Lineage, PreparedOperation>,
    transaction_limits: TransactionLimits,
}

impl Composition {
    /// Joins one edit when its exact artifact lineage and semantic facet are
    /// disjoint from every accepted edit.
    ///
    /// # Errors
    ///
    /// Returns a structured common composition refusal while retaining the
    /// rejected work.
    #[expect(
        clippy::result_large_err,
        reason = "composition refusals intentionally return the rejected prepared edit to the caller"
    )]
    pub(crate) fn join(&mut self, incoming: SubEdit) -> std::result::Result<&mut Self, JoinError> {
        let transaction_limits = incoming.transaction_limits;
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure,
                rejected: SubEdit {
                    inner: rejected,
                    transaction_limits,
                },
            });
        }
        self.transaction_limits = self.transaction_limits.intersect(transaction_limits);
        Ok(self)
    }

    /// Number of accepted independently prepared edits.
    #[must_use]
    pub(crate) fn len(&self) -> usize {
        self.joined.len()
    }

    /// Whether no edits have been accepted.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.joined.is_empty()
    }

    /// Commits all accepted edits atomically in stable identifier order.
    ///
    /// # Errors
    ///
    /// Returns the same staging, publication, and full-reopen errors as an
    /// ordinary multi-operation [`Edit`].
    pub(crate) fn into_edit(self) -> Result<Edit> {
        let source = self.source.with_transaction_limits(self.transaction_limits);
        let mut edit = source.edit()?;
        for prepared in self.joined.into_sub_edits() {
            match prepared.into_payload() {
                PreparedOperation::Text {
                    target,
                    replacement,
                } => edit.replace_text(target, &replacement)?,
                PreparedOperation::Format {
                    target,
                    property,
                    enabled,
                } => {
                    edit.set_character_property(target, property, enabled)?;
                },
            }
        }
        Ok(edit)
    }
}

impl std::fmt::Debug for Composition {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Composition")
            .field("edits", &self.joined.len())
            .finish_non_exhaustive()
    }
}

/// Dependency-closed inert text transfer prepared for one exact receiver.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferPlan {
    target: Lineage,
    destination: TextTarget,
    value: String,
    transaction_limits: TransactionLimits,
}

impl TransferPlan {
    #[must_use]
    pub const fn destination(&self) -> TextTarget {
        self.destination
    }

    #[must_use]
    pub fn text(&self) -> &str {
        &self.value
    }
}

/// Complete inert embedded-object closure prepared for one exact receiver.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedTransferPlan {
    target: Lineage,
    options: crate::embedded_object::WriteOptions,
    transaction_limits: TransactionLimits,
}

/// Canonical re-homed inline/floating picture closure prepared for one exact
/// receiver artifact.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PictureTransferPlan {
    target: Lineage,
    destination: TextTarget,
    graph: crate::tracked_revision::PictureGraph,
    transaction_limits: TransactionLimits,
}

impl PictureTransferPlan {
    #[must_use]
    pub const fn destination(&self) -> TextTarget {
        self.destination
    }

    #[must_use]
    pub const fn is_floating(&self) -> bool {
        self.graph.floating
    }

    #[must_use]
    pub fn picture_block_size(&self) -> usize {
        self.graph.picture_block.len()
    }
}

impl EmbeddedTransferPlan {
    /// Destination semantic storage identifier.
    #[must_use]
    pub const fn storage_id(&self) -> u32 {
        self.options.storage_id
    }

    /// Size of the standalone bounded object CFB.
    #[must_use]
    pub fn compound_size(&self) -> usize {
        self.options.compound_file.len()
    }

    /// Size of the exact `PICFAndOfficeArtData` preview dependency.
    #[must_use]
    pub fn preview_size(&self) -> usize {
        self.options.picture_data.len()
    }
}

/// One semantic facet changed differently by both sides of a three-way plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ThreeWayConflict {
    effect: String,
}

impl ThreeWayConflict {
    #[must_use]
    pub fn effect(&self) -> &str {
        &self.effect
    }
}

/// Non-mutating three-way merge plan over one exact immutable DOC source.
#[derive(Clone, Debug)]
pub(crate) struct ThreeWayPlan {
    source: Snapshot,
    changes: Vec<Change>,
    conflicts: Vec<ThreeWayConflict>,
}

impl ThreeWayPlan {
    fn new(source: Snapshot, left: &Patch, right: &Patch) -> Self {
        let transaction_limits = source
            .transaction_limits
            .intersect(left.before.transaction_limits)
            .intersect(left.after.transaction_limits)
            .intersect(right.before.transaction_limits)
            .intersect(right.after.transaction_limits);
        let source = source.with_transaction_limits(transaction_limits);
        let mut changes = left.changes.clone();
        let mut conflicts = Vec::new();
        for incoming in &right.changes {
            let effect = incoming.effect();
            if let Some(existing) = changes.iter().find(|change| change.effect() == effect) {
                if !existing.same_after(incoming) {
                    conflicts.push(ThreeWayConflict { effect });
                }
            } else {
                changes.push(incoming.clone());
            }
        }
        Self {
            source,
            changes,
            conflicts,
        }
    }

    #[must_use]
    pub(crate) fn conflicts(&self) -> &[ThreeWayConflict] {
        &self.conflicts
    }

    #[must_use]
    pub(crate) fn is_conflict_free(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Publishes the conflict-free combined plan through the ordinary bounded
    /// transaction and full reopen boundary.
    pub(crate) fn into_edit(self) -> Result<Edit> {
        if !self.conflicts.is_empty() {
            return Err(Error::Conflict);
        }
        let mut edit = self.source.edit()?;
        for change in &self.changes {
            apply_change_after(&mut edit, change)?;
        }
        Ok(edit)
    }
}

fn durable_operation(
    limits: PatchLimits,
    change: &Change,
    artifact: &str,
    blobs: &mut BlobBundle,
) -> std::result::Result<PatchOperation, PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        serde_json::Value::String(artifact.to_string()),
    );
    let (op, target, value) = match change {
        Change::Text {
            target,
            before,
            after,
        } => {
            preconditions.insert(
                "text".to_string(),
                serde_json::Value::String(before.clone()),
            );
            (
                "text.set".to_string(),
                durable_target(*target),
                serde_json::Value::String(after.clone()),
            )
        },
        Change::Format {
            target,
            property,
            before,
            after,
        } => {
            preconditions.insert(
                property.wire_name().to_string(),
                optional_bool_value(*before),
            );
            (
                format!("character-{}.set", property.wire_name()),
                durable_target(*target),
                optional_bool_value(*after),
            )
        },
        Change::RevisionMark {
            identity,
            before,
            after,
        } => {
            preconditions.insert("revision".to_string(), revision_spec_value(before.as_ref()));
            (
                "revision-mark.set".to_string(),
                identity.clone(),
                revision_spec_value(after.as_ref()),
            )
        },
        Change::EmbeddedDisplay {
            storage_id,
            before,
            after,
        } => {
            preconditions.insert("display_as_icon".to_string(), (*before).into());
            (
                "embedded-display.set".to_string(),
                format!("resource/embedded:{storage_id}"),
                (*after).into(),
            )
        },
        Change::EmbeddedObject {
            storage_id,
            before,
            after,
        } => {
            preconditions.insert("present".to_string(), before.is_some().into());
            (
                "embedded-object.set".to_string(),
                format!("resource/embedded:{storage_id}"),
                embedded_options_value(after.as_ref(), blobs)?,
            )
        },
        Change::Picture {
            target,
            before,
            after,
        } => {
            preconditions.insert(
                "picture_slot".to_string(),
                picture_slot_value(before, blobs)?,
            );
            (
                "picture.set".to_string(),
                durable_target(*target),
                picture_slot_value(after, blobs)?,
            )
        },
        Change::CharacterValue { .. } => {
            return Err(PatchError::InvalidText {
                field: "DOC valued character-format durable patch",
            });
        },
        Change::Paragraph { .. } | Change::ParagraphMove { .. } => {
            return Err(PatchError::InvalidText {
                field: "DOC paragraph-structure durable patch",
            });
        },
        Change::Mark(_) => {
            return Err(PatchError::InvalidText {
                field: "DOC style, tab, field, bookmark or section durable patch",
            });
        },
    };
    PatchOperation::new(limits, op, target, preconditions, value)
}

fn embedded_options_value(
    options: Option<&crate::embedded_object::WriteOptions>,
    blobs: &mut BlobBundle,
) -> std::result::Result<serde_json::Value, PatchError> {
    let Some(options) = options else {
        return Ok(serde_json::Value::Null);
    };
    let compound = blobs.insert(&options.compound_file)?.as_hex();
    let preview = blobs.insert(&options.picture_data)?.as_hex();
    let mut value = serde_json::Map::new();
    value.insert("compound_blob".to_string(), compound.into());
    value.insert("preview_blob".to_string(), preview.into());
    Ok(serde_json::Value::Object(value))
}

fn picture_slot_value(
    slot: &PictureSlot,
    blobs: &mut BlobBundle,
) -> std::result::Result<serde_json::Value, PatchError> {
    let mut value = serde_json::Map::new();
    match slot {
        PictureSlot::Text(text) => {
            value.insert("kind".to_string(), "text".into());
            value.insert("text".to_string(), text.clone().into());
        },
        PictureSlot::Graph(graph) => {
            let formatting = graph
                .replaced_grpprl
                .as_deref()
                .ok_or(PatchError::InvalidText {
                    field: "picture displaced formatting",
                })?;
            let data_offset = graph.data_offset.ok_or(PatchError::InvalidText {
                field: "picture Data offset",
            })?;
            value.insert("kind".to_string(), "graph".into());
            value.insert("floating".to_string(), graph.floating.into());
            value.insert(
                "picture_blob".to_string(),
                blobs.insert(&graph.picture_block)?.as_hex().into(),
            );
            let spa_blob = if let Some(spa) = graph.spa {
                blobs.insert(spa.to_bytes())?.as_hex().into()
            } else {
                serde_json::Value::Null
            };
            value.insert("spa_blob".to_string(), spa_blob);
            value.insert(
                "dgg_blob".to_string(),
                if graph.dgg_info.is_empty() {
                    serde_json::Value::Null
                } else {
                    blobs.insert(&graph.dgg_info)?.as_hex().into()
                },
            );
            value.insert(
                "format_blob".to_string(),
                blobs.insert(formatting)?.as_hex().into(),
            );
            value.insert("data_offset".to_string(), data_offset.into());
        },
    }
    Ok(serde_json::Value::Object(value))
}

fn apply_durable_embedded_object(
    edit: &mut Edit,
    operation: &PatchOperation,
    blobs: &BlobBundle,
    used_blobs: &mut BTreeSet<String>,
) -> Result<()> {
    let storage_id = operation
        .target
        .strip_prefix("resource/embedded:")
        .and_then(|value| value.parse::<u32>().ok())
        .ok_or_else(|| invalid_durable_patch("invalid embedded-resource target"))?;
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| invalid_durable_patch("embedded presence precondition is invalid"))?;
    let actual = embedded_object_value(edit, storage_id)?.is_some();
    if actual != expected {
        return Err(Error::PatchConflict);
    }
    if operation.value.is_null() {
        if !expected {
            return Err(invalid_durable_patch(
                "embedded-object removal requires an existing resource",
            ));
        }
        return edit.remove_embedded_object(storage_id);
    }
    if expected {
        return Err(invalid_durable_patch(
            "embedded-object replacement is outside this add/remove vocabulary",
        ));
    }
    let object = operation
        .value
        .as_object()
        .filter(|object| object.len() == 2)
        .ok_or_else(|| invalid_durable_patch("embedded-object value has invalid fields"))?;
    let compound = required_blob(object, "compound_blob", blobs, used_blobs)?;
    let preview = required_blob(object, "preview_blob", blobs, used_blobs)?;
    edit.add_embedded_object(crate::embedded_object::WriteOptions::new(
        storage_id, compound, preview,
    ))
}

fn required_blob(
    object: &serde_json::Map<String, serde_json::Value>,
    key: &str,
    blobs: &BlobBundle,
    used_blobs: &mut BTreeSet<String>,
) -> Result<Vec<u8>> {
    let id = object
        .get(key)
        .and_then(serde_json::Value::as_str)
        .filter(|value| value.len() == 64 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .ok_or_else(|| invalid_durable_patch("blob reference is invalid"))?;
    let blob_id = blobs
        .ids()
        .find(|candidate| candidate.as_hex() == id)
        .ok_or_else(|| invalid_durable_patch("referenced blob is missing"))?;
    let bytes = blobs
        .get(blob_id)
        .ok_or_else(|| invalid_durable_patch("referenced blob is missing"))?
        .to_vec();
    used_blobs.insert(id.to_string());
    Ok(bytes)
}

fn parse_picture_slot(
    value: &serde_json::Value,
    blobs: &BlobBundle,
    used_blobs: &mut BTreeSet<String>,
) -> Result<PictureSlot> {
    let object = value
        .as_object()
        .ok_or_else(|| invalid_durable_patch("picture slot is not an object"))?;
    match object.get("kind").and_then(serde_json::Value::as_str) {
        Some("text") if object.len() == 2 => object
            .get("text")
            .and_then(serde_json::Value::as_str)
            .map(|text| PictureSlot::Text(text.to_string()))
            .ok_or_else(|| invalid_durable_patch("picture text slot is invalid")),
        Some("graph") if object.len() == 7 => {
            let floating = object
                .get("floating")
                .and_then(serde_json::Value::as_bool)
                .ok_or_else(|| invalid_durable_patch("picture floating flag is invalid"))?;
            let picture_block = required_blob(object, "picture_blob", blobs, used_blobs)?;
            let dgg_info = if object
                .get("dgg_blob")
                .is_some_and(serde_json::Value::is_null)
            {
                Vec::new()
            } else {
                required_blob(object, "dgg_blob", blobs, used_blobs)?
            };
            let spa = if object
                .get("spa_blob")
                .is_some_and(serde_json::Value::is_null)
            {
                None
            } else {
                let bytes = required_blob(object, "spa_blob", blobs, used_blobs)?;
                if bytes.len() != crate::parts::spa::SPA_LEN {
                    return Err(invalid_durable_patch("picture SPA blob length is invalid"));
                }
                let spa = crate::parts::spa::Spa::parse(&bytes)
                    .map_err(|error| invalid_durable_patch(&error.to_string()))?;
                if spa.to_bytes().as_slice() != bytes.as_slice() {
                    return Err(invalid_durable_patch("picture SPA blob is noncanonical"));
                }
                Some(spa)
            };
            let formatting = required_blob(object, "format_blob", blobs, used_blobs)?;
            if formatting.len() > 255 {
                return Err(invalid_durable_patch(
                    "picture formatting blob exceeds CHPX limit",
                ));
            }
            let replaced_grpprl = Some(formatting);
            let data_offset = Some(required_u32(object, "data_offset")?);
            if floating != spa.is_some() || floating == dgg_info.is_empty() {
                return Err(invalid_durable_patch(
                    "picture graph SPA/Dgg closure is inconsistent",
                ));
            }
            let graph = crate::tracked_revision::PictureGraph {
                floating,
                picture_block,
                spa,
                dgg_info,
                replaced_grpprl,
                data_offset,
            };
            graph
                .validate_rehomed()
                .map_err(|error| invalid_durable_patch(&error.to_string()))?;
            Ok(PictureSlot::Graph(graph))
        },
        _ => Err(invalid_durable_patch("picture slot fields are invalid")),
    }
}

fn apply_durable_picture(
    edit: &mut Edit,
    operation: &PatchOperation,
    blobs: &BlobBundle,
    used_blobs: &mut BTreeSet<String>,
) -> Result<()> {
    let target = parse_durable_target(&operation.target)?;
    let expected = parse_picture_slot(
        operation
            .preconditions
            .get("picture_slot")
            .ok_or_else(|| invalid_durable_patch("picture slot precondition is missing"))?,
        blobs,
        used_blobs,
    )?;
    let value = parse_picture_slot(&operation.value, blobs, used_blobs)?;
    set_picture_slot(edit, target, expected, value)
}

fn optional_bool_value(value: Option<bool>) -> serde_json::Value {
    value.map_or(serde_json::Value::Null, serde_json::Value::Bool)
}

fn apply_durable_format(
    edit: &mut Edit,
    target: TextTarget,
    property: CharacterProperty,
    operation: &PatchOperation,
) -> Result<()> {
    let expected = parse_optional_bool(
        operation
            .preconditions
            .get(property.wire_name())
            .ok_or_else(|| invalid_durable_patch("missing character-format precondition"))?,
    )?;
    let span = resolve_target(&edit.editor, target)?;
    let actual = uniform_property_override(&edit.editor, span.start_cp, span.end_cp, property)
        .map_err(Error::Invalid)?
        .ok_or(Error::Refused(Refusal::FormattingDependency))?;
    if actual != expected {
        return Err(Error::PatchConflict);
    }
    let value = parse_optional_bool(&operation.value)?;
    edit.set_character_property_override(target, property, value)
}

fn revision_spec_value(spec: Option<&RevisionSpec>) -> serde_json::Value {
    let Some(spec) = spec else {
        return serde_json::Value::Null;
    };
    let mut value = serde_json::Map::new();
    value.insert(
        "kind".to_string(),
        serde_json::Value::String(revision_kind_name(spec.kind).to_string()),
    );
    value.insert("start_cp".to_string(), spec.start_cp.into());
    value.insert("end_cp".to_string(), spec.end_cp.into());
    value.insert(
        "author".to_string(),
        serde_json::Value::String(spec.metadata.author.clone()),
    );
    value.insert(
        "timestamp".to_string(),
        spec.metadata
            .timestamp
            .map_or(serde_json::Value::Null, timestamp_value),
    );
    value.insert(
        "reason".to_string(),
        spec.metadata
            .reason
            .map_or(serde_json::Value::Null, Into::into),
    );
    value.insert(
        "revision_save_id".to_string(),
        spec.metadata
            .revision_save_id
            .map_or(serde_json::Value::Null, Into::into),
    );
    serde_json::Value::Object(value)
}

fn timestamp_value(timestamp: DateTime) -> serde_json::Value {
    let mut value = serde_json::Map::new();
    for (key, number) in [
        ("year", u64::from(timestamp.year)),
        ("month", u64::from(timestamp.month)),
        ("day", u64::from(timestamp.day)),
        ("hour", u64::from(timestamp.hour)),
        ("minute", u64::from(timestamp.minute)),
        ("weekday", u64::from(timestamp.weekday)),
    ] {
        value.insert(key.to_string(), number.into());
    }
    serde_json::Value::Object(value)
}

fn parse_revision_spec(value: &serde_json::Value) -> Result<Option<RevisionSpec>> {
    if value.is_null() {
        return Ok(None);
    }
    let object = value
        .as_object()
        .filter(|object| object.len() == 7)
        .ok_or_else(|| invalid_durable_patch("revision mark value has invalid fields"))?;
    let kind = parse_revision_kind(required_string(object, "kind")?)?;
    let start_cp = required_u32(object, "start_cp")?;
    let end_cp = required_u32(object, "end_cp")?;
    let author = required_string(object, "author")?.to_string();
    let timestamp = parse_timestamp(
        object
            .get("timestamp")
            .ok_or_else(|| invalid_durable_patch("revision timestamp is missing"))?,
    )?;
    let reason = optional_u16(
        object
            .get("reason")
            .ok_or_else(|| invalid_durable_patch("revision reason is missing"))?,
    )?;
    let revision_save_id = optional_u32(
        object
            .get("revision_save_id")
            .ok_or_else(|| invalid_durable_patch("revision save ID is missing"))?,
    )?;
    Ok(Some(RevisionSpec {
        kind,
        start_cp,
        end_cp,
        metadata: RevisionMetadata {
            author,
            timestamp,
            reason,
            revision_save_id,
        },
    }))
}

fn required_string<'a>(
    object: &'a serde_json::Map<String, serde_json::Value>,
    key: &str,
) -> Result<&'a str> {
    object
        .get(key)
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("revision string field is invalid"))
}

fn required_u32(object: &serde_json::Map<String, serde_json::Value>, key: &str) -> Result<u32> {
    object
        .get(key)
        .and_then(serde_json::Value::as_u64)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| invalid_durable_patch("revision integer field is invalid"))
}

fn optional_u16(value: &serde_json::Value) -> Result<Option<u16>> {
    if value.is_null() {
        Ok(None)
    } else {
        value
            .as_u64()
            .and_then(|value| u16::try_from(value).ok())
            .map(Some)
            .ok_or_else(|| invalid_durable_patch("revision reason is invalid"))
    }
}

fn optional_u32(value: &serde_json::Value) -> Result<Option<u32>> {
    if value.is_null() {
        Ok(None)
    } else {
        value
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .map(Some)
            .ok_or_else(|| invalid_durable_patch("revision save ID is invalid"))
    }
}

fn parse_timestamp(value: &serde_json::Value) -> Result<Option<DateTime>> {
    if value.is_null() {
        return Ok(None);
    }
    let object = value
        .as_object()
        .filter(|object| object.len() == 6)
        .ok_or_else(|| invalid_durable_patch("revision timestamp has invalid fields"))?;
    Ok(Some(DateTime {
        year: required_u32(object, "year")?.try_into().map_err(|error| {
            invalid_durable_patch(&format!("revision year exceeds u16: {error}"))
        })?,
        month: required_u32(object, "month")?.try_into().map_err(|error| {
            invalid_durable_patch(&format!("revision month exceeds u8: {error}"))
        })?,
        day: required_u32(object, "day")?
            .try_into()
            .map_err(|error| invalid_durable_patch(&format!("revision day exceeds u8: {error}")))?,
        hour: required_u32(object, "hour")?.try_into().map_err(|error| {
            invalid_durable_patch(&format!("revision hour exceeds u8: {error}"))
        })?,
        minute: required_u32(object, "minute")?
            .try_into()
            .map_err(|error| {
                invalid_durable_patch(&format!("revision minute exceeds u8: {error}"))
            })?,
        weekday: required_u32(object, "weekday")?
            .try_into()
            .map_err(|error| {
                invalid_durable_patch(&format!("revision weekday exceeds u8: {error}"))
            })?,
    }))
}

fn revision_kind_name(kind: RevisionKind) -> &'static str {
    match kind {
        RevisionKind::Insertion => "insertion",
        RevisionKind::Deletion => "deletion",
        RevisionKind::MoveFrom => "move-from",
        RevisionKind::MoveTo => "move-to",
        RevisionKind::CharacterFormatting => "character-formatting",
        RevisionKind::ParagraphFormatting => "paragraph-formatting",
        RevisionKind::TableRowFormatting => "table-row-formatting",
    }
}

fn parse_revision_kind(value: &str) -> Result<RevisionKind> {
    Ok(match value {
        "insertion" => RevisionKind::Insertion,
        "deletion" => RevisionKind::Deletion,
        "move-from" => RevisionKind::MoveFrom,
        "move-to" => RevisionKind::MoveTo,
        "character-formatting" => RevisionKind::CharacterFormatting,
        "paragraph-formatting" => RevisionKind::ParagraphFormatting,
        "table-row-formatting" => RevisionKind::TableRowFormatting,
        _ => return Err(invalid_durable_patch("unknown revision kind")),
    })
}

fn apply_durable_revision(edit: &mut Edit, operation: &PatchOperation) -> Result<()> {
    let expected = parse_revision_spec(
        operation
            .preconditions
            .get("revision")
            .ok_or_else(|| invalid_durable_patch("revision precondition is missing"))?,
    )?;
    let value = parse_revision_spec(&operation.value)?;
    set_revision_mark(edit, &operation.target, expected, value)
}

fn apply_durable_embedded_display(edit: &mut Edit, operation: &PatchOperation) -> Result<()> {
    let storage_id = operation
        .target
        .strip_prefix("resource/embedded:")
        .and_then(|value| value.parse::<u32>().ok())
        .ok_or_else(|| invalid_durable_patch("invalid embedded-resource target"))?;
    let expected = operation
        .preconditions
        .get("display_as_icon")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| invalid_durable_patch("embedded display precondition is invalid"))?;
    if embedded_display_value(edit, storage_id)? != expected {
        return Err(Error::PatchConflict);
    }
    let enabled = operation
        .value
        .as_bool()
        .ok_or_else(|| invalid_durable_patch("embedded display value is invalid"))?;
    edit.set_embedded_display_as_icon(storage_id, enabled)
}

fn embedded_display_value(edit: &Edit, storage_id: u32) -> Result<bool> {
    let bytes = edit.finish_candidate()?;
    let snapshot = crate::embedded_object::Snapshot::open_with_policy(
        bytes,
        edit.source.limits,
        edit.source.protection_policy.clone(),
    )
    .map_err(Error::Invalid)?;
    snapshot
        .inventory()
        .map_err(Error::Invalid)?
        .get(storage_id)
        .and_then(|entry| entry.metadata().obj_info())
        .map(|info| info.display_as_icon)
        .ok_or(Error::Refused(Refusal::TargetNotFound))
}

fn uniform_property_override(
    editor: &RevisionEditor,
    start: u32,
    end: u32,
    property: CharacterProperty,
) -> crate::package::Result<Option<Option<bool>>> {
    match property {
        CharacterProperty::Bold => editor.uniform_bold_override(start, end),
        CharacterProperty::Italic => editor.uniform_italic_override(start, end),
        CharacterProperty::Underline => editor.uniform_underline_override(start, end),
    }
}

fn set_property_override(
    editor: &mut RevisionEditor,
    start: u32,
    end: u32,
    property: CharacterProperty,
    value: Option<bool>,
) -> crate::package::Result<()> {
    match property {
        CharacterProperty::Bold => editor.set_character_bold_override(start, end, value),
        CharacterProperty::Italic => editor.set_character_italic_override(start, end, value),
        CharacterProperty::Underline => editor.set_character_underline_override(start, end, value),
    }
}

fn apply_change_after(edit: &mut Edit, change: &Change) -> Result<()> {
    match change {
        Change::CharacterValue {
            target,
            property,
            after,
            ..
        } => edit.set_character_fragment(*target, *property, after.clone()),
        Change::Paragraph {
            position,
            before: None,
            after: Some(copy),
        } => edit.insert_paragraph_copy(*position, copy),
        Change::Paragraph {
            position,
            before: Some(_),
            after: None,
        } => edit.remove_paragraph(*position),
        Change::Paragraph { .. } => Err(invalid_durable_patch("invalid paragraph transition")),
        Change::ParagraphMove { from, to } => edit.move_paragraph(*from, *to),
        Change::Mark(change) => change.apply_after(edit),
        Change::Text { target, after, .. } => edit.replace_text(*target, after),
        Change::Format {
            target,
            property,
            after,
            ..
        } => edit.set_character_property_override(*target, *property, *after),
        Change::RevisionMark {
            identity,
            before,
            after,
        } => set_revision_mark(edit, identity, before.clone(), after.clone()),
        Change::EmbeddedDisplay {
            storage_id, after, ..
        } => edit.set_embedded_display_as_icon(*storage_id, *after),
        Change::EmbeddedObject {
            storage_id,
            before,
            after,
        } => set_embedded_object(edit, *storage_id, before.clone(), after.clone()),
        Change::Picture {
            target,
            before,
            after,
        } => match (before, after) {
            (PictureSlot::Text(expected), PictureSlot::Graph(graph)) => {
                let actual = resolve_target(&edit.editor, *target)?;
                if actual.text != expected.as_str() {
                    return Err(Error::PatchConflict);
                }
                edit.install_picture(*target, graph.clone())
            },
            (PictureSlot::Graph(graph), PictureSlot::Text(text)) => {
                edit.restore_picture_text(*target, graph.clone(), text.clone())
            },
            _ => Err(invalid_durable_patch("invalid picture slot transition")),
        },
    }
}

fn set_picture_slot(
    edit: &mut Edit,
    target: TextTarget,
    expected: PictureSlot,
    value: PictureSlot,
) -> Result<()> {
    match (&expected, &value) {
        (PictureSlot::Text(text), PictureSlot::Graph(graph)) => {
            let actual = resolve_target(&edit.editor, target)?;
            if actual.text != text.as_str() {
                return Err(Error::PatchConflict);
            }
            edit.install_picture(target, graph.clone())
        },
        (PictureSlot::Graph(graph), PictureSlot::Text(text)) => {
            edit.restore_picture_text(target, graph.clone(), text.clone())
        },
        _ => Err(invalid_durable_patch("invalid picture slot transition")),
    }
}

fn set_embedded_object(
    edit: &mut Edit,
    storage_id: u32,
    expected: Option<crate::embedded_object::WriteOptions>,
    value: Option<crate::embedded_object::WriteOptions>,
) -> Result<()> {
    let actual = embedded_object_value(edit, storage_id)?;
    if actual != expected {
        return Err(Error::PatchConflict);
    }
    if expected == value {
        return Ok(());
    }
    match value {
        Some(options) if options.storage_id == storage_id => edit.add_embedded_object(options),
        None if expected.is_some() => edit.remove_embedded_object(storage_id),
        _ => Err(invalid_durable_patch(
            "embedded-object transition or storage identity is invalid",
        )),
    }
}

fn embedded_object_value(
    edit: &Edit,
    storage_id: u32,
) -> Result<Option<crate::embedded_object::WriteOptions>> {
    let bytes = edit.finish_candidate()?;
    let snapshot = crate::embedded_object::Snapshot::open_with_policy(
        bytes,
        edit.source.limits,
        edit.source.protection_policy.clone(),
    )
    .map_err(Error::Invalid)?;
    if snapshot
        .inventory()
        .map_err(Error::Invalid)?
        .get(storage_id)
        .is_none()
    {
        return Ok(None);
    }
    snapshot
        .export_for_transfer(storage_id, storage_id)
        .map(Some)
        .map_err(Error::Invalid)
}

fn set_revision_mark(
    edit: &mut Edit,
    identity: &str,
    expected: Option<RevisionSpec>,
    value: Option<RevisionSpec>,
) -> Result<()> {
    if expected
        .as_ref()
        .is_some_and(|spec| spec.identity() != identity)
        || value
            .as_ref()
            .is_some_and(|spec| spec.identity() != identity)
    {
        return Err(invalid_durable_patch(
            "revision target does not match its semantic value",
        ));
    }
    let revisions = edit.editor.revisions().map_err(Error::Invalid)?;
    let matches = revisions
        .iter()
        .enumerate()
        .filter_map(|(index, revision)| {
            let spec = RevisionSpec::from_revision(revision);
            (spec.identity() == identity).then_some((index, spec))
        })
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(Error::PatchConflict);
    }
    let actual = matches.first().map(|(_index, spec)| spec.clone());
    if actual != expected {
        return Err(Error::PatchConflict);
    }
    if expected == value {
        return Ok(());
    }
    edit.ensure_operation_capacity()?;
    match (&expected, &value) {
        (Some(_), None) => {
            let index = matches
                .first()
                .map(|(index, _spec)| *index)
                .ok_or(Error::PatchConflict)?;
            edit.editor.remove(index).map_err(Error::Invalid)?;
        },
        (None, Some(spec)) => {
            edit.editor
                .add(spec.start_cp, spec.end_cp, spec.kind, spec.metadata.clone())
                .map_err(Error::Invalid)?;
        },
        _ => {
            return Err(invalid_durable_patch(
                "revision mark transition must add or remove one mark",
            ));
        },
    }
    edit.changes.push(Change::RevisionMark {
        identity: identity.to_string(),
        before: expected,
        after: value,
    });
    Ok(())
}

fn parse_optional_bool(value: &serde_json::Value) -> Result<Option<bool>> {
    match value {
        serde_json::Value::Null => Ok(None),
        serde_json::Value::Bool(value) => Ok(Some(*value)),
        _ => Err(invalid_durable_patch(
            "bold value must be a Boolean or null",
        )),
    }
}

fn durable_target(target: TextTarget) -> String {
    match target {
        TextTarget::StoryParagraph { story, position } => {
            format!("story/{}/paragraph:{}", story.wire_name(), position.get())
        },
        TextTarget::TableCell(position) => format!("body/table-cell:{}", position.get()),
        TextTarget::FieldResult(position) => format!("body/field-result:{}", position.get()),
        TextTarget::BodySpan {
            paragraph,
            start,
            end,
        } => format!(
            "story/main/paragraph:{}/span:{start}-{end}",
            paragraph.get()
        ),
        TextTarget::Field { start, part } => format!("body/field:{start}/{}", part.wire_name()),
    }
}

fn parse_durable_target(target: &str) -> Result<TextTarget> {
    if let Some(position) = target.strip_prefix("body/table-cell:") {
        return parse_position(position).map(TextTarget::TableCell);
    }
    if let Some(position) = target.strip_prefix("body/field-result:") {
        return parse_position(position).map(TextTarget::FieldResult);
    }
    if let Some(position) = target.strip_prefix("body/paragraph:") {
        return parse_position(position).map(TextTarget::body_paragraph);
    }
    if let Some(remainder) = target.strip_prefix("body/field:")
        && let Some((start, part)) = remainder.split_once('/')
    {
        let part = match part {
            "instruction" => FieldPart::Instruction,
            "result" => FieldPart::Result,
            _ => return Err(invalid_durable_patch("unknown DOC field part")),
        };
        let start = u32::try_from(parse_position(start)?.get())
            .map_err(|_error| invalid_durable_patch("DOC field CP exceeds u32"))?;
        return Ok(TextTarget::Field { start, part });
    }
    if let Some(remainder) = target.strip_prefix("story/main/paragraph:")
        && let Some((paragraph, span)) = remainder.split_once("/span:")
    {
        let (start, end) = span
            .split_once('-')
            .ok_or_else(|| invalid_durable_patch("invalid DOC span target"))?;
        let unit = |value: &str| {
            parse_position(value).and_then(|value| {
                u32::try_from(value.get())
                    .map_err(|_error| invalid_durable_patch("DOC span offset exceeds u32"))
            })
        };
        return Ok(TextTarget::BodySpan {
            paragraph: parse_position(paragraph)?,
            start: unit(start)?,
            end: unit(end)?,
        });
    }
    let remainder = target
        .strip_prefix("story/")
        .ok_or_else(|| invalid_durable_patch("invalid DOC text target"))?;
    let (story_name, position) = remainder
        .split_once("/paragraph:")
        .ok_or_else(|| invalid_durable_patch("invalid DOC story paragraph target"))?;
    let story = match story_name {
        "main" => Story::Main,
        "footnote" => Story::Footnote,
        "header" => Story::Header,
        "comment" => Story::Comment,
        "endnote" => Story::Endnote,
        "textbox" => Story::Textbox,
        "header-textbox" => Story::HeaderTextbox,
        _ => return Err(invalid_durable_patch("unknown DOC story target")),
    };
    parse_position(position).map(|position| TextTarget::StoryParagraph { story, position })
}

fn parse_position(value: &str) -> Result<Position> {
    if value.is_empty()
        || !value.bytes().all(|byte| byte.is_ascii_digit())
        || (value != "0" && value.starts_with('0'))
    {
        return Err(invalid_durable_patch("invalid canonical DOC position"));
    }
    value.parse::<usize>().map(Position::new).map_err(|error| {
        invalid_durable_patch(&format!("DOC position exceeds this platform: {error}"))
    })
}

fn invalid_durable_patch(message: &str) -> Error {
    Error::Invalid(PackageError::InvalidFormat(format!(
        "invalid DOC body durable patch: {message}"
    )))
}

type Result<T> = std::result::Result<T, Error>;

#[derive(Clone)]
struct SourceParagraph {
    start_cp: u32,
    end_cp: u32,
    text: String,
}

#[derive(Clone)]
struct SourceSpan {
    target: TextTarget,
    start_cp: u32,
    end_cp: u32,
    text: String,
}

#[derive(Clone, Copy)]
enum TargetCollection {
    Story(Story),
    TableCells,
    FieldResults,
}

fn target_story(target: TextTarget) -> Story {
    match target {
        TextTarget::StoryParagraph { story, .. } => story,
        TextTarget::TableCell(_)
        | TextTarget::FieldResult(_)
        | TextTarget::BodySpan { .. }
        | TextTarget::Field { .. } => Story::Main,
    }
}

fn resolve_target(editor: &RevisionEditor, target: TextTarget) -> Result<SourceSpan> {
    if let TextTarget::BodySpan {
        paragraph,
        start,
        end,
    } = target
    {
        return resolve_body_span(editor, target, paragraph, start, end);
    }
    if let TextTarget::Field { start, part } = target {
        return properties::field_span(editor, start, part);
    }
    let spans = match target {
        TextTarget::StoryParagraph { story, .. } => {
            target_spans(editor, TargetCollection::Story(story))?
        },
        TextTarget::TableCell(_) => target_spans(editor, TargetCollection::TableCells)?,
        TextTarget::FieldResult(_) => target_spans(editor, TargetCollection::FieldResults)?,
        TextTarget::BodySpan { .. } | TextTarget::Field { .. } => {
            return Err(Error::Refused(Refusal::TargetNotFound));
        },
    };
    let position = match target {
        TextTarget::StoryParagraph { position, .. }
        | TextTarget::TableCell(position)
        | TextTarget::FieldResult(position) => position,
        TextTarget::BodySpan { paragraph, .. } => paragraph,
        TextTarget::Field { .. } => return Err(Error::Refused(Refusal::TargetNotFound)),
    };
    spans
        .get(position.get())
        .cloned()
        .ok_or(Error::Refused(Refusal::TargetNotFound))
}

/// Resolve a UTF-16 range of one ordinary main-story paragraph. The range
/// must be non-empty and lie on character boundaries of the paragraph text.
fn resolve_body_span(
    editor: &RevisionEditor,
    target: TextTarget,
    paragraph: Position,
    start: u32,
    end: u32,
) -> Result<SourceSpan> {
    let paragraphs = source_paragraphs(editor)?;
    let source = paragraphs
        .get(paragraph.get())
        .ok_or(Error::Refused(Refusal::TargetNotFound))?;
    if start >= end {
        return Err(Error::Refused(Refusal::TargetNotFound));
    }
    let mut units = 0u32;
    let mut start_byte = None;
    let mut end_byte = None;
    for (byte, character) in source.text.char_indices() {
        if units == start {
            start_byte = Some(byte);
        }
        if units == end {
            end_byte = Some(byte);
        }
        units = units
            .checked_add(u32::try_from(character.len_utf16()).unwrap_or(2))
            .ok_or_else(cp_overflow)?;
    }
    if units == end {
        end_byte = Some(source.text.len());
    }
    let (Some(start_byte), Some(end_byte)) = (start_byte, end_byte) else {
        return Err(Error::Refused(Refusal::TargetNotFound));
    };
    Ok(SourceSpan {
        target,
        start_cp: source.start_cp.checked_add(start).ok_or_else(cp_overflow)?,
        end_cp: source.start_cp.checked_add(end).ok_or_else(cp_overflow)?,
        text: source.text[start_byte..end_byte].to_string(),
    })
}

fn target_spans(editor: &RevisionEditor, collection: TargetCollection) -> Result<Vec<SourceSpan>> {
    match collection {
        TargetCollection::Story(Story::Main) => source_paragraphs(editor).map(|paragraphs| {
            paragraphs
                .into_iter()
                .enumerate()
                .map(|(position, paragraph)| SourceSpan {
                    target: TextTarget::body_paragraph(Position::new(position)),
                    start_cp: paragraph.start_cp,
                    end_cp: paragraph.end_cp,
                    text: paragraph.text,
                })
                .collect()
        }),
        TargetCollection::Story(story) => story_paragraph_spans(editor, story),
        TargetCollection::TableCells => simple_table_cell_spans(editor),
        TargetCollection::FieldResults => simple_field_result_spans(editor),
    }
}

fn items_from_spans(spans: Vec<SourceSpan>) -> Vec<TextItem> {
    spans
        .into_iter()
        .map(|span| TextItem {
            target: span.target,
            text: span.text,
        })
        .collect()
}

fn story_paragraph_spans(editor: &RevisionEditor, story: Story) -> Result<Vec<SourceSpan>> {
    let (origin, text) = editor.story_text(story.index()).map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut start_cp = origin;
    let mut start_byte = 0usize;
    let mut cp = origin;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(cp_overflow)?;
        if character == '\r' {
            output.push(SourceSpan {
                target: TextTarget::StoryParagraph {
                    story,
                    position: Position::new(output.len()),
                },
                start_cp,
                end_cp: cp,
                text: text[start_byte..byte].to_string(),
            });
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        } else if character == '\u{7}' {
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        }
        cp = next_cp;
    }
    Ok(output)
}

fn simple_table_cell_spans(editor: &RevisionEditor) -> Result<Vec<SourceSpan>> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut segment_cp = 0u32;
    let mut segment_byte = 0usize;
    let mut pending: Option<(u32, u32, usize, usize)> = None;
    let mut cp = 0u32;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(cp_overflow)?;
        if character == '\r' {
            if editor.is_in_table_at_cp(cp).map_err(Error::Invalid)? {
                if pending.is_some() {
                    return Err(Error::Refused(Refusal::ComplexTableCell));
                }
                pending = Some((segment_cp, cp, segment_byte, byte));
            } else {
                pending = None;
            }
            segment_cp = next_cp;
            segment_byte = byte + character.len_utf8();
        } else if character == '\u{7}' {
            if editor.is_in_table_at_cp(cp).map_err(Error::Invalid)? {
                let (start_cp, end_cp, start_byte, end_byte) = match pending.take() {
                    Some(span) if segment_cp == cp => span,
                    Some(_) => return Err(Error::Refused(Refusal::ComplexTableCell)),
                    None => (segment_cp, cp, segment_byte, byte),
                };
                output.push(SourceSpan {
                    target: TextTarget::TableCell(Position::new(output.len())),
                    start_cp,
                    end_cp,
                    text: text[start_byte..end_byte].to_string(),
                });
            }
            pending = None;
            segment_cp = next_cp;
            segment_byte = byte + character.len_utf8();
        }
        cp = next_cp;
    }
    Ok(output)
}

struct FieldFrame {
    separator: Option<(u32, usize)>,
    complex: bool,
}

fn simple_field_result_spans(editor: &RevisionEditor) -> Result<Vec<SourceSpan>> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut stack = Vec::<FieldFrame>::new();
    let mut cp = 0u32;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(cp_overflow)?;
        match character {
            '\u{13}' => {
                if let Some(parent) = stack.last_mut() {
                    parent.complex = true;
                }
                stack.push(FieldFrame {
                    separator: None,
                    complex: false,
                });
            },
            '\u{14}' => {
                let frame = stack
                    .last_mut()
                    .ok_or(Error::Refused(Refusal::ComplexField))?;
                if frame
                    .separator
                    .replace((next_cp, byte + character.len_utf8()))
                    .is_some()
                {
                    return Err(Error::Refused(Refusal::ComplexField));
                }
            },
            '\u{15}' => {
                let frame = stack.pop().ok_or(Error::Refused(Refusal::ComplexField))?;
                if stack.is_empty() {
                    let (start_cp, start_byte) = frame
                        .separator
                        .filter(|_| !frame.complex)
                        .ok_or(Error::Refused(Refusal::ComplexField))?;
                    output.push(SourceSpan {
                        target: TextTarget::FieldResult(Position::new(output.len())),
                        start_cp,
                        end_cp: cp,
                        text: text[start_byte..byte].to_string(),
                    });
                }
            },
            _ => {},
        }
        cp = next_cp;
    }
    if !stack.is_empty() {
        return Err(Error::Refused(Refusal::ComplexField));
    }
    Ok(output)
}

fn cp_overflow() -> Error {
    Error::Invalid(PackageError::Corrupted(
        "DOC semantic target CP overflow".to_string(),
    ))
}

fn source_paragraphs(editor: &RevisionEditor) -> Result<Vec<SourceParagraph>> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut start_cp = 0u32;
    let mut start_byte = 0usize;
    let mut cp = 0u32;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "main-story CP overflow".to_string(),
            ))
        })?;
        if character == '\r' {
            if !editor.is_in_table_at_cp(cp).map_err(Error::Invalid)? {
                output.push(SourceParagraph {
                    start_cp,
                    end_cp: cp,
                    text: text[start_byte..byte].to_string(),
                });
            }
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        } else if character == '\u{7}' {
            // A table cell marker is never an ordinary body paragraph.
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        }
        cp = next_cp;
    }
    if cp != editor.main_story_cp_len() {
        return Err(Error::Invalid(PackageError::Corrupted(
            "decoded main story has an inconsistent CP length".to_string(),
        )));
    }
    Ok(output)
}

fn projected_paragraphs(editor: &RevisionEditor, projection: Projection) -> Result<Vec<Paragraph>> {
    let source = source_paragraphs(editor)?;
    if projection == Projection::All {
        return Ok(source
            .into_iter()
            .enumerate()
            .map(|(position, paragraph)| Paragraph {
                position: Position::new(position),
                text: paragraph.text,
            })
            .collect());
    }
    let hidden = hidden_ranges(editor, projection)?;
    let mut output = Vec::new();
    for paragraph in source {
        let text = project_text(&paragraph, &hidden)?;
        output.push(Paragraph {
            position: Position::new(output.len()),
            text,
        });
    }
    Ok(output)
}

fn hidden_ranges(editor: &RevisionEditor, projection: Projection) -> Result<Vec<(u32, u32)>> {
    let mut ranges = Vec::new();
    for revision in editor.revisions().map_err(Error::Invalid)? {
        let hide = matches!(
            (projection, revision.kind),
            (
                Projection::Accepted,
                RevisionKind::Deletion | RevisionKind::MoveFrom
            ) | (
                Projection::Rejected,
                RevisionKind::Insertion | RevisionKind::MoveTo
            )
        );
        if hide && revision.start_cp < revision.end_cp {
            ranges.push((revision.start_cp, revision.end_cp));
        }
    }
    ranges.sort_unstable();
    if ranges.windows(2).any(|pair| pair[0].1 > pair[1].0) {
        return Err(Error::Refused(Refusal::AmbiguousReviewRanges));
    }
    Ok(ranges)
}

fn project_text(paragraph: &SourceParagraph, hidden: &[(u32, u32)]) -> Result<String> {
    let mut output = String::new();
    let mut cp = paragraph.start_cp;
    for character in paragraph.text.chars() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let end = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "projection CP overflow".to_string(),
            ))
        })?;
        if !hidden
            .iter()
            .any(|(start, finish)| *start < end && cp < *finish)
        {
            output.push(character);
        }
        cp = end;
    }
    Ok(output)
}

fn text_revision_intersects(editor: &RevisionEditor, start: u32, end: u32) -> Result<bool> {
    Ok(editor
        .revisions()
        .map_err(Error::Invalid)?
        .into_iter()
        .any(|revision| {
            matches!(
                revision.kind,
                RevisionKind::Insertion
                    | RevisionKind::Deletion
                    | RevisionKind::MoveFrom
                    | RevisionKind::MoveTo
            ) && revision.start_cp < end
                && start < revision.end_cp
        }))
}

/// The keyword of a field instruction, upper-cased.
fn keyword_of(instruction: &str) -> Option<String> {
    instruction
        .split_whitespace()
        .next()
        .map(str::to_ascii_uppercase)
}

fn has_structural_content(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(character, '\r' | '\u{7}' | '\u{13}'..='\u{15}' | '\u{fffc}')
            || (character.is_control() && character != '\t')
    })
}

/// Return a private mutation-capable copy when `bytes` contains legacy CFB
/// signature state. The signature editor is used only as a bounded detector
/// and candidate builder; callers retain and publish the original bytes.
fn unsigned_doc_source(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut editor = litchi_sign::cfb::Editor::open(
        bytes.to_vec(),
        litchi_sign::cfb::Format::Doc,
        litchi_sign::Limits::standard(),
    )
    .ok()?;
    editor.clear().ok()?;
    let unsigned = editor.finish().ok()?;
    (unsigned != bytes).then_some(unsigned)
}

/// Re-apply the caller's physical source policy to the original CFB before a
/// signed fallback strips signature metadata. The strict owner may have
/// failed during its admission callback, so validating only the stripped
/// candidate would allow an oversized signature stream to disappear from the
/// policy accounting.
fn validate_original_source_for_fallback(
    bytes: &[u8],
    source_limits: crate::package::SourceLimits,
) -> Result<()> {
    let ole = crate::raw::Reader::<Cursor<Vec<u8>>>::open_ole_file_with_limits(
        Cursor::new(bytes.to_vec()),
        source_limits,
    )
    .map_err(Error::Invalid)?;
    reject_malformed_signature_containers(&ole)?;
    crate::package::preflight_declared_streams(&ole, source_limits)
        .map(|_| ())
        .map_err(Error::Invalid)
}

/// Reject signature names that are present in an impossible CFB shape before
/// the unsigned fallback can discard them. A legacy signature is a root
/// stream; an XML signature is a root storage containing only streams. An
/// empty or recursively nested signature storage cannot be authenticated or
/// safely preserved by the bounded fallback editor.
fn reject_malformed_signature_containers<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.directory_exists(&["_signatures"]) {
        return Err(Error::Invalid(PackageError::InvalidFormat(
            "legacy DOC signature marker is a storage".into(),
        )));
    }
    if ole.directory_exists(&["_xmlsignatures"]) {
        let entries = ole
            .list_directory_entries(&["_xmlsignatures"])
            .map_err(|error| Error::Invalid(PackageError::from(error)))?;
        if entries.is_empty() || entries.iter().any(|entry| entry.entry_type != 2) {
            return Err(Error::Invalid(PackageError::InvalidFormat(
                "DOC XML signature storage has an invalid entry shape".into(),
            )));
        }
    }
    Ok(())
}

fn drawing_dependency(text: &str) -> Option<DrawingDependency> {
    if text.contains('\u{1}') {
        Some(DrawingDependency::InlinePictureOrObjectPreview)
    } else if text.contains('\u{8}') {
        Some(DrawingDependency::FloatingOfficeArt)
    } else if text.contains('\u{fffc}') {
        Some(DrawingDependency::ObjectReplacement)
    } else if text
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\r'))
    {
        Some(DrawingDependency::UnsupportedControl)
    } else {
        None
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn same_source_allocation(left: &Arc<[u8]>, right: &Arc<[u8]>) -> bool {
    Arc::ptr_eq(left, right) && left.len() == right.len()
}

#[cfg(test)]
impl Snapshot {
    /// Unit-test shorthand: replay a durable patch through the engine.
    pub(crate) fn apply_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<Self> {
        self.stage_durable(patch)?
            .commit()
            .map(Commit::into_package)
    }
}

#[cfg(test)]
impl Composition {
    /// Unit-test shorthand: commit through the engine.
    pub(crate) fn commit(self) -> Result<Commit> {
        self.into_edit()?.commit()
    }
}

#[cfg(test)]
impl ThreeWayPlan {
    /// Unit-test shorthand: commit through the engine.
    pub(crate) fn commit(self) -> Result<Commit> {
        self.into_edit()?.commit()
    }
}

#[cfg(test)]
impl Edit {
    /// Unit-test shorthand: the immutable source snapshot.
    pub(crate) const fn source(&self) -> &Snapshot {
        &self.source
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "unit-test fixtures use contextual fail-fast assertions"
    )]

    #[cfg(feature = "performance-diagnostics")]
    use super::{
        CfbParseEvent, CfbParseOutcome, DiagnosticEvent, DiagnosticOutcome, DiagnosticPhase,
        observe_phase,
    };
    use super::{
        CharacterProperty, DrawingDependency, Error, OpenOptions, Projection, Refusal,
        RevisionDisposition, Snapshot, Story, TextTarget, TransactionLimits, fingerprint,
    };
    use crate::parts::dofr::{DofrFrameKind, DofrPayload};
    use crate::parts::fib::FileInformationBlock;
    #[cfg(feature = "performance-diagnostics")]
    use crate::parts::protection::ProtectionPolicy;
    use crate::tracked_revision::{Limits, RevisionEditor};
    use crate::writer::{
        BookmarkEntry, Builder, CharacterFormatting, FloatingPosition, ParagraphFormatting,
        Picture, TextRevision,
    };
    use litchi_cfb::{OleFile, OleWriter};
    use litchi_core::Position;
    use litchi_core::patch::{
        BlobLimits, CompositionLimits, HistoryLimits, Patch, PatchLimits, Reversible,
        SubEditJoinFailure,
    };
    use litchi_ole_common::object::{
        Edit as PackageEditor, Limits as ObjectLimits, OpenOptions as ObjectOpenOptions,
        Package as ObjectPackage, Targets,
    };
    use std::io::Cursor;
    use std::sync::{Arc, OnceLock};

    fn open_object(
        bytes: Vec<u8>,
        targets: Targets,
        limits: ObjectLimits,
    ) -> Result<PackageEditor, litchi_ole_common::Error> {
        ObjectPackage::open_with_options(
            bytes,
            ObjectOpenOptions::new().targets(targets).limits(limits),
        )
        .and_then(|package| package.edit())
    }

    fn doc(paragraphs: &[&str]) -> Vec<u8> {
        let mut writer = Builder::new();
        for paragraph in paragraphs {
            writer
                .add_paragraph_runs(
                    vec![(paragraph.to_string(), CharacterFormatting::default())],
                    ParagraphFormatting::default(),
                )
                .expect("fixture paragraph must be valid");
        }
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("fixture DOC must serialize");
        normalize_word2002_dop(output.into_inner())
    }

    fn formatted_empty_doc() -> Vec<u8> {
        let mut writer = Builder::new();
        writer
            .add_paragraph_runs(
                vec![(
                    String::new(),
                    CharacterFormatting {
                        bold: Some(true),
                        italic: Some(true),
                        ..CharacterFormatting::default()
                    },
                )],
                ParagraphFormatting::default(),
            )
            .expect("formatted empty paragraph must be accepted");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .expect("DOC fixture builds")
            .write_to(&mut output)
            .expect("formatted empty DOC must serialize");
        normalize_word2002_dop(output.into_inner())
    }

    fn empty_body_and_table_doc() -> Vec<u8> {
        let mut writer = Builder::new();
        writer
            .add_paragraph("")
            .expect("empty body paragraph must be accepted");
        let table = writer.add_table(1, 1).expect("table");
        writer
            .set_table_cell_text(table, 0, 0, "")
            .expect("empty table cell must be accepted");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .expect("DOC fixture builds")
            .write_to(&mut output)
            .expect("empty body/table DOC must serialize");
        normalize_word2002_dop(output.into_inner())
    }

    fn normalize_word2002_dop(bytes: Vec<u8>) -> Vec<u8> {
        let mut package = open_object(bytes, Targets::default(), Limits::default())
            .expect("fixture package should open");
        let word_path = ["WordDocument".to_string()];
        let word = package.stream(&word_path).expect("WordDocument");
        let fib = FileInformationBlock::parse(word).expect("fixture FIB");
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_string()];
        let mut word = word.to_vec();
        let mut table = package
            .stream(&table_path)
            .expect("selected table stream")
            .to_vec();
        let (offset, length) = fib.get_table_pointer(31).expect("DOP pointer");
        let offset = usize::try_from(offset).expect("DOP offset");
        let length = usize::try_from(length).expect("DOP length");
        let mut dop = crate::parts::document_properties::DocumentProperties::word97_writer_bytes(
            false, false, false,
        );
        dop.resize(594, 0);
        if length < dop.len() {
            let insertion = offset + length;
            let extra = dop.len() - length;
            table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
            let count = fib.table_pointer_count().expect("FIB pointer count");
            for index in 0..count {
                let pointer = 154 + index * 8;
                let current = usize::try_from(u32::from_le_bytes(
                    word[pointer..pointer + 4]
                        .try_into()
                        .expect("FIB pointer offset"),
                ))
                .expect("FIB pointer offset");
                if current >= insertion {
                    let shifted = u32::try_from(current + extra).expect("shifted FIB pointer");
                    word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
                }
            }
        }
        assert!(table.len() >= offset + dop.len());
        table[offset..offset + dop.len()].copy_from_slice(&dop);
        let pointer = 154 + 31 * 8;
        word[pointer + 4..pointer + 8]
            .copy_from_slice(&u32::try_from(dop.len()).expect("DOP length").to_le_bytes());
        package
            .put_stream(&word_path, word)
            .expect("normalized WordDocument stream");
        package
            .put_stream(&table_path, table)
            .expect("normalized table stream");
        crate::package::object_edit_bytes(&package).expect("normalized fixture publication")
    }

    fn doc_with_opaque_stream(paragraphs: &[&str]) -> Vec<u8> {
        let mut package = open_object(doc(paragraphs), Targets::default(), Limits::default())
            .expect("fixture package should open");
        package
            .add_stream(vec!["OpaqueVendorData".to_string()], b"untouched".to_vec())
            .expect("opaque stream should be admitted");
        crate::package::object_edit_bytes(&package).expect("fixture package should finish")
    }

    fn doc_with_signature_stream(payload: &[u8]) -> Vec<u8> {
        let mut source = OleFile::open(
            Cursor::new(doc(&["alpha"])),
            litchi_cfb::OleFileLimits::default(),
        )
        .expect("fixture CFB should open");
        let paths = source.list_streams();
        let mut writer = OleWriter::new();
        for path in paths {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let stream = source
                .open_stream(&refs)
                .expect("fixture stream should open");
            writer
                .create_stream(&refs, &stream)
                .expect("fixture stream should copy");
        }
        writer
            .create_stream(&["_xmlsignatures"], payload)
            .expect("signature stream should be admitted");
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("signed fixture should finish");
        output.into_inner()
    }

    fn dofr_record(kind: u32, payload: &[u8]) -> Vec<u8> {
        let size = u32::try_from(8usize + payload.len()).expect("Dofr record fits");
        let mut record = Vec::with_capacity(usize::try_from(size).expect("record size fits"));
        record.extend_from_slice(&size.to_le_bytes());
        record.extend_from_slice(&kind.to_le_bytes());
        record.extend_from_slice(payload);
        record
    }

    fn doc_with_auxiliary_tables() -> Vec<u8> {
        let mut package = open_object(doc(&["alpha bravo"]), Targets::default(), Limits::default())
            .expect("fixture package should open");
        let word_path = ["WordDocument".to_string()];
        let mut word = package
            .stream(&word_path)
            .expect("WordDocument stream")
            .to_vec();
        let fib = FileInformationBlock::parse(&word).expect("fixture FIB");
        assert!(fib.table_pointer_count().expect("FIB pair count") > 99);
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_string()];
        let mut table = package.stream(&table_path).expect("table stream").to_vec();

        let mut selection = vec![0u8; 36];
        selection[2] = 1;
        selection[4..8].copy_from_slice(&4i32.to_le_bytes());
        selection[8..12].copy_from_slice(&8i32.to_le_bytes());
        selection[20..24].copy_from_slice(&4i32.to_le_bytes());
        selection[24..26].copy_from_slice(&1u16.to_le_bytes());
        selection[32..34].copy_from_slice(&(-100i16).to_le_bytes());
        selection[34..36].copy_from_slice(&100i16.to_le_bytes());
        let selection_offset = u32::try_from(table.len()).expect("table offset fits");
        table.extend_from_slice(&selection);

        let mut frame_payload = vec![0u8; 36];
        frame_payload[12..16].copy_from_slice(&2u32.to_le_bytes());
        let mut dofr = dofr_record(0, &[]);
        dofr.extend_from_slice(&dofr_record(1, &frame_payload));
        let dofr_offset = u32::try_from(table.len()).expect("table offset fits");
        table.extend_from_slice(&dofr);

        for (index, offset, length) in [
            (30usize, selection_offset, selection.len()),
            (99usize, dofr_offset, dofr.len()),
        ] {
            let pair = 154usize
                .checked_add(index.checked_mul(8).expect("FIB pair offset"))
                .expect("FIB pair offset");
            word[pair..pair + 4].copy_from_slice(&offset.to_le_bytes());
            word[pair + 4..pair + 8].copy_from_slice(
                &u32::try_from(length)
                    .expect("table length fits")
                    .to_le_bytes(),
            );
        }
        package
            .put_stream(&word_path, word)
            .expect("replace WordDocument stream");
        package
            .put_stream(&table_path, table)
            .expect("replace table stream");
        package
            .add_stream(
                vec!["OpaqueVendorData".to_string()],
                b"untouched auxiliary stream".to_vec(),
            )
            .expect("opaque auxiliary stream");
        crate::package::object_edit_bytes(&package).expect("fixture package should finish")
    }

    fn picture_doc(floating: bool) -> Vec<u8> {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/images/png/lena.png"),
        )
        .expect("PNG fixture");
        let picture = Picture::new(bytes).expect("picture fixture");
        let mut writer = Builder::new();
        if floating {
            writer
                .insert_floating_picture(picture, FloatingPosition::new(1440, 720))
                .expect("floating picture");
        } else {
            writer.insert_picture(picture).expect("inline picture");
        }
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("picture DOC");
        normalize_word2002_dop(output.into_inner())
    }

    fn formatted_picture_receiver_doc() -> Vec<u8> {
        let mut writer = Builder::new();
        writer
            .add_paragraph_runs(
                vec![(
                    "placeholder".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        ..CharacterFormatting::default()
                    },
                )],
                ParagraphFormatting::default(),
            )
            .expect("formatted placeholder");
        writer.add_paragraph("other").expect("second paragraph");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("formatted receiver DOC");
        normalize_word2002_dop(output.into_inner())
    }

    fn structured_doc() -> Vec<u8> {
        let mut writer = Builder::new();
        writer.add_paragraph("Body").expect("body paragraph");
        let table = writer.add_table(1, 1).expect("table");
        writer
            .set_table_cell_text(table, 0, 0, "Cell")
            .expect("cell text");
        writer
            .add_hyperlink(
                "Link",
                "https://example.invalid",
                ParagraphFormatting::default(),
            )
            .expect("hyperlink field");
        writer.set_odd_header("Header");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("structured DOC");
        normalize_word2002_dop(output.into_inner())
    }

    fn embedded_doc() -> Vec<u8> {
        let mut object = OleWriter::new();
        object
            .create_stream(
                &["\u{1}CompObj"],
                &crate::writer::ole_metadata::generate_compobj_stream(),
            )
            .expect("CompObj stream");
        object
            .create_stream(
                &["\u{1}Ole"],
                &crate::writer::ole_metadata::generate_ole_stream(),
            )
            .expect("Ole stream");
        object
            .create_stream(&["\u{3}ObjInfo"], &[0x00, 0x00, 0x03, 0x00, 0x00, 0x00])
            .expect("ObjInfo stream");
        let mut object_output = Cursor::new(Vec::new());
        object
            .write_to(&mut object_output)
            .expect("object CFB serialization");

        let mut picture = 12u32.to_le_bytes().to_vec();
        picture.extend_from_slice(&77u32.to_le_bytes());
        picture.extend_from_slice(&[0; 4]);
        let mut editor =
            crate::embedded_object::Editor::open(doc(&["embedded"]), Limits::default())
                .expect("embedded-object editor");
        editor
            .add(crate::embedded_object::WriteOptions::new(
                77,
                object_output.into_inner(),
                picture,
            ))
            .expect("embedded object");
        editor.finish().expect("embedded DOC serialization")
    }

    fn patch_limits() -> PatchLimits {
        PatchLimits::new(
            BlobLimits::new(0, 0, 0),
            128 * 1024,
            32,
            8,
            16 * 1024,
            64 * 1024,
        )
    }

    fn resource_patch_limits() -> PatchLimits {
        PatchLimits::new(
            BlobLimits::new(8, 4 * 1024 * 1024, 8 * 1024 * 1024),
            12 * 1024 * 1024,
            32,
            8,
            16 * 1024,
            12 * 1024 * 1024,
        )
    }

    #[cfg(feature = "performance-diagnostics")]
    fn public_reader_failure_doc() -> Vec<u8> {
        let mut package = open_object(doc(&["alpha"]), Targets::default(), Limits::default())
            .expect("fixture package should open");
        let mut word = package
            .stream(&["WordDocument".to_string()])
            .expect("WordDocument stream")
            .to_vec();
        // FIB pair 1 (fcStshf/lcbStshf) is owned by the public reader's
        // stylesheet parse, not by the strict body-text owner. A one-byte
        // stylesheet is a framing error, which stays fatal under every
        // leniency (optional tables such as the RSID list now degrade instead,
        // see LENIENCY_AUDIT F1.1), while the strict owner ignores it.
        let pointer = 154 + 8;
        word[pointer + 4..pointer + 8].copy_from_slice(&1_u32.to_le_bytes());
        package
            .put_stream(&["WordDocument".to_string()], word)
            .expect("replace WordDocument stream");
        crate::package::object_edit_bytes(&package).expect("fixture package should finish")
    }

    fn protected_dop_doc() -> Vec<u8> {
        let mut package = open_object(doc(&["alpha"]), Targets::default(), Limits::default())
            .expect("fixture package should open");
        let word_path = ["WordDocument".to_string()];
        let word = package
            .stream(&word_path)
            .expect("WordDocument stream")
            .to_vec();
        let flags = u16::from_le_bytes([word[10], word[11]]);
        let table_name = if flags & 0x0200 != 0 {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_string()];
        let mut table = package
            .stream(&table_path)
            .expect("selected Table stream")
            .to_vec();
        let pointer = 154 + 31 * 8;
        let offset = usize::try_from(u32::from_le_bytes(
            word[pointer..pointer + 4]
                .try_into()
                .expect("DOP offset bytes"),
        ))
        .expect("DOP offset fits usize");
        let length = usize::try_from(u32::from_le_bytes(
            word[pointer + 4..pointer + 8]
                .try_into()
                .expect("DOP length bytes"),
        ))
        .expect("DOP length fits usize");
        assert!(length >= 84, "fixture DOP should have protection fields");
        table[offset + 6] |= 0x10;
        package
            .put_stream(&table_path, table)
            .expect("protected DOP should remain a valid package");
        crate::package::object_edit_bytes(&package).expect("fixture package should finish")
    }

    #[test]
    fn protected_snapshot_reads_and_noops_by_default_but_changed_commit_is_denied() {
        let bytes = protected_dop_doc();
        let snapshot = Snapshot::open(bytes.clone(), Limits::default()).expect("protected open");
        assert_eq!(snapshot.finish(), bytes);
        let mut edit = snapshot.edit().expect("protected snapshot edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "omega"),
            Err(Error::Invalid(crate::package::Error::ProtectionDenied(_)))
        ));
        assert!(matches!(
            edit.commit(),
            Ok(commit) if !commit.diagnostics().changed()
        ));

        let noop = snapshot
            .edit()
            .expect("protected snapshot edit")
            .commit()
            .expect("protected no-op");
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.package().finish(), bytes);
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_open_matches_ordinary_open_and_balances_success_events() {
        let bytes = doc(&["alpha", "bravo"]);
        let ordinary = Snapshot::open(bytes.clone(), Limits::default()).expect("ordinary open");
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = Snapshot::open_bounded_profiled_with_cfb_observer(
            bytes,
            Limits::default(),
            TransactionLimits::default(),
            crate::package::SourceLimits::default(),
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        )
        .expect("profiled open");

        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
        assert_eq!(profiled, ordinary);
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::PublicReaderValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::PublicReaderValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::SourceRetention,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::SourceRetention,
                    outcome: DiagnosticOutcome::Success,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_open_matches_ordinary_error_and_closes_failed_phase() {
        let bytes = vec![0_u8; 8];
        let ordinary =
            Snapshot::open(bytes.clone(), Limits::default()).map_err(|error| error.to_string());
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = Snapshot::open_bounded_profiled_with_cfb_observer(
            bytes,
            Limits::default(),
            TransactionLimits::default(),
            crate::package::SourceLimits::default(),
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        );

        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Error,
                },
            ]
        );
        assert!(matches!(profiled, Err(Error::Invalid(_))));
        assert_eq!(profiled.map_err(|error| error.to_string()), ordinary);
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Error,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_open_preflights_source_before_any_cfb_parse() {
        let source_limits = crate::package::SourceLimits::builder()
            .max_input_bytes(1)
            .expect("one-byte input limit is within the safety ceiling")
            .build()
            .expect("source limits");
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let result = Snapshot::open_bounded_profiled_with_source_limits(
            doc(&["alpha"]),
            Limits::default(),
            TransactionLimits::default(),
            source_limits,
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        );

        assert!(cfb_events.is_empty());
        assert!(matches!(result, Err(Error::Invalid(_))));
        assert!(events.is_empty());
    }

    #[test]
    fn editor_reopens_preflight_outer_input_before_container_parsing() {
        let mut source =
            Snapshot::open(doc_with_opaque_stream(&["alpha"]), Limits::default()).unwrap();
        let mut edit = source.edit().unwrap();
        let limits = crate::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        source.source_limits = limits;
        assert!(matches!(
            source.editor(),
            Err(Error::Invalid(crate::Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                maximum: 1,
                ..
            }))
        ));
        edit.source.source_limits = limits;
        // Invalid CFB bytes distinguish the input gate from container parsing.
        assert!(matches!(
            edit.reopen_editor(vec![0; 2]),
            Err(Error::Invalid(crate::Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: 2,
                maximum: 1
            }))
        ));
    }

    #[test]
    fn source_stream_preflight_precedes_strict_owner_capture() {
        let source_limits = crate::package::SourceLimits::builder()
            .max_entry_bytes(1)
            .expect("one-byte entry limit is within the safety ceiling")
            .build()
            .expect("source limits");
        let strict_limits = Limits {
            max_stream_size: 1,
            ..Limits::default()
        };
        let result = Snapshot::from_bytes_with_options(
            doc_with_opaque_stream(&["alpha"]),
            OpenOptions::default()
                .source_limits(source_limits)
                .revision_limits(strict_limits),
        );

        assert!(matches!(
            result,
            Err(Error::Invalid(crate::package::Error::Limit {
                kind: crate::package::SourceLimitKind::EntryBytes,
                observed,
                maximum: 1,
            })) if observed > 1
        ));
    }

    #[test]
    fn source_aggregate_preflight_precedes_strict_owner_capture() {
        let source_limits = crate::package::SourceLimits::builder()
            .max_aggregate_bytes(1)
            .expect("one-byte aggregate limit is within the safety ceiling")
            .build()
            .expect("source limits");
        let strict_limits = Limits {
            max_total_size: 1,
            ..Limits::default()
        };
        let result = Snapshot::from_bytes_with_options(
            doc_with_opaque_stream(&["alpha"]),
            OpenOptions::default()
                .source_limits(source_limits)
                .revision_limits(strict_limits),
        );

        assert!(matches!(
            result,
            Err(Error::Invalid(crate::package::Error::Limit {
                kind: crate::package::SourceLimitKind::AggregateBytes,
                observed,
                maximum: 1,
            })) if observed > 1
        ));
    }

    #[test]
    fn signed_fallback_preserves_original_entry_and_aggregate_limits() {
        let bytes = doc_with_signature_stream(&vec![0_u8; 128 * 1024]);
        let entry_limits = crate::package::SourceLimits::builder()
            .max_entry_bytes(64 * 1024)
            .expect("entry limit is within the safety ceiling")
            .build()
            .expect("entry source limits");
        assert!(matches!(
            Snapshot::from_bytes_with_options(
                bytes.clone(),
                OpenOptions::default().source_limits(entry_limits),
            ),
            Err(Error::Invalid(crate::package::Error::Limit {
                kind: crate::package::SourceLimitKind::EntryBytes,
                observed,
                maximum,
            })) if maximum == 64 * 1024 && observed > maximum
        ));

        let aggregate_limits = crate::package::SourceLimits::builder()
            .max_aggregate_bytes(64 * 1024)
            .expect("aggregate limit is within the safety ceiling")
            .build()
            .expect("aggregate source limits");
        assert!(matches!(
            Snapshot::from_bytes_with_options(
                bytes,
                OpenOptions::default().source_limits(aggregate_limits),
            ),
            Err(Error::Invalid(crate::package::Error::Limit {
                kind: crate::package::SourceLimitKind::AggregateBytes,
                observed,
                maximum,
            })) if maximum == 64 * 1024 && observed > maximum
        ));
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_signed_fallback_preserves_original_source_limits() {
        let bytes = doc_with_signature_stream(&vec![0_u8; 128 * 1024]);
        let source_limits = crate::package::SourceLimits::builder()
            .max_entry_bytes(64 * 1024)
            .expect("entry limit is within the safety ceiling")
            .build()
            .expect("entry source limits");
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let result = Snapshot::open_bounded_profiled_with_source_limits(
            bytes,
            Limits::default(),
            TransactionLimits::default(),
            source_limits,
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        );

        assert!(matches!(
            result,
            Err(Error::Invalid(crate::package::Error::Limit {
                kind: crate::package::SourceLimitKind::EntryBytes,
                observed,
                maximum,
            })) if maximum == 64 * 1024 && observed > maximum
        ));
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Error,
                },
            ]
        );
        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_source_stream_preflight_matches_ordinary_error_before_capture() {
        let source_limits = crate::package::SourceLimits::builder()
            .max_entry_bytes(1)
            .expect("one-byte entry limit is within the safety ceiling")
            .build()
            .expect("source limits");
        let strict_limits = Limits {
            max_stream_size: 1,
            ..Limits::default()
        };
        let ordinary = Snapshot::from_bytes_with_options(
            doc_with_opaque_stream(&["alpha"]),
            OpenOptions::default()
                .source_limits(source_limits)
                .revision_limits(strict_limits),
        )
        .map_err(|error| error.to_string());

        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = Snapshot::open_bounded_profiled_with_source_limits(
            doc_with_opaque_stream(&["alpha"]),
            strict_limits,
            TransactionLimits::default(),
            source_limits,
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        )
        .map_err(|error| error.to_string());

        assert_eq!(profiled, ordinary);
        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Error,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn retained_render_still_requires_independent_public_reader_validation() {
        let source = Snapshot::parse(&doc(&["alpha"])).expect("valid source");
        let mut edit = source.edit().expect("edit");
        // Test-only substitution: the strict owner accepts this package, but
        // its malformed RSID payload must fail the independent public reader.
        edit.editor = RevisionEditor::open(public_reader_failure_doc(), Limits::default())
            .expect("strict owner accepts the public-reader failure fixture");
        edit.editor.set_retained_render_limit(8 * 1024 * 1024);
        edit.replace_paragraph(Position::new(0), "replacement")
            .expect("common candidate validates");
        assert!(edit.retained_render_bytes().is_some());
        assert!(matches!(edit.commit(), Err(Error::Invalid(_))));
        assert_eq!(
            source.paragraphs(Projection::All).expect("source intact")[0].text(),
            "alpha"
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_open_public_reader_error_closes_only_public_phase() {
        let bytes = public_reader_failure_doc();
        assert!(RevisionEditor::open(bytes.clone(), Limits::default()).is_ok());
        let ordinary =
            Snapshot::open(bytes.clone(), Limits::default()).map_err(|error| error.to_string());
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = Snapshot::open_bounded_profiled_with_cfb_observer(
            bytes,
            Limits::default(),
            TransactionLimits::default(),
            crate::package::SourceLimits::default(),
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        )
        .map_err(|error| error.to_string());

        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
        assert_eq!(profiled, ordinary);
        assert!(profiled.is_err());
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::PublicReaderValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::PublicReaderValidation,
                    outcome: DiagnosticOutcome::Error,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_open_protected_dop_is_readable_and_balances_success_phases() {
        let bytes = protected_dop_doc();
        let ordinary =
            Snapshot::open(bytes.clone(), Limits::default()).map_err(|error| error.to_string());
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = Snapshot::open_bounded_profiled_with_cfb_observer(
            bytes,
            Limits::default(),
            TransactionLimits::default(),
            crate::package::SourceLimits::default(),
            ProtectionPolicy::default(),
            |event| events.push(event),
            |event| cfb_events.push(event),
        )
        .map_err(|error| error.to_string());

        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
        assert_eq!(profiled, ordinary);
        assert!(profiled.is_ok());
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::PublicReaderValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::PublicReaderValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::SourceRetention,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::SourceRetention,
                    outcome: DiagnosticOutcome::Success,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn failed_finish_phase_is_balanced_when_render_returns_error() {
        // The public Edit API validates every binary mutation before storing
        // it, so a renderer failure cannot be injected without relying on an
        // allocator or lower-level CFB failure. Exercise the same phase
        // boundary with the typed error returned by such a renderer.
        let mut events = Vec::new();
        let result: super::Result<()> = observe_phase(
            &mut |event| events.push(event),
            DiagnosticPhase::Finish,
            || Err(Error::PatchConflict),
        );

        assert!(result.is_err());
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::Finish,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::Finish,
                    outcome: DiagnosticOutcome::Error,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_changed_commit_matches_ordinary_commit() {
        let bytes = doc(&["alpha", "bravo"]);
        let ordinary_source = Snapshot::open(bytes.clone(), Limits::default()).expect("source");
        let mut ordinary_edit = ordinary_source.edit().expect("ordinary edit");
        ordinary_edit
            .replace_paragraph(Position::new(0), "omega")
            .expect("ordinary replacement");
        let ordinary = ordinary_edit.commit().expect("ordinary commit");

        let profiled_source = Snapshot::open(bytes, Limits::default()).expect("source");
        let mut profiled_edit = profiled_source.edit().expect("profiled edit");
        profiled_edit
            .replace_paragraph(Position::new(0), "omega")
            .expect("profiled replacement");
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = profiled_edit
            .commit_profiled_with_cfb_observer(
                |event| events.push(event),
                |event| cfb_events.push(event),
            )
            .expect("profiled commit");

        assert_eq!(
            cfb_events,
            vec![
                CfbParseEvent::Started,
                CfbParseEvent::Finished {
                    outcome: CfbParseOutcome::Success,
                },
            ]
        );
        assert_eq!(profiled, ordinary);
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::Finish,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::Finish,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::StrictOwnerValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::PublicReaderValidation,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::PublicReaderValidation,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::SourceRetention,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::SourceRetention,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::Patch,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::Patch,
                    outcome: DiagnosticOutcome::Success,
                },
            ]
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_exact_noop_commit_matches_ordinary_and_skips_reopen() {
        let bytes = doc(&["alpha", "bravo"]);
        let ordinary_source = Snapshot::open(bytes.clone(), Limits::default()).expect("source");
        let ordinary = ordinary_source
            .edit()
            .expect("ordinary edit")
            .commit()
            .expect("ordinary commit");

        let profiled_source = Snapshot::open(bytes, Limits::default()).expect("source");
        let mut profiled_edit = profiled_source.edit().expect("profiled edit");
        profiled_edit
            .replace_paragraph(Position::new(0), "alpha")
            .expect("same text replacement");
        let mut events = Vec::new();
        let mut cfb_events = Vec::new();
        let profiled = profiled_edit
            .commit_profiled_with_cfb_observer(
                |event| events.push(event),
                |event| cfb_events.push(event),
            )
            .expect("profiled no-op commit");

        assert!(cfb_events.is_empty());
        assert_eq!(profiled, ordinary);
        assert!(!profiled.diagnostics().changed());
        assert_eq!(
            events,
            vec![
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::Finish,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::Finish,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::ExactNoOp,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::ExactNoOp,
                    outcome: DiagnosticOutcome::Success,
                },
                DiagnosticEvent::Started {
                    phase: DiagnosticPhase::Patch,
                },
                DiagnosticEvent::Finished {
                    phase: DiagnosticPhase::Patch,
                    outcome: DiagnosticOutcome::Success,
                },
            ]
        );
    }

    #[test]
    fn same_shape_body_edit_is_reversible_and_source_checked() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo"])).expect("snapshot");
        assert_eq!(
            source
                .paragraphs(Projection::All)
                .expect("paragraphs")
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["alpha", "bravo"]
        );
        assert_eq!(
            source.paragraphs(Projection::All).expect("paragraphs")[0].position(),
            Position::new(0)
        );

        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("same shape edit");
        let commit = edit.commit().expect("commit");
        assert!(commit.diagnostics().changed());
        assert_eq!(
            commit
                .package()
                .paragraphs(Projection::All)
                .expect("changed paragraphs")[0]
                .text(),
            "omega"
        );

        let applied = commit.patch().apply(&source).expect("exact source applies");
        assert_eq!(applied, *commit.package());
        let restored = commit
            .patch()
            .inverse()
            .apply(&applied)
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());

        let other = Snapshot::open(doc(&["other"]), Limits::default()).expect("other source");
        assert!(matches!(
            commit.patch().apply(&other),
            Err(Error::PatchConflict)
        ));
    }

    #[test]
    fn patch_apply_revalidates_and_retains_intersected_source_limits() {
        let source_bytes = doc(&["alpha"]);
        let source = Snapshot::from_bytes(source_bytes.clone()).expect("source");
        let mut edit = source.edit().expect("edit");
        let replacement = "expanded ".repeat(4_096);
        edit.replace_paragraph(Position::new(0), &replacement)
            .expect("large replacement");
        let commit = edit.commit().expect("changed commit");
        assert!(commit.package().bytes().len() > source.bytes().len());

        let source_limit = crate::package::SourceLimits::builder()
            .max_input_bytes(source.bytes().len())
            .expect("source limit")
            .build()
            .expect("source limits");
        let strict_source = Snapshot::from_bytes_with_options(
            source_bytes.clone(),
            OpenOptions::default().source_limits(source_limit),
        )
        .expect("source fits strict package limit");
        assert!(matches!(
            commit.patch().apply(&strict_source),
            Err(Error::Invalid(crate::package::Error::Limit { .. }))
        ));

        let target_limit = crate::package::SourceLimits::builder()
            .max_input_bytes(commit.package().bytes().len())
            .expect("target limit")
            .build()
            .expect("target limits");
        let bounded_source = Snapshot::from_bytes_with_options(
            source_bytes,
            OpenOptions::default().source_limits(target_limit),
        )
        .expect("source fits target package limit");
        let applied = commit
            .patch()
            .apply(&bounded_source)
            .expect("target fits retained package limit");
        assert_eq!(applied.source_limits(), target_limit);
        assert_eq!(applied.bytes(), commit.package().bytes());
    }

    #[test]
    fn auxiliary_tables_publish_through_body_commit_and_reopen() {
        let source =
            Snapshot::open(doc_with_auxiliary_tables(), Limits::default()).expect("auxiliary DOC");
        let selection = source
            .saved_selection()
            .expect("saved-selection read")
            .expect("saved-selection record");
        assert_eq!(selection.cp_first(), 4);
        let dofr = source
            .dofr_records()
            .expect("Dofr read")
            .expect("Dofr records");
        assert_eq!(dofr.len(), 2);

        let mut no_op = source.edit().expect("no-op body edit");
        assert!(
            !no_op
                .edit_saved_selection(|_| Ok(()))
                .expect("no-op Selsf edit")
        );
        let no_op_commit = no_op.commit().expect("no-op auxiliary commit");
        assert!(!no_op_commit.diagnostics().changed());
        assert_eq!(no_op_commit.package().bytes(), source.bytes());

        let mut edit = source.edit().expect("body edit");
        assert!(
            edit.edit_saved_selection(|transaction| {
                transaction.set_cp_anchor(5)?.set_range(5, 9).map(|_| ())
            })
            .expect("Selsf edit")
        );
        let mut replacement = dofr.get(1).expect("frame record").bytes().to_vec();
        replacement[20..24].copy_from_slice(&1u32.to_le_bytes());
        assert!(
            edit.replace_dofr_record(1, &replacement)
                .expect("Dofr edit")
        );
        let commit = edit.commit().expect("auxiliary commit");
        assert!(commit.diagnostics().changed());
        assert!(commit.patch().has_auxiliary_changes());
        assert!(commit.patch().to_durable(patch_limits()).is_err());
        let published = commit.package();
        let package = open_object(published.finish(), Targets::default(), Limits::default())
            .expect("published package");
        assert_eq!(
            package
                .stream(&["OpaqueVendorData".to_string()])
                .expect("opaque stream"),
            b"untouched auxiliary stream"
        );
        assert_eq!(
            published
                .saved_selection()
                .expect("published Selsf")
                .expect("published record")
                .cp_first(),
            5
        );
        let published_dofr = published
            .dofr_records()
            .expect("published Dofr")
            .expect("published records");
        let DofrPayload::Frame(frame) = published_dofr
            .get(1)
            .expect("published frame")
            .payload()
            .expect("published frame payload")
        else {
            panic!("expected frame payload");
        };
        assert_eq!(frame.frame_kind(), DofrFrameKind::FrameSet);

        let reopened =
            Snapshot::open(published.finish(), Limits::default()).expect("reopen published DOC");
        assert_eq!(
            reopened
                .saved_selection()
                .expect("reopened Selsf")
                .expect("reopened record")
                .cp_first(),
            5
        );
        let reopened_dofr = reopened
            .dofr_records()
            .expect("reopened Dofr")
            .expect("reopened records");
        let DofrPayload::Frame(frame) = reopened_dofr
            .get(1)
            .expect("reopened frame")
            .payload()
            .expect("reopened payload")
        else {
            panic!("expected reopened frame payload");
        };
        assert_eq!(frame.frame_kind(), DofrFrameKind::FrameSet);

        let inverse = commit.patch().inverse();
        let reverted = inverse.apply(published).expect("inverse auxiliary patch");
        assert_eq!(reverted.bytes(), source.bytes());

        let mut stale = published.edit().expect("stale edit");
        assert!(
            stale
                .edit_saved_selection(|transaction| {
                    transaction.set_cp_anchor(6)?.set_range(6, 10).map(|_| ())
                })
                .expect("second selection edit")
        );
        let stale_commit = stale.commit().expect("second edit");
        let mut mismatched = stale_commit.package().edit().expect("mismatched edit");
        let source_patch = selection
            .edit()
            .commit()
            .expect("no-op selection patch")
            .patch()
            .clone();
        assert!(
            mismatched
                .apply_saved_selection_patch(&source_patch)
                .is_err()
        );
    }

    #[test]
    fn patch_construction_and_same_lineage_apply_keep_fingerprint_lazy() {
        let source = Snapshot::from_bytes(doc(&["alpha", "bravo"]))
            .expect("source should open without fingerprinting");
        assert!(source.fingerprint_cache.get().is_none());

        let mut edit = source.edit().expect("edit should open");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("replacement should stage");
        let commit = edit.commit().expect("commit should reopen");
        let patch = commit.patch();
        assert!(patch.before.fingerprint_cache.get().is_none());
        assert!(patch.after.fingerprint_cache.get().is_none());

        let applied = patch
            .apply(&source)
            .expect("same-lineage application should use allocation identity");
        assert_eq!(applied.bytes(), commit.package().bytes());
        assert!(source.fingerprint_cache.get().is_none());
        assert!(patch.before.fingerprint_cache.get().is_none());
        assert!(patch.after.fingerprint_cache.get().is_none());

        let source_fingerprint = patch.source_fingerprint();
        let target_fingerprint = patch.target_fingerprint();
        assert_eq!(source_fingerprint, fingerprint(patch.before.bytes()));
        assert_eq!(target_fingerprint, fingerprint(patch.after.bytes()));
        assert_eq!(fingerprint(b""), 0xcbf2_9ce4_8422_2325);
        assert_eq!(fingerprint(b"hello"), 0xa430_d846_80aa_bd0b);
        assert!(patch.before.fingerprint_cache.get().is_some());
        assert!(patch.after.fingerprint_cache.get().is_some());
    }

    #[test]
    fn independently_reopened_identical_source_uses_fingerprint_then_exact_bytes() {
        let source = Snapshot::from_bytes(doc(&["alpha", "bravo"])).expect("source");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("replacement");
        let commit = edit.commit().expect("commit");
        let patch = commit.patch();
        let reopened = Snapshot::from_bytes(source.finish()).expect("reopened source");
        assert!(!Arc::ptr_eq(
            &source.bytes_shared(),
            &reopened.bytes_shared()
        ));
        assert_eq!(reopened.fingerprint(), patch.source_fingerprint());

        let applied = patch
            .apply(&reopened)
            .expect("identical independently reopened source should apply");
        assert_eq!(applied.bytes(), commit.package().bytes());

        let mut different_limits = Limits::default();
        different_limits.max_streams = Limits::default().max_streams - 1;
        let differently_limited = Snapshot::open(source.finish(), different_limits)
            .expect("same source should open under a different limit profile");
        assert_eq!(differently_limited.bytes(), source.bytes());
        assert_eq!(
            differently_limited.fingerprint(),
            patch.source_fingerprint()
        );
        assert_eq!(
            patch
                .apply(&differently_limited)
                .expect("byte-identical source with different limits should apply")
                .bytes(),
            commit.package().bytes()
        );
    }

    #[test]
    fn snapshot_clone_copies_initialized_fingerprint_without_extra_cache_allocation() {
        let source = Snapshot::from_bytes(doc(&["alpha"])).expect("source");
        let uninitialized_clone = source.clone();
        assert!(source.fingerprint_cache.get().is_none());
        assert!(uninitialized_clone.fingerprint_cache.get().is_none());

        let expected = source.fingerprint();
        let initialized_clone = source.clone();
        assert_eq!(initialized_clone.fingerprint_cache.get(), Some(&expected));
        assert_eq!(initialized_clone.fingerprint(), expected);
    }

    #[test]
    fn equal_fingerprint_collision_still_requires_exact_source_bytes() {
        let source = Snapshot::from_bytes(doc(&["alpha", "bravo"])).expect("source");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("replacement");
        let commit = edit.commit().expect("commit");
        let patch = commit.patch();
        let source_fingerprint = patch.source_fingerprint();

        let mut collision = Snapshot::from_bytes(doc(&["other"])).expect("collision source");
        assert_ne!(fingerprint(collision.bytes()), source_fingerprint);
        let cache = OnceLock::new();
        cache
            .set(source_fingerprint)
            .expect("test cache should be initialized once");
        collision.fingerprint_cache = cache;
        assert_eq!(collision.fingerprint(), source_fingerprint);
        assert_ne!(collision.bytes(), patch.before.bytes());
        assert!(matches!(patch.apply(&collision), Err(Error::PatchConflict)));
    }

    #[test]
    fn empty_and_same_text_commits_share_exact_source_allocations() {
        let source =
            Snapshot::from_bytes(doc(&["alpha", "bravo"])).expect("owned source should open");
        let source_bytes = source.bytes_shared();

        let empty = source.edit().expect("empty edit").commit().expect("no-op");
        assert!(!empty.diagnostics().changed());
        assert!(empty.patch().is_noop());
        assert_eq!(empty.patch().changes().len(), 0);
        assert!(Arc::ptr_eq(&source_bytes, &empty.package().bytes_shared()));
        let applied = empty.patch().apply(&source).expect("no-op patch applies");
        assert!(Arc::ptr_eq(&source_bytes, &applied.bytes_shared()));
        let restored = empty
            .patch()
            .inverse()
            .apply(&applied)
            .expect("no-op inverse applies");
        assert!(Arc::ptr_eq(&source_bytes, &restored.bytes_shared()));

        let mut same_text = source.edit().expect("same-text edit");
        same_text
            .replace_paragraph(Position::new(0), "alpha")
            .expect("same text is accepted");
        let same_text = same_text.commit().expect("same-text no-op");
        assert!(!same_text.diagnostics().changed());
        assert_eq!(same_text.patch().changes().len(), 0);
        assert!(Arc::ptr_eq(
            &source_bytes,
            &same_text.package().bytes_shared()
        ));
    }

    #[test]
    fn empty_body_paragraph_insertion_round_trips_and_inverts() {
        let source = Snapshot::parse(&doc(&[""])).expect("empty paragraph source");
        assert_eq!(
            source
                .paragraphs(Projection::All)
                .expect("source paragraphs")[0]
                .text(),
            ""
        );
        let mut edit = source.edit().expect("empty paragraph edit");
        edit.replace_paragraph(Position::new(0), "inserted")
            .expect("empty paragraph accepts insertion");
        let commit = edit.commit().expect("empty paragraph commit");
        assert_eq!(
            commit
                .package()
                .paragraphs(Projection::All)
                .expect("changed paragraphs")[0]
                .text(),
            "inserted"
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.package())
                .expect("empty insertion inverse"),
            source
        );
    }

    #[test]
    fn formatted_empty_body_paragraph_copies_paragraph_mark_character_format() {
        let source = Snapshot::parse(&formatted_empty_doc()).expect("formatted empty source");
        let mut edit = source.edit().expect("formatted empty edit");
        edit.replace_paragraph(Position::new(0), "formatted")
            .expect("formatted empty paragraph accepts insertion");
        let commit = edit.commit().expect("formatted empty commit");
        let mut package = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("formatted candidate opens");
        let document = package.document().expect("formatted candidate document");
        let paragraph = document
            .paragraphs()
            .expect("formatted candidate paragraphs")
            .into_iter()
            .next()
            .expect("formatted candidate paragraph");
        let run = paragraph
            .runs()
            .expect("formatted candidate runs")
            .into_iter()
            .next()
            .expect("formatted candidate run");
        assert_eq!(run.bold(), Some(true));
        assert_eq!(run.italic(), Some(true));
    }

    #[test]
    fn empty_body_paragraph_does_not_address_a_table_cell() {
        let source = Snapshot::parse(&empty_body_and_table_doc()).expect("body/table source");
        assert_eq!(
            source
                .paragraphs(Projection::All)
                .expect("ordinary paragraphs")
                .len(),
            2
        );
        assert_eq!(source.table_cells().expect("table cells")[0].text(), "");
        let mut edit = source.edit().expect("body/table edit");
        edit.replace_paragraph(Position::new(0), "body")
            .expect("ordinary empty body paragraph accepts insertion");
        assert!(matches!(
            edit.replace_paragraph(Position::new(2), "table"),
            Err(Error::Refused(Refusal::ParagraphNotFound))
        ));
        let commit = edit.commit().expect("body-only commit");
        let paragraphs = commit
            .package()
            .paragraphs(Projection::All)
            .expect("changed ordinary paragraphs");
        assert_eq!(
            paragraphs
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["body", ""]
        );
        assert_eq!(
            commit.package().table_cells().expect("changed table cells")[0].text(),
            ""
        );
    }

    #[test]
    fn empty_body_insertion_updates_bookmark_cp_tables() {
        let mut writer = Builder::new();
        writer.add_paragraph("").expect("empty paragraph");
        writer.add_bookmark(BookmarkEntry::new("Empty", 0, 0));
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .expect("DOC fixture builds")
            .write_to(&mut output)
            .expect("bookmark DOC");
        let source =
            Snapshot::parse(&normalize_word2002_dop(output.into_inner())).expect("bookmark source");
        let mut edit = source.edit().expect("bookmark edit");
        edit.replace_paragraph(Position::new(0), "inside")
            .expect("bookmark-protected empty insertion");
        let commit = edit.commit().expect("bookmark commit");
        let mut package = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("bookmark candidate opens");
        let bookmark = package
            .document()
            .expect("bookmark candidate document")
            .bookmarks()
            .expect("bookmark candidate bookmarks")
            .into_iter()
            .next()
            .expect("bookmark");
        assert_eq!((bookmark.start, bookmark.end), (6, 6));
    }

    #[test]
    fn exact_text_noop_bypasses_replacement_quota() {
        let source = Snapshot::open_bounded(
            doc(&["already stored"]),
            Limits::default(),
            TransactionLimits::new(1, 0, 0),
        )
        .expect("zero replacement quota source");
        let mut edit = source.edit().expect("zero replacement quota edit");
        edit.replace_paragraph(Position::new(0), "already stored")
            .expect("exact text no-op bypasses quota");
        let commit = edit.commit().expect("exact text no-op commit");
        assert!(!commit.diagnostics().changed());
        assert!(Arc::ptr_eq(
            &source.bytes_shared(),
            &commit.package().bytes_shared()
        ));
    }

    #[test]
    fn patch_and_commit_debug_redact_authored_text() {
        let source = Snapshot::parse(&doc(&["private-before"])).expect("debug source");
        let mut edit = source.edit().expect("debug edit");
        edit.replace_paragraph(Position::new(0), "private-after")
            .expect("debug replacement");
        let commit = edit.commit().expect("debug commit");
        let debug = format!("{commit:?} {:?}", commit.patch());
        assert!(!debug.contains("private-before"));
        assert!(!debug.contains("private-after"));
        let change = commit.patch().changes().next().expect("debug change");
        assert!(!format!("{change:?}").contains("private-before"));
        assert!(!format!("{change:?}").contains("private-after"));
    }

    #[test]
    fn paragraph_commit_preserves_every_unmodeled_stream_payload() {
        let source = Snapshot::from_bytes(doc_with_opaque_stream(&["alpha", "bravo"]))
            .expect("source with opaque stream should open");
        let mut edit = source.edit().expect("edit should open");
        edit.replace_paragraph(Position::new(0), "alpha expanded")
            .expect("paragraph replacement should stage");
        let commit = edit.commit().expect("changed package should reopen");

        let mut before = OleFile::open(
            Cursor::new(source.bytes()),
            litchi_cfb::OleFileLimits::default(),
        )
        .expect("source CFB should open");
        let mut after = OleFile::open(
            Cursor::new(commit.package().bytes()),
            litchi_cfb::OleFileLimits::default(),
        )
        .expect("changed CFB should open");
        let before_paths = before.list_streams();
        assert_eq!(after.list_streams(), before_paths);
        for path in before_paths {
            if matches!(
                path.last().map(String::as_str),
                Some("WordDocument" | "0Table" | "1Table")
            ) {
                continue;
            }
            let path = path.iter().map(String::as_str).collect::<Vec<_>>();
            assert_eq!(
                before
                    .open_stream(&path)
                    .expect("source stream should read"),
                after
                    .open_stream(&path)
                    .expect("changed stream should read"),
                "unmodeled stream payload changed at {path:?}"
            );
        }
    }

    #[test]
    fn shared_physical_validation_preserves_candidate_semantics_and_opaque_streams() {
        let source = doc_with_opaque_stream(&["alpha", "bravo"]);
        let (_strict_editor, mut ole) =
            RevisionEditor::open_with_ole_file(source.clone(), Limits::default())
                .expect("physical DOC context should open");
        crate::raw::Reader::<Cursor<Vec<u8>>>::validate_source_len(
            source.len(),
            crate::package::SourceLimits::default(),
        )
        .expect("public package preflight should pass");
        crate::raw::Reader::validate_ole_file(&mut ole, crate::package::SourceLimits::default())
            .expect("public reader should validate the shared context");

        let source = Snapshot::from_bytes(source).expect("source should open");
        let mut edit = source.edit().expect("edit should open");
        edit.replace_paragraph(Position::new(0), "alpha expanded")
            .expect("paragraph replacement should stage");
        let commit = edit.commit().expect("changed candidate should reopen");
        assert_eq!(
            commit.package().paragraphs(Projection::All).unwrap()[0].text(),
            "alpha expanded"
        );
        let mut public = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("changed candidate should pass the public package gate");
        assert_eq!(
            public
                .document()
                .expect("changed candidate should pass public document parsing")
                .paragraphs()
                .expect("changed candidate paragraphs")
                .first()
                .expect("changed candidate first paragraph")
                .text()
                .expect("changed candidate paragraph text"),
            "alpha expanded"
        );
        let package = open_object(
            commit.package().finish(),
            Targets::default(),
            Limits::default(),
        )
        .expect("candidate package should open");
        assert_eq!(
            package
                .stream(&["OpaqueVendorData".to_string()])
                .expect("opaque stream should survive"),
            b"untouched"
        );
    }

    #[test]
    fn length_changes_publish_while_structural_changes_are_refused() {
        let source = Snapshot::parse(&doc(&["alpha"])).expect("snapshot");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "a much longer paragraph")
            .expect("length-changing edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "a\rpha"),
            Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent
            ))
        ));
        assert!(matches!(
            edit.replace_paragraph(Position::new(1), "alpha"),
            Err(Error::Refused(Refusal::ParagraphNotFound))
        ));
        let commit = edit.commit().expect("changed commit");
        assert!(commit.diagnostics().changed());
        let paragraphs = commit
            .package()
            .paragraphs(Projection::All)
            .expect("changed paragraphs");
        assert_eq!(paragraphs[0].text(), "a much longer paragraph");
    }

    #[test]
    fn accepted_and_rejected_projections_hide_text_revisions() {
        let mut writer = Builder::new();
        writer
            .add_paragraph_runs(
                vec![
                    ("kept ".to_string(), CharacterFormatting::default()),
                    (
                        "old".to_string(),
                        CharacterFormatting {
                            deletion_revision: Some(TextRevision::new("Reviewer")),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (" new".to_string(), CharacterFormatting::default()),
                ],
                ParagraphFormatting::default(),
            )
            .expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("fixture DOC");
        let bytes = normalize_word2002_dop(output.into_inner());
        let snapshot = Snapshot::parse(&bytes).expect("snapshot");
        assert_eq!(
            snapshot.paragraphs(Projection::Accepted).expect("accepted")[0].text(),
            "kept  new"
        );
        assert_eq!(
            snapshot.paragraphs(Projection::Rejected).expect("rejected")[0].text(),
            "kept old new"
        );
    }

    #[test]
    fn multi_paragraph_text_and_bold_commit_reopens_and_inverts() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo", "charlie"])).expect("snapshot");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "alpha expanded")
            .expect("grow first paragraph");
        edit.replace_paragraph(Position::new(2), "c")
            .expect("shrink third paragraph");
        edit.set_paragraph_bold(Position::new(0), true)
            .expect("bold first paragraph");
        let commit = edit.commit().expect("commit and full reopen");
        let texts = commit
            .package()
            .paragraphs(Projection::All)
            .expect("readback")
            .into_iter()
            .map(|paragraph| paragraph.text().to_string())
            .collect::<Vec<_>>();
        assert_eq!(texts, ["alpha expanded", "bravo", "c"]);

        let mut package = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("CFB reopens");
        let document = package.document().expect("DOC reopens");
        let paragraphs = document.paragraphs().expect("paragraphs read back");
        assert_eq!(paragraphs[0].runs().expect("runs")[0].bold(), Some(true));

        let restored = commit
            .patch()
            .inverse()
            .apply(commit.package())
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());
    }

    #[test]
    fn deferred_body_formatting_flushes_before_table_edits_and_keeps_failures_atomic() {
        let source = Snapshot::parse(&structured_doc()).expect("structured snapshot");
        let mut edit = source.edit().expect("edit");
        edit.set_paragraph_bold(Position::new(0), true)
            .expect("body bold");
        edit.set_character_property(
            TextTarget::body_paragraph(Position::new(0)),
            CharacterProperty::Italic,
            true,
        )
        .expect("body italic");

        // A failed selector must not discard or partially publish the pending
        // CHPX changes. Rendering a clone is the same output check used by
        // the source-backed publication path, while leaving `edit` staged.
        let before_refusal = edit.editor.clone().finish().expect("pending output");
        assert!(matches!(
            edit.set_paragraph_bold(Position::new(99), true),
            Err(Error::Refused(Refusal::TargetNotFound))
        ));
        assert_eq!(
            edit.editor.clone().finish().expect("output after refusal"),
            before_refusal
        );

        // The table-cell edit is a serialized-stream dependency. It must see
        // the deferred body format and retain both changes in one final
        // candidate before the package commit.
        edit.replace_text(TextTarget::TableCell(Position::new(0)), "Cell!")
            .expect("table cell");
        let commit = edit.commit().expect("mixed commit");
        assert_eq!(
            commit.package().table_cells().expect("table cells")[0].text(),
            "Cell!"
        );
        let body = commit
            .package()
            .paragraphs(Projection::All)
            .expect("body paragraphs");
        assert_eq!(body[0].text(), "Body");
        let mut package = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("DOC reopens");
        let document = package.document().expect("document");
        let runs = document.paragraphs().expect("paragraphs")[0]
            .runs()
            .expect("runs");
        assert_eq!(runs[0].bold(), Some(true));
        assert_eq!(runs[0].italic(), Some(true));

        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.package())
            .expect("inverse");
        assert_eq!(inverse.bytes(), source.bytes());

        let noop = source.edit().expect("no-op edit").commit().expect("no-op");
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.package().bytes(), source.bytes());
    }

    #[test]
    fn durable_patch_composition_and_history_are_bounded() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo"])).expect("snapshot");
        let composition_limits = CompositionLimits::new(4, 2, 8, 8);
        let text = source
            .prepare_replace(
                composition_limits,
                "replace-alpha",
                Position::new(0),
                "alpha grows",
            )
            .expect("prepare text");
        let bold = source
            .prepare_bold(composition_limits, "bold-alpha", Position::new(0), true)
            .expect("prepare bold");
        let conflict = source
            .prepare_replace(
                composition_limits,
                "replace-alpha-again",
                Position::new(0),
                "other",
            )
            .expect("prepare conflicting text");
        let mut composition = source.compose(composition_limits);
        composition.join(text).expect("text joins");
        composition.join(bold).expect("disjoint bold joins");
        let failure = composition
            .join(conflict)
            .expect_err("same facet conflicts");
        assert!(matches!(failure.failure(), SubEditJoinFailure::Overlap(_)));

        let commit = composition.commit().expect("composition commits");
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("durable patch");
        let wire = durable.to_deterministic_json().expect("canonical JSON");
        let decoded = Patch::<Reversible>::from_deterministic_json(&wire, patch_limits())
            .expect("durable decode");
        assert_eq!(
            decoded.to_deterministic_json().expect("durable re-encode"),
            wire
        );
        assert_eq!(
            source.apply_durable(&decoded).expect("durable apply"),
            *commit.package()
        );
        let semantic_inverse = commit
            .package()
            .apply_durable(&decoded.inverse())
            .expect("durable inverse");
        assert_eq!(
            semantic_inverse
                .paragraphs(Projection::All)
                .expect("inverse paragraphs")
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["alpha", "bravo"]
        );
        let mut inverse_package =
            crate::raw::Reader::from_reader(Cursor::new(semantic_inverse.finish()))
                .expect("durable inverse CFB reopens");
        let inverse_document = inverse_package
            .document()
            .expect("durable inverse DOC reopens");
        assert_eq!(
            inverse_document.paragraphs().expect("inverse paragraphs")[0]
                .runs()
                .expect("inverse runs")[0]
                .bold(),
            None
        );

        let wire_bytes = u64::try_from(wire.len()).expect("wire length fits u64");
        let mut history = source.history(HistoryLimits::new(1, wire_bytes));
        history
            .record(commit.package().clone(), wire_bytes)
            .expect("record history");
        assert!(history.undo());
        assert_eq!(history.current(), &source);
        assert!(history.redo());
        assert_eq!(history.current(), commit.package());

        let mut too_small = source.history(HistoryLimits::new(1, wire_bytes - 1));
        assert!(
            too_small
                .record(commit.package().clone(), wire_bytes)
                .is_err()
        );
        assert_eq!(too_small.current(), &source);
    }

    #[test]
    fn transaction_limits_fail_before_mutation() {
        let source = Snapshot::open_bounded(
            doc(&["alpha", "bravo"]),
            Limits::default(),
            TransactionLimits::new(1, 5, 5),
        )
        .expect("bounded source");
        let mut edit = source.edit().expect("edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "too long"),
            Err(Error::Refused(Refusal::ReplacementLimit { .. }))
        ));
        edit.replace_paragraph(Position::new(0), "short")
            .expect("within bounds");
        assert!(matches!(
            edit.set_paragraph_bold(Position::new(1), true),
            Err(Error::Refused(Refusal::OperationLimit { .. }))
        ));
    }

    #[test]
    fn stories_cells_fields_format_merge_transfer_and_durable_reopen() {
        let source = Snapshot::parse(&structured_doc()).expect("structured snapshot");
        let header = source
            .story_paragraphs(Story::Header)
            .expect("header paragraphs")
            .into_iter()
            .find(|item| item.text() == "Header")
            .expect("odd header");
        let cell = source
            .table_cells()
            .expect("table cells")
            .into_iter()
            .find(|item| item.text() == "Cell")
            .expect("simple cell");
        let field = source
            .field_results()
            .expect("field results")
            .into_iter()
            .find(|item| item.text() == "Link")
            .expect("simple field result");

        let mut refused_story_resize = source.edit().expect("story resize refusal");
        assert!(matches!(
            refused_story_resize.replace_text(header.target(), "Longer header"),
            Err(Error::Refused(Refusal::StoryLengthChange {
                story: Story::Header,
                ..
            }))
        ));

        let mut edit = source.edit().expect("edit");
        edit.replace_text(header.target(), "Banner")
            .expect("same-length header edit");
        edit.replace_text(cell.target(), "Cell expanded")
            .expect("length-changing cell edit");
        edit.replace_text(field.target(), "Result")
            .expect("length-changing field-result edit");
        edit.set_character_property(cell.target(), CharacterProperty::Italic, true)
            .expect("cell italic");
        edit.set_character_property(field.target(), CharacterProperty::Underline, true)
            .expect("field underline");
        let commit = edit.commit().expect("full reopen commit");
        assert_eq!(
            commit
                .package()
                .story_paragraphs(Story::Header)
                .expect("changed header")
                .into_iter()
                .find(|item| item.target() == header.target())
                .expect("same header selector")
                .text(),
            "Banner"
        );
        assert_eq!(
            commit.package().table_cells().expect("changed cells")[0].text(),
            "Cell expanded"
        );
        assert_eq!(
            commit.package().field_results().expect("changed fields")[0].text(),
            "Result"
        );
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("durable patch");
        let replay = source.apply_durable(&durable).expect("durable replay");
        assert_eq!(replay, *commit.package());
        let inverse = replay
            .apply_durable(&durable.inverse())
            .expect("durable inverse");
        assert_eq!(
            inverse.table_cells().expect("inverse cells")[0].text(),
            "Cell"
        );

        let mut left = source.edit().expect("left");
        left.replace_text(cell.target(), "Left cell")
            .expect("left edit");
        let left = left.commit().expect("left commit");
        let mut right = source.edit().expect("right");
        right
            .replace_text(field.target(), "Right result")
            .expect("right edit");
        let right = right.commit().expect("right commit");
        let merged = source
            .plan_three_way(left.patch(), right.patch())
            .expect("three-way plan");
        assert!(merged.is_conflict_free());
        let merged = merged.commit().expect("merged commit");
        assert_eq!(
            merged.package().table_cells().expect("merged cells")[0].text(),
            "Left cell"
        );
        assert_eq!(
            merged.package().field_results().expect("merged fields")[0].text(),
            "Right result"
        );

        let mut competing = source.edit().expect("competing");
        competing
            .replace_text(cell.target(), "Competing cell")
            .expect("competing edit");
        let competing = competing.commit().expect("competing commit");
        let conflicted = source
            .plan_three_way(left.patch(), competing.patch())
            .expect("conflict plan");
        assert_eq!(conflicted.conflicts().len(), 1);
        assert!(matches!(conflicted.commit(), Err(Error::Conflict)));

        let donor = Snapshot::parse(&doc(&["Donor"])).expect("donor");
        let transfer = source
            .plan_text_transfer_from(
                &donor,
                TextTarget::body_paragraph(Position::new(0)),
                TextTarget::body_paragraph(Position::new(0)),
            )
            .expect("text transfer plan");
        let mut transfer_edit = source.edit().expect("transfer edit");
        transfer_edit
            .apply_transfer(&transfer)
            .expect("apply transfer");
        assert_eq!(
            transfer_edit
                .commit()
                .expect("transfer commit")
                .package()
                .paragraphs(Projection::All)
                .expect("transferred body")[0]
                .text(),
            "Donor"
        );
        let mut stale_receiver = Snapshot::parse(&doc(&["different"]))
            .expect("stale receiver")
            .edit()
            .expect("stale edit");
        assert!(matches!(
            stale_receiver.apply_transfer(&transfer),
            Err(Error::PatchConflict)
        ));
    }

    #[test]
    fn embedded_display_metadata_uses_root_patch_and_durable_inverse() {
        let source = Snapshot::parse(&embedded_doc()).expect("embedded snapshot");
        let before = crate::embedded_object::Snapshot::open(source.finish(), Limits::default())
            .expect("embedded owner snapshot")
            .inventory()
            .expect("embedded inventory")
            .get(77)
            .and_then(|entry| entry.metadata().obj_info())
            .expect("typed ObjInfo")
            .display_as_icon;

        let mut edit = source.edit().expect("root edit");
        edit.set_embedded_display_as_icon(77, !before)
            .expect("resource metadata edit");
        let commit = edit.commit().expect("resource root commit");
        let mut package = crate::raw::Reader::from_reader(Cursor::new(commit.package().finish()))
            .expect("resource CFB reopens");
        package.document().expect("resource DOC reopens");
        let changed =
            crate::embedded_object::Snapshot::open(commit.package().finish(), Limits::default())
                .expect("changed embedded owner snapshot")
                .inventory()
                .expect("changed embedded inventory")
                .get(77)
                .and_then(|entry| entry.metadata().obj_info())
                .expect("changed typed ObjInfo")
                .display_as_icon;
        assert_eq!(changed, !before);
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("resource durable patch");
        assert_eq!(
            source.apply_durable(&durable).expect("resource replay"),
            *commit.package()
        );
        assert_eq!(
            commit
                .package()
                .apply_durable(&durable.inverse())
                .expect("resource inverse"),
            source
        );
    }

    #[test]
    fn embedded_transfer_closes_field_preview_storage_merge_and_history() {
        let donor = Snapshot::parse(&embedded_doc()).expect("embedded donor");
        let receiver = Snapshot::parse(&doc(&["receiver"])).expect("receiver");
        let plan = receiver
            .plan_embedded_transfer_from(&donor, 77, 88)
            .expect("dependency-closed transfer plan");
        assert_eq!(plan.storage_id(), 88);
        assert!(plan.compound_size() > 0);
        assert!(plan.preview_size() >= 12);

        let mut transfer = receiver.edit().expect("transfer edit");
        transfer
            .apply_embedded_transfer(&plan)
            .expect("embedded transfer");
        let transfer = transfer.commit().expect("transfer commit");
        assert!(
            transfer
                .package()
                .embedded_objects()
                .expect("transferred inventory")
                .get(88)
                .is_some()
        );
        let exact_inverse = transfer
            .patch()
            .inverse()
            .apply(transfer.package())
            .expect("exact in-memory inverse");
        assert_eq!(exact_inverse.bytes(), receiver.bytes());

        let durable = transfer
            .patch()
            .to_durable(resource_patch_limits())
            .expect("resource durable patch");
        assert_eq!(durable.blobs().len(), 2);
        let wire = durable
            .to_deterministic_json()
            .expect("resource canonical JSON");
        let decoded = Patch::<Reversible>::from_deterministic_json(&wire, resource_patch_limits())
            .expect("resource durable decode");
        let replay = receiver
            .apply_durable(&decoded)
            .expect("resource durable replay");
        assert!(
            replay
                .embedded_objects()
                .expect("replayed inventory")
                .get(88)
                .is_some()
        );
        let semantic_inverse = replay
            .apply_durable(&decoded.inverse())
            .expect("resource durable inverse");
        assert!(
            semantic_inverse
                .embedded_objects()
                .expect("inverse inventory")
                .is_empty()
        );

        let mut text = receiver.edit().expect("text edit");
        text.replace_paragraph(Position::new(0), "receiver changed")
            .expect("disjoint text change");
        let text = text.commit().expect("text commit");
        let merged = receiver
            .plan_three_way(transfer.patch(), text.patch())
            .expect("resource/text three-way plan");
        assert!(merged.is_conflict_free());
        let merged = merged.commit().expect("resource/text merge");
        assert!(
            merged
                .package()
                .embedded_objects()
                .expect("merged inventory")
                .get(88)
                .is_some()
        );
        assert_eq!(
            merged
                .package()
                .paragraphs(Projection::All)
                .expect("merged text")[0]
                .text(),
            "receiver changed"
        );

        let wire_bytes = u64::try_from(wire.len()).expect("wire length fits u64");
        let mut history = receiver.history(HistoryLimits::new(1, wire_bytes));
        history
            .record(transfer.package().clone(), wire_bytes)
            .expect("resource history record");
        assert!(history.undo());
        assert_eq!(history.current(), &receiver);
        assert!(history.redo());
        assert_eq!(history.current(), transfer.package());

        assert!(matches!(
            transfer
                .package()
                .plan_embedded_transfer_from(&donor, 77, 88),
            Err(Error::Refused(Refusal::ResourceCollision {
                storage_id: 88
            }))
        ));
    }

    #[test]
    fn drawing_and_preview_dependencies_have_precise_atomic_refusals() {
        let embedded = Snapshot::parse(&embedded_doc()).expect("embedded snapshot");
        let preview = embedded
            .field_results()
            .expect("embedded result")
            .into_iter()
            .find(|item| item.text().contains('\u{1}'))
            .expect("embedded preview character");
        let mut edit = embedded.edit().expect("preview edit");
        assert!(matches!(
            edit.replace_text(preview.target(), "replacement"),
            Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::InlinePictureOrObjectPreview
            }))
        ));

        let mut writer = Builder::new();
        writer.add_paragraph("body").expect("body");
        writer
            .insert_floating_shape(
                crate::writer::Shape::new(crate::writer::Kind::Rectangle, 720, 360).expect("shape"),
                FloatingPosition::new(120, 240),
            )
            .expect("floating shape");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("drawing DOC");
        let bytes = normalize_word2002_dop(output.into_inner());
        let drawing = Snapshot::parse(&bytes).expect("drawing snapshot");
        let anchor = drawing
            .paragraphs(Projection::All)
            .expect("drawing paragraphs")
            .into_iter()
            .find(|paragraph| paragraph.text().contains('\u{8}'))
            .expect("floating anchor");
        let mut edit = drawing.edit().expect("drawing edit");
        assert!(matches!(
            edit.replace_paragraph(anchor.position(), "replacement"),
            Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::FloatingOfficeArt
            }))
        ));
    }

    #[test]
    fn non_destructive_revision_disposition_is_durable_and_reversible() {
        let mut writer = Builder::new();
        writer
            .add_paragraph_runs(
                vec![(
                    "deleted".to_string(),
                    CharacterFormatting {
                        deletion_revision: Some(TextRevision::new("Reviewer")),
                        ..CharacterFormatting::default()
                    },
                )],
                ParagraphFormatting::default(),
            )
            .expect("revision paragraph");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("revision DOC");
        let bytes = normalize_word2002_dop(output.into_inner());
        let source = Snapshot::parse(&bytes).expect("revision snapshot");
        let position = source
            .revisions()
            .expect("revisions")
            .iter()
            .position(|revision| revision.kind == crate::tracked_revision::RevisionKind::Deletion)
            .map(Position::new)
            .expect("deletion revision");

        let mut destructive = source.edit().expect("destructive edit");
        assert!(matches!(
            destructive.dispose_revision(position, RevisionDisposition::Accept),
            Err(Error::Refused(
                Refusal::DestructiveRevisionDisposition { .. }
            ))
        ));

        let mut edit = source.edit().expect("revision edit");
        edit.dispose_revision(position, RevisionDisposition::Reject)
            .expect("reject deletion mark without deleting text");
        let commit = edit.commit().expect("revision commit");
        assert!(
            commit
                .package()
                .revisions()
                .expect("changed revisions")
                .is_empty()
        );
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("revision durable patch");
        assert!(
            source
                .apply_durable(&durable)
                .expect("revision durable replay")
                .revisions()
                .expect("replayed revisions")
                .is_empty()
        );
        let restored = commit
            .package()
            .apply_durable(&durable.inverse())
            .expect("revision durable inverse");
        assert_eq!(restored.revisions().expect("restored revisions").len(), 1);
    }

    #[test]
    fn retained_picture_mismatch_is_atomic_and_missing_data_handoff_is_current() {
        for floating in [false, true] {
            let donor = Snapshot::parse(&picture_doc(floating)).expect("donor");
            let mut package = open_object(
                formatted_picture_receiver_doc(),
                Targets::default(),
                Limits::default(),
            )
            .expect("receiver package");
            package
                .remove_stream(&["Data".to_string()])
                .expect("remove unused Data");
            let receiver = Snapshot::parse(
                &crate::package::object_edit_bytes(&package).expect("receiver bytes"),
            )
            .expect("receiver");
            let plan = receiver
                .plan_picture_transfer_from(
                    &donor,
                    TextTarget::body_paragraph(Position::new(0)),
                    TextTarget::body_paragraph(Position::new(0)),
                )
                .expect("picture plan");
            let mut edit = receiver.edit().expect("edit");
            edit.replace_paragraph(Position::new(1), "changed other")
                .expect("first edit");
            let held = edit.retained_render_bytes();
            assert!(held.is_some());
            let before = edit.editor.clone().finish().expect("before bytes");
            let changes = edit.changes.len();
            let mut mismatched = plan.clone();
            mismatched.graph.data_offset = Some(u32::MAX);
            assert!(matches!(
                edit.apply_picture_transfer(&mismatched),
                Err(Error::PatchConflict)
            ));
            assert_eq!(edit.retained_render_bytes(), held);
            assert_eq!(edit.changes.len(), changes);
            assert_eq!(edit.editor.clone().finish().expect("after refusal"), before);
            edit.apply_picture_transfer(&plan)
                .expect("valid missing-Data insertion");
            assert!(edit.retained_render_bytes().is_some());
            let commit = edit.commit().expect("retained commit");
            let zero = receiver.with_transaction_limits(
                receiver
                    .transaction_limits
                    .with_max_retained_render_bytes(0),
            );
            let mut control = zero.edit().expect("control");
            control
                .replace_paragraph(Position::new(1), "changed other")
                .expect("control first");
            control
                .apply_picture_transfer(&plan)
                .expect("control picture");
            assert_eq!(control.retained_render_bytes(), None);
            assert_eq!(
                control.commit().expect("control commit").package().bytes(),
                commit.package().bytes()
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.package())
                    .expect("inverse"),
                receiver
            );
        }
    }

    #[test]
    fn retained_render_is_cleared_by_resource_reopen_and_formatting_flushes_at_commit() {
        let donor = Snapshot::parse(&embedded_doc()).expect("embedded donor");
        let receiver = Snapshot::parse(&doc(&["placeholder", "other"])).expect("receiver");
        let plan = receiver
            .plan_embedded_transfer_from(&donor, 77, 88)
            .expect("embedded plan");
        let mut outputs = Vec::new();
        for ceiling in [0, 8 * 1024 * 1024] {
            let source = receiver.with_transaction_limits(
                receiver
                    .transaction_limits
                    .with_max_retained_render_bytes(ceiling),
            );
            let mut edit = source.edit().expect("edit");
            edit.replace_paragraph(Position::new(1), "changed other")
                .expect("text edit");
            assert_eq!(edit.retained_render_bytes().is_some(), ceiling != 0);
            edit.apply_embedded_transfer(&plan).expect("resource edit");
            assert_eq!(edit.retained_render_bytes(), None);
            edit.set_paragraph_bold(Position::new(1), true)
                .expect("format after resource reopen");
            // Direct format changes stay in the editor's CHPX model until
            // the transaction renders. This is what lets a run of toggles
            // rebuild and publish its FKP once; the final output still uses
            // the configured retention ceiling.
            assert_eq!(edit.retained_render_bytes(), None);
            outputs.push(edit.commit().expect("commit").package().bytes().to_vec());
        }
        assert_eq!(outputs[0], outputs[1]);
    }

    #[test]
    fn canonical_inline_and_floating_picture_graphs_transfer_and_reopen() {
        for floating in [false, true] {
            let donor = Snapshot::parse(&picture_doc(floating)).expect("picture donor");
            let receiver = Snapshot::parse(&formatted_picture_receiver_doc()).expect("receiver");
            let plan = receiver
                .plan_picture_transfer_from(
                    &donor,
                    TextTarget::body_paragraph(Position::new(0)),
                    TextTarget::body_paragraph(Position::new(0)),
                )
                .expect("bounded picture transfer plan");
            assert_eq!(plan.is_floating(), floating);
            assert!(plan.picture_block_size() > 68);

            let mut edit = receiver.edit().expect("picture edit");
            edit.apply_picture_transfer(&plan)
                .expect("install picture graph");
            let commit = edit.commit().expect("picture commit and full reopen");
            let marker = if floating { '\u{0008}' } else { '\u{0001}' };
            assert_eq!(
                commit.package().paragraphs(Projection::All).unwrap()[0].text(),
                marker.to_string()
            );
            assert_eq!(
                commit
                    .patch()
                    .changes()
                    .filter(|change| matches!(change, super::ChangeRef::Picture { .. }))
                    .count(),
                1
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.package())
                    .expect("exact picture inverse"),
                receiver
            );
            let durable = commit
                .patch()
                .to_durable(resource_patch_limits())
                .expect("picture durable patch");
            let replay = receiver
                .apply_durable(&durable)
                .expect("picture durable replay and reopen");
            assert_eq!(&replay, commit.package());
            let durable_restored = replay
                .apply_durable(&durable.inverse())
                .expect("picture durable inverse and reopen");
            assert_eq!(
                durable_restored.paragraphs(Projection::All).unwrap()[0].text(),
                "placeholder"
            );
            let mut restored_package =
                crate::raw::Reader::from_reader(Cursor::new(durable_restored.finish()))
                    .expect("durable inverse CFB");
            assert_eq!(
                restored_package
                    .document()
                    .expect("durable inverse DOC")
                    .paragraphs()
                    .expect("durable inverse paragraphs")[0]
                    .runs()
                    .expect("durable inverse runs")[0]
                    .bold(),
                Some(true)
            );

            let mut right = receiver.edit().expect("disjoint text edit");
            right
                .replace_paragraph(Position::new(1), "changed")
                .expect("disjoint replacement");
            let right = right.commit().expect("disjoint text commit");
            let merged = receiver
                .plan_three_way(commit.patch(), right.patch())
                .expect("picture/text three-way plan");
            assert!(merged.is_conflict_free());
            let merged = merged.commit().expect("picture/text merge");
            assert_eq!(
                merged.package().paragraphs(Projection::All).unwrap()[1].text(),
                "changed"
            );

            let mut history = receiver.history(HistoryLimits::new(1, u64::MAX));
            history
                .record(commit.package().clone(), 1)
                .expect("picture history record");
            assert!(history.undo());
            assert_eq!(history.current(), &receiver);
            assert!(history.redo());
            assert_eq!(history.current(), commit.package());
        }
    }

    #[test]
    fn picture_transfer_rehomes_shared_pictures_beside_main_and_header_drawings() {
        let singleton = Snapshot::parse(&picture_doc(false)).expect("singleton donor");
        let receiver_with_picture =
            Snapshot::parse(&picture_doc(false)).expect("occupied receiver");
        assert!(matches!(
            receiver_with_picture.plan_picture_transfer_from(
                &singleton,
                TextTarget::body_paragraph(Position::new(0)),
                TextTarget::body_paragraph(Position::new(0)),
            ),
            Err(Error::Refused(Refusal::DrawingDependency {
                dependency: DrawingDependency::PictureGraphCollision
            }))
        ));

        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/images/png/lena.png"),
        )
        .expect("PNG fixture");
        let mut writer = Builder::new();
        writer
            .insert_floating_shape(
                crate::writer::Shape::new(crate::writer::Kind::Rectangle, 640, 320)
                    .expect("neighbor shape"),
                FloatingPosition::new(20, 40),
            )
            .expect("neighbor shape run");
        writer
            .insert_floating_text_box(
                crate::writer::Shape::new(crate::writer::Kind::Rectangle, 680, 360)
                    .expect("neighbor textbox"),
                FloatingPosition::new(60, 80),
                "unrelated textbox content",
            )
            .expect("neighbor textbox run");
        writer
            .insert_header_text_box(
                crate::HeaderKind::Odd,
                crate::writer::Shape::new(crate::writer::Kind::Rectangle, 720, 400)
                    .expect("header textbox"),
                FloatingPosition::new(90, 110),
                "unrelated header textbox content",
            )
            .expect("header textbox run");
        writer
            .insert_header_picture(
                crate::HeaderKind::Odd,
                Picture::new(bytes.clone()).expect("header picture"),
                FloatingPosition::new(130, 150),
            )
            .expect("header picture run");
        writer
            .insert_picture(Picture::new(bytes.clone()).expect("first picture"))
            .expect("first picture run");
        writer
            .insert_floating_picture(
                Picture::new(bytes.clone()).expect("first floating picture"),
                FloatingPosition::new(100, 200),
            )
            .expect("first floating picture run");
        writer
            .insert_picture(Picture::new(bytes.clone()).expect("second picture"))
            .expect("second picture run");
        writer
            .insert_floating_picture(
                Picture::new(bytes).expect("second floating picture"),
                FloatingPosition::new(300, 400),
            )
            .expect("second floating picture run");
        let mut output = Cursor::new(Vec::new());
        writer
            .build()
            .unwrap()
            .write_to(&mut output)
            .expect("mixed drawing donor DOC");
        let bytes = normalize_word2002_dop(output.into_inner());
        let shared = Snapshot::parse(&bytes).expect("mixed shared-store donor");
        for (position, floating) in [(Position::new(4), false), (Position::new(5), true)] {
            let empty = Snapshot::parse(&doc(&["placeholder"])).expect("empty receiver");
            let plan = empty
                .plan_picture_transfer_from(
                    &shared,
                    TextTarget::body_paragraph(position),
                    TextTarget::body_paragraph(Position::new(0)),
                )
                .expect("selected graph is safely re-homed from shared donor");
            assert_eq!(plan.is_floating(), floating);
            let mut edit = empty.edit().expect("multi-picture transfer edit");
            edit.apply_picture_transfer(&plan)
                .expect("re-home selected picture");
            let commit = edit.commit().expect("re-homed picture fully reopens");
            let durable = commit
                .patch()
                .to_durable(resource_patch_limits())
                .expect("re-homed picture durable patch");
            assert_eq!(
                &empty
                    .apply_durable(&durable)
                    .expect("re-homed picture durable replay"),
                commit.package()
            );
        }

        for position in [Position::new(0), Position::new(1)] {
            let empty = Snapshot::parse(&doc(&["placeholder"])).expect("empty receiver");
            assert!(matches!(
                empty.plan_picture_transfer_from(
                    &shared,
                    TextTarget::body_paragraph(position),
                    TextTarget::body_paragraph(Position::new(0)),
                ),
                Err(Error::Refused(Refusal::DrawingDependency {
                    dependency: DrawingDependency::UnsupportedPictureGraph
                }))
            ));
        }
    }
}
