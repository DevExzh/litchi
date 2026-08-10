//! Bounded, source-preserving edits for ordinary main-story DOC paragraphs.
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

use crate::DateTime;
use crate::package::Error as PackageError;
use crate::tracked_revision::{Limits, Revision, RevisionEditor, RevisionKind, RevisionMetadata};
use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, CompositionError, JoinedSubEdits, PatchError, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation, SubEdit,
};
pub use litchi_core::patch::{CompositionLimits, HistoryLimits, SubEditJoinFailure};
use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;
use std::sync::Arc;

/// Bounded undo/redo history over immutable DOC snapshots.
pub type History = litchi_core::patch::History<Snapshot>;

/// Finite limits for one body transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TransactionLimits {
    operations: usize,
    replacement_units: usize,
    total_units: usize,
}

impl TransactionLimits {
    /// Creates explicit per-transaction bounds.
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
        }
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
}

impl Default for TransactionLimits {
    fn default() -> Self {
        Self::new(256, 1024 * 1024, 4 * 1024 * 1024)
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
    /// A picture graph uses noncanonical/reordered sharing, grouping,
    /// textboxes, producer extensions, PICF transforms, or a delayed BLIP.
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
    /// snapshot collection is checked by [`Edit::replace_paragraph`].
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
    /// The configured operation count was exhausted.
    OperationLimit { observed: usize, limit: usize },
    /// One or all replacement payloads exceed the configured UTF-16 bound.
    ReplacementLimit { observed: usize, limit: usize },
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
            Self::OperationLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} operations; limit is {limit}"
            ),
            Self::ReplacementLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} UTF-16 replacement units; limit is {limit}"
            ),
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
    /// A patch was presented with any snapshot other than its exact source.
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
            Self::Conflict => formatter.write_str("body-text patch source conflict"),
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
            Self::Refused(_) | Self::Conflict => None,
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

/// Immutable, exact-source snapshot for the body-text transaction seam.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
    transaction_limits: TransactionLimits,
}

impl Snapshot {
    /// Opens an owned Word 97+ DOC source after validating its safe edit basis.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the CFB, Word 97+ FIB, selected table
    /// stream, piece table, or FKP basis cannot support safe editing.
    pub fn open(input: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::open_bounded(input, limits, TransactionLimits::default())
    }

    /// Opens an owned DOC source with explicit package and transaction bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the complete package or its safe edit
    /// basis cannot be reopened.
    pub fn open_bounded(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
    ) -> Result<Self> {
        let bytes = input.into();
        RevisionEditor::open(bytes.clone(), limits).map_err(Error::Invalid)?;
        let mut package =
            crate::Package::from_reader(Cursor::new(bytes.clone())).map_err(Error::Invalid)?;
        package.document().map_err(Error::Invalid)?;
        Ok(Self {
            source: Arc::from(bytes.into_boxed_slice()),
            limits,
            transaction_limits,
        })
    }

    /// Parses a borrowed DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn parse(input: &[u8]) -> Result<Self> {
        Self::open(input.to_vec(), Limits::default())
    }

    /// Opens an owned DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn from_bytes(input: Vec<u8>) -> Result<Self> {
        Self::open(input, Limits::default())
    }

    /// Exact CFB source bytes retained for source checks and byte-exact no-ops.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Stable first-stage fingerprint for diagnostics and stale-source checks.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        fingerprint(&self.source)
    }

    /// Configured finite semantic-operation and text-payload bounds.
    #[must_use]
    pub const fn transaction_limits(&self) -> TransactionLimits {
        self.transaction_limits
    }

    /// Lists ordinary source-body paragraphs under the requested review projection.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an invalid source or [`Error::Refused`]
    /// when tracked text ranges overlap ambiguously for the projection.
    pub fn paragraphs(&self, projection: Projection) -> Result<Vec<Paragraph>> {
        let editor = self.editor()?;
        projected_paragraphs(&editor, projection)
    }

    /// Lists ordinary paragraphs from any DOC story. Main-story review
    /// projection remains available through [`Self::paragraphs`]; non-main
    /// stories expose their exact stored text.
    pub fn story_paragraphs(&self, story: Story) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::Story(story)).map(items_from_spans)
    }

    /// Lists simple main-story table cells that can be edited without changing
    /// table marks or PAPX/TAP structure.
    pub fn table_cells(&self) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::TableCells).map(items_from_spans)
    }

    /// Lists cached results of balanced, non-nested main-story fields.
    pub fn field_results(&self) -> Result<Vec<TextItem>> {
        let editor = self.editor()?;
        target_spans(&editor, TargetCollection::FieldResults).map(items_from_spans)
    }

    /// Managed embedded fields and their passive `ObjectPool` metadata.
    pub fn embedded_objects(&self) -> Result<crate::embedded_object::Inventory> {
        crate::embedded_object::Snapshot::open(self.finish(), self.limits)
            .map_err(Error::Invalid)?
            .inventory()
            .map_err(Error::Invalid)
    }

    /// Lists tracked main-story ranges in stable CP/kind order.
    pub fn revisions(&self) -> Result<Vec<Revision>> {
        self.editor()?.revisions().map_err(Error::Invalid)
    }

    /// Starts a staged bounded body text-and-formatting transaction.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] if the retained source no longer validates.
    pub fn edit(&self) -> Result<Edit> {
        Edit::new(self.clone())
    }

    /// Alias for [`Self::edit`].
    ///
    /// # Errors
    ///
    /// Returns the same failure as [`Self::edit`].
    pub fn transaction(&self) -> Result<Edit> {
        self.edit()
    }

    /// Starts a bounded disjoint composition for this exact artifact.
    #[must_use]
    pub fn compose(&self, limits: CompositionLimits) -> Composition {
        Composition {
            source: self.clone(),
            joined: JoinedSubEdits::new(self.lineage(), limits),
        }
    }

    /// Starts explicit bounded undo/redo history at this immutable snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Prepares one independently composable paragraph replacement against
    /// this exact immutable artifact.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as [`Edit::replace_paragraph`]
    /// or a common composition-bound error.
    pub fn prepare_replace(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        replacement: impl Into<String>,
    ) -> Result<PreparedEdit> {
        let replacement = replacement.into();
        let mut validation = self.edit()?;
        let target = TextTarget::body_paragraph(position);
        validation.replace_text(target, &replacement)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Text {
                target,
                replacement,
            },
        )
    }

    /// Prepares one independently composable direct-bold change.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as
    /// [`Edit::set_paragraph_bold`] or a common composition-bound error.
    pub fn prepare_bold(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        enabled: bool,
    ) -> Result<PreparedEdit> {
        let mut validation = self.edit()?;
        let target = TextTarget::body_paragraph(position);
        validation.set_character_property(target, CharacterProperty::Bold, enabled)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Format {
                target,
                property: CharacterProperty::Bold,
                enabled,
            },
        )
    }

    /// Prepares one independently composable semantic text replacement.
    pub fn prepare_text(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        replacement: impl Into<String>,
    ) -> Result<PreparedEdit> {
        let replacement = replacement.into();
        let mut validation = self.edit()?;
        validation.replace_text(target, &replacement)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Text {
                target,
                replacement,
            },
        )
    }

    /// Prepares one independently composable direct-format mutation.
    pub fn prepare_format(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    ) -> Result<PreparedEdit> {
        let mut validation = self.edit()?;
        validation.set_character_property(target, property, enabled)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Format {
                target,
                property,
                enabled,
            },
        )
    }

    /// Applies supported durable body-text operations to this exact artifact.
    ///
    /// # Errors
    ///
    /// Returns an error for a foreign vocabulary, malformed selector,
    /// artifact or semantic precondition conflict, bound violation, or typed
    /// body-edit refusal.
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Self> {
        if patch.format() != "litchi-doc-body" {
            return Err(invalid_durable_patch("unsupported format"));
        }
        if patch.operations().is_empty() {
            if !patch.blobs().is_empty() {
                return Err(invalid_durable_patch(
                    "operation-free durable patch has unreferenced blobs",
                ));
            }
            return Ok(self.clone());
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
                return Err(Error::Conflict);
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
                        return Err(Error::Conflict);
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
        edit.commit().map(|commit| commit.snapshot)
    }

    /// Plans an inert text transfer while proving the receiving selector and
    /// exact target artifact. Text carries no DOC style, drawing, or resource
    /// identifiers, so its dependency closure is empty.
    pub fn plan_text_transfer_from(
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
        let mut validation = self.edit()?;
        validation.replace_text(destination, &value)?;
        Ok(TransferPlan {
            target: self.lineage(),
            destination,
            value,
        })
    }

    /// Plans transfer of a complete inert embedded-object dependency closure.
    ///
    /// The donor's standalone object CFB and exact preview block are retained;
    /// the receiver regenerates the field text, field PLCF, CHPX picture
    /// references, Data-stream offset, and `ObjectPool` storage identity.
    pub fn plan_embedded_transfer_from(
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
        let mut validation = self.edit()?;
        validation.add_embedded_object(options.clone())?;
        Ok(EmbeddedTransferPlan {
            target: self.lineage(),
            options,
        })
    }

    /// Plans transfer of one canonical inline or floating picture into a
    /// non-empty inert main-story placeholder.
    ///
    /// The bounded closure owns the marker CHPX, exact PICF/Data block, and,
    /// for floating pictures, a re-homed singleton `PlcfSpaMom` and `DggInfo`.
    /// Multi-picture donors are accepted after their complete shared graph is
    /// proved canonical. Groups, textboxes, producer extensions, delay-loaded
    /// images, auxiliary stories, and occupied receivers are refused.
    pub fn plan_picture_transfer_from(
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
        let mut validation = self.edit()?;
        validation.install_picture(destination, graph.clone())?;
        Ok(PictureTransferPlan {
            target: self.lineage(),
            destination,
            graph,
        })
    }

    /// Non-mutating three-way plan for two patches based on this exact source.
    pub fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
        if left.before != *self || right.before != *self {
            return Err(Error::Conflict);
        }
        Ok(ThreeWayPlan::new(self.clone(), left, right))
    }

    /// Exact source bytes. A snapshot has no implicit serialization step.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    fn editor(&self) -> Result<RevisionEditor> {
        RevisionEditor::open(self.source.as_ref().to_vec(), self.limits).map_err(Error::Invalid)
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
            .field("limits", &self.limits)
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

/// Clone-first staged text edit over one immutable source snapshot.
pub struct Edit {
    source: Snapshot,
    editor: RevisionEditor,
    changes: Vec<Change>,
    replacement_units: usize,
}

impl Edit {
    fn new(source: Snapshot) -> Result<Self> {
        let editor = source.editor()?;
        Ok(Self {
            source,
            editor,
            changes: Vec::new(),
            replacement_units: 0,
        })
    }

    /// Immutable source snapshot that authorizes this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
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
    pub fn replace_paragraph(&mut self, position: Position, replacement: &str) -> Result<()> {
        self.replace_text(TextTarget::body_paragraph(position), replacement)
            .map_err(|error| match error {
                Error::Refused(Refusal::TargetNotFound) => {
                    Error::Refused(Refusal::ParagraphNotFound)
                },
                other => other,
            })
    }

    /// Replaces one checked ordinary paragraph, simple table-cell value, or
    /// simple field cached result. Main-story targets may change length;
    /// non-main story paragraphs require equal UTF-16 length.
    pub fn replace_text(&mut self, target: TextTarget, replacement: &str) -> Result<()> {
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
        let actual = replacement.encode_utf16().count();
        self.ensure_replacement_capacity(actual)?;
        if span.start_cp == span.end_cp {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        if text_revision_intersects(&self.editor, span.start_cp, span.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        if span.text == replacement {
            return Ok(());
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
        if actual != expected
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()?;
        if !self
            .editor
            .has_uniform_character_format(span.start_cp, span.end_cp)
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
    pub fn set_paragraph_bold(&mut self, position: Position, enabled: bool) -> Result<()> {
        self.set_character_property(
            TextTarget::body_paragraph(position),
            CharacterProperty::Bold,
            enabled,
        )
    }

    /// Sets a direct bold, italic, or single-underline override on one semantic
    /// text target without creating or rewriting style-table resources.
    pub fn set_character_property(
        &mut self,
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    ) -> Result<()> {
        self.set_character_property_override(target, property, Some(enabled))
    }

    /// Applies a dependency-checked text transfer prepared for this exact
    /// receiving artifact.
    pub fn apply_transfer(&mut self, plan: &TransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::Conflict);
        }
        self.replace_text(plan.destination, &plan.value)
    }

    /// Applies a dependency-closed embedded-object transfer prepared for this
    /// exact receiving artifact.
    pub fn apply_embedded_transfer(&mut self, plan: &EmbeddedTransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::Conflict);
        }
        self.add_embedded_object(plan.options.clone())
    }

    /// Applies a canonical picture graph prepared for this exact receiver.
    pub fn apply_picture_transfer(&mut self, plan: &PictureTransferPlan) -> Result<()> {
        if plan.target != self.source.lineage() {
            return Err(Error::Conflict);
        }
        self.install_picture(plan.destination, plan.graph.clone())
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
        self.ensure_operation_capacity()?;
        if !self
            .editor
            .has_uniform_character_format(span.start_cp, span.end_cp)
            .map_err(Error::Invalid)?
        {
            return Err(Error::Refused(Refusal::FormattingDependency));
        }
        let expected = span.text.encode_utf16().count();
        if expected != actual
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        let before = PictureSlot::Text(span.text);
        let installed = self
            .editor
            .replace_with_picture_graph(span.start_cp, span.end_cp, &graph)
            .map_err(Error::Invalid)?;
        if graph.data_offset.is_some() && installed != graph {
            return Err(Error::Conflict);
        }
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
            return Err(Error::Conflict);
        }
        if has_structural_content(&replacement) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        let actual = replacement.encode_utf16().count();
        self.ensure_replacement_capacity(actual)?;
        self.ensure_operation_capacity()?;
        if actual != 1
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
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
    pub fn add_embedded_object(
        &mut self,
        mut options: crate::embedded_object::WriteOptions,
    ) -> Result<()> {
        let storage_id = options.storage_id;
        options.instruction = format!(" EMBED LITCHI_OBJECT _{storage_id} ");
        let bytes = self.editor.clone().finish().map_err(Error::Invalid)?;
        let snapshot = crate::embedded_object::Snapshot::open(bytes, self.source.limits)
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
        self.editor = RevisionEditor::open(bytes, self.source.limits).map_err(Error::Invalid)?;
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
    pub fn remove_embedded_object(&mut self, storage_id: u32) -> Result<()> {
        let bytes = self.editor.clone().finish().map_err(Error::Invalid)?;
        let snapshot = crate::embedded_object::Snapshot::open(bytes, self.source.limits)
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
        self.editor = RevisionEditor::open(bytes, self.source.limits).map_err(Error::Invalid)?;
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
    pub fn dispose_revision(
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
    pub fn set_embedded_display_as_icon(&mut self, storage_id: u32, enabled: bool) -> Result<()> {
        let bytes = self.editor.clone().finish().map_err(Error::Invalid)?;
        let snapshot = crate::embedded_object::Snapshot::open(bytes, self.source.limits)
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
        self.editor = RevisionEditor::open(bytes, self.source.limits).map_err(Error::Invalid)?;
        self.changes.push(Change::EmbeddedDisplay {
            storage_id,
            before,
            after: enabled,
        });
        Ok(())
    }

    /// Discards staged changes and returns the original immutable snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Publishes a validated snapshot and its reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the rendered candidate cannot be
    /// reopened with the original safety limits.
    pub fn commit(self) -> Result<Commit> {
        let bytes = self.editor.finish().map_err(Error::Invalid)?;
        let snapshot = if bytes == self.source.bytes() {
            self.source.clone()
        } else {
            Snapshot::open_bounded(bytes, self.source.limits, self.source.transaction_limits)?
        };
        let patch = Patch::new(self.source, snapshot.clone(), self.changes);
        Ok(Commit { snapshot, patch })
    }

    fn ensure_operation_capacity(&self) -> Result<()> {
        let observed = self.changes.len().saturating_add(1);
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

/// Validated commit result for one body-text transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether a DOC byte changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Published post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits a commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// In-memory reversible replacement guarded by exact source bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    before_fingerprint: u64,
    after_fingerprint: u64,
    changes: Vec<Change>,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot, changes: Vec<Change>) -> Self {
        Self {
            before_fingerprint: before.fingerprint(),
            after_fingerprint: after.fingerprint(),
            before,
            after,
            changes,
        }
    }

    /// Semantic text and formatting changes in application order.
    #[must_use]
    pub fn changes(&self) -> impl ExactSizeIterator<Item = ChangeRef<'_>> {
        self.changes.iter().map(Change::as_ref)
    }

    /// Exact source snapshot required for application.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot produced by the transaction.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Fast stale-source precheck; exact bytes remain authoritative.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.before_fingerprint
    }

    /// Target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.after_fingerprint
    }

    /// Whether this patch preserves the exact artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] unless `source` has byte-for-byte equality
    /// with this patch's captured source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.before_fingerprint || source.bytes() != self.before.bytes()
        {
            return Err(Error::Conflict);
        }
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }

    /// Exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
        }
    }

    /// Converts this exact-source patch to the shared bounded deterministic
    /// semantic envelope.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError`] when the requested wire limits cannot represent
    /// every semantic operation and inverse.
    pub fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> std::result::Result<litchi_core::patch::Patch<Reversible>, PatchError> {
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
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
pub struct PreparedEdit {
    inner: SubEdit<Lineage, PreparedOperation>,
}

impl PreparedEdit {
    fn new(
        lineage: Lineage,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        operation: PreparedOperation,
    ) -> Result<Self> {
        let writes = [operation.effect()];
        let inner = SubEdit::new(
            lineage,
            limits,
            identifier,
            std::iter::empty(),
            writes,
            operation,
        )?;
        Ok(Self { inner })
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

impl std::fmt::Debug for PreparedEdit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PreparedEdit")
            .field("identifier", &self.identifier())
            .field("target", &self.target())
            .finish_non_exhaustive()
    }
}

/// Recoverable failure to join one independently prepared edit.
pub struct JoinError {
    failure: SubEditJoinFailure,
    rejected: PreparedEdit,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Recovers the rejected prepared edit.
    #[must_use]
    pub fn into_rejected(self) -> PreparedEdit {
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
pub struct Composition {
    source: Snapshot,
    joined: JoinedSubEdits<Lineage, PreparedOperation>,
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
    pub fn join(&mut self, incoming: PreparedEdit) -> std::result::Result<&mut Self, JoinError> {
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure,
                rejected: PreparedEdit { inner: rejected },
            });
        }
        Ok(self)
    }

    /// Number of accepted independently prepared edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.joined.len()
    }

    /// Whether no edits have been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.joined.is_empty()
    }

    /// Commits all accepted edits atomically in stable identifier order.
    ///
    /// # Errors
    ///
    /// Returns the same staging, publication, and full-reopen errors as an
    /// ordinary multi-operation [`Edit`].
    pub fn commit(self) -> Result<Commit> {
        let mut edit = self.source.edit()?;
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
        edit.commit()
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
}

/// Canonical re-homed inline/floating picture closure prepared for one exact
/// receiver artifact.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PictureTransferPlan {
    target: Lineage,
    destination: TextTarget,
    graph: crate::tracked_revision::PictureGraph,
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
pub struct ThreeWayPlan {
    source: Snapshot,
    changes: Vec<Change>,
    conflicts: Vec<ThreeWayConflict>,
}

impl ThreeWayPlan {
    fn new(source: Snapshot, left: &Patch, right: &Patch) -> Self {
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
    pub fn conflicts(&self) -> &[ThreeWayConflict] {
        &self.conflicts
    }

    #[must_use]
    pub fn is_conflict_free(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Publishes the conflict-free combined plan through the ordinary bounded
    /// transaction and full reopen boundary.
    pub fn commit(self) -> Result<Commit> {
        if !self.conflicts.is_empty() {
            return Err(Error::Conflict);
        }
        let mut edit = self.source.edit()?;
        for change in &self.changes {
            apply_change_after(&mut edit, change)?;
        }
        edit.commit()
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
        return Err(Error::Conflict);
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
        return Err(Error::Conflict);
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
        return Err(Error::Conflict);
    }
    let enabled = operation
        .value
        .as_bool()
        .ok_or_else(|| invalid_durable_patch("embedded display value is invalid"))?;
    edit.set_embedded_display_as_icon(storage_id, enabled)
}

fn embedded_display_value(edit: &Edit, storage_id: u32) -> Result<bool> {
    let bytes = edit.editor.clone().finish().map_err(Error::Invalid)?;
    let snapshot = crate::embedded_object::Snapshot::open(bytes, edit.source.limits)
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
                    return Err(Error::Conflict);
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
                return Err(Error::Conflict);
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
        return Err(Error::Conflict);
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
    let bytes = edit.editor.clone().finish().map_err(Error::Invalid)?;
    let snapshot = crate::embedded_object::Snapshot::open(bytes, edit.source.limits)
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
        return Err(Error::Conflict);
    }
    let actual = matches.first().map(|(_index, spec)| spec.clone());
    if actual != expected {
        return Err(Error::Conflict);
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
                .ok_or(Error::Conflict)?;
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
        TextTarget::TableCell(_) | TextTarget::FieldResult(_) => Story::Main,
    }
}

fn resolve_target(editor: &RevisionEditor, target: TextTarget) -> Result<SourceSpan> {
    let spans = match target {
        TextTarget::StoryParagraph { story, .. } => {
            target_spans(editor, TargetCollection::Story(story))?
        },
        TextTarget::TableCell(_) => target_spans(editor, TargetCollection::TableCells)?,
        TextTarget::FieldResult(_) => target_spans(editor, TargetCollection::FieldResults)?,
    };
    let position = match target {
        TextTarget::StoryParagraph { position, .. }
        | TextTarget::TableCell(position)
        | TextTarget::FieldResult(position) => position,
    };
    spans
        .get(position.get())
        .cloned()
        .ok_or(Error::Refused(Refusal::TargetNotFound))
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

fn has_structural_content(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(character, '\r' | '\u{7}' | '\u{13}'..='\u{15}' | '\u{fffc}')
            || (character.is_control() && character != '\t')
    })
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

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "unit-test fixtures use contextual fail-fast assertions"
    )]

    use super::{
        CharacterProperty, DrawingDependency, Error, Projection, Refusal, RevisionDisposition,
        Snapshot, Story, TextTarget, TransactionLimits,
    };
    use crate::tracked_revision::Limits;
    use crate::writer::{
        CharacterFormatting, FloatingPosition, ParagraphFormatting, Picture, TextRevision, Writer,
    };
    use litchi_cfb::OleWriter;
    use litchi_core::Position;
    use litchi_core::patch::{
        BlobLimits, CompositionLimits, HistoryLimits, Patch, PatchLimits, Reversible,
        SubEditJoinFailure,
    };
    use std::io::Cursor;

    fn doc(paragraphs: &[&str]) -> Vec<u8> {
        let mut writer = Writer::new();
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
            .write_to(&mut output)
            .expect("fixture DOC must serialize");
        output.into_inner()
    }

    fn picture_doc(floating: bool) -> Vec<u8> {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/images/png/lena.png"),
        )
        .expect("PNG fixture");
        let picture = Picture::new(bytes).expect("picture fixture");
        let mut writer = Writer::new();
        if floating {
            writer
                .insert_floating_picture(picture, FloatingPosition::new(1440, 720))
                .expect("floating picture");
        } else {
            writer.insert_picture(picture).expect("inline picture");
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("picture DOC");
        output.into_inner()
    }

    fn formatted_picture_receiver_doc() -> Vec<u8> {
        let mut writer = Writer::new();
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
            .write_to(&mut output)
            .expect("formatted receiver DOC");
        output.into_inner()
    }

    fn structured_doc() -> Vec<u8> {
        let mut writer = Writer::new();
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
        writer.write_to(&mut output).expect("structured DOC");
        output.into_inner()
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
        assert!(commit.changed());
        assert_eq!(
            commit
                .snapshot()
                .paragraphs(Projection::All)
                .expect("changed paragraphs")[0]
                .text(),
            "omega"
        );

        let applied = commit.patch().apply(&source).expect("exact source applies");
        assert_eq!(applied, *commit.snapshot());
        let restored = commit
            .patch()
            .inverse()
            .apply(&applied)
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());

        let other = Snapshot::open(doc(&["other"]), Limits::default()).expect("other source");
        assert!(matches!(commit.patch().apply(&other), Err(Error::Conflict)));
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
        assert!(commit.changed());
        let paragraphs = commit
            .snapshot()
            .paragraphs(Projection::All)
            .expect("changed paragraphs");
        assert_eq!(paragraphs[0].text(), "a much longer paragraph");
    }

    #[test]
    fn accepted_and_rejected_projections_hide_text_revisions() {
        let mut writer = Writer::new();
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
        writer.write_to(&mut output).expect("fixture DOC");
        let snapshot = Snapshot::parse(&output.into_inner()).expect("snapshot");
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
            .snapshot()
            .paragraphs(Projection::All)
            .expect("readback")
            .into_iter()
            .map(|paragraph| paragraph.text().to_string())
            .collect::<Vec<_>>();
        assert_eq!(texts, ["alpha expanded", "bravo", "c"]);

        let mut package = crate::Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("CFB reopens");
        let document = package.document().expect("DOC reopens");
        let paragraphs = document.paragraphs().expect("paragraphs read back");
        assert_eq!(paragraphs[0].runs().expect("runs")[0].bold(), Some(true));

        let restored = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());
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
            *commit.snapshot()
        );
        let semantic_inverse = commit
            .snapshot()
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
            crate::Package::from_reader(Cursor::new(semantic_inverse.finish()))
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
            .record(commit.snapshot().clone(), wire_bytes)
            .expect("record history");
        assert!(history.undo());
        assert_eq!(history.current(), &source);
        assert!(history.redo());
        assert_eq!(history.current(), commit.snapshot());

        let mut too_small = source.history(HistoryLimits::new(1, wire_bytes - 1));
        assert!(
            too_small
                .record(commit.snapshot().clone(), wire_bytes)
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
                .snapshot()
                .story_paragraphs(Story::Header)
                .expect("changed header")
                .into_iter()
                .find(|item| item.target() == header.target())
                .expect("same header selector")
                .text(),
            "Banner"
        );
        assert_eq!(
            commit.snapshot().table_cells().expect("changed cells")[0].text(),
            "Cell expanded"
        );
        assert_eq!(
            commit.snapshot().field_results().expect("changed fields")[0].text(),
            "Result"
        );
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("durable patch");
        let replay = source.apply_durable(&durable).expect("durable replay");
        assert_eq!(replay, *commit.snapshot());
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
            merged.snapshot().table_cells().expect("merged cells")[0].text(),
            "Left cell"
        );
        assert_eq!(
            merged.snapshot().field_results().expect("merged fields")[0].text(),
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
                .snapshot()
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
            Err(Error::Conflict)
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
        let mut package = crate::Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("resource CFB reopens");
        package.document().expect("resource DOC reopens");
        let changed =
            crate::embedded_object::Snapshot::open(commit.snapshot().finish(), Limits::default())
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
            *commit.snapshot()
        );
        assert_eq!(
            commit
                .snapshot()
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
                .snapshot()
                .embedded_objects()
                .expect("transferred inventory")
                .get(88)
                .is_some()
        );
        let exact_inverse = transfer
            .patch()
            .inverse()
            .apply(transfer.snapshot())
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
                .snapshot()
                .embedded_objects()
                .expect("merged inventory")
                .get(88)
                .is_some()
        );
        assert_eq!(
            merged
                .snapshot()
                .paragraphs(Projection::All)
                .expect("merged text")[0]
                .text(),
            "receiver changed"
        );

        let wire_bytes = u64::try_from(wire.len()).expect("wire length fits u64");
        let mut history = receiver.history(HistoryLimits::new(1, wire_bytes));
        history
            .record(transfer.snapshot().clone(), wire_bytes)
            .expect("resource history record");
        assert!(history.undo());
        assert_eq!(history.current(), &receiver);
        assert!(history.redo());
        assert_eq!(history.current(), transfer.snapshot());

        assert!(matches!(
            transfer
                .snapshot()
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

        let mut writer = Writer::new();
        writer.add_paragraph("body").expect("body");
        writer
            .insert_floating_shape(
                crate::writer::Shape::new(crate::writer::Kind::Rectangle, 720, 360).expect("shape"),
                FloatingPosition::new(120, 240),
            )
            .expect("floating shape");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("drawing DOC");
        let drawing = Snapshot::parse(&output.into_inner()).expect("drawing snapshot");
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
        let mut writer = Writer::new();
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
        writer.write_to(&mut output).expect("revision DOC");
        let source = Snapshot::parse(&output.into_inner()).expect("revision snapshot");
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
                .snapshot()
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
            .snapshot()
            .apply_durable(&durable.inverse())
            .expect("revision durable inverse");
        assert_eq!(restored.revisions().expect("restored revisions").len(), 1);
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
                commit.snapshot().paragraphs(Projection::All).unwrap()[0].text(),
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
                    .apply(commit.snapshot())
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
            assert_eq!(&replay, commit.snapshot());
            let durable_restored = replay
                .apply_durable(&durable.inverse())
                .expect("picture durable inverse and reopen");
            assert_eq!(
                durable_restored.paragraphs(Projection::All).unwrap()[0].text(),
                "placeholder"
            );
            let mut restored_package =
                crate::Package::from_reader(Cursor::new(durable_restored.finish()))
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
                merged.snapshot().paragraphs(Projection::All).unwrap()[1].text(),
                "changed"
            );

            let mut history = receiver.history(HistoryLimits::new(1, u64::MAX));
            history
                .record(commit.snapshot().clone(), 1)
                .expect("picture history record");
            assert!(history.undo());
            assert_eq!(history.current(), &receiver);
            assert!(history.redo());
            assert_eq!(history.current(), commit.snapshot());
        }
    }

    #[test]
    fn picture_transfer_refuses_receiver_collision_and_rehomes_multi_picture_graphs() {
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
        let mut writer = Writer::new();
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
            .write_to(&mut output)
            .expect("multi-picture donor DOC");
        let shared = Snapshot::parse(&output.into_inner()).expect("shared-store donor");
        for (position, floating) in [(Position::new(2), false), (Position::new(3), true)] {
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
                commit.snapshot()
            );
        }
    }
}
