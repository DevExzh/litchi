//! Archive-free Keynote soundtrack item values and transaction vocabulary.
//!
//! This module contains the semantic side of soundtrack media lifecycle
//! operations. Package selection, native object ownership, data allocation,
//! metadata repair, ZIP reassembly, and candidate verification stay in the
//! private package adapter. Consequently no native identifier, component
//! name, archive member, protobuf value, or package byte slice is part of this
//! public API.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::media::Type as MediaType;
use thiserror::Error as ThisError;

use crate::Package;

/// Maximum filename storage accepted by AudioSource.
///
/// The limit matches the archive-wide member-name ceiling and is checked
/// before the source enters a package transaction. A package-specific limit
/// may still reject a source during commit when the caller selected a smaller
/// crate::Limits profile.
pub const MAX_FILENAME_BYTES: usize = 4 * 1024;

/// Maximum audio payload retained by one AudioSource.
///
/// The limit matches the archive-wide uncompressed entry ceiling. Package
/// limits are checked again by the physical transaction because the complete
/// candidate and aggregate package budgets also apply.
pub const MAX_AUDIO_BYTES: usize = 512 * 1024 * 1024;

/// A content-redacted failure while constructing or publishing soundtrack
/// items.
///
/// Native identifiers, package paths, source bytes, and parser diagnostics are
/// intentionally not carried by this error. This keeps failures stable at
/// the semantic boundary and prevents accidental disclosure of package
/// internals through Display or Debug.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The source package has no physical representation that can be changed.
    #[error("this Keynote source does not support physical soundtrack-item edits")]
    UnsupportedSource,
    /// The presentation has no rooted soundtrack.
    #[error("the Keynote presentation has no soundtrack")]
    SoundtrackNotFound,
    /// The selected source item position does not exist.
    #[error("soundtrack item source position {position:?} does not exist")]
    SourcePositionNotFound {
        /// Missing source position.
        position: Position,
    },
    /// The selected item handle does not belong to this exact source snapshot.
    #[error("the soundtrack item handle does not belong to this package snapshot")]
    ItemHandleConflict,
    /// A selector resolved to more than one semantic item.
    #[error("the Keynote soundtrack item selector is ambiguous")]
    AmbiguousSelector,
    /// The requested insertion position is outside the final sequence.
    #[error("soundtrack item insertion position {position:?} is outside {item_count} items")]
    InsertionOutOfRange {
        /// Requested final insertion position.
        position: Position,
        /// Existing item count used for the inclusive insertion range.
        item_count: usize,
    },
    /// One edit may stage only one item operation.
    #[error("a soundtrack-item operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested without a staged item operation.
    #[error("no soundtrack-item operation is staged")]
    NoStagedOperation,
    /// The existing soundtrack graph cannot be edited without risking data.
    #[error("the Keynote soundtrack item graph cannot be edited safely")]
    InvalidSource,
    /// The new filename is empty.
    #[error("soundtrack audio filename cannot be empty")]
    EmptyFilename,
    /// The filename contains a path separator, NUL, or control character.
    #[error("soundtrack audio filename must be one safe relative filename")]
    InvalidFilename,
    /// The filename exceeds the bounded semantic storage budget.
    #[error("soundtrack audio filename exceeds {maximum} bytes")]
    FilenameTooLong {
        /// Observed UTF-8 byte length.
        observed: usize,
        /// Maximum accepted filename length.
        maximum: usize,
    },
    /// The filename extension is not a recognized audio extension.
    #[error("soundtrack audio filename must use a recognized audio extension")]
    UnsupportedAudioFilename,
    /// The payload is empty.
    #[error("soundtrack audio payload cannot be empty")]
    EmptyAudio,
    /// The payload exceeds the bounded semantic storage budget.
    #[error("soundtrack audio payload exceeds {maximum} bytes")]
    AudioTooLarge {
        /// Observed payload length.
        observed: usize,
        /// Maximum accepted payload length.
        maximum: usize,
    },
    /// The payload does not have a recognized audio signature.
    #[error("soundtrack audio payload has no recognized audio signature")]
    UnsupportedAudioData,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Keynote soundtrack-item {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote soundtrack-item transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Candidate readback did not reproduce the requested item state.
    #[error("the edited Keynote soundtrack items failed verification")]
    Verification,
    /// The exact-source patch was applied to a different package snapshot.
    #[error("the Keynote soundtrack-item patch does not match the exact source package")]
    PatchConflict,
}

/// A finite resource governed by a soundtrack-item transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// Retained package entries.
    Entries,
    /// Bytes in one package entry or encoded record.
    EntryBytes,
    /// Aggregate package bytes.
    TotalBytes,
    /// Semantic soundtrack items.
    Items,
    /// Audio payload bytes retained for insertion or replacement.
    AudioBytes,
    /// Package-metadata records inspected or rewritten.
    MetadataRecords,
    /// Data members inspected or rewritten.
    DataMembers,
    /// Graph references traversed during candidate reopen.
    References,
    /// Parsed wire bytes.
    WireBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Items => "soundtrack items",
            Self::AudioBytes => "audio bytes",
            Self::MetadataRecords => "metadata records",
            Self::DataMembers => "data members",
            Self::References => "references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
        })
    }
}

/// An opaque handle for one item in one immutable package snapshot.
///
/// Handles are package-local capabilities, not durable identifiers. They
/// carry no public native ID and cannot be manufactured by callers. Obtain a
/// handle from Item::handle and use it only with an edit bound to the same
/// immutable package snapshot. A handle remains useful when an edit's
/// operation changes neighbouring positions; the package adapter validates
/// its lineage before publication.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct ItemHandle {
    pub(crate) lineage: [u8; 20],
    pub(crate) occurrence: usize,
}

impl fmt::Debug for ItemHandle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("ItemHandle(..)")
    }
}

impl ItemHandle {
    /// Return the source position captured when this handle was created.
    ///
    /// The position is semantic and zero-based. It is only a diagnostic
    /// convenience; the handle's opaque lineage remains the authoritative
    /// selection capability for an edit.
    #[must_use]
    pub const fn position(self) -> Position {
        Position::new(self.occurrence)
    }

    pub(crate) const fn new(lineage: [u8; 20], occurrence: usize) -> Self {
        Self {
            lineage,
            occurrence,
        }
    }
}

/// A semantic selector for an existing soundtrack item.
///
/// Positions are convenient for callers that already have a checked list;
/// handles provide stronger protection against accidentally applying an item
/// operation to another immutable package snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ItemSelector {
    /// Select an item by its zero-based source-order position.
    Position(Position),
    /// Select an item through a package-local opaque handle.
    Handle(ItemHandle),
}

impl ItemSelector {
    /// Construct a position selector.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Construct a position selector from a zero-based index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Position(Position::new(index))
    }

    /// Construct an opaque-handle selector.
    #[must_use]
    pub const fn handle(handle: ItemHandle) -> Self {
        Self::Handle(handle)
    }

    /// Return the contained position when this is a positional selector.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Position(position) => Some(position),
            Self::Handle(_) => None,
        }
    }
}

impl From<usize> for ItemSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for ItemSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<ItemHandle> for ItemSelector {
    fn from(handle: ItemHandle) -> Self {
        Self::handle(handle)
    }
}

impl From<&Item> for ItemSelector {
    fn from(item: &Item) -> Self {
        Self::handle(item.handle)
    }
}

impl From<Item> for ItemSelector {
    fn from(item: Item) -> Self {
        Self::handle(item.handle)
    }
}

/// A semantic description of one existing soundtrack item.
///
/// Item metadata is intentionally compact: package reads do not retain a
/// second copy of the media payload merely to describe the sequence. The
/// payload can be supplied separately through AudioSource when staging an
/// insertion or replacement.
#[derive(Clone)]
pub struct Item {
    pub(crate) handle: ItemHandle,
    pub(crate) position: Position,
    pub(crate) filename: Box<str>,
    pub(crate) byte_length: usize,
}

impl PartialEq for Item {
    fn eq(&self, other: &Self) -> bool {
        self.position == other.position
            && self.filename == other.filename
            && self.byte_length == other.byte_length
    }
}

impl Eq for Item {}

impl fmt::Debug for Item {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Item")
            .field("position", &self.position)
            .field("filename", &self.filename)
            .field("byte_length", &self.byte_length)
            .finish_non_exhaustive()
    }
}

impl Item {
    /// Return the package-local opaque handle for this item.
    #[must_use]
    pub const fn handle(&self) -> ItemHandle {
        self.handle
    }

    /// Return the item's zero-based playback position in the source snapshot.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the item's zero-based playback index.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.position.get()
    }

    /// Return the safe leaf filename associated with the media payload.
    #[must_use]
    pub fn filename(&self) -> &str {
        &self.filename
    }

    /// Return the filename as the item's semantic name.
    ///
    /// This alias is useful when presenting the sequence in a UI; it has the
    /// same exact value as Item::filename.
    #[must_use]
    pub fn name(&self) -> &str {
        self.filename()
    }

    /// Return the uncompressed media payload length in bytes.
    #[must_use]
    pub const fn byte_length(&self) -> usize {
        self.byte_length
    }

    /// Return the uncompressed media payload length in bytes.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.byte_length()
    }

    /// Return whether the media payload is empty.
    ///
    /// Valid package items are non-empty, but this conventional predicate is
    /// useful to generic collection code and costs no additional traversal.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.byte_length == 0
    }

    pub(crate) fn new(
        lineage: [u8; 20],
        position: Position,
        filename: impl Into<Box<str>>,
        byte_length: usize,
    ) -> Self {
        Self {
            handle: ItemHandle::new(lineage, position.get()),
            position,
            filename: filename.into(),
            byte_length,
        }
    }
}

/// Owned, validated audio content for insertion or replacement.
///
/// The source owns an Arc<[u8]>, making repeated staging and retry paths
/// cheap without exposing mutable bytes. The package adapter independently
/// validates the payload against the selected package's physical limits before
/// allocating data identifiers or publishing a candidate.
#[derive(Clone, PartialEq, Eq)]
pub struct AudioSource {
    pub(crate) filename: Box<str>,
    pub(crate) bytes: Arc<[u8]>,
}

impl fmt::Debug for AudioSource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("AudioSource")
            .field("filename", &self.filename)
            .field("byte_length", &self.bytes.len())
            .finish_non_exhaustive()
    }
}

impl AudioSource {
    /// Construct an owned, bounded audio source.
    ///
    /// Both the filename extension and a bounded byte-prefix signature must
    /// identify audio. The filename is retained as one safe relative leaf;
    /// path separators, NUL, controls, and overlong names are refused before
    /// the source can enter a package edit. Data may be an owned Vec or a
    /// borrowed slice; conversion to the immutable shared representation is
    /// performed exactly once.
    ///
    /// Returns a typed error for an unsafe or overlong filename, an empty or
    /// oversized payload, or a filename/payload that is not recognized as
    /// audio.
    pub fn new(filename: impl Into<Box<str>>, data: impl Into<Box<[u8]>>) -> Result<Self, Error> {
        let filename = filename.into();
        validate_filename(&filename)?;
        let data = data.into();
        validate_payload(&data)?;
        Ok(Self {
            filename,
            bytes: Arc::from(data),
        })
    }

    /// Construct an owned source by copying a borrowed payload.
    ///
    /// This spelling makes the allocation explicit at call sites that hold a
    /// borrowed byte slice while AudioSource::new remains zero-copy for an
    /// owned Vec converted into a boxed slice.
    pub fn from_bytes(filename: impl Into<Box<str>>, data: &[u8]) -> Result<Self, Error> {
        let filename = filename.into();
        validate_filename(&filename)?;
        validate_payload(data)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(data.len())
            .map_err(|_| Error::Allocation { amount: data.len() })?;
        owned.extend_from_slice(data);
        Ok(Self {
            filename,
            bytes: Arc::from(owned.into_boxed_slice()),
        })
    }

    /// Return the safe leaf filename.
    #[must_use]
    pub fn filename(&self) -> &str {
        &self.filename
    }

    /// Return the immutable audio payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the payload length in bytes.
    #[must_use]
    pub fn byte_length(&self) -> usize {
        self.bytes.len()
    }

    /// Return whether the payload is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }
}

impl AsRef<[u8]> for AudioSource {
    fn as_ref(&self) -> &[u8] {
        self.bytes()
    }
}

/// The semantic kind of one staged item operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OperationKind {
    /// Append a new item after all existing items.
    Add,
    /// Insert a new item at a checked sequence position.
    Insert,
    /// Replace one existing item with fresh audio content.
    Replace,
    /// Remove one existing item from the soundtrack sequence.
    ///
    /// Physical media reclamation is deliberately conservative: storage is
    /// retained unless the package owner can prove every native dependency.
    Remove,
}

/// One operation staged against an immutable source snapshot.
pub(crate) enum StagedOperation {
    Add(AudioSource),
    Insert {
        position: Position,
        source: AudioSource,
    },
    Replace {
        selector: ItemSelector,
        source: AudioSource,
    },
    Remove {
        selector: ItemSelector,
    },
}

impl StagedOperation {
    pub(crate) const fn kind(&self) -> OperationKind {
        match self {
            Self::Add(_) => OperationKind::Add,
            Self::Insert { .. } => OperationKind::Insert,
            Self::Replace { .. } => OperationKind::Replace,
            Self::Remove { .. } => OperationKind::Remove,
        }
    }
}

/// A one-operation immutable soundtrack-item edit.
///
/// The edit borrows its exact source package and owns at most one validated
/// AudioSource. Staging never mutates the source. The private package adapter
/// performs selection, physical planning, and commit verification.
pub struct Edit<'a> {
    pub(crate) source: &'a Package,
    pub(crate) operation: Option<StagedOperation>,
    pub(crate) item_count: Option<usize>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "operation",
                &self.operation.as_ref().map(StagedOperation::kind),
            )
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// Append one new item after the existing soundtrack sequence.
    pub fn add(&mut self, source: AudioSource) -> Result<&mut Self, Error> {
        self.stage(StagedOperation::Add(source))
    }

    /// Append one new item after the existing soundtrack sequence.
    pub fn append(&mut self, source: AudioSource) -> Result<&mut Self, Error> {
        self.add(source)
    }

    /// Insert one new item at a final zero-based sequence position.
    pub fn insert(&mut self, position: Position, source: AudioSource) -> Result<&mut Self, Error> {
        self.validate_insertion_position(position)?;
        self.stage(StagedOperation::Insert { position, source })
    }

    /// Replace one existing item selected by a position, handle, or item.
    pub fn replace(
        &mut self,
        selector: impl Into<ItemSelector>,
        source: AudioSource,
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        self.validate_item_selector(selector)?;
        self.stage(StagedOperation::Replace { selector, source })
    }

    /// Remove one existing item selected by a position, handle, or item.
    pub fn remove(&mut self, selector: impl Into<ItemSelector>) -> Result<&mut Self, Error> {
        let selector = selector.into();
        self.validate_item_selector(selector)?;
        self.stage(StagedOperation::Remove { selector })
    }

    /// Return the staged operation kind, if any.
    #[must_use]
    pub fn operation(&self) -> Option<OperationKind> {
        self.operation.as_ref().map(StagedOperation::kind)
    }

    fn stage(&mut self, operation: StagedOperation) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        self.operation = Some(operation);
        Ok(self)
    }

    /// Construct an edit after the package adapter has bounded its source
    /// sequence. The count is retained only for stage-time positional checks;
    /// the physical adapter repeats the complete selection during commit.
    pub(crate) const fn new_with_item_count(
        source: &'a Package,
        item_count: Option<usize>,
    ) -> Self {
        Self {
            source,
            operation: None,
            item_count,
        }
    }

    fn validate_insertion_position(&self, position: Position) -> Result<(), Error> {
        let Some(item_count) = self.item_count else {
            return Ok(());
        };
        if position.get() > item_count {
            return Err(Error::InsertionOutOfRange {
                position,
                item_count,
            });
        }
        Ok(())
    }

    fn validate_item_selector(&self, selector: ItemSelector) -> Result<(), Error> {
        let Some(item_count) = self.item_count else {
            return Ok(());
        };
        let position = match selector {
            ItemSelector::Position(position) => position,
            ItemSelector::Handle(handle) => handle.position(),
        };
        if position.get() >= item_count {
            return Err(Error::SourcePositionNotFound { position });
        }
        Ok(())
    }
}

/// An exact-source-checked reversible soundtrack-item patch.
///
/// The complete source and target package artifacts remain private and are
/// shared on clone/inversion. Semantic item summaries are retained so callers
/// can inspect a patch without receiving package bytes or native identities.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    pub(crate) artifacts: ExactArtifacts,
    pub(crate) before: Box<[Item]>,
    pub(crate) after: Box<[Item]>,
    pub(crate) operation: OperationKind,
    pub(crate) inverse_operation: OperationKind,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("operation", &self.operation)
            .field("before_items", &self.before.len())
            .field("after_items", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the complete semantic item sequence required from the source.
    #[must_use]
    pub fn before(&self) -> &[Item] {
        &self.before
    }

    /// Return the complete semantic item sequence produced by the target.
    #[must_use]
    pub fn after(&self) -> &[Item] {
        &self.after
    }

    /// Return the staged operation kind.
    #[must_use]
    pub const fn operation(&self) -> OperationKind {
        self.operation
    }

    /// Return the source package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact semantic and byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact target-to-source inverse in shared-handle work.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            before: self.after.clone(),
            after: self.before.clone(),
            operation: self.inverse_operation,
            inverse_operation: self.operation,
        }
    }
}

/// Compact evidence about one soundtrack-item publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    pub(crate) changed: bool,
    pub(crate) touched_components: usize,
    pub(crate) full_reparse_performed: bool,
}

impl Diagnostics {
    pub(crate) const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    pub(crate) const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten package components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether a changed candidate was fully reopened for verification.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one immutable soundtrack-item transaction.
#[must_use = "a soundtrack-item commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    pub(crate) package: Package,
    pub(crate) patch: Patch,
    pub(crate) diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the verified immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn validate_filename(filename: &str) -> Result<(), Error> {
    if filename.is_empty() {
        return Err(Error::EmptyFilename);
    }
    let length = filename.len();
    if length > MAX_FILENAME_BYTES {
        return Err(Error::FilenameTooLong {
            observed: length,
            maximum: MAX_FILENAME_BYTES,
        });
    }
    if filename
        .bytes()
        .any(|byte| byte == b'/' || byte == b'\\' || byte == 0 || byte.is_ascii_control())
        || filename == "."
        || filename == ".."
    {
        return Err(Error::InvalidFilename);
    }
    let Some(dot) = filename.rfind('.') else {
        return Err(Error::UnsupportedAudioFilename);
    };
    if dot == 0 || dot + 1 == filename.len() {
        return Err(Error::UnsupportedAudioFilename);
    }
    if MediaType::from_extension(&filename[dot + 1..]) != MediaType::Audio {
        return Err(Error::UnsupportedAudioFilename);
    }
    Ok(())
}

fn validate_payload(data: &[u8]) -> Result<(), Error> {
    if data.is_empty() {
        return Err(Error::EmptyAudio);
    }
    if data.len() > MAX_AUDIO_BYTES {
        return Err(Error::AudioTooLarge {
            observed: data.len(),
            maximum: MAX_AUDIO_BYTES,
        });
    }
    if MediaType::from_bytes(data) != MediaType::Audio {
        return Err(Error::UnsupportedAudioData);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn audio(filename: &str, bytes: &[u8]) -> AudioSource {
        AudioSource::from_bytes(filename, bytes).expect("audio source")
    }

    #[test]
    fn source_owns_bounded_audio_without_exposing_mutable_bytes() {
        let source = audio("theme.mp3", b"ID3\x04\0\0audio");
        assert_eq!(source.filename(), "theme.mp3");
        assert_eq!(source.bytes(), b"ID3\x04\0\0audio");
        assert_eq!(source.byte_length(), 11);
        assert!(!source.is_empty());
    }

    #[test]
    fn source_rejects_path_traversal_and_non_audio_values() {
        assert_eq!(
            AudioSource::from_bytes("../theme.mp3", b"ID3\x04\0\0"),
            Err(Error::InvalidFilename)
        );
        assert_eq!(
            AudioSource::from_bytes("theme.bin", b"ID3\x04\0\0"),
            Err(Error::UnsupportedAudioFilename)
        );
        assert_eq!(
            AudioSource::from_bytes("theme.mp3", b"not audio"),
            Err(Error::UnsupportedAudioData)
        );
        assert_eq!(
            AudioSource::from_bytes("theme.mp3", &[]),
            Err(Error::EmptyAudio)
        );
    }

    #[test]
    fn selector_conversions_remain_position_or_opaque_handle_only() {
        let handle = ItemHandle::new([42; 20], 3);
        assert_eq!(ItemSelector::from(3usize), ItemSelector::index(3));
        assert_eq!(ItemSelector::from(Position::new(3)), ItemSelector::index(3));
        assert_eq!(ItemSelector::from(handle), ItemSelector::handle(handle));
        assert_eq!(handle.position(), Position::new(3));
        assert_eq!(format!("{handle:?}"), "ItemHandle(..)");
    }

    #[test]
    fn item_debug_does_not_contain_opaque_handle_identity() {
        let item = Item::new([7; 20], Position::new(2), "theme.m4a", 32);
        let debug = format!("{item:?}");
        assert!(debug.contains("position"));
        assert!(debug.contains("theme.m4a"));
        assert!(!debug.contains("lineage"));
        assert!(!debug.contains("ItemHandle"));
        assert_eq!(item.handle().position(), Position::new(2));
        assert_eq!(item.index(), 2);
        assert_eq!(item.name(), item.filename());
        assert_eq!(item.len(), 32);
        assert!(!item.is_empty());
    }
}
