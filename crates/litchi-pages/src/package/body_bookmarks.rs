//! Selector-first ranged bookmarks of the rooted Pages body.
//!
//! A bookmark is a `TSWP.BookmarkFieldArchive` object referenced by the body
//! storage's bookmark table over a nonempty UTF-16 range. The public values
//! carry the range and the archive-free settings; bookmarks are selected by
//! source order or exact name. Adding one creates the object above the
//! package watermark; removing one deletes it with its metadata
//! registrations (see [`super::body_storage_objects`]).

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_text::TextRange;
use litchi_iwa_text::bookmark::{Name, Settings};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{self as graph, GraphError, GraphLimit, map_wire_error};
use super::body_storage_objects::{self as objects, Storage};
use litchi_iwa_text_wire::{ranged, smart_field};

const BOOKMARK_MESSAGE_TYPE: u32 = 2_035;

/// A finite resource charged by one body-bookmark operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and reassembly work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure from a body-bookmark read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The body storage or a bookmark object is malformed.
    #[error("the Pages body bookmark source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages body bookmark {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body bookmarks")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested bookmarks.
    #[error("the edited Pages body bookmarks failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body bookmark patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No bookmark matches the selector.
    #[error("no Pages body bookmark matches the selector")]
    NotFound,
    /// Two bookmarks share the selected name.
    #[error("several Pages body bookmarks share the selected name")]
    AmbiguousName,
    /// The range is outside the body text or splits a character.
    #[error("the bookmark range is outside the body text or splits a character")]
    InvalidRange,
    /// The range overlaps another bookmark.
    #[error("the bookmark range overlaps another bookmark")]
    Overlap,
    /// The bookmark object is referenced by more than the body storage.
    #[error("the bookmark object is shared outside the body storage")]
    SharedObject,
    /// A second operation was staged on one edit.
    #[error("one body bookmark edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body-bookmark edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound => RefusalKind::NotFound,
            Self::AmbiguousName => RefusalKind::UnsupportedTarget,
            Self::InvalidRange | Self::Overlap => RefusalKind::InvalidValue,
            Self::SharedObject | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::OperationAlreadyStaged => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

/// One ranged bookmark of the body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bookmark {
    range: TextRange,
    settings: Settings,
}

impl Bookmark {
    /// The bookmarked UTF-16 range of the body text.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

    /// The bookmark's name and visibility.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }
}

/// Selects one body bookmark.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// The bookmark at this zero-based position in body order.
    Position(Position),
    /// The bookmark with this exact name.
    Name(&'a str),
}

impl Selector<'_> {
    /// Select by zero-based body-order index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Position(Position::new(index))
    }
}

impl From<usize> for Selector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(name: &'a str) -> Self {
        Self::Name(name)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    Add(TextRange, Settings),
    Update(usize, TextRange, Settings),
    Remove(usize),
}

/// One bookmark operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    before: Arc<[Bookmark]>,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operation", &self.operation)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The bookmarks when the edit began.
    #[must_use]
    pub fn bookmarks(&self) -> &[Bookmark] {
        &self.before
    }

    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.operation = Some(operation);
        Ok(self)
    }

    fn select(&self, selector: Selector<'_>) -> Result<usize, Error> {
        select(&self.before, selector)
    }

    /// Stage a new bookmark over a nonempty body range.
    ///
    /// # Errors
    ///
    /// Refuses a second operation; range checks run at commit.
    pub fn add(&mut self, range: TextRange, settings: Settings) -> Result<&mut Self, Error> {
        self.stage(Operation::Add(range, settings))
    }

    /// Stage a new range and settings for the selected bookmark.
    ///
    /// # Errors
    ///
    /// Refuses a missing or ambiguous selection and a second operation.
    pub fn update<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        range: TextRange,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        let position = self.select(selector.into())?;
        self.stage(Operation::Update(position, range, settings))
    }

    /// Stage removal of the selected bookmark (its text stays).
    ///
    /// # Errors
    ///
    /// Refuses a missing or ambiguous selection and a second operation.
    pub fn remove<'s>(&mut self, selector: impl Into<Selector<'s>>) -> Result<&mut Self, Error> {
        let position = self.select(selector.into())?;
        self.stage(Operation::Remove(position))
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-bookmark patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    before: Arc<[Bookmark]>,
    after: Arc<[Bookmark]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before.len())
            .field("after", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The bookmarks required before this patch applies.
    #[must_use]
    pub fn before(&self) -> &[Bookmark] {
        &self.before
    }

    /// The bookmarks this patch produces.
    #[must_use]
    pub fn after(&self) -> &[Bookmark] {
        &self.after
    }

    /// Whether both the bookmarks and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Content-free diagnostics of one body-bookmark publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether package bytes and bookmarks changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// The validated immutable result of one body-bookmark transaction.
#[must_use = "a body-bookmark commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the body's ranged bookmarks in body order.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the body storage or a bookmark is malformed.
    pub fn body_bookmarks(&self) -> Result<Vec<Bookmark>, Error> {
        let (storage, _) = objects::body_storage(self)?;
        Ok(read(self, &storage)?
            .into_iter()
            .map(|(_, bookmark)| bookmark)
            .collect())
    }

    /// Begin one bookmark add, update or removal on the body.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the body storage or a bookmark is malformed.
    pub fn edit_body_bookmarks(&self) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            before: self.body_bookmarks()?.into(),
            operation: None,
        })
    }

    /// Replay an exact-source body-bookmark patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the bookmarks no longer match its `before` state.
    pub fn apply_body_bookmarks(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        if *self.body_bookmarks()? != *patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics { changed: false },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        if *candidate.body_bookmarks()? != *patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}

fn select(bookmarks: &[Bookmark], selector: Selector<'_>) -> Result<usize, Error> {
    match selector {
        Selector::Position(position) => {
            if position.get() < bookmarks.len() {
                Ok(position.get())
            } else {
                Err(Refusal::NotFound.into())
            }
        },
        Selector::Name(name) => {
            let mut matches = bookmarks
                .iter()
                .enumerate()
                .filter(|(_, bookmark)| bookmark.settings.name().map(Name::as_str) == Some(name));
            match (matches.next(), matches.next()) {
                (Some((position, _)), None) => Ok(position),
                (Some(_), Some(_)) => Err(Refusal::AmbiguousName.into()),
                (None, _) => Err(Refusal::NotFound.into()),
            }
        },
    }
}

/// The body's bookmarks with their native object identifiers.
fn read(package: &Package, storage: &Storage) -> Result<Vec<(u64, Bookmark)>, Error> {
    let entries = ranged::entries(&storage.payload, ranged::Table::Bookmark, wire_limits())
        .map_err(map_ranged_error)?;
    let archive = package
        .state
        .source
        .components()
        .get(storage.component.as_ref())
        .ok_or(Error::InvalidSource)?
        .archive();
    let mut bookmarks = Vec::new();
    let mut seen = std::collections::HashSet::new();
    for entry in entries {
        let object = entry.object();
        let Some(found) = archive.object(object) else {
            return Err(Error::InvalidSource);
        };
        let Some(settings) = bookmark_settings(found)? else {
            continue;
        };
        if !seen.insert(object) {
            return Err(Error::InvalidSource);
        }
        bookmarks.push((
            object,
            Bookmark {
                range: entry.range(),
                settings,
            },
        ));
    }
    Ok(bookmarks)
}

/// The settings of a bookmark object, `None` for another smart object.
fn bookmark_settings(object: &litchi_iwa_core::ArchiveObject) -> Result<Option<Settings>, Error> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == BOOKMARK_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let message = match payloads.as_slice() {
        [] => return Ok(None),
        [message] if object.messages.len() == 1 => message,
        _ => return Err(Error::InvalidSource),
    };
    let (_, settings) =
        smart_field::bookmark(&message.data, wire_limits()).map_err(|_| Error::InvalidSource)?;
    Ok(Some(settings))
}

const fn wire_limits() -> litchi_iwa_common::WireLimits {
    litchi_iwa_common::WireLimits::DEFAULT
}

fn map_ranged_error(error: ranged::Error) -> Error {
    match error {
        ranged::Error::Refused(ranged::Refusal::NotABoundary) => Refusal::InvalidRange.into(),
        ranged::Error::Refused(ranged::Refusal::Occupied) => Refusal::Overlap.into(),
        ranged::Error::Refused(_) => Refusal::SharedObject.into(),
        ranged::Error::Allocation { amount } => Error::Allocation { amount },
        ranged::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        map_wire_error(error).into()
    }
}

fn check_range(storage: &Storage, range: TextRange) -> Result<(u32, u32), Error> {
    let (start, end) = (range.start().utf16_index(), range.end().utf16_index());
    if end > storage.utf16_len() || !objects::scalar_boundaries(&storage.text, [start, end]) {
        return Err(Refusal::InvalidRange.into());
    }
    Ok((start, end))
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let unchanged = |before: Arc<[Bookmark]>| Commit {
        package: source.clone(),
        patch: Patch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), Arc::clone(&source_bytes)),
            before: Arc::clone(&before),
            after: before,
        },
        diagnostics: Diagnostics { changed: false },
    };
    let Some(operation) = edit.operation else {
        return Ok(unchanged(edit.before));
    };
    if let Operation::Update(position, range, settings) = &operation
        && edit.before[*position].range == *range
        && edit.before[*position].settings == *settings
    {
        return Ok(unchanged(edit.before));
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let (storage, _) = objects::body_storage(source)?;
    let current = read(source, &storage)?;
    if current
        .iter()
        .map(|(_, bookmark)| bookmark)
        .ne(edit.before.iter())
    {
        return Err(Error::PatchConflict);
    }
    let metadata = objects::metadata(source)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archive = graph::editable_archive(source, &storage.component)?;
    let mut metadata_payload = metadata.as_ref().map(|metadata| metadata.payload.clone());
    let mut expected = edit.before.to_vec();
    let mut added = None;
    let mut removed = None;
    let rewrite;
    match operation {
        Operation::Add(range, settings) => {
            check_range(&storage, range)?;
            let identifier = objects::next_identifier(source, metadata.as_ref())?;
            rewrite = Some(
                ranged::rewrite(
                    &storage.payload,
                    ranged::Table::Bookmark,
                    &ranged::Change::Add {
                        object: identifier,
                        range,
                    },
                    wire_limits(),
                )
                .map_err(map_ranged_error)?,
            );
            let uuid = smart_field::TextAttributeUuid::new(&objects::text_attribute_uuid()?)
                .map_err(|_| Error::InvalidSource)?;
            let object = objects::new_object(
                source,
                identifier,
                BOOKMARK_MESSAGE_TYPE,
                smart_field::encode_bookmark(&uuid, &settings, wire_limits())
                    .map_err(|_| Error::InvalidSource)?,
            )?;
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
            if let Some(metadata) = &metadata {
                metadata_payload = Some(metadata.raised(source, identifier)?);
            }
            added = Some(identifier);
            let position =
                expected.partition_point(|bookmark| bookmark.range.start() < range.start());
            expected.insert(position, Bookmark { range, settings });
        },
        Operation::Update(position, range, settings) => {
            let identifier = current[position].0;
            check_range(&storage, range)?;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            rewrite = Some(
                ranged::rewrite(
                    &storage.payload,
                    ranged::Table::Bookmark,
                    &ranged::Change::Move {
                        object: identifier,
                        range,
                    },
                    wire_limits(),
                )
                .map_err(map_ranged_error)?,
            );
            let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
            let data = smart_field::bookmark_with_settings(
                &object.messages[0].data,
                &settings,
                wire_limits(),
            )
            .map_err(|_| Error::InvalidSource)?;
            object
                .replace_message_preserving_header(
                    0,
                    litchi_iwa_core::RawMessage {
                        type_: BOOKMARK_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
            expected[position] = Bookmark { range, settings };
            expected.sort_by_key(|bookmark| bookmark.range.start());
        },
        Operation::Remove(position) => {
            let identifier = current[position].0;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            rewrite = Some(
                ranged::rewrite(
                    &storage.payload,
                    ranged::Table::Bookmark,
                    &ranged::Change::Remove { object: identifier },
                    wire_limits(),
                )
                .map_err(map_ranged_error)?,
            );
            if let Some(metadata) = &metadata {
                let mut payload = metadata.payload.clone();
                if let Some(component) = metadata.component_for_entry(&storage.component)? {
                    payload = objects::remove_external_references(&payload, component, identifier)?;
                }
                if let Some(component) = metadata.uuid_owner(identifier)? {
                    payload = objects::remove_object_uuid(&payload, component, identifier)?;
                }
                metadata_payload = Some(payload);
            }
            archive
                .remove_object(identifier)
                .ok_or(Error::InvalidSource)?;
            removed = Some(identifier);
            expected.remove(position);
        },
    }
    // The storage's bookmark table and its object-reference metadata.
    let payload = rewrite.ok_or(Error::InvalidSource)?.into_payload();
    let storage_object = archive
        .object_mut(storage.identifier.get())
        .ok_or(Error::InvalidSource)?;
    storage_object
        .replace_message_preserving_header(
            storage.message_index,
            litchi_iwa_core::RawMessage {
                type_: storage.message_type,
                data: payload,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let info = storage_object
        .archive_info
        .message_infos
        .get_mut(storage.message_index)
        .ok_or(Error::InvalidSource)?;
    if let Some(identifier) = added {
        if info.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(identifier);
    }
    if let Some(identifier) = removed {
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
        {
            return Err(Error::InvalidSource);
        }
        info.object_references
            .retain(|candidate| *candidate != identifier);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|candidate| *candidate != identifier);
        }
    }
    let mut archives = vec![(Arc::clone(&storage.component), archive)];
    if let (Some(metadata), Some(mut payload)) = (&metadata, metadata_payload) {
        if let Some(identifier) = removed {
            let remaining = objects::max_identifier(source, &archives);
            payload = objects::released(&payload, metadata.last, &[identifier], remaining)?;
        }
        if payload != metadata.payload {
            let mut metadata_archive = graph::editable_archive(source, &metadata.component)?;
            metadata_archive
                .object_mut(metadata.object)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    metadata.message_index,
                    litchi_iwa_core::RawMessage {
                        type_: 11_006,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
            archives.push((Arc::clone(&metadata.component), metadata_archive));
        }
    }
    let candidate = graph::publish_archives(source, archives)?;
    let after = candidate.body_bookmarks()?;
    if after != expected {
        return Err(Error::Verification);
    }
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            before: edit.before,
            after: after.into(),
        },
        diagnostics: Diagnostics { changed: true },
    })
}

/// Refuse a bookmark object that another object's metadata references.
fn require_exclusive(package: &Package, storage: u64, object: u64) -> Result<(), Error> {
    let owners = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter(|candidate| {
            candidate
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&object))
        })
        .filter_map(|candidate| candidate.archive_info.identifier)
        .collect::<Vec<_>>();
    if owners != [storage] {
        return Err(Refusal::SharedObject.into());
    }
    Ok(())
}
