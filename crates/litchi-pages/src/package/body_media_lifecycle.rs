//! Selector-first lifecycle transactions of body images, movies and audio
//! clips: removal (this module's first slice), with duplication, data
//! replacement and creation following the same private graph proof.
//!
//! The graph of one body media drawable is what the migration host
//! deletes (record 0371): the drawable, the private objects of its title
//! and caption (a stand-in, or the caption graph), and its body attachment,
//! all in the drawable's component. Removal deletes the anchoring U+FFFC
//! and its attachment entry, drops the drawable from the document z-order,
//! releases the graph's data-reference owners, UUID registrations and
//! identifier watermark, and deletes every embedded data asset no object,
//! component or data-metadata entry references any more.

use std::collections::HashSet;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::wire::remove_repeated_length_delimited_field_where;
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_text::TextPosition;
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_caption::{self as caption, Kind as LabelKind};
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, ObjectIndex, kind_name,
};
use super::body_storage_objects as objects;
use crate::selector::DrawableSelector;

const ZORDER_MESSAGE_TYPE: u32 = 10_015;
const ZORDER_DRAWABLES_FIELD: u32 = 1;
const DRAWABLE_COMMENT_FIELD: u32 = 6;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
/// The document component, which registers body drawables' UUIDs and data
/// references.
const DOCUMENT_COMPONENT_IDENTIFIER: u64 = 1;

/// A finite resource charged by one body-media lifecycle operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected by the graph proof.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal, text rewrite and reassembly work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a body-media lifecycle transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted media graph is malformed.
    #[error("the selected Pages body-media source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages body-media {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages body-media operation")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested change.
    #[error("the edited Pages body media failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-media patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body drawable of the selected kind is at the selected position.
    #[error("the Pages body has no {kind} at position {position:?}")]
    NotFound {
        /// The selected kind.
        kind: &'static str,
        /// The selected position.
        position: Position,
    },
    /// The selected kind is not body media (shapes have their own family).
    #[error("a Pages body {kind} is not body media")]
    NotMedia {
        /// The selected kind.
        kind: &'static str,
    },
    /// The drawable carries a comment, which this family does not remove.
    #[error("the selected Pages body media carries a comment")]
    Commented,
    /// The media graph shares an object or spans components in a way this
    /// family does not prove.
    #[error("the selected Pages body media has a shared or split graph")]
    SharedGraph,
    /// No operation was staged.
    #[error("no body-media operation is staged")]
    NoStagedOperation,
    /// A second operation was staged on one edit.
    #[error("one body-media edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body-media edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::NotMedia { .. } => RefusalKind::UnsupportedTarget,
            Self::Commented | Self::SharedGraph | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::NoStagedOperation | Self::OperationAlreadyStaged => {
                RefusalKind::UnsupportedOperation
            },
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
                },
                observed,
                maximum,
            },
        }
    }
}

impl From<caption::Error> for Error {
    fn from(error: caption::Error) -> Self {
        match error {
            caption::Error::Allocation { amount } => Self::Allocation { amount },
            caption::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: LimitKind::WireWork,
                observed,
                maximum,
            },
            _ => Self::InvalidSource,
        }
    }
}

/// The staged operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operation {
    Remove(DrawableSelector),
    Duplicate(DrawableSelector, TextPosition),
}

/// One body-media lifecycle operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
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
    /// Stage the removal of one body image, movie or audio clip with its
    /// private graph and the data assets only it uses.
    ///
    /// # Errors
    ///
    /// Refuses a shape and a second operation.
    pub fn remove(&mut self, selector: impl Into<DrawableSelector>) -> Result<&mut Self, Error> {
        let selector = selector.into();
        media_kind(selector)?;
        self.stage(Operation::Remove(selector))
    }

    /// Stage a copy of one body image, movie or audio clip anchored at a
    /// UTF-16 body position: its graph is cloned with fresh identifiers,
    /// offset as Pages places a duplicate (12 pt, 30 pt for audio), and its
    /// data assets are shared with the source.
    ///
    /// # Errors
    ///
    /// Refuses a shape and a second operation.
    pub fn duplicate(
        &mut self,
        selector: impl Into<DrawableSelector>,
        anchor: TextPosition,
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        media_kind(selector)?;
        self.stage(Operation::Duplicate(selector, anchor))
    }

    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.operation = Some(operation);
        Ok(self)
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self.operation.ok_or(Refusal::NoStagedOperation)?;
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let (package, diagnostics) = match operation {
            Operation::Remove(selector) => remove(self.source, selector)?,
            Operation::Duplicate(selector, anchor) => duplicate(self.source, selector, anchor)?,
        };
        let target = package.state.source.shared_source();
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics,
        })
    }
}

/// Exact-source reversible body-media lifecycle patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the exact artifact is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics of one body-media lifecycle publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    created_objects: usize,
    removed_objects: usize,
    removed_assets: usize,
}

impl Diagnostics {
    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Native objects the publication created.
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }

    /// Native objects the publication deleted.
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }

    /// Embedded data assets the publication deleted.
    #[must_use]
    pub const fn removed_assets(self) -> usize {
        self.removed_assets
    }
}

/// The validated immutable result of one body-media lifecycle transaction.
#[must_use = "a body-media commit contains the validated package snapshot"]
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
    /// Begin a lifecycle operation on the body's images, movies and audio
    /// clips.
    #[must_use]
    pub const fn edit_body_media(&self) -> Edit<'_> {
        Edit {
            source: self,
            operation: None,
        }
    }

    /// Replay an exact-source body-media lifecycle patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_media(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            graph::reopen(self, patch.artifacts.target())?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                ..Diagnostics::default()
            },
        })
    }
}

fn media_kind(selector: DrawableSelector) -> Result<(), Error> {
    if matches!(selector, DrawableSelector::Shape(_)) {
        return Err(Refusal::NotMedia {
            kind: kind_name(selector),
        }
        .into());
    }
    Ok(())
}

/// The private graph of one body media drawable.
struct MediaGraph {
    drawable: NonZeroU64,
    attachment: NonZeroU64,
    character_index: u32,
    component: Arc<str>,
    /// Drawable, title objects, caption objects, attachment (host order).
    objects: Vec<u64>,
    /// `(data, object)` for every data reference the graph's messages list.
    data_references: Vec<(u64, u64)>,
    /// Graph objects registered with a UUID in the document component.
    uuid_objects: Vec<u64>,
    selector: DrawableSelector,
    zorder: NonZeroU64,
}

fn resolve(package: &Package, selector: DrawableSelector) -> Result<MediaGraph, Error> {
    media_kind(selector)?;
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let view = budget.parse(selected.drawable_payload, 3)?;
    if graph::unique_field(&view, DRAWABLE_COMMENT_FIELD, 2)?.is_some() {
        return Err(Refusal::Commented.into());
    }
    let drawable = selected.anchor.drawable;
    let mut graph_objects = vec![drawable.get()];
    for kind in [LabelKind::Title, LabelKind::Caption] {
        graph_objects.extend(
            caption::label_objects(package, selector, kind)?
                .into_iter()
                .map(NonZeroU64::get),
        );
    }
    graph_objects.push(selected.anchor.attachment.get());
    let unique = graph_objects.iter().collect::<HashSet<_>>();
    if unique.len() != graph_objects.len() {
        return Err(Refusal::SharedGraph.into());
    }
    let mut data_references = Vec::new();
    for identifier in &graph_objects {
        let located = index.get(NonZeroU64::new(*identifier).ok_or(Error::InvalidSource)?)?;
        if located.component_index != selected.drawable.component_index {
            return Err(Refusal::SharedGraph.into());
        }
        for info in &located.object.archive_info.message_infos {
            data_references.extend(
                info.data_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.data_references),
                    )
                    .map(|data| (*data, *identifier)),
            );
        }
    }
    let (_, zorder) = root_zorder(package, &mut budget)?;
    let mut uuid_objects = Vec::new();
    if let Some(metadata) = objects::metadata(package)? {
        for identifier in &graph_objects {
            if metadata.uuid_owner(*identifier)? == Some(DOCUMENT_COMPONENT_IDENTIFIER) {
                uuid_objects.push(*identifier);
            }
        }
    }
    Ok(MediaGraph {
        uuid_objects,
        selector,
        drawable,
        attachment: selected.anchor.attachment,
        character_index: selected.anchor.character_index,
        component: index.component_name(selected.drawable.component_index)?,
        objects: graph_objects,
        data_references,
        zorder: zorder.ok_or(Error::InvalidSource)?,
    })
}

fn root_zorder(package: &Package, budget: &mut Budget) -> Result<((), Option<NonZeroU64>), Error> {
    Ok(((), graph::root_zorder(package, budget)?))
}

/// The z-order payload without `drawable`, and the message index.
fn zorder_without(
    archive: &mut Archive,
    zorder: NonZeroU64,
    drawable: u64,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let object = archive
        .object_mut(zorder.get())
        .ok_or(Error::InvalidSource)?;
    let (message_index, message) = graph::unique_message(object, ZORDER_MESSAGE_TYPE)?;
    let payload = message.data.clone();
    let mut count = 0usize;
    for reference in litchi_iwa_common::wire::repeated_length_delimited_payloads(
        &payload,
        ZORDER_DRAWABLES_FIELD,
    )
    .map_err(graph::map_wire_error)?
    {
        let budget_view = litchi_iwa_common::wire::WireView::parse(
            reference,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(graph::map_wire_error)?;
        let identifier = graph::unique_field(&budget_view, 1, 0)?
            .map(|field| graph::canonical_varint(field.payload()))
            .transpose()?;
        count += usize::from(identifier == Some(drawable));
    }
    if count != 1 {
        return Err(Refusal::SharedGraph.into());
    }
    let data = remove_repeated_length_delimited_field_where(
        &payload,
        ZORDER_DRAWABLES_FIELD,
        |reference| {
            let view = litchi_iwa_common::wire::WireView::parse(
                reference,
                litchi_iwa_common::WireLimits::default(),
            )?;
            Ok::<bool, litchi_iwa_common::Error>(
                view.fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .is_some_and(|(value, _)| value == drawable),
            )
        },
    )
    .map_err(graph::map_wire_error)?;
    object
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: ZORDER_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(graph::map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Error::InvalidSource)?;
    info.object_references
        .retain(|identifier| *identifier != drawable);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|identifier| *identifier != drawable);
    }
    Ok(())
}

fn editable<'a>(
    source: &Package,
    archives: &'a mut Vec<(Arc<str>, Archive)>,
    component: &str,
) -> Result<&'a mut Archive, Error> {
    let position = match archives
        .iter()
        .position(|(name, _)| name.as_ref() == component)
    {
        Some(position) => position,
        None => {
            archives.push((
                Arc::from(component),
                graph::editable_archive(source, component)?,
            ));
            archives.len() - 1
        },
    };
    Ok(&mut archives[position].1)
}

fn remove(source: &Package, selector: DrawableSelector) -> Result<(Package, Diagnostics), Error> {
    let media = resolve(source, selector)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();

    // The anchoring U+FFFC and its attachment entry.
    let (mut storage, _) = objects::body_storage(source)?;
    let at = media.character_index;
    if storage.text.encode_utf16().nth(at as usize) != Some(0xfffc) {
        return Err(Error::InvalidSource);
    }
    let removed = objects::splice_text(source, &mut storage, at, at + 1, "")?;
    if removed != [media.attachment.get()] {
        return Err(Refusal::SharedGraph.into());
    }
    {
        let archive = editable(source, &mut archives, &storage.component)?;
        archive
            .object_mut(storage.identifier.get())
            .ok_or(Error::InvalidSource)?
            .replace_message_pruning_object_references_preserving_header(
                storage.message_index,
                RawMessage {
                    type_: storage.message_type,
                    data: storage.payload.clone(),
                },
                &[media.attachment.get()],
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }

    // The document z-order.
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let component = index.component_name(index.get(media.zorder)?.component_index)?;
        let archive = editable(source, &mut archives, &component)?;
        zorder_without(archive, media.zorder, media.drawable.get(), archive_limits)?;
    }

    // The graph objects.
    {
        let archive = editable(source, &mut archives, &media.component)?;
        for identifier in &media.objects {
            archive
                .remove_object(*identifier)
                .ok_or(Error::InvalidSource)?;
        }
    }
    let removed_set = media.objects.iter().copied().collect::<HashSet<_>>();
    for (_, archive) in &archives {
        if archive.objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|info| {
                info.object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| removed_set.contains(identifier))
            })
        }) {
            return Err(Refusal::SharedGraph.into());
        }
    }
    for component in source.state.source.components().iter() {
        if archives
            .iter()
            .any(|(name, _)| name.as_ref() == component.name())
        {
            continue;
        }
        if component.archive().objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|info| {
                info.object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| removed_set.contains(identifier))
            })
        }) {
            return Err(Refusal::SharedGraph.into());
        }
    }

    // Package metadata and data assets.
    let mut deletions = Vec::new();
    let mut removed_assets = 0usize;
    if let Some(metadata) = objects::metadata(source)? {
        let mut payload = metadata.payload.clone();
        for (data, object) in &media.data_references {
            payload = objects::adjust_data_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                *data,
                *object,
                false,
            )?;
        }
        for identifier in &media.objects {
            if metadata.uuid_owner(*identifier)? == Some(DOCUMENT_COMPONENT_IDENTIFIER) {
                payload = objects::remove_object_uuid(
                    &payload,
                    DOCUMENT_COMPONENT_IDENTIFIER,
                    *identifier,
                )?;
            }
        }
        let remaining = objects::max_identifier(source, &archives);
        payload = objects::released(&payload, metadata.last, &media.objects, remaining)?;

        let mut referenced = objects::referenced_data(&payload)?;
        for component in source.state.source.components().iter() {
            let staged = archives
                .iter()
                .find(|(name, _)| name.as_ref() == component.name())
                .map_or(component.archive(), |(_, archive)| archive);
            for object in &staged.objects {
                for info in &object.archive_info.message_infos {
                    referenced.extend(info.data_references.iter().copied());
                    referenced.extend(
                        info.field_infos
                            .iter()
                            .flat_map(|field| field.data_references.iter().copied()),
                    );
                }
                if metadata.data_metadata_map.is_some()
                    && object.archive_info.identifier == metadata.data_metadata_map
                {
                    for message in &object.messages {
                        if message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE {
                            referenced.extend(objects::data_metadata_entries(&message.data)?);
                        }
                    }
                }
            }
        }
        let mut candidates = media
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<Vec<_>>();
        candidates.sort_unstable();
        candidates.dedup();
        for data in candidates {
            if referenced.contains(&data) {
                continue;
            }
            if let Some(member) = objects::data_member(&payload, data)? {
                if source
                    .state
                    .source
                    .package()
                    .iter()
                    .any(|entry| entry.name() == member)
                {
                    deletions.push(member);
                }
            }
            payload = objects::remove_data_info(&payload, data)?;
            removed_assets += 1;
        }
        if payload != metadata.payload {
            let archive = editable(source, &mut archives, &metadata.component)?;
            archive
                .object_mut(metadata.object)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    metadata.message_index,
                    RawMessage {
                        type_: METADATA_MESSAGE_TYPE,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
        }
    }
    let deletions = deletions.iter().map(String::as_str).collect::<Vec<_>>();
    let candidate = graph::publish_changes(source, archives, &[], &deletions)?;
    verify_removed(&candidate, selector, &media)?;
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            removed_objects: media.objects.len(),
            removed_assets,
            ..Diagnostics::default()
        },
    ))
}

fn verify_removed(
    candidate: &Package,
    selector: DrawableSelector,
    media: &MediaGraph,
) -> Result<(), Error> {
    let mut budget = Budget::new(candidate)?;
    let index = ObjectIndex::new(candidate, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    if body
        .anchors
        .iter()
        .any(|anchor| anchor.drawable == media.drawable)
    {
        return Err(Error::Verification);
    }
    // The selector still resolves (the next drawable of the kind) or not;
    // either is valid, so only the graph's absence is checked.
    let _ = selector;
    for identifier in &media.objects {
        if NonZeroU64::new(*identifier).is_some_and(|identifier| index.get(identifier).is_ok()) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const BODY_DUPLICATE_OFFSET: f32 = 12.0;
const AUDIO_DUPLICATE_OFFSET: f32 = 30.0;
const ATTACHMENT_H_OFFSET_FIELD: u32 = 3;
const ATTACHMENT_V_OFFSET_FIELD: u32 = 5;
const STORAGE_ATTACHMENTS_FIELD: u32 = 9;

/// The reference paths of a message type the host's graph clone remaps.
fn reference_paths(message_type: u32) -> Option<&'static [&'static [u32]]> {
    const IMAGE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    const MOVIE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    const STORAGE: &[&[u32]] = &[
        &[2],
        &[5, 1, 2],
        &[7, 1, 2],
        &[8, 1, 2],
        &[9, 1, 2],
        &[11, 1, 2],
        &[12, 1, 2],
        &[15, 1, 2],
        &[16, 1, 2],
        &[17, 1, 2],
        &[18, 1, 2],
        &[21, 1, 2],
        &[22, 1, 2],
        &[23, 1, 2],
        &[27, 1, 2],
        &[28, 1, 2],
        &[25, 1, 2],
        &[26, 1, 2],
    ];
    const ATTACHMENT: &[&[u32]] = &[&[1]];
    match message_type {
        IMAGE_MESSAGE_TYPE => Some(IMAGE),
        MOVIE_MESSAGE_TYPE => Some(MOVIE),
        2_001 | 2_022 => Some(STORAGE),
        graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE => Some(ATTACHMENT),
        _ => None,
    }
}

fn remap_paths(
    data: &[u8],
    paths: &[&[u32]],
    remap: &std::collections::HashMap<u64, u64>,
) -> Result<Vec<u8>, Error> {
    let mut output = data.to_vec();
    for path in paths {
        output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
            &output,
            path,
            |reference| {
                let view = litchi_iwa_common::wire::WireView::parse(
                    reference,
                    litchi_iwa_common::WireLimits::default(),
                )?;
                let current = view
                    .fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map(|(value, _)| value);
                let Some(replacement) = current.and_then(|current| remap.get(&current)) else {
                    return Ok(reference.to_vec());
                };
                litchi_iwa_common::wire::patch_varint_field(reference, 1, true, Some(*replacement))
            },
        )
        .map_err(graph::map_wire_error)?;
    }
    Ok(output)
}

/// A clone of one graph object under `remap` (the host's
/// `clone_pages_drawable_graph_object`).
fn clone_object(
    source: &Package,
    object: &litchi_iwa_core::ArchiveObject,
    remap: &std::collections::HashMap<u64, u64>,
) -> Result<litchi_iwa_core::ArchiveObject, Error> {
    let old = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
    let new = *remap.get(&old).ok_or(Error::InvalidSource)?;
    let mut messages = Vec::with_capacity(object.messages.len());
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        let data = match message.type_ {
            CAPTION_INFO_MESSAGE_TYPE => {
                let pairs = remap
                    .iter()
                    .map(|(old, new)| (*old, *new))
                    .collect::<Vec<_>>();
                litchi_iwa_protos::pages_movie_caption_codec::rewrite_caption_info(
                    &message.data,
                    litchi_iwa_protos::pages_movie_caption_codec::CaptionInfoWrite::new(&pairs),
                    litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::for_source(
                        &message.data,
                    )
                    .with_max_output_bytes(message.data.len().saturating_mul(10)),
                )
                .map_err(|_| Error::InvalidSource)?
            },
            STANDIN_MESSAGE_TYPE => message.data.clone(),
            message_type => match reference_paths(message_type) {
                Some(paths) => remap_paths(&message.data, paths, remap)?,
                None => {
                    if info
                        .object_references
                        .iter()
                        .any(|identifier| remap.contains_key(identifier))
                    {
                        return Err(Refusal::SharedGraph.into());
                    }
                    message.data.clone()
                },
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    let limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut cloned = litchi_iwa_core::ArchiveObject::new(new, messages, limits)
        .map_err(graph::map_core_error)?;
    cloned.archive_info.should_merge = object.archive_info.should_merge;
    for ((target, original), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&object.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_| Error::InvalidSource)?;
        *target = original.clone();
        target.length = length;
        for identifier in &mut target.object_references {
            if let Some(replacement) = remap.get(identifier) {
                *identifier = *replacement;
            }
        }
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

/// The host's `update_reference_list` with only an addition.
fn add_reference(references: &mut Vec<u64>, new: u64) {
    if !references.contains(&new) {
        references.push(new);
    }
}

fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut payload = vec![0x08];
    litchi_iwa_common::varint::encode_varint_into(&mut payload, identifier);
    payload
}

/// The body storage payload with a new attachment entry at `anchor`
/// (the host's `add_body_drawable_attachment`).
fn with_attachment_entry(payload: &[u8], anchor: u32, attachment: u64) -> Result<Vec<u8>, Error> {
    let mut entry = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut entry, 1, u64::from(anchor))
        .map_err(graph::map_wire_error)?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut entry,
        2,
        &reference_payload(attachment),
        litchi_iwa_common::WireLimits::default(),
    )
    .map_err(graph::map_wire_error)?;
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::default())
            .map_err(graph::map_wire_error)?;
    if graph::unique_field(&view, STORAGE_ATTACHMENTS_FIELD, 2)?.is_none() {
        let mut table = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut table,
            1,
            &entry,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(graph::map_wire_error)?;
        return litchi_iwa_common::wire::patch_length_delimited_field(
            payload,
            STORAGE_ATTACHMENTS_FIELD,
            false,
            Some(&table),
        )
        .map_err(|error| graph::map_wire_error(error).into());
    }
    litchi_iwa_common::wire::transform_length_delimited_field(
        payload,
        STORAGE_ATTACHMENTS_FIELD,
        |table| {
            let mut entries = Vec::new();
            for (order, data) in
                litchi_iwa_common::wire::repeated_length_delimited_payloads(table, 1)?
                    .into_iter()
                    .enumerate()
            {
                let view = litchi_iwa_common::wire::WireView::parse(
                    data,
                    litchi_iwa_common::WireLimits::default(),
                )?;
                let index = view
                    .fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map_or(0, |(value, _)| value);
                if index == u64::from(anchor) {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "attachment already at the anchor".to_owned(),
                    ));
                }
                entries.push((index, order, data.to_vec()));
            }
            let order = entries.len();
            entries.push((u64::from(anchor), order, entry.clone()));
            entries.sort_by_key(|(index, order, _)| (*index, *order));
            let entries = entries
                .into_iter()
                .map(|(_, _, data)| data)
                .collect::<Vec<_>>();
            litchi_iwa_common::wire::rewrite_repeated_length_delimited_fields(table, 1, &entries)
        },
    )
    .map_err(|error: litchi_iwa_common::Error| graph::map_wire_error(error).into())
}

fn duplicate(
    source: &Package,
    selector: DrawableSelector,
    anchor: TextPosition,
) -> Result<(Package, Diagnostics), Error> {
    let media = resolve(source, selector)?;
    let anchor = anchor.utf16_index();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let metadata = objects::metadata(source)?;
    let first = objects::next_identifier(source, metadata.as_ref())?;
    let mut remap = std::collections::HashMap::with_capacity(media.objects.len());
    for (offset, identifier) in media.objects.iter().enumerate() {
        let replacement = first
            .checked_add(offset as u64)
            .ok_or(Error::InvalidSource)?;
        remap.insert(*identifier, replacement);
    }
    let new_drawable = remap[&media.drawable.get()];
    let new_attachment = remap[&media.attachment.get()];
    let offset = if matches!(selector, DrawableSelector::Audio(_)) {
        AUDIO_DUPLICATE_OFFSET
    } else {
        BODY_DUPLICATE_OFFSET
    };
    let position = source
        .body_drawable_properties(selector)
        .map_err(|_| Error::InvalidSource)?
        .geometry()
        .and_then(|geometry| geometry.position)
        .ok_or(Error::InvalidSource)?;
    let moved = litchi_iwa_protos::shape_properties_codec::Point::new(
        position.x + offset,
        position.y + offset,
    );
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    {
        let archive = editable(source, &mut archives, &media.component)?;
        for identifier in &media.objects {
            let object = source
                .state
                .source
                .components()
                .get(media.component.as_ref())
                .and_then(|component| component.archive().object(*identifier))
                .ok_or(Error::InvalidSource)?;
            let mut cloned = clone_object(source, object, &remap)?;
            if *identifier == media.drawable.get() {
                let message_type = graph::message_type(selector);
                let index = cloned
                    .messages
                    .iter()
                    .position(|message| message.type_ == message_type)
                    .ok_or(Error::InvalidSource)?;
                let data = litchi_iwa_common::wire::transform_length_delimited_field(
                    &cloned.messages[index].data,
                    1,
                    |drawable| {
                        litchi_iwa_protos::shape_properties_codec::rewrite_drawable_archive(
                            drawable,
                            litchi_iwa_protos::shape_properties_codec::GeometryWrite::preserve()
                                .with_position(Some(moved)),
                            litchi_iwa_protos::shape_properties_codec::DrawablePropertiesWrite::preserve(),
                            litchi_iwa_protos::shape_properties_codec::DecodeOptions::for_source(
                                drawable,
                            ),
                        )
                        .map(litchi_iwa_protos::shape_properties_codec::RewriteOutput::into_bytes)
                        .map_err(|_| {
                            litchi_iwa_common::Error::InvalidFormat("drawable geometry".to_owned())
                        })
                    },
                )
                .map_err(graph::map_wire_error)?;
                cloned
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: message_type,
                            data,
                        },
                        archive_limits,
                    )
                    .map_err(graph::map_core_error)?;
            }
            if *identifier == media.attachment.get() {
                let (index, message) =
                    graph::unique_message(&cloned, graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE)?;
                let payload = message.data.clone();
                let view = litchi_iwa_common::wire::WireView::parse(
                    &payload,
                    litchi_iwa_common::WireLimits::default(),
                )
                .map_err(graph::map_wire_error)?;
                let offset_of = |number| -> Result<f32, Error> {
                    let field =
                        graph::unique_field(&view, number, 5)?.ok_or(Error::InvalidSource)?;
                    let bytes: [u8; 4] = field
                        .payload()
                        .try_into()
                        .map_err(|_| Error::InvalidSource)?;
                    Ok(f32::from_le_bytes(bytes) + offset)
                };
                let h = offset_of(ATTACHMENT_H_OFFSET_FIELD)?;
                let v = offset_of(ATTACHMENT_V_OFFSET_FIELD)?;
                let data = litchi_iwa_common::wire::patch_fixed32_field(
                    &payload,
                    ATTACHMENT_H_OFFSET_FIELD,
                    true,
                    Some(h.to_bits()),
                )
                .and_then(|data| {
                    litchi_iwa_common::wire::patch_fixed32_field(
                        &data,
                        ATTACHMENT_V_OFFSET_FIELD,
                        true,
                        Some(v.to_bits()),
                    )
                })
                .map_err(graph::map_wire_error)?;
                cloned
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                            data,
                        },
                        archive_limits,
                    )
                    .map_err(graph::map_core_error)?;
            }
            archive
                .insert_object(cloned, archive_limits)
                .map_err(graph::map_core_error)?;
        }
    }

    // The anchoring U+FFFC and the attachment entry.
    let (mut storage, _) = objects::body_storage(source)?;
    let removed = objects::splice_text(source, &mut storage, anchor, anchor, "\u{fffc}")?;
    if !removed.is_empty() {
        return Err(Error::InvalidSource);
    }
    let payload = with_attachment_entry(&storage.payload, anchor, new_attachment)?;
    {
        let existing = {
            let mut budget = Budget::new(source)?;
            let index = ObjectIndex::new(source, &mut budget)?;
            graph::resolve_body(&index, &mut budget)?
                .anchors
                .iter()
                .map(|anchor| anchor.attachment.get())
                .collect::<HashSet<_>>()
        };
        let archive = editable(source, &mut archives, &storage.component)?;
        let object = archive
            .object_mut(storage.identifier.get())
            .ok_or(Error::InvalidSource)?;
        object
            .replace_message_preserving_header(
                storage.message_index,
                RawMessage {
                    type_: storage.message_type,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(storage.message_index)
            .ok_or(Error::InvalidSource)?;
        add_reference(&mut info.object_references, new_attachment);
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| existing.contains(identifier))
            {
                add_reference(&mut field.object_references, new_attachment);
            }
        }
    }

    // The document z-order.
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let component = index.component_name(index.get(media.zorder)?.component_index)?;
        let archive = editable(source, &mut archives, &component)?;
        let object = archive
            .object_mut(media.zorder.get())
            .ok_or(Error::InvalidSource)?;
        let (message_index, message) = graph::unique_message(object, ZORDER_MESSAGE_TYPE)?;
        let previous = litchi_iwa_common::wire::repeated_length_delimited_payloads(
            &message.data,
            ZORDER_DRAWABLES_FIELD,
        )
        .map_err(graph::map_wire_error)?
        .into_iter()
        .filter_map(|reference| {
            litchi_iwa_common::wire::WireView::parse(
                reference,
                litchi_iwa_common::WireLimits::default(),
            )
            .ok()
            .and_then(|view| {
                view.fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map(|(value, _)| value)
            })
        })
        .collect::<HashSet<_>>();
        let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
            &message.data,
            ZORDER_DRAWABLES_FIELD,
            &reference_payload(new_drawable),
        )
        .map_err(graph::map_wire_error)?;
        object
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: ZORDER_MESSAGE_TYPE,
                    data,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or(Error::InvalidSource)?;
        add_reference(&mut info.object_references, new_drawable);
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| previous.contains(identifier))
            {
                add_reference(&mut field.object_references, new_drawable);
            }
        }
    }

    // Package metadata: watermark, UUIDs, shared data-reference owners.
    if let Some(metadata) = &metadata {
        let last = remap.values().copied().max().ok_or(Error::InvalidSource)?;
        let uuids = media
            .uuid_objects
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        let mut payload =
            metadata.registered(source, DOCUMENT_COMPONENT_IDENTIFIER, &uuids, last)?;
        for (data, object) in &media.data_references {
            payload = objects::adjust_data_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                *data,
                remap[object],
                true,
            )?;
        }
        let archive = editable(source, &mut archives, &metadata.component)?;
        archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let candidate = graph::publish_archives(source, archives)?;
    let created = resolve_drawable(&candidate, media.selector, NonZeroU64::new(new_drawable))?;
    if !created {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: media.objects.len(),
            ..Diagnostics::default()
        },
    ))
}

/// Whether the candidate body anchors `drawable` as media of the selected
/// kind.
fn resolve_drawable(
    candidate: &Package,
    selector: DrawableSelector,
    drawable: Option<NonZeroU64>,
) -> Result<bool, Error> {
    let drawable = drawable.ok_or(Error::InvalidSource)?;
    let mut budget = Budget::new(candidate)?;
    let index = ObjectIndex::new(candidate, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let _ = selector;
    Ok(body
        .anchors
        .iter()
        .any(|anchor| anchor.drawable == drawable))
}
