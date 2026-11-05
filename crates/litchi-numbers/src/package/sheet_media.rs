//! Implementation of the public [`crate::sheet::drawable::media`] family:
//! the embedded media assets reachable from the document or one sheet, their
//! bytes, and replacing the data of an asset or of a sheet image, movie
//! (data or poster) or audio clip (record 0378).
//!
//! An asset is reachable when an object reachable from the root through
//! object references carries a data reference to it, directly or through a
//! component data-reference owner (the host editor's
//! `reachable_embedded_assets`). A replacement keeps the asset's identifier
//! (every drawable sharing it shows the new bytes): the `Data/` member is
//! rewritten and its `DataInfo` digest and materialized length are patched in
//! place, as the host's media editor does.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::{HashMap, HashSet, VecDeque};
use std::fmt;
use std::sync::Arc;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::media::Type;
use litchi_iwa_common::media::embedded::{AssetId, Embedded};
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, append_varint_field,
    transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::RawMessage;
use thiserror::Error as ThisError;

use super::Package;
use super::metadata::{ENTRY_NAME, MESSAGE_TYPE};
use super::sheet_drawable_graph as graph;
use super::structure::{
    self, Fault, FaultResult, Published, Working,
    metadata::{self, Transition},
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{DrawableSelector, Kind};

structure::structural_family!("media", Wire(litchi_iwa_common::LimitKind));

/// `TSP.PackageMetadata.datas`.
const METADATA_DATAS_FIELD: u32 = 4;
/// `TSP.DataInfo.identifier`, `.digest` and `.materialized_length`.
const DATA_INFO_IDENTIFIER_FIELD: u32 = 1;
const DATA_INFO_DIGEST_FIELD: u32 = 2;
const DATA_INFO_LENGTH_FIELD: u32 = 18;
/// `TSD.ImageArchive.data`.
const IMAGE_DATA_FIELD: u32 = 11;
/// `TSD.MovieArchive.movieData` and `.posterImageData`.
const MOVIE_DATA_FIELD: u32 = 14;
const MOVIE_POSTER_FIELD: u32 = 15;
/// The document root object.
const DOCUMENT_ROOT: u64 = 1;

/// A typed, content-redacted media failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted media source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} media items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("media candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("media patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable has no media data of the requested role (only images,
    /// movies and audio clips carry data; only movies carry a poster).
    #[error("a Numbers sheet {kind} has no such media data")]
    NoMedia {
        /// The selected kind.
        kind: &'static str,
    },
    /// The asset is not reachable from the document.
    #[error("the media asset is not reachable from the document")]
    AssetNotFound,
    /// The asset has no `Data/` member.
    #[error("the media asset is not materialized")]
    NotMaterialized,
    /// The replacement is empty.
    #[error("the replacement media is empty")]
    EmptyData,
    /// The replacement's signature names another media type.
    #[error("the replacement's media type differs from the asset's")]
    MismatchedType,
    /// Commit was requested before staging a change.
    #[error("no media change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the media graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support media edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } | Self::AssetNotFound => {
                RefusalKind::NotFound
            },
            Self::NoMedia { .. } | Self::NotMaterialized => RefusalKind::UnsupportedTarget,
            Self::EmptyData | Self::MismatchedType => RefusalKind::InvalidValue,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// Which media data of a drawable a replacement targets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Role {
    Data,
    Poster,
}

#[derive(Debug, Clone)]
enum Target {
    Asset(AssetId),
    Prune,
    Drawable(usize, DrawableSelector, Role),
}

/// Media replacements staged against one package.
pub struct Edit<'a> {
    source: &'a Package,
    changes: Vec<(Target, Arc<[u8]>)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(&mut self, target: Target, data: &[u8]) -> Result<&mut Self, Error> {
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push((target, Arc::from(data)));
        Ok(self)
    }

    /// Stage new bytes for embedded asset `asset`, which keeps its
    /// identifier.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyData`], [`Error::Allocation`].
    pub fn replace_asset(&mut self, asset: AssetId, data: &[u8]) -> Result<&mut Self, Error> {
        self.stage(Target::Asset(asset), data)
    }

    /// Stage the removal of every data record nothing references: no
    /// component data-reference owner, object data reference or
    /// `DataMetadataMap` entry names it. Each record's `DataInfo` leaves the
    /// package metadata and its `Data/` member, when it has one, is
    /// deleted (the host's `media::Draft::remove_unreferenced`, over the
    /// shared `litchi_iwa_style::media::unreferenced_data` rule and the
    /// shared data-owner transition's data removals, record 0417). A
    /// package without such records is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn prune_unreferenced(&mut self) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push((Target::Prune, Arc::from(&[][..])));
        Ok(self)
    }

    /// Stage new bytes for the media data of one sheet image, movie or
    /// audio clip.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::EmptyData`], [`Error::Allocation`].
    pub fn replace_data<'sheet>(
        &mut self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        data: &[u8],
    ) -> Result<&mut Self, Error> {
        let sheet = resolve_sheet(self.source, sheet.into())?;
        self.stage(Target::Drawable(sheet, drawable.into(), Role::Data), data)
    }

    /// Stage new bytes for the poster image of one sheet movie.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::EmptyData`], [`Error::Allocation`].
    pub fn replace_poster<'sheet>(
        &mut self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        data: &[u8],
    ) -> Result<&mut Self, Error> {
        let sheet = resolve_sheet(self.source, sheet.into())?;
        self.stage(Target::Drawable(sheet, drawable.into(), Role::Poster), data)
    }

    /// Validate and atomically publish the staged replacements, in order;
    /// a replacement with the current bytes is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::NoMedia`],
    /// [`Refusal::AssetNotFound`], [`Refusal::NotMaterialized`],
    /// [`Refusal::MismatchedType`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (target, data) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let asset = match target {
                Target::Prune => {
                    if let Some(published) = prune(base)? {
                        touched = touched.saturating_add(published.touched_components);
                        current = Some(published.package);
                    }
                    continue;
                },
                Target::Asset(asset) => *asset,
                Target::Drawable(sheet, selector, role) => {
                    drawable_asset(base, *sheet, *selector, *role)?
                },
            };
            if let Some(published) = replace(base, asset, data)? {
                if extract(&published.package, asset)? != **data {
                    return Err(Error::Verification);
                }
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The asset of a sheet drawable's media data `role`.
fn drawable_asset(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
    role: Role,
) -> Result<AssetId, Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    let refused = Refusal::NoMedia {
        kind: entry.kind.name(),
    };
    let field = match (entry.kind, role) {
        (Kind::Image, Role::Data) => IMAGE_DATA_FIELD,
        (Kind::Movie | Kind::Audio, Role::Data) => MOVIE_DATA_FIELD,
        (Kind::Movie, Role::Poster) => MOVIE_POSTER_FIELD,
        _ => return Err(refused.into()),
    };
    let (index, _) = entry.message.ok_or(Error::InvalidSource)?;
    let object = graph::object(package, entry.identifier)?;
    let records = structure::references::records(&object.messages[index].data, &[field])?;
    let [record] = records.as_slice() else {
        return Err(refused.into());
    };
    let view = WireView::parse(record, WireLimits::DEFAULT)?;
    let identifier =
        graph::unique_field(&view, DATA_INFO_IDENTIFIER_FIELD, 0)?.ok_or(Error::InvalidSource)?;
    let (value, _) = litchi_iwa_common::decode_varint_from_bytes(identifier.payload())
        .map_err(|_error| Error::InvalidSource)?;
    AssetId::new(value).map_err(|_error| Error::InvalidSource)
}

/// Remove every unreferenced data record through the shared
/// data-owner transition (see [`Edit::prune_unreferenced`]).
fn prune(package: &Package) -> Result<Option<Published>, Error> {
    let mut working = Working::new(package)?;
    working.keep_previews();
    let Ok((_, _, payload)) = metadata_payload(&working) else {
        return Ok(None);
    };
    let all: Vec<_> = package.state.components.iter_objects().collect();
    let Some(pruning) =
        litchi_iwa_style::media::unreferenced_data(&payload, all.iter().copied(), |path| {
            member_size(package, path)
        })
        .map_err(|_error| Error::InvalidSource)?
    else {
        return Ok(None);
    };
    let mut transition = Transition::default();
    for (data, member) in pruning.records {
        if let Some(member) = member {
            structure::reserve(&mut transition.data_removals, 1)?;
            transition.data_removals.push((data.get(), member));
        } else {
            structure::reserve(&mut transition.data_record_removals, 1)?;
            transition.data_record_removals.push(data.get());
        }
    }
    Ok(Some(working.publish(&transition)?))
}

/// The metadata payload's (object, message index, payload).
fn metadata_payload(working: &Working<'_>) -> FaultResult<(u64, usize, Vec<u8>)> {
    let archive = working.archive(ENTRY_NAME)?;
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == MESSAGE_TYPE
                && found
                    .replace((
                        object.archive_info.identifier.ok_or(Fault::Invalid)?,
                        index,
                        message.data.clone(),
                    ))
                    .is_some()
            {
                return Err(Fault::Invalid);
            }
        }
    }
    found.ok_or(Fault::Unsupported)
}

fn member_size(package: &Package, path: &str) -> Option<u64> {
    let catalog = structure::exact_source(package).ok()?;
    catalog
        .package()
        .iter()
        .find(|entry| entry.name() == path)
        .map(|entry| entry.data().len() as u64)
}

/// The assets reachable from `root` (the host's
/// `reachable_embedded_assets`), ordered by identifier.
fn assets(package: &Package, root: u64) -> Result<Vec<Embedded>, Error> {
    let working = Working::new(package)?;
    let Ok((_, _, metadata_payload)) = metadata_payload(&working) else {
        return Ok(Vec::new());
    };
    let all: Vec<_> = package.state.components.iter_objects().collect();
    let objects = || all.iter().copied();
    let embedded = litchi_iwa_style::media::embedded_assets(&metadata_payload, objects(), |path| {
        member_size(package, path)
    })
    .map_err(|_error| Error::InvalidSource)?;
    let mut outgoing = HashMap::<u64, Vec<u64>>::new();
    let mut data = HashMap::<u64, Vec<u64>>::new();
    for object in objects() {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            outgoing
                .entry(identifier)
                .or_default()
                .extend_from_slice(&info.object_references);
            data.entry(identifier)
                .or_default()
                .extend_from_slice(&info.data_references);
        }
    }
    let mut reachable = HashSet::new();
    let mut queue = VecDeque::from([root]);
    while let Some(identifier) = queue.pop_front() {
        if !reachable.insert(identifier) {
            continue;
        }
        if let Some(references) = outgoing.get(&identifier) {
            queue.extend(references.iter().copied());
        }
    }
    let mut live = HashSet::new();
    for identifier in &reachable {
        if let Some(references) = data.get(identifier) {
            live.extend(references.iter().copied());
        }
    }
    if let Some(registry) = metadata::registry(&working)? {
        for owner in &registry.data_references {
            if reachable.contains(&owner.object) {
                live.insert(owner.data);
            }
        }
    }
    let mut found: Vec<Embedded> = embedded
        .into_iter()
        .filter(|asset| live.contains(&asset.data().get()))
        .collect();
    found.sort_by_key(|asset| asset.data().get());
    Ok(found)
}

fn find(package: &Package, asset: AssetId) -> Result<Embedded, Error> {
    assets(package, DOCUMENT_ROOT)?
        .into_iter()
        .find(|embedded| embedded.data() == asset)
        .ok_or_else(|| Refusal::AssetNotFound.into())
}

fn extract(package: &Package, asset: AssetId) -> Result<Vec<u8>, Error> {
    let embedded = find(package, asset)?;
    let path = embedded.package_path().ok_or(Refusal::NotMaterialized)?;
    let catalog = structure::exact_source(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == path)
        .ok_or(Refusal::NotMaterialized)?;
    structure::copy_bytes(entry.data()).map_err(Into::into)
}

/// The `DataInfo` with a new digest and materialized length: existing
/// fields rewritten in place, missing ones appended (the host's
/// `patch_data_info`).
fn patched_data_info(info: &[u8], digest: &[u8], length: u64) -> FaultResult<Vec<u8>> {
    let limits = WireLimits::DEFAULT;
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    let view = WireView::parse(info, limits).map_err(wire)?;
    let mut output = Vec::new();
    structure::reserve(&mut output, info.len().saturating_add(digest.len()))?;
    let (mut digests, mut lengths) = (0_usize, 0_usize);
    for field in view.fields() {
        match field.number() {
            DATA_INFO_DIGEST_FIELD => {
                digests += 1;
                append_length_delimited_field(&mut output, DATA_INFO_DIGEST_FIELD, digest, limits)
                    .map_err(wire)?;
            },
            DATA_INFO_LENGTH_FIELD => {
                lengths += 1;
                append_varint_field(&mut output, DATA_INFO_LENGTH_FIELD, length).map_err(wire)?;
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if digests > 1 || lengths > 1 {
        return Err(Fault::Invalid);
    }
    if digests == 0 {
        append_length_delimited_field(&mut output, DATA_INFO_DIGEST_FIELD, digest, limits)
            .map_err(wire)?;
    }
    if lengths == 0 {
        append_varint_field(&mut output, DATA_INFO_LENGTH_FIELD, length).map_err(wire)?;
    }
    Ok(output)
}

/// Replace the bytes of `asset` (the host's media `replace`).
fn replace(package: &Package, asset: AssetId, data: &[u8]) -> Result<Option<Published>, Error> {
    use sha1::Digest as _;
    let embedded = find(package, asset)?;
    let path = embedded
        .package_path()
        .ok_or(Refusal::NotMaterialized)?
        .to_owned();
    let detected = Type::from_bytes(data);
    if embedded.media_type() != Type::Unknown
        && detected != Type::Unknown
        && embedded.media_type() != detected
    {
        return Err(Refusal::MismatchedType.into());
    }
    if extract(package, asset)? == data {
        return Ok(None);
    }
    let digest = sha1::Sha1::digest(data);
    let length = u64::try_from(data.len()).map_err(|_error| Error::InvalidSource)?;
    let mut working = Working::new(package)?;
    working.keep_previews();
    let (object, index, payload) = metadata_payload(&working)?;
    let mut patched = 0_usize;
    let payload =
        transform_length_delimited_fields_at_path(&payload, &[METADATA_DATAS_FIELD], |info| {
            let view = WireView::parse(info, WireLimits::DEFAULT)?;
            let identifier = view
                .fields()
                .find(|field| field.number() == DATA_INFO_IDENTIFIER_FIELD)
                .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
                .map(|(value, _)| value);
            if identifier != Some(asset.get()) {
                return Ok(info.to_vec());
            }
            patched += 1;
            patched_data_info(info, &digest, length)
                .map_err(|_error| litchi_iwa_common::Error::InvalidFormat("DataInfo".to_owned()))
        })?;
    if patched != 1 {
        return Err(Error::InvalidSource);
    }
    let limits = working.archive_limits();
    working
        .archive_mut(ENTRY_NAME)?
        .object_mut(object)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    let mut transition = Transition::default();
    structure::reserve(&mut transition.entry_writes, 1)?;
    transition
        .entry_writes
        .push((path, structure::copy_bytes(data)?));
    Ok(Some(working.publish(&transition)?))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The embedded media assets reachable from the document's object
    /// graph, ordered by identifier.
    ///
    /// # Errors
    ///
    /// Source errors.
    pub fn media_assets(&self) -> Result<Vec<Embedded>, Error> {
        assets(self, DOCUMENT_ROOT)
    }

    /// The embedded media assets reachable from one sheet's object graph.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], source errors.
    pub fn sheet_media_assets<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Vec<Embedded>, Error> {
        let position = resolve_sheet(self, sheet.into())?;
        let sheet = graph::sheet_target(self, position)?;
        assets(self, sheet.identifier)
    }

    /// The bytes of one embedded asset reachable from the document.
    ///
    /// # Errors
    ///
    /// [`Refusal::AssetNotFound`], [`Refusal::NotMaterialized`].
    pub fn extract_media(&self, asset: AssetId) -> Result<Vec<u8>, Error> {
        extract(self, asset)
    }

    /// Start an immutable media replacement edit.
    #[must_use]
    pub const fn edit_media(&self) -> Edit<'_> {
        Edit {
            source: self,
            changes: Vec::new(),
        }
    }

    /// Apply a media patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_media(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
