//! Embedded media reachable from a Keynote document or one slide (record
//! 0372; a port of the host's `media_assets`, `slide_media_assets` and
//! `extract_media`), on the W1-E embedded-media inventory
//! (`litchi_iwa_common::media::embedded`).
//!
//! An asset is reachable when an object reachable from the roots (through
//! archive-header object references) names it in a header data reference
//! or owns it in the package registry.

use std::collections::{BTreeMap, BTreeSet, VecDeque};

use litchi_iwa_common::media::Type;
use litchi_iwa_common::media::embedded::{AssetId, Embedded};
use thiserror::Error as ThisError;

use super::structure::workspace::Workspace;
use super::{Package, ReadError};
use crate::SlideSelector;

const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const DOCUMENT_ROOT: u64 = 1;

/// A media-inventory read failure that never exposes native identifiers.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The slide selector did not resolve.
    #[error("the Keynote show has no such slide")]
    SlideNotFound,
    /// The asset is not reachable from the document.
    #[error("the media asset is not reachable from the Keynote document")]
    NotReachable,
    /// The asset has no materialized package member.
    #[error("the media asset is not materialized in the package")]
    NotMaterialized,
    /// The package cannot be read safely.
    #[error("the Keynote media inventory cannot be read")]
    InvalidSource,
}

/// The media family of an asset, from its file name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// An image (PNG, JPEG, TIFF, HEIF and related formats).
    Image,
    /// A video (MP4, MOV and related formats).
    Video,
    /// An audio clip (MP3, AAC, WAV and related formats).
    Audio,
    /// A PDF document.
    Pdf,
    /// An unrecognized file type.
    Unknown,
}

impl From<Type> for Kind {
    fn from(value: Type) -> Self {
        match value {
            Type::Image => Self::Image,
            Type::Video => Self::Video,
            Type::Audio => Self::Audio,
            Type::Pdf => Self::Pdf,
            Type::Unknown => Self::Unknown,
        }
    }
}

/// One embedded media asset reachable from the document. The value is an
/// opaque handle for [`Package::media_asset_data`]; it carries no native
/// identifier.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaAsset {
    data: AssetId,
    preferred_file_name: String,
    package_path: Option<String>,
    kind: Kind,
    size: Option<u64>,
    declared_size: Option<u64>,
    digest: Vec<u8>,
    component_reference_count: u64,
    message_reference_count: usize,
    has_data_metadata: bool,
}

impl MediaAsset {
    fn new(embedded: Embedded) -> Self {
        Self {
            data: embedded.data(),
            preferred_file_name: embedded.preferred_file_name().to_owned(),
            package_path: embedded.package_path().map(str::to_owned),
            kind: embedded.media_type().into(),
            size: embedded.size(),
            declared_size: embedded.declared_size(),
            digest: embedded.digest().to_vec(),
            component_reference_count: embedded.component_reference_count(),
            message_reference_count: embedded.message_reference_count(),
            has_data_metadata: embedded.has_data_metadata(),
        }
    }

    /// The private data identifier (never public; the replacement
    /// transaction keys its patch on it).
    pub(in crate::package) const fn data_identifier(&self) -> AssetId {
        self.data
    }

    /// The name Keynote prefers when exporting the asset.
    #[must_use]
    pub fn preferred_file_name(&self) -> &str {
        &self.preferred_file_name
    }

    /// The materialized `Data/` member, when present.
    #[must_use]
    pub fn package_path(&self) -> Option<&str> {
        self.package_path.as_deref()
    }

    /// The media family of the materialized (else the preferred) file name.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The materialized length in bytes.
    #[must_use]
    pub const fn size(&self) -> Option<u64> {
        self.size
    }

    /// The length the package metadata declares.
    #[must_use]
    pub const fn declared_size(&self) -> Option<u64> {
        self.declared_size
    }

    /// The declared digest (normally SHA-1).
    #[must_use]
    pub fn digest(&self) -> &[u8] {
        &self.digest
    }

    /// The sum of the component-level reference counts.
    #[must_use]
    pub const fn component_reference_count(&self) -> u64 {
        self.component_reference_count
    }

    /// The number of object-level references.
    #[must_use]
    pub const fn message_reference_count(&self) -> usize {
        self.message_reference_count
    }

    /// Whether the package's data-metadata map lists the asset.
    #[must_use]
    pub const fn has_data_metadata(&self) -> bool {
        self.has_data_metadata
    }

    /// Whether the asset has a materialized member.
    #[must_use]
    pub const fn is_materialized(&self) -> bool {
        self.package_path.is_some() && self.size.is_some()
    }
}

fn map_read_error(_error: ReadError) -> Error {
    Error::InvalidSource
}

fn embedded(package: &Package) -> Result<Vec<Embedded>, Error> {
    let metadata = package
        .structure_component(METADATA_MEMBER)
        .and_then(|archive| {
            archive.objects.iter().find_map(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
            })
        })
        .ok_or(Error::InvalidSource)?;
    let objects = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .collect::<Vec<_>>();
    litchi_iwa_style::media::embedded_assets(&metadata.data, objects.iter().copied(), |name| {
        package
            .structure_member_bytes(name)
            .map(|bytes| bytes.len() as u64)
    })
    .map_err(|_| Error::InvalidSource)
}

/// The assets reachable from `roots`, in data-identifier order.
fn reachable(package: &Package, roots: &[u64]) -> Result<Vec<MediaAsset>, Error> {
    let assets = embedded(package)?;
    let mut outgoing = BTreeMap::<u64, Vec<u64>>::new();
    let mut referencing = BTreeMap::<u64, BTreeSet<u64>>::new();
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(Error::InvalidSource);
            };
            let entry = outgoing.entry(identifier).or_default();
            for info in &object.archive_info.message_infos {
                entry.extend(info.object_references.iter().copied());
                for data in &info.data_references {
                    referencing.entry(*data).or_default().insert(identifier);
                }
            }
        }
    }
    let workspace = Workspace::new(package).map_err(|_| Error::InvalidSource)?;
    for asset in &assets {
        let data = asset.data().get();
        for component in workspace.registry.components() {
            for object in component.data_owned_objects(data) {
                referencing.entry(data).or_default().insert(object);
            }
        }
    }
    let mut seen = BTreeSet::new();
    let mut queue = roots.iter().copied().collect::<VecDeque<_>>();
    while let Some(identifier) = queue.pop_front() {
        if !seen.insert(identifier) {
            continue;
        }
        if let Some(references) = outgoing.get(&identifier) {
            queue.extend(references.iter().copied());
        }
    }
    Ok(assets
        .into_iter()
        .filter(|asset| {
            referencing
                .get(&asset.data().get())
                .is_some_and(|objects| objects.iter().any(|object| seen.contains(object)))
        })
        .map(MediaAsset::new)
        .collect())
}

impl Package {
    /// The embedded media reachable from the document (host
    /// `media_assets`), in data order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] for a malformed package.
    pub fn media_assets(&self) -> Result<Vec<MediaAsset>, Error> {
        reachable(self, &[DOCUMENT_ROOT])
    }

    /// The embedded media reachable from one slide and its node (host
    /// `slide_media_assets`).
    ///
    /// # Errors
    ///
    /// Returns [`Error::SlideNotFound`] and [`Error::InvalidSource`].
    pub fn slide_media_assets<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
    ) -> Result<Vec<MediaAsset>, Error> {
        let position = self
            .show()
            .map_err(map_read_error)?
            .select_slide(slide.into())
            .map_err(|_| Error::SlideNotFound)?
            .ok_or(Error::SlideNotFound)?
            .position();
        let record = self
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .ok_or(Error::SlideNotFound)?;
        reachable(self, &[record.node_identifier, record.slide_identifier])
    }

    /// The bytes of one reachable, materialized asset (host
    /// `extract_media`).
    ///
    /// # Errors
    ///
    /// Returns [`Error::NotReachable`] and [`Error::NotMaterialized`].
    pub fn media_asset_data(&self, asset: &MediaAsset) -> Result<&[u8], Error> {
        let found = self
            .media_assets()?
            .into_iter()
            .find(|candidate| candidate.data == asset.data)
            .ok_or(Error::NotReachable)?;
        let path = found.package_path().ok_or(Error::NotMaterialized)?;
        self.structure_member_bytes(path)
            .ok_or(Error::NotMaterialized)
    }
}
