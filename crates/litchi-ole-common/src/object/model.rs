//! Host-neutral object and captured-stream views.

use super::directory::{self, Metadata};
use super::link::{self, Link};
use super::target::Target;
use crate::ole_streams::{self, native, presentation};
use crate::property_set::Guid;
use crate::{Error, LimitKind};
use litchi_cfb::OleError;
use std::collections::HashSet;
use std::sync::Arc;

/// Maximum number of OLEDS presentation streams allowed in one storage by
/// [MS-OLEDS] section 2.3.4.
const MAX_PRESENTATION_STREAMS: usize = 999;

/// Applies codec-owned limit validation even when the requested stream is absent.
pub(crate) fn validate_ole_stream_limits(limits: ole_streams::Limits) -> Result<(), Error> {
    limits.validate().map(|_| ())
}

/// Resource ceilings applied before CFB bytes are retained or rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of target-selected object storages.
    pub max_objects: usize,
    /// Maximum storage nesting depth captured from the CFB root. Also bounds
    /// the number of descendant storages retained for one selected object.
    pub max_storage_depth: usize,
    /// Maximum streams captured for one selected object.
    pub max_streams_per_object: usize,
    /// Maximum streams retained while capturing one complete editable package.
    pub max_streams: usize,
    /// Maximum size of one CFB stream.
    pub max_stream_size: u64,
    /// Maximum rendered size of one selected object compound file.
    pub max_object_size: u64,
    /// Maximum aggregate captured stream bytes for one operation.
    pub max_total_size: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_objects: 1_024,
            max_storage_depth: 32,
            max_streams_per_object: 4_096,
            max_streams: 65_536,
            max_stream_size: 128 * 1024 * 1024,
            max_object_size: 256 * 1024 * 1024,
            max_total_size: 512 * 1024 * 1024,
        }
    }
}

impl Limits {
    #[inline(always)]
    pub(crate) fn validate(self) -> Result<(), Error> {
        // Check fields directly: constructing a table here makes the success
        // path materialize all error metadata before it knows it is needed.
        macro_rules! nonzero {
            ($value:expr, $kind:ident, $maximum:expr) => {
                if $value == 0 {
                    return Err(Error::InvalidLimit {
                        kind: LimitKind::$kind,
                        value: 0,
                        maximum: $maximum,
                    });
                }
            };
        }
        nonzero!(self.max_objects, Objects, usize::MAX as u64);
        nonzero!(self.max_storage_depth, StorageDepth, usize::MAX as u64);
        nonzero!(
            self.max_streams_per_object,
            ObjectStreams,
            usize::MAX as u64
        );
        nonzero!(self.max_streams, PackageStreams, usize::MAX as u64);
        nonzero!(self.max_stream_size, StreamBytes, u64::MAX);
        nonzero!(self.max_object_size, ObjectBytes, u64::MAX);
        nonzero!(self.max_total_size, TotalBytes, u64::MAX);
        Ok(())
    }

    /// Maximum number of directory entries that one complete package
    /// capture may visit in any one sibling tree.  A package capture admits
    /// all selected objects together, so its storage ceiling is the product
    /// of the object and depth limits.
    pub(crate) const fn max_package_storages(self) -> usize {
        self.max_objects.saturating_mul(self.max_storage_depth)
    }

    /// Maximum direct children admitted while walking one complete package
    /// directory.  This is the package-level storage plus stream budget.
    pub(crate) const fn max_package_directory_entries(self) -> usize {
        self.max_package_storages().saturating_add(self.max_streams)
    }

    /// Maximum number of descendant storages retained for one selected or
    /// standalone object compound file.
    pub(crate) const fn max_object_storages(self) -> usize {
        self.max_storage_depth
    }

    /// Maximum streams admitted for one selected or standalone object. The
    /// per-object and complete-package ceilings both apply when a standalone
    /// object is captured as part of one edit operation.
    pub(crate) const fn max_object_streams(self) -> usize {
        if self.max_streams < self.max_streams_per_object {
            self.max_streams
        } else {
            self.max_streams_per_object
        }
    }

    /// Intersects this requested policy with a retained ingress policy.
    ///
    /// An edit may tighten a package's bounds, but it can never widen the
    /// limits under which the package was admitted.
    pub(crate) fn constrained_by(self, retained: Self) -> Self {
        Self {
            max_objects: self.max_objects.min(retained.max_objects),
            max_storage_depth: self.max_storage_depth.min(retained.max_storage_depth),
            max_streams_per_object: self
                .max_streams_per_object
                .min(retained.max_streams_per_object),
            max_streams: self.max_streams.min(retained.max_streams),
            max_stream_size: self.max_stream_size.min(retained.max_stream_size),
            max_object_size: self.max_object_size.min(retained.max_object_size),
            max_total_size: self.max_total_size.min(retained.max_total_size),
        }
    }
}

/// One captured CFB storage below a selected object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Storage {
    path: Vec<String>,
    directory: Metadata,
    clsid: Option<String>,
}

impl Storage {
    pub(crate) fn new(path: Vec<String>, directory: Metadata) -> Self {
        let clsid = directory.class_id().map(directory::format_class_id);
        Self {
            path,
            directory,
            clsid,
        }
    }

    /// Exact path relative to the selected object's storage.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Typed CFB directory metadata for this storage.
    #[must_use]
    pub fn directory(&self) -> &Metadata {
        &self.directory
    }

    /// CFB class identifier, when the directory entry contained one.
    #[must_use]
    pub fn class_id(&self) -> Option<Guid> {
        self.directory.class_id()
    }

    /// Canonical CFB CLSID text retained for existing host diagnostics.
    /// Prefer [`Self::class_id`] for typed comparisons.
    #[must_use]
    pub fn clsid(&self) -> Option<&str> {
        self.clsid.as_deref()
    }
}

/// One captured CFB stream with opaque, lossless bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Stream {
    path: Vec<String>,
    data: Arc<[u8]>,
    directory: Option<Metadata>,
}

impl Stream {
    pub(crate) fn new(path: Vec<String>, data: Arc<[u8]>, directory: Option<Metadata>) -> Self {
        Self {
            path,
            data,
            directory,
        }
    }

    /// Exact path relative to the selected object's storage.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Final CFB stream name, when this stream is directly below the object.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.path.last().map(String::as_str)
    }

    /// Borrowed opaque stream bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.data
    }

    /// Shared ownership of the captured stream allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.data)
    }

    /// Typed physical CFB directory metadata, when this stream came from a
    /// parsed directory entry.  Newly staged streams receive their metadata
    /// after the next atomic CFB publication.
    #[must_use]
    pub fn directory(&self) -> Option<&Metadata> {
        self.directory.as_ref()
    }

    /// Parses this stream as OLEDS link metadata when it is the standard
    /// `\x01Ole` stream.  The returned value shares the captured allocation.
    ///
    /// No link is resolved and no embedded payload is activated.
    ///
    /// # Errors
    ///
    /// Returns an error when a standard OLEDS stream is malformed or exceeds
    /// the OLEDS metadata limit.
    pub fn link(&self, limits: link::Limits) -> Result<Option<Link>, Error> {
        if self.name() != Some(link::NAME) {
            return Ok(None);
        }
        Link::parse_shared(Arc::clone(&self.data), limits).map(Some)
    }

    /// Parses this stream as an OLEDS presentation stream under explicit
    /// stream limits.
    ///
    /// The source allocation is shared with the captured object stream.
    pub fn presentation(
        &self,
        limits: ole_streams::Limits,
    ) -> Result<Option<presentation::Snapshot>, Error> {
        validate_ole_stream_limits(limits)?;
        let Some(name) = self.name() else {
            return Ok(None);
        };
        if !name.starts_with(ole_streams::PRESENTATION_STREAM_PREFIX) {
            return Ok(None);
        }
        ole_streams::parse_named_presentation(name, Arc::clone(&self.data), limits).map(Some)
    }

    /// Parses this stream as OLEDS native data under explicit stream limits.
    ///
    /// The source allocation is shared with the captured object stream.
    pub fn native(&self, limits: ole_streams::Limits) -> Result<Option<native::Snapshot>, Error> {
        validate_ole_stream_limits(limits)?;
        if self.name() != Some(ole_streams::NATIVE_STREAM_NAME) {
            return Ok(None);
        }
        ole_streams::parse_named_native(
            ole_streams::NATIVE_STREAM_NAME,
            Arc::clone(&self.data),
            limits,
        )
        .map(Some)
    }

    pub(crate) fn replace_data(&mut self, data: Arc<[u8]>) {
        self.data = data;
    }
}

/// A target-selected OLE storage and its opaque captured descendants.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Object {
    target: Target,
    storage: Storage,
    storages: Vec<Storage>,
    streams: Vec<Stream>,
    compound: Arc<[u8]>,
}

impl Object {
    pub(crate) fn new(
        target: Target,
        storage: Storage,
        storages: Vec<Storage>,
        streams: Vec<Stream>,
        compound: Arc<[u8]>,
    ) -> Self {
        Self {
            target,
            storage,
            storages,
            streams,
            compound,
        }
    }

    /// Host-owned semantic target key.
    #[must_use]
    pub fn key(&self) -> &str {
        self.target.key()
    }

    /// Exact absolute CFB path selected by the host target.
    #[must_use]
    pub fn path(&self) -> &[String] {
        self.target.path()
    }

    /// The target descriptor used to discover this object.
    #[must_use]
    pub fn target(&self) -> &Target {
        &self.target
    }

    /// Selected storage directory metadata.
    #[must_use]
    pub fn storage(&self) -> &Storage {
        &self.storage
    }

    /// Descendant storage directory metadata, in CFB discovery order.
    #[must_use]
    pub fn storages(&self) -> &[Storage] {
        &self.storages
    }

    /// All captured streams, including format-owned metadata streams.
    #[must_use]
    pub fn streams(&self) -> &[Stream] {
        &self.streams
    }

    /// Parses the direct-child OLEDS link stream, when present.
    ///
    /// This is a format-neutral view used by DOC, PPT, and XLS owners.  It
    /// never resolves external references or activates embedded content.
    ///
    /// # Errors
    ///
    /// Returns an error when the direct-child OLEDS stream is malformed or
    /// exceeds the OLEDS metadata limit.
    pub fn link(&self, limits: link::Limits) -> Result<Option<Link>, Error> {
        match self
            .streams
            .iter()
            .find(|stream| stream.path().len() == 1 && stream.name() == Some(link::NAME))
        {
            Some(stream) => stream.link(limits),
            None => Ok(None),
        }
    }

    /// Returns one direct-child OLEDS presentation under explicit limits.
    pub fn presentation(
        &self,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<presentation::Snapshot>, Error> {
        validate_ole_stream_limits(limits)?;
        self.validate_presentation_count()?;
        let name = ole_streams::presentation_name(index)?;
        self.streams
            .iter()
            .find(|stream| stream.path().len() == 1 && stream.name() == Some(name.as_str()))
            .map_or(Ok(None), |stream| stream.presentation(limits))
    }

    /// Returns all direct-child OLEDS presentations under explicit limits.
    pub fn presentations(
        &self,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, presentation::Snapshot)>, Error> {
        validate_ole_stream_limits(limits)?;
        self.validate_presentation_count()?;
        let mut presentations = Vec::new();
        for stream in self
            .streams
            .iter()
            .filter(|stream| stream.path().len() == 1)
            .filter(|stream| {
                stream
                    .name()
                    .is_some_and(|name| name.starts_with(ole_streams::PRESENTATION_STREAM_PREFIX))
            })
        {
            let name = stream.name().ok_or_else(|| {
                OleError::InvalidFormat("OLEDS presentation stream has no name".into())
            })?;
            let index = ole_streams::presentation_index(name)?;
            let snapshot = stream.presentation(limits)?.ok_or_else(|| {
                OleError::InvalidFormat("OLEDS presentation stream name disappeared".into())
            })?;
            if presentations.iter().any(|(existing, _)| *existing == index) {
                return Err(OleError::InvalidFormat(format!(
                    "duplicate OLEDS presentation stream index {index}"
                ))
                .into());
            }
            presentations.push((index, snapshot));
        }
        presentations.sort_unstable_by_key(|(index, _)| *index);
        Ok(presentations)
    }

    pub(crate) fn validate_presentation_count(&self) -> Result<(), Error> {
        let count = self
            .streams
            .iter()
            .filter(|stream| stream.path().len() == 1)
            .filter(|stream| {
                stream
                    .name()
                    .is_some_and(|name| name.starts_with(ole_streams::PRESENTATION_STREAM_PREFIX))
            })
            .count();
        if count > MAX_PRESENTATION_STREAMS {
            return Err(OleError::InvalidFormat(format!(
                "OLEDS storage contains more than {MAX_PRESENTATION_STREAMS} presentation streams"
            ))
            .into());
        }
        Ok(())
    }

    /// Returns the direct-child OLEDS native-data stream under explicit limits.
    pub fn native(&self, limits: ole_streams::Limits) -> Result<Option<native::Snapshot>, Error> {
        validate_ole_stream_limits(limits)?;
        self.validate_presentation_count()?;
        self.streams
            .iter()
            .find(|stream| {
                stream.path().len() == 1 && stream.name() == Some(ole_streams::NATIVE_STREAM_NAME)
            })
            .map_or(Ok(None), |stream| stream.native(limits))
    }

    /// Finds a captured stream by its path relative to the selected storage.
    #[must_use]
    pub fn stream(&self, path: &[&str]) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| {
                stream
                    .path
                    .iter()
                    .map(String::as_str)
                    .eq(path.iter().copied())
            })
            .map(Stream::bytes)
    }

    /// Standalone CFB bytes containing the selected storage's descendants.
    #[must_use]
    pub fn compound(&self) -> &[u8] {
        &self.compound
    }

    /// Shared ownership of the standalone CFB allocation.
    #[must_use]
    pub fn compound_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.compound)
    }
}

static EMPTY_OBJECTS: Objects = Objects {
    objects: Vec::new(),
};

impl super::snapshot::EmptyValue for Objects {
    fn empty() -> &'static Self {
        &EMPTY_OBJECTS
    }

    fn is_empty_value(&self) -> bool {
        self.objects.is_empty()
    }
}

/// Ordered, immutable object discovery results.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Objects {
    objects: Vec<Object>,
}

impl Objects {
    /// Moves these results into shared owner state; a package without
    /// discovered objects allocates nothing.
    pub(crate) fn into_shared(self) -> super::snapshot::Shared<Self> {
        super::snapshot::Shared::new(self)
    }

    pub(crate) fn new(objects: Vec<Object>) -> Result<Self, Error> {
        let mut keys = HashSet::new();
        let mut paths = HashSet::new();
        for object in &objects {
            if !keys.insert(object.key()) || !paths.insert(object.path()) {
                return Err(OleError::InvalidFormat(
                    "object discovery produced duplicate target keys or paths".into(),
                )
                .into());
            }
        }
        Ok(Self { objects })
    }

    /// Returns all discovered objects in target order.
    #[must_use]
    pub fn as_slice(&self) -> &[Object] {
        &self.objects
    }

    /// Number of discovered objects.
    #[must_use]
    pub fn len(&self) -> usize {
        self.objects.len()
    }

    /// Whether discovery returned no objects.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.objects.is_empty()
    }

    /// Finds an object by the host-owned semantic key.
    #[must_use]
    pub fn get(&self, key: &str) -> Option<&Object> {
        self.objects.iter().find(|object| object.key() == key)
    }

    /// Finds an object by checked discovery order.
    #[must_use]
    pub fn at(&self, index: usize) -> Option<&Object> {
        self.objects.get(index)
    }

    /// Borrows the ordered object iterator.
    pub fn iter(&self) -> std::slice::Iter<'_, Object> {
        self.objects.iter()
    }

    /// Consumes the catalog into its owned vector.
    #[must_use]
    pub fn into_vec(self) -> Vec<Object> {
        self.objects
    }
}

impl AsRef<[Object]> for Objects {
    fn as_ref(&self) -> &[Object] {
        self.as_slice()
    }
}

impl<'a> IntoIterator for &'a Objects {
    type Item = &'a Object;
    type IntoIter = std::slice::Iter<'a, Object>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}
