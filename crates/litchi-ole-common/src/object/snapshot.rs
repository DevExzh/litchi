//! Immutable, shareable object-package state.
//!
//! A package is the read side of the object owner. It keeps the package
//! topology and stream allocations shared across clones while leaving all
//! mutation inside [`super::Edit`].  Format crates can therefore retain a
//! stable view between semantic reads without holding a mutable editor.

use super::codec::Package as CapturedPackage;
use super::editor::Edit;
use super::link::Link;
use super::model::Objects;
use super::options::{EditOptions, OpenOptions};
use super::patch::{Commit, Patch};
use super::target::Targets;
use crate::Error;
use crate::ole_streams::{self, native, presentation};
use litchi_cfb::{OleError, OleFile, OleFileLimits, SectorLayoutPolicy};
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

/// Immutable, cheap-to-clone owner of a captured OLE object package.
///
/// Cloning a package shares the original/current bytes, captured topology and
/// every stream allocation.  The small owner handle stays inline, so opening
/// a package does not allocate a second owner-state object when the caller
/// immediately starts an edit.  Cloning the handle remains O(1): it only
/// clones the already shared catalogs and byte owners.
/// An edit clones those `Arc` handles and copies no package topology until a
/// mutation first needs copy-on-write.
#[derive(Clone)]
pub struct Package {
    targets: Arc<Targets>,
    limits: super::model::Limits,
    cfb_limits: OleFileLimits,
    original: Arc<Vec<u8>>,
    current: Arc<Vec<u8>>,
    base_package: Arc<CapturedPackage>,
    package: Arc<CapturedPackage>,
    objects: Arc<Objects>,
    changed: bool,
    layout: SectorLayoutPolicy,
    validated_layout: SectorLayoutPolicy,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("target_count", &self.targets.len())
            .field("object_count", &self.objects.len())
            .field("changed", &self.changed)
            .field(
                "source_is_current",
                &Arc::ptr_eq(&self.original, &self.current),
            )
            .finish()
    }
}

impl Package {
    pub(crate) fn new(
        targets: Arc<Targets>,
        limits: super::model::Limits,
        cfb_limits: OleFileLimits,
        original: Arc<Vec<u8>>,
        current: Arc<Vec<u8>>,
        base_package: Arc<CapturedPackage>,
        package: Arc<CapturedPackage>,
        objects: Arc<Objects>,
        changed: bool,
        layout: SectorLayoutPolicy,
        validated_layout: SectorLayoutPolicy,
    ) -> Self {
        Self {
            targets,
            limits,
            cfb_limits,
            original,
            current,
            base_package,
            package,
            objects,
            changed,
            layout,
            validated_layout,
        }
    }

    /// Opens an immutable package under the default target and limit policy.
    ///
    /// Use [`Self::open_with_options`] when the caller needs to select a
    /// target set or impose explicit admission limits.
    pub fn open(bytes: Vec<u8>) -> Result<Self, Error> {
        Self::open_with_options(bytes, OpenOptions::default())
    }

    /// Opens an immutable package under the caller's target and limit policy.
    ///
    /// # Errors
    ///
    /// Returns bounded CFB and target-validation errors from the requested
    /// [`OpenOptions`].
    pub fn open_with_options(bytes: Vec<u8>, options: OpenOptions) -> Result<Self, Error> {
        Ok(Edit::open_with_cfb_limits(
            bytes,
            options.targets_value().clone(),
            options.limits_value(),
            options.cfb_limits_value(),
        )?
        .staged())
    }

    /// Opens an immutable package and returns the admitted physical CFB
    /// context used by the owner.
    pub fn open_with_ole_file(
        bytes: Vec<u8>,
        options: OpenOptions,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), Error> {
        Self::open_with_ole_file_with_admission(bytes, options, |_| Ok(()))
    }

    /// Opens an immutable package with one owner-supplied admission callback.
    /// The callback runs after bounded CFB parsing and before selected stream
    /// capture, so the owner can enforce its own policy without reparsing.
    pub fn open_with_ole_file_with_admission<E, F>(
        bytes: Vec<u8>,
        options: OpenOptions,
        admission: F,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), E>
    where
        E: From<Error>,
        F: FnOnce(&OleFile<Cursor<Vec<u8>>>) -> Result<(), E>,
    {
        Edit::open_with_ole_file_with_admission_and_cfb_limits(
            bytes,
            options.targets_value().clone(),
            options.limits_value(),
            options.cfb_limits_value(),
            admission,
        )
        .map(|(edit, ole)| (edit.staged(), ole))
    }

    /// Feature-gated profiled equivalent of [`Self::open_with_ole_file`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn open_with_ole_file_profiled(
        bytes: Vec<u8>,
        options: OpenOptions,
        observer: impl FnMut(super::CfbParseEvent),
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), Error> {
        Self::open_with_ole_file_profiled_with_admission(bytes, options, observer, |_| Ok(()))
    }

    /// Feature-gated profiled equivalent of
    /// [`Self::open_with_ole_file_with_admission`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn open_with_ole_file_profiled_with_admission<E, F>(
        bytes: Vec<u8>,
        options: OpenOptions,
        observer: impl FnMut(super::CfbParseEvent),
        admission: F,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), E>
    where
        E: From<Error>,
        F: FnOnce(&OleFile<Cursor<Vec<u8>>>) -> Result<(), E>,
    {
        Edit::open_with_ole_file_profiled_with_admission_and_cfb_limits(
            bytes,
            options.targets_value().clone(),
            options.limits_value(),
            options.cfb_limits_value(),
            observer,
            admission,
        )
        .map(|(edit, ole)| (edit.staged(), ole))
    }

    /// The target catalog used by this package.
    #[must_use]
    pub fn targets(&self) -> &Targets {
        self.targets.as_ref()
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        self.objects.as_ref()
    }

    /// The sector-layout policy captured from the edit that produced this
    /// package.
    #[must_use]
    pub fn sector_layout_policy(&self) -> SectorLayoutPolicy {
        self.layout
    }

    /// Borrows an opaque package stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.package.stream(path)
    }

    /// Returns shared ownership of an opaque package stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.package.stream_shared(path)
    }

    /// Returns shared ownership of the exact bytes published by this package.
    ///
    /// Cloning this handle is O(1); the immutable current content remains
    /// available to deferred owners and exact no-op publications without a
    /// second copy.
    #[must_use]
    pub fn source_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.original)
    }

    /// Object-capture bounds retained by this immutable package.
    #[must_use]
    pub fn limits_value(&self) -> super::model::Limits {
        self.limits
    }

    /// Physical CFB bounds retained by this immutable package.
    #[must_use]
    pub fn cfb_limits_value(&self) -> OleFileLimits {
        self.cfb_limits
    }

    /// Parses one selected object's OLEDS `\x01Ole` metadata, when present.
    ///
    /// The returned link is inert and retains unknown wire bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when `key` is absent or its OLEDS stream is malformed
    /// or exceeds the metadata limit.
    pub fn link(&self, key: &str, limits: super::link::Limits) -> Result<Option<Link>, Error> {
        self.objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .link(limits)
    }

    /// Returns one selected object's presentation under explicit OLEDS limits.
    pub fn presentation(
        &self,
        key: &str,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<presentation::Snapshot>, Error> {
        self.objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .presentation(index, limits)
    }

    /// Returns all selected-object OLEDS presentations under explicit limits.
    pub fn presentations(
        &self,
        key: &str,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, presentation::Snapshot)>, Error> {
        self.objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .presentations(limits)
    }

    /// Returns a selected object's native-data stream under explicit limits.
    pub fn native(
        &self,
        key: &str,
        limits: ole_streams::Limits,
    ) -> Result<Option<native::Snapshot>, Error> {
        self.objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .native(limits)
    }

    /// Starts an independent edit with the retained package policy.
    ///
    /// The edit shares captured stream allocations until a write replaces
    /// one of them, so starting an edit does not copy large payloads.
    pub fn edit(&self) -> Result<Edit, Error> {
        self.edit_with_options(EditOptions::default())
    }

    /// Starts an independent edit under a policy constrained by this package.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit, Error> {
        let requested_limits = if options.has_limits_override() {
            options.limits_value()
        } else {
            self.limits
        };
        let limits = requested_limits.constrained_by(self.limits);
        limits.validate()?;
        let cfb_limits = if options.has_cfb_limits_override() {
            options.cfb_limits_value()
        } else {
            self.cfb_limits
        };
        let cfb_limits = OleFileLimits::new(
            cfb_limits
                .max_input_bytes()
                .min(self.cfb_limits.max_input_bytes()),
        )?
        .with_max_directory_bytes(
            cfb_limits
                .max_directory_bytes()
                .min(self.cfb_limits.max_directory_bytes()),
        )?
        .with_max_allocation_table_bytes(
            cfb_limits
                .max_allocation_table_bytes()
                .min(self.cfb_limits.max_allocation_table_bytes()),
        )?;
        let layout = if options.has_sector_layout_override() {
            options.sector_layout_value()
        } else {
            self.layout
        };
        Ok(Edit::from_package_with_options(
            self, limits, cfb_limits, layout,
        ))
    }

    /// Replays a patch against a package with matching recorded content.
    pub fn apply(&self, patch: &Patch) -> Result<Commit, Error> {
        patch.apply(self)
    }

    /// Streams the package bytes to a sequential sink.
    pub fn write_to(&self, mut writer: impl std::io::Write) -> Result<(), Error> {
        writer
            .write_all(self.current.as_slice())
            .map_err(|source| Error::Cfb(OleError::Io(source)))
    }

    pub(crate) fn validated_layout(&self) -> SectorLayoutPolicy {
        self.validated_layout
    }

    pub(crate) fn cfb_limits(&self) -> OleFileLimits {
        self.cfb_limits
    }

    pub(crate) fn targets_shared(&self) -> Arc<Targets> {
        Arc::clone(&self.targets)
    }

    pub(crate) fn limits(&self) -> super::model::Limits {
        self.limits
    }

    pub(crate) fn original(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.original)
    }

    pub(crate) fn current_bytes(&self) -> &[u8] {
        self.current.as_slice()
    }

    pub(crate) fn package(&self) -> Arc<CapturedPackage> {
        Arc::clone(&self.package)
    }

    pub(crate) fn base_package(&self) -> Arc<CapturedPackage> {
        Arc::clone(&self.base_package)
    }

    pub(crate) fn objects_clone(&self) -> Arc<Objects> {
        Arc::clone(&self.objects)
    }

    pub(crate) fn changed(&self) -> bool {
        self.changed
    }

    pub(crate) fn layout(&self) -> SectorLayoutPolicy {
        self.layout
    }
}
