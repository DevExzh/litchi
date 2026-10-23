//! Immutable, shareable object-package state.
//!
//! A snapshot is the read side of the object owner.  It keeps the package
//! topology and stream allocations shared across clones while leaving all
//! mutation inside [`super::Edit`].  Format crates can therefore retain a
//! stable view between semantic reads without holding a mutable editor.

use super::codec::Package;
use super::editor::Edit;
use super::link::Link;
use super::model::Objects;
use super::target::Targets;
use crate::Error;
use crate::ole_streams::{self, native, presentation};
use crate::protection::admit_readonly_container;
use litchi_cfb::{OleError, OleFile, OleFileLimits, SectorLayoutPolicy};
use std::io::Cursor;
use std::sync::Arc;

/// A bounded, read-only CFB capture used when a format owner must inspect a
/// signed source without opening a mutation-capable object editor.
///
/// `Probe` deliberately exposes stream reads and retained source bytes only.
/// It has no edit, finish, replacement, or publication operation, so a
/// signature-bearing source cannot accidentally enter the ordinary object
/// mutation path. Encryption and DRM containers are rejected during the same
/// bounded directory admission used by the object owner.
#[derive(Debug, Clone)]
pub struct Probe {
    original: Arc<[u8]>,
    package: Package,
    limits: super::model::Limits,
    cfb_limits: OleFileLimits,
    signed: bool,
}

impl Probe {
    /// Opens and captures a CFB source under explicit object and CFB limits.
    ///
    /// Signature streams are retained byte-for-byte for readback. The probe
    /// refuses encrypted and DRM components before package capture.
    pub fn open_with_cfb_limits(
        bytes: Vec<u8>,
        limits: super::model::Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self, Error> {
        limits.validate()?;
        let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits)?;
        let signed = admit_readonly_container(
            &ole,
            "read-only signed probe",
            limits.max_package_directory_entries(),
        )?;
        let package = Package::capture(&mut ole, limits)?;
        package.check(limits)?;
        let original: Arc<[u8]> = Arc::from(ole.into_inner().into_inner());
        Ok(Self {
            original,
            package,
            limits,
            cfb_limits,
            signed,
        })
    }

    /// Returns whether the captured directory contains signature state.
    #[must_use]
    pub const fn is_signed(&self) -> bool {
        self.signed
    }

    /// Borrows one captured stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.package.stream(path)
    }

    /// Returns shared ownership of one captured stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.package.stream_shared(path)
    }

    /// Returns the exact source bytes retained by the probe.
    #[must_use]
    pub fn source_bytes(&self) -> Arc<[u8]> {
        Arc::clone(&self.original)
    }

    /// Returns the object-capture profile retained by the probe.
    #[must_use]
    pub const fn limits(&self) -> super::model::Limits {
        self.limits
    }

    /// Returns the low-level CFB profile retained by the probe.
    #[must_use]
    pub const fn cfb_limits(&self) -> OleFileLimits {
        self.cfb_limits
    }
}

#[derive(Debug, Clone)]
struct State {
    targets: Targets,
    limits: super::model::Limits,
    cfb_limits: OleFileLimits,
    original: Arc<Vec<u8>>,
    base_package: Package,
    package: Package,
    objects: Objects,
    changed: bool,
    layout: SectorLayoutPolicy,
    validated_layout: SectorLayoutPolicy,
}

/// Immutable, cheap-to-clone view of a captured OLE object package.
///
/// Cloning a snapshot shares the original package bytes and every captured
/// stream allocation.  Only small topology vectors are cloned, so callers can
/// pass read state between format-specific semantic layers without copying
/// payloads.  Use [`Snapshot::edit`] to start an independent transactional
/// edit.
#[derive(Debug, Clone)]
pub struct Snapshot {
    state: Arc<State>,
}

impl Snapshot {
    pub(crate) fn new(
        targets: Targets,
        limits: super::model::Limits,
        cfb_limits: OleFileLimits,
        original: Arc<Vec<u8>>,
        base_package: Package,
        package: Package,
        objects: Objects,
        changed: bool,
        layout: SectorLayoutPolicy,
        validated_layout: SectorLayoutPolicy,
    ) -> Self {
        Self {
            state: Arc::new(State {
                targets,
                limits,
                cfb_limits,
                original,
                base_package,
                package,
                objects,
                changed,
                layout,
                validated_layout,
            }),
        }
    }

    /// Opens a package as an immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns the same bounded CFB and target-validation errors as
    /// [`super::Edit::open`].
    pub fn open(
        bytes: Vec<u8>,
        targets: Targets,
        limits: super::model::Limits,
    ) -> Result<Self, Error> {
        Ok(Edit::open(bytes, targets, limits)?.staged())
    }

    /// The target catalog used by this snapshot.
    #[must_use]
    pub fn targets(&self) -> &Targets {
        &self.state.targets
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        &self.state.objects
    }

    /// The sector-layout policy captured from the editor that produced this
    /// snapshot.
    #[must_use]
    pub fn sector_layout_policy(&self) -> SectorLayoutPolicy {
        self.state.layout
    }

    /// Borrows an opaque package stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.state.package.stream(path)
    }

    /// Returns shared ownership of an opaque package stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.state.package.stream_shared(path)
    }

    /// Parses one selected object's OLEDS `\x01Ole` metadata, when present.
    ///
    /// The returned link is inert and retains unknown wire bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when `key` is absent or its OLEDS stream is malformed
    /// or exceeds the metadata limit.
    pub fn link(&self, key: &str) -> Result<Option<Link>, Error> {
        self.state
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .link()
    }

    /// Returns one selected object's presentation under explicit OLEDS limits.
    pub fn presentation(
        &self,
        key: &str,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<presentation::Snapshot>, Error> {
        self.state
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .presentation(index, limits)
    }

    /// Returns all selected-object OLEDS presentations under explicit limits.
    pub fn presentations(
        &self,
        key: &str,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, presentation::Snapshot)>, Error> {
        self.state
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .presentations(limits)
    }

    /// Returns a selected object's native-data stream under explicit limits.
    pub fn native(
        &self,
        key: &str,
        limits: ole_streams::Limits,
    ) -> Result<Option<native::Snapshot>, Error> {
        self.state
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .native(limits)
    }

    /// Starts an independent edit from this snapshot.
    ///
    /// The edit shares captured stream allocations until a write replaces
    /// one of them, so starting an edit does not copy large payloads.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit::from_snapshot(self)
    }

    /// Renders the snapshot, preserving the original bytes for a true no-op.
    ///
    /// # Errors
    ///
    /// Returns an error when the captured package cannot be rendered.
    pub fn finish(&self) -> Result<Vec<u8>, Error> {
        if self.state.changed {
            if self.state.layout == SectorLayoutPolicy::Reuse
                && let Some(rendered) = self.state.package.render_copy_through(
                    &self.state.base_package,
                    &self.state.original,
                    self.state.limits,
                    self.state.cfb_limits,
                )?
            {
                return Ok(rendered);
            }
            let rendered = self
                .state
                .package
                .render_with_layout(Some(self.state.original.as_slice()), self.state.layout)?;
            if self.state.layout != self.state.validated_layout {
                OleFile::open_with_limits(Cursor::new(rendered.as_slice()), self.state.cfb_limits)?;
            }
            Ok(rendered)
        } else {
            Ok(self.state.original.as_ref().clone())
        }
    }

    pub(crate) fn validated_layout(&self) -> SectorLayoutPolicy {
        self.state.validated_layout
    }

    pub(crate) fn cfb_limits(&self) -> OleFileLimits {
        self.state.cfb_limits
    }

    pub(crate) fn limits(&self) -> super::model::Limits {
        self.state.limits
    }

    pub(crate) fn original(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.state.original)
    }

    pub(crate) fn package(&self) -> Package {
        self.state.package.clone()
    }

    pub(crate) fn base_package(&self) -> Package {
        self.state.base_package.clone()
    }

    pub(crate) fn objects_clone(&self) -> Objects {
        self.state.objects.clone()
    }

    pub(crate) fn changed(&self) -> bool {
        self.state.changed
    }

    pub(crate) fn layout(&self) -> SectorLayoutPolicy {
        self.state.layout
    }
}
