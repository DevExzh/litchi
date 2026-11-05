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
use crate::protection::{Limits as ProtectionLimits, admit_readonly_container};
use litchi_cfb::{OleError, OleFile, OleFileLimits, SectorLayoutPolicy};
use std::fmt;
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
    original: Arc<Vec<u8>>,
    package: CapturedPackage,
    limits: super::model::Limits,
    cfb_limits: OleFileLimits,
    signed: bool,
}

impl Probe {
    /// Opens and captures a CFB source under explicit object and CFB limits.
    ///
    /// Signature streams are retained byte-for-byte for readback. The probe
    /// refuses encrypted and DRM components before package capture.
    /// The probe captures the complete CFB read-only. Host edit targets in
    /// `options` do not select streams; object and CFB limits apply in full.
    pub fn open_with_options(bytes: Vec<u8>, options: OpenOptions) -> Result<Self, Error> {
        let limits = options.limits_value();
        let cfb_limits = options.cfb_limits_value();
        limits.validate()?;
        let mut ole = OleFile::open(Cursor::new(bytes), cfb_limits)?;
        let signed = admit_readonly_container(
            &ole,
            "read-only signed probe",
            ProtectionLimits {
                max_entries: limits.max_package_directory_entries(),
            },
        )?;
        let package = CapturedPackage::capture(&mut ole, limits)?;
        package.check(limits)?;
        let original: Arc<Vec<u8>> = Arc::from(ole.into_inner().into_inner());
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
    pub fn source_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.original)
    }

    /// Returns the object-capture profile retained by the probe.
    #[must_use]
    pub const fn limits_value(&self) -> super::model::Limits {
        self.limits
    }

    /// Returns the low-level CFB profile retained by the probe.
    #[must_use]
    pub const fn cfb_limits_value(&self) -> OleFileLimits {
        self.cfb_limits
    }
}

/// A value with a static empty instance.
pub(crate) trait EmptyValue: Sized + 'static {
    /// The process-wide empty value.
    fn empty() -> &'static Self;

    /// Whether this value equals [`Self::empty`].
    fn is_empty_value(&self) -> bool;
}

/// One shared owner field whose empty value needs no allocation.
///
/// Most OLE packages have no host targets and no discovered objects; holding
/// those as `None` keeps an open or publication from allocating an `Arc` for
/// an empty collection, while a populated value is shared exactly as an
/// `Arc` field would be.  The field stays one pointer wide.
pub(crate) struct Shared<T: EmptyValue>(Option<Arc<T>>);

impl<T: EmptyValue> Shared<T> {
    pub(crate) fn new(value: T) -> Self {
        if value.is_empty_value() {
            Self(None)
        } else {
            Self(Some(Arc::new(value)))
        }
    }
}

impl<T: EmptyValue + Clone> Shared<T> {
    /// Copy-on-write access, as [`Arc::make_mut`]; an empty value gains its
    /// own allocation only here, when it is about to change.
    pub(crate) fn make_mut(&mut self) -> &mut T {
        Arc::make_mut(self.0.get_or_insert_with(|| Arc::new(T::empty().clone())))
    }
}

impl<T: EmptyValue> Clone for Shared<T> {
    fn clone(&self) -> Self {
        Self(self.0.clone())
    }
}

impl<T: EmptyValue> std::ops::Deref for Shared<T> {
    type Target = T;

    fn deref(&self) -> &T {
        match &self.0 {
            Some(value) => value,
            None => T::empty(),
        }
    }
}

/// Shared immutable owner state.
///
/// `Package` and an edit that starts from it hold one `Arc` to this state.
/// An edit obtains a private state only when a mutation needs copy-on-write;
/// cloning a read owner therefore has the same one-counter cost as the
/// former `Snapshot` owner while staged publication still shares all capture
/// allocations.
#[derive(Clone)]
pub(crate) struct PackageState {
    pub(crate) targets: Shared<Targets>,
    pub(crate) limits: super::model::Limits,
    pub(crate) cfb_limits: OleFileLimits,
    pub(crate) original: Arc<Vec<u8>>,
    pub(crate) package: Arc<CapturedPackage>,
    pub(crate) objects: Shared<Objects>,
    pub(crate) changed: bool,
    pub(crate) layout: SectorLayoutPolicy,
    pub(crate) validated_layout: SectorLayoutPolicy,
}

/// Immutable, cheap-to-clone owner of a captured OLE object package.
///
/// The owner keeps one [`Arc`] to its immutable state. Cloning a package is
/// therefore O(1) and shares the published source bytes, captured topology,
/// stream allocations, limits, and layout policy through that state.  A
/// default edit retains the same state until its first mutation; an edit that
/// needs copy-on-write detaches a private state while continuing to share the
/// immutable byte and catalog allocations.  A changed publication creates one
/// fresh state only when the edit state is still shared; otherwise it
/// normalizes that state in place.  A no-op publication reuses the existing
/// state Arc.
#[derive(Clone)]
pub struct Package {
    pub(crate) state: Arc<PackageState>,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("target_count", &self.state.targets.len())
            .field("object_count", &self.state.objects.len())
            .field("changed", &self.state.changed)
            .finish()
    }
}

impl Package {
    pub(crate) fn new(
        targets: Shared<Targets>,
        limits: super::model::Limits,
        cfb_limits: OleFileLimits,
        original: Arc<Vec<u8>>,
        package: Arc<CapturedPackage>,
        objects: Shared<Objects>,
        changed: bool,
        layout: SectorLayoutPolicy,
        validated_layout: SectorLayoutPolicy,
    ) -> Self {
        Self {
            state: Arc::new(PackageState {
                targets,
                limits,
                cfb_limits,
                original,
                package,
                objects,
                changed,
                layout,
                validated_layout,
            }),
        }
    }

    pub(crate) fn from_state(state: Arc<PackageState>) -> Self {
        Self { state }
    }

    pub(crate) fn state_shared(&self) -> Arc<PackageState> {
        Arc::clone(&self.state)
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
        &self.state.targets
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        &self.state.objects
    }

    /// The sector-layout policy captured from the edit that produced this
    /// package.
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

    /// Returns shared ownership of the exact bytes published by this package.
    ///
    /// Cloning this handle is O(1); the immutable current content remains
    /// available to deferred owners and exact no-op publications without a
    /// second copy.
    #[must_use]
    pub fn source_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.state.original)
    }

    /// Object-capture bounds retained by this immutable package.
    #[must_use]
    pub fn limits_value(&self) -> super::model::Limits {
        self.state.limits
    }

    /// Physical CFB bounds retained by this immutable package.
    #[must_use]
    pub fn cfb_limits_value(&self) -> OleFileLimits {
        self.state.cfb_limits
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
        self.state
            .objects
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
        self.state
            .objects
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
        self.state
            .objects
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
        self.state
            .objects
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
            self.state.limits
        };
        let limits = requested_limits.constrained_by(self.state.limits);
        limits.validate()?;
        let cfb_limits = if options.has_cfb_limits_override() {
            options.cfb_limits_value()
        } else {
            self.state.cfb_limits
        };
        let cfb_limits = OleFileLimits::new(
            cfb_limits
                .max_input_bytes()
                .min(self.state.cfb_limits.max_input_bytes()),
        )?
        .with_max_directory_bytes(
            cfb_limits
                .max_directory_bytes()
                .min(self.state.cfb_limits.max_directory_bytes()),
        )?
        .with_max_allocation_table_bytes(
            cfb_limits
                .max_allocation_table_bytes()
                .min(self.state.cfb_limits.max_allocation_table_bytes()),
        )?;
        let layout = if options.has_sector_layout_override() {
            options.sector_layout_value()
        } else {
            self.state.layout
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
            .write_all(self.state.original.as_slice())
            .map_err(|source| Error::Cfb(OleError::Io(source)))
    }

    pub(crate) fn cfb_limits(&self) -> OleFileLimits {
        self.state.cfb_limits
    }

    pub(crate) fn limits(&self) -> super::model::Limits {
        self.state.limits
    }

    pub(crate) fn current_bytes(&self) -> &[u8] {
        self.state.original.as_slice()
    }
}
