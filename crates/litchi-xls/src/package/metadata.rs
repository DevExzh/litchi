//! Opened-package metadata transactions (ADR 0034 §7).
//!
//! [`Package::edit_metadata`] stages edits of the workbook properties held
//! in `\x05SummaryInformation`, `\x05DocumentSummaryInformation` and its
//! user-defined section, [`Edit::commit`] publishes a new [`Package`], and
//! [`Package::apply_metadata`] replays a [`Patch`] onto the exact source it
//! was made from. The source package is never mutated; dropping an [`Edit`]
//! is the rollback.
//!
//! The edit reads the three sections into the shared [`SummaryProperties`]
//! model and stages back only the sections that changed, through the XLS
//! property-set transaction, so an exact no-op commits the source package
//! unchanged and a changed commit keeps every other stream and the root
//! class identifier. A missing stream is created with the Windows-1252 code
//! page. Encrypted and signed workbooks are refused.
//!
//! ```
//! let package = litchi_xls::Package::new()?;
//! let mut edit = package.edit_metadata()?;
//! edit.merge_metadata(&litchi_core::Metadata {
//!     title: Some("Budget".to_owned()),
//!     ..litchi_core::Metadata::default()
//! })?;
//! let (edited, patch) = edit.commit()?.into_parts();
//! assert_eq!(edited.metadata()?.title.as_deref(), Some("Budget"));
//! let reverted = edited.apply_metadata(&patch.inverse())?;
//! assert_eq!(reverted.to_bytes()?, package.to_bytes()?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::sync::Arc;

use litchi_cfb::{OleError, SharedOleFile};
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::MetadataError;

use super::{Error as PackageError, Package, RetainedPolicy};
use crate::OpenOptions;
use crate::workbook::package::property_set;

/// A metadata transaction limit dimension.
///
/// The property sets bound their own text (1 MiB per value) and the
/// republished package is admitted under the source package's
/// [`crate::SourceLimits`] or source-backed limits, so no caller-selected dimension exists yet.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {}

/// A metadata transaction failure.
///
/// Messages name structural facts only (the unified field, the stream);
/// authored property values are never included.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The source package could not be read, or the edited bytes could not
    /// be republished as a XLS package.
    Package(PackageError),
    /// A unified metadata field cannot be stored, or its value was refused.
    Metadata(MetadataError),
    /// A property-set edit was refused (an invalid value or name, or a source
    /// stream that cannot be rewritten without loss).
    Property(OleError),
    /// The workbook is encrypted or signed; its properties cannot be
    /// rewritten without storing them in plaintext or breaking the signature.
    Protected,
    /// The patch was made from a different source package.
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => write!(formatter, "XLS package error: {error}"),
            Self::Metadata(error) => fmt::Display::fmt(error, formatter),
            Self::Property(_) => formatter.write_str("XLS workbook property edit was refused"),
            Self::Protected => formatter.write_str(
                "the metadata of an encrypted or signed XLS workbook cannot be rewritten",
            ),
            Self::PatchConflict => {
                formatter.write_str("the XLS metadata patch was made from a different package")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Metadata(error) => Some(error),
            Self::Property(error) => Some(error),
            Self::Protected | Self::PatchConflict => None,
        }
    }
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        match error {
            crate::Error::Cfb(error) => Self::Property(error),
            other => Self::Package(PackageError::Reader(other)),
        }
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<MetadataError> for Error {
    fn from(error: MetadataError) -> Self {
        Self::Metadata(error)
    }
}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Property(error)
    }
}

/// A metadata transaction result.
pub type Result<T> = std::result::Result<T, Error>;

/// What a metadata commit did.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// An exact-source, reversible metadata patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .finish()
    }
}

impl Patch {
    /// Whether applying the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// The patch that undoes this one, applicable to its result.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    /// The exact package bytes the patch applies to.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }

    /// The exact package bytes the patch produces.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }
}

/// A published metadata transaction.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl Commit {
    /// The published package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The reversible patch from the source package.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the commit did.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit into its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// A staged metadata transaction over one immutable [`Package`].
pub struct Edit {
    base: Package,
    transaction: property_set::Transaction,
    initial: SummaryProperties,
    properties: SummaryProperties,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("changed", &(self.initial != self.properties))
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(base: &Package) -> Result<Self> {
        if base.is_encrypted() {
            return Err(Error::Protected);
        }
        let bytes = base.to_bytes()?;
        let transaction = property_set::Snapshot::from_bytes(bytes)?
            .transaction()
            .map_err(protected)?;
        let initial = SummaryProperties::from_editor(transaction.editor())?;
        Ok(Self {
            base: base.clone(),
            transaction,
            properties: initial.clone(),
            initial,
        })
    }

    /// The staged workbook properties.
    #[must_use]
    pub const fn properties(&self) -> &SummaryProperties {
        &self.properties
    }

    /// Mutable access to the staged properties: set or remove one property
    /// by name ([`SummaryProperties::set`], [`SummaryProperties::remove`]),
    /// custom properties, and typed section edits.
    pub fn properties_mut(&mut self) -> &mut SummaryProperties {
        &mut self.properties
    }

    /// Set every field `metadata` carries, keeping the others.
    ///
    /// Returns whether a staged property changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Metadata`], leaving the staged properties unchanged,
    /// for a field the property sets cannot store or a refused value.
    pub fn merge_metadata(&mut self, metadata: &litchi_core::Metadata) -> Result<bool> {
        Ok(self.properties.merge_metadata(metadata)?)
    }

    /// Publish the staged properties as a new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Property`] when a changed source stream cannot be
    /// rewritten without loss, and [`Error::Package`] when the edited bytes
    /// cannot be republished.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            base,
            mut transaction,
            initial,
            properties,
        } = self;
        if properties == initial {
            return unchanged(base);
        }
        properties.stage_changes(&initial, transaction.editor_mut())?;
        let patch = transaction.commit()?.into_patch();
        let before: Arc<[u8]> = Arc::from(patch.before());
        let after: Arc<[u8]> = Arc::from(patch.after());
        let package = if before == after {
            base
        } else {
            republish(&base, after.to_vec())?
        };
        Ok(Commit {
            package,
            diagnostics: Diagnostics {
                changed: before != after,
            },
            patch: Patch { before, after },
        })
    }
}

impl Package {
    /// Start a metadata transaction over this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Protected`] for an encrypted or signed workbook, and
    /// an error when the property streams cannot be read or are invalid.
    pub fn edit_metadata(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply a metadata patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, and [`Error::Package`] when its result cannot be republished.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        let bytes = self.to_bytes()?;
        if bytes.as_slice() != patch.before_bytes() {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(self.clone());
        }
        republish(self, patch.after_bytes().to_vec())
    }
}

fn protected(error: crate::Error) -> Error {
    match error {
        crate::Error::Cfb(OleError::InvalidFormat(_)) => Error::Protected,
        other => Error::from(other),
    }
}

fn unchanged(base: Package) -> Result<Commit> {
    let bytes: Arc<[u8]> = Arc::from(base.to_bytes()?);
    Ok(Commit {
        package: base,
        diagnostics: Diagnostics { changed: false },
        patch: Patch {
            before: Arc::clone(&bytes),
            after: bytes,
        },
    })
}

/// Admit edited bytes under the source package's retained policy.
fn republish(source: &Package, bytes: Vec<u8>) -> Result<Package> {
    match source.inner.policy {
        RetainedPolicy::Eager {
            source_limits,
            leniency,
            compatibility_profile,
        } => Ok(Package::from_owned_bytes(
            bytes,
            OpenOptions::default()
                .source_limits(source_limits)
                .leniency(leniency)
                .compatibility_profile(compatibility_profile),
        )?),
        RetainedPolicy::Source { limits, leniency } => {
            let cfb_limits =
                source
                    .source()
                    .map(SharedOleFile::limits)
                    .ok_or(PackageError::InvalidData(
                        "source-backed XLS package lost its CFB",
                    ))?;
            let cfb = Arc::new(
                SharedOleFile::open_with_limits(
                    Arc::new(litchi_core::OwnedSource::new(bytes)),
                    cfb_limits,
                )
                .map_err(PackageError::Cfb)?,
            );
            let workbook = crate::SourceBackedWorkbook::from_shared_ole_file_with_options(
                cfb,
                crate::SourceBackedOptions::default()
                    .with_limits(limits)
                    .with_leniency(leniency),
            )
            .map_err(PackageError::Source)?;
            Ok(Package::from_source(workbook)?)
        },
    }
}
