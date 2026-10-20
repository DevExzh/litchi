//! Opened-package metadata transactions (ADR 0034 §7).
//!
//! [`Package::edit_metadata`] stages edits of the presentation properties
//! held in `\x05SummaryInformation`, `\x05DocumentSummaryInformation` and its
//! user-defined section, [`Edit::commit`] publishes a new [`Package`], and
//! [`Package::apply_metadata`] replays a [`Patch`] onto the exact source it
//! was made from. The source package is never mutated; dropping an [`Edit`]
//! is the rollback.
//!
//! The edit reads the three sections into the shared [`SummaryProperties`]
//! model and stages back only the sections that changed, through the PPT
//! property-set transaction ([`super::property_set`]), so an exact no-op
//! commits the source package unchanged and a changed commit keeps every other
//! stream and the root class identifier. A missing stream, or the blank
//! section a presentation built without properties carries, is replaced by
//! one with the Windows-1252 code page. Encrypted presentations are refused.
//!
//! ```
//! let package = litchi_ppt::Package::new()?;
//! let mut edit = package.edit_metadata()?;
//! edit.merge_metadata(&litchi_core::Metadata {
//!     title: Some("Roadmap".to_owned()),
//!     ..litchi_core::Metadata::default()
//! })?;
//! let (edited, patch) = edit.commit()?.into_parts();
//! assert_eq!(edited.metadata()?.title.as_deref(), Some("Roadmap"));
//! let reverted = edited.apply_metadata(&patch.inverse())?.into_package();
//! assert_eq!(reverted.to_bytes()?, package.to_bytes()?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::sync::Arc;

use litchi_cfb::OleError;
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::MetadataError;

use super::property_set;
use super::{Error as PackageError, Package};
use crate::OpenOptions;

/// A metadata transaction limit dimension.
///
/// The property sets bound their own text (1 MiB per value) and the
/// republished package is admitted under the source package's
/// [`super::SourceLimits`], so no caller-selected dimension exists yet.
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
    /// be republished as a PPT package.
    Package(PackageError),
    /// A unified metadata field cannot be stored, or its value was refused.
    Metadata(MetadataError),
    /// A property-set edit was refused (an invalid value or name, or a source
    /// stream that cannot be rewritten without loss).
    Property(OleError),
    /// The document is encrypted; its properties cannot be rewritten without
    /// storing them in plaintext.
    Encrypted,
    /// The patch was made from a different source package.
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => write!(formatter, "PPT package error: {error}"),
            Self::Metadata(error) => fmt::Display::fmt(error, formatter),
            Self::Property(_) => formatter.write_str("PPT presentation property edit was refused"),
            Self::Encrypted => formatter
                .write_str("the metadata of an encrypted PPT presentation cannot be rewritten"),
            Self::PatchConflict => {
                formatter.write_str("the PPT metadata patch was made from a different package")
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
            Self::Encrypted | Self::PatchConflict => None,
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
            return Err(Error::Encrypted);
        }
        let bytes = base.to_bytes()?;
        let transaction = property_set::Snapshot::from_bytes(bytes)?.transaction()?;
        let initial = SummaryProperties::from_editor(transaction.editor())?;
        Ok(Self {
            base: base.clone(),
            transaction,
            properties: initial.clone(),
            initial,
        })
    }

    /// The staged presentation properties.
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
        let (_, patch) = transaction.commit()?.into_parts();
        let before: Arc<[u8]> = Arc::from(patch.source());
        let after: Arc<[u8]> = Arc::from(patch.replacement());
        let package = if before == after {
            base
        } else {
            reopen(&base, &after)?
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
    /// Returns [`Error::Encrypted`] for an encrypted presentation, and an
    /// error when the property streams cannot be read or are invalid.
    pub fn edit_metadata(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply a metadata patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, and [`Error::Package`] when its result cannot be republished.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        let bytes = self.to_bytes()?;
        if bytes.as_slice() != patch.before_bytes() {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            reopen(self, patch.after_bytes())?
        };
        Ok(Commit {
            package,
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
            },
            patch: patch.clone(),
        })
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

fn reopen(base: &Package, bytes: &[u8]) -> Result<Package> {
    Ok(Package::from_bytes_with_options(
        bytes,
        OpenOptions::default()
            .source_limits(base.source_limits())
            .leniency(base.leniency()),
    )?)
}
