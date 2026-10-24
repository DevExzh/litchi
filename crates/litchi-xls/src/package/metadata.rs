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
//! page. A signed workbook admits reads and exact no-op commits, while a
//! changed commit is refused before publication.
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
//! let reverted = edited.apply_metadata(&patch.inverse())?.into_package();
//! assert_eq!(reverted.to_bytes()?, package.to_bytes()?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::sync::Arc;

use litchi_cfb::OleError;
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::MetadataError;

use super::{Error as PackageError, Package};
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
    /// A shared property-set or object refusal with its typed classification.
    Object(litchi_ole_common::Error),
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
            Self::Object(error) => write!(formatter, "XLS property-set error: {error}"),
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
            Self::Object(error) => Some(error),
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

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        Self::Object(error)
    }
}

/// A metadata transaction result.
pub type Result<T> = std::result::Result<T, Error>;

pub use litchi_core::edit::Diagnostics;

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

/// A published metadata transaction with the shared commit triad.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A staged metadata transaction over one immutable [`Package`].
pub struct Edit {
    base: Package,
    transaction: property_set::Edit,
    initial: Option<SummaryProperties>,
    properties: Option<SummaryProperties>,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "changed",
                &self
                    .initial
                    .as_ref()
                    .zip(self.properties.as_ref())
                    .is_some_and(|(initial, properties)| initial != properties),
            )
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(base: &Package) -> Result<Self> {
        if base.is_encrypted() {
            return Err(Error::Protected);
        }
        let transaction = property_set::Edit::new(base).map_err(transaction_error)?;
        Ok(Self {
            base: base.clone(),
            transaction,
            initial: None,
            properties: None,
        })
    }

    fn ensure_properties(&mut self) -> Result<()> {
        if self.properties.is_some() {
            return Ok(());
        }
        let initial = {
            if self.base.is_signed() {
                let source = self.transaction.snapshot().map_err(transaction_error)?;
                SummaryProperties::from_sections(
                    source
                        .summary_information()
                        .map(|snapshot| snapshot.section()),
                    source
                        .document_summary_information()
                        .map(|snapshot| snapshot.section()),
                    source.user_defined_properties(),
                )?
            } else {
                let editor = self.transaction.editor().map_err(transaction_error)?;
                SummaryProperties::from_editor(editor)?
            }
        };
        self.initial = Some(initial.clone());
        self.properties = Some(initial);
        Ok(())
    }

    fn properties_ref(&mut self) -> Result<&SummaryProperties> {
        self.ensure_properties()?;
        self.properties
            .as_ref()
            .ok_or(Error::Package(PackageError::InvalidData(
                "metadata edit lost its staged properties",
            )))
    }

    fn properties_mut_ref(&mut self) -> Result<&mut SummaryProperties> {
        self.ensure_properties()?;
        self.properties
            .as_mut()
            .ok_or(Error::Package(PackageError::InvalidData(
                "metadata edit lost its staged properties",
            )))
    }

    /// The staged workbook properties.
    /// # Errors
    ///
    /// Returns an error when the property streams cannot be parsed under the
    /// package's retained CFB policy.
    pub fn properties(&mut self) -> Result<&SummaryProperties> {
        self.properties_ref()
    }

    /// Mutable access to the staged properties: set or remove one property
    /// by name ([`SummaryProperties::set`], [`SummaryProperties::remove`]),
    /// custom properties, and typed section edits.
    /// # Errors
    ///
    /// Returns an error when the property streams cannot be parsed under the
    /// package's retained CFB policy.
    pub fn properties_mut(&mut self) -> Result<&mut SummaryProperties> {
        self.properties_mut_ref()
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
        Ok(self.properties_mut_ref()?.merge_metadata(metadata)?)
    }

    /// Publish the staged properties as a new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Property`] when a changed source stream cannot be
    /// rewritten without loss, and [`Error::Package`] when the edited bytes
    /// cannot be republished.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_properties()?;
        let base = self.base;
        let mut transaction = self.transaction;
        let initial = self
            .initial
            .take()
            .ok_or(Error::Package(PackageError::InvalidData(
                "metadata edit lost its initial properties",
            )))?;
        let properties =
            self.properties
                .take()
                .ok_or(Error::Package(PackageError::InvalidData(
                    "metadata edit lost its staged properties",
                )))?;
        if properties != initial {
            // Preserve the established source-conflict precedence before the
            // signed-write policy rejects publication.
            transaction.check_source().map_err(transaction_error)?;
            if base.is_signed() {
                return Err(Error::Protected);
            }
            properties.stage_changes(
                &initial,
                transaction.editor_mut().map_err(transaction_error)?,
            )?;
        }
        let (patch, _) = transaction.finish()?;
        let before: Arc<[u8]> = Arc::from(patch.before());
        let after: Arc<[u8]> = Arc::from(patch.after());
        let package = if before == after {
            base
        } else {
            republish(&base, after.to_vec())?
        };
        let changed = before != after;
        Ok(Commit::new(
            package,
            Patch { before, after },
            Diagnostics::new(changed),
        ))
    }
}

impl Package {
    /// Start a metadata transaction over this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Protected`] for an encrypted workbook. Property
    /// stream errors are returned by the first accessor, verb, or commit.
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
        let changed = !patch.is_noop();
        if changed && self.is_signed() {
            return Err(Error::Protected);
        }
        let package = if changed {
            republish(self, patch.after_bytes().to_vec())?
        } else {
            self.clone()
        };
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(changed),
        ))
    }
}

fn protected(error: crate::Error) -> Error {
    match error {
        crate::Error::Cfb(OleError::InvalidFormat(_)) | crate::Error::Signed => Error::Protected,
        other => Error::from(other),
    }
}

fn transaction_error(error: crate::package::Error) -> Error {
    match error {
        crate::package::Error::Reader(error) => protected(error),
        other => Error::from(other),
    }
}

/// Admit edited bytes under the source package's retained policy.
fn republish(source: &Package, bytes: Vec<u8>) -> Result<Package> {
    crate::package::republish(source, bytes).map_err(Error::Package)
}

#[cfg(test)]
mod tests {
    use super::Edit;

    #[test]
    fn package_edit_entry_keeps_properties_cold() {
        let package = crate::Builder::new().build().unwrap();
        let edit = Edit::new(&package).unwrap();
        assert!(edit.initial.is_none());
        assert!(edit.properties.is_none());
    }
}
