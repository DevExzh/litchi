//! Opened-package metadata transactions (ADR 0034 §7).
//!
//! [`Package::edit_metadata`] stages edits of the document properties held
//! in `\x05SummaryInformation`, `\x05DocumentSummaryInformation` and its
//! user-defined section, [`Edit::commit`] publishes a new [`Package`], and
//! [`Package::apply_metadata`] replays a [`Patch`] onto the exact source it
//! was made from. The source package is never mutated; dropping an [`Edit`]
//! discards it.
//!
//! The edit defers reading the three sections until the first property
//! operation or commit, then stages back only the sections that changed,
//! through the DOC property-set transaction
//! ([`super::property_set`]), so an exact no-op commits the source package
//! unchanged and a changed commit keeps every other stream and the root
//! class identifier. A missing stream is created with the Windows-1252 code
//! page. Encrypted documents are refused: their properties would have to be
//! written as plaintext next to the protected content.
//!
//! ```
//! let mut builder = litchi_doc::Builder::new();
//! builder.add_paragraph("Body")?;
//! let package = builder.build()?;
//!
//! let mut edit = package.edit_metadata()?;
//! edit.merge_metadata(&litchi_core::Metadata {
//!     title: Some("Quarterly report".to_owned()),
//!     ..litchi_core::Metadata::default()
//! })?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let (edited, patch) = commit.into_parts();
//! assert_eq!(edited.metadata().title.as_deref(), Some("Quarterly report"));
//!
//! // The patch replays onto its source and reverts through its inverse.
//! let replayed = package.apply_metadata(&patch)?.into_package();
//! assert_eq!(replayed.to_bytes()?, edited.to_bytes()?);
//! let reverted = edited.apply_metadata(&patch.inverse())?.into_package();
//! assert_eq!(reverted.to_bytes()?, package.to_bytes()?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::MetadataError;

use super::property_set;
use super::{Error as PackageError, OpenOptions, Package};
use crate::parts::fib::FileInformationBlock;

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
    /// be republished as a DOC package.
    Package(PackageError),
    /// A unified metadata field cannot be stored, or its value was refused.
    Metadata(MetadataError),
    /// A property-set edit was refused (an invalid value or name, or a source
    /// stream that cannot be rewritten without loss).
    Property(OleError),
    /// A shared property-set or object refusal with its typed classification.
    Object(litchi_ole_common::Error),
    /// The document is encrypted; its properties cannot be rewritten without
    /// storing them in plaintext.
    Encrypted,
    /// The patch was made from a different source package.
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => write!(formatter, "DOC package error: {error}"),
            Self::Metadata(error) => fmt::Display::fmt(error, formatter),
            Self::Property(_) => formatter.write_str("DOC document property edit was refused"),
            Self::Object(error) => write!(formatter, "DOC property-set error: {error}"),
            Self::Encrypted => {
                formatter.write_str("the metadata of an encrypted DOC document cannot be rewritten")
            },
            Self::PatchConflict => {
                formatter.write_str("the DOC metadata patch was made from a different package")
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

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        Self::Object(error)
    }
}

/// A metadata transaction result.
pub type Result<T> = std::result::Result<T, Error>;

/// Content-free diagnostics of one metadata commit.
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

/// A published metadata edit: the new package, its exact reversible patch
/// and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A staged metadata transaction over one immutable [`Package`].
pub struct Edit {
    base: Package,
    options: super::EditOptions,
    before: Option<Arc<[u8]>>,
    transaction: Option<property_set::Staging>,
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
    fn new(base: &Package, options: super::EditOptions) -> Result<Self> {
        Ok(Self {
            base: base.clone(),
            options,
            before: None,
            transaction: None,
            properties: None,
            initial: None,
        })
    }

    fn ensure_loaded(&mut self) -> Result<()> {
        if self.transaction.is_some() {
            return Ok(());
        }
        let bytes = self.base.to_bytes()?;
        if is_encrypted(&bytes)? {
            return Err(Error::Encrypted);
        }
        let source = property_set::Snapshot::from_bytes(bytes)?;
        let before = source.bytes_shared();
        let transaction =
            property_set::Staging::new(source, self.options.protection_policy_value())?;
        let initial = SummaryProperties::from_editor(transaction.editor())?;
        self.before = Some(before);
        self.properties = Some(initial.clone());
        self.initial = Some(initial);
        self.transaction = Some(transaction);
        Ok(())
    }

    /// The staged document properties.
    pub fn properties(&mut self) -> Result<&SummaryProperties> {
        self.ensure_loaded()?;
        self.properties.as_ref().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "metadata properties were not initialized".into(),
            ))
        })
    }

    /// Mutable access to the staged properties: set or remove one property
    /// by name ([`SummaryProperties::set`], [`SummaryProperties::remove`]),
    /// custom properties, and typed section edits.
    pub fn properties_mut(&mut self) -> Result<&mut SummaryProperties> {
        self.ensure_loaded()?;
        self.properties.as_mut().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "metadata properties were not initialized".into(),
            ))
        })
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
        Ok(self.properties_mut()?.merge_metadata(metadata)?)
    }

    /// Publish the staged properties as a new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Property`] when a changed source stream cannot be
    /// rewritten without loss, and [`Error::Package`] when the edited bytes
    /// cannot be republished (for example under the protection policy).
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_loaded()?;
        let Self {
            base,
            before: Some(before),
            options,
            transaction: Some(mut transaction),
            initial: Some(initial),
            properties: Some(properties),
        } = self
        else {
            return Err(Error::Package(PackageError::InvalidFormat(
                "metadata edit was not initialized".into(),
            )));
        };
        if properties == initial {
            return Ok(unchanged(base, before));
        }
        properties.stage_changes(&initial, transaction.editor_mut())?;
        let after: Arc<[u8]> = Arc::from(transaction.finish()?);
        if before == after {
            return Ok(unchanged(base, before));
        }
        let package = reopen(&base, &after)?.with_edit_options(&options);
        Ok(Commit::new(
            package,
            Patch { before, after },
            Diagnostics::CHANGED,
        ))
    }
}

impl Package {
    /// Start a metadata transaction over this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Encrypted`] for an encrypted document, and an error
    /// when the property streams cannot be read or are invalid.
    pub fn edit_metadata(&self) -> Result<Edit> {
        Edit::new(self, self.retained_edit_options())
    }

    /// Start a metadata edit of this package under explicit edit options
    /// (their protection policy authorizes a changed write).
    ///
    /// # Errors
    ///
    /// As [`Package::edit_metadata`].
    pub fn edit_metadata_with_options(&self, options: super::EditOptions) -> Result<Edit> {
        Edit::new(self, self.constrained_edit_options(options))
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
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        let protection = property_set::classify_host(&bytes)?;
        self.retained_edit_options()
            .protection_policy_value()
            .authorize(protection)?;
        let package =
            reopen(self, patch.after_bytes())?.with_edit_options(&self.retained_edit_options());
        Ok(Commit::new(package, patch.clone(), Diagnostics::CHANGED))
    }
}

fn unchanged(base: Package, bytes: Arc<[u8]>) -> Commit {
    Commit::new(
        base,
        Patch {
            before: Arc::clone(&bytes),
            after: bytes,
        },
        Diagnostics::UNCHANGED,
    )
}

fn reopen(base: &Package, bytes: &[u8]) -> Result<Package> {
    Ok(Package::from_bytes_with_options(
        bytes,
        OpenOptions::default()
            .source_limits(base.source_limits())
            .leniency(base.leniency()),
    )?)
}

/// Whether the FIB declares the document encrypted (XOR or RC4).
fn is_encrypted(bytes: &[u8]) -> Result<bool> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let word = ole.open_stream(&["WordDocument"])?;
    let fib = FileInformationBlock::parse(&word)?;
    Ok(fib.is_encrypted())
}
