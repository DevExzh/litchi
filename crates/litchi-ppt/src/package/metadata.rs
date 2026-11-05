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
//! property-set transaction, so an exact no-op
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
use litchi_core::SourceVersion;
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::MetadataError;

use super::property_set;
use super::{Error as PackageError, Package};
use crate::OpenOptions;
use crate::scoped::Lazy;

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
            Self::Package(error) => write!(formatter, "PPT package error: {error}"),
            Self::Metadata(error) => fmt::Display::fmt(error, formatter),
            Self::Property(_) => formatter.write_str("PPT presentation property edit was refused"),
            Self::Object(error) => write!(formatter, "PPT property-set error: {error}"),
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

/// A published metadata transaction: the new package, exact patch and
/// content-free diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A staged metadata transaction over one immutable [`Package`].
pub struct Edit {
    base: Package,
    expected_source: SourceVersion,
    state: Lazy<State>,
}

struct State {
    transaction: property_set::Transaction,
    initial: SummaryProperties,
    properties: SummaryProperties,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "changed",
                &self
                    .state
                    .get()
                    .is_some_and(|state| state.initial != state.properties),
            )
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(base: &Package) -> Result<Self> {
        if base.is_encrypted() {
            return Err(Error::Encrypted);
        }
        let expected_source = base.source_version().map_err(Error::Package)?;
        Ok(Self {
            base: base.clone(),
            expected_source,
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> Result<State> {
        let bytes = source_bytes(&self.base, self.expected_source)?;
        let transaction = property_set::Snapshot::from_bytes(bytes)?.transaction()?;
        let initial = SummaryProperties::from_editor(transaction.editor())?;
        Ok(State {
            transaction,
            properties: initial.clone(),
            initial,
        })
    }

    fn state(&self) -> Result<&State> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> Result<&mut State> {
        let base = self.base.clone();
        let expected_source = self.expected_source;
        self.state.get_or_try_init_mut(|| {
            let bytes = source_bytes(&base, expected_source)?;
            let transaction = property_set::Snapshot::from_bytes(bytes)?.transaction()?;
            let initial = SummaryProperties::from_editor(transaction.editor())?;
            Ok::<State, Error>(State {
                transaction,
                properties: initial.clone(),
                initial,
            })
        })
    }

    /// The staged presentation properties.
    #[must_use = "handle the fallible metadata projection result"]
    pub fn properties(&self) -> Result<&SummaryProperties> {
        Ok(&self.state()?.properties)
    }

    /// Mutable access to the staged properties: set or remove one property
    /// by name ([`SummaryProperties::set`], [`SummaryProperties::remove`]),
    /// custom properties, and typed section edits.
    pub fn properties_mut(&mut self) -> Result<&mut SummaryProperties> {
        Ok(&mut self.state_mut()?.properties)
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
        Ok(self.state_mut()?.properties.merge_metadata(metadata)?)
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
            expected_source,
            state,
        } = self;
        ensure_source_version(&base, expected_source)?;
        state.get_or_try_init(|| {
            let bytes = source_bytes(&base, expected_source)?;
            let transaction = property_set::Snapshot::from_bytes(bytes)?.transaction()?;
            let initial = SummaryProperties::from_editor(transaction.editor())?;
            Ok::<State, Error>(State {
                transaction,
                properties: initial.clone(),
                initial,
            })
        })?;
        let State {
            mut transaction,
            initial,
            properties,
        } = state.into_inner().ok_or_else(|| {
            Error::Package(PackageError::Corrupted(
                "PPT metadata state was not initialized".into(),
            ))
        })?;
        if properties == initial {
            return unchanged(base, expected_source);
        }
        properties.stage_changes(&initial, transaction.editor_mut())?;
        let (_, patch) = transaction.commit()?.into_parts();
        ensure_source_version(&base, expected_source)?;
        let before: Arc<[u8]> = Arc::from(patch.source());
        let after: Arc<[u8]> = Arc::from(patch.replacement());
        let changed = before != after;
        let package = if changed {
            reopen(&base, &after)?
        } else {
            base
        };
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
        let changed = !patch.is_noop();
        let package = if changed {
            reopen(self, patch.after_bytes())?
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

fn unchanged(base: Package, expected: SourceVersion) -> Result<Commit> {
    let bytes: Arc<[u8]> = Arc::from(source_bytes(&base, expected)?);
    Ok(Commit::new(
        base,
        Patch {
            before: Arc::clone(&bytes),
            after: bytes,
        },
        Diagnostics::UNCHANGED,
    ))
}

fn source_bytes(base: &Package, expected: SourceVersion) -> Result<Vec<u8>> {
    ensure_source_version(base, expected)?;
    let bytes = base.to_bytes()?;
    ensure_source_version(base, expected)?;
    Ok(bytes)
}

fn ensure_source_version(base: &Package, expected: SourceVersion) -> Result<()> {
    let observed = base.source_version().map_err(Error::Package)?;
    if observed != expected {
        return Err(Error::Package(PackageError::Ole(OleError::SourceChanged {
            expected,
            observed,
        })));
    }
    Ok(())
}

fn reopen(base: &Package, bytes: &[u8]) -> Result<Package> {
    let options = OpenOptions::default()
        .source_limits(base.source_limits())
        .leniency(base.leniency());
    #[cfg(feature = "vba-inspection")]
    let options = options.vba_project_limits(base.vba_project_limits());
    Ok(Package::from_bytes_with_options(bytes, options)?)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "the lazy-entry regression is a direct state assertion"
)]
mod tests {
    use super::{Edit, Package};

    #[test]
    fn metadata_edit_construction_and_drop_do_not_read_property_streams() {
        let package = Package::new().unwrap();
        let edit = Edit::new(&package).unwrap();
        assert!(edit.state.get().is_none());
        drop(edit);
    }
}
