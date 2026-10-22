//! Document-property edits of a [`SourceBackedWorkbook`] that publish over
//! the retained source.
//!
//! [`SourceBackedWorkbook::edit_metadata`] starts an [`Edit`] over the
//! property parts read from the source; [`Edit::commit`] returns a
//! [`Commit`] whose new owner keeps every workbook part, cache and semantic
//! read of the source owner and replaces only the property parts when it is
//! published ([`SourceBackedWorkbook::write_to`]). Nothing is materialized:
//! no worksheet, style or shared-string part is read, and the workbook is
//! not revalidated, because the property parts carry no workbook semantics.
//! The commit checks the patch against the source and refuses a signed
//! source; the source topology that replaces the parts is planned once, on
//! the first publication or materialization.
//!
//! The transaction is the shared OOXML engine's
//! ([`litchi_ooxml_common::properties`]), so field routing, lexical
//! preservation, refusals ([`Error::Unsupported`], [`Error::Signed`]) and
//! the exact no-op contract are those of [`super::Edit`].
//!
//! ```
//! use litchi_xlsb::{Package, SourceBackedWorkbook};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let bytes = Package::new()?.to_bytes()?;
//! let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(bytes))?;
//! let mut edit = source.edit_metadata()?;
//! edit.update(|metadata| metadata.title = Some("Budget".to_owned()))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let edited = commit.into_package();
//! assert_eq!(edited.metadata()?.title.as_deref(), Some("Budget"));
//! assert_eq!(source.metadata()?.title, None);
//! let mut published = Vec::new();
//! edited.write_to(&mut published)?;
//! assert_eq!(Package::from_bytes(&published)?.metadata()?.title.as_deref(), Some("Budget"));
//! # Ok(())
//! # }
//! ```

use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_ooxml_common::custom::{self, Host};
use litchi_ooxml_common::properties::{self, Snapshot, app};
use litchi_opc::PreparedTopology;

use super::{Diagnostics, Error, Patch};
use crate::SourceBackedWorkbook;

type Result<T> = std::result::Result<T, Error>;

/// The document-property state a source owner publishes over its source.
pub(crate) struct Revision {
    /// The properties read from the source.
    base: Snapshot,
    /// The properties after every published edit.
    current: Snapshot,
    /// The composed patch from `base` to `current`, checked against the
    /// source when it was committed.
    patch: Patch,
    /// `patch` planned over the source package, on first publication or
    /// materialization.
    topology: OnceLock<PreparedTopology>,
}

impl std::fmt::Debug for Revision {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Revision")
            .field("changed_families", &self.patch.changes().len())
            .finish_non_exhaustive()
    }
}

impl Revision {
    /// The revision planned over `source`, prepared once.
    pub(crate) fn topology(
        &self,
        source: &litchi_opc::SourceBackedPackage,
    ) -> Result<&PreparedTopology> {
        if let Some(topology) = self.topology.get() {
            return Ok(topology);
        }
        let prepared = self.patch.prepare_source_backed(source)?;
        // A concurrent first publication may have won; either plan is the
        // same deterministic topology.
        let _ = self.topology.set(prepared);
        self.topology
            .get()
            .ok_or_else(|| Error::Invalid("the metadata topology was not retained".to_owned()))
    }
}

/// A staged metadata edit of one [`SourceBackedWorkbook`].
///
/// Dropping the edit discards it.
pub struct Edit {
    /// The source owner this edit stages over (a share of it).
    workbook: SourceBackedWorkbook,
    /// The properties this edit started from.
    start: Snapshot,
    inner: properties::Plan,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The staged metadata.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn metadata(&self) -> Result<Metadata> {
        self.inner.metadata()
    }

    /// Stage the complete metadata.
    pub fn set(&mut self, metadata: Metadata) -> &mut Self {
        self.inner.set_metadata(metadata);
        self
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn update(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        self.inner.update_metadata(update)?;
        Ok(self)
    }

    /// Stage the complete application properties; `None` removes the part.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> &mut Self {
        self.inner.set_app_props(props);
        self
    }

    /// The staged custom properties.
    #[must_use]
    pub fn custom_props(&self) -> &custom::Props {
        self.inner.custom_props()
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.inner.set_custom_props(props);
        self
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.inner = self.start.plan();
    }

    /// Compute the patch and publish it over the source.
    ///
    /// An exact no-op returns the same owner. A changed edit plans the
    /// property parts over the source package; the returned owner shares
    /// every other part and cache with this one.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, [`Error::Signed`] for a changed edit of a signed
    /// workbook, and the source-change and topology errors of the planner.
    pub fn commit(self) -> Result<Commit> {
        let (current, patch, diagnostics) = self.inner.finish()?;
        if patch.is_noop() {
            return Ok(Commit {
                package: self.workbook,
                patch,
                diagnostics,
            });
        }
        let source = self.workbook.source_backed_package();
        let (base, composed) = match &self.workbook.metadata {
            Some(revision) => (revision.base.clone(), revision.patch.compose(&patch)?),
            None => (self.start, patch.clone()),
        };
        // The plan itself is prepared on first publication; its preconditions
        // are checked now, so a stale or signed source fails the commit.
        if !composed.is_noop() {
            composed.check_source_backed(source)?;
            if source.has_signature_infrastructure() {
                return Err(Error::Signed);
            }
        }
        let package = self
            .workbook
            .with_metadata_revision((!composed.is_noop()).then(|| {
                Arc::new(Revision {
                    base,
                    current,
                    patch: composed,
                    topology: OnceLock::new(),
                })
            }));
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// The result of [`Edit::commit`].
pub struct Commit {
    package: SourceBackedWorkbook,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl Commit {
    /// The source owner after the edit.
    #[must_use]
    pub const fn package(&self) -> &SourceBackedWorkbook {
        &self.package
    }

    /// Consume the commit into the source owner after the edit.
    #[must_use]
    pub fn into_package(self) -> SourceBackedWorkbook {
        self.package
    }

    /// The exact, reversible patch of this edit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the edit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into the source owner and the patch.
    #[must_use]
    pub fn into_parts(self) -> (SourceBackedWorkbook, Patch) {
        (self.package, self.patch)
    }
}

impl SourceBackedWorkbook {
    fn metadata_snapshot(&self) -> Result<Snapshot> {
        match &self.metadata {
            Some(revision) => Ok(revision.current.clone()),
            None => Snapshot::read_source_backed(self.source_backed_package(), Host::Excel),
        }
    }

    /// The document metadata: the core and application properties.
    ///
    /// Only the property parts are read from the source.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the property parts,
    /// and a source-change or cancellation error.
    pub fn metadata(&self) -> Result<Metadata> {
        self.metadata_snapshot()?.metadata()
    }

    /// The OOXML application properties (`docProps/app.xml`), when present.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::metadata`].
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.metadata_snapshot()?.app_props()
    }

    /// Start a metadata edit that publishes over the retained source.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::metadata`].
    pub fn edit_metadata(&self) -> Result<Edit> {
        let start = self.metadata_snapshot()?;
        Ok(Edit {
            workbook: self.clone(),
            inner: start.plan(),
            start,
        })
    }

    /// Apply a metadata patch to a new source-backed owner.
    ///
    /// The patch is checked against the recorded property-part bytes before
    /// the new owner is returned. Replaying a patch therefore reads only the
    /// property families it changes and retains every other source member
    /// lazily for [`SourceBackedWorkbook::write_to`]. A patch created from an
    /// eager package can be replayed here when its recorded content base is
    /// the same.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a different content base,
    /// [`Error::Signed`] for a changed signed source, and the source and
    /// property decoding errors of the patch.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: patch.diagnostics(),
            });
        }
        let source = self.source_backed_package();
        let (base, composed) = match &self.metadata {
            Some(revision) => (revision.base.clone(), revision.patch.compose(patch)?),
            None => (self.metadata_snapshot()?, patch.clone()),
        };
        let current = Snapshot::read_source_backed_after(source, Host::Excel, &composed)?;
        if source.has_signature_infrastructure() {
            return Err(Error::Signed);
        }
        let package = self.with_metadata_revision(Some(Arc::new(Revision {
            base,
            current,
            patch: composed,
            topology: OnceLock::new(),
        })));
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: patch.diagnostics(),
        })
    }
}
