//! Document metadata of a workbook: the core, application and custom
//! property families, edited as one exact and reversible transaction.
//!
//! [`Workbook::metadata`] reads the unified [`Metadata`];
//! [`Workbook::edit_metadata`] starts an [`Edit`] whose [`Edit::commit`]
//! returns a [`Commit`] holding the new snapshot, the [`Patch`] and the
//! [`Diagnostics`]. [`Workbook::apply_metadata`] replays a patch onto an
//! equal snapshot. The source snapshot never changes.
//!
//! The shared OOXML engine ([`litchi_ooxml_common::properties`]) routes each
//! field to the part that stores it and keeps the lexical form of every
//! unchanged value. A field no OOXML family stores is refused with
//! [`Error::Unsupported`]; a changed edit of a signed workbook is refused
//! with [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! This module names the current owner, [`Workbook`]. When the package
//! becomes the owner of edits (ADR 0034 W2), the entry points move to it
//! unchanged.
//!
//! ```
//! use litchi_xlsx::Workbook;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let workbook = Workbook::new()?;
//! let mut edit = workbook.edit_metadata()?;
//! edit.update(|metadata| metadata.title = Some("Budget".to_owned()))?;
//! let after = edit.commit()?.into_package();
//! assert_eq!(after.metadata()?.title.as_deref(), Some("Budget"));
//! assert_eq!(workbook.metadata()?.title, None);
//! # Ok(())
//! # }
//! ```

use litchi_core::Metadata;
use litchi_ooxml_common::custom::{self, Host};
use litchi_ooxml_common::properties::{self, Snapshot, app};

pub use litchi_ooxml_common::properties::{Diagnostics, Error, LimitKind, Patch};

use crate::Workbook;

type Result<T> = std::result::Result<T, Error>;

/// A staged metadata edit of one [`Workbook`] snapshot.
///
/// Dropping the edit discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    workbook: &'a Workbook,
    inner: properties::Edit,
}

impl Edit<'_> {
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

    /// Stage the complete custom properties; an empty collection removes the
    /// part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.inner.set_custom_props(props);
        self
    }

    /// Compute the patch and publish it to a new snapshot.
    ///
    /// An exact no-op returns the same snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// workbook.
    pub fn commit(self) -> Result<Commit> {
        let (_, patch, diagnostics) = self.inner.commit()?.into_parts();
        let workbook = self.workbook.apply_metadata(&patch)?;
        Ok(Commit {
            workbook,
            patch,
            diagnostics,
        })
    }
}

/// The result of [`Edit::commit`].
#[derive(Debug)]
pub struct Commit {
    workbook: Workbook,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The snapshot after the edit.
    #[must_use]
    pub const fn package(&self) -> &Workbook {
        &self.workbook
    }

    /// Consume the commit into the snapshot after the edit.
    #[must_use]
    pub fn into_package(self) -> Workbook {
        self.workbook
    }

    /// The exact, reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the edit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into the snapshot and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Workbook, Patch) {
        (self.workbook, self.patch)
    }
}

impl Workbook {
    fn metadata_snapshot(&self) -> Result<Snapshot> {
        Snapshot::read(self.pivot_package(), Host::Excel)
    }

    /// The document metadata: the core and application properties.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the property parts.
    pub fn metadata(&self) -> Result<Metadata> {
        self.metadata_snapshot()?.metadata()
    }

    /// The OOXML application properties (`docProps/app.xml`), when present.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the part.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.metadata_snapshot()?.app_props()
    }

    /// Start a metadata edit.
    ///
    /// # Errors
    ///
    /// Returns the strict graph errors of the property families.
    pub fn edit_metadata(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            workbook: self,
            inner: self.metadata_snapshot()?.edit(),
        })
    }

    /// Apply a metadata patch to a new snapshot; this one is unchanged.
    ///
    /// An empty patch returns this snapshot, a signed one included. The new
    /// snapshot keeps this one's encryption provenance and repair policy.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] when this snapshot is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        if patch.is_noop() {
            return Ok(self.clone());
        }
        let mut package = self.pivot_package().clone();
        patch.apply(&mut package)?;
        self.adopt_published_package(package)
            .map_err(|error| Error::Invalid(error.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edits_publish_a_new_snapshot_and_replay_exactly() {
        let workbook = Workbook::new().unwrap();
        let mut edit = workbook.edit_metadata().unwrap();
        edit.update(|metadata| {
            metadata.author = Some("Ada".to_owned());
            metadata.company = Some("Litchi".to_owned());
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (after, patch) = commit.into_parts();
        assert_eq!(after.metadata().unwrap().author.as_deref(), Some("Ada"));
        assert_eq!(
            after.app_props().unwrap().unwrap().company.as_deref(),
            Some("Litchi")
        );
        assert!(!workbook.same_content(&after).unwrap());

        let replayed = workbook.apply_metadata(&patch).unwrap();
        assert!(replayed.same_content(&after).unwrap());
        let restored = after.apply_metadata(&patch.inverse()).unwrap();
        assert!(restored.same_content(&workbook).unwrap());
    }

    #[test]
    fn no_op_edits_return_the_same_snapshot() {
        let workbook = Workbook::new().unwrap();
        let mut edit = workbook.edit_metadata().unwrap();
        edit.set(workbook.metadata().unwrap());
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert!(commit.package().same_content(&workbook).unwrap());
    }
}
