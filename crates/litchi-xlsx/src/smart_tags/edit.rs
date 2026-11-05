//! Immutable worksheet smart-tag edits of a [`crate::Package`] (record
//! 0360).
//!
//! [`crate::Package::edit_smart_tags`] starts an [`Edit`] that retains an
//! owned selector without loading the worksheet's smart tags; every
//! operation validates a private candidate, so an error cannot stage a
//! partial worksheet.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_smart_tags`] replays a patch onto an equal
//! package. The source package never changes.
//!
//! An operation that would change the worksheet of a signed package is
//! refused with [`Error::Signed`] and stages nothing; one that leaves the
//! worksheet bytes as they are (removing an absent annotation, clearing an
//! empty collection) publishes nothing and keeps the signature byte for byte.
//! Remove the signature explicitly first ([`crate::Package::unsign`]) to edit
//! a signed workbook.

use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::At;

use super::model::{Cell, Collection};
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("smart-tag");

pub use crate::opc::LimitKind;

/// A staged smart-tag edit of one worksheet of a [`Package`]. Dropping it
/// discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    package: Package,
    selector: crate::workbook::Selector<'static>,
}

impl Package {
    /// Start a smart-tag edit of a semantic worksheet without loading its
    /// worksheet part or smart-tag collection.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Selector, worksheet-kind, and smart-tag parse errors are
    /// reported by the first operation or by [`Edit::commit`].
    pub fn edit_smart_tags<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Edit> {
        let staged = Staged::new(self, "edit_smart_tags")?;
        let selector = crate::workbook::source::own_selector(sheet.into())?;
        Ok(Edit {
            staged,
            package: self.clone(),
            selector,
        })
    }

    /// Apply a smart-tag patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_smart_tags(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_smart_tags")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// Return the current staged collection. Absence is represented by an
    /// empty semantic collection.
    pub fn collection(&self) -> Result<Collection> {
        let worksheet = self.worksheet()?;
        Ok(super::package::load(self.staged.opc(), &worksheet)?.unwrap_or_default())
    }

    /// Replace the complete collection. An empty value removes `smartTags`.
    pub fn replace(&mut self, value: Collection) -> Result<()> {
        self.apply(move |candidate, worksheet| {
            super::package::store(candidate, worksheet, Some(&value))
        })
    }

    /// Add or replace all annotations for one checked cell.
    pub fn set(&mut self, value: Cell) -> Result<()> {
        let mut collection = self.collection()?;
        collection.upsert(value);
        super::validation::collection(&collection)?;
        self.replace(collection)
    }

    /// Remove one cell's annotations and return the removed value.
    pub fn remove<'key>(&mut self, at: impl Into<At<'key>>) -> Result<Option<Cell>> {
        let address = at.into().resolve()?;
        let mut collection = self.collection()?;
        let removed = collection.remove(address);
        if removed.is_some() {
            self.replace(collection)?;
        }
        Ok(removed)
    }

    /// Remove the complete worksheet collection.
    pub fn clear(&mut self) -> Result<bool> {
        if self.collection()?.is_empty() {
            return Ok(false);
        }
        self.apply(|candidate, worksheet| super::package::store(candidate, worksheet, None))?;
        Ok(true)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Validate the staged worksheet with a final bounded parse and publish
    /// it to a new package. An unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the parse errors, [`Error::Signed`] for a changed signed
    /// package and the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        let worksheet = self.worksheet()?;
        super::package::load(self.staged.opc(), &worksheet)?;
        Ok(Commit::new(self.staged.commit()?))
    }

    fn worksheet(&self) -> Result<PackURI> {
        let worksheet = self
            .package
            .workbook_ref()
            .sheet(self.selector.clone())?
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                "edit_smart_tags",
                "the worksheet selector matches no sheet",
            ))?;
        if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: worksheet.name().to_owned(),
            });
        }
        Ok(worksheet.part_uri().clone())
    }

    fn apply<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage, &PackURI) -> Result<T>,
    ) -> Result<T> {
        let worksheet = self.worksheet()?;
        self.staged.update_part(&worksheet, |candidate| {
            let value = operation(candidate, &worksheet)?;
            super::package::load(candidate, &worksheet)?;
            Ok(value)
        })
    }
}
