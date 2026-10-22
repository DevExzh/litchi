//! Immutable worksheet slicer edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_slicers`] starts an [`Edit`] over the validated
//! slicer graph; every operation works on a private candidate graph, so an
//! error cannot stage a partial graph. [`Edit::commit`] validates the
//! complete owner graph and returns a [`Commit`] holding the new package,
//! the exact part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_slicers`] replays a patch onto an equal package.
//! The source package never changes.
//!
//! An operation that would change a signed package is refused with
//! [`Error::Signed`] and stages nothing (ADR 0006); an operation that changes
//! nothing (an absent name, a reorder into the stored bytes) keeps the
//! signature byte for byte. Remove the signature explicitly first
//! ([`crate::Package::unsign`]) to edit a signed workbook.

use super::model::{Cache, Definition, Part, Slicer};
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};
use litchi_opc::{OpcPackage, PackURI};

family!("slicer");

pub use crate::opc::LimitKind;

/// A staged slicer edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a slicer edit after validating the existing slicer graph.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and the validation errors of the slicer graph.
    pub fn edit_slicers(&self) -> Result<Edit> {
        let staged = Staged::new(self, "edit_slicers")?;
        super::validation::graph(staged.opc())?;
        Ok(Edit { staged })
    }

    /// Apply a slicer patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_slicers(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_slicers")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged OPC graph, for diagnostics only.
    #[must_use]
    pub fn opc(&self) -> &OpcPackage {
        self.staged.opc()
    }

    /// Find a cache by its case-insensitive semantic name.
    pub fn cache(&self, name: &str) -> Result<Option<Cache>> {
        crate::slicer_cache::crud::find_slicer_cache(self.staged.opc(), name)
    }

    /// Find a worksheet slicer by its case-insensitive semantic name.
    pub fn view(&self, worksheet: &PackURI, name: &str) -> Result<Option<Slicer>> {
        crate::slicer_cache::crud::find_slicer(self.staged.opc(), worksheet, name)
    }

    pub fn add_cache(&mut self, definition: Definition) -> Result<Cache> {
        self.apply(|candidate| crate::slicer_cache::crud::add_slicer_cache(candidate, definition))
    }

    pub fn update_cache<F>(&mut self, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Definition),
    {
        self.apply(|candidate| {
            crate::slicer_cache::crud::update_slicer_cache(candidate, name, update)
        })
    }

    pub fn replace_cache(&mut self, name: &str, replacement: Definition) -> Result<bool> {
        self.apply(|candidate| {
            crate::slicer_cache::crud::replace_slicer_cache(candidate, name, replacement)
        })
    }

    pub fn reorder_caches(&mut self, names: &[String]) -> Result<Vec<Cache>> {
        let names = names.to_vec();
        self.apply(|candidate| crate::slicer_cache::crud::reorder_slicer_caches(candidate, &names))
    }

    pub fn remove_cache(&mut self, name: &str) -> Result<bool> {
        self.apply(|candidate| crate::slicer_cache::crud::remove_slicer_cache(candidate, name))
    }

    pub fn add_view(&mut self, worksheet: &PackURI, view: Slicer) -> Result<Part> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_slicer(candidate, &worksheet, view)
        })
    }

    pub fn update_view<F>(&mut self, worksheet: &PackURI, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Slicer),
    {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::update_slicer(candidate, &worksheet, name, update)
        })
    }

    pub fn replace_view(
        &mut self,
        worksheet: &PackURI,
        name: &str,
        replacement: Slicer,
    ) -> Result<bool> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::replace_slicer(candidate, &worksheet, name, replacement)
        })
    }

    pub fn reorder_views(&mut self, worksheet: &PackURI, names: &[String]) -> Result<Vec<Slicer>> {
        let worksheet = worksheet.clone();
        let names = names.to_vec();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::reorder_slicers(candidate, &worksheet, &names)
        })
    }

    pub fn remove_view(&mut self, worksheet: &PackURI, name: &str) -> Result<bool> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::remove_slicer(candidate, &worksheet, name)
        })
    }

    /// Slicer filtering and rendering are intentionally outside this inert
    /// package owner.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::Unsupported`].
    pub fn apply_filter(&mut self) -> Result<()> {
        crate::slicer::unsupported_ui()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Validate the staged slicer graph and publish it to a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors, [`Error::Signed`] for a changed
    /// signed package and the workbook validation errors of the staged
    /// package.
    pub fn commit(self) -> Result<Commit> {
        super::validation::graph(self.staged.opc())?;
        Ok(Commit::new(self.staged.commit()?))
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.staged.update(|candidate| {
            let value = operation(candidate)?;
            super::validation::graph(candidate)?;
            Ok(value)
        })
    }
}
