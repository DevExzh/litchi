//! Clone-staged slicer package transaction.

use super::model::{Cache, Definition, Part, Slicer};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

/// A slicer-only transaction over an OPC snapshot.
///
/// Every operation works on a private candidate. An error therefore cannot
/// publish a partial graph, and dropping the transaction is an implicit
/// rollback. The commit boundary performs the complete owner graph check.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    staged: OpcPackage,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating the existing slicer graph.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        super::validation::graph(target)?;
        Ok(Self {
            staged: target.clone(),
            target,
        })
    }

    /// Borrow the staged package for diagnostics only.
    #[must_use]
    pub fn package(&self) -> &OpcPackage {
        &self.staged
    }

    /// Find a cache by its case-insensitive semantic name.
    pub fn cache(&self, name: &str) -> Result<Option<Cache>> {
        crate::slicer_cache::crud::find_slicer_cache(&self.staged, name)
    }

    /// Find a worksheet slicer by its case-insensitive semantic name.
    pub fn view(&self, worksheet: &PackURI, name: &str) -> Result<Option<Slicer>> {
        crate::slicer_cache::crud::find_slicer(&self.staged, worksheet, name)
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
    pub fn apply_filter(&mut self) -> Result<()> {
        crate::slicer::unsupported_ui()
    }

    /// Publish the staged graph. Dropping without calling this discards it.
    pub fn commit(self) -> Result<()> {
        super::validation::graph(&self.staged)?;
        *self.target = self.staged;
        Ok(())
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.staged.clone();
        let value = operation(&mut candidate)?;
        super::validation::graph(&candidate)?;
        self.staged = candidate;
        Ok(value)
    }
}
