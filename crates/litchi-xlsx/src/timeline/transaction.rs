//! Clone-staged timeline package transaction.

use super::model::{Cache, CacheDefinition, Part, View};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

/// A timeline-only transaction over an OPC snapshot.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    workbook: PackURI,
    staged: OpcPackage,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating the existing timeline graph.
    pub fn new(target: &'a mut OpcPackage, workbook: &PackURI) -> Result<Self> {
        super::validation::graph(target, workbook)?;
        Ok(Self {
            staged: target.clone(),
            target,
            workbook: workbook.clone(),
        })
    }

    /// Borrow the staged package for diagnostics only.
    #[must_use]
    pub fn package(&self) -> &OpcPackage {
        &self.staged
    }

    /// Find a cache by its case-insensitive semantic name.
    pub fn cache(&self, name: &str) -> Result<Option<Cache>> {
        crate::slicer_cache::crud::find_timeline_cache(&self.staged, name)
    }

    /// Find a worksheet timeline by its case-insensitive semantic name.
    pub fn view(&self, worksheet: &PackURI, name: &str) -> Result<Option<View>> {
        crate::slicer_cache::crud::find_timeline(&self.staged, worksheet, name)
    }

    pub fn add_cache(&mut self, definition: CacheDefinition) -> Result<Cache> {
        self.apply(|candidate| crate::slicer_cache::crud::add_timeline_cache(candidate, definition))
    }

    pub fn update_cache<F>(&mut self, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut CacheDefinition),
    {
        self.apply(|candidate| {
            crate::slicer_cache::crud::update_timeline_cache(candidate, name, update)
        })
    }

    pub fn replace_cache(&mut self, name: &str, replacement: CacheDefinition) -> Result<bool> {
        self.apply(|candidate| {
            crate::slicer_cache::crud::replace_timeline_cache(candidate, name, replacement)
        })
    }

    pub fn reorder_caches(&mut self, names: &[String]) -> Result<Vec<Cache>> {
        let names = names.to_vec();
        self.apply(|candidate| {
            crate::slicer_cache::crud::reorder_timeline_caches(candidate, &names)
        })
    }

    pub fn remove_cache(&mut self, name: &str) -> Result<bool> {
        self.apply(|candidate| crate::slicer_cache::crud::remove_timeline_cache(candidate, name))
    }

    pub fn add_view(&mut self, worksheet: &PackURI, view: View) -> Result<Part> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::add_timeline(candidate, &worksheet, view)
        })
    }

    pub fn update_view<F>(&mut self, worksheet: &PackURI, name: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut View),
    {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::update_timeline(candidate, &worksheet, name, update)
        })
    }

    pub fn replace_view(
        &mut self,
        worksheet: &PackURI,
        name: &str,
        replacement: View,
    ) -> Result<bool> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::replace_timeline(candidate, &worksheet, name, replacement)
        })
    }

    pub fn reorder_views(&mut self, worksheet: &PackURI, names: &[String]) -> Result<Vec<View>> {
        let worksheet = worksheet.clone();
        let names = names.to_vec();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::reorder_timelines(candidate, &worksheet, &names)
        })
    }

    pub fn remove_view(&mut self, worksheet: &PackURI, name: &str) -> Result<bool> {
        let worksheet = worksheet.clone();
        self.apply(move |candidate| {
            crate::slicer_cache::crud::remove_timeline(candidate, &worksheet, name)
        })
    }

    /// Timeline filtering, refresh, and rendering are intentionally outside
    /// this inert package owner.
    pub fn apply_filter(&mut self) -> Result<()> {
        crate::timeline::unsupported_ui()
    }

    /// Publish the staged graph. Dropping without calling this discards it.
    pub fn commit(self) -> Result<()> {
        super::validation::graph(&self.staged, &self.workbook)?;
        *self.target = self.staged;
        Ok(())
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.staged.clone();
        let value = operation(&mut candidate)?;
        super::validation::graph(&candidate, &self.workbook)?;
        self.staged = candidate;
        Ok(value)
    }
}
