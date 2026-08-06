//! Clone-staged worksheet smart-tag edits.

use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::At;

use super::model::{Cell, Collection};
use crate::error::Result;

/// An atomic transaction over one worksheet's inert smart tags.
///
/// Every operation validates a cloned candidate. Dropping without calling
/// [`commit`](Self::commit) leaves the source package unchanged.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    staged: OpcPackage,
    worksheet: PackURI,
}

impl<'a> Transaction<'a> {
    pub(crate) fn new(target: &'a mut OpcPackage, worksheet: PackURI) -> Result<Self> {
        super::package::load(target, &worksheet)?;
        Ok(Self {
            staged: target.clone(),
            target,
            worksheet,
        })
    }

    /// Return the current staged collection. Absence is represented by an
    /// empty semantic collection.
    pub fn collection(&self) -> Result<Collection> {
        Ok(super::package::load(&self.staged, &self.worksheet)?.unwrap_or_default())
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

    /// Publish the staged worksheet after a final bounded parse and
    /// validation pass.
    pub fn commit(self) -> Result<()> {
        super::package::load(&self.staged, &self.worksheet)?;
        *self.target = self.staged;
        Ok(())
    }

    fn apply<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage, &PackURI) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.staged.clone();
        let value = operation(&mut candidate, &self.worksheet)?;
        super::package::load(&candidate, &self.worksheet)?;
        self.staged = candidate;
        Ok(value)
    }
}
