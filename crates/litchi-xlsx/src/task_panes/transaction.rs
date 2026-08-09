//! Clone-staged task-pane CRUD.

use litchi_ooxml_common::web::{self as common_web, Conformance, Pane, Panes, Selector};
use litchi_opc::OpcPackage;

use crate::error::Result;

/// An atomic transaction over the package-level task-pane graph.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    staged: OpcPackage,
    conformance: Conformance,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating any existing task-pane graph.
    pub fn new(target: &'a mut OpcPackage, conformance: Conformance) -> Result<Self> {
        let staged = target.clone();
        common_web::load(&staged).map_err(crate::Error::from)?;
        Ok(Self {
            target,
            staged,
            conformance,
        })
    }

    /// Borrow the staged physical package for diagnostics.
    #[must_use]
    pub fn package(&self) -> &OpcPackage {
        &self.staged
    }

    /// Return a semantic snapshot of the current task panes.
    pub fn panes(&self) -> Result<Panes> {
        Ok(common_web::load(&self.staged)
            .map_err(crate::Error::from)?
            .unwrap_or_default())
    }

    /// Replace the complete graph. An empty collection removes the graph.
    pub fn replace(&mut self, panes: Panes) -> Result<()> {
        let conformance = self.conformance;
        self.apply(move |candidate| {
            if panes.is_empty() {
                common_web::remove(candidate)
                    .map(|_| ())
                    .map_err(crate::Error::from)
            } else {
                common_web::put(candidate, panes, conformance).map_err(crate::Error::from)
            }
        })
    }

    /// Add one task pane, allocating a deterministic relationship ID when
    /// the model leaves it empty.
    pub fn add(&mut self, pane: Pane) -> Result<()> {
        let mut panes = self.panes()?;
        panes.push(pane).map_err(crate::Error::from)?;
        self.replace(panes)
    }

    /// Edit one pane by semantic add-in ID or collection index.
    ///
    /// The common model edits a cloned pane and validates collection-wide
    /// uniqueness before replacing it, so a failed closure leaves this
    /// transaction's staged snapshot unchanged.
    pub fn edit<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
        edit: impl FnOnce(&mut Pane) -> litchi_ooxml_common::Result<()>,
    ) -> Result<bool> {
        let mut panes = self.panes()?;
        let changed = panes.edit(selector, edit).map_err(crate::Error::from)?;
        if changed {
            self.replace(panes)?;
        }
        Ok(changed)
    }

    /// Remove one pane by semantic add-in ID or collection index.
    pub fn remove<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Result<Option<Pane>> {
        let mut panes = self.panes()?;
        let removed = panes.remove(selector);
        if removed.is_some() {
            self.replace(panes)?;
        }
        Ok(removed)
    }

    /// Remove the complete task-pane graph.
    pub fn clear(&mut self) -> Result<bool> {
        self.apply(|candidate| common_web::remove(candidate).map_err(crate::Error::from))
    }

    /// Publish the staged package. Dropping without committing rolls back.
    pub fn commit(self) -> Result<()> {
        common_web::load(&self.staged).map_err(crate::Error::from)?;
        *self.target = self.staged;
        Ok(())
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.staged.clone();
        let value = operation(&mut candidate)?;
        common_web::load(&candidate).map_err(crate::Error::from)?;
        self.staged = candidate;
        Ok(value)
    }
}
