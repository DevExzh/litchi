//! Immutable task-pane edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_task_panes`] starts an [`Edit`] of the
//! package-level Office Add-in task-pane graph; every operation works on a
//! private candidate graph, so an error cannot stage a partial graph.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] (the task-pane parts and the package-level
//! relationships) and the [`Diagnostics`];
//! [`crate::Package::apply_task_panes`] replays a patch onto an equal
//! package. The source package never changes. A changed edit of a signed
//! workbook is refused with [`Error::Signed`]; an exact no-op keeps every
//! byte.

use litchi_ooxml_common::web::{self as common_web, Conformance, Pane, Panes, Selector};
use litchi_opc::OpcPackage;

use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("task-pane");

pub use crate::opc::LimitKind;

/// The configuration of a task-pane edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    conformance: Option<Conformance>,
}

impl EditOptions {
    /// Write new task-pane relationships in `conformance` instead of the
    /// existing graph's conformance (Transitional when the package has no
    /// task-pane graph).
    #[must_use]
    pub const fn conformance(mut self, conformance: Conformance) -> Self {
        self.conformance = Some(conformance);
        self
    }
}

/// A staged task-pane edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    package: Package,
    conformance: Option<Conformance>,
}

impl Package {
    /// Start a task-pane edit without loading any existing graph. Existing
    /// XML conformance is retained when a non-empty graph is written; new
    /// graphs use Transitional `SpreadsheetML` relationships unless
    /// [`Self::edit_task_panes_with_options`] selects another.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Existing graph errors are reported by the first operation or
    /// by [`Edit::commit`].
    pub fn edit_task_panes(&self) -> Result<Edit> {
        self.edit_task_panes_with_options(EditOptions::default())
    }

    /// Start a task-pane edit with explicit options.
    ///
    /// # Errors
    ///
    /// See [`Self::edit_task_panes`].
    pub fn edit_task_panes_with_options(&self, options: EditOptions) -> Result<Edit> {
        let staged = Staged::new(self, "edit_task_panes")?;
        Ok(Edit {
            staged,
            package: self.clone(),
            conformance: options.conformance,
        })
    }

    /// Apply a task-pane patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_task_panes(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_task_panes")?,
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

    /// Return a semantic snapshot of the current task panes.
    pub fn panes(&self) -> Result<Panes> {
        Ok(
            common_web::load(self.staged.opc(), &common_web::Limits::standard())
                .map_err(Error::from)?
                .unwrap_or_default(),
        )
    }

    /// Replace the complete graph. An empty collection removes the graph.
    pub fn replace(&mut self, panes: Panes) -> Result<()> {
        let conformance = if panes.is_empty() {
            None
        } else {
            let conformance = match self.conformance {
                Some(conformance) => conformance,
                None => super::existing_conformance(self.package.opc())?,
            };
            self.conformance = Some(conformance);
            Some(conformance)
        };
        self.apply(move |candidate| {
            if let Some(conformance) = conformance {
                let patch = common_web::plan_put(
                    candidate,
                    panes,
                    conformance,
                    &common_web::Limits::standard(),
                )
                .map_err(Error::from)?;
                let commit = patch.apply(candidate).map_err(Error::from)?;
                let changed = commit.diagnostics().changed();
                *candidate = commit.into_package();
                Ok(changed)
            } else {
                let patch = common_web::plan_remove(candidate, &common_web::Limits::standard())
                    .map_err(Error::from)?;
                let commit = patch.apply(candidate).map_err(Error::from)?;
                let changed = commit.diagnostics().changed();
                *candidate = commit.into_package();
                Ok(changed)
            }
        })
        .map(|_| ())
    }

    /// Add one task pane, allocating a deterministic relationship ID when
    /// the model leaves it empty.
    pub fn add(&mut self, pane: Pane) -> Result<()> {
        let mut panes = self.panes()?;
        panes.push(pane).map_err(Error::from)?;
        self.replace(panes)
    }

    /// Edit one pane by semantic add-in ID or collection index.
    ///
    /// The common model edits a cloned pane and validates collection-wide
    /// uniqueness before replacing it, so a failed closure leaves this
    /// edit's staged graph unchanged.
    pub fn edit<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
        edit: impl FnOnce(&mut Pane) -> litchi_ooxml_common::Result<()>,
    ) -> Result<bool> {
        let mut panes = self.panes()?;
        let changed = panes.edit(selector, edit).map_err(Error::from)?;
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
        self.apply(|candidate| {
            let patch = common_web::plan_remove(candidate, &common_web::Limits::standard())
                .map_err(Error::from)?;
            let commit = patch.apply(candidate).map_err(Error::from)?;
            let changed = commit.diagnostics().changed();
            *candidate = commit.into_package();
            Ok(changed)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Validate the staged graph and publish it to a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors, [`Error::Signed`] for a changed
    /// signed package and the workbook validation errors of the staged
    /// package.
    pub fn commit(self) -> Result<Commit> {
        common_web::load(self.staged.opc(), &common_web::Limits::standard())
            .map_err(Error::from)?;
        Ok(Commit::new(self.staged.commit()?))
    }

    fn apply<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.staged.update(|candidate| {
            let value = operation(candidate)?;
            common_web::load(candidate, &common_web::Limits::standard()).map_err(Error::from)?;
            Ok(value)
        })
    }
}
