//! The task-pane family (record 0355): the package's inert Office Add-in
//! web extensions and task panes, shared with the other OOXML hosts
//! ([`crate::web_extensions`]).
//!
//! [`Package::task_panes`](crate::Package::task_panes) reads the panes and
//! [`Package::plan_task_panes`](crate::Package::plan_task_panes) plans a
//! replacement; [`Package::edit_task_panes`](crate::Package::edit_task_panes)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_task_panes`](crate::Package::apply_task_panes) replays its
//! [`Patch`]. Add-in code is never fetched or run.
//!
//! ```
//! use litchi_docx::{Package, web_extensions as extensions};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::new()?;
//! let reference = extensions::Reference::new("addin", "1.0", extensions::Store::Omex)?;
//! let mut panes = extensions::Panes::new();
//! panes.push(extensions::Pane::new(extensions::AddIn::new("addin-instance", reference)?))?;
//! let mut edit = package.edit_task_panes();
//! edit.put(panes, extensions::Conformance::Transitional)?;
//! let commit = edit.commit()?;
//! assert!(commit.package().task_panes()?.is_some());
//! let undone = commit.package().apply_task_panes(&commit.patch().inverse())?;
//! assert!(undone.package().task_panes()?.is_none());
//! # Ok(())
//! # }
//! ```

use crate::package::family::{Staged, graph_family};
use crate::web_extensions::{self as web, Conformance, Limits, Panes};

graph_family!("task-pane", own_limits);

/// The typed budgets of the task-pane family: its limit errors are
/// [`litchi_ooxml_common::Error::Limit`] carried by
/// [`crate::Error::Common`] (record 0360).
pub use litchi_ooxml_common::web::LimitKind;

/// The configuration of a task-pane edit.
#[derive(Debug, Clone, Default)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Bound every planned graph by `limits` instead of the package's
    /// semantic task-pane limits.
    #[must_use]
    pub fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// The configured limits; `None` uses the package's semantic limits.
    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// A staged task-pane change over one package snapshot. The source is never
/// changed: [`Self::commit`] publishes a new package, and dropping the edit
/// discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    limits: Limits,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package with the changes staged so far.
    #[must_use]
    pub const fn package(&self) -> &crate::Package {
        self.staged.working()
    }

    fn stage(&mut self, plan: &web::Patch) -> Result<bool> {
        let mut working = self.staged.working().clone();
        let changed = working.apply_task_panes_patch(plan)?;
        *self.staged.working_mut() = working;
        Ok(changed)
    }

    /// Replace the task panes with `panes`; returns whether the package
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for panes the configured limits refuse.
    pub fn put(&mut self, panes: Panes, conformance: Conformance) -> Result<bool> {
        let plan = web::plan_put(
            self.staged.working().opc_package(),
            panes,
            conformance,
            &self.limits,
        )?;
        self.stage(&plan)
    }

    /// Remove the task panes and their unreferenced web extensions; returns
    /// whether the package changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a graph the configured limits refuse.
    pub fn remove(&mut self) -> Result<bool> {
        let plan = web::plan_remove(self.staged.working().opc_package(), &self.limits)?;
        self.stage(&plan)
    }

    /// Stage a planned graph patch ([`crate::Package::plan_task_panes`] and
    /// [`crate::Package::plan_remove_task_panes`]); returns whether the
    /// package changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged package does not hold the plan's
    /// before state.
    pub fn apply_plan(&mut self, plan: &web::Patch) -> Result<bool> {
        self.stage(plan)
    }

    /// Plan and stage a common durable patch; returns whether the package
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a durable vocabulary or graph the configured
    /// limits refuse.
    pub fn apply_durable<Mode>(&mut self, patch: &litchi_core::patch::Patch<Mode>) -> Result<bool> {
        let plan = web::plan_durable(self.staged.working().opc_package(), patch, &self.limits)?;
        self.stage(&plan)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes as a new package with their exact patch.
    ///
    /// # Errors
    ///
    /// Returns a package-graph failure while capturing the patch and the
    /// signed-package refusal for a change to a signed package.
    pub fn commit(self) -> Result<Commit> {
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a task-pane change over this package under its semantic
    /// task-pane limits. O(1).
    #[must_use]
    pub fn edit_task_panes(&self) -> Edit {
        self.edit_task_panes_with_options(EditOptions::default())
    }

    /// Stage a task-pane change over this package under `options`. O(1).
    #[must_use]
    pub fn edit_task_panes_with_options(&self, options: EditOptions) -> Edit {
        Edit {
            staged: Staged::new(self),
            limits: options
                .limits_value()
                .unwrap_or_else(|| self.semantic_limits().task_panes_value()),
        }
    }

    /// Replay a task-pane patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_task_panes(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_task_panes")
    }
}
