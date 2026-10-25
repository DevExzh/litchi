//! Coherent publication of inert, shared Office Add-in metadata.

use crate::web_extensions::{self as web, Conformance, Limits, Panes, Patch};
use crate::{Package, Result};

impl Package {
    /// Load inert persisted Office Add-in task panes with standard limits.
    pub fn task_panes(&self) -> Result<Option<Panes>> {
        self.task_panes_bounded(&self.semantic_limits().task_panes_value())
    }

    /// Load task panes with explicit XML, metadata and package-graph limits.
    ///
    /// Runtime IDs, custom-function IDs and external references are inert data.
    /// This operation never starts an add-in or fetches an external target.
    pub(crate) fn task_panes_bounded(&self, limits: &Limits) -> Result<Option<Panes>> {
        self.ensure_story_opc_current("task_panes_bounded")?;
        Ok(web::load(self.opc_package(), limits)?)
    }

    /// Plan a reversible task-pane graph replacement without changing this package.
    pub fn plan_task_panes(&self, panes: Panes, conformance: Conformance) -> Result<Patch> {
        self.plan_task_panes_bounded(
            panes,
            conformance,
            &self.semantic_limits().task_panes_value(),
        )
    }

    /// Plan a task-pane replacement with explicit resource limits.
    ///
    /// The returned shared patch keeps physical part and relationship selection
    /// within the graph owner. Stage it with [`crate::task_panes::Edit::apply_plan`].
    /// For a changed signed package, call [`Self::unsign`] before planning:
    /// removing signatures can change the source metadata guarded by a patch.
    pub(crate) fn plan_task_panes_bounded(
        &self,
        panes: Panes,
        conformance: Conformance,
        limits: &Limits,
    ) -> Result<Patch> {
        self.ensure_story_opc_current("plan_task_panes_bounded")?;
        Ok(web::plan_put(
            self.opc_package(),
            panes,
            conformance,
            limits,
        )?)
    }

    /// Plan removal of task panes and exclusively owned graph resources.
    pub fn plan_remove_task_panes(&self) -> Result<Patch> {
        self.plan_remove_task_panes_bounded(&self.semantic_limits().task_panes_value())
    }

    /// Plan task-pane removal with explicit resource limits.
    pub(crate) fn plan_remove_task_panes_bounded(&self, limits: &Limits) -> Result<Patch> {
        self.ensure_story_opc_current("plan_remove_task_panes_bounded")?;
        Ok(web::plan_remove(self.opc_package(), limits)?)
    }

    /// Apply a task-pane patch after validating its source and package policy.
    ///
    /// Exact no-ops preserve signatures. Changed signed packages require
    /// [`Self::unsign`] followed by a new plan before publication. Failure
    /// leaves this package intact.
    /// The same entry point applies [`Patch::inverse`].
    pub(crate) fn apply_task_panes_patch(&mut self, patch: &Patch) -> Result<bool> {
        self.ensure_story_opc_current("apply_task_panes_patch")?;
        if patch.is_noop() {
            let commit = patch.apply(&self.opc)?;
            let changed = commit.diagnostics().changed();
            self.opc = commit.into_package().into();
            return Ok(changed);
        }
        self.edit_semantic_opc("apply_task_panes_patch", |candidate| {
            let commit = patch.apply(candidate)?;
            let changed = commit.diagnostics().changed();
            *candidate = commit.into_package();
            Ok(changed)
        })
    }
}
