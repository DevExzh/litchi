//! Coherent publication of inert, shared Office Add-in metadata.

use crate::web_extensions::{self as web, Conformance, Limits, Panes, Patch};
use crate::{Package, Result};

impl Package {
    /// Load inert persisted Office Add-in task panes with standard limits.
    pub fn task_panes(&self) -> Result<Option<Panes>> {
        self.task_panes_with_limits(&Limits::standard())
    }

    /// Load task panes with explicit XML, metadata and package-graph limits.
    ///
    /// Runtime IDs, custom-function IDs and external references are inert data.
    /// This operation never starts an add-in or fetches an external target.
    pub fn task_panes_with_limits(&self, limits: &Limits) -> Result<Option<Panes>> {
        self.ensure_story_opc_current("task_panes_with_limits")?;
        Ok(web::load_with(self.opc_package(), limits)?)
    }

    /// Plan a reversible task-pane graph replacement without changing this package.
    pub fn plan_task_panes(&self, panes: Panes, conformance: Conformance) -> Result<Patch> {
        self.plan_task_panes_with_limits(panes, conformance, &Limits::standard())
    }

    /// Plan a task-pane replacement with explicit resource limits.
    ///
    /// The returned shared patch keeps physical part and relationship selection
    /// within the graph owner. Apply it with [`Self::apply_task_panes_patch`].
    /// For a changed signed package, call [`Self::unsign`] before planning:
    /// removing signatures can change the source metadata guarded by a patch.
    pub fn plan_task_panes_with_limits(
        &self,
        panes: Panes,
        conformance: Conformance,
        limits: &Limits,
    ) -> Result<Patch> {
        self.ensure_story_opc_current("plan_task_panes_with_limits")?;
        Ok(web::plan_put_with(
            self.opc_package(),
            panes,
            conformance,
            limits,
        )?)
    }

    /// Plan removal of task panes and exclusively owned graph resources.
    pub fn plan_remove_task_panes(&self) -> Result<Patch> {
        self.plan_remove_task_panes_with_limits(&Limits::standard())
    }

    /// Plan task-pane removal with explicit resource limits.
    pub fn plan_remove_task_panes_with_limits(&self, limits: &Limits) -> Result<Patch> {
        self.ensure_story_opc_current("plan_remove_task_panes_with_limits")?;
        Ok(web::plan_remove_with(self.opc_package(), limits)?)
    }

    /// Apply a task-pane patch after validating its source and package policy.
    ///
    /// Exact no-ops preserve signatures. Changed signed packages require
    /// [`Self::unsign`] followed by a new plan before publication. Failure
    /// leaves this package intact.
    /// The same entry point applies [`Patch::inverse`].
    pub fn apply_task_panes_patch(&mut self, patch: &Patch) -> Result<bool> {
        self.ensure_story_opc_current("apply_task_panes_patch")?;
        if patch.is_empty() {
            return Ok(patch.apply(&mut self.opc)?);
        }
        self.edit_semantic_opc("apply_task_panes_patch", |candidate| {
            Ok(patch.apply(candidate)?)
        })
    }

    /// Verify and apply a durable task-pane patch with standard graph limits.
    ///
    /// Obtain the durable value with [`Patch::to_durable`], or decode it with
    /// the core patch API and explicit wire limits. Application replays the
    /// typed edit and checks its source before publishing. The same entry point
    /// accepts the durable inverse.
    pub fn apply_durable_task_panes_patch<Mode>(
        &mut self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<bool> {
        self.apply_durable_task_panes_patch_with_limits(patch, &Limits::standard())
    }

    /// Verify and apply a durable task-pane patch with caller graph limits.
    ///
    /// Restoration data is checked by replaying the original semantic edit.
    /// Signed no-ops preserve their source; changed signed sources require
    /// explicit unsigned disposition before creating a new patch. A stale
    /// source, invalid proof or exceeded limit leaves this package intact.
    pub fn apply_durable_task_panes_patch_with_limits<Mode>(
        &mut self,
        patch: &litchi_core::patch::Patch<Mode>,
        limits: &Limits,
    ) -> Result<bool> {
        self.ensure_story_opc_current("apply_durable_task_panes_patch_with_limits")?;
        let planned = web::plan_durable_with_limits(self.opc_package(), patch, limits)?;
        self.apply_task_panes_patch(&planned)
    }

    /// Store a validated task-pane graph by moving it into package ownership.
    pub fn put_task_panes(&mut self, panes: Panes, conformance: Conformance) -> Result<&mut Self> {
        self.put_task_panes_with_limits(panes, conformance, &Limits::standard())
    }

    /// Store task panes atomically with explicit resource limits.
    pub fn put_task_panes_with_limits(
        &mut self,
        panes: Panes,
        conformance: Conformance,
        limits: &Limits,
    ) -> Result<&mut Self> {
        let patch = self.plan_task_panes_with_limits(panes, conformance, limits)?;
        self.apply_task_panes_patch(&patch)?;
        Ok(self)
    }

    /// Remove task panes and resources no longer shared elsewhere.
    pub fn remove_task_panes(&mut self) -> Result<bool> {
        self.remove_task_panes_with_limits(&Limits::standard())
    }

    /// Remove task panes atomically with explicit resource limits.
    pub fn remove_task_panes_with_limits(&mut self, limits: &Limits) -> Result<bool> {
        let patch = self.plan_remove_task_panes_with_limits(limits)?;
        self.apply_task_panes_patch(&patch)
    }
}
