//! Package-level cell-value routes (record 0359a F1).
//!
//! The cell-value engine (`cell_values::{Edit, Commit, Patch}` and the
//! snapshot constructors) is crate-private: [`Package::edit`] is its only
//! staging route and every commit is the primary [`Commit`]. This module
//! publishes on [`Package`] what the engine had beyond the primary edit's
//! verbs: durable semantic-patch replay, the three-way merge of two semantic
//! patches, dependency-checked transfer to another workbook, and an
//! undo/redo [`History`] of packages. [`Package::cell_values`] is the read
//! view and [`super::edit::Edit::join`] composes disjoint edits.

use litchi_core::patch::HistoryLimits;

use super::Package;
use super::edit::{Commit, Edit, Error, Result};
use crate::cell_values::{SemanticPatch, ThreeWayPlan, TransferPlan};

/// Bounded undo/redo history over immutable XLS packages: record each
/// [`Commit::package`] with its weight (for example its byte length).
pub type History = litchi_core::patch::History<Package>;

impl Package {
    /// Replay a durable semantic patch (`Patch::semantic`, or one read back
    /// with [`SemanticPatch::from_deterministic_json`]) onto this package.
    ///
    /// # Errors
    ///
    /// Returns a malformed-operation, stale-precondition, storage or commit
    /// error; this package is unchanged.
    pub fn apply_semantic(&self, patch: &SemanticPatch) -> Result<Commit> {
        let (snapshot, options) = self.cell_source()?;
        let inner = patch.stage(&snapshot).map_err(Error::from)?;
        Edit::from_leaf(self.clone(), inner, options)?.commit()
    }

    /// Plan the deterministic three-way merge of two semantic patches made
    /// from this package.
    ///
    /// # Errors
    ///
    /// Returns a stale-precondition, malformed-operation, allocation or
    /// patch-bound error.
    pub fn plan_three_way(
        &self,
        left: &SemanticPatch,
        right: &SemanticPatch,
    ) -> Result<ThreeWayPlan> {
        let (snapshot, _) = self.cell_source()?;
        SemanticPatch::plan_three_way(&snapshot, left, right).map_err(Error::from)
    }

    /// Preflight the dependency-aware transfer of `patch` onto this
    /// package; the plan records each refusal.
    ///
    /// # Errors
    ///
    /// Returns an error when this workbook cannot be read as cell values.
    pub fn plan_transfer(&self, patch: &SemanticPatch) -> Result<TransferPlan> {
        let (snapshot, _) = self.cell_source()?;
        Ok(patch.plan_transfer(&snapshot))
    }

    /// Execute an executable transfer plan onto this package.
    ///
    /// # Errors
    ///
    /// Returns the plan's first refusal or a typed commit error.
    pub fn execute_transfer(&self, plan: &TransferPlan) -> Result<Commit> {
        let (snapshot, options) = self.cell_source()?;
        let inner = plan.stage(&snapshot).map_err(Error::from)?;
        Edit::from_leaf(self.clone(), inner, options)?.commit()
    }

    /// A bounded undo/redo history starting at this package.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }
}
