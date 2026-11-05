//! The classic-chart family (record 0355): the main document's chart
//! ownership graph (chart parts, embedded workbooks and companions).
//!
//! [`Package::chart_graph`](crate::Package::chart_graph) reads the graph;
//! [`Package::edit_chart`](crate::Package::edit_chart) stages an [`Edit`]
//! whose commit publishes a new package, and
//! [`Package::apply_chart`](crate::Package::apply_chart) replays its
//! [`Patch`]. Chart part identities and ownership never change.

use super::transaction::Transaction as Draft;
use super::{Graph, Resource};
use crate::package::family::{Staged, graph_family};

graph_family!("chart", own_limits);

pub use super::limit::LimitKind;

/// A staged chart-graph change over one package snapshot. The source is
/// never changed: [`Self::commit`] publishes a new package, and dropping the
/// edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    draft: Option<Draft>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn draft(&mut self) -> Result<&mut Draft> {
        if self.draft.is_none() {
            let opc = self.staged.working().opc_package();
            let document = opc.main_document_part()?.partname().clone();
            self.draft = Some(Draft::new(opc, &document)?);
        }
        Ok(self.draft.as_mut().expect("chart draft was created above"))
    }

    /// The staged graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the chart graph cannot be read.
    pub fn graph(&mut self) -> Result<&Graph> {
        Ok(self.draft()?.graph())
    }

    /// The staged charts in document anchor order.
    ///
    /// # Errors
    ///
    /// Returns an error when the chart graph cannot be read.
    pub fn charts(&mut self) -> Result<&[Resource]> {
        Ok(self.draft()?.charts())
    }

    /// Replace the complete graph after bounded validation; returns whether
    /// it changed. Chart identities and ownership must stay the same.
    ///
    /// # Errors
    ///
    /// Returns an error for a graph the validation refuses.
    pub fn replace_graph(&mut self, graph: Graph) -> Result<bool> {
        self.draft()?.replace_graph(graph)
    }

    /// Apply a checked, failure-atomic mutation to the staged graph.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or a graph the validation refuses; the
    /// staged graph is then unchanged.
    pub fn edit_graph(&mut self, edit: impl FnOnce(&mut Graph) -> Result<()>) -> Result<&mut Self> {
        self.draft()?.edit_graph(edit)?;
        Ok(self)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Publish the staged graph as a new package with its exact patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale source or a graph the publication
    /// refuses, and the signed-package refusal for a change to a signed
    /// package.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some(draft) = self.draft.take() {
            let mut working = self.staged.working().clone();
            working.edit_semantic_opc("chart.edit", |opc| draft.commit(opc).map(|_| ()))?;
            *self.staged.working_mut() = working;
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a chart-graph change over this package. O(1).
    #[must_use]
    pub fn edit_chart(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            draft: None,
        }
    }

    /// Replay a chart-graph patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_chart(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_chart")
    }
}
