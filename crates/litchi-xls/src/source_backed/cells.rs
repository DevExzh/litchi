//! The source-backed fixed-width cell-value family: number edits of existing
//! `Number`, `RK` and `MulRk` value fields published as same-length splices
//! (the former `cell_values::Edit::commit_source_backed_plan`).

use crate::Result;
use crate::SourceBackedWorkbook;
use crate::cell_values::{Reference, Selector, Value};

source_family!("cell", "edit cells");

/// Content-free source-backed cell publication diagnostics.
pub use crate::cell_values::source_diagnostics::Diagnostics;

/// A staged source-backed cell edit. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    source: SourceBackedWorkbook,
    inner: Option<crate::cell_values::Edit>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("materialized", &self.inner.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn inner_mut(&mut self) -> Result<&mut crate::cell_values::Edit> {
        if self.inner.is_none() {
            let package = self
                .source
                .staging_package()
                .map_err(|error| crate::source_backed::protected_edit_error(error, "edit cells"))?;
            let snapshot = package.cell_values().map_err(|error| {
                crate::source_backed::protected_edit_error(
                    crate::scoped::package_error(error),
                    "edit cells",
                )
            })?;
            self.inner = Some(snapshot.edit());
        }
        self.inner.as_mut().ok_or(Error::Internal(
            "source-backed cell edit engine was not retained",
        ))
    }

    /// Stage a number into an existing numeric cell.
    ///
    /// # Errors
    ///
    /// Returns a selector or cell refusal; the staged state is unchanged.
    pub fn set_number<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        self.inner_mut()?
            .set_number(selector.into(), reference, value)?;
        Ok(())
    }

    /// Stage a numeric value, choosing the record form the cell holds.
    ///
    /// # Errors
    ///
    /// Returns a selector or cell refusal; the staged state is unchanged.
    pub fn set_numeric<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        self.inner_mut()?
            .set_numeric(selector.into(), reference, value)?;
        Ok(())
    }

    /// Stage a typed value; only numbers into existing numeric fields
    /// publish source-backed.
    ///
    /// # Errors
    ///
    /// Returns a selector or cell refusal; the staged state is unchanged.
    pub fn set_value<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        self.inner_mut()?
            .set_value(selector.into(), reference, value)?;
        Ok(())
    }

    /// Validate the staged edits and commit them as a workbook that writes
    /// the checked overlay plan. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns a refusal for an unsupported, protected, signed, encrypted,
    /// macro-bearing, stale or length-changing edit, or a failed readback.
    pub fn commit(self) -> Result<Commit> {
        // The cell-values edit owns a parsed snapshot, while this family edit
        // retains the dynamic source-backed owner. Fence the latter around
        // planning and publication so stage -> mutate -> commit cannot
        // publish an overlay validated against stale source bytes.
        let source = self.source.clone();
        source.ensure_original_source_current()?;
        let inner = match self.inner {
            Some(inner) => inner,
            None => {
                let package = source.staging_package().map_err(|error| {
                    crate::source_backed::protected_edit_error(error, "edit cells")
                })?;
                let snapshot = package.cell_values().map_err(|error| {
                    crate::source_backed::protected_edit_error(
                        crate::scoped::package_error(error),
                        "edit cells",
                    )
                })?;
                snapshot.edit()
            },
        };
        let planned = inner.commit_source_backed_plan()?;
        source.ensure_original_source_current()?;
        let (plan, diagnostics) = planned.into_plan();
        let commit = published_deferred(&source, plan, diagnostics)?;
        source.ensure_original_source_current()?;
        Ok(commit)
    }
}

fn published_deferred(
    before: &SourceBackedWorkbook,
    plan: litchi_cfb::ValidatedOverlayPlan,
    diagnostics: Diagnostics,
) -> Result<Commit> {
    let after = before
        .committed_deferred(plan)
        .map_err(crate::source_backed::workbook_error)?;
    let patch = Patch {
        before: before.clone(),
        after: after.clone(),
        diagnostics,
    };
    Ok(Commit::new(after, patch, diagnostics))
}

impl SourceBackedWorkbook {
    /// Start a source-backed cell edit of this workbook.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook cannot be read as cell values.
    pub fn edit_cells(&self) -> Result<Edit> {
        self.ensure_original_source_current()?;
        Ok(Edit {
            source: self.clone(),
            inner: None,
        })
    }

    /// Replay a source-backed cell patch onto matching workbook content,
    /// including an independently opened equal-content workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for another workbook.
    pub fn apply_cells(&self, patch: &Patch) -> Result<Commit> {
        replay(self, patch)
    }
}
