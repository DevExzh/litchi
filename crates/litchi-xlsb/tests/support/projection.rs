//! An owned workbook projection of a package, for integration tests that
//! keep a [`Workbook`] value across steps. A package's own projection is
//! borrowed through `Package::workbook`; this parses a fresh one from the
//! same OPC graph under the package's retained policies.

use litchi_xlsb::package::PackageResult;
use litchi_xlsb::{OpenOptions, Package, Workbook};

/// Parse an owned [`Workbook`] projection of a package.
pub(crate) trait OwnedWorkbook {
    /// A fresh projection of this package's graph under its retained
    /// external-link limits, leniency, drawing policy and semantic limits.
    fn owned_workbook(&self) -> PackageResult<Workbook>;
}

impl OwnedWorkbook for Package {
    fn owned_workbook(&self) -> PackageResult<Workbook> {
        Workbook::from_opc_package_with_options(
            self.opc_package().clone(),
            OpenOptions::new()
                .external_link_limits(self.external_link_limits())
                .leniency(self.leniency())
                .parse_drawings(self.drawings_parsed())
                .semantic_limits(self.semantic_limits()),
        )
    }
}

/// A projection of the same graph under other semantic budgets.
#[allow(dead_code, reason = "each test target uses the steps it needs")]
pub(crate) trait Reprojected {
    /// This workbook's graph projected under `limits`.
    fn reprojected(&self, limits: litchi_xlsb::SemanticLimits) -> PackageResult<Workbook>;
}

impl Reprojected for Workbook {
    fn reprojected(&self, limits: litchi_xlsb::SemanticLimits) -> PackageResult<Workbook> {
        self.to_package()
            .with_semantic_limits(limits)
            .owned_workbook()
    }
}

/// In-place publication of package edits onto an owned [`Workbook`]
/// projection, for integration tests written as a sequence of workbook
/// states: each step runs the public package route on the workbook's
/// package and adopts the projection of the published package.
#[allow(dead_code, reason = "each test target uses the steps it needs")]
pub(crate) trait InPlace {
    /// Run one raw OPC update through `Package::edit_opc` and adopt the
    /// committed package.
    fn edit_opc_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_opc::OpcPackage) -> PackageResult<T>,
    ) -> PackageResult<T>;

    /// Replay a structural patch through `Package::apply` and adopt the
    /// published package.
    fn apply_structure(
        &mut self,
        patch: &litchi_xlsb::cell_values::structure::Patch,
    ) -> PackageResult<()>;

    /// Undo the history tip through `History::undo` and adopt the result.
    fn undo_with(
        &mut self,
        history: &mut litchi_xlsb::cell_values::structure::History,
    ) -> PackageResult<()>;

    /// Redo through `History::redo` and adopt the result.
    fn redo_with(
        &mut self,
        history: &mut litchi_xlsb::cell_values::structure::History,
    ) -> PackageResult<()>;
}

impl InPlace for Workbook {
    fn edit_opc_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_opc::OpcPackage) -> PackageResult<T>,
    ) -> PackageResult<T> {
        let mut edit = self.to_package().edit_opc()?;
        let value = edit.update(operation)?;
        *self = edit.commit()?.into_package().owned_workbook()?;
        Ok(value)
    }

    fn apply_structure(
        &mut self,
        patch: &litchi_xlsb::cell_values::structure::Patch,
    ) -> PackageResult<()> {
        *self = self
            .to_package()
            .apply(patch)?
            .into_package()
            .owned_workbook()?;
        Ok(())
    }

    fn undo_with(
        &mut self,
        history: &mut litchi_xlsb::cell_values::structure::History,
    ) -> PackageResult<()> {
        *self = history
            .undo(&self.to_package())?
            .into_package()
            .owned_workbook()?;
        Ok(())
    }

    fn redo_with(
        &mut self,
        history: &mut litchi_xlsb::cell_values::structure::History,
    ) -> PackageResult<()> {
        *self = history
            .redo(&self.to_package())?
            .into_package()
            .owned_workbook()?;
        Ok(())
    }
}
