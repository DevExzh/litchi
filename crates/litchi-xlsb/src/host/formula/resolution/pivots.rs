#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! PivotTable formula-scope validation and calculated-name resolution.

use crate::package::error::{Error, Result};
use crate::package::formula::Scope;

use super::Context;

impl Context {
    pub(super) fn validate_active_pivot_scope(&self) -> Result<&Scope> {
        let scope_key = self.active_pivot_scope.as_ref().ok_or({
            Error::UnresolvedDependency(
                "PtgSxName requires an explicit pivot cache, sheet, and view scope",
            )
        })?;
        let cache_id = scope_key.cache_id();
        let sheet_index = scope_key.sheet_index();
        let view_name = scope_key.view_name();
        if sheet_index >= self.worksheet_names.len() {
            return Err(Error::UnresolvedDependency(
                "pivot sheet index is outside the workbook sheet range",
            ));
        }
        if self.current_sheet != Some(sheet_index) {
            return Err(Error::InvalidFormula(
                "pivot scope sheet does not match the formula sheet",
            ));
        }

        let mut views = self.pivot_views.iter().filter(|view| {
            view.cache_id == cache_id
                && view.sheet_index == sheet_index
                && view.name.eq_ignore_ascii_case(view_name)
        });
        let _view = views.next().ok_or(Error::UnresolvedDependency(
            "PivotTable view on sheet does not use cache",
        ))?;
        if views.next().is_some() {
            return Err(Error::UnresolvedDependency(
                "PivotTable view on sheet and cache is ambiguous",
            ));
        }

        let mut scopes = self.pivot_name_scopes.iter().filter(|scope| {
            scope.cache_id == cache_id
                && scope.sheet_index == sheet_index
                && scope.view_name.eq_ignore_ascii_case(view_name)
        });
        let scope = scopes.next().ok_or(Error::UnresolvedDependency(
            "calculated-name metadata is missing for PivotTable view",
        ))?;
        if scopes.next().is_some() {
            return Err(Error::UnresolvedDependency(
                "calculated-name metadata for PivotTable view is ambiguous",
            ));
        }
        Ok(scope)
    }

    pub(super) fn resolve_pivot_name(&self, index: u32) -> Result<String> {
        let scope = self.validate_active_pivot_scope()?;
        let index = usize::try_from(index)
            .map_err(|_| Error::InvalidFormula("pivot calculated-name index overflow"))?;
        let reference = scope
            .references
            .get(index)
            .ok_or(Error::UnresolvedDependency(
                "pivot calculated-name index is outside 0..",
            ))?;
        Ok(reference.to_formula_text())
    }
}
