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
        let scope_key = self.active_pivot_scope.as_ref().ok_or_else(|| {
            Error::InvalidFormula(
                "PtgSxName requires an explicit pivot cache, sheet, and view scope".to_string(),
            )
        })?;
        let cache_id = scope_key.cache_id();
        let sheet_index = scope_key.sheet_index();
        let view_name = scope_key.view_name();
        if sheet_index >= self.worksheet_names.len() {
            return Err(Error::InvalidFormula(format!(
                "pivot sheet index {sheet_index} is outside the workbook sheet range"
            )));
        }
        if self.current_sheet != Some(sheet_index) {
            return Err(Error::InvalidFormula(format!(
                "pivot scope sheet {sheet_index} does not match the formula sheet {:?}",
                self.current_sheet
            )));
        }

        let mut views = self.pivot_views.iter().filter(|view| {
            view.cache_id == cache_id
                && view.sheet_index == sheet_index
                && view.name.eq_ignore_ascii_case(view_name)
        });
        let _view = views.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} does not use cache {cache_id}"
            ))
        })?;
        if views.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} and cache {cache_id} is ambiguous"
            )));
        }

        let mut scopes = self.pivot_name_scopes.iter().filter(|scope| {
            scope.cache_id == cache_id
                && scope.sheet_index == sheet_index
                && scope.view_name.eq_ignore_ascii_case(view_name)
        });
        let scope = scopes.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "calculated-name metadata is missing for PivotTable view {view_name:?}"
            ))
        })?;
        if scopes.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "calculated-name metadata for PivotTable view {view_name:?} is ambiguous"
            )));
        }
        Ok(scope)
    }

    pub(super) fn resolve_pivot_name(&self, index: u32) -> Result<String> {
        let scope = self.validate_active_pivot_scope()?;
        let index = usize::try_from(index).map_err(|_| {
            Error::InvalidFormula("pivot calculated-name index overflow".to_string())
        })?;
        let reference = scope.references.get(index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "pivot calculated-name index {index} is outside 0..{}",
                scope.references.len()
            ))
        })?;
        Ok(reference.to_formula_text())
    }
}
