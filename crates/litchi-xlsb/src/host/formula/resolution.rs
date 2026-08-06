//! Contextual relationship resolution for XLSB formulas.
//!
//! The public surface remains intentionally small: [`Context`] is the
//! workbook-owned semantic facade, while the private children keep XTI,
//! external-name, structured-table, and PivotTable rules close to the domain
//! they serve.

mod context;
mod names;
mod pivots;
mod relationships;
mod tables;
mod validation;

pub use context::{Context, ExternalBook};
pub use relationships::SupportingLink;

use crate::formula::{Resolution, Result as FormulaResult, TableReference};

use self::validation::owner_formula_resolution;

impl Resolution for Context {
    fn sheet_prefix(&self, index: u16) -> FormulaResult<String> {
        owner_formula_resolution(self.resolve_sheet_prefix(index))
    }

    fn defined_name(&self, index: u32) -> FormulaResult<String> {
        owner_formula_resolution(self.resolve_defined_name(index))
    }

    fn external_name(&self, sheet_index: u16, name_index: u32) -> FormulaResult<String> {
        owner_formula_resolution(self.resolve_external_name(sheet_index, name_index))
    }

    fn table_reference(&self, reference: &TableReference) -> FormulaResult<String> {
        owner_formula_resolution(self.resolve_table_reference(reference))
    }

    fn pivot_name(&self, index: u32) -> FormulaResult<String> {
        owner_formula_resolution(self.resolve_pivot_name(index))
    }
}
