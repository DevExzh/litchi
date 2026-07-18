//! Typed formula CRUD for Pages body tables.

use super::*;

/// A typed whole-row or whole-column reference used by a table formula.
pub type PagesTableFormulaAxisReference = crate::numbers::FormulaAxisReference;
/// A typed binary operator used by a table formula.
pub type PagesTableFormulaBinaryOperator = crate::numbers::FormulaBinaryOperator;
/// A typed formula AST compiled into native iWork table storage.
pub type PagesTableFormulaExpression = crate::numbers::FormulaExpression;
/// A typed formula result displayed before Pages next recalculates the cell.
pub type PagesTableFormulaCachedValue = crate::numbers::FormulaCachedValue;
/// A typed cell address used by a table formula.
pub type PagesTableFormulaCellReference = crate::numbers::FormulaCellReference;

impl PagesEditor {
    /// Read a cell's canonical formula source, including the leading equals sign.
    pub fn table_formula(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<String>> {
        let table = self.table(model_object_id)?;
        Ok(match table.get_cell(row, column) {
            Some(PagesCellValue::Formula(formula)) => Some(formula.clone()),
            _ => None,
        })
    }

    /// Create or replace a cell formula transactionally.
    pub fn set_table_formula(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        expression: PagesTableFormulaExpression,
        cached_value: PagesTableFormulaCachedValue,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_formula_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            expression,
            cached_value,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified
            .table_formula(model_object_id, row, column)?
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "Pages table formula update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete a cell formula and return its source text.
    pub fn clear_table_formula(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<String> {
        let formula = self
            .table_formula(model_object_id, row, column)?
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table cell ({row}, {column}) does not contain a formula"
                ))
            })?;
        self.clear_table_cell(model_object_id, row, column)?;
        Ok(formula)
    }
}
