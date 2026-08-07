//! Typed formula CRUD for Keynote slide tables.

use super::*;
use litchi_keynote::slide::table::formula::{FormulaCachedValue, FormulaExpression};

impl KeynoteEditor {
    /// Read a cell's canonical formula source, including the leading equals sign.
    pub fn slide_table_formula(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<String>> {
        let table = self.slide_table(slide_index, model_object_id)?;
        Ok(match table.get_cell(row, column) {
            Some(KeynoteTableCellValue::Formula(formula)) => Some(formula.clone()),
            _ => None,
        })
    }

    /// Create or replace a cell formula transactionally.
    pub fn set_slide_table_formula(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        expression: FormulaExpression,
        cached_value: FormulaCachedValue,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
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
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_formula(slide_index, model_object_id, row, column)?
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "Keynote table formula update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete a cell formula and return its canonical source text.
    pub fn clear_slide_table_formula(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<String> {
        let formula = self
            .slide_table_formula(slide_index, model_object_id, row, column)?
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote table cell ({row}, {column}) does not contain a formula"
                ))
            })?;
        self.clear_slide_table_cell(slide_index, model_object_id, row, column)?;
        Ok(formula)
    }
}
