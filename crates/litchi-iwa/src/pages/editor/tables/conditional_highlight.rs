//! Conditional-highlight inspection and deletion for Pages body-table cells.

use super::*;

/// Storage identity and rule count for one conditionally highlighted Pages cell.
pub type PagesTableCellConditionalHighlightInfo =
    crate::numbers::editor::TableCellConditionalHighlightInfo;

impl PagesEditor {
    /// Inspect conditional highlighting attached to a reachable body-table cell.
    pub fn table_cell_conditional_highlighting(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<PagesTableCellConditionalHighlightInfo>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_conditional_highlighting_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Read supported ordered conditional-highlight rules from a body-table cell.
    pub fn table_cell_conditional_highlight_rules(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<
        Option<Vec<crate::table_cell_conditional_highlight::TableCellConditionalHighlightRule>>,
    > {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_conditional_highlight_rules_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Delete conditional highlighting without changing the cell value or base style.
    pub fn clear_table_cell_conditional_highlighting(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::clear_table_cell_conditional_highlighting_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .table_cell_conditional_highlighting(model_object_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Pages conditional-highlight deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a body-table cell's conditional highlighting and return its storage identity.
    pub fn set_table_cell_conditional_highlighting(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        rules: &[crate::table_cell_conditional_highlight::TableCellConditionalHighlightRule],
    ) -> Result<PagesTableCellConditionalHighlightInfo> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_conditional_highlighting_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            rules,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let actual = verified
            .table_cell_conditional_highlighting(model_object_id, row, column)?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Pages conditional-highlight creation failed validation".to_owned(),
                )
            })?;
        if actual.rule_count as usize != rules.len() {
            return Err(Error::InvalidFormat(
                "Pages conditional-highlight rule count failed validation".to_owned(),
            ));
        }
        if verified
            .table_cell_conditional_highlight_rules(model_object_id, row, column)?
            .as_deref()
            != Some(rules)
        {
            return Err(Error::InvalidFormat(
                "Pages conditional-highlight rules failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(actual)
    }
}
