//! Conditional-highlight inspection and deletion for Keynote slide-table cells.

use super::*;

/// Storage identity and rule count for one conditionally highlighted Keynote cell.
pub type KeynoteTableCellConditionalHighlightInfo =
    crate::numbers::editor::TableCellConditionalHighlightInfo;

impl KeynoteEditor {
    /// Inspect conditional highlighting attached to a reachable slide-table cell.
    pub fn slide_table_cell_conditional_highlighting(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<KeynoteTableCellConditionalHighlightInfo>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_conditional_highlighting_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Delete conditional highlighting without changing the cell value or base style.
    pub fn clear_slide_table_cell_conditional_highlighting(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::clear_table_cell_conditional_highlighting_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_table_cell_conditional_highlighting(slide_index, model_object_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Keynote conditional-highlight deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}
