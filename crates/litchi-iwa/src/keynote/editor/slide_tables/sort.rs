//! Typed full-table sort-rule editing for Keynote slide tables.

use super::*;

/// A validated zero-based physical column index used by a Keynote sort rule.
pub type KeynoteTableSortColumnIndex = crate::numbers::NumbersTableSortColumnIndex;
/// Sort direction for one Keynote table column.
pub type KeynoteTableSortDirection = crate::numbers::NumbersTableSortDirection;
/// One full-table sort-configuration rule in priority order.
pub type KeynoteTableSortRule = crate::numbers::NumbersTableSortRule;
/// An ordered, non-empty full-table Keynote sort-rule configuration.
pub type KeynoteTableSortOrder = crate::numbers::NumbersTableSortOrder;

impl KeynoteEditor {
    /// Read a slide table's full-table native sort-rule configuration.
    ///
    /// An empty native order is reported as `None`. Row-range sorts depend on
    /// transient selection state and are rejected rather than guessed.
    pub fn slide_table_sort_order(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<Option<KeynoteTableSortOrder>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_sort_order_in_package(self.package(), model_object_id)
    }

    /// Set a slide table's full-table native sort-rule configuration transactionally.
    ///
    /// This only configures the stored native rule. Use
    /// [`Self::apply_slide_table_sort_order`] to physically reorder the body
    /// rows.
    pub fn set_slide_table_sort_order(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        order: KeynoteTableSortOrder,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self
            .slide_table_sort_order(slide_index, model_object_id)?
            .as_ref()
            == Some(&order)
        {
            return Ok(());
        }

        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_sort_order_in_package(
            &mut staged,
            model_object_id,
            &order,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_sort_order(slide_index, model_object_id)?
            .as_ref()
            != Some(&order)
        {
            return Err(Error::InvalidFormat(
                "Keynote table sort order failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear a slide table's stored native sort rules transactionally.
    pub fn clear_slide_table_sort_order(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self
            .slide_table_sort_order(slide_index, model_object_id)?
            .is_none()
        {
            return Ok(());
        }

        let mut staged = self.package().clone();
        if !crate::numbers::editor::clear_table_sort_order_in_package(&mut staged, model_object_id)?
        {
            return Err(Error::InvalidFormat(
                "Keynote table sort order unexpectedly had no rules to clear".to_owned(),
            ));
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_sort_order(slide_index, model_object_id)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Keynote table sort-order clear failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Execute a slide table's configured full-table sort order.
    ///
    /// This physically reorders only body rows and retains the native rule.
    /// The supported scalar subset and safety checks are the same as for
    /// [`crate::numbers::NumbersEditor::apply_table_sort_order`]. Returns
    /// `true` when one or more rows moved and `false` for an already stable
    /// body order.
    pub fn apply_slide_table_sort_order(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let order = self
            .slide_table_sort_order(slide_index, model_object_id)?
            .ok_or_else(|| {
                Error::ParseError(
                    "Cannot execute a Keynote table sort without a configured table sort order"
                        .to_owned(),
                )
            })?;
        let mut staged = self.package().clone();
        if !crate::numbers::editor::apply_table_sort_order_in_package(
            &mut staged,
            model_object_id,
            &order,
        )? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified
            .slide_table_sort_order(slide_index, model_object_id)?
            .as_ref()
            != Some(&order)
        {
            return Err(Error::InvalidFormat(
                "Keynote table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }
}
