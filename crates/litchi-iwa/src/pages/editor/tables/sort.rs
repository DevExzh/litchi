//! Typed full-table sort-rule editing for Pages body tables.

use super::*;

/// A validated zero-based physical column index used by a Pages sort rule.
pub type PagesTableSortColumnIndex = crate::numbers::NumbersTableSortColumnIndex;
/// Sort direction for one Pages table column.
pub type PagesTableSortDirection = crate::numbers::NumbersTableSortDirection;
/// One full-table sort-configuration rule in priority order.
pub type PagesTableSortRule = crate::numbers::NumbersTableSortRule;
/// An ordered, non-empty full-table Pages sort-rule configuration.
pub type PagesTableSortOrder = crate::numbers::NumbersTableSortOrder;

impl PagesEditor {
    /// Read a body table's full-table native sort-rule configuration.
    ///
    /// An empty native order is reported as `None`. Row-range sorts depend on
    /// transient selection state and are rejected rather than guessed.
    pub fn table_sort_order(&self, model_object_id: u64) -> Result<Option<PagesTableSortOrder>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_sort_order_in_package(self.package(), model_object_id)
    }

    /// Set a body table's full-table native sort-rule configuration transactionally.
    ///
    /// This only configures the stored native rule. Use
    /// [`Self::apply_table_sort_order`] to physically reorder the body rows.
    pub fn set_table_sort_order(
        &mut self,
        model_object_id: u64,
        order: PagesTableSortOrder,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_sort_order(model_object_id)?.as_ref() == Some(&order) {
            return Ok(());
        }

        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_sort_order_in_package(
            &mut staged,
            model_object_id,
            &order,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_sort_order(model_object_id)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Pages table sort order failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear a body table's stored native sort rules transactionally.
    pub fn clear_table_sort_order(&mut self, model_object_id: u64) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_sort_order(model_object_id)?.is_none() {
            return Ok(());
        }

        let mut staged = self.package().clone();
        if !crate::numbers::editor::clear_table_sort_order_in_package(&mut staged, model_object_id)?
        {
            return Err(Error::InvalidFormat(
                "Pages table sort order unexpectedly had no rules to clear".to_owned(),
            ));
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_sort_order(model_object_id)?.is_some() {
            return Err(Error::InvalidFormat(
                "Pages table sort-order clear failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Execute a body table's configured full-table sort order.
    ///
    /// This physically reorders only body rows and retains the native rule.
    /// The supported scalar subset and safety checks are the same as for
    /// [`crate::numbers::NumbersEditor::apply_table_sort_order`]. Returns
    /// `true` when one or more rows moved and `false` for an already stable
    /// body order.
    pub fn apply_table_sort_order(&mut self, model_object_id: u64) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let order = self.table_sort_order(model_object_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Pages table sort without a configured table sort order"
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
        verified.require_body_table(model_object_id)?;
        if verified.table_sort_order(model_object_id)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Pages table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }
}
