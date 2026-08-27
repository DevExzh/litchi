//! Typed sort-rule editing and execution for Keynote slide tables.

use super::*;
use litchi_keynote::slide::table::{
    lock::State as FocusedTableLockState,
    sort::{Order, RowRange},
};

impl KeynoteEditor {
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
        let (order, lock_state) = focused_table_sort_context(self, slide_index, model_object_id)?;
        reject_locked_table(lock_state)?;
        let order = order.ok_or_else(|| {
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
        if focused_table_sort_order(&verified, slide_index, model_object_id)?.as_ref()
            != Some(&order)
        {
            return Err(Error::InvalidFormat(
                "Keynote table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Execute a slide table's configured selected-row sort over one range.
    ///
    /// The half-open range is body-relative, excluding header and footer rows.
    /// Returns `true` when one or more selected rows moved.
    pub fn apply_slide_table_sort_order_to_rows(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        rows: RowRange,
    ) -> Result<bool> {
        require_table_model(self, slide_index, model_object_id)?;
        let (order, lock_state) = focused_table_sort_context(self, slide_index, model_object_id)?;
        reject_locked_table(lock_state)?;
        let order = order.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Keynote table sort without a configured table sort order"
                    .to_owned(),
            )
        })?;
        let mut staged = self.package().clone();
        if !crate::numbers::editor::apply_table_sort_order_to_rows_in_package(
            &mut staged,
            model_object_id,
            &order,
            rows,
        )? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if focused_table_sort_order(&verified, slide_index, model_object_id)?.as_ref()
            != Some(&order)
        {
            return Err(Error::InvalidFormat(
                "Keynote selected-row table sort did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }
}

/// Read persisted sort configuration through the focused package owner while
/// retaining the historical model identifier at the physical executor edge.
fn focused_table_sort_order(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<Option<Order>> {
    Ok(focused_table_sort_context(editor, slide_index, model_object_id)?.0)
}

/// Read the persisted order and lock state before handing the model to the
/// legacy physical executor.  The focused package owns both persisted facts;
/// the executor must not move rows in a locked table.
fn focused_table_sort_context(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<(Option<Order>, FocusedTableLockState)> {
    let tables = editor.slide_tables(slide_index)?;
    let mut table_index = None;
    for (index, table) in tables.iter().enumerate() {
        if table.model_object_id != model_object_id {
            continue;
        }
        if table_index.replace(index).is_some() {
            return Err(Error::ParseError(format!(
                "Keynote object {model_object_id} has ambiguous table ownership on slide {slide_index}"
            )));
        }
    }
    let table_index = table_index.ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote table model {model_object_id} is not owned by slide {slide_index}"
        ))
    })?;
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote sort source: {error}")))?;
    let order = package
        .slide_table_sort_order(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote sort read: {error}")))?;
    let lock_state = package
        .slide_table_lock_state(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote lock read: {error}")))?;
    Ok((order, lock_state))
}

fn reject_locked_table(lock_state: FocusedTableLockState) -> Result<()> {
    if lock_state == FocusedTableLockState::Locked {
        return Err(Error::InvalidFormat(
            "Cannot execute a Keynote table sort on a locked table".to_owned(),
        ));
    }
    Ok(())
}
