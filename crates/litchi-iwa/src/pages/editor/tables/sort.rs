//! Typed sort-rule execution for Pages body tables.

use super::*;
use litchi_pages::table::sort::{Order, RowRange};
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

/// Rows targeted by a persisted Pages table sort configuration.
pub type PagesTableSortRowRange = RowRange;

fn focused_table_sort_source(
    editor: &PagesEditor,
    model_object_id: u64,
) -> Result<(PagesPackage, BodyTableSelector<'static>)> {
    let table_index = editor
        .tables()?
        .iter()
        .position(|table| table.model_object_id == model_object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table model object {model_object_id} is missing"
            ))
        })?;
    let package = PagesPackage::from_bytes(&editor.to_bytes()?)
        .map_err(|error| Error::InvalidFormat(format!("Pages table sort package: {error}")))?;
    Ok((package, BodyTableSelector::index(table_index)))
}

fn focused_table_sort_order(editor: &PagesEditor, model_object_id: u64) -> Result<Option<Order>> {
    let (package, selector) = focused_table_sort_source(editor, model_object_id)?;
    package
        .body_table_sort_order(selector)
        .map_err(|error| Error::InvalidFormat(format!("Pages table sort: {error}")))
}

impl PagesEditor {
    /// Execute a body table's configured full-table sort order.
    ///
    /// Persisted sort configuration is read through the selector-first
    /// `litchi-pages` package owner. This compatibility executor remains in
    /// the umbrella crate because it physically reorders Pages body rows.
    pub fn apply_table_sort_order(&mut self, model_object_id: u64) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let order = focused_table_sort_order(self, model_object_id)?.ok_or_else(|| {
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
        if focused_table_sort_order(&verified, model_object_id)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Pages table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Execute a body table's configured selected-row sort over one range.
    ///
    /// The half-open range is body-relative, excluding header and footer rows.
    /// Persisted configuration is owned by `litchi-pages`; row movement remains
    /// in this compatibility host until Pages owns the physical table graph.
    pub fn apply_table_sort_order_to_rows(
        &mut self,
        model_object_id: u64,
        rows: PagesTableSortRowRange,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let order = focused_table_sort_order(self, model_object_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Pages table sort without a configured table sort order"
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
        if focused_table_sort_order(&verified, model_object_id)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Pages selected-row table sort did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }
}
