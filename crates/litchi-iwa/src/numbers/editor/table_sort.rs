//! Typed sort-rule configuration and execution for Numbers tables.

use super::*;
use litchi_numbers::TableSelector;
use litchi_numbers::table::sort::{self, ColumnIndex, Direction, Order, RowRange, Rule, Scope};

mod apply;
mod wire;

pub use litchi_numbers::table::sort::{
    ColumnIndex as NumbersTableSortColumnIndex, Direction as NumbersTableSortDirection,
    Order as NumbersTableSortOrder, RowRange as NumbersTableSortRowRange,
    Rule as NumbersTableSortRule, Scope as NumbersTableSortScope,
};

use apply::apply_attached_table_sort_order;
use wire::{
    clear_table_sort_order_wire, delete_table_sort_column_wire, read_native_table_sort_order_wire,
    write_table_sort_order_wire,
};

fn invalid_stored_sort(error: sort::Error) -> Error {
    Error::InvalidFormat(format!(
        "Numbers table has an invalid stored sort order: {error}"
    ))
}

fn order_from_native(sort: &tst::TableSortOrderArchive) -> Result<Option<Order>> {
    if sort.rules.is_empty() {
        return Ok(None);
    }
    let scope = Scope::from_native(sort.r#type).map_err(invalid_stored_sort)?;
    let rules = sort
        .rules
        .iter()
        .map(|rule| {
            let direction = Direction::from_native(rule.direction).map_err(invalid_stored_sort)?;
            let column = ColumnIndex::from_native(rule.index).map_err(invalid_stored_sort)?;
            Ok(Rule::new(column, direction))
        })
        .collect::<Result<Vec<_>>>()?;
    Order::with_scope(scope, rules)
        .map(Some)
        .map_err(invalid_stored_sort)
}

fn order_as_native(order: &Order) -> tst::TableSortOrderArchive {
    tst::TableSortOrderArchive {
        r#type: order.scope().native_value(),
        rules: order
            .rules()
            .iter()
            .map(|rule| tst::table_sort_order_archive::SortRuleArchive {
                index: rule.column().native_value(),
                direction: rule.direction().native_value(),
            })
            .collect(),
    }
}

impl NumbersEditor {
    /// Read an attached table's persisted sort-rule configuration.
    ///
    /// An empty native order is reported as `None`, matching the state shown
    /// by Numbers after its last sort rule is removed. Selected-row orders
    /// expose their persisted [`Scope::SelectedRows`] scope;
    /// their view-state selected interval is intentionally not guessed.
    pub fn table_sort_order(&self, selector: TableSelector) -> Result<Option<Order>> {
        let table_id = super::selectors::table_id(self, selector)?;
        table_sort_order_in_package(&self.package, table_id)
    }

    /// Set the persisted sort-rule configuration transactionally.
    ///
    /// The resulting file stores the same table-level order exposed in
    /// Numbers' **Organize → Sort** pane, including for spreadsheets created
    /// entirely by this crate. This operation configures the native rule; it
    /// does not execute it or reorder stored rows. Numbers exposes that
    /// separate action as **Sort Now**.
    pub fn set_table_sort_order(&mut self, selector: TableSelector, order: Order) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        set_table_sort_order_in_package(&mut staged, table_id, &order)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(selector)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Numbers table sort order failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Clear an attached table's stored sort rules transactionally.
    ///
    /// When a table already carries native sort metadata, this preserves
    /// Numbers' empty-order marker and any associated reference tracker,
    /// exactly as removing the final rule in the Numbers UI does.
    pub fn clear_table_sort_order(&mut self, selector: TableSelector) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        if !clear_table_sort_order_in_package(&mut staged, table_id)? {
            return Ok(());
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(selector)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table sort-order clear failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Execute the attached table's configured full-table sort order.
    ///
    /// This is the programmatic equivalent of Numbers' **Organize → Sort →
    /// Sort Now** action. It moves only body rows, leaving configured header
    /// and footer rows in place, and retains the native sort configuration for
    /// subsequent use in Numbers.
    ///
    /// The current executor deliberately supports the scalar, non-formula
    /// subset that can be moved without rewriting a formula graph: every sort
    /// key in the body must be a complete plain Text, finite Number, Boolean,
    /// Date, or Duration column of one consistent type. Cell comment threads
    /// move with their rows. It rejects formula and error body cells, merged
    /// cells, filters, grouping, pivots, spill state, and conditional styles
    /// transactionally rather than risking a semantically partial rewrite.
    /// Explicit cell borders and comment threads move with their rows, while
    /// user-hidden row and column positions remain fixed to match native iWork.
    ///
    /// Returns `true` when one or more body rows were physically reordered,
    /// and `false` when the body was already in the requested stable order.
    pub fn apply_table_sort_order(&mut self, selector: TableSelector) -> Result<bool> {
        let table_id = super::selectors::table_id(self, selector)?;
        let order = table_sort_order_in_package(&self.package, table_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Numbers sort without a configured table sort order".to_owned(),
            )
        })?;
        if order.scope() != Scope::EntireTable {
            return Err(Error::ParseError(
                "Cannot execute a selected-row Numbers sort without an explicit row range; use apply_table_sort_order_to_rows"
                    .to_owned(),
            ));
        }
        let mut staged = self.package.clone();
        if !apply_table_sort_order_in_package(&mut staged, table_id, &order)? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(selector)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Numbers table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        self.package = staged;
        Ok(true)
    }

    /// Execute a configured selected-row sort over one explicit body-row range.
    ///
    /// The range is body-relative and half-open, so header and footer rows
    /// cannot be included. This is the deterministic programmatic equivalent
    /// of selecting rows in Numbers, choosing **Sort Selected Rows**, and
    /// applying the stored rules. The range is supplied here because Numbers
    /// keeps it in view state rather than in the table sort archive.
    ///
    /// Returns `true` when one or more selected rows moved and `false` for a
    /// one-row or already stable selection.
    pub fn apply_table_sort_order_to_rows(
        &mut self,
        selector: TableSelector,
        rows: RowRange,
    ) -> Result<bool> {
        let table_id = super::selectors::table_id(self, selector)?;
        let order = table_sort_order_in_package(&self.package, table_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Numbers sort without a configured table sort order".to_owned(),
            )
        })?;
        if order.scope() != Scope::SelectedRows {
            return Err(Error::ParseError(
                "Cannot execute an entire-table Numbers sort through a selected-row range"
                    .to_owned(),
            ));
        }
        let mut staged = self.package.clone();
        if !apply_table_sort_order_to_rows_in_package(&mut staged, table_id, &order, rows)? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(selector)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Numbers selected-row sort execution did not preserve its sort order".to_owned(),
            ));
        }
        self.package = staged;
        Ok(true)
    }
}

/// Read an attached native iWork table's persisted sort-rule configuration.
pub(crate) fn table_sort_order_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Option<Order>> {
    read_attached_table_sort_order(package, table_id)
}

/// Set an attached native iWork table's persisted sort-rule configuration.
pub(crate) fn set_table_sort_order_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
) -> Result<()> {
    set_attached_table_sort_order(package, table_id, order)
}

/// Clear an attached native iWork table's stored sort rules.
///
/// Returns whether a non-empty native order was cleared.
pub(crate) fn clear_table_sort_order_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
) -> Result<bool> {
    clear_attached_table_sort_order(package, table_id)
}

/// Execute a validated full-table sort on an attached native iWork table.
///
/// The caller supplies the configuration it has already read from or assigned
/// to the table so presentation-specific editors can preserve it while they
/// validate their own ownership graph.
pub(crate) fn apply_table_sort_order_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
) -> Result<bool> {
    apply_attached_table_sort_order(package, table_id, order)
}

/// Execute a validated selected-row sort on an attached native iWork table.
pub(crate) fn apply_table_sort_order_to_rows_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
    rows: RowRange,
) -> Result<bool> {
    apply::apply_attached_table_sort_order_to_rows(package, table_id, order, rows)
}

pub(super) fn read_attached_table_sort_order(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Option<Order>> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let native = read_native_table_sort_order(package, &descriptor)?;
    native
        .as_ref()
        .map(order_from_native)
        .transpose()
        .map(Option::flatten)
}

fn set_attached_table_sort_order(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_sort_order(&descriptor.model, order)?;
    let current = read_native_table_sort_order(package, &descriptor)?;
    if current
        .as_ref()
        .map(order_from_native)
        .transpose()?
        .flatten()
        .as_ref()
        == Some(order)
    {
        return Ok(());
    }
    update_table_sort_order(package, table_id, |original, model| {
        write_table_sort_order_wire(original, model, order)
    })
}

fn clear_attached_table_sort_order(package: &mut IWorkPackage, table_id: u64) -> Result<bool> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let Some(native) = read_native_table_sort_order(package, &descriptor)? else {
        return Ok(false);
    };
    if native.rules.is_empty() {
        return Ok(false);
    }
    update_table_sort_order(package, table_id, |original, model| {
        clear_table_sort_order_wire(original, model)
    })?;
    Ok(true)
}

fn read_native_table_sort_order(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<Option<tst::TableSortOrderArchive>> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let message_index = find_table_model_message(object)?;
    read_native_table_sort_order_wire(
        object.messages[message_index].data.as_slice(),
        &descriptor.model,
    )
}

fn update_table_sort_order<F>(package: &mut IWorkPackage, table_id: u64, update: F) -> Result<()>
where
    F: FnOnce(&[u8], &TableModelArchive) -> Result<Vec<u8>>,
{
    let locations = object_locations(package)?;
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
        })?;
        let message_index = find_table_model_message(object)?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let model = TableModelArchive::decode(original)?;
        let data = update(original, &model)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn validate_sort_order(model: &TableModelArchive, order: &Order) -> Result<()> {
    let columns = model.number_of_columns as usize;
    for rule in order.rules() {
        if rule.column().get() >= columns {
            return Err(Error::ParseError(format!(
                "Numbers table sort column {} is outside the table's {columns} columns",
                rule.column().get()
            )));
        }
    }
    Ok(())
}

/// Validate that a table has either no sort order or a supported full-table order.
pub(super) fn validate_table_sort_order_for_topology(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let Some(native) = read_native_table_sort_order(package, &descriptor)? else {
        return Ok(());
    };
    let Some(order) = order_from_native(&native)? else {
        return Ok(());
    };
    if order.scope() == Scope::SelectedRows {
        return Err(Error::ParseError(
            "Cannot yet edit table topology while a selected-row sort order is configured"
                .to_owned(),
        ));
    }
    validate_sort_order(&descriptor.model, &order)
}

/// Remove sort rules whose physical slot disappears with a column deletion.
///
/// Numbers keeps every other rule index unchanged, including rules after a
/// deleted earlier column. A rule therefore belongs to its physical slot
/// rather than following the cells that shift through that slot.
pub(super) fn delete_table_sort_column(
    package: &mut IWorkPackage,
    table_id: u64,
    column: usize,
    new_columns: usize,
) -> Result<()> {
    validate_table_sort_order_for_topology(package, table_id)?;
    let column = u32::try_from(column)
        .map_err(|_| Error::ParseError("Numbers deleted sort column exceeds u32".to_owned()))?;
    let new_columns = u32::try_from(new_columns)
        .map_err(|_| Error::ParseError("Numbers table column count exceeds u32".to_owned()))?;
    let descriptor = attached_table_descriptor(package, table_id)?;
    let Some(native) = read_native_table_sort_order(package, &descriptor)? else {
        return Ok(());
    };
    if !native
        .rules
        .iter()
        .any(|rule| rule.index == column || rule.index >= new_columns)
    {
        return Ok(());
    }
    update_table_sort_order(package, table_id, |original, model| {
        delete_table_sort_column_wire(original, model, column, new_columns)
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use litchi_numbers::table::topology::{
        ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
    };

    #[test]
    fn sort_column_index_rejects_values_outside_native_range() {
        assert_eq!(ColumnIndex::new(0).unwrap().get(), 0);
        if let Ok(too_large) = usize::try_from(u64::from(u32::MAX) + 1) {
            assert!(ColumnIndex::new(too_large).is_err());
        }
    }

    #[test]
    fn sort_order_requires_non_empty_unique_rules() {
        assert!(Order::new([]).is_err());
        let column = ColumnIndex::new(1).unwrap();
        let duplicate = Order::new([
            Rule::new(column, Direction::Ascending),
            Rule::new(column, Direction::Descending),
        ]);
        assert!(duplicate.is_err());
    }

    #[test]
    fn sort_scope_and_selected_row_range_are_strict_typed() {
        let rule = Rule::new(ColumnIndex::new(0).unwrap(), Direction::Ascending);
        let entire = Order::new([rule]).unwrap();
        assert_eq!(entire.scope(), Scope::EntireTable);
        let selected = Order::selected_rows([rule]).unwrap();
        assert_eq!(selected.scope(), Scope::SelectedRows);

        assert!(RowRange::new(0, 0).is_err());
        assert!(RowRange::new(2, 1).is_err());
        let range = RowRange::new(2, 5).unwrap();
        assert_eq!(range.start(), 2);
        assert_eq!(range.end(), 5);
        assert_eq!(range.len(), 3);
        assert!(!range.is_empty());
    }

    #[test]
    fn table_sort_selector_resolves_by_name_and_catalog_index() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(0).unwrap(),
            NumbersTableSortDirection::Ascending,
        )])
        .unwrap();

        editor
            .set_table_sort_order(TableSelector::name("Revenue"), order.clone())
            .unwrap();
        assert_eq!(
            editor.table_sort_order(TableSelector::index(0)).unwrap(),
            Some(order)
        );
        assert!(editor.table_sort_order(TableSelector::index(1)).is_err());
        assert!(
            editor
                .table_sort_order(TableSelector::name("Missing"))
                .is_err()
        );
    }

    #[test]
    fn full_table_sort_rules_survive_native_topology_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 4)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let selector = TableSelector::name("Table 1");
        let initial = Order::new([
            Rule::new(ColumnIndex::new(1).unwrap(), Direction::Ascending),
            Rule::new(ColumnIndex::new(3).unwrap(), Direction::Descending),
        ])
        .unwrap();
        editor
            .set_table_sort_order(selector, initial.clone())
            .unwrap();

        editor
            .insert_table_row(
                test_table_selector(&editor, table_id),
                RowInsertion::body(0),
            )
            .unwrap();
        editor
            .insert_table_column(
                test_table_selector(&editor, table_id),
                ColumnInsertion::body(0),
            )
            .unwrap();
        assert_eq!(editor.table_sort_order(selector).unwrap(), Some(initial));

        editor
            .remove_table_row(test_table_selector(&editor, table_id), RowDeletion::body(0))
            .unwrap();
        editor
            .remove_table_column(
                test_table_selector(&editor, table_id),
                ColumnDeletion::body(0),
            )
            .unwrap();
        let remaining = Order::new([Rule::new(
            ColumnIndex::new(3).unwrap(),
            Direction::Descending,
        )])
        .unwrap();
        assert_eq!(editor.table_sort_order(selector).unwrap(), Some(remaining));

        editor
            .remove_table_column(
                test_table_selector(&editor, table_id),
                ColumnDeletion::body(2),
            )
            .unwrap();
        assert_eq!(editor.table_sort_order(selector).unwrap(), None);
    }
}
