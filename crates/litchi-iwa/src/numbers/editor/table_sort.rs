//! Typed sort-rule configuration and execution for Numbers tables.

use super::*;
use litchi_iwa_protos::table_sort_order_codec as codec;
use litchi_numbers::Package as FocusedNumbersPackage;
use litchi_numbers::SheetSelector;
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
use wire::{delete_table_sort_column_wire, read_native_table_sort_order_wire};

fn invalid_stored_sort(error: sort::Error) -> Error {
    Error::InvalidFormat(format!(
        "Numbers table has an invalid stored sort order: {error}"
    ))
}

fn order_from_native(sort: &codec::SortOrderSnapshot) -> Result<Option<Order>> {
    if sort.rules().is_empty() {
        return Ok(None);
    }
    let scope = match sort.scope() {
        codec::SortScope::EntireTable => Scope::EntireTable,
        codec::SortScope::SelectedRows => Scope::SelectedRows,
    };
    let mut rules = Vec::new();
    rules
        .try_reserve_exact(sort.rules().len())
        .map_err(|_allocation| {
            invalid_stored_sort(sort::Error::Allocation {
                amount: sort.rules().len(),
            })
        })?;
    for rule in sort.rules() {
        let direction = match rule.direction() {
            codec::SortDirection::Ascending => Direction::Ascending,
            codec::SortDirection::Descending => Direction::Descending,
        };
        let column = ColumnIndex::from_native(rule.column()).map_err(invalid_stored_sort)?;
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules)
        .map(Some)
        .map_err(invalid_stored_sort)
}

impl NumbersEditor {
    /// Read an attached table's persisted sort-rule configuration.
    ///
    /// An empty native order is reported as `None`, matching the state shown
    /// by Numbers after its last sort rule is removed. Selected-row orders
    /// expose their persisted [`Scope::SelectedRows`] scope;
    /// their view-state selected interval is intentionally not guessed.
    #[cfg(test)]
    pub fn table_sort_order(&self, selector: TableSelector) -> Result<Option<Order>> {
        let table_id = super::selectors::table_id(self, selector)?;
        let (source, sheet, table) = focused_table_sort_source(self, table_id)?;
        source
            .table_sort_order(sheet, table)
            .map_err(focused_sort_error)
    }

    /// Set the persisted sort-rule configuration transactionally.
    ///
    /// The resulting file stores the same table-level order exposed in
    /// Numbers' **Organize → Sort** pane, including for spreadsheets created
    /// entirely by this crate. This operation configures the native rule; it
    /// does not execute it or reorder stored rows. Numbers exposes that
    /// separate action as **Sort Now**.
    #[cfg(test)]
    pub fn set_table_sort_order(&mut self, selector: TableSelector, order: Order) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let (source, sheet, table) = focused_table_sort_source(self, table_id)?;
        let expected = order.clone();
        let commit = source
            .edit_table_sort_order(sheet, table)
            .map_err(focused_sort_error)?
            .set(order)
            .commit()
            .map_err(focused_sort_error)?;
        let verified = focused_sort_commit_editor(&commit)?;
        if verified.table_sort_order(selector)?.as_ref() != Some(&expected) {
            return Err(Error::InvalidFormat(
                "focused Numbers table sort order failed round-trip validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }

    /// Clear an attached table's stored sort rules transactionally.
    ///
    /// When a table already carries native sort metadata, this preserves
    /// Numbers' empty-order marker and any associated reference tracker,
    /// exactly as removing the final rule in the Numbers UI does.
    #[cfg(test)]
    pub fn clear_table_sort_order(&mut self, selector: TableSelector) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let (source, sheet, table) = focused_table_sort_source(self, table_id)?;
        let commit = source
            .edit_table_sort_order(sheet, table)
            .map_err(focused_sort_error)?
            .clear()
            .commit()
            .map_err(focused_sort_error)?;
        let verified = focused_sort_commit_editor(&commit)?;
        if verified.table_sort_order(selector)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table sort-order clear failed round-trip validation".to_owned(),
            ));
        }
        self.package = verified.package;
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
        let order = focused_table_sort_order_for_apply(self, table_id)?.ok_or_else(|| {
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
        if table_sort_order_in_package(&verified.package, table_id)?.as_ref() != Some(&order) {
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
        let order = focused_table_sort_order_for_apply(self, table_id)?.ok_or_else(|| {
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
        if table_sort_order_in_package(&verified.package, table_id)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Numbers selected-row sort execution did not preserve its sort order".to_owned(),
            ));
        }
        self.package = staged;
        Ok(true)
    }
}

fn focused_sort_error(error: litchi_numbers::table::sort::transaction::Error) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers persisted sort operation failed: {error}"
    ))
}

#[cfg(test)]
fn focused_table_sort_source(
    editor: &NumbersEditor,
    table_id: u64,
) -> Result<(
    FocusedNumbersPackage,
    SheetSelector<'static>,
    TableSelector<'static>,
)> {
    let (sheet, table) = super::selectors::focused_table_location(editor, table_id)?;
    let bytes = editor.to_bytes()?;
    let source = FocusedNumbersPackage::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers persisted-sort source validation failed: {error}"
        ))
    })?;
    Ok((source, sheet, table))
}

fn focused_table_sort_order_for_apply(
    editor: &NumbersEditor,
    table_id: u64,
) -> Result<Option<Order>> {
    let source_built = !editor.package.source_is_exact();
    let (sheet, table) = super::selectors::focused_table_location(editor, table_id)?;
    let bytes = editor.to_bytes()?;
    match FocusedNumbersPackage::from_bytes(&bytes) {
        Ok(source) => match (FocusedNumbersPackage::table_sort_order)(&source, sheet, table) {
            Ok(order) => Ok(order),
            Err(strict_error) if source_built => {
                FocusedNumbersPackage::__table_sort_order_from_bytes_for_compatibility(
                    &bytes, sheet, table,
                )
                .map_err(|_| focused_sort_error(strict_error))
            },
            Err(strict_error) => Err(focused_sort_error(strict_error)),
        },
        Err(strict_error) if source_built => {
            FocusedNumbersPackage::__table_sort_order_from_bytes_for_compatibility(
                &bytes, sheet, table,
            )
            .map_err(|_| {
                Error::InvalidFormat(format!(
                    "focused Numbers persisted-sort source validation failed: {strict_error}"
                ))
            })
        },
        Err(strict_error) => Err(Error::InvalidFormat(format!(
            "focused Numbers persisted-sort source validation failed: {strict_error}"
        ))),
    }
}

#[cfg(test)]
fn focused_sort_commit_editor(
    commit: &litchi_numbers::table::sort::transaction::Commit,
) -> Result<NumbersEditor> {
    let mut bytes = Vec::new();
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

/// Read an attached native iWork table's persisted sort-rule configuration.
pub(crate) fn table_sort_order_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Option<Order>> {
    read_attached_table_sort_order(package, table_id)
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

fn read_native_table_sort_order(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<Option<codec::SortOrderSnapshot>> {
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
        .rules()
        .iter()
        .any(|rule| rule.column() == column || rule.column() >= new_columns)
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
