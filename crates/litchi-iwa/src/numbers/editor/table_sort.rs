//! Typed sort-rule configuration and execution for Numbers tables.

use std::collections::BTreeSet;

use super::*;

mod apply;
mod wire;

use apply::apply_attached_table_sort_order;
use wire::{
    clear_table_sort_order_wire, delete_table_sort_column_wire, read_native_table_sort_order_wire,
    write_table_sort_order_wire,
};

const NATIVE_ENTIRE_TABLE_SORT: i32 = tst::table_sort_order_archive::SortType::EntireTable as i32;
const NATIVE_ROW_RANGE_SORT: i32 = tst::table_sort_order_archive::SortType::RowRange as i32;

/// Rows targeted by a persisted Numbers table sort configuration.
///
/// Numbers stores this scope with the rules. For [`Self::SelectedRows`], the
/// actual selected row interval belongs to document view state and is
/// therefore supplied explicitly when the sort is executed.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum NumbersTableSortScope {
    /// Apply the rules to every body row, excluding headers and footers.
    #[default]
    EntireTable,
    /// Apply the rules to the rows selected in the document view.
    SelectedRows,
}

impl NumbersTableSortScope {
    const fn as_native(self) -> i32 {
        match self {
            Self::EntireTable => NATIVE_ENTIRE_TABLE_SORT,
            Self::SelectedRows => NATIVE_ROW_RANGE_SORT,
        }
    }

    fn from_native(value: i32) -> Result<Self> {
        match tst::table_sort_order_archive::SortType::try_from(value) {
            Ok(tst::table_sort_order_archive::SortType::EntireTable) => Ok(Self::EntireTable),
            Ok(tst::table_sort_order_archive::SortType::RowRange) => Ok(Self::SelectedRows),
            Err(_) => Err(Error::InvalidFormat(format!(
                "Numbers table sort order has unknown scope {value}"
            ))),
        }
    }
}

/// A non-empty, body-relative half-open row range for selected-row sorting.
///
/// `start` is inclusive and `end` is exclusive. Header and footer rows are
/// deliberately outside this coordinate system, so callers cannot
/// accidentally move them through a selected-row sort.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NumbersTableSortRowRange {
    start: usize,
    end: usize,
}

impl NumbersTableSortRowRange {
    /// Construct a non-empty body-relative range `[start, end)`.
    pub fn new(start: usize, end: usize) -> Result<Self> {
        if start >= end {
            return Err(Error::ParseError(
                "Numbers selected-row sort range must be non-empty".to_owned(),
            ));
        }
        Ok(Self { start, end })
    }

    /// Return the inclusive body-relative start row.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Return the exclusive body-relative end row.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// Return the number of selected body rows.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end - self.start
    }

    /// Return whether this range is empty.
    ///
    /// Construction rejects empty ranges, so this is always `false`; the
    /// method makes the type convenient in generic range-oriented code.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        false
    }
}

/// A validated zero-based physical column index used by a Numbers sort rule.
///
/// The index is checked against a table's current column count when an order
/// is assigned. Constructing this type also rejects values that do not fit
/// Numbers' native `uint32` representation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct NumbersTableSortColumnIndex(usize);

impl NumbersTableSortColumnIndex {
    /// Construct a native-compatible zero-based column index.
    pub fn new(index: usize) -> Result<Self> {
        u32::try_from(index).map_err(|_| {
            Error::ParseError(
                "Numbers table sort column index exceeds the native u32 range".to_owned(),
            )
        })?;
        Ok(Self(index))
    }

    /// Return the zero-based column index.
    #[must_use]
    pub const fn get(self) -> usize {
        self.0
    }

    pub(super) fn as_native(self) -> Result<u32> {
        u32::try_from(self.0).map_err(|_| {
            Error::InvalidFormat(
                "Numbers table sort column index exceeds the native u32 range".to_owned(),
            )
        })
    }

    pub(super) fn from_native(index: u32) -> Result<Self> {
        Self::new(usize::try_from(index).map_err(|_| {
            Error::InvalidFormat("Numbers table sort column index exceeds usize".to_owned())
        })?)
        .map_err(|_| {
            Error::InvalidFormat(
                "Numbers table sort column index exceeds the native u32 range".to_owned(),
            )
        })
    }
}

impl TryFrom<usize> for NumbersTableSortColumnIndex {
    type Error = Error;

    fn try_from(index: usize) -> Result<Self> {
        Self::new(index)
    }
}

impl From<NumbersTableSortColumnIndex> for usize {
    fn from(index: NumbersTableSortColumnIndex) -> Self {
        index.get()
    }
}

/// Sort direction for one Numbers table column.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum NumbersTableSortDirection {
    /// Sort low-to-high, alphabetically A-to-Z, or oldest-to-newest.
    Ascending,
    /// Sort high-to-low, alphabetically Z-to-A, or newest-to-oldest.
    Descending,
}

/// One full-table sort-configuration rule in priority order.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NumbersTableSortRule {
    column: NumbersTableSortColumnIndex,
    direction: NumbersTableSortDirection,
}

impl NumbersTableSortRule {
    /// Construct a rule for one physical table column.
    #[must_use]
    pub const fn new(
        column: NumbersTableSortColumnIndex,
        direction: NumbersTableSortDirection,
    ) -> Self {
        Self { column, direction }
    }

    /// Return the column selected by this rule.
    #[must_use]
    pub const fn column(self) -> NumbersTableSortColumnIndex {
        self.column
    }

    /// Return this rule's direction.
    #[must_use]
    pub const fn direction(self) -> NumbersTableSortDirection {
        self.direction
    }
}

/// An ordered, non-empty Numbers sort-rule configuration.
///
/// Rules are evaluated in slice order. Numbers does not accept the same
/// column more than once, so construction rejects duplicate columns.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NumbersTableSortOrder {
    scope: NumbersTableSortScope,
    rules: Vec<NumbersTableSortRule>,
}

impl NumbersTableSortOrder {
    /// Construct a native full-table sort-rule configuration.
    pub fn new(rules: impl IntoIterator<Item = NumbersTableSortRule>) -> Result<Self> {
        Self::with_scope(NumbersTableSortScope::EntireTable, rules)
    }

    /// Construct a native selected-row sort-rule configuration.
    ///
    /// The selected interval is not persisted in the table archive. Supply a
    /// [`NumbersTableSortRowRange`] to
    /// [`NumbersEditor::apply_table_sort_order_to_rows`] when executing it.
    pub fn selected_rows(rules: impl IntoIterator<Item = NumbersTableSortRule>) -> Result<Self> {
        Self::with_scope(NumbersTableSortScope::SelectedRows, rules)
    }

    /// Construct a native sort-rule configuration with an explicit scope.
    pub fn with_scope(
        scope: NumbersTableSortScope,
        rules: impl IntoIterator<Item = NumbersTableSortRule>,
    ) -> Result<Self> {
        let rules = rules.into_iter().collect::<Vec<_>>();
        if rules.is_empty() {
            return Err(Error::ParseError(
                "Numbers table sort order must contain at least one rule".to_owned(),
            ));
        }
        let mut columns = BTreeSet::new();
        if rules.iter().any(|rule| !columns.insert(rule.column)) {
            return Err(Error::ParseError(
                "Numbers table sort order cannot contain the same column more than once".to_owned(),
            ));
        }
        Ok(Self { scope, rules })
    }

    /// Return the persisted sort scope.
    #[must_use]
    pub const fn scope(&self) -> NumbersTableSortScope {
        self.scope
    }

    /// Return the rules in native evaluation order without cloning them.
    #[must_use]
    pub fn rules(&self) -> &[NumbersTableSortRule] {
        &self.rules
    }

    fn from_native(sort: &tst::TableSortOrderArchive) -> Result<Option<Self>> {
        if sort.rules.is_empty() {
            return Ok(None);
        }
        let scope = NumbersTableSortScope::from_native(sort.r#type)?;
        Self::with_scope(scope, sort.rules.iter().map(|rule| {
            let direction = match tst::table_sort_order_archive::sort_rule_archive::Direction::try_from(
                rule.direction,
            ) {
                Ok(tst::table_sort_order_archive::sort_rule_archive::Direction::Ascending) => {
                    NumbersTableSortDirection::Ascending
                },
                Ok(tst::table_sort_order_archive::sort_rule_archive::Direction::Descending) => {
                    NumbersTableSortDirection::Descending
                },
                Err(_) => {
                    return Err(Error::InvalidFormat(
                        "Numbers table sort rule has an unknown direction".to_owned(),
                    ));
                },
            };
            Ok(NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::from_native(rule.index)?,
                direction,
            ))
        })
        .collect::<Result<Vec<_>>>()?)
        .map(Some)
        .map_err(|error| {
            Error::InvalidFormat(format!("Numbers table has an invalid stored sort order: {error}"))
        })
    }

    fn as_native(&self) -> Result<tst::TableSortOrderArchive> {
        Ok(tst::TableSortOrderArchive {
            r#type: self.scope.as_native(),
            rules: self
                .rules
                .iter()
                .map(|rule| {
                    Ok(tst::table_sort_order_archive::SortRuleArchive {
                        index: rule.column.as_native()?,
                        direction: match rule.direction {
                            NumbersTableSortDirection::Ascending => {
                                tst::table_sort_order_archive::sort_rule_archive::Direction::Ascending
                                    as i32
                            },
                            NumbersTableSortDirection::Descending => {
                                tst::table_sort_order_archive::sort_rule_archive::Direction::Descending
                                    as i32
                            },
                        },
                    })
                })
                .collect::<Result<Vec<_>>>()?,
        })
    }
}

impl NumbersEditor {
    /// Read an attached table's persisted sort-rule configuration.
    ///
    /// An empty native order is reported as `None`, matching the state shown
    /// by Numbers after its last sort rule is removed. Selected-row orders
    /// expose their persisted [`NumbersTableSortScope::SelectedRows`] scope;
    /// their view-state selected interval is intentionally not guessed.
    pub fn table_sort_order(&self, table_id: u64) -> Result<Option<NumbersTableSortOrder>> {
        table_sort_order_in_package(&self.package, table_id)
    }

    /// Set the persisted sort-rule configuration transactionally.
    ///
    /// The resulting file stores the same table-level order exposed in
    /// Numbers' **Organize → Sort** pane, including for spreadsheets created
    /// entirely by this crate. This operation configures the native rule; it
    /// does not execute it or reorder stored rows. Numbers exposes that
    /// separate action as **Sort Now**.
    pub fn set_table_sort_order(
        &mut self,
        table_id: u64,
        order: NumbersTableSortOrder,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_table_sort_order_in_package(&mut staged, table_id, &order)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(table_id)?.as_ref() != Some(&order) {
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
    pub fn clear_table_sort_order(&mut self, table_id: u64) -> Result<()> {
        let mut staged = self.package.clone();
        if !clear_table_sort_order_in_package(&mut staged, table_id)? {
            return Ok(());
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_sort_order(table_id)?.is_some() {
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
    pub fn apply_table_sort_order(&mut self, table_id: u64) -> Result<bool> {
        let order = self.table_sort_order(table_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Numbers sort without a configured table sort order".to_owned(),
            )
        })?;
        if order.scope() != NumbersTableSortScope::EntireTable {
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
        if verified.table_sort_order(table_id)?.as_ref() != Some(&order) {
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
        table_id: u64,
        rows: NumbersTableSortRowRange,
    ) -> Result<bool> {
        let order = self.table_sort_order(table_id)?.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Numbers sort without a configured table sort order".to_owned(),
            )
        })?;
        if order.scope() != NumbersTableSortScope::SelectedRows {
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
        if verified.table_sort_order(table_id)?.as_ref() != Some(&order) {
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
) -> Result<Option<NumbersTableSortOrder>> {
    read_attached_table_sort_order(package, table_id)
}

/// Set an attached native iWork table's persisted sort-rule configuration.
pub(crate) fn set_table_sort_order_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &NumbersTableSortOrder,
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
    order: &NumbersTableSortOrder,
) -> Result<bool> {
    apply_attached_table_sort_order(package, table_id, order)
}

/// Execute a validated selected-row sort on an attached native iWork table.
pub(crate) fn apply_table_sort_order_to_rows_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &NumbersTableSortOrder,
    rows: NumbersTableSortRowRange,
) -> Result<bool> {
    apply::apply_attached_table_sort_order_to_rows(package, table_id, order, rows)
}

pub(super) fn read_attached_table_sort_order(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Option<NumbersTableSortOrder>> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let native = read_native_table_sort_order(package, &descriptor)?;
    native
        .as_ref()
        .map(NumbersTableSortOrder::from_native)
        .transpose()
        .map(Option::flatten)
}

fn set_attached_table_sort_order(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &NumbersTableSortOrder,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_sort_order(&descriptor.model, order)?;
    let current = read_native_table_sort_order(package, &descriptor)?;
    if current
        .as_ref()
        .map(NumbersTableSortOrder::from_native)
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

fn validate_sort_order(model: &TableModelArchive, order: &NumbersTableSortOrder) -> Result<()> {
    let columns = model.number_of_columns as usize;
    for rule in order.rules() {
        if rule.column.get() >= columns {
            return Err(Error::ParseError(format!(
                "Numbers table sort column {} is outside the table's {columns} columns",
                rule.column.get()
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
    let Some(order) = NumbersTableSortOrder::from_native(&native)? else {
        return Ok(());
    };
    if order.scope() == NumbersTableSortScope::SelectedRows {
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
    use crate::numbers::{
        NumbersDocumentBuilder, TableColumnDeletion, TableColumnInsertion, TableRowDeletion,
        TableRowInsertion,
    };

    #[test]
    fn sort_column_index_rejects_values_outside_native_range() {
        assert_eq!(NumbersTableSortColumnIndex::new(0).unwrap().get(), 0);
        if let Ok(too_large) = usize::try_from(u64::from(u32::MAX) + 1) {
            assert!(NumbersTableSortColumnIndex::new(too_large).is_err());
        }
    }

    #[test]
    fn sort_order_requires_non_empty_unique_rules() {
        assert!(NumbersTableSortOrder::new([]).is_err());
        let column = NumbersTableSortColumnIndex::new(1).unwrap();
        let duplicate = NumbersTableSortOrder::new([
            NumbersTableSortRule::new(column, NumbersTableSortDirection::Ascending),
            NumbersTableSortRule::new(column, NumbersTableSortDirection::Descending),
        ]);
        assert!(duplicate.is_err());
    }

    #[test]
    fn sort_scope_and_selected_row_range_are_strict_typed() {
        let rule = NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(0).unwrap(),
            NumbersTableSortDirection::Ascending,
        );
        let entire = NumbersTableSortOrder::new([rule]).unwrap();
        assert_eq!(entire.scope(), NumbersTableSortScope::EntireTable);
        let selected = NumbersTableSortOrder::selected_rows([rule]).unwrap();
        assert_eq!(selected.scope(), NumbersTableSortScope::SelectedRows);

        assert!(NumbersTableSortRowRange::new(0, 0).is_err());
        assert!(NumbersTableSortRowRange::new(2, 1).is_err());
        let range = NumbersTableSortRowRange::new(2, 5).unwrap();
        assert_eq!(range.start(), 2);
        assert_eq!(range.end(), 5);
        assert_eq!(range.len(), 3);
        assert!(!range.is_empty());
    }

    #[test]
    fn full_table_sort_rules_survive_native_topology_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 4)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let initial = NumbersTableSortOrder::new([
            NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Ascending,
            ),
            NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(3).unwrap(),
                NumbersTableSortDirection::Descending,
            ),
        ])
        .unwrap();
        editor
            .set_table_sort_order(table_id, initial.clone())
            .unwrap();

        editor
            .insert_table_row(table_id, TableRowInsertion::body(0))
            .unwrap();
        editor
            .insert_table_column(table_id, TableColumnInsertion::body(0))
            .unwrap();
        assert_eq!(editor.table_sort_order(table_id).unwrap(), Some(initial));

        editor
            .remove_table_row(table_id, TableRowDeletion::body(0))
            .unwrap();
        editor
            .remove_table_column(table_id, TableColumnDeletion::body(0))
            .unwrap();
        let remaining = NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(3).unwrap(),
            NumbersTableSortDirection::Descending,
        )])
        .unwrap();
        assert_eq!(editor.table_sort_order(table_id).unwrap(), Some(remaining));

        editor
            .remove_table_column(table_id, TableColumnDeletion::body(2))
            .unwrap();
        assert_eq!(editor.table_sort_order(table_id).unwrap(), None);
    }
}
