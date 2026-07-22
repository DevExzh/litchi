//! Typed full-table sort-rule configuration for Numbers tables.

use std::collections::BTreeSet;

use super::*;

mod wire;

use wire::{
    clear_table_sort_order_wire, read_native_table_sort_order_wire, write_table_sort_order_wire,
};

const NATIVE_ENTIRE_TABLE_SORT: i32 = tst::table_sort_order_archive::SortType::EntireTable as i32;

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

/// An ordered, non-empty full-table Numbers sort-rule configuration.
///
/// Rules are evaluated in slice order. Numbers does not accept the same
/// column more than once, so construction rejects duplicate columns.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NumbersTableSortOrder {
    rules: Vec<NumbersTableSortRule>,
}

impl NumbersTableSortOrder {
    /// Construct a native full-table sort-rule configuration.
    pub fn new(rules: impl IntoIterator<Item = NumbersTableSortRule>) -> Result<Self> {
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
        Ok(Self { rules })
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
        if sort.r#type != NATIVE_ENTIRE_TABLE_SORT {
            return Err(Error::InvalidFormat(
                "Numbers row-range sort orders are not yet supported".to_owned(),
            ));
        }
        Self::new(sort.rules.iter().map(|rule| {
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
            r#type: NATIVE_ENTIRE_TABLE_SORT,
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
    /// Read an attached table's full-table sort-rule configuration.
    ///
    /// An empty native order is reported as `None`, matching the state shown
    /// by Numbers after its last sort rule is removed. Row-range sorts require
    /// transient native selection state and are rejected rather than guessed.
    pub fn table_sort_order(&self, table_id: u64) -> Result<Option<NumbersTableSortOrder>> {
        read_attached_table_sort_order(&self.package, table_id)
    }

    /// Set the full-table sort-rule configuration transactionally.
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
        set_attached_table_sort_order(&mut staged, table_id, &order)?;
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
        if !clear_attached_table_sort_order(&mut staged, table_id)? {
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
