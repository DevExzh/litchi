//! Human-readable and checked positional selectors for Numbers objects.

/// Selects one sheet by its exact visible name or checked zero-based document
/// position without allocating for the selector itself.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector names intentionally identify their selected Numbers object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SheetSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in stable document order.
    Index(usize),
}

impl<'a> SheetSelector<'a> {
    /// Creates a name-first sheet selector.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Creates a checked zero-based sheet selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: litchi_core::Position) -> Self {
        Self::index(position.get())
    }

    /// Borrows the selected exact name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Index(_) => None,
        }
    }

    /// Returns the selected zero-based index, if present.
    #[must_use]
    pub const fn as_index(self) -> Option<usize> {
        match self {
            Self::Name(_) => None,
            Self::Index(index) => Some(index),
        }
    }

    /// Returns the selected typed zero-based collection position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<litchi_core::Position> {
        match self {
            Self::Name(_) => None,
            Self::Index(index) => Some(litchi_core::Position::new(index)),
        }
    }
}

impl<'a> From<&'a str> for SheetSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl<'a> From<&'a String> for SheetSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::name(name)
    }
}

impl From<usize> for SheetSelector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<litchi_core::Position> for SheetSelector<'_> {
    fn from(position: litchi_core::Position) -> Self {
        Self::position(position)
    }
}

/// Selects one chart by its checked zero-based position within a sheet's
/// source-order chart sequence.
///
/// Numbers does not currently expose a stable archive-free chart-name
/// catalog. Keeping this selector positional prevents native drawable IDs,
/// archive names, and generated payloads from leaking through the semantic
/// package boundary.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector name intentionally identifies the selected Numbers object."
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSelector {
    /// Select by zero-based source-order position among ordinary charts.
    Index(usize),
}

impl ChartSelector {
    /// Creates a checked zero-based chart selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: litchi_core::Position) -> Self {
        Self::index(position.get())
    }

    /// Returns the selected zero-based index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        match self {
            Self::Index(index) => index,
        }
    }

    /// Returns the selected typed zero-based collection position.
    #[must_use]
    pub const fn as_position(self) -> litchi_core::Position {
        litchi_core::Position::new(self.as_index())
    }
}

impl From<usize> for ChartSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<litchi_core::Position> for ChartSelector {
    fn from(position: litchi_core::Position) -> Self {
        Self::position(position)
    }
}

/// Selects one table by its exact visible name or checked zero-based position
/// within a sheet without allocating for the selector itself.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector names intentionally identify their selected Numbers object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in stable sheet table order.
    Index(usize),
}

impl<'a> TableSelector<'a> {
    /// Creates a name-first table selector.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Creates a checked zero-based table selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: litchi_core::Position) -> Self {
        Self::index(position.get())
    }

    /// Borrows the selected exact name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Index(_) => None,
        }
    }

    /// Returns the selected zero-based index, if present.
    #[must_use]
    pub const fn as_index(self) -> Option<usize> {
        match self {
            Self::Name(_) => None,
            Self::Index(index) => Some(index),
        }
    }

    /// Returns the selected typed zero-based collection position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<litchi_core::Position> {
        match self {
            Self::Name(_) => None,
            Self::Index(index) => Some(litchi_core::Position::new(index)),
        }
    }
}

impl<'a> From<&'a str> for TableSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl<'a> From<&'a String> for TableSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::name(name)
    }
}

impl From<usize> for TableSelector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<litchi_core::Position> for TableSelector<'_> {
    fn from(position: litchi_core::Position) -> Self {
        Self::position(position)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selectors_keep_sheet_and_table_lookup_typed() {
        assert_eq!(
            SheetSelector::name("Summary"),
            SheetSelector::Name("Summary")
        );
        assert_eq!(SheetSelector::index(1), SheetSelector::Index(1));
        assert_eq!(ChartSelector::index(2), ChartSelector::Index(2));
    }

    #[test]
    fn table_selectors_keep_name_lookup_primary_and_index_lookup_typed() {
        assert_eq!(
            TableSelector::name("Revenue"),
            TableSelector::Name("Revenue")
        );
        assert_eq!(TableSelector::index(2), TableSelector::Index(2));
    }

    #[test]
    fn borrowed_names_and_indexes_convert_without_allocating() {
        let sheet: SheetSelector<'_> = "Summary".into();
        let sheet_index: SheetSelector<'_> = 1.into();
        let table: TableSelector<'_> = "Revenue".into();
        let table_index: TableSelector<'_> = 2.into();
        assert_eq!(sheet, SheetSelector::Name("Summary"));
        assert_eq!(sheet_index, SheetSelector::Index(1));
        assert_eq!(table, TableSelector::Name("Revenue"));
        assert_eq!(table_index, TableSelector::Index(2));
    }

    #[test]
    fn borrowed_owned_names_convert_without_allocating() {
        let sheet_name = String::from("Summary");
        let table_name = String::from("Revenue");
        let sheet: SheetSelector<'_> = (&sheet_name).into();
        let table: TableSelector<'_> = (&table_name).into();

        assert_eq!(sheet, SheetSelector::Name("Summary"));
        assert_eq!(table, TableSelector::Name("Revenue"));
    }

    #[test]
    fn selectors_expose_typed_position_and_variant_accessors() {
        let position = litchi_core::Position::new(3);
        let sheet = SheetSelector::position(position);
        let table = TableSelector::position(position);

        assert_eq!(sheet, SheetSelector::Index(3));
        assert_eq!(sheet.as_index(), Some(3));
        assert_eq!(sheet.as_position(), Some(position));
        assert_eq!(sheet.as_name(), None);
        assert_eq!(table, TableSelector::Index(3));
        assert_eq!(table.as_index(), Some(3));
        assert_eq!(table.as_position(), Some(position));
        assert_eq!(table.as_name(), None);
    }

    #[test]
    fn names_expose_only_the_name_accessor() {
        let sheet = SheetSelector::name("Summary");
        let table = TableSelector::name("Revenue");

        assert_eq!(sheet.as_name(), Some("Summary"));
        assert_eq!(sheet.as_index(), None);
        assert_eq!(sheet.as_position(), None);
        assert_eq!(table.as_name(), Some("Revenue"));
        assert_eq!(table.as_index(), None);
        assert_eq!(table.as_position(), None);
    }

    #[test]
    fn core_positions_convert_without_changing_selector_variants() {
        let position = litchi_core::Position::new(7);
        let sheet: SheetSelector<'_> = position.into();
        let table: TableSelector<'_> = position.into();
        let chart: ChartSelector = position.into();

        assert_eq!(sheet, SheetSelector::Index(7));
        assert_eq!(table, TableSelector::Index(7));
        assert_eq!(chart, ChartSelector::Index(7));
    }
}
