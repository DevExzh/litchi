//! Human-readable and checked positional selectors for Numbers objects.

use litchi_core::Position;

/// Selects one sheet by its exact visible name or checked zero-based document
/// position without allocating for the selector itself.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector names intentionally identify their selected Numbers object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SheetSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in stable document order.
    Position(Position),
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
        Self::Position(Position::new(index))
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Borrows the selected exact name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Position(_) => None,
        }
    }

    /// Returns the selected typed zero-based collection position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Name(_) => None,
            Self::Position(position) => Some(position),
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

impl From<Position> for SheetSelector<'_> {
    fn from(position: Position) -> Self {
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
    Position(Position),
}

impl ChartSelector {
    /// Creates a checked zero-based chart selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Position(Position::new(index))
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Returns the selected typed zero-based collection position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Position(position) => position,
        }
    }
}

impl From<usize> for ChartSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for ChartSelector {
    fn from(position: Position) -> Self {
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
#[non_exhaustive]
pub enum TableSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in stable sheet table order.
    Position(Position),
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
        Self::Position(Position::new(index))
    }

    /// Creates a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Borrows the selected exact name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Position(_) => None,
        }
    }

    /// Returns the selected typed zero-based collection position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Name(_) => None,
            Self::Position(position) => Some(position),
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

impl From<Position> for TableSelector<'_> {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

/// An owned sheet selector retained by a deferred edit entry.
///
/// Name selectors are copied once at the API boundary so the edit can defer
/// package-owned catalog resolution without borrowing a caller-owned string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum OwnedSheetSelector {
    Name(Box<str>),
    Position(Position),
}

impl OwnedSheetSelector {
    pub(crate) fn name_bytes(selector: SheetSelector<'_>) -> Option<usize> {
        match selector {
            SheetSelector::Name(name) => Some(name.len()),
            SheetSelector::Position(_) => None,
        }
    }

    pub(crate) fn try_from(
        selector: SheetSelector<'_>,
        maximum_name_bytes: usize,
    ) -> Result<Self, usize> {
        match selector {
            SheetSelector::Name(name) => {
                if name.len() > maximum_name_bytes {
                    return Err(name.len());
                }
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|_| name.len())?;
                owned.push_str(name);
                Ok(Self::Name(owned.into_boxed_str()))
            },
            SheetSelector::Position(position) => Ok(Self::Position(position)),
        }
    }

    pub(crate) fn as_selector(&self) -> SheetSelector<'_> {
        match self {
            Self::Name(name) => SheetSelector::Name(name.as_ref()),
            Self::Position(position) => SheetSelector::Position(*position),
        }
    }
}

/// An owned table selector retained by a deferred edit entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum OwnedTableSelector {
    Name(Box<str>),
    Position(Position),
}

impl OwnedTableSelector {
    pub(crate) fn name_bytes(selector: TableSelector<'_>) -> Option<usize> {
        match selector {
            TableSelector::Name(name) => Some(name.len()),
            TableSelector::Position(_) => None,
        }
    }

    pub(crate) fn try_from(
        selector: TableSelector<'_>,
        maximum_name_bytes: usize,
    ) -> Result<Self, usize> {
        match selector {
            TableSelector::Name(name) => {
                if name.len() > maximum_name_bytes {
                    return Err(name.len());
                }
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|_| name.len())?;
                owned.push_str(name);
                Ok(Self::Name(owned.into_boxed_str()))
            },
            TableSelector::Position(position) => Ok(Self::Position(position)),
        }
    }

    pub(crate) fn as_selector(&self) -> TableSelector<'_> {
        match self {
            Self::Name(name) => TableSelector::Name(name.as_ref()),
            Self::Position(position) => TableSelector::Position(*position),
        }
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
        assert_eq!(
            SheetSelector::index(1),
            SheetSelector::Position(Position::new(1))
        );
        assert_eq!(
            ChartSelector::index(2),
            ChartSelector::Position(Position::new(2))
        );
    }

    #[test]
    fn table_selectors_keep_name_lookup_primary_and_index_lookup_typed() {
        assert_eq!(
            TableSelector::name("Revenue"),
            TableSelector::Name("Revenue")
        );
        assert_eq!(
            TableSelector::index(2),
            TableSelector::Position(Position::new(2))
        );
    }

    #[test]
    fn borrowed_names_and_indexes_convert_without_allocating() {
        let sheet: SheetSelector<'_> = "Summary".into();
        let sheet_index: SheetSelector<'_> = 1.into();
        let table: TableSelector<'_> = "Revenue".into();
        let table_index: TableSelector<'_> = 2.into();
        assert_eq!(sheet, SheetSelector::Name("Summary"));
        assert_eq!(sheet_index, SheetSelector::Position(Position::new(1)));
        assert_eq!(table, TableSelector::Name("Revenue"));
        assert_eq!(table_index, TableSelector::Position(Position::new(2)));
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
        let position = Position::new(3);
        let sheet = SheetSelector::position(position);
        let table = TableSelector::position(position);

        assert_eq!(sheet, SheetSelector::Position(Position::new(3)));
        assert_eq!(sheet.as_position(), Some(position));
        assert_eq!(sheet.as_name(), None);
        assert_eq!(table, TableSelector::Position(Position::new(3)));
        assert_eq!(table.as_position(), Some(position));
        assert_eq!(table.as_name(), None);
    }

    #[test]
    fn names_expose_only_the_name_accessor() {
        let sheet = SheetSelector::name("Summary");
        let table = TableSelector::name("Revenue");

        assert_eq!(sheet.as_name(), Some("Summary"));
        assert_eq!(sheet.as_position(), None);
        assert_eq!(table.as_name(), Some("Revenue"));
        assert_eq!(table.as_position(), None);
    }

    #[test]
    fn core_positions_convert_without_changing_selector_variants() {
        let position = Position::new(7);
        let sheet: SheetSelector<'_> = position.into();
        let table: TableSelector<'_> = position.into();
        let chart: ChartSelector = position.into();

        assert_eq!(sheet, SheetSelector::Position(Position::new(7)));
        assert_eq!(table, TableSelector::Position(Position::new(7)));
        assert_eq!(chart, ChartSelector::Position(Position::new(7)));
    }

    #[test]
    fn owned_names_require_admitted_bytes_before_copying() {
        let sheet = SheetSelector::name("Summary");
        let table = TableSelector::name("Revenue");
        assert_eq!(OwnedSheetSelector::name_bytes(sheet), Some(7));
        assert_eq!(OwnedTableSelector::name_bytes(table), Some(7));
        assert_eq!(OwnedSheetSelector::try_from(sheet, 6), Err("Summary".len()));
        assert_eq!(OwnedTableSelector::try_from(table, 6), Err("Revenue".len()));
        assert!(matches!(
            OwnedSheetSelector::try_from(sheet, 7),
            Ok(OwnedSheetSelector::Name(_))
        ));
        assert!(matches!(
            OwnedTableSelector::try_from(table, 7),
            Ok(OwnedTableSelector::Name(_))
        ));
    }
}
