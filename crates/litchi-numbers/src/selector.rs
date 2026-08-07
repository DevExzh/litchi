//! Human-readable and checked positional selectors for Numbers objects.

/// Selects one sheet by its exact visible name or checked zero-based catalog
/// position without allocating for the selector itself.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector names intentionally identify their selected Numbers object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SheetSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in the editor's sheet catalog.
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
}

impl<'a> From<&'a str> for SheetSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

/// Selects one table by its exact visible name or checked zero-based catalog
/// position without allocating for the selector itself.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector names intentionally identify their selected Numbers object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableSelector<'a> {
    /// Select by the exact name shown by Numbers.
    Name(&'a str),
    /// Select by zero-based position in the editor's table catalog.
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
}

impl<'a> From<&'a str> for TableSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
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
    fn borrowed_names_convert_without_allocating() {
        let sheet: SheetSelector<'_> = "Summary".into();
        let table: TableSelector<'_> = "Revenue".into();
        assert_eq!(sheet, SheetSelector::Name("Summary"));
        assert_eq!(table, TableSelector::Name("Revenue"));
    }
}
