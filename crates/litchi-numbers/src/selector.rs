//! Human-readable and checked positional selectors for Numbers objects.

/// Selects one table by its exact visible name or checked zero-based catalog
/// position without allocating for the selector itself.
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
    fn selectors_keep_name_lookup_primary_and_index_lookup_typed() {
        assert_eq!(
            TableSelector::name("Revenue"),
            TableSelector::Name("Revenue")
        );
        assert_eq!(TableSelector::index(2), TableSelector::Index(2));
    }
}
