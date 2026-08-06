//! Archive-free semantic selectors for Keynote charts.

/// Selects one chart by its visible native title or checked zero-based position.
///
/// A selector carries no native object identifier and does not depend on a
/// package or archive representation. The concrete Keynote adapter resolves
/// the selector against the charts owned by a slide, checking the position or
/// exact name there.
#[allow(
    clippy::module_name_repetitions,
    reason = "ChartSelector keeps the chart semantic domain explicit at the crate boundary"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSelector<'a> {
    /// Select the chart at a checked zero-based position in slide chart order.
    Index(usize),
    /// Select the chart with this exact visible native title.
    Name(&'a str),
}

impl<'a> ChartSelector<'a> {
    /// Create a zero-based positional chart selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }

    /// Create an exact-name chart selector without allocating.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Return the selected zero-based position, if this is an index selector.
    #[must_use]
    pub const fn as_index(self) -> Option<usize> {
        match self {
            Self::Index(index) => Some(index),
            Self::Name(_) => None,
        }
    }

    /// Borrow the selected exact name, if this is a name selector.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Index(_) => None,
            Self::Name(name) => Some(name),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::ChartSelector;

    #[test]
    fn index_selector_preserves_checked_position() {
        let selector = ChartSelector::index(3);

        assert_eq!(selector, ChartSelector::Index(3));
        assert_eq!(selector.as_index(), Some(3));
        assert_eq!(selector.as_name(), None);
    }

    #[test]
    fn name_selector_borrows_exact_name() {
        let name = String::from("Revenue chart");
        let selector = ChartSelector::name(name.as_str());

        assert_eq!(selector, ChartSelector::Name("Revenue chart"));
        assert_eq!(selector.as_name(), Some("Revenue chart"));
        assert_eq!(selector.as_index(), None);
    }

    #[test]
    fn selectors_are_copyable_value_inputs() {
        const INDEX: ChartSelector<'static> = ChartSelector::index(0);
        const NAME: ChartSelector<'static> = ChartSelector::name("Chart");

        assert_eq!(INDEX, INDEX);
        assert_eq!(NAME, NAME);
    }
}
