//! Archive-free semantic selectors and catalogs for Keynote charts.
//!
//! The types in this module intentionally contain only chart order and the
//! visible native title. An adapter may use [`ChartCatalog`] to resolve a
//! [`ChartSelector`] before it enters its native graph, but native object IDs,
//! archive names, and protobuf values do not cross this boundary.

#![allow(
    clippy::module_name_repetitions,
    reason = "chart semantic types keep their domain explicit at the crate boundary"
)]

use std::collections::TryReserveError;

use litchi_core::Position;

pub use litchi_iwa_common::chart::axis::Axis;

/// Selects one chart by its visible native title or checked zero-based position.
///
/// A selector carries no native object identifier and does not depend on a
/// package or archive representation. The concrete Keynote adapter resolves
/// the selector against the charts owned by a slide, checking the position or
/// exact name there.
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

    /// Create a selector from a typed zero-based position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::index(position.get())
    }

    /// Create an exact-name chart selector without allocating.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Create an exact-name selector after checking that the name is usable.
    ///
    /// This checked constructor is useful at input boundaries where an empty
    /// visible chart title should be rejected before a chart catalog or native
    /// adapter is consulted. [`Self::name`] remains available for callers that
    /// intentionally preserve an empty title as a positional-only selector.
    ///
    /// # Errors
    ///
    /// Returns [`ChartSelectorError::EmptyName`] when `name` is empty.
    pub const fn try_name(name: &'a str) -> Result<Self, ChartSelectorError> {
        if name.is_empty() {
            Err(ChartSelectorError::EmptyName)
        } else {
            Ok(Self::Name(name))
        }
    }

    /// Return the selected zero-based position, if this is an index selector.
    #[must_use]
    pub const fn as_index(self) -> Option<usize> {
        match self {
            Self::Index(index) => Some(index),
            Self::Name(_) => None,
        }
    }

    /// Return the selected typed zero-based position, if this is an index selector.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Index(index) => Some(Position::new(index)),
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

impl<'a> From<&'a str> for ChartSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl<'a> From<&'a String> for ChartSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::name(name)
    }
}

impl From<usize> for ChartSelector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for ChartSelector<'_> {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

/// A semantic chart summary in one slide's stable source order.
///
/// The summary deliberately has no native identity. Its position is only
/// meaningful within the [`ChartCatalog`] that produced it, and its title is
/// the visible native title rather than a generated object or component name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartDescriptor {
    position: usize,
    title: Option<Box<str>>,
}

impl ChartDescriptor {
    /// Return this chart's zero-based position in its catalog.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Return the optional visible native title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Return a selector local to this catalog entry.
    ///
    /// A non-empty title is preferred for readability. Callers that need a
    /// selector immune to duplicate or later-renamed titles should use
    /// [`Self::position_selector`] instead.
    #[must_use]
    pub fn selector(&self) -> ChartSelector<'_> {
        self.title
            .as_deref()
            .filter(|title| !title.is_empty())
            .map_or_else(|| ChartSelector::index(self.position), ChartSelector::name)
    }

    /// Return a selector for this chart's checked catalog position.
    #[must_use]
    pub const fn position_selector(&self) -> ChartSelector<'static> {
        ChartSelector::index(self.position)
    }
}

/// Errors raised while resolving a semantic chart selector.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartSelectorError {
    /// A name selector cannot select an empty visible title.
    EmptyName,
    /// More than one chart has the requested exact visible title.
    DuplicateChartTitle {
        /// The ambiguous visible title.
        name: Box<str>,
    },
}

impl std::fmt::Display for ChartSelectorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::EmptyName => formatter.write_str("chart selector name cannot be empty"),
            Self::DuplicateChartTitle { name } => {
                write!(formatter, "chart catalog contains duplicate title {name:?}")
            },
        }
    }
}

impl std::error::Error for ChartSelectorError {}

/// An immutable semantic catalog of charts owned by one slide.
///
/// The catalog stores only source order and optional visible titles. It is a
/// safe hand-off object for a native adapter: resolving a selector yields a
/// [`ChartDescriptor`] or its semantic position, never a native object ID.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartCatalog {
    charts: Box<[ChartDescriptor]>,
}

impl ChartCatalog {
    /// Build a catalog from chart titles in slide source order.
    ///
    /// Missing titles are represented by `None`; an empty title is retained as
    /// an existing native title but cannot be used as a name selector. Use a
    /// positional selector for that entry.
    #[must_use]
    pub fn from_titles<T>(titles: impl IntoIterator<Item = Option<T>>) -> Self
    where
        T: AsRef<str>,
    {
        let charts = titles
            .into_iter()
            .enumerate()
            .map(|(position, title)| ChartDescriptor {
                position,
                title: title.map(|title| title.as_ref().into()),
            })
            .collect::<Vec<_>>()
            .into_boxed_slice();
        Self { charts }
    }

    /// Build a catalog from owned chart titles without copying their strings.
    ///
    /// This variant is intended for adapters that already decoded titles into
    /// [`String`] values. Converting each string directly into its boxed
    /// semantic representation reuses the existing allocation when possible;
    /// [`Self::from_titles`] remains the convenient borrowed-input form.
    #[must_use]
    pub fn from_owned_titles(titles: impl IntoIterator<Item = Option<String>>) -> Self {
        let charts = titles
            .into_iter()
            .enumerate()
            .map(|(position, title)| ChartDescriptor {
                position,
                title: title.map(String::into_boxed_str),
            })
            .collect::<Vec<_>>()
            .into_boxed_slice();
        Self { charts }
    }

    /// Build a catalog from owned chart titles with fallible collection growth.
    ///
    /// This variant is intended for adapters that decoded titles from an
    /// untrusted native source. Existing `String` allocations are transferred
    /// into the catalog without copying; only the descriptor collection is
    /// grown here.
    ///
    /// # Errors
    ///
    /// Returns the standard allocation error when the descriptor collection
    /// cannot reserve its next entry.
    pub fn try_from_owned_titles(
        titles: impl IntoIterator<Item = Option<String>>,
    ) -> Result<Self, TryReserveError> {
        let mut charts = Vec::new();
        for (position, title) in titles.into_iter().enumerate() {
            charts.try_reserve(1)?;
            charts.push(ChartDescriptor {
                position,
                title: title.map(String::into_boxed_str),
            });
        }
        Ok(Self {
            charts: charts.into_boxed_slice(),
        })
    }

    /// Build a catalog from borrowed chart titles with fallible ownership.
    ///
    /// Each title is copied into exactly sized storage before its descriptor
    /// is published. Use this when title values originate outside the
    /// caller's trusted semantic model.
    ///
    /// # Errors
    ///
    /// Returns the standard allocation error when either a title or the
    /// descriptor collection cannot reserve its required storage.
    pub fn try_from_titles<T>(
        titles: impl IntoIterator<Item = Option<T>>,
    ) -> Result<Self, TryReserveError>
    where
        T: AsRef<str>,
    {
        let mut charts = Vec::new();
        for (position, title) in titles.into_iter().enumerate() {
            charts.try_reserve(1)?;
            let title = title
                .map(|title| try_boxed_str(title.as_ref()))
                .transpose()?;
            charts.push(ChartDescriptor { position, title });
        }
        Ok(Self {
            charts: charts.into_boxed_slice(),
        })
    }

    /// Borrow chart summaries in source order.
    #[must_use]
    pub fn charts(&self) -> &[ChartDescriptor] {
        &self.charts
    }

    /// Return the number of charts in this catalog.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.charts.len()
    }

    /// Return whether this catalog contains no charts.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.charts.is_empty()
    }

    /// Resolve a selector against this immutable semantic catalog.
    ///
    /// Name matching is exact and case sensitive. Missing charts and
    /// out-of-range positions are represented by `None`; duplicate titles and
    /// empty name selectors are rejected rather than selecting arbitrarily.
    ///
    /// # Errors
    ///
    /// Returns [`ChartSelectorError::EmptyName`] for an empty name selector or
    /// [`ChartSelectorError::DuplicateChartTitle`] when the requested title is
    /// not unique in this catalog.
    pub fn select<'selector>(
        &self,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<Option<&ChartDescriptor>, ChartSelectorError> {
        match selector.into() {
            ChartSelector::Index(index) => Ok(self.charts.get(index)),
            ChartSelector::Name(name) => {
                if name.is_empty() {
                    return Err(ChartSelectorError::EmptyName);
                }
                let mut matches = self
                    .charts
                    .iter()
                    .filter(|chart| chart.title() == Some(name));
                let Some(chart) = matches.next() else {
                    return Ok(None);
                };
                if matches.next().is_some() {
                    return Err(ChartSelectorError::DuplicateChartTitle { name: name.into() });
                }
                Ok(Some(chart))
            },
        }
    }

    /// Resolve a selector to its semantic zero-based chart position.
    ///
    /// # Errors
    ///
    /// Returns the same selector errors as [`Self::select`].
    pub fn select_position<'selector>(
        &self,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<Option<usize>, ChartSelectorError> {
        self.select(selector)
            .map(|chart| chart.map(ChartDescriptor::position))
    }
}

fn try_boxed_str(value: &str) -> Result<Box<str>, TryReserveError> {
    let mut owned = String::new();
    owned.try_reserve_exact(value.len())?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::{ChartCatalog, ChartSelector, ChartSelectorError};
    use litchi_core::Position;

    #[test]
    fn index_selector_preserves_checked_position() {
        let selector = ChartSelector::index(3);

        assert_eq!(selector, ChartSelector::Index(3));
        assert_eq!(selector.as_index(), Some(3));
        assert_eq!(selector.as_position(), Some(Position::new(3)));
        assert_eq!(selector.as_name(), None);
    }

    #[test]
    fn typed_positions_round_trip_without_native_identity() {
        let position = Position::new(3);
        let selector = ChartSelector::position(position);
        let from_position: ChartSelector<'_> = position.into();

        assert_eq!(selector, ChartSelector::index(3));
        assert_eq!(from_position, selector);
        assert_eq!(selector.as_position(), Some(position));
        assert_eq!(ChartSelector::name("Chart").as_position(), None);
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
    fn checked_name_rejects_empty_input_without_changing_name_constructor() {
        assert_eq!(
            ChartSelector::try_name(""),
            Err(ChartSelectorError::EmptyName)
        );
        assert_eq!(
            ChartSelector::try_name("Revenue"),
            Ok(ChartSelector::name("Revenue"))
        );
        assert_eq!(ChartSelector::name(""), ChartSelector::Name(""));
    }

    #[test]
    fn selectors_are_copyable_value_inputs() {
        const INDEX: ChartSelector<'static> = ChartSelector::index(0);
        const NAME: ChartSelector<'static> = ChartSelector::name("Chart");

        assert_eq!(INDEX, INDEX);
        assert_eq!(NAME, NAME);
    }

    #[test]
    fn selectors_accept_semantic_name_and_position_inputs() {
        let name: ChartSelector<'_> = "Revenue".into();
        let position: ChartSelector<'_> = 2usize.into();

        assert_eq!(name, ChartSelector::name("Revenue"));
        assert_eq!(position, ChartSelector::index(2));
    }

    #[test]
    fn borrowed_owned_names_convert_without_allocating() {
        let chart_name = String::from("Revenue");
        let selector: ChartSelector<'_> = (&chart_name).into();

        assert_eq!(selector, ChartSelector::Name("Revenue"));
    }

    #[test]
    fn catalog_resolves_exact_titles_and_positions_without_native_identity() {
        let titles = [None, Some("Revenue"), Some("Cost")];
        let catalog = ChartCatalog::from_titles(titles);

        assert_eq!(catalog.len(), 3);
        assert!(!catalog.is_empty());
        assert_eq!(catalog.select_position(1usize), Ok(Some(1)));
        assert_eq!(catalog.select_position("Revenue"), Ok(Some(1)));
        assert_eq!(catalog.select_position("revenue"), Ok(None));
        assert_eq!(catalog.select_position(99usize), Ok(None));
        assert_eq!(catalog.charts()[1].title(), Some("Revenue"));
        assert_eq!(
            catalog.charts()[1].selector(),
            ChartSelector::name("Revenue")
        );
        assert_eq!(
            catalog.charts()[1].position_selector(),
            ChartSelector::index(1)
        );
    }

    #[test]
    fn catalog_accepts_owned_titles_without_changing_selection() {
        let catalog = ChartCatalog::from_owned_titles(vec![
            Some("Revenue".to_owned()),
            None,
            Some("Cost".to_owned()),
        ]);

        assert_eq!(catalog.len(), 3);
        assert_eq!(catalog.select_position("Revenue"), Ok(Some(0)));
        assert_eq!(catalog.select_position(2usize), Ok(Some(2)));
        assert_eq!(catalog.charts()[1].title(), None);
    }

    #[test]
    fn fallible_catalog_constructors_preserve_titles()
    -> Result<(), std::collections::TryReserveError> {
        let borrowed = ChartCatalog::try_from_titles([Some("Revenue"), None, Some("Cost")])?;
        assert_eq!(borrowed.charts()[0].title(), Some("Revenue"));
        assert_eq!(borrowed.charts()[1].title(), None);
        assert_eq!(borrowed.select_position("Cost"), Ok(Some(2)));

        let owned = ChartCatalog::try_from_owned_titles(vec![
            Some("Revenue".to_owned()),
            None,
            Some("Cost".to_owned()),
        ])?;
        assert_eq!(owned.charts(), borrowed.charts());
        Ok(())
    }

    #[test]
    fn catalog_rejects_empty_and_duplicate_name_selectors() {
        let duplicate = ChartCatalog::from_titles([Some("Revenue"), Some("Revenue")]);
        assert_eq!(
            duplicate.select_position("Revenue"),
            Err(ChartSelectorError::DuplicateChartTitle {
                name: "Revenue".into()
            })
        );

        let empty = ChartCatalog::from_titles([Some("")]);
        assert_eq!(
            empty.select_position(""),
            Err(ChartSelectorError::EmptyName)
        );
        assert_eq!(empty.select_position(0usize), Ok(Some(0)));
        assert_eq!(empty.charts()[0].selector(), ChartSelector::index(0));
    }

    #[test]
    fn duplicate_name_resolution_does_not_poison_positions_or_other_names() {
        let catalog = ChartCatalog::from_titles([Some("Revenue"), Some("Revenue"), Some("Costs")]);

        assert_eq!(
            catalog.select_position("Revenue"),
            Err(ChartSelectorError::DuplicateChartTitle {
                name: "Revenue".into()
            })
        );
        assert_eq!(catalog.select_position("Costs"), Ok(Some(2)));
        assert_eq!(catalog.select_position(0usize), Ok(Some(0)));
        assert_eq!(catalog.select_position(1usize), Ok(Some(1)));
        assert_eq!(catalog.select_position(99usize), Ok(None));
    }
}
