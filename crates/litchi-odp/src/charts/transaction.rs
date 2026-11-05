//! Immutable ODP chart inventories.

use super::codec::inventory;
use super::model::{Chart, Limits, Selector};
use crate::core::OwnedPackage;
use litchi_core::{Error, Result};

/// Immutable chart inventory bound to one ODP package snapshot.
pub struct Inventory<'source> {
    pub(crate) _source: std::marker::PhantomData<&'source OwnedPackage>,
    pub(crate) limits: Limits,
    pub(crate) charts: Vec<Chart>,
}

impl<'source> Inventory<'source> {
    pub(crate) fn load(source: &'source OwnedPackage, limits: Limits) -> Result<Self> {
        Ok(Self {
            _source: std::marker::PhantomData,
            limits,
            charts: inventory(source, limits)?,
        })
    }

    /// Return the resource budget used for this inventory.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the number of embedded charts in drawing order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.charts.len()
    }

    /// Return whether no chart occurrence was discovered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.charts.is_empty()
    }

    /// Iterate chart occurrences in source order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Chart> {
        self.charts.iter()
    }

    /// Select by checked zero-based discovery order.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn at(&self, index: usize) -> Result<Option<&Chart>> {
        Ok(self.charts.get(index))
    }

    /// Select by exact producer-visible frame name.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn named(&self, name: &str) -> Result<Option<&Chart>> {
        select(&self.charts, Selector::Name(name))
            .map(|index| index.map(|selected| &self.charts[selected]))
    }

    /// Select by exact name or checked zero-based position.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Chart>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.charts, selector.into())
            .map(|index| index.map(|selected| &self.charts[selected]))
    }
}

fn select<'a, S>(charts: &[Chart], selector: S) -> Result<Option<usize>>
where
    S: Into<Selector<'a>>,
{
    match selector.into() {
        Selector::Index(index) => Ok((index < charts.len()).then_some(index)),
        Selector::Name(name) => {
            let mut selected = None;
            for (index, chart) in charts.iter().enumerate() {
                if chart.name() == Some(name) {
                    if selected.is_some() {
                        return invalid("ODP embedded chart name is ambiguous");
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
