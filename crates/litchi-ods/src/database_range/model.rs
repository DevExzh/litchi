//! Immutable database-range catalogs and checked selectors.

use crate::model::database_range::{self as vocabulary, Range};
use crate::package::Package;
use litchi_core::{Error, Result};

use super::codec::{Location, locate};
use super::transaction::Transaction;
use super::validation;

/// Immutable database-range declarations bound to one ODS package snapshot.
pub struct Catalog<'source> {
    pub(crate) source: &'source Package,
    pub(crate) source_xml: &'source str,
    pub(crate) location: Location,
    pub(crate) ranges: Vec<Range>,
    pub(crate) present: bool,
}

impl<'source> Catalog<'source> {
    pub(crate) fn load(source: &'source Package) -> Result<Self> {
        let source_xml = source.content_xml();
        let location = locate(source_xml)?;
        let ranges = if let Some(container) = &location.container {
            let fragment = super::codec::owner_fragment(source_xml, container)?;
            vocabulary::parse_database_ranges(&fragment)?
        } else {
            Vec::new()
        };
        validation::validate_snapshot(source_xml, &location, &ranges)?;
        Ok(Self {
            source,
            source_xml,
            present: location.container.is_some(),
            location,
            ranges,
        })
    }

    /// Borrow declarations in source order.
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }

    /// Alias for [`Self::ranges`].
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Range> {
        self.ranges.iter()
    }

    /// Number of declarations in this catalog.
    #[must_use]
    pub fn len(&self) -> usize {
        self.ranges.len()
    }

    /// Whether a physical `table:database-ranges` owner exists, including an
    /// explicitly empty owner.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.present
    }

    /// Whether no typed declarations are present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }

    /// Select a declaration by checked source position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn at(&self, index: usize) -> Result<Option<&Range>> {
        Ok(self.ranges.get(index))
    }

    /// Select one declaration by exact producer-visible name.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn named(&self, name: &str) -> Result<Option<&Range>> {
        self.get(Selector::Name(name))
    }

    /// Select by exact name or checked zero-based source position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Range>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.ranges, selector.into()).map(|index| index.map(|index| &self.ranges[index]))
    }

    /// Start an isolated clone-staged transaction over this catalog.
    #[must_use]
    pub fn transaction(&self) -> Transaction<'source> {
        Transaction::from_catalog(self)
    }
}

/// Primary semantic selector for database-range declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Checked zero-based source order.
    Index(usize),
    /// Exact producer-visible database-range name.
    Name(&'a str),
}

impl From<usize> for Selector<'static> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

pub(crate) fn select(ranges: &[Range], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < ranges.len()).then_some(index)),
        Selector::Name(name) => {
            let mut selected = None;
            for (index, range) in ranges.iter().enumerate() {
                if range.name.as_deref() == Some(name) {
                    if selected.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "ODS database-range name '{name}' is ambiguous"
                        )));
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}
