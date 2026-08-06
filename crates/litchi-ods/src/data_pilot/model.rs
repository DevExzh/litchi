//! Immutable DataPilot catalogs and semantic selectors.

use crate::model::data_pilot::{self as vocabulary, Table};
use crate::package::Package;
use litchi_core::Result;

use super::codec::{Location, locate};
use super::transaction::Transaction;
use super::validation;

/// Immutable DataPilot declarations bound to one ODS package snapshot.
pub struct Catalog<'source> {
    pub(crate) source: &'source Package,
    pub(crate) source_xml: &'source str,
    pub(crate) location: Location,
    pub(crate) tables: Vec<Table>,
    pub(crate) present: bool,
}

impl<'source> Catalog<'source> {
    pub(crate) fn load(source: &'source Package) -> Result<Self> {
        let source_xml = source.content_xml();
        let location = locate(source_xml)?;
        let tables = vocabulary::parse_data_pilot_tables(source_xml)?;
        validation::validate_snapshot(source_xml, &location, &tables)?;
        Ok(Self {
            source,
            source_xml,
            present: location.container.is_some(),
            location,
            tables,
        })
    }

    /// Borrow the declarations in source order.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Iterate declarations without copying their nested source metadata.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Table> {
        self.tables.iter()
    }

    /// Return the number of declarations in this catalog.
    #[must_use]
    pub fn len(&self) -> usize {
        self.tables.len()
    }

    /// Return whether the source contains a physical `table:data-pilot-tables`
    /// owner, including an explicitly empty owner.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.present
    }

    /// Return whether the catalog contains no typed declarations.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// Select a declaration by checked source position.
    pub fn at(&self, index: usize) -> Result<Option<&Table>> {
        Ok(self.tables.get(index))
    }

    /// Select one declaration by its exact producer-visible name.
    pub fn named(&self, name: &str) -> Result<Option<&Table>> {
        self.get(Selector::Name(name))
    }

    /// Select by exact name or checked zero-based source position.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Table>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.tables, selector.into()).map(|index| index.map(|index| &self.tables[index]))
    }

    /// Start an isolated clone-staged transaction over this catalog.
    pub fn transaction(&self) -> Transaction<'source> {
        Transaction::from_catalog(self)
    }
}

/// Primary semantic selector for DataPilot declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Checked zero-based source order.
    Index(usize),
    /// Exact producer-visible DataPilot name.
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

pub(crate) fn select<'a>(tables: &[Table], selector: Selector<'a>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < tables.len()).then_some(index)),
        Selector::Name(name) => {
            let mut selected = None;
            for (index, table) in tables.iter().enumerate() {
                if table.name == name {
                    if selected.is_some() {
                        return Err(litchi_core::Error::InvalidFormat(format!(
                            "ODS DataPilot name '{name}' is ambiguous"
                        )));
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}
