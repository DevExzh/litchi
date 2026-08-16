//! Exact row-visibility state over the source-backed scalar worksheet closure.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::Row;

use super::rewrite;
use crate::Selector;
use crate::cell_values;
use crate::error::Result;

/// Exact source-bound row-owner visibility state.
#[derive(Clone, Debug)]
pub struct Snapshot {
    inner: cell_values::Snapshot,
    rows: Arc<[(Row, bool)]>,
}

impl Snapshot {
    /// Load the conservative row-visibility closure from an owning package.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        Self::from_inner(cell_values::Snapshot::load(package, selector)?)
    }

    pub(crate) fn from_inner(inner: cell_values::Snapshot) -> Result<Self> {
        let rows = rewrite::scan(inner.source_xml())?;
        inner.check_execution()?;
        Ok(Self {
            inner,
            rows: Arc::from(rows),
        })
    }

    pub(crate) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        source.check_execution()?;
        Self::from_inner(cell_values::Snapshot::from_rewritten_source(
            &source.inner,
            bytes,
        )?)
    }

    /// Selected worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        self.inner.sheet_name()
    }

    /// Selected zero-based sheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.inner.sheet_position()
    }

    /// Selected worksheet Part URI.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        self.inner.worksheet_part_name()
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.inner.source_xml()
    }

    /// Whether an explicit row owner exists at this coordinate.
    #[must_use]
    pub fn contains_row(&self, row: Row) -> bool {
        self.rows
            .binary_search_by_key(&row, |(row, _)| *row)
            .is_ok()
    }

    /// Effective direct visibility of an existing row owner.
    ///
    /// `None` means no explicit `<row>` owner exists at this coordinate.
    #[must_use]
    pub fn is_hidden(&self, row: Row) -> Option<bool> {
        self.rows
            .binary_search_by_key(&row, |(row, _)| *row)
            .ok()
            .map(|index| self.rows[index].1)
    }

    pub(crate) const fn inner(&self) -> &cell_values::Snapshot {
        &self.inner
    }

    pub(crate) fn check_execution(&self) -> Result<()> {
        self.inner.check_execution()
    }
}
