//! Immutable source-bound worksheet conditional-formatting state.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, SourceBackedPackage};

use super::Formatting;
use crate::error::Result;
use crate::{Selector, auto_filter};

/// Complete core conditional-formatting collections and their exact owner closure.
#[derive(Clone, Debug)]
pub struct Snapshot {
    values: Arc<Vec<Formatting>>,
    closure: auto_filter::Snapshot,
    format_cells_locked: bool,
}

impl Snapshot {
    /// Resolve and capture one existing normal worksheet.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        Self::from_closure(auto_filter::Snapshot::load(package, selector)?)
    }

    pub(super) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        Self::from_closure(auto_filter::Snapshot::load_source_backed(
            package, selector,
        )?)
    }

    fn from_closure(closure: auto_filter::Snapshot) -> Result<Self> {
        let values = super::package::parse_editable_conditional_formattings(
            closure.source_xml(),
            closure.differential_format_count(),
        )?;
        let protection = crate::sheet_protection::parse_protection(closure.source_xml())?;
        let format_cells_locked = protection
            .sheet_protection()
            .is_some_and(crate::sheet_protection::Protection::format_cells_locked);
        Ok(Self {
            values: Arc::new(values),
            closure,
            format_cells_locked,
        })
    }

    pub(super) fn from_rewritten_source(
        source: &Self,
        bytes: Vec<u8>,
        readback: Vec<Formatting>,
    ) -> Result<Self> {
        let closure = source.closure.rebind_worksheet_xml(bytes)?;
        Ok(Self {
            values: Arc::new(readback),
            closure,
            format_cells_locked: source.format_cells_locked,
        })
    }

    /// Complete ordered core `conditionalFormatting` collection.
    #[must_use]
    pub fn collections(&self) -> &[Formatting] {
        self.values.as_slice()
    }

    /// Developer-facing worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        self.closure.sheet_name()
    }

    /// Checked zero-based worksheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.closure.sheet_position()
    }

    /// Selected worksheet Part name.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        self.closure.worksheet_part_name()
    }

    /// Exact selected worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.closure.source_xml()
    }

    /// Shared exact selected worksheet XML.
    ///
    /// Managed snapshots return `ManagedPartDataArcEscape` instead of
    /// detaching the payload reservation.
    pub fn source_arc(&self) -> Result<Arc<Vec<u8>>> {
        self.closure.source_arc()
    }

    pub(super) fn check_execution(&self) -> Result<()> {
        self.closure.check_execution()
    }

    pub(super) fn differential_format_count(&self) -> usize {
        self.closure.differential_format_count()
    }

    pub(super) const fn mutation_locked(&self) -> bool {
        self.format_cells_locked
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.closure.same_source(&other.closure)
            && self.format_cells_locked == other.format_cells_locked
    }

    pub(super) fn matches_source_backed(&self, package: &SourceBackedPackage) -> Result<bool> {
        self.closure.matches_source_backed(package)
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        self.closure.matches_current_source(package)
    }
}
