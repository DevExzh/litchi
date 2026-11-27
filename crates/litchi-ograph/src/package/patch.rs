//! Source-checked whole-package patches for Graph chart metadata edits.

use std::sync::Arc;

use crate::{Error, Result, chart};

use super::Package;

/// A published standalone Graph package edit: the new package, its
/// source-checked patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// Content-free diagnostics of one package commit.
pub use litchi_core::edit::Diagnostics;

/// A reversible, source-bound whole-package patch.
#[derive(Debug, Clone)]
pub struct Patch {
    source: Arc<Vec<u8>>,
    replacement: Arc<Vec<u8>>,
    chart: chart::Patch,
}

impl Patch {
    pub(super) fn new(
        source: Arc<Vec<u8>>,
        replacement: Arc<Vec<u8>>,
        chart: chart::Patch,
    ) -> Self {
        Self {
            source,
            replacement,
            chart,
        }
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }

    /// Exact replacement bytes produced by this patch.
    #[must_use]
    pub fn replacement(&self) -> &[u8] {
        &self.replacement
    }

    /// The typed chart patch nested inside this package publication.
    #[must_use]
    pub const fn chart(&self) -> &chart::Patch {
        &self.chart
    }

    /// Whether the typed package edit was a semantic no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        (Arc::ptr_eq(&self.source, &self.replacement) || self.source == self.replacement)
            && self.chart.is_noop()
    }

    /// Returns the source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.replacement),
            replacement: Arc::clone(&self.source),
            chart: self.chart.inverse(),
        }
    }

    pub(super) fn apply_to(&self, source: &Package) -> Result<Commit> {
        if !Arc::ptr_eq(source.bytes.shared(), &self.source)
            && **source.bytes.shared() != *self.source
        {
            return Err(Error::PatchConflict {
                operation: "package-patch",
                reason: "patch source package does not match the target package",
            });
        }
        let package = Package::from_shared(Arc::clone(&self.replacement), source.limits)?;
        Ok(Commit::new(
            package,
            self.clone(),
            Diagnostics::new(!self.is_noop()),
        ))
    }
}
