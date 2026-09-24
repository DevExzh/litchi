use crate::{Result, ink};

impl super::Package {
    /// Inventory active InkML annotations across all reachable Word stories.
    ///
    /// Only story XML and candidate Ink payloads are loaded. Shared targets are
    /// parsed once per query; returned annotations pin their immutable payloads
    /// and managed memory reservations after this package is dropped. Orphan
    /// Ink parts are outside this active-annotation inventory.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, cancellation, invalid ownership or
    /// XML, unsupported Ink dependencies, or exhausted resource budgets.
    pub fn ink(&self) -> Result<ink::Snapshot> {
        self.ink_with_limits(ink::Limits::default())
    }

    /// Inventory active InkML using explicit finite resource limits.
    ///
    /// Declared payload sizes are checked before decompression, then actual
    /// lengths are checked before parsing. No managed payload is detached into
    /// an uncharged byte allocation and no unrelated payload is materialized.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, stale source, cancellation, invalid
    /// package/XML data, unsupported dependencies, or an exceeded budget.
    pub fn ink_with_limits(&self, limits: ink::Limits) -> Result<ink::Snapshot> {
        self.package.check_execution()?;
        self.package.source_version()?;
        let result = ink::package::load_source(&self.package, limits);
        // A mutation during semantic validation must win over a stale parse
        // error, just as it does for the pinned main-document reader.
        self.package.source_version()?;
        self.package.check_execution()?;
        result
    }
}
