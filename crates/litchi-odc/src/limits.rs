//! Caller-selected resource ceilings for chart parsing, authoring, and history.

use litchi_core::{Error, Result};

/// Absolute package byte ceiling accepted by this crate.
const HARD_MAX_PACKAGE_BYTES: usize = 256 * 1024 * 1024;
/// Absolute content XML byte ceiling accepted by this crate.
const HARD_MAX_CONTENT_BYTES: usize = 64 * 1024 * 1024;
/// Absolute XML nesting ceiling accepted by this crate.
const HARD_MAX_DEPTH: usize = 4_096;
/// Absolute semantic collection ceiling.
const HARD_MAX_ITEMS: usize = 16_777_216;

/// Resource limits retained by snapshots and reused by edits and patches.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    package_bytes: usize,
    content_bytes: usize,
    depth: usize,
    axes: usize,
    series: usize,
    data_points: usize,
    cached_rows: usize,
    cached_cells: usize,
    range_items: usize,
    resources: usize,
    history: usize,
    scalar_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            package_bytes: HARD_MAX_PACKAGE_BYTES,
            content_bytes: HARD_MAX_CONTENT_BYTES,
            depth: 256,
            axes: 16_384,
            series: 65_536,
            data_points: HARD_MAX_ITEMS,
            cached_rows: HARD_MAX_ITEMS,
            cached_cells: HARD_MAX_ITEMS,
            range_items: 65_536,
            resources: 100_000,
            history: 4_096,
            scalar_bytes: 64 * 1024,
        }
    }
}

impl Limits {
    /// Return the default retained limits.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the package byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_package_bytes(mut self, value: usize) -> Result<Self> {
        self.package_bytes = checked(value, HARD_MAX_PACKAGE_BYTES, "package bytes")?;
        Ok(self)
    }

    /// Set the XML content byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_content_bytes(mut self, value: usize) -> Result<Self> {
        self.content_bytes = checked(value, HARD_MAX_CONTENT_BYTES, "content bytes")?;
        Ok(self)
    }

    /// Set the XML nesting ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_depth(mut self, value: usize) -> Result<Self> {
        self.depth = checked(value, HARD_MAX_DEPTH, "XML depth")?;
        Ok(self)
    }

    /// Set the retained axis ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_axes(mut self, value: usize) -> Result<Self> {
        self.axes = checked(value, HARD_MAX_ITEMS, "axis count")?;
        Ok(self)
    }

    /// Set the retained series ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_series(mut self, value: usize) -> Result<Self> {
        self.series = checked(value, HARD_MAX_ITEMS, "series count")?;
        Ok(self)
    }

    /// Set the expanded data-point ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_data_points(mut self, value: usize) -> Result<Self> {
        self.data_points = checked(value, HARD_MAX_ITEMS, "data-point count")?;
        Ok(self)
    }

    /// Set the expanded cached-row ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_cached_rows(mut self, value: usize) -> Result<Self> {
        self.cached_rows = checked(value, HARD_MAX_ITEMS, "cached-row count")?;
        Ok(self)
    }

    /// Set the expanded cached-cell ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_cached_cells(mut self, value: usize) -> Result<Self> {
        self.cached_cells = checked(value, HARD_MAX_ITEMS, "cached-cell count")?;
        Ok(self)
    }

    /// Set the number of addresses accepted in one range list.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_range_items(mut self, value: usize) -> Result<Self> {
        self.range_items = checked(value, HARD_MAX_ITEMS, "range-list item count")?;
        Ok(self)
    }

    /// Set the auxiliary-resource count ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_resources(mut self, value: usize) -> Result<Self> {
        self.resources = checked(value, HARD_MAX_ITEMS, "resource count")?;
        Ok(self)
    }

    /// Set the undo-history entry ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_history(mut self, value: usize) -> Result<Self> {
        self.history = checked(value, HARD_MAX_ITEMS, "history length")?;
        Ok(self)
    }

    /// Set the scalar string byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error for zero or a value above the hard safety ceiling.
    pub fn with_scalar_bytes(mut self, value: usize) -> Result<Self> {
        self.scalar_bytes = checked(value, HARD_MAX_CONTENT_BYTES, "scalar bytes")?;
        Ok(self)
    }

    #[must_use]
    pub const fn max_package_bytes(self) -> usize {
        self.package_bytes
    }

    #[must_use]
    pub const fn max_content_bytes(self) -> usize {
        self.content_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.depth
    }

    #[must_use]
    pub const fn max_axes(self) -> usize {
        self.axes
    }

    #[must_use]
    pub const fn max_series(self) -> usize {
        self.series
    }

    #[must_use]
    pub const fn max_data_points(self) -> usize {
        self.data_points
    }

    #[must_use]
    pub const fn max_cached_rows(self) -> usize {
        self.cached_rows
    }

    #[must_use]
    pub const fn max_cached_cells(self) -> usize {
        self.cached_cells
    }

    #[must_use]
    pub const fn max_range_items(self) -> usize {
        self.range_items
    }

    #[must_use]
    pub const fn max_resources(self) -> usize {
        self.resources
    }

    #[must_use]
    pub const fn max_history(self) -> usize {
        self.history
    }

    #[must_use]
    pub const fn max_scalar_bytes(self) -> usize {
        self.scalar_bytes
    }
}

fn checked(value: usize, hard: usize, label: &str) -> Result<usize> {
    if value == 0 || value > hard {
        return Err(Error::InvalidFormat(format!(
            "ODC {label} limit must be between 1 and {hard}"
        )));
    }
    Ok(value)
}
