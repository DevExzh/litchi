//! Caller-selected resource ceilings for chart parsing, authoring, and history.
#![allow(
    clippy::missing_errors_doc,
    reason = "all fluent limit setters share the documented checked-ceiling contract"
)]

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
    max_package_bytes: usize,
    max_content_bytes: usize,
    max_depth: usize,
    max_axes: usize,
    max_series: usize,
    max_data_points: usize,
    max_resources: usize,
    max_history: usize,
    max_scalar_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_package_bytes: HARD_MAX_PACKAGE_BYTES,
            max_content_bytes: HARD_MAX_CONTENT_BYTES,
            max_depth: 256,
            max_axes: 16_384,
            max_series: 65_536,
            max_data_points: HARD_MAX_ITEMS,
            max_resources: 100_000,
            max_history: 4_096,
            max_scalar_bytes: 64 * 1024,
        }
    }
}

impl Limits {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_package_bytes(mut self, value: usize) -> Result<Self> {
        self.max_package_bytes = checked(value, HARD_MAX_PACKAGE_BYTES, "package bytes")?;
        Ok(self)
    }

    pub fn with_content_bytes(mut self, value: usize) -> Result<Self> {
        self.max_content_bytes = checked(value, HARD_MAX_CONTENT_BYTES, "content bytes")?;
        Ok(self)
    }

    pub fn with_depth(mut self, value: usize) -> Result<Self> {
        self.max_depth = checked(value, HARD_MAX_DEPTH, "XML depth")?;
        Ok(self)
    }

    pub fn with_axes(mut self, value: usize) -> Result<Self> {
        self.max_axes = checked(value, HARD_MAX_ITEMS, "axis count")?;
        Ok(self)
    }

    pub fn with_series(mut self, value: usize) -> Result<Self> {
        self.max_series = checked(value, HARD_MAX_ITEMS, "series count")?;
        Ok(self)
    }

    pub fn with_data_points(mut self, value: usize) -> Result<Self> {
        self.max_data_points = checked(value, HARD_MAX_ITEMS, "data-point count")?;
        Ok(self)
    }

    pub fn with_resources(mut self, value: usize) -> Result<Self> {
        self.max_resources = checked(value, HARD_MAX_ITEMS, "resource count")?;
        Ok(self)
    }

    pub fn with_history(mut self, value: usize) -> Result<Self> {
        self.max_history = checked(value, HARD_MAX_ITEMS, "history length")?;
        Ok(self)
    }

    pub fn with_scalar_bytes(mut self, value: usize) -> Result<Self> {
        self.max_scalar_bytes = checked(value, HARD_MAX_CONTENT_BYTES, "scalar bytes")?;
        Ok(self)
    }

    #[must_use]
    pub const fn max_package_bytes(self) -> usize {
        self.max_package_bytes
    }

    #[must_use]
    pub const fn max_content_bytes(self) -> usize {
        self.max_content_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub const fn max_axes(self) -> usize {
        self.max_axes
    }

    #[must_use]
    pub const fn max_series(self) -> usize {
        self.max_series
    }

    #[must_use]
    pub const fn max_data_points(self) -> usize {
        self.max_data_points
    }

    #[must_use]
    pub const fn max_resources(self) -> usize {
        self.max_resources
    }

    #[must_use]
    pub const fn max_history(self) -> usize {
        self.max_history
    }

    #[must_use]
    pub const fn max_scalar_bytes(self) -> usize {
        self.max_scalar_bytes
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
