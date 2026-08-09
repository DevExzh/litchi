//! Resource limits for calculation metadata codecs.

use crate::error::{Result, invalid};

const DEFAULT_BYTES: usize = 32 * 1024 * 1024;

/// Resource limits for reading, processing, retaining, and writing calculation metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_raw_bytes: usize,
    max_mce_bytes: usize,
    max_output_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_attributes: usize,
    max_features: usize,
    max_feature_name_bytes: usize,
    max_feature_names_bytes: usize,
    max_opaque_bytes: usize,
}

impl Limits {
    /// Returns the checked default resource limits.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_raw_bytes: DEFAULT_BYTES,
            max_mce_bytes: DEFAULT_BYTES,
            max_output_bytes: DEFAULT_BYTES,
            max_depth: 256,
            max_events: 1_000_000,
            max_attributes: 256,
            max_features: 65_536,
            max_feature_name_bytes: 64 * 1024,
            max_feature_names_bytes: 4 * 1024 * 1024,
            max_opaque_bytes: DEFAULT_BYTES,
        }
    }

    #[must_use]
    pub const fn max_raw_bytes(self) -> usize {
        self.max_raw_bytes
    }
    #[must_use]
    pub const fn max_mce_bytes(self) -> usize {
        self.max_mce_bytes
    }
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }
    #[must_use]
    pub const fn max_features(self) -> usize {
        self.max_features
    }
    #[must_use]
    pub const fn max_feature_name_bytes(self) -> usize {
        self.max_feature_name_bytes
    }
    #[must_use]
    pub const fn max_feature_names_bytes(self) -> usize {
        self.max_feature_names_bytes
    }
    #[must_use]
    pub const fn max_opaque_bytes(self) -> usize {
        self.max_opaque_bytes
    }

    pub fn with_max_raw_bytes(mut self, value: usize) -> Result<Self> {
        self.max_raw_bytes = nonzero(value, "max_raw_bytes")?;
        Ok(self)
    }

    pub fn with_max_mce_bytes(mut self, value: usize) -> Result<Self> {
        self.max_mce_bytes = nonzero(value, "max_mce_bytes")?;
        Ok(self)
    }

    pub fn with_max_output_bytes(mut self, value: usize) -> Result<Self> {
        self.max_output_bytes = nonzero(value, "max_output_bytes")?;
        Ok(self)
    }

    pub fn with_max_depth(mut self, value: usize) -> Result<Self> {
        self.max_depth = nonzero(value, "max_depth")?;
        Ok(self)
    }

    pub fn with_max_events(mut self, value: usize) -> Result<Self> {
        self.max_events = nonzero(value, "max_events")?;
        Ok(self)
    }

    pub fn with_max_attributes(mut self, value: usize) -> Result<Self> {
        self.max_attributes = nonzero(value, "max_attributes")?;
        Ok(self)
    }

    pub fn with_max_features(mut self, value: usize) -> Result<Self> {
        self.max_features = nonzero(value, "max_features")?;
        Ok(self)
    }

    pub fn with_max_feature_name_bytes(mut self, value: usize) -> Result<Self> {
        self.max_feature_name_bytes = nonzero(value, "max_feature_name_bytes")?;
        Ok(self)
    }

    pub fn with_max_feature_names_bytes(mut self, value: usize) -> Result<Self> {
        self.max_feature_names_bytes = nonzero(value, "max_feature_names_bytes")?;
        Ok(self)
    }

    pub fn with_max_opaque_bytes(mut self, value: usize) -> Result<Self> {
        self.max_opaque_bytes = nonzero(value, "max_opaque_bytes")?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

fn nonzero(value: usize, name: &str) -> Result<usize> {
    if value == 0 {
        Err(invalid(format!(
            "calculation metadata limit {name} must be nonzero"
        )))
    } else {
        Ok(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_match_existing_parser_bounds() {
        let limits = Limits::default();
        assert_eq!(limits.max_raw_bytes(), 32 * 1024 * 1024);
        assert_eq!(limits.max_mce_bytes(), 32 * 1024 * 1024);
        assert_eq!(limits.max_output_bytes(), 32 * 1024 * 1024);
        assert_eq!(limits.max_depth(), 256);
        assert_eq!(limits.max_events(), 1_000_000);
    }

    #[test]
    fn rejects_zero_for_every_limit() {
        let limits = Limits::new();
        assert!(limits.with_max_raw_bytes(0).is_err());
        assert!(limits.with_max_mce_bytes(0).is_err());
        assert!(limits.with_max_output_bytes(0).is_err());
        assert!(limits.with_max_depth(0).is_err());
        assert!(limits.with_max_events(0).is_err());
        assert!(limits.with_max_attributes(0).is_err());
        assert!(limits.with_max_features(0).is_err());
        assert!(limits.with_max_feature_name_bytes(0).is_err());
        assert!(limits.with_max_feature_names_bytes(0).is_err());
        assert!(limits.with_max_opaque_bytes(0).is_err());
    }
}
