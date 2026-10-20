//! Finite physical ingress limits for legacy Excel workbooks.
//!
//! These limits are applied to the complete CFB source before the container
//! parser is entered and to every declared CFB stream before an eager XLS reader
//! materializes its payload. They are deliberately separate from the older
//! positional-reader policy in [`crate::SourceBackedLimits`], which also
//! limits worksheet scans, query indexes, and eager materialization work.

use crate::error::{Error, Result};

/// Finite physical ingress limits for a legacy Excel package.
///
/// `input_bytes` bounds the complete CFB source before it is parsed,
/// `entry_bytes` bounds each declared CFB stream before a reader materializes
/// it, and `aggregate_bytes` bounds the declared stream payloads together.
/// The root mini-stream backing is excluded because mini-stream payload
/// lengths are already charged through their stream entries. Selecting a
/// profile stores only these three integer ceilings; it does not reserve any
/// payload memory.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: usize,
    entry_bytes: usize,
    aggregate_bytes: usize,
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl SourceLimits {
    pub(crate) const DEFAULT: Self = Self {
        input_bytes: Self::DEFAULT_MAX_INPUT_BYTES,
        entry_bytes: Self::DEFAULT_MAX_ENTRY_BYTES,
        aggregate_bytes: Self::DEFAULT_MAX_AGGREGATE_BYTES,
    };

    /// Shared default complete-source input limit.
    pub const DEFAULT_MAX_INPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Shared default per-entry input limit.
    pub const DEFAULT_MAX_ENTRY_BYTES: usize = 512 * 1024 * 1024;
    /// Shared default aggregate input limit.
    pub const DEFAULT_MAX_AGGREGATE_BYTES: usize = 2 * 1024 * 1024 * 1024;
    /// Hard safety ceiling for the complete source.
    pub const MAX_INPUT_BYTES: usize = 1024 * 1024 * 1024;
    /// Hard safety ceiling for one declared stream entry.
    pub const MAX_ENTRY_BYTES: usize = 512 * 1024 * 1024;
    /// Hard safety ceiling for aggregate declared stream payloads.
    pub const MAX_AGGREGATE_BYTES: usize = 2 * 1024 * 1024 * 1024;

    /// Start a checked builder from the shared defaults.
    ///
    /// ```
    /// use litchi_xls::{OpenOptions, SourceLimits};
    /// let limits = SourceLimits::builder()
    ///     .max_input_bytes(64 * 1024 * 1024)?
    ///     .max_entry_bytes(16 * 1024 * 1024)?
    ///     .max_aggregate_bytes(64 * 1024 * 1024)?
    ///     .build()?;
    /// let options = OpenOptions::new().source_limits(limits);
    /// assert_eq!(options.source_limits_value(), limits);
    /// # Ok::<(), litchi_xls::Error>(())
    /// ```
    #[must_use]
    pub fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::default(),
        }
    }

    /// Return the complete-source input limit.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Return the per-entry input limit.
    #[must_use]
    pub const fn max_entry_bytes(self) -> usize {
        self.entry_bytes
    }

    /// Return the aggregate input limit.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> usize {
        self.aggregate_bytes
    }
}

/// Consuming builder for a checked [`SourceLimits`] profile.
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Set the complete-source input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_INPUT_BYTES`].
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::InputBytes,
            value,
            SourceLimits::MAX_INPUT_BYTES,
        )?;
        self.limits.input_bytes = value;
        Ok(self)
    }

    /// Set the per-entry input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_ENTRY_BYTES`].
    pub fn max_entry_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::EntryBytes,
            value,
            SourceLimits::MAX_ENTRY_BYTES,
        )?;
        self.limits.entry_bytes = value;
        Ok(self)
    }

    /// Set the aggregate input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_AGGREGATE_BYTES`].
    pub fn max_aggregate_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::AggregateBytes,
            value,
            SourceLimits::MAX_AGGREGATE_BYTES,
        )?;
        self.limits.aggregate_bytes = value;
        Ok(self)
    }

    /// Finish the checked profile.
    pub fn build(self) -> Result<SourceLimits> {
        Ok(self.limits)
    }
}

fn validate_source_limit(kind: SourceLimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

/// Physical resource dimension governed by an XLS source limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Complete source bytes, before the CFB container is parsed.
    InputBytes,
    /// Bytes in one named CFB stream entry.
    EntryBytes,
    /// Aggregate bytes in all declared CFB stream entries.
    AggregateBytes,
}

impl SourceLimitKind {
    pub(crate) const fn description(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
        }
    }
}

impl std::fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.description())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_and_hard_boundaries_are_finite() {
        let limits = SourceLimits::default();
        assert_eq!(limits.max_input_bytes(), 512 * 1024 * 1024);
        assert_eq!(limits.max_entry_bytes(), 512 * 1024 * 1024);
        assert_eq!(limits.max_aggregate_bytes(), 2 * 1024 * 1024 * 1024);

        let maximum = SourceLimits::builder()
            .max_input_bytes(SourceLimits::MAX_INPUT_BYTES)
            .unwrap()
            .max_entry_bytes(SourceLimits::MAX_ENTRY_BYTES)
            .unwrap()
            .max_aggregate_bytes(SourceLimits::MAX_AGGREGATE_BYTES)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(maximum.max_input_bytes(), SourceLimits::MAX_INPUT_BYTES);
        assert_eq!(maximum.max_entry_bytes(), SourceLimits::MAX_ENTRY_BYTES);
        assert_eq!(
            maximum.max_aggregate_bytes(),
            SourceLimits::MAX_AGGREGATE_BYTES
        );
    }

    #[test]
    fn builder_rejects_zero_and_values_above_hard_bounds() {
        let cases = [
            (
                SourceLimitKind::InputBytes,
                0,
                SourceLimits::MAX_INPUT_BYTES,
            ),
            (
                SourceLimitKind::InputBytes,
                SourceLimits::MAX_INPUT_BYTES.saturating_add(1),
                SourceLimits::MAX_INPUT_BYTES,
            ),
            (
                SourceLimitKind::EntryBytes,
                0,
                SourceLimits::MAX_ENTRY_BYTES,
            ),
            (
                SourceLimitKind::EntryBytes,
                SourceLimits::MAX_ENTRY_BYTES.saturating_add(1),
                SourceLimits::MAX_ENTRY_BYTES,
            ),
            (
                SourceLimitKind::AggregateBytes,
                0,
                SourceLimits::MAX_AGGREGATE_BYTES,
            ),
            (
                SourceLimitKind::AggregateBytes,
                SourceLimits::MAX_AGGREGATE_BYTES.saturating_add(1),
                SourceLimits::MAX_AGGREGATE_BYTES,
            ),
        ];
        for (kind, value, maximum) in cases {
            let result = match kind {
                SourceLimitKind::InputBytes => SourceLimits::builder().max_input_bytes(value),
                SourceLimitKind::EntryBytes => SourceLimits::builder().max_entry_bytes(value),
                SourceLimitKind::AggregateBytes => {
                    SourceLimits::builder().max_aggregate_bytes(value)
                },
            };
            assert!(matches!(
                result,
                Err(Error::InvalidLimit {
                    kind: observed,
                    value: actual,
                    maximum: bound,
                }) if observed == kind && actual == value && bound == maximum
            ));
        }
    }
}
