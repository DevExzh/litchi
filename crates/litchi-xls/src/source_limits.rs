//! Finite physical ingress limits for legacy Excel workbooks.
//!
//! These limits are applied to the complete CFB source before the container
//! parser is entered and to every declared CFB stream before an eager XLS reader
//! materializes its payload. They are deliberately separate from the older
//! positional-reader policy in [`crate::SourceBackedLimits`], which also
//! limits worksheet scans, query indexes, and eager materialization work.

use crate::error::{Error, LimitKind, Result};

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

    /// Return the component-wise stricter combination of two profiles.
    ///
    /// A raw reader may accept a narrower request than the profile that
    /// admitted its source, but it must never widen that retained ingress
    /// policy. This helper keeps that rule explicit at every raw boundary.
    #[must_use]
    pub const fn constrained_by(self, other: Self) -> Self {
        Self {
            input_bytes: min_usize(self.input_bytes, other.input_bytes),
            entry_bytes: min_usize(self.entry_bytes, other.entry_bytes),
            aggregate_bytes: min_usize(self.aggregate_bytes, other.aggregate_bytes),
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
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
        validate_source_limit(LimitKind::InputBytes, value, SourceLimits::MAX_INPUT_BYTES)?;
        self.limits.input_bytes = value;
        Ok(self)
    }

    /// Set the per-entry input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_ENTRY_BYTES`].
    pub fn max_entry_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(LimitKind::EntryBytes, value, SourceLimits::MAX_ENTRY_BYTES)?;
        self.limits.entry_bytes = value;
        Ok(self)
    }

    /// Set the aggregate input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_AGGREGATE_BYTES`].
    pub fn max_aggregate_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            LimitKind::AggregateBytes,
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

fn validate_source_limit(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value: value as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
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
            (LimitKind::InputBytes, 0, SourceLimits::MAX_INPUT_BYTES),
            (
                LimitKind::InputBytes,
                SourceLimits::MAX_INPUT_BYTES.saturating_add(1),
                SourceLimits::MAX_INPUT_BYTES,
            ),
            (LimitKind::EntryBytes, 0, SourceLimits::MAX_ENTRY_BYTES),
            (
                LimitKind::EntryBytes,
                SourceLimits::MAX_ENTRY_BYTES.saturating_add(1),
                SourceLimits::MAX_ENTRY_BYTES,
            ),
            (
                LimitKind::AggregateBytes,
                0,
                SourceLimits::MAX_AGGREGATE_BYTES,
            ),
            (
                LimitKind::AggregateBytes,
                SourceLimits::MAX_AGGREGATE_BYTES.saturating_add(1),
                SourceLimits::MAX_AGGREGATE_BYTES,
            ),
        ];
        for (kind, value, maximum) in cases {
            let result = match kind {
                LimitKind::InputBytes => SourceLimits::builder().max_input_bytes(value),
                LimitKind::EntryBytes => SourceLimits::builder().max_entry_bytes(value),
                LimitKind::AggregateBytes => SourceLimits::builder().max_aggregate_bytes(value),
                _ => panic!("source-limit test contains an edit dimension"),
            };
            assert!(matches!(
                result,
                Err(Error::InvalidLimit {
                    kind: observed,
                    value: actual,
                    maximum: bound,
                }) if observed == kind && actual == value as u64 && bound == maximum as u64
            ));
        }
    }

    #[test]
    fn constrained_profile_never_widens_ingress() {
        let retained = SourceLimits::builder()
            .max_input_bytes(64)
            .unwrap()
            .max_entry_bytes(32)
            .unwrap()
            .max_aggregate_bytes(48)
            .unwrap()
            .build()
            .unwrap();
        let requested = SourceLimits::default();
        let effective = retained.constrained_by(requested);
        assert_eq!(effective, retained);
        assert_eq!(requested.constrained_by(retained), retained);
    }
}
