//! Limits for retaining immutable transaction steps.

use std::collections::TryReserveError;
use std::fmt;

/// A bounded resource of an explicit transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Number of mutation calls.
    Operations,
    /// Retained native snapshot and patch data.
    SnapshotBytes,
    /// Distinct XLSX grid or PPTX shape targets expanded by one mutation.
    ExpandedTargets,
    /// Authored XLSX/PPTX payload and action charge staged by one mutation.
    StagedBytes,
}

/// Content-free transaction admission or replay failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A patch belongs to another immutable owner or security disposition.
    PatchConflict,
    /// A configured bound is zero or exceeds the supported integer range.
    InvalidLimit {
        /// Resource being configured.
        kind: LimitKind,
        /// Rejected value.
        value: usize,
        /// Largest supported value.
        maximum: usize,
    },
    /// A transaction exceeds its caller-selected bound.
    Limit {
        /// Resource consumed.
        kind: LimitKind,
        /// Observed count, saturated on overflow.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Allocation of the bounded step list failed.
    Allocation {
        /// Allocator refusal.
        source: TryReserveError,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::PatchConflict => {
                formatter.write_str("the patch belongs to a different immutable snapshot")
            },
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid transaction {kind:?} limit {value}; expected 1..={maximum}"
            ),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "transaction {kind:?} limit exceeded: {observed} > {maximum}"
            ),
            Self::Allocation { .. } => formatter.write_str("cannot allocate transaction steps"),
        }
    }
}
impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Allocation { source } => Some(source),
            _ => None,
        }
    }
}

/// Bounds on an explicit transaction: 10,000 mutation calls and 256 MiB of
/// retained native snapshots by default. XLSX and PPTX mutations additionally default
/// to 65,536 expanded targets and a 16 MiB staged payload/action charge.
/// Coalesced XLSX cell and PPTX shape runs close before their summed charge
/// exceeds the staged-byte bound; an individual oversized mutation is refused.
///
/// Snapshot bytes are a conservative logical charge, counting shared data at
/// every step. This is not a measurement of heap use: native owners retain
/// source/package data under their existing ingress and execution limits.
///
/// ```
/// use litchi_ops::transaction::SemanticLimits;
/// let limits = SemanticLimits::builder().max_operations(100)?.max_snapshot_bytes(1024 * 1024)?.build()?;
/// assert_eq!(limits.max_operations(), 100);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy)]
pub struct SemanticLimits {
    operations: usize,
    snapshot_bytes: usize,
    expanded_targets: usize,
    staged_bytes: usize,
}
impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            operations: 10_000,
            snapshot_bytes: 256 * 1024 * 1024,
            expanded_targets: 65_536,
            staged_bytes: 16 * 1024 * 1024,
        }
    }
}
impl SemanticLimits {
    /// Start a consuming builder with default bounds.
    #[must_use]
    pub fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder(Self::default())
    }
    /// Maximum mutation calls, including no-ops.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.operations
    }
    /// Maximum distinct XLSX grid or PPTX shape targets expanded by one mutation.
    /// This adapter admission bound is independent of native transfer limits.
    #[must_use]
    pub const fn max_expanded_targets(self) -> usize {
        self.expanded_targets
    }
    /// Maximum authored payload/action charge staged by one XLSX or PPTX mutation.
    /// Counts repeated scalar payloads before native staging. Coalesced cell
    /// runs also close at this bound. This is separate from retained snapshots
    /// and does not measure total allocator use.
    #[must_use]
    pub const fn max_staged_bytes(self) -> usize {
        self.staged_bytes
    }
    /// Maximum summed retained native snapshot charge.
    ///
    /// DOCX charges before/after main-document XML lengths. XLSX charges
    /// retained package data and native patch buffers, including no-op steps.
    /// PPTX charges the retained native patch wire size.
    /// Shared endpoints are charged repeatedly; this is a semantic retention
    /// budget, not a complete bound on allocator or derived-cache memory.
    #[must_use]
    pub const fn max_snapshot_bytes(self) -> usize {
        self.snapshot_bytes
    }
}

/// Consuming builder for [`SemanticLimits`].
#[derive(Debug, Clone, Copy)]
pub struct SemanticLimitsBuilder(SemanticLimits);
impl SemanticLimitsBuilder {
    /// Set the mutation-call bound.
    ///
    /// # Errors
    /// Refuses zero and values larger than `isize::MAX`.
    pub fn max_operations(mut self, value: usize) -> crate::Result<Self> {
        check(LimitKind::Operations, value)?;
        self.0.operations = value;
        Ok(self)
    }
    /// Set the conservative retained native snapshot byte bound.
    ///
    /// # Errors
    /// Refuses zero and values larger than `isize::MAX`.
    pub fn max_snapshot_bytes(mut self, value: usize) -> crate::Result<Self> {
        check(LimitKind::SnapshotBytes, value)?;
        self.0.snapshot_bytes = value;
        Ok(self)
    }
    /// Set the expanded-target bound for each XLSX or PPTX mutation.
    ///
    /// # Errors
    /// Refuses zero and values larger than `isize::MAX`.
    pub fn max_expanded_targets(mut self, value: usize) -> crate::Result<Self> {
        check(LimitKind::ExpandedTargets, value)?;
        self.0.expanded_targets = value;
        Ok(self)
    }
    /// Set the authored payload/action byte bound for each XLSX or PPTX mutation.
    ///
    /// # Errors
    /// Refuses zero and values larger than `isize::MAX`.
    pub fn max_staged_bytes(mut self, value: usize) -> crate::Result<Self> {
        check(LimitKind::StagedBytes, value)?;
        self.0.staged_bytes = value;
        Ok(self)
    }
    /// Finish the checked profile.
    ///
    /// # Errors
    /// Refuses a profile outside the supported integer range.
    pub fn build(self) -> crate::Result<SemanticLimits> {
        check(LimitKind::Operations, self.0.operations)?;
        check(LimitKind::SnapshotBytes, self.0.snapshot_bytes)?;
        check(LimitKind::ExpandedTargets, self.0.expanded_targets)?;
        check(LimitKind::StagedBytes, self.0.staged_bytes)?;
        Ok(self.0)
    }
}
fn check(kind: LimitKind, value: usize) -> crate::Result<()> {
    let maximum = isize::MAX as usize;
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        }
        .into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Code, error::Source};

    #[test]
    fn checked_profiles_reject_zero_and_out_of_range_for_each_dimension() {
        for value in [0, usize::MAX] {
            for error in [
                SemanticLimits::builder().max_operations(value).unwrap_err(),
                SemanticLimits::builder()
                    .max_snapshot_bytes(value)
                    .unwrap_err(),
                SemanticLimits::builder()
                    .max_expanded_targets(value)
                    .unwrap_err(),
                SemanticLimits::builder()
                    .max_staged_bytes(value)
                    .unwrap_err(),
            ] {
                assert_eq!(error.code(), Code::InvalidArgument);
                assert!(
                    matches!(error.leaf(), Source::Transaction(Error::InvalidLimit { value: rejected, .. }) if *rejected == value)
                );
                assert!(std::error::Error::source(&error).is_some());
            }
        }
        for bound in [1, isize::MAX as usize] {
            let limits = SemanticLimits::builder()
                .max_operations(bound)
                .unwrap()
                .max_snapshot_bytes(bound)
                .unwrap()
                .max_expanded_targets(bound)
                .unwrap()
                .max_staged_bytes(bound)
                .unwrap()
                .build()
                .unwrap();
            assert_eq!(limits.max_operations(), bound);
            assert_eq!(limits.max_snapshot_bytes(), bound);
            assert_eq!(limits.max_expanded_targets(), bound);
            assert_eq!(limits.max_staged_bytes(), bound);
        }
    }

    #[test]
    fn explicit_transaction_defaults_and_send_sync_contract() {
        let limits = SemanticLimits::default();
        assert_eq!(limits.max_operations(), 10_000);
        assert_eq!(limits.max_snapshot_bytes(), 256 * 1024 * 1024);
        fn send_sync<T: Send + Sync>() {}
        send_sync::<crate::transaction::Patch>();
        send_sync::<crate::transaction::Commit>();
        send_sync::<crate::transaction::Edit>();
    }
}
