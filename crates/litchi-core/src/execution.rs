//! Runtime-neutral execution policy, cancellation, and resource charging.
//!
//! This module deliberately describes *what* an operation is allowed to use,
//! rather than providing a runtime or executor. Container and format crates
//! may adapt an [`ExecutionContext`] to a local scheduler without making
//! `litchi-core` depend on Rayon, Tokio, or a platform thread-affinity API.

use std::{
    num::{NonZeroU64, NonZeroUsize},
    sync::{
        Arc,
        atomic::{AtomicBool, Ordering},
    },
};

use thiserror::Error;

use crate::{Budget, Reservation, Resource, ResourceLimit};

/// CPU-affinity policy for workers created by a runtime adapter.
///
/// The neutral core currently supports only inheriting the operating system's
/// placement decision. Runtime adapters must record this choice rather than
/// claiming that workers were pinned.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum AffinityPolicy {
    /// Do not change the operating system's worker affinity.
    Inherit,
}

/// A validated, finite policy for an explicitly scheduled operation.
///
/// It has no default: callers must choose every bound when they opt into
/// managed execution. `workers` may not exceed `max_in_flight_tasks`, and the
/// parallel-work threshold may not exceed the in-flight byte ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ExecutionLimits {
    workers: NonZeroUsize,
    max_in_flight_tasks: NonZeroUsize,
    max_in_flight_bytes: NonZeroU64,
    min_parallel_bytes: u64,
    affinity: AffinityPolicy,
}

impl ExecutionLimits {
    /// Creates an execution policy using [`AffinityPolicy::Inherit`].
    ///
    /// # Errors
    ///
    /// Returns an [`ExecutionError`] when the requested worker count cannot
    /// make progress under the task cap, or when the parallel-work threshold
    /// exceeds the finite in-flight byte ceiling.
    pub fn new(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
        max_in_flight_bytes: NonZeroU64,
        min_parallel_bytes: u64,
    ) -> Result<Self, ExecutionError> {
        Self::with_affinity(
            workers,
            max_in_flight_tasks,
            max_in_flight_bytes,
            min_parallel_bytes,
            AffinityPolicy::Inherit,
        )
    }

    /// Creates an execution policy with an explicit affinity policy.
    ///
    /// # Errors
    ///
    /// Returns an [`ExecutionError`] when the requested worker count cannot
    /// make progress under the task cap, or when the parallel-work threshold
    /// exceeds the finite in-flight byte ceiling.
    pub fn with_affinity(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
        max_in_flight_bytes: NonZeroU64,
        min_parallel_bytes: u64,
        affinity: AffinityPolicy,
    ) -> Result<Self, ExecutionError> {
        if workers > max_in_flight_tasks {
            return Err(ExecutionError::WorkersExceedInFlightTasks {
                workers,
                max_in_flight_tasks,
            });
        }
        if min_parallel_bytes > max_in_flight_bytes.get() {
            return Err(ExecutionError::ParallelThresholdExceedsInFlightBytes {
                min_parallel_bytes,
                max_in_flight_bytes,
            });
        }
        Ok(Self {
            workers,
            max_in_flight_tasks,
            max_in_flight_bytes,
            min_parallel_bytes,
            affinity,
        })
    }

    /// Maximum number of workers a runtime adapter may create.
    #[must_use]
    pub const fn workers(self) -> NonZeroUsize {
        self.workers
    }

    /// Maximum tasks that may be outstanding at once.
    #[must_use]
    pub const fn max_in_flight_tasks(self) -> NonZeroUsize {
        self.max_in_flight_tasks
    }

    /// Maximum bytes that may be retained by in-flight work.
    #[must_use]
    pub const fn max_in_flight_bytes(self) -> NonZeroU64 {
        self.max_in_flight_bytes
    }

    /// Smallest aggregate task size for which parallel scheduling is eligible.
    #[must_use]
    pub const fn min_parallel_bytes(self) -> u64 {
        self.min_parallel_bytes
    }

    /// Affinity policy selected by the caller.
    #[must_use]
    pub const fn affinity(self) -> AffinityPolicy {
        self.affinity
    }
}

/// Handle used by an operation owner to request cooperative cancellation.
#[derive(Debug, Clone)]
pub struct CancellationSource {
    cancelled: Arc<AtomicBool>,
}

impl CancellationSource {
    /// Creates a source and the token supplied to an operation.
    #[must_use]
    pub fn pair() -> (Self, CancellationToken) {
        let cancelled = Arc::new(AtomicBool::new(false));
        (
            Self {
                cancelled: Arc::clone(&cancelled),
            },
            CancellationToken { cancelled },
        )
    }

    /// Requests cancellation.
    ///
    /// Requesting cancellation is idempotent. Runtimes must check the paired
    /// token at their defined cooperative interruption points.
    pub fn cancel(&self) {
        self.cancelled.store(true, Ordering::Release);
    }

    /// Returns whether this source has requested cancellation.
    #[must_use]
    pub fn is_cancelled(&self) -> bool {
        self.cancelled.load(Ordering::Acquire)
    }
}

/// Clone-cheap cooperative-cancellation token supplied to an operation.
#[derive(Debug, Clone)]
pub struct CancellationToken {
    cancelled: Arc<AtomicBool>,
}

impl CancellationToken {
    /// Returns whether cancellation has been requested.
    #[must_use]
    pub fn is_cancelled(&self) -> bool {
        self.cancelled.load(Ordering::Acquire)
    }

    /// Returns a typed error when cancellation has been requested.
    ///
    /// # Errors
    ///
    /// Returns [`ExecutionError::Cancelled`] if the paired source has
    /// requested cancellation.
    pub fn check(&self) -> Result<(), ExecutionError> {
        if self.is_cancelled() {
            Err(ExecutionError::Cancelled)
        } else {
            Ok(())
        }
    }
}

/// Explicit runtime-neutral policy and shared resource budget for one class of
/// operations.
///
/// Constructing a context consumes a caller-supplied [`Budget`]. Clone that
/// budget before construction when multiple contexts must share a hierarchical
/// root. The context itself creates no threads, installs no global runtime,
/// and does not choose implicit limits.
#[derive(Debug, Clone)]
pub struct ExecutionContext {
    budget: Budget,
    cancellation: CancellationToken,
    limits: ExecutionLimits,
}

impl ExecutionContext {
    /// Creates a context from a caller-selected budget, token, and limits.
    #[must_use]
    pub fn new(budget: Budget, cancellation: CancellationToken, limits: ExecutionLimits) -> Self {
        Self {
            budget,
            cancellation,
            limits,
        }
    }

    /// Shared hierarchical budget charged by this context.
    #[must_use]
    pub const fn budget(&self) -> &Budget {
        &self.budget
    }

    /// Cooperative-cancellation token for this context.
    #[must_use]
    pub const fn cancellation(&self) -> &CancellationToken {
        &self.cancellation
    }

    /// Validated execution policy for this context.
    #[must_use]
    pub const fn limits(&self) -> ExecutionLimits {
        self.limits
    }

    /// Checks whether cancellation has been requested.
    ///
    /// # Errors
    ///
    /// Returns [`ExecutionError::Cancelled`] when the paired source requested
    /// cancellation.
    pub fn check(&self) -> Result<(), ExecutionError> {
        self.cancellation.check()
    }

    /// Reserves a resource after checking for cancellation.
    ///
    /// The returned reservation releases its charge when dropped.
    ///
    /// # Errors
    ///
    /// Returns [`ExecutionError::Cancelled`] if cancellation was requested, or
    /// [`ExecutionError::ResourceLimit`] if this context's budget or an
    /// ancestor cannot accept the charge.
    pub fn reserve(&self, resource: Resource, amount: u64) -> Result<Reservation, ExecutionError> {
        self.check()?;
        self.budget.reserve(resource, amount).map_err(Into::into)
    }

    /// Consumes cumulative resource capacity after checking for cancellation.
    ///
    /// # Errors
    ///
    /// Returns [`ExecutionError::Cancelled`] if cancellation was requested, or
    /// [`ExecutionError::ResourceLimit`] if this context's budget or an
    /// ancestor cannot accept the charge.
    pub fn consume(&self, resource: Resource, amount: u64) -> Result<(), ExecutionError> {
        self.check()?;
        self.budget.consume(resource, amount).map_err(Into::into)
    }
}

/// Typed execution-policy, cancellation, and budget errors.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ExecutionError {
    /// The operation was cooperatively cancelled before its next check.
    #[error("operation cancelled")]
    Cancelled,

    /// A worker policy would permit fewer outstanding tasks than workers.
    #[error(
        "execution policy has {workers} worker(s) but only {max_in_flight_tasks} in-flight task slot(s)"
    )]
    WorkersExceedInFlightTasks {
        /// Requested worker count.
        workers: NonZeroUsize,
        /// Configured outstanding-task ceiling.
        max_in_flight_tasks: NonZeroUsize,
    },

    /// The parallel-work threshold would never fit in the byte ceiling.
    #[error(
        "execution policy requires {min_parallel_bytes} parallel byte(s), exceeding the {max_in_flight_bytes} in-flight byte ceiling"
    )]
    ParallelThresholdExceedsInFlightBytes {
        /// Minimum aggregate work size eligible for parallel scheduling.
        min_parallel_bytes: u64,
        /// Configured outstanding-byte ceiling.
        max_in_flight_bytes: NonZeroU64,
    },

    /// A hierarchical resource budget rejected the requested charge.
    #[error(transparent)]
    ResourceLimit(#[from] ResourceLimit),
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use std::{
        num::{NonZeroU64, NonZeroUsize},
        sync::{Arc, Barrier},
    };

    use super::*;
    use crate::Limits;

    fn limits() -> ExecutionLimits {
        ExecutionLimits::new(
            NonZeroUsize::new(2).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(1024).unwrap(),
            256,
        )
        .unwrap()
    }

    fn budget(memory: u64) -> Budget {
        Budget::root("root", Limits::new(memory, 100, 100, 100, 100, 100))
    }

    fn context(budget: Budget, token: CancellationToken) -> ExecutionContext {
        ExecutionContext::new(budget, token, limits())
    }

    #[test]
    fn limits_expose_every_explicit_policy_choice() {
        let limits = limits();
        assert_eq!(limits.workers().get(), 2);
        assert_eq!(limits.max_in_flight_tasks().get(), 4);
        assert_eq!(limits.max_in_flight_bytes().get(), 1024);
        assert_eq!(limits.min_parallel_bytes(), 256);
        assert_eq!(limits.affinity(), AffinityPolicy::Inherit);
    }

    #[test]
    fn limits_reject_worker_count_above_task_cap() {
        assert_eq!(
            ExecutionLimits::new(
                NonZeroUsize::new(3).unwrap(),
                NonZeroUsize::new(2).unwrap(),
                NonZeroU64::new(1024).unwrap(),
                0,
            ),
            Err(ExecutionError::WorkersExceedInFlightTasks {
                workers: NonZeroUsize::new(3).unwrap(),
                max_in_flight_tasks: NonZeroUsize::new(2).unwrap(),
            })
        );
    }

    #[test]
    fn limits_reject_parallel_threshold_above_byte_cap() {
        assert_eq!(
            ExecutionLimits::new(
                NonZeroUsize::new(1).unwrap(),
                NonZeroUsize::new(1).unwrap(),
                NonZeroU64::new(7).unwrap(),
                8,
            ),
            Err(ExecutionError::ParallelThresholdExceedsInFlightBytes {
                min_parallel_bytes: 8,
                max_in_flight_bytes: NonZeroU64::new(7).unwrap(),
            })
        );
    }

    #[test]
    fn source_cancellation_is_visible_to_another_thread() {
        let (source, token) = CancellationSource::pair();
        let barrier = Arc::new(Barrier::new(2));
        let waiting = barrier.clone();
        let worker = std::thread::spawn(move || {
            waiting.wait();
            token.check()
        });

        source.cancel();
        barrier.wait();

        assert_eq!(worker.join().unwrap(), Err(ExecutionError::Cancelled));
        assert!(source.is_cancelled());
    }

    #[test]
    fn cancellation_precedes_budget_charges() {
        let (source, token) = CancellationSource::pair();
        let budget = budget(10);
        let context = context(budget.clone(), token);
        source.cancel();

        assert_eq!(context.check(), Err(ExecutionError::Cancelled));
        assert!(matches!(
            context.reserve(Resource::Memory, 1),
            Err(ExecutionError::Cancelled)
        ));
        assert_eq!(
            context.consume(Resource::Work, 1),
            Err(ExecutionError::Cancelled)
        );
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Work), 0);
    }

    #[test]
    fn context_charges_shared_hierarchical_budget() {
        let root = budget(5);
        let child = root.child("operation", Limits::new(10, 100, 100, 100, 100, 100));
        let (_source, token) = CancellationSource::pair();
        let context = context(child.clone(), token);

        let reservation = context.reserve(Resource::Memory, 5).unwrap();
        assert_eq!(root.used(Resource::Memory), 5);
        assert_eq!(child.used(Resource::Memory), 5);
        let error = context
            .reserve(Resource::Memory, 1)
            .expect_err("parent limit must apply through the context");
        assert_eq!(
            error,
            ExecutionError::ResourceLimit(ResourceLimit {
                resource: Resource::Memory,
                observed: 6,
                limit: 5,
                scope: Arc::<str>::from("root"),
            })
        );
        drop(reservation);
        assert_eq!(root.used(Resource::Memory), 0);
        assert_eq!(child.used(Resource::Memory), 0);
    }

    #[test]
    fn public_handles_are_send_and_sync() {
        const fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<CancellationSource>();
        assert_send_sync::<CancellationToken>();
        assert_send_sync::<ExecutionContext>();
        assert_send_sync::<ExecutionLimits>();
    }
}
