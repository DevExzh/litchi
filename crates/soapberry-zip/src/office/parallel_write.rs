//! Bounded, caller-scheduled parallel compression for archive rewrites.
//!
//! A [`ParallelWriteSession`] is the write-side counterpart of
//! [`ParallelReadSession`](super::ParallelReadSession): an explicit, reusable
//! scheduler a caller opts into. It owns at most one local worker pool, builds
//! that pool lazily on the first wave that qualifies, and never initializes or
//! installs Rayon's process-global pool. A caller that supplies its own
//! [`ScopedWorkers`] facility gets no pool at all.
//!
//! The session schedules compression only. Emission order, framing and every
//! byte a member contributes are decided elsewhere and are not affected by how
//! many workers compressed the members.

use std::num::NonZeroUsize;
use std::sync::{Arc, Mutex};

use rayon::prelude::*;

use crate::{Error, ErrorKind};

/// Default smallest changed-set remainder eligible for parallel compression.
///
/// The *remainder* is the changed set's total payload minus its largest
/// member: the work that remains once the longest pole is placed. A set whose
/// remainder is below this floor cannot repay the cost of taking the parallel
/// path, whatever the largest member's size.
///
/// Measured by change 0662 over sixteen changed-set shapes published through
/// this writer: a wave costs about 9 µs of fixed scheduling, so a remainder of
/// 768 bytes loses (0.75× at width 2) while 871 bytes already wins (1.48×).
/// This value is an order of magnitude above that crossover; every shape
/// measured at or above it returned at least 1.77× at width 2.
pub const DEFAULT_MIN_PARALLELIZABLE_BYTES: u64 = 8 * 1024;

/// Default smallest remainder eligible for a wave the session must build a
/// worker pool for.
///
/// Building a pool costs about 120 µs on this host, and a publication owns its
/// session, so a caller that does not attach its own facility pays that once
/// per publication. Change 0662 measured a 28 KiB remainder at 0.93× and 0.85×
/// at widths 4 and 8 on a cold pool against 1.28× and 1.22× on a caller's, and
/// a 60 KiB remainder at 1.19× cold. This threshold keeps the cold path out of
/// the range where the pool costs more than the wave saves.
pub const DEFAULT_MIN_POOL_PARALLELIZABLE_BYTES: u64 = 64 * 1024;

/// Default per-task size floor used to narrow a wave: none.
///
/// The granted width never exceeds `1 + remainder / min_task_bytes`, so a
/// positive floor narrows a set of one dominant member and a few crumbs to a
/// serial wave. Change 0662 measured no floor worth imposing on *deflate*: a
/// changed set of 128 members of 871 bytes each — every task three orders of
/// magnitude below any plausible floor — scales 6.06× at width 8, because the
/// cost a wave must repay belongs to the wave, not to the task. A caller whose
/// own workload disagrees sets its own floor.
pub const DEFAULT_MIN_TASK_BYTES: u64 = 0;

/// Upper estimate of the resident memory one worker of a wave adds, in bytes.
///
/// zlib's documented deflate footprint at window bits 15 and memory level 8 is
/// `(1 << 17) + (1 << 17)` bytes, and a worker also touches its own stack and
/// the wave's bookkeeping. Change 0662 measured peak RSS over a real
/// publication rising 2.1 MiB at width 4 and 5.0 MiB at width 8 — 525 KiB and
/// 625 KiB per worker — so this value is set above the larger of the two. A
/// caller that charges a wave against a memory budget multiplies it by the
/// granted width.
pub const WORKER_STATE_BYTES: u64 = 640 * 1024;

/// A caller-provided facility that runs a bounded set of borrowed tasks.
///
/// This trait is defined here rather than taken from `litchi-core` for the
/// same reason [`CancellationProbe`](super::CancellationProbe) is: this crate
/// is a standalone ZIP crate and depends on no litchi crate. A caller that
/// owns both bridges one to the other.
pub trait ScopedWorkers: Send + Sync + std::fmt::Debug {
    /// Runs every task exactly once and returns only after all have returned.
    ///
    /// Tasks may run on any thread, in any order, concurrently or serially.
    /// Tasks handed to this facility have already caught their own panics, so
    /// an implementation never observes one unwinding out of a task.
    fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]);
}

/// Validated finite limits for a local [`ParallelWriteSession`].
///
/// `workers` is the widest wave the session may run; `max_in_flight_tasks`
/// caps the width independently of the member count. The two byte thresholds
/// decide whether a changed set is worth splitting at all — see
/// [`DEFAULT_MIN_PARALLELIZABLE_BYTES`] and [`DEFAULT_MIN_TASK_BYTES`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParallelWriteLimits {
    workers: NonZeroUsize,
    max_in_flight_tasks: NonZeroUsize,
    min_parallelizable_bytes: u64,
    min_pool_parallelizable_bytes: u64,
    min_task_bytes: u64,
}

impl ParallelWriteLimits {
    /// Creates validated limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the worker count exceeds the in-flight task cap.
    pub fn new(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
        min_parallelizable_bytes: u64,
        min_pool_parallelizable_bytes: u64,
        min_task_bytes: u64,
    ) -> Result<Self, Error> {
        if workers > max_in_flight_tasks {
            return Err(ErrorKind::InvalidParallelWriteLimits {
                reason: "workers must not exceed max_in_flight_tasks",
            }
            .into());
        }
        Ok(Self {
            workers,
            max_in_flight_tasks,
            min_parallelizable_bytes,
            min_pool_parallelizable_bytes,
            min_task_bytes,
        })
    }

    /// Returns this policy with its warm-path remainder threshold replaced.
    #[must_use]
    pub const fn with_min_parallelizable_bytes(mut self, bytes: u64) -> Self {
        self.min_parallelizable_bytes = bytes;
        self
    }

    /// Creates limits with the measured default thresholds.
    ///
    /// # Errors
    ///
    /// Returns an error when the worker count exceeds the in-flight task cap.
    pub fn with_default_thresholds(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
    ) -> Result<Self, Error> {
        Self::new(
            workers,
            max_in_flight_tasks,
            DEFAULT_MIN_PARALLELIZABLE_BYTES,
            DEFAULT_MIN_POOL_PARALLELIZABLE_BYTES,
            DEFAULT_MIN_TASK_BYTES,
        )
    }

    /// Widest wave this policy admits.
    #[must_use]
    pub const fn workers(self) -> NonZeroUsize {
        self.workers
    }

    /// Maximum tasks that may be outstanding at once.
    #[must_use]
    pub const fn max_in_flight_tasks(self) -> NonZeroUsize {
        self.max_in_flight_tasks
    }

    /// Smallest changed-set remainder eligible for a parallel wave when the
    /// workers already exist.
    #[must_use]
    pub const fn min_parallelizable_bytes(self) -> u64 {
        self.min_parallelizable_bytes
    }

    /// Smallest changed-set remainder eligible for a parallel wave that must
    /// first build a worker pool.
    #[must_use]
    pub const fn min_pool_parallelizable_bytes(self) -> u64 {
        self.min_pool_parallelizable_bytes
    }

    /// Per-task size floor used to narrow a wave. Zero imposes no floor.
    #[must_use]
    pub const fn min_task_bytes(self) -> u64 {
        self.min_task_bytes
    }
}

/// Reusable local scheduler for explicit archive rewrites.
///
/// The session is inert until a wave qualifies: constructing one creates no
/// threads. A caller-supplied [`ScopedWorkers`] facility replaces the local
/// pool entirely, so a caller can bound, instrument or serialize every task
/// this crate schedules.
pub struct ParallelWriteSession {
    limits: ParallelWriteLimits,
    workers: Option<Arc<dyn ScopedWorkers>>,
    pool: Mutex<Option<Arc<rayon::ThreadPool>>>,
}

impl ParallelWriteSession {
    /// Creates a session that builds its own pool lazily.
    #[must_use]
    pub const fn new(limits: ParallelWriteLimits) -> Self {
        Self {
            limits,
            workers: None,
            pool: Mutex::new(None),
        }
    }

    /// Creates a session that runs every task on a caller-provided facility.
    #[must_use]
    pub fn with_scoped_workers(
        limits: ParallelWriteLimits,
        workers: Arc<dyn ScopedWorkers>,
    ) -> Self {
        Self {
            limits,
            workers: Some(workers),
            pool: Mutex::new(None),
        }
    }

    /// Validated policy used by this session.
    #[must_use]
    pub const fn limits(&self) -> ParallelWriteLimits {
        self.limits
    }

    /// Caller-provided worker facility, if one is attached.
    #[must_use]
    pub fn scoped_workers(&self) -> Option<&Arc<dyn ScopedWorkers>> {
        self.workers.as_ref()
    }

    /// Worker count of the local pool, or zero while none has been built.
    ///
    /// A session that never ran a qualifying wave, and a session running on a
    /// caller-provided facility, both report zero.
    #[must_use]
    pub fn local_pool_worker_count(&self) -> usize {
        self.pool.lock().map_or(0, |pool| {
            pool.as_ref().map_or(0, |pool| pool.current_num_threads())
        })
    }

    /// Decides the wave this session would run for `plan`.
    ///
    /// This is [`crate::PreservationPlan::deflate_wave`] with the one threshold
    /// the session can choose: a session that already holds its workers — a
    /// caller's facility, or a pool an earlier wave built — admits a wave from
    /// [`ParallelWriteLimits::min_parallelizable_bytes`], while one that must
    /// build a pool first admits from the larger
    /// [`ParallelWriteLimits::min_pool_parallelizable_bytes`], because that
    /// wave has to repay the pool as well as itself.
    #[must_use]
    pub fn wave_for(&self, plan: &crate::PreservationPlan) -> crate::DeflateWave {
        let limits = if self.workers.is_some() || self.local_pool_worker_count() > 0 {
            self.limits
        } else {
            self.limits
                .with_min_parallelizable_bytes(self.limits.min_pool_parallelizable_bytes)
        };
        plan.deflate_wave(limits)
    }

    /// Runs `tasks` on at most `width` workers, in an unspecified order.
    ///
    /// A width of one, a single task, or an empty task list runs the tasks in
    /// place on the calling thread and builds no pool. Every task has already
    /// caught its own panics, so this never unwinds through a worker.
    pub(crate) fn run_tasks(
        &self,
        width: NonZeroUsize,
        tasks: &mut [&mut (dyn FnMut() + Send)],
    ) -> Result<(), Error> {
        let width = width.get().min(self.limits.workers.get());
        if width <= 1 || tasks.len() <= 1 {
            for task in tasks.iter_mut() {
                task();
            }
            return Ok(());
        }
        if let Some(workers) = self.workers.as_ref() {
            workers.run_all(tasks);
            return Ok(());
        }
        let pool = self.pool(width)?;
        pool.install(|| tasks.par_iter_mut().for_each(|task| task()));
        Ok(())
    }

    fn pool(&self, workers: usize) -> Result<Arc<rayon::ThreadPool>, Error> {
        let mut cached = self.pool.lock().map_err(|_error| {
            Error::from(ErrorKind::ParallelWriteWorkerPool {
                workers,
                message: "local write pool cache is poisoned".to_string(),
            })
        })?;
        if let Some(pool) = cached.as_ref() {
            return Ok(Arc::clone(pool));
        }
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(workers)
            .build()
            .map_err(|error| ErrorKind::ParallelWriteWorkerPool {
                workers,
                message: error.to_string(),
            })?;
        let pool = Arc::new(pool);
        *cached = Some(Arc::clone(&pool));
        Ok(pool)
    }
}

impl std::fmt::Debug for ParallelWriteSession {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ParallelWriteSession")
            .field("limits", &self.limits)
            .field("uses_local_pool", &(self.local_pool_worker_count() > 0))
            .field("uses_caller_workers", &self.workers.is_some())
            .finish()
    }
}

/// A [`ScopedWorkers`] facility that runs every task on the calling thread.
///
/// It exists so a caller, a test or a determinism oracle can force every
/// scheduled task onto one thread without changing any other policy.
#[derive(Debug, Clone, Copy, Default)]
pub struct SerialScopedWorkers;

impl ScopedWorkers for SerialScopedWorkers {
    fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]) {
        for task in tasks.iter_mut() {
            task();
        }
    }
}
