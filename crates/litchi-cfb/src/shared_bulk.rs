//! Explicit, local-runtime bulk reads for [`crate::SharedOleFile`].

use crate::{OleError, SharedOleFile};
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use rayon::prelude::*;
use std::{
    fmt,
    panic::{AssertUnwindSafe, catch_unwind},
    sync::{Arc, Mutex},
};

/// Error from one bounded [`SharedOleBulkRead`] operation.
#[derive(Debug)]
pub enum SharedOleBulkError {
    /// A CFB lookup, payload, or source-version check failed.
    Ole(OleError),
    /// Cancellation or the caller's resource budget rejected the operation.
    Execution(ExecutionError),
    /// One stream cannot fit within the configured in-flight byte bound.
    StreamExceedsInFlightBytes {
        /// Declared byte length of the requested stream.
        declared: u64,
        /// Maximum bytes allowed in one scheduled batch.
        maximum: u64,
    },
    /// The crate-local worker pool could not be constructed.
    Scheduler(String),
}

impl fmt::Display for SharedOleBulkError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => error.fmt(formatter),
            Self::Execution(error) => error.fmt(formatter),
            Self::StreamExceedsInFlightBytes { declared, maximum } => write!(
                formatter,
                "CFB stream declares {declared} bytes, exceeding the {maximum}-byte in-flight limit"
            ),
            Self::Scheduler(error) => write!(formatter, "CFB local worker pool failed: {error}"),
        }
    }
}

impl std::error::Error for SharedOleBulkError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Ole(error) => Some(error),
            Self::Execution(error) => Some(error),
            Self::StreamExceedsInFlightBytes { .. } | Self::Scheduler(_) => None,
        }
    }
}

impl From<OleError> for SharedOleBulkError {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

impl From<ExecutionError> for SharedOleBulkError {
    fn from(error: ExecutionError) -> Self {
        Self::Execution(error)
    }
}

/// Reusable, caller-configured local bulk-read session for [`SharedOleFile`].
///
/// The supplied [`ExecutionContext`] selects worker count, positional-read
/// concurrency, maximum outstanding tasks and bytes, cooperative cancellation,
/// the parallel threshold, and the `Inherit` affinity policy. The
/// implementation creates a private Rayon pool only for eligible multi-worker
/// calls; it never installs or uses Rayon’s global pool. A caller facility
/// attached to the context replaces that private pool.
pub struct SharedOleBulkRead<'file> {
    file: &'file SharedOleFile,
    context: ExecutionContext,
    /// Successful construction is retained for this session. A failed build
    /// leaves this empty so a later eligible call can retry.
    pool: Mutex<Option<Arc<rayon::ThreadPool>>>,
    worker_permits: Mutex<Option<WorkerPermits>>,
    #[cfg(test)]
    pool_builds: std::sync::atomic::AtomicUsize,
}

impl<'file> SharedOleBulkRead<'file> {
    pub(crate) fn new(file: &'file SharedOleFile, context: ExecutionContext) -> Self {
        Self {
            file,
            context,
            pool: Mutex::new(None),
            worker_permits: Mutex::new(None),
            #[cfg(test)]
            pool_builds: std::sync::atomic::AtomicUsize::new(0),
        }
    }

    /// Reads requested streams in the same order as `paths`.
    ///
    /// All names and declared lengths are checked before the first payload
    /// read. Work is then partitioned into batches bounded by the supplied
    /// task/byte limits. No partially completed collection is returned on a
    /// stream, source-change, cancellation, or resource-budget error.
    ///
    /// # Errors
    ///
    /// Returns an error if a path is invalid, cancellation is requested, a
    /// requested stream exceeds the in-flight byte ceiling, source data
    /// changes, or the context budget cannot reserve a scheduled batch.
    pub fn read_streams(&self, paths: &[&[&str]]) -> Result<Vec<Vec<u8>>, SharedOleBulkError> {
        self.context.check()?;
        let limits = self.context.limits();
        let max_tasks = limits.max_in_flight_tasks().get();
        let max_bytes = limits.max_in_flight_bytes().get();

        let mut requests = Vec::new();
        requests
            .try_reserve_exact(paths.len())
            .map_err(|source| OleError::allocation("shared bulk stream requests", source))?;
        let mut total_bytes = 0u64;
        for &path in paths {
            self.context.check()?;
            let entry = self.file.find_entry(path)?;
            if entry.entry_type != crate::consts::STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()).into());
            }
            let size = entry.size;
            if size > max_bytes {
                return Err(Self::too_large(size, max_bytes));
            }
            total_bytes = total_bytes.checked_add(size).ok_or_else(|| {
                OleError::InvalidData("shared bulk stream sizes overflow u64".to_string())
            })?;
            requests.push(StreamRequest {
                path,
                size,
                is_minifat: entry.is_minifat,
            });
        }
        self.context.check()?;
        // Work units are declared stream bytes. Charge the complete request
        // before any payload read so a rejected budget has no read side effect.
        self.context.consume(Resource::Work, total_bytes)?;
        self.context
            .consume(Resource::CpuTasks, requests.len() as u64)?;

        let mut results = Vec::new();
        results
            .try_reserve_exact(requests.len())
            .map_err(|source| OleError::allocation("shared bulk stream results", source))?;
        let mut next = 0usize;
        while next < requests.len() {
            self.context.check()?;
            let (end, batch_bytes) = batch_end(&requests, next, max_tasks, max_bytes)?;
            // The reservation both checks the hierarchical caller budget and
            // makes the declared input retained by this scheduled batch an
            // explicit bounded resource. It drops before the next batch.
            let _in_flight = self.context.reserve(Resource::Memory, batch_bytes)?;
            let batch = &requests[next..end];
            // A batch with multiple MiniFAT requests already needs a shared
            // convergence point: serial execution would otherwise consume a
            // direct read for the first target and cache for the next, while
            // parallel execution would make the choice scheduler-dependent.
            // A one-item MiniFAT batch keeps the prior bounded direct behavior
            // and therefore does not retain an unaccounted root cache.
            let force_minifat_cache = batch.iter().filter(|request| request.is_minifat).count() > 1;
            let parallel_candidate = limits.workers().get() > 1
                && batch.len() > 1
                && batch_bytes >= limits.min_parallel_bytes()
                && (limits.min_task_bytes() == 0
                    || batch
                        .iter()
                        .all(|request| request.size >= limits.min_task_bytes()));
            let requested_width = limits
                .workers()
                .get()
                .min(limits.max_in_flight_tasks().get())
                .min(batch.len());
            let admission = if parallel_candidate && requested_width > 1 {
                self.admit_workers(requested_width)?
            } else {
                self.admit_serial()?
            };
            let width = admission.width;
            let _workers = admission.workers;
            let _io = admission.io;
            let batch_results = if width > 1 {
                self.read_parallel_batch(batch, width, force_minifat_cache)?
            } else {
                batch
                    .iter()
                    .map(|request| self.read_one(request, force_minifat_cache))
                    .collect()
            };
            self.context.check()?;
            for result in batch_results {
                results.push(result?);
            }
            next = end;
        }
        Ok(results)
    }

    fn read_one(
        &self,
        request: &StreamRequest<'_>,
        force_minifat_cache: bool,
    ) -> Result<Vec<u8>, SharedOleBulkError> {
        self.context.check()?;
        let result = if force_minifat_cache && request.is_minifat {
            // Multi-target MiniFAT batches deliberately converge on one root
            // cache. Single-target batches retain target-aware direct reads.
            self.file.open_stream_force_cache(request.path)
        } else {
            self.file.open_stream(request.path)
        };
        self.context.check()?;
        result.map_err(Into::into)
    }

    fn too_large(declared: u64, maximum: u64) -> SharedOleBulkError {
        SharedOleBulkError::StreamExceedsInFlightBytes { declared, maximum }
    }

    fn admit_serial(&self) -> Result<WorkerAdmission, SharedOleBulkError> {
        if let Ok(permits) = self.worker_permits.lock() {
            if permits.is_some() {
                let io = self.context.reserve(Resource::IoConcurrency, 1)?;
                return Ok(WorkerAdmission {
                    width: 1,
                    workers: None,
                    io,
                });
            }
        } else {
            return Err(SharedOleBulkError::Scheduler(
                "worker admission cache is poisoned".to_string(),
            ));
        }
        let workers = self.context.reserve(Resource::Workers, 1)?;
        let io = match self.context.reserve(Resource::IoConcurrency, 1) {
            Ok(io) => io,
            Err(error) => {
                drop(workers);
                return Err(error.into());
            },
        };
        Ok(WorkerAdmission {
            width: 1,
            workers: Some(workers),
            io,
        })
    }

    fn admit_workers(&self, requested: usize) -> Result<WorkerAdmission, SharedOleBulkError> {
        let mut permits = self.worker_permits.lock().map_err(|_error| {
            SharedOleBulkError::Scheduler("worker admission cache is poisoned".to_string())
        })?;
        if let Some(permits) = permits.as_ref() {
            let maximum = permits.width.min(requested);
            let mut last_limit = None;
            for width in (1..=maximum).rev() {
                match self.context.reserve(Resource::IoConcurrency, width as u64) {
                    Ok(io) => {
                        return Ok(WorkerAdmission {
                            width,
                            workers: None,
                            io,
                        });
                    },
                    Err(ExecutionError::ResourceLimit(limit))
                        if limit.resource == Resource::IoConcurrency =>
                    {
                        last_limit = Some(limit);
                    },
                    Err(error) => return Err(error.into()),
                }
            }
            if let Some(limit) = last_limit {
                return Err(ExecutionError::ResourceLimit(limit).into());
            }
        }
        for width in (2..=requested).rev() {
            let workers = match self.context.reserve(Resource::Workers, width as u64) {
                Ok(workers) => workers,
                Err(ExecutionError::ResourceLimit(limit))
                    if limit.resource == Resource::Workers =>
                {
                    continue;
                },
                Err(error) => return Err(error.into()),
            };
            let io = match self.context.reserve(Resource::IoConcurrency, width as u64) {
                Ok(io) => io,
                Err(ExecutionError::ResourceLimit(limit))
                    if limit.resource == Resource::IoConcurrency =>
                {
                    drop(workers);
                    continue;
                },
                Err(error) => {
                    drop(workers);
                    return Err(error.into());
                },
            };
            if self.context.scoped_workers().is_none() {
                // Construct the private pool while the reservation is local;
                // a failed build releases it before the error is returned.
                if let Err(error) = self.pool(width) {
                    drop(io);
                    drop(workers);
                    return Err(error);
                }
            }
            if self.context.scoped_workers().is_some() {
                return Ok(WorkerAdmission {
                    width,
                    workers: Some(workers),
                    io,
                });
            }
            *permits = Some(WorkerPermits {
                width,
                _workers: workers,
            });
            return Ok(WorkerAdmission {
                width,
                workers: None,
                io,
            });
        }
        drop(permits);
        self.admit_serial()
    }

    fn read_parallel_batch(
        &self,
        batch: &[StreamRequest<'_>],
        width: usize,
        force_minifat_cache: bool,
    ) -> Result<Vec<Result<Vec<u8>, SharedOleBulkError>>, SharedOleBulkError> {
        let mut slots: Vec<Option<Result<Vec<u8>, SharedOleBulkError>>> = Vec::new();
        slots
            .try_reserve_exact(batch.len())
            .map_err(|error| SharedOleBulkError::Scheduler(error.to_string()))?;
        slots.resize_with(batch.len(), || None);
        {
            let mut tasks: Vec<Box<dyn FnMut() + Send + '_>> = batch
                .iter()
                .zip(slots.iter_mut())
                .map(|(request, slot)| {
                    let task: Box<dyn FnMut() + Send + '_> = Box::new(move || {
                        let result = catch_unwind(AssertUnwindSafe(|| {
                            self.read_one(request, force_minifat_cache)
                        }))
                        .unwrap_or_else(|_| {
                            Err(SharedOleBulkError::Scheduler(
                                "CFB bulk worker panicked".to_string(),
                            ))
                        });
                        *slot = Some(result);
                    });
                    task
                })
                .collect();
            let mut task_refs: Vec<&mut (dyn FnMut() + Send + '_)> =
                tasks.iter_mut().map(|task| &mut **task).collect();
            if let Some(workers) = self.context.scoped_workers() {
                for wave in task_refs.chunks_mut(width) {
                    workers.run_all(wave);
                }
            } else {
                let pool = self.pool(width)?;
                pool.install(|| task_refs.par_iter_mut().for_each(|task| task()));
            }
        }
        Ok(slots
            .into_iter()
            .map(|result| {
                result.unwrap_or_else(|| {
                    Err(SharedOleBulkError::Scheduler(
                        "CFB bulk worker did not produce a result".to_string(),
                    ))
                })
            })
            .collect())
    }

    fn pool(&self, workers: usize) -> Result<Arc<rayon::ThreadPool>, SharedOleBulkError> {
        let mut cached = self.pool.lock().map_err(|_error| {
            SharedOleBulkError::Scheduler("pool cache is poisoned".to_string())
        })?;
        if let Some(pool) = cached.as_ref() {
            return Ok(Arc::clone(pool));
        }
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(workers)
            .build()
            .map_err(|error| SharedOleBulkError::Scheduler(error.to_string()))?;
        let pool = Arc::new(pool);
        #[cfg(test)]
        self.pool_builds
            .fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        *cached = Some(Arc::clone(&pool));
        Ok(pool)
    }

    #[cfg(test)]
    pub(crate) fn pool_build_count(&self) -> usize {
        self.pool_builds.load(std::sync::atomic::Ordering::Relaxed)
    }
}

#[derive(Debug)]
struct WorkerPermits {
    width: usize,
    _workers: Reservation,
}

struct WorkerAdmission {
    width: usize,
    workers: Option<Reservation>,
    io: Reservation,
}

#[derive(Clone, Copy)]
struct StreamRequest<'path> {
    path: &'path [&'path str],
    size: u64,
    is_minifat: bool,
}

fn batch_end(
    requests: &[StreamRequest<'_>],
    start: usize,
    max_tasks: usize,
    max_bytes: u64,
) -> Result<(usize, u64), SharedOleBulkError> {
    let mut end = start;
    let mut bytes = 0u64;
    while end < requests.len() && end - start < max_tasks {
        let request = requests[end];
        let next = bytes.checked_add(request.size).ok_or_else(|| {
            OleError::InvalidData("shared bulk batch size overflows u64".to_string())
        })?;
        if end > start && next > max_bytes {
            break;
        }
        if next > max_bytes {
            return Err(SharedOleBulkRead::too_large(request.size, max_bytes));
        }
        bytes = next;
        end += 1;
    }
    Ok((end, bytes))
}
