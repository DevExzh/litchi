//! Bounded, opt-in source-backed multi-Part reads.
//!
//! This module owns only orchestration.  Source freshness, ZIP validation,
//! cache admission, payload reservations, and cumulative physical accounting
//! remain in [`super::SourceBackedPackage`]'s one-Part path.

use super::{PartData, SourceBackedPackage};
use crate::error::{OpcError, Result};
use crate::limits::ReadResource;
use crate::packuri::PackURI;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use soapberry_zip::office::EntryId;
use std::mem::size_of;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::mpsc::{self, Receiver, SyncSender};
use std::thread;

// Keep the platform's ordinary scoped-worker depth envelope conservative. The
// reservation is explicit even though the OS owns the actual stack mapping.
const THREAD_STACK_BYTES: usize = 2 * 1024 * 1024;
const SCHEDULER_CONTROL_BYTES: usize = 1024;
// Conservative per-worker allowance for the two bounded std channels and
// their cacheline-padded control/wait state. This is an admission envelope,
// not a portable exact allocator-size claim.
const WORKER_CHANNEL_CONTROL_BYTES: usize = 4096;

/// An ordered collection returned by [`SourceBackedPackage::read_parts_ordered`].
///
/// The collection retains the managed reservation for its `PartData` slots
/// until it is dropped.  Payload reservations are owned independently by the
/// individual `PartData` handles and the package cache.  The wrapper exposes
/// borrowing accessors instead of an unbudgeted `Vec` escape, so callers keep
/// the collection's accounting attached to the returned owner.
#[derive(Debug)]
pub struct PartBatch {
    parts: Vec<PartData>,
    _memory: Option<Reservation>,
    _objects: Option<Reservation>,
}

impl PartBatch {
    fn empty() -> Self {
        Self {
            parts: Vec::new(),
            _memory: None,
            _objects: None,
        }
    }

    fn new(
        parts: Vec<PartData>,
        memory: Option<Reservation>,
        objects: Option<Reservation>,
    ) -> Self {
        Self {
            parts,
            _memory: memory,
            _objects: objects,
        }
    }

    /// Number of requested Part occurrences in this batch.
    #[must_use]
    pub fn len(&self) -> usize {
        self.parts.len()
    }

    /// Whether the batch contains no requested Part occurrences.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
    }

    /// Borrow the ordered Part handles.
    #[must_use]
    pub fn as_slice(&self) -> &[PartData] {
        &self.parts
    }

    /// Borrow one ordered Part handle by occurrence index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&PartData> {
        self.parts.get(index)
    }

    /// Iterate over the ordered Part handles without detaching their owner.
    pub fn iter(&self) -> std::slice::Iter<'_, PartData> {
        self.parts.iter()
    }
}

impl AsRef<[PartData]> for PartBatch {
    fn as_ref(&self) -> &[PartData] {
        self.as_slice()
    }
}

#[derive(Clone, Copy)]
struct PreparedRequest {
    index: usize,
    entry_id: EntryId,
    declared_bytes: Option<u64>,
}

struct PreparedBatch {
    requests: Vec<PreparedRequest>,
    declared_total: u64,
    all_fit_inflight: bool,
    _memory: Option<Reservation>,
    _objects: Option<Reservation>,
}

struct OutputAdmission {
    charged_bytes: usize,
    memory: Option<Reservation>,
    objects: Option<Reservation>,
}

struct SchedulerAdmission {
    handle_bytes: usize,
    command_vec_bytes: usize,
    reply_vec_bytes: usize,
    ordinal_vec_bytes: usize,
    _workers: Reservation,
    _io: Reservation,
    _memory: Reservation,
    _objects: Reservation,
}

enum WorkerCommand {
    Read {
        ordinal: usize,
        request: PreparedRequest,
    },
    Shutdown,
}

/// Read a bounded ordered request set through the package's existing Part
/// cache and source fences.
pub(super) fn read_parts_ordered(
    package: &SourceBackedPackage,
    partnames: &[PackURI],
) -> Result<PartBatch> {
    let context = package.cache.context().cloned();
    fence(package)?;

    if partnames.is_empty() {
        return finish(package, Ok(PartBatch::empty()));
    }

    if let Err(error) = check_request_count(package, partnames.len()) {
        return finish(package, Err(error));
    }

    let prepared = match PreparedBatch::build(package, partnames, context.as_ref()) {
        Ok(prepared) => prepared,
        Err(error) => return finish(package, Err(error)),
    };

    if let Some(context) = context.as_ref()
        && let Err(error) = context.consume(Resource::CpuTasks, prepared.requests.len() as u64)
    {
        drop(prepared);
        return finish(package, Err(super::map_execution_error(error)));
    }

    let output = match OutputAdmission::reserve(context.as_ref(), prepared.requests.len()) {
        Ok(output) => output,
        Err(error) => {
            drop(prepared);
            return finish(package, Err(error));
        },
    };
    let mut parts = match output.allocate(prepared.requests.len()) {
        Ok(parts) => parts,
        Err(error) => {
            drop(output);
            drop(prepared);
            return finish(package, Err(error));
        },
    };

    let workers = context.as_ref().map_or(0, |context| {
        let limits = context.limits();
        limits
            .workers()
            .get()
            .min(limits.max_in_flight_tasks().get())
            .min(prepared.requests.len())
    });
    let parallel = context.as_ref().is_some_and(|context| {
        let limits = context.limits();
        workers > 1
            && prepared.requests.len() > 1
            && prepared.declared_total >= limits.min_parallel_bytes()
            && (limits.min_task_bytes() == 0
                || prepared.requests.iter().all(|request| {
                    request
                        .declared_bytes
                        .is_some_and(|bytes| bytes >= limits.min_task_bytes())
                }))
            && prepared.all_fit_inflight
    });

    let result = if parallel {
        let context = context
            .as_ref()
            .ok_or_else(|| batch_refusal("parallel scheduling requires an execution context"))?;
        let caller_workers = context.scoped_workers().is_some();
        let mut selected = None;
        for width in (2..=workers).rev() {
            let reuse_workers = wave_end(
                &prepared.requests,
                0,
                width,
                context.limits().max_in_flight_bytes().get(),
            )
            .map(|end| end < prepared.requests.len())
            .unwrap_or(false);
            match SchedulerAdmission::reserve(context, width, reuse_workers, caller_workers) {
                Ok(scheduler) => {
                    selected = Some((width, scheduler));
                    break;
                },
                Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                    if matches!(
                        limit.resource,
                        Resource::Workers
                            | Resource::IoConcurrency
                            | Resource::Memory
                            | Resource::Objects
                    ) => {},
                Err(error) => {
                    drop(parts);
                    drop(output);
                    drop(prepared);
                    return finish(package, Err(error));
                },
            }
        }
        if let Some((width, scheduler)) = selected {
            let result = read_parallel(
                package,
                context,
                &prepared.requests,
                width,
                &scheduler,
                &mut parts,
            );
            drop(scheduler);
            result
        } else {
            let workers = context
                .reserve(Resource::Workers, 1)
                .map_err(super::map_execution_error);
            let io = match workers {
                Ok(workers) => match context.reserve(Resource::IoConcurrency, 1) {
                    Ok(io) => Some((workers, io)),
                    Err(error) => {
                        drop(workers);
                        return finish(package, Err(super::map_execution_error(error)));
                    },
                },
                Err(error) => return finish(package, Err(error)),
            };
            let result = read_serial(package, &prepared.requests, &mut parts);
            drop(io);
            result
        }
    } else {
        let admission = if let Some(context) = context.as_ref() {
            let workers = context
                .reserve(Resource::Workers, 1)
                .map_err(super::map_execution_error);
            match workers {
                Ok(workers) => match context.reserve(Resource::IoConcurrency, 1) {
                    Ok(io) => Some((workers, io)),
                    Err(error) => {
                        drop(workers);
                        drop(parts);
                        drop(output);
                        drop(prepared);
                        return finish(package, Err(super::map_execution_error(error)));
                    },
                },
                Err(error) => {
                    drop(parts);
                    drop(output);
                    drop(prepared);
                    return finish(package, Err(error));
                },
            }
        } else {
            None
        };
        let result = read_serial(package, &prepared.requests, &mut parts);
        drop(admission);
        result
    };

    drop(prepared);
    match result {
        Ok(()) => finish(package, Ok(output.into_batch(parts))),
        Err(error) => {
            drop(parts);
            drop(output);
            finish(package, Err(error))
        },
    }
}

impl PreparedBatch {
    fn build(
        package: &SourceBackedPackage,
        names: &[PackURI],
        context: Option<&ExecutionContext>,
    ) -> Result<Self> {
        let descriptor_bytes = checked_collection_bytes(
            names.len(),
            size_of::<PreparedRequest>(),
            "source-backed batch descriptors",
        )?;
        let (memory, objects) = reserve_shape(context, descriptor_bytes, 1)?;
        let mut requests = Vec::new();
        requests
            .try_reserve_exact(names.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch descriptors",
                source,
            })?;
        capacity_fits::<PreparedRequest>(
            requests.capacity(),
            descriptor_bytes,
            "source-backed batch descriptors",
        )?;

        let max_inflight = context.map_or(u64::MAX, |context| {
            context.limits().max_in_flight_bytes().get()
        });
        let mut declared_total = 0_u64;
        let mut all_fit_inflight = true;
        for name in names {
            fence(package)?;
            let index = resolve_part_index(package, name, context)?;
            let entry_id = package
                .parts
                .get(index)
                .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?
                .entry_id;
            let declared_bytes = if context.is_some() {
                let declared = declared_size(package, entry_id)?;
                package.limits.check(
                    ReadResource::PartBytes,
                    declared,
                    package.limits.max_part_bytes(),
                )?;
                declared_total = declared_total
                    .checked_add(declared)
                    .ok_or_else(|| batch_refusal("prepared batch byte sum overflows"))?;
                all_fit_inflight &= declared <= max_inflight;
                Some(declared)
            } else {
                None
            };
            requests.push(PreparedRequest {
                index,
                entry_id,
                declared_bytes,
            });
        }
        fence(package)?;
        Ok(Self {
            requests,
            declared_total,
            all_fit_inflight,
            _memory: memory,
            _objects: objects,
        })
    }
}

impl OutputAdmission {
    fn reserve(context: Option<&ExecutionContext>, count: usize) -> Result<Self> {
        let charged_bytes =
            checked_collection_bytes(count, size_of::<PartData>(), "source-backed batch output")?;
        let Some(context) = context else {
            return Ok(Self {
                charged_bytes: 0,
                memory: None,
                objects: None,
            });
        };
        let memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(charged_bytes)
                    .map_err(|_| batch_refusal("batch output exceeds u64"))?,
            )
            .map_err(super::map_execution_error)?;
        let objects = match context.reserve(Resource::Objects, 1) {
            Ok(objects) => objects,
            Err(error) => {
                drop(memory);
                return Err(super::map_execution_error(error));
            },
        };
        Ok(Self {
            charged_bytes,
            memory: Some(memory),
            objects: Some(objects),
        })
    }

    fn allocate(&self, count: usize) -> Result<Vec<PartData>> {
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(count)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch output",
                source,
            })?;
        if self.memory.is_some() {
            capacity_fits::<PartData>(
                parts.capacity(),
                self.charged_bytes,
                "source-backed batch output",
            )?;
        }
        Ok(parts)
    }

    fn into_batch(self, parts: Vec<PartData>) -> PartBatch {
        PartBatch::new(parts, self.memory, self.objects)
    }
}

impl SchedulerAdmission {
    fn reserve(
        context: &ExecutionContext,
        workers: usize,
        reuse_workers: bool,
        caller_workers: bool,
    ) -> Result<Self> {
        let handle_bytes = if caller_workers {
            0
        } else {
            checked_collection_bytes(
                workers,
                if reuse_workers {
                    size_of::<(usize, thread::ScopedJoinHandle<'static, ()>)>()
                } else {
                    size_of::<(usize, thread::ScopedJoinHandle<'static, Result<PartData>>)>()
                },
                "source-backed batch worker handles",
            )?
        };
        let (wave_bytes, command_vec_bytes, reply_vec_bytes, ordinal_vec_bytes) = if reuse_workers
            && !caller_workers
        {
            let command_message_bytes = checked_collection_bytes(
                workers,
                size_of::<WorkerCommand>(),
                "source-backed batch worker commands",
            )?;
            let reply_message_bytes = checked_collection_bytes(
                workers,
                size_of::<Result<PartData>>(),
                "source-backed batch worker results",
            )?;
            let command_vec_bytes = checked_collection_bytes(
                workers,
                size_of::<SyncSender<WorkerCommand>>(),
                "source-backed batch worker command senders",
            )?;
            let reply_vec_bytes = checked_collection_bytes(
                workers,
                size_of::<Receiver<Result<PartData>>>(),
                "source-backed batch worker reply receivers",
            )?;
            let ordinal_vec_bytes = checked_collection_bytes(
                workers,
                size_of::<Option<usize>>(),
                "source-backed batch worker ordinals",
            )?;
            let endpoint_per_worker = size_of::<SyncSender<WorkerCommand>>()
                .checked_add(size_of::<Receiver<WorkerCommand>>())
                .and_then(|bytes| bytes.checked_add(size_of::<SyncSender<Result<PartData>>>()))
                .and_then(|bytes| bytes.checked_add(size_of::<Receiver<Result<PartData>>>()))
                .ok_or_else(|| batch_refusal("source-backed batch worker endpoints overflow"))?;
            let endpoint_bytes = checked_collection_bytes(
                workers,
                endpoint_per_worker,
                "source-backed batch worker endpoints",
            )?;
            let channel_control_bytes = checked_collection_bytes(
                workers,
                WORKER_CHANNEL_CONTROL_BYTES,
                "source-backed batch worker channels",
            )?;
            let wave_bytes = command_message_bytes
                .checked_add(reply_message_bytes)
                .and_then(|bytes| bytes.checked_add(endpoint_bytes))
                .and_then(|bytes| bytes.checked_add(channel_control_bytes))
                .and_then(|bytes| bytes.checked_add(ordinal_vec_bytes))
                .ok_or_else(|| batch_refusal("source-backed batch worker channels overflow"))?;
            (
                wave_bytes,
                command_vec_bytes,
                reply_vec_bytes,
                ordinal_vec_bytes,
            )
        } else {
            (0, 0, 0, 0)
        };
        let stack_bytes = if caller_workers {
            0
        } else {
            workers
                .checked_mul(THREAD_STACK_BYTES)
                .ok_or_else(|| batch_refusal("source-backed batch worker stacks overflow"))?
        };
        let control_bytes = SCHEDULER_CONTROL_BYTES
            .checked_add(handle_bytes)
            .and_then(|bytes| bytes.checked_add(wave_bytes))
            .and_then(|bytes| bytes.checked_add(stack_bytes))
            .ok_or_else(|| batch_refusal("source-backed batch scheduler overflows"))?;
        let worker_permits = context
            .reserve(Resource::Workers, workers as u64)
            .map_err(super::map_execution_error)?;
        let io = match context.reserve(Resource::IoConcurrency, workers as u64) {
            Ok(io) => io,
            Err(error) => {
                drop(worker_permits);
                return Err(super::map_execution_error(error));
            },
        };
        let memory = match context.reserve(
            Resource::Memory,
            u64::try_from(control_bytes)
                .map_err(|_| batch_refusal("batch scheduler exceeds u64"))?,
        ) {
            Ok(memory) => memory,
            Err(error) => {
                drop(io);
                drop(worker_permits);
                return Err(super::map_execution_error(error));
            },
        };
        let channel_objects = if reuse_workers && !caller_workers {
            workers
                .checked_mul(4)
                .ok_or_else(|| batch_refusal("source-backed batch channel objects overflow"))?
        } else {
            0
        };
        let object_count = u64::try_from(
            workers
                .checked_add(1)
                .and_then(|count| count.checked_add(channel_objects))
                .ok_or_else(|| batch_refusal("batch scheduler objects overflow"))?,
        )
        .map_err(|_| batch_refusal("batch scheduler objects exceed u64"))?;
        let objects = match context.reserve(Resource::Objects, object_count) {
            Ok(objects) => objects,
            Err(error) => {
                drop(memory);
                drop(io);
                drop(worker_permits);
                return Err(super::map_execution_error(error));
            },
        };
        Ok(Self {
            handle_bytes,
            command_vec_bytes,
            reply_vec_bytes,
            ordinal_vec_bytes,
            _workers: worker_permits,
            _io: io,
            _memory: memory,
            _objects: objects,
        })
    }
}

fn read_serial(
    package: &SourceBackedPackage,
    requests: &[PreparedRequest],
    output: &mut Vec<PartData>,
) -> Result<()> {
    // The serial wave is one sequential operation over distinct members, so an
    // unmanaged package long enough to repay the retained workspace may reset
    // one Deflate decoder across it instead of building one per cold member.
    // Store members bypass the decoder, cache hits stay cache-only, and a
    // managed package keeps the one-shot path so that no decoder workspace
    // outlives a managed load (change 0402).
    let mut session = package.sequential_read_session(requests.len());
    for request in requests {
        let data = match request.declared_bytes {
            Some(declared) => package.read_part_prepared_with_session(
                request.index,
                request.entry_id,
                declared,
                session.as_mut(),
            )?,
            None => match session.as_mut() {
                Some(session) => package.read_part_with_session(request.index, session, None)?,
                None => package.read_part(request.index)?,
            },
        };
        output.push(data);
    }
    Ok(())
}

fn read_parallel(
    package: &SourceBackedPackage,
    context: &ExecutionContext,
    requests: &[PreparedRequest],
    workers: usize,
    scheduler: &SchedulerAdmission,
    output: &mut Vec<PartData>,
) -> Result<()> {
    let max_bytes = context.limits().max_in_flight_bytes().get();
    fence(package)?;
    if context.scoped_workers().is_some() {
        return read_with_scoped_workers(package, context, requests, workers, max_bytes, output);
    }
    let first_end = wave_end(requests, 0, workers, max_bytes)?;
    if first_end == requests.len() {
        return read_one_wave(package, requests, 0, first_end, scheduler, output);
    }
    read_reused_workers(package, requests, workers, max_bytes, scheduler, output)
}

fn read_with_scoped_workers(
    package: &SourceBackedPackage,
    context: &ExecutionContext,
    requests: &[PreparedRequest],
    workers: usize,
    max_bytes: u64,
    output: &mut Vec<PartData>,
) -> Result<()> {
    let facility = context
        .scoped_workers()
        .ok_or_else(|| batch_refusal("caller worker facility disappeared"))?;
    let mut start = 0;
    while start < requests.len() {
        fence(package)?;
        let end = wave_end(requests, start, workers, max_bytes)?;
        let mut slots: Vec<Option<Result<PartData>>> = Vec::new();
        slots
            .try_reserve_exact(end - start)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch caller worker slots",
                source,
            })?;
        slots.resize_with(end - start, || None);
        {
            let mut tasks: Vec<Box<dyn FnMut() + Send + '_>> = requests[start..end]
                .iter()
                .zip(slots.iter_mut())
                .enumerate()
                .map(|(offset, (request, slot))| {
                    let ordinal = start + offset;
                    let task: Box<dyn FnMut() + Send + '_> = Box::new(move || {
                        let result = catch_unwind(AssertUnwindSafe(|| {
                            let declared = request.declared_bytes.ok_or_else(|| {
                                batch_refusal("parallel batch request lacks prepared metadata")
                            })?;
                            package.read_part_prepared(request.index, request.entry_id, declared)
                        }))
                        .unwrap_or_else(|_| {
                            Err(OpcError::SourceBackedBatchWorkerPanic { ordinal })
                        });
                        *slot = Some(result);
                    });
                    task
                })
                .collect();
            let mut task_refs: Vec<&mut (dyn FnMut() + Send + '_)> =
                tasks.iter_mut().map(|task| &mut **task).collect();
            facility.run_all(&mut task_refs);
        }
        let mut selected = None;
        for (offset, result) in slots.into_iter().enumerate() {
            match result.unwrap_or_else(|| {
                Err(OpcError::SourceBackedBatchWorkerPanic {
                    ordinal: start + offset,
                })
            }) {
                Ok(data) => output.push(data),
                Err(error) => select_error(&mut selected, start + offset, error),
            }
        }
        if let Some((_, error)) = selected {
            return Err(error);
        }
        start = end;
    }
    Ok(())
}

fn read_one_wave(
    package: &SourceBackedPackage,
    requests: &[PreparedRequest],
    start: usize,
    end: usize,
    scheduler: &SchedulerAdmission,
    output: &mut Vec<PartData>,
) -> Result<()> {
    let selected = thread::scope(|scope| {
        let width = end - start;
        let mut handles = Vec::new();
        if let Err(source) = handles.try_reserve_exact(width) {
            return Some((
                start,
                OpcError::Allocation {
                    resource: "source-backed batch worker handles",
                    source,
                },
            ));
        }
        if let Err(error) =
            capacity_fits::<(usize, thread::ScopedJoinHandle<'static, Result<PartData>>)>(
                handles.capacity(),
                scheduler.handle_bytes,
                "source-backed batch worker handles",
            )
        {
            return Some((start, error));
        }
        let mut selected = None;
        for (ordinal, request) in requests.iter().copied().enumerate().take(end).skip(start) {
            let result = thread::Builder::new()
                .stack_size(THREAD_STACK_BYTES)
                .spawn_scoped(scope, move || {
                    let declared = request.declared_bytes.ok_or_else(|| {
                        batch_refusal("parallel batch request lacks prepared metadata")
                    })?;
                    package.read_part_prepared(request.index, request.entry_id, declared)
                });
            match result {
                Ok(handle) => handles.push((ordinal, handle)),
                Err(error) => {
                    select_error(&mut selected, ordinal, OpcError::IoError(error));
                    break;
                },
            }
        }
        for (ordinal, handle) in handles.drain(..) {
            match handle.join() {
                Ok(Ok(data)) => output.push(data),
                Ok(Err(error)) => select_error(&mut selected, ordinal, error),
                Err(_) => select_error(
                    &mut selected,
                    ordinal,
                    OpcError::SourceBackedBatchWorkerPanic { ordinal },
                ),
            }
        }
        selected
    });
    selected.map_or(Ok(()), |(_, error)| Err(error))
}

fn read_reused_workers(
    package: &SourceBackedPackage,
    requests: &[PreparedRequest],
    workers: usize,
    max_bytes: u64,
    scheduler: &SchedulerAdmission,
    output: &mut Vec<PartData>,
) -> Result<()> {
    thread::scope(|scope| {
        let mut command_senders = Vec::new();
        let mut reply_receivers = Vec::new();
        let mut worker_handles = Vec::new();
        let mut active_ordinals = Vec::new();
        command_senders
            .try_reserve_exact(workers)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch worker channels",
                source,
            })?;
        reply_receivers
            .try_reserve_exact(workers)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch worker reply receivers",
                source,
            })?;
        worker_handles
            .try_reserve_exact(workers)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch worker handles",
                source,
            })?;
        active_ordinals
            .try_reserve_exact(workers)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed batch worker ordinals",
                source,
            })?;
        capacity_fits::<SyncSender<WorkerCommand>>(
            command_senders.capacity(),
            scheduler.command_vec_bytes,
            "source-backed batch worker command senders",
        )?;
        capacity_fits::<Receiver<Result<PartData>>>(
            reply_receivers.capacity(),
            scheduler.reply_vec_bytes,
            "source-backed batch worker reply receivers",
        )?;
        capacity_fits::<(usize, thread::ScopedJoinHandle<'static, ()>)>(
            worker_handles.capacity(),
            scheduler.handle_bytes,
            "source-backed batch worker handles",
        )?;
        capacity_fits::<Option<usize>>(
            active_ordinals.capacity(),
            scheduler.ordinal_vec_bytes,
            "source-backed batch worker ordinals",
        )?;
        for _ in 0..workers {
            active_ordinals.push(None);
        }

        let mut startup_error = None;
        for worker in 0..workers {
            let (command_sender, command_receiver) = mpsc::sync_channel::<WorkerCommand>(1);
            let (reply_sender, reply_receiver) = mpsc::sync_channel::<Result<PartData>>(1);
            match thread::Builder::new()
                .stack_size(THREAD_STACK_BYTES)
                .spawn_scoped(scope, move || {
                    worker_loop(package, command_receiver, reply_sender)
                }) {
                Ok(handle) => {
                    command_senders.push(command_sender);
                    reply_receivers.push(reply_receiver);
                    worker_handles.push((worker, handle));
                },
                Err(error) => {
                    startup_error = Some((worker, OpcError::IoError(error)));
                    break;
                },
            }
        }
        if let Some((ordinal, error)) = startup_error {
            shutdown_workers(&command_senders);
            let mut selected = Some((ordinal, error));
            join_workers(worker_handles, &active_ordinals, &mut selected, ordinal);
            return selected.map_or_else(
                || Err(batch_refusal("source-backed batch worker startup failed")),
                |(_, error)| Err(error),
            );
        }

        let run_result = run_reused_waves(
            package,
            requests,
            max_bytes,
            &command_senders,
            &reply_receivers,
            &mut active_ordinals,
            output,
        );
        shutdown_workers(&command_senders);
        let mut selected = run_result.err();
        let fallback_ordinal = selected
            .as_ref()
            .map_or(requests.len(), |(ordinal, _)| *ordinal);
        join_workers(
            worker_handles,
            &active_ordinals,
            &mut selected,
            fallback_ordinal,
        );
        selected.map_or(Ok(()), |(_, error)| Err(error))
    })
}

fn run_reused_waves(
    package: &SourceBackedPackage,
    requests: &[PreparedRequest],
    max_bytes: u64,
    command_senders: &[SyncSender<WorkerCommand>],
    reply_receivers: &[Receiver<Result<PartData>>],
    active_ordinals: &mut [Option<usize>],
    output: &mut Vec<PartData>,
) -> std::result::Result<(), (usize, OpcError)> {
    let workers = command_senders.len();
    let mut start = 0;
    while start < requests.len() {
        fence(package).map_err(|error| (start, error))?;
        let end = wave_end(requests, start, workers, max_bytes).map_err(|error| (start, error))?;
        let width = end - start;
        for ordinal in active_ordinals.iter_mut() {
            *ordinal = None;
        }

        let mut sent = 0;
        let mut selected = None;
        for lane in 0..width {
            let ordinal = start + lane;
            active_ordinals[lane] = Some(ordinal);
            if command_senders[lane]
                .send(WorkerCommand::Read {
                    ordinal,
                    request: requests[ordinal],
                })
                .is_err()
            {
                select_error(
                    &mut selected,
                    ordinal,
                    OpcError::SourceBackedBatchWorkerPanic { ordinal },
                );
                break;
            }
            sent += 1;
        }

        // Each lane has its own one-slot reply queue. Receive in input order
        // so no shared transport can strand a later idle worker's reply.
        for lane in 0..sent {
            match reply_receivers[lane].recv() {
                Ok(result) => {
                    active_ordinals[lane] = None;
                    match result {
                        Ok(data) => output.push(data),
                        Err(error) => select_error(&mut selected, start + lane, error),
                    }
                },
                Err(_) => {
                    let ordinal = active_ordinals[lane].unwrap_or(start + lane);
                    select_error(
                        &mut selected,
                        ordinal,
                        OpcError::SourceBackedBatchWorkerPanic { ordinal },
                    );
                },
            }
        }

        if let Some((ordinal, error)) = selected {
            return Err((ordinal, error));
        }
        if sent != width {
            return Err((
                start + sent,
                OpcError::SourceBackedBatchWorkerPanic {
                    ordinal: start + sent,
                },
            ));
        }
        start = end;
    }
    Ok(())
}

fn worker_loop(
    package: &SourceBackedPackage,
    commands: Receiver<WorkerCommand>,
    replies: SyncSender<Result<PartData>>,
) {
    while let Ok(command) = commands.recv() {
        match command {
            WorkerCommand::Shutdown => return,
            WorkerCommand::Read { ordinal, request } => {
                let outcome = catch_unwind(AssertUnwindSafe(|| {
                    let declared = request.declared_bytes.ok_or_else(|| {
                        batch_refusal("parallel batch request lacks prepared metadata")
                    })?;
                    package.read_part_prepared(request.index, request.entry_id, declared)
                }));
                let panicked = outcome.is_err();
                let result = match outcome {
                    Ok(result) => result,
                    Err(_) => Err(OpcError::SourceBackedBatchWorkerPanic { ordinal }),
                };
                if replies.send(result).is_err() {
                    return;
                }
                if panicked {
                    return;
                }
            },
        }
    }
}

fn shutdown_workers(senders: &[SyncSender<WorkerCommand>]) {
    for sender in senders {
        let _ = sender.send(WorkerCommand::Shutdown);
    }
}

fn join_workers<'scope>(
    handles: Vec<(usize, thread::ScopedJoinHandle<'scope, ()>)>,
    active_ordinals: &[Option<usize>],
    selected: &mut Option<(usize, OpcError)>,
    fallback_ordinal: usize,
) {
    for (worker, handle) in handles {
        if handle.join().is_err() {
            let ordinal = active_ordinals
                .get(worker)
                .and_then(|ordinal| *ordinal)
                .unwrap_or(fallback_ordinal);
            select_error(
                selected,
                ordinal,
                OpcError::SourceBackedBatchWorkerPanic { ordinal },
            );
        }
    }
}

fn select_error(selected: &mut Option<(usize, OpcError)>, ordinal: usize, error: OpcError) {
    if selected
        .as_ref()
        .is_none_or(|(selected_ordinal, _)| ordinal < *selected_ordinal)
    {
        *selected = Some((ordinal, error));
    }
}

fn wave_end(
    requests: &[PreparedRequest],
    start: usize,
    workers: usize,
    max_bytes: u64,
) -> Result<usize> {
    let mut end = start;
    let mut bytes = 0_u64;
    while end < requests.len() && end - start < workers {
        let declared = requests[end]
            .declared_bytes
            .ok_or_else(|| batch_refusal("parallel batch request lacks declared bytes"))?;
        let next_bytes = bytes
            .checked_add(declared)
            .ok_or_else(|| batch_refusal("parallel batch wave bytes overflow"))?;
        if end > start && next_bytes > max_bytes {
            break;
        }
        if declared > max_bytes {
            return Err(batch_refusal("parallel batch admitted an oversized Part"));
        }
        bytes = next_bytes;
        end += 1;
    }
    if end == start {
        return Err(batch_refusal("parallel batch wave made no progress"));
    }
    Ok(end)
}

fn check_request_count(package: &SourceBackedPackage, actual: usize) -> Result<()> {
    let actual = u64::try_from(actual)
        .map_err(|_| batch_refusal("source-backed batch request count exceeds u64"))?;
    let maximum = u64::try_from(package.limits.max_parts())
        .map_err(|_| batch_refusal("source-backed batch Part limit exceeds u64"))?;
    if actual > maximum {
        return Err(OpcError::ReadLimit {
            resource: ReadResource::Parts,
            actual,
            maximum,
        });
    }
    Ok(())
}

fn resolve_part_index(
    package: &SourceBackedPackage,
    partname: &PackURI,
    context: Option<&ExecutionContext>,
) -> Result<usize> {
    if let Some(index) = package.parts_by_name.get(partname).copied() {
        return Ok(index);
    }
    if package.casefold_order.is_some() {
        return package
            .part_index(partname)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()));
    }
    for (position, part) in package.parts.iter().enumerate() {
        if let Some(context) = context
            && position & 0xff == 0
        {
            context.check().map_err(super::map_execution_error)?;
        }
        if part
            .partname
            .as_str()
            .eq_ignore_ascii_case(partname.as_str())
        {
            return Ok(position);
        }
    }
    Err(OpcError::PartNotFound(partname.to_string()))
}

fn declared_size(package: &SourceBackedPackage, entry_id: EntryId) -> Result<u64> {
    fence(package)?;
    let metadata = package
        .archive
        .metadata_for(entry_id)
        .map(|metadata| metadata.uncompressed_size())
        .map_err(OpcError::from);
    fence(package)?;
    metadata
}

fn checked_collection_bytes(
    count: usize,
    element_size: usize,
    resource: &'static str,
) -> Result<usize> {
    count
        .checked_mul(element_size)
        .ok_or_else(|| batch_refusal(resource))
}

fn capacity_fits<T>(capacity: usize, charged_bytes: usize, resource: &'static str) -> Result<()> {
    let actual = capacity
        .checked_mul(size_of::<T>())
        .ok_or_else(|| batch_refusal("source-backed batch capacity overflows"))?;
    if actual > charged_bytes {
        return Err(batch_refusal(resource));
    }
    Ok(())
}

fn reserve_shape(
    context: Option<&ExecutionContext>,
    memory_bytes: usize,
    object_count: u64,
) -> Result<(Option<Reservation>, Option<Reservation>)> {
    let Some(context) = context else {
        return Ok((None, None));
    };
    let memory = context
        .reserve(
            Resource::Memory,
            u64::try_from(memory_bytes)
                .map_err(|_| batch_refusal("source-backed batch shape exceeds u64"))?,
        )
        .map_err(super::map_execution_error)?;
    match context.reserve(Resource::Objects, object_count) {
        Ok(objects) => Ok((Some(memory), Some(objects))),
        Err(error) => {
            drop(memory);
            Err(super::map_execution_error(error))
        },
    }
}

fn fence(package: &SourceBackedPackage) -> Result<()> {
    package.source.ensure_current()?;
    package
        .cache
        .check_context()
        .map_err(super::map_execution_error)
}

fn finish(package: &SourceBackedPackage, result: Result<PartBatch>) -> Result<PartBatch> {
    // Source freshness remains the outermost decision, then cancellation;
    // only when both fences pass is the prepared/task result exposed.
    package.source.ensure_current()?;
    package
        .cache
        .check_context()
        .map_err(super::map_execution_error)?;
    result
}

fn batch_refusal(reason: &'static str) -> OpcError {
    OpcError::SourceBackedBatchInvariant { reason }
}
