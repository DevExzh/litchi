//! Bounded, opt-in source-backed multi-Part reads.
//!
//! This module owns only orchestration.  Source freshness, ZIP validation,
//! cache admission, payload reservations, and cumulative physical accounting
//! remain in [`super::SourceBackedPackage`]'s one-Part path.

use super::{PartData, SourceBackedPackage};
use crate::error::{OpcError, Result};
use crate::limits::ReadResource;
use crate::packuri::PackURI;
use litchi_core::{ExecutionContext, Reservation, Resource};
use soapberry_zip::office::EntryId;
use std::mem::size_of;
use std::thread;

// Keep the platform's ordinary scoped-worker depth envelope conservative. The
// reservation is explicit even though the OS owns the actual stack mapping.
const THREAD_STACK_BYTES: usize = 2 * 1024 * 1024;
const SCHEDULER_CONTROL_BYTES: usize = 1024;

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
    _memory: Reservation,
    _objects: Reservation,
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
            && prepared.all_fit_inflight
    });

    let result = if parallel {
        let context = context
            .as_ref()
            .ok_or_else(|| batch_refusal("parallel scheduling requires an execution context"))?;
        let scheduler = match SchedulerAdmission::reserve(context, workers) {
            Ok(scheduler) => scheduler,
            Err(error) => {
                drop(parts);
                drop(output);
                drop(prepared);
                return finish(package, Err(error));
            },
        };
        let result = read_parallel(
            package,
            context,
            &prepared.requests,
            workers,
            &scheduler,
            &mut parts,
        );
        drop(scheduler);
        result
    } else {
        read_serial(package, &prepared.requests, &mut parts)
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
    fn reserve(context: &ExecutionContext, workers: usize) -> Result<Self> {
        let handle_bytes = checked_collection_bytes(
            workers,
            size_of::<(usize, thread::ScopedJoinHandle<'static, Result<PartData>>)>(),
            "source-backed batch worker handles",
        )?;
        let stack_bytes = workers
            .checked_mul(THREAD_STACK_BYTES)
            .ok_or_else(|| batch_refusal("source-backed batch worker stacks overflow"))?;
        let control_bytes = SCHEDULER_CONTROL_BYTES
            .checked_add(handle_bytes)
            .and_then(|bytes| bytes.checked_add(stack_bytes))
            .ok_or_else(|| batch_refusal("source-backed batch scheduler overflows"))?;
        let memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(control_bytes)
                    .map_err(|_| batch_refusal("batch scheduler exceeds u64"))?,
            )
            .map_err(super::map_execution_error)?;
        let object_count = u64::try_from(
            workers
                .checked_add(1)
                .ok_or_else(|| batch_refusal("batch scheduler objects overflow"))?,
        )
        .map_err(|_| batch_refusal("batch scheduler objects exceed u64"))?;
        let objects = match context.reserve(Resource::Objects, object_count) {
            Ok(objects) => objects,
            Err(error) => {
                drop(memory);
                return Err(super::map_execution_error(error));
            },
        };
        Ok(Self {
            handle_bytes,
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
    for request in requests {
        let data = match request.declared_bytes {
            Some(declared) => {
                package.read_part_prepared(request.index, request.entry_id, declared)?
            },
            None => package.read_part(request.index)?,
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
    let mut start = 0;
    while start < requests.len() {
        // Recheck before admitting each later wave.  Workers already admitted
        // in the current wave are joined below; this check prevents a
        // cancellation or source transition from starting fresh work.
        fence(package)?;
        let end = wave_end(requests, start, workers, max_bytes)?;
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
            let mut setup_error = None;
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
                        setup_error = Some((ordinal, OpcError::IoError(error)));
                        break;
                    },
                }
            }

            let mut selected = setup_error;
            for (ordinal, handle) in handles.drain(..) {
                match handle.join() {
                    Ok(Ok(data)) => output.push(data),
                    Ok(Err(error)) => {
                        if selected
                            .as_ref()
                            .is_none_or(|(selected_ordinal, _)| ordinal < *selected_ordinal)
                        {
                            selected = Some((ordinal, error));
                        }
                    },
                    Err(_) => {
                        let error = OpcError::SourceBackedBatchWorkerPanic { ordinal };
                        if selected
                            .as_ref()
                            .is_none_or(|(selected_ordinal, _)| ordinal < *selected_ordinal)
                        {
                            selected = Some((ordinal, error));
                        }
                    },
                }
            }
            selected
        });
        if let Some((_, error)) = selected {
            return Err(error);
        }
        start = end;
    }
    Ok(())
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
