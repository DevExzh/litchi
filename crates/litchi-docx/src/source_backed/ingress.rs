//! Bounded sequential ingestion whose source owns every retained input charge.

use std::io::{self, Read};
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::{
    ExecutionContext, ExecutionError, OwnedSource, ReadAt, Reservation, Resource, SourceVersion,
};
use litchi_opc::{OpcError, SourceLimitKind};

use crate::{Error, OpenOptions, Result};

const CHUNK: usize = 8 * 1024;

// Field drop order frees the allocation before releasing its reservation.
struct Input {
    source: OwnedSource,
    _buffer: Option<Reservation>,
    _owner: Option<Reservation>,
}

impl ReadAt for Input {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.source.read_at(offset, output)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

fn execution(error: ExecutionError) -> Error {
    match error {
        ExecutionError::Cancelled => OpcError::Cancelled.into(),
        error => OpcError::Execution(error).into(),
    }
}

fn reserve(
    context: Option<&ExecutionContext>,
    resource: Resource,
    amount: usize,
) -> Result<Option<Reservation>> {
    context
        .map(|context| context.reserve(resource, amount as u64).map_err(execution))
        .transpose()
}

fn work(context: Option<&ExecutionContext>, amount: usize) -> Result<()> {
    if let Some(context) = context {
        context
            .consume(Resource::Work, amount as u64)
            .map_err(execution)?;
    }
    Ok(())
}

// A large caller buffer must not defeat a smaller remaining input budget.
fn input_window(
    context: Option<&ExecutionContext>,
    requested: usize,
) -> Result<(usize, Option<Reservation>)> {
    let Some(context) = context else {
        return Ok((requested, None));
    };
    let mut attempt = requested as u64;
    loop {
        match context.reserve(Resource::InputBytes, attempt) {
            Ok(reservation) => return Ok((attempt as usize, Some(reservation))),
            Err(error) => {
                let ExecutionError::ResourceLimit(limit) = &error else {
                    return Err(execution(error));
                };
                let previous = limit.observed.saturating_sub(attempt);
                let available = limit.limit.saturating_sub(previous);
                let next = available.min(attempt.saturating_sub(1));
                if next == 0 {
                    return Err(execution(error));
                }
                attempt = next;
            },
        }
    }
}

fn read<R: Read>(
    reader: &mut R,
    output: &mut [u8],
    context: Option<&ExecutionContext>,
) -> Result<usize> {
    loop {
        let (length, reservation) = input_window(context, output.len())?;
        work(context, 1)?;
        let count = match reader.read(&mut output[..length]) {
            Ok(count) => count,
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            Err(error) => return Err(error.into()),
        };
        if count > length {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "reader returned more bytes than supplied capacity",
            )
            .into());
        }
        if let Some(reservation) = reservation {
            if !reservation.commit(count as u64) {
                return Err(io::Error::other("DOCX input reservation underflow").into());
            }
        }
        if let Some(context) = context {
            context.check().map_err(execution)?;
        }
        return Ok(count);
    }
}

pub(super) fn from_reader<R: Read>(
    mut reader: R,
    options: &OpenOptions,
) -> Result<Arc<dyn ReadAt>> {
    let context = options.execution_context.as_ref();
    let maximum = usize::try_from(options.source_limits.max_input_bytes()).map_err(|_| {
        OpcError::InvalidLimit {
            kind: SourceLimitKind::InputBytes,
            value: options.source_limits.max_input_bytes(),
            maximum: usize::MAX.saturating_sub(1) as u64,
        }
    })?;
    // Both Arc control blocks, the source object and its Vec header stay charged.
    let owner = reserve(
        context,
        Resource::Memory,
        size_of::<Input>() + size_of::<Vec<u8>>() + 4 * size_of::<usize>(),
    )?;
    let scratch = reserve(context, Resource::Memory, CHUNK)?;
    let mut buffer_charge = None;
    let mut bytes = Vec::new();
    let mut buffer = [0; CHUNK];
    loop {
        let remaining = maximum - bytes.len();
        let requested = remaining.clamp(1, CHUNK);
        let count = read(&mut reader, &mut buffer[..requested], context)?;
        if count == 0 {
            break;
        }
        if remaining == 0 {
            return Err(OpcError::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: maximum as u64 + 1,
                maximum: maximum as u64,
            }
            .into());
        }
        let needed = bytes.len() + count;
        if needed > bytes.capacity() {
            let capacity = bytes.capacity().saturating_mul(2).max(needed).min(maximum);
            // Reserve the complete replacement while the old allocation is live.
            // This accounts for the copy's peak, not merely the capacity delta.
            let replacement_charge = reserve(context, Resource::Memory, capacity)?;
            work(context, bytes.len())?;
            let mut replacement = Vec::new();
            replacement
                .try_reserve_exact(capacity)
                .map_err(|source| Error::Allocation {
                    resource: "sequential DOCX input",
                    source,
                })?;
            if replacement.capacity() != capacity {
                return Err(Error::Other(
                    "DOCX input allocator exceeded admitted capacity".into(),
                ));
            }
            replacement.extend_from_slice(&bytes);
            drop(std::mem::replace(&mut bytes, replacement));
            buffer_charge = replacement_charge;
        }
        work(context, count)?;
        bytes.extend_from_slice(&buffer[..count]);
    }
    drop(scratch);
    Ok(Arc::new(Input {
        source: OwnedSource::new(bytes),
        _buffer: buffer_charge,
        _owner: owner,
    }))
}

#[cfg(test)]
mod tests;
