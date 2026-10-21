//! Authenticated exact restoration using an explicitly supplied original artifact.

use super::splice::{
    check_inverse_state, fingerprint_inverse_snapshot, finish_inverse_publication,
};
use super::{
    ContextCheckedSink, Counted, OpcError, OutputBudgetedSink, Result,
    SOURCE_PUBLICATION_CHUNK_BYTES, SourceArtifact, SourceArtifactFingerprint, SourceBackedPackage,
    SourceCheckedSink, SourceSnapshot, map_execution_error, map_io_error, overlay_unavailable,
    read_source_at_with_context,
};
use crate::SpliceResource;
use litchi_core::{ExecutionContext, ExecutionError, Resource};
use sha2::{Digest as _, Sha256};
use std::io::{self, Write};
use std::sync::{Arc, Mutex};

/// Complete artifact identities for a durable exact inverse.
///
/// These identities survive reopening through a different positional source.
/// Runtime source versions are checked separately throughout the operation.
/// A format-owned patch must authenticate this proof before invoking the
/// low-level restoration capability.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceArtifactRestoreProof {
    pub current_len: u64,
    pub current_sha256: SourceArtifactFingerprint,
    pub original_len: u64,
    pub original_sha256: SourceArtifactFingerprint,
}

impl SourceBackedPackage {
    /// Restore an explicit original artifact after authenticating both complete
    /// artifacts, without materializing either artifact.
    ///
    /// The current package's execution context controls the operation. The
    /// retained provider is still checked for freshness and cancellation, but
    /// its resource policy is never inherited by the recipient. Both
    /// providers must remain fresh and uncancelled. A single 64 KiB workspace
    /// is reserved when the current context has one; output is charged once by
    /// that context. The original bytes are authenticated again as they are
    /// emitted. Any failure after accepted output is reported with its exact
    /// accepted byte count.
    pub fn restore_source_artifact_to_stream<W: Write>(
        &self,
        original: &SourceArtifact,
        proof: SourceArtifactRestoreProof,
        max_output_bytes: u64,
        writer: W,
    ) -> Result<()> {
        self.disable_read_ahead_for_publication()?;
        let retained = &original.snapshot;
        let context = self.cache.context().cloned();
        check_inverse_state(self, retained, context.as_ref())?;
        if max_output_bytes == 0 || max_output_bytes == u64::MAX {
            return Err(OpcError::InvalidSourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                value: max_output_bytes,
            });
        }
        if retained.length > max_output_bytes {
            return Err(OpcError::SourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                actual: retained.length,
                maximum: max_output_bytes,
            });
        }
        if self.source.length != proof.current_len {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "length",
            });
        }
        if retained.length != proof.original_len {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "length",
            });
        }
        let output_failures = if self.cache.context().is_some() {
            self.source.output_reservation_failures.clone()
        } else {
            retained.output_reservation_failures.clone()
        };
        if context.is_some() && output_failures.is_none() {
            return Err(overlay_unavailable(
                "managed restore output counter is unavailable",
            ));
        }
        let _monitored_current = self.source.monitor_publication();
        let _monitored_retained = retained.monitor_publication();
        let mut use_heap_buffer = context.is_none();
        let _workspace = if let Some(context) = context.as_ref() {
            let requested = SOURCE_PUBLICATION_CHUNK_BYTES;
            match context.reserve(Resource::Memory, requested as u64) {
                Ok(reservation) => Some(reservation),
                Err(ExecutionError::ResourceLimit(limit)) if limit.resource == Resource::Memory => {
                    let used = limit.observed.saturating_sub(requested as u64);
                    let available = limit.limit.saturating_sub(used);
                    if available == 0 {
                        // The package may already retain its complete
                        // catalog under an exact memory allowance. A
                        // bounded stack window keeps authentication and
                        // streaming available without an uncharged heap
                        // allocation or a false policy refusal.
                        None
                    } else {
                        return Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)));
                    }
                },
                Err(error) => return Err(map_execution_error(error)),
            }
        } else {
            None
        };
        if _workspace.is_some() {
            use_heap_buffer = true;
        }
        const RESTORE_STACK_BYTES: usize = 4 * 1024;
        let mut stack_buffer = [0_u8; RESTORE_STACK_BYTES];
        let mut heap_buffer = Vec::new();
        let buffer: &mut [u8] = if use_heap_buffer {
            heap_buffer
                .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
                .map_err(|source| OpcError::Allocation {
                    resource: "exact artifact restore buffer",
                    source,
                })?;
            if heap_buffer.capacity() > SOURCE_PUBLICATION_CHUNK_BYTES {
                return Err(overlay_unavailable(
                    "exact artifact restore allocation exceeds its reserved window",
                ));
            }
            heap_buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
            heap_buffer.as_mut_slice()
        } else {
            &mut stack_buffer
        };
        let current_hash =
            fingerprint_inverse_snapshot(self, &self.source, retained, context.as_ref(), buffer)?;
        if current_hash != proof.current_sha256 {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "fingerprint",
            });
        }
        let original_hash =
            fingerprint_inverse_snapshot(self, retained, retained, context.as_ref(), buffer)?;
        if original_hash != proof.original_sha256 {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "fingerprint",
            });
        }
        let failure = Arc::new(Mutex::new(None));
        let mut written = 0;
        let counted = Counted::new(writer, &mut written);
        let checked_current = SourceCheckedSink {
            inner: counted,
            snapshot: self.source.clone(),
        };
        let checked_original = SourceCheckedSink {
            inner: checked_current,
            snapshot: retained.clone(),
        };
        let original_cooperative = ContextCheckedSink {
            inner: checked_original,
            context: retained.context.clone(),
            failure: Arc::clone(&failure),
        };
        let cooperative = ContextCheckedSink {
            inner: original_cooperative,
            context: context.clone(),
            failure: Arc::clone(&failure),
        };
        let result = if let (Some(operation_context), Some(output_reservation_failures)) =
            (context.clone(), output_failures)
        {
            let mut output = OutputBudgetedSink {
                inner: cooperative,
                context: operation_context,
                failure: Arc::clone(&failure),
                output_reservation_failures,
            };
            copy_authenticated(
                self,
                retained,
                context.as_ref(),
                proof.original_sha256,
                buffer,
                &mut output,
            )
            .and_then(|()| output.flush().map_err(map_io_error))
        } else {
            let mut output = cooperative;
            copy_authenticated(
                self,
                retained,
                context.as_ref(),
                proof.original_sha256,
                buffer,
                &mut output,
            )
            .and_then(|()| output.flush().map_err(map_io_error))
        };
        let result = failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = match (
            result,
            check_inverse_state(self, retained, context.as_ref()),
        ) {
            (_, Err(error @ OpcError::SourceChanged { .. })) => Err(error),
            (Err(error), _) | (Ok(()), Err(error)) => Err(error),
            (Ok(()), Ok(())) => Ok(()),
        };
        finish_inverse_publication(result, &self.source, retained, written)
    }
}

fn copy_authenticated(
    current: &SourceBackedPackage,
    retained: &SourceSnapshot,
    context: Option<&ExecutionContext>,
    expected: SourceArtifactFingerprint,
    buffer: &mut [u8],
    output: &mut dyn Write,
) -> Result<()> {
    let mut offset = 0_u64;
    let mut hash = Sha256::new();
    while offset < retained.length {
        check_inverse_state(current, retained, context)?;
        let count = usize::try_from((retained.length - offset).min(buffer.len() as u64))
            .map_err(|_| overlay_unavailable("artifact restore range exceeds usize"))?;
        let read = read_source_at_with_context(
            retained,
            context,
            offset,
            &mut buffer[..count],
            "authenticated restore",
        )
        .map_err(super::map_source_backed_error)?;
        if read == 0 {
            return Err(OpcError::IoError(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "original artifact ended during restore",
            )));
        }
        check_inverse_state(current, retained, context)?;
        if let Some(context) = context {
            context
                .consume(Resource::Work, read as u64)
                .map_err(map_execution_error)?;
        }
        hash.update(&buffer[..read]);
        output.write_all(&buffer[..read]).map_err(map_io_error)?;
        check_inverse_state(current, retained, context)?;
        offset = offset
            .checked_add(read as u64)
            .ok_or_else(|| overlay_unavailable("artifact restore offset overflow"))?;
    }
    if SourceArtifactFingerprint::from_sha256(hash.finalize().into()) != expected {
        return Err(OpcError::SourceArtifactMismatch {
            artifact: "original",
            field: "replayed fingerprint",
        });
    }
    Ok(())
}
