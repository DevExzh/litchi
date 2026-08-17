//! Checked, source-backed publication of same-length CFB stream overlays.
//!
//! This deliberately does not accept caller-selected physical offsets. Plans
//! are derived from stream chains already validated by [`SharedOleFile`], then
//! the composed artifact is reopened through the normal CFB parser before a
//! sink can observe a byte.

use crate::consts::{ENDOFCHAIN, MAXREGSECT, STGTY_STREAM};
use crate::file::{OleError, OleFile, ParsedOleIndex};
use crate::shared::SharedOleFile;
use crate::writer::{atomic_replace, create_sibling_temp_file, parent_directory, sync_parent};
use litchi_core::{ReadAt, SourceVersion};
use sha2::{Digest as _, Sha256};
use std::collections::TryReserveError;
use std::fs;
use std::io::{self, BufWriter, Read, Seek, SeekFrom, Write};
use std::ops::Range;
use std::path::Path;
use std::sync::Arc;

const PUBLICATION_CHUNK_BYTES: usize = 64 * 1024;
const FINGERPRINT_CHUNK_BYTES: usize = 1024 * 1024;
const FINGERPRINT_CHUNK_BYTES_U64: u64 = 1024 * 1024;

/// SHA-256 identity of a complete source or composed CFB artifact.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ArtifactFingerprint([u8; 32]);

impl ArtifactFingerprint {
    /// Returns the digest bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 32] {
        &self.0
    }
}

/// Finite resource bounds for one same-length overlay plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct OverlayLimits {
    max_streams: usize,
    max_spans: usize,
    max_overlay_bytes: u64,
}

impl OverlayLimits {
    /// Creates a non-zero bounded overlay policy.
    ///
    /// # Errors
    ///
    /// Returns an error if any limit is zero.
    pub fn new(
        max_streams: usize,
        max_spans: usize,
        max_overlay_bytes: u64,
    ) -> Result<Self, OverlayError> {
        if max_streams == 0 || max_spans == 0 || max_overlay_bytes == 0 {
            return Err(OverlayError::Unavailable {
                reason: "same-length overlay limits must be non-zero".to_string(),
            });
        }
        Ok(Self {
            max_streams,
            max_spans,
            max_overlay_bytes,
        })
    }

    /// Maximum number of logical stream replacements.
    #[must_use]
    pub const fn max_streams(self) -> usize {
        self.max_streams
    }

    /// Maximum number of physical changed spans retained by the plan.
    #[must_use]
    pub const fn max_spans(self) -> usize {
        self.max_spans
    }

    /// Maximum aggregate logical replacement bytes.
    #[must_use]
    pub const fn max_overlay_bytes(self) -> u64 {
        self.max_overlay_bytes
    }
}

impl Default for OverlayLimits {
    fn default() -> Self {
        Self {
            max_streams: 64,
            max_spans: 65_536,
            max_overlay_bytes: 256 * 1024 * 1024,
        }
    }
}

/// One existing logical stream and its equal-length replacement.
#[derive(Clone, Debug)]
pub struct SameLengthStreamOverlay {
    path: Vec<String>,
    replacement: Arc<[u8]>,
}

impl SameLengthStreamOverlay {
    /// Creates a logical overlay without copying `replacement`.
    #[must_use]
    pub fn new(path: Vec<String>, replacement: Arc<[u8]>) -> Self {
        Self { path, replacement }
    }

    /// Selected CFB path.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Equal-length replacement bytes.
    #[must_use]
    pub fn replacement(&self) -> &Arc<[u8]> {
        &self.replacement
    }
}

/// What a non-atomic sequential sink may already contain after failure.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OutputProgress {
    /// The sink accepted no bytes.
    Untouched,
    /// An exact artifact prefix was accepted.
    Prefix {
        /// Bytes definitely accepted.
        accepted: u64,
        /// Complete artifact length.
        expected: u64,
    },
    /// Every artifact byte was accepted but the final flush did not complete.
    CompleteUnflushed {
        /// Complete artifact length.
        bytes: u64,
    },
    /// A hostile sink over-reported one write, so exact progress is unknowable.
    Indeterminate {
        /// Bytes definitely accepted before the invalid report.
        accepted_before: u64,
    },
}

/// Successful direct or atomic-path publication evidence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PublishReport {
    bytes: u64,
    changed_spans: usize,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
}

impl PublishReport {
    /// Complete published artifact length.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.bytes
    }

    /// Number of sorted physical spans that differed from the source.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Exact source identity checked before and during publication.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact composed target identity checked before publication.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }
}

/// Planning or publication failure for a same-length overlay.
#[derive(Debug)]
#[non_exhaustive]
pub enum OverlayError {
    /// The validated CFB reader rejected the source or composed artifact.
    Ole(OleError),
    /// Positional source or sequential sink I/O failed.
    Io(io::Error),
    /// The requested operation is outside the narrow same-length contract.
    Unavailable { reason: String },
    /// The positional adapter reported a different source revision.
    SourceChanged {
        expected: SourceVersion,
        observed: SourceVersion,
    },
    /// Full-artifact bytes differed despite an unchanged version token.
    SourceFingerprintChanged {
        expected: ArtifactFingerprint,
        observed: ArtifactFingerprint,
    },
    /// An expected stream-relative byte range did not match the source.
    PreconditionFailed {
        /// Selected logical stream.
        path: Vec<String>,
        /// Stream-relative start of the failed range.
        offset: u64,
    },
    /// The composed artifact identity differed from the validated plan.
    TargetFingerprintChanged {
        expected: ArtifactFingerprint,
        observed: ArtifactFingerprint,
    },
    /// A bounded planning or publication allocation failed.
    Allocation {
        resource: &'static str,
        source: TryReserveError,
    },
    /// A non-atomic sink may contain a typed prefix or indeterminate progress.
    IncompleteOutput {
        progress: OutputProgress,
        source: Box<OverlayError>,
    },
    /// The destination name was replaced but parent durability is unknown.
    Committed { source: io::Error },
}

impl std::fmt::Display for OverlayError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ole(error) => write!(formatter, "CFB overlay validation failed: {error}"),
            Self::Io(error) => write!(formatter, "CFB overlay I/O failed: {error}"),
            Self::Unavailable { reason } => {
                write!(
                    formatter,
                    "same-length CFB overlay is unavailable: {reason}"
                )
            },
            Self::SourceChanged { expected, observed } => write!(
                formatter,
                "CFB overlay source changed from {expected:?} to {observed:?}"
            ),
            Self::SourceFingerprintChanged { .. } => {
                formatter.write_str("CFB overlay source fingerprint changed")
            },
            Self::PreconditionFailed { path, offset } => write!(
                formatter,
                "CFB stream splice precondition failed for {path:?} at byte {offset}"
            ),
            Self::TargetFingerprintChanged { .. } => {
                formatter.write_str("CFB overlay target fingerprint changed")
            },
            Self::Allocation { resource, source } => {
                write!(formatter, "could not reserve {resource}: {source}")
            },
            Self::IncompleteOutput { progress, source } => {
                write!(
                    formatter,
                    "incomplete CFB overlay output ({progress:?}): {source}"
                )
            },
            Self::Committed { source } => write!(
                formatter,
                "CFB overlay destination was replaced but directory durability could not be confirmed: {source}"
            ),
        }
    }
}

impl std::error::Error for OverlayError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Ole(error) => Some(error),
            Self::Io(error) | Self::Committed { source: error } => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::IncompleteOutput { source, .. } => Some(source),
            Self::Unavailable { .. }
            | Self::SourceChanged { .. }
            | Self::SourceFingerprintChanged { .. }
            | Self::PreconditionFailed { .. }
            | Self::TargetFingerprintChanged { .. } => None,
        }
    }
}

impl From<OleError> for OverlayError {
    fn from(error: OleError) -> Self {
        match error {
            OleError::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            other => Self::Ole(other),
        }
    }
}

impl From<io::Error> for OverlayError {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

#[derive(Clone)]
pub(crate) struct SourceSnapshot {
    pub(crate) source: Arc<dyn ReadAt>,
    pub(crate) version: SourceVersion,
    pub(crate) length: u64,
    pub(crate) source_is_owned_immutable: bool,
}

impl SourceSnapshot {
    pub(crate) fn ensure_current(&self) -> Result<(), OverlayError> {
        let observed = self.source.version()?;
        if observed == self.version {
            Ok(())
        } else {
            Err(OverlayError::SourceChanged {
                expected: self.version,
                observed,
            })
        }
    }

    pub(crate) fn ensure_length(&self) -> Result<(), OverlayError> {
        self.ensure_current()?;
        let length = self.source.len()?;
        self.ensure_current()?;
        if length == self.length {
            Ok(())
        } else {
            Err(OverlayError::Unavailable {
                reason: format!(
                    "source length changed from {} to {length} without a new version token",
                    self.length
                ),
            })
        }
    }

    pub(crate) fn read_exact(&self, offset: u64, output: &mut [u8]) -> Result<(), OverlayError> {
        self.ensure_current()?;
        self.source.read_exact_at(offset, output)?;
        self.ensure_current()
    }
}

#[derive(Clone)]
pub(crate) struct PhysicalSpan {
    pub(crate) offset: u64,
    pub(crate) replacement: Arc<[u8]>,
    pub(crate) replacement_range: Range<usize>,
}

impl PhysicalSpan {
    pub(crate) fn len(&self) -> usize {
        self.replacement_range.len()
    }

    pub(crate) fn end(&self) -> Result<u64, OverlayError> {
        self.offset
            .checked_add(self.len() as u64)
            .ok_or_else(|| unavailable("physical span end overflows u64"))
    }
}

struct Selection {
    path: Vec<String>,
    replacement: Arc<[u8]>,
    sid: u32,
    start_sector: u32,
    size: u64,
    is_minifat: bool,
}

/// A fully validated, immutable streaming publication plan.
///
/// The plan owns only the positional source handle, changed physical spans,
/// and replacement allocations. It never constructs or retains a complete
/// replacement artifact.
pub struct ValidatedOverlayPlan {
    source: SourceSnapshot,
    spans: Vec<PhysicalSpan>,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
}

/// Immutable positional view of a fully validated composed overlay target.
///
/// The view keeps the source snapshot and replacement spans alive, performs no
/// whole-artifact copy, and supports concurrent positional reads. Creating it
/// through [`ValidatedOverlayPlan::composed_source`] rechecks both complete
/// artifact fingerprints before returning it.
#[derive(Clone)]
pub struct ComposedOverlaySource {
    source: SourceSnapshot,
    spans: Arc<[PhysicalSpan]>,
    version: SourceVersion,
}

impl std::fmt::Debug for ComposedOverlaySource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ComposedOverlaySource")
            .field("length", &self.source.length)
            .field("changed_spans", &self.spans.len())
            .finish_non_exhaustive()
    }
}

impl ReadAt for ComposedOverlaySource {
    fn len(&self) -> io::Result<u64> {
        self.source.ensure_length().map_err(overlay_io_error)?;
        Ok(self.source.length)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || offset >= self.source.length {
            self.source.ensure_length().map_err(overlay_io_error)?;
            return Ok(0);
        }
        self.source.ensure_current().map_err(overlay_io_error)?;
        let count = usize::try_from((self.source.length - offset).min(output.len() as u64))
            .map_err(|_| io::Error::other("composed overlay range does not fit usize"))?;
        let read = self.source.source.read_at(offset, &mut output[..count])?;
        if read > count {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        apply_spans(&mut output[..read], offset, &self.spans).map_err(overlay_io_error)?;
        self.source.ensure_current().map_err(overlay_io_error)?;
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.ensure_length().map_err(overlay_io_error)?;
        Ok(self.version)
    }
}

impl std::fmt::Debug for ValidatedOverlayPlan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ValidatedOverlayPlan")
            .field("source_length", &self.source.length)
            .field("changed_spans", &self.spans.len())
            .field("source_fingerprint", &self.source_fingerprint)
            .field("target_fingerprint", &self.target_fingerprint)
            .finish_non_exhaustive()
    }
}

impl SharedOleFile {
    /// Derives and fully validates physical overlays for existing equal-length
    /// streams.
    ///
    /// Logical targets are resolved through the validated CFB directory and
    /// their FAT or MiniFAT chains. Duplicate targets, length changes,
    /// generated overlaps, and resource-limit violations are rejected. The
    /// composed artifact is reopened through [`OleFile`] and every selected
    /// stream is checked before a reusable plan is returned.
    ///
    /// This CFB-level operation is policy-neutral. Mutation-capable format
    /// owners must retain their existing signature/encryption/DRM checks.
    pub fn plan_same_length_stream_overlays(
        &self,
        overlays: Vec<SameLengthStreamOverlay>,
        limits: OverlayLimits,
    ) -> Result<ValidatedOverlayPlan, OverlayError> {
        self.check_source_version()?;
        if overlays.len() > limits.max_streams {
            return Err(unavailable(format!(
                "replacement count {} exceeds limit {}",
                overlays.len(),
                limits.max_streams
            )));
        }

        let source = SourceSnapshot {
            source: Arc::clone(&self.source),
            version: self.expected_version,
            length: self.index.file_size,
            source_is_owned_immutable: self.source_is_owned_immutable,
        };
        source.ensure_length()?;

        let mut selections = Vec::new();
        selections
            .try_reserve_exact(overlays.len())
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB logical overlay selections",
                source,
            })?;
        let mut aggregate_bytes = 0_u64;
        for overlay in overlays {
            if overlay.path.is_empty() || overlay.path.iter().any(String::is_empty) {
                return Err(unavailable("overlay path must contain non-empty names"));
            }
            let refs = path_refs(&overlay.path)?;
            let entry = self.find_entry(&refs)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(unavailable(format!(
                    "overlay path {:?} does not identify a stream",
                    overlay.path
                )));
            }
            if selections
                .iter()
                .any(|selection: &Selection| selection.sid == entry.sid)
            {
                return Err(unavailable(format!(
                    "overlay set contains duplicate logical target {:?}",
                    overlay.path
                )));
            }
            let replacement_len = u64::try_from(overlay.replacement.len())
                .map_err(|_| unavailable("replacement length does not fit u64"))?;
            if replacement_len != entry.size {
                return Err(unavailable(format!(
                    "overlay for {:?} changes stream length from {} to {replacement_len}",
                    overlay.path, entry.size
                )));
            }
            aggregate_bytes = aggregate_bytes
                .checked_add(replacement_len)
                .ok_or_else(|| unavailable("aggregate replacement bytes overflow u64"))?;
            if aggregate_bytes > limits.max_overlay_bytes {
                return Err(unavailable(format!(
                    "aggregate replacement bytes {aggregate_bytes} exceed limit {}",
                    limits.max_overlay_bytes
                )));
            }
            selections.push(Selection {
                path: overlay.path,
                replacement: overlay.replacement,
                sid: entry.sid,
                start_sector: entry.start_sector,
                size: entry.size,
                is_minifat: entry.is_minifat,
            });
        }

        let mut spans = Vec::new();
        let mut comparison = publication_buffer()?;
        for selection in &selections {
            derive_selection_spans(
                &source,
                &self.index,
                selection,
                limits.max_spans,
                &mut comparison,
                &mut spans,
            )?;
        }
        spans.sort_unstable_by_key(|span| span.offset);
        validate_and_coalesce_spans(&source, limits, &mut spans)?;
        drop(comparison);

        let (source_fingerprint, target_fingerprint) = fingerprints(&source, &spans)?;
        validate_composed_artifact(&source, &spans, &selections)?;
        // An explicitly owned source and every retained replacement/span are
        // immutable. Candidate validation can only read that sealed composed
        // view, so a second complete fingerprint cannot observe a change.
        // Generic positional adapters retain the final bracket against a
        // dishonest stable version token and interior byte mutation.
        if !source.source_is_owned_immutable {
            let (observed_source, observed_target) = fingerprints(&source, &spans)?;
            if observed_source != source_fingerprint {
                return Err(OverlayError::SourceFingerprintChanged {
                    expected: source_fingerprint,
                    observed: observed_source,
                });
            }
            if observed_target != target_fingerprint {
                return Err(OverlayError::TargetFingerprintChanged {
                    expected: target_fingerprint,
                    observed: observed_target,
                });
            }
        }

        Ok(ValidatedOverlayPlan {
            source,
            spans,
            source_fingerprint,
            target_fingerprint,
        })
    }
}

impl ValidatedOverlayPlan {
    /// Whether every supplied replacement was an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.spans.is_empty()
    }

    /// Number of retained sorted, non-overlapping physical spans.
    #[must_use]
    pub fn changed_spans(&self) -> usize {
        self.spans.len()
    }

    /// Exact source artifact fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact composed target artifact fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.target_fingerprint
    }

    /// Returns a checked, read-only positional view of the composed target.
    ///
    /// Planning has already reopened the complete candidate through the normal
    /// CFB validator. This method additionally rechecks the complete source and
    /// target fingerprints, so a stale or dishonest source adapter cannot hand
    /// a caller an unvalidated view.
    pub fn composed_source(&self) -> Result<ComposedOverlaySource, OverlayError> {
        self.preflight_fingerprints()?;
        Ok(composed_source(
            self.source.clone(),
            Arc::from(self.spans.clone()),
            self.target_fingerprint,
        ))
    }

    /// Streams the complete composed artifact to a sequential sink.
    ///
    /// For a generic positional adapter, the full source and target
    /// fingerprints are rechecked before the first sink byte and again after
    /// emission. A plan created from explicitly owned immutable bytes relies
    /// on that sealed ownership instead. Both paths read and hash the source
    /// and target while emitting. Generic sinks are not called atomic:
    /// failures after progress return [`OverlayError::IncompleteOutput`].
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport, OverlayError> {
        if self.source.source_is_owned_immutable {
            return self.write_validated(writer, false);
        }
        self.preflight_fingerprints()?;
        self.write_validated(writer, true)
    }

    /// Publishes through a synced sibling temporary file and atomic rename.
    ///
    /// Before rename, every failure leaves the destination unchanged. After a
    /// successful rename, parent-directory sync failure is returned as
    /// [`OverlayError::Committed`] so callers do not retry as if the old file
    /// were still present.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<PublishReport, OverlayError> {
        let path = path.as_ref();
        let parent = parent_directory(path);
        // Explicitly owned immutable bytes cannot change while this plan is
        // alive. Skip only the two outer complete-artifact fences in that
        // sealed case; `write_validated` still hashes every source and target
        // byte while emitting, and publication keeps all flush/fsync/rename
        // durability steps below.
        if !self.source.source_is_owned_immutable {
            // Complete the potentially expensive source/fingerprint
            // validation before creating even a temporary file.
            self.preflight_fingerprints()?;
        }
        let (temporary_path, file) = create_sibling_temp_file(path)?;
        let result = (|| {
            let mut buffered = BufWriter::new(file);
            let report = self
                // `save` has its own mandatory pre-rename preflight below. The
                // direct sink path keeps the post-emission preflight inside
                // `write_validated`; skipping only that duplicate scan here
                // leaves source/target hashing during emission intact.
                .write_validated(&mut buffered, false)
                .map_err(strip_staging_progress)?;
            buffered.flush()?;
            buffered.get_ref().sync_all()?;

            if !self.source.source_is_owned_immutable {
                // Close the staging window with the same full identity check.
                // A stable but dishonest version token must not hide a late
                // source byte change immediately before the atomic rename.
                self.preflight_fingerprints()?;
            }
            drop(buffered);
            atomic_replace(&temporary_path, path)?;
            sync_parent(parent).map_err(|source| OverlayError::Committed { source })?;
            Ok(report)
        })();
        if result.is_err() {
            drop(fs::remove_file(&temporary_path));
        }
        result
    }

    fn preflight_fingerprints(&self) -> Result<(), OverlayError> {
        let (source, target) = fingerprints(&self.source, &self.spans)?;
        if source != self.source_fingerprint {
            return Err(OverlayError::SourceFingerprintChanged {
                expected: self.source_fingerprint,
                observed: source,
            });
        }
        if target != self.target_fingerprint {
            return Err(OverlayError::TargetFingerprintChanged {
                expected: self.target_fingerprint,
                observed: target,
            });
        }
        Ok(())
    }

    fn write_validated<W: Write>(
        &self,
        writer: &mut W,
        post_emission_preflight: bool,
    ) -> Result<PublishReport, OverlayError> {
        let mut buffer = publication_buffer()?;
        let mut source_hasher = Sha256::new();
        let mut target_hasher = Sha256::new();
        let mut offset = 0_u64;
        let mut accepted = 0_u64;

        while offset < self.source.length {
            let count = chunk_len(self.source.length, offset, buffer.len())?;
            if let Err(error) = self.source.read_exact(offset, &mut buffer[..count]) {
                return Err(with_progress(error, accepted, self.source.length, false));
            }
            source_hasher.update(&buffer[..count]);
            apply_spans(&mut buffer[..count], offset, &self.spans)?;
            target_hasher.update(&buffer[..count]);
            if let Err(failure) = write_all_checked(writer, &buffer[..count], &mut accepted) {
                return Err(with_progress(
                    OverlayError::Io(failure.error),
                    accepted,
                    self.source.length,
                    failure.indeterminate,
                ));
            }
            offset = offset
                .checked_add(count as u64)
                .ok_or_else(|| unavailable("publication offset overflow"))?;
        }

        let observed_source = ArtifactFingerprint(source_hasher.finalize().into());
        if observed_source != self.source_fingerprint {
            return Err(with_progress(
                OverlayError::SourceFingerprintChanged {
                    expected: self.source_fingerprint,
                    observed: observed_source,
                },
                accepted,
                self.source.length,
                false,
            ));
        }
        let observed_target = ArtifactFingerprint(target_hasher.finalize().into());
        if observed_target != self.target_fingerprint {
            return Err(with_progress(
                OverlayError::TargetFingerprintChanged {
                    expected: self.target_fingerprint,
                    observed: observed_target,
                },
                accepted,
                self.source.length,
                false,
            ));
        }
        // Direct sequential publication must retain this post-emission check:
        // a dishonest adapter may keep a stable version token while mutating a
        // chunk that has already been hashed and accepted by the sink. Atomic
        // `save` skips only this duplicate because it performs the same full
        // check after flush/fsync and immediately before rename.
        drop(buffer);
        if post_emission_preflight {
            if let Err(error) = self.preflight_fingerprints() {
                return Err(with_progress(error, accepted, self.source.length, false));
            }
        }
        if let Err(error) = writer.flush() {
            return Err(OverlayError::IncompleteOutput {
                progress: OutputProgress::CompleteUnflushed { bytes: accepted },
                source: Box::new(OverlayError::Io(error)),
            });
        }

        Ok(PublishReport {
            bytes: accepted,
            changed_spans: self.spans.len(),
            source_fingerprint: self.source_fingerprint,
            target_fingerprint: self.target_fingerprint,
        })
    }
}

pub(crate) fn finish_overlay_plan<F>(
    source: SourceSnapshot,
    spans: Vec<PhysicalSpan>,
    verify: F,
) -> Result<ValidatedOverlayPlan, OverlayError>
where
    F: FnOnce(&SourceSnapshot, &SharedOleFile) -> Result<(), OverlayError>,
{
    let (plan, _owner) = finish_overlay_plan_with_owner(source, spans, verify, |_candidate| {
        Ok::<(), OverlayError>(())
    })?;
    Ok(plan)
}

pub(crate) fn finish_overlay_plan_with_owner<T, E, F, V>(
    source: SourceSnapshot,
    spans: Vec<PhysicalSpan>,
    verify: F,
    validate_owner: V,
) -> Result<(ValidatedOverlayPlan, Option<T>), E>
where
    F: FnOnce(&SourceSnapshot, &SharedOleFile) -> Result<(), OverlayError>,
    V: FnOnce(&ComposedOverlaySource) -> Result<T, E>,
    E: From<OverlayError>,
{
    let (source_fingerprint, target_fingerprint) =
        fingerprints(&source, &spans).map_err(E::from)?;
    let view = Arc::new(composed_source(
        source.clone(),
        Arc::from(spans.clone()),
        target_fingerprint,
    ));
    let candidate_source: Arc<dyn ReadAt> = view.clone();
    let candidate = SharedOleFile::open(candidate_source)
        .map_err(OverlayError::from)
        .map_err(E::from)?;
    // Bracket caller-specific preconditions with complete fingerprints. This
    // ties checked logical ranges to the same stable source bytes retained by
    // the returned plan, even for adapters whose version token dishonestly
    // remains unchanged while bytes mutate.
    verify(&source, &candidate).map_err(E::from)?;
    let owner = if spans.is_empty() {
        None
    } else {
        Some(validate_owner(view.as_ref())?)
    };
    // Owned bytes and immutable spans cannot change while the owner callback
    // reads the composed view. Keep the final complete bracket for arbitrary
    // positional adapters, whose stable version token may be dishonest.
    if !source.source_is_owned_immutable {
        let (observed_source, observed_target) = fingerprints(&source, &spans).map_err(E::from)?;
        if observed_source != source_fingerprint {
            return Err(E::from(OverlayError::SourceFingerprintChanged {
                expected: source_fingerprint,
                observed: observed_source,
            }));
        }
        if observed_target != target_fingerprint {
            return Err(E::from(OverlayError::TargetFingerprintChanged {
                expected: target_fingerprint,
                observed: observed_target,
            }));
        }
    }
    Ok((
        ValidatedOverlayPlan {
            source,
            spans,
            source_fingerprint,
            target_fingerprint,
        },
        owner,
    ))
}

fn composed_source(
    source: SourceSnapshot,
    spans: Arc<[PhysicalSpan]>,
    target: ArtifactFingerprint,
) -> ComposedOverlaySource {
    let digest = target.as_bytes();
    let id = source.version.id()
        ^ u64::from_le_bytes(digest[..8].try_into().expect("fixed digest prefix"));
    let revision = source.version.revision()
        ^ u64::from_le_bytes(digest[8..16].try_into().expect("fixed digest prefix"));
    ComposedOverlaySource {
        source,
        spans,
        version: SourceVersion::new(id, revision),
    }
}

fn overlay_io_error(error: OverlayError) -> io::Error {
    match error {
        OverlayError::Io(source) => source,
        other => io::Error::other(other.to_string()),
    }
}

fn derive_selection_spans(
    source: &SourceSnapshot,
    index: &ParsedOleIndex,
    selection: &Selection,
    maximum: usize,
    comparison: &mut [u8],
    spans: &mut Vec<PhysicalSpan>,
) -> Result<(), OverlayError> {
    if selection.size == 0 {
        return Ok(());
    }
    let logical_size = usize::try_from(selection.size)
        .map_err(|_| unavailable("stream length does not fit this platform"))?;
    if selection.is_minifat {
        let root = index
            .root
            .as_ref()
            .ok_or_else(|| unavailable("mini stream has no root entry"))?;
        let root_size = usize::try_from(root.size)
            .map_err(|_| unavailable("root mini-stream length does not fit this platform"))?;
        let root_count = root_size.div_ceil(index.sector_size);
        let root_chain = collect_chain_exact(&index.fat, root.start_sector, root_count, "FAT")?;
        let mini_count = logical_size.div_ceil(index.mini_sector_size);
        let mini_chain = collect_chain_exact(
            &index.minifat,
            selection.start_sector,
            mini_count,
            "MiniFAT",
        )?;
        let mut logical = 0usize;
        for mini_sector in mini_chain {
            let mini_offset = usize::try_from(mini_sector)
                .map_err(|_| unavailable("mini-sector index does not fit usize"))?
                .checked_mul(index.mini_sector_size)
                .ok_or_else(|| unavailable("mini-sector offset overflow"))?;
            let count = index.mini_sector_size.min(logical_size - logical);
            if mini_offset
                .checked_add(count)
                .is_none_or(|end| end > root_size)
            {
                return Err(unavailable(
                    "mini-sector payload exceeds the declared root mini stream",
                ));
            }
            let root_ordinal = mini_offset / index.sector_size;
            let within_sector = mini_offset % index.sector_size;
            let root_sector = *root_chain
                .get(root_ordinal)
                .ok_or_else(|| unavailable("mini-sector is outside the root mini stream"))?;
            let physical = sector_offset(root_sector, index.sector_size)?
                .checked_add(within_sector as u64)
                .ok_or_else(|| unavailable("mini-sector physical offset overflow"))?;
            push_candidate_span(
                source,
                spans,
                maximum,
                comparison,
                physical,
                Arc::clone(&selection.replacement),
                logical..logical + count,
            )?;
            logical += count;
        }
    } else {
        let sector_count = logical_size.div_ceil(index.sector_size);
        let chain = collect_chain_exact(&index.fat, selection.start_sector, sector_count, "FAT")?;
        let mut logical = 0usize;
        for sector in chain {
            let count = index.sector_size.min(logical_size - logical);
            push_candidate_span(
                source,
                spans,
                maximum,
                comparison,
                sector_offset(sector, index.sector_size)?,
                Arc::clone(&selection.replacement),
                logical..logical + count,
            )?;
            logical += count;
        }
    }
    Ok(())
}

fn push_candidate_span(
    source: &SourceSnapshot,
    spans: &mut Vec<PhysicalSpan>,
    maximum: usize,
    comparison: &mut [u8],
    offset: u64,
    replacement: Arc<[u8]>,
    replacement_range: Range<usize>,
) -> Result<(), OverlayError> {
    let candidate = PhysicalSpan {
        offset,
        replacement,
        replacement_range,
    };
    if span_matches_source(source, &candidate, comparison)? {
        return Ok(());
    }
    if let Some(previous) = spans.last_mut()
        && Arc::ptr_eq(&previous.replacement, &candidate.replacement)
        && previous.end()? == candidate.offset
        && previous.replacement_range.end == candidate.replacement_range.start
    {
        previous.replacement_range.end = candidate.replacement_range.end;
        return Ok(());
    }
    if spans.len() == maximum {
        return Err(unavailable(format!(
            "changed physical span count exceeds limit {maximum}"
        )));
    }
    spans
        .try_reserve(1)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB physical overlay spans",
            source,
        })?;
    spans.push(candidate);
    Ok(())
}

fn span_matches_source(
    source: &SourceSnapshot,
    span: &PhysicalSpan,
    scratch: &mut [u8],
) -> Result<bool, OverlayError> {
    let replacement = &span.replacement[span.replacement_range.clone()];
    let mut compared = 0usize;
    while compared < replacement.len() {
        let count = scratch.len().min(replacement.len() - compared);
        let offset = span
            .offset
            .checked_add(compared as u64)
            .ok_or_else(|| unavailable("overlay comparison offset overflow"))?;
        source.read_exact(offset, &mut scratch[..count])?;
        if scratch[..count] != replacement[compared..compared + count] {
            return Ok(false);
        }
        compared += count;
    }
    Ok(true)
}

pub(crate) fn validate_and_coalesce_spans(
    source: &SourceSnapshot,
    limits: OverlayLimits,
    spans: &mut Vec<PhysicalSpan>,
) -> Result<(), OverlayError> {
    let mut checked = Vec::new();
    checked
        .try_reserve_exact(spans.len())
        .map_err(|source| OverlayError::Allocation {
            resource: "validated CFB physical overlay spans",
            source,
        })?;
    let mut changed_bytes = 0_u64;
    for span in spans.drain(..) {
        if span.len() == 0 {
            return Err(unavailable("physical overlay span is empty"));
        }
        let end = span.end()?;
        if end > source.length {
            return Err(unavailable("physical overlay span exceeds source length"));
        }
        changed_bytes = changed_bytes
            .checked_add(span.len() as u64)
            .ok_or_else(|| unavailable("changed physical byte total overflows u64"))?;
        if changed_bytes > limits.max_overlay_bytes {
            return Err(unavailable(format!(
                "changed physical bytes {changed_bytes} exceed limit {}",
                limits.max_overlay_bytes
            )));
        }
        if let Some(previous) = checked.last_mut() {
            let previous: &mut PhysicalSpan = previous;
            let previous_end = previous.end()?;
            if previous_end > span.offset {
                return Err(unavailable("physical overlay spans overlap"));
            }
            if previous_end == span.offset
                && Arc::ptr_eq(&previous.replacement, &span.replacement)
                && previous.replacement_range.end == span.replacement_range.start
            {
                previous.replacement_range.end = span.replacement_range.end;
                continue;
            }
        }
        checked.push(span);
    }
    if checked.len() > limits.max_spans {
        return Err(unavailable(format!(
            "changed physical span count {} exceeds limit {}",
            checked.len(),
            limits.max_spans
        )));
    }
    *spans = checked;
    Ok(())
}

fn validate_composed_artifact(
    source: &SourceSnapshot,
    spans: &[PhysicalSpan],
    selections: &[Selection],
) -> Result<(), OverlayError> {
    source.ensure_length()?;
    let cursor = OverlayCursor::new(source, spans);
    let mut candidate = OleFile::open(cursor)?;
    source.ensure_length()?;
    for selection in selections {
        let refs = path_refs(&selection.path)?;
        let bytes = candidate.open_stream(&refs)?;
        if bytes.as_slice() != selection.replacement.as_ref() {
            return Err(unavailable(format!(
                "composed stream {:?} differs after CFB reopen",
                selection.path
            )));
        }
    }
    source.ensure_length()
}

fn fingerprints(
    source: &SourceSnapshot,
    spans: &[PhysicalSpan],
) -> Result<(ArtifactFingerprint, ArtifactFingerprint), OverlayError> {
    source.ensure_length()?;
    let mut buffer = fingerprint_buffer(source.length)?;
    let mut source_hasher = Sha256::new();
    let mut target_hasher = Sha256::new();
    let mut offset = 0_u64;
    while offset < source.length {
        let count = chunk_len(source.length, offset, buffer.len())?;
        source.read_exact(offset, &mut buffer[..count])?;
        source_hasher.update(&buffer[..count]);
        apply_spans(&mut buffer[..count], offset, spans)?;
        target_hasher.update(&buffer[..count]);
        offset = offset
            .checked_add(count as u64)
            .ok_or_else(|| unavailable("fingerprint offset overflow"))?;
    }
    source.ensure_length()?;
    Ok((
        ArtifactFingerprint(source_hasher.finalize().into()),
        ArtifactFingerprint(target_hasher.finalize().into()),
    ))
}

pub(crate) fn path_refs(path: &[String]) -> Result<Vec<&str>, OverlayError> {
    let mut refs = Vec::new();
    refs.try_reserve_exact(path.len())
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB overlay path references",
            source,
        })?;
    refs.extend(path.iter().map(String::as_str));
    Ok(refs)
}

fn apply_spans(bytes: &mut [u8], offset: u64, spans: &[PhysicalSpan]) -> Result<(), OverlayError> {
    let end = offset
        .checked_add(bytes.len() as u64)
        .ok_or_else(|| unavailable("publication chunk end overflow"))?;
    let first =
        spans.partition_point(|span| span.offset.saturating_add(span.len() as u64) <= offset);
    for span in &spans[first..] {
        let span_end = span.end()?;
        if span.offset >= end {
            break;
        }
        let intersection_start = span.offset.max(offset);
        let intersection_end = span_end.min(end);
        let output_start = usize::try_from(intersection_start - offset)
            .map_err(|_| unavailable("chunk-relative overlay offset does not fit usize"))?;
        let output_end = usize::try_from(intersection_end - offset)
            .map_err(|_| unavailable("chunk-relative overlay end does not fit usize"))?;
        let replacement_start = span
            .replacement_range
            .start
            .checked_add(
                usize::try_from(intersection_start - span.offset)
                    .map_err(|_| unavailable("replacement-relative offset does not fit usize"))?,
            )
            .ok_or_else(|| unavailable("replacement range offset overflow"))?;
        let replacement_end = replacement_start
            .checked_add(output_end - output_start)
            .ok_or_else(|| unavailable("replacement range end overflow"))?;
        bytes[output_start..output_end]
            .copy_from_slice(&span.replacement[replacement_start..replacement_end]);
    }
    Ok(())
}

pub(crate) fn collect_chain_exact(
    table: &[u32],
    start: u32,
    expected: usize,
    name: &str,
) -> Result<Vec<u32>, OverlayError> {
    if expected == 0 {
        if start != ENDOFCHAIN {
            return Err(unavailable(format!(
                "empty {name} chain does not start with ENDOFCHAIN"
            )));
        }
        return Ok(Vec::new());
    }
    if start >= MAXREGSECT || expected > table.len() {
        return Err(unavailable(format!("invalid {name} chain length")));
    }
    let mut chain = Vec::new();
    chain
        .try_reserve_exact(expected)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB overlay sector chain",
            source,
        })?;
    let mut sector = start;
    for ordinal in 0..expected {
        let index = usize::try_from(sector)
            .map_err(|_| unavailable(format!("{name} sector does not fit usize")))?;
        let next = *table
            .get(index)
            .ok_or_else(|| unavailable(format!("{name} sector is outside its table")))?;
        chain.push(sector);
        if ordinal + 1 == expected {
            if next != ENDOFCHAIN {
                return Err(unavailable(format!(
                    "{name} chain exceeds declared stream length"
                )));
            }
        } else {
            if next == ENDOFCHAIN || next >= MAXREGSECT {
                return Err(unavailable(format!(
                    "{name} chain ends before declared stream length"
                )));
            }
            sector = next;
        }
    }
    Ok(chain)
}

pub(crate) fn sector_offset(sector: u32, sector_size: usize) -> Result<u64, OverlayError> {
    (u64::from(sector) + 1)
        .checked_mul(sector_size as u64)
        .ok_or_else(|| unavailable("physical sector offset overflow"))
}

fn publication_buffer() -> Result<Vec<u8>, OverlayError> {
    fixed_buffer(PUBLICATION_CHUNK_BYTES, "CFB overlay publication buffer")
}

fn fingerprint_buffer(source_length: u64) -> Result<Vec<u8>, OverlayError> {
    let bytes = usize::try_from(source_length.min(FINGERPRINT_CHUNK_BYTES_U64))
        .map_err(|_| unavailable("CFB overlay fingerprint buffer length does not fit usize"))?;
    debug_assert!(bytes <= FINGERPRINT_CHUNK_BYTES);
    fixed_buffer(bytes, "CFB overlay fingerprint buffer")
}

fn fixed_buffer(bytes: usize, resource: &'static str) -> Result<Vec<u8>, OverlayError> {
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(bytes)
        .map_err(|source| OverlayError::Allocation { resource, source })?;
    buffer.resize(bytes, 0);
    Ok(buffer)
}

fn chunk_len(length: u64, offset: u64, maximum: usize) -> Result<usize, OverlayError> {
    usize::try_from((length - offset).min(maximum as u64))
        .map_err(|_| unavailable("publication chunk length does not fit usize"))
}

struct WriteFailure {
    error: io::Error,
    indeterminate: bool,
}

fn write_all_checked<W: Write>(
    writer: &mut W,
    mut bytes: &[u8],
    accepted: &mut u64,
) -> Result<(), WriteFailure> {
    while !bytes.is_empty() {
        match writer.write(bytes) {
            Ok(0) => {
                return Err(WriteFailure {
                    error: io::Error::new(
                        io::ErrorKind::WriteZero,
                        "CFB overlay sink accepted no bytes",
                    ),
                    indeterminate: false,
                });
            },
            Ok(written) if written > bytes.len() => {
                return Err(WriteFailure {
                    error: io::Error::new(
                        io::ErrorKind::InvalidData,
                        "CFB overlay sink reported more bytes than offered",
                    ),
                    indeterminate: true,
                });
            },
            Ok(written) => {
                *accepted = accepted.saturating_add(written as u64);
                bytes = &bytes[written..];
            },
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => {
                return Err(WriteFailure {
                    error,
                    indeterminate: false,
                });
            },
        }
    }
    Ok(())
}

fn with_progress(
    error: OverlayError,
    accepted: u64,
    expected: u64,
    indeterminate: bool,
) -> OverlayError {
    if indeterminate {
        return OverlayError::IncompleteOutput {
            progress: OutputProgress::Indeterminate {
                accepted_before: accepted,
            },
            source: Box::new(error),
        };
    }
    if accepted == 0 {
        return error;
    }
    let progress = if accepted == expected {
        OutputProgress::CompleteUnflushed { bytes: accepted }
    } else {
        OutputProgress::Prefix { accepted, expected }
    };
    OverlayError::IncompleteOutput {
        progress,
        source: Box::new(error),
    }
}

fn strip_staging_progress(error: OverlayError) -> OverlayError {
    match error {
        OverlayError::IncompleteOutput { source, .. } => strip_staging_progress(*source),
        other => other,
    }
}

pub(crate) fn unavailable(reason: impl Into<String>) -> OverlayError {
    OverlayError::Unavailable {
        reason: reason.into(),
    }
}

struct OverlayCursor<'a> {
    source: &'a SourceSnapshot,
    spans: &'a [PhysicalSpan],
    position: u64,
}

impl<'a> OverlayCursor<'a> {
    fn new(source: &'a SourceSnapshot, spans: &'a [PhysicalSpan]) -> Self {
        Self {
            source,
            spans,
            position: 0,
        }
    }
}

impl Read for OverlayCursor<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || self.position >= self.source.length {
            return Ok(0);
        }
        let count = usize::try_from((self.source.length - self.position).min(output.len() as u64))
            .map_err(|_| io::Error::other("overlay cursor range does not fit usize"))?;
        self.source
            .ensure_current()
            .map_err(|error| io::Error::other(error.to_string()))?;
        let read = self
            .source
            .source
            .read_at(self.position, &mut output[..count])?;
        if read > count {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        apply_spans(&mut output[..read], self.position, self.spans)
            .map_err(|error| io::Error::other(error.to_string()))?;
        self.position = self
            .position
            .checked_add(read as u64)
            .ok_or_else(|| io::Error::other("overlay cursor position overflow"))?;
        self.source
            .ensure_current()
            .map_err(|error| io::Error::other(error.to_string()))?;
        Ok(read)
    }
}

impl Seek for OverlayCursor<'_> {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let position = match from {
            SeekFrom::Start(position) => i128::from(position),
            SeekFrom::End(delta) => i128::from(self.source.length) + i128::from(delta),
            SeekFrom::Current(delta) => i128::from(self.position) + i128::from(delta),
        };
        self.position = u64::try_from(position).map_err(|_| {
            io::Error::new(io::ErrorKind::InvalidInput, "invalid overlay cursor seek")
        })?;
        Ok(self.position)
    }
}
