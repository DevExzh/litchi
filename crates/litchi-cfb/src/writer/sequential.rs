//! Forward-only, predeclared-layout CFB authoring.
//!
//! [`SequentialOleWriter`] is deliberately narrower than [`super::OleWriter`].
//! It accepts each stream once together with its exact length, plans every
//! directory and allocation-table byte before touching the destination, and
//! then emits the physical file in sector order.  The source payloads are not
//! retained by the plan; only bounded metadata and one reusable publication
//! buffer are held in memory.  This is payload-bounded rather than O(1) in
//! the number of directory or allocation-table entries: metadata grows with
//! the declared topology and is rejected before serialized buffers are built.

use super::difat::DifatBuilder;
use super::directory::DirectoryBuilder;
use super::fat::FatBuilder;
use super::header::HeaderBuilder;
use super::minifat::MiniFatBuilder;
use super::{atomic_replace, create_sibling_temp_file, parent_directory, sync_parent};
use crate::consts::{DIFSECT, ENDOFCHAIN, FATSECT, MAXREGSECT};
use crate::directory_name::directory_name_data;
use crate::file::{OleError, OleFile};
use litchi_core::CancellationToken;
use std::fmt;
use std::fs::{self, File};
use std::io::{self, BufWriter, ErrorKind, Read, Write};
use std::path::Path;

const MINI_STREAM_CUTOFF: u64 = 4096;
const MINI_SECTOR_SIZE: usize = 64;
const DEFAULT_PUBLICATION_BUFFER_BYTES: usize = 64 * 1024;
const DEFAULT_MAX_STREAMS: u64 = 1_048_576;
const DEFAULT_MAX_DIRECTORY_ENTRIES: u64 = 1_048_576;
const DEFAULT_MAX_PATH_COMPONENTS: u64 = 1024;
const DEFAULT_MAX_PATH_BYTES: u64 = 256 * 1024;
const DEFAULT_MAX_STREAM_BYTES: u64 = 2 * 1024 * 1024 * 1024 - 1;
const DEFAULT_MAX_METADATA_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TempFileIdentity {
    #[cfg(unix)]
    Unix { device: u64, inode: u64 },
    #[cfg(windows)]
    Windows { volume: u32, index: u64 },
    // Targets without a native file identity get only a length discriminator.
    // The public save contract therefore requires a trusted/private parent;
    // this fallback is not protection against a hostile shared directory.
    #[cfg(not(any(unix, windows)))]
    Fallback { length: u64 },
}

fn file_identity(file: &File) -> io::Result<TempFileIdentity> {
    identity_from_metadata(&file.metadata()?)
}

fn path_identity(path: &Path) -> io::Result<TempFileIdentity> {
    let metadata = fs::symlink_metadata(path)?;
    if !metadata.file_type().is_file() {
        return Err(io::Error::new(
            ErrorKind::InvalidData,
            "CFB temporary path is not a regular file",
        ));
    }
    identity_from_metadata(&metadata)
}

fn identity_from_metadata(metadata: &fs::Metadata) -> io::Result<TempFileIdentity> {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;

        return Ok(TempFileIdentity::Unix {
            device: metadata.dev(),
            inode: metadata.ino(),
        });
    }

    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;

        let volume = metadata.volume_serial_number().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "CFB temporary file has no volume identity",
            )
        })?;
        let index = metadata.file_index().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "CFB temporary file has no file identity",
            )
        })?;
        return Ok(TempFileIdentity::Windows { volume, index });
    }

    #[cfg(not(any(unix, windows)))]
    {
        Ok(TempFileIdentity::Fallback {
            length: metadata.len(),
        })
    }
}

fn ensure_temp_identity(temporary_path: &Path, expected: TempFileIdentity) -> io::Result<()> {
    let observed = path_identity(temporary_path)?;
    if observed == expected {
        Ok(())
    } else {
        Err(io::Error::new(
            ErrorKind::AlreadyExists,
            "CFB temporary path changed while staged output was open",
        ))
    }
}

fn cleanup_owned_temp(temporary_path: &Path, expected: TempFileIdentity) {
    // Best-effort identity checking preserves a known replacement.  The save
    // contract still requires a trusted/private parent because portable,
    // path-based cleanup and replacement cannot close a hostile-directory
    // race between this check and the filesystem operation.
    if path_identity(temporary_path).is_ok_and(|observed| observed == expected) {
        drop(fs::remove_file(temporary_path));
    }
}

struct TemporaryCleanupGuard<'a> {
    path: &'a Path,
    identity: Option<TempFileIdentity>,
    published: bool,
}

impl<'a> TemporaryCleanupGuard<'a> {
    fn new(path: &'a Path) -> Self {
        Self {
            path,
            identity: None,
            published: false,
        }
    }

    fn set_identity(&mut self, identity: TempFileIdentity) {
        self.identity = Some(identity);
    }

    fn mark_published(&mut self) {
        self.published = true;
    }
}

impl Drop for TemporaryCleanupGuard<'_> {
    fn drop(&mut self) {
        if self.published {
            return;
        }

        if let Some(identity) = self.identity {
            cleanup_owned_temp(self.path, identity);
        } else {
            // Identity acquisition failed before we could compare the path.
            // This exact-name cleanup is correct only under save's
            // trusted/private-parent contract, but prevents an orphaned temp.
            drop(fs::remove_file(self.path));
        }
    }
}

/// What a forward-only sink may already contain after a failed publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SequentialWriteProgress {
    /// No output byte was accepted.
    Untouched,
    /// An exact prefix of the planned artifact was accepted.
    Prefix {
        /// Bytes definitely accepted by the sink.
        accepted: u64,
        /// Complete planned artifact length.
        expected: u64,
    },
    /// Every artifact byte was accepted and the sink flush completed.
    Complete {
        /// Complete planned artifact length.
        bytes: u64,
    },
    /// Every artifact byte was accepted, but the final sink flush failed.
    CompleteUnflushed {
        /// Complete planned artifact length.
        bytes: u64,
    },
    /// A hostile sink reported more bytes than it was given.
    Indeterminate {
        /// Bytes definitely accepted before the invalid report.
        accepted_before: u64,
    },
}

impl SequentialWriteProgress {
    /// Returns bytes definitely accepted before the failure.
    #[must_use]
    pub const fn accepted(self) -> u64 {
        match self {
            Self::Untouched => 0,
            Self::Prefix { accepted, .. } => accepted,
            Self::Complete { bytes } => bytes,
            Self::CompleteUnflushed { bytes } => bytes,
            Self::Indeterminate { accepted_before } => accepted_before,
        }
    }

    /// Returns whether the sink accepted every artifact byte.
    #[must_use]
    pub const fn is_complete(self) -> bool {
        matches!(self, Self::Complete { .. } | Self::CompleteUnflushed { .. })
    }
}

/// Finite bounds for one [`SequentialOleWriter`] plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SequentialWriterLimits {
    /// Maximum number of registered streams.
    pub max_streams: u64,
    /// Maximum number of directory entries, including implicit storages.
    pub max_directory_entries: u64,
    /// Maximum path depth in components.
    pub max_path_components: u64,
    /// Maximum aggregate UTF-8 path bytes.
    pub max_path_bytes: u64,
    /// Maximum length of one stream payload.
    pub max_stream_bytes: u64,
    /// Maximum serialized metadata bytes retained by the plan.
    pub max_metadata_bytes: u64,
    /// Maximum complete artifact bytes, including the 512-byte/4096-byte header.
    pub max_output_bytes: u64,
}

impl SequentialWriterLimits {
    /// Creates explicit finite limits.
    #[must_use]
    pub const fn new(
        max_streams: u64,
        max_directory_entries: u64,
        max_path_components: u64,
        max_path_bytes: u64,
        max_stream_bytes: u64,
        max_metadata_bytes: u64,
        max_output_bytes: u64,
    ) -> Self {
        Self {
            max_streams,
            max_directory_entries,
            max_path_components,
            max_path_bytes,
            max_stream_bytes,
            max_metadata_bytes,
            max_output_bytes,
        }
    }
}

impl Default for SequentialWriterLimits {
    fn default() -> Self {
        Self {
            max_streams: DEFAULT_MAX_STREAMS,
            max_directory_entries: DEFAULT_MAX_DIRECTORY_ENTRIES,
            max_path_components: DEFAULT_MAX_PATH_COMPONENTS,
            max_path_bytes: DEFAULT_MAX_PATH_BYTES,
            max_stream_bytes: DEFAULT_MAX_STREAM_BYTES,
            max_metadata_bytes: DEFAULT_MAX_METADATA_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
        }
    }
}

/// Options for one predeclared-layout writer.
#[derive(Clone, Debug)]
pub struct SequentialWriterOptions {
    /// Regular CFB sector size.  Only 512 and 4096 are valid.
    pub sector_size: usize,
    /// Finite semantic and output limits.
    pub limits: SequentialWriterLimits,
    /// Capacity of the one reusable source/publication buffer.
    pub publication_buffer_bytes: usize,
    /// Optional cooperative cancellation token.
    pub cancellation: Option<CancellationToken>,
}

impl SequentialWriterOptions {
    /// Creates options with explicit geometry, limits, and buffer size.
    #[must_use]
    pub fn new(
        sector_size: usize,
        limits: SequentialWriterLimits,
        publication_buffer_bytes: usize,
    ) -> Self {
        Self {
            sector_size,
            limits,
            publication_buffer_bytes,
            cancellation: None,
        }
    }

    /// Sets the cancellation token used at publication checkpoints.
    #[must_use]
    pub fn with_cancellation(mut self, cancellation: CancellationToken) -> Self {
        self.cancellation = Some(cancellation);
        self
    }

    /// Sets a sector size and returns the updated options.
    #[must_use]
    pub fn with_sector_size(mut self, sector_size: usize) -> Self {
        self.sector_size = sector_size;
        self
    }

    /// Sets the finite limits.
    #[must_use]
    pub fn with_limits(mut self, limits: SequentialWriterLimits) -> Self {
        self.limits = limits;
        self
    }

    /// Sets the reusable publication buffer capacity.
    #[must_use]
    pub fn with_publication_buffer_bytes(mut self, bytes: usize) -> Self {
        self.publication_buffer_bytes = bytes;
        self
    }
}

impl Default for SequentialWriterOptions {
    fn default() -> Self {
        Self {
            sector_size: 512,
            limits: SequentialWriterLimits::default(),
            publication_buffer_bytes: DEFAULT_PUBLICATION_BUFFER_BYTES,
            cancellation: None,
        }
    }
}

/// Exact accounting returned after a successful sequential publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SequentialWriteReport {
    output_bytes: u64,
    payload_bytes: u64,
    metadata_bytes: u64,
    stream_count: u64,
    publication_buffer_bytes: u64,
}

impl SequentialWriteReport {
    /// Complete artifact length.
    #[must_use]
    pub const fn output_bytes(self) -> u64 {
        self.output_bytes
    }

    /// Sum of declared stream payload lengths, excluding CFB padding.
    #[must_use]
    pub const fn payload_bytes(self) -> u64 {
        self.payload_bytes
    }

    /// Artifact bytes not counted as raw stream payload, including padding and metadata.
    #[must_use]
    pub const fn metadata_bytes(self) -> u64 {
        self.metadata_bytes
    }

    /// Number of registered streams.
    #[must_use]
    pub const fn stream_count(self) -> u64 {
        self.stream_count
    }

    /// Fixed reusable publication-buffer capacity.
    #[must_use]
    pub const fn publication_buffer_bytes(self) -> u64 {
        self.publication_buffer_bytes
    }

    /// Complete artifact length (compatibility shorthand).
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.output_bytes
    }
}

/// Typed failures from the predeclared-layout writer.
#[derive(Debug)]
pub enum SequentialWriteError {
    /// A planning, validation, allocation, or metadata-generation failure.
    Planning(OleError),
    /// A source reader failed after publication began.
    SourceIo {
        /// Logical stream path.
        path: Vec<String>,
        /// Reader failure.
        source: io::Error,
        /// Sink progress at the failure.
        progress: SequentialWriteProgress,
    },
    /// A source ended early or supplied an extra byte.
    SourceLength {
        /// Logical stream path.
        path: Vec<String>,
        /// Declared source length.
        expected: u64,
        /// Bytes observed from the source (for an extra byte this is expected + 1).
        observed: u64,
        /// Sink progress at the failure.
        progress: SequentialWriteProgress,
    },
    /// The destination sink returned an I/O failure after accepting a prefix.
    Sink {
        /// Sink failure.
        source: io::Error,
        /// Sink progress at the failure.
        progress: SequentialWriteProgress,
    },
    /// The destination sink returned zero for a nonempty write.
    WriteZero {
        /// Sink progress at the failure.
        progress: SequentialWriteProgress,
    },
    /// A sink reported an impossible byte count, so exact progress is unknown.
    Indeterminate {
        /// Bytes definitely accepted before the invalid report.
        accepted_before: u64,
    },
    /// Cooperative cancellation was observed.
    Cancelled {
        /// Sink progress at cancellation.
        progress: SequentialWriteProgress,
    },
    /// A finite semantic or output bound rejected the plan.
    LimitExceeded {
        /// Bound name.
        resource: &'static str,
        /// Requested/observed amount.
        observed: u64,
        /// Configured maximum.
        limit: u64,
    },
    /// All bytes were accepted, but the final sink flush failed.
    Flush {
        /// Flush failure.
        source: io::Error,
        /// Always [`SequentialWriteProgress::CompleteUnflushed`].
        progress: SequentialWriteProgress,
    },
    /// The destination was replaced but the parent-directory durability step failed.
    Committed {
        /// Parent synchronization failure.
        source: io::Error,
        /// Complete staged artifact report.
        report: SequentialWriteReport,
        /// Publication state after replacement.
        progress: SequentialWriteProgress,
    },
    /// The staged artifact was complete, but synchronization or replacement failed.
    Stage {
        /// Staging failure.
        source: io::Error,
        /// Staged output state.
        progress: SequentialWriteProgress,
        /// Complete staged artifact report.
        report: SequentialWriteReport,
    },
    /// The staged candidate was read successfully but failed full CFB validation.
    CandidateValidation {
        /// Candidate validation failure.
        source: OleError,
        /// Staged output state.
        progress: SequentialWriteProgress,
        /// Complete staged artifact report.
        report: SequentialWriteReport,
    },
}

impl SequentialWriteError {
    /// Returns the best available sink-progress evidence.
    #[must_use]
    pub const fn progress(&self) -> SequentialWriteProgress {
        match self {
            Self::Planning(_) | Self::LimitExceeded { .. } => SequentialWriteProgress::Untouched,
            Self::SourceIo { progress, .. }
            | Self::SourceLength { progress, .. }
            | Self::Sink { progress, .. }
            | Self::WriteZero { progress }
            | Self::Cancelled { progress }
            | Self::Flush { progress, .. } => *progress,
            Self::Indeterminate { accepted_before } => SequentialWriteProgress::Indeterminate {
                accepted_before: *accepted_before,
            },
            Self::Committed { progress, .. }
            | Self::Stage { progress, .. }
            | Self::CandidateValidation { progress, .. } => *progress,
        }
    }
}

impl fmt::Display for SequentialWriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Planning(error) => {
                write!(formatter, "CFB sequential writer planning failed: {error}")
            },
            Self::SourceIo { path, source, .. } => {
                write!(
                    formatter,
                    "CFB source reader failed at {}: {source}",
                    display_path(path)
                )
            },
            Self::SourceLength {
                path,
                expected,
                observed,
                ..
            } => write!(
                formatter,
                "CFB source length mismatch at {}: expected {expected} bytes, observed {observed}",
                display_path(path)
            ),
            Self::Sink { source, .. } => write!(formatter, "CFB sequential sink failed: {source}"),
            Self::WriteZero { .. } => {
                formatter.write_str("CFB sequential sink returned a zero-length write")
            },
            Self::Indeterminate { accepted_before } => write!(
                formatter,
                "CFB sequential sink over-reported a write after {accepted_before} accepted bytes"
            ),
            Self::Cancelled { .. } => {
                formatter.write_str("CFB sequential publication was cancelled")
            },
            Self::LimitExceeded {
                resource,
                observed,
                limit,
            } => write!(
                formatter,
                "CFB sequential writer limit exceeded for {resource}: observed {observed}, limit {limit}"
            ),
            Self::Flush { source, .. } => {
                write!(formatter, "CFB sequential sink flush failed: {source}")
            },
            Self::Committed { source, .. } => write!(
                formatter,
                "CFB destination was replaced but directory durability could not be confirmed: {source}"
            ),
            Self::Stage { source, .. } => {
                write!(
                    formatter,
                    "CFB staged CFB artifact could not be published: {source}"
                )
            },
            Self::CandidateValidation { source, .. } => {
                write!(
                    formatter,
                    "CFB staged candidate failed validation: {source}"
                )
            },
        }
    }
}

impl std::error::Error for SequentialWriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Planning(error) => Some(error),
            Self::SourceIo { source, .. }
            | Self::Sink { source, .. }
            | Self::Flush { source, .. }
            | Self::Committed { source, .. }
            | Self::Stage { source, .. } => Some(source),
            Self::CandidateValidation { source, .. } => Some(source),
            Self::SourceLength { .. }
            | Self::WriteZero { .. }
            | Self::Indeterminate { .. }
            | Self::Cancelled { .. }
            | Self::LimitExceeded { .. } => None,
        }
    }
}

impl From<OleError> for SequentialWriteError {
    fn from(error: OleError) -> Self {
        Self::Planning(error)
    }
}

struct StreamInput<'a> {
    path: Vec<String>,
    declared_len: u64,
    source: Box<dyn Read + 'a>,
    start_sector: u32,
    start_mini_sector: u32,
}

struct StorageInput {
    path: Vec<String>,
    clsid: Option<[u8; 16]>,
}

/// A single-use writer for a fresh CFB artifact.
///
/// The writer's memory contract is bounded by the finite metadata and path
/// limits in [`SequentialWriterLimits`] plus one reusable publication buffer;
/// it is not constant-memory in the number of streams or sectors. Declared
/// payload bytes are read once from their sources and are never retained by
/// the plan.
pub struct SequentialOleWriter<'a> {
    options: SequentialWriterOptions,
    root_clsid: Option<[u8; 16]>,
    storages: Vec<StorageInput>,
    streams: Vec<StreamInput<'a>>,
    path_bytes: u64,
}

impl<'a> SequentialOleWriter<'a> {
    /// Creates a writer with 512-byte sectors and finite default limits.
    #[must_use]
    pub fn new() -> Self {
        Self {
            options: SequentialWriterOptions::default(),
            root_clsid: None,
            storages: Vec::new(),
            streams: Vec::new(),
            path_bytes: 0,
        }
    }

    /// Creates a writer with explicit options.
    pub fn with_options(options: SequentialWriterOptions) -> Result<Self, SequentialWriteError> {
        validate_options(&options)?;
        Ok(Self {
            options,
            root_clsid: None,
            storages: Vec::new(),
            streams: Vec::new(),
            path_bytes: 0,
        })
    }

    /// Creates a writer for one of the two legal CFB regular sector sizes.
    pub fn with_sector_size(sector_size: usize) -> Result<Self, SequentialWriteError> {
        Self::with_options(SequentialWriterOptions::default().with_sector_size(sector_size))
    }

    /// Returns the retained writer options.
    #[must_use]
    pub const fn options(&self) -> &SequentialWriterOptions {
        &self.options
    }

    /// Sets the root storage CLSID.
    pub fn set_root_clsid(&mut self, clsid: [u8; 16]) {
        self.root_clsid = Some(clsid);
    }

    /// Declares a storage path.  Missing parent storages are created during planning.
    pub fn create_storage(&mut self, path: &[&str]) -> Result<(), SequentialWriteError> {
        let path_bytes_before = self.path_bytes;
        let path = self.own_path(path, "storage path")?;
        if self.storages.iter().any(|storage| storage.path == path) {
            self.path_bytes = path_bytes_before;
            return Ok(());
        }
        if let Err(source) = self.storages.try_reserve(1) {
            self.path_bytes = path_bytes_before;
            return Err(OleError::allocation("sequential storage table", source).into());
        }
        self.storages.push(StorageInput { path, clsid: None });
        Ok(())
    }

    /// Sets a storage CLSID, declaring the storage when necessary.
    pub fn set_storage_clsid(
        &mut self,
        path: &[&str],
        clsid: [u8; 16],
    ) -> Result<(), SequentialWriteError> {
        let path_bytes_before = self.path_bytes;
        let path = self.own_path(path, "storage path")?;
        if path.is_empty() {
            self.path_bytes = path_bytes_before;
            return Err(planning("CFB storage CLSID path must not be empty"));
        }
        if let Some(storage) = self
            .storages
            .iter_mut()
            .find(|storage| storage.path == path)
        {
            storage.clsid = Some(clsid);
            self.path_bytes = path_bytes_before;
            return Ok(());
        }
        if let Err(source) = self.storages.try_reserve(1) {
            self.path_bytes = path_bytes_before;
            return Err(OleError::allocation("sequential storage table", source).into());
        }
        self.storages.push(StorageInput {
            path,
            clsid: Some(clsid),
        });
        Ok(())
    }

    /// Registers one single-use source and its exact byte length.
    pub fn add_stream<R: Read + 'a>(
        &mut self,
        path: &[&str],
        declared_len: u64,
        source: R,
    ) -> Result<(), SequentialWriteError> {
        let path_bytes_before = self.path_bytes;
        let path = self.own_path(path, "stream path")?;
        let limits = self.options.limits;
        let stream_count = u64::try_from(self.streams.len()).unwrap_or(u64::MAX);
        if stream_count >= limits.max_streams {
            self.path_bytes = path_bytes_before;
            return Err(limit(
                "streams",
                stream_count.saturating_add(1),
                limits.max_streams,
            ));
        }
        if declared_len > limits.max_stream_bytes {
            self.path_bytes = path_bytes_before;
            return Err(limit("stream bytes", declared_len, limits.max_stream_bytes));
        }
        if self.streams.iter().any(|stream| stream.path == path) {
            self.path_bytes = path_bytes_before;
            return Err(planning("duplicate CFB stream path"));
        }
        if let Err(source) = self.streams.try_reserve(1) {
            self.path_bytes = path_bytes_before;
            return Err(OleError::allocation("sequential stream table", source).into());
        }
        self.streams.push(StreamInput {
            path,
            declared_len,
            source: Box::new(source),
            start_sector: ENDOFCHAIN,
            start_mini_sector: ENDOFCHAIN,
        });
        Ok(())
    }

    /// Emits the complete artifact to a non-seek sink and flushes it.
    ///
    /// Planning consumes no source bytes and performs no sink call.  The
    /// writer is consumed on success or failure because every source is
    /// single-use.
    pub fn write_to<W: Write>(
        self,
        sink: &mut W,
    ) -> Result<SequentialWriteReport, SequentialWriteError> {
        let mut plan = self.plan()?;
        if plan
            .cancellation
            .as_ref()
            .is_some_and(CancellationToken::is_cancelled)
        {
            return Err(SequentialWriteError::Cancelled {
                progress: SequentialWriteProgress::Untouched,
            });
        }

        let mut accepted = 0_u64;
        let mut buffer = plan.publication_buffer;
        let expected = plan.output_bytes;

        publish_segment(
            sink,
            &plan.header,
            &mut accepted,
            expected,
            &plan.cancellation,
        )?;

        for &index in &plan.large_streams {
            emit_stream(
                sink,
                &mut plan.streams[index],
                plan.sector_size,
                &mut buffer,
                &mut accepted,
                expected,
                &plan.cancellation,
            )?;
        }
        if !plan.small_streams.is_empty() {
            for &index in &plan.small_streams {
                emit_stream(
                    sink,
                    &mut plan.streams[index],
                    MINI_SECTOR_SIZE,
                    &mut buffer,
                    &mut accepted,
                    expected,
                    &plan.cancellation,
                )?;
            }
            let mini_stream_bytes = plan
                .streams
                .iter()
                .filter(|stream| stream.declared_len < MINI_STREAM_CUTOFF)
                .try_fold(0_u64, |total, stream| {
                    let padded =
                        padded_len(stream.declared_len, MINI_SECTOR_SIZE).ok_or_else(|| {
                            SequentialWriteError::Planning(OleError::InvalidData(
                                "CFB ministream padding overflows u64".to_string(),
                            ))
                        })?;
                    total.checked_add(padded).ok_or_else(|| {
                        SequentialWriteError::Planning(OleError::InvalidData(
                            "CFB ministream size overflows u64".to_string(),
                        ))
                    })
                })?;
            let regular_padding = padded_len(mini_stream_bytes, plan.sector_size)
                .and_then(|padded| padded.checked_sub(mini_stream_bytes))
                .ok_or_else(|| {
                    SequentialWriteError::Planning(OleError::InvalidData(
                        "CFB ministream regular padding overflows u64".to_string(),
                    ))
                })?;
            publish_zeroes(
                sink,
                regular_padding,
                &mut buffer,
                &mut accepted,
                expected,
                &plan.cancellation,
            )?;
        }

        publish_padded(
            sink,
            &plan.directory,
            plan.sector_size,
            &mut buffer,
            &mut accepted,
            expected,
            &plan.cancellation,
        )?;
        for sector in &plan.minifat_sectors {
            publish_segment(sink, sector, &mut accepted, expected, &plan.cancellation)?;
        }
        for sector in &plan.difat_sectors {
            publish_segment(sink, sector, &mut accepted, expected, &plan.cancellation)?;
        }
        for sector in &plan.fat_sectors {
            publish_segment(sink, sector, &mut accepted, expected, &plan.cancellation)?;
        }

        if accepted != expected {
            return Err(SequentialWriteError::Planning(OleError::InvalidData(
                "CFB sequential physical emission did not reach planned output size".to_string(),
            )));
        }
        if let Some(token) = &plan.cancellation {
            if token.is_cancelled() {
                return Err(SequentialWriteError::Cancelled {
                    progress: SequentialWriteProgress::Prefix { accepted, expected },
                });
            }
        }
        sink.flush().map_err(|source| SequentialWriteError::Flush {
            source,
            progress: SequentialWriteProgress::CompleteUnflushed { bytes: expected },
        })?;
        Ok(plan.report)
    }

    /// Saves atomically through the established sibling-temp helper.
    ///
    /// The destination parent must be trusted and private to this operation
    /// for the duration of the call.  Portable path-based replacement cannot
    /// guarantee safety if another actor can unlink and recreate sibling
    /// temporary names.  Identity checks detect known substitutions as a
    /// best-effort defense; this API does not promise hostile shared-directory
    /// safety.
    ///
    /// The established replacement helper replaces the destination name with
    /// a sibling temporary file.  Existing destination contents and metadata
    /// (including permission bits) are therefore not preserved; callers that
    /// require a specific mode should apply it after a successful save.
    pub fn save<P: AsRef<Path>>(
        self,
        path: P,
    ) -> Result<SequentialWriteReport, SequentialWriteError> {
        self.save_with_hooks(path, atomic_replace, sync_parent)
    }

    fn save_with_hooks<P, A, S>(
        self,
        path: P,
        replace: A,
        sync: S,
    ) -> Result<SequentialWriteReport, SequentialWriteError>
    where
        P: AsRef<Path>,
        A: FnOnce(&Path, &Path) -> io::Result<()>,
        S: FnOnce(&Path) -> io::Result<()>,
    {
        let destination = path.as_ref();
        let parent = parent_directory(destination);
        let (temporary_path, file) =
            create_sibling_temp_file(destination).map_err(SequentialWriteError::Planning)?;
        let mut cleanup = TemporaryCleanupGuard::new(&temporary_path);
        let temporary_identity = match file_identity(&file) {
            Ok(identity) => {
                cleanup.set_identity(identity);
                identity
            },
            Err(source) => {
                // Drop the open handle before the guard attempts exact-name
                // cleanup (notably required by Windows file-sharing rules).
                drop(file);
                return Err(SequentialWriteError::Planning(OleError::Io(source)));
            },
        };
        (|| {
            let mut buffered = BufWriter::new(file);
            let report = self.write_to(&mut buffered)?;
            buffered
                .flush()
                .map_err(|source| SequentialWriteError::Flush {
                    source,
                    progress: SequentialWriteProgress::CompleteUnflushed {
                        bytes: report.output_bytes,
                    },
                })?;
            let staged_file =
                buffered
                    .into_inner()
                    .map_err(|error| SequentialWriteError::Flush {
                        source: error.into_error(),
                        progress: SequentialWriteProgress::CompleteUnflushed {
                            bytes: report.output_bytes,
                        },
                    })?;
            staged_file
                .sync_all()
                .map_err(|source| SequentialWriteError::Stage {
                    source,
                    progress: SequentialWriteProgress::Complete {
                        bytes: report.output_bytes,
                    },
                    report,
                })?;

            // Validate through a clone of the file we created, never by
            // reopening the mutable temporary pathname.  A sibling process
            // may unlink and recreate that name while we are staging.
            let candidate =
                staged_file
                    .try_clone()
                    .map_err(|source| SequentialWriteError::Stage {
                        source,
                        progress: SequentialWriteProgress::Complete {
                            bytes: report.output_bytes,
                        },
                        report,
                    })?;
            let validated = OleFile::open(candidate).map_err(|error| match error {
                OleError::Io(source) => SequentialWriteError::Stage {
                    source,
                    progress: SequentialWriteProgress::Complete {
                        bytes: report.output_bytes,
                    },
                    report,
                },
                source => SequentialWriteError::CandidateValidation {
                    source,
                    progress: SequentialWriteProgress::Complete {
                        bytes: report.output_bytes,
                    },
                    report,
                },
            })?;
            // Windows requires all validation handles to be closed before the
            // sibling can replace the destination name.
            drop(validated);
            drop(staged_file);

            // Best-effort identity checking rejects a known replacement and
            // lets cleanup preserve it.  The public save contract requires a
            // trusted/private parent because portable path APIs cannot close
            // every hostile-directory race.
            ensure_temp_identity(&temporary_path, temporary_identity).map_err(|source| {
                SequentialWriteError::Stage {
                    source,
                    progress: SequentialWriteProgress::Complete {
                        bytes: report.output_bytes,
                    },
                    report,
                }
            })?;

            replace(&temporary_path, destination).map_err(|source| {
                SequentialWriteError::Stage {
                    source,
                    progress: SequentialWriteProgress::Complete {
                        bytes: report.output_bytes,
                    },
                    report,
                }
            })?;
            cleanup.mark_published();
            sync(parent).map_err(|source| SequentialWriteError::Committed {
                source,
                report,
                progress: SequentialWriteProgress::Complete {
                    bytes: report.output_bytes,
                },
            })?;
            Ok(report)
        })()
    }

    fn own_path(
        &mut self,
        path: &[&str],
        resource: &'static str,
    ) -> Result<Vec<String>, SequentialWriteError> {
        let limits = self.options.limits;
        if path.is_empty() {
            return Err(planning(format!("CFB {resource} must not be empty")));
        }
        let components = u64::try_from(path.len()).unwrap_or(u64::MAX);
        if components > limits.max_path_components {
            return Err(limit(
                "path components",
                components,
                limits.max_path_components,
            ));
        }
        let mut path_bytes = 0_u64;
        for component in path {
            directory_name_data(component).map_err(|error| planning(error.to_string()))?;
            let bytes = u64::try_from(component.len()).unwrap_or(u64::MAX);
            path_bytes = path_bytes
                .checked_add(bytes)
                .ok_or_else(|| limit("path bytes", u64::MAX, limits.max_path_bytes))?;
        }
        let next_total = self
            .path_bytes
            .checked_add(path_bytes)
            .ok_or_else(|| limit("path bytes", u64::MAX, limits.max_path_bytes))?;
        if next_total > limits.max_path_bytes {
            return Err(limit("path bytes", next_total, limits.max_path_bytes));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(path.len())
            .map_err(|source| OleError::allocation("sequential path", source))?;
        for component in path {
            let mut owned_component = String::new();
            owned_component
                .try_reserve(component.len())
                .map_err(|source| OleError::allocation("sequential path component", source))?;
            owned_component.push_str(component);
            owned.push(owned_component);
        }
        self.path_bytes = next_total;
        Ok(owned)
    }

    fn plan(mut self) -> Result<SequentialPlan<'a>, SequentialWriteError> {
        validate_options(&self.options)?;
        let sector_size = self.options.sector_size;
        let limits = self.options.limits;
        let stream_count = u64::try_from(self.streams.len()).unwrap_or(u64::MAX);
        if stream_count > limits.max_streams {
            return Err(limit("streams", stream_count, limits.max_streams));
        }

        // Compute every serialized-sector bound before growing any writer
        // builder or allocating a directory/FAT/MiniFAT byte buffer.  This is
        // the hard `max_metadata_bytes` contract, not a post-plan report.
        let preflight = preflight_layout(&self.streams, &self.storages, sector_size, limits)?;

        let mut fat =
            FatBuilder::new_with_size(sector_size).map_err(SequentialWriteError::Planning)?;
        let mut minifat = MiniFatBuilder::new(MINI_SECTOR_SIZE);
        let mut large_streams = Vec::new();
        let mut small_streams = Vec::new();
        large_streams
            .try_reserve(self.streams.len())
            .map_err(|source| OleError::allocation("sequential large-stream plan", source))?;
        small_streams
            .try_reserve(self.streams.len())
            .map_err(|source| OleError::allocation("sequential small-stream plan", source))?;

        for (index, stream) in self.streams.iter_mut().enumerate() {
            if stream.declared_len < MINI_STREAM_CUTOFF {
                let length = usize::try_from(stream.declared_len)
                    .map_err(|_err| planning("CFB mini stream length does not fit usize"))?;
                stream.start_mini_sector = minifat
                    .allocate_mini_chain_len(length)
                    .map_err(SequentialWriteError::Planning)?;
                small_streams.push(index);
            } else {
                stream.start_sector = fat
                    .allocate_chain_u64(stream.declared_len)
                    .map_err(SequentialWriteError::Planning)?;
                large_streams.push(index);
            }
        }
        if u64::from(fat.total_sectors()) != preflight.large_sector_count {
            return Err(planning(
                "CFB large-stream preflight disagrees with allocation",
            ));
        }

        let ministream_size = minifat
            .ministream_size()
            .map_err(SequentialWriteError::Planning)?;
        if u64::from(minifat.mini_sector_count()) != preflight.mini_sector_count {
            return Err(planning("CFB MiniFAT preflight disagrees with allocation"));
        }
        if ministream_size != preflight.ministream_bytes {
            return Err(planning(
                "CFB ministream preflight disagrees with allocation",
            ));
        }
        let ministream_start = if minifat.is_empty() {
            ENDOFCHAIN
        } else {
            fat.allocate_chain_u64(ministream_size)
                .map_err(SequentialWriteError::Planning)?
        };
        let sectors_after_ministream = u64::from(fat.total_sectors());
        let expected_after_ministream = preflight
            .large_sector_count
            .checked_add(preflight.ministream_sector_count)
            .ok_or_else(|| planning("CFB ministream sector count overflows u64"))?;
        if sectors_after_ministream != expected_after_ministream {
            return Err(planning(
                "CFB ministream sector preflight disagrees with allocation",
            ));
        }

        let mut directory = DirectoryBuilder::try_new(ministream_start, ministream_size)
            .map_err(SequentialWriteError::Planning)?;
        if let Some(clsid) = self.root_clsid {
            directory.set_root_clsid(clsid);
        }
        for storage in &self.storages {
            directory
                .add_storage_path(&storage.path)
                .map_err(SequentialWriteError::Planning)?;
        }
        for storage in &self.storages {
            if let Some(clsid) = storage.clsid {
                directory
                    .set_storage_clsid(&storage.path, clsid)
                    .map_err(SequentialWriteError::Planning)?;
            }
        }
        for &index in &large_streams {
            let stream = &self.streams[index];
            directory
                .add_stream_path(&stream.path, stream.start_sector, stream.declared_len)
                .map_err(SequentialWriteError::Planning)?;
        }
        for &index in &small_streams {
            let stream = &self.streams[index];
            directory
                .add_stream_path(&stream.path, stream.start_mini_sector, stream.declared_len)
                .map_err(SequentialWriteError::Planning)?;
        }
        let directory_entries = u64::try_from(directory.entry_count()).unwrap_or(u64::MAX);
        if directory_entries != preflight.directory_entries {
            return Err(planning(
                "CFB directory preflight disagrees with builder entry count",
            ));
        }
        let directory_bytes = directory
            .generate_directory_stream()
            .map_err(SequentialWriteError::Planning)?;
        if u64::try_from(directory_bytes.len()).unwrap_or(u64::MAX) != preflight.directory_bytes {
            return Err(planning(
                "CFB directory preflight disagrees with serialized length",
            ));
        }
        let directory_start = fat
            .allocate_chain_u64(u64::try_from(directory_bytes.len()).unwrap_or(u64::MAX))
            .map_err(SequentialWriteError::Planning)?;
        if u64::from(fat.total_sectors())
            != expected_after_ministream
                .checked_add(preflight.directory_sector_count)
                .ok_or_else(|| planning("CFB directory sector count overflows u64"))?
        {
            return Err(planning(
                "CFB directory sector preflight disagrees with allocation",
            ));
        }

        let minifat_sectors = minifat
            .generate_minifat_sectors(sector_size)
            .map_err(SequentialWriteError::Planning)?;
        if u64::try_from(minifat_sectors.len()).unwrap_or(u64::MAX)
            != preflight.minifat_sector_count
        {
            return Err(planning(
                "CFB MiniFAT preflight disagrees with serialized length",
            ));
        }
        let minifat_start = if minifat_sectors.is_empty() {
            ENDOFCHAIN
        } else {
            let bytes = checked_byte_product(minifat_sectors.len(), sector_size, "MiniFAT")?;
            fat.allocate_chain_u64(bytes)
                .map_err(SequentialWriteError::Planning)?
        };
        if u64::from(fat.total_sectors())
            != expected_after_ministream
                .checked_add(preflight.directory_sector_count)
                .and_then(|value| value.checked_add(preflight.minifat_sector_count))
                .ok_or_else(|| planning("CFB MiniFAT sector count overflows u64"))?
        {
            return Err(planning(
                "CFB MiniFAT sector preflight disagrees with allocation",
            ));
        }

        let used = fat.total_sectors();
        if u64::from(used) != preflight.used_sectors {
            return Err(planning("CFB sector preflight disagrees with allocation"));
        }
        let (fat_count, difat_count) = allocation_table_sector_counts(used, sector_size)?;
        if u64::from(fat_count) != preflight.fat_sector_count
            || u64::from(difat_count) != preflight.difat_sector_count
        {
            return Err(planning(
                "CFB FAT/DIFAT preflight disagrees with allocation",
            ));
        }
        let difat_start = if difat_count == 0 {
            ENDOFCHAIN
        } else {
            fat.allocate_special(difat_count, DIFSECT)
                .map_err(SequentialWriteError::Planning)?
        };
        let fat_start = if fat_count == 0 {
            ENDOFCHAIN
        } else {
            fat.allocate_special(fat_count, FATSECT)
                .map_err(SequentialWriteError::Planning)?
        };
        validate_output_size(sector_size, fat.total_sectors(), limits.max_output_bytes)?;
        fat.validate().map_err(SequentialWriteError::Planning)?;

        let fat_sector_ids = sector_ids(fat_start, fat_count)?;
        let mut difat = DifatBuilder::new(sector_size).map_err(SequentialWriteError::Planning)?;
        difat
            .set_fat_sectors(&fat_sector_ids)
            .map_err(SequentialWriteError::Planning)?;
        let generated_difat_count = difat
            .calculate_difat_sector_count()
            .map_err(SequentialWriteError::Planning)?;
        if generated_difat_count != difat_count {
            return Err(planning("CFB DIFAT planning mismatch"));
        }
        let difat_sectors = if difat_count == 0 {
            Vec::new()
        } else {
            difat
                .generate_difat_sectors(difat_start)
                .map_err(SequentialWriteError::Planning)?
        };
        let fat_sectors = fat
            .generate_fat_sectors()
            .map_err(SequentialWriteError::Planning)?;
        if u32::try_from(fat_sectors.len()).unwrap_or(u32::MAX) != fat_count {
            return Err(planning("CFB FAT planning mismatch"));
        }

        let mut header_builder =
            HeaderBuilder::new(sector_size).map_err(SequentialWriteError::Planning)?;
        header_builder.set_first_dir_sector(directory_start);
        let directory_sector_count = sectors_for_len(directory_bytes.len(), sector_size)?;
        header_builder.set_num_dir_sectors(directory_sector_count);
        let minifat_sector_count = u32::try_from(minifat_sectors.len())
            .map_err(|_err| planning("CFB MiniFAT sector count exceeds u32"))?;
        header_builder.set_minifat(minifat_start, minifat_sector_count);
        header_builder
            .set_fat_sectors(&fat_sector_ids)
            .map_err(SequentialWriteError::Planning)?;
        if difat_count != 0 {
            header_builder.set_difat(difat_start, difat_count);
        }
        let header = header_builder
            .generate()
            .map_err(SequentialWriteError::Planning)?;

        let output_bytes = (u64::from(fat.total_sectors()) + 1)
            .checked_mul(u64::try_from(sector_size).unwrap_or(u64::MAX))
            .ok_or_else(|| planning("CFB output size overflows u64"))?;
        if output_bytes != preflight.output_bytes {
            return Err(planning("CFB output preflight disagrees with allocation"));
        }
        let mut physical_bytes = u64::try_from(header.len()).unwrap_or(u64::MAX);
        for stream in &self.streams {
            let alignment = if stream.declared_len < MINI_STREAM_CUTOFF {
                MINI_SECTOR_SIZE
            } else {
                sector_size
            };
            let bytes = padded_len(stream.declared_len, alignment)
                .ok_or_else(|| planning("CFB stream physical length overflows u64"))?;
            physical_bytes = physical_bytes
                .checked_add(bytes)
                .ok_or_else(|| planning("CFB physical length overflows u64"))?;
        }
        let ministream_bytes = self
            .streams
            .iter()
            .filter(|stream| stream.declared_len < MINI_STREAM_CUTOFF)
            .try_fold(0_u64, |total, stream| {
                let bytes = padded_len(stream.declared_len, MINI_SECTOR_SIZE)
                    .ok_or_else(|| planning("CFB ministream length overflows u64"))?;
                total
                    .checked_add(bytes)
                    .ok_or_else(|| planning("CFB ministream length overflows u64"))
            })?;
        let ministream_regular_bytes = padded_len(ministream_bytes, sector_size)
            .ok_or_else(|| planning("CFB ministream regular length overflows u64"))?;
        let ministream_regular_padding = ministream_regular_bytes
            .checked_sub(ministream_bytes)
            .ok_or_else(|| planning("CFB ministream padding underflows u64"))?;
        physical_bytes = physical_bytes
            .checked_add(ministream_regular_padding)
            .and_then(|bytes| {
                bytes.checked_add(padded_len(
                    u64::try_from(directory_bytes.len()).unwrap_or(u64::MAX),
                    sector_size,
                )?)
            })
            .ok_or_else(|| planning("CFB directory physical length overflows u64"))?;
        for sector in minifat_sectors
            .iter()
            .chain(difat_sectors.iter())
            .chain(fat_sectors.iter())
        {
            physical_bytes = physical_bytes
                .checked_add(u64::try_from(sector.len()).unwrap_or(u64::MAX))
                .ok_or_else(|| planning("CFB metadata physical length overflows u64"))?;
        }
        if physical_bytes != output_bytes {
            return Err(planning("CFB sequential physical layout length mismatch"));
        }
        let payload_bytes = preflight.payload_bytes;
        let metadata_bytes = preflight.metadata_bytes;
        let publication_buffer_bytes =
            plan_buffer_size(self.options.publication_buffer_bytes, sector_size)?;
        let report = SequentialWriteReport {
            output_bytes,
            payload_bytes,
            metadata_bytes,
            stream_count,
            publication_buffer_bytes: u64::try_from(publication_buffer_bytes).unwrap_or(u64::MAX),
        };
        let mut publication_buffer = Vec::new();
        publication_buffer
            .try_reserve_exact(publication_buffer_bytes)
            .map_err(|source| OleError::allocation("sequential publication buffer", source))?;
        publication_buffer.resize(publication_buffer_bytes, 0);

        let cancellation = self.options.cancellation.clone();
        Ok(SequentialPlan {
            sector_size,
            cancellation,
            streams: self.streams,
            large_streams,
            small_streams,
            header,
            directory: directory_bytes,
            minifat_sectors,
            difat_sectors,
            fat_sectors,
            output_bytes,
            report,
            publication_buffer,
        })
    }
}

impl<'a> Default for SequentialOleWriter<'a> {
    fn default() -> Self {
        Self::new()
    }
}

struct LayoutPreflight {
    directory_entries: u64,
    directory_bytes: u64,
    large_sector_count: u64,
    mini_sector_count: u64,
    ministream_bytes: u64,
    ministream_sector_count: u64,
    directory_sector_count: u64,
    minifat_sector_count: u64,
    used_sectors: u64,
    fat_sector_count: u64,
    difat_sector_count: u64,
    output_bytes: u64,
    payload_bytes: u64,
    metadata_bytes: u64,
}

fn preflight_layout(
    streams: &[StreamInput<'_>],
    storages: &[StorageInput],
    sector_size: usize,
    limits: SequentialWriterLimits,
) -> Result<LayoutPreflight, SequentialWriteError> {
    let directory_entries = preflight_directory_entries(streams, storages, limits)?;
    let directory_bytes = directory_entries
        .checked_mul(128)
        .ok_or_else(|| planning("CFB directory byte count overflows u64"))?;
    let directory_sector_count = sectors_for_u64(directory_bytes, sector_size)?;

    let mut large_sector_count = 0_u64;
    let mut mini_sector_count = 0_u64;
    let mut payload_bytes = 0_u64;
    for stream in streams {
        payload_bytes = payload_bytes
            .checked_add(stream.declared_len)
            .ok_or_else(|| planning("CFB payload byte count overflows u64"))?;
        if stream.declared_len < MINI_STREAM_CUTOFF {
            mini_sector_count = mini_sector_count
                .checked_add(ceil_div_u64(stream.declared_len, MINI_SECTOR_SIZE)?)
                .ok_or_else(|| planning("CFB MiniFAT sector count overflows u64"))?;
        } else {
            large_sector_count = large_sector_count
                .checked_add(ceil_div_u64(stream.declared_len, sector_size)?)
                .ok_or_else(|| planning("CFB FAT sector count overflows u64"))?;
        }
    }
    if mini_sector_count >= u64::from(MAXREGSECT) {
        return Err(planning("CFB MiniFAT sector count reaches MAXREGSECT"));
    }
    let ministream_bytes = mini_sector_count
        .checked_mul(u64::try_from(MINI_SECTOR_SIZE).unwrap_or(u64::MAX))
        .ok_or_else(|| planning("CFB ministream byte count overflows u64"))?;
    let ministream_sector_count = ceil_div_u64(ministream_bytes, sector_size)?;
    let minifat_bytes = mini_sector_count
        .checked_mul(4)
        .ok_or_else(|| planning("CFB MiniFAT byte count overflows u64"))?;
    let minifat_sector_count = ceil_div_u64(minifat_bytes, sector_size)?;
    let used_sectors = large_sector_count
        .checked_add(ministream_sector_count)
        .and_then(|value| value.checked_add(directory_sector_count))
        .and_then(|value| value.checked_add(minifat_sector_count))
        .ok_or_else(|| planning("CFB used-sector count overflows u64"))?;
    let used_u32 = u32::try_from(used_sectors)
        .map_err(|_err| planning("CFB used-sector count exceeds u32"))?;
    let (fat_sector_count, difat_sector_count) =
        allocation_table_sector_counts(used_u32, sector_size)?;
    let fat_sector_count = u64::from(fat_sector_count);
    let difat_sector_count = u64::from(difat_sector_count);
    let total_sectors = used_sectors
        .checked_add(fat_sector_count)
        .and_then(|value| value.checked_add(difat_sector_count))
        .ok_or_else(|| planning("CFB total-sector count overflows u64"))?;
    let total_u32 = u32::try_from(total_sectors)
        .map_err(|_err| planning("CFB total-sector count exceeds u32"))?;
    validate_output_size(sector_size, total_u32, limits.max_output_bytes)?;
    let output_bytes = (total_sectors + 1)
        .checked_mul(u64::try_from(sector_size).unwrap_or(u64::MAX))
        .ok_or_else(|| planning("CFB output size overflows u64"))?;
    let metadata_bytes = output_bytes
        .checked_sub(payload_bytes)
        .ok_or_else(|| planning("CFB payload exceeds planned output"))?;
    if metadata_bytes > limits.max_metadata_bytes {
        return Err(limit(
            "metadata bytes",
            metadata_bytes,
            limits.max_metadata_bytes,
        ));
    }
    Ok(LayoutPreflight {
        directory_entries,
        directory_bytes,
        large_sector_count,
        mini_sector_count,
        ministream_bytes,
        ministream_sector_count,
        directory_sector_count,
        minifat_sector_count,
        used_sectors,
        fat_sector_count,
        difat_sector_count,
        output_bytes,
        payload_bytes,
        metadata_bytes,
    })
}

fn preflight_directory_entries(
    streams: &[StreamInput<'_>],
    storages: &[StorageInput],
    limits: SequentialWriterLimits,
) -> Result<u64, SequentialWriteError> {
    let capacity = storages
        .len()
        .checked_add(streams.len())
        .ok_or_else(|| planning("CFB directory path count overflows usize"))?;
    let mut unique_storages: Vec<&[String]> = Vec::new();
    unique_storages
        .try_reserve(capacity)
        .map_err(|source| OleError::allocation("CFB directory preflight paths", source))?;
    for storage in storages {
        add_unique_storage_prefixes(&mut unique_storages, &storage.path)?;
    }
    for stream in streams {
        if stream.path.len() > 1 {
            add_unique_storage_prefixes(
                &mut unique_storages,
                &stream.path[..stream.path.len() - 1],
            )?;
        }
    }
    let stream_count = u64::try_from(streams.len()).unwrap_or(u64::MAX);
    let storage_count = u64::try_from(unique_storages.len()).unwrap_or(u64::MAX);
    let entries = 1_u64
        .checked_add(stream_count)
        .and_then(|value| value.checked_add(storage_count))
        .ok_or_else(|| planning("CFB directory entry count overflows u64"))?;
    if entries > limits.max_directory_entries {
        return Err(limit(
            "directory entries",
            entries,
            limits.max_directory_entries,
        ));
    }
    Ok(entries)
}

fn add_unique_storage_prefixes<'a>(
    unique_storages: &mut Vec<&'a [String]>,
    path: &'a [String],
) -> Result<(), SequentialWriteError> {
    for depth in 1..=path.len() {
        let prefix = &path[..depth];
        if !unique_storages.contains(&prefix) {
            unique_storages
                .try_reserve(1)
                .map_err(|source| OleError::allocation("CFB directory preflight paths", source))?;
            unique_storages.push(prefix);
        }
    }
    Ok(())
}

fn ceil_div_u64(value: u64, divisor: usize) -> Result<u64, SequentialWriteError> {
    let divisor =
        u64::try_from(divisor).map_err(|_err| planning("CFB sector divisor does not fit u64"))?;
    if divisor == 0 {
        return Err(planning("CFB sector divisor must be nonzero"));
    }
    value
        .checked_add(divisor - 1)
        .map(|value| value / divisor)
        .ok_or_else(|| planning("CFB sector count overflows u64"))
}

fn sectors_for_u64(value: u64, sector_size: usize) -> Result<u64, SequentialWriteError> {
    ceil_div_u64(value, sector_size)
}

struct SequentialPlan<'a> {
    sector_size: usize,
    cancellation: Option<CancellationToken>,
    streams: Vec<StreamInput<'a>>,
    large_streams: Vec<usize>,
    small_streams: Vec<usize>,
    header: Vec<u8>,
    directory: Vec<u8>,
    minifat_sectors: Vec<Vec<u8>>,
    difat_sectors: Vec<Vec<u8>>,
    fat_sectors: Vec<Vec<u8>>,
    output_bytes: u64,
    report: SequentialWriteReport,
    publication_buffer: Vec<u8>,
}

fn emit_stream<W: Write>(
    sink: &mut W,
    stream: &mut StreamInput<'_>,
    alignment: usize,
    buffer: &mut [u8],
    accepted: &mut u64,
    expected_output: u64,
    cancellation: &Option<CancellationToken>,
) -> Result<(), SequentialWriteError> {
    let expected = stream.declared_len;
    let mut remaining = expected;
    let mut observed = 0_u64;
    while remaining != 0 {
        check_cancel(cancellation, *accepted, expected_output)?;
        let requested = usize::try_from(remaining)
            .unwrap_or(buffer.len())
            .min(buffer.len());
        let read = loop {
            match stream.source.read(&mut buffer[..requested]) {
                Ok(read) => break read,
                Err(error) if error.kind() == ErrorKind::Interrupted => {
                    check_cancel(cancellation, *accepted, expected_output)?;
                    continue;
                },
                Err(source) => {
                    return Err(SequentialWriteError::SourceIo {
                        path: std::mem::take(&mut stream.path),
                        source,
                        progress: prefix_progress(*accepted, expected_output),
                    });
                },
            }
        };
        if read == 0 {
            return Err(SequentialWriteError::SourceLength {
                path: std::mem::take(&mut stream.path),
                expected,
                observed,
                progress: prefix_progress(*accepted, expected_output),
            });
        }
        if read > requested {
            return Err(SequentialWriteError::SourceLength {
                path: std::mem::take(&mut stream.path),
                expected,
                observed: expected.saturating_add(1),
                progress: prefix_progress(*accepted, expected_output),
            });
        }
        let read_u64 = u64::try_from(read).unwrap_or(u64::MAX);
        observed = observed.saturating_add(read_u64);
        remaining = remaining.saturating_sub(read_u64);
        publish_segment(
            sink,
            &buffer[..read],
            accepted,
            expected_output,
            cancellation,
        )?;
    }

    // Probe exactly one byte.  This catches a producer that declared a short
    // length without retaining or buffering the unplanned tail.
    check_cancel(cancellation, *accepted, expected_output)?;
    let mut probe = [0_u8; 1];
    let extra = loop {
        match stream.source.read(&mut probe) {
            Ok(read) => break read,
            Err(error) if error.kind() == ErrorKind::Interrupted => {
                check_cancel(cancellation, *accepted, expected_output)?;
                continue;
            },
            Err(source) => {
                return Err(SequentialWriteError::SourceIo {
                    path: std::mem::take(&mut stream.path),
                    source,
                    progress: prefix_progress(*accepted, expected_output),
                });
            },
        }
    };
    if extra != 0 {
        return Err(SequentialWriteError::SourceLength {
            path: std::mem::take(&mut stream.path),
            expected,
            observed: expected.saturating_add(1),
            progress: prefix_progress(*accepted, expected_output),
        });
    }

    let padding = padded_len(expected, alignment)
        .and_then(|padded| padded.checked_sub(expected))
        .ok_or_else(|| {
            SequentialWriteError::Planning(OleError::InvalidData(
                "CFB stream padding overflows u64".to_string(),
            ))
        })?;
    publish_zeroes(
        sink,
        padding,
        buffer,
        accepted,
        expected_output,
        cancellation,
    )
}

fn publish_padded<W: Write>(
    sink: &mut W,
    bytes: &[u8],
    alignment: usize,
    buffer: &mut [u8],
    accepted: &mut u64,
    expected_output: u64,
    cancellation: &Option<CancellationToken>,
) -> Result<(), SequentialWriteError> {
    publish_segment(sink, bytes, accepted, expected_output, cancellation)?;
    let bytes_len = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
    let padding = padded_len(bytes_len, alignment)
        .and_then(|padded| padded.checked_sub(bytes_len))
        .ok_or_else(|| {
            SequentialWriteError::Planning(OleError::InvalidData(
                "CFB serialized stream padding overflows u64".to_string(),
            ))
        })?;
    publish_zeroes(
        sink,
        padding,
        buffer,
        accepted,
        expected_output,
        cancellation,
    )
}

fn publish_zeroes<W: Write>(
    sink: &mut W,
    mut bytes: u64,
    buffer: &mut [u8],
    accepted: &mut u64,
    expected_output: u64,
    cancellation: &Option<CancellationToken>,
) -> Result<(), SequentialWriteError> {
    buffer.fill(0);
    while bytes != 0 {
        check_cancel(cancellation, *accepted, expected_output)?;
        let count = usize::try_from(bytes)
            .unwrap_or(buffer.len())
            .min(buffer.len());
        publish_segment(
            sink,
            &buffer[..count],
            accepted,
            expected_output,
            cancellation,
        )?;
        bytes = bytes.saturating_sub(u64::try_from(count).unwrap_or(u64::MAX));
    }
    Ok(())
}

fn publish_segment<W: Write>(
    sink: &mut W,
    bytes: &[u8],
    accepted: &mut u64,
    expected_output: u64,
    cancellation: &Option<CancellationToken>,
) -> Result<(), SequentialWriteError> {
    let mut offset = 0_usize;
    while offset < bytes.len() {
        check_cancel(cancellation, *accepted, expected_output)?;
        let written = match sink.write(&bytes[offset..]) {
            Ok(written) => written,
            Err(error) if error.kind() == ErrorKind::Interrupted => {
                check_cancel(cancellation, *accepted, expected_output)?;
                continue;
            },
            Err(source) => {
                return Err(SequentialWriteError::Sink {
                    source,
                    progress: prefix_progress(*accepted, expected_output),
                });
            },
        };
        if written == 0 {
            return Err(SequentialWriteError::WriteZero {
                progress: prefix_progress(*accepted, expected_output),
            });
        }
        let remaining = bytes.len() - offset;
        if written > remaining {
            return Err(SequentialWriteError::Indeterminate {
                accepted_before: *accepted,
            });
        }
        offset += written;
        *accepted = accepted.saturating_add(u64::try_from(written).unwrap_or(u64::MAX));
    }
    Ok(())
}

fn check_cancel(
    cancellation: &Option<CancellationToken>,
    accepted: u64,
    expected: u64,
) -> Result<(), SequentialWriteError> {
    if cancellation
        .as_ref()
        .is_some_and(CancellationToken::is_cancelled)
    {
        return Err(SequentialWriteError::Cancelled {
            progress: prefix_progress(accepted, expected),
        });
    }
    Ok(())
}

fn prefix_progress(accepted: u64, expected: u64) -> SequentialWriteProgress {
    if accepted == 0 {
        SequentialWriteProgress::Untouched
    } else {
        SequentialWriteProgress::Prefix { accepted, expected }
    }
}

fn validate_options(options: &SequentialWriterOptions) -> Result<(), SequentialWriteError> {
    if !matches!(options.sector_size, 512 | 4096) {
        return Err(planning(format!(
            "CFB sector size must be 512 or 4096 bytes, got {}",
            options.sector_size
        )));
    }
    if options.publication_buffer_bytes == 0 || options.publication_buffer_bytes > 16 * 1024 * 1024
    {
        return Err(limit(
            "publication buffer bytes",
            u64::try_from(options.publication_buffer_bytes).unwrap_or(u64::MAX),
            16 * 1024 * 1024,
        ));
    }
    let limits = options.limits;
    for (resource, value) in [
        ("streams", limits.max_streams),
        ("directory entries", limits.max_directory_entries),
        ("path components", limits.max_path_components),
        ("path bytes", limits.max_path_bytes),
        ("stream bytes", limits.max_stream_bytes),
        ("metadata bytes", limits.max_metadata_bytes),
        ("output bytes", limits.max_output_bytes),
    ] {
        if value == 0 {
            return Err(limit(resource, value, 1));
        }
    }
    Ok(())
}

fn plan_buffer_size(requested: usize, sector_size: usize) -> Result<usize, SequentialWriteError> {
    if requested < sector_size {
        return Err(limit(
            "publication buffer bytes",
            u64::try_from(requested).unwrap_or(u64::MAX),
            u64::try_from(sector_size).unwrap_or(u64::MAX),
        ));
    }
    Ok(requested)
}

fn planning(message: impl Into<String>) -> SequentialWriteError {
    SequentialWriteError::Planning(OleError::InvalidData(message.into()))
}

fn limit(resource: &'static str, observed: u64, limit: u64) -> SequentialWriteError {
    SequentialWriteError::LimitExceeded {
        resource,
        observed,
        limit,
    }
}

fn display_path(path: &[String]) -> String {
    path.join("/")
}

fn padded_len(len: u64, alignment: usize) -> Option<u64> {
    let alignment = u64::try_from(alignment).ok()?;
    if alignment == 0 {
        return None;
    }
    let remainder = len % alignment;
    if remainder == 0 {
        Some(len)
    } else {
        len.checked_add(alignment - remainder)
    }
}

fn checked_byte_product(
    count: usize,
    size: usize,
    resource: &'static str,
) -> Result<u64, SequentialWriteError> {
    let count = u64::try_from(count)
        .map_err(|_err| planning(format!("CFB {resource} sector count does not fit u64")))?;
    let size = u64::try_from(size)
        .map_err(|_err| planning(format!("CFB {resource} sector size does not fit u64")))?;
    count
        .checked_mul(size)
        .ok_or_else(|| planning(format!("CFB {resource} byte count overflows u64")))
}

fn sectors_for_len(len: usize, sector_size: usize) -> Result<u32, SequentialWriteError> {
    let count = len.div_ceil(sector_size);
    let count = u32::try_from(count)
        .map_err(|_err| planning("CFB serialized stream has too many sectors"))?;
    if count > MAXREGSECT {
        return Err(planning("CFB serialized stream exceeds MAXREGSECT"));
    }
    Ok(count)
}

fn validate_output_size(
    sector_size: usize,
    sector_count: u32,
    maximum: u64,
) -> Result<(), SequentialWriteError> {
    // MAXREGSECT is the first reserved sector marker, not a usable sector ID.
    // Require the complete regular-sector count to remain strictly below it.
    if sector_count >= MAXREGSECT {
        return Err(planning(
            "CFB output requires a sector count below MAXREGSECT",
        ));
    }
    let bytes = (u64::from(sector_count) + 1)
        .checked_mul(u64::try_from(sector_size).unwrap_or(u64::MAX))
        .ok_or_else(|| planning("CFB output size overflows u64"))?;
    if bytes > maximum {
        return Err(limit("output bytes", bytes, maximum));
    }
    if sector_size == 512 && bytes > DEFAULT_MAX_OUTPUT_BYTES {
        return Err(planning("version 3 CFB output cannot exceed 2 GiB"));
    }
    Ok(())
}

fn allocation_table_sector_counts(
    used: u32,
    sector_size: usize,
) -> Result<(u32, u32), SequentialWriteError> {
    let fat_entries =
        u32::try_from(sector_size / 4).map_err(|_err| planning("CFB FAT geometry exceeds u32"))?;
    let difat_entries = fat_entries
        .checked_sub(1)
        .ok_or_else(|| planning("CFB DIFAT sector has no ID slots"))?;
    let mut fat = 0_u32;
    let mut difat = 0_u32;
    for _ in 0..32 {
        let total = used
            .checked_add(fat)
            .and_then(|value| value.checked_add(difat))
            .ok_or_else(|| planning("CFB sector count overflows u32"))?;
        let next_fat = total.div_ceil(fat_entries);
        let next_difat = next_fat.saturating_sub(109).div_ceil(difat_entries);
        if next_fat == fat && next_difat == difat {
            return Ok((fat, difat));
        }
        fat = next_fat;
        difat = next_difat;
    }
    Err(planning("CFB FAT/DIFAT planning did not converge"))
}

fn sector_ids(start: u32, count: u32) -> Result<Vec<u32>, SequentialWriteError> {
    if count == 0 {
        return Ok(Vec::new());
    }
    let end = start
        .checked_add(count)
        .ok_or_else(|| planning("CFB FAT sector range overflows u32"))?;
    if start >= MAXREGSECT || end > MAXREGSECT {
        return Err(planning("CFB FAT sector range exceeds MAXREGSECT"));
    }
    let count =
        usize::try_from(count).map_err(|_err| planning("CFB FAT sector count exceeds usize"))?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(count)
        .map_err(|source| OleError::allocation("sequential FAT sector IDs", source))?;
    ids.extend(start..end);
    Ok(ids)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused sequential-writer tests use panic-on-failure assertions"
    )]

    use super::*;
    use std::path::PathBuf;
    use std::sync::atomic::{AtomicU64, Ordering};

    fn test_destination(label: &str) -> PathBuf {
        static NEXT_ID: AtomicU64 = AtomicU64::new(0);
        let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
        std::env::temp_dir().join(format!(
            "litchi-cfb-sequential-{label}-{}-{id}.ole",
            std::process::id()
        ))
    }

    fn writer_with_payload(
        name: &'static str,
        payload: &'static [u8],
    ) -> SequentialOleWriter<'static> {
        let mut writer = SequentialOleWriter::new();
        writer
            .add_stream(&[name], payload.len() as u64, io::Cursor::new(payload))
            .unwrap();
        writer
    }

    #[test]
    fn exact_two_gib_and_reserved_sector_boundaries_are_explicit() {
        let exact_two_gib_sector_count = u32::try_from(DEFAULT_MAX_OUTPUT_BYTES / 512 - 1).unwrap();
        assert!(
            validate_output_size(512, exact_two_gib_sector_count, DEFAULT_MAX_OUTPUT_BYTES).is_ok()
        );

        let one_over =
            validate_output_size(512, exact_two_gib_sector_count + 1, u64::MAX).unwrap_err();
        assert!(
            one_over
                .to_string()
                .contains("version 3 CFB output cannot exceed 2 GiB")
        );

        assert!(validate_output_size(4096, MAXREGSECT - 1, u64::MAX).is_ok());
        let reserved = validate_output_size(4096, MAXREGSECT, u64::MAX).unwrap_err();
        assert!(
            reserved
                .to_string()
                .contains("sector count below MAXREGSECT")
        );
    }

    #[test]
    fn four_k_geometry_covers_fat_and_difat_transition() {
        let fat_entries = 4096 / 4;
        // The FAT sectors themselves consume entries, so 109 FAT sectors
        // cover used sectors only through `109 * entries - 109`.
        let exact_used = 109 * fat_entries - 109;
        assert_eq!(
            allocation_table_sector_counts(exact_used, 4096).unwrap(),
            (109, 0)
        );
        assert_eq!(
            allocation_table_sector_counts(exact_used + 1, 4096).unwrap(),
            (110, 1)
        );
    }

    #[test]
    fn failed_replace_cleans_only_unpublished_temporary_file() {
        let destination = test_destination("replace-failure");
        fs::write(&destination, b"old destination").unwrap();
        let writer = writer_with_payload("Saved", b"saved");
        let mut temporary = None;
        let error = writer
            .save_with_hooks(
                &destination,
                |staged, target| {
                    temporary = Some(staged.to_path_buf());
                    assert_eq!(target, destination.as_path());
                    assert_eq!(fs::read(target).unwrap(), b"old destination");
                    Err(io::Error::other("injected replace failure"))
                },
                |_parent| Ok(()),
            )
            .unwrap_err();
        assert!(matches!(
            error,
            SequentialWriteError::Stage {
                progress: SequentialWriteProgress::Complete { .. },
                ..
            }
        ));
        assert!(!temporary.unwrap().exists());
        assert_eq!(fs::read(&destination).unwrap(), b"old destination");
        fs::remove_file(destination).unwrap();
    }

    #[test]
    fn detected_temp_substitution_is_not_deleted_by_cleanup() {
        // This is a best-effort identity regression, not a guarantee against
        // an adversary racing portable path-based replacement and cleanup.
        let destination = test_destination("hostile-temp");
        fs::write(&destination, b"old destination").unwrap();
        let writer = writer_with_payload("Saved", b"saved");
        let mut attacker_path = None;
        let error = writer
            .save_with_hooks(
                &destination,
                |staged, _target| {
                    attacker_path = Some(staged.to_path_buf());
                    fs::remove_file(staged).unwrap();
                    fs::write(staged, b"attacker replacement").unwrap();
                    Err(io::Error::other("injected replacement failure"))
                },
                |_parent| Ok(()),
            )
            .unwrap_err();
        assert!(matches!(error, SequentialWriteError::Stage { .. }));
        let attacker_path = attacker_path.unwrap();
        assert_eq!(fs::read(&attacker_path).unwrap(), b"attacker replacement");
        assert_eq!(fs::read(&destination).unwrap(), b"old destination");
        fs::remove_file(attacker_path).unwrap();
        fs::remove_file(destination).unwrap();
    }

    #[test]
    fn parent_sync_failure_reports_committed_and_keeps_destination() {
        let destination = test_destination("parent-sync-failure");
        fs::write(&destination, b"old destination").unwrap();
        let writer = writer_with_payload("Saved", b"saved");
        let mut temporary = None;
        let error = writer
            .save_with_hooks(
                &destination,
                |staged, target| {
                    temporary = Some(staged.to_path_buf());
                    atomic_replace(staged, target)
                },
                |_parent| Err(io::Error::other("injected parent sync failure")),
            )
            .unwrap_err();
        assert!(matches!(
            error,
            SequentialWriteError::Committed {
                progress: SequentialWriteProgress::Complete { .. },
                ..
            }
        ));
        assert!(!temporary.unwrap().exists());
        let mut parsed = OleFile::open(File::open(&destination).unwrap()).unwrap();
        assert_eq!(parsed.open_stream(&["Saved"]).unwrap(), b"saved");
        fs::remove_file(destination).unwrap();
    }
}
