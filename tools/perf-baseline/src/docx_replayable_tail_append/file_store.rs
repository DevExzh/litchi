//! Harness-only bounded file storage for one-shot authored replay.
//!
//! This module deliberately lives below the performance harness rather than in
//! `litchi-docx`.  The file is an explicitly supplied caller-owned provider:
//! construction does not touch the filesystem, `prepare_for_operation` creates
//! one new file, and cleanup is always an explicit operation.  A sealed handle
//! keeps the creating descriptor alive so replacing the pathname cannot redirect
//! an already sealed replay.

#![allow(clippy::module_name_repetitions)]

use litchi_core::{CancellationToken, ExecutionContext, Resource};
use litchi_docx::source_backed::tail_append_stream::{
    AuthoredPassProof, AuthoredReplayError, AuthoredReplayHandle, AuthoredReplayReader,
    AuthoredReplayReference, AuthoredReplayStore, AuthoredStreamProof, ParagraphStreamLimits,
};
use sha2::{Digest as _, Sha256};
use std::fs::{self, File, Metadata, OpenOptions};
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU8, AtomicU64, Ordering};

const HASH_WINDOW_BYTES: usize = 64 * 1024;
const ABSOLUTE_REFERENCE_BYTES: u64 = 4 * 1024 * 1024;
const REFERENCE_MAGIC: &[u8] = b"litchi-perf-file-replay-v1\0";

/// Whether a finished replay file is merely flushed or data-synchronized.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileSyncPolicy {
    /// Flush userspace buffers without asking the filesystem to synchronize data.
    None,
    /// Flush and call [`File::sync_data`] before the file is sealed.
    Data,
}

/// Counters and physical facts collected by one file-store operation.
///
/// The values are observations of the store.  In particular, the configured
/// maximum is not reported as retained bytes or allocated filesystem bytes.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct FileReplayStats {
    /// Number of append calls admitted under the byte ceiling.
    pub append_calls: u64,
    /// Sum of bytes supplied to append calls that completed successfully.
    pub appended_bytes: u64,
    /// Number of `Write::write` calls attempted by append loops.
    pub write_calls: u64,
    /// Number of data synchronization calls.
    pub sync_calls: u64,
    /// Number of replay handles opened.
    pub replay_opens: u64,
    /// Number of reader calls, including terminal EOF calls.
    pub replay_read_calls: u64,
    /// Number of bytes returned by replay readers.
    pub replay_returned_bytes: u64,
    /// Number of complete-file SHA-256 checks performed while sealing.
    pub seal_sha256_checks: u64,
    /// Number of same-pass SHA-256 checks completed by replay readers.
    pub replay_sha256_checks: u64,
    /// Number of complete-file SHA-256 checks performed during cleanup.
    pub cleanup_sha256_checks: u64,
    /// Actual sealed file length, when the file has been sealed.
    pub file_logical_bytes: Option<u64>,
    /// Actual filesystem allocation, when the target exposes it.
    pub file_allocated_bytes: Option<u64>,
    /// Whether explicit cleanup verified and removed the owned path.
    pub file_cleanup_verified: bool,
}

/// Copyable identity and digest facts used for explicit post-operation cleanup.
///
/// This observation owns no path and allocates no heap storage.  Keep it after
/// all replay owners drop, then pass the caller-owned path and this value to
/// [`FileReplayStore::cleanup`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FileReplayObservation {
    /// Stable descriptor identity: device/inode on Unix or volume/file index
    /// on Windows, encoded as two little-endian 64-bit words.
    pub identity: [u8; 16],
    /// Exact file length observed after sealing.
    pub length: u64,
    /// SHA-256 over the exact sealed file bytes.
    pub sha256: [u8; 32],
    /// Filesystem allocation, when exposed by the target platform.
    pub allocated_bytes: Option<u64>,
}

/// Externally owned counter storage for one measured file-store operation.
///
/// Construct this before entering the timed or allocator region and pass a
/// reference to [`FileReplayStore::new_with_monitor`].  The monitor remains
/// usable after the store, handle, and readers have dropped without retaining
/// a cleanup or file owner.
#[derive(Clone, Debug)]
pub struct FileReplayMonitor {
    counters: Arc<FileReplayCounters>,
}

impl FileReplayMonitor {
    /// Allocate an empty monitor for caller-owned lifecycle accounting.
    #[must_use]
    pub fn new() -> Self {
        Self {
            counters: fresh_counters(),
        }
    }

    /// Return the latest counter snapshot.
    #[must_use]
    pub fn stats(&self) -> FileReplayStats {
        self.counters.snapshot()
    }
}

impl Default for FileReplayMonitor {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Default)]
struct FileReplayCounters {
    append_calls: AtomicU64,
    appended_bytes: AtomicU64,
    write_calls: AtomicU64,
    sync_calls: AtomicU64,
    replay_opens: AtomicU64,
    replay_read_calls: AtomicU64,
    replay_returned_bytes: AtomicU64,
    seal_sha256_checks: AtomicU64,
    replay_sha256_checks: AtomicU64,
    cleanup_sha256_checks: AtomicU64,
    file_logical_bytes: AtomicU64,
    file_allocated_bytes: AtomicU64,
    file_allocated_known: AtomicU8,
    file_cleanup_verified: AtomicU8,
}

impl FileReplayCounters {
    fn snapshot(&self) -> FileReplayStats {
        FileReplayStats {
            append_calls: self.append_calls.load(Ordering::Relaxed),
            appended_bytes: self.appended_bytes.load(Ordering::Relaxed),
            write_calls: self.write_calls.load(Ordering::Relaxed),
            sync_calls: self.sync_calls.load(Ordering::Relaxed),
            replay_opens: self.replay_opens.load(Ordering::Relaxed),
            replay_read_calls: self.replay_read_calls.load(Ordering::Relaxed),
            replay_returned_bytes: self.replay_returned_bytes.load(Ordering::Relaxed),
            seal_sha256_checks: self.seal_sha256_checks.load(Ordering::Relaxed),
            replay_sha256_checks: self.replay_sha256_checks.load(Ordering::Relaxed),
            cleanup_sha256_checks: self.cleanup_sha256_checks.load(Ordering::Relaxed),
            file_logical_bytes: (self.file_logical_bytes.load(Ordering::Relaxed) != u64::MAX)
                .then(|| self.file_logical_bytes.load(Ordering::Relaxed)),
            file_allocated_bytes: (self.file_allocated_known.load(Ordering::Relaxed) != 0)
                .then(|| self.file_allocated_bytes.load(Ordering::Relaxed)),
            file_cleanup_verified: self.file_cleanup_verified.load(Ordering::Relaxed) != 0,
        }
    }

    fn set_file_length(&self, length: u64) {
        self.file_logical_bytes.store(length, Ordering::Relaxed);
    }

    fn set_allocated_bytes(&self, bytes: Option<u64>) {
        if let Some(bytes) = bytes {
            self.file_allocated_bytes.store(bytes, Ordering::Relaxed);
            self.file_allocated_known.store(1, Ordering::Relaxed);
        }
    }
}

fn fresh_counters() -> Arc<FileReplayCounters> {
    Arc::new(FileReplayCounters {
        file_logical_bytes: AtomicU64::new(u64::MAX),
        ..FileReplayCounters::default()
    })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct FileIdentity {
    #[cfg(unix)]
    device: u64,
    #[cfg(unix)]
    inode: u64,
    #[cfg(windows)]
    volume_serial: u32,
    #[cfg(windows)]
    file_index: u64,
}

impl FileIdentity {
    fn from_metadata(metadata: &Metadata) -> io::Result<Self> {
        if !metadata.is_file() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "replay path is not a regular file",
            ));
        }
        #[cfg(unix)]
        {
            use std::os::unix::fs::MetadataExt;
            Ok(Self {
                device: metadata.dev(),
                inode: metadata.ino(),
            })
        }
        #[cfg(windows)]
        {
            let _ = metadata;
            Err(io::Error::new(
                io::ErrorKind::Unsupported,
                "stable Windows replay file identity is unavailable",
            ))
        }
        #[cfg(not(any(unix, windows)))]
        {
            let _ = metadata;
            Err(io::Error::new(
                io::ErrorKind::Unsupported,
                "replay file identity is unsupported on this platform",
            ))
        }
    }

    fn bytes(self) -> [u8; 16] {
        let mut bytes = [0_u8; 16];
        #[cfg(unix)]
        {
            bytes[..8].copy_from_slice(&self.device.to_le_bytes());
            bytes[8..].copy_from_slice(&self.inode.to_le_bytes());
        }
        #[cfg(windows)]
        {
            bytes[..8].copy_from_slice(&u64::from(self.volume_serial).to_le_bytes());
            bytes[8..].copy_from_slice(&self.file_index.to_le_bytes());
        }
        bytes
    }
}

fn descriptor_identity(file: &File) -> io::Result<(FileIdentity, u64)> {
    let metadata = file.metadata()?;
    let length = metadata.len();
    Ok((FileIdentity::from_metadata(&metadata)?, length))
}

fn path_metadata(path: &Path) -> io::Result<(FileIdentity, u64)> {
    let metadata = fs::symlink_metadata(path)?;
    let length = metadata.len();
    Ok((FileIdentity::from_metadata(&metadata)?, length))
}

fn same_object(left: FileIdentity, right: FileIdentity) -> bool {
    left == right
}

fn map_operation_error(error: litchi_core::ExecutionError) -> AuthoredReplayError {
    AuthoredReplayError::Provider(Box::new(error))
}

fn check_operation(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<(), AuthoredReplayError> {
    if let Some(context) = context {
        context.check().map_err(map_operation_error)?;
    }
    if let Some(cancellation) = cancellation {
        cancellation.check().map_err(map_operation_error)?;
    }
    Ok(())
}

fn map_io(error: io::Error) -> AuthoredReplayError {
    AuthoredReplayError::Io(error)
}

fn changed_io() -> io::Error {
    io::Error::other(AuthoredReplayError::Changed)
}

fn add_work(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    bytes: usize,
) -> Result<(), AuthoredReplayError> {
    check_operation(context, cancellation)?;
    if let Some(context) = context {
        context
            .consume(Resource::Work, bytes as u64)
            .map_err(map_operation_error)?;
    }
    Ok(())
}

/// A finite replay store backed by one explicitly supplied path.
///
/// `new` only validates arguments and stores the path.  The exclusive file
/// creation occurs in `prepare_for_operation`, which is the timed operation
/// boundary used by the harness.
#[derive(Debug)]
pub struct FileReplayStore {
    path: PathBuf,
    maximum: u64,
    sync_policy: FileSyncPolicy,
    file: Option<File>,
    identity: Option<FileIdentity>,
    logical_len: u64,
    max_patch_bytes: Option<u64>,
    operation_context: Option<ExecutionContext>,
    operation_cancellation: Option<CancellationToken>,
    counters: Arc<FileReplayCounters>,
}

impl FileReplayStore {
    /// Construct an empty store without touching `path`.
    #[cfg(test)]
    pub fn new(
        path: impl AsRef<Path>,
        maximum: u64,
        sync_policy: FileSyncPolicy,
    ) -> Result<Self, AuthoredReplayError> {
        if maximum == 0 || maximum == u64::MAX {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: maximum,
                maximum: maximum.saturating_sub(1),
            });
        }
        Self::new_with_monitor(path, maximum, sync_policy, &FileReplayMonitor::new())
    }

    /// Construct a store using caller-owned counters allocated before timing.
    pub fn new_with_monitor(
        path: impl AsRef<Path>,
        maximum: u64,
        sync_policy: FileSyncPolicy,
        monitor: &FileReplayMonitor,
    ) -> Result<Self, AuthoredReplayError> {
        if maximum == 0 || maximum == u64::MAX {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: maximum,
                maximum: maximum.saturating_sub(1),
            });
        }
        Ok(Self {
            path: path.as_ref().to_path_buf(),
            maximum,
            sync_policy,
            file: None,
            identity: None,
            logical_len: 0,
            max_patch_bytes: None,
            operation_context: None,
            operation_cancellation: None,
            counters: Arc::clone(&monitor.counters),
        })
    }

    /// Verify and remove one sealed replay file after all replay owners drop.
    ///
    /// The caller retains the explicit path and copyable observation outside
    /// the measured owner region, drops every handle and reader, then invokes
    /// this method.  Failed verification leaves the artifact in place.
    pub fn cleanup(
        path: impl AsRef<Path>,
        observation: FileReplayObservation,
    ) -> Result<FileReplayStats, AuthoredReplayError> {
        let path = path.as_ref();
        let file = OpenOptions::new().read(true).open(path).map_err(map_io)?;
        let (identity, length) = descriptor_identity(&file).map_err(map_io)?;
        if identity.bytes() != observation.identity || length != observation.length {
            return Err(AuthoredReplayError::Changed);
        }
        let counters = FileReplayCounters {
            file_logical_bytes: AtomicU64::new(u64::MAX),
            ..FileReplayCounters::default()
        };
        let (actual_length, actual_hash, allocated) =
            Self::sync_and_hash(&file, identity, None, None, &counters.cleanup_sha256_checks)?;
        if actual_length != observation.length || actual_hash != observation.sha256 {
            return Err(AuthoredReplayError::Changed);
        }
        drop(file);
        let (path_identity, path_length) = path_metadata(path).map_err(map_io)?;
        if path_identity.bytes() != observation.identity || path_length != observation.length {
            return Err(AuthoredReplayError::Changed);
        }
        fs::remove_file(path).map_err(map_io)?;
        match path_metadata(path) {
            Ok(_) => {
                return Err(AuthoredReplayError::Store(
                    "replay file cleanup did not remove the owned path",
                ));
            },
            Err(error) if error.kind() == io::ErrorKind::NotFound => {},
            Err(error) => return Err(map_io(error)),
        }
        Ok(FileReplayStats {
            file_logical_bytes: Some(actual_length),
            file_allocated_bytes: allocated,
            file_cleanup_verified: true,
            cleanup_sha256_checks: counters.cleanup_sha256_checks.load(Ordering::Relaxed),
            ..FileReplayStats::default()
        })
    }

    fn ensure_prepared(&self) -> Result<(), AuthoredReplayError> {
        if self.file.is_none() {
            return Err(AuthoredReplayError::Store(
                "replay file was not prepared before use",
            ));
        }
        Ok(())
    }

    fn check_current_operation(&self) -> Result<(), AuthoredReplayError> {
        check_operation(
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
        )
    }

    fn sync_and_hash(
        file: &File,
        expected_identity: FileIdentity,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        hash_checks: &AtomicU64,
    ) -> Result<(u64, [u8; 32], Option<u64>), AuthoredReplayError> {
        let (identity, length) = descriptor_identity(file).map_err(map_io)?;
        if !same_object(identity, expected_identity) {
            return Err(AuthoredReplayError::Changed);
        }
        if length > u64::MAX - 1 {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: length,
                maximum: u64::MAX - 1,
            });
        }
        let mut window = [0_u8; HASH_WINDOW_BYTES];
        let mut offset = 0_u64;
        let mut digest = Sha256::new();
        while offset < length {
            let remaining = length - offset;
            let requested = usize::try_from(remaining)
                .unwrap_or(HASH_WINDOW_BYTES)
                .min(HASH_WINDOW_BYTES);
            let count = read_at(file, &mut window[..requested], offset).map_err(map_io)?;
            if count == 0 || count > requested {
                return Err(AuthoredReplayError::Changed);
            }
            add_work(context, cancellation, count)?;
            digest.update(&window[..count]);
            offset = offset
                .checked_add(count as u64)
                .ok_or(AuthoredReplayError::Changed)?;
        }
        let (after_identity, after_length) = descriptor_identity(file).map_err(map_io)?;
        if !same_object(after_identity, expected_identity) || after_length != length {
            return Err(AuthoredReplayError::Changed);
        }
        hash_checks.fetch_add(1, Ordering::Relaxed);
        let allocated = allocated_bytes(&file.metadata().map_err(map_io)?);
        Ok((length, digest.finalize().into(), allocated))
    }

    fn build_reference(
        &self,
        identity: FileIdentity,
        proof: AuthoredStreamProof,
        maximum: u64,
    ) -> Result<AuthoredReplayReference, AuthoredReplayError> {
        let path = self.path.to_string_lossy();
        let path_bytes = path.as_bytes();
        let total = REFERENCE_MAGIC
            .len()
            .checked_add(8)
            .and_then(|value| value.checked_add(path_bytes.len()))
            .and_then(|value| value.checked_add(16 + 8 + 32))
            .and_then(|value| value.checked_add(32))
            .ok_or(AuthoredReplayError::Limit {
                resource: "durable replay reference bytes",
                actual: u64::MAX,
                maximum: ABSOLUTE_REFERENCE_BYTES,
            })?;
        let total_u64 = u64::try_from(total).map_err(|_| AuthoredReplayError::Limit {
            resource: "durable replay reference bytes",
            actual: u64::MAX,
            maximum: ABSOLUTE_REFERENCE_BYTES,
        })?;
        if total_u64 == 0 || total_u64 > ABSOLUTE_REFERENCE_BYTES || total_u64 > maximum {
            return Err(AuthoredReplayError::Limit {
                resource: "durable replay reference bytes",
                actual: total_u64,
                maximum: maximum.min(ABSOLUTE_REFERENCE_BYTES),
            });
        }
        let mut token = Vec::new();
        token.try_reserve_exact(total).map_err(|_| {
            AuthoredReplayError::Store("durable replay reference allocation failed")
        })?;
        if token.capacity() != total {
            return Err(AuthoredReplayError::Store(
                "durable replay reference allocation exceeded its reservation",
            ));
        }
        token.extend_from_slice(REFERENCE_MAGIC);
        token.extend_from_slice(&(path_bytes.len() as u64).to_le_bytes());
        token.extend_from_slice(path_bytes);
        append_identity(&mut token, identity);
        token.extend_from_slice(&proof.encoded_xml_bytes.to_le_bytes());
        token.extend_from_slice(&proof.encoded_sha256);
        AuthoredReplayReference::try_from_bytes(&token, ABSOLUTE_REFERENCE_BYTES)
    }
}

impl AuthoredReplayStore for FileReplayStore {
    type Handle = FileReplayHandle;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        if self.file.is_some() {
            return Err(AuthoredReplayError::Store(
                "replay file was prepared more than once",
            ));
        }
        if self.maximum > limits.max_replay_bytes {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: self.maximum,
                maximum: limits.max_replay_bytes,
            });
        }
        check_operation(context, cancellation)?;
        let file = OpenOptions::new()
            .create_new(true)
            .read(true)
            .write(true)
            .open(&self.path)
            .map_err(map_io)?;
        let (identity, length) = descriptor_identity(&file).map_err(map_io)?;
        if length != 0 {
            return Err(AuthoredReplayError::Changed);
        }
        self.file = Some(file);
        self.identity = Some(identity);
        self.max_patch_bytes = Some(limits.max_patch_bytes);
        self.operation_context = context.cloned();
        self.operation_cancellation = cancellation.cloned();
        self.logical_len = 0;
        check_operation(context, cancellation)
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.ensure_prepared()?;
        self.check_current_operation()?;
        let file = self.file.as_ref().ok_or(AuthoredReplayError::Store(
            "replay file was not prepared before use",
        ))?;
        let (identity, length) = descriptor_identity(file).map_err(map_io)?;
        if self.identity != Some(identity) || length != self.logical_len {
            return Err(AuthoredReplayError::Changed);
        }
        let amount = u64::try_from(chunk.len()).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: u64::MAX,
            maximum: self.maximum,
        })?;
        let next = self
            .logical_len
            .checked_add(amount)
            .ok_or(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: u64::MAX,
                maximum: self.maximum,
            })?;
        if next > self.maximum {
            return Err(AuthoredReplayError::Limit {
                resource: "replay bytes",
                actual: next,
                maximum: self.maximum,
            });
        }
        self.counters.append_calls.fetch_add(1, Ordering::Relaxed);
        let file = self.file.as_mut().ok_or(AuthoredReplayError::Store(
            "replay file was not prepared before use",
        ))?;
        let mut written = 0_usize;
        while written < chunk.len() {
            let count = file.write(&chunk[written..]).map_err(map_io)?;
            if count == 0 || count > chunk.len() - written {
                return Err(AuthoredReplayError::Io(io::Error::new(
                    io::ErrorKind::WriteZero,
                    "replay file write returned an invalid byte count",
                )));
            }
            self.counters.write_calls.fetch_add(1, Ordering::Relaxed);
            written += count;
        }
        self.logical_len = next;
        self.counters
            .appended_bytes
            .fetch_add(amount, Ordering::Relaxed);
        self.check_current_operation()
    }

    fn finish(mut self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        self.ensure_prepared()?;
        self.check_current_operation()?;
        {
            let file = self.file.as_mut().ok_or(AuthoredReplayError::Store(
                "replay file was not prepared before use",
            ))?;
            file.flush().map_err(map_io)?;
            if self.sync_policy == FileSyncPolicy::Data {
                self.counters.sync_calls.fetch_add(1, Ordering::Relaxed);
                file.sync_data().map_err(map_io)?;
            }
        }
        self.check_current_operation()?;
        let identity = self.identity.ok_or(AuthoredReplayError::Store(
            "replay file identity was not captured",
        ))?;
        let file = self.file.as_ref().ok_or(AuthoredReplayError::Store(
            "replay file was not prepared before use",
        ))?;
        let (length, digest, allocated) = Self::sync_and_hash(
            file,
            identity,
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
            &self.counters.seal_sha256_checks,
        )?;
        self.counters.set_file_length(length);
        self.counters.set_allocated_bytes(allocated);
        if length != self.logical_len
            || length != proof.encoded_xml_bytes
            || digest != proof.encoded_sha256
        {
            return Err(AuthoredReplayError::Changed);
        }
        let durable_reference = self.build_reference(
            identity,
            proof,
            self.max_patch_bytes.unwrap_or(ABSOLUTE_REFERENCE_BYTES),
        )?;
        let file = self.file.take().ok_or(AuthoredReplayError::Store(
            "replay file was not prepared before use",
        ))?;
        let storage = Arc::new(FileReplayStorage {
            file,
            identity,
            counters: Arc::clone(&self.counters),
            operation_context: self.operation_context,
            operation_cancellation: self.operation_cancellation,
        });
        Ok(FileReplayHandle {
            storage,
            proof,
            durable_reference,
        })
    }
}

#[derive(Debug)]
struct FileReplayStorage {
    file: File,
    identity: FileIdentity,
    counters: Arc<FileReplayCounters>,
    operation_context: Option<ExecutionContext>,
    operation_cancellation: Option<CancellationToken>,
}

/// Sealed positional replay handle retained by a file descriptor.
#[derive(Debug, Clone)]
pub struct FileReplayHandle {
    storage: Arc<FileReplayStorage>,
    proof: AuthoredStreamProof,
    durable_reference: AuthoredReplayReference,
}

impl FileReplayHandle {
    /// Return copyable identity and digest facts for explicit cleanup.
    #[must_use]
    pub fn observation(&self) -> FileReplayObservation {
        FileReplayObservation {
            identity: self.storage.identity.bytes(),
            length: self.proof.encoded_xml_bytes,
            sha256: self.proof.encoded_sha256,
            allocated_bytes: self.storage.counters.snapshot().file_allocated_bytes,
        }
    }

    fn open_reader<'a>(
        &'a self,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        charge_work: bool,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        check_operation(context, cancellation)?;
        let (identity, length) = descriptor_identity(&self.storage.file).map_err(map_io)?;
        if !same_object(identity, self.storage.identity) || length != self.proof.encoded_xml_bytes {
            return Err(AuthoredReplayError::Changed);
        }
        let object_reservation = context
            .map(|context| {
                context
                    .reserve(Resource::Objects, 1)
                    .map_err(map_operation_error)
            })
            .transpose()?;
        self.storage
            .counters
            .replay_opens
            .fetch_add(1, Ordering::Relaxed);
        Ok(Box::new(FileReplayReader {
            _storage: Arc::clone(&self.storage),
            expected_identity: self.storage.identity,
            proof: self.proof,
            position: 0,
            bytes: 0,
            hash: Sha256::new(),
            eof: false,
            failed: false,
            operation_context: context.cloned(),
            operation_cancellation: cancellation.cloned(),
            charge_work,
            _object_reservation: object_reservation,
        }))
    }
}

impl AuthoredReplayHandle for FileReplayHandle {
    fn proof(&self) -> AuthoredStreamProof {
        self.proof
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        self.open_reader(
            self.storage.operation_context.as_ref(),
            self.storage.operation_cancellation.as_ref(),
            true,
        )
    }

    fn open_for_package<'a>(
        &'a self,
        context: Option<&'a ExecutionContext>,
        cancellation: Option<&'a CancellationToken>,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        self.open_reader(context, cancellation, false)
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        Some(self.durable_reference.clone())
    }
}

struct FileReplayReader {
    _storage: Arc<FileReplayStorage>,
    expected_identity: FileIdentity,
    proof: AuthoredStreamProof,
    position: u64,
    bytes: u64,
    hash: Sha256,
    eof: bool,
    failed: bool,
    operation_context: Option<ExecutionContext>,
    operation_cancellation: Option<CancellationToken>,
    charge_work: bool,
    _object_reservation: Option<litchi_core::Reservation>,
}

impl Read for FileReplayReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.failed {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        check_operation(
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
        )
        .map_err(|error| {
            self.failed = true;
            io::Error::other(error)
        })?;
        self.storage_counters()
            .replay_read_calls
            .fetch_add(1, Ordering::Relaxed);
        if output.is_empty() {
            return Ok(0);
        }
        let (identity, length) = descriptor_identity(&self._storage.file).map_err(|error| {
            self.failed = true;
            io::Error::other(AuthoredReplayError::Io(error))
        })?;
        if !same_object(identity, self.expected_identity) || length != self.proof.encoded_xml_bytes
        {
            self.failed = true;
            return Err(changed_io());
        }
        if self.position == self.proof.encoded_xml_bytes {
            self.eof = true;
            return Ok(0);
        }
        let remaining = self.proof.encoded_xml_bytes - self.position;
        let requested = usize::try_from(remaining)
            .unwrap_or(usize::MAX)
            .min(output.len());
        let count = read_at(&self._storage.file, &mut output[..requested], self.position).map_err(
            |error| {
                self.failed = true;
                io::Error::other(AuthoredReplayError::Io(error))
            },
        )?;
        if count == 0 || count > requested {
            self.failed = true;
            return Err(changed_io());
        }
        if self.charge_work {
            add_work(
                self.operation_context.as_ref(),
                self.operation_cancellation.as_ref(),
                count,
            )
            .map_err(|error| {
                self.failed = true;
                io::Error::other(error)
            })?;
        }
        self.hash.update(&output[..count]);
        self.position = self.position.checked_add(count as u64).ok_or_else(|| {
            self.failed = true;
            io::Error::other("authored replay length overflow")
        })?;
        self.bytes = self.position;
        self.storage_counters()
            .replay_returned_bytes
            .fetch_add(count as u64, Ordering::Relaxed);
        Ok(count)
    }
}

impl FileReplayReader {
    fn storage_counters(&self) -> &FileReplayCounters {
        &self._storage.counters
    }
}

impl AuthoredReplayReader for FileReplayReader {
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
        if self.failed {
            return Err(AuthoredReplayError::Io(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            )));
        }
        check_operation(
            self.operation_context.as_ref(),
            self.operation_cancellation.as_ref(),
        )?;
        if !self.eof || self.position != self.proof.encoded_xml_bytes {
            return Err(AuthoredReplayError::Invalid(
                "authored replay reader did not reach EOF before finish",
            ));
        }
        let returned_hash: [u8; 32] = self.hash.finalize().into();
        self._storage
            .counters
            .replay_sha256_checks
            .fetch_add(1, Ordering::Relaxed);
        if self.bytes != self.proof.encoded_xml_bytes || returned_hash != self.proof.encoded_sha256
        {
            return Err(AuthoredReplayError::Changed);
        }
        let (identity, length) = descriptor_identity(&self._storage.file).map_err(map_io)?;
        if !same_object(identity, self.expected_identity) || length != self.proof.encoded_xml_bytes
        {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(AuthoredPassProof(self.proof))
    }
}

fn append_identity(token: &mut Vec<u8>, identity: FileIdentity) {
    #[cfg(unix)]
    {
        token.extend_from_slice(&identity.device.to_le_bytes());
        token.extend_from_slice(&identity.inode.to_le_bytes());
    }
    #[cfg(windows)]
    {
        token.extend_from_slice(&u64::from(identity.volume_serial).to_le_bytes());
        token.extend_from_slice(&identity.file_index.to_le_bytes());
    }
    #[cfg(not(any(unix, windows)))]
    {
        let _ = (token, identity);
    }
}

fn allocated_bytes(metadata: &Metadata) -> Option<u64> {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        metadata.blocks().checked_mul(512)
    }
    #[cfg(not(unix))]
    {
        let _ = metadata;
        None
    }
}

#[cfg(unix)]
fn read_at(file: &File, output: &mut [u8], offset: u64) -> io::Result<usize> {
    use std::os::unix::fs::FileExt;
    file.read_at(output, offset)
}

#[cfg(windows)]
fn read_at(file: &File, output: &mut [u8], offset: u64) -> io::Result<usize> {
    use std::os::windows::fs::FileExt;
    file.seek_read(output, offset)
}

#[cfg(not(any(unix, windows)))]
fn read_at(_file: &File, _output: &mut [u8], _offset: u64) -> io::Result<usize> {
    Err(io::Error::new(
        io::ErrorKind::Unsupported,
        "positional replay file reads are unsupported on this platform",
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Seek as _, SeekFrom};
    use std::time::{SystemTime, UNIX_EPOCH};

    fn unique_directory() -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock before epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("litchi-file-replay-{nanos}"));
        fs::create_dir(&path).expect("create test directory");
        path
    }

    fn proof(bytes: &[u8]) -> AuthoredStreamProof {
        AuthoredStreamProof {
            strict_namespace: false,
            paragraph_count: 1,
            event_count: 3,
            text_bytes: 0,
            encoded_xml_bytes: bytes.len() as u64,
            event_sha256: [7; 32],
            encoded_sha256: Sha256::digest(bytes).into(),
        }
    }

    fn prepared(path: &Path, maximum: u64) -> FileReplayStore {
        let mut store = FileReplayStore::new(path, maximum, FileSyncPolicy::None).unwrap();
        store
            .prepare_for_operation(ParagraphStreamLimits::default(), None, None)
            .unwrap();
        store
    }

    fn seal(path: &Path, bytes: &[u8]) -> (FileReplayHandle, FileReplayObservation) {
        let mut store = prepared(path, bytes.len() as u64 + 1);
        store.append(bytes).unwrap();
        let handle = store.finish(proof(bytes)).unwrap();
        let observation = handle.observation();
        (handle, observation)
    }

    #[test]
    fn append_refuses_before_crossing_cap() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let mut store = prepared(&path, 3);
        let error = store.append(b"four").unwrap_err();
        assert!(matches!(error, AuthoredReplayError::Limit { .. }));
        assert_eq!(fs::metadata(&path).unwrap().len(), 0);
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn finish_rejects_wrong_length_and_hash() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let mut store = prepared(&path, 32);
        store.append(b"payload").unwrap();
        let mut wrong = proof(b"payload");
        wrong.encoded_xml_bytes += 1;
        assert!(matches!(
            store.finish(wrong),
            Err(AuthoredReplayError::Changed)
        ));
        assert!(path.exists());
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn finish_rejects_wrong_hash() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let mut store = prepared(&path, 32);
        store.append(b"payload").unwrap();
        let mut wrong = proof(b"payload");
        wrong.encoded_sha256[0] ^= 1;
        assert!(matches!(
            store.finish(wrong),
            Err(AuthoredReplayError::Changed)
        ));
        assert!(path.exists());
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn changed_file_is_refused_before_first_byte() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let (handle, observation) = seal(&path, b"payload");
        let file = OpenOptions::new().write(true).open(&path).unwrap();
        file.set_len(6).unwrap();
        drop(file);
        assert!(matches!(handle.open(), Err(AuthoredReplayError::Changed)));
        drop(handle);
        // The changed artifact is deliberately preserved by the failed
        // preflight; explicit cleanup refuses it rather than deleting it.
        assert!(matches!(
            FileReplayStore::cleanup(&path, observation),
            Err(AuthoredReplayError::Changed)
        ));
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn mutation_after_prefix_preserves_typed_partial_failure() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let (handle, observation) = seal(&path, b"payload");
        let mut reader = handle.open().unwrap();
        let mut prefix = [0; 3];
        reader.read_exact(&mut prefix).unwrap();
        assert_eq!(&prefix, b"pay");
        let mut file = OpenOptions::new().write(true).open(&path).unwrap();
        file.seek(SeekFrom::Start(3)).unwrap();
        file.write_all(b"X").unwrap();
        drop(file);
        let mut suffix = Vec::new();
        reader.read_to_end(&mut suffix).unwrap();
        assert_eq!(suffix, b"Xoad");
        assert!(matches!(reader.finish(), Err(AuthoredReplayError::Changed)));
        drop(handle);
        assert!(matches!(
            FileReplayStore::cleanup(&path, observation),
            Err(AuthoredReplayError::Changed)
        ));
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn readers_have_independent_positional_offsets() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let (handle, observation) = seal(&path, b"0123456789");
        let mut first = handle.open().unwrap();
        let mut second = handle.open().unwrap();
        let mut first_bytes = [0; 3];
        let mut second_bytes = [0; 5];
        first.read_exact(&mut first_bytes).unwrap();
        second.read_exact(&mut second_bytes).unwrap();
        assert_eq!(&first_bytes, b"012");
        assert_eq!(&second_bytes, b"01234");
        let mut rest = Vec::new();
        first.read_to_end(&mut rest).unwrap();
        assert_eq!(rest, b"3456789");
        let mut rest = Vec::new();
        second.read_to_end(&mut rest).unwrap();
        assert_eq!(rest, b"56789");
        first.finish().unwrap();
        second.finish().unwrap();
        drop(handle);
        FileReplayStore::cleanup(&path, observation).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn preexisting_path_is_rejected_exclusively() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        fs::write(&path, b"existing").unwrap();
        let mut store = FileReplayStore::new(&path, 32, FileSyncPolicy::None).unwrap();
        let error = store
            .prepare_for_operation(ParagraphStreamLimits::default(), None, None)
            .unwrap_err();
        assert!(matches!(error, AuthoredReplayError::Io(_)));
        assert_eq!(fs::read(&path).unwrap(), b"existing");
        fs::remove_file(path).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn pathname_replacement_does_not_redirect_retained_handle() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let (handle, observation) = seal(&path, b"payload");
        let moved = directory.join("old.bin");
        fs::rename(&path, &moved).unwrap();
        fs::write(&path, b"payload").unwrap();
        let mut reader = handle.open().unwrap();
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes).unwrap();
        assert_eq!(bytes, b"payload");
        reader.finish().unwrap();
        drop(handle);
        assert!(matches!(
            FileReplayStore::cleanup(&path, observation),
            Err(AuthoredReplayError::Changed)
        ));
        fs::remove_file(path).unwrap();
        fs::remove_file(moved).unwrap();
        fs::remove_dir(directory).unwrap();
    }

    #[test]
    fn cleanup_waits_for_reader_and_removes_only_owned_file() {
        let directory = unique_directory();
        let path = directory.join("payload.bin");
        let (handle, observation) = seal(&path, b"payload");
        let reader = handle.open().unwrap();
        // The static cleanup API deliberately leaves ownership ordering to the
        // harness.  The reader is still alive here, so retain the path and
        // defer cleanup until all owners drop.
        drop(reader);
        drop(handle);
        FileReplayStore::cleanup(&path, observation).unwrap();
        assert!(!path.exists());
        fs::remove_dir(directory).unwrap();
    }
}
