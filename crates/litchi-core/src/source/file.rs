use std::{
    fs::{File, Metadata},
    io,
    path::Path,
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, Ordering, fence},
    },
};

use super::{ReadAt, SourceVersion};

/// Metadata used to revise a [`FileSource`]'s process-local version token.
///
/// Neither policy hashes file contents. They detect ordinary writes and
/// truncations, but a writer capable of restoring all tracked metadata can
/// evade detection. Filesystems may also expose timestamps at a coarser
/// resolution than the operating-system API. Callers that do not trust other
/// writers must additionally validate the bytes they consume. Each metadata
/// transition observed by [`ReadAt::version`] increments the revision; multiple
/// transitions between observations may be reported as one, and a transition
/// fully reverted between observations is not visible.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FileVersionPolicy {
    /// Unix device/inode, length, and modification time.
    ///
    /// The open descriptor remains pinned to its original inode across an
    /// atomic pathname replacement; a new [`FileSource`] opened by pathname
    /// observes the replacement instead.
    UnixMetadata,
    /// Windows length, creation time, and last-write time.
    ///
    /// Stable Rust does not expose a persistent Windows file identifier. The
    /// [`SourceVersion`] identity is therefore process-local to the open
    /// source and must not be persisted or compared with an independently
    /// opened source.
    WindowsTimestampMetadata,
}

#[derive(Debug)]
struct FileState {
    file: File,
    source_id: u64,
    version_policy: FileVersionPolicy,
    version: Mutex<FileVersionState>,
    /// A lock-free copy of `version` for the common case of an unchanged
    /// file (record 0280): readers compare the fingerprint they observe with
    /// it and return its revision without taking the mutex.
    published: Published,
}

/// A sequence-locked copy of the last [`FileVersionState`].
///
/// Only the slow path writes it, while it holds the `version` mutex: it
/// makes `sequence` odd, stores the words, then makes it even again. A
/// reader that sees the same even sequence before and after loading the
/// words has a consistent copy; any other reader takes the slow path. Every
/// word is an atomic, so a torn read is detected, never undefined.
#[derive(Debug)]
struct Published {
    sequence: AtomicU64,
    words: [AtomicU64; FINGERPRINT_WORDS + 1],
}

impl Published {
    fn new(state: &FileVersionState) -> Self {
        let published = Self {
            sequence: AtomicU64::new(0),
            words: std::array::from_fn(|_| AtomicU64::new(0)),
        };
        published.store(state);
        published
    }

    /// Publishes `state`; the caller holds the `version` mutex, so writers
    /// never race each other.
    fn store(&self, state: &FileVersionState) {
        let sequence = self.sequence.load(Ordering::Relaxed);
        self.sequence
            .store(sequence.wrapping_add(1), Ordering::Relaxed);
        fence(Ordering::Release);
        let words = state.fingerprint.words();
        for (slot, word) in self.words.iter().zip(words) {
            slot.store(word, Ordering::Relaxed);
        }
        self.words[FINGERPRINT_WORDS].store(state.revision, Ordering::Relaxed);
        self.sequence
            .store(sequence.wrapping_add(2), Ordering::Release);
    }

    /// The published revision when the published fingerprint is
    /// `fingerprint`; `None` when it differs or a writer is active.
    fn revision_if(&self, fingerprint: &MetadataFingerprint) -> Option<u64> {
        let before = self.sequence.load(Ordering::Acquire);
        if before % 2 == 1 {
            return None;
        }
        let expected = fingerprint.words();
        let mut same = true;
        for (slot, word) in self.words.iter().zip(expected) {
            same &= slot.load(Ordering::Relaxed) == word;
        }
        let revision = self.words[FINGERPRINT_WORDS].load(Ordering::Relaxed);
        fence(Ordering::Acquire);
        let after = self.sequence.load(Ordering::Relaxed);
        (same && before == after).then_some(revision)
    }
}

#[derive(Debug)]
struct FileVersionState {
    fingerprint: MetadataFingerprint,
    revision: u64,
}

/// An O(1)-cloneable positional source backed by an open regular file.
///
/// Reads use the platform's positional file API and never share or mutate a
/// seek cursor. On Windows, construction reopens the supplied handle for
/// overlapped I/O and can fail if the filesystem cannot provide that handle.
/// [`FileSource::open`] follows symlinks in the same way as
/// [`File::open`], validates that the resolved object is a regular file, and
/// then pins that open handle. Replacing the pathname therefore cannot switch
/// an existing source to the replacement file; opening the path again creates
/// a source with a distinct process-local identity.
///
/// The adapter only performs local filesystem API calls. It does not resolve
/// URLs or create network clients; paths on an OS-mounted remote filesystem
/// retain that filesystem's normal operating-system semantics.
#[derive(Debug, Clone)]
pub struct FileSource {
    state: Arc<FileState>,
}

impl FileSource {
    /// Opens `path` without reading its contents into memory.
    ///
    /// Symlinks are followed by the operating system. The resolved object must
    /// be a regular file.
    ///
    /// # Errors
    ///
    /// Returns an error from opening or inspecting the file, or
    /// [`io::ErrorKind::InvalidInput`] when the resolved object is not a
    /// regular file.
    pub fn open(path: impl AsRef<Path>) -> io::Result<Self> {
        Self::from_file(File::open(path)?)
    }

    /// Takes ownership of an already-open regular file without reading it.
    ///
    /// # Errors
    ///
    /// Returns an error from inspecting the file, or
    /// [`io::ErrorKind::InvalidInput`] when the handle is not a regular file.
    pub fn from_file(file: File) -> io::Result<Self> {
        let metadata = file.metadata()?;
        if !metadata.is_file() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "positional filesystem source must be a regular file",
            ));
        }

        let file = prepare_file(file)?;
        let metadata = file.metadata()?;
        let version_policy = version_policy(&metadata);
        let fingerprint = metadata_fingerprint(&metadata);
        let source_id = SourceVersion::fresh().id();
        let state = FileVersionState {
            fingerprint,
            revision: 0,
        };
        Ok(Self {
            state: Arc::new(FileState {
                file,
                source_id,
                version_policy,
                published: Published::new(&state),
                version: Mutex::new(state),
            }),
        })
    }

    /// Returns the metadata fields used to detect external changes.
    #[must_use]
    pub fn version_policy(&self) -> FileVersionPolicy {
        self.state.version_policy
    }
}

impl ReadAt for FileSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.state.file.metadata()?.len())
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        validate_range(offset, output.len())?;
        if output.is_empty() {
            return Ok(0);
        }

        loop {
            let result = positional_read(&self.state.file, output, offset);
            match result {
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                other => return other,
            }
        }
    }

    fn version(&self) -> io::Result<SourceVersion> {
        // Fast path (record 0280): one metadata read and a lock-free compare
        // with the last published fingerprint. Any difference, including a
        // concurrent update, falls through to the locked path, which reads
        // the metadata again under the mutex exactly as before.
        if let Some(fingerprint) = probe_fingerprint(&self.state.file)
            && let Some(revision) = self.state.published.revision_if(&fingerprint)
        {
            return Ok(SourceVersion::new(self.state.source_id, revision));
        }
        self.version_locked()
    }
}

impl FileSource {
    /// The locked revision check: reads the metadata under the mutex,
    /// advances the revision on a fingerprint change and publishes it.
    fn version_locked(&self) -> io::Result<SourceVersion> {
        let mut version = self
            .state
            .version
            .lock()
            .map_err(|_error| io::Error::other("filesystem source version lock is poisoned"))?;
        let metadata = self.state.file.metadata()?;
        if !metadata.is_file() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional filesystem source is no longer a regular file",
            ));
        }
        let fingerprint = metadata_fingerprint(&metadata);
        if version.fingerprint != fingerprint {
            version.revision = version.revision.checked_add(1).ok_or_else(|| {
                io::Error::other("filesystem source revision counter is exhausted")
            })?;
            version.fingerprint = fingerprint;
            self.state.published.store(&version);
        }
        Ok(SourceVersion::new(self.state.source_id, version.revision))
    }
}

fn validate_range(offset: u64, output_len: usize) -> io::Result<()> {
    let Some(last_index) = output_len.checked_sub(1) else {
        return Ok(());
    };
    let last_index = u64::try_from(last_index).map_err(|_error| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "positional read length does not fit u64",
        )
    })?;
    offset
        .checked_add(last_index)
        .map(|_end| ())
        .ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "positional read range exceeds u64",
            )
        })
}

#[cfg(unix)]
fn positional_read(file: &File, output: &mut [u8], offset: u64) -> io::Result<usize> {
    std::os::unix::fs::FileExt::read_at(file, output, offset)
}

#[cfg(windows)]
fn positional_read(file: &File, output: &mut [u8], offset: u64) -> io::Result<usize> {
    use std::{os::windows::io::AsRawHandle, ptr};
    use windows_sys::Win32::{
        Foundation::{ERROR_HANDLE_EOF, ERROR_IO_PENDING, HANDLE},
        Storage::FileSystem::ReadFile,
        System::IO::{GetOverlappedResult, OVERLAPPED},
    };

    let event = EventHandle::create()?;
    let mut overlapped = OVERLAPPED {
        hEvent: event.0,
        ..OVERLAPPED::default()
    };
    // Selecting the offset fields initializes the active union member. Both
    // halves are derived from the same caller-validated `u64` offset.
    overlapped.Anonymous.Anonymous.Offset = offset as u32;
    overlapped.Anonymous.Anonymous.OffsetHigh = (offset >> 32) as u32;
    let amount = u32::try_from(output.len()).unwrap_or(u32::MAX);
    let handle: HANDLE = file.as_raw_handle();

    // SAFETY: `handle` is the live overlapped file owned by `FileSource`;
    // `output` and `overlapped` remain alive and exclusively borrowed until
    // `GetOverlappedResult` confirms completion below. The requested amount is
    // bounded by both `output.len()` and `u32::MAX`.
    let started = unsafe {
        ReadFile(
            handle,
            output.as_mut_ptr(),
            amount,
            ptr::null_mut(),
            &raw mut overlapped,
        )
    };
    if started == 0 {
        let error = io::Error::last_os_error();
        match error.raw_os_error().map(|code| code as u32) {
            Some(ERROR_HANDLE_EOF) => return Ok(0),
            Some(ERROR_IO_PENDING) => {},
            _ => return Err(error),
        }
    }

    let mut transferred = 0_u32;
    // SAFETY: the event and buffers associated with `overlapped` are still
    // alive. Waiting here ensures the kernel no longer accesses either stack
    // value before this function returns.
    let completed =
        unsafe { GetOverlappedResult(handle, &raw const overlapped, &raw mut transferred, 1) };
    if completed == 0 {
        let error = io::Error::last_os_error();
        if error.raw_os_error().map(|code| code as u32) == Some(ERROR_HANDLE_EOF) {
            return Ok(0);
        }
        return Err(error);
    }

    usize::try_from(transferred).map_err(|_error| {
        io::Error::new(
            io::ErrorKind::InvalidData,
            "Windows positional read count does not fit usize",
        )
    })
}

#[cfg(unix)]
fn prepare_file(file: File) -> io::Result<File> {
    Ok(file)
}

#[cfg(windows)]
fn prepare_file(file: File) -> io::Result<File> {
    use std::os::windows::io::{AsRawHandle, FromRawHandle};
    use windows_sys::Win32::{
        Foundation::{GENERIC_READ, HANDLE, INVALID_HANDLE_VALUE},
        Storage::FileSystem::{
            FILE_FLAG_OVERLAPPED, FILE_SHARE_DELETE, FILE_SHARE_READ, FILE_SHARE_WRITE, ReOpenFile,
        },
    };

    let original: HANDLE = file.as_raw_handle();
    // SAFETY: `original` remains live for this call. On success, `ReOpenFile`
    // returns a distinct owned handle referring to the same file object.
    let reopened = unsafe {
        ReOpenFile(
            original,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            FILE_FLAG_OVERLAPPED,
        )
    };
    if reopened == INVALID_HANDLE_VALUE {
        return Err(io::Error::last_os_error());
    }

    // SAFETY: `reopened` is a fresh, valid, uniquely owned file handle. Moving
    // it into `File` transfers responsibility for closing it exactly once.
    Ok(unsafe { File::from_raw_handle(reopened) })
}

#[cfg(windows)]
struct EventHandle(windows_sys::Win32::Foundation::HANDLE);

#[cfg(windows)]
impl EventHandle {
    fn create() -> io::Result<Self> {
        use std::ptr;
        use windows_sys::Win32::System::Threading::CreateEventW;

        // SAFETY: null security attributes and name request a process-local,
        // unnamed manual-reset event with an initially nonsignaled state.
        let handle = unsafe { CreateEventW(ptr::null(), 1, 0, ptr::null()) };
        if handle.is_null() {
            Err(io::Error::last_os_error())
        } else {
            Ok(Self(handle))
        }
    }
}

#[cfg(windows)]
impl Drop for EventHandle {
    fn drop(&mut self) {
        // SAFETY: this wrapper exclusively owns the valid event handle.
        let _closed = unsafe { windows_sys::Win32::Foundation::CloseHandle(self.0) };
    }
}

#[cfg(unix)]
const fn version_policy(_metadata: &Metadata) -> FileVersionPolicy {
    FileVersionPolicy::UnixMetadata
}

#[cfg(windows)]
const fn version_policy(_metadata: &Metadata) -> FileVersionPolicy {
    FileVersionPolicy::WindowsTimestampMetadata
}

#[cfg(unix)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MetadataFingerprint {
    device: u64,
    inode: u64,
    length: u64,
    modified_seconds: i64,
    modified_nanoseconds: i64,
}

/// The fingerprint of a regular file from one `fstat`, or `None` when the
/// call fails or the file is no longer regular (the locked path then reports
/// the error). It reads the same fields as [`metadata_fingerprint`] without
/// the `statx` field set and the `Metadata` conversion of
/// [`File::metadata`], about a fifth of the check's cost (record 0280). A
/// value that ever differed from [`metadata_fingerprint`] would only send
/// every check to the locked path.
#[cfg(unix)]
#[allow(
    unsafe_code,
    trivial_numeric_casts,
    clippy::cast_sign_loss,
    clippy::cast_possible_wrap,
    clippy::useless_conversion,
    clippy::unnecessary_cast,
    reason = "one FFI call; the casts copy platform-sized stat fields bit for bit"
)]
fn probe_fingerprint(file: &File) -> Option<MetadataFingerprint> {
    use std::os::fd::AsRawFd;

    let mut stat = std::mem::MaybeUninit::<libc::stat>::uninit();
    // SAFETY: the descriptor is owned by `file`, which outlives the call, and
    // `stat` points to writable storage of the type `fstat` fills.
    let status = unsafe { libc::fstat(file.as_raw_fd(), stat.as_mut_ptr()) };
    if status != 0 {
        return None;
    }
    // SAFETY: `fstat` returned 0, so it initialized the whole structure.
    let stat = unsafe { stat.assume_init() };
    if stat.st_mode & libc::S_IFMT != libc::S_IFREG {
        return None;
    }
    Some(MetadataFingerprint {
        device: stat.st_dev as u64,
        inode: stat.st_ino as u64,
        length: stat.st_size as u64,
        modified_seconds: stat.st_mtime as i64,
        modified_nanoseconds: stat.st_mtime_nsec as i64,
    })
}

/// The fingerprint of `file` through [`File::metadata`], or `None`.
#[cfg(windows)]
fn probe_fingerprint(file: &File) -> Option<MetadataFingerprint> {
    let metadata = file.metadata().ok()?;
    metadata.is_file().then(|| metadata_fingerprint(&metadata))
}

/// The number of `u64` words a [`MetadataFingerprint`] publishes as.
#[cfg(unix)]
const FINGERPRINT_WORDS: usize = 5;

#[cfg(unix)]
impl MetadataFingerprint {
    #[allow(
        clippy::cast_sign_loss,
        reason = "bit-preserving: the words are only compared for equality"
    )]
    const fn words(&self) -> [u64; FINGERPRINT_WORDS] {
        [
            self.device,
            self.inode,
            self.length,
            self.modified_seconds as u64,
            self.modified_nanoseconds as u64,
        ]
    }
}

#[cfg(unix)]
fn metadata_fingerprint(metadata: &Metadata) -> MetadataFingerprint {
    use std::os::unix::fs::MetadataExt;

    // Keep a pinned descriptor readable across atomic pathname replacement.
    // Unlinking the old inode changes ctime even though its contents do not;
    // length and mtime still detect ordinary in-place writes.
    MetadataFingerprint {
        device: metadata.dev(),
        inode: metadata.ino(),
        length: metadata.len(),
        modified_seconds: metadata.mtime(),
        modified_nanoseconds: metadata.mtime_nsec(),
    }
}

#[cfg(windows)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MetadataFingerprint {
    length: u64,
    creation_time: u64,
    last_write_time: u64,
}

/// The number of `u64` words a [`MetadataFingerprint`] publishes as.
#[cfg(windows)]
const FINGERPRINT_WORDS: usize = 3;

#[cfg(windows)]
impl MetadataFingerprint {
    const fn words(&self) -> [u64; FINGERPRINT_WORDS] {
        [self.length, self.creation_time, self.last_write_time]
    }
}

#[cfg(windows)]
fn metadata_fingerprint(metadata: &Metadata) -> MetadataFingerprint {
    use std::os::windows::fs::MetadataExt;

    MetadataFingerprint {
        length: metadata.file_size(),
        creation_time: metadata.creation_time(),
        last_write_time: metadata.last_write_time(),
    }
}

#[cfg(all(test, unix))]
mod tests {
    #![allow(clippy::expect_used, reason = "test assertions panic by design")]

    use super::{metadata_fingerprint, probe_fingerprint};

    /// The fast probe reads the fingerprint the locked path compares, so an
    /// unchanged file takes the fast path (record 0280).
    #[test]
    fn the_fast_probe_matches_the_metadata_fingerprint() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let path = directory.path().join("probe.bin");
        std::fs::write(&path, b"probe").expect("write");
        let file = std::fs::File::open(&path).expect("open");
        let metadata = file.metadata().expect("metadata");
        assert_eq!(
            probe_fingerprint(&file),
            Some(metadata_fingerprint(&metadata))
        );
        let directory_handle = std::fs::File::open(directory.path()).expect("open directory");
        assert_eq!(probe_fingerprint(&directory_handle), None);
    }
}
