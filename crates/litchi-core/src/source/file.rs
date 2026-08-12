use std::{
    fs::{File, Metadata},
    io,
    path::Path,
    sync::{Arc, Mutex},
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
    /// Unix device/inode, length, modification time, and status-change time.
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
        Ok(Self {
            state: Arc::new(FileState {
                file,
                source_id,
                version_policy,
                version: Mutex::new(FileVersionState {
                    fingerprint,
                    revision: 0,
                }),
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
    changed_seconds: i64,
    changed_nanoseconds: i64,
}

#[cfg(unix)]
fn metadata_fingerprint(metadata: &Metadata) -> MetadataFingerprint {
    use std::os::unix::fs::MetadataExt;

    MetadataFingerprint {
        device: metadata.dev(),
        inode: metadata.ino(),
        length: metadata.len(),
        modified_seconds: metadata.mtime(),
        modified_nanoseconds: metadata.mtime_nsec(),
        changed_seconds: metadata.ctime(),
        changed_nanoseconds: metadata.ctime_nsec(),
    }
}

#[cfg(windows)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MetadataFingerprint {
    length: u64,
    creation_time: u64,
    last_write_time: u64,
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
