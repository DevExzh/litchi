use crate::{
    CENTRAL_HEADER_SIGNATURE, CompressionMethod, DataDescriptor,
    END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE, END_OF_CENTRAL_DIR_SIGNATURE64,
    END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES, Error, Header, ZipFileHeaderFixed, ZipLocalFileHeaderFixed,
    accounting::{AccountingWriteKind, ZipOperationAccounting, usize_to_u64, write_all_counted},
    crc,
    errors::ErrorKind,
    extra_fields::{ExtraFieldId, ExtraFieldsContainer},
    mode::CREATOR_UNIX,
    path::{NormalizedPath, ZipFilePath},
    time::{DosDateTime, UtcDateTime},
};
use flate2::{Compress, Compression, FlushCompress, Status};
use std::io::{self, Read, Seek, SeekFrom, Write};

// ZIP64 constants
const ZIP64_VERSION_NEEDED: u16 = 45; // 4.5
const ZIP64_EOCD_SIZE: usize = 56;

// General purpose bit flags
const FLAG_DATA_DESCRIPTOR: u16 = 0x08; // bit 3: data descriptor present
const FLAG_UTF8_ENCODING: u16 = 0x800; // bit 11: UTF-8 encoding flag (EFS)

// ZIP64 thresholds - when to switch to ZIP64 format
const ZIP64_THRESHOLD_FILE_SIZE: u64 = u32::MAX as u64;
const ZIP64_THRESHOLD_OFFSET: u64 = u32::MAX as u64;
const ZIP64_THRESHOLD_ENTRIES: usize = u16::MAX as usize;
const ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN: u16 = 16;
const ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN: u16 = 20;
const ZIP64_LOCAL_SIZE_EXTRA_MAX_LEN: usize = 20;

fn io_error_is_interrupted(error: &io::Error) -> bool {
    error.kind() == io::ErrorKind::Interrupted
}

fn zip_error_is_interrupted(error: &Error) -> bool {
    matches!(
        error.kind(),
        ErrorKind::IO(source) | ErrorKind::Io(source)
            if source.kind() == io::ErrorKind::Interrupted
    )
}

/// Limits for a caller-provided central-directory spool.
///
/// The spool stores finalized central-directory records while member payloads
/// are written to the output sink. `max_bytes` bounds the complete serialized
/// directory extent; `buffer_bytes` bounds the fixed replay buffer used while
/// copying that extent to the final output. The replay buffer is operation
/// scratch and is not counted against the serialized extent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DirectorySpoolLimits {
    /// Maximum serialized central-directory bytes retained by the spool.
    pub max_bytes: u64,
    /// Size of the bounded buffer used to replay the spool at finalization.
    pub buffer_bytes: usize,
}

impl DirectorySpoolLimits {
    /// Creates a central-directory spool limit policy.
    pub const fn new(max_bytes: u64, buffer_bytes: usize) -> Self {
        Self {
            max_bytes,
            buffer_bytes,
        }
    }

    /// Returns the maximum serialized central-directory extent.
    pub const fn max_bytes(self) -> u64 {
        self.max_bytes
    }

    /// Returns the fixed replay buffer size.
    pub const fn buffer_bytes(self) -> usize {
        self.buffer_bytes
    }

    fn validate(self) -> Result<(), Error> {
        if self.buffer_bytes == 0 {
            return Err(ErrorKind::InvalidInput {
                msg: "central-directory spool replay buffer must be non-zero".to_string(),
            }
            .into());
        }
        Ok(())
    }
}

/// The erased caller-owned storage capability used by the optional central
/// directory spool. The bound intentionally includes `Send` and `Sync` so
/// adding the optional mode does not weaken the writer's existing auto-trait
/// behavior.
trait DirectorySpoolStore: Read + Write + Seek + Send + Sync + 'static {}

impl<T> DirectorySpoolStore for T where T: Read + Write + Seek + Send + Sync + 'static {}

struct DirectorySpool {
    store: Box<dyn DirectorySpoolStore>,
    base_offset: u64,
    extent: u64,
    entries: u64,
    saw_zip64: bool,
    limits: DirectorySpoolLimits,
    replay_buffer: Vec<u8>,
    poisoned: bool,
}

impl std::fmt::Debug for DirectorySpool {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DirectorySpool")
            .field("base_offset", &self.base_offset)
            .field("extent", &self.extent)
            .field("entries", &self.entries)
            .field("saw_zip64", &self.saw_zip64)
            .field("limits", &self.limits)
            .field("poisoned", &self.poisoned)
            .finish()
    }
}

impl DirectorySpool {
    fn new<S>(mut store: S, limits: DirectorySpoolLimits) -> Result<Self, Error>
    where
        S: DirectorySpoolStore,
    {
        limits.validate()?;
        let mut replay_buffer = Vec::new();
        replay_buffer
            .try_reserve_exact(limits.buffer_bytes)
            .map_err(|source| ErrorKind::Allocation {
                resource: "central-directory spool replay buffer",
                source,
            })?;
        replay_buffer.resize(limits.buffer_bytes, 0);
        let base_offset = loop {
            match store.seek(SeekFrom::End(0)) {
                Ok(offset) => break offset,
                Err(source) if source.kind() == io::ErrorKind::Interrupted => continue,
                Err(source) => {
                    return Err(ErrorKind::CentralDirectorySpool {
                        operation: "position",
                        source,
                    }
                    .into());
                },
            }
        };
        if base_offset.checked_add(limits.max_bytes).is_none() {
            return Err(ErrorKind::InvalidInput {
                msg: "central-directory spool base offset plus maximum extent overflows u64"
                    .to_string(),
            }
            .into());
        }
        Ok(Self {
            store: Box::new(store),
            base_offset,
            extent: 0,
            entries: 0,
            saw_zip64: false,
            limits,
            replay_buffer,
            poisoned: false,
        })
    }

    fn healthy(&self) -> Result<(), Error> {
        if self.poisoned {
            return Err(ErrorKind::InvalidInput {
                msg: "central-directory spool is poisoned after a prior failure".to_string(),
            }
            .into());
        }
        Ok(())
    }

    fn capacity_for(&self, additional: u64) -> Result<u64, Error> {
        let actual =
            self.extent
                .checked_add(additional)
                .ok_or_else(|| ErrorKind::InvalidInput {
                    msg: "central-directory spool extent overflows u64".to_string(),
                })?;
        if actual > self.limits.max_bytes {
            return Err(ErrorKind::CentralDirectorySpoolLimitExceeded {
                actual,
                maximum: self.limits.max_bytes,
            }
            .into());
        }
        Ok(actual)
    }

    fn check_capacity(&self, additional: u64) -> Result<(), Error> {
        self.healthy()?;
        self.capacity_for(additional).map(|_| ())
    }

    fn record_len(file: &FileHeader, name: &[u8]) -> Result<u64, Error> {
        let size = ZipFileHeaderFixed::SIZE
            .checked_add(name.len())
            .and_then(|size| size.checked_add(usize::from(file.extra_fields.central_size)))
            .ok_or_else(|| ErrorKind::InvalidInput {
                msg: "central-directory record length overflow".to_string(),
            })?;
        usize_to_u64(size, "central-directory record length")
    }

    fn record_bytes(file: &FileHeader, name: &[u8]) -> Result<Vec<u8>, Error> {
        let capacity = usize::try_from(Self::record_len(file, name)?).map_err(|_| {
            ErrorKind::InvalidInput {
                msg: "central-directory record length does not fit usize".to_string(),
            }
        })?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|source| ErrorKind::Allocation {
                resource: "central-directory spool record",
                source,
            })?;
        file.central_fixed().write(&mut bytes)?;
        bytes.extend_from_slice(name);
        file.extra_fields
            .write_extra_fields(&mut bytes, Header::CENTRAL)?;
        Ok(bytes)
    }

    fn append_record(&mut self, file: &FileHeader, name: &[u8]) -> Result<(), Error> {
        self.healthy()?;
        let record_len = match Self::record_len(file, name) {
            Ok(length) => length,
            Err(error) => {
                self.poisoned = true;
                return Err(error);
            },
        };
        if let Err(error) = self.capacity_for(record_len) {
            self.poisoned = true;
            return Err(error);
        }
        let bytes = match Self::record_bytes(file, name) {
            Ok(bytes) => bytes,
            Err(error) => {
                self.poisoned = true;
                return Err(error);
            },
        };
        if let Err(error) = self.write_bytes(&bytes) {
            self.poisoned = true;
            return Err(error);
        }
        self.extent += record_len;
        self.entries = match self.entries.checked_add(1) {
            Some(entries) => entries,
            None => {
                self.poisoned = true;
                return Err(ErrorKind::InvalidInput {
                    msg: "central-directory entry count overflows u64".to_string(),
                }
                .into());
            },
        };
        self.saw_zip64 |= file.needs_zip64();
        Ok(())
    }

    fn write_bytes(&mut self, bytes: &[u8]) -> Result<(), Error> {
        let mut written = 0;
        while written < bytes.len() {
            let remaining = &bytes[written..];
            let count = loop {
                match self.store.write(remaining) {
                    Ok(count) => break count,
                    Err(source) if source.kind() == io::ErrorKind::Interrupted => continue,
                    Err(source) => {
                        return Err(ErrorKind::CentralDirectorySpool {
                            operation: "write central-directory record",
                            source,
                        }
                        .into());
                    },
                }
            };
            if count == 0 {
                return Err(ErrorKind::CentralDirectorySpool {
                    operation: "write central-directory record",
                    source: io::Error::new(
                        io::ErrorKind::WriteZero,
                        "central-directory spool accepted no bytes",
                    ),
                }
                .into());
            }
            if count > remaining.len() {
                return Err(ErrorKind::CentralDirectorySpool {
                    operation: "write central-directory record",
                    source: io::Error::new(
                        io::ErrorKind::InvalidData,
                        "central-directory spool returned more bytes than requested",
                    ),
                }
                .into());
            }
            written += count;
        }
        Ok(())
    }

    fn replay_into<W: Write>(&mut self, output: &mut W) -> Result<u64, Error> {
        self.healthy()?;
        loop {
            match self.store.flush() {
                Ok(()) => break,
                Err(source) if source.kind() == io::ErrorKind::Interrupted => continue,
                Err(source) => {
                    self.poisoned = true;
                    return Err(ErrorKind::CentralDirectorySpool {
                        operation: "flush before replay",
                        source,
                    }
                    .into());
                },
            }
        }
        loop {
            match self.store.seek(SeekFrom::Start(self.base_offset)) {
                Ok(offset) if offset == self.base_offset => break,
                Ok(offset) => {
                    self.poisoned = true;
                    return Err(ErrorKind::CentralDirectorySpool {
                        operation: "seek for replay",
                        source: io::Error::new(
                            io::ErrorKind::InvalidData,
                            format!(
                                "central-directory spool seek returned {offset}, requested {}",
                                self.base_offset
                            ),
                        ),
                    }
                    .into());
                },
                Err(source) if source.kind() == io::ErrorKind::Interrupted => continue,
                Err(source) => {
                    self.poisoned = true;
                    return Err(ErrorKind::CentralDirectorySpool {
                        operation: "seek for replay",
                        source,
                    }
                    .into());
                },
            }
        }

        let mut remaining = self.extent;
        while remaining != 0 {
            let requested = remaining.min(self.replay_buffer.len() as u64) as usize;
            let read = loop {
                match self.store.read(&mut self.replay_buffer[..requested]) {
                    Ok(read) => break read,
                    Err(source) if source.kind() == io::ErrorKind::Interrupted => continue,
                    Err(source) => {
                        self.poisoned = true;
                        return Err(ErrorKind::CentralDirectorySpool {
                            operation: "read for replay",
                            source,
                        }
                        .into());
                    },
                }
            };
            if read == 0 {
                self.poisoned = true;
                return Err(ErrorKind::CentralDirectorySpool {
                    operation: "read for replay",
                    source: io::Error::new(
                        io::ErrorKind::UnexpectedEof,
                        "central-directory spool ended before its admitted extent",
                    ),
                }
                .into());
            }
            if read > requested {
                self.poisoned = true;
                return Err(ErrorKind::CentralDirectorySpool {
                    operation: "read for replay",
                    source: io::Error::new(
                        io::ErrorKind::InvalidData,
                        "central-directory spool returned more bytes than requested",
                    ),
                }
                .into());
            }
            output.write_all(&self.replay_buffer[..read])?;
            remaining -= read as u64;
        }
        Ok(self.extent)
    }
}

#[derive(Debug)]
struct CountWriter<W> {
    writer: W,
    count: u64,
}

impl<W> CountWriter<W> {
    fn new(writer: W, count: u64) -> Self {
        CountWriter { writer, count }
    }

    fn count(&self) -> u64 {
        self.count
    }
}

impl<W: Write> Write for CountWriter<W> {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        let requested = u64::try_from(buf.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP output byte count does not fit in u64",
            )
        })?;
        if self.count.checked_add(requested).is_none() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP output offset overflows u64",
            ));
        }
        let bytes_written = self.writer.write(buf)?;
        if bytes_written > buf.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP output sink returned more bytes than requested",
            ));
        }
        let accepted = u64::try_from(bytes_written).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP output byte count does not fit in u64",
            )
        })?;
        self.count = self.count.checked_add(accepted).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP output offset overflows u64",
            )
        })?;
        Ok(bytes_written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }
}

/// Builds a `ZipArchiveWriter`.
#[derive(Debug, Default)]
pub struct ZipArchiveWriterBuilder {
    count: u64,
    capacity: usize,
}

impl ZipArchiveWriterBuilder {
    /// Creates a new `ZipArchiveWriterBuilder`.
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the anticipated number of files to optimize memory allocation.
    pub fn with_capacity(mut self, capacity: usize) -> Self {
        self.capacity = capacity;
        self
    }

    /// Sets the starting offset for writing. Useful when there is prelude data
    /// prior to the zip archive.
    ///
    /// When there is prelude data, setting the offset may not technically be
    /// required, but it is recommended. For standard zip files, many zip
    /// readers can self correct when the prelude data isn't properly declared.
    /// However for zip64 archives, setting the correct offset is required.
    ///
    /// # Example: Appending ZIP to existing data
    /// ```rust
    /// use std::io::{Cursor, Write, Seek, SeekFrom};
    ///
    /// // Create a file with some prefix data
    /// let mut output = Cursor::new(Vec::new());
    /// output.write_all(b"This is a custom header or prefix data\n").unwrap();
    /// let zip_start_offset = output.position();
    ///
    /// // Create ZIP archive starting after the prefix data
    /// let mut archive = soapberry_zip::ZipArchiveWriter::builder()
    ///     .with_offset(zip_start_offset)  // Tell the archive where it starts
    ///     .build(&mut output);
    ///
    /// // Add files normally
    /// let (mut file, config) = archive.new_file("data.txt").start().unwrap();
    /// let mut writer = config.wrap(&mut file);
    /// writer.write_all(b"File content").unwrap();
    /// let (_, desc) = writer.finish().unwrap();
    /// file.finish(desc).unwrap();
    /// archive.finish().unwrap();
    ///
    /// // The resulting file contains both prefix data and the ZIP archive
    /// let final_data = output.into_inner();
    /// assert!(final_data.starts_with(b"This is a custom header"));
    /// ```
    pub fn with_offset(mut self, offset: u64) -> Self {
        self.count = offset;
        self
    }

    /// Builds a `ZipArchiveWriter` that writes to `writer`.
    pub fn build<W>(&self, writer: W) -> ZipArchiveWriter<W> {
        ZipArchiveWriter {
            writer: CountWriter::new(writer, self.count),
            files: Vec::with_capacity(self.capacity),
            file_names: Vec::new(),
            reusable_deflate: None,
            directory_spool: None,
            pending_borrowed_entry: false,
            poisoned: false,
        }
    }

    /// Builds a `ZipArchiveWriter` with an explicit caller-provided store for
    /// finalized central-directory records.
    ///
    /// The store is positioned at its current end and the writer owns it for
    /// the lifetime of the archive. The store must remain exclusively owned by
    /// this writer for that lifetime; callers must not mutate or reposition it
    /// through another alias. It is never replaced with an ambient file or
    /// temporary path. `finish` replays only the bytes admitted by `limits`
    /// through its fixed replay buffer. The buffer is one part of the bounded
    /// working set: one active member name and one serialized central record
    /// are also retained, with their sizes bounded by ZIP field limits.
    pub fn build_with_spool<W, S>(
        &self,
        writer: W,
        spool: S,
        limits: DirectorySpoolLimits,
    ) -> Result<ZipArchiveWriter<W>, Error>
    where
        S: Read + Write + Seek + Send + Sync + 'static,
    {
        let directory_spool = Some(Box::new(DirectorySpool::new(spool, limits)?));
        Ok(ZipArchiveWriter {
            writer: CountWriter::new(writer, self.count),
            files: Vec::new(),
            file_names: Vec::new(),
            reusable_deflate: None,
            directory_spool,
            pending_borrowed_entry: false,
            poisoned: false,
        })
    }
}

/// Create a new Zip archive.
///
/// Basic usage:
/// ```rust
/// use std::io::Write;
///
/// let mut output = std::io::Cursor::new(Vec::new());
/// let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
/// let (mut entry, config) = archive.new_file("file.txt").start().unwrap();
/// let mut writer = config.wrap(&mut entry);
/// writer.write_all(b"Hello, world!").unwrap();
/// let (_, output) = writer.finish().unwrap();
/// entry.finish(output).unwrap();
/// archive.finish().unwrap();
/// ```
///
/// Use the builder for customization:
/// ```rust
/// use std::io::Write;
///
/// let mut output = std::io::Cursor::new(Vec::<u8>::new());
/// let mut _archive = soapberry_zip::ZipArchiveWriter::builder()
///     .with_capacity(1000)  // Optimize for 1000 anticipated files
///     .build(&mut output);
/// // ... add files as usual
/// ```
#[derive(Debug)]
pub struct ZipArchiveWriter<W> {
    files: Vec<FileHeader>,
    file_names: Vec<u8>,
    writer: CountWriter<W>,
    reusable_deflate: Option<Box<ReusableDeflateState>>,
    directory_spool: Option<Box<DirectorySpool>>,
    pending_borrowed_entry: bool,
    poisoned: bool,
}

struct SizedLocalHeader {
    fixed: ZipLocalFileHeaderFixed,
    extra: [u8; ZIP64_LOCAL_SIZE_EXTRA_MAX_LEN],
}

/// Sized member metadata shared by the ordinary writer and preservation's
/// direct-payload path.
///
/// The local header and central record carry the same grammar used by
/// [`ZipArchiveWriter::write_precompressed_file`].  The local-header offset is
/// supplied by the caller; preservation prepares it at zero and applies its
/// checked output-offset patch during layout preflight.
pub(crate) struct PreparedSizedMember<'a> {
    path: ZipFilePath<NormalizedPath<'a>>,
    local_header: SizedLocalHeader,
    file_header: FileHeader,
}

impl PreparedSizedMember<'_> {
    fn name_bytes(&self) -> &[u8] {
        self.path.as_ref().as_bytes()
    }

    fn local_len(&self) -> Result<usize, Error> {
        ZipLocalFileHeaderFixed::SIZE
            .checked_add(self.name_bytes().len())
            .and_then(|size| size.checked_add(usize::from(self.local_header.fixed.extra_field_len)))
            .ok_or_else(|| {
                ErrorKind::InvalidInput {
                    msg: "sized local header length overflow".to_string(),
                }
                .into()
            })
    }

    /// Materialize only the local framing; the compressed payload is supplied
    /// separately by the preservation writer.
    pub(crate) fn local_bytes(&self) -> Result<Vec<u8>, Error> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.local_len()?)
            .map_err(|source| ErrorKind::Allocation {
                resource: "generated local framing",
                source,
            })?;
        self.write_local(&mut bytes)?;
        Ok(bytes)
    }

    /// Materialize one central-directory record with the prepared local-header
    /// offset.  Preservation prepares the offset as zero and patches it only
    /// after all member spans have been measured.
    pub(crate) fn central_bytes(&mut self) -> Result<Vec<u8>, Error> {
        self.finalize_central()?;
        let capacity = ZipFileHeaderFixed::SIZE
            .checked_add(self.name_bytes().len())
            .and_then(|size| {
                size.checked_add(usize::from(self.file_header.extra_fields.central_size))
            })
            .ok_or_else(|| ErrorKind::InvalidInput {
                msg: "sized central record length overflow".to_string(),
            })?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|source| ErrorKind::Allocation {
                resource: "generated central framing",
                source,
            })?;
        self.file_header.central_fixed().write(&mut bytes)?;
        bytes.extend_from_slice(self.name_bytes());
        self.file_header
            .extra_fields
            .write_extra_fields(&mut bytes, Header::CENTRAL)?;
        Ok(bytes)
    }

    fn finalize_central(&mut self) -> Result<(), Error> {
        self.file_header.finalize_extra_fields()
    }

    fn write_local<W: Write>(&self, writer: &mut W) -> Result<(), Error> {
        self.local_header.fixed.write(&mut *writer)?;
        writer.write_all(self.name_bytes())?;
        let extra_len = usize::from(self.local_header.fixed.extra_field_len);
        writer.write_all(&self.local_header.extra[..extra_len])?;
        Ok(())
    }

    fn into_file_header(self) -> FileHeader {
        self.file_header
    }
}

fn sized_local_header(
    flags: u16,
    compression_method: CompressionMethod,
    crc32: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    file_name_len: u16,
) -> Result<SizedLocalHeader, Error> {
    let needs_uncompressed = uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE;
    let needs_compressed = compressed_size >= ZIP64_THRESHOLD_FILE_SIZE;
    let uses_zip64 = needs_uncompressed || needs_compressed;
    let mut extra = [0u8; ZIP64_LOCAL_SIZE_EXTRA_MAX_LEN];
    let (version_needed, compressed_size32, uncompressed_size32, extra_field_len) = if uses_zip64 {
        // A local ZIP64 size extra carries both sizes whenever either one
        // requires ZIP64. Central records may omit the non-sentinel value.
        extra[..2].copy_from_slice(&ExtraFieldId::ZIP64.as_u16().to_le_bytes());
        extra[2..4].copy_from_slice(&ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN.to_le_bytes());
        extra[4..12].copy_from_slice(&uncompressed_size.to_le_bytes());
        extra[12..20].copy_from_slice(&compressed_size.to_le_bytes());
        (
            ZIP64_VERSION_NEEDED,
            u32::MAX,
            u32::MAX,
            ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN,
        )
    } else {
        let compressed_size32 = u32::try_from(compressed_size).map_err(|_err| {
            Error::from(ErrorKind::InvalidInput {
                msg: "compressed payload length does not fit ZIP32".to_string(),
            })
        })?;
        let uncompressed_size32 = u32::try_from(uncompressed_size).map_err(|_err| {
            Error::from(ErrorKind::InvalidInput {
                msg: "uncompressed payload length does not fit ZIP32".to_string(),
            })
        })?;
        (20, compressed_size32, uncompressed_size32, 0)
    };

    Ok(SizedLocalHeader {
        fixed: ZipLocalFileHeaderFixed {
            signature: ZipLocalFileHeaderFixed::SIGNATURE,
            version_needed,
            flags,
            compression_method: compression_method.as_id(),
            last_mod_time: 0,
            last_mod_date: 0,
            crc32,
            compressed_size: compressed_size32,
            uncompressed_size: uncompressed_size32,
            file_name_len,
            extra_field_len,
        },
        extra,
    })
}

fn sized_member_path(name: &str) -> Result<ZipFilePath<NormalizedPath<'_>>, Error> {
    let path = ZipFilePath::from_str(name.trim_end_matches('/'));
    if path.len() > u16::MAX as usize {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "file name too long".to_string(),
        }));
    }
    Ok(path)
}

fn prepare_sized_member_with_path(
    path: ZipFilePath<NormalizedPath<'_>>,
    compression_method: CompressionMethod,
    crc32: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    local_header_offset: u64,
) -> Result<PreparedSizedMember<'_>, Error> {
    if !matches!(
        compression_method,
        CompressionMethod::Store | CompressionMethod::Deflate
    ) {
        return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
            compression_method.as_id().as_u16(),
        )));
    }

    let mut flags = 0u16;
    if path.needs_utf8_encoding() {
        flags |= FLAG_UTF8_ENCODING;
    }
    let name_len = path.len() as u16;
    let local_header = sized_local_header(
        flags,
        compression_method,
        crc32,
        compressed_size,
        uncompressed_size,
        name_len,
    )?;
    let file_header = FileHeader {
        name_len,
        compression_method,
        local_header_offset,
        compressed_size,
        uncompressed_size,
        crc: crc32,
        flags,
        zip64: false,
        modification_time: None,
        unix_permissions: None,
        extra_fields: ExtraFieldsContainer::new(),
    };
    Ok(PreparedSizedMember {
        path,
        local_header,
        file_header,
    })
}

pub(crate) fn prepare_sized_member(
    name: &str,
    compression_method: CompressionMethod,
    crc32: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    local_header_offset: u64,
) -> Result<PreparedSizedMember<'_>, Error> {
    let path = sized_member_path(name)?;
    prepare_sized_member_with_path(
        path,
        compression_method,
        crc32,
        compressed_size,
        uncompressed_size,
        local_header_offset,
    )
}

/// Prepare a stored member while preserving the writer's cheap name-refusal
/// ordering before scanning the payload for its CRC-32.
pub(crate) fn prepare_stored_member<'name>(
    name: &'name str,
    data: &[u8],
    local_header_offset: u64,
) -> Result<PreparedSizedMember<'name>, Error> {
    let path = sized_member_path(name)?;
    let size = usize_to_u64(data.len(), "stored payload length")?;
    let crc32 = crc::crc32(data);
    prepare_sized_member_with_path(
        path,
        CompressionMethod::Store,
        crc32,
        size,
        size,
        local_header_offset,
    )
}

impl ZipArchiveWriter<()> {
    /// Creates a `ZipArchiveWriterBuilder` for configuring the writer.
    pub fn builder() -> ZipArchiveWriterBuilder {
        ZipArchiveWriterBuilder::new()
    }
}

impl<W> ZipArchiveWriter<W> {
    /// Creates a new `ZipArchiveWriter` that writes to `writer`.
    pub fn new(writer: W) -> Self {
        ZipArchiveWriterBuilder::new().build(writer)
    }

    /// Creates a new writer whose finalized central-directory records are
    /// retained in the supplied explicit store. The store is consumed and
    /// remains exclusively owned by this writer; it must not be changed or
    /// repositioned through another alias until the archive is finished.
    pub fn new_with_spool<S>(
        writer: W,
        spool: S,
        limits: DirectorySpoolLimits,
    ) -> Result<Self, Error>
    where
        S: Read + Write + Seek + Send + Sync + 'static,
    {
        ZipArchiveWriterBuilder::new().build_with_spool(writer, spool, limits)
    }

    /// Returns the current offset in the output stream.
    ///
    /// Analagous to [`std::io::Cursor::position`].
    ///
    /// This can be used to determine various offsets during ZIP archive
    /// creation:
    ///
    /// - Local header offset
    /// - Start of compressed data offset
    /// - End of compressed data offset
    /// - End of data descriptor offset / next file's local header offset
    ///
    /// # Example
    ///
    /// ```rust
    /// use std::io::Write;
    ///
    /// let mut output = std::io::Cursor::new(Vec::new());
    /// let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
    ///
    /// // 1. Get local header offset
    /// let local_header_offset = archive.stream_offset();
    /// let (mut file, config) = archive.new_file("test.txt").start().unwrap();
    ///
    /// // 2. Get start of data offset
    /// let data_start_offset = file.stream_offset();
    ///
    /// // Write some data
    /// let mut writer = config.wrap(&mut file);
    /// writer.write_all(b"Hello World").unwrap();
    /// let (_, desc) = writer.finish().unwrap();
    ///
    /// // 3. Get end of compressed data offset
    /// let end_data_offset = file.stream_offset();
    ///
    /// let compressed_bytes = file.finish(desc).unwrap();
    ///
    /// // 4. Get end of data descriptor offset (next file's local header offset)
    /// let end_descriptor_offset = archive.stream_offset();
    ///
    /// archive.finish().unwrap();
    ///
    /// assert_eq!(local_header_offset, 0);
    /// assert!(data_start_offset > local_header_offset);
    /// assert_eq!(end_data_offset, data_start_offset + b"Hello World".len() as u64);
    /// assert_eq!(end_descriptor_offset, end_data_offset + 16); // 16 bytes for data descriptor
    /// assert_eq!(compressed_bytes, end_data_offset - data_start_offset);
    /// ```
    pub fn stream_offset(&self) -> u64 {
        self.writer.count()
    }

    /// Borrows the archive's reusable Deflate state, constructing it on first
    /// use.
    ///
    /// The state leaves the archive for the lifetime of one member so that an
    /// abandoned member drops an unfinished compressor instead of returning it.
    /// Callers must return it with [`Self::restore_reusable_deflate`] only
    /// after the member's final Deflate output succeeded.
    pub(crate) fn take_reusable_deflate(&mut self) -> Box<ReusableDeflateState> {
        let mut state = self
            .reusable_deflate
            .take()
            .unwrap_or_else(|| Box::new(ReusableDeflateState::new()));
        state.begin_member();
        state
    }

    /// Returns a finished Deflate state to the archive for the next member.
    pub(crate) fn restore_reusable_deflate(&mut self, state: Box<ReusableDeflateState>) {
        self.reusable_deflate = Some(state);
    }
}

/// Options for CRC32 calculation in ZIP files.
#[derive(Debug, Clone, Copy, Default)]
pub enum Crc32Option {
    /// Calculate CRC32 automatically from the data.
    #[default]
    Calculate,
    /// Use a custom CRC32 value and skip calculation.
    Custom(u32),
    /// Skip CRC32 calculation entirely (sets CRC32 to 0).
    Skip,
}

impl Crc32Option {
    /// Returns the initial CRC32 value for this option.
    #[inline]
    pub fn initial_value(&self) -> u32 {
        match self {
            Crc32Option::Calculate => 0,
            Crc32Option::Custom(value) => *value,
            Crc32Option::Skip => 0,
        }
    }
}

/// A builder for creating a new file entry in a ZIP archive.
#[derive(Debug)]
pub struct ZipFileBuilder<'archive, 'name, W> {
    archive: &'archive mut ZipArchiveWriter<W>,
    name: &'name str,
    compression_method: CompressionMethod,
    zip64: bool,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
    crc32_option: Crc32Option,
}

impl<'archive, W> ZipFileBuilder<'archive, '_, W>
where
    W: Write,
{
    /// Sets the compression method for the file entry.
    #[must_use]
    #[inline]
    pub fn compression_method(mut self, compression_method: CompressionMethod) -> Self {
        self.compression_method = compression_method;
        self
    }

    /// Selects ZIP64 framing for this streaming entry before its local header
    /// is emitted.
    ///
    /// The default is `false`, which preserves the compact ZIP32 streaming
    /// header and descriptor. Set this to `true` when the final sizes are not
    /// known in advance and the entry may require ZIP64. The opt-in mode uses
    /// a version-4.5 local header, a signed 64-bit data descriptor, and
    /// ZIP64 size fields in the central directory even when the final payload
    /// happens to fit in ZIP32.
    #[must_use]
    #[inline]
    pub fn zip64(mut self, zip64: bool) -> Self {
        self.zip64 = zip64;
        self
    }

    /// Sets the modification time for the file entry.
    ///
    /// Only accepts UTC timestamps to ensure Extended Timestamp fields are written correctly.
    #[must_use]
    #[inline]
    pub fn last_modified(mut self, modification_time: UtcDateTime) -> Self {
        self.modification_time = Some(modification_time);
        self
    }

    /// Sets the Unix permissions for the file entry.
    ///
    /// Accepts either:
    /// - Basic permission bits (e.g., 0o644 for rw-r--r--, 0o755 for rwxr-xr-x)
    /// - Full Unix mode including file type (e.g., 0o100644 for regular file, 0o040755 for directory)
    /// - Special permission bits are preserved (SUID: 0o4000, SGID: 0o2000, sticky: 0o1000)
    ///
    /// When set, the archive will be created with Unix-compatible "version made by" field
    /// to ensure proper interpretation of the permissions by zip readers.
    #[must_use]
    #[inline]
    pub fn unix_permissions(mut self, permissions: u32) -> Self {
        self.unix_permissions = Some(permissions);
        self
    }

    /// Adds an extra field to this file entry.
    ///
    /// Extra fields contain additional metadata about files in ZIP archives,
    /// such as timestamps, alignment information, and platform-specific data.
    ///
    /// No deduplication is performed - duplicate field IDs will result in
    /// multiple entries
    ///
    /// Will return an error if the total size exceeds 65,535 bytes for the
    /// specified headers.
    ///
    /// Rawzip will automatically add extra fields:
    ///
    /// - `EXTENDED_TIMESTAMP` when `last_modified()` is set
    /// - `ZIP64` when 32-bit thresholds are met
    ///
    /// # Examples
    ///
    /// Create files with different extra field headers and verify the
    /// behavior. Only the central directory is checked. To check the local
    /// extra fields, see
    /// [`ZipEntry::local_header`](crate::ZipEntry::local_header)
    ///
    /// ```rust
    /// # use std::io::{Cursor, Write};
    /// # use soapberry_zip::{ZipArchive, ZipArchiveWriter, ZipDataWriter, extra_fields::ExtraFieldId, Header};
    /// let mut output = Cursor::new(Vec::new());
    /// let mut archive = ZipArchiveWriter::new(&mut output);
    ///
    /// let my_custom_field = ExtraFieldId::new(0x6666);
    ///
    /// // File with extra fields only in the local file header
    /// let (mut local_file, local_config) = archive
    ///     .new_file("video.mp4")
    ///     .extra_field(my_custom_field, b"field1", Header::LOCAL)?
    ///     .start()?;
    /// let mut writer = local_config.wrap(&mut local_file);
    /// writer.write_all(b"video data")?;
    /// let (_, desc) = writer.finish()?;
    /// local_file.finish(desc)?;
    ///
    /// // File with extra fields only in the central directory
    /// let (mut central_file, central_config) = archive
    ///     .new_file("document.pdf")
    ///     .extra_field(my_custom_field, b"field2", Header::CENTRAL)?
    ///     .start()?;
    /// let mut writer = central_config.wrap(&mut central_file);
    /// writer.write_all(b"PDF content")?;
    /// let (_, desc) = writer.finish()?;
    /// central_file.finish(desc)?;
    ///
    /// // File with extra fields in both headers for maximum compatibility
    /// assert_eq!(Header::default(), Header::LOCAL | Header::CENTRAL);
    /// let (mut both_file, both_config) = archive
    ///     .new_file("important.dat")
    ///     .extra_field(my_custom_field, b"field3", Header::default())?
    ///     .start()?;
    /// let mut writer = both_config.wrap(&mut both_file);
    /// writer.write_all(b"important data")?;
    /// let (_, desc) = writer.finish()?;
    /// both_file.finish(desc)?;
    ///
    /// archive.finish()?;
    ///
    /// // Verify the behavior when reading back the central directory
    /// let zip_data = output.into_inner();
    /// let archive = ZipArchive::from_slice(&zip_data)?;
    ///
    /// for entry_result in archive.entries() {
    ///     let entry = entry_result?;
    ///     
    ///     // Find our custom field in the central directory
    ///     let custom_field_data = entry.extra_fields()
    ///         .find(|(id, _)| *id == my_custom_field)
    ///         .map(|(_, data)| data);
    ///     
    ///     match entry.file_path().as_ref() {
    ///         b"video.mp4" => {
    ///             // local only field should not be in central directory
    ///             assert_eq!(custom_field_data, None);
    ///         }
    ///         b"document.pdf" => {
    ///             // central only field should be in central directory
    ///             assert_eq!(custom_field_data, Some(b"field2".as_slice()));
    ///         }
    ///         b"important.dat" => {
    ///             // both location field should be in central directory
    ///             assert_eq!(custom_field_data, Some(b"field3".as_slice()));
    ///         }
    ///         _ => {}
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn extra_field(
        mut self,
        id: ExtraFieldId,
        data: &[u8],
        location: Header,
    ) -> Result<Self, Error> {
        self.extra_fields.add_field(id, data, location)?;
        Ok(self)
    }

    /// Sets the CRC32 calculation option for the file entry.
    ///
    /// By default, CRC32 is calculated automatically from the data. Use this
    /// method to:
    ///
    /// - Skip CRC32 calculation entirely (for performance or when verification
    ///   isn't desired)
    /// - Provide a pre-calculated CRC32 value
    #[must_use]
    #[inline]
    pub fn crc32(mut self, crc32_option: Crc32Option) -> Self {
        self.crc32_option = crc32_option;
        self
    }

    /// Creates the file entry and returns a writer for the file's content.
    #[deprecated(
        since = "0.4.0",
        note = "Use `start()` method instead as it allows for more flexibility (ie: CRC configuration)"
    )]
    pub fn create(self) -> Result<ZipEntryWriter<'archive, W>, Error> {
        let (entry_writer, _) = self.start()?;
        Ok(entry_writer)
    }

    /// Mark the start of file data
    ///
    /// Returns a tuple:
    ///
    /// - `entry` handles the ZIP format and writes compressed data to the archive
    /// - `config` constructs data writers that handle uncompressed data and CRC32 calculation
    ///
    /// # Examples
    ///
    /// For stored (uncompressed) files:
    /// ```
    /// # use std::io::Write;
    /// # let mut output = std::io::Cursor::new(Vec::new());
    /// # let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
    /// let (mut entry, config) = archive.new_file("file.txt").start().unwrap();
    /// let mut writer = config.wrap(&mut entry);
    /// writer.write_all(b"Hello").unwrap();
    /// let (_, output) = writer.finish().unwrap();
    /// entry.finish(output).unwrap();
    /// # archive.finish().unwrap();
    /// ```
    ///
    /// For deflate compression:
    /// ```
    /// # use std::io::Write;
    /// # let mut output = std::io::Cursor::new(Vec::new());
    /// # let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
    /// let (mut entry, config) = archive.new_file("file.txt").start().unwrap();
    /// let encoder = flate2::write::DeflateEncoder::new(&mut entry, flate2::Compression::default());
    /// let mut writer = config.wrap(encoder);
    /// writer.write_all(b"Hello").unwrap();
    /// let (encoder, output) = writer.finish().unwrap();
    /// encoder.finish().unwrap();
    /// entry.finish(output).unwrap();
    /// # archive.finish().unwrap();
    /// ```
    pub fn start(self) -> Result<(ZipEntryWriter<'archive, W>, ZipDataWriterConfig), Error> {
        let crc32_option = self.crc32_option;
        let options = ZipEntryOptions {
            compression_method: self.compression_method,
            zip64: self.zip64,
            modification_time: self.modification_time,
            unix_permissions: self.unix_permissions,
            extra_fields: self.extra_fields,
        };
        let entry_writer = self.archive.new_file_with_options(self.name, options)?;

        let data_writer_config = ZipDataWriterConfig { crc32_option };

        Ok((entry_writer, data_writer_config))
    }
}

/// A builder for creating a new directory entry in a ZIP archive.
#[derive(Debug)]
pub struct ZipDirBuilder<'a, W> {
    archive: &'a mut ZipArchiveWriter<W>,
    name: &'a str,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

impl<W> ZipDirBuilder<'_, W>
where
    W: Write,
{
    /// Sets the modification time for the directory entry.
    ///
    /// See [`ZipFileBuilder::last_modified`] for details.
    #[must_use]
    #[inline]
    pub fn last_modified(mut self, modification_time: UtcDateTime) -> Self {
        self.modification_time = Some(modification_time);
        self
    }

    /// Sets the Unix permissions for the directory entry.
    ///
    /// See [`ZipFileBuilder::unix_permissions`] for details.
    #[must_use]
    #[inline]
    pub fn unix_permissions(mut self, permissions: u32) -> Self {
        self.unix_permissions = Some(permissions);
        self
    }

    /// Adds an extra field to this directory entry.
    ///
    /// See [`ZipFileBuilder::extra_field`] for details and examples.
    /// The same behavior notes apply: append-only, no deduplication, and automatic fields.
    pub fn extra_field(
        mut self,
        id: ExtraFieldId,
        data: &[u8],
        location: Header,
    ) -> Result<Self, Error> {
        self.extra_fields.add_field(id, data, location)?;
        Ok(self)
    }

    /// Creates the directory entry.
    pub fn create(self) -> Result<(), Error> {
        let options = ZipEntryOptions {
            compression_method: CompressionMethod::Store, // Directories always use Store
            zip64: false,
            modification_time: self.modification_time,
            unix_permissions: self.unix_permissions,
            extra_fields: self.extra_fields,
        };
        self.archive.new_dir_with_options(self.name, options)
    }
}

impl<W> ZipArchiveWriter<W>
where
    W: Write,
{
    fn ensure_usable(&self) -> Result<(), Error> {
        if self.poisoned {
            return Err(ErrorKind::InvalidInput {
                msg: "ZIP archive is poisoned after a prior failure".to_string(),
            }
            .into());
        }
        Ok(())
    }

    fn ensure_no_pending_borrowed_entry(&self) -> Result<(), Error> {
        if self.pending_borrowed_entry {
            return Err(ErrorKind::InvalidInput {
                msg: "ZIP archive has an unfinished borrowed entry".to_string(),
            }
            .into());
        }
        Ok(())
    }

    fn has_directory_spool(&self) -> bool {
        self.directory_spool.is_some()
    }

    fn check_directory_spool_capacity(
        &self,
        name_len: usize,
        central_extra_len: u16,
    ) -> Result<(), Error> {
        let Some(spool) = self.directory_spool.as_ref() else {
            return Ok(());
        };
        let record_len = ZipFileHeaderFixed::SIZE
            .checked_add(name_len)
            .and_then(|size| size.checked_add(usize::from(central_extra_len)))
            .ok_or_else(|| ErrorKind::InvalidInput {
                msg: "central-directory record length overflow".to_string(),
            })?;
        spool.check_capacity(usize_to_u64(record_len, "central-directory record length")?)
    }

    fn check_directory_spool_file(&self, file: &FileHeader, name: &[u8]) -> Result<(), Error> {
        let Some(spool) = self.directory_spool.as_ref() else {
            return Ok(());
        };
        spool.check_capacity(DirectorySpool::record_len(file, name)?)
    }

    fn directory_spool_extra_len(
        options: &ZipEntryOptions,
        local_header_offset: u64,
    ) -> Result<u16, Error> {
        let automatic_timestamp = if options.modification_time.is_some() {
            4usize + 5
        } else {
            0
        };
        let automatic_zip64 = if options.zip64 {
            4usize
                + 16
                + if local_header_offset >= ZIP64_THRESHOLD_OFFSET {
                    8
                } else {
                    0
                }
        } else if local_header_offset >= ZIP64_THRESHOLD_OFFSET {
            4usize + 8
        } else {
            0
        };
        let size = usize::from(options.extra_fields.central_size)
            .checked_add(automatic_timestamp)
            .and_then(|size| size.checked_add(automatic_zip64))
            .ok_or_else(|| ErrorKind::InvalidInput {
                msg: "central-directory extra-field length overflow".to_string(),
            })?;
        u16::try_from(size).map_err(|_| {
            ErrorKind::InvalidInput {
                msg: "central-directory extra fields exceed ZIP limits".to_string(),
            }
            .into()
        })
    }

    fn owned_spool_name(&self, name: &[u8]) -> Result<Option<Vec<u8>>, Error> {
        if !self.has_directory_spool() {
            return Ok(None);
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(name.len())
            .map_err(|source| ErrorKind::Allocation {
                resource: "active central-directory member name",
                source,
            })?;
        owned.extend_from_slice(name);
        Ok(Some(owned))
    }

    fn publish_file_header(&mut self, name: &[u8], file: FileHeader) -> Result<(), Error> {
        if let Some(spool) = self.directory_spool.as_mut() {
            match spool.append_record(&file, name) {
                Ok(()) => Ok(()),
                Err(error) => {
                    self.poisoned = true;
                    Err(error)
                },
            }
        } else {
            self.files.push(file);
            Ok(())
        }
    }

    fn poison_directory_spool(&mut self) {
        self.poisoned = true;
        if let Some(spool) = self.directory_spool.as_mut() {
            spool.poisoned = true;
        }
    }

    fn publish_prepared_member(&mut self, prepared: PreparedSizedMember<'_>) -> Result<(), Error> {
        if let Some(spool) = self.directory_spool.as_mut() {
            let PreparedSizedMember {
                path, file_header, ..
            } = prepared;
            match spool.append_record(&file_header, path.as_ref().as_bytes()) {
                Ok(()) => Ok(()),
                Err(error) => {
                    self.poisoned = true;
                    Err(error)
                },
            }
        } else {
            self.files.push(prepared.into_file_header());
            Ok(())
        }
    }

    fn reserve_member_metadata(&mut self, name_bytes: &[u8]) -> Result<(), Error> {
        if self.has_directory_spool() {
            return Ok(());
        }
        self.file_names
            .try_reserve(name_bytes.len())
            .map_err(|error| ErrorKind::InvalidInput {
                msg: format!("could not reserve ZIP member-name storage: {error}"),
            })?;
        self.files
            .try_reserve(1)
            .map_err(|error| ErrorKind::InvalidInput {
                msg: format!("could not reserve ZIP file-header storage: {error}"),
            })?;
        self.file_names.extend_from_slice(name_bytes);
        Ok(())
    }

    pub fn write_stored_file(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_stored_file_with_accounting(name, data, &mut accounting)
    }

    /// Write a stored file while recording payload bytes accepted by the
    /// archive sink. ZIP framing bytes are excluded from the counter.
    pub fn write_stored_file_with_accounting(
        &mut self,
        name: &str,
        data: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        let path = sized_member_path(name)?;
        let crc32 = crc::crc32(data);
        let size_u64 = usize_to_u64(data.len(), "stored payload length")?;
        let mut prepared = prepare_sized_member_with_path(
            path,
            CompressionMethod::Store,
            crc32,
            size_u64,
            size_u64,
            self.writer.count(),
        )?;
        prepared.finalize_central()?;
        self.check_directory_spool_file(&prepared.file_header, prepared.name_bytes())?;
        self.reserve_member_metadata(prepared.name_bytes())?;
        if let Err(error) = prepared.write_local(&mut self.writer) {
            self.poison_directory_spool();
            return Err(error);
        }
        if let Err(error) = write_all_counted(
            &mut self.writer,
            data,
            accounting,
            AccountingWriteKind::Stored,
        ) {
            self.poison_directory_spool();
            return Err(error);
        }
        self.publish_prepared_member(prepared)?;

        Ok(())
    }

    /// Writes one member whose compressed bytes, CRC-32, and sizes are already
    /// known, declaring the final values in the local header.
    ///
    /// This is the sized counterpart of the streaming entry API: because the
    /// local header carries the real CRC-32 and both sizes, no data descriptor
    /// is emitted and the resulting archive can be probed from its central
    /// directory alone. `crc32` and `uncompressed_size` must describe the
    /// plaintext that `compressed` inflates to under `compression_method`;
    /// only `Store` and `Deflate` are accepted.
    pub fn write_precompressed_file(
        &mut self,
        name: &str,
        compression_method: CompressionMethod,
        crc32: u32,
        uncompressed_size: u64,
        compressed: &[u8],
    ) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_precompressed_file_with_accounting(
            name,
            compression_method,
            crc32,
            uncompressed_size,
            compressed,
            &mut accounting,
        )
    }

    /// Write a precompressed member while recording payload bytes accepted by
    /// the archive sink. ZIP framing bytes are excluded from the counter.
    pub fn write_precompressed_file_with_accounting(
        &mut self,
        name: &str,
        compression_method: CompressionMethod,
        crc32: u32,
        uncompressed_size: u64,
        compressed: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        let accounting_kind = match compression_method {
            CompressionMethod::Store => AccountingWriteKind::Stored,
            CompressionMethod::Deflate => AccountingWriteKind::Precompressed,
            _ => AccountingWriteKind::Precompressed,
        };
        self.write_precompressed_file_classified(
            name,
            compression_method,
            crc32,
            uncompressed_size,
            compressed,
            accounting,
            accounting_kind,
        )
    }

    pub(crate) fn write_generated_deflate_file_with_accounting(
        &mut self,
        name: &str,
        crc32: u32,
        uncompressed_size: u64,
        compressed: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.write_precompressed_file_classified(
            name,
            CompressionMethod::Deflate,
            crc32,
            uncompressed_size,
            compressed,
            accounting,
            AccountingWriteKind::GeneratedDeflate,
        )
    }

    fn write_precompressed_file_classified(
        &mut self,
        name: &str,
        compression_method: CompressionMethod,
        crc32: u32,
        uncompressed_size: u64,
        compressed: &[u8],
        accounting: &mut ZipOperationAccounting,
        accounting_kind: AccountingWriteKind,
    ) -> Result<(), Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                compression_method.as_id().as_u16(),
            )));
        }
        let compressed_size = usize_to_u64(compressed.len(), "precompressed payload length")?;
        let mut prepared = prepare_sized_member(
            name,
            compression_method,
            crc32,
            compressed_size,
            uncompressed_size,
            self.writer.count(),
        )?;
        prepared.finalize_central()?;
        self.check_directory_spool_file(&prepared.file_header, prepared.name_bytes())?;
        self.reserve_member_metadata(prepared.name_bytes())?;
        if let Err(error) = prepared.write_local(&mut self.writer) {
            self.poison_directory_spool();
            return Err(error);
        }
        if let Err(error) =
            write_all_counted(&mut self.writer, compressed, accounting, accounting_kind)
        {
            self.poison_directory_spool();
            return Err(error);
        }
        self.publish_prepared_member(prepared)?;

        Ok(())
    }

    /// Writes a local file header with filtered extra fields.
    fn write_local_header(
        &mut self,
        file_path: &ZipFilePath<NormalizedPath>,
        flags: u16,
        compression_method: CompressionMethod,
        options: &mut ZipEntryOptions,
    ) -> Result<(), Error> {
        if options.zip64
            && options
                .extra_fields
                .contains_id(ExtraFieldId::ZIP64, Header::default())
        {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "ZIP64 streaming mode reserves its own ZIP64 extra field".to_string(),
            }));
        }

        // Get DOS timestamp from options or use 0 as default
        let (dos_time, dos_date) = options
            .modification_time
            .as_ref()
            .map(|dt| DosDateTime::from(dt).into_parts())
            .unwrap_or((0, 0));

        if let Some(datetime) = options.modification_time.as_ref() {
            let unix_time = datetime.to_unix().max(0) as u32;
            let mut data = [0u8; 5];
            data[0] = 1; // Flags: modification time present
            data[1..].copy_from_slice(&unix_time.to_le_bytes());
            options.extra_fields.add_field(
                ExtraFieldId::EXTENDED_TIMESTAMP,
                &data,
                Header::CENTRAL,
            )?;
        }

        // A ZIP64 streaming local header reserves both size values in the
        // ZIP64 extra field.  The values are intentionally zero until the
        // descriptor is emitted; the central record carries the resolved
        // values.  Selecting this layout before writing the header keeps the
        // descriptor width fixed for the lifetime of the entry.
        if options.zip64 {
            options.extra_fields.add_field(
                ExtraFieldId::ZIP64,
                &[0u8; ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN as usize],
                Header::LOCAL,
            )?;
        }

        let header = ZipLocalFileHeaderFixed {
            signature: ZipLocalFileHeaderFixed::SIGNATURE,
            version_needed: if options.zip64 {
                ZIP64_VERSION_NEEDED
            } else {
                20
            },
            flags,
            compression_method: compression_method.as_id(),
            last_mod_time: dos_time,
            last_mod_date: dos_date,
            crc32: 0, // must be zero if data descriptor is used (4.4.4)
            compressed_size: if options.zip64 { u32::MAX } else { 0 },
            uncompressed_size: if options.zip64 { u32::MAX } else { 0 },
            file_name_len: file_path.len() as u16,
            extra_field_len: options.extra_fields.local_size,
        };

        header.write(&mut self.writer)?;
        self.writer.write_all(file_path.as_ref().as_bytes())?;
        options
            .extra_fields
            .write_extra_fields(&mut self.writer, Header::LOCAL)?;
        Ok(())
    }

    /// Creates a builder for adding a new directory to the archive.
    ///
    /// The name of the directory must end with a `/`.
    ///
    /// # Example
    ///
    /// ```rust
    /// # use std::io::Cursor;
    /// # let mut output = Cursor::new(Vec::new());
    /// # let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
    /// archive.new_dir("my-dir/")
    ///     .unix_permissions(0o755)
    ///     .create()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn new_dir<'a>(&'a mut self, name: &'a str) -> ZipDirBuilder<'a, W> {
        ZipDirBuilder {
            archive: self,
            name,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
        }
    }

    /// Adds a new directory to the archive with options (internal method).
    ///
    /// The name of the directory must end with a `/`.
    fn new_dir_with_options(
        &mut self,
        name: &str,
        mut options: ZipEntryOptions,
    ) -> Result<(), Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        let file_path = ZipFilePath::from_str(name);
        if !file_path.is_dir() {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "not a directory".to_string(),
            }));
        }

        if file_path.len() > u16::MAX as usize {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "directory name too long".to_string(),
            }));
        }

        let local_header_offset = self.writer.count();
        if local_header_offset >= ZIP64_THRESHOLD_OFFSET
            && options
                .extra_fields
                .contains_id(ExtraFieldId::ZIP64, Header::CENTRAL)
        {
            return Err(ErrorKind::InvalidInput {
                msg: "ZIP64 extra field collides with automatic central metadata".to_string(),
            }
            .into());
        }
        let mut flags = 0u16;
        if file_path.needs_utf8_encoding() {
            flags |= FLAG_UTF8_ENCODING;
        } else {
            flags &= !FLAG_UTF8_ENCODING;
        }

        // Store the name bytes in the central buffer
        let name_bytes = file_path.as_ref().as_bytes();
        let name_len = name_bytes.len() as u16;
        let spool_extra_len = Self::directory_spool_extra_len(&options, local_header_offset)?;
        self.check_directory_spool_capacity(name_bytes.len(), spool_extra_len)?;
        self.reserve_member_metadata(name_bytes)?;

        if let Err(error) =
            self.write_local_header(&file_path, flags, CompressionMethod::Store, &mut options)
        {
            self.poison_directory_spool();
            return Err(error);
        }

        let file_header = FileHeader {
            name_len,
            compression_method: CompressionMethod::Store,
            local_header_offset,
            compressed_size: 0,
            uncompressed_size: 0,
            crc: 0,
            flags,
            zip64: false,
            modification_time: options.modification_time,
            unix_permissions: options.unix_permissions,
            extra_fields: options.extra_fields,
        };
        let mut file_header = file_header;
        if let Err(error) = file_header.finalize_extra_fields() {
            self.poison_directory_spool();
            return Err(error);
        }
        if let Err(error) = self.check_directory_spool_file(&file_header, name_bytes) {
            self.poison_directory_spool();
            return Err(error);
        }
        self.publish_file_header(name_bytes, file_header)?;

        Ok(())
    }

    /// Creates a builder for adding a new file to the archive.
    ///
    /// # Example
    ///
    /// ```rust
    /// # use std::io::{Cursor, Write};
    /// # let mut output = Cursor::new(Vec::new());
    /// # let mut archive = soapberry_zip::ZipArchiveWriter::new(&mut output);
    /// let (mut entry, config) = archive.new_file("my-file")
    ///     .compression_method(soapberry_zip::CompressionMethod::Deflate)
    ///     .unix_permissions(0o644)
    ///     .start()?;
    /// let mut writer = config.wrap(&mut entry);
    /// writer.write_all(b"Hello, world!")?;
    /// let (_, output) = writer.finish()?;
    /// entry.finish(output)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn new_file<'name>(&mut self, name: &'name str) -> ZipFileBuilder<'_, 'name, W> {
        ZipFileBuilder {
            archive: self,
            name,
            compression_method: CompressionMethod::Store,
            zip64: false,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
            crc32_option: Crc32Option::default(),
        }
    }

    /// Adds a new file to the archive with options (internal method).
    fn new_file_with_options(
        &mut self,
        name: &str,
        mut options: ZipEntryOptions,
    ) -> Result<ZipEntryWriter<'_, W>, Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        let file_path = ZipFilePath::from_str(name.trim_end_matches('/'));

        if file_path.len() > u16::MAX as usize {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "file name too long".to_string(),
            }));
        }

        let local_header_offset = self.writer.count();
        let mut flags = FLAG_DATA_DESCRIPTOR;
        if file_path.needs_utf8_encoding() {
            flags |= FLAG_UTF8_ENCODING;
        } else {
            flags &= !FLAG_UTF8_ENCODING;
        }

        // Store the name bytes in the central buffer
        let name_bytes = file_path.as_ref().as_bytes();
        let name_len = name_bytes.len() as u16;
        let spool_extra_len = Self::directory_spool_extra_len(&options, local_header_offset)?;
        self.check_directory_spool_capacity(name_bytes.len(), spool_extra_len)?;
        let spool_name = self.owned_spool_name(name_bytes)?;
        self.reserve_member_metadata(name_bytes)?;

        if let Err(error) =
            self.write_local_header(&file_path, flags, options.compression_method, &mut options)
        {
            self.poison_directory_spool();
            return Err(error);
        }
        self.pending_borrowed_entry = true;

        Ok(ZipEntryWriter {
            inner: self,
            compressed_bytes: 0,
            name_len,
            local_header_offset,
            compression_method: options.compression_method,
            zip64: options.zip64,
            flags,
            modification_time: options.modification_time,
            unix_permissions: options.unix_permissions,
            extra_fields: options.extra_fields,
            name: spool_name,
        })
    }

    /// Starts a file entry while consuming this archive writer.
    ///
    /// Unlike [`ZipFileBuilder::start`], this API does not borrow the archive
    /// while the entry is being written. The returned entry owns the archive
    /// state and returns it from [`ZipOwnedEntryWriter::finish`]. This makes it
    /// possible to pass an entry writer through a streaming pipeline without
    /// constructing a self-referential wrapper around `ZipArchiveWriter`.
    ///
    /// The entry writer accepts uncompressed bytes through [`Write`]. `Store`
    /// forwards those bytes directly and `Deflate` compresses them
    /// incrementally. Other compression methods are rejected before the local
    /// header is emitted.
    pub fn start_file_owned(
        self,
        name: &str,
        compression_method: CompressionMethod,
    ) -> Result<ZipOwnedEntryWriter<W>, Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                compression_method.as_id().as_u16(),
            )));
        }

        let options = ZipEntryOptions {
            compression_method,
            zip64: false,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
        };
        self.start_file_owned_with_options(name, options)
    }

    /// Alias for [`Self::start_file_owned`] using the entry-oriented name used
    /// by higher-level package writers.
    pub fn start_entry_owned(
        self,
        name: &str,
        compression_method: CompressionMethod,
    ) -> Result<ZipOwnedEntryWriter<W>, Error> {
        self.start_file_owned(name, compression_method)
    }

    /// Starts an owned streaming entry with explicit ZIP64 framing.
    ///
    /// This is the consuming counterpart of [`ZipFileBuilder::zip64`]. The
    /// ZIP64 local header and 64-bit descriptor are selected before any entry
    /// bytes are accepted, so the entry remains valid when its final sizes are
    /// not known to the caller.
    pub fn start_file_owned_zip64(
        self,
        name: &str,
        compression_method: CompressionMethod,
    ) -> Result<ZipOwnedEntryWriter<W>, Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                compression_method.as_id().as_u16(),
            )));
        }

        let options = ZipEntryOptions {
            compression_method,
            zip64: true,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
        };
        self.start_file_owned_with_options(name, options)
    }

    /// Alias for [`Self::start_file_owned_zip64`] using the entry-oriented
    /// name used by higher-level package writers.
    pub fn start_entry_owned_zip64(
        self,
        name: &str,
        compression_method: CompressionMethod,
    ) -> Result<ZipOwnedEntryWriter<W>, Error> {
        self.start_file_owned_zip64(name, compression_method)
    }

    fn start_file_owned_with_options(
        mut self,
        name: &str,
        mut options: ZipEntryOptions,
    ) -> Result<ZipOwnedEntryWriter<W>, Error> {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        let file_path = ZipFilePath::from_str(name.trim_end_matches('/'));
        if file_path.len() > u16::MAX as usize {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "file name too long".to_string(),
            }));
        }

        let local_header_offset = self.writer.count();
        let mut flags = FLAG_DATA_DESCRIPTOR;
        if file_path.needs_utf8_encoding() {
            flags |= FLAG_UTF8_ENCODING;
        }

        let name_bytes = file_path.as_ref().as_bytes();
        let name_len = name_bytes.len() as u16;
        let spool_extra_len = Self::directory_spool_extra_len(&options, local_header_offset)?;
        self.check_directory_spool_capacity(name_bytes.len(), spool_extra_len)?;
        let spool_name = self.owned_spool_name(name_bytes)?;
        self.reserve_member_metadata(name_bytes)?;
        if let Err(error) =
            self.write_local_header(&file_path, flags, options.compression_method, &mut options)
        {
            self.poison_directory_spool();
            return Err(error);
        }

        let reusable_deflate = if options.compression_method == CompressionMethod::Deflate {
            Some(self.take_reusable_deflate())
        } else {
            None
        };

        let state = OwnedEntryState {
            name: spool_name,
            name_len,
            local_header_offset,
            compression_method: options.compression_method,
            zip64: options.zip64,
            flags,
            modification_time: options.modification_time,
            unix_permissions: options.unix_permissions,
            extra_fields: options.extra_fields,
        };
        let compressed = OwnedCompressedEntry {
            archive: self,
            state,
            compressed_bytes: 0,
            compressed_limit: None,
        };
        let compressor = match options.compression_method {
            CompressionMethod::Store => OwnedCompressor::Store(compressed),
            CompressionMethod::Deflate => {
                let state = match reusable_deflate {
                    Some(state) => state,
                    None => {
                        return Err(ErrorKind::InvalidInput {
                            msg: "owned Deflate state was not prepared".to_string(),
                        }
                        .into());
                    },
                };
                OwnedCompressor::Deflate(OwnedDeflateCompressor {
                    state,
                    entry: compressed,
                })
            },
            // `start_file_owned` checks this before calling this helper. Keep
            // the internal helper defensive for future callers with custom
            // options.
            other => {
                return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                    other.as_id().as_u16(),
                )));
            },
        };

        Ok(ZipOwnedEntryWriter {
            inner: Some(ZipDataWriter::with_crc32(
                compressor,
                Crc32Option::default(),
            )),
        })
    }

    /// Finishes writing the archive and returns the underlying writer.
    ///
    /// This writes the central directory and the end of central directory
    /// record. ZIP64 format is used automatically when thresholds are exceeded.
    pub fn finish(mut self) -> Result<W, Error>
    where
        W: Write,
    {
        self.ensure_usable()?;
        self.ensure_no_pending_borrowed_entry()?;
        let central_directory_offset = self.writer.count();
        let (total_entries, needs_zip64_before_central_directory, central_directory_size) =
            if let Some(mut spool) = self.directory_spool.take() {
                let total_entries = spool.entries;
                let needs_zip64_before_central_directory = total_entries
                    >= ZIP64_THRESHOLD_ENTRIES as u64
                    || central_directory_offset >= ZIP64_THRESHOLD_OFFSET
                    || spool.saw_zip64;
                let central_directory_size = spool.replay_into(&mut self.writer)?;
                (
                    total_entries,
                    needs_zip64_before_central_directory,
                    central_directory_size,
                )
            } else {
                let total_entries = self.files.len() as u64;
                let needs_zip64_before_central_directory = total_entries
                    >= ZIP64_THRESHOLD_ENTRIES as u64
                    || central_directory_offset >= ZIP64_THRESHOLD_OFFSET
                    || self.files.iter().any(|f| f.needs_zip64());

                let mut name_offset = 0;

                // Write central directory entries
                for file in &self.files {
                    let header = file.central_fixed();

                    header.write(&mut self.writer)?;

                    // File name
                    let new_name_offset = name_offset + file.name_len as usize;
                    self.writer
                        .write_all(&self.file_names[name_offset..new_name_offset])?;
                    name_offset = new_name_offset;

                    // Extra fields
                    file.extra_fields
                        .write_extra_fields(&mut self.writer, Header::CENTRAL)?;
                }

                let central_directory_end = self.writer.count();
                let central_directory_size = central_directory_end - central_directory_offset;
                (
                    total_entries,
                    needs_zip64_before_central_directory,
                    central_directory_size,
                )
            };
        let needs_zip64 = needs_zip64_before_central_directory
            || central_directory_size >= ZIP64_THRESHOLD_OFFSET;

        // Write ZIP64 structures if needed
        if needs_zip64 {
            let zip64_eocd_offset = self.writer.count();

            // Write ZIP64 End of Central Directory Record
            write_zip64_eocd(
                &mut self.writer,
                total_entries,
                central_directory_size,
                central_directory_offset,
            )?;

            // Write ZIP64 End of Central Directory Locator
            write_zip64_eocd_locator(&mut self.writer, zip64_eocd_offset)?;
        }

        // Write regular End of Central Directory Record
        self.writer.write_all(&END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES)?;

        // Disk numbers
        self.writer.write_all(&[0u8; 4])?;

        // A ZIP64 tail must be discoverable through a classic EOCD sentinel,
        // even when only an entry's size forced ZIP64 and the directory-level
        // count, size, and offset still fit in ZIP32.
        let entries_count = if needs_zip64 {
            u16::MAX
        } else {
            total_entries.min(ZIP64_THRESHOLD_ENTRIES as u64) as u16
        };
        self.writer.write_all(&entries_count.to_le_bytes())?;
        self.writer.write_all(&entries_count.to_le_bytes())?;

        // Central directory size - use 0xFFFFFFFF if ZIP64
        let cd_size = central_directory_size.min(ZIP64_THRESHOLD_OFFSET) as u32;
        self.writer.write_all(&cd_size.to_le_bytes())?;

        // Central directory offset - use 0xFFFFFFFF if ZIP64
        let cd_offset = central_directory_offset.min(ZIP64_THRESHOLD_OFFSET) as u32;
        self.writer.write_all(&cd_offset.to_le_bytes())?;

        // Comment length
        self.writer.write_all(&0u16.to_le_bytes())?;

        self.writer.flush()?;
        Ok(self.writer.writer)
    }
}

/// A writer for a file in a ZIP archive.
///
/// This writer is created by `ZipArchiveWriter::new_file`.
/// Data written to this writer is compressed and written to the underlying archive.
///
/// After writing all data, call `finish` to complete the entry.
#[derive(Debug)]
pub struct ZipEntryWriter<'a, W> {
    inner: &'a mut ZipArchiveWriter<W>,
    compressed_bytes: u64,
    name: Option<Vec<u8>>,
    name_len: u16,
    local_header_offset: u64,
    compression_method: CompressionMethod,
    zip64: bool,
    flags: u16,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

/// Configuration for creating data writers that handle uncompressed data and CRC32 calculation.
#[derive(Debug)]
pub struct ZipDataWriterConfig {
    crc32_option: Crc32Option,
}

impl ZipDataWriterConfig {
    /// Wraps an encoder with a data writer configured with this builder's options.
    pub fn wrap<E>(self, encoder: E) -> ZipDataWriter<E> {
        ZipDataWriter::with_crc32(encoder, self.crc32_option)
    }
}

impl<'a, W> ZipEntryWriter<'a, W> {
    /// Returns the total number of bytes successfully written (bytes out).
    pub fn compressed_bytes(&self) -> u64 {
        self.compressed_bytes
    }

    /// Returns the current offset in the output stream.
    ///
    /// See [`ZipArchiveWriter::stream_offset`] for more information.
    pub fn stream_offset(&self) -> u64 {
        self.inner.stream_offset()
    }

    /// Finishes writing the file entry.
    ///
    /// This writes the data descriptor if necessary and adds the file entry to the central directory.
    pub fn finish(self, mut output: DataDescriptorOutput) -> Result<u64, Error>
    where
        W: Write,
    {
        self.inner.ensure_usable()?;
        output.compressed_size = self.compressed_bytes;
        let mut file_header = FileHeader {
            name_len: self.name_len,
            compression_method: self.compression_method,
            local_header_offset: self.local_header_offset,
            compressed_size: output.compressed_size,
            uncompressed_size: output.uncompressed_size,
            crc: output.crc,
            flags: self.flags,
            zip64: self.zip64,
            modification_time: self.modification_time,
            unix_permissions: self.unix_permissions,
            extra_fields: self.extra_fields,
        };
        if let Err(error) = file_header.finalize_extra_fields() {
            if !zip_error_is_interrupted(&error) {
                self.inner.poison_directory_spool();
            }
            return Err(error);
        }
        if let Err(error) = write_data_descriptor(
            &mut self.inner.writer,
            output.crc,
            output.compressed_size,
            output.uncompressed_size,
            self.zip64,
        ) {
            if !zip_error_is_interrupted(&error) {
                self.inner.poison_directory_spool();
            }
            return Err(error);
        }
        let name = self.name.unwrap_or_default();
        self.inner.publish_file_header(&name, file_header)?;
        self.inner.pending_borrowed_entry = false;

        Ok(self.compressed_bytes)
    }
}

impl<W> Write for ZipEntryWriter<'_, W>
where
    W: Write,
{
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        let requested = match u64::try_from(buf.len()) {
            Ok(requested) => requested,
            Err(_) => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "ZIP compressed byte count does not fit in u64",
                );
                self.inner.poison_directory_spool();
                return Err(error);
            },
        };
        if self.compressed_bytes.checked_add(requested).is_none() {
            let error = io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP compressed byte count overflows u64",
            );
            self.inner.poison_directory_spool();
            return Err(error);
        }
        let bytes_written = match self.inner.writer.write(buf) {
            Ok(bytes_written) => bytes_written,
            Err(error) => {
                if !io_error_is_interrupted(&error) {
                    self.inner.poison_directory_spool();
                }
                return Err(error);
            },
        };
        if bytes_written > buf.len() {
            self.inner.poison_directory_spool();
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP compressed sink returned more bytes than requested",
            ));
        }
        if !buf.is_empty() && bytes_written == 0 {
            self.inner.poison_directory_spool();
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "ZIP compressed sink accepted no bytes",
            ));
        }
        let accepted = match u64::try_from(bytes_written) {
            Ok(accepted) => accepted,
            Err(_) => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidData,
                    "ZIP compressed byte count does not fit in u64",
                );
                self.inner.poison_directory_spool();
                return Err(error);
            },
        };
        self.compressed_bytes = match self.compressed_bytes.checked_add(accepted) {
            Some(total) => total,
            None => {
                let error = io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "ZIP compressed byte count overflows u64",
                );
                self.inner.poison_directory_spool();
                return Err(error);
            },
        };
        Ok(bytes_written)
    }

    fn flush(&mut self) -> io::Result<()> {
        match self.inner.writer.flush() {
            Ok(()) => Ok(()),
            Err(error) => {
                if !io_error_is_interrupted(&error) {
                    self.inner.poison_directory_spool();
                }
                Err(error)
            },
        }
    }
}

/// Internal state retained by an owned entry while its payload is emitted.
#[derive(Debug)]
struct OwnedEntryState {
    name: Option<Vec<u8>>,
    name_len: u16,
    local_header_offset: u64,
    compression_method: CompressionMethod,
    zip64: bool,
    flags: u16,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

/// A compressed payload sink that owns the archive it is writing into.
#[derive(Debug)]
struct OwnedCompressedEntry<W> {
    archive: ZipArchiveWriter<W>,
    state: OwnedEntryState,
    compressed_bytes: u64,
    compressed_limit: Option<u64>,
}

impl<W> OwnedCompressedEntry<W> {
    fn compressed_bytes(&self) -> u64 {
        self.compressed_bytes
    }

    fn set_compressed_limit(&mut self, maximum: u64) {
        self.compressed_limit = Some(maximum);
    }

    fn finish(self, mut output: DataDescriptorOutput) -> Result<ZipArchiveWriter<W>, Error>
    where
        W: Write,
    {
        output.compressed_size = self.compressed_bytes;
        let mut archive = self.archive;
        archive.ensure_usable()?;
        let mut file_header = FileHeader {
            name_len: self.state.name_len,
            compression_method: self.state.compression_method,
            local_header_offset: self.state.local_header_offset,
            compressed_size: output.compressed_size,
            uncompressed_size: output.uncompressed_size,
            crc: output.crc,
            flags: self.state.flags,
            zip64: self.state.zip64,
            modification_time: self.state.modification_time,
            unix_permissions: self.state.unix_permissions,
            extra_fields: self.state.extra_fields,
        };
        if let Err(error) = file_header.finalize_extra_fields() {
            if !zip_error_is_interrupted(&error) {
                archive.poison_directory_spool();
            }
            return Err(error);
        }
        if let Err(error) = write_data_descriptor(
            &mut archive.writer,
            output.crc,
            output.compressed_size,
            output.uncompressed_size,
            self.state.zip64,
        ) {
            if !zip_error_is_interrupted(&error) {
                archive.poison_directory_spool();
            }
            return Err(error);
        }
        let name = self.state.name.unwrap_or_default();
        archive.publish_file_header(&name, file_header)?;

        Ok(archive)
    }
}

/// The output capacity used by flate2's allocating write adapter.
///
/// `DeflateEncoder` uses a `Vec` with this initial capacity.  Keeping the same
/// capacity here preserves its input/output boundaries while avoiding one
/// fresh vector allocation for every member.
const REUSABLE_DEFLATE_OUTPUT_BUFFER_SIZE: usize = 32 * 1024;

/// Reusable state for one raw Deflate stream at a time.
///
/// The state is held by its owner — the archive writer between successfully
/// finalized owned entries, the streaming Office writer and the preservation
/// writer between members — so that one save constructs one compressor
/// instead of one per member.  An active entry owns it, so dropping an
/// unfinished entry drops the compressor instead of returning a partially
/// finished stream to its owner.
///
/// Reuse is byte-transparent: [`Compress::reset`] restores the same level,
/// strategy and window the constructor selects, and the pending-output
/// boundaries below reproduce flate2's `zio::Writer` call sequence exactly, so
/// a reused stream emits the same bytes a fresh [`DeflateEncoder`] would.
#[derive(Debug)]
pub(crate) struct ReusableDeflateState {
    compressor: Compress,
    // Heap-resident so the struct itself stays a few words wide: it is moved
    // into a `Box` on construction, and a 32 KiB inline array would be
    // memset on the stack and then copied into that box on every save.
    output: Box<[u8]>,
    pending_start: usize,
    pending_end: usize,
    /// Whether a member has already been fed through this state.
    ///
    /// The reset is paid when the *next* member starts, not when the previous
    /// one finishes, so a save with a single Deflate member pays exactly the
    /// one construction it paid before this state existed and no reset at all.
    used: bool,
}

impl ReusableDeflateState {
    pub(crate) fn new() -> Self {
        Self {
            compressor: Compress::new(Compression::default(), false),
            output: vec![0; REUSABLE_DEFLATE_OUTPUT_BUFFER_SIZE].into_boxed_slice(),
            pending_start: 0,
            pending_end: 0,
            used: false,
        }
    }

    /// Readies the state for one member, resetting the compressor only when a
    /// previous member left a stream in it.
    ///
    /// `Compress::reset` restores the level, strategy and window the
    /// constructor selected, so the member that follows emits exactly the
    /// bytes a freshly constructed encoder would.
    pub(crate) fn begin_member(&mut self) {
        if self.used {
            self.compressor.reset();
            self.pending_start = 0;
            self.pending_end = 0;
        }
        self.used = true;
    }

    fn compress_once(
        &mut self,
        input: &[u8],
        flush: FlushCompress,
    ) -> io::Result<(Status, usize, usize)> {
        let before_in = self.compressor.total_in();
        let before_out = self.compressor.total_out();
        let status = self
            .compressor
            .compress(input, &mut self.output[self.pending_end..], flush)
            .map_err(|_| reusable_deflate_error())?;
        let consumed = self
            .compressor
            .total_in()
            .checked_sub(before_in)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(reusable_deflate_progress_error)?;
        let produced = self
            .compressor
            .total_out()
            .checked_sub(before_out)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(reusable_deflate_progress_error)?;
        let available = self.output.len() - self.pending_end;
        if consumed > input.len() || produced > available {
            return Err(reusable_deflate_progress_error());
        }
        self.pending_end += produced;
        Ok((status, consumed, produced))
    }

    fn compact_pending(&mut self) {
        if self.pending_start != 0 {
            let remaining = self.pending_end - self.pending_start;
            self.output
                .copy_within(self.pending_start..self.pending_end, 0);
            self.pending_start = 0;
            self.pending_end = remaining;
        }
    }

    fn drain_pending<W: Write + ?Sized>(&mut self, entry: &mut W) -> io::Result<()> {
        while self.pending_start < self.pending_end {
            let start = self.pending_start;
            let end = self.pending_end;
            let pending_len = end - start;
            match entry.write(&self.output[start..end]) {
                Ok(0) => {
                    self.compact_pending();
                    return Err(io::ErrorKind::WriteZero.into());
                },
                Ok(written) => {
                    if written > pending_len {
                        self.compact_pending();
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "ZIP compressed sink returned more bytes than requested",
                        ));
                    }
                    self.pending_start += written;
                },
                Err(error) => {
                    self.compact_pending();
                    return Err(error);
                },
            }
        }
        self.pending_start = 0;
        self.pending_end = 0;
        Ok(())
    }

    pub(crate) fn write_to<W: Write + ?Sized>(
        &mut self,
        entry: &mut W,
        input: &[u8],
    ) -> io::Result<usize> {
        // Keep the same write boundary as flate2's zio writer: one codec
        // call per `Write::write`, draining output left by the preceding
        // codec call before retrying a zero-consumption call. Newly produced
        // bytes remain pending until the next drain, as they do in zio.
        loop {
            self.drain_pending(entry)?;
            let (status, consumed, produced) = self.compress_once(input, FlushCompress::None)?;
            if !input.is_empty() && consumed == 0 && status != Status::StreamEnd {
                if produced == 0 {
                    return Err(reusable_deflate_progress_error());
                }
                continue;
            }
            return Ok(consumed);
        }
    }

    pub(crate) fn flush_to<W: Write + ?Sized>(&mut self, entry: &mut W) -> io::Result<()> {
        // zio runs the initial sync flush before dumping output already
        // buffered by the preceding write. This intentionally uses only the
        // remaining scratch capacity, then drains the combined range.
        let (_, consumed, _) = self.compress_once(&[], FlushCompress::Sync)?;
        if consumed != 0 {
            return Err(reusable_deflate_progress_error());
        }

        // zio::Writer drains any bytes left by the sync flush with no-flush
        // calls before forwarding `flush` to its wrapped writer.
        loop {
            self.drain_pending(entry)?;
            let before_out = self.compressor.total_out();
            let (_, consumed, _) = self.compress_once(&[], FlushCompress::None)?;
            if consumed != 0 {
                return Err(reusable_deflate_progress_error());
            }
            if self.compressor.total_out() == before_out {
                break;
            }
        }
        entry.flush()
    }

    pub(crate) fn finish_to<W: Write + ?Sized>(&mut self, entry: &mut W) -> io::Result<()> {
        loop {
            self.drain_pending(entry)?;
            let before_out = self.compressor.total_out();
            let (status, consumed, _) = self.compress_once(&[], FlushCompress::Finish)?;
            if consumed != 0 {
                return Err(reusable_deflate_progress_error());
            }
            if status == Status::StreamEnd {
                self.drain_pending(entry)?;
                return Ok(());
            }
            if self.compressor.total_out() == before_out {
                return Err(reusable_deflate_progress_error());
            }
        }
    }
}

fn reusable_deflate_error() -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, "corrupt deflate stream")
}

fn reusable_deflate_progress_error() -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidData,
        "owned Deflate compressor made no valid progress",
    )
}

/// A raw Deflate encoder that borrows a [`ReusableDeflateState`] instead of
/// constructing its own compressor.
///
/// This is a drop-in replacement for `flate2::write::DeflateEncoder` at the
/// Office writer sites that compress one whole member at a time: it keeps the
/// same `Write` boundaries and the same `finish` semantics, but the compressor
/// and its 32 KiB output buffer belong to the caller and outlive the member.
/// The caller readies the state with [`ReusableDeflateState::begin_member`];
/// a member that fails mid-stream leaves the state unfinished, so its owner
/// discards it rather than handing it to the next member.
pub(crate) struct ReusedDeflateEncoder<'state, W: Write> {
    state: &'state mut ReusableDeflateState,
    // `None` only after a successful `finish`, exactly as flate2's
    // `zio::Writer` clears its object after `take_inner`.  `Drop` below reads
    // it to decide whether the stream still needs a final block.
    sink: Option<W>,
}

impl<'state, W: Write> ReusedDeflateEncoder<'state, W> {
    pub(crate) fn new(state: &'state mut ReusableDeflateState, sink: W) -> Self {
        Self {
            state,
            sink: Some(sink),
        }
    }

    /// Emits the final Deflate block and returns the wrapped sink, matching
    /// `DeflateEncoder::finish`.
    pub(crate) fn finish(mut self) -> io::Result<W> {
        match self.sink.as_mut() {
            Some(sink) => self.state.finish_to(sink)?,
            None => return Err(reused_deflate_finished_error()),
        }
        match self.sink.take() {
            Some(sink) => Ok(sink),
            None => Err(reused_deflate_finished_error()),
        }
    }
}

impl<W: Write> Write for ReusedDeflateEncoder<'_, W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        match self.sink.as_mut() {
            Some(sink) => self.state.write_to(sink, buffer),
            None => Err(reused_deflate_finished_error()),
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        match self.sink.as_mut() {
            Some(sink) => self.state.flush_to(sink),
            None => Err(reused_deflate_finished_error()),
        }
    }
}

impl<W: Write> Drop for ReusedDeflateEncoder<'_, W> {
    fn drop(&mut self) {
        // flate2's `zio::Writer` finishes an unfinished stream on drop and
        // discards the result. An entry abandoned mid-member keeps that
        // behaviour so a failed member emits the same bytes it did before,
        // and its owner discards the borrowed state rather than reusing it.
        if let Some(sink) = self.sink.as_mut() {
            let _ = self.state.finish_to(sink);
        }
    }
}

fn reused_deflate_finished_error() -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidInput,
        "reused Deflate encoder was already finished",
    )
}

#[derive(Debug)]
struct OwnedDeflateCompressor<W: Write> {
    state: Box<ReusableDeflateState>,
    entry: OwnedCompressedEntry<W>,
}

#[derive(Debug)]
enum OwnedCompressor<W: Write> {
    Store(OwnedCompressedEntry<W>),
    Deflate(OwnedDeflateCompressor<W>),
}

impl<W: Write> OwnedCompressor<W> {
    fn compressed_bytes(&self) -> u64 {
        match self {
            Self::Store(entry) => entry.compressed_bytes(),
            Self::Deflate(compressor) => compressor.entry.compressed_bytes(),
        }
    }

    fn set_compressed_limit(&mut self, maximum: u64) {
        match self {
            Self::Store(entry) => entry.set_compressed_limit(maximum),
            Self::Deflate(compressor) => compressor.entry.set_compressed_limit(maximum),
        }
    }

    fn finish(self, output: DataDescriptorOutput) -> Result<ZipArchiveWriter<W>, Error>
    where
        W: Write,
    {
        match self {
            Self::Store(entry) => entry.finish(output),
            Self::Deflate(mut compressor) => {
                compressor.state.finish_to(&mut compressor.entry)?;
                let mut archive = compressor.entry.finish(output)?;
                archive.restore_reusable_deflate(compressor.state);
                Ok(archive)
            },
        }
    }
}

impl<W: Write> Write for OwnedCompressor<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        match self {
            Self::Store(entry) => entry.write(buffer),
            Self::Deflate(compressor) => {
                let result = compressor.state.write_to(&mut compressor.entry, buffer);
                if let Err(error) = &result {
                    if !io_error_is_interrupted(error) {
                        compressor.entry.archive.poison_directory_spool();
                    }
                }
                result
            },
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        match self {
            Self::Store(entry) => entry.flush(),
            Self::Deflate(compressor) => {
                let result = compressor.state.flush_to(&mut compressor.entry);
                if let Err(error) = &result {
                    if !io_error_is_interrupted(error) {
                        compressor.entry.archive.poison_directory_spool();
                    }
                }
                result
            },
        }
    }
}

#[derive(Debug)]
struct OwnedEntryLimitMarker {
    actual: u64,
    maximum: u64,
}

impl std::fmt::Display for OwnedEntryLimitMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "owned ZIP entry compressed limit exceeded: attempted {}, maximum {}",
            self.actual, self.maximum
        )
    }
}

impl std::error::Error for OwnedEntryLimitMarker {}

fn owned_entry_limit_io_error(actual: u64, maximum: u64) -> io::Error {
    io::Error::new(
        io::ErrorKind::Other,
        OwnedEntryLimitMarker { actual, maximum },
    )
}

pub(crate) fn owned_entry_limit_from_io_error(error: &io::Error) -> Option<(u64, u64)> {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<OwnedEntryLimitMarker>())
        .map(|marker| (marker.actual, marker.maximum))
}

/// Writes a signed data descriptor with the width selected before payload
/// emission. A streaming entry that was not explicitly opted into ZIP64 must
/// refuse a size that reaches the ZIP64 sentinel boundary rather than emit a
/// descriptor whose width disagrees with its local header.
fn write_data_descriptor<W: Write>(
    writer: &mut W,
    crc: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    zip64: bool,
) -> Result<(), Error> {
    let sizes_need_zip64 = compressed_size >= ZIP64_THRESHOLD_FILE_SIZE
        || uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE;
    if sizes_need_zip64 && !zip64 {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "streaming entry reached ZIP64 size without zip64(true)".to_string(),
        }));
    }

    let mut buffer = [0u8; 24];
    buffer[0..4].copy_from_slice(&DataDescriptor::SIGNATURE.to_le_bytes());
    buffer[4..8].copy_from_slice(&crc.to_le_bytes());
    if zip64 || sizes_need_zip64 {
        buffer[8..16].copy_from_slice(&compressed_size.to_le_bytes());
        buffer[16..24].copy_from_slice(&uncompressed_size.to_le_bytes());
        writer.write_all(&buffer)?;
    } else {
        buffer[8..12].copy_from_slice(&(compressed_size as u32).to_le_bytes());
        buffer[12..16].copy_from_slice(&(uncompressed_size as u32).to_le_bytes());
        writer.write_all(&buffer[..16])?;
    }
    Ok(())
}

impl<W: Write> Write for OwnedCompressedEntry<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if buffer.is_empty() {
            return Ok(0);
        }
        let requested = u64::try_from(buffer.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP compressed byte count does not fit in u64",
            )
        })?;
        let attempted = self
            .compressed_bytes
            .checked_add(requested)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "ZIP compressed byte count overflows u64",
                )
            })?;
        if let Some(maximum) = self.compressed_limit {
            if attempted > maximum {
                return Err(owned_entry_limit_io_error(attempted, maximum));
            }
        }
        let written = match self.archive.writer.write(buffer) {
            Ok(written) => written,
            Err(error) => {
                if !io_error_is_interrupted(&error) {
                    self.archive.poison_directory_spool();
                }
                return Err(error);
            },
        };
        if written > buffer.len() {
            self.archive.poison_directory_spool();
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP compressed sink returned more bytes than requested",
            ));
        }
        if written == 0 {
            self.archive.poison_directory_spool();
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "ZIP compressed sink accepted no bytes",
            ));
        }
        let accepted = u64::try_from(written).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP compressed byte count does not fit in u64",
            )
        })?;
        self.compressed_bytes = self.compressed_bytes.checked_add(accepted).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP compressed byte count overflows u64",
            )
        })?;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        match self.archive.writer.flush() {
            Ok(()) => Ok(()),
            Err(error) => {
                if !io_error_is_interrupted(&error) {
                    self.archive.poison_directory_spool();
                }
                Err(error)
            },
        }
    }
}

/// A consuming ZIP entry writer that owns its parent archive.
///
/// The entry implements [`Write`] for uncompressed payload bytes. Store data
/// is forwarded directly; Deflate data is compressed incrementally with a
/// bounded working buffer. Calling [`Self::finish`] consumes the entry and
/// returns the archive writer so another entry can be started without a
/// borrow tied to the original archive value.
#[derive(Debug)]
pub struct ZipOwnedEntryWriter<W: Write> {
    inner: Option<ZipDataWriter<OwnedCompressor<W>>>,
}

impl<W: Write> ZipOwnedEntryWriter<W> {
    /// Sets a compressed-payload ceiling enforced before bytes reach the
    /// archive sink.
    #[must_use]
    pub fn with_compressed_limit(mut self, maximum: u64) -> Self {
        if let Some(inner) = self.inner.as_mut() {
            inner.inner.set_compressed_limit(maximum);
        }
        self
    }

    /// Number of uncompressed bytes accepted by this entry.
    #[must_use]
    pub fn uncompressed_bytes(&self) -> u64 {
        self.inner
            .as_ref()
            .map(|inner| inner.uncompressed_bytes)
            .unwrap_or(0)
    }

    /// Number of compressed payload bytes accepted by this entry.
    #[must_use]
    pub fn compressed_bytes(&self) -> u64 {
        self.inner
            .as_ref()
            .map(|inner| inner.inner.compressed_bytes())
            .unwrap_or(0)
    }

    /// Finishes the entry and recovers the parent archive writer.
    pub fn finish(mut self) -> Result<ZipArchiveWriter<W>, Error>
    where
        W: Write,
    {
        let inner = self.inner.take().ok_or_else(|| ErrorKind::InvalidInput {
            msg: "owned ZIP entry writer was already finished".to_string(),
        })?;
        let (compressor, descriptor) = inner.finish()?;
        compressor.finish(descriptor)
    }
}

impl<W: Write> Write for ZipOwnedEntryWriter<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        self.inner
            .as_mut()
            .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "owned ZIP entry finished"))?
            .write(buffer)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner
            .as_mut()
            .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "owned ZIP entry finished"))?
            .flush()
    }
}

/// A writer for the uncompressed data of a Zip file entry.
///
/// This writer will keep track of the data necessary to write the data
/// descriptor (ie: number of bytes written and the CRC32 checksum).
///
/// Once all the data has been written, invoke the `finish` method to receive the
/// `DataDescriptorOutput` necessary to finalize the entry.
#[derive(Debug)]
pub struct ZipDataWriter<W> {
    inner: W,
    uncompressed_bytes: u64,
    crc: u32,
    crc32_option: Crc32Option,
}

impl<W> ZipDataWriter<W> {
    /// Creates a new `ZipDataWriter` that writes to an underlying writer.
    #[deprecated(
        since = "0.4.0",
        note = "Use the tuple-based API: `ZipFileBuilder::start()` returns `(writer, builder)` which can propagate the CRC32 option"
    )]
    pub fn new(inner: W) -> Self {
        Self::with_crc32_option(inner, Crc32Option::default())
    }

    /// Creates a new `ZipDataWriter` with the specified CRC32 option.
    ///
    /// This is an internal method. Use the tuple-based API via
    /// `ZipFileBuilder::start()` instead.
    pub(crate) fn with_crc32(inner: W, crc32_option: Crc32Option) -> Self {
        Self::with_crc32_option(inner, crc32_option)
    }

    /// Creates a new `ZipDataWriter` with a specific CRC32 calculation option.
    fn with_crc32_option(inner: W, crc32_option: Crc32Option) -> Self {
        let crc = crc32_option.initial_value();
        ZipDataWriter {
            inner,
            uncompressed_bytes: 0,
            crc,
            crc32_option,
        }
    }

    /// Gets a mutable reference to the underlying writer.
    pub fn get_mut(&mut self) -> &mut W {
        &mut self.inner
    }

    /// Consumes self and returns the inner writer and the data descriptor to be
    /// passed to a `ZipEntryWriter`.
    ///
    /// The writer is returned to facilitate situations where the underlying
    /// compressor needs to be notified that no more data will be written so it
    /// can write any sort of necessary epilogue (think zstd).
    ///
    /// The `DataDescriptorOutput` contains the CRC32 checksum and uncompressed size,
    /// which is needed by `ZipEntryWriter::finish`.
    pub fn finish(mut self) -> Result<(W, DataDescriptorOutput), Error>
    where
        W: Write,
    {
        self.flush()?;
        let output = DataDescriptorOutput {
            crc: self.crc,
            compressed_size: 0,
            uncompressed_size: self.uncompressed_bytes,
        };

        Ok((self.inner, output))
    }
}

impl<W> Write for ZipDataWriter<W>
where
    W: Write,
{
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        let requested = u64::try_from(buf.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP uncompressed byte count does not fit in u64",
            )
        })?;
        if self.uncompressed_bytes.checked_add(requested).is_none() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "ZIP uncompressed byte count overflows u64",
            ));
        }
        let bytes_written = self.inner.write(buf)?;
        if bytes_written > buf.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP uncompressed sink returned more bytes than requested",
            ));
        }
        let accepted = u64::try_from(bytes_written).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP uncompressed byte count does not fit in u64",
            )
        })?;
        self.uncompressed_bytes =
            self.uncompressed_bytes
                .checked_add(accepted)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "ZIP uncompressed byte count overflows u64",
                    )
                })?;

        // Only calculate CRC32 if the option is Calculate
        if matches!(self.crc32_option, Crc32Option::Calculate) {
            self.crc = crc::crc32_chunk(&buf[..bytes_written], self.crc);
        }

        Ok(bytes_written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

/// Contains information written in the data descriptor after the file data.
#[derive(Debug, Clone)]
pub struct DataDescriptorOutput {
    crc: u32,
    compressed_size: u64,
    uncompressed_size: u64,
}

impl DataDescriptorOutput {
    /// Returns the CRC32 checksum of the uncompressed data.
    pub fn crc(&self) -> u32 {
        self.crc
    }

    /// Returns the uncompressed size of the data.
    pub fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }
}

#[derive(Debug)]
struct FileHeader {
    name_len: u16,
    compression_method: CompressionMethod,
    local_header_offset: u64,
    compressed_size: u64,
    uncompressed_size: u64,
    crc: u32,
    flags: u16,
    zip64: bool,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

impl FileHeader {
    fn central_fixed(&self) -> ZipFileHeaderFixed {
        // Version made by and version needed to extract.
        let version_needed = if self.needs_zip64() {
            ZIP64_VERSION_NEEDED
        } else {
            20
        };

        // Set version_made_by to indicate Unix when Unix permissions are present.
        let version_made_by_hi = self.unix_permissions.map(|_| CREATOR_UNIX).unwrap_or(0);
        let version_made_by = (version_made_by_hi << 8) | version_needed;

        let (dos_time, dos_date) = self
            .modification_time
            .as_ref()
            .map(|dt| DosDateTime::from(dt).into_parts())
            .unwrap_or((0, 0));

        ZipFileHeaderFixed {
            signature: CENTRAL_HEADER_SIGNATURE,
            version_made_by,
            version_needed,
            flags: self.flags,
            compression_method: self.compression_method.as_id(),
            last_mod_time: dos_time,
            last_mod_date: dos_date,
            crc32: self.crc,
            compressed_size: if self.zip64 {
                u32::MAX
            } else {
                self.compressed_size.min(ZIP64_THRESHOLD_FILE_SIZE) as u32
            },
            uncompressed_size: if self.zip64 {
                u32::MAX
            } else {
                self.uncompressed_size.min(ZIP64_THRESHOLD_FILE_SIZE) as u32
            },
            file_name_len: self.name_len,
            extra_field_len: self.extra_fields.central_size,
            file_comment_len: 0,
            disk_number_start: 0,
            internal_file_attrs: 0,
            external_file_attrs: self.unix_permissions.map(|x| x << 16).unwrap_or(0),
            local_header_offset: self.local_header_offset.min(ZIP64_THRESHOLD_OFFSET) as u32,
        }
    }

    fn needs_zip64(&self) -> bool {
        self.zip64
            || self.compressed_size >= ZIP64_THRESHOLD_FILE_SIZE
            || self.uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE
            || self.local_header_offset >= ZIP64_THRESHOLD_OFFSET
    }

    fn finalize_extra_fields(&mut self) -> Result<(), Error> {
        if self.needs_zip64() {
            if self
                .extra_fields
                .contains_id(ExtraFieldId::ZIP64, Header::CENTRAL)
            {
                return Err(Error::from(ErrorKind::InvalidInput {
                    msg: "ZIP64 extra field collides with automatic central metadata".to_string(),
                }));
            }
            let mut sink = [0u8; 24];
            let mut pos = 0;
            if self.zip64 || self.uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE {
                sink[pos..pos + 8].copy_from_slice(&self.uncompressed_size.to_le_bytes());
                pos += 8;
            }
            if self.zip64 || self.compressed_size >= ZIP64_THRESHOLD_FILE_SIZE {
                sink[pos..pos + 8].copy_from_slice(&self.compressed_size.to_le_bytes());
                pos += 8;
            }
            if self.local_header_offset >= ZIP64_THRESHOLD_OFFSET {
                sink[pos..pos + 8].copy_from_slice(&self.local_header_offset.to_le_bytes());
                pos += 8;
            }
            self.extra_fields
                .add_field(ExtraFieldId::ZIP64, &sink[..pos], Header::CENTRAL)?;
        }

        Ok(())
    }
}

/// Writes the ZIP64 End of Central Directory Record
fn write_zip64_eocd<W>(
    writer: &mut W,
    total_entries: u64,
    central_directory_size: u64,
    central_directory_offset: u64,
) -> Result<(), Error>
where
    W: Write,
{
    // ZIP64 End of Central Directory Record signature
    writer.write_all(&END_OF_CENTRAL_DIR_SIGNATURE64.to_le_bytes())?;

    // Size of ZIP64 end of central directory record (excluding signature and this field)
    let record_size = (ZIP64_EOCD_SIZE - 12) as u64;
    writer.write_all(&record_size.to_le_bytes())?;

    // Version made by
    writer.write_all(&ZIP64_VERSION_NEEDED.to_le_bytes())?;

    // Version needed to extract
    writer.write_all(&ZIP64_VERSION_NEEDED.to_le_bytes())?;

    // Number of this disk
    writer.write_all(&0u32.to_le_bytes())?;

    // Number of the disk with the start of the central directory
    writer.write_all(&0u32.to_le_bytes())?;

    // Total number of entries in the central directory on this disk
    writer.write_all(&total_entries.to_le_bytes())?;

    // Total number of entries in the central directory
    writer.write_all(&total_entries.to_le_bytes())?;

    // Size of the central directory
    writer.write_all(&central_directory_size.to_le_bytes())?;

    // Offset of start of central directory with respect to the starting disk number
    writer.write_all(&central_directory_offset.to_le_bytes())?;

    Ok(())
}

/// Writes the ZIP64 End of Central Directory Locator
fn write_zip64_eocd_locator<W>(writer: &mut W, zip64_eocd_offset: u64) -> Result<(), Error>
where
    W: Write,
{
    // ZIP64 End of Central Directory Locator signature
    writer.write_all(&END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE.to_le_bytes())?;

    // Number of the disk with the start of the ZIP64 end of central directory
    writer.write_all(&0u32.to_le_bytes())?;

    // Relative offset of the ZIP64 end of central directory record
    writer.write_all(&zip64_eocd_offset.to_le_bytes())?;

    // Total number of disks
    writer.write_all(&1u32.to_le_bytes())?;

    Ok(())
}

#[derive(Debug, Clone)]
struct ZipEntryOptions {
    compression_method: CompressionMethod,
    zip64: bool,
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ZipArchive;
    use flate2::{Compression, write::DeflateEncoder};
    use std::io::{Cursor, Write};

    #[test]
    fn accounting_distinguishes_low_level_payload_kinds() {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let mut accounting = ZipOperationAccounting::default();
        let deflate = |data: &[u8]| {
            let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
            encoder.write_all(data).unwrap();
            encoder.finish().unwrap()
        };
        let generated_payload = b"generated payload";
        let generated = deflate(generated_payload);
        let precompressed_payload = b"precompressed payload";
        let precompressed = deflate(precompressed_payload);
        let stored_precompressed = b"stored precompressed payload";

        archive
            .write_stored_file_with_accounting("stored.bin", b"stored", &mut accounting)
            .unwrap();
        archive
            .write_generated_deflate_file_with_accounting(
                "generated.bin",
                crate::crc32(generated_payload),
                generated_payload.len() as u64,
                &generated,
                &mut accounting,
            )
            .unwrap();
        archive
            .write_precompressed_file_with_accounting(
                "precompressed.bin",
                CompressionMethod::Deflate,
                crate::crc32(precompressed_payload),
                precompressed_payload.len() as u64,
                &precompressed,
                &mut accounting,
            )
            .unwrap();
        archive
            .write_precompressed_file_with_accounting(
                "stored-precompressed.bin",
                CompressionMethod::Store,
                crate::crc32(stored_precompressed),
                stored_precompressed.len() as u64,
                stored_precompressed,
                &mut accounting,
            )
            .unwrap();
        archive.finish().unwrap();
        let bytes = output.into_inner();
        let reader = crate::office::ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.read("generated.bin").unwrap(), generated_payload);
        assert_eq!(
            reader.read("precompressed.bin").unwrap(),
            precompressed_payload
        );
        assert_eq!(
            reader.read("stored-precompressed.bin").unwrap(),
            stored_precompressed
        );

        assert_eq!(
            accounting.stored_payload_bytes_emitted(),
            6 + stored_precompressed.len() as u64
        );
        assert_eq!(
            accounting.generated_deflate_payload_bytes_emitted(),
            generated.len() as u64
        );
        assert_eq!(
            accounting.precompressed_payload_bytes_emitted(),
            precompressed.len() as u64
        );
    }

    #[test]
    fn sized_header_zip64_sizes_use_the_sentinel_boundary() {
        let crc32 = 0x1234_5678;
        let file_name_len = 9;
        for (size, uses_zip64) in [
            (u64::from(u32::MAX) - 1, false),
            (u64::from(u32::MAX), true),
            (u64::from(u32::MAX) + 1, true),
        ] {
            let local = sized_local_header(
                FLAG_UTF8_ENCODING,
                CompressionMethod::Store,
                crc32,
                size,
                size,
                file_name_len,
            )
            .expect("stored header metadata");
            let expected_size32 = if uses_zip64 { u32::MAX } else { size as u32 };
            let expected_extra_len = if uses_zip64 {
                usize::from(ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN)
            } else {
                0
            };
            assert_eq!(local.fixed.version_needed, if uses_zip64 { 45 } else { 20 });
            assert_eq!(local.fixed.compressed_size, expected_size32);
            assert_eq!(local.fixed.uncompressed_size, expected_size32);
            assert_eq!(usize::from(local.fixed.extra_field_len), expected_extra_len);

            let mut expected_extra = [0u8; ZIP64_LOCAL_SIZE_EXTRA_MAX_LEN];
            if uses_zip64 {
                expected_extra[..2].copy_from_slice(&ExtraFieldId::ZIP64.as_u16().to_le_bytes());
                expected_extra[2..4]
                    .copy_from_slice(&ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN.to_le_bytes());
                expected_extra[4..12].copy_from_slice(&size.to_le_bytes());
                expected_extra[12..20].copy_from_slice(&size.to_le_bytes());
            }
            assert_eq!(
                &local.extra[..expected_extra_len],
                &expected_extra[..expected_extra_len]
            );

            let mut central = FileHeader {
                name_len: file_name_len,
                compression_method: CompressionMethod::Store,
                local_header_offset: 0,
                compressed_size: size,
                uncompressed_size: size,
                crc: crc32,
                flags: FLAG_UTF8_ENCODING,
                zip64: false,
                modification_time: None,
                unix_permissions: None,
                extra_fields: ExtraFieldsContainer::new(),
            };
            assert_eq!(central.needs_zip64(), uses_zip64);
            central
                .finalize_extra_fields()
                .expect("central ZIP64 metadata");
            let mut central_extra = Vec::new();
            central
                .extra_fields
                .write_extra_fields(&mut central_extra, Header::CENTRAL)
                .expect("central extra fields");
            assert_eq!(central_extra, &expected_extra[..expected_extra_len]);
        }
    }

    #[test]
    fn sized_header_local_zip64_extra_carries_both_sizes() {
        let maximum = u64::from(u32::MAX);
        for (compressed_size, uncompressed_size) in [(maximum, maximum - 1), (maximum - 1, maximum)]
        {
            let local = sized_local_header(
                0,
                CompressionMethod::Deflate,
                0,
                compressed_size,
                uncompressed_size,
                4,
            )
            .expect("stored header metadata");
            assert_eq!(local.fixed.version_needed, ZIP64_VERSION_NEEDED);
            assert_eq!(local.fixed.compressed_size, u32::MAX);
            assert_eq!(local.fixed.uncompressed_size, u32::MAX);
            assert_eq!(
                local.fixed.extra_field_len,
                ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN
            );

            let mut expected_extra = [0u8; ZIP64_LOCAL_SIZE_EXTRA_MAX_LEN];
            expected_extra[..2].copy_from_slice(&ExtraFieldId::ZIP64.as_u16().to_le_bytes());
            expected_extra[2..4].copy_from_slice(&ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN.to_le_bytes());
            expected_extra[4..12].copy_from_slice(&uncompressed_size.to_le_bytes());
            expected_extra[12..20].copy_from_slice(&compressed_size.to_le_bytes());
            assert_eq!(local.extra, expected_extra);

            let mut central = FileHeader {
                name_len: 4,
                compression_method: CompressionMethod::Deflate,
                local_header_offset: 0,
                compressed_size,
                uncompressed_size,
                crc: 0,
                flags: 0,
                zip64: false,
                modification_time: None,
                unix_permissions: None,
                extra_fields: ExtraFieldsContainer::new(),
            };
            central
                .finalize_extra_fields()
                .expect("central ZIP64 metadata");
            let mut central_extra = Vec::new();
            central
                .extra_fields
                .write_extra_fields(&mut central_extra, Header::CENTRAL)
                .expect("central extra fields");
            let central_size = if compressed_size >= ZIP64_THRESHOLD_FILE_SIZE {
                compressed_size
            } else {
                uncompressed_size
            };
            let mut expected_central_extra = [0u8; 12];
            expected_central_extra[..2]
                .copy_from_slice(&ExtraFieldId::ZIP64.as_u16().to_le_bytes());
            expected_central_extra[2..4].copy_from_slice(&8u16.to_le_bytes());
            expected_central_extra[4..12].copy_from_slice(&central_size.to_le_bytes());
            assert_eq!(central_extra, expected_central_extra);
        }
    }

    #[test]
    fn precompressed_zip64_metadata_reopens_with_a_small_payload() {
        let compressed = b"small stored payload";
        let declared_uncompressed_size = u64::from(u32::MAX);
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let mut accounting = ZipOperationAccounting::default();
        archive
            .write_precompressed_file_with_accounting(
                "large.bin",
                CompressionMethod::Store,
                crate::crc32(compressed),
                declared_uncompressed_size,
                compressed,
                &mut accounting,
            )
            .expect("ZIP64 precompressed member");
        archive.finish().expect("ZIP64 archive finish");
        assert_eq!(
            accounting.stored_payload_bytes_emitted(),
            compressed.len() as u64
        );

        // This intentionally checks framing and source-index metadata only:
        // the tiny stored payload cannot satisfy the declared 4 GiB size.
        let bytes = output.into_inner();
        let archive = ZipArchive::from_slice(&bytes).expect("reopen ZIP64 archive");
        assert!(archive.is_zip64());
        let mut entries = archive.entries();
        let record = entries
            .next_entry()
            .expect("central entry")
            .expect("one central entry");
        assert!(
            entries
                .next_entry()
                .expect("end of central entries")
                .is_none()
        );
        assert!(record.is_zip64());
        assert_eq!(record.compressed_size_hint(), compressed.len() as u64);
        assert_eq!(record.uncompressed_size_hint(), declared_uncompressed_size);

        let local = ZipLocalFileHeaderFixed::parse(&bytes).expect("local header");
        assert_eq!(local.version_needed, ZIP64_VERSION_NEEDED);
        assert_eq!(local.compressed_size, u32::MAX);
        assert_eq!(local.uncompressed_size, u32::MAX);
        assert_eq!(local.extra_field_len, ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN);

        let entry = archive
            .get_entry(record.wayfinder())
            .expect("indexed local span");
        assert_eq!(entry.data(), compressed);
        let mut local_fields = entry.extra_fields();
        let (local_id, local_data) = local_fields.next().expect("local ZIP64 field");
        assert_eq!(local_id, ExtraFieldId::ZIP64);
        assert_eq!(
            local_data.len(),
            usize::from(ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN)
        );
        let declared_bytes = declared_uncompressed_size.to_le_bytes();
        let compressed_bytes = (compressed.len() as u64).to_le_bytes();
        assert_eq!(&local_data[..8], declared_bytes.as_slice());
        assert_eq!(&local_data[8..], compressed_bytes.as_slice());
        assert!(local_fields.next().is_none());

        let central_offset = usize::try_from(record.central_directory_offset()).unwrap();
        let central = ZipFileHeaderFixed::parse(&bytes[central_offset..]).expect("central header");
        assert_eq!(central.version_needed, ZIP64_VERSION_NEEDED);
        assert_eq!(central.compressed_size, compressed.len() as u32);
        assert_eq!(central.uncompressed_size, u32::MAX);
        assert_eq!(central.extra_field_len, 12);
        let mut central_fields = record.extra_fields();
        let (central_id, central_data) = central_fields.next().expect("central ZIP64 field");
        assert_eq!(central_id, ExtraFieldId::ZIP64);
        assert_eq!(central_data, declared_bytes.as_slice());
        assert!(central_fields.next().is_none());
    }

    #[test]
    fn explicit_zip64_streaming_deflate_reopens_with_a_small_payload() {
        let payload = b"streamed ZIP64 payload";
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);

        let (mut entry, config) = archive
            .new_file("small.txt")
            .compression_method(CompressionMethod::Deflate)
            .zip64(true)
            .start()
            .expect("ZIP64 streaming entry");
        let encoder = DeflateEncoder::new(&mut entry, Compression::default());
        let mut data_writer = config.wrap(encoder);
        data_writer.write_all(payload).expect("entry payload");
        let (encoder, descriptor) = data_writer.finish().expect("entry data finish");
        encoder.finish().expect("deflate finish");
        entry.finish(descriptor).expect("entry finish");
        archive.finish().expect("archive finish");

        let bytes = output.into_inner();
        let archive = ZipArchive::from_slice(&bytes).expect("reopen ZIP64 archive");
        assert!(archive.is_zip64());
        let mut entries = archive.entries();
        let record = entries
            .next_entry()
            .expect("central entry")
            .expect("one central entry");
        assert!(
            entries
                .next_entry()
                .expect("end of central entries")
                .is_none()
        );
        assert!(record.is_zip64());
        assert!(record.compressed_size_hint() > 0);
        assert_eq!(record.uncompressed_size_hint(), payload.len() as u64);

        let local = ZipLocalFileHeaderFixed::parse(&bytes).expect("local header");
        assert_eq!(local.version_needed, ZIP64_VERSION_NEEDED);
        assert_eq!(local.flags & FLAG_DATA_DESCRIPTOR, FLAG_DATA_DESCRIPTOR);
        assert_eq!(local.crc32, 0);
        assert_eq!(local.compressed_size, u32::MAX);
        assert_eq!(local.uncompressed_size, u32::MAX);
        assert_eq!(local.extra_field_len, ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN);
        let local_extra_start = ZipLocalFileHeaderFixed::SIZE + usize::from(local.file_name_len);
        let local_extra =
            &bytes[local_extra_start..local_extra_start + usize::from(local.extra_field_len)];
        assert_eq!(
            u16::from_le_bytes(local_extra[..2].try_into().unwrap()),
            ExtraFieldId::ZIP64.as_u16()
        );
        assert_eq!(
            u16::from_le_bytes(local_extra[2..4].try_into().unwrap()),
            ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN
        );
        assert!(local_extra[4..].iter().all(|byte| *byte == 0));

        let central_offset = usize::try_from(record.central_directory_offset()).unwrap();
        let central = ZipFileHeaderFixed::parse(&bytes[central_offset..]).expect("central header");
        assert_eq!(central.version_needed, ZIP64_VERSION_NEEDED);
        assert_eq!(central.compressed_size, u32::MAX);
        assert_eq!(central.uncompressed_size, u32::MAX);
        assert_eq!(central.extra_field_len, ZIP64_LOCAL_SIZE_EXTRA_FIELD_LEN);
        let mut central_fields = record.extra_fields();
        let (central_id, central_data) = central_fields.next().expect("central ZIP64 field");
        assert_eq!(central_id, ExtraFieldId::ZIP64);
        assert_eq!(
            central_data.len(),
            usize::from(ZIP64_LOCAL_SIZE_EXTRA_DATA_LEN)
        );
        assert_eq!(
            u64::from_le_bytes(central_data[..8].try_into().unwrap()),
            payload.len() as u64
        );
        assert_eq!(
            u64::from_le_bytes(central_data[8..].try_into().unwrap()),
            record.compressed_size_hint()
        );
        assert!(central_fields.next().is_none());

        let data_start = ZipLocalFileHeaderFixed::SIZE
            + usize::from(local.file_name_len)
            + usize::from(local.extra_field_len);
        let descriptor_start = data_start + record.compressed_size_hint() as usize;
        let descriptor = &bytes[descriptor_start..descriptor_start + 24];
        assert_eq!(
            u32::from_le_bytes(descriptor[..4].try_into().unwrap()),
            DataDescriptor::SIGNATURE
        );
        assert_eq!(
            u64::from_le_bytes(descriptor[8..16].try_into().unwrap()),
            record.compressed_size_hint()
        );
        assert_eq!(
            u64::from_le_bytes(descriptor[16..24].try_into().unwrap()),
            payload.len() as u64
        );

        let reader = crate::office::ArchiveReader::new(&bytes).expect("archive reader");
        assert_eq!(reader.read("small.txt").expect("inflate payload"), payload);
    }

    #[test]
    fn explicit_zip64_owned_streaming_deflate_reopens_with_a_small_payload() {
        let payload = b"owned streamed ZIP64 payload";
        let archive = ZipArchiveWriter::new(Vec::new());
        let mut entry = archive
            .start_file_owned_zip64("owned.txt", CompressionMethod::Deflate)
            .expect("owned ZIP64 streaming entry");
        entry.write_all(payload).expect("entry payload");
        let archive = entry.finish().expect("entry finish");
        let bytes = archive.finish().expect("archive finish");

        let reader = crate::office::ArchiveReader::new(&bytes).expect("archive reader");
        assert_eq!(reader.read("owned.txt").expect("inflate payload"), payload);
        let archive = ZipArchive::from_slice(&bytes).expect("reopen ZIP64 archive");
        assert!(archive.is_zip64());
        let record = archive
            .entries()
            .next_entry()
            .expect("central entry")
            .expect("one central entry");
        assert!(record.is_zip64());
    }

    #[test]
    fn explicit_zip64_rejects_user_zip64_extra_before_local_header() {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let error = archive
            .new_file("duplicate.txt")
            .zip64(true)
            .extra_field(ExtraFieldId::ZIP64, &[0u8; 16], Header::CENTRAL)
            .expect("user extra field");
        let error = error
            .start()
            .expect_err("explicit ZIP64 mode rejects a user ZIP64 field");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        assert_eq!(archive.stream_offset(), 0);
        drop(archive);
        assert!(output.into_inner().is_empty());
    }

    #[test]
    fn automatic_zip64_rejects_a_central_zip64_extra_collision() {
        let mut file_header = FileHeader {
            name_len: 1,
            compression_method: CompressionMethod::Store,
            local_header_offset: ZIP64_THRESHOLD_OFFSET,
            compressed_size: 0,
            uncompressed_size: 0,
            crc: 0,
            flags: 0,
            zip64: false,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
        };
        file_header
            .extra_fields
            .add_field(ExtraFieldId::ZIP64, &[0u8; 8], Header::CENTRAL)
            .expect("user central ZIP64 field");
        let error = file_header
            .finalize_extra_fields()
            .expect_err("automatic ZIP64 metadata must not duplicate user data");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        assert_eq!(file_header.extra_fields.central_size, 12);
    }

    #[test]
    fn zip32_streaming_size_boundary_refuses_before_descriptor() {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let entry = ZipEntryWriter {
            inner: &mut archive,
            compressed_bytes: ZIP64_THRESHOLD_FILE_SIZE,
            name: None,
            name_len: 1,
            local_header_offset: 0,
            compression_method: CompressionMethod::Store,
            zip64: false,
            flags: FLAG_DATA_DESCRIPTOR,
            modification_time: None,
            unix_permissions: None,
            extra_fields: ExtraFieldsContainer::new(),
        };
        let error = entry
            .finish(DataDescriptorOutput {
                crc: 0,
                compressed_size: 0,
                uncompressed_size: 0,
            })
            .expect_err("ZIP32 streaming entry must refuse ZIP64 size");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        assert_eq!(archive.stream_offset(), 0);
        drop(archive);
        assert!(output.into_inner().is_empty());
    }

    #[test]
    fn streaming_counters_reject_u64_overflow_before_sink_write() {
        let mut output = Vec::new();
        let mut count_writer = CountWriter::new(&mut output, u64::MAX);
        let error = count_writer.write(b"x").expect_err("offset overflow");
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
        assert!(output.is_empty());

        let mut output = Vec::new();
        let mut archive = ZipArchiveWriter::builder()
            .with_offset(u64::MAX)
            .build(&mut output);
        let error = archive
            .new_file("x")
            .start()
            .expect_err("offset overflow before local header");
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(archive.stream_offset(), u64::MAX);
        drop(archive);
        assert!(output.is_empty());

        let mut output = Vec::new();
        let mut data_writer = ZipDataWriter {
            inner: &mut output,
            uncompressed_bytes: u64::MAX,
            crc: 0,
            crc32_option: Crc32Option::Skip,
        };
        let error = data_writer
            .write(b"x")
            .expect_err("uncompressed count overflow");
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
        assert!(output.is_empty());

        let mut output = Vec::new();
        let archive = ZipArchiveWriter::new(&mut output);
        let mut compressed_entry = OwnedCompressedEntry {
            archive,
            state: OwnedEntryState {
                name: None,
                name_len: 1,
                local_header_offset: 0,
                compression_method: CompressionMethod::Store,
                zip64: false,
                flags: FLAG_DATA_DESCRIPTOR,
                modification_time: None,
                unix_permissions: None,
                extra_fields: ExtraFieldsContainer::new(),
            },
            compressed_bytes: u64::MAX,
            compressed_limit: None,
        };
        let error = compressed_entry
            .write(b"x")
            .expect_err("compressed count overflow");
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
        drop(compressed_entry);
        assert!(output.is_empty());
    }

    struct PartialPrecompressedSink<'a> {
        payload: &'a [u8],
        bytes: Vec<u8>,
        payload_started: bool,
    }

    impl Write for PartialPrecompressedSink<'_> {
        fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
            if !self.payload_started && buffer == self.payload {
                self.payload_started = true;
                let accepted = buffer.len().min(3);
                self.bytes.extend_from_slice(&buffer[..accepted]);
                return Ok(accepted);
            }
            if self.payload_started {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    "precompressed payload sink failure",
                ));
            }
            self.bytes.extend_from_slice(buffer);
            Ok(buffer.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn precompressed_accounting_charges_partial_payload_with_provenance() {
        let payload = b"precompressed partial payload";
        let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
        encoder.write_all(payload).unwrap();
        let compressed = encoder.finish().unwrap();
        assert!(compressed.len() > 3);

        let mut sink = PartialPrecompressedSink {
            payload: &compressed,
            bytes: Vec::new(),
            payload_started: false,
        };
        let mut archive = ZipArchiveWriter::new(&mut sink);
        let mut accounting = ZipOperationAccounting::default();
        let error = archive
            .write_precompressed_file_with_accounting(
                "partial.bin",
                CompressionMethod::Deflate,
                crate::crc32(payload),
                payload.len() as u64,
                &compressed,
                &mut accounting,
            )
            .unwrap_err();

        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(accounting.precompressed_payload_bytes_emitted(), 3);
        assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
        assert_eq!(accounting.stored_payload_bytes_emitted(), 0);
        assert_eq!(&sink.bytes[sink.bytes.len() - 3..], &compressed[..3]);
    }

    #[test]
    fn test_name_lifetime_independence() {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);

        // Test file builder with temporary name
        {
            let (mut entry, config) = {
                let temp_name = format!("temp-{}.txt", 42);
                archive.new_file(&temp_name).start().unwrap()
            };
            let mut writer = config.wrap(&mut entry);
            writer.write_all(b"test").unwrap();
            let (_, desc) = writer.finish().unwrap();
            entry.finish(desc).unwrap();
        }

        archive.finish().unwrap();
    }

    #[test]
    fn test_builder_with_offset_and_capacity() {
        let mut output = Cursor::new(Vec::new());

        output.write_all(b"PREFIX DATA").unwrap();
        let offset = output.position();

        let mut archive = ZipArchiveWriterBuilder::new()
            .with_capacity(5)
            .with_offset(offset)
            .build(&mut output);

        let (mut entry, config) = archive.new_file("test.txt").start().unwrap();
        let mut writer = config.wrap(&mut entry);
        writer.write_all(b"Hello World").unwrap();
        let (_, desc) = writer.finish().unwrap();
        entry.finish(desc).unwrap();

        archive.finish().unwrap();
    }

    #[test]
    fn test_stream_offset_methods() {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);

        // Test case 1: Get local header offset
        let local_header_offset = archive.stream_offset();
        let (mut file, config) = archive.new_file("test.txt").start().unwrap();

        // Test case 2: Get start of data offset
        let data_start_offset = file.stream_offset();

        // Write some data
        let mut writer = config.wrap(&mut file);
        writer.write_all(b"Hello World").unwrap();
        let (_, desc) = writer.finish().unwrap();

        // Test case 3: Get end of compressed data offset
        let end_data_offset = file.stream_offset();

        let compressed_bytes = file.finish(desc).unwrap();

        // Test case 4: Get end of data descriptor offset (next file's local header offset)
        let end_descriptor_offset = archive.stream_offset();

        archive.finish().unwrap();

        // Verify the offsets make sense
        assert_eq!(local_header_offset, 0);
        assert!(data_start_offset > local_header_offset);
        assert_eq!(
            end_data_offset,
            data_start_offset + b"Hello World".len() as u64
        );
        assert_eq!(end_descriptor_offset, end_data_offset + 16); // 16 bytes for data descriptor
        assert_eq!(compressed_bytes, end_data_offset - data_start_offset);
    }

    #[test]
    fn test_crc32_options() {
        use std::io::Write;

        let data = b"Hello, world!";
        let correct_crc = crate::crc32(data);
        let incorrect_crc = 0x12345678u32;

        // Test with default CRC calculation
        {
            let mut output = Cursor::new(Vec::new());
            let mut archive = ZipArchiveWriter::new(&mut output);
            let (mut entry, config) = archive.new_file("normal.txt").start().unwrap();
            let mut writer = config.wrap(&mut entry);
            writer.write_all(data).unwrap();
            let (_, descriptor) = writer.finish().unwrap();
            entry.finish(descriptor).unwrap();
            archive.finish().unwrap();
        }

        // Test with correct custom CRC - should succeed
        {
            let mut output = Cursor::new(Vec::new());
            let mut archive = ZipArchiveWriter::new(&mut output);
            let (mut entry, config) = archive
                .new_file("correct.txt")
                .crc32(Crc32Option::Custom(correct_crc))
                .start()
                .unwrap();
            let mut writer = config.wrap(&mut entry);
            writer.write_all(data).unwrap();
            let (_, descriptor) = writer.finish().unwrap();
            entry.finish(descriptor).unwrap();
            archive.finish().unwrap();

            // Verify the archive can be read
            let output = output.into_inner();
            let archive = ZipArchive::from_slice(&output).unwrap();
            let mut entries = archive.entries();
            let entry = entries.next_entry().unwrap().unwrap();
            let wayfinder = entry.wayfinder();
            let entry = archive.get_entry(wayfinder).unwrap();
            let mut verifier = entry.verifying_reader(entry.data());
            let mut actual = Vec::new();
            std::io::copy(&mut verifier, &mut actual).unwrap();
            assert_eq!(&actual, data);
        }

        // Test with incorrect custom CRC - verification should fail
        {
            let mut output = Cursor::new(Vec::new());
            let mut archive = ZipArchiveWriter::new(&mut output);
            let (mut entry, config) = archive
                .new_file("incorrect.txt")
                .crc32(Crc32Option::Custom(incorrect_crc))
                .start()
                .unwrap();
            let mut writer = config.wrap(&mut entry);
            writer.write_all(data).unwrap();
            let (_, descriptor) = writer.finish().unwrap();
            entry.finish(descriptor).unwrap();
            archive.finish().unwrap();

            // Verify the archive fails verification
            let output = output.into_inner();
            let archive = ZipArchive::from_slice(&output).unwrap();
            let mut entries = archive.entries();
            let entry = entries.next_entry().unwrap().unwrap();
            let wayfinder = entry.wayfinder();
            let entry = archive.get_entry(wayfinder).unwrap();
            let mut verifier = entry.verifying_reader(entry.data());
            let mut actual = Vec::new();
            let result = std::io::copy(&mut verifier, &mut actual);

            // Verification should fail with InvalidChecksum error
            assert!(result.is_err());
            let err = result.unwrap_err();
            assert_eq!(err.kind(), std::io::ErrorKind::InvalidData);
            let source = err.into_inner().unwrap();
            let zip_error = source.downcast::<crate::Error>().unwrap();
            match zip_error.kind() {
                ErrorKind::InvalidChecksum { expected, actual } => {
                    assert_eq!(*expected, incorrect_crc);
                    assert_eq!(*actual, correct_crc);
                },
                _ => panic!("Expected InvalidChecksum error, got {:?}", zip_error.kind()),
            }
        }

        // Test with skipped CRC - should have CRC of 0, and should validate fine
        {
            let mut output = Cursor::new(Vec::new());
            let mut archive = ZipArchiveWriter::new(&mut output);
            let (mut entry, config) = archive
                .new_file("skipped.txt")
                .crc32(Crc32Option::Skip)
                .start()
                .unwrap();
            let mut writer = config.wrap(&mut entry);
            writer.write_all(data).unwrap();
            let (_, descriptor) = writer.finish().unwrap();
            entry.finish(descriptor).unwrap();
            archive.finish().unwrap();

            // Verify the archive can be read
            let output = output.into_inner();
            let archive = ZipArchive::from_slice(&output).unwrap();
            let mut entries = archive.entries();
            let entry = entries.next_entry().unwrap().unwrap();
            let wayfinder = entry.wayfinder();
            let entry = archive.get_entry(wayfinder).unwrap();
            let mut verifier = entry.verifying_reader(entry.data());
            let mut actual = Vec::new();
            std::io::copy(&mut verifier, &mut actual).unwrap();
            assert_eq!(&actual, data);
        }
    }

    #[test]
    fn test_tuple_api() {
        use std::io::Write;

        let data = b"Hello, world!";
        let custom_crc = 0x12345678u32;

        // Test the new tuple-based API with custom CRC
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let (mut entry, config) = archive
            .new_file("test.txt")
            .crc32(Crc32Option::Custom(custom_crc))
            .start()
            .unwrap();

        // Using the new unified API - the CRC option is automatically configured
        let mut writer = config.wrap(&mut entry);
        writer.write_all(data).unwrap();
        let (_, descriptor) = writer.finish().unwrap();

        // Verify the CRC was correctly applied
        assert_eq!(descriptor.crc, custom_crc);

        entry.finish(descriptor).unwrap();
        archive.finish().unwrap();
    }

    #[test]
    fn test_owned_entry_writer_recovers_archive_for_store_and_deflate() {
        use std::io::Write;

        let mut output = Cursor::new(Vec::new());
        let archive = ZipArchiveWriter::new(&mut output);
        let mut entry = archive
            .start_file_owned("stored.txt", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"stored payload").unwrap();
        let archive = entry.finish().unwrap();

        let mut entry = archive
            .start_entry_owned("deflated.txt", CompressionMethod::Deflate)
            .unwrap();
        entry.write_all(b"deflated payload").unwrap();
        let archive = entry.finish().unwrap();
        archive.finish().unwrap();
        let reader = crate::office::ArchiveReader::new(output.get_ref()).unwrap();
        assert_eq!(reader.read("stored.txt").unwrap(), b"stored payload");
        assert_eq!(reader.read("deflated.txt").unwrap(), b"deflated payload");
    }

    #[test]
    #[allow(deprecated)]
    fn test_deprecated_create_method() {
        use std::io::Write;

        let data = b"Hello, deprecated API!";

        // Test that deprecated create() method still works
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        let mut entry = archive.new_file("deprecated.txt").create().unwrap();
        let mut writer = ZipDataWriter::new(&mut entry);
        writer.write_all(data).unwrap();
        let (_, descriptor) = writer.finish().unwrap();
        entry.finish(descriptor).unwrap();
        archive.finish().unwrap();

        // Verify the archive can be read
        let output = output.into_inner();
        let archive = ZipArchive::from_slice(&output).unwrap();
        let mut entries = archive.entries();
        let entry = entries.next_entry().unwrap().unwrap();
        let wayfinder = entry.wayfinder();
        let entry = archive.get_entry(wayfinder).unwrap();
        let mut verifier = entry.verifying_reader(entry.data());
        let mut actual = Vec::new();
        std::io::copy(&mut verifier, &mut actual).unwrap();
        assert_eq!(&actual, data);
    }
}
