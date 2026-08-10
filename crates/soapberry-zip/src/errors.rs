/// An error that occurred while reading or writing a zip file
#[derive(Debug)]
pub struct Error {
    inner: Box<ErrorInner>,
}

impl Error {
    /// Returns the offset of the end of central directory (EOCD) signature
    ///
    /// Useful for reparsing input that contains a false EOCD signature.
    pub fn eocd_offset(&self) -> Option<u64> {
        self.inner.eocd_offset
    }

    /// Sets the false signature offset on this error
    pub(crate) fn with_eocd_offset(mut self, offset: u64) -> Self {
        self.inner.eocd_offset = Some(offset);
        self
    }
}

impl Error {
    pub(crate) fn io(err: std::io::Error) -> Error {
        Error::from(ErrorKind::IO(err))
    }

    pub(crate) fn utf8(err: std::str::Utf8Error) -> Error {
        Error::from(ErrorKind::InvalidUtf8(err))
    }

    pub(crate) fn is_eof(&self) -> bool {
        matches!(self.inner.kind, ErrorKind::Eof)
    }

    /// The kind of error that occurred
    pub fn kind(&self) -> &ErrorKind {
        &self.inner.kind
    }
}

#[derive(Debug)]
struct ErrorInner {
    kind: ErrorKind,
    eocd_offset: Option<u64>,
}

/// The kind of error that occurred
#[derive(Debug)]
#[non_exhaustive]
pub enum ErrorKind {
    /// Missing end of central directory
    MissingEndOfCentralDirectory,

    /// Missing zip64 end of central directory
    MissingZip64EndOfCentralDirectory,

    /// Buffer size too small
    BufferTooSmall,

    /// Invalid end of central directory signature
    InvalidSignature { expected: u32, actual: u32 },

    /// Invalid inflated file crc checksum
    InvalidChecksum { expected: u32, actual: u32 },

    /// An unexpected inflated file size
    InvalidSize { expected: u64, actual: u64 },

    /// Invalid UTF-8 sequence
    InvalidUtf8(std::str::Utf8Error),

    /// An invalid input error with associated message
    InvalidInput { msg: String },

    /// An explicit parallel-read policy was internally inconsistent.
    InvalidParallelReadLimits { reason: &'static str },

    /// A member cannot fit within the caller-selected in-flight byte budget.
    ParallelReadInFlightBytesExceeded { actual: u64, maximum: u64 },

    /// A local parallel-read worker pool could not be created.
    ParallelReadWorkerPool { workers: usize, message: String },

    /// A parallel-read operation observed cooperative cancellation.
    Cancelled,

    /// A declared archive resource exceeds the caller-selected ceiling.
    LimitExceeded {
        /// The resource whose declared size or count exceeded its ceiling.
        resource: LimitResource,
        /// The declared or observed value that exceeded the ceiling.
        actual: u64,
        /// The caller-selected ceiling.
        maximum: u64,
    },

    /// Could not construct an archive with the given end of central directory
    InvalidEndOfCentralDirectory,

    /// An IO error
    IO(std::io::Error),

    /// An IO error (alias for compatibility)
    Io(std::io::Error),

    /// An unexpected end of file
    Eof,

    /// File not found in archive
    FileNotFound(String),

    /// Unsupported compression method
    UnsupportedCompressionMethod(u16),

    /// A ZIP layout cannot be safely preserved by the raw-copy writer.
    UnsupportedPreservation { reason: &'static str },
}

impl std::error::Error for Error {}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.inner.kind)?;
        Ok(())
    }
}

impl std::fmt::Display for ErrorKind {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match *self {
            ErrorKind::IO(ref err) => err.fmt(f),
            ErrorKind::MissingEndOfCentralDirectory => {
                write!(f, "Missing end of central directory")
            },
            ErrorKind::MissingZip64EndOfCentralDirectory => {
                write!(f, "Missing zip64 end of central directory")
            },
            ErrorKind::BufferTooSmall => {
                write!(f, "Buffer size too small")
            },
            ErrorKind::Eof => {
                write!(f, "Unexpected end of file")
            },
            ErrorKind::InvalidSignature { expected, actual } => {
                write!(
                    f,
                    "Invalid signature: expected 0x{:08x}, got 0x{:08x}",
                    expected, actual
                )
            },
            ErrorKind::InvalidChecksum { expected, actual } => {
                write!(
                    f,
                    "Invalid checksum: expected 0x{:08x}, got 0x{:08x}",
                    expected, actual
                )
            },
            ErrorKind::InvalidSize { expected, actual } => {
                write!(f, "Invalid size: expected {}, got {}", expected, actual)
            },
            ErrorKind::InvalidUtf8(ref err) => {
                write!(f, "Invalid UTF-8: {}", err)
            },
            ErrorKind::InvalidInput { ref msg } => {
                write!(f, "Invalid input: {}", msg)
            },
            ErrorKind::InvalidParallelReadLimits { reason } => {
                write!(f, "Invalid parallel read limits: {reason}")
            },
            ErrorKind::ParallelReadInFlightBytesExceeded { actual, maximum } => {
                write!(
                    f,
                    "Parallel read in-flight byte limit exceeded: declared {actual}, maximum {maximum}"
                )
            },
            ErrorKind::ParallelReadWorkerPool {
                workers,
                ref message,
            } => {
                write!(
                    f,
                    "Could not create local parallel read pool with {workers} worker(s): {message}"
                )
            },
            ErrorKind::Cancelled => write!(f, "Operation cancelled"),
            ErrorKind::LimitExceeded {
                resource,
                actual,
                maximum,
            } => {
                write!(
                    f,
                    "ZIP {} limit exceeded: declared {}, maximum {}",
                    resource, actual, maximum
                )
            },
            ErrorKind::InvalidEndOfCentralDirectory => {
                write!(f, "Invalid end of central directory")
            },
            ErrorKind::Io(ref err) => err.fmt(f),
            ErrorKind::FileNotFound(ref name) => {
                write!(f, "File not found in archive: {}", name)
            },
            ErrorKind::UnsupportedCompressionMethod(method) => {
                write!(f, "Unsupported compression method: {}", method)
            },
            ErrorKind::UnsupportedPreservation { reason } => {
                write!(f, "Unsupported ZIP preservation layout: {reason}")
            },
        }
    }
}

/// A resource governed by [`ErrorKind::LimitExceeded`].
///
/// Values are declared ZIP metadata, except [`Self::FileCount`], which is the
/// number of non-directory members accepted into an Office archive index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitResource {
    /// Number of non-directory members.
    FileCount,
    /// Bytes in one raw member name.
    MemberNameBytes,
    /// Aggregate central-directory variable metadata bytes.
    ///
    /// This includes member names, extra fields, and file comments, including
    /// those on directory entries.
    MetadataBytes,
    /// Declared compressed bytes for one non-directory member.
    CompressedSize,
    /// Declared uncompressed bytes for one non-directory member.
    EntrySize,
    /// Aggregate declared uncompressed bytes for non-directory members.
    TotalSize,
}

impl std::fmt::Display for LimitResource {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::FileCount => "file count",
            Self::MemberNameBytes => "member name bytes",
            Self::MetadataBytes => "central-directory metadata bytes",
            Self::CompressedSize => "compressed member size",
            Self::EntrySize => "uncompressed member size",
            Self::TotalSize => "total uncompressed size",
        })
    }
}

impl From<ErrorKind> for Error {
    fn from(kind: ErrorKind) -> Error {
        Error {
            inner: Box::new(ErrorInner {
                kind,
                eocd_offset: None,
            }),
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(err: std::io::Error) -> Error {
        Error::from(ErrorKind::IO(err))
    }
}
