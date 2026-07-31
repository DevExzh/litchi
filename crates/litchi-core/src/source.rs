//! Immutable positional byte sources.

use std::{
    io,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
};

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

/// Opaque identity and revision for one stable source snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SourceVersion {
    id: u64,
    revision: u64,
}

impl SourceVersion {
    /// Creates a version token for a custom source adapter.
    ///
    /// The adapter owns the meaning of `id` and must keep it stable for one
    /// source while increasing `revision` whenever observable bytes change.
    pub const fn new(id: u64, revision: u64) -> Self {
        Self { id, revision }
    }

    fn fresh() -> Self {
        Self::new(NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed), 0)
    }

    /// Opaque process-local source identity.
    pub const fn id(self) -> u64 {
        self.id
    }

    /// Source revision captured by the adapter.
    pub const fn revision(self) -> u64 {
        self.revision
    }
}

/// Thread-safe immutable positional input.
///
/// Implementations must not maintain a shared cursor. A document records
/// `version()` before work and compares it later when the adapter can detect
/// external mutation.
pub trait ReadAt: Send + Sync {
    /// Current byte length of the source snapshot.
    fn len(&self) -> io::Result<u64>;

    /// Returns whether this source is empty.
    fn is_empty(&self) -> io::Result<bool> {
        self.len().map(|length| length == 0)
    }

    /// Reads bytes at `offset` without changing shared state.
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize>;

    /// Returns the adapter's current source identity/revision.
    fn version(&self) -> io::Result<SourceVersion>;

    /// Fills `output` or returns `UnexpectedEof` without panicking.
    fn read_exact_at(&self, mut offset: u64, mut output: &mut [u8]) -> io::Result<()> {
        while !output.is_empty() {
            match self.read_at(offset, output) {
                Ok(0) => {
                    return Err(io::Error::new(
                        io::ErrorKind::UnexpectedEof,
                        "positional source ended before the requested range",
                    ));
                },
                Ok(read) => {
                    if read > output.len() {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "positional source reported more bytes than requested",
                        ));
                    }
                    offset = offset.checked_add(read as u64).ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "source offset overflow")
                    })?;
                    output = &mut output[read..];
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(error) => return Err(error),
            }
        }
        Ok(())
    }
}

/// Move-owned in-memory source with O(1) clones and no source-byte copy.
#[derive(Debug, Clone)]
pub struct OwnedSource {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
}

impl OwnedSource {
    /// Moves a vector into shared immutable ownership.
    pub fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            version: SourceVersion::fresh(),
        }
    }

    /// Wraps an existing shared vector without copying its bytes.
    pub fn from_arc(bytes: Arc<Vec<u8>>) -> Self {
        Self {
            bytes,
            version: SourceVersion::fresh(),
        }
    }

    /// Direct borrowed access for parsers able to use a contiguous source.
    pub fn as_slice(&self) -> &[u8] {
        self.bytes.as_slice()
    }
}

impl From<Vec<u8>> for OwnedSource {
    fn from(bytes: Vec<u8>) -> Self {
        Self::new(bytes)
    }
}

impl ReadAt for OwnedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        read_slice_at(self.as_slice(), offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

/// Scoped zero-copy positional view over borrowed bytes.
#[derive(Debug, Clone, Copy)]
pub struct SliceSource<'a> {
    bytes: &'a [u8],
    version: SourceVersion,
}

impl<'a> SliceSource<'a> {
    /// Creates a scoped source without copying the bytes.
    pub fn new(bytes: &'a [u8]) -> Self {
        Self {
            bytes,
            version: SourceVersion::fresh(),
        }
    }

    /// Returns the borrowed source bytes.
    pub const fn as_slice(&self) -> &'a [u8] {
        self.bytes
    }
}

impl ReadAt for SliceSource<'_> {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        read_slice_at(self.bytes, offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

fn read_slice_at(input: &[u8], offset: u64, output: &mut [u8]) -> io::Result<usize> {
    let Ok(start) = usize::try_from(offset) else {
        return Ok(0);
    };
    let Some(input) = input.get(start..) else {
        return Ok(0);
    };
    let count = input.len().min(output.len());
    output[..count].copy_from_slice(&input[..count]);
    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn positional_reads_are_independent() {
        let source = OwnedSource::new(b"abcdef".to_vec());
        let mut first = [0; 2];
        let mut second = [0; 3];

        source.read_exact_at(4, &mut first).expect("valid range");
        source.read_exact_at(1, &mut second).expect("valid range");

        assert_eq!(&first, b"ef");
        assert_eq!(&second, b"bcd");
    }

    #[test]
    fn exact_read_reports_eof() {
        let source = SliceSource::new(b"abc");
        let mut output = [0; 4];
        let error = source
            .read_exact_at(0, &mut output)
            .expect_err("range exceeds source");
        assert_eq!(error.kind(), io::ErrorKind::UnexpectedEof);
    }

    #[test]
    fn owned_clones_keep_snapshot_identity() {
        let source = OwnedSource::new(vec![1, 2, 3]);
        let clone = source.clone();
        assert_eq!(source.version().ok(), clone.version().ok());
        assert!(Arc::ptr_eq(&source.bytes, &clone.bytes));
    }

    #[test]
    fn custom_adapters_can_report_versions() {
        let version = SourceVersion::new(41, 7);
        assert_eq!(version.id(), 41);
        assert_eq!(version.revision(), 7);
    }
}
