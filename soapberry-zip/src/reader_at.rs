use crate::errors::{Error, ErrorKind};
use std::io::Read;
use std::ops::Range;
#[cfg(unix)]
use std::os::unix::fs::FileExt;
#[cfg(windows)]
use std::os::windows::fs::FileExt;
use std::{rc::Rc, sync::Arc};

/// Provides reading bytes at a specific offset
///
/// This trait is similar to [`std::io::Read`] but with an additional offset
/// parameter that signals where the read should begin offset from the start of
/// the data. This allows methods to not require a mutable reference to the
/// reader, which is critical for zip files to easily offer decompression of
/// multiple files simultaneously without needing to store them in memory.
///
/// This trait is modelled after Go's
/// [`io.ReaderAt`](https://pkg.go.dev/io#ReaderAt) interface, which is used by
/// their own [Zip implementation](https://pkg.go.dev/archive/zip#NewReader).
pub trait ReaderAt {
    /// Read bytes from the reader at a specific offset
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize>;

    /// Sibling to [`read_exact`](std::io::Read::read_exact), but at an offset
    fn read_exact_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<()> {
        let mut read = 0;
        while read < buf.len() {
            let latest = self.read_at(&mut buf[read..], offset + (read as u64))?;
            if latest == 0 {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::UnexpectedEof,
                    "failed to fill whole buffer",
                ));
            }
            read += latest;
        }
        Ok(())
    }
}

pub(crate) trait ReaderAtExt {
    fn try_read_at_least_at(
        &self,
        buffer: &mut [u8],
        size: usize,
        offset: u64,
    ) -> std::io::Result<usize>;

    fn read_at_least_at(&self, buffer: &mut [u8], size: usize, offset: u64)
    -> Result<usize, Error>;
}

impl<T: ReaderAt> ReaderAtExt for T {
    fn try_read_at_least_at(
        &self,
        buffer: &mut [u8],
        mut size: usize,
        offset: u64,
    ) -> std::io::Result<usize> {
        size = size.min(buffer.len());
        let mut pos = 0;
        while pos < size {
            let read = self.read_at(&mut buffer[pos..], offset + pos as u64)?;
            if read == 0 {
                return Ok(pos);
            }
            pos += read;
        }
        Ok(pos)
    }

    fn read_at_least_at(
        &self,
        buffer: &mut [u8],
        size: usize,
        offset: u64,
    ) -> Result<usize, Error> {
        if buffer.len() < size {
            return Err(Error::from(ErrorKind::BufferTooSmall));
        }

        let read = self.try_read_at_least_at(buffer, size, offset)?;

        if read < size {
            return Err(Error::from(ErrorKind::Eof));
        }

        Ok(read)
    }
}

#[cfg(not(any(unix, windows)))]
#[derive(Debug)]
pub struct FileReader(MutexReader<std::fs::File>);

/// A file wrapper that implements [`ReaderAt`] across platforms.
#[cfg(any(unix, windows))]
#[derive(Debug)]
pub struct FileReader(std::fs::File);

impl FileReader {
    pub fn into_inner(self) -> std::fs::File {
        #[cfg(not(any(unix, windows)))]
        return self.0.into_inner();
        #[cfg(any(unix, windows))]
        return self.0;
    }
}

impl ReaderAt for FileReader {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        #[cfg(unix)]
        return self.0.read_at(buf, offset);
        #[cfg(windows)]
        return self.0.seek_read(buf, offset);
        #[cfg(not(any(unix, windows)))]
        return self.0.read_at(buf, offset);
    }
}

impl std::io::Seek for FileReader {
    #[inline]
    fn seek(&mut self, pos: std::io::SeekFrom) -> std::io::Result<u64> {
        self.0.seek(pos)
    }
}

impl From<std::fs::File> for FileReader {
    #[cfg(not(any(unix, windows)))]
    fn from(file: std::fs::File) -> Self {
        Self(MutexReader(std::sync::Mutex::new(file)))
    }

    #[cfg(any(unix, windows))]
    fn from(file: std::fs::File) -> Self {
        Self(file)
    }
}

/// A reader that is wrapped in a mutex to allow for concurrent reads.
#[derive(Debug)]
pub struct MutexReader<R>(std::sync::Mutex<R>);

impl<R> MutexReader<R> {
    pub fn new(inner: R) -> Self {
        Self(std::sync::Mutex::new(inner))
    }

    pub fn into_inner(self) -> R {
        self.0.into_inner().unwrap()
    }
}

impl<R> ReaderAt for MutexReader<R>
where
    R: std::io::Read + std::io::Seek,
{
    /// For seekable implementations, we can emulate the read_at method by
    /// seeking to the offset, reading the data, and then seeking back to the
    /// original position within a mutex.
    ///
    /// This is how Go implements the `io.ReaderAt` interface for filed on
    /// Windows:
    /// https://github.com/golang/go/blob/70b603f4d295573197b43ad090d7cad21895144e/src/internal/poll/fd_windows.go#L525
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        let mut lock = self.0.lock().unwrap();
        let original_position = lock.stream_position()?;
        lock.seek(std::io::SeekFrom::Start(offset))?;
        let result = lock.read(buf);
        lock.seek(std::io::SeekFrom::Start(original_position))?;
        result
    }
}

impl<R> std::io::Read for MutexReader<R>
where
    R: std::io::Read,
{
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        self.0.lock().unwrap().read(buf)
    }
}

impl<R> std::io::Seek for MutexReader<R>
where
    R: std::io::Seek,
{
    fn seek(&mut self, pos: std::io::SeekFrom) -> std::io::Result<u64> {
        self.0.lock().unwrap().seek(pos)
    }
}

impl<T: ReaderAt> ReaderAt for &'_ T {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        (*self).read_at(buf, offset)
    }
}

impl<T: ReaderAt> ReaderAt for &'_ mut T {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        (**self).read_at(buf, offset)
    }
}

impl ReaderAt for &[u8] {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        let skip = self.len().min(offset as usize);
        let data = &self[skip..];
        let len = data.len().min(buf.len());
        buf[..len].copy_from_slice(&data[..len]);
        Ok(len)
    }
}

impl<R> ReaderAt for std::io::Cursor<R>
where
    R: AsRef<[u8]>,
{
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        let data = self.get_ref().as_ref();
        data.read_at(buf, offset)
    }
}

impl ReaderAt for Vec<u8> {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        self.as_slice().read_at(buf, offset)
    }
}

impl<T: ReaderAt + ?Sized> ReaderAt for Arc<T> {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        (**self).read_at(buf, offset)
    }
}

impl<T: ReaderAt + ?Sized> ReaderAt for Rc<T> {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        (**self).read_at(buf, offset)
    }
}

impl<T: ReaderAt + ?Sized> ReaderAt for Box<T> {
    #[inline]
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        (**self).read_at(buf, offset)
    }
}

/// A reader that reads a specific range of data from a [`ReaderAt`] source.
///
/// `RangeReader` implements [`std::io::Read`] and provides bounded reading
/// within a specified range of offsets. It maintains its current position and
/// ensures reads don't exceed the defined end boundary.
///
/// Useful when working with APIs that operate on [`std::io::Read`] instead of
/// [`ReaderAt`]. For instance, incrementally reading large prelude and trailing
/// data of a ZIP file.
///
/// # Examples
///
/// Reading prelude data from a zip file:
///
/// ```
/// use std::io::Read;
/// use rawzip::{ZipArchive, RangeReader, RECOMMENDED_BUFFER_SIZE};
/// use std::fs::File;
///
/// let file = File::open("assets/test-prefix.zip")?;
/// let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
/// let archive = ZipArchive::from_file(file, &mut buffer)?;
///
/// // Typically you only need the first entry to find where the zip data starts
/// // but this is the longer form that examines every entry in case they are
/// // out of order
/// let mut zip_start_offset = archive.directory_offset();
/// let mut entries = archive.entries(&mut buffer);
/// while let Some(entry) = entries.next_entry()? {
///     zip_start_offset = zip_start_offset.min(entry.local_header_offset());
/// }
///
/// // For example purposes, just slurp up all the prelude data
/// let mut prelude_reader = RangeReader::new(archive.get_ref(), 0..zip_start_offset);
/// prelude_reader.read_exact(&mut buffer[..zip_start_offset as usize])?;
/// assert_eq!(
///     &buffer[..zip_start_offset as usize],
///     b"prefix that could be an executable jar file"
/// );
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct RangeReader<R> {
    archive: R,
    offset: u64,
    end_offset: u64,
}

impl<R> RangeReader<R> {
    /// Creates a new `RangeReader` that will read data from the specified range.
    #[inline]
    pub fn new(archive: R, range: Range<u64>) -> Self {
        Self {
            archive,
            offset: range.start,
            end_offset: range.end,
        }
    }

    /// Returns the current read position within the range.
    #[inline]
    pub fn position(&self) -> u64 {
        self.offset
    }

    /// Returns the remaining bytes that are expected to be read from the
    /// current position.
    ///
    /// When a range reader is constructed with a range that exceeds the
    /// underlying reader, remaining will be non-zero when `read()` returns zero
    /// signalling the end of the stream.
    #[inline]
    pub fn remaining(&self) -> u64 {
        self.end_offset - self.offset
    }

    /// Returns the end offset of the range.
    #[inline]
    pub fn end_offset(&self) -> u64 {
        self.end_offset
    }

    /// Returns a reference to the underlying reader.
    #[inline]
    pub fn get_ref(&self) -> &R {
        &self.archive
    }

    /// Consumes the self and returns the underlying reader.
    #[inline]
    pub fn into_inner(self) -> R {
        self.archive
    }
}

impl<R> Read for RangeReader<R>
where
    R: ReaderAt,
{
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let read_size = buf.len().min(self.remaining() as usize);
        let read = self.archive.read_at(&mut buf[..read_size], self.offset)?;
        self.offset += read as u64;
        Ok(read)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    const TEST_DATA: &[u8] = b"Hello, World! This is test data for ReaderAt implementations.";

    fn test_reader_at_impl<R: ReaderAt>(reader: R, data_len: usize) {
        let mut buf = [0u8; 5];

        // Test reading from start
        assert_eq!(reader.read_at(&mut buf, 0).unwrap(), 5);
        assert_eq!(&buf, b"Hello");

        // Test reading from offset
        buf.fill(0);
        assert_eq!(reader.read_at(&mut buf, 7).unwrap(), 5);
        assert_eq!(&buf, b"World");

        // Test read beyond data length
        buf.fill(0);
        let bytes_read = reader.read_at(&mut buf, data_len as u64).unwrap();
        assert_eq!(bytes_read, 0);

        // Test partial read at end of data
        buf.fill(0);
        let bytes_read = reader.read_at(&mut buf, (data_len - 3) as u64).unwrap();
        assert_eq!(bytes_read, 3);
        assert_eq!(&buf[..3], &TEST_DATA[data_len - 3..]);
    }

    #[test]
    fn test_smart_pointer_implementations() {
        let data = TEST_DATA.to_vec();

        // Test Arc<Vec<u8>>
        let arc_reader = Arc::new(data.clone());
        test_reader_at_impl(&*arc_reader, data.len());
        test_reader_at_impl(arc_reader, data.len());

        // Test Rc<Vec<u8>>
        let rc_reader = Rc::new(data.clone());
        test_reader_at_impl(&*rc_reader, data.len());
        test_reader_at_impl(rc_reader, data.len());

        // Test Box<Vec<u8>>
        let box_reader = Box::new(data.clone());
        test_reader_at_impl(&*box_reader, data.len());
        test_reader_at_impl(box_reader, data.len());
    }

    #[test]
    fn test_reference_implementations() {
        let mut data = TEST_DATA.to_vec();
        let data_len = data.len();

        test_reader_at_impl(&data, data_len);
        test_reader_at_impl(&mut data, data_len);
    }

    #[test]
    fn test_byte_slice_implementation() {
        let data = TEST_DATA;
        test_reader_at_impl(data, data.len());
    }

    #[test]
    fn test_cursor_implementation() {
        let data = TEST_DATA.to_vec();
        let cursor = Cursor::new(data.clone());
        test_reader_at_impl(&cursor, data.len());
    }

    #[test]
    fn test_vec_implementation() {
        let data = TEST_DATA.to_vec();
        test_reader_at_impl(&data, data.len());
    }

    #[test]
    fn test_range_reader_basic() {
        let data = b"Hello, World! This is test data.";
        let mut range_reader = RangeReader::new(data.as_slice(), 7..13);

        let mut buffer = [0u8; 10];
        let bytes_read = range_reader.read(&mut buffer).unwrap();

        assert_eq!(bytes_read, 6);
        assert_eq!(&buffer[..bytes_read], b"World!");
    }

    #[test]
    fn test_range_reader_multiple_reads() {
        let data = b"0123456789";
        let mut range_reader = RangeReader::new(data.as_slice(), 2..8);

        let mut buffer = [0u8; 3];
        let bytes_read1 = range_reader.read(&mut buffer).unwrap();
        assert_eq!(bytes_read1, 3);
        assert_eq!(&buffer[..bytes_read1], b"234");
        assert_eq!(range_reader.position(), 5);

        let bytes_read2 = range_reader.read(&mut buffer).unwrap();
        assert_eq!(bytes_read2, 3);
        assert_eq!(&buffer[..bytes_read2], b"567");
        assert_eq!(range_reader.position(), 8);

        // Should return 0 when at end
        let bytes_read3 = range_reader.read(&mut buffer).unwrap();
        assert_eq!(bytes_read3, 0);
    }

    #[test]
    fn test_range_reader_empty_range() {
        let data = b"Hello, World!";
        let mut range_reader = RangeReader::new(data.as_slice(), 5..5);

        let mut buffer = [0u8; 10];
        let bytes_read = range_reader.read(&mut buffer).unwrap();

        assert_eq!(bytes_read, 0);
        assert_eq!(range_reader.remaining(), 0);
    }

    #[test]
    fn test_range_reader_get_ref_and_into_inner() {
        let data = b"Hello, World!";
        let range_reader = RangeReader::new(data.as_slice(), 0..5);

        assert_eq!(range_reader.get_ref(), &data.as_slice());
        let inner = range_reader.into_inner();
        assert_eq!(inner, data.as_slice());
    }

    #[test]
    fn test_range_reader_clone() {
        let data = b"Hello, World!";
        let range_reader = RangeReader::new(data.as_slice(), 0..5);
        let cloned = range_reader.clone();

        assert_eq!(range_reader.position(), cloned.position());
        assert_eq!(range_reader.remaining(), cloned.remaining());
    }

    #[test]
    fn test_range_reader_range_exceeds_data() {
        let data = b"Hello";

        // Test range that starts within data but extends beyond
        let mut reader1 = RangeReader::new(data.as_slice(), 3..10);
        let mut buf1 = [0u8; 10];
        let read1 = reader1.read(&mut buf1).unwrap();
        assert_eq!(read1, 2); // Only reads "lo"
        assert_eq!(&buf1[..read1], b"lo");

        // Test range that starts at end of data
        let mut reader2 = RangeReader::new(data.as_slice(), 5..10);
        let mut buf2 = [0u8; 10];
        let read2 = reader2.read(&mut buf2).unwrap();
        assert_eq!(read2, 0); // No data to read

        // Test range that starts beyond data
        let mut reader3 = RangeReader::new(data.as_slice(), 10..20);
        let mut buf3 = [0u8; 10];
        let read3 = reader3.read(&mut buf3).unwrap();
        assert_eq!(read3, 0); // No data to read
    }
}
