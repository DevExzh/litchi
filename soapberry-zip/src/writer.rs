use crate::{
    CENTRAL_HEADER_SIGNATURE, CompressionMethod, DataDescriptor,
    END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE, END_OF_CENTRAL_DIR_SIGNATURE64,
    END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES, Error, Header, ZipFileHeaderFixed, ZipLocalFileHeaderFixed,
    crc,
    errors::ErrorKind,
    extra_fields::{ExtraFieldId, ExtraFieldsContainer},
    mode::CREATOR_UNIX,
    path::{NormalizedPath, ZipFilePath},
    time::{DosDateTime, UtcDateTime},
};
use std::io::{self, Write};

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
        let bytes_written = self.writer.write(buf)?;
        self.count += bytes_written as u64;
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
    /// let mut archive = rawzip::ZipArchiveWriter::builder()
    ///     .with_offset(zip_start_offset)  // Tell the archive where it starts
    ///     .build(&mut output);
    ///
    /// // Add files normally
    /// let mut file = archive.new_file("data.txt").create().unwrap();
    /// let mut writer = rawzip::ZipDataWriter::new(&mut file);
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
        }
    }
}

/// Create a new Zip archive.
///
/// Basic usage:
/// ```rust
/// use std::io::Write;
///
/// let mut output = std::io::Cursor::new(Vec::new());
/// let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
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
/// let mut _archive = rawzip::ZipArchiveWriter::builder()
///     .with_capacity(1000)  // Optimize for 1000 anticipated files
///     .build(&mut output);
/// // ... add files as usual
/// ```
#[derive(Debug)]
pub struct ZipArchiveWriter<W> {
    files: Vec<FileHeader>,
    file_names: Vec<u8>,
    writer: CountWriter<W>,
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
    /// let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
    ///
    /// // 1. Get local header offset
    /// let local_header_offset = archive.stream_offset();
    /// let mut file = archive.new_file("test.txt").create().unwrap();
    ///
    /// // 2. Get start of data offset
    /// let data_start_offset = file.stream_offset();
    ///
    /// // Write some data
    /// let mut writer = rawzip::ZipDataWriter::new(&mut file);
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
    /// # use rawzip::{ZipArchive, ZipArchiveWriter, ZipDataWriter, extra_fields::ExtraFieldId, Header};
    /// let mut output = Cursor::new(Vec::new());
    /// let mut archive = ZipArchiveWriter::new(&mut output);
    ///
    /// let my_custom_field = ExtraFieldId::new(0x6666);
    ///
    /// // File with extra fields only in the local file header
    /// let mut local_file = archive.new_file("video.mp4")
    ///     .extra_field(my_custom_field, b"field1", Header::LOCAL)?
    ///     .create()?;
    /// let mut writer = ZipDataWriter::new(&mut local_file);
    /// writer.write_all(b"video data")?;
    /// let (_, desc) = writer.finish()?;
    /// local_file.finish(desc)?;
    ///
    /// // File with extra fields only in the central directory
    /// let mut central_file = archive.new_file("document.pdf")
    ///     .extra_field(my_custom_field, b"field2", Header::CENTRAL)?
    ///     .create()?;
    /// let mut writer = ZipDataWriter::new(&mut central_file);
    /// writer.write_all(b"PDF content")?;
    /// let (_, desc) = writer.finish()?;
    /// central_file.finish(desc)?;
    ///
    /// // File with extra fields in both headers for maximum compatibility
    /// assert_eq!(Header::default(), Header::LOCAL | Header::CENTRAL);
    /// let mut both_file = archive.new_file("important.dat")
    ///     .extra_field(my_custom_field, b"field3", Header::default())?
    ///     .create()?;
    /// let mut writer = ZipDataWriter::new(&mut both_file);
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
    /// # let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
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
    /// # let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
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
    /// Writes a local file header with filtered extra fields.
    fn write_local_header(
        &mut self,
        file_path: &ZipFilePath<NormalizedPath>,
        flags: u16,
        compression_method: CompressionMethod,
        options: &mut ZipEntryOptions,
    ) -> Result<(), Error> {
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

        let header = ZipLocalFileHeaderFixed {
            signature: ZipLocalFileHeaderFixed::SIGNATURE,
            version_needed: 20,
            flags,
            compression_method: compression_method.as_id(),
            last_mod_time: dos_time,
            last_mod_date: dos_date,
            crc32: 0, // must be zero if data descriptor is used (4.4.4)
            compressed_size: 0,
            uncompressed_size: 0,
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
    /// # let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
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
        let mut flags = 0u16;
        if file_path.needs_utf8_encoding() {
            flags |= FLAG_UTF8_ENCODING;
        } else {
            flags &= !FLAG_UTF8_ENCODING;
        }

        // Store the name bytes in the central buffer
        let name_bytes = file_path.as_ref().as_bytes();
        let name_len = name_bytes.len() as u16;
        self.file_names.extend_from_slice(name_bytes);

        self.write_local_header(&file_path, flags, CompressionMethod::Store, &mut options)?;

        let file_header = FileHeader {
            name_len,
            compression_method: CompressionMethod::Store,
            local_header_offset,
            compressed_size: 0,
            uncompressed_size: 0,
            crc: 0,
            flags,
            modification_time: options.modification_time,
            unix_permissions: options.unix_permissions,
            extra_fields: options.extra_fields,
        };
        self.files.push(file_header);

        Ok(())
    }

    /// Creates a builder for adding a new file to the archive.
    ///
    /// # Example
    ///
    /// ```rust
    /// # use std::io::{Cursor, Write};
    /// # let mut output = Cursor::new(Vec::new());
    /// # let mut archive = rawzip::ZipArchiveWriter::new(&mut output);
    /// let (mut entry, config) = archive.new_file("my-file")
    ///     .compression_method(rawzip::CompressionMethod::Deflate)
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
        self.file_names.extend_from_slice(name_bytes);

        self.write_local_header(&file_path, flags, options.compression_method, &mut options)?;

        Ok(ZipEntryWriter {
            inner: self,
            compressed_bytes: 0,
            name_len,
            local_header_offset,
            compression_method: options.compression_method,
            flags,
            modification_time: options.modification_time,
            unix_permissions: options.unix_permissions,
            extra_fields: options.extra_fields,
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
        let central_directory_offset = self.writer.count();
        let total_entries = self.files.len();

        // Determine if we need ZIP64 format
        let needs_zip64 = total_entries >= ZIP64_THRESHOLD_ENTRIES
            || central_directory_offset >= ZIP64_THRESHOLD_OFFSET
            || self.files.iter().any(|f| f.needs_zip64());

        let mut name_offset = 0;

        // Write central directory entries
        for file in &self.files {
            // Version made by and version needed to extract
            let version_needed = if file.needs_zip64() {
                ZIP64_VERSION_NEEDED
            } else {
                20
            };

            // Set version_made_by to indicate Unix when Unix permissions are present
            let version_made_by_hi = file.unix_permissions.map(|_| CREATOR_UNIX).unwrap_or(0);
            let version_made_by = (version_made_by_hi << 8) | version_needed;

            let (dos_time, dos_date) = file
                .modification_time
                .as_ref()
                .map(|dt| DosDateTime::from(dt).into_parts())
                .unwrap_or((0, 0));

            let header = ZipFileHeaderFixed {
                signature: CENTRAL_HEADER_SIGNATURE,
                version_made_by,
                version_needed,
                flags: file.flags,
                compression_method: file.compression_method.as_id(),
                last_mod_time: dos_time,
                last_mod_date: dos_date,
                crc32: file.crc,
                compressed_size: file.compressed_size.min(ZIP64_THRESHOLD_FILE_SIZE) as u32,
                uncompressed_size: file.uncompressed_size.min(ZIP64_THRESHOLD_FILE_SIZE) as u32,
                file_name_len: file.name_len,
                extra_field_len: file.extra_fields.central_size,
                file_comment_len: 0,
                disk_number_start: 0,
                internal_file_attrs: 0,
                external_file_attrs: file.unix_permissions.map(|x| x << 16).unwrap_or(0),
                local_header_offset: file.local_header_offset.min(ZIP64_THRESHOLD_OFFSET) as u32,
            };

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

        // Write ZIP64 structures if needed
        if needs_zip64 {
            let zip64_eocd_offset = self.writer.count();

            // Write ZIP64 End of Central Directory Record
            write_zip64_eocd(
                &mut self.writer,
                total_entries as u64,
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

        // Number of entries - use 0xFFFF if ZIP64
        let entries_count = total_entries.min(ZIP64_THRESHOLD_ENTRIES) as u16;
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
    name_len: u16,
    local_header_offset: u64,
    compression_method: CompressionMethod,
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
        output.compressed_size = self.compressed_bytes;
        let mut buffer = [0u8; 24];
        buffer[0..4].copy_from_slice(&DataDescriptor::SIGNATURE.to_le_bytes());
        buffer[4..8].copy_from_slice(&output.crc.to_le_bytes());

        let out_data = if output.compressed_size >= ZIP64_THRESHOLD_FILE_SIZE
            || output.uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE
        {
            // Use 64-bit sizes for ZIP64
            buffer[8..16].copy_from_slice(&output.compressed_size.to_le_bytes());
            buffer[16..24].copy_from_slice(&output.uncompressed_size.to_le_bytes());
            &buffer[..]
        } else {
            // Use 32-bit sizes for standard ZIP
            buffer[8..12].copy_from_slice(&(output.compressed_size as u32).to_le_bytes());
            buffer[12..16].copy_from_slice(&(output.uncompressed_size as u32).to_le_bytes());
            &buffer[..16]
        };

        self.inner.writer.write_all(out_data)?;

        let mut file_header = FileHeader {
            name_len: self.name_len,
            compression_method: self.compression_method,
            local_header_offset: self.local_header_offset,
            compressed_size: output.compressed_size,
            uncompressed_size: output.uncompressed_size,
            crc: output.crc,
            flags: self.flags,
            modification_time: self.modification_time,
            unix_permissions: self.unix_permissions,
            extra_fields: self.extra_fields,
        };
        file_header.finalize_extra_fields()?;
        self.inner.files.push(file_header);

        Ok(self.compressed_bytes)
    }
}

impl<W> Write for ZipEntryWriter<'_, W>
where
    W: Write,
{
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        let bytes_written = self.inner.writer.write(buf)?;
        self.compressed_bytes += bytes_written as u64;
        Ok(bytes_written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.writer.flush()
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
        let bytes_written = self.inner.write(buf)?;
        self.uncompressed_bytes += bytes_written as u64;

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
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

impl FileHeader {
    fn needs_zip64(&self) -> bool {
        self.compressed_size >= ZIP64_THRESHOLD_FILE_SIZE
            || self.uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE
            || self.local_header_offset >= ZIP64_THRESHOLD_OFFSET
    }

    fn finalize_extra_fields(&mut self) -> Result<(), Error> {
        if self.needs_zip64() {
            let mut sink = [0u8; 24];
            let mut pos = 0;
            if self.uncompressed_size >= ZIP64_THRESHOLD_FILE_SIZE {
                sink[pos..pos + 8].copy_from_slice(&self.uncompressed_size.to_le_bytes());
                pos += 8;
            }
            if self.compressed_size >= ZIP64_THRESHOLD_FILE_SIZE {
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
    modification_time: Option<UtcDateTime>,
    unix_permissions: Option<u32>,
    extra_fields: ExtraFieldsContainer,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ZipArchive;
    use std::io::Cursor;

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
