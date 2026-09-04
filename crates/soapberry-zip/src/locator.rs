use crate::errors::{Error, ErrorKind};
use crate::reader_at::{FileReader, ReaderAtExt};
use crate::utils::{le_u16, le_u32, le_u64};
use crate::{
    END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE, ReaderAt, Zip64EndOfCentralDirectory,
    Zip64EndOfCentralDirectoryRecord, ZipArchive, ZipFileHeaderFixed, ZipSliceArchive,
};
use std::cell::RefCell;
use std::fs::File;
use std::io::Seek;

const END_OF_CENTRAL_DIR_SIGNAUTRE: u32 = 0x06054b50;
pub(crate) const END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES: [u8; 4] =
    END_OF_CENTRAL_DIR_SIGNAUTRE.to_le_bytes();

// https://github.com/zlib-ng/minizip-ng/blob/55db144e03027b43263e5ebcb599bf0878ba58de/mz_zip.c#L78
const END_OF_CENTRAL_DIR_MAX_OFFSET: u64 = 1 << 20;

fn reject_multi_disk_zip() -> Error {
    Error::from(ErrorKind::InvalidInput {
        msg: "multi-disk ZIP archives are not supported".to_string(),
    })
}

fn validate_classic_single_disk(
    record: &EndOfCentralDirectoryRecordFixed,
    is_zip64: bool,
) -> std::result::Result<(), Error> {
    if !is_zip64 {
        if record.disk_number != 0
            || record.eocd_disk != 0
            || record.num_entries != record.total_entries
        {
            return Err(reject_multi_disk_zip());
        }
    } else if (record.disk_number != 0 && record.disk_number != u16::MAX)
        || (record.eocd_disk != 0 && record.eocd_disk != u16::MAX)
    {
        return Err(reject_multi_disk_zip());
    }
    Ok(())
}

fn validate_zip64_locator_single_disk(
    locator: &Zip64EndOfCentralDirectoryLocatorRecord,
) -> std::result::Result<(), Error> {
    if locator.eocd_disk != 0 || locator.total_disks != 1 {
        return Err(reject_multi_disk_zip());
    }
    Ok(())
}

fn validate_zip64_record_single_disk(
    record: &Zip64EndOfCentralDirectoryRecord,
) -> std::result::Result<(), Error> {
    if record.disk_number != 0 || record.cd_disk != 0 || record.num_entries != record.total_entries
    {
        return Err(reject_multi_disk_zip());
    }
    Ok(())
}

fn validate_zip64_record_fits_before_locator(
    record_offset: u64,
    locator_offset: u64,
) -> std::result::Result<(), Error> {
    let fixed_record_end = record_offset
        .checked_add(Zip64EndOfCentralDirectoryRecord::SIZE as u64)
        .ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "ZIP64 end-of-central-directory offset overflows".to_string(),
            })
        })?;
    if fixed_record_end > locator_offset {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "ZIP64 end-of-central-directory record does not fit before its locator"
                .to_string(),
        }));
    }
    Ok(())
}

fn validate_zip64_record_layout(
    record: &Zip64EndOfCentralDirectoryRecord,
    record_offset: u64,
    locator_offset: u64,
) -> std::result::Result<(), Error> {
    let fixed_payload_size = (Zip64EndOfCentralDirectoryRecord::SIZE - 12) as u64;
    if record.size < fixed_payload_size {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "ZIP64 end-of-central-directory record is too short".to_string(),
        }));
    }

    let record_end = record_offset
        .checked_add(12)
        .and_then(|offset| offset.checked_add(record.size))
        .ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "ZIP64 end-of-central-directory record length overflows".to_string(),
            })
        })?;
    if record_end != locator_offset {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "ZIP64 end-of-central-directory record is not adjacent to its locator".to_string(),
        }));
    }

    Ok(())
}

/// Locates the End of Central Directory (EOCD) record in a ZIP archive.
///
/// The `ZipLocator` is responsible for finding the EOCD record, which is
/// crucial for reading the contents of a ZIP file.
///
/// In the event, that the comment or tailing data contains the EOCD signature,
/// causing the zip locator to fail to parse. One can reparse the data starting
/// from the false EOCD offset using the reported offset
/// [`Error::eocd_offset()`]
#[derive(Debug)]
pub struct ZipLocator {
    max_search_space: u64,
}

impl Default for ZipLocator {
    fn default() -> Self {
        Self::new()
    }
}

impl ZipLocator {
    /// Creates a new `ZipLocator` with a default maximum search space of 1 MiB
    pub fn new() -> Self {
        ZipLocator {
            max_search_space: END_OF_CENTRAL_DIR_MAX_OFFSET,
        }
    }

    /// Sets the maximum number of bytes to search for the EOCD signature.
    ///
    /// The search is performed backwards from the end of the data source.
    ///
    /// ```rust
    /// use soapberry_zip::ZipLocator;
    ///
    /// let locator = ZipLocator::new().max_search_space(1024 * 64); // 64 KiB
    /// ```
    pub fn max_search_space(mut self, max_search_space: u64) -> Self {
        self.max_search_space = max_search_space;
        self
    }

    fn locate_in_byte_slice(&self, data: &[u8]) -> Result<EndOfCentralDirectory, Error> {
        let max_search_space = usize::try_from(self.max_search_space).unwrap_or(usize::MAX);
        let location = find_end_of_central_dir_signature(data, max_search_space)
            .ok_or(ErrorKind::MissingEndOfCentralDirectory)?;

        let mut eocd = self
            .locate_in_byte_slice_impl(data, location)
            .map_err(|e| e.with_eocd_offset(location as u64))?;
        eocd.validate_source(data.len())?;

        // Transparently verify that the self reported central directory points
        // to a valid entry. If it is not a valid entry, we can attempt to
        // correct offsets when there is undeclared prelude data by testing if
        // the central directory directly precedes the end of central directory
        // marker, which should hold true in the vast majority of cases. If both
        // checks fail, defer returning an error until the user explicitly wants
        // to iterate through the central directory.
        let first_entry = usize::try_from(eocd.central_dir_offset)
            .ok()
            .and_then(|offset| data.get(offset..))
            .filter(|d| ZipFileHeaderFixed::parse(d).is_ok());

        match first_entry {
            None => {
                let cd_offset = eocd
                    .head_eocd_offset()
                    .checked_sub(eocd.central_dir_size)
                    .ok_or_else(|| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;

                let first_entry = usize::try_from(cd_offset)
                    .ok()
                    .and_then(|offset| data.get(offset..))
                    .filter(|d| ZipFileHeaderFixed::parse(d).is_ok());

                if first_entry.is_some() {
                    eocd.base_offset = cd_offset.saturating_sub(eocd.central_dir_offset);
                    eocd.central_dir_offset = cd_offset;
                    eocd.validate_source(data.len())?;
                }

                Ok(eocd)
            },
            _ => Ok(eocd),
        }
    }

    fn locate_in_byte_slice_impl(
        &self,
        data: &[u8],
        location: usize,
    ) -> Result<EndOfCentralDirectory, Error> {
        let eocd = EndOfCentralDirectoryRecordFixed::parse(&data[location..])?;
        let is_zip64 = eocd.is_zip64();
        if is_zip64 {
            validate_classic_single_disk(&eocd, is_zip64)?;
        }
        let eocd_fixed = eocd.clone();
        let eocd = EndOfCentralDirectoryRecord::from_parts(location as u64, eocd);

        // Validate comment is completely present in the slice
        let comment_start = location
            .checked_add(EndOfCentralDirectoryRecordFixed::SIZE)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let comment_len = eocd.comment_len as usize;
        let comment_end = comment_start
            .checked_add(comment_len)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        if comment_end > data.len() {
            return Err(Error::from(ErrorKind::Eof));
        }

        if !is_zip64 {
            let eocd = EndOfCentralDirectory::create(eocd)?;
            validate_classic_single_disk(&eocd_fixed, false)?;
            return Ok(eocd);
        }

        let locator_offset = location
            .checked_sub(Zip64EndOfCentralDirectoryLocatorRecord::SIZE)
            .ok_or_else(|| Error::from(ErrorKind::MissingZip64EndOfCentralDirectory))?;
        let zip64l = &data[locator_offset..];
        let zip64_locator = Zip64EndOfCentralDirectoryLocatorRecord::parse(zip64l)?;
        validate_zip64_locator_single_disk(&zip64_locator)?;
        validate_zip64_record_fits_before_locator(
            zip64_locator.directory_offset,
            locator_offset as u64,
        )?;
        let zip64_offset = usize::try_from(zip64_locator.directory_offset)
            .map_err(|_| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;
        let zip64_eocd = data
            .get(zip64_offset..)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let (zip64_record_offset, zip64_record) =
            match Zip64EndOfCentralDirectoryRecord::parse(zip64_eocd) {
                Ok(record) => (zip64_locator.directory_offset, record),
                Err(error) => {
                    let physical_offset = (locator_offset as u64)
                        .checked_sub(Zip64EndOfCentralDirectoryRecord::SIZE as u64)
                        .ok_or_else(|| {
                            Error::from(ErrorKind::InvalidInput {
                                msg: "ZIP64 end-of-central-directory record is before its locator"
                                    .to_string(),
                            })
                        })?;
                    if physical_offset == zip64_locator.directory_offset {
                        return Err(error);
                    }
                    validate_zip64_record_fits_before_locator(
                        physical_offset,
                        locator_offset as u64,
                    )?;
                    let physical_offset = usize::try_from(physical_offset)
                        .map_err(|_| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;
                    let physical_eocd = data
                        .get(physical_offset..)
                        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
                    (
                        physical_offset as u64,
                        Zip64EndOfCentralDirectoryRecord::parse(physical_eocd)?,
                    )
                },
            };
        validate_zip64_record_single_disk(&zip64_record)?;
        validate_zip64_record_layout(&zip64_record, zip64_record_offset, locator_offset as u64)?;

        let zip64 = Zip64EndOfCentralDirectory::from_parts(zip64_record_offset, zip64_record);
        EndOfCentralDirectory::create_zip64(eocd, zip64)
    }

    /// Locates the EOCD record within a byte slice.
    ///
    /// On success, returns a `ZipSliceArchive` which allows reading the archive
    /// directly from the slice. On failure, returns the original slice and an `Error`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use soapberry_zip::ZipLocator;
    /// use std::fs;
    /// use std::io::Read;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let mut file = fs::File::open("assets/readme.zip")?;
    /// let mut data = Vec::new();
    /// file.read_to_end(&mut data)?;
    ///
    /// let locator = ZipLocator::new();
    /// match locator.locate_in_slice(&data) {
    ///     Ok(archive) => {
    ///         println!("Found EOCD in slice, archive has {} files.", archive.entries_hint());
    ///     }
    ///     Err((_data, e)) => {
    ///         eprintln!("Failed to locate EOCD in slice: {:?}", e);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn locate_in_slice<T: AsRef<[u8]>>(
        &self,
        data: T,
    ) -> Result<ZipSliceArchive<T>, (T, Error)> {
        match self.locate_in_byte_slice(data.as_ref()) {
            Ok(eocd) => Ok(ZipSliceArchive::new(data, eocd)),
            Err(e) => Err((data, e)),
        }
    }

    /// Locates the EOCD record within a file.
    ///
    /// A mutable byte slice to use for reading data from the file. The buffer
    /// should be large enough to hold the EOCD record and potentially parts of
    /// the ZIP64 EOCD locator if present. A common size might be a few
    /// kilobytes.
    ///
    /// On failure, returns the original file and an `Error`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use soapberry_zip::ZipLocator;
    /// use std::fs::File;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let file = File::open("assets/readme.zip")?;
    /// let mut buffer = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    /// let locator = ZipLocator::new();
    ///
    /// match locator.locate_in_file(file, &mut buffer) {
    ///     Ok(archive) => {
    ///         println!("Found EOCD in file, archive has {} files.", archive.entries_hint());
    ///     }
    ///     Err((_file, e)) => {
    ///         eprintln!("Failed to locate EOCD in file: {:?}", e);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn locate_in_file(
        &self,
        file: std::fs::File,
        buffer: &mut [u8],
    ) -> Result<ZipArchive<FileReader>, (File, Error)> {
        let mut reader = FileReader::from(file);
        let end_offset = match reader.seek(std::io::SeekFrom::End(0)) {
            Ok(offset) => offset,
            Err(e) => return Err((reader.into_inner(), Error::from(e))),
        };
        self.locate_in_reader(reader, buffer, end_offset)
            .map_err(|(fr, e)| (fr.into_inner(), e))
    }

    /// Locates the EOCD record in a reader, treating the specified end offset
    /// as the starting point when searching backwards.
    ///
    /// This method is useful for several scenarios:
    ///
    /// - Zip archive is nowhere near the end of the reader
    /// - Zip archives are concatenated
    ///
    /// For seekable readers, you can determine the end_offset by seeking to the
    /// end of the stream.
    ///
    /// Note that the zip locator may request data passed the end offset in
    /// order to read the entire end of the central directory record + comment.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use soapberry_zip::{ZipLocator, FileReader};
    /// use std::fs::File;
    /// use std::io::Seek;
    ///
    /// # fn main() -> Result<(), soapberry_zip::Error> {
    /// let file = File::open("assets/test.zip").unwrap();
    /// let mut reader = FileReader::from(file);
    /// let mut buffer = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    /// let locator = ZipLocator::new();
    ///
    /// // An example of determining the end offset when you don't
    /// // the length but have a seekable reader.
    /// let end_offset = reader.seek(std::io::SeekFrom::End(0)).unwrap();
    /// let archive = locator.locate_in_reader(reader, &mut buffer, end_offset)
    ///     .map_err(|(_, e)| e)?;
    ///
    /// // Maybe there is another zip archive to be found.
    /// // To find where the current archive starts, we need the minimum local header
    /// // offset. Below we are being conservative and iterating through the entire central
    /// // directory for the start offset, but in reality out of order central directories
    /// // are an edge case.
    /// let zip_start = {
    ///     let mut min_offset = u64::MAX;
    ///     let mut entries = archive.entries(&mut buffer);
    ///     while let Ok(Some(entry)) = entries.next_entry() {
    ///         min_offset = min_offset.min(entry.local_header_offset());
    ///     }
    ///     if min_offset == u64::MAX { 0 } else { min_offset }
    /// };
    /// match locator.locate_in_reader(archive.get_ref(), &mut buffer, zip_start) {
    ///    Ok(previous_archive) => {
    ///        println!("Found previous ZIP archive!");
    ///    }
    ///    Err((_, _)) => println!("No previous ZIP archive found"),
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn locate_in_reader<R>(
        &self,
        mut reader: R,
        buffer: &mut [u8],
        end_offset: u64,
    ) -> Result<ZipArchive<R>, (R, Error)>
    where
        R: ReaderAt,
    {
        // A no-comment EOCD ends exactly at `end_offset`. Probe that common
        // case before reading an entire search buffer. Failed reads and rejected
        // fixed records deliberately fall through to the established backwards
        // search so comments, ZIP64, suffixes, and false signatures keep its
        // semantics. Once a fixed record is accepted, its validation result is
        // definitive, just as it would be after the backwards search finds it.
        if self.max_search_space >= EndOfCentralDirectoryRecordFixed::SIZE as u64
            && buffer.len() >= EndOfCentralDirectoryRecordFixed::SIZE
            && buffer.len() >= ZipFileHeaderFixed::SIZE
        {
            if let Some(eocd_offset) =
                end_offset.checked_sub(EndOfCentralDirectoryRecordFixed::SIZE as u64)
            {
                if reader
                    .read_exact_at(
                        &mut buffer[..EndOfCentralDirectoryRecordFixed::SIZE],
                        eocd_offset,
                    )
                    .is_ok()
                    && matches!(
                        EndOfCentralDirectoryRecordFixed::parse(
                            &buffer[..EndOfCentralDirectoryRecordFixed::SIZE]
                        ),
                        Ok(record) if record.comment_len == 0 && !record.is_zip64()
                    )
                {
                    match self.locate_in_reader_impl(
                        reader,
                        buffer,
                        eocd_offset,
                        0,
                        EndOfCentralDirectoryRecordFixed::SIZE,
                    ) {
                        Ok((reader, eocd)) => {
                            return Ok(self.finish_locate_in_reader(reader, buffer, eocd));
                        },
                        Err((reader, error)) => {
                            return Err((reader, error.with_eocd_offset(eocd_offset)));
                        },
                    }
                }
            }
        }

        let location_result =
            find_end_of_central_dir(&mut reader, buffer, self.max_search_space, end_offset);

        let (eocd_offset, buffer_pos, buffer_valid_len) = match location_result {
            Ok(Some(location_tuple)) => location_tuple,
            Ok(None) => {
                return Err((reader, Error::from(ErrorKind::MissingEndOfCentralDirectory)));
            },
            Err(error) => {
                return Err((reader, Error::io(error)));
            },
        };

        let (reader, eocd) = self
            .locate_in_reader_impl(reader, buffer, eocd_offset, buffer_pos, buffer_valid_len)
            .map_err(|(reader, e)| (reader, e.with_eocd_offset(eocd_offset)))?;

        Ok(self.finish_locate_in_reader(reader, buffer, eocd))
    }

    fn finish_locate_in_reader<R>(
        &self,
        reader: R,
        _buffer: &mut [u8],
        mut eocd: EndOfCentralDirectory,
    ) -> ZipArchive<R>
    where
        R: ReaderAt,
    {
        // Check first entry in central directory, see
        // `ZipLocator::locate_in_byte_slice` for more info
        let mut first_entry_buffer = [0u8; ZipFileHeaderFixed::SIZE];
        let first_entry = reader
            .read_exact_at(&mut first_entry_buffer, eocd.central_dir_offset)
            .ok()
            .filter(|_| ZipFileHeaderFixed::parse(&first_entry_buffer).is_ok());

        match first_entry {
            None => {
                let Some(cd_offset) = eocd.head_eocd_offset().checked_sub(eocd.central_dir_size)
                else {
                    return ZipArchive::new(reader, eocd);
                };

                let first_entry = reader
                    .read_exact_at(&mut first_entry_buffer, cd_offset)
                    .ok()
                    .filter(|_| ZipFileHeaderFixed::parse(&first_entry_buffer).is_ok());

                if first_entry.is_some() {
                    eocd.base_offset = cd_offset.saturating_sub(eocd.central_dir_offset);
                    eocd.central_dir_offset = cd_offset;
                }

                ZipArchive::new(reader, eocd)
            },
            _ => ZipArchive::new(reader, eocd),
        }
    }

    fn locate_in_reader_impl<R>(
        &self,
        reader: R,
        buffer: &mut [u8],
        eocd_offset: u64,
        buffer_pos: usize,
        buffer_valid_len: usize,
    ) -> Result<(R, EndOfCentralDirectory), (R, Error)>
    where
        R: ReaderAt,
    {
        // Most likely the single read to find the end of the central directory
        // will fill the buffer with entire end of the central directory (and
        // optionally zip64 end of central directory). So let's try and reuse
        // the the data already in memory as much as possible.
        let reader = Marker::new(reader);

        let mut end_of_central_directory = &buffer[buffer_pos..buffer_valid_len];
        let eocd = loop {
            match EndOfCentralDirectoryRecordFixed::parse(end_of_central_directory) {
                Ok(record) => break record,
                Err(e) if e.is_eof() => {
                    // Unhappy path: the end of central directory crossed over read boundaries
                    let read = reader.read_at_least_at(
                        buffer,
                        EndOfCentralDirectoryRecordFixed::SIZE,
                        eocd_offset,
                    );

                    let read = match read {
                        Ok(read) => read,
                        Err(e) => return Err((reader.inner, e)),
                    };

                    end_of_central_directory = &buffer[..read];
                },
                Err(e) => return Err((reader.inner, e)),
            }
        };

        let is_zip64 = eocd.is_zip64();
        if is_zip64 {
            if let Err(error) = validate_classic_single_disk(&eocd, is_zip64) {
                return Err((reader.inner, error));
            }
        }

        end_of_central_directory =
            &end_of_central_directory[EndOfCentralDirectoryRecordFixed::SIZE..];

        let comment_len = eocd.comment_len as usize;

        // Check if the rest of the buffer doesn't completely contain the comment.
        if end_of_central_directory.len() < comment_len {
            let pos = end_of_central_directory.len();
            let comment_offset = eocd_offset
                .checked_add(EndOfCentralDirectoryRecordFixed::SIZE as u64)
                .and_then(|offset| offset.checked_add(pos as u64));
            let remaining_comment_len = comment_len - pos;

            // Try to read a single byte to validate the rest of the comment is accessible
            let mut temp_buf = [0u8; 1];
            let Some(end_comment_offset) = comment_offset
                .and_then(|offset| offset.checked_add(remaining_comment_len as u64))
                .and_then(|offset| offset.checked_sub(1))
            else {
                return Err((reader.inner, Error::from(ErrorKind::Eof)));
            };
            if let Err(e) = reader.read_exact_at(&mut temp_buf, end_comment_offset) {
                return Err((reader.inner, Error::io(e)));
            }
        }

        let eocd_fixed = eocd.clone();
        let eocd = EndOfCentralDirectoryRecord::from_parts(eocd_offset, eocd);
        if !is_zip64 {
            let eocd = match EndOfCentralDirectory::create(eocd) {
                Ok(eocd) => eocd,
                Err(e) => return Err((reader.inner, e)),
            };
            if let Err(error) = validate_classic_single_disk(&eocd_fixed, false) {
                return Err((reader.inner, error));
            }
            return Ok((reader.inner, eocd));
        }

        let eocd64l_size = Zip64EndOfCentralDirectoryLocatorRecord::SIZE;

        // Unhappy path: if we needed to issue any reads since the original
        // eocd or don't have enough data in the buffer
        let eocd64l_pos = if reader.is_marked() || eocd64l_size > buffer_pos {
            if (eocd64l_size as u64) > eocd_offset {
                return Err((
                    reader.inner,
                    Error::from(ErrorKind::MissingZip64EndOfCentralDirectory),
                ));
            }

            let read = reader.read_exact_at(
                &mut buffer[..eocd64l_size],
                eocd_offset - eocd64l_size as u64,
            );

            match read {
                Ok(_) => 0,
                Err(e) => return Err((reader.inner, Error::io(e))),
            }
        } else {
            buffer_pos - eocd64l_size
        };

        let zip64l_eocd = &buffer[eocd64l_pos..eocd64l_pos + eocd64l_size];
        let zip64_locator = match Zip64EndOfCentralDirectoryLocatorRecord::parse(zip64l_eocd) {
            Ok(locator) => locator,
            Err(e) => return Err((reader.inner, e)),
        };
        if let Err(error) = validate_zip64_locator_single_disk(&zip64_locator) {
            return Err((reader.inner, error));
        }

        let locator_offset = match eocd_offset.checked_sub(eocd64l_size as u64) {
            Some(offset) => offset,
            None => {
                return Err((
                    reader.inner,
                    Error::from(ErrorKind::InvalidInput {
                        msg: "ZIP64 locator offset is before the reader start".to_string(),
                    }),
                ));
            },
        };
        if let Err(error) = validate_zip64_record_fits_before_locator(
            zip64_locator.directory_offset,
            locator_offset,
        ) {
            return Err((reader.inner, error));
        }

        let zip64_eocd_fixed_size = Zip64EndOfCentralDirectoryRecord::SIZE;
        let buffered_record_start = if buffer.len() >= zip64_eocd_fixed_size
            && !reader.is_marked()
            && zip64_locator.directory_offset <= eocd_offset
        {
            eocd_offset
                .checked_sub(zip64_locator.directory_offset)
                .and_then(|distance| usize::try_from(distance).ok())
                .and_then(|distance| buffer_pos.checked_sub(distance))
                .filter(|start| {
                    start
                        .checked_add(zip64_eocd_fixed_size)
                        .is_some_and(|end| end <= buffer_valid_len)
                })
        } else {
            None
        };

        let mut zip64_record_offset = zip64_locator.directory_offset;
        let mut zip64_eocd = [0u8; Zip64EndOfCentralDirectoryRecord::SIZE];
        let zip64_record_bytes = if buffer.len() < zip64_eocd_fixed_size {
            if let Err(e) = reader.read_exact_at(&mut zip64_eocd, zip64_record_offset) {
                return Err((reader.inner, Error::io(e)));
            }
            &zip64_eocd[..]
        } else if let Some(start) = buffered_record_start {
            &buffer[start..start + zip64_eocd_fixed_size]
        } else {
            let read =
                reader.try_read_at_least_at(buffer, zip64_eocd_fixed_size, zip64_record_offset);
            let read = match read {
                Ok(read) => read,
                Err(e) => return Err((reader.inner, Error::io(e))),
            };
            &buffer[..read]
        };
        let zip64_record_result = Zip64EndOfCentralDirectoryRecord::parse(zip64_record_bytes);
        let zip64_record = match zip64_record_result {
            Ok(record) => record,
            Err(error) => {
                let physical_offset = match locator_offset
                    .checked_sub(Zip64EndOfCentralDirectoryRecord::SIZE as u64)
                {
                    Some(offset) => offset,
                    None => {
                        return Err((
                            reader.inner,
                            Error::from(ErrorKind::InvalidInput {
                                msg: "ZIP64 end-of-central-directory record is before its locator"
                                    .to_string(),
                            }),
                        ));
                    },
                };
                if physical_offset == zip64_locator.directory_offset {
                    return Err((reader.inner, error));
                }
                if let Err(error) =
                    validate_zip64_record_fits_before_locator(physical_offset, locator_offset)
                {
                    return Err((reader.inner, error));
                }
                zip64_record_offset = physical_offset;
                if let Err(e) = reader.read_exact_at(&mut zip64_eocd, physical_offset) {
                    return Err((reader.inner, Error::io(e)));
                }
                match Zip64EndOfCentralDirectoryRecord::parse(&zip64_eocd) {
                    Ok(record) => record,
                    Err(e) => return Err((reader.inner, e)),
                }
            },
        };

        if let Err(error) = validate_zip64_record_single_disk(&zip64_record) {
            return Err((reader.inner, error));
        }

        if let Err(error) =
            validate_zip64_record_layout(&zip64_record, zip64_record_offset, locator_offset)
        {
            return Err((reader.inner, error));
        }

        // todo: zip64 extensible data sector

        let zip_eocd = Zip64EndOfCentralDirectory::from_parts(zip64_record_offset, zip64_record);
        match EndOfCentralDirectory::create_zip64(eocd, zip_eocd) {
            Ok(eocd) => Ok((reader.inner, eocd)),
            Err(e) => Err((reader.inner, e)),
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct EndOfCentralDirectory {
    eocd_offset: u64,
    // `0` is a valid ZIP64 EOCD offset for an empty archive.  Do not encode
    // this as `NonZeroU64`: the locator's offset is an ordinary u64 and the
    // zero value must remain distinguishable from the ZIP32/no-record case.
    zip64_eocd_offset: Option<u64>,
    central_dir_size: u64,
    central_dir_offset: u64,
    num_entries: u64,
    comment_len: u16,
    base_offset: u64,
}

impl EndOfCentralDirectory {
    pub(crate) fn create(eocd: EndOfCentralDirectoryRecord) -> Result<Self, Error> {
        let result = EndOfCentralDirectory {
            eocd_offset: eocd.offset,
            zip64_eocd_offset: None,
            central_dir_size: u64::from(eocd.central_dir_size),
            central_dir_offset: u64::from(eocd.central_dir_offset),
            num_entries: u64::from(eocd.num_entries),
            comment_len: eocd.comment_len,
            base_offset: 0,
        };

        result.validate()?;
        Ok(result)
    }

    pub(crate) fn create_zip64(
        eocd: EndOfCentralDirectoryRecord,
        zip64: Zip64EndOfCentralDirectory,
    ) -> Result<Self, Error> {
        let result = EndOfCentralDirectory {
            eocd_offset: eocd.offset,
            zip64_eocd_offset: Some(zip64.offset),
            central_dir_size: zip64.central_dir_size,
            central_dir_offset: zip64.central_dir_offset,
            num_entries: zip64.num_entries,
            comment_len: eocd.comment_len,
            base_offset: 0,
        };

        result.validate()?;
        Ok(result)
    }

    fn validate(&self) -> Result<(), Error> {
        let head_eocd_offset = self.head_eocd_offset();
        let tail_eocd_offset = self.tail_eocd_offset();
        let Some(central_directory_end) =
            self.central_dir_offset.checked_add(self.central_dir_size)
        else {
            return Err(Error::from(ErrorKind::InvalidEndOfCentralDirectory));
        };

        // The ZIP64 EOCD, when present, must precede the terminal EOCD. The
        // central directory itself must fit entirely between its start and
        // the first EOCD record; accepting only the start offset leaves
        // malformed sizes to panic later while constructing slice ranges.
        if head_eocd_offset > tail_eocd_offset || central_directory_end > head_eocd_offset {
            return Err(Error::from(ErrorKind::InvalidEndOfCentralDirectory));
        }

        // Keep comment-end arithmetic checked even for reader-backed inputs;
        // slice-backed inputs additionally validate the source length below.
        let Some(_) = self
            .eocd_offset
            .checked_add(EndOfCentralDirectoryRecordFixed::SIZE as u64)
            .and_then(|offset| offset.checked_add(self.comment_len as u64))
        else {
            return Err(Error::from(ErrorKind::InvalidEndOfCentralDirectory));
        };

        Ok(())
    }

    fn validate_source(&self, source_len: usize) -> Result<(), Error> {
        let source_len = u64::try_from(source_len)
            .map_err(|_| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;
        let central_directory_end = self
            .central_dir_offset
            .checked_add(self.central_dir_size)
            .ok_or_else(|| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;
        let eocd_end = self
            .eocd_offset
            .checked_add(EndOfCentralDirectoryRecordFixed::SIZE as u64)
            .and_then(|offset| offset.checked_add(self.comment_len as u64))
            .ok_or_else(|| Error::from(ErrorKind::InvalidEndOfCentralDirectory))?;

        if central_directory_end > self.head_eocd_offset()
            || eocd_end > source_len
            || self.head_eocd_offset() > source_len
        {
            return Err(Error::from(ErrorKind::InvalidEndOfCentralDirectory));
        }

        Ok(())
    }

    #[inline]
    pub(crate) fn directory_end_offset(&self) -> u64 {
        // `validate` is called before an EndOfCentralDirectory is stored in an
        // archive. Saturating keeps this accessor panic-free if an internal
        // caller ever constructs one without validation.
        self.central_dir_offset
            .saturating_add(self.central_dir_size)
            .min(self.head_eocd_offset())
    }

    #[inline]
    pub(crate) fn is_zip64(&self) -> bool {
        self.zip64_eocd_offset.is_some()
    }

    pub(crate) fn base_offset(&self) -> u64 {
        self.base_offset
    }

    /// The first end of the central directory signature offsets.
    ///
    /// This is offset where no new central directory records are expected.
    ///
    /// Will be equivalent to [`Self::tail_eocd_offset`] eocd for non-zip64 files
    #[inline]
    pub(crate) fn head_eocd_offset(&self) -> u64 {
        self.zip64_eocd_offset.unwrap_or(self.eocd_offset)
    }

    /// The last end of the central directory signature offsets.
    ///
    /// This will always be the byte offset of 0x06054b50
    #[inline]
    pub(crate) fn tail_eocd_offset(&self) -> u64 {
        self.eocd_offset
    }

    #[inline]
    pub(crate) fn central_directory_size(&self) -> u64 {
        self.central_dir_size
    }

    /// offset of the start of the central directory
    #[inline]
    pub(crate) fn directory_offset(&self) -> u64 {
        self.central_dir_offset
    }

    #[inline]
    pub(crate) fn entries(&self) -> u64 {
        self.num_entries
    }

    #[inline]
    pub(crate) fn comment_len(&self) -> usize {
        self.comment_len as usize
    }
}

struct Marker<T> {
    inner: T,
    marked: RefCell<bool>,
}

impl<T> Marker<T> {
    fn new(inner: T) -> Self {
        Self {
            inner,
            marked: RefCell::new(false),
        }
    }

    fn is_marked(&self) -> bool {
        *self.marked.borrow()
    }
}

impl<T> ReaderAt for Marker<T>
where
    T: ReaderAt,
{
    fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
        match self.inner.read_at(buf, offset) {
            Ok(n) if n > 0 => {
                *self.marked.borrow_mut() = true;
                Ok(n)
            },
            x => x,
        }
    }
}

impl<T> std::io::Seek for Marker<T>
where
    T: std::io::Seek,
{
    fn seek(&mut self, pos: std::io::SeekFrom) -> std::io::Result<u64> {
        self.inner.seek(pos)
    }
}

/// A non-zip64 end of central directory
#[derive(Debug, Clone)]
pub(crate) struct EndOfCentralDirectoryRecord {
    pub(crate) offset: u64,
    pub(crate) central_dir_size: u32,
    pub(crate) central_dir_offset: u32,
    pub(crate) num_entries: u16,
    pub(crate) comment_len: u16,
}

impl EndOfCentralDirectoryRecord {
    #[inline]
    pub fn from_parts(offset: u64, eocd: EndOfCentralDirectoryRecordFixed) -> Self {
        Self {
            offset,
            central_dir_size: eocd.central_dir_size,
            central_dir_offset: eocd.central_dir_offset,
            num_entries: eocd.total_entries,
            comment_len: eocd.comment_len,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct EndOfCentralDirectoryRecordFixed {
    pub(crate) signature: u32,
    #[allow(dead_code)]
    pub(crate) disk_number: u16,
    #[allow(dead_code)]
    pub(crate) eocd_disk: u16,
    pub(crate) num_entries: u16,
    pub(crate) total_entries: u16,
    pub(crate) central_dir_size: u32,
    pub(crate) central_dir_offset: u32,
    pub(crate) comment_len: u16,
}

impl EndOfCentralDirectoryRecordFixed {
    pub(crate) const SIZE: usize = 22;
    pub fn parse(data: &[u8]) -> Result<EndOfCentralDirectoryRecordFixed, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let result = EndOfCentralDirectoryRecordFixed {
            signature: le_u32(&data[0..4]),
            disk_number: le_u16(&data[4..6]),
            eocd_disk: le_u16(&data[6..8]),
            num_entries: le_u16(&data[8..10]),
            total_entries: le_u16(&data[10..12]),
            central_dir_size: le_u32(&data[12..16]),
            central_dir_offset: le_u32(&data[16..20]),
            comment_len: le_u16(&data[20..22]),
        };

        if result.signature != END_OF_CENTRAL_DIR_SIGNAUTRE {
            return Err(Error::from(ErrorKind::InvalidSignature {
                expected: END_OF_CENTRAL_DIR_SIGNAUTRE,
                actual: result.signature,
            }));
        }

        Ok(result)
    }

    pub fn is_zip64(&self) -> bool {
        // https://github.com/zlib-ng/minizip-ng/blob/55db144e03027b43263e5ebcb599bf0878ba58de/mz_zip.c#L1011
        // The classic EOCD has four independent ZIP64 sentinels.  A partial
        // sentinel set is still ZIP64 and must resolve through the ZIP64
        // record; checking only the per-disk count and offset loses valid
        // archives whose total count or central-directory size overflowed the
        // ZIP32 representation first.
        self.num_entries == u16::MAX // 4.4.22
            || self.total_entries == u16::MAX // 4.4.23
            || self.central_dir_size == u32::MAX // 4.4.23
            || self.central_dir_offset == u32::MAX // 4.4.24
    }
}

///
///
/// 4.3.15
#[derive(Debug)]
#[allow(dead_code)]
struct Zip64EndOfCentralDirectoryLocatorRecord {
    /// zip64 end of central dir locator signature
    pub signature: u32,

    /// number of the disk with the start of the zip64 end of central directory
    pub eocd_disk: u32,

    /// relative offset of the zip64 end of central directory record
    pub directory_offset: u64,

    /// total number of disks
    pub total_disks: u32,
}

impl Zip64EndOfCentralDirectoryLocatorRecord {
    const SIZE: usize = 20;

    pub fn parse(data: &[u8]) -> Result<Zip64EndOfCentralDirectoryLocatorRecord, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let result = Zip64EndOfCentralDirectoryLocatorRecord {
            signature: le_u32(&data[0..4]),
            eocd_disk: le_u32(&data[4..8]),
            directory_offset: le_u64(&data[8..16]),
            total_disks: le_u32(&data[16..20]),
        };

        if result.signature != END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE {
            return Err(Error::from(ErrorKind::InvalidSignature {
                expected: END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE,
                actual: result.signature,
            }));
        }

        Ok(result)
    }
}

pub(crate) fn find_end_of_central_dir_signature(
    data: &[u8],
    max_search_space: usize,
) -> Option<usize> {
    let start_search = data.len().saturating_sub(max_search_space);
    backwards_find(
        &data[start_search..],
        &END_OF_CENTRAL_DIR_SIGNAUTRE.to_le_bytes(),
    )
    .map(|pos| pos + start_search)
}

pub(crate) fn find_end_of_central_dir<T>(
    reader: T,
    buffer: &mut [u8],
    max_search_space: u64,
    end_offset: u64,
) -> std::io::Result<Option<(u64, usize, usize)>>
where
    T: ReaderAt,
{
    if buffer.len() < END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES.len() {
        debug_assert!(false, "buffer not big enough to hold signature");
        return Ok(None);
    }

    let max_back = end_offset.saturating_sub(max_search_space);
    let mut offset = end_offset;

    // The amount of data the remains in the stream
    let mut remaining = end_offset - max_back;

    // The number of bytes that were translated from the front to the back
    let mut carry_over = 0;
    loop {
        // We either want to read into the entire buffer (sans the bytes that
        // were carried over from the last read). Or we want to read the remainder
        let read_size =
            (buffer.len() - carry_over).min(usize::try_from(remaining).unwrap_or(usize::MAX));

        // Need to jump back to the start of the previous read and then how much
        // we want to read
        offset -= read_size as u64;

        // reader.seek_relative(-offset)?;
        reader.read_exact_at(&mut buffer[..read_size], offset)?;
        remaining -= read_size as u64;

        let haystack = &buffer[..read_size + carry_over];
        if let Some(i) = backwards_find(haystack, &END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES) {
            let eocd_offset = (max_back + remaining) + (i as u64);
            return Ok(Some((eocd_offset, i, read_size + carry_over)));
        }

        if remaining == 0 {
            return Ok(None);
        }

        // Since the signature may be across read boundaries, match how much the
        // end of the signature matches the start of the buffer
        carry_over = match buffer {
            [b0, b1, b2, ..] if [*b0, *b1, *b2] == END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES[1..4] => 3,
            [b0, b1, ..] if [*b0, *b1] == END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES[2..4] => 2,
            [b0, ..] if *b0 == END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES[3] => 1,
            _ => 0,
        };

        if carry_over > 0 {
            // place the carry over bytes at the end of the buffer for the next read
            let remaining_for_buffer = usize::try_from(remaining).unwrap_or(usize::MAX);
            let dest = (buffer.len() - carry_over).min(remaining_for_buffer);
            buffer.copy_within(..carry_over, dest);
        }
    }
}

fn backwards_find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .rposition(|window| window == needle)
}

#[cfg(test)]
mod tests {
    use super::*;
    use quickcheck_macros::quickcheck;
    use rstest::rstest;
    use std::{cell::RefCell, io::Cursor};

    #[derive(Debug)]
    struct CountingReader {
        data: Vec<u8>,
        reads: RefCell<Vec<(u64, usize)>>,
    }

    impl CountingReader {
        fn new(data: Vec<u8>) -> Self {
            Self {
                data,
                reads: RefCell::new(Vec::new()),
            }
        }

        fn read_log(self) -> Vec<(u64, usize)> {
            self.reads.into_inner()
        }
    }

    impl ReaderAt for CountingReader {
        fn read_at(&self, buf: &mut [u8], offset: u64) -> std::io::Result<usize> {
            self.reads.borrow_mut().push((offset, buf.len()));
            let data = self.data.get(offset as usize..).unwrap_or_default();
            let len = data.len().min(buf.len());
            buf[..len].copy_from_slice(&data[..len]);
            Ok(len)
        }
    }

    #[cfg(target_pointer_width = "32")]
    #[derive(Debug)]
    struct LargeWindowReader {
        reads: std::cell::Cell<u8>,
    }

    #[cfg(target_pointer_width = "32")]
    impl ReaderAt for LargeWindowReader {
        fn read_at(&self, buf: &mut [u8], _offset: u64) -> std::io::Result<usize> {
            let call = self.reads.get();
            self.reads.set(call.saturating_add(1));
            if call >= 2 {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    "large-window regression issued an unexpected read",
                ));
            }

            buf.fill(0);
            if call == 0 {
                buf[..3].copy_from_slice(&END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES[1..]);
            } else {
                buf[4] = END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES[0];
            }
            Ok(buf.len())
        }
    }

    #[cfg(target_pointer_width = "32")]
    #[test]
    fn find_end_of_central_dir_saturates_large_remaining_window() {
        let end_offset = u64::from(u32::MAX) + 8;
        let mut buffer = [0u8; 8];
        let result = find_end_of_central_dir(
            LargeWindowReader {
                reads: std::cell::Cell::new(0),
            },
            &mut buffer,
            u64::MAX,
            end_offset,
        )
        .expect("synthetic large-window reader should locate the signature");

        assert!(result.is_some());
    }

    fn central_directory_entry() -> Vec<u8> {
        let mut entry = vec![0; ZipFileHeaderFixed::SIZE];
        entry[..4].copy_from_slice(&0x0201_4b50u32.to_le_bytes());
        entry
    }

    fn append_eocd(
        data: &mut Vec<u8>,
        entries: u16,
        central_dir_size: u32,
        central_dir_offset: u32,
        comment: &[u8],
    ) {
        data.extend_from_slice(&END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES);
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(&entries.to_le_bytes());
        data.extend_from_slice(&entries.to_le_bytes());
        data.extend_from_slice(&central_dir_size.to_le_bytes());
        data.extend_from_slice(&central_dir_offset.to_le_bytes());
        data.extend_from_slice(&(comment.len() as u16).to_le_bytes());
        data.extend_from_slice(comment);
    }

    fn ordinary_archive(comment: &[u8]) -> Vec<u8> {
        let mut data = central_directory_entry();
        append_eocd(&mut data, 1, ZipFileHeaderFixed::SIZE as u32, 0, comment);
        data
    }

    fn zip64_archive() -> Vec<u8> {
        let mut data = central_directory_entry();
        let zip64_eocd_offset = data.len() as u64;

        data.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        data.extend_from_slice(&44u64.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&1u64.to_le_bytes());
        data.extend_from_slice(&1u64.to_le_bytes());
        data.extend_from_slice(&(ZipFileHeaderFixed::SIZE as u64).to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());

        data.extend_from_slice(&END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&zip64_eocd_offset.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());

        append_eocd(
            &mut data,
            u16::MAX,
            ZipFileHeaderFixed::SIZE as u32,
            u32::MAX,
            &[],
        );
        data
    }

    fn zip64_empty_archive_at_zero() -> Vec<u8> {
        let mut data = Vec::new();

        // ZIP64 EOCD at the beginning of the source is valid for an empty
        // archive.  The locator below deliberately points at offset zero.
        data.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        data.extend_from_slice(&44u64.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&0u64.to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());

        data.extend_from_slice(&END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());

        append_eocd(&mut data, u16::MAX, u32::MAX, u32::MAX, &[]);
        data
    }

    fn zip64_archive_with_classic_fields(
        num_entries: u16,
        total_entries: u16,
        central_dir_size: u32,
        central_dir_offset: u32,
    ) -> Vec<u8> {
        let mut data = zip64_archive();
        let (_, _, classic_eocd) = zip64_offsets();
        data[classic_eocd + 8..classic_eocd + 10].copy_from_slice(&num_entries.to_le_bytes());
        data[classic_eocd + 10..classic_eocd + 12].copy_from_slice(&total_entries.to_le_bytes());
        data[classic_eocd + 12..classic_eocd + 16].copy_from_slice(&central_dir_size.to_le_bytes());
        data[classic_eocd + 16..classic_eocd + 20]
            .copy_from_slice(&central_dir_offset.to_le_bytes());
        data
    }

    fn zip64_offsets() -> (usize, usize, usize) {
        let zip64_eocd = ZipFileHeaderFixed::SIZE;
        let locator = zip64_eocd + Zip64EndOfCentralDirectoryRecord::SIZE;
        let classic_eocd = locator + Zip64EndOfCentralDirectoryLocatorRecord::SIZE;
        (zip64_eocd, locator, classic_eocd)
    }

    #[test]
    fn classic_eocd_zip64_detection_covers_each_count_size_and_offset_sentinel() {
        let ordinary = EndOfCentralDirectoryRecordFixed {
            signature: END_OF_CENTRAL_DIR_SIGNAUTRE,
            disk_number: 0,
            eocd_disk: 0,
            num_entries: 1,
            total_entries: 1,
            central_dir_size: 46,
            central_dir_offset: 0,
            comment_len: 0,
        };
        assert!(!ordinary.is_zip64());

        let mut per_disk_count = ordinary.clone();
        per_disk_count.num_entries = u16::MAX;
        assert!(per_disk_count.is_zip64());

        let mut total_count = ordinary.clone();
        total_count.total_entries = u16::MAX;
        assert!(total_count.is_zip64());

        let mut directory_size = ordinary.clone();
        directory_size.central_dir_size = u32::MAX;
        assert!(directory_size.is_zip64());

        let mut directory_offset = ordinary;
        directory_offset.central_dir_offset = u32::MAX;
        assert!(directory_offset.is_zip64());
    }

    #[test]
    fn zip64_empty_archive_preserves_a_zero_eocd_offset() {
        let data = zip64_empty_archive_at_zero();
        let archive = ZipLocator::new()
            .locate_in_slice(data.clone())
            .expect("offset-zero ZIP64 archive should locate");
        assert!(archive.is_zip64());
        assert_eq!(archive.entries_hint(), 0);
        assert_eq!(archive.directory_offset(), 0);
        assert_eq!(archive.head_eocd_offset(), 0);
        assert_eq!(archive.eocd_offset(), 76);
        assert!(archive.entries().next().is_none());

        let mut buffer = [0u8; 128];
        let reader_archive = ZipLocator::new()
            .locate_in_reader(
                CountingReader::new(data.clone()),
                &mut buffer,
                data.len() as u64,
            )
            .expect("reader locator should preserve offset-zero ZIP64");
        assert!(reader_archive.is_zip64());
        assert_eq!(reader_archive.entries_hint(), 0);
        assert_eq!(reader_archive.head_eocd_offset(), 0);
    }

    #[test]
    fn zip64_partial_classic_sentinels_resolve_through_the_zip64_record() {
        let ordinary = (1, 1, ZipFileHeaderFixed::SIZE as u32, 0);
        let cases = [
            (u16::MAX, ordinary.1, ordinary.2, ordinary.3),
            (ordinary.0, u16::MAX, ordinary.2, ordinary.3),
            (ordinary.0, ordinary.1, u32::MAX, ordinary.3),
            (ordinary.0, ordinary.1, ordinary.2, u32::MAX),
        ];

        for fields in cases {
            let data = zip64_archive_with_classic_fields(fields.0, fields.1, fields.2, fields.3);
            let archive = ZipLocator::new()
                .locate_in_slice(data)
                .expect("each ZIP64 classic sentinel should resolve");
            assert!(archive.is_zip64());
            assert_eq!(archive.entries_hint(), 1);
            assert_eq!(archive.directory_offset(), 0);
        }
    }

    #[test]
    fn zip64_central_directory_must_end_at_or_before_the_zip64_eocd() {
        let (zip64_eocd, _, _) = zip64_offsets();
        let mut data = zip64_archive();
        // The valid central directory occupies [0, 46), immediately before
        // the ZIP64 EOCD.  A size of 47 would intrude into that tail and must
        // be rejected before an entries iterator is constructed.
        data[zip64_eocd + 40..zip64_eocd + 48].copy_from_slice(&47u64.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(
            slice_error.kind(),
            ErrorKind::InvalidEndOfCentralDirectory
        ));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidEndOfCentralDirectory
        ));
    }

    fn locate_slice_error(data: Vec<u8>) -> Error {
        match ZipLocator::new().locate_in_slice(data) {
            Ok(_) => panic!("the synthetic archive unexpectedly located"),
            Err((_data, error)) => error,
        }
    }

    fn locate_reader_error(data: Vec<u8>) -> Error {
        let end_offset = data.len() as u64;
        let mut buffer = vec![0; 128];
        match ZipLocator::new().locate_in_reader(CountingReader::new(data), &mut buffer, end_offset)
        {
            Ok(_) => panic!("the synthetic archive unexpectedly located"),
            Err((_reader, error)) => error,
        }
    }

    fn prefixed_suffixed_zip64_archive() -> Vec<u8> {
        const PREFIX_LEN: usize = 3;

        let archive = zip64_archive();
        let mut data = vec![0xa5; PREFIX_LEN];
        data.extend_from_slice(&archive);
        data.extend_from_slice(&[0x5a; 7]);
        data
    }

    #[test]
    fn classic_eocd_rejects_mismatched_entry_counts() {
        let mut data = ordinary_archive(&[]);
        let eocd_offset = ZipFileHeaderFixed::SIZE;
        data[eocd_offset + 8..eocd_offset + 10].copy_from_slice(&0u16.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
    }

    #[test]
    fn zip64_eocd_rejects_mismatched_entry_counts() {
        let (zip64_eocd, _, _) = zip64_offsets();
        let mut data = zip64_archive();
        data[zip64_eocd + 24..zip64_eocd + 32].copy_from_slice(&0u64.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
    }

    #[test]
    fn classic_eocd_rejects_nonzero_disk_metadata() {
        for field_offset in [4, 6] {
            let mut data = Vec::new();
            append_eocd(&mut data, 0, 0, 0, &[]);
            data[field_offset..field_offset + 2].copy_from_slice(&1u16.to_le_bytes());

            let slice_error = locate_slice_error(data.clone());
            assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

            let reader_error = locate_reader_error(data);
            assert!(matches!(
                reader_error.kind(),
                ErrorKind::InvalidInput { .. }
            ));
        }
    }

    #[test]
    fn zip64_locator_rejects_non_single_disk_metadata() {
        let (_, locator, _) = zip64_offsets();
        for (field_offset, value) in [(4, 1u32), (16, 2u32)] {
            let mut data = zip64_archive();
            data[locator + field_offset..locator + field_offset + 4]
                .copy_from_slice(&value.to_le_bytes());

            let slice_error = locate_slice_error(data.clone());
            assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

            let reader_error = locate_reader_error(data);
            assert!(matches!(
                reader_error.kind(),
                ErrorKind::InvalidInput { .. }
            ));
        }
    }

    #[test]
    fn zip64_eocd_rejects_nonzero_disk_metadata() {
        let (zip64_eocd, _, _) = zip64_offsets();
        for field_offset in [16, 20] {
            let mut data = zip64_archive();
            data[zip64_eocd + field_offset..zip64_eocd + field_offset + 4]
                .copy_from_slice(&1u32.to_le_bytes());

            let slice_error = locate_slice_error(data.clone());
            assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

            let reader_error = locate_reader_error(data);
            assert!(matches!(
                reader_error.kind(),
                ErrorKind::InvalidInput { .. }
            ));
        }
    }

    #[test]
    fn reader_zip64_rejects_nonzero_locator_disk_metadata() {
        let mut data = zip64_archive();
        let (_, locator, _) = zip64_offsets();
        data[locator + 4..locator + 8].copy_from_slice(&1u32.to_le_bytes());
        let end_offset = data.len() as u64;
        let error = match ZipLocator::new().locate_in_reader(
            CountingReader::new(data),
            &mut [0; 128],
            end_offset,
        ) {
            Ok(_) => panic!("the synthetic archive unexpectedly located"),
            Err((_reader, error)) => error,
        };

        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[test]
    fn zip64_locator_disk_metadata_precedes_invalid_offset() {
        let (_, locator, _) = zip64_offsets();
        let mut data = zip64_archive();
        data[locator + 4..locator + 8].copy_from_slice(&1u32.to_le_bytes());
        data[locator + 8..locator + 16].copy_from_slice(&u64::MAX.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
    }

    #[test]
    fn zip64_eocd_rejects_invalid_record_length() {
        let (zip64_eocd, _, _) = zip64_offsets();
        let mut data = zip64_archive();
        data[zip64_eocd + 4..zip64_eocd + 12].copy_from_slice(&43u64.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
    }

    #[test]
    fn zip64_eocd_rejects_nonadjacent_locator() {
        let (_, locator, _) = zip64_offsets();
        let mut data = zip64_archive();
        data.insert(locator, 0);

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let reader_error = locate_reader_error(data);
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
    }

    #[test]
    fn zip64_invalid_offset_is_rejected_before_reader_access() {
        let (_, locator, _) = zip64_offsets();
        let mut data = zip64_archive();
        data[locator + 8..locator + 16].copy_from_slice(&u64::MAX.to_le_bytes());

        let slice_error = locate_slice_error(data.clone());
        assert!(matches!(slice_error.kind(), ErrorKind::InvalidInput { .. }));

        let end_offset = data.len() as u64;
        let mut buffer = vec![0; 128];
        let (reader, reader_error) = match ZipLocator::new().locate_in_reader(
            CountingReader::new(data),
            &mut buffer,
            end_offset,
        ) {
            Ok(_) => panic!("the synthetic archive unexpectedly located"),
            Err((reader, error)) => (reader, error),
        };
        assert!(matches!(
            reader_error.kind(),
            ErrorKind::InvalidInput { .. }
        ));
        assert!(
            reader
                .read_log()
                .iter()
                .all(|(offset, _)| *offset != u64::MAX)
        );
    }

    #[test]
    fn terminal_zip64_eocd_with_short_buffer_never_panics() {
        let data = zip64_archive();
        let end_offset = data.len() as u64;
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            ZipLocator::new().locate_in_reader(
                CountingReader::new(data),
                &mut [0; EndOfCentralDirectoryRecordFixed::SIZE],
                end_offset,
            )
        }));

        assert!(matches!(result, Ok(Ok(_))));
    }

    #[test]
    fn zip64_archive_accepts_prefix_and_suffix_in_slice_and_reader() {
        let data = prefixed_suffixed_zip64_archive();
        assert!(ZipLocator::new().locate_in_slice(data.clone()).is_ok());

        let end_offset = data.len() as u64;
        let mut buffer = vec![0; 128];
        assert!(
            ZipLocator::new()
                .locate_in_reader(CountingReader::new(data), &mut buffer, end_offset)
                .is_ok()
        );
    }

    #[test]
    fn terminal_zero_comment_eocd_probe_reads_only_the_fixed_record_first() {
        let data = ordinary_archive(&[]);
        let end_offset = data.len() as u64;
        let archive = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap();
        let reads = archive.into_inner().read_log();

        assert_eq!(reads.first(), Some(&(end_offset - 22, 22)));
        assert_eq!(reads[1], (0, ZipFileHeaderFixed::SIZE));
    }

    #[test]
    fn terminal_comment_falls_back_to_backward_search() {
        let data = ordinary_archive(b"comment");
        let end_offset = data.len() as u64;
        let archive = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap();
        let reads = archive.into_inner().read_log();

        assert_eq!(reads[0], (end_offset - 22, 22));
        assert_eq!(reads[1], (end_offset - 64, 64));
        assert_eq!(reads[2], (0, ZipFileHeaderFixed::SIZE));
    }

    #[test]
    fn terminal_zip64_eocd_falls_back_to_backward_search() {
        let data = zip64_archive();
        let end_offset = data.len() as u64;
        let archive = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 128], end_offset)
            .unwrap();
        let reads = archive.into_inner().read_log();

        assert_eq!(reads[0], (end_offset - 22, 22));
        assert_eq!(reads[1], (end_offset - 128, 128));
        assert_eq!(reads[2], (0, ZipFileHeaderFixed::SIZE));
    }

    #[test]
    fn malformed_terminal_eocd_returns_the_existing_error_offset_without_rescanning() {
        let mut data = Vec::new();
        append_eocd(&mut data, 0, 0, 1, &[]);
        let end_offset = data.len() as u64;
        let (reader, error) = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap_err();

        assert!(matches!(
            error.kind(),
            ErrorKind::InvalidEndOfCentralDirectory
        ));
        assert_eq!(error.eocd_offset(), Some(0));
        assert_eq!(reader.read_log(), vec![(0, 22)]);
    }

    #[test]
    fn explicit_end_offset_preserves_archives_with_suffixes() {
        let mut data = ordinary_archive(&[]);
        let end_offset = data.len() as u64;
        data.extend_from_slice(b"suffix");
        let archive = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap();
        let reads = archive.into_inner().read_log();

        assert_eq!(reads[0], (end_offset - 22, 22));
    }

    #[test]
    fn explicit_end_offset_preserves_the_first_of_concatenated_archives() {
        let first = ordinary_archive(&[]);
        let end_offset = first.len() as u64;
        let mut data = first;
        data.extend_from_slice(&ordinary_archive(&[]));
        let archive = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap();
        let reads = archive.into_inner().read_log();

        assert_eq!(reads[0], (end_offset - 22, 22));
    }

    #[test]
    fn insufficient_search_space_uses_the_existing_search_path() {
        let data = ordinary_archive(&[]);
        let end_offset = data.len() as u64;
        let (reader, error) = ZipLocator::new()
            .max_search_space(21)
            .locate_in_reader(CountingReader::new(data), &mut [0; 64], end_offset)
            .unwrap_err();

        assert!(matches!(
            error.kind(),
            ErrorKind::MissingEndOfCentralDirectory
        ));
        assert_eq!(reader.read_log(), vec![(end_offset - 21, 21)]);
    }

    #[test]
    fn short_buffer_uses_the_existing_search_path() {
        let data = ordinary_archive(&[]);
        let end_offset = data.len() as u64;
        let (reader, error) = ZipLocator::new()
            .locate_in_reader(CountingReader::new(data), &mut [0; 21], end_offset)
            .unwrap_err();

        assert!(matches!(error.kind(), ErrorKind::BufferTooSmall));
        assert_eq!(error.eocd_offset(), Some(ZipFileHeaderFixed::SIZE as u64));
        assert_eq!(reader.read_log().first(), Some(&(end_offset - 21, 21)));
    }

    #[test]
    fn slice_locator_rejects_central_directory_range_past_eocd_without_panicking() {
        let mut data = ordinary_archive(&[]);
        let eocd_offset = data.len() - EndOfCentralDirectoryRecordFixed::SIZE;
        // Keep this malformed archive ZIP32: `u32::MAX` is a valid ZIP64
        // sentinel and now deliberately exercises the ZIP64 path.
        data[eocd_offset + 12..eocd_offset + 16].copy_from_slice(&47u32.to_le_bytes());

        let result = std::panic::catch_unwind(|| ZipLocator::new().locate_in_slice(data));
        assert!(result.is_ok(), "malformed EOCD must not panic");
        let result = result.unwrap();
        let Err((_data, error)) = result else {
            panic!("central-directory range must be rejected before archive construction");
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::InvalidEndOfCentralDirectory
        ));
    }

    #[test]
    fn zip64_central_directory_range_overflow_is_a_typed_rejection() {
        let eocd = EndOfCentralDirectoryRecord {
            offset: u64::MAX,
            central_dir_size: 0,
            central_dir_offset: 0,
            num_entries: 0,
            comment_len: 0,
        };
        let zip64 = Zip64EndOfCentralDirectory {
            offset: u64::MAX,
            central_dir_offset: u64::MAX - 1,
            central_dir_size: 2,
            num_entries: 0,
        };

        let result = EndOfCentralDirectory::create_zip64(eocd, zip64);
        assert!(matches!(
            result,
            Err(error) if matches!(error.kind(), ErrorKind::InvalidEndOfCentralDirectory)
        ));
    }

    #[test]
    fn oversized_search_space_remains_bounded_by_the_source_without_panicking() {
        let data = ordinary_archive(&[]);
        let result = std::panic::catch_unwind(|| {
            ZipLocator::new()
                .max_search_space(u64::MAX)
                .locate_in_slice(data)
        });

        assert!(result.is_ok(), "an oversized search limit must not panic");
        assert!(result.unwrap().is_ok());
    }

    #[quickcheck]
    fn test_find_end_of_central_dir_signature(mut data: Vec<u8>, offset: usize, chunk_size: u16) {
        if data.len() < 4 {
            return;
        }

        let max_search_space = END_OF_CENTRAL_DIR_MAX_OFFSET;
        let pos = (offset % data.len()).saturating_sub(END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES.len());
        data[pos..pos + 4].copy_from_slice(&END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES);

        let result = find_end_of_central_dir_signature(&data, max_search_space as usize).unwrap();

        let mut buffer = vec![0u8; chunk_size.max(4) as usize];
        let reader = std::io::Cursor::new(&data);
        let (index, buffer_index, buffer_valid_len) =
            find_end_of_central_dir(reader, &mut buffer, max_search_space, data.len() as u64)
                .unwrap()
                .unwrap();

        assert_eq!(index, result as u64);
        assert!(buffer_valid_len > 0, "buffer_valid_len should be positive");
        assert!(
            buffer_valid_len <= buffer.len(),
            "buffer_valid_len should not exceed buffer capacity"
        );
        assert!(
            buffer_index < buffer_valid_len,
            "buffer_index should be within buffer_valid_len"
        );
        assert!(
            buffer_index + END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES.len() <= buffer_valid_len,
            "signature should be within valid part of buffer"
        );
        assert_eq!(
            buffer[buffer_index..buffer_index + 4],
            END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES
        );
    }

    #[quickcheck]
    fn test_find_end_of_central_dir_signature_random(
        data: Vec<u8>,
        chunk_size: u16,
        max_search_space: u64,
    ) {
        let mem = find_end_of_central_dir_signature(&data, max_search_space as usize);

        let mut buffer = vec![0u8; chunk_size.max(4) as usize];
        let reader = std::io::Cursor::new(&data);
        let curse =
            find_end_of_central_dir(reader, &mut buffer, max_search_space, data.len() as u64)
                .unwrap();

        let mem_result = mem.map(|x| x as u64);
        let curse_result = curse.map(|(a, _, _)| a);
        assert_eq!(mem_result, curse_result);

        if let Some((_, buffer_index, buffer_valid_len)) = curse {
            assert!(buffer_valid_len > 0, "buffer_valid_len should be positive");
            assert!(
                buffer_valid_len <= buffer.len(),
                "buffer_valid_len should not exceed buffer capacity"
            );
            assert!(
                buffer_index < buffer_valid_len,
                "buffer_index should be within buffer_valid_len"
            );
            assert!(
                buffer_index + END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES.len() <= buffer_valid_len,
                "signature should be within valid part of buffer"
            );
        }
    }

    #[rstest]
    #[case(&[], 4, 1000, None)]
    #[case(&[6], 4, 1000, None)]
    #[case(&[5, 6], 4, 1000, None)]
    #[case(&[b'K', 5, 6], 4, 1000, None)]
    #[case(&[0, 6, 0, 0, 0], 4, 1000, None)]
    #[case(&[b'P', b'K', 5, 6], 4, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6], 5, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 5, 6], 5, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 6, 0, 0, 0], 4, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 0, 0, 0, 0], 4, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 0, 0, 0], 4, 1000, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 0], 4, 1000, Some(0))]
    #[case(&[5, 6, b'P', b'K', 5, 6], 4, 1000, Some(2))]
    #[case(&[5, 6, b'P', b'K', 5, 6], 5, 1000, Some(2))]
    #[case(&[5, 6, b'P', b'K', 5, 6, 5, 6], 4, 1000, Some(2))]
    #[case(&[5, 6, b'P', b'K', 5, 6, 5, 6], 5, 1000, Some(2))]
    #[case(&[b'P', b'K', 5, 6, b'P', b'K', 5, 6, 5, 6], 5, 1000, Some(4))]
    #[case(&[b'P', b'K', 5, 6, b'P', b'K', 5, 6, 5, 6], 32, 1000, Some(4))]
    #[case(&[b'P', b'K', 5, 6], 5, 4, Some(0))] // start of max search space tests
    #[case(&[b'P', b'K', 5, 6, 5, 6], 5, 5, None)]
    #[case(&[b'P', b'K', 5, 6, 6, 0, 0, 0], 4, 8, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 0, 0, 0], 4, 8, Some(0))]
    #[case(&[b'P', b'K', 5, 6, 0], 4, 4, None)]
    #[case(&[5, 6, b'P', b'K', 5, 6], 4, 4, Some(2))]
    #[case(&[5, 6, b'P', b'K', 5, 6], 5, 4, Some(2))]
    #[case(&[5, 6, b'P', b'K', 5, 6, 5, 6], 4, 4, None)]
    #[case(&[5, 6, b'P', b'K', 5, 6, 5, 6], 5, 4, None)]
    #[case(&[b'P', b'K', 5, 6, b'P', b'K', 5, 6, 5, 6], 5, 6, Some(4))]
    #[case(&[b'P', b'K', 5, 6, b'P', b'K', 5, 6, 5, 6], 32, 10, Some(4))]
    #[test]
    fn test_find_end_of_central_dir_signature_cases(
        #[case] input: &[u8],
        #[case] buffer_size: usize,
        #[case] max_search_space: u64,
        #[case] expected: Option<u64>,
    ) {
        let result = find_end_of_central_dir_signature(input, max_search_space as usize);
        assert_eq!(result.map(|x| x as u64), expected);

        let cursor = Cursor::new(&input);
        let mut buffer = vec![0u8; buffer_size];
        let found =
            find_end_of_central_dir(cursor, &mut buffer, max_search_space, input.len() as u64)
                .unwrap();
        let found_result = found.map(|(a, _, _)| a);
        assert_eq!(found_result, expected);

        if expected.is_some() {
            let (_, buffer_pos, buffer_valid_len) = found.unwrap();
            assert!(buffer_valid_len > 0, "buffer_valid_len should be positive");
            assert!(
                buffer_valid_len <= buffer_size,
                "buffer_valid_len should not exceed buffer capacity"
            );
            assert!(
                buffer_pos < buffer_valid_len,
                "buffer_index should be within buffer_valid_len"
            );
            assert!(
                buffer_pos + END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES.len() <= buffer_valid_len,
                "signature should be within valid part of buffer"
            );
            assert_eq!(
                buffer[buffer_pos..buffer_pos + 4],
                END_OF_CENTRAL_DIR_SIGNAUTRE_BYTES
            );
        }
    }
}
