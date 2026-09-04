use crate::crc::crc32_chunk;
use crate::errors::{Error, ErrorKind};
use crate::extra_fields::{ExtraFieldId, ExtraFields};
use crate::mode::{
    CREATOR_FAT, CREATOR_MACOS, CREATOR_NTFS, CREATOR_UNIX, CREATOR_VFAT, EntryMode,
    msdos_mode_to_file_mode, unix_mode_to_file_mode,
};
use crate::path::{RawPath, ZipFilePath};
use crate::reader_at::{
    FileReader, MutexReader, RangeReader, ReaderAt, ReaderAtExt, validate_read_count,
};
use crate::time::{ZipDateTimeKind, extract_best_timestamp};
use crate::utils::{le_u16, le_u32, le_u64};
use crate::{
    EndOfCentralDirectory, EndOfCentralDirectoryRecordFixed, ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIZE,
    ZipLocator,
};
use std::io::{Read, Seek, Write};
use std::ops::Range;

pub(crate) const END_OF_CENTRAL_DIR_SIGNATURE64: u32 = 0x06064b50;
pub(crate) const END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE: u32 = 0x07064b50;
pub(crate) const CENTRAL_HEADER_SIGNATURE: u32 = 0x02014b50;
/// The recommended buffer size to use when reading from a zip file.
///
/// This buffer size was chosen as it can hold an entire central directory
/// record as the spec states (4.4.10):
///
/// > the combined length of any directory and these three fields SHOULD NOT
/// > generally exceed 65,535 bytes.
pub const RECOMMENDED_BUFFER_SIZE: usize = 1 << 16;

/// Represents a Zip archive that operates on an in-memory data.
///
/// A [`ZipSliceArchive`] is more efficient and easier to use than a [`ZipArchive`],
/// as there is no buffer management and memory copying involved.
///
/// # Examples
///
/// ```rust
/// use soapberry_zip::{ZipArchive, ZipSliceArchive, Error};
///
/// fn process_zip_slice(data: &[u8]) -> Result<(), Error> {
///     let archive = ZipArchive::from_slice(data)?;
///     println!("Found {} entries.", archive.entries_hint());
///     for entry_result in archive.entries() {
///         let entry = entry_result?;
///         println!("File: {}", entry.file_path().try_normalize()?.as_ref());
///     }
///     Ok(())
/// }
/// ```
#[derive(Debug, Clone)]
pub struct ZipSliceArchive<T: AsRef<[u8]>> {
    data: T,
    eocd: EndOfCentralDirectory,
}

impl<T: AsRef<[u8]>> ZipSliceArchive<T> {
    pub(crate) fn new(data: T, eocd: EndOfCentralDirectory) -> Self {
        ZipSliceArchive { data, eocd }
    }

    /// Returns an iterator over the entries in the central directory of the archive.
    pub fn entries(&self) -> ZipSliceEntries<'_> {
        let data = self.data.as_ref();
        let directory_start = self.eocd.directory_offset();
        let directory_end = self.eocd.directory_end_offset();
        let entry_data = usize::try_from(directory_start)
            .ok()
            .and_then(|start| {
                usize::try_from(directory_end)
                    .ok()
                    .and_then(|end| data.get(start..end))
            })
            .unwrap_or_default();
        ZipSliceEntries {
            entry_data,
            base_offset: self.eocd.base_offset(),
            current_offset: directory_start,
        }
    }

    /// Returns the byte slice that represents the zip file.
    ///
    /// This will include the entire input slice.
    pub fn as_bytes(&self) -> &[u8] {
        self.data.as_ref()
    }

    /// Returns a hint for the total number of entries in the archive.
    ///
    /// This value is read from the End of Central Directory record.
    pub fn entries_hint(&self) -> u64 {
        self.eocd.entries()
    }

    /// Returns the resolved, declared central-directory byte size.
    pub(crate) fn central_directory_size(&self) -> u64 {
        self.eocd.central_directory_size()
    }

    /// Whether this archive uses ZIP64 end-of-central-directory metadata.
    #[inline]
    pub fn is_zip64(&self) -> bool {
        self.eocd.is_zip64()
    }

    /// Returns the first EOCD offset that bounds the central directory.
    ///
    /// For ZIP32 this is the terminal EOCD offset.  For ZIP64 it is the
    /// ZIP64 EOCD offset, which keeps the central directory separate from the
    /// ZIP64 record, locator, and terminal EOCD tail.
    pub fn head_eocd_offset(&self) -> u64 {
        self.eocd.head_eocd_offset()
    }

    /// Returns the offset of the End of Central Directory (EOCD) signature.
    ///
    /// See [`ZipArchive::eocd_offset()`] for more details.
    pub fn eocd_offset(&self) -> u64 {
        self.eocd.tail_eocd_offset()
    }

    /// The declared offset of the start of the central directory.
    ///
    /// See [`ZipArchive::directory_offset()`] for more details.
    pub fn directory_offset(&self) -> u64 {
        self.eocd.directory_offset()
    }

    /// Returns the offset where the ZIP archive ends.
    ///
    /// See [`ZipArchive::end_offset`] for more details.
    pub fn end_offset(&self) -> u64 {
        self.eocd.tail_eocd_offset()
            + EndOfCentralDirectoryRecordFixed::SIZE as u64
            + self.comment().as_bytes().len() as u64
    }

    /// The comment of the zip file.
    pub fn comment(&self) -> ZipStr<'_> {
        let data = self.data.as_ref();
        let Some(comment_start) = usize::try_from(self.eocd.tail_eocd_offset())
            .ok()
            .and_then(|offset| offset.checked_add(EndOfCentralDirectoryRecordFixed::SIZE))
        else {
            return ZipStr::new(&[]);
        };
        let comment_len = self.eocd.comment_len();
        let Some(comment_end) = comment_start.checked_add(comment_len) else {
            return ZipStr::new(&[]);
        };
        ZipStr::new(data.get(comment_start..comment_end).unwrap_or_default())
    }

    /// Converts the [`ZipSliceArchive`] into a general [`ZipArchive`].
    ///
    /// This is useful for unifying code that might handle both slice-based
    /// and reader-based archives.
    #[deprecated(note = "Use `ZipSliceArchive::into_zip_archive` instead")]
    pub fn into_reader(self) -> ZipArchive<T> {
        ZipArchive {
            reader: self.data,
            eocd: self.eocd,
        }
    }

    /// Converts the [`ZipSliceArchive`] into a general [`ZipArchive`].
    ///
    /// This is useful for unifying code that might handle both slice-based and
    /// reader-based archives. The data is wrapped in a [`std::io::Cursor`] to
    /// provide the [`ReaderAt`] implementation needed for [`ZipArchive`].
    pub fn into_zip_archive(self) -> ZipArchive<std::io::Cursor<T>> {
        ZipArchive {
            reader: std::io::Cursor::new(self.data),
            eocd: self.eocd,
        }
    }

    /// Seeks to the given file entry in the zip archive.
    ///
    /// See [`ZipArchive::get_entry`] for more details. The biggest difference
    /// between the reader and slice APIs is that the slice APIs will eagerly
    /// validate that the entire compressed data is present.
    pub fn get_entry(&self, entry: ZipArchiveEntryWayfinder) -> Result<ZipSliceEntry<'_>, Error> {
        slice_entry(self.data.as_ref(), entry)
    }
}

impl<'data> ZipSliceArchive<&'data [u8]> {
    /// Seeks to the given file entry, borrowing from the archive's source
    /// slice for the source lifetime rather than the `&self` borrow.
    ///
    /// This behaves exactly like [`Self::get_entry`], but the returned entry
    /// remains valid for the complete source lifetime, which lets a caller
    /// retain borrowed member bytes beyond the archive borrow.
    /// The returned entry exposes raw compressed bytes; its CRC and
    /// uncompressed-size claims are not verified until the caller validates
    /// [`ZipSliceEntry::claim_verifier`] or consumes a verifying reader. The
    /// stricter trusted borrowed Store contract is provided by
    /// `office::ArchiveReader::read_stored_borrowed`.
    pub fn get_entry_borrowed(
        &self,
        entry: ZipArchiveEntryWayfinder,
    ) -> Result<ZipSliceEntry<'data>, Error> {
        slice_entry(self.data, entry)
    }

    /// Seeks to a stored entry and validates the local record before exposing
    /// its source slice.  The central name is supplied by the index because a
    /// wayfinder deliberately does not own a copy of the member name.
    pub(crate) fn get_stored_entry_borrowed(
        &self,
        entry: ZipArchiveEntryWayfinder,
        central_name: &[u8],
    ) -> Result<ZipSliceEntry<'data>, Error> {
        validate_single_disk_archive(self.data, &self.eocd)?;
        slice_stored_entry(
            self.data,
            entry,
            central_name,
            self.eocd.directory_offset(),
            self.eocd.is_zip64(),
        )
    }

    pub(crate) fn validate_strict_stream_target(
        &self,
        entry: ZipArchiveEntryWayfinder,
    ) -> Result<(), Error> {
        validate_borrowed_wayfinder(&entry)
    }

    pub(crate) fn validate_strict_entry_layout(
        &self,
        entry: ZipArchiveEntryWayfinder,
        central_name: &[u8],
    ) -> Result<StrictEntryLayout, Error> {
        let metadata = validate_borrowed_entry_metadata(
            self.data,
            &entry,
            central_name,
            self.eocd.directory_offset(),
            self.eocd.is_zip64(),
        )?;
        Ok(StrictEntryLayout {
            local_header_offset: u64::try_from(metadata.header_offset)
                .map_err(|_| Error::from(ErrorKind::Eof))?,
            data_start_offset: u64::try_from(metadata.data_start_offset)
                .map_err(|_| Error::from(ErrorKind::Eof))?,
            data_end_offset: u64::try_from(metadata.data_end_offset)
                .map_err(|_| Error::from(ErrorKind::Eof))?,
            span_end: u64::try_from(metadata.span_end).map_err(|_| Error::from(ErrorKind::Eof))?,
            verifier: ZipVerification {
                crc: entry.crc,
                uncompressed_size: entry.uncompressed_size,
            },
        })
    }

    pub(crate) fn strict_payload(&self, layout: StrictEntryLayout) -> Result<&'data [u8], Error> {
        let start =
            usize::try_from(layout.data_start_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
        let end =
            usize::try_from(layout.data_end_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
        self.data
            .get(start..end)
            .ok_or_else(|| Error::from(ErrorKind::Eof))
    }

    /// Validate archive-level metadata used by the borrowed Store path.
    pub(crate) fn validate_borrowed_layout(&self) -> Result<(), Error> {
        validate_single_disk_archive(self.data, &self.eocd)
    }

    /// Return the validated local span of one entry for borrowed-access
    /// overlap admission.  This does not alter ordinary owned reads.
    pub(crate) fn borrowed_entry_span(
        &self,
        entry: ZipArchiveEntryWayfinder,
        central_name: &[u8],
    ) -> Result<(u64, u64), Error> {
        let (start, end) = borrowed_entry_span(
            self.data,
            entry,
            central_name,
            self.eocd.directory_offset(),
            self.eocd.is_zip64(),
        )?;
        Ok((
            u64::try_from(start).map_err(|_| Error::from(ErrorKind::Eof))?,
            u64::try_from(end).map_err(|_| Error::from(ErrorKind::Eof))?,
        ))
    }
}

/// Locates and validates one member's local record within a slice archive.
fn slice_entry(data: &[u8], entry: ZipArchiveEntryWayfinder) -> Result<ZipSliceEntry<'_>, Error> {
    let header_offset = usize::try_from(entry.local_header_offset)
        .unwrap_or(data.len())
        .min(data.len());
    let header = &data[header_offset..];
    let file_header = ZipLocalFileHeaderFixed::parse(header)?;
    let variable_length = file_header.variable_length();

    let header_size = (ZipLocalFileHeaderFixed::SIZE + variable_length) as u32;
    let (total_size, o1) = (u64::from(header_size)).overflowing_add(entry.compressed_size_hint());

    if o1 || (header.len() as u64) < total_size {
        return Err(Error::from(ErrorKind::Eof));
    }

    let (entire_entry, rest) = header.split_at(total_size as usize);

    let expected_crc = if entry.has_data_descriptor {
        DataDescriptor::parse(rest)?.crc
    } else {
        entry.crc
    };

    Ok(ZipSliceEntry {
        data: entire_entry,
        verifier: ZipVerification {
            crc: expected_crc,
            uncompressed_size: entry.uncompressed_size_hint(),
        },
        local_header_offset: entry.local_header_offset,
        data_start_offset: header_size,
    })
}

#[derive(Debug, Clone, Copy)]
struct BorrowedEntryMetadata {
    header_offset: usize,
    data_start_offset: usize,
    data_end_offset: usize,
    span_end: usize,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct StrictEntryLayout {
    pub(crate) local_header_offset: u64,
    pub(crate) data_start_offset: u64,
    pub(crate) data_end_offset: u64,
    pub(crate) span_end: u64,
    pub(crate) verifier: ZipVerification,
}

/// Validate all metadata needed to use an entry's local span for borrowed
/// access.  Encryption eligibility is intentionally handled by the target
/// path separately so an unrelated encrypted member cannot poison the global
/// non-overlap proof.
fn validate_borrowed_entry_metadata(
    data: &[u8],
    entry: &ZipArchiveEntryWayfinder,
    central_name: &[u8],
    central_directory_offset: u64,
    archive_is_zip64: bool,
) -> Result<BorrowedEntryMetadata, Error> {
    validate_borrowed_wayfinder_metadata(entry)?;
    let header_offset =
        usize::try_from(entry.local_header_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
    let central_directory_offset =
        usize::try_from(central_directory_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
    if header_offset > central_directory_offset || central_directory_offset > data.len() {
        return Err(Error::from(ErrorKind::Eof));
    }

    let header = data
        .get(header_offset..central_directory_offset)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let file_header = ZipLocalFileHeaderFixed::parse(header)?;
    if file_header.compression_method != entry.compression_method {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central compression methods differ".to_string(),
        }));
    }
    if file_header.flags != entry.flags {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central flags differ".to_string(),
        }));
    }

    let variable_length = file_header.variable_length();
    let variable_end = ZipLocalFileHeaderFixed::SIZE
        .checked_add(variable_length)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let variable_data = header
        .get(ZipLocalFileHeaderFixed::SIZE..variable_end)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let file_name_len = usize::from(file_header.file_name_len);
    let local_name = variable_data
        .get(..file_name_len)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if local_name != central_name {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central names differ".to_string(),
        }));
    }
    let local_extra = variable_data
        .get(file_name_len..)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;

    let local_has_zip64_sentinel =
        file_header.compressed_size == u32::MAX || file_header.uncompressed_size == u32::MAX;
    if local_has_zip64_sentinel && !entry.zip64_sizes {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local header uses ZIP64 sizes absent from the central entry".to_string(),
        }));
    }
    let (local_compressed_size, local_uncompressed_size) = if local_has_zip64_sentinel {
        local_header_sizes(&file_header, local_extra)?
    } else {
        (
            u64::from(file_header.compressed_size),
            u64::from(file_header.uncompressed_size),
        )
    };

    if !entry.has_data_descriptor || local_has_zip64_sentinel {
        if local_compressed_size != entry.compressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.compressed_size,
                actual: local_compressed_size,
            }));
        }
        if local_uncompressed_size != entry.uncompressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.uncompressed_size,
                actual: local_uncompressed_size,
            }));
        }
    }
    if !entry.has_data_descriptor && file_header.crc32 != entry.crc {
        return Err(Error::from(ErrorKind::InvalidChecksum {
            expected: entry.crc,
            actual: file_header.crc32,
        }));
    }

    let compressed_size =
        usize::try_from(entry.compressed_size).map_err(|_| Error::from(ErrorKind::Eof))?;
    let payload_end_relative = variable_end
        .checked_add(compressed_size)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if payload_end_relative > header.len() {
        return Err(Error::from(ErrorKind::Eof));
    }
    let payload_end = header_offset
        .checked_add(payload_end_relative)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;

    let span_end = if entry.has_data_descriptor {
        let descriptor_data = data
            .get(payload_end..central_directory_offset)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let descriptor = DataDescriptor::parse_complete(descriptor_data, entry, archive_is_zip64)?;
        let descriptor_end = payload_end
            .checked_add(descriptor.encoded_size)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        if descriptor_end > central_directory_offset {
            return Err(Error::from(ErrorKind::Eof));
        }
        descriptor_end
    } else {
        payload_end
    };

    Ok(BorrowedEntryMetadata {
        header_offset,
        data_start_offset: header_offset
            .checked_add(variable_end)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?,
        data_end_offset: payload_end,
        span_end,
    })
}

/// Locates a stored member and validates all metadata that can affect safe
/// immutable-slice publication.  This is intentionally separate from
/// `slice_entry`: the latter is the compatibility path used by decompression
/// and keeps its historical central-directory-authoritative behavior.
fn slice_stored_entry<'data>(
    data: &'data [u8],
    entry: ZipArchiveEntryWayfinder,
    central_name: &[u8],
    central_directory_offset: u64,
    archive_is_zip64: bool,
) -> Result<ZipSliceEntry<'data>, Error> {
    validate_borrowed_wayfinder(&entry)?;
    validate_borrowed_entry_metadata(
        data,
        &entry,
        central_name,
        central_directory_offset,
        archive_is_zip64,
    )?;
    let header_offset =
        usize::try_from(entry.local_header_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
    let central_directory_offset =
        usize::try_from(central_directory_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
    if header_offset > central_directory_offset || central_directory_offset > data.len() {
        return Err(Error::from(ErrorKind::Eof));
    }

    let header = data
        .get(header_offset..central_directory_offset)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let file_header = ZipLocalFileHeaderFixed::parse(header)?;
    if file_header.compression_method != entry.compression_method {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central compression methods differ".to_string(),
        }));
    }
    if file_header.flags != entry.flags {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central flags differ".to_string(),
        }));
    }

    let variable_length = file_header.variable_length();
    let variable_end = ZipLocalFileHeaderFixed::SIZE
        .checked_add(variable_length)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let variable_data = header
        .get(ZipLocalFileHeaderFixed::SIZE..variable_end)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let file_name_len = usize::from(file_header.file_name_len);
    let local_name = variable_data
        .get(..file_name_len)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if local_name != central_name {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local and central names differ".to_string(),
        }));
    }
    let local_extra = variable_data
        .get(file_name_len..)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;

    let local_has_zip64_sentinel =
        file_header.compressed_size == u32::MAX || file_header.uncompressed_size == u32::MAX;
    if local_has_zip64_sentinel && !entry.zip64_sizes {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local header uses ZIP64 sizes absent from the central entry".to_string(),
        }));
    }
    let (local_compressed_size, local_uncompressed_size) = if local_has_zip64_sentinel {
        local_header_sizes(&file_header, local_extra)?
    } else {
        (
            u64::from(file_header.compressed_size),
            u64::from(file_header.uncompressed_size),
        )
    };

    if !entry.has_data_descriptor {
        if file_header.crc32 != entry.crc {
            return Err(Error::from(ErrorKind::InvalidChecksum {
                expected: entry.crc,
                actual: file_header.crc32,
            }));
        }
        if local_compressed_size != entry.compressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.compressed_size,
                actual: local_compressed_size,
            }));
        }
        if local_uncompressed_size != entry.uncompressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.uncompressed_size,
                actual: local_uncompressed_size,
            }));
        }
    }

    let compressed_size =
        usize::try_from(entry.compressed_size).map_err(|_| Error::from(ErrorKind::Eof))?;
    let payload_end_relative = variable_end
        .checked_add(compressed_size)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if payload_end_relative > header.len() {
        return Err(Error::from(ErrorKind::Eof));
    }
    let payload_end = header_offset
        .checked_add(payload_end_relative)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;

    let expected_crc = if entry.has_data_descriptor {
        let descriptor_data = data
            .get(payload_end..central_directory_offset)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let descriptor = DataDescriptor::parse_complete(descriptor_data, &entry, archive_is_zip64)?;
        let descriptor_end = payload_end
            .checked_add(descriptor.encoded_size)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        if descriptor_end > central_directory_offset {
            return Err(Error::from(ErrorKind::Eof));
        }

        descriptor.crc
    } else {
        entry.crc
    };

    let data_start_offset = u32::try_from(variable_end).map_err(|_| Error::from(ErrorKind::Eof))?;
    let entire_entry = data
        .get(header_offset..payload_end)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    Ok(ZipSliceEntry {
        data: entire_entry,
        verifier: ZipVerification {
            crc: expected_crc,
            uncompressed_size: entry.uncompressed_size,
        },
        local_header_offset: entry.local_header_offset,
        data_start_offset,
    })
}

fn borrowed_entry_span(
    data: &[u8],
    entry: ZipArchiveEntryWayfinder,
    central_name: &[u8],
    central_directory_offset: u64,
    archive_is_zip64: bool,
) -> Result<(usize, usize), Error> {
    let metadata = validate_borrowed_entry_metadata(
        data,
        &entry,
        central_name,
        central_directory_offset,
        archive_is_zip64,
    )?;
    Ok((metadata.header_offset, metadata.span_end))
}

fn validate_reader_entry_layout<R: ReaderAt>(
    reader: &R,
    entry: &ZipArchiveEntryWayfinder,
    central_name: &[u8],
    central_directory_offset: u64,
    archive_is_zip64: bool,
) -> Result<StrictEntryLayout, Error> {
    let (layout, local_name_mismatch) = validate_reader_entry_layout_with_name_policy(
        reader,
        entry,
        central_name,
        central_directory_offset,
        archive_is_zip64,
        false,
    )?;
    debug_assert!(!local_name_mismatch);
    Ok(layout)
}

fn validate_reader_entry_layout_with_name_policy<R: ReaderAt>(
    reader: &R,
    entry: &ZipArchiveEntryWayfinder,
    central_name: &[u8],
    central_directory_offset: u64,
    archive_is_zip64: bool,
    allow_name_mismatch: bool,
) -> Result<(StrictEntryLayout, bool), Error> {
    validate_borrowed_wayfinder_metadata(entry)?;
    if entry.local_header_offset > central_directory_offset {
        return Err(Error::from(ErrorKind::Eof));
    }

    let mut fixed_buffer = [0u8; ZipLocalFileHeaderFixed::SIZE];
    reader.read_exact_at(&mut fixed_buffer, entry.local_header_offset)?;
    let file_header = ZipLocalFileHeaderFixed::parse(&fixed_buffer)?;
    let variable_length = file_header.variable_length();
    let variable_end = ZipLocalFileHeaderFixed::SIZE
        .checked_add(variable_length)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let variable_end_u64 = u64::try_from(variable_end).map_err(|_| Error::from(ErrorKind::Eof))?;
    let local_end = entry
        .local_header_offset
        .checked_add(variable_end_u64)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if local_end > central_directory_offset {
        return Err(Error::from(ErrorKind::Eof));
    }

    let mut variable_data = Vec::new();
    variable_data
        .try_reserve_exact(variable_length)
        .map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "strict ZIP local metadata",
                source,
            })
        })?;
    variable_data.resize(variable_length, 0);
    if variable_length != 0 {
        let variable_offset = entry
            .local_header_offset
            .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        reader.read_exact_at(&mut variable_data, variable_offset)?;
    }

    if file_header.compression_method != entry.compression_method {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "strict local and central compression methods differ".to_string(),
        }));
    }
    if file_header.flags != entry.flags {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "strict local and central flags differ".to_string(),
        }));
    }
    let file_name_len = usize::from(file_header.file_name_len);
    let local_name = variable_data
        .get(..file_name_len)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let local_name_mismatch = local_name != central_name;
    if local_name_mismatch && !allow_name_mismatch {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "strict local and central names differ".to_string(),
        }));
    }
    let local_extra = variable_data
        .get(file_name_len..)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let local_has_zip64_sentinel =
        file_header.compressed_size == u32::MAX || file_header.uncompressed_size == u32::MAX;
    if local_has_zip64_sentinel && !entry.zip64_sizes {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "strict local header uses ZIP64 sizes absent from the central entry".to_string(),
        }));
    }
    let (local_compressed_size, local_uncompressed_size) = if local_has_zip64_sentinel {
        local_header_sizes(&file_header, local_extra)?
    } else {
        (
            u64::from(file_header.compressed_size),
            u64::from(file_header.uncompressed_size),
        )
    };
    if !entry.has_data_descriptor || local_has_zip64_sentinel {
        if local_compressed_size != entry.compressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.compressed_size,
                actual: local_compressed_size,
            }));
        }
        if local_uncompressed_size != entry.uncompressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.uncompressed_size,
                actual: local_uncompressed_size,
            }));
        }
    }
    if !entry.has_data_descriptor && file_header.crc32 != entry.crc {
        return Err(Error::from(ErrorKind::InvalidChecksum {
            expected: entry.crc,
            actual: file_header.crc32,
        }));
    }

    let data_start_offset = local_end;
    let data_end_offset = data_start_offset
        .checked_add(entry.compressed_size)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if data_end_offset > central_directory_offset {
        return Err(Error::from(ErrorKind::Eof));
    }
    let span_end = if entry.has_data_descriptor {
        let descriptor = DataDescriptor::parse_complete_at(
            reader,
            data_end_offset,
            central_directory_offset,
            entry,
            archive_is_zip64,
        )?;
        data_end_offset
            .checked_add(
                u64::try_from(descriptor.encoded_size).map_err(|_| Error::from(ErrorKind::Eof))?,
            )
            .ok_or_else(|| Error::from(ErrorKind::Eof))?
    } else {
        data_end_offset
    };
    if span_end > central_directory_offset {
        return Err(Error::from(ErrorKind::Eof));
    }

    Ok((
        StrictEntryLayout {
            local_header_offset: entry.local_header_offset,
            data_start_offset,
            data_end_offset,
            span_end,
            verifier: ZipVerification {
                crc: entry.crc,
                uncompressed_size: entry.uncompressed_size,
            },
        },
        local_name_mismatch,
    ))
}

fn validate_borrowed_wayfinder(entry: &ZipArchiveEntryWayfinder) -> Result<(), Error> {
    validate_borrowed_wayfinder_metadata(entry)?;
    if entry.flags & ((1 << 0) | (1 << 6)) != 0 {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses encrypted Store entries".to_string(),
        }));
    }
    Ok(())
}

fn validate_borrowed_wayfinder_metadata(entry: &ZipArchiveEntryWayfinder) -> Result<(), Error> {
    if !entry.zip64_local_header_offset_resolved || !entry.zip64_disk_start_resolved {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses unresolved ZIP64 offset or disk metadata".to_string(),
        }));
    }
    if entry.disk_number_start != 0 {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses nonzero disk-start metadata".to_string(),
        }));
    }
    if !entry.zip64_compressed_size_resolved || !entry.zip64_uncompressed_size_resolved {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses unresolved ZIP64 size metadata".to_string(),
        }));
    }
    Ok(())
}

fn validate_single_disk_archive(data: &[u8], eocd: &EndOfCentralDirectory) -> Result<(), Error> {
    let tail_offset =
        usize::try_from(eocd.tail_eocd_offset()).map_err(|_| Error::from(ErrorKind::Eof))?;
    let tail_end = tail_offset
        .checked_add(EndOfCentralDirectoryRecordFixed::SIZE)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let tail = data
        .get(tail_offset..tail_end)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let fixed = EndOfCentralDirectoryRecordFixed::parse(tail)?;

    if !eocd.is_zip64() {
        if fixed.disk_number != 0 || fixed.eocd_disk != 0 {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "borrowed access refuses multi-disk archive metadata".to_string(),
            }));
        }
        return Ok(());
    }

    // In a ZIP64 archive the classic EOCD uses 0xffff sentinels.  Any other
    // nonzero value is still a multi-disk declaration and is refused.
    if (fixed.disk_number != 0 && fixed.disk_number != u16::MAX)
        || (fixed.eocd_disk != 0 && fixed.eocd_disk != u16::MAX)
    {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses multi-disk archive metadata".to_string(),
        }));
    }

    let zip64_offset =
        usize::try_from(eocd.head_eocd_offset()).map_err(|_| Error::from(ErrorKind::Eof))?;
    let zip64_fixed_end = zip64_offset
        .checked_add(Zip64EndOfCentralDirectoryRecord::SIZE)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let zip64 = data
        .get(zip64_offset..zip64_fixed_end)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let zip64 = Zip64EndOfCentralDirectoryRecord::parse(zip64)?;
    let locator_offset = eocd
        .tail_eocd_offset()
        .checked_sub(ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIZE as u64)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let zip64_offset_u64 = u64::try_from(zip64_offset).map_err(|_| Error::from(ErrorKind::Eof))?;
    let record_end = zip64_offset_u64
        .checked_add(12)
        .and_then(|offset| offset.checked_add(zip64.size))
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
    if zip64.disk_number != 0 || zip64.cd_disk != 0 {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "borrowed access refuses multi-disk archive metadata".to_string(),
        }));
    }
    Ok(())
}

fn local_header_sizes(
    file_header: &ZipLocalFileHeaderFixed,
    extra: &[u8],
) -> Result<(u64, u64), Error> {
    let needs_uncompressed = file_header.uncompressed_size == u32::MAX;
    let needs_compressed = file_header.compressed_size == u32::MAX;
    if !needs_uncompressed && !needs_compressed {
        return Ok((
            u64::from(file_header.compressed_size),
            u64::from(file_header.uncompressed_size),
        ));
    }

    let mut zip64_data = None;
    let mut fields = ExtraFields::new(extra);
    for (field_id, field_data) in fields.by_ref() {
        if field_id != ExtraFieldId::ZIP64 {
            continue;
        }
        if zip64_data.replace(field_data).is_some() {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "stored local header contains duplicate ZIP64 extra fields".to_string(),
            }));
        }
    }
    if !fields.remaining_bytes().is_empty() {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local header contains malformed extra fields".to_string(),
        }));
    }
    let zip64_data = zip64_data.ok_or_else(|| {
        Error::from(ErrorKind::InvalidInput {
            msg: "stored local header is missing its ZIP64 size extra field".to_string(),
        })
    })?;

    let required_len = usize::from(u8::from(needs_uncompressed))
        .checked_add(usize::from(u8::from(needs_compressed)))
        .and_then(|count| count.checked_mul(8))
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if zip64_data.len() != required_len {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "stored local ZIP64 size extra field has unexpected length".to_string(),
        }));
    }

    let mut pos = 0usize;
    let mut uncompressed_size = u64::from(file_header.uncompressed_size);
    let mut compressed_size = u64::from(file_header.compressed_size);
    if needs_uncompressed {
        let end = pos
            .checked_add(8)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        uncompressed_size = zip64_data.get(pos..end).map(le_u64).ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "stored local ZIP64 uncompressed size is truncated".to_string(),
            })
        })?;
        pos = end;
    }
    if needs_compressed {
        let end = pos
            .checked_add(8)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        compressed_size = zip64_data.get(pos..end).map(le_u64).ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "stored local ZIP64 compressed size is truncated".to_string(),
            })
        })?;
    }
    Ok((compressed_size, uncompressed_size))
}

/// Represents a single entry (file or directory) within a `ZipSliceArchive`.
///
/// It provides access to the raw compressed data of the entry.
#[derive(Debug, Clone)]
pub struct ZipSliceEntry<'a> {
    // From local header offset to end of compressed data
    data: &'a [u8],
    verifier: ZipVerification,
    local_header_offset: u64,
    // self.data[self.data_start_offset] is the start of compressed data
    data_start_offset: u32,
}

impl<'a> ZipSliceEntry<'a> {
    /// Returns the raw, compressed data of the entry as a byte slice.
    pub fn data(&self) -> &'a [u8] {
        self.data
            .get(self.data_start_offset as usize..)
            .unwrap_or_default()
    }

    /// Returns a verifier for the CRC and uncompressed size of the entry.
    ///
    /// Useful when it's more practical to oneshot decompress the data,
    /// otherwise use [`ZipSliceEntry::verifying_reader`] to stream
    /// decompression and verification.
    pub fn claim_verifier(&self) -> ZipVerification {
        self.verifier
    }

    /// Returns a reader that checks the declared size and, when supplied, CRC
    /// of the decompressed data once finished. A declared CRC of zero is
    /// treated as unavailable by the compatibility verifier.
    pub fn verifying_reader<D>(&self, reader: D) -> ZipSliceVerifier<D>
    where
        D: std::io::Read,
    {
        ZipSliceVerifier {
            reader,
            verifier: self.verifier,
            crc: 0,
            size: 0,
        }
    }

    /// Returns the byte range of the compressed data within the archive.
    ///
    /// See [`ZipEntry::compressed_data_range`] for more details.
    pub fn compressed_data_range(&self) -> (u64, u64) {
        let compressed_data_start = self
            .local_header_offset
            .saturating_add(self.data_start_offset as u64);
        let compressed_data_end = compressed_data_start.saturating_add(
            self.data
                .len()
                .saturating_sub(self.data_start_offset as usize) as u64,
        );
        (compressed_data_start, compressed_data_end)
    }

    /// Returns an iterator over the extra fields from the local file header.
    ///
    /// See [`ZipLocalFileHeader`] for more details.
    pub fn extra_fields(&self) -> ExtraFields<'_> {
        let Ok(header) = ZipLocalFileHeaderFixed::parse(self.data) else {
            return ExtraFields::new(&[]);
        };
        let file_name_len = header.file_name_len as usize;
        let extra_field_len = header.extra_field_len as usize;
        let Some(extra_field_start) = ZipLocalFileHeaderFixed::SIZE.checked_add(file_name_len)
        else {
            return ExtraFields::new(&[]);
        };
        let Some(extra_field_end) = extra_field_start.checked_add(extra_field_len) else {
            return ExtraFields::new(&[]);
        };
        ExtraFields::new(
            self.data
                .get(extra_field_start..extra_field_end)
                .unwrap_or_default(),
        )
    }

    /// Returns the file path from the local file header.
    ///
    /// See [`ZipLocalFileHeader`] for more details.
    pub fn file_path(&self) -> ZipFilePath<RawPath<'_>> {
        let Ok(header) = ZipLocalFileHeaderFixed::parse(self.data) else {
            return ZipFilePath::from_bytes(&[]);
        };
        let file_name_len = header.file_name_len as usize;
        let filename_start = ZipLocalFileHeaderFixed::SIZE;
        let Some(filename_end) = filename_start.checked_add(file_name_len) else {
            return ZipFilePath::from_bytes(&[]);
        };
        ZipFilePath::from_bytes(
            self.data
                .get(filename_start..filename_end)
                .unwrap_or_default(),
        )
    }
}

/// Checks the wrapped reader returns the expected size and, when supplied, CRC.
/// A declared CRC of zero is treated as unavailable.
#[derive(Debug, Clone)]
pub struct ZipSliceVerifier<D> {
    reader: D,
    crc: u32,
    size: u64,
    verifier: ZipVerification,
}

impl<D> ZipSliceVerifier<D> {
    /// Consumes the `ZipSliceVerifier`, returning the underlying reader.
    pub fn into_inner(self) -> D {
        self.reader
    }
}

impl<D> std::io::Read for ZipSliceVerifier<D>
where
    D: std::io::Read,
{
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let read = validate_read_count(self.reader.read(buf)?, buf.len())?;
        self.crc = crc32_chunk(&buf[..read], self.crc);
        let read_u64 = u64::try_from(read).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "ZIP verifier read count overflows u64",
            )
        })?;
        self.size = self.size.checked_add(read_u64).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "ZIP verifier size overflows u64",
            )
        })?;

        if read == 0 || self.size >= self.verifier.size() {
            self.verifier
                .valid(ZipVerification {
                    crc: self.crc,
                    uncompressed_size: self.size,
                })
                .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
        }

        Ok(read)
    }
}

/// An iterator over the central directory file header records.
///
/// Created from [`ZipSliceArchive::entries`].
#[derive(Debug, Clone)]
pub struct ZipSliceEntries<'data> {
    entry_data: &'data [u8],
    base_offset: u64,
    current_offset: u64,
}

impl<'data> ZipSliceEntries<'data> {
    /// Yield the next zip file entry in the central directory if there is any
    #[inline]
    pub fn next_entry(&mut self) -> Result<Option<ZipFileHeaderRecord<'data>>, Error> {
        if self.entry_data.is_empty() {
            return Ok(None);
        }

        let file_header = ZipFileHeaderFixed::parse(self.entry_data)?;
        let Some((file_name, extra_field, file_comment, entry_data)) =
            file_header.parse_variable_length(&self.entry_data[ZipFileHeaderFixed::SIZE..])
        else {
            return Err(Error::from(ErrorKind::Eof));
        };

        let mut entry = ZipFileHeaderRecord::from_parts(
            file_header,
            file_name,
            extra_field,
            file_comment,
            self.current_offset,
        )?;
        entry.local_header_offset = entry
            .local_header_offset
            .checked_add(self.base_offset)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        self.current_offset += (self.entry_data.len() - entry_data.len()) as u64;
        self.entry_data = entry_data;
        Ok(Some(entry))
    }
}

impl<'data> Iterator for ZipSliceEntries<'data> {
    type Item = Result<ZipFileHeaderRecord<'data>, Error>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        self.next_entry().transpose()
    }
}

/// The main entrypoint for reading a Zip archive.
///
/// It can be created from a slice, a file, or any `Read + Seek` source.
///
/// # Examples
///
/// Creating from a file:
///
/// ```rust
/// # use soapberry_zip::{ZipArchive, Error, RECOMMENDED_BUFFER_SIZE};
/// # use std::fs::File;
/// # use std::io;
/// fn example_from_file(file: File) -> Result<(), Error> {
///     let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
///     let archive = ZipArchive::from_file(file, &mut buffer)?;
///     Ok(())
/// }
/// ```
///
/// For more complex use cases, use the [`ZipLocator`] to locate an archive.
#[derive(Debug, Clone)]
pub struct ZipArchive<R> {
    reader: R,
    eocd: EndOfCentralDirectory,
}

impl ZipArchive<()> {
    /// Creates a [`ZipLocator`] configured with a maximum search space for the
    /// End of Central Directory Record (EOCD).
    pub fn with_max_search_space(max_search_space: u64) -> ZipLocator {
        ZipLocator::new().max_search_space(max_search_space)
    }

    /// Parses an archive from in-memory data.
    pub fn from_slice<T: AsRef<[u8]>>(data: T) -> Result<ZipSliceArchive<T>, Error> {
        ZipLocator::new().locate_in_slice(data).map_err(|(_, e)| e)
    }

    /// Parses an archive from a file by reading the End of Central Directory.
    ///
    /// A buffer is required to read parts of the file.
    /// [`RECOMMENDED_BUFFER_SIZE`] can be used to construct this buffer.
    pub fn from_file(
        file: std::fs::File,
        buffer: &mut [u8],
    ) -> Result<ZipArchive<FileReader>, Error> {
        ZipLocator::new()
            .locate_in_file(file, buffer)
            .map_err(|(_, e)| e)
    }

    /// Parses an archive from a seekable reader.
    ///
    /// Prefer [`ZipArchive::from_file`] and [`ZipArchive::from_slice`] when
    /// possible, as they are more efficient due to not wrapping the underlying
    /// reader in a mutex to support positioned io.
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, Error, RECOMMENDED_BUFFER_SIZE, ZipFileHeaderRecord};
    /// # use std::io::Cursor;
    /// fn example(zip_data: &[u8]) -> Result<(), Error> {
    ///     let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    ///     let archive = ZipArchive::from_seekable(Cursor::new(zip_data), &mut buffer)?;
    ///     Ok(())
    /// }
    /// ```
    pub fn from_seekable<R>(
        mut reader: R,
        buffer: &mut [u8],
    ) -> Result<ZipArchive<MutexReader<R>>, Error>
    where
        R: Read + Seek,
    {
        let end_offset = reader.seek(std::io::SeekFrom::End(0))?;
        let reader = MutexReader::new(reader);
        ZipLocator::new()
            .locate_in_reader(reader, buffer, end_offset)
            .map_err(|(_, e)| e)
    }
}

impl<R> ZipArchive<R> {
    pub(crate) fn new(reader: R, eocd: EndOfCentralDirectory) -> Self {
        ZipArchive { reader, eocd }
    }

    /// Returns a reference to the underlying reader.
    pub fn get_ref(&self) -> &R {
        &self.reader
    }

    /// Consumes this archive and returns the underlying reader.
    pub fn into_inner(self) -> R {
        self.reader
    }

    /// Returns a lending iterator over the entries in the central directory of
    /// the archive.
    ///
    /// Requires a mutable buffer to read directory entries from the underlying
    /// reader.
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, Error, RECOMMENDED_BUFFER_SIZE, ZipFileHeaderRecord};
    /// # use std::fs::File;
    /// fn example(file: File) -> Result<(), Error> {
    ///     let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    ///     let archive = ZipArchive::from_file(file, &mut buffer)?;
    ///     let entries_hint = archive.entries_hint();
    ///     let mut actual_entries = 0;
    ///     let mut entries_iterator = archive.entries(&mut buffer);
    ///     while let Some(_) = entries_iterator.next_entry()? {
    ///         actual_entries += 1;
    ///     }
    ///     println!("Found {} entries (hint: {})", actual_entries, entries_hint);
    ///     Ok(())
    /// }
    /// ```
    pub fn entries<'archive, 'buf>(
        &'archive self,
        buffer: &'buf mut [u8],
    ) -> ZipEntries<'archive, 'buf, R> {
        self.entries_with_metadata_limit(buffer, u64::MAX)
    }

    /// Returns a lending iterator over the central-directory records with an
    /// aggregate metadata ceiling.
    ///
    /// The caller-provided buffer remains the hot path for ordinary records.
    /// A valid record whose name, extra fields, and comment do not fit in that
    /// buffer is read into a fallible, per-record spill buffer instead. The
    /// spill buffer is bounded by the ZIP `u16` variable-field widths and is
    /// released for reuse on the next oversized record. The supplied limit is
    /// checked before reserving that buffer, so a rejected record cannot cause
    /// an allocation beyond the caller's metadata policy.
    pub fn entries_with_metadata_limit<'archive, 'buf>(
        &'archive self,
        buffer: &'buf mut [u8],
        max_metadata_bytes: u64,
    ) -> ZipEntries<'archive, 'buf, R> {
        ZipEntries {
            buffer,
            metadata_buffer: Vec::new(),
            archive: self,
            pos: 0,
            end: 0,
            offset: self.eocd.directory_offset(),
            base_offset: self.eocd.base_offset(),
            central_dir_end_pos: self.eocd.directory_end_offset(),
            metadata_bytes: 0,
            max_metadata_bytes,
        }
    }

    /// Returns a hint for the total number of entries in the archive.
    ///
    /// This value is read from the End of Central Directory record.
    pub fn entries_hint(&self) -> u64 {
        self.eocd.entries()
    }

    /// Returns the resolved, declared central-directory byte size.
    pub(crate) fn central_directory_size(&self) -> u64 {
        self.eocd.central_directory_size()
    }

    /// Whether this archive uses ZIP64 end-of-central-directory metadata.
    #[inline]
    pub fn is_zip64(&self) -> bool {
        self.eocd.is_zip64()
    }

    /// Returns the first EOCD offset that bounds the central directory.
    ///
    /// For ZIP32 this is the terminal EOCD offset.  For ZIP64 it is the
    /// ZIP64 EOCD offset, which keeps the central directory separate from the
    /// ZIP64 record, locator, and terminal EOCD tail.
    pub fn head_eocd_offset(&self) -> u64 {
        self.eocd.head_eocd_offset()
    }

    /// Returns a Read implementation for the comment of the zip archive.
    ///
    /// Use [`RangeReader::remaining()`] to get the comment length before
    /// reading. It is guaranteed to be less than `u16::MAX`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use soapberry_zip::{ZipArchive, ZipStr, RECOMMENDED_BUFFER_SIZE};
    /// use std::io::Read;
    /// use std::fs::File;
    ///
    /// let file = File::open("assets/test.zip")?;
    /// let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    /// let archive = ZipArchive::from_file(file, &mut buffer)?;
    ///
    /// let mut comment_reader = archive.comment();
    /// let comment_len = comment_reader.remaining() as usize;
    /// comment_reader.read_exact(&mut buffer[..comment_len])?;
    ///
    /// let actual = ZipStr::new(&buffer[..comment_len]);
    /// let expected = ZipStr::new(b"This is a zipfile comment.");
    /// assert_eq!(expected, actual);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn comment(&self) -> RangeReader<&R> {
        let comment_start =
            self.eocd.tail_eocd_offset() + EndOfCentralDirectoryRecordFixed::SIZE as u64;
        let comment_end = comment_start + self.eocd.comment_len() as u64;
        RangeReader::new(&self.reader, comment_start..comment_end)
    }

    /// Returns the offset of the End of Central Directory (EOCD) signature.
    ///
    /// This is the byte position where the EOCD signature (0x06054b50) was found.
    /// Useful for recovery scenarios when dealing with false EOCD signatures or
    /// when restarting archive searches from a known position.
    ///
    /// # Examples
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, ZipLocator, RECOMMENDED_BUFFER_SIZE};
    /// # use std::fs::File;
    /// # fn example() -> Result<(), Box<dyn std::error::Error>> {
    /// # let file = File::open("assets/test.zip")?;
    /// # let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    /// let archive = ZipArchive::from_file(file, &mut buffer)?;
    /// let eocd_position = archive.eocd_offset();
    ///
    /// let locator = ZipLocator::new();
    /// let reader = archive.get_ref();
    /// let maybe_previous = locator.locate_in_reader(reader, &mut buffer, eocd_position);
    /// # Ok(())
    /// # }
    /// ```
    pub fn eocd_offset(&self) -> u64 {
        self.eocd.tail_eocd_offset()
    }

    /// The declared offset of the start of the central directory.
    ///
    /// To verify the validity of this offset, start iterating through the
    /// central directory via `entries()`. Ensure no errors are returned on the
    /// first entry.
    ///
    /// This value is useful when calculating the amount of prelude data exists
    /// in the data, as it will serve as the upper bound until each file's
    /// [`ZipFileHeaderRecord::local_header_offset`] can be examined.
    pub fn directory_offset(&self) -> u64 {
        self.eocd.directory_offset()
    }

    /// Returns the offset where the ZIP archive ends.
    ///
    /// This returns the position immediately after the last byte of the ZIP
    /// archive, including the End of Central Directory record and any comment.
    /// This is useful for extracting trailing data.
    ///
    /// The calculation does not rely on any self reported values from the
    /// archive.
    ///
    /// This can be used in conjunction with the starting offset calculation
    /// start offset as shown in [`RangeReader`] to determine the exact byte
    /// range (and thus size) of the ZIP archive within a context of a larger
    /// file.
    pub fn end_offset(&self) -> u64 {
        self.eocd.tail_eocd_offset()
            + EndOfCentralDirectoryRecordFixed::SIZE as u64
            + self.comment().remaining()
    }
}

impl<R> ZipArchive<R>
where
    R: ReaderAt,
{
    pub(crate) fn validate_strict_stream_target(
        &self,
        entry: ZipArchiveEntryWayfinder,
    ) -> Result<(), Error> {
        validate_borrowed_wayfinder(&entry)
    }

    pub(crate) fn validate_strict_entry_layout(
        &self,
        entry: ZipArchiveEntryWayfinder,
        central_name: &[u8],
    ) -> Result<StrictEntryLayout, Error> {
        validate_reader_entry_layout(
            &self.reader,
            &entry,
            central_name,
            self.eocd.directory_offset(),
            self.eocd.is_zip64(),
        )
    }

    /// Validate one source entry for raw preservation.
    ///
    /// Raw preservation intentionally retains the historical ability to copy
    /// a member whose local and central names differ when the caller does not
    /// regenerate or append members.  All other strict local-header, ZIP64
    /// size, and data-descriptor checks are shared with the trusted reader
    /// path.  The returned boolean reports that name mismatch.
    pub(crate) fn validate_preservation_entry_layout(
        &self,
        entry: ZipArchiveEntryWayfinder,
        central_name: &[u8],
    ) -> Result<(StrictEntryLayout, bool), Error> {
        validate_reader_entry_layout_with_name_policy(
            &self.reader,
            &entry,
            central_name,
            self.eocd.directory_offset(),
            self.eocd.is_zip64(),
            true,
        )
    }

    pub(crate) fn strict_payload_reader(
        &self,
        entry: ZipArchiveEntryWayfinder,
        layout: StrictEntryLayout,
    ) -> ZipReader<&R> {
        ZipReader {
            entry,
            range_reader: RangeReader::new(
                self.get_ref(),
                layout.data_start_offset..layout.data_end_offset,
            ),
        }
    }

    /// Seeks to the given file entry in the zip archive.
    pub fn get_entry(&self, entry: ZipArchiveEntryWayfinder) -> Result<ZipEntry<'_, R>, Error> {
        let mut buffer = [0u8; ZipLocalFileHeaderFixed::SIZE];
        self.reader
            .read_exact_at(&mut buffer, entry.local_header_offset)?;

        // The central directory is the source of truth so we really only parse
        // out the local file header to verify the signature and understand the
        // variable length. Not everyone uses this as the source of truth:
        // https://labs.redyops.com/index.php/2020/04/30/spending-a-night-reading-the-zip-file-format-specification/
        let file_header = ZipLocalFileHeaderFixed::parse(&buffer)?;
        let (body_offset, o1) = entry
            .local_header_offset
            .overflowing_add(ZipLocalFileHeaderFixed::SIZE as u64);
        let (body_offset, o2) = body_offset.overflowing_add(file_header.variable_length() as u64);
        let (body_end_offset, o3) = body_offset.overflowing_add(entry.compressed_size);

        if o1 || o2 || o3 {
            return Err(Error::from(ErrorKind::Eof));
        }

        Ok(ZipEntry {
            archive: self,
            entry,
            body_offset,
            body_end_offset,
        })
    }
}

/// Represents a single entry (file or directory) within a [`ZipArchive`]
#[derive(Debug, Clone)]
pub struct ZipEntry<'archive, R> {
    archive: &'archive ZipArchive<R>,
    body_offset: u64,
    body_end_offset: u64,
    entry: ZipArchiveEntryWayfinder,
}

impl<'archive, R> ZipEntry<'archive, R>
where
    R: ReaderAt,
{
    pub(crate) fn local_header_fixed(&self) -> Result<ZipLocalFileHeaderFixed, Error> {
        let mut buffer = [0u8; ZipLocalFileHeaderFixed::SIZE];
        self.archive
            .get_ref()
            .read_exact_at(&mut buffer, self.entry.local_header_offset)?;
        ZipLocalFileHeaderFixed::parse(&buffer)
    }

    /// Returns a [`ZipReader`] for reading the compressed data of this entry.
    pub fn reader(&self) -> ZipReader<&'archive R> {
        ZipReader {
            entry: self.entry,
            range_reader: RangeReader::new(
                self.archive.get_ref(),
                self.body_offset..self.body_end_offset,
            ),
        }
    }

    /// Returns a reader that checks the declared size and, when supplied, CRC
    /// of the decompressed data once finished. A declared CRC of zero is
    /// treated as unavailable by the compatibility verifier.
    pub fn verifying_reader<D>(&self, reader: D) -> ZipVerifier<D, &'archive R>
    where
        D: std::io::Read,
    {
        ZipVerifier {
            reader,
            crc: 0,
            size: 0,
            archive: self.archive.get_ref(),
            end_offset: self.body_end_offset,
            wayfinder: self.entry,
        }
    }

    /// Returns a tuple of start and end byte offsets for the compressed data
    /// within the underlying reader.
    ///
    /// This method uses the information from the local file header in its
    /// calculations.
    ///
    /// # Security Usage
    ///
    /// This method is useful for detecting overlapping entries, which are often
    /// used in zip bombs. By comparing the ranges returned by this method
    /// across multiple entries, you can identify when entries share compressed
    /// data:
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, Error};
    /// # fn example(data: &[u8]) -> Result<(), Error> {
    /// let archive = ZipArchive::from_slice(data)?;
    /// let mut ranges = Vec::new();
    ///
    /// for entry_result in archive.entries() {
    ///     let entry = entry_result?;
    ///     let wayfinder = entry.wayfinder();
    ///     if let Ok(zip_entry) = archive.get_entry(wayfinder) {
    ///         ranges.push(zip_entry.compressed_data_range());
    ///     }
    /// }
    ///
    /// // Check for overlapping ranges
    /// ranges.sort_by_key(|&(start, _)| start);
    /// for window in ranges.windows(2) {
    ///     let (_, end1) = window[0];
    ///     let (start2, _) = window[1];
    ///     if end1 > start2 {
    ///         panic!("Warning: Overlapping entries detected!");
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn compressed_data_range(&self) -> (u64, u64) {
        (self.body_offset, self.body_end_offset)
    }

    /// Returns the local file header information.
    ///
    /// This method reads the local file header to which may differ from the
    /// central directory data. Most ZIP tools use the central directory as
    /// authoritative, but access to local header data can be useful:
    ///
    /// The local header may contain:
    /// - Additional or different extra fields (richer timestamp data, etc.)
    /// - Different filename than the central directory (security concern)
    ///
    /// The buffer argument must be large enough to hold both the filename and
    /// extra fields from the local header or a too small error will be
    /// returned.
    ///
    /// # Examples
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, RECOMMENDED_BUFFER_SIZE, extra_fields::ExtraFieldId};
    /// # use std::fs::File;
    /// # fn example() -> Result<(), Box<dyn std::error::Error>> {
    /// // Test with filename mismatch test fixture
    /// let file = File::open("assets/filename_mismatch_test.zip")?;
    /// let mut buf = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    /// let archive = ZipArchive::from_file(file, &mut buf)?;
    ///
    /// let mut entries = archive.entries(&mut buf);
    /// let entry_header = entries.next_entry()?.unwrap();
    ///
    /// // Central directory shows one filename
    /// assert_eq!(entry_header.file_path().as_ref(), b"malware.exe");
    /// let wayfinder = entry_header.wayfinder();
    /// let entry = archive.get_entry(wayfinder)?;
    ///
    /// // Read the local header
    /// let mut local_buffer = vec![0u8; 1024];
    /// let local_header = entry.local_header(&mut local_buffer)?;
    ///
    /// // Local header shows different filename
    /// assert_eq!(local_header.file_path().as_ref(), b"safe_file.txt");
    ///
    /// // Access extra fields from local header
    /// let mut found_fields = 0;
    /// for (field_id, _data) in local_header.extra_fields() {
    ///     found_fields += 1;
    ///     // Could check for specific extra field types here
    ///     println!("Found extra field: {:04x}", field_id.as_u16());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn local_header<'a>(&self, buffer: &'a mut [u8]) -> Result<ZipLocalFileHeader<'a>, Error> {
        let mut header_buffer = [0u8; ZipLocalFileHeaderFixed::SIZE];

        // Read the local file header
        self.archive
            .get_ref()
            .read_exact_at(&mut header_buffer, self.entry.local_header_offset)?;

        let local_header_fixed = ZipLocalFileHeaderFixed::parse(&header_buffer)?;
        let file_name_len = local_header_fixed.file_name_len as usize;
        let extra_field_len = local_header_fixed.extra_field_len as usize;
        let total_variable_len = file_name_len + extra_field_len;

        // Check if buffer is large enough for both filename and extra fields
        if buffer.len() < total_variable_len {
            return Err(Error::from(ErrorKind::BufferTooSmall));
        }

        let variable_data = &mut buffer[..total_variable_len];
        let variable_data_offset =
            self.entry.local_header_offset + ZipLocalFileHeaderFixed::SIZE as u64;
        self.archive
            .get_ref()
            .read_exact_at(variable_data, variable_data_offset)?;

        let (filename_data, extra_field_data) = variable_data.split_at(file_name_len);
        Ok(ZipLocalFileHeader {
            file_path: ZipFilePath::from_bytes(filename_data),
            extra_fields: ExtraFields::new(extra_field_data),
        })
    }
}

/// Holds the expected CRC32 checksum and uncompressed size for a Zip entry.
///
/// This struct is used to verify the integrity of decompressed data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ZipVerification {
    pub crc: u32,
    pub uncompressed_size: u64,
}

impl ZipVerification {
    /// Returns the expected CRC32 checksum.
    pub fn crc(&self) -> u32 {
        self.crc
    }

    /// Returns the expected uncompressed size.
    pub fn size(&self) -> u64 {
        self.uncompressed_size
    }

    /// Validates the size and CRC of the entry.
    ///
    /// This function will return an error if the size or CRC does not match
    /// the expected values.
    pub fn valid(&self, rhs: ZipVerification) -> Result<(), Error> {
        self.valid_with_crc_policy(rhs, false)
    }

    pub(crate) fn valid_strict(&self, rhs: ZipVerification) -> Result<(), Error> {
        self.valid_with_crc_policy(rhs, true)
    }

    fn valid_with_crc_policy(&self, rhs: ZipVerification, require_crc: bool) -> Result<(), Error> {
        if self.size() != rhs.size() {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: self.size(),
                actual: rhs.size(),
            }));
        }

        // If the CRC is 0, then it is not verified.
        if (require_crc || self.crc() != 0) && self.crc() != rhs.crc() {
            return Err(Error::from(ErrorKind::InvalidChecksum {
                expected: self.crc(),
                actual: rhs.crc(),
            }));
        }

        Ok(())
    }
}

/// Verifies the checksum of the decompressed data matches the checksum listed in the zip
#[derive(Debug, Clone)]
pub struct ZipVerifier<Decompressor, ReaderAt> {
    reader: Decompressor,
    crc: u32,
    size: u64,
    archive: ReaderAt,
    end_offset: u64,
    wayfinder: ZipArchiveEntryWayfinder,
}

impl<Decompressor, ReaderAt> ZipVerifier<Decompressor, ReaderAt> {
    /// Consumes the [`ZipVerifier`], returning the underlying decompressor.
    pub fn into_inner(self) -> Decompressor {
        self.reader
    }
}

impl<Decompressor, Reader> std::io::Read for ZipVerifier<Decompressor, Reader>
where
    Decompressor: std::io::Read,
    Reader: ReaderAt,
{
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let read = validate_read_count(self.reader.read(buf)?, buf.len())?;
        self.crc = crc32_chunk(&buf[..read], self.crc);
        let read_u64 = u64::try_from(read).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "ZIP verifier read count overflows u64",
            )
        })?;
        self.size = self.size.checked_add(read_u64).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "ZIP verifier size overflows u64",
            )
        })?;

        if read == 0 || self.size >= self.wayfinder.uncompressed_size_hint() {
            let expected_crc = if self.wayfinder.has_data_descriptor {
                DataDescriptor::read_at(&self.archive, self.end_offset).map(|x| x.crc)
            } else {
                Ok(self.wayfinder.crc)
            };

            expected_crc
                .and_then(|expected_crc| {
                    let expected = ZipVerification {
                        crc: expected_crc,
                        uncompressed_size: self.wayfinder.uncompressed_size_hint(),
                    };

                    expected.valid(ZipVerification {
                        crc: self.crc,
                        uncompressed_size: self.size,
                    })
                })
                .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
        }

        Ok(read)
    }
}

/// A reader for a Zip entry's compressed data.
#[derive(Debug, Clone)]
pub struct ZipReader<R> {
    entry: ZipArchiveEntryWayfinder,
    range_reader: RangeReader<R>,
}

impl<R> ZipReader<R>
where
    R: ReaderAt,
{
    /// Returns an object that can be used to verify the size and checksum of
    /// inflated data
    ///
    /// Consumes the reader, so this should be called after all data has been read from the entry.
    ///
    /// The function will read the data descriptor if one is expected to exist.
    pub fn claim_verifier(self) -> Result<ZipVerification, Error> {
        let expected_size = self.entry.uncompressed_size_hint();

        let expected_crc = if self.entry.has_data_descriptor {
            let end_offset = self.range_reader.end_offset();
            let archive = self.range_reader.into_inner();
            DataDescriptor::read_at(archive, end_offset).map(|x| x.crc)?
        } else {
            self.entry.crc
        };

        Ok(ZipVerification {
            crc: expected_crc,
            uncompressed_size: expected_size,
        })
    }
}

impl<R> Read for ZipReader<R>
where
    R: ReaderAt,
{
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        self.range_reader.read(buf)
    }
}

/// Local file header information from a ZIP archive entry.
///
/// This struct provides access to data stored in the local file header of a ZIP entry,
/// which may differ from the information in the central directory. The local header
/// contains the filename and extra fields as they appear at the start of each entry's
/// data within the ZIP file.
///
/// Most ZIP tools use the central directory as authoritative, but access to local
/// header data is useful for validation, security analysis, and forensic purposes.
#[derive(Debug)]
pub struct ZipLocalFileHeader<'a> {
    file_path: ZipFilePath<RawPath<'a>>,
    extra_fields: ExtraFields<'a>,
}

impl<'a> ZipLocalFileHeader<'a> {
    /// Returns the file path from the local file header.
    ///
    /// This may differ from the central directory file path.
    pub fn file_path(&self) -> ZipFilePath<RawPath<'a>> {
        self.file_path
    }

    /// Returns an iterator over the extra fields from the local file header.
    ///
    /// Extra fields in the local header may differ from those in the central directory.
    /// The local header may contain additional or different metadata compared to the
    /// central directory entry.
    pub fn extra_fields(&self) -> ExtraFields<'a> {
        self.extra_fields
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct DataDescriptor {
    crc: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    encoded_size: usize,
}

impl DataDescriptor {
    const SIZE: usize = 8;
    pub const SIGNATURE: u32 = 0x08074b50;

    fn parse(data: &[u8]) -> Result<DataDescriptor, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let mut pos = 0;

        let potential_signature = le_u32(&data[0..4]);
        if potential_signature == Self::SIGNATURE {
            pos += 4;
        }

        // The crc is followed by the compressed_size and then the
        // uncompressed_size but the spec allows for the sizes to be either 4
        // bytes each or 8 bytes in Zip64 mode. (spec 4.3.9.1). They aren't
        // needed, so we skip them.
        Ok(DataDescriptor {
            crc: le_u32(&data[pos..pos + 4]),
            compressed_size: 0,
            uncompressed_size: 0,
            encoded_size: pos + 4,
        })
    }

    fn parse_complete(
        data: &[u8],
        entry: &ZipArchiveEntryWayfinder,
        archive_is_zip64: bool,
    ) -> Result<DataDescriptor, Error> {
        let (widths, width_count) = if entry.zip64_sizes {
            ([8usize, 0], 1)
        } else if archive_is_zip64 {
            ([4usize, 8], 2)
        } else {
            ([4usize, 0], 1)
        };
        let has_signature = data.get(..4).map(le_u32) == Some(Self::SIGNATURE);

        // A signature can be either the marker on a signed descriptor or the
        // CRC of an unsigned descriptor.  Keep the signed interpretation
        // first, then the unsigned interpretation, and try narrower framing
        // before wider framing within each interpretation.  For a ZIP64
        // archive with ordinary per-entry sizes both descriptor widths are
        // seen in the wild, so only a unique central-record match is safe.
        let mut starts = [0usize; 2];
        let start_count = if has_signature {
            starts[0] = 4;
            if entry.crc == Self::SIGNATURE {
                starts[1] = 0;
                2
            } else {
                1
            }
        } else {
            1
        };
        if !has_signature {
            starts[0] = 0;
        }

        let mut parsed = [None; 4];
        let mut parsed_count = 0usize;
        for start in starts[..start_count].iter().copied() {
            for width in widths[..width_count].iter().copied() {
                if let Some(descriptor) = Self::parse_fields(data, start, width) {
                    parsed[parsed_count] = Some(descriptor);
                    parsed_count += 1;
                }
            }
        }

        let mut matching = None;
        for (index, descriptor) in parsed[..parsed_count].iter().enumerate() {
            if descriptor
                .as_ref()
                .is_some_and(|descriptor| descriptor.matches(entry))
                && matching.replace(index).is_some()
            {
                return Err(Error::from(ErrorKind::InvalidInput {
                    msg: "ambiguous stored data descriptor framing".to_string(),
                }));
            }
        }
        if let Some(index) = matching {
            return Ok(parsed[index]
                .take()
                .expect("matching descriptor was parsed"));
        }

        // Preserve the historical error ordering: report the first complete
        // candidate's checksum/size error in signed-before-unsigned and
        // narrow-before-wide order.  If no candidate is complete, report EOF.
        for descriptor in parsed[..parsed_count].iter().flatten() {
            descriptor.validate(entry)?;
        }
        Err(Error::from(ErrorKind::Eof))
    }

    fn parse_complete_at<R: ReaderAt>(
        reader: &R,
        offset: u64,
        central_directory_offset: u64,
        entry: &ZipArchiveEntryWayfinder,
        archive_is_zip64: bool,
    ) -> Result<DataDescriptor, Error> {
        if offset > central_directory_offset {
            return Err(Error::from(ErrorKind::Eof));
        }
        let max_width: usize = if entry.zip64_sizes || archive_is_zip64 {
            8
        } else {
            4
        };
        let maximum = 4usize
            .checked_add(4)
            .and_then(|size| size.checked_add(max_width.checked_mul(2)?))
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let available = central_directory_offset - offset;
        let read_length = usize::try_from(
            available.min(u64::try_from(maximum).map_err(|_| Error::from(ErrorKind::Eof))?),
        )
        .map_err(|_| Error::from(ErrorKind::Eof))?;
        let mut buffer = [0u8; 24];
        if read_length != 0 {
            reader.read_exact_at(&mut buffer[..read_length], offset)?;
        }
        Self::parse_complete(&buffer[..read_length], entry, archive_is_zip64)
    }

    fn parse_fields(data: &[u8], offset: usize, width: usize) -> Option<DataDescriptor> {
        let crc_end = offset.checked_add(4)?;
        let compressed_end = crc_end.checked_add(width)?;
        let uncompressed_end = compressed_end.checked_add(width)?;
        let crc = le_u32(data.get(offset..crc_end)?);
        let compressed_size = match width {
            4 => u64::from(le_u32(data.get(crc_end..compressed_end)?)),
            8 => le_u64(data.get(crc_end..compressed_end)?),
            _ => return None,
        };
        let uncompressed_size = match width {
            4 => u64::from(le_u32(data.get(compressed_end..uncompressed_end)?)),
            8 => le_u64(data.get(compressed_end..uncompressed_end)?),
            _ => return None,
        };
        Some(DataDescriptor {
            crc,
            compressed_size,
            uncompressed_size,
            encoded_size: uncompressed_end,
        })
    }

    fn matches(&self, entry: &ZipArchiveEntryWayfinder) -> bool {
        self.crc == entry.crc
            && self.compressed_size == entry.compressed_size
            && self.uncompressed_size == entry.uncompressed_size
    }

    fn validate(&self, entry: &ZipArchiveEntryWayfinder) -> Result<(), Error> {
        if self.crc != entry.crc {
            return Err(Error::from(ErrorKind::InvalidChecksum {
                expected: entry.crc,
                actual: self.crc,
            }));
        }
        if self.compressed_size != entry.compressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.compressed_size,
                actual: self.compressed_size,
            }));
        }
        if self.uncompressed_size != entry.uncompressed_size {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: entry.uncompressed_size,
                actual: self.uncompressed_size,
            }));
        }
        Ok(())
    }

    fn read_at<R>(reader: R, offset: u64) -> Result<DataDescriptor, Error>
    where
        R: ReaderAt,
    {
        let mut buffer = [0u8; Self::SIZE];
        reader.read_exact_at(&mut buffer, offset)?;
        Self::parse(&buffer)
    }
}

/// A lending iterator over file header records in a [`ZipArchive`].
#[derive(Debug)]
pub struct ZipEntries<'archive, 'buf, R> {
    buffer: &'buf mut [u8],
    metadata_buffer: Vec<u8>,
    archive: &'archive ZipArchive<R>,
    pos: usize,
    end: usize,
    offset: u64,
    base_offset: u64,
    central_dir_end_pos: u64,
    metadata_bytes: u64,
    max_metadata_bytes: u64,
}

impl<R> ZipEntries<'_, '_, R>
where
    R: ReaderAt,
{
    /// Yield the next zip file entry in the central directory if there is any
    ///
    /// This method reads from the underlying archive reader into the provided
    /// buffer to parse entry headers.
    #[inline]
    pub fn next_entry(&mut self) -> Result<Option<ZipFileHeaderRecord<'_>>, Error> {
        if self.pos + ZipFileHeaderFixed::SIZE >= self.end {
            if self.offset >= self.central_dir_end_pos {
                return Ok(None);
            }

            let remaining = self.end - self.pos;
            self.buffer.copy_within(self.pos..self.end, 0);
            let remaining_central = self
                .central_dir_end_pos
                .checked_sub(self.offset)
                .ok_or_else(|| Error::from(ErrorKind::Eof))?;
            let max_read = usize::try_from(remaining_central)
                .unwrap_or(usize::MAX)
                .min(self.buffer.len().saturating_sub(remaining));
            let read = self.archive.reader.read_at_least_at(
                &mut self.buffer[remaining..][..max_read],
                ZipFileHeaderFixed::SIZE,
                self.offset,
            )?;
            self.offset += read as u64;
            self.pos = 0;
            self.end = remaining + read;
        }

        let central_directory_offset = self.offset - (self.end - self.pos) as u64;
        let data = &self.buffer[self.pos..self.end];
        let file_header = ZipFileHeaderFixed::parse(data)?;
        self.pos += ZipFileHeaderFixed::SIZE;

        let variable_length = file_header.variable_length();
        let metadata_length = u64::try_from(variable_length).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: "ZIP central-directory metadata length does not fit u64".to_string(),
            })
        })?;
        let next_metadata_bytes = self
            .metadata_bytes
            .checked_add(metadata_length)
            .ok_or_else(|| {
                Error::from(ErrorKind::InvalidInput {
                    msg: "ZIP central-directory metadata total overflows u64".to_string(),
                })
            })?;
        if next_metadata_bytes > self.max_metadata_bytes {
            return Err(ErrorKind::LimitExceeded {
                resource: crate::LimitResource::MetadataBytes,
                actual: next_metadata_bytes,
                maximum: self.max_metadata_bytes,
            }
            .into());
        }
        self.metadata_bytes = next_metadata_bytes;

        // A ZIP32 central record can contain three independent u16-sized
        // variable fields. The ordinary scratch buffer is intentionally only
        // a recommended size, not a validity requirement. Read an oversized
        // record into an owned spill buffer so valid metadata is not rejected
        // merely because it crosses that recommendation.
        if variable_length > self.buffer.len() {
            let variable_data_offset = central_directory_offset
                .checked_add(ZipFileHeaderFixed::SIZE as u64)
                .ok_or_else(|| Error::from(ErrorKind::Eof))?;
            let variable_data_end = variable_data_offset
                .checked_add(metadata_length)
                .ok_or_else(|| Error::from(ErrorKind::Eof))?;
            if variable_data_end > self.central_dir_end_pos {
                return Err(Error::from(ErrorKind::Eof));
            }

            if self.metadata_buffer.len() < variable_length {
                self.metadata_buffer
                    .try_reserve_exact(variable_length - self.metadata_buffer.len())
                    .map_err(|source| ErrorKind::Allocation {
                        resource: "ZIP central-directory metadata buffer",
                        source,
                    })?;
                self.metadata_buffer.resize(variable_length, 0);
            }
            self.archive.reader.read_exact_at(
                &mut self.metadata_buffer[..variable_length],
                variable_data_offset,
            )?;

            // Discard any prefetched bytes. The spill read starts at the
            // logical record boundary and makes the next read begin exactly
            // after the record, avoiding stale-buffer arithmetic.
            self.offset = variable_data_end;
            self.pos = self.end;
            let data = &self.metadata_buffer[..variable_length];
            let Some((file_name, extra_field, file_comment, _)) =
                file_header.parse_variable_length(data)
            else {
                return Err(Error::from(ErrorKind::Eof));
            };
            let mut file_header = ZipFileHeaderRecord::from_parts(
                file_header,
                file_name,
                extra_field,
                file_comment,
                central_directory_offset,
            )?;
            file_header.local_header_offset = file_header
                .local_header_offset
                .checked_add(self.base_offset)
                .ok_or_else(|| Error::from(ErrorKind::Eof))?;
            return Ok(Some(file_header));
        }

        if self.pos + variable_length > self.end {
            // Need to read more data
            let remaining = self.end - self.pos;
            self.buffer.copy_within(self.pos..self.end, 0);
            let remaining_central = self
                .central_dir_end_pos
                .checked_sub(self.offset)
                .ok_or_else(|| Error::from(ErrorKind::Eof))?;
            let max_read = usize::try_from(remaining_central)
                .unwrap_or(usize::MAX)
                .min(self.buffer.len().saturating_sub(remaining));
            let read = self.archive.reader.read_at_least_at(
                &mut self.buffer[remaining..][..max_read],
                variable_length - remaining,
                self.offset,
            )?;
            self.offset += read as u64;
            self.pos = 0;
            self.end = remaining + read;
        }

        let data = &self.buffer[self.pos..self.end];
        let Some((file_name, extra_field, file_comment, _)) =
            file_header.parse_variable_length(data)
        else {
            return Err(Error::from(ErrorKind::Eof));
        };
        let mut file_header = ZipFileHeaderRecord::from_parts(
            file_header,
            file_name,
            extra_field,
            file_comment,
            central_directory_offset,
        )?;
        file_header.local_header_offset = file_header
            .local_header_offset
            .checked_add(self.base_offset)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        self.pos += variable_length;
        Ok(Some(file_header))
    }
}

/// 4.4.2
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct VersionMadeBy(u16);

#[allow(dead_code)]
impl VersionMadeBy {
    pub fn as_u16(&self) -> u16 {
        self.0
    }

    /// The (major, minor) ZIP specification version supported by the software
    /// used to encode the file.
    ///
    /// 4.4.2.3: The lower byte, The value / 10 indicates the major version
    /// number, and the value mod 10 is the minor version number.
    pub fn version(&self) -> (u8, u8) {
        let v = (self.0 >> 8) as u8;
        (v / 10, v % 10)
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Zip64EndOfCentralDirectory {
    pub offset: u64,
    pub central_dir_offset: u64,
    pub central_dir_size: u64,
    pub num_entries: u64,
}

impl Zip64EndOfCentralDirectory {
    #[inline]
    pub fn from_parts(offset: u64, record: Zip64EndOfCentralDirectoryRecord) -> Self {
        Self {
            offset,
            central_dir_offset: record.central_dir_offset,
            central_dir_size: record.central_dir_size,
            num_entries: record.num_entries,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Zip64EndOfCentralDirectoryRecord {
    /// zip64 end of central dir signature
    pub signature: u32,

    /// size of zip64 end of central directory record
    #[allow(dead_code)]
    pub size: u64,

    /// version made by
    #[allow(dead_code)]
    pub version_made_by: VersionMadeBy,

    /// version needed to extract
    #[allow(dead_code)]
    pub version_needed: u16,

    /// number of this disk
    #[allow(dead_code)]
    pub disk_number: u32,

    /// number of the disk with the start of the central directory
    #[allow(dead_code)]
    pub cd_disk: u32,

    /// total number of entries in the central directory on this disk
    pub num_entries: u64,

    /// total number of entries in the central directory
    #[allow(dead_code)]
    pub total_entries: u64,

    /// size of the central directory
    pub central_dir_size: u64,

    /// offset of start of central directory with respect to the starting disk number
    pub central_dir_offset: u64,
    // zip64 extensible data sector
    // pub extensible_data: Vec<u8>,
}

impl Zip64EndOfCentralDirectoryRecord {
    pub(crate) const SIZE: usize = 56;

    #[inline]
    pub fn parse(data: &[u8]) -> Result<Zip64EndOfCentralDirectoryRecord, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let result = Zip64EndOfCentralDirectoryRecord {
            signature: le_u32(&data[0..4]),
            size: le_u64(&data[4..12]),
            version_made_by: VersionMadeBy(le_u16(&data[12..14])),
            version_needed: le_u16(&data[14..16]),
            disk_number: le_u32(&data[16..20]),
            cd_disk: le_u32(&data[20..24]),
            num_entries: le_u64(&data[24..32]),
            total_entries: le_u64(&data[32..40]),
            central_dir_size: le_u64(&data[40..48]),
            central_dir_offset: le_u64(&data[48..56]),
        };

        if result.signature != END_OF_CENTRAL_DIR_SIGNATURE64 {
            return Err(Error::from(ErrorKind::InvalidSignature {
                expected: END_OF_CENTRAL_DIR_SIGNATURE64,
                actual: result.signature,
            }));
        }

        Ok(result)
    }
}

/// A numeric identifier for a compression method used in a Zip archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CompressionMethodId(u16);

impl CompressionMethodId {
    /// Returns the raw `u16` value of the compression method ID.
    #[inline]
    pub fn as_u16(&self) -> u16 {
        self.0
    }

    /// Converts the numeric ID to a `CompressionMethod` enum.
    #[inline]
    pub fn as_method(&self) -> CompressionMethod {
        match self.0 {
            0 => CompressionMethod::Store,
            1 => CompressionMethod::Shrunk,
            2 => CompressionMethod::Reduce1,
            3 => CompressionMethod::Reduce2,
            4 => CompressionMethod::Reduce3,
            5 => CompressionMethod::Reduce4,
            6 => CompressionMethod::Imploded,
            7 => CompressionMethod::Tokenizing,
            8 => CompressionMethod::Deflate,
            9 => CompressionMethod::Deflate64,
            10 => CompressionMethod::Terse,
            12 => CompressionMethod::Bzip2,
            14 => CompressionMethod::Lzma,
            18 => CompressionMethod::Lz77,
            20 => CompressionMethod::ZstdDeprecated,
            93 => CompressionMethod::Zstd,
            94 => CompressionMethod::Mp3,
            95 => CompressionMethod::Xz,
            96 => CompressionMethod::Jpeg,
            97 => CompressionMethod::WavPack,
            98 => CompressionMethod::Ppmd,
            99 => CompressionMethod::Aes,
            _ => CompressionMethod::Unknown(self.0),
        }
    }
}

/// The compression method used on an individual Zip archive entry
///
/// Documented in the spec under: 4.4.5
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum CompressionMethod {
    Store = 0,
    Shrunk = 1,
    Reduce1 = 2,
    Reduce2 = 3,
    Reduce3 = 4,
    Reduce4 = 5,
    Imploded = 6,
    Tokenizing = 7,
    Deflate = 8,
    Deflate64 = 9,
    Terse = 10,
    Bzip2 = 12,
    Lzma = 14,
    Lz77 = 18,
    ZstdDeprecated = 20,
    Zstd = 93,
    Mp3 = 94,
    Xz = 95,
    Jpeg = 96,
    WavPack = 97,
    Ppmd = 98,
    Aes = 99,
    Unknown(u16),
}

impl CompressionMethod {
    /// Return the numeric id of this compression method.
    #[inline]
    pub fn as_id(&self) -> CompressionMethodId {
        let value = match self {
            CompressionMethod::Store => 0,
            CompressionMethod::Shrunk => 1,
            CompressionMethod::Reduce1 => 2,
            CompressionMethod::Reduce2 => 3,
            CompressionMethod::Reduce3 => 4,
            CompressionMethod::Reduce4 => 5,
            CompressionMethod::Imploded => 6,
            CompressionMethod::Tokenizing => 7,
            CompressionMethod::Deflate => 8,
            CompressionMethod::Deflate64 => 9,
            CompressionMethod::Terse => 10,
            CompressionMethod::Bzip2 => 12,
            CompressionMethod::Lzma => 14,
            CompressionMethod::Lz77 => 18,
            CompressionMethod::ZstdDeprecated => 20,
            CompressionMethod::Zstd => 93,
            CompressionMethod::Mp3 => 94,
            CompressionMethod::Xz => 95,
            CompressionMethod::Jpeg => 96,
            CompressionMethod::WavPack => 97,
            CompressionMethod::Ppmd => 98,
            CompressionMethod::Aes => 99,
            CompressionMethod::Unknown(id) => *id,
        };
        CompressionMethodId(value)
    }
}

impl From<u16> for CompressionMethod {
    fn from(id: u16) -> Self {
        CompressionMethodId(id).as_method()
    }
}

/// A borrowed data from a Zip archive, typically for comments or non-path text.
///
/// Zip archives may contain text that is not strictly UTF-8. This type
/// represents such text as a byte slice.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ZipStr<'a>(&'a [u8]);

impl<'a> ZipStr<'a> {
    /// Creates a new `ZipStr` from a byte slice.
    #[inline]
    pub fn new(data: &'a [u8]) -> Self {
        Self(data)
    }

    /// Returns the underlying byte slice.
    #[inline]
    pub fn as_bytes(&self) -> &'a [u8] {
        self.0
    }

    /// Converts the borrowed `ZipStr` into an owned `ZipString` by cloning the
    /// data.
    #[inline]
    pub fn into_owned(&self) -> ZipString {
        ZipString::new(self.0.to_vec())
    }
}

/// An owned string (`Vec<u8>`) from a Zip archive, typically for comments or non-path text.
///
/// Similar to `ZipStr`, but owns its data.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ZipString(Vec<u8>);

impl ZipString {
    /// Creates a new `ZipString` from a vector of bytes.
    #[inline]
    pub fn new(data: Vec<u8>) -> Self {
        Self(data)
    }

    /// Returns a borrowed `ZipStr` view of this `ZipString`.
    #[inline]
    pub fn as_str(&self) -> ZipStr<'_> {
        ZipStr::new(self.0.as_slice())
    }
}

/// Represents a record from the Zip archive's central directory for a single
/// file
///
/// This contains metadata about the file. If interested in navigating to the
/// file contents, use `[ZipFileHeaderRecord::wayfinder]`.
///
/// Reference 4.3.12 in the zip specification
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct ZipFileHeaderRecord<'a> {
    signature: u32,
    version_made_by: u16,
    version_needed: u16,
    flags: u16,
    compression_method: CompressionMethodId,
    last_mod_time: u16,
    last_mod_date: u16,
    crc32: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    file_name_len: u16,
    extra_field_len: u16,
    file_comment_len: u16,
    disk_number_start: u32,
    internal_file_attrs: u16,
    external_file_attrs: u32,
    local_header_offset: u64,
    central_directory_offset: u64,
    file_name: ZipFilePath<RawPath<'a>>,
    extra_field: &'a [u8],
    file_comment: ZipStr<'a>,
    is_zip64: bool,
    zip64_sizes: bool,
    zip64_uncompressed_size_resolved: bool,
    zip64_compressed_size_resolved: bool,
    zip64_local_header_offset_resolved: bool,
    zip64_disk_start_resolved: bool,
    /// Absolute source range of the ZIP64 local-header-offset value in the
    /// central record, when the record used that value instead of its ZIP32
    /// sentinel.  Keeping this span lets a preservation writer patch the
    /// existing value in place without rediscovering the ZIP64 field.
    zip64_local_header_offset_range: Option<Range<u64>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Zip64ExtraValues {
    uncompressed_size: Option<u64>,
    compressed_size: Option<u64>,
    local_header_offset: Option<u64>,
    disk_number_start: Option<u32>,
    zip64_field_start: usize,
    local_header_offset_range: Option<(usize, usize)>,
}

/// Resolve the ZIP64 extra field only when a central-directory fixed field
/// carries its ZIP64 sentinel.  Generic extra fields intentionally retain the
/// historical permissive iterator behavior; this parser is the strict path
/// for metadata that controls archive layout or entry navigation.
fn parse_zip64_extra_values(
    header: &ZipFileHeaderFixed,
    extra_field: &[u8],
) -> Result<Option<Zip64ExtraValues>, Error> {
    let needs_uncompressed = header.uncompressed_size == u32::MAX;
    let needs_compressed = header.compressed_size == u32::MAX;
    let needs_local_header_offset = header.local_header_offset == u32::MAX;
    let needs_disk_number_start = header.disk_number_start == u16::MAX;

    if !needs_uncompressed
        && !needs_compressed
        && !needs_local_header_offset
        && !needs_disk_number_start
    {
        return Ok(None);
    }

    let mut zip64_data = None;
    let mut fields = ExtraFields::new(extra_field);
    while !fields.remaining_bytes().is_empty() {
        let field_start = extra_field
            .len()
            .checked_sub(fields.remaining_bytes().len())
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let Some((field_id, field_data)) = fields.next() else {
            break;
        };
        if field_id != ExtraFieldId::ZIP64 {
            continue;
        }
        if zip64_data.replace((field_start, field_data)).is_some() {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "central header contains duplicate ZIP64 extra fields".to_string(),
            }));
        }
    }
    if !fields.remaining_bytes().is_empty() {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "central header contains malformed extra fields".to_string(),
        }));
    }
    let (zip64_field_start, zip64_data) = zip64_data.ok_or_else(|| {
        Error::from(ErrorKind::InvalidInput {
            msg: "central header is missing its required ZIP64 extra field".to_string(),
        })
    })?;

    let required_len = usize::from(u8::from(needs_uncompressed))
        .checked_add(usize::from(u8::from(needs_compressed)))
        .and_then(|count| count.checked_mul(8))
        .and_then(|size| {
            size.checked_add(usize::from(u8::from(needs_local_header_offset)).checked_mul(8)?)
        })
        .and_then(|size| {
            size.checked_add(usize::from(u8::from(needs_disk_number_start)).checked_mul(4)?)
        })
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if zip64_data.len() < required_len {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "central header ZIP64 extra field is truncated".to_string(),
        }));
    }
    if zip64_data.len() != required_len {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "central header ZIP64 extra field has unexpected trailing data".to_string(),
        }));
    }

    let mut pos = 0usize;

    let uncompressed_size = needs_uncompressed
        .then(|| {
            take_zip64_u64(
                zip64_data,
                &mut pos,
                "central header ZIP64 uncompressed size is truncated",
            )
        })
        .transpose()?;
    let compressed_size = needs_compressed
        .then(|| {
            take_zip64_u64(
                zip64_data,
                &mut pos,
                "central header ZIP64 compressed size is truncated",
            )
        })
        .transpose()?;
    let local_header_offset_start = needs_local_header_offset.then_some(pos);
    let local_header_offset = needs_local_header_offset
        .then(|| {
            take_zip64_u64(
                zip64_data,
                &mut pos,
                "central header ZIP64 local-header offset is truncated",
            )
        })
        .transpose()?;
    let local_header_offset_range = local_header_offset_start.map(|start| (start, pos));
    let disk_number_start = if needs_disk_number_start {
        let end = pos
            .checked_add(4)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        let value = zip64_data.get(pos..end).map(le_u32).ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "central header ZIP64 disk-start value is truncated".to_string(),
            })
        })?;
        pos = end;
        Some(value)
    } else {
        None
    };

    debug_assert_eq!(pos, required_len);
    Ok(Some(Zip64ExtraValues {
        uncompressed_size,
        compressed_size,
        local_header_offset,
        disk_number_start,
        zip64_field_start,
        local_header_offset_range,
    }))
}

fn take_zip64_u64(data: &[u8], pos: &mut usize, message: &'static str) -> Result<u64, Error> {
    let end = pos
        .checked_add(8)
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let value = data.get(*pos..end).map(le_u64).ok_or_else(|| {
        Error::from(ErrorKind::InvalidInput {
            msg: message.to_string(),
        })
    })?;
    *pos = end;
    Ok(value)
}

impl<'a> ZipFileHeaderRecord<'a> {
    #[inline]
    fn from_parts(
        header: ZipFileHeaderFixed,
        file_name: &'a [u8],
        extra_field: &'a [u8],
        file_comment: &'a [u8],
        central_directory_offset: u64,
    ) -> Result<Self, Error> {
        let mut result = Self {
            signature: header.signature,
            version_made_by: header.version_made_by,
            version_needed: header.version_needed,
            flags: header.flags,
            compression_method: header.compression_method,
            last_mod_time: header.last_mod_time,
            last_mod_date: header.last_mod_date,
            crc32: header.crc32,
            compressed_size: u64::from(header.compressed_size),
            uncompressed_size: u64::from(header.uncompressed_size),
            file_name_len: header.file_name_len,
            extra_field_len: header.extra_field_len,
            file_comment_len: header.file_comment_len,
            disk_number_start: u32::from(header.disk_number_start),
            internal_file_attrs: header.internal_file_attrs,
            external_file_attrs: header.external_file_attrs,
            local_header_offset: u64::from(header.local_header_offset),
            central_directory_offset,
            file_name: ZipFilePath::from_bytes(file_name),
            extra_field,
            file_comment: ZipStr::new(file_comment),
            is_zip64: false,
            zip64_sizes: header.uncompressed_size == u32::MAX || header.compressed_size == u32::MAX,
            zip64_uncompressed_size_resolved: header.uncompressed_size != u32::MAX,
            zip64_compressed_size_resolved: header.compressed_size != u32::MAX,
            zip64_local_header_offset_resolved: header.local_header_offset != u32::MAX,
            zip64_disk_start_resolved: header.disk_number_start != u16::MAX,
            zip64_local_header_offset_range: None,
        };

        if let Some(values) = parse_zip64_extra_values(&header, extra_field)? {
            result.is_zip64 = true;
            if let Some(value) = values.uncompressed_size {
                result.uncompressed_size = value;
                result.zip64_uncompressed_size_resolved = true;
            }
            if let Some(value) = values.compressed_size {
                result.compressed_size = value;
                result.zip64_compressed_size_resolved = true;
            }
            if let Some(value) = values.local_header_offset {
                result.local_header_offset = value;
                result.zip64_local_header_offset_resolved = true;
            }
            if let Some(value) = values.disk_number_start {
                result.disk_number_start = value;
                result.zip64_disk_start_resolved = true;
            }
            if let Some((start, end)) = values.local_header_offset_range {
                let width = end
                    .checked_sub(start)
                    .ok_or_else(|| Error::from(ErrorKind::Eof))?;
                let zip64_field_start = u64::try_from(values.zip64_field_start)
                    .map_err(|_| Error::from(ErrorKind::Eof))?;
                let start = u64::try_from(start)
                    .map_err(|_| Error::from(ErrorKind::Eof))?
                    .checked_add(ZipFileHeaderFixed::SIZE as u64)
                    .and_then(|offset| offset.checked_add(u64::from(header.file_name_len)))
                    .and_then(|offset| offset.checked_add(zip64_field_start))
                    .and_then(|offset| offset.checked_add(4))
                    .and_then(|offset| central_directory_offset.checked_add(offset))
                    .ok_or_else(|| Error::from(ErrorKind::Eof))?;
                let width = u64::try_from(width).map_err(|_| Error::from(ErrorKind::Eof))?;
                let end = start
                    .checked_add(width)
                    .ok_or_else(|| Error::from(ErrorKind::Eof))?;
                result.zip64_local_header_offset_range = Some(start..end);
            }
        }

        Ok(result)
    }

    /// Describes if the file is a directory.
    ///
    /// See [`ZipFilePath::is_dir`] for more information.
    #[inline]
    pub fn is_dir(&self) -> bool {
        self.file_name.is_dir()
    }

    /// Returns true if the entry has a data descriptor that follows its
    /// compressed data.
    ///
    /// From the spec (4.3.9.1):
    ///
    /// > This descriptor MUST exist if bit 3 of the general purpose bit flag is
    /// > set
    #[inline]
    pub fn has_data_descriptor(&self) -> bool {
        self.flags & 0x08 != 0
    }

    /// Returns the central-directory general-purpose bit flags.
    #[inline]
    pub(crate) fn flags(&self) -> u16 {
        self.flags
    }

    /// Describes where the file's data is located within the archive.
    #[inline]
    pub fn wayfinder(&self) -> ZipArchiveEntryWayfinder {
        ZipArchiveEntryWayfinder {
            uncompressed_size: self.uncompressed_size,
            compressed_size: self.compressed_size,
            local_header_offset: self.local_header_offset,
            has_data_descriptor: self.has_data_descriptor(),
            crc: self.crc32,
            flags: self.flags,
            compression_method: self.compression_method,
            zip64_sizes: self.zip64_sizes,
            zip64_uncompressed_size_resolved: self.zip64_uncompressed_size_resolved,
            zip64_compressed_size_resolved: self.zip64_compressed_size_resolved,
            zip64_local_header_offset_resolved: self.zip64_local_header_offset_resolved,
            zip64_disk_start_resolved: self.zip64_disk_start_resolved,
            disk_number_start: self.disk_number_start,
        }
    }

    /// The purported number of bytes of the uncompressed data.
    ///
    /// **WARNING**: this number has not yet been validated, so don't trust it
    /// to make allocation decisions.
    #[inline]
    pub fn uncompressed_size_hint(&self) -> u64 {
        self.uncompressed_size
    }

    /// The purported number of bytes of the compressed data.
    ///
    /// **WARNING**: this number has not yet been validated, so don't trust it
    /// to make allocation decisions.
    #[inline]
    pub fn compressed_size_hint(&self) -> u64 {
        self.compressed_size
    }

    /// Returns the number of variable central-directory bytes for this entry.
    ///
    /// The count includes the raw member name, extra fields, and file comment.
    /// It is available without allocating or normalizing any of those fields.
    #[inline]
    pub fn metadata_size_hint(&self) -> u64 {
        u64::from(self.file_name_len)
            + u64::from(self.extra_field_len)
            + u64::from(self.file_comment_len)
    }

    /// Whether the central-directory record declares traditional ZIP
    /// encryption through general-purpose flag bit zero.
    #[inline]
    pub fn is_encrypted(&self) -> bool {
        self.flags & 1 != 0
    }

    /// The declared offset to the local file header within the Zip archive.
    ///
    /// To verify the validity of this offset, call
    /// [`ZipSliceArchive::get_entry`] or [`ZipArchive::get_entry`].
    ///
    /// The minimum of all local header offsets (or `directory_offset()` when a
    /// zip is empty), will be the length of prelude data in a zip archive (data
    /// that is unrelated to the zip archive).
    ///
    /// See [`RangeReader`] for an example.
    #[inline]
    pub fn local_header_offset(&self) -> u64 {
        self.local_header_offset
    }

    /// The compression method used to compress the data
    #[inline]
    pub fn compression_method(&self) -> CompressionMethod {
        self.compression_method.as_method()
    }

    /// Returns the file path in its raw form.
    ///
    /// # Safety
    ///
    /// The raw path may contain unsafe components like:
    /// - Absolute paths (`/etc/passwd`)
    /// - Directory traversal (`../../../etc/passwd`)
    /// - Invalid UTF-8 sequences
    ///
    /// # Example
    /// ```rust
    /// # use soapberry_zip::ZipArchive;
    /// # fn example() -> Result<(), Box<dyn std::error::Error>> {
    /// # let data = include_bytes!("../assets/test.zip");
    /// # let archive = ZipArchive::from_slice(data)?;
    /// # let mut entries = archive.entries();
    /// # let entry = entries.next_entry()?.unwrap();
    /// // Get raw path (potentially unsafe)
    /// let raw_path = entry.file_path();
    ///
    /// // Convert to safe path
    /// let safe_path = raw_path.try_normalize()?;
    /// println!("Safe path: {}", safe_path.as_ref());
    ///
    /// // Check if it's a directory
    /// if safe_path.is_dir() {
    ///     println!("This is a directory");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[inline]
    pub fn file_path(&self) -> ZipFilePath<RawPath<'a>> {
        self.file_name
    }

    /// Returns the last modification date and time.
    ///
    /// This method parses the extra field data to locate more accurate timestamps.
    #[inline]
    pub fn last_modified(&self) -> ZipDateTimeKind {
        extract_best_timestamp(self.extra_fields(), self.last_mod_time, self.last_mod_date)
    }

    /// Returns the file mode information extracted from the external file attributes.
    #[inline]
    pub fn mode(&self) -> EntryMode {
        let creator_version = self.version_made_by >> 8;

        let mut mode = match creator_version {
            // Unix and macOS
            CREATOR_UNIX | CREATOR_MACOS => unix_mode_to_file_mode(self.external_file_attrs >> 16),
            // NTFS, VFAT, FAT
            CREATOR_NTFS | CREATOR_VFAT | CREATOR_FAT => {
                msdos_mode_to_file_mode(self.external_file_attrs)
            },
            // default to basic permissions
            _ => 0o644,
        };

        // Check if it's a directory by filename ending with '/'
        if self.is_dir() {
            mode |= 0o040000; // S_IFDIR
        }

        EntryMode::new(mode)
    }

    /// The declared CRC32 checksum of the uncompressed data.
    ///
    /// To verify the validity of this value, [`ZipEntry::verifying_reader`]
    /// will return an error if when the decompressed data does not match this
    /// checksum.
    #[inline]
    pub fn crc32(&self) -> u32 {
        self.crc32
    }

    /// Whether this central-directory record uses ZIP64 fields.
    #[inline]
    pub fn is_zip64(&self) -> bool {
        self.is_zip64
    }

    /// Returns the absolute source range containing the ZIP64 local-header
    /// offset value in this central-directory record.
    ///
    /// The range is present only when the record's 32-bit local-header offset
    /// is the ZIP64 sentinel and the corresponding value was resolved from a
    /// single, valid ZIP64 extra field.  It is expressed as a half-open byte
    /// range in the same source coordinate system as
    /// [`Self::central_directory_offset`].
    #[inline]
    pub fn zip64_local_header_offset_range(&self) -> Option<Range<u64>> {
        self.zip64_local_header_offset_range.clone()
    }

    /// Returns the offset from the start of reader where this central directory
    /// record was parsed from.
    #[inline]
    pub fn central_directory_offset(&self) -> u64 {
        self.central_directory_offset
    }

    /// Returns an iterator over the extra fields in this file header record.
    ///
    /// Extra fields contain additional metadata about files in ZIP archives,
    /// such as timestamps, alignment information, and platform-specific data.
    ///
    /// # Examples
    ///
    /// ```rust
    /// # use soapberry_zip::{ZipArchive, extra_fields::ExtraFieldId};
    /// # fn example(data: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    /// let archive = ZipArchive::from_slice(data)?;
    /// for entry_result in archive.entries() {
    ///     let entry = entry_result?;
    ///     let mut extra_fields = entry.extra_fields();
    ///     for (field_id, field_data) in extra_fields.by_ref() {
    ///         match field_id {
    ///             ExtraFieldId::JAVA_JAR => {
    ///                 println!("Handle jar CAFE field with {} bytes", field_data.len());
    ///             }
    ///             _ => {
    ///                 println!("Found extra field ID: 0x{:04x}", field_id.as_u16());
    ///             }
    ///         }
    ///     }
    ///
    ///     // If desired, check for truncated data
    ///     if !extra_fields.remaining_bytes().is_empty() {
    ///         println!("Warning: Some extra field data was truncated");
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// Raw access to the entire extra field data is available when
    /// `remaining_bytes` is called prior to any iteration.
    #[inline]
    pub fn extra_fields(&self) -> ExtraFields<'_> {
        ExtraFields::new(self.extra_field)
    }
}

/// Contains directions to where the Zip entry's data is located within the Zip archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ZipArchiveEntryWayfinder {
    uncompressed_size: u64,
    compressed_size: u64,
    local_header_offset: u64,
    crc: u32,
    has_data_descriptor: bool,
    flags: u16,
    compression_method: CompressionMethodId,
    zip64_sizes: bool,
    zip64_uncompressed_size_resolved: bool,
    zip64_compressed_size_resolved: bool,
    zip64_local_header_offset_resolved: bool,
    zip64_disk_start_resolved: bool,
    disk_number_start: u32,
}

impl ZipArchiveEntryWayfinder {
    /// Equivalent to [`ZipFileHeaderRecord::compressed_size_hint`]
    ///
    /// This is a convenience method to avoid having to deal with lifetime
    /// issues on a `ZipFileHeaderRecord`
    #[inline]
    pub fn uncompressed_size_hint(&self) -> u64 {
        self.uncompressed_size
    }

    /// Equivalent to [`ZipFileHeaderRecord::compressed_size_hint`]
    ///
    /// This is a convenience method to avoid having to deal with lifetime
    /// issues on a `ZipFileHeaderRecord`
    #[inline]
    pub fn compressed_size_hint(&self) -> u64 {
        self.compressed_size
    }

    /// Returns the local-header offset retained by this wayfinder.
    #[inline]
    pub(crate) fn local_header_offset(&self) -> u64 {
        self.local_header_offset
    }

    pub(crate) fn borrowed_provenance_supported(&self) -> bool {
        self.zip64_uncompressed_size_resolved
            && self.zip64_compressed_size_resolved
            && self.zip64_local_header_offset_resolved
            && self.zip64_disk_start_resolved
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ZipLocalFileHeaderFixed {
    pub(crate) signature: u32,
    pub(crate) version_needed: u16,
    pub(crate) flags: u16,
    pub(crate) compression_method: CompressionMethodId,
    pub(crate) last_mod_time: u16,
    pub(crate) last_mod_date: u16,
    pub(crate) crc32: u32,
    pub(crate) compressed_size: u32,
    pub(crate) uncompressed_size: u32,
    pub(crate) file_name_len: u16,
    pub(crate) extra_field_len: u16,
}

impl ZipLocalFileHeaderFixed {
    pub(crate) const SIZE: usize = 30;
    pub const SIGNATURE: u32 = 0x04034b50;

    pub fn parse(data: &[u8]) -> Result<ZipLocalFileHeaderFixed, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let result = ZipLocalFileHeaderFixed {
            signature: le_u32(&data[0..4]),
            version_needed: le_u16(&data[4..6]),
            flags: le_u16(&data[6..8]),
            compression_method: CompressionMethodId(le_u16(&data[8..10])),
            last_mod_time: le_u16(&data[10..12]),
            last_mod_date: le_u16(&data[12..14]),
            crc32: le_u32(&data[14..18]),
            compressed_size: le_u32(&data[18..22]),
            uncompressed_size: le_u32(&data[22..26]),
            file_name_len: le_u16(&data[26..28]),
            extra_field_len: le_u16(&data[28..30]),
        };

        if result.signature != Self::SIGNATURE {
            return Err(Error::from(ErrorKind::InvalidSignature {
                expected: Self::SIGNATURE,
                actual: result.signature,
            }));
        }

        Ok(result)
    }

    pub fn variable_length(&self) -> usize {
        self.file_name_len as usize + self.extra_field_len as usize
    }

    pub fn write<W>(&self, mut writer: W) -> Result<(), Error>
    where
        W: Write,
    {
        // Batch writes with a fixed size buffer. Improved throughput 25%
        let mut buffer = [0u8; 30];
        buffer[..4].copy_from_slice(&self.signature.to_le_bytes());
        buffer[4..6].copy_from_slice(&self.version_needed.to_le_bytes());
        buffer[6..8].copy_from_slice(&self.flags.to_le_bytes());
        buffer[8..10].copy_from_slice(&self.compression_method.0.to_le_bytes());
        buffer[10..12].copy_from_slice(&self.last_mod_time.to_le_bytes());
        buffer[12..14].copy_from_slice(&self.last_mod_date.to_le_bytes());
        buffer[14..18].copy_from_slice(&self.crc32.to_le_bytes());
        buffer[18..22].copy_from_slice(&self.compressed_size.to_le_bytes());
        buffer[22..26].copy_from_slice(&self.uncompressed_size.to_le_bytes());
        buffer[26..28].copy_from_slice(&self.file_name_len.to_le_bytes());
        buffer[28..30].copy_from_slice(&self.extra_field_len.to_le_bytes());
        writer.write_all(&buffer)?;
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ZipFileHeaderFixed {
    pub signature: u32,
    pub version_made_by: u16,
    pub version_needed: u16,
    pub flags: u16,
    pub compression_method: CompressionMethodId,
    pub last_mod_time: u16,
    pub last_mod_date: u16,
    pub crc32: u32,
    pub compressed_size: u32,
    pub uncompressed_size: u32,
    pub file_name_len: u16,
    pub extra_field_len: u16,
    pub file_comment_len: u16,
    pub disk_number_start: u16,
    pub internal_file_attrs: u16,
    pub external_file_attrs: u32,
    pub local_header_offset: u32,
}

impl ZipFileHeaderFixed {
    pub fn variable_length(&self) -> usize {
        self.file_name_len as usize + self.extra_field_len as usize + self.file_comment_len as usize
    }
}

type VariableFields<'a> = (
    &'a [u8], // file_name
    &'a [u8], // extra_field
    &'a [u8], // file_comment
    &'a [u8], // rest of the data
);

impl ZipFileHeaderFixed {
    pub(crate) const SIZE: usize = 46;

    #[inline]
    pub fn parse(data: &[u8]) -> Result<ZipFileHeaderFixed, Error> {
        if data.len() < Self::SIZE {
            return Err(Error::from(ErrorKind::Eof));
        }

        let result = ZipFileHeaderFixed {
            signature: le_u32(&data[0..4]),
            version_made_by: le_u16(&data[4..6]),
            version_needed: le_u16(&data[6..8]),
            flags: le_u16(&data[8..10]),
            compression_method: CompressionMethodId(le_u16(&data[10..12])),
            last_mod_time: le_u16(&data[12..14]),
            last_mod_date: le_u16(&data[14..16]),
            crc32: le_u32(&data[16..20]),
            compressed_size: le_u32(&data[20..24]),
            uncompressed_size: le_u32(&data[24..28]),
            file_name_len: le_u16(&data[28..30]),
            extra_field_len: le_u16(&data[30..32]),
            file_comment_len: le_u16(&data[32..34]),
            disk_number_start: le_u16(&data[34..36]),
            internal_file_attrs: le_u16(&data[36..38]),
            external_file_attrs: le_u32(&data[38..42]),
            local_header_offset: le_u32(&data[42..46]),
        };

        if result.signature != CENTRAL_HEADER_SIGNATURE {
            return Err(Error::from(ErrorKind::InvalidSignature {
                expected: CENTRAL_HEADER_SIGNATURE,
                actual: result.signature,
            }));
        }

        Ok(result)
    }

    #[inline]
    fn parse_variable_length<'a>(&self, data: &'a [u8]) -> Option<VariableFields<'a>> {
        if data.len() < self.file_name_len as usize {
            return None;
        }
        let (file_name, rest) = data.split_at(self.file_name_len as usize);

        if rest.len() < self.extra_field_len as usize {
            return None;
        }
        let (extra_field, rest) = rest.split_at(self.extra_field_len as usize);

        if rest.len() < self.file_comment_len as usize {
            return None;
        }
        let (file_comment, rest) = rest.split_at(self.file_comment_len as usize);

        Some((file_name, extra_field, file_comment, rest))
    }

    pub fn write<W>(&self, mut writer: W) -> Result<(), Error>
    where
        W: Write,
    {
        // Batch writes with a fixed size buffer. Improved throughput 25%
        let mut buffer = [0u8; Self::SIZE];
        buffer[0..4].copy_from_slice(&self.signature.to_le_bytes());
        buffer[4..6].copy_from_slice(&self.version_made_by.to_le_bytes());
        buffer[6..8].copy_from_slice(&self.version_needed.to_le_bytes());
        buffer[8..10].copy_from_slice(&self.flags.to_le_bytes());
        buffer[10..12].copy_from_slice(&self.compression_method.0.to_le_bytes());
        buffer[12..14].copy_from_slice(&self.last_mod_time.to_le_bytes());
        buffer[14..16].copy_from_slice(&self.last_mod_date.to_le_bytes());
        buffer[16..20].copy_from_slice(&self.crc32.to_le_bytes());
        buffer[20..24].copy_from_slice(&self.compressed_size.to_le_bytes());
        buffer[24..28].copy_from_slice(&self.uncompressed_size.to_le_bytes());
        buffer[28..30].copy_from_slice(&self.file_name_len.to_le_bytes());
        buffer[30..32].copy_from_slice(&self.extra_field_len.to_le_bytes());
        buffer[32..34].copy_from_slice(&self.file_comment_len.to_le_bytes());
        buffer[34..36].copy_from_slice(&self.disk_number_start.to_le_bytes());
        buffer[36..38].copy_from_slice(&self.internal_file_attrs.to_le_bytes());
        buffer[38..42].copy_from_slice(&self.external_file_attrs.to_le_bytes());
        buffer[42..46].copy_from_slice(&self.local_header_offset.to_le_bytes());
        writer.write_all(&buffer)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn complete_descriptor_rejects_ambiguous_signature() {
        let signature = DataDescriptor::SIGNATURE;
        let entry = ZipArchiveEntryWayfinder {
            uncompressed_size: u64::from(signature),
            compressed_size: u64::from(signature),
            local_header_offset: 0,
            crc: signature,
            has_data_descriptor: true,
            flags: 0,
            compression_method: CompressionMethodId(0),
            zip64_sizes: false,
            zip64_uncompressed_size_resolved: true,
            zip64_compressed_size_resolved: true,
            zip64_local_header_offset_resolved: true,
            zip64_disk_start_resolved: true,
            disk_number_start: 0,
        };
        let descriptor = signature.to_le_bytes().repeat(4);
        let error = DataDescriptor::parse_complete(&descriptor, &entry, false).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    fn descriptor_entry(
        crc: u32,
        compressed_size: u64,
        uncompressed_size: u64,
        zip64_sizes: bool,
    ) -> ZipArchiveEntryWayfinder {
        ZipArchiveEntryWayfinder {
            uncompressed_size,
            compressed_size,
            local_header_offset: 0,
            crc,
            has_data_descriptor: true,
            flags: 0x08,
            compression_method: CompressionMethodId(0),
            zip64_sizes,
            zip64_uncompressed_size_resolved: true,
            zip64_compressed_size_resolved: true,
            zip64_local_header_offset_resolved: true,
            zip64_disk_start_resolved: true,
            disk_number_start: 0,
        }
    }

    fn descriptor_bytes(
        width: usize,
        signed: bool,
        crc: u32,
        compressed_size: u64,
        uncompressed_size: u64,
    ) -> Vec<u8> {
        let mut descriptor = Vec::new();
        if signed {
            descriptor.extend_from_slice(&DataDescriptor::SIGNATURE.to_le_bytes());
        }
        descriptor.extend_from_slice(&crc.to_le_bytes());
        match width {
            4 => {
                descriptor.extend_from_slice(
                    &u32::try_from(compressed_size)
                        .expect("test descriptor compressed size fits ZIP32")
                        .to_le_bytes(),
                );
                descriptor.extend_from_slice(
                    &u32::try_from(uncompressed_size)
                        .expect("test descriptor uncompressed size fits ZIP32")
                        .to_le_bytes(),
                );
            },
            8 => {
                descriptor.extend_from_slice(&compressed_size.to_le_bytes());
                descriptor.extend_from_slice(&uncompressed_size.to_le_bytes());
            },
            _ => panic!("unsupported test descriptor width"),
        }
        descriptor
    }

    fn push_u16(output: &mut Vec<u8>, value: u16) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(output: &mut Vec<u8>, value: u32) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u64(output: &mut Vec<u8>, value: u64) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn archive_zip64_descriptor_fixture(width: usize, signed: bool) -> Vec<u8> {
        let payload = b"hello";
        let name = b"x";
        let crc = crate::crc32(payload);
        let size = u32::try_from(payload.len()).unwrap();
        let mut archive = Vec::new();

        // Local header with ordinary (zero) descriptor sizes.
        push_u32(&mut archive, 0x0403_4b50);
        push_u16(&mut archive, if width == 8 { 45 } else { 20 });
        push_u16(&mut archive, 0x08);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u16(&mut archive, u16::try_from(name.len()).unwrap());
        push_u16(&mut archive, 0);
        archive.extend_from_slice(name);
        archive.extend_from_slice(payload);
        archive.extend_from_slice(&descriptor_bytes(
            width,
            signed,
            crc,
            u64::from(size),
            u64::from(size),
        ));

        let central_directory_offset = u64::try_from(archive.len()).unwrap();
        let mut central = Vec::new();
        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, if width == 8 { 45 } else { 20 });
        push_u16(&mut central, if width == 8 { 45 } else { 20 });
        push_u16(&mut central, 0x08);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc);
        push_u32(&mut central, size);
        push_u32(&mut central, size);
        push_u16(&mut central, u16::try_from(name.len()).unwrap());
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, 0);
        central.extend_from_slice(name);
        let central_directory_size = u64::try_from(central.len()).unwrap();
        archive.extend_from_slice(&central);

        // A ZIP64 EOCD is required by the classic central-directory offset
        // sentinel, while this member itself remains entirely ZIP32.
        let zip64_eocd_offset = u64::try_from(archive.len()).unwrap();
        push_u32(&mut archive, END_OF_CENTRAL_DIR_SIGNATURE64);
        push_u64(&mut archive, 44);
        push_u16(&mut archive, 45);
        push_u16(&mut archive, 45);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u64(&mut archive, 1);
        push_u64(&mut archive, 1);
        push_u64(&mut archive, central_directory_size);
        push_u64(&mut archive, central_directory_offset);
        push_u32(&mut archive, END_OF_CENTRAL_DIR_LOCATOR_SIGNATURE);
        push_u32(&mut archive, 0);
        push_u64(&mut archive, zip64_eocd_offset);
        push_u32(&mut archive, 1);
        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 1);
        push_u16(&mut archive, 1);
        push_u32(&mut archive, u32::try_from(central_directory_size).unwrap());
        push_u32(&mut archive, u32::MAX);
        push_u16(&mut archive, 0);
        archive
    }

    #[test]
    fn complete_descriptor_uses_archive_zip64_context_without_changing_zip32_empty_entries() {
        let crc = 0x3610_a686;
        let entry = descriptor_entry(crc, 5, 5, false);
        for signed in [false, true] {
            for width in [4, 8] {
                let descriptor = descriptor_bytes(width, signed, crc, 5, 5);
                let parsed = DataDescriptor::parse_complete(&descriptor, &entry, true).unwrap();
                assert_eq!(parsed.crc, crc);
                assert_eq!(parsed.compressed_size, 5);
                assert_eq!(parsed.uncompressed_size, 5);
                assert_eq!(
                    parsed.encoded_size,
                    width * 2 + 4 + if signed { 4 } else { 0 }
                );
            }
        }

        let empty_entry = descriptor_entry(0, 0, 0, false);
        let mut descriptor = descriptor_bytes(4, false, 0, 0, 0);
        descriptor.extend_from_slice(&[0; 8]);
        let parsed = DataDescriptor::parse_complete(&descriptor, &empty_entry, false).unwrap();
        assert_eq!(parsed.encoded_size, 12);
    }

    #[test]
    fn strict_layout_accepts_archive_zip64_descriptors_at_reader_boundaries() {
        for signed in [false, true] {
            for width in [4, 8] {
                let data = archive_zip64_descriptor_fixture(width, signed);
                let archive = ZipArchive::from_slice(data.as_slice()).unwrap();
                assert!(archive.is_zip64());
                let entry = archive.entries().next_entry().unwrap().unwrap();
                assert!(!entry.is_zip64());
                let layout = archive
                    .validate_strict_entry_layout(entry.wayfinder(), b"x")
                    .unwrap();
                assert_eq!(layout.data_end_offset - layout.data_start_offset, 5);
                assert_eq!(layout.span_end, archive.directory_offset());

                let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
                let reader_archive =
                    ZipArchive::from_seekable(Cursor::new(data.as_slice()), &mut buffer).unwrap();
                let wayfinder = {
                    let mut entries = reader_archive.entries(&mut buffer);
                    entries.next_entry().unwrap().unwrap().wayfinder()
                };
                let layout = reader_archive
                    .validate_strict_entry_layout(wayfinder, b"x")
                    .unwrap();
                assert_eq!(layout.span_end, reader_archive.directory_offset());
            }
        }
    }

    #[test]
    fn complete_descriptor_requires_zip64_width_for_zip64_size_sentinels() {
        let entry = descriptor_entry(0x3610_a686, 5, 5, true);
        for signed in [false, true] {
            let descriptor = descriptor_bytes(8, signed, entry.crc, 5, 5);
            assert!(DataDescriptor::parse_complete(&descriptor, &entry, false).is_ok());

            let descriptor = descriptor_bytes(4, signed, entry.crc, 5, 5);
            let error = DataDescriptor::parse_complete(&descriptor, &entry, false).unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::Eof));
        }
    }

    #[test]
    fn complete_descriptor_rejects_malformed_and_ambiguous_width_candidates() {
        let entry = descriptor_entry(0, 0, 0, false);
        let ambiguous = descriptor_bytes(4, false, 0, 0, 0);
        let mut ambiguous = ambiguous;
        ambiguous.extend_from_slice(&[0; 8]);
        let error = DataDescriptor::parse_complete(&ambiguous, &entry, true).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let malformed = descriptor_bytes(8, true, 0x3610_a686, 5, 5);
        let error = DataDescriptor::parse_complete(&malformed[..malformed.len() - 4], &entry, true)
            .unwrap_err();
        assert!(matches!(
            error.kind(),
            ErrorKind::InvalidChecksum { .. } | ErrorKind::InvalidSize { .. } | ErrorKind::Eof
        ));

        let entry = descriptor_entry(0x3610_a686, 5, 5, false);
        let malformed = descriptor_bytes(4, false, 0x3610_a686, 4, 5);
        let error = DataDescriptor::parse_complete(&malformed, &entry, true).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
    }

    use super::*;
    use std::io::{self, Cursor, Read};

    fn central_fixed_with_zip64_extra(
        uncompressed_size: u32,
        compressed_size: u32,
        disk_number_start: u16,
        local_header_offset: u32,
        extra: &[u8],
    ) -> ZipFileHeaderFixed {
        central_fixed_with_zip64_extra_name_len(
            4,
            uncompressed_size,
            compressed_size,
            disk_number_start,
            local_header_offset,
            extra,
        )
    }

    fn central_fixed_with_zip64_extra_name_len(
        name_len: u16,
        uncompressed_size: u32,
        compressed_size: u32,
        disk_number_start: u16,
        local_header_offset: u32,
        extra: &[u8],
    ) -> ZipFileHeaderFixed {
        let mut bytes = vec![0u8; ZipFileHeaderFixed::SIZE];
        bytes[0..4].copy_from_slice(&CENTRAL_HEADER_SIGNATURE.to_le_bytes());
        bytes[20..24].copy_from_slice(&compressed_size.to_le_bytes());
        bytes[24..28].copy_from_slice(&uncompressed_size.to_le_bytes());
        bytes[28..30].copy_from_slice(&name_len.to_le_bytes());
        bytes[30..32].copy_from_slice(&(extra.len() as u16).to_le_bytes());
        bytes[34..36].copy_from_slice(&disk_number_start.to_le_bytes());
        bytes[42..46].copy_from_slice(&local_header_offset.to_le_bytes());
        ZipFileHeaderFixed::parse(&bytes).unwrap()
    }

    fn zip64_extra(body: &[u8]) -> Vec<u8> {
        let mut extra = Vec::with_capacity(4 + body.len());
        extra.extend_from_slice(&ExtraFieldId::ZIP64.as_u16().to_le_bytes());
        extra.extend_from_slice(&(body.len() as u16).to_le_bytes());
        extra.extend_from_slice(body);
        extra
    }

    fn custom_extra(id: u16, body: &[u8]) -> Vec<u8> {
        let mut extra = Vec::with_capacity(4 + body.len());
        extra.extend_from_slice(&id.to_le_bytes());
        extra.extend_from_slice(&(body.len() as u16).to_le_bytes());
        extra.extend_from_slice(body);
        extra
    }

    #[test]
    fn central_zip64_extra_resolves_sentinels_and_returns_offset_field_range() {
        let mut body = Vec::new();
        body.extend_from_slice(&5u64.to_le_bytes());
        body.extend_from_slice(&4u64.to_le_bytes());
        body.extend_from_slice(&1234u64.to_le_bytes());
        let extra = zip64_extra(&body);
        let header = central_fixed_with_zip64_extra(u32::MAX, u32::MAX, 0, u32::MAX, &extra);
        let record = ZipFileHeaderRecord::from_parts(header, b"item", &extra, b"", 123).unwrap();

        assert!(record.is_zip64());
        assert_eq!(record.uncompressed_size_hint(), 5);
        assert_eq!(record.compressed_size_hint(), 4);
        assert_eq!(record.local_header_offset(), 1234);
        assert_eq!(
            record.zip64_local_header_offset_range(),
            Some(123 + ZipFileHeaderFixed::SIZE as u64 + 4 + 4 + 16..123 + 46 + 4 + 4 + 24)
        );
    }

    #[test]
    fn central_zip64_offset_range_tracks_variable_raw_file_names() {
        let names = [
            Vec::new(),
            b"item".to_vec(),
            vec![0xe6, 0x96, 0x87, 0xf0, 0x9f, 0x92, 0xa9],
            vec![0xa5; usize::from(u16::MAX)],
        ];
        let central_directory_offset = 123u64;
        let value = 0x1122_3344_5566_7788u64;
        let body = value.to_le_bytes();

        for name in names {
            let extra = zip64_extra(&body);
            let name_len = u16::try_from(name.len()).unwrap();
            let header =
                central_fixed_with_zip64_extra_name_len(name_len, 1, 1, 0, u32::MAX, &extra);
            let record = ZipFileHeaderRecord::from_parts(
                header,
                &name,
                &extra,
                b"",
                central_directory_offset,
            )
            .unwrap();
            let range = record
                .zip64_local_header_offset_range()
                .expect("ZIP64 local-header offset should have a patch range");
            assert_eq!(range.end - range.start, 8);

            let central_start = usize::try_from(central_directory_offset).unwrap();
            let extra_start = central_start + ZipFileHeaderFixed::SIZE + name.len();
            let mut source = vec![0u8; extra_start + extra.len()];
            source[extra_start..].copy_from_slice(&extra);
            let start = usize::try_from(range.start).unwrap();
            let end = usize::try_from(range.end).unwrap();
            assert_eq!(&source[start..end], &body);
        }
    }

    #[test]
    fn central_zip64_offset_range_skips_preceding_extra_fields() {
        let custom_fields = [
            custom_extra(0xcafe, &[]),
            custom_extra(0xbeef, &[1, 2, 3]),
            {
                let mut fields = custom_extra(0x4242, &[4, 5, 6, 7, 8]);
                fields.extend_from_slice(&custom_extra(0x5151, &[9, 10]));
                fields
            },
        ];
        let name = b"raw-name";
        let central_directory_offset = 211u64;
        let mut body = Vec::new();
        body.extend_from_slice(&0x0102_0304_0506_0708u64.to_le_bytes());
        body.extend_from_slice(&0x1112_1314_1516_1718u64.to_le_bytes());
        body.extend_from_slice(&0xa1a2_a3a4_a5a6_a7a8u64.to_le_bytes());
        let zip64 = zip64_extra(&body);

        for custom in custom_fields {
            let mut extra = custom.clone();
            extra.extend_from_slice(&zip64);
            let header = central_fixed_with_zip64_extra_name_len(
                u16::try_from(name.len()).unwrap(),
                u32::MAX,
                u32::MAX,
                0,
                u32::MAX,
                &extra,
            );
            let record = ZipFileHeaderRecord::from_parts(
                header,
                name,
                &extra,
                b"",
                central_directory_offset,
            )
            .unwrap();
            let range = record
                .zip64_local_header_offset_range()
                .expect("ZIP64 local-header offset should have a patch range");
            let range_start = central_directory_offset
                + ZipFileHeaderFixed::SIZE as u64
                + name.len() as u64
                + custom.len() as u64
                + 4
                + 16;
            assert_eq!(range, range_start..range_start + 8);

            let central_start = usize::try_from(central_directory_offset).unwrap();
            let extra_start = central_start + ZipFileHeaderFixed::SIZE + name.len();
            let mut source = vec![0u8; extra_start + extra.len()];
            source[extra_start..].copy_from_slice(&extra);
            let start = usize::try_from(range.start).unwrap();
            let end = usize::try_from(range.end).unwrap();
            assert_eq!(&source[start..end], &body[16..24]);
        }
    }

    #[test]
    fn central_zip64_extra_rejects_duplicate_truncated_and_missing_values() {
        let valid_body = 7u64.to_le_bytes();
        let valid = zip64_extra(&valid_body);

        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&valid);
        let header = central_fixed_with_zip64_extra(1, u32::MAX, 0, 0, &duplicate);
        let error = ZipFileHeaderRecord::from_parts(header, b"item", &duplicate, b"", 0)
            .expect_err("duplicate ZIP64 fields must be rejected");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let truncated = zip64_extra(&[1, 2, 3, 4]);
        let header = central_fixed_with_zip64_extra(1, u32::MAX, 0, 0, &truncated);
        let error = ZipFileHeaderRecord::from_parts(header, b"item", &truncated, b"", 0)
            .expect_err("truncated ZIP64 values must be rejected");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let missing = zip64_extra(&[]);
        let header = central_fixed_with_zip64_extra(1, u32::MAX, 0, 0, &missing);
        let error = ZipFileHeaderRecord::from_parts(header, b"item", &missing, b"", 0)
            .expect_err("missing ZIP64 values must be rejected");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[derive(Debug)]
    struct OverreportingReader;

    impl Read for OverreportingReader {
        fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
            Ok(buffer.len().saturating_add(1))
        }
    }

    #[test]
    pub fn blank_zip_archive() {
        let data = [80, 75, 5, 6];
        let mut buf = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(data), &mut buf);
        assert!(archive.is_err());
    }

    #[test]
    pub fn trunc_comment_zips() {
        let data = [
            80, 75, 6, 7, 21, 0, 0, 0, 34, 0, 0, 0, 0, 0, 0, 0, 10, 0, 59, 59, 80, 75, 5, 6, 0,
            255, 255, 255, 255, 255, 255, 0, 0, 0, 80, 75, 6, 6, 0, 0, 0, 10,
        ];
        let mut buf = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(data), &mut buf);
        assert!(archive.is_err());

        let archive = ZipArchive::from_slice(data);
        assert!(archive.is_err());
    }

    #[test]
    pub fn trunc_eocd64() {
        let data = [
            80, 75, 6, 7, 21, 0, 0, 0, 34, 0, 0, 0, 0, 0, 0, 0, 10, 0, 59, 59, 80, 75, 5, 6, 0,
            255, 255, 255, 255, 255, 255, 0, 0, 0, 80, 75, 6, 6, 0, 0, 6, 0, 0, 250, 255, 255, 255,
            255, 251, 0, 0, 0, 0, 80, 5, 6, 0, 0, 0, 0, 56, 0, 0, 0, 0, 10,
        ];

        let archive = ZipArchive::from_slice(data);
        assert!(archive.is_err());

        let mut buf = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(data), &mut buf);
        assert!(archive.is_err());
    }

    #[test]
    pub fn trunc_eocd_entry() {
        let data = [
            80, 75, 1, 2, 159, 159, 159, 159, 159, 159, 159, 159, 159, 0, 241, 205, 0, 80, 75, 5,
            6, 0, 48, 249, 0, 250, 255, 255, 255, 255, 251, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
            35, 0,
        ];

        assert!(matches!(
            ZipArchive::from_slice(data),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));

        let mut buf = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        assert!(matches!(
            ZipArchive::from_seekable(Cursor::new(data), &mut buf),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
    }

    #[test]
    fn verifier_wrappers_reject_overreported_reads_without_panicking() {
        let bytes = std::fs::read("assets/test.zip").unwrap();
        let slice_archive = ZipArchive::from_slice(bytes.clone()).unwrap();
        let mut slice_entries = slice_archive.entries();
        let wayfinder = slice_entries.next_entry().unwrap().unwrap().wayfinder();
        let slice_entry = slice_archive.get_entry(wayfinder).unwrap();
        let mut slice_verifier = slice_entry.verifying_reader(OverreportingReader);
        let mut output = [0u8; 1];
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            slice_verifier.read(&mut output)
        }))
        .expect("slice verifier must not panic");
        let error = result.unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);

        let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(bytes), &mut scratch).unwrap();
        let mut entries = archive.entries(&mut scratch);
        let wayfinder = entries.next_entry().unwrap().unwrap().wayfinder();
        let entry = archive.get_entry(wayfinder).unwrap();
        let mut verifier = entry.verifying_reader(OverreportingReader);
        let result =
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| verifier.read(&mut output)))
                .expect("seekable verifier must not panic");
        let error = result.unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
    }

    #[test]
    fn strict_verification_compares_a_zero_crc() {
        let expected = ZipVerification {
            crc: 0,
            uncompressed_size: 3,
        };
        assert!(
            expected
                .valid(ZipVerification {
                    crc: 7,
                    uncompressed_size: 3,
                })
                .is_ok()
        );
        assert!(matches!(
            expected.valid_strict(ZipVerification {
                crc: 7,
                uncompressed_size: 3,
            }),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. })
        ));
        assert!(
            expected
                .valid_strict(ZipVerification {
                    crc: 0,
                    uncompressed_size: 3,
                })
                .is_ok()
        );
        assert!(
            (ZipVerification {
                crc: 0,
                uncompressed_size: 0,
            })
            .valid_strict(ZipVerification {
                crc: 0,
                uncompressed_size: 0,
            })
            .is_ok()
        );
    }

    #[test]
    fn test_compressed_data_range() {
        let test_zip = std::fs::read("assets/test.zip").unwrap();

        // Test ZipSliceEntry API (from slice)
        let slice_archive = ZipArchive::from_slice(&test_zip).unwrap();
        let slice_header_records: Vec<_> = slice_archive
            .entries()
            .collect::<Result<Vec<_>, _>>()
            .unwrap();
        assert_eq!(slice_header_records.len(), 2);

        let entry1_wayfinder = slice_header_records[0].wayfinder();
        let slice_entry1 = slice_archive.get_entry(entry1_wayfinder).unwrap();
        let slice_range1 = slice_entry1.compressed_data_range();
        assert_eq!(
            slice_range1,
            (66, 91),
            "test.txt compressed data should be at bytes 66-91"
        );

        let entry2_wayfinder = slice_header_records[1].wayfinder();
        let slice_entry2 = slice_archive.get_entry(entry2_wayfinder).unwrap();
        let slice_range2 = slice_entry2.compressed_data_range();
        assert_eq!(
            slice_range2,
            (169, 954),
            "gophercolor16x16.png compressed data should be at bytes 169-954"
        );

        // Test ZipEntry API
        let file = std::fs::File::open("assets/test.zip").unwrap();
        let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
        let reader_archive = ZipArchive::from_file(file, &mut buffer).unwrap();

        // Get wayfinders from the slice archive since they should be identical
        let reader_entry1 = reader_archive.get_entry(entry1_wayfinder).unwrap();
        let reader_range1 = reader_entry1.compressed_data_range();

        let reader_entry2 = reader_archive.get_entry(entry2_wayfinder).unwrap();
        let reader_range2 = reader_entry2.compressed_data_range();

        // Verify both APIs return identical ranges
        assert_eq!(slice_range1, reader_range1);
        assert_eq!(slice_range2, reader_range2);
    }
}
