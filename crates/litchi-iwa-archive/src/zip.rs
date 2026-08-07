use std::ops::Range;

use litchi_iwa_core::{Archive, SnappyStream};
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{ZipArchive as RawZipArchive, ZipFileHeaderRecord};

use crate::catalog::Component;
use crate::{Error, Limits, Result};

/// Opaque ZIP reader used by the physical component catalog.
pub(crate) struct ZipArchive<'data> {
    reader: ArchiveReader<'data>,
    physical_entries: Vec<PhysicalEntry>,
    central_order: Vec<usize>,
    data: &'data [u8],
    directory_offset: usize,
    eocd_offset: usize,
    base_offset: u64,
}

/// The fields of one physical ZIP header retained independently of the
/// dependency's normalized high-level view.
#[derive(Debug, Clone)]
pub(crate) struct PhysicalHeader {
    pub(crate) version_needed: u16,
    pub(crate) flags: u16,
    pub(crate) compression_method: u16,
    pub(crate) last_mod_time: u16,
    pub(crate) last_mod_date: u16,
    pub(crate) name: Box<[u8]>,
    pub(crate) extra: Box<[u8]>,
    pub(crate) comment: Box<[u8]>,
}

/// One ZIP member with ranges into the immutable source archive.
#[derive(Debug, Clone)]
pub(crate) struct PhysicalEntry {
    name: String,
    local_record: Range<usize>,
    compressed_data: Range<usize>,
    central_record: Range<usize>,
    local_header: PhysicalHeader,
    central_header: PhysicalHeader,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
}

impl PhysicalEntry {
    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    pub(crate) fn is_directory(&self) -> bool {
        self.central_header.name.last() == Some(&b'/')
    }

    pub(crate) fn is_encrypted(&self) -> bool {
        self.central_header.flags & 0x0001 != 0
    }

    pub(crate) fn is_supported(&self) -> bool {
        matches!(self.central_header.compression_method, 0 | 8)
    }

    pub(crate) fn local_header(&self) -> &PhysicalHeader {
        &self.local_header
    }

    pub(crate) fn central_header(&self) -> &PhysicalHeader {
        &self.central_header
    }

    pub(crate) fn raw_name(&self) -> &[u8] {
        &self.central_header.name
    }

    pub(crate) fn local_record(&self) -> Range<usize> {
        self.local_record.clone()
    }

    pub(crate) fn compressed_data_range(&self) -> Range<usize> {
        self.compressed_data.clone()
    }

    pub(crate) fn central_record(&self) -> Range<usize> {
        self.central_record.clone()
    }

    pub(crate) fn compressed_data<'a>(&self, data: &'a [u8]) -> &'a [u8] {
        &data[self.compressed_data.clone()]
    }

    pub(crate) fn compressed_size(&self) -> u64 {
        self.compressed_size
    }

    pub(crate) fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }

    pub(crate) fn crc32(&self) -> u32 {
        self.crc32
    }
}

impl<'data> ZipArchive<'data> {
    pub(crate) fn new_with_limits(data: &'data [u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let input_size = u64::try_from(data.len()).map_err(|_error| {
            Error::InvalidBundle("ZIP input length does not fit u64".to_owned())
        })?;
        validated_limits.check_input_size(input_size, "ZIP input")?;
        let raw_archive = RawZipArchive::from_slice(data)?;
        let physical_entries = parse_physical_entries(data, &raw_archive)?;
        let mut central_order = Vec::new();
        central_order
            .try_reserve_exact(physical_entries.len())
            .map_err(|_error| Error::Allocation {
                resource: "physical ZIP central order",
                amount: physical_entries.len(),
            })?;
        central_order.extend(0..physical_entries.len());
        central_order.sort_unstable_by_key(|&index| physical_entries[index].central_record.start);

        let directory_offset = checked_offset(raw_archive.directory_offset(), "central directory")?;
        let eocd_offset = checked_offset(raw_archive.eocd_offset(), "end of central directory")?;
        let tail = data.get(eocd_offset..).ok_or_else(|| {
            Error::InvalidBundle("ZIP end of central directory is truncated".to_owned())
        })?;
        let raw_central_offset = read_u32(tail, 16, "end of central directory offset")?;
        let base_offset = u64::try_from(directory_offset)
            .map_err(|_error| {
                Error::InvalidBundle("ZIP central directory offset does not fit u64".to_owned())
            })?
            .checked_sub(u64::from(raw_central_offset))
            .ok_or_else(|| {
                Error::InvalidBundle("ZIP central directory offset has an invalid base".to_owned())
            })?;
        Ok(Self {
            reader: ArchiveReader::new_with_limits(data, validated_limits.zip_limits())?,
            physical_entries,
            central_order,
            data,
            directory_offset,
            eocd_offset,
            base_offset,
        })
    }

    pub(crate) fn file_names(&self) -> impl Iterator<Item = &str> {
        self.reader.file_names()
    }

    pub(crate) fn read(&self, name: &str) -> Result<Vec<u8>> {
        Ok(self.reader.read(name)?)
    }

    pub(crate) fn read_entry(&self, entry: &PhysicalEntry) -> Result<Vec<u8>> {
        self.read(entry.name())
    }

    pub(crate) fn source(&self) -> &[u8] {
        self.data
    }

    pub(crate) fn physical_entries(&self) -> impl Iterator<Item = &PhysicalEntry> {
        self.physical_entries.iter()
    }

    pub(crate) fn physical_entry(&self, index: usize) -> Option<&PhysicalEntry> {
        self.physical_entries.get(index)
    }

    pub(crate) fn physical_indices_in_central_order(&self) -> impl Iterator<Item = usize> + '_ {
        self.central_order.iter().copied()
    }

    pub(crate) const fn directory_offset(&self) -> usize {
        self.directory_offset
    }

    pub(crate) const fn eocd_offset(&self) -> usize {
        self.eocd_offset
    }

    pub(crate) const fn base_offset(&self) -> u64 {
        self.base_offset
    }
}

fn parse_physical_entries(
    data: &[u8],
    archive: &soapberry_zip::ZipSliceArchive<&[u8]>,
) -> Result<Vec<PhysicalEntry>> {
    let mut entries = Vec::new();
    for result in archive.entries() {
        let entry = result?;
        entries.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "physical ZIP entry metadata",
            amount: 1,
        })?;
        entries.push(parse_physical_entry(data, &entry)?);
    }

    entries.sort_unstable_by_key(|entry| entry.local_record.start);
    let directory_offset = checked_offset(archive.directory_offset(), "central directory")?;
    for index in 0..entries.len() {
        let local_end = entries
            .get(index + 1)
            .map_or(directory_offset, |entry| entry.local_record.start);
        let entry = &mut entries[index];
        if entry.local_record.start >= local_end
            || entry.compressed_data.end > local_end
            || local_end > data.len()
        {
            return Err(Error::InvalidBundle(
                "ZIP local member records overlap or extend into the central directory".to_owned(),
            ));
        }
        entry.local_record.end = local_end;
    }
    Ok(entries)
}

fn parse_physical_entry(data: &[u8], entry: &ZipFileHeaderRecord<'_>) -> Result<PhysicalEntry> {
    let local_start = checked_offset(entry.local_header_offset(), "local header")?;
    let central_start =
        checked_offset(entry.central_directory_offset(), "central directory record")?;

    let local_fixed = checked_slice(data, local_start, 30, "local file header")?;
    if read_u32(local_fixed, 0, "local file header signature")? != 0x0403_4b50 {
        return Err(Error::InvalidBundle(
            "ZIP local file header has an invalid signature".to_owned(),
        ));
    }
    let local_name_len = usize::from(read_u16(local_fixed, 26, "local file name length")?);
    let local_extra_len = usize::from(read_u16(local_fixed, 28, "local extra length")?);
    let local_variable_len = local_name_len.checked_add(local_extra_len).ok_or_else(|| {
        Error::InvalidBundle("ZIP local header length overflows usize".to_owned())
    })?;
    let local_header_len = 30usize.checked_add(local_variable_len).ok_or_else(|| {
        Error::InvalidBundle("ZIP local header length overflows usize".to_owned())
    })?;
    let local_header = checked_slice(data, local_start, local_header_len, "local file header")?;
    let local_name_start = 30;
    let local_extra_start = local_name_start + local_name_len;
    let local_header_metadata = PhysicalHeader {
        version_needed: read_u16(local_fixed, 4, "local version")?,
        flags: read_u16(local_fixed, 6, "local flags")?,
        compression_method: read_u16(local_fixed, 8, "local compression method")?,
        last_mod_time: read_u16(local_fixed, 10, "local modification time")?,
        last_mod_date: read_u16(local_fixed, 12, "local modification date")?,
        name: copy_slice(
            &local_header[local_name_start..local_name_start + local_name_len],
            "local file name",
        )?,
        extra: copy_slice(
            &local_header[local_extra_start..local_extra_start + local_extra_len],
            "local extra fields",
        )?,
        comment: Box::default(),
    };

    let central_fixed = checked_slice(data, central_start, 46, "central directory record")?;
    if read_u32(central_fixed, 0, "central directory signature")? != 0x0201_4b50 {
        return Err(Error::InvalidBundle(
            "ZIP central directory record has an invalid signature".to_owned(),
        ));
    }
    let central_name_len = usize::from(read_u16(central_fixed, 28, "central file name length")?);
    let central_extra_len = usize::from(read_u16(central_fixed, 30, "central extra length")?);
    let central_comment_len =
        usize::from(read_u16(central_fixed, 32, "central file comment length")?);
    let central_variable_len = central_name_len
        .checked_add(central_extra_len)
        .and_then(|length| length.checked_add(central_comment_len))
        .ok_or_else(|| {
            Error::InvalidBundle("ZIP central directory record length overflows usize".to_owned())
        })?;
    let central_record_len = 46usize.checked_add(central_variable_len).ok_or_else(|| {
        Error::InvalidBundle("ZIP central directory record length overflows usize".to_owned())
    })?;
    let central_record = checked_slice(
        data,
        central_start,
        central_record_len,
        "central directory record",
    )?;
    let central_name_start = 46;
    let central_extra_start = central_name_start + central_name_len;
    let central_comment_start = central_extra_start + central_extra_len;
    let central_header = PhysicalHeader {
        version_needed: read_u16(central_fixed, 6, "central version")?,
        flags: read_u16(central_fixed, 8, "central flags")?,
        compression_method: read_u16(central_fixed, 10, "central compression method")?,
        last_mod_time: read_u16(central_fixed, 12, "central modification time")?,
        last_mod_date: read_u16(central_fixed, 14, "central modification date")?,
        name: copy_slice(
            &central_record[central_name_start..central_name_start + central_name_len],
            "central file name",
        )?,
        extra: copy_slice(
            &central_record[central_extra_start..central_extra_start + central_extra_len],
            "central extra fields",
        )?,
        comment: copy_slice(
            &central_record[central_comment_start..central_comment_start + central_comment_len],
            "central file comment",
        )?,
    };

    let compressed_start = local_start.checked_add(local_header_len).ok_or_else(|| {
        Error::InvalidBundle("ZIP compressed data offset overflows usize".to_owned())
    })?;
    let compressed_size = entry.compressed_size_hint();
    let compressed_len = usize::try_from(compressed_size).map_err(|_error| {
        Error::InvalidBundle("ZIP compressed member length does not fit usize".to_owned())
    })?;
    let compressed_end = compressed_start
        .checked_add(compressed_len)
        .ok_or_else(|| {
            Error::InvalidBundle("ZIP compressed data range overflows usize".to_owned())
        })?;
    checked_slice(
        data,
        compressed_start,
        compressed_len,
        "compressed member data",
    )?;

    let name = match std::str::from_utf8(&central_header.name) {
        Ok(name) => soapberry_zip::path::ZipFilePath::from_str(name)
            .as_ref()
            .to_owned(),
        Err(_error) => String::from_utf8_lossy(&central_header.name).into_owned(),
    };
    Ok(PhysicalEntry {
        name,
        local_record: local_start..compressed_end,
        compressed_data: compressed_start..compressed_end,
        central_record: central_start..central_start + central_record_len,
        local_header: local_header_metadata,
        central_header,
        compressed_size,
        uncompressed_size: entry.uncompressed_size_hint(),
        crc32: entry.crc32(),
    })
}

fn checked_offset(offset: u64, label: &str) -> Result<usize> {
    usize::try_from(offset)
        .map_err(|_error| Error::InvalidBundle(format!("ZIP {label} offset does not fit usize")))
}

fn checked_slice<'a>(data: &'a [u8], start: usize, length: usize, label: &str) -> Result<&'a [u8]> {
    let end = start
        .checked_add(length)
        .ok_or_else(|| Error::InvalidBundle(format!("ZIP {label} range overflows usize")))?;
    data.get(start..end)
        .ok_or_else(|| Error::InvalidBundle(format!("ZIP {label} is truncated")))
}

fn copy_slice(data: &[u8], label: &'static str) -> Result<Box<[u8]>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(data.len())
        .map_err(|_error| Error::Allocation {
            resource: label,
            amount: data.len(),
        })?;
    copy.extend_from_slice(data);
    Ok(copy.into_boxed_slice())
}

fn read_u16(data: &[u8], start: usize, label: &str) -> Result<u16> {
    let bytes = checked_slice(data, start, 2, label)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], start: usize, label: &str) -> Result<u32> {
    let bytes = checked_slice(data, start, 4, label)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

pub(crate) fn parse_iwa_components(
    archive: &ZipArchive<'_>,
    limits: Limits,
) -> Result<Vec<Component>> {
    let validated_limits = limits.validate()?;
    if is_encrypted(archive) {
        return Err(Error::Encrypted);
    }

    let has_direct_iwa = archive.file_names().any(is_iwa_name);
    let nested_name = nested_index_name(archive)?;
    if has_direct_iwa && nested_name.is_some() {
        return Err(Error::InvalidBundle(
            "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
        ));
    }
    if has_direct_iwa {
        return parse_direct_iwa_components(archive, validated_limits);
    }

    let Some(index_name) = nested_name else {
        return Ok(Vec::new());
    };
    let index_data = archive.read(&index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|_error| {
        Error::InvalidBundle("legacy iWork Index.zip length does not fit u64".to_owned())
    })?;
    validated_limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index = ZipArchive::new_with_limits(&index_data, validated_limits)?;
    let components = parse_direct_iwa_components(&index, validated_limits)?;
    if components.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }
    Ok(components)
}

fn parse_direct_iwa_components(archive: &ZipArchive<'_>, limits: Limits) -> Result<Vec<Component>> {
    let mut components = Vec::new();
    for name in archive.file_names() {
        if !is_iwa_name(name) {
            continue;
        }
        let compressed_data = archive.read(name)?;

        // OperationStorage is a separate persistence format despite its `.iwa`
        // suffix in legacy documents. It remains a raw package member but is
        // not part of the IWA object graph.
        if name.rsplit('/').next() == Some("OperationStorage.iwa")
            && compressed_data.starts_with(b"bvxn")
        {
            continue;
        }

        let decompressed =
            SnappyStream::decompress_with_limits(&compressed_data, limits.snappy_limits()?)?;
        let parsed = Archive::parse_with_limits(
            decompressed.as_bytes(),
            limits.effective_archive_limits()?,
        )?;
        components.push(Component::new(name, parsed));
    }
    components.sort_unstable_by(|left, right| left.name().cmp(right.name()));
    Ok(components)
}

pub(crate) fn is_encrypted(archive: &ZipArchive<'_>) -> bool {
    archive
        .file_names()
        .any(|name| matches!(name.rsplit('/').next(), Some(".iwpv2" | ".iwph")))
        || archive.physical_entries().any(PhysicalEntry::is_encrypted)
}

pub(crate) fn nested_index_name(archive: &ZipArchive<'_>) -> Result<Option<String>> {
    let mut candidates = archive
        .file_names()
        .filter(|name| name.rsplit('/').next() == Some("Index.zip"));
    let first = candidates.next().map(str::to_owned);
    if let Some(second) = candidates.next() {
        return Err(Error::InvalidBundle(format!(
            "iWork package contains ambiguous nested indexes: {} and {second}",
            first.as_deref().unwrap_or("Index.zip")
        )));
    }
    Ok(first)
}

pub(crate) fn is_iwa_name(name: &str) -> bool {
    #[allow(
        clippy::case_sensitive_file_extension_comparisons,
        reason = "IWA member names are case-sensitive protocol names."
    )]
    {
        name.ends_with(".iwa")
    }
}
