//! Conservative raw-member preservation for ZIP archive rewrites.
//!
//! This module intentionally accepts only ordinary single-disk ZIP archives.
//! It validates source layout/range declarations and every requested action
//! before writing to the destination, so an unsupported layout never produces
//! a partial preserved archive.

use crate::{
    CompressionMethod, EndOfCentralDirectoryRecordFixed, Error, ErrorKind, ReaderAt, ZipArchive,
    ZipArchiveWriter, ZipFileHeaderFixed, ZipLocalFileHeaderFixed,
};
use std::io::Write;
use std::ops::Range;
use std::sync::Arc;

const COPY_CHUNK_SIZE: usize = 32 * 1024;
const CENTRAL_LOCAL_HEADER_OFFSET: Range<usize> = 42..46;

/// An opaque, stable identifier for an entry in a [`PreservationIndex`].
///
/// IDs are assigned in central-directory order and cannot be constructed by
/// callers. They remain valid for the lifetime of their index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct PreservationEntryId(u32);

/// A validated source member available to a preservation plan.
#[derive(Debug, Clone)]
pub struct PreservedEntry {
    id: PreservationEntryId,
    local_span: Range<u64>,
    central_record: Range<u64>,
    central_bytes: Vec<u8>,
    compression_method: CompressionMethod,
}

impl PreservedEntry {
    /// The stable opaque ID used by [`PreservationPlan`].
    pub fn id(&self) -> PreservationEntryId {
        self.id
    }

    /// The raw local-member span: local header through the next local header
    /// (or the central directory for the final member).
    pub fn local_span(&self) -> Range<u64> {
        self.local_span.clone()
    }

    /// The exact source range of this central-directory record.
    pub fn central_record(&self) -> Range<u64> {
        self.central_record.clone()
    }

    /// The compression method declared by this entry's central-directory
    /// record.
    ///
    /// The central record is parsed and structurally validated while the
    /// [`PreservationIndex`] is built, so this accessor does not read the
    /// member payload.
    pub fn compression_method(&self) -> CompressionMethod {
        self.compression_method
    }

    /// The exact raw member-name bytes from this entry's central-directory
    /// record. These bytes are not normalized or decoded as UTF-8.
    pub fn raw_name_bytes(&self) -> &[u8] {
        let name_len =
            u16::from_le_bytes([self.central_bytes[28], self.central_bytes[29]]) as usize;
        &self.central_bytes[ZipFileHeaderFixed::SIZE..ZipFileHeaderFixed::SIZE + name_len]
    }
}

/// A newly generated ordinary ZIP member.
///
/// Regeneration uses [`ZipArchiveWriter`] for the member grammar. Only Store
/// and Deflate are supported by this low-level primitive; callers needing any
/// other generated representation should choose a different writer path.
#[derive(Debug, Clone)]
pub struct RegeneratedEntry {
    name: String,
    data: RegeneratedPayload,
    compression: CompressionMethod,
}

#[derive(Debug, Clone)]
enum RegeneratedPayload {
    Owned(Vec<u8>),
    Shared(Arc<Vec<u8>>),
}

impl RegeneratedPayload {
    fn as_slice(&self) -> &[u8] {
        match self {
            Self::Owned(data) => data,
            Self::Shared(data) => data,
        }
    }
}

impl RegeneratedEntry {
    pub fn new(name: impl Into<String>, data: impl Into<Vec<u8>>) -> Self {
        Self {
            name: name.into(),
            data: RegeneratedPayload::Owned(data.into()),
            compression: CompressionMethod::Store,
        }
    }

    /// Create a regenerated member that shares an immutable payload.
    ///
    /// This is useful when a caller already retains the complete generated
    /// payload and the ZIP writer only needs to borrow it during publication.
    pub fn new_shared(name: impl Into<String>, data: Arc<Vec<u8>>) -> Self {
        Self {
            name: name.into(),
            data: RegeneratedPayload::Shared(data),
            compression: CompressionMethod::Store,
        }
    }

    #[must_use]
    pub fn compression_method(mut self, compression: CompressionMethod) -> Self {
        self.compression = compression;
        self
    }
}

/// One action in a [`PreservationPlan`].
#[derive(Debug, Clone)]
pub enum PreservationAction {
    /// Copy the validated local span and raw central record for this entry.
    Copy(PreservationEntryId),
    /// Replace this entry using the existing ZIP writer's ordinary semantics.
    Regenerate {
        id: PreservationEntryId,
        entry: RegeneratedEntry,
    },
}

/// A complete rewrite plan for a [`PreservationIndex`].
///
/// A plan must mention every source ID exactly once. The action list itself is
/// not an ordering control: raw local spans retain source physical order and
/// central records retain source central-directory order, even when those two
/// orders differ.
#[derive(Debug, Clone, Default)]
pub struct PreservationPlan {
    actions: Vec<PreservationAction>,
}

impl PreservationPlan {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn copy_all<R>(index: &PreservationIndex<'_, R>) -> Self
    where
        R: ReaderAt,
    {
        Self {
            actions: index
                .entries
                .iter()
                .map(|entry| PreservationAction::Copy(entry.id))
                .collect(),
        }
    }

    pub fn push(&mut self, action: PreservationAction) {
        self.actions.push(action);
    }

    /// Fallibly reserve action capacity before building a bounded plan.
    pub fn try_reserve_exact(
        &mut self,
        additional: usize,
    ) -> Result<(), std::collections::TryReserveError> {
        self.actions.try_reserve_exact(additional)
    }

    pub fn actions(&self) -> &[PreservationAction] {
        &self.actions
    }
}

/// An indexed, structurally validated ordinary ZIP source.
pub struct PreservationIndex<'source, R> {
    source: &'source R,
    entries: Vec<PreservedEntry>,
    local_order: Vec<usize>,
    archive_comment: Vec<u8>,
    archive_end_offset: u64,
}

impl<'source, R> PreservationIndex<'source, R>
where
    R: ReaderAt,
{
    /// Builds a preservation index without reading member bodies.
    ///
    /// ZIP64, multi-disk, prefixed, ambiguous, overlapping, and truncated
    /// layouts are rejected with [`ErrorKind::UnsupportedPreservation`] before
    /// a caller can begin writing a plan.
    pub fn new(archive: &'source ZipArchive<R>, buffer: &mut [u8]) -> Result<Self, Error> {
        if archive.is_zip64() {
            return Err(unsupported("ZIP64 source archives"));
        }

        let central_start = archive.directory_offset();
        let eocd_offset = archive.eocd_offset();
        let archive_end = archive.end_offset();
        if central_start > eocd_offset || archive_end < eocd_offset {
            return Err(unsupported("invalid central-directory bounds"));
        }

        let mut eocd_bytes = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
        archive
            .get_ref()
            .read_exact_at(&mut eocd_bytes, eocd_offset)?;
        let eocd = EndOfCentralDirectoryRecordFixed::parse(&eocd_bytes)?;
        if eocd.disk_number != 0 || eocd.eocd_disk != 0 {
            return Err(unsupported("multi-disk archives"));
        }

        let comment_len = usize::try_from(archive_end - eocd_offset - eocd_bytes.len() as u64)
            .map_err(|_| unsupported("archive comment length"))?;
        let archive_comment = read_vec(
            archive.get_ref(),
            eocd_offset + EndOfCentralDirectoryRecordFixed::SIZE as u64,
            comment_len,
        )?;

        let entry_count = usize::try_from(archive.entries_hint())
            .map_err(|_| unsupported("central-directory entry count"))?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(entry_count)
            .map_err(|source| allocation("preservation-index entries", source))?;
        let mut iterator = archive.entries(buffer);
        while let Some(record) = iterator.next_entry()? {
            if record.is_zip64() {
                return Err(unsupported("ZIP64 entry records"));
            }

            let central_offset = record.central_directory_offset();
            let central_len = ZipFileHeaderFixed::SIZE
                .checked_add(
                    usize::try_from(record.metadata_size_hint())
                        .map_err(|_| unsupported("central-directory record length"))?,
                )
                .ok_or_else(|| unsupported("central-directory record length"))?;
            let central_end = central_offset
                .checked_add(central_len as u64)
                .ok_or_else(|| unsupported("central-directory record range"))?;
            if central_end > eocd_offset {
                return Err(unsupported("truncated central-directory record"));
            }

            let central_bytes = read_vec(archive.get_ref(), central_offset, central_len)?;
            let central_fixed = ZipFileHeaderFixed::parse(&central_bytes)?;
            if central_fixed.disk_number_start != 0 {
                return Err(unsupported("multi-disk entry records"));
            }
            if central_fixed.local_header_offset == u32::MAX
                || central_fixed.compressed_size == u32::MAX
                || central_fixed.uncompressed_size == u32::MAX
            {
                return Err(unsupported("ZIP64 entry fields"));
            }

            let id = PreservationEntryId(
                u32::try_from(entries.len()).map_err(|_| unsupported("too many entries"))?,
            );
            entries.push(PreservedEntry {
                id,
                local_span: record.local_header_offset()..record.local_header_offset(),
                central_record: central_offset..central_end,
                central_bytes,
                compression_method: central_fixed.compression_method.as_method(),
            });
        }

        if entries.len() as u64 != archive.entries_hint() {
            return Err(unsupported("central-directory entry count mismatch"));
        }

        let mut expected_central = central_start;
        for entry in &entries {
            if entry.central_record.start != expected_central {
                return Err(unsupported("non-contiguous central-directory records"));
            }
            expected_central = entry.central_record.end;
        }
        if expected_central != eocd_offset {
            return Err(unsupported("central-directory trailing bytes"));
        }

        let mut local_order = Vec::new();
        local_order
            .try_reserve_exact(entries.len())
            .map_err(|source| allocation("preservation local order", source))?;
        local_order.extend(0..entries.len());
        local_order.sort_unstable_by_key(|&index| entries[index].local_span.start);
        if let Some(&first) = local_order.first() {
            if entries[first].local_span.start != 0 {
                return Err(unsupported("archive prelude data"));
            }
        }

        for position in 0..local_order.len() {
            let index = local_order[position];
            let local_start = entries[index].local_span.start;
            let local_end = local_order
                .get(position + 1)
                .map(|next| entries[*next].local_span.start)
                .unwrap_or(central_start);
            if local_start >= local_end || local_end > central_start {
                return Err(unsupported("overlapping or empty local-member spans"));
            }

            validate_local_span(archive.get_ref(), &entries[index], local_end)?;
            entries[index].local_span.end = local_end;
        }

        Ok(Self {
            source: archive.get_ref(),
            entries,
            local_order,
            archive_comment,
            archive_end_offset: archive_end,
        })
    }

    /// The source members in original central-directory order.
    pub fn entries(&self) -> &[PreservedEntry] {
        &self.entries
    }

    /// The exact exclusive end offset of the located ZIP archive.
    ///
    /// This is the EOCD signature offset plus the fixed EOCD record and its
    /// comment. It deliberately excludes any bytes trailing the located
    /// archive in the underlying source.
    pub fn archive_end_offset(&self) -> u64 {
        self.archive_end_offset
    }

    /// Validates `plan` completely, then writes it to a non-seekable sink.
    pub fn write_to<W>(&self, plan: &PreservationPlan, mut sink: W) -> Result<W, Error>
    where
        W: Write,
    {
        let prepared = self.prepare(plan)?;
        self.validate_output_layout(&prepared)?;

        let mut local_offsets = Vec::new();
        local_offsets
            .try_reserve_exact(self.entries.len())
            .map_err(|source| allocation("preservation local offsets", source))?;
        local_offsets.resize(self.entries.len(), 0_u64);
        let mut output_offset = 0u64;
        let mut copy_buffer = [0u8; COPY_CHUNK_SIZE];
        for &index in &self.local_order {
            local_offsets[index] = output_offset;
            match &prepared[index].local {
                PreparedLocal::Copy(range) => {
                    copy_range(self.source, range.clone(), &mut sink, &mut copy_buffer)?;
                    output_offset = output_offset
                        .checked_add(range.end - range.start)
                        .ok_or_else(|| unsupported("output offset overflow"))?;
                },
                PreparedLocal::Generated(bytes) => {
                    sink.write_all(bytes)?;
                    output_offset = output_offset
                        .checked_add(bytes.len() as u64)
                        .ok_or_else(|| unsupported("output offset overflow"))?;
                },
            }
        }

        let central_start = output_offset;
        for (index, entry) in prepared.iter().enumerate() {
            let local_offset = u32::try_from(local_offsets[index])
                .map_err(|_| unsupported("ZIP64 output promotion"))?;
            let central = entry.central.bytes(&self.entries);
            sink.write_all(&central[..CENTRAL_LOCAL_HEADER_OFFSET.start])?;
            sink.write_all(&local_offset.to_le_bytes())?;
            sink.write_all(&central[CENTRAL_LOCAL_HEADER_OFFSET.end..])?;
            output_offset = output_offset
                .checked_add(central.len() as u64)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        let central_size = output_offset - central_start;

        let entry_count =
            u16::try_from(self.entries.len()).map_err(|_| unsupported("ZIP64 output promotion"))?;
        let mut eocd = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
        eocd[..4].copy_from_slice(&0x0605_4b50u32.to_le_bytes());
        eocd[8..10].copy_from_slice(&entry_count.to_le_bytes());
        eocd[10..12].copy_from_slice(&entry_count.to_le_bytes());
        eocd[12..16].copy_from_slice(&(central_size as u32).to_le_bytes());
        eocd[16..20].copy_from_slice(&(central_start as u32).to_le_bytes());
        eocd[20..22].copy_from_slice(&(self.archive_comment.len() as u16).to_le_bytes());
        sink.write_all(&eocd)?;
        sink.write_all(&self.archive_comment)?;
        sink.flush()?;
        Ok(sink)
    }

    fn prepare(&self, plan: &PreservationPlan) -> Result<Vec<PreparedEntry>, Error> {
        if plan.actions.len() != self.entries.len() {
            return Err(unsupported("plan does not cover every entry exactly once"));
        }

        let mut prepared = Vec::new();
        prepared
            .try_reserve_exact(self.entries.len())
            .map_err(|source| allocation("preservation plan", source))?;
        prepared.resize_with(self.entries.len(), || None);
        for action in &plan.actions {
            let (id, generated) = match action {
                PreservationAction::Copy(id) => (*id, None),
                PreservationAction::Regenerate { id, entry } => (*id, Some(entry)),
            };
            let index = usize::try_from(id.0).map_err(|_| unsupported("invalid entry ID"))?;
            let Some(source_entry) = self.entries.get(index) else {
                return Err(unsupported("invalid entry ID"));
            };
            if prepared[index].is_some() || source_entry.id != id {
                return Err(unsupported("duplicate entry ID in plan"));
            }

            prepared[index] = Some(match generated {
                None => PreparedEntry {
                    local: PreparedLocal::Copy(source_entry.local_span.clone()),
                    central: PreparedCentral::Copy(index),
                },
                Some(entry) => generated_entry(entry)?,
            });
        }

        let mut complete = Vec::new();
        complete
            .try_reserve_exact(self.entries.len())
            .map_err(|source| allocation("prepared preservation plan", source))?;
        for entry in prepared {
            complete.push(
                entry.ok_or_else(|| unsupported("plan does not cover every entry exactly once"))?,
            );
        }
        Ok(complete)
    }

    fn validate_output_layout(&self, prepared: &[PreparedEntry]) -> Result<(), Error> {
        let mut local_size = 0u64;
        for &index in &self.local_order {
            local_size = local_size
                .checked_add(prepared[index].local.len())
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        let central_size = prepared.iter().try_fold(0u64, |size, entry| {
            size.checked_add(entry.central.bytes(&self.entries).len() as u64)
                .ok_or_else(|| unsupported("output offset overflow"))
        })?;
        if local_size > u64::from(u32::MAX)
            || central_size > u64::from(u32::MAX)
            || local_size
                .checked_add(central_size)
                .ok_or_else(|| unsupported("output offset overflow"))?
                > u64::from(u32::MAX)
            || self.entries.len() > u16::MAX as usize
            || self.archive_comment.len() > u16::MAX as usize
        {
            return Err(unsupported("ZIP64 output promotion"));
        }
        Ok(())
    }
}

struct PreparedEntry {
    local: PreparedLocal,
    central: PreparedCentral,
}

enum PreparedCentral {
    Copy(usize),
    Generated(Vec<u8>),
}

impl PreparedCentral {
    fn bytes<'a>(&'a self, entries: &'a [PreservedEntry]) -> &'a [u8] {
        match self {
            Self::Copy(index) => &entries[*index].central_bytes,
            Self::Generated(bytes) => bytes,
        }
    }
}

enum PreparedLocal {
    Copy(Range<u64>),
    Generated(Vec<u8>),
}

impl PreparedLocal {
    fn len(&self) -> u64 {
        match self {
            Self::Copy(range) => range.end - range.start,
            Self::Generated(bytes) => bytes.len() as u64,
        }
    }
}

fn generated_entry(entry: &RegeneratedEntry) -> Result<PreparedEntry, Error> {
    let payload_len = entry.data.as_slice().len();
    let capacity = payload_len
        .checked_add(payload_len / 8)
        .and_then(|size| size.checked_add(entry.name.len()))
        .and_then(|size| size.checked_add(4 * 1024))
        .ok_or_else(|| unsupported("generated member allocation size"))?;
    let mut generated = Vec::new();
    generated
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("generated member", source))?;
    let mut writer = ZipArchiveWriter::new(generated);
    match entry.compression {
        CompressionMethod::Store => writer.write_stored_file(&entry.name, entry.data.as_slice())?,
        CompressionMethod::Deflate => {
            use flate2::Compression;
            use flate2::write::DeflateEncoder;

            let (mut file, config) = writer
                .new_file(&entry.name)
                .compression_method(CompressionMethod::Deflate)
                .start()?;
            let encoder = DeflateEncoder::new(&mut file, Compression::default());
            let mut data_writer = config.wrap(encoder);
            data_writer.write_all(entry.data.as_slice())?;
            let (encoder, descriptor) = data_writer.finish()?;
            encoder.finish()?;
            file.finish(descriptor)?;
        },
        _ => return Err(unsupported("generated compression method")),
    }
    let mut bytes = writer.finish()?;
    let (directory_offset, eocd_offset) = {
        let archive = ZipArchive::from_slice(&bytes)?;
        if archive.is_zip64() || archive.entries_hint() != 1 {
            return Err(unsupported("generated ZIP64 output"));
        }
        let directory_offset = usize::try_from(archive.directory_offset())
            .map_err(|_| unsupported("generated archive layout"))?;
        let eocd_offset = usize::try_from(archive.eocd_offset())
            .map_err(|_| unsupported("generated archive layout"))?;
        if directory_offset > eocd_offset || eocd_offset > bytes.len() {
            return Err(unsupported("generated archive layout"));
        }
        let mut entries = archive.entries();
        let record = entries
            .next_entry()?
            .ok_or_else(|| unsupported("generated archive entry"))?;
        if entries.next_entry()?.is_some() || record.is_zip64() {
            return Err(unsupported("generated archive layout"));
        }
        (directory_offset, eocd_offset)
    };
    let central_len = eocd_offset - directory_offset;
    let mut central = Vec::new();
    central
        .try_reserve_exact(central_len)
        .map_err(|source| allocation("generated central record", source))?;
    central.extend_from_slice(&bytes[directory_offset..eocd_offset]);
    bytes.truncate(directory_offset);
    Ok(PreparedEntry {
        local: PreparedLocal::Generated(bytes),
        central: PreparedCentral::Generated(central),
    })
}

fn validate_local_span<R: ReaderAt>(
    source: &R,
    entry: &PreservedEntry,
    local_end: u64,
) -> Result<(), Error> {
    let mut local_bytes = [0u8; ZipLocalFileHeaderFixed::SIZE];
    source.read_exact_at(&mut local_bytes, entry.local_span.start)?;
    let local = ZipLocalFileHeaderFixed::parse(&local_bytes)?;
    let central = ZipFileHeaderFixed::parse(&entry.central_bytes)?;
    if local.flags != central.flags || local.compression_method != central.compression_method {
        return Err(unsupported("local and central header mismatch"));
    }
    let header_end = entry
        .local_span
        .start
        .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
        .and_then(|offset| offset.checked_add(local.variable_length() as u64))
        .ok_or_else(|| unsupported("local header range"))?;
    let payload_end = header_end
        .checked_add(u64::from(central.compressed_size))
        .ok_or_else(|| unsupported("local payload range"))?;
    if payload_end > local_end {
        return Err(unsupported("truncated or overlapping local member"));
    }
    if local.flags & 0x08 == 0
        && (local.crc32 != central.crc32
            || local.compressed_size != central.compressed_size
            || local.uncompressed_size != central.uncompressed_size)
    {
        return Err(unsupported("local and central sizes mismatch"));
    }
    Ok(())
}

fn copy_range<R, W>(
    source: &R,
    range: Range<u64>,
    sink: &mut W,
    buffer: &mut [u8],
) -> Result<(), Error>
where
    R: ReaderAt,
    W: Write,
{
    let mut offset = range.start;
    while offset < range.end {
        let len = usize::try_from((range.end - offset).min(buffer.len() as u64))
            .map_err(|_| unsupported("copy range length"))?;
        source.read_exact_at(&mut buffer[..len], offset)?;
        sink.write_all(&buffer[..len])?;
        offset += len as u64;
    }
    Ok(())
}

fn read_vec<R: ReaderAt>(source: &R, offset: u64, len: usize) -> Result<Vec<u8>, Error> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(len)
        .map_err(|source| allocation("preservation byte buffer", source))?;
    bytes.resize(len, 0);
    source.read_exact_at(&mut bytes, offset)?;
    Ok(bytes)
}

fn unsupported(reason: &'static str) -> Error {
    Error::from(ErrorKind::UnsupportedPreservation { reason })
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::from(ErrorKind::Allocation { resource, source })
}

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::{Compression, write::DeflateEncoder};
    use std::{
        cell::RefCell,
        io::{self, Cursor},
    };

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

        fn clear_reads(&self) {
            self.reads.borrow_mut().clear();
        }

        fn reads(&self) -> Vec<(u64, usize)> {
            self.reads.borrow().clone()
        }
    }

    impl ReaderAt for CountingReader {
        fn read_at(&self, buffer: &mut [u8], offset: u64) -> io::Result<usize> {
            self.reads.borrow_mut().push((offset, buffer.len()));
            let bytes = self.data.get(offset as usize..).unwrap_or_default();
            let len = bytes.len().min(buffer.len());
            buffer[..len].copy_from_slice(&bytes[..len]);
            Ok(len)
        }
    }

    fn deflated_file(writer: &mut ZipArchiveWriter<Vec<u8>>, name: &str, data: &[u8]) {
        let (mut file, config) = writer
            .new_file(name)
            .compression_method(CompressionMethod::Deflate)
            .start()
            .unwrap();
        let encoder = DeflateEncoder::new(&mut file, Compression::default());
        let mut data_writer = config.wrap(encoder);
        data_writer.write_all(data).unwrap();
        let (encoder, descriptor) = data_writer.finish().unwrap();
        encoder.finish().unwrap();
        file.finish(descriptor).unwrap();
    }

    fn ordinary_archive() -> Vec<u8> {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        writer
            .write_stored_file("first.bin", b"stored data")
            .unwrap();
        deflated_file(&mut writer, "second.bin", b"deflated data");
        writer.new_dir("folder/").create().unwrap();
        writer.finish().unwrap()
    }

    fn without_descriptor_signature(mut data: Vec<u8>) -> Vec<u8> {
        let archive = ZipArchive::from_slice(&data).unwrap();
        let central = archive.directory_offset() as usize;
        let descriptor = data[..central]
            .windows(4)
            .position(|window| window == crate::DataDescriptor::SIGNATURE.to_le_bytes())
            .unwrap();
        data.drain(descriptor..descriptor + 4);

        let new_central = central - 4;
        let eocd = data.len() - EndOfCentralDirectoryRecordFixed::SIZE;
        data[eocd + 16..eocd + 20].copy_from_slice(&(new_central as u32).to_le_bytes());
        let mut offset = new_central;
        while offset < eocd {
            let header = ZipFileHeaderFixed::parse(&data[offset..]).unwrap();
            let length = ZipFileHeaderFixed::SIZE + header.variable_length();
            let local_offset =
                u32::from_le_bytes(data[offset + 42..offset + 46].try_into().unwrap());
            if local_offset as usize > descriptor {
                data[offset + 42..offset + 46].copy_from_slice(&(local_offset - 4).to_le_bytes());
            }
            offset += length;
        }
        data
    }

    fn zip64_archive() -> Vec<u8> {
        let mut data = vec![0];
        data.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        data.extend_from_slice(&44u64.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&[0; 32]);
        data.extend_from_slice(&0x0706_4b50u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u64.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data
    }

    fn indexed(data: &[u8]) -> (ZipArchive<Cursor<&[u8]>>, Vec<u8>) {
        let archive = ZipArchive::from_slice(data).unwrap().into_zip_archive();
        (archive, vec![0; crate::RECOMMENDED_BUFFER_SIZE])
    }

    #[test]
    fn copies_store_deflate_descriptor_and_directory_members_byte_for_byte() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();

        assert_eq!(output, data);
    }

    #[test]
    fn exposes_compression_methods_and_excludes_trailing_source_suffix() {
        let archive_bytes = ordinary_archive();
        let mut source_bytes = archive_bytes.clone();
        source_bytes.extend_from_slice(b"source suffix");

        let (archive, mut buffer) = indexed(&source_bytes);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        assert_eq!(index.archive_end_offset(), archive_bytes.len() as u64);
        assert_eq!(
            index
                .entries()
                .iter()
                .map(PreservedEntry::compression_method)
                .collect::<Vec<_>>(),
            vec![
                CompressionMethod::Store,
                CompressionMethod::Deflate,
                CompressionMethod::Store,
            ]
        );

        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();
        assert_eq!(output, archive_bytes);
    }

    #[test]
    fn copy_all_reads_each_local_span_once_during_write() {
        let data = ordinary_archive();
        let end_offset = data.len() as u64;
        let archive = crate::ZipLocator::new()
            .locate_in_reader(CountingReader::new(data.clone()), &mut [0; 64], end_offset)
            .unwrap();
        let mut buffer = vec![0; crate::RECOMMENDED_BUFFER_SIZE];
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let expected: Vec<_> = index
            .entries
            .iter()
            .map(|entry| entry.local_span())
            .collect();
        archive.get_ref().clear_reads();

        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();
        let reads = archive.get_ref().reads();

        assert_eq!(output, data);
        assert_eq!(reads.len(), expected.len());
        for (read, span) in reads.iter().zip(expected) {
            assert_eq!(*read, (span.start, (span.end - span.start) as usize));
        }
    }

    #[test]
    fn preserves_unsigned_data_descriptors() {
        let data = without_descriptor_signature(ordinary_archive());
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();

        assert_eq!(output, data);
    }

    #[test]
    fn preserves_archive_comments_and_different_local_central_metadata() {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        let (mut file, config) = writer
            .new_file("central-a")
            .extra_field(
                crate::extra_fields::ExtraFieldId::new(0xaaaa),
                b"local",
                crate::Header::LOCAL,
            )
            .unwrap()
            .extra_field(
                crate::extra_fields::ExtraFieldId::new(0xbbbb),
                b"central",
                crate::Header::CENTRAL,
            )
            .unwrap()
            .start()
            .unwrap();
        let mut data_writer = config.wrap(&mut file);
        data_writer.write_all(b"payload").unwrap();
        let (_, descriptor) = data_writer.finish().unwrap();
        file.finish(descriptor).unwrap();
        let mut data = writer.finish().unwrap();

        let archive = ZipArchive::from_slice(&data).unwrap();
        let eocd = archive.eocd_offset() as usize;
        data[eocd + 20..eocd + 22].copy_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(b"ok");
        data[30..39].copy_from_slice(b"local-nam");

        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();

        assert_eq!(output, data);
    }

    #[test]
    fn preserves_physical_local_order_when_central_order_differs() {
        let mut data = ordinary_archive();
        let archive = ZipArchive::from_slice(&data).unwrap();
        let central = archive.directory_offset() as usize;
        let eocd = archive.eocd_offset() as usize;
        let first_len = ZipFileHeaderFixed::SIZE
            + ZipFileHeaderFixed::parse(&data[central..])
                .unwrap()
                .variable_length();
        let first = data[central..central + first_len].to_vec();
        let rest = data[central + first_len..eocd].to_vec();
        data[central..eocd].copy_from_slice(&[rest, first].concat());

        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();

        assert_eq!(output, data);
    }

    #[test]
    fn regenerates_a_selected_member_with_writer_semantics() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let mut plan = PreservationPlan::copy_all(&index);
        plan.actions[1] = PreservationAction::Regenerate {
            id: index.entries()[1].id(),
            entry: RegeneratedEntry::new("replacement.bin", b"new content".as_slice())
                .compression_method(CompressionMethod::Deflate),
        };

        let output = index.write_to(&plan, Vec::new()).unwrap();
        let output_archive = ZipArchive::from_slice(&output).unwrap();
        let names: Vec<_> = output_archive
            .entries()
            .map(|entry| entry.unwrap().file_path().as_ref().to_vec())
            .collect();
        assert_eq!(
            names,
            vec![
                b"first.bin".to_vec(),
                b"replacement.bin".to_vec(),
                b"folder/".to_vec(),
            ]
        );
    }

    #[test]
    fn regenerated_entry_can_retain_a_shared_payload() {
        let data = Arc::new(b"shared generated content".to_vec());
        let entry = RegeneratedEntry::new_shared("shared.bin", Arc::clone(&data));

        assert_eq!(Arc::strong_count(&data), 2);

        let prepared = generated_entry(&entry).unwrap();
        assert!(matches!(prepared.local, PreparedLocal::Generated(_)));
        assert_eq!(Arc::strong_count(&data), 2);
    }

    #[test]
    fn generated_entry_splits_large_local_and_central_spans() {
        let data = Arc::new(
            (0..256 * 1024)
                .map(|index| u8::try_from(index % 251).unwrap())
                .collect::<Vec<_>>(),
        );
        let entry = RegeneratedEntry::new_shared("large.bin", Arc::clone(&data))
            .compression_method(CompressionMethod::Deflate);

        let prepared = generated_entry(&entry).unwrap();
        let PreparedLocal::Generated(local) = prepared.local else {
            panic!("generated entry must retain generated local bytes");
        };
        let local_header = ZipLocalFileHeaderFixed::parse(&local).unwrap();
        let central = prepared.central.bytes(&[]);
        let central_header = ZipFileHeaderFixed::parse(central).unwrap();

        assert_eq!(local_header.file_name_len as usize, b"large.bin".len());
        assert_eq!(
            &local
                [ZipLocalFileHeaderFixed::SIZE..ZipLocalFileHeaderFixed::SIZE + b"large.bin".len()],
            b"large.bin"
        );
        assert_eq!(
            central.len(),
            ZipFileHeaderFixed::SIZE + central_header.variable_length()
        );
        assert_eq!(central_header.local_header_offset, 0);
        assert_eq!(central_header.uncompressed_size as usize, data.len());
        assert_eq!(Arc::strong_count(&data), 2);
    }

    #[test]
    fn rejects_overlapping_and_truncated_source_spans_before_writing() {
        let mut overlapping = ordinary_archive();
        let archive = ZipArchive::from_slice(&overlapping).unwrap();
        let central = archive.directory_offset() as usize;
        let first_len = ZipFileHeaderFixed::SIZE
            + ZipFileHeaderFixed::parse(&overlapping[central..])
                .unwrap()
                .variable_length();
        overlapping[central + first_len + 42..central + first_len + 46]
            .copy_from_slice(&0u32.to_le_bytes());
        let (archive, mut buffer) = indexed(&overlapping);
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("overlapping spans must be rejected"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));

        let mut truncated = ordinary_archive();
        let archive = ZipArchive::from_slice(&truncated).unwrap();
        let central = archive.directory_offset() as usize;
        truncated.remove(30 + b"first.bin".len() + b"stored data".len() - 1);
        // Reconstruct an EOCD whose central-directory offset follows the removed byte.
        let old_eocd = truncated.len() - 22;
        truncated[old_eocd + 16..old_eocd + 20]
            .copy_from_slice(&((central - 1) as u32).to_le_bytes());
        let first_central_len = ZipFileHeaderFixed::SIZE
            + ZipFileHeaderFixed::parse(&truncated[central - 1..])
                .unwrap()
                .variable_length();
        truncated[central - 1 + first_central_len + 42..central - 1 + first_central_len + 46]
            .copy_from_slice(&(49u32).to_le_bytes());
        let (archive, mut buffer) = indexed(&truncated);
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("truncated spans must be rejected"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
    }

    #[derive(Debug, Default)]
    struct FailingSink;

    impl Write for FailingSink {
        fn write(&mut self, _buf: &[u8]) -> io::Result<usize> {
            Err(io::Error::new(io::ErrorKind::BrokenPipe, "sink failed"))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn propagates_sink_failure_after_preflight() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let error = index
            .write_to(&PreservationPlan::copy_all(&index), FailingSink)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
    }

    #[test]
    fn rejects_zip64_sources_with_a_typed_fallback() {
        let data = zip64_archive();
        let archive = ZipArchive::from_slice(&data).unwrap().into_zip_archive();
        let mut buffer = vec![0; crate::RECOMMENDED_BUFFER_SIZE];
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("ZIP64 sources must be rejected"),
            Err(error) => error,
        };

        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
    }
}
