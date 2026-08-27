//! Conservative raw-member preservation for ZIP archive rewrites.
//!
//! This module intentionally accepts only ordinary single-disk ZIP archives.
//! It validates source layout/range declarations and every requested action
//! before writing to the destination, so an unsupported layout never produces
//! a partial preserved archive.

use crate::office::ArchiveLimits;
use crate::{
    CompressionMethod, EndOfCentralDirectoryRecordFixed, Error, ErrorKind, LimitResource, ReaderAt,
    ZipArchive, ZipArchiveWriter, ZipFileHeaderFixed, ZipLocalFileHeaderFixed,
    accounting::{AccountingWriteKind, ZipOperationAccounting, usize_to_u64, write_all_counted},
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
    local_central_name_mismatch: bool,
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
    /// Omit this source entry from the rewritten archive.
    Omit(PreservationEntryId),
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
/// orders differ. Appended local records are emitted after all source local
/// records and before the central directory; appended central records are
/// emitted after all source central records.
#[derive(Debug, Clone, Default)]
pub struct PreservationPlan {
    actions: Vec<PreservationAction>,
    appended: Vec<RegeneratedEntry>,
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
            appended: Vec::new(),
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

    /// Fallibly reserve capacity for appended generated members.
    pub fn try_reserve_appended(
        &mut self,
        additional: usize,
    ) -> Result<(), std::collections::TryReserveError> {
        self.appended.try_reserve_exact(additional)
    }

    /// Append one generated member in the order it should appear after all
    /// source local records and before the central directory, with its central
    /// record after all source central records.
    ///
    /// The source action list remains separate and must still mention every
    /// source ID exactly once. Appended members are generated before the sink
    /// sees any bytes, so a member-generation failure cannot leave partial
    /// output behind.
    pub fn try_append(
        &mut self,
        entry: RegeneratedEntry,
    ) -> Result<(), std::collections::TryReserveError> {
        self.appended.try_reserve(1)?;
        self.appended.push(entry);
        Ok(())
    }

    pub fn actions(&self) -> &[PreservationAction] {
        &self.actions
    }

    /// Generated members in their requested append order.
    pub fn appended(&self) -> &[RegeneratedEntry] {
        &self.appended
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
        Self::new_with_limits(archive, buffer, ArchiveLimits::default())
    }

    /// Builds a preservation index under explicit archive metadata limits.
    ///
    /// Preservation retains every source central-directory record verbatim,
    /// including records for directory members and the EOCD comment. The
    /// archive profile's `max_files` bounds non-directory source members;
    /// directory records are retained too, but consume only the metadata
    /// budget. `max_metadata_bytes` covers the aggregate variable central
    /// metadata plus the EOCD comment. Limits are validated in a metadata-only
    /// pass before any owned entry or raw-record buffers are reserved.
    pub fn new_with_limits(
        archive: &'source ZipArchive<R>,
        buffer: &mut [u8],
        limits: ArchiveLimits,
    ) -> Result<Self, Error> {
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

        let comment_len_u64 = archive_end
            .checked_sub(eocd_offset)
            .and_then(|length| length.checked_sub(eocd_bytes.len() as u64))
            .ok_or_else(|| unsupported("archive comment length"))?;
        let comment_len =
            usize::try_from(comment_len_u64).map_err(|_| unsupported("archive comment length"))?;
        if comment_len_u64 > limits.max_metadata_bytes {
            return Err(limit_error(
                LimitResource::MetadataBytes,
                comment_len_u64,
                limits.max_metadata_bytes,
            ));
        }

        let entry_count_hint = archive.entries_hint();
        let max_files_u64 = u64::try_from(limits.max_files).unwrap_or(u64::MAX);
        let central_metadata_budget = limits
            .max_metadata_bytes
            .checked_sub(comment_len_u64)
            .ok_or_else(|| unsupported("central-directory metadata budget"))?;

        // Validate every central record's bounded metadata and structural
        // declaration before reserving the owned preservation index. This
        // keeps a hostile count or oversized metadata from causing ownership
        // allocation before the corresponding limit is reported.
        let mut iterator = archive.entries_with_metadata_limit(buffer, central_metadata_budget);
        let mut entry_count_u64 = 0u64;
        let mut file_count_u64 = 0u64;
        while let Some(record) = iterator
            .next_entry()
            .map_err(|error| map_metadata_limit_error(error, comment_len_u64, limits))?
        {
            if record.is_zip64() {
                return Err(unsupported("ZIP64 entry records"));
            }

            let raw_name_bytes = u64::try_from(record.file_path().as_ref().len())
                .map_err(|_| unsupported("member name length"))?;
            if raw_name_bytes > limits.max_member_name_bytes {
                return Err(limit_error(
                    LimitResource::MemberNameBytes,
                    raw_name_bytes,
                    limits.max_member_name_bytes,
                ));
            }

            if !record.is_dir() {
                file_count_u64 = file_count_u64
                    .checked_add(1)
                    .ok_or_else(|| unsupported("central-directory file count"))?;
                if file_count_u64 > max_files_u64 {
                    return Err(limit_error(
                        LimitResource::FileCount,
                        file_count_u64,
                        max_files_u64,
                    ));
                }
            }

            central_record_span(&record, eocd_offset)?;
            entry_count_u64 = entry_count_u64
                .checked_add(1)
                .ok_or_else(|| unsupported("central-directory entry count"))?;
        }

        if entry_count_u64 != entry_count_hint {
            return Err(unsupported("central-directory entry count mismatch"));
        }

        // Fixed central headers are retained alongside variable metadata. The
        // saturated arithmetic intentionally preserves UNBOUNDED semantics,
        // while the checked multiplication still rejects impossible ownership
        // sizes before any reservation.
        let fixed_central_bytes = entry_count_u64
            .checked_mul(ZipFileHeaderFixed::SIZE as u64)
            .ok_or_else(|| unsupported("central-directory source budget"))?;
        let source_budget = limits
            .max_metadata_bytes
            .saturating_add(fixed_central_bytes);

        let comment_offset = eocd_offset
            .checked_add(EndOfCentralDirectoryRecordFixed::SIZE as u64)
            .ok_or_else(|| unsupported("archive comment offset"))?;
        let archive_comment = read_vec(archive.get_ref(), comment_offset, comment_len)?;

        let entry_count = usize::try_from(entry_count_u64)
            .map_err(|_| unsupported("central-directory entry count"))?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(entry_count)
            .map_err(|source| allocation("preservation-index entries", source))?;

        let mut iterator = archive.entries_with_metadata_limit(buffer, central_metadata_budget);
        let mut retained_source_bytes = comment_len_u64;
        while let Some(record) = iterator
            .next_entry()
            .map_err(|error| map_metadata_limit_error(error, comment_len_u64, limits))?
        {
            if record.is_zip64() {
                return Err(unsupported("ZIP64 entry records"));
            }

            let (central_offset, central_len, central_end) =
                central_record_span(&record, eocd_offset)?;
            let next_retained_source_bytes = retained_source_bytes
                .checked_add(central_len as u64)
                .ok_or_else(|| unsupported("central-directory source budget"))?;
            if next_retained_source_bytes > source_budget {
                return Err(limit_error(
                    LimitResource::MetadataBytes,
                    next_retained_source_bytes,
                    source_budget,
                ));
            }

            let central_bytes = read_vec(archive.get_ref(), central_offset, central_len)?;
            retained_source_bytes = next_retained_source_bytes;
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
                local_central_name_mismatch: false,
            });
        }

        if entries.len() as u64 != entry_count_u64 {
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

            let local_central_name_mismatch =
                validate_local_span(archive.get_ref(), &entries[index], local_end)?;
            entries[index].local_central_name_mismatch = local_central_name_mismatch;
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
    pub fn write_to<W>(&self, plan: &PreservationPlan, sink: W) -> Result<W, Error>
    where
        W: Write,
    {
        let mut accounting = ZipOperationAccounting::default();
        self.write_to_with_accounting(plan, sink, &mut accounting)
    }

    /// Validate and write a preservation plan while recording unchanged source
    /// bytes accepted by the publication sink. The counter intentionally
    /// includes unchanged local-member spans, unchanged central-record bytes
    /// outside regenerated local offsets, and the source archive comment; it
    /// is not payload-only. Generated member framing and payload are excluded
    /// from this raw-source counter.
    pub fn write_to_with_accounting<W>(
        &self,
        plan: &PreservationPlan,
        mut sink: W,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<W, Error>
    where
        W: Write,
    {
        let prepared = self.prepare(plan)?;
        self.validate_output_layout(&prepared)?;

        let mut local_offsets = Vec::new();
        local_offsets
            .try_reserve_exact(prepared.len())
            .map_err(|source| allocation("preservation local offsets", source))?;
        local_offsets.resize(prepared.len(), 0_u64);
        let mut output_offset = 0u64;
        let mut copy_buffer = [0u8; COPY_CHUNK_SIZE];
        for &index in &self.local_order {
            if prepared[index].omitted {
                continue;
            }
            local_offsets[index] = output_offset;
            let local_size = write_prepared_local(
                &prepared[index].local,
                prepared[index].generated_payload.as_ref(),
                self.source,
                &mut sink,
                &mut copy_buffer,
                accounting,
            )?;
            output_offset = output_offset
                .checked_add(local_size)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        for index in self.entries.len()..prepared.len() {
            if prepared[index].omitted {
                continue;
            }
            local_offsets[index] = output_offset;
            let local_size = write_prepared_local(
                &prepared[index].local,
                prepared[index].generated_payload.as_ref(),
                self.source,
                &mut sink,
                &mut copy_buffer,
                accounting,
            )?;
            output_offset = output_offset
                .checked_add(local_size)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }

        let central_start = output_offset;
        for (index, entry) in prepared.iter().enumerate() {
            if entry.omitted {
                continue;
            }
            let local_offset = u32::try_from(local_offsets[index])
                .map_err(|_| unsupported("ZIP64 output promotion"))?;
            let central = entry.central.bytes(&self.entries);
            let is_unchanged = matches!(&entry.central, PreparedCentral::Copy(_));
            if is_unchanged {
                write_all_counted(
                    &mut sink,
                    &central[..CENTRAL_LOCAL_HEADER_OFFSET.start],
                    accounting,
                    AccountingWriteKind::RawUnchangedSource,
                )?;
            } else {
                sink.write_all(&central[..CENTRAL_LOCAL_HEADER_OFFSET.start])?;
            }
            sink.write_all(&local_offset.to_le_bytes())?;
            if is_unchanged {
                write_all_counted(
                    &mut sink,
                    &central[CENTRAL_LOCAL_HEADER_OFFSET.end..],
                    accounting,
                    AccountingWriteKind::RawUnchangedSource,
                )?;
            } else {
                sink.write_all(&central[CENTRAL_LOCAL_HEADER_OFFSET.end..])?;
            }
            output_offset = output_offset
                .checked_add(central.len() as u64)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        let central_size = output_offset - central_start;

        let entry_count = u16::try_from(retained_entry_count(&prepared))
            .map_err(|_| unsupported("ZIP64 output promotion"))?;
        let mut eocd = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
        eocd[..4].copy_from_slice(&0x0605_4b50u32.to_le_bytes());
        eocd[8..10].copy_from_slice(&entry_count.to_le_bytes());
        eocd[10..12].copy_from_slice(&entry_count.to_le_bytes());
        eocd[12..16].copy_from_slice(
            &u32::try_from(central_size)
                .map_err(|_| unsupported("ZIP64 output promotion"))?
                .to_le_bytes(),
        );
        eocd[16..20].copy_from_slice(
            &u32::try_from(central_start)
                .map_err(|_| unsupported("ZIP64 output promotion"))?
                .to_le_bytes(),
        );
        eocd[20..22].copy_from_slice(
            &u16::try_from(self.archive_comment.len())
                .map_err(|_| unsupported("ZIP64 output promotion"))?
                .to_le_bytes(),
        );
        sink.write_all(&eocd)?;
        write_all_counted(
            &mut sink,
            &self.archive_comment,
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        )?;
        sink.flush()?;
        Ok(sink)
    }

    fn prepare(&self, plan: &PreservationPlan) -> Result<Vec<PreparedEntry>, Error> {
        if plan.actions.len() != self.entries.len() {
            return Err(unsupported("plan does not cover every entry exactly once"));
        }

        let mut preserved_name_mismatch = false;
        let mut generated_new_member = !plan.appended.is_empty();
        for action in &plan.actions {
            let (id, is_copy, generated) = match action {
                PreservationAction::Copy(id) => (*id, true, None),
                PreservationAction::Omit(id) => (*id, false, None),
                PreservationAction::Regenerate { id, entry } => (*id, false, Some(entry)),
            };
            let index = usize::try_from(id.0).map_err(|_| unsupported("invalid entry ID"))?;
            let Some(source_entry) = self.entries.get(index) else {
                return Err(unsupported("invalid entry ID"));
            };
            if is_copy && source_entry.local_central_name_mismatch {
                preserved_name_mismatch = true;
            }
            if let Some(entry) = generated {
                if entry.name.as_bytes() != source_entry.raw_name_bytes() {
                    generated_new_member = true;
                }
            }
        }
        if preserved_name_mismatch && generated_new_member {
            return Err(unsupported(
                "new members are unsupported with local and central filename mismatch",
            ));
        }

        let mut prepared = Vec::new();
        prepared
            .try_reserve_exact(self.entries.len())
            .map_err(|source| allocation("preservation plan", source))?;
        prepared.resize_with(self.entries.len(), || None);
        for action in &plan.actions {
            let (id, generated, omitted) = match action {
                PreservationAction::Copy(id) => (*id, None, false),
                PreservationAction::Omit(id) => (*id, None, true),
                PreservationAction::Regenerate { id, entry } => (*id, Some(entry), false),
            };
            let index = usize::try_from(id.0).map_err(|_| unsupported("invalid entry ID"))?;
            let Some(source_entry) = self.entries.get(index) else {
                return Err(unsupported("invalid entry ID"));
            };
            if prepared[index].is_some() || source_entry.id != id {
                return Err(unsupported("duplicate entry ID in plan"));
            }

            prepared[index] = Some(if omitted {
                PreparedEntry {
                    local: PreparedLocal::Generated(Vec::new()),
                    central: PreparedCentral::Generated(Vec::new()),
                    generated_payload: None,
                    omitted: true,
                }
            } else {
                match generated {
                    None => PreparedEntry {
                        local: PreparedLocal::Copy(source_entry.local_span.clone()),
                        central: PreparedCentral::Copy(index),
                        generated_payload: None,
                        omitted: false,
                    },
                    Some(entry) => generated_entry(entry)?,
                }
            });
        }

        let complete_len = self
            .entries
            .len()
            .checked_add(plan.appended.len())
            .ok_or_else(|| unsupported("prepared preservation plan length"))?;
        let mut complete = Vec::new();
        complete
            .try_reserve_exact(complete_len)
            .map_err(|source| allocation("prepared preservation plan", source))?;
        for entry in prepared {
            complete.push(
                entry.ok_or_else(|| unsupported("plan does not cover every entry exactly once"))?,
            );
        }
        for entry in &plan.appended {
            complete.push(generated_entry(entry)?);
        }
        Ok(complete)
    }

    fn validate_output_layout(&self, prepared: &[PreparedEntry]) -> Result<(), Error> {
        let mut local_size = 0u64;
        for &index in &self.local_order {
            if prepared[index].omitted {
                continue;
            }
            local_size = local_size
                .checked_add(prepared[index].local.len())
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        for entry in &prepared[self.entries.len()..] {
            if entry.omitted {
                continue;
            }
            local_size = local_size
                .checked_add(entry.local.len())
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        let central_size = prepared.iter().try_fold(0u64, |size, entry| {
            if entry.omitted {
                return Ok(size);
            }
            size.checked_add(entry.central.bytes(&self.entries).len() as u64)
                .ok_or_else(|| unsupported("output offset overflow"))
        })?;
        let output_size = local_size
            .checked_add(central_size)
            .and_then(|size| size.checked_add(EndOfCentralDirectoryRecordFixed::SIZE as u64))
            .and_then(|size| size.checked_add(self.archive_comment.len() as u64))
            .ok_or_else(|| unsupported("output offset overflow"))?;
        if local_size > u64::from(u32::MAX)
            || central_size > u64::from(u32::MAX)
            || output_size > u64::from(u32::MAX)
            || retained_entry_count(prepared) >= u16::MAX as usize
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
    generated_payload: Option<GeneratedPayload>,
    omitted: bool,
}

struct GeneratedPayload {
    range: Range<usize>,
    kind: AccountingWriteKind,
}

fn retained_entry_count(prepared: &[PreparedEntry]) -> usize {
    prepared.iter().filter(|entry| !entry.omitted).count()
}

enum PreparedCentral {
    Copy(usize),
    Generated(Vec<u8>),
    /// Generated local and central records may share the one archive buffer.
    ///
    /// Keeping ranges into the finished one-entry archive avoids copying the
    /// central record into a second allocation merely to retain it while the
    /// local record is emitted to a forward-only sink.
    Shared {
        bytes: Arc<Vec<u8>>,
        range: Range<usize>,
    },
}

impl PreparedCentral {
    fn bytes<'a>(&'a self, entries: &'a [PreservedEntry]) -> &'a [u8] {
        match self {
            Self::Copy(index) => &entries[*index].central_bytes,
            Self::Generated(bytes) => bytes,
            Self::Shared { bytes, range } => &bytes[range.clone()],
        }
    }
}

enum PreparedLocal {
    Copy(Range<u64>),
    Generated(Vec<u8>),
    /// A range into the same finished archive buffer retained by the central
    /// record. This keeps generated-member publication forward-only without
    /// copying the central directory bytes.
    Shared {
        bytes: Arc<Vec<u8>>,
        range: Range<usize>,
    },
}

impl PreparedLocal {
    fn len(&self) -> u64 {
        match self {
            Self::Copy(range) => range.end - range.start,
            Self::Generated(bytes) => bytes.len() as u64,
            Self::Shared { range, .. } => (range.end - range.start) as u64,
        }
    }
}

fn generated_entry(entry: &RegeneratedEntry) -> Result<PreparedEntry, Error> {
    let payload_len = entry.data.as_slice().len();
    let name_bytes = entry
        .name
        .len()
        .checked_mul(2)
        .ok_or_else(|| unsupported("generated member allocation size"))?;
    let capacity = payload_len
        .checked_add(payload_len / 8)
        .and_then(|size| size.checked_add(name_bytes))
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
    let bytes = writer.finish()?;
    let (directory_offset, eocd_offset, payload_start, payload_end) = {
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
        let entry = archive.get_entry(record.wayfinder())?;
        let (payload_start, payload_end) = entry.compressed_data_range();
        let payload_start =
            usize::try_from(payload_start).map_err(|_| unsupported("generated payload range"))?;
        let payload_end =
            usize::try_from(payload_end).map_err(|_| unsupported("generated payload range"))?;
        if payload_start > payload_end || payload_end > directory_offset {
            return Err(unsupported("generated payload range"));
        }
        (directory_offset, eocd_offset, payload_start, payload_end)
    };
    // Retain one owned archive buffer and publish disjoint ranges from it.
    // The previous implementation copied the central record into another
    // `Vec`, even though the finished one-entry archive already contained the
    // exact bytes needed for both forward-only output phases.
    let bytes = Arc::new(bytes);
    Ok(PreparedEntry {
        local: PreparedLocal::Shared {
            bytes: Arc::clone(&bytes),
            range: 0..directory_offset,
        },
        central: PreparedCentral::Shared {
            bytes,
            range: directory_offset..eocd_offset,
        },
        generated_payload: Some(GeneratedPayload {
            range: payload_start..payload_end,
            kind: match entry.compression {
                CompressionMethod::Store => AccountingWriteKind::Stored,
                CompressionMethod::Deflate => AccountingWriteKind::GeneratedDeflate,
                _ => return Err(unsupported("generated compression method")),
            },
        }),
        omitted: false,
    })
}

fn validate_local_span<R: ReaderAt>(
    source: &R,
    entry: &PreservedEntry,
    local_end: u64,
) -> Result<bool, Error> {
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
    let central_name_start = ZipFileHeaderFixed::SIZE;
    let central_name_end = central_name_start
        .checked_add(usize::from(central.file_name_len))
        .ok_or_else(|| unsupported("central filename range"))?;
    if central_name_end > entry.central_bytes.len() {
        return Err(unsupported("central filename range"));
    }
    let local_central_name_mismatch = if local.file_name_len != central.file_name_len {
        true
    } else {
        let local_name_offset = entry
            .local_span
            .start
            .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
            .ok_or_else(|| unsupported("local filename range"))?;
        let local_name = read_vec(source, local_name_offset, usize::from(local.file_name_len))?;
        local_name.as_slice() != &entry.central_bytes[central_name_start..central_name_end]
    };
    if local.flags & 0x08 == 0
        && (local.crc32 != central.crc32
            || local.compressed_size != central.compressed_size
            || local.uncompressed_size != central.uncompressed_size)
    {
        return Err(unsupported("local and central sizes mismatch"));
    }
    Ok(local_central_name_mismatch)
}

fn write_prepared_local<R, W>(
    local: &PreparedLocal,
    generated_payload: Option<&GeneratedPayload>,
    source: &R,
    sink: &mut W,
    buffer: &mut [u8],
    accounting: &mut ZipOperationAccounting,
) -> Result<u64, Error>
where
    R: ReaderAt,
    W: Write,
{
    match (local, generated_payload) {
        (PreparedLocal::Copy(range), None) => {
            copy_range(source, range.clone(), sink, buffer, accounting)?;
            Ok(range.end - range.start)
        },
        (PreparedLocal::Copy(_), Some(_)) => Err(unsupported("generated payload on copied member")),
        (PreparedLocal::Generated(bytes), None) => {
            sink.write_all(bytes)?;
            usize_to_u64(bytes.len(), "generated local bytes")
        },
        (PreparedLocal::Generated(_), Some(_)) => {
            Err(unsupported("generated payload metadata on owned local"))
        },
        (PreparedLocal::Shared { bytes, range }, Some(payload)) => {
            if payload.range.start < range.start
                || payload.range.end > range.end
                || payload.range.start > payload.range.end
            {
                return Err(unsupported("generated payload outside local member"));
            }
            let local_bytes = &bytes[range.clone()];
            let payload_start = payload.range.start - range.start;
            let payload_end = payload.range.end - range.start;
            sink.write_all(&local_bytes[..payload_start])?;
            write_all_counted(
                sink,
                &local_bytes[payload_start..payload_end],
                accounting,
                payload.kind,
            )?;
            sink.write_all(&local_bytes[payload_end..])?;
            usize_to_u64(local_bytes.len(), "generated local bytes")
        },
        (PreparedLocal::Shared { bytes, range }, None) => {
            let local_bytes = &bytes[range.clone()];
            sink.write_all(local_bytes)?;
            usize_to_u64(local_bytes.len(), "generated local bytes")
        },
    }
}

fn copy_range<R, W>(
    source: &R,
    range: Range<u64>,
    sink: &mut W,
    buffer: &mut [u8],
    accounting: &mut ZipOperationAccounting,
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
        write_all_counted(
            sink,
            &buffer[..len],
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        )?;
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

fn central_record_span(
    record: &crate::ZipFileHeaderRecord<'_>,
    eocd_offset: u64,
) -> Result<(u64, usize, u64), Error> {
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
    Ok((central_offset, central_len, central_end))
}

fn map_metadata_limit_error(error: Error, comment_len: u64, limits: ArchiveLimits) -> Error {
    if let ErrorKind::LimitExceeded {
        resource: LimitResource::MetadataBytes,
        actual,
        ..
    } = error.kind()
    {
        let actual = comment_len.saturating_add(*actual);
        return limit_error(
            LimitResource::MetadataBytes,
            actual,
            limits.max_metadata_bytes,
        );
    }
    error
}

fn unsupported(reason: &'static str) -> Error {
    Error::from(ErrorKind::UnsupportedPreservation { reason })
}

fn limit_error(resource: LimitResource, actual: u64, maximum: u64) -> Error {
    Error::from(ErrorKind::LimitExceeded {
        resource,
        actual,
        maximum,
    })
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

    fn central_records(data: &[u8]) -> Vec<Vec<u8>> {
        let archive = ZipArchive::from_slice(data).unwrap();
        archive
            .entries()
            .map(|entry| {
                let entry = entry.unwrap();
                let start = usize::try_from(entry.central_directory_offset()).unwrap();
                let length =
                    ZipFileHeaderFixed::SIZE + usize::try_from(entry.metadata_size_hint()).unwrap();
                data[start..start + length].to_vec()
            })
            .collect()
    }

    fn central_record_without_offset(mut record: Vec<u8>) -> Vec<u8> {
        record[CENTRAL_LOCAL_HEADER_OFFSET].fill(0);
        record
    }

    fn with_comment(mut data: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let archive = ZipArchive::from_slice(&data).unwrap();
        let eocd = usize::try_from(archive.eocd_offset()).unwrap();
        data[eocd + 20..eocd + 22].copy_from_slice(&(comment.len() as u16).to_le_bytes());
        data.extend_from_slice(comment);
        data
    }

    fn with_file_comment(mut data: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let archive = ZipArchive::from_slice(&data).unwrap();
        let central = usize::try_from(archive.directory_offset()).unwrap();
        let eocd = usize::try_from(archive.eocd_offset()).unwrap();
        data[central + 32..central + 34].copy_from_slice(&(comment.len() as u16).to_le_bytes());
        let central_size = u32::from_le_bytes(data[eocd + 12..eocd + 16].try_into().unwrap());
        data[eocd + 12..eocd + 16]
            .copy_from_slice(&(central_size + comment.len() as u32).to_le_bytes());
        data.splice(eocd..eocd, comment.iter().copied());
        data
    }

    fn central_metadata_bytes(data: &[u8]) -> u64 {
        let archive = ZipArchive::from_slice(data).unwrap();
        archive
            .entries()
            .map(|entry| entry.unwrap().metadata_size_hint())
            .sum::<u64>()
            + archive.comment().as_bytes().len() as u64
    }

    fn stored_archive_with_metadata(name: &str, extra: &[u8]) -> Vec<u8> {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        let (mut file, config) = writer
            .new_file(name)
            .extra_field(
                crate::extra_fields::ExtraFieldId::new(0xaaaa),
                extra,
                crate::Header::CENTRAL,
            )
            .unwrap()
            .start()
            .unwrap();
        let mut data_writer = config.wrap(&mut file);
        data_writer.write_all(b"payload").unwrap();
        let (_, descriptor) = data_writer.finish().unwrap();
        file.finish(descriptor).unwrap();
        writer.finish().unwrap()
    }

    fn many_small_archive(count: usize) -> Vec<u8> {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        for index in 0..count {
            let name = format!("f{index}");
            writer.write_stored_file(&name, &[]).unwrap();
        }
        writer.finish().unwrap()
    }

    fn with_reordered_central(mut data: Vec<u8>) -> Vec<u8> {
        let archive = ZipArchive::from_slice(&data).unwrap();
        let central = usize::try_from(archive.directory_offset()).unwrap();
        let eocd = usize::try_from(archive.eocd_offset()).unwrap();
        let first_len = ZipFileHeaderFixed::SIZE
            + ZipFileHeaderFixed::parse(&data[central..])
                .unwrap()
                .variable_length();
        let first = data[central..central + first_len].to_vec();
        let rest = data[central + first_len..eocd].to_vec();
        data[central..eocd].copy_from_slice(&[rest, first].concat());
        data
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
    fn bounded_preservation_counts_files_not_directory_records() {
        let data = with_comment(ordinary_archive(), b"preserved comment");
        let (archive, mut buffer) = indexed(&data);
        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_files = 2;

        let index = PreservationIndex::new_with_limits(&archive, &mut buffer, limits).unwrap();
        assert_eq!(index.entries().len(), 3);
        assert_eq!(
            index
                .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                .unwrap(),
            data
        );
    }

    #[test]
    fn bounded_preservation_reports_member_and_metadata_limits_before_ownership() {
        let data = stored_archive_with_metadata("payload.bin", b"metadata");
        let metadata_bytes = central_metadata_bytes(&data);

        let (archive, mut buffer) = indexed(&data);
        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_member_name_bytes = b"payload.bi".len() as u64;
        let error = match PreservationIndex::new_with_limits(&archive, &mut buffer, limits) {
            Ok(_) => panic!("member-name limit should reject the source"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::MemberNameBytes,
                actual,
                maximum,
            } if *actual == b"payload.bin".len() as u64 && *maximum == b"payload.bi".len() as u64
        ));

        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_metadata_bytes = metadata_bytes - 1;
        let (archive, mut buffer) = indexed(&data);
        let error = match PreservationIndex::new_with_limits(&archive, &mut buffer, limits) {
            Ok(_) => panic!("metadata limit should reject the source"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::MetadataBytes,
                actual,
                maximum,
            } if *actual == metadata_bytes && *maximum == metadata_bytes - 1
        ));
    }

    #[test]
    fn bounded_preservation_includes_eocd_comment_in_metadata_and_unbounded_is_exact() {
        let data = with_file_comment(
            with_comment(
                stored_archive_with_metadata("payload.bin", b"metadata"),
                b"eocd",
            ),
            b"file",
        );
        let metadata_bytes = central_metadata_bytes(&data);

        let (archive, mut buffer) = indexed(&data);
        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_metadata_bytes = metadata_bytes - 1;
        let error = match PreservationIndex::new_with_limits(&archive, &mut buffer, limits) {
            Ok(_) => panic!("EOCD/comment metadata limit should reject the source"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::MetadataBytes,
                actual,
                maximum,
            } if *actual == metadata_bytes && *maximum == metadata_bytes - 1
        ));

        let (archive, mut buffer) = indexed(&data);
        let index =
            PreservationIndex::new_with_limits(&archive, &mut buffer, ArchiveLimits::UNBOUNDED)
                .unwrap();
        assert_eq!(
            index
                .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                .unwrap(),
            data
        );
    }

    #[test]
    fn bounded_preservation_rejects_file_count_before_owned_entry_reservation() {
        let data = many_small_archive(4);
        let (archive, mut buffer) = indexed(&data);
        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_files = 0;
        let error = match PreservationIndex::new_with_limits(&archive, &mut buffer, limits) {
            Ok(_) => panic!("file-count limit should reject the source"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::FileCount,
                actual: 1,
                maximum: 0,
            }
        ));
    }

    #[test]
    fn bounded_preservation_keeps_non_metadata_error_identity() {
        let original = Error::from(io::Error::new(io::ErrorKind::BrokenPipe, "central read"));
        let mapped = map_metadata_limit_error(original, 7, ArchiveLimits::UNBOUNDED);
        assert!(matches!(
            mapped.kind(),
            ErrorKind::IO(error)
                if error.kind() == io::ErrorKind::BrokenPipe && error.to_string() == "central read"
        ));
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

        let mut plan = PreservationPlan::copy_all(&index);
        plan.try_append(RegeneratedEntry::new("appended.bin", b"payload".to_vec()))
            .unwrap();
        let mut sink = b"untouched".to_vec();
        let error = index
            .write_to(&plan, &mut sink)
            .expect_err("new members must reject mismatched source provenance");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(sink, b"untouched");
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

        let replacement = output_archive
            .entries()
            .map(|entry| entry.unwrap())
            .find(|entry| entry.file_path().as_ref() == b"replacement.bin")
            .unwrap();
        assert_eq!(replacement.compression_method(), CompressionMethod::Deflate);
        assert!(replacement.has_data_descriptor());
    }

    #[test]
    fn accounting_distinguishes_changed_store_and_deflate_members() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let mut plan = PreservationPlan::copy_all(&index);
        let store = b"replacement store";
        let deflate = b"replacement deflate payload";
        plan.actions[0] = PreservationAction::Regenerate {
            id: index.entries()[0].id(),
            entry: RegeneratedEntry::new("changed-store", store.as_slice()),
        };
        plan.actions[1] = PreservationAction::Regenerate {
            id: index.entries()[1].id(),
            entry: RegeneratedEntry::new("changed-deflate", deflate.as_slice())
                .compression_method(CompressionMethod::Deflate),
        };

        let mut accounting = ZipOperationAccounting::default();
        let output = index
            .write_to_with_accounting(&plan, Vec::new(), &mut accounting)
            .unwrap();
        let reader = crate::office::ArchiveReader::new(&output).unwrap();
        assert_eq!(reader.read("changed-store").unwrap(), store);
        assert_eq!(reader.read("changed-deflate").unwrap(), deflate);
        assert_eq!(
            accounting.stored_payload_bytes_emitted(),
            store.len() as u64
        );
        assert!(accounting.generated_deflate_payload_bytes_emitted() > 0);
        assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);

        let unchanged = &index.entries()[2];
        let unchanged_local = unchanged.local_span();
        let unchanged_central = unchanged.central_record();
        let expected_raw = unchanged_local.end - unchanged_local.start + unchanged_central.end
            - unchanged_central.start
            - CENTRAL_LOCAL_HEADER_OFFSET.len() as u64;
        assert_eq!(
            accounting.raw_unchanged_source_bytes_accepted(),
            expected_raw
        );
    }

    #[test]
    fn accounting_charges_generated_payloads_on_partial_publication() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let store = b"partial store payload";
        let mut store_plan = PreservationPlan::copy_all(&index);
        store_plan.actions[0] = PreservationAction::Regenerate {
            id: index.entries()[0].id(),
            entry: RegeneratedEntry::new("partial-store", store.as_slice()),
        };
        let store_payload_start = ZipLocalFileHeaderFixed::SIZE + b"partial-store".len();
        let mut store_sink = PartialFailingSink::new(store_payload_start + 3, usize::MAX);
        let mut store_accounting = ZipOperationAccounting::default();
        let error = index
            .write_to_with_accounting(&store_plan, &mut store_sink, &mut store_accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(store_accounting.stored_payload_bytes_emitted(), 3);

        let deflate = b"partial deflate payload with repeated repeated repeated bytes";
        let mut deflate_plan = PreservationPlan::copy_all(&index);
        deflate_plan.actions[0] = PreservationAction::Regenerate {
            id: index.entries()[0].id(),
            entry: RegeneratedEntry::new("partial-deflate", deflate.as_slice())
                .compression_method(CompressionMethod::Deflate),
        };
        let deflate_payload_start = ZipLocalFileHeaderFixed::SIZE + b"partial-deflate".len();
        let mut deflate_sink = PartialFailingSink::new(deflate_payload_start + 3, usize::MAX);
        let mut deflate_accounting = ZipOperationAccounting::default();
        let error = index
            .write_to_with_accounting(&deflate_plan, &mut deflate_sink, &mut deflate_accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(
            deflate_accounting.generated_deflate_payload_bytes_emitted(),
            3
        );
    }

    #[test]
    fn regenerated_entry_can_retain_a_shared_payload() {
        let data = Arc::new(b"shared generated content".to_vec());
        let entry = RegeneratedEntry::new_shared("shared.bin", Arc::clone(&data));

        assert_eq!(Arc::strong_count(&data), 2);

        let prepared = generated_entry(&entry).unwrap();
        let PreparedLocal::Shared {
            bytes: local_bytes, ..
        } = &prepared.local
        else {
            panic!("generated entry must retain the shared archive buffer");
        };
        let PreparedCentral::Shared {
            bytes: central_bytes,
            ..
        } = &prepared.central
        else {
            panic!("generated entry must retain the shared archive buffer");
        };
        assert!(Arc::ptr_eq(local_bytes, central_bytes));
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
        let PreparedLocal::Shared {
            bytes: local_bytes,
            range: local_range,
        } = &prepared.local
        else {
            panic!("generated entry must retain the shared archive buffer");
        };
        let local = &local_bytes[local_range.clone()];
        let local_header = ZipLocalFileHeaderFixed::parse(local).unwrap();
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
    fn generated_entry_reserves_both_long_member_name_copies() {
        let name = "n".repeat(u16::MAX as usize);
        let entry = RegeneratedEntry::new(name.clone(), Vec::new());
        let prepared = generated_entry(&entry).unwrap();
        let PreparedLocal::Shared { bytes, range } = &prepared.local else {
            panic!("generated entry must retain the shared archive buffer");
        };
        let expected_capacity = name
            .len()
            .checked_mul(2)
            .and_then(|size| size.checked_add(4 * 1024))
            .unwrap();
        assert!(bytes.capacity() >= expected_capacity);
        assert!(range.end - range.start <= expected_capacity);
    }

    #[test]
    fn appends_store_and_deflate_members_after_preserved_source_order() {
        let source = with_comment(
            with_reordered_central(ordinary_archive()),
            b"append comment",
        );
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let source_local_spans: Vec<_> = index
            .entries()
            .iter()
            .map(|entry| {
                let span = entry.local_span();
                (
                    span.clone(),
                    source[span.start as usize..span.end as usize].to_vec(),
                )
            })
            .collect();
        let source_central = central_records(&source)
            .into_iter()
            .map(central_record_without_offset)
            .collect::<Vec<_>>();

        let mut plan = PreservationPlan::copy_all(&index);
        plan.try_reserve_appended(2).unwrap();
        plan.try_append(RegeneratedEntry::new(
            "appended-store.bin",
            b"stored".to_vec(),
        ))
        .unwrap();
        plan.try_append(
            RegeneratedEntry::new("appended-deflate.bin", b"deflated".to_vec())
                .compression_method(CompressionMethod::Deflate),
        )
        .unwrap();
        assert_eq!(plan.appended().len(), 2);

        let output = index.write_to(&plan, Vec::new()).unwrap();
        let output_archive = ZipArchive::from_slice(&output).unwrap();
        let output_names: Vec<_> = output_archive
            .entries()
            .map(|entry| entry.unwrap().file_path().as_ref().to_vec())
            .collect();
        let mut expected_names: Vec<_> = ZipArchive::from_slice(&source)
            .unwrap()
            .entries()
            .map(|entry| entry.unwrap().file_path().as_ref().to_vec())
            .collect();
        expected_names.extend([
            b"appended-store.bin".to_vec(),
            b"appended-deflate.bin".to_vec(),
        ]);
        assert_eq!(output_names, expected_names);

        for (span, bytes) in source_local_spans {
            assert_eq!(&output[span.start as usize..span.end as usize], bytes);
        }
        let output_central = central_records(&output);
        assert_eq!(
            output_central[..source_central.len()]
                .iter()
                .cloned()
                .map(central_record_without_offset)
                .collect::<Vec<_>>(),
            source_central
        );
        assert_eq!(output_archive.comment().as_bytes(), b"append comment");

        let output_entries: Vec<_> = output_archive
            .entries()
            .map(|entry| entry.unwrap())
            .collect();
        let appended_start = output_entries.len() - 2;
        let source_local_end = index
            .entries()
            .iter()
            .map(|entry| entry.local_span().end)
            .max()
            .unwrap_or_default();
        assert!(
            output_entries[appended_start].local_header_offset()
                < output_entries[appended_start + 1].local_header_offset()
        );
        assert!(output_entries[appended_start].local_header_offset() >= source_local_end);
        let output_reader = crate::office::ArchiveReader::new(&output).unwrap();
        assert_eq!(
            output_reader
                .read(
                    std::str::from_utf8(output_entries[appended_start].file_path().as_ref())
                        .unwrap()
                )
                .unwrap(),
            b"stored"
        );
        assert_eq!(
            output_reader
                .read(
                    std::str::from_utf8(output_entries[appended_start + 1].file_path().as_ref(),)
                        .unwrap(),
                )
                .unwrap(),
            b"deflated"
        );
    }

    #[test]
    fn omitting_an_appended_suffix_restores_the_exact_source_archive() {
        let source = with_comment(
            with_reordered_central(ordinary_archive()),
            b"append then omit",
        );
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let mut append_plan = PreservationPlan::copy_all(&index);
        append_plan
            .try_append(RegeneratedEntry::new(
                "appended-store.bin",
                b"stored".to_vec(),
            ))
            .unwrap();
        append_plan
            .try_append(
                RegeneratedEntry::new("appended-deflate.bin", b"deflated".to_vec())
                    .compression_method(CompressionMethod::Deflate),
            )
            .unwrap();
        let appended = index.write_to(&append_plan, Vec::new()).unwrap();

        let (appended_archive, mut appended_buffer) = indexed(&appended);
        let appended_index = PreservationIndex::new(&appended_archive, &mut appended_buffer)
            .expect("appended archive remains preservable");
        let mut restore_plan = PreservationPlan::copy_all(&appended_index);
        for entry in appended_index.entries() {
            if matches!(
                entry.raw_name_bytes(),
                b"appended-store.bin" | b"appended-deflate.bin"
            ) {
                let id = entry.id();
                restore_plan.actions[usize::try_from(id.0).unwrap()] = PreservationAction::Omit(id);
            }
        }

        let restored = appended_index.write_to(&restore_plan, Vec::new()).unwrap();
        assert_eq!(restored, source);
    }

    #[test]
    fn omitting_a_middle_source_member_reopens_and_preserves_retained_raw_records() {
        let source = with_comment(ordinary_archive(), b"middle omit");
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let omitted_id = index.entries()[1].id();
        let omitted_name = index.entries()[1].raw_name_bytes().to_vec();

        let mut plan = PreservationPlan::copy_all(&index);
        plan.actions[1] = PreservationAction::Omit(omitted_id);
        let output = index.write_to(&plan, Vec::new()).unwrap();

        let output_archive = ZipArchive::from_slice(&output).unwrap();
        let output_names: Vec<_> = output_archive
            .entries()
            .map(|entry| entry.unwrap().file_path().as_ref().to_vec())
            .collect();
        assert!(!output_names.iter().any(|name| name == &omitted_name));
        assert_eq!(output_archive.comment().as_bytes(), b"middle omit");

        let (output_archive, mut output_buffer) = indexed(&output);
        let output_index = PreservationIndex::new(&output_archive, &mut output_buffer).unwrap();
        for source_entry in index
            .entries()
            .iter()
            .filter(|entry| entry.id() != omitted_id)
        {
            let output_entry = output_index
                .entries()
                .iter()
                .find(|entry| entry.raw_name_bytes() == source_entry.raw_name_bytes())
                .expect("retained member remains present");
            let source_span = source_entry.local_span();
            let output_span = output_entry.local_span();
            assert_eq!(
                &source[usize::try_from(source_span.start).unwrap()
                    ..usize::try_from(source_span.end).unwrap()],
                &output[usize::try_from(output_span.start).unwrap()
                    ..usize::try_from(output_span.end).unwrap()]
            );
            assert_eq!(
                central_record_without_offset(
                    source[usize::try_from(source_entry.central_record().start).unwrap()
                        ..usize::try_from(source_entry.central_record().end).unwrap()]
                        .to_vec()
                ),
                central_record_without_offset(
                    output[usize::try_from(output_entry.central_record().start).unwrap()
                        ..usize::try_from(output_entry.central_record().end).unwrap()]
                        .to_vec()
                )
            );
        }
        let output_reader = crate::office::ArchiveReader::new(&output).unwrap();
        assert_eq!(output_reader.read("first.bin").unwrap(), b"stored data");
    }

    #[test]
    fn replacement_then_append_patches_source_and_appended_offsets() {
        let source = ordinary_archive();
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let mut plan = PreservationPlan::copy_all(&index);
        plan.actions[0] = PreservationAction::Regenerate {
            id: index.entries()[0].id(),
            entry: RegeneratedEntry::new("replaced.bin", vec![b'x'; 128 * 1024])
                .compression_method(CompressionMethod::Deflate),
        };
        plan.try_append(RegeneratedEntry::new("tail.bin", b"tail".to_vec()))
            .unwrap();

        let output = index.write_to(&plan, Vec::new()).unwrap();
        let output_archive = ZipArchive::from_slice(&output).unwrap();
        let entries: Vec<_> = output_archive
            .entries()
            .map(|entry| entry.unwrap())
            .collect();
        let names: Vec<_> = entries
            .iter()
            .map(|entry| entry.file_path().as_ref().to_vec())
            .collect();
        assert_eq!(
            names,
            vec![
                b"replaced.bin".to_vec(),
                b"second.bin".to_vec(),
                b"folder/".to_vec(),
                b"tail.bin".to_vec(),
            ]
        );
        assert!(entries[1].local_header_offset() > entries[0].local_header_offset());
        assert!(entries[3].local_header_offset() > entries[2].local_header_offset());

        let reader = crate::office::ArchiveReader::new(&output).unwrap();
        assert_eq!(reader.read("replaced.bin").unwrap(), vec![b'x'; 128 * 1024]);
        assert_eq!(reader.read("tail.bin").unwrap(), b"tail");
    }

    #[test]
    fn incomplete_or_unrepresentable_append_plan_leaves_sink_untouched() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let mut incomplete_sink = b"untouched".to_vec();
        let incomplete = PreservationPlan::new();
        let error = index
            .write_to(&incomplete, &mut incomplete_sink)
            .expect_err("incomplete source coverage must fail before output");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(incomplete_sink, b"untouched");

        let mut unsupported_sink = b"untouched".to_vec();
        let mut unsupported = PreservationPlan::copy_all(&index);
        unsupported
            .try_append(
                RegeneratedEntry::new("unsupported.bin", b"payload".to_vec())
                    .compression_method(CompressionMethod::Bzip2),
            )
            .unwrap();
        let error = index
            .write_to(&unsupported, &mut unsupported_sink)
            .expect_err("unsupported generated compression must fail before output");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(unsupported_sink, b"untouched");
    }

    #[test]
    fn rejects_the_zip32_entry_count_sentinel() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let mut prepared = Vec::with_capacity(u16::MAX as usize);
        prepared.resize_with(u16::MAX as usize, || PreparedEntry {
            local: PreparedLocal::Generated(Vec::new()),
            central: PreparedCentral::Generated(Vec::new()),
            generated_payload: None,
            omitted: false,
        });
        let error = index
            .validate_output_layout(&prepared)
            .expect_err("0xffff entries require ZIP64");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
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

    #[derive(Debug)]
    struct PartialFailingSink {
        bytes: Vec<u8>,
        fail_after: usize,
        max_write: usize,
    }

    impl PartialFailingSink {
        fn new(fail_after: usize, max_write: usize) -> Self {
            Self {
                bytes: Vec::new(),
                fail_after,
                max_write,
            }
        }
    }

    impl Write for PartialFailingSink {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            if self.bytes.len() >= self.fail_after {
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, "sink failed"));
            }
            let available = self.fail_after - self.bytes.len();
            let written = buffer.len().min(self.max_write).min(available);
            self.bytes.extend_from_slice(&buffer[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn accounting_counts_exact_raw_preservation_and_partial_sink_progress() {
        let data = with_comment(ordinary_archive(), b"raw comment");
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let output = index
            .write_to_with_accounting(
                &PreservationPlan::copy_all(&index),
                Vec::new(),
                &mut accounting,
            )
            .unwrap();
        let expected = index
            .entries()
            .iter()
            .map(|entry| entry.local_span().end - entry.local_span().start)
            .sum::<u64>()
            + index
                .entries()
                .iter()
                .map(|entry| {
                    (entry.central_record().end - entry.central_record().start)
                        - CENTRAL_LOCAL_HEADER_OFFSET.len() as u64
                })
                .sum::<u64>();
        let expected = expected + b"raw comment".len() as u64;
        assert_eq!(output, data);
        assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), expected);

        let mut sink = PartialFailingSink::new(8, 5);
        let mut partial_accounting = ZipOperationAccounting::default();
        let error = index
            .write_to_with_accounting(
                &PreservationPlan::copy_all(&index),
                &mut sink,
                &mut partial_accounting,
            )
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(
            partial_accounting.raw_unchanged_source_bytes_accepted(),
            sink.bytes.len() as u64
        );
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
