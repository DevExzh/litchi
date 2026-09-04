//! Conservative raw-member preservation for ZIP archive rewrites.
//!
//! This module intentionally accepts only single-disk ZIP archives whose
//! framing can be copied without interpretation or normalization.
//! It validates source layout/range declarations and every requested action
//! before writing to the destination, so an unsupported layout never produces
//! a partial preserved archive.

use crate::office::ArchiveLimits;
use crate::{
    CompressionMethod, EndOfCentralDirectoryRecordFixed, Error, ErrorKind, LimitResource, ReaderAt,
    ZipArchive, ZipArchiveWriter, ZipFileHeaderFixed,
    accounting::{AccountingWriteKind, ZipOperationAccounting, usize_to_u64, write_all_counted},
};
use std::io::Write;
use std::ops::Range;
use std::sync::Arc;

const COPY_CHUNK_SIZE: usize = 32 * 1024;
const CENTRAL_LOCAL_HEADER_OFFSET: Range<usize> = 42..46;
const ZIP64_LOCATOR_SIZE: usize = 20;
const ZIP64_EOCD_FIXED_SIZE: usize = 56;
const ZIP64_EOCD_FIXED_PAYLOAD_SIZE: u64 = 44;
const ZIP64_LOCATOR_SIGNATURE: u32 = 0x0706_4b50;

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
    wayfinder: crate::ZipArchiveEntryWayfinder,
    central_local_header_offset_patch: CentralLocalHeaderOffsetPatch,
    compression_method: CompressionMethod,
    local_central_name_mismatch: bool,
}

#[derive(Debug, Clone)]
enum CentralLocalHeaderOffsetPatch {
    Fixed32,
    Zip64(Range<usize>),
}

#[derive(Debug, Clone)]
struct Zip64Tail {
    eocd: Vec<u8>,
    locator: Vec<u8>,
    classic_eocd: Vec<u8>,
}

impl Zip64Tail {
    fn extensible_data_len(&self) -> Result<usize, Error> {
        self.eocd
            .len()
            .checked_sub(ZIP64_EOCD_FIXED_SIZE)
            .ok_or_else(|| unsupported("ZIP64 EOCD framing"))
    }

    fn len(&self) -> Result<u64, Error> {
        let byte_len = self
            .eocd
            .len()
            .checked_add(self.locator.len())
            .and_then(|length| length.checked_add(self.classic_eocd.len()))
            .ok_or_else(|| unsupported("ZIP64 tail length"))?;
        usize_to_u64(byte_len, "ZIP64 tail length")
    }
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
        let name_len = usize::from(u16::from_le_bytes([
            self.central_bytes[28],
            self.central_bytes[29],
        ]));
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

/// The policy governing the preservation format accepted by the index.
#[allow(dead_code)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PreservationPolicy {
    /// Preserve ZIP32 framing only. This is the public/default contract.
    Zip32Only,
    /// Permit the internal ZIP64 framing-preservation foundation.
    AllowZip64,
}

/// An indexed, structurally validated single-disk ZIP source.
pub struct PreservationIndex<'source, R> {
    source: &'source R,
    entries: Vec<PreservedEntry>,
    local_order: Vec<usize>,
    archive_comment: Vec<u8>,
    zip64_tail: Option<Zip64Tail>,
    archive_end_offset: u64,
}

impl<'source, R> PreservationIndex<'source, R>
where
    R: ReaderAt,
{
    /// Builds a preservation index without reading member bodies.
    ///
    /// Multi-disk, prefixed, ambiguous, overlapping, and truncated layouts are
    /// rejected with [`ErrorKind::UnsupportedPreservation`] before a caller
    /// can begin writing a plan. ZIP64 sources are rejected by this
    /// public/default constructor until a higher-level integration opts into
    /// the internal ZIP64 preservation foundation.
    pub fn new(archive: &'source ZipArchive<R>, buffer: &mut [u8]) -> Result<Self, Error> {
        Self::new_with_limits_and_policy(
            archive,
            buffer,
            ArchiveLimits::default(),
            PreservationPolicy::Zip32Only,
        )
    }

    /// Builds a preservation index under explicit archive metadata limits.
    ///
    /// Preservation retains every source central-directory record verbatim,
    /// including records for directory members and the EOCD comment. The
    /// archive profile's `max_files` bounds non-directory source members;
    /// directory records are retained too, but consume only the metadata
    /// budget. `max_metadata_bytes` covers the aggregate variable central
    /// metadata, ZIP64 extensible data, and EOCD comment. Limits are validated
    /// in a metadata-only pass before any owned entry or raw-record buffers are
    /// reserved. ZIP64 sources remain refused by this public constructor; the
    /// internal [`PreservationPolicy::AllowZip64`] policy is intentionally not
    /// used by OPC yet.
    pub fn new_with_limits(
        archive: &'source ZipArchive<R>,
        buffer: &mut [u8],
        limits: ArchiveLimits,
    ) -> Result<Self, Error> {
        Self::new_with_limits_and_policy(archive, buffer, limits, PreservationPolicy::Zip32Only)
    }

    /// Builds the internal ZIP64 preservation foundation under explicit
    /// archive metadata limits. OPC currently deliberately does not call this
    /// constructor; it is kept crate-private while the integration contract
    /// remains refusal for mutated ZIP64 sources.
    #[allow(dead_code)]
    pub(crate) fn new_with_policy(
        archive: &'source ZipArchive<R>,
        buffer: &mut [u8],
        policy: PreservationPolicy,
    ) -> Result<Self, Error> {
        Self::new_with_limits_and_policy(archive, buffer, ArchiveLimits::default(), policy)
    }

    pub(crate) fn new_with_limits_and_policy(
        archive: &'source ZipArchive<R>,
        buffer: &mut [u8],
        limits: ArchiveLimits,
        policy: PreservationPolicy,
    ) -> Result<Self, Error> {
        if archive.is_zip64() && policy == PreservationPolicy::Zip32Only {
            return Err(unsupported("ZIP64 preservation is not enabled"));
        }

        let central_start = archive.directory_offset();
        let head_eocd_offset = archive.head_eocd_offset();
        let eocd_offset = archive.eocd_offset();
        let archive_end = archive.end_offset();
        if central_start > head_eocd_offset
            || head_eocd_offset > eocd_offset
            || archive_end < eocd_offset
        {
            return Err(unsupported("invalid central-directory bounds"));
        }

        let mut eocd_bytes = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
        archive
            .get_ref()
            .read_exact_at(&mut eocd_bytes, eocd_offset)?;
        let eocd = EndOfCentralDirectoryRecordFixed::parse(&eocd_bytes)?;
        if !valid_classic_disk_fields(&eocd, archive.is_zip64()) {
            return Err(unsupported("multi-disk archives"));
        }
        if archive.is_zip64()
            && !valid_classic_zip64_fields(
                &eocd,
                archive.entries_hint(),
                archive.central_directory_size(),
                archive.directory_offset(),
            )
        {
            return Err(unsupported("classic EOCD disagrees with ZIP64 metadata"));
        }

        let eocd_fixed_len = usize_to_u64(eocd_bytes.len(), "archive comment length")?;
        let comment_len_u64 = archive_end
            .checked_sub(eocd_offset)
            .and_then(|length| length.checked_sub(eocd_fixed_len))
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

        let zip64_tail = if archive.is_zip64() {
            let max_extensible_data = limits
                .max_metadata_bytes
                .checked_sub(comment_len_u64)
                .ok_or_else(|| unsupported("ZIP64 metadata budget"))?;
            Some(
                read_zip64_tail(
                    archive.get_ref(),
                    archive.head_eocd_offset(),
                    archive.eocd_offset(),
                    archive.entries_hint(),
                    archive.central_directory_size(),
                    archive.directory_offset(),
                    max_extensible_data,
                )
                .map_err(|error| map_metadata_limit_error(error, comment_len_u64, limits))?,
            )
        } else {
            None
        };
        let zip64_extensible_data_u64 = zip64_tail
            .as_ref()
            .map(|tail| {
                tail.extensible_data_len()
                    .and_then(|length| usize_to_u64(length, "ZIP64 extensible data"))
            })
            .transpose()?
            .unwrap_or(0);
        let metadata_prefix_u64 = comment_len_u64
            .checked_add(zip64_extensible_data_u64)
            .ok_or_else(|| unsupported("metadata budget overflow"))?;
        if metadata_prefix_u64 > limits.max_metadata_bytes {
            return Err(limit_error(
                LimitResource::MetadataBytes,
                metadata_prefix_u64,
                limits.max_metadata_bytes,
            ));
        }

        let entry_count_hint = archive.entries_hint();
        let max_files_u64 = match u64::try_from(limits.max_files) {
            Ok(value) => value,
            Err(_) => u64::MAX,
        };
        let central_metadata_budget = limits
            .max_metadata_bytes
            .checked_sub(metadata_prefix_u64)
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
            .map_err(|error| map_metadata_limit_error(error, metadata_prefix_u64, limits))?
        {
            if policy == PreservationPolicy::Zip32Only && record.is_zip64() {
                return Err(unsupported("ZIP64 entry preservation is not enabled"));
            }
            if !record.wayfinder().borrowed_provenance_supported() {
                return Err(unsupported("unresolved ZIP64 entry fields"));
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

            central_record_span(&record, head_eocd_offset)?;
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
        let central_fixed_record_len = usize_to_u64(
            ZipFileHeaderFixed::SIZE,
            "central-directory fixed record length",
        )?;
        let fixed_central_bytes = entry_count_u64
            .checked_mul(central_fixed_record_len)
            .ok_or_else(|| unsupported("central-directory source budget"))?;
        let source_budget = limits
            .max_metadata_bytes
            .saturating_add(fixed_central_bytes);

        let comment_offset = eocd_offset
            .checked_add(eocd_fixed_len)
            .ok_or_else(|| unsupported("archive comment offset"))?;
        let archive_comment = read_vec(archive.get_ref(), comment_offset, comment_len)?;

        let entry_count = usize::try_from(entry_count_u64)
            .map_err(|_| unsupported("central-directory entry count"))?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(entry_count)
            .map_err(|source| allocation("preservation-index entries", source))?;

        let mut iterator = archive.entries_with_metadata_limit(buffer, central_metadata_budget);
        let mut retained_source_bytes = metadata_prefix_u64;
        while let Some(record) = iterator
            .next_entry()
            .map_err(|error| map_metadata_limit_error(error, metadata_prefix_u64, limits))?
        {
            if policy == PreservationPolicy::Zip32Only && record.is_zip64() {
                return Err(unsupported("ZIP64 entry preservation is not enabled"));
            }
            let wayfinder = record.wayfinder();
            if !wayfinder.borrowed_provenance_supported() {
                return Err(unsupported("unresolved ZIP64 entry fields"));
            }

            let (central_offset, central_len, central_end) =
                central_record_span(&record, head_eocd_offset)?;
            let central_len_u64 = usize_to_u64(central_len, "central-directory record length")?;
            let next_retained_source_bytes = retained_source_bytes
                .checked_add(central_len_u64)
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
            let central_local_header_offset_patch = central_local_header_offset_patch(
                &record,
                &central_fixed,
                &central_bytes,
                central_offset,
                central_end,
            )?;

            let id = PreservationEntryId(
                u32::try_from(entries.len()).map_err(|_| unsupported("too many entries"))?,
            );
            entries.push(PreservedEntry {
                id,
                local_span: record.local_header_offset()..record.local_header_offset(),
                central_record: central_offset..central_end,
                central_bytes,
                wayfinder,
                central_local_header_offset_patch,
                compression_method: central_fixed.compression_method.as_method(),
                local_central_name_mismatch: false,
            });
        }

        if usize_to_u64(entries.len(), "central-directory entry count")? != entry_count_u64 {
            return Err(unsupported("central-directory entry count mismatch"));
        }

        let mut expected_central = central_start;
        for entry in &entries {
            if entry.central_record.start != expected_central {
                return Err(unsupported("non-contiguous central-directory records"));
            }
            expected_central = entry.central_record.end;
        }
        if expected_central != head_eocd_offset {
            return Err(unsupported("central-directory trailing bytes"));
        }

        let mut local_order = Vec::new();
        local_order
            .try_reserve_exact(entries.len())
            .map_err(|source| allocation("preservation local order", source))?;
        local_order.extend(0..entries.len());
        local_order.sort_unstable_by_key(|&index| entries[index].local_span.start);
        if entries.is_empty() && (central_start != 0 || head_eocd_offset != 0) {
            return Err(unsupported("archive prelude data"));
        }
        if let Some(&first) = local_order.first() {
            if entries[first].local_span.start != 0 {
                return Err(unsupported("archive prelude data"));
            }
        }

        for position in 0..local_order.len() {
            let index = local_order[position];
            let local_start = entries[index].local_span.start;
            let local_end = local_order
                .get(
                    position
                        .checked_add(1)
                        .ok_or_else(|| unsupported("local-member order"))?,
                )
                .map(|next| entries[*next].local_span.start)
                .unwrap_or(central_start);
            if local_start >= local_end || local_end > central_start {
                return Err(unsupported("overlapping or empty local-member spans"));
            }

            let local_central_name_mismatch =
                validate_local_span(archive, &entries[index], local_end)?;
            entries[index].local_central_name_mismatch = local_central_name_mismatch;
            entries[index].local_span.end = local_end;
        }

        Ok(Self {
            source: archive.get_ref(),
            entries,
            local_order,
            archive_comment,
            zip64_tail,
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
        let layout = self.validate_output_layout(&prepared)?;

        let mut copy_buffer = [0u8; COPY_CHUNK_SIZE];
        for &index in &self.local_order {
            if prepared[index].omitted {
                continue;
            }
            write_prepared_local(
                &prepared[index].local,
                prepared[index].generated_payload.as_ref(),
                self.source,
                &mut sink,
                &mut copy_buffer,
                accounting,
            )?;
        }
        for index in self.entries.len()..prepared.len() {
            if prepared[index].omitted {
                continue;
            }
            write_prepared_local(
                &prepared[index].local,
                prepared[index].generated_payload.as_ref(),
                self.source,
                &mut sink,
                &mut copy_buffer,
                accounting,
            )?;
        }

        for (entry, patch) in prepared
            .iter()
            .filter(|entry| !entry.omitted)
            .zip(layout.central_patches.iter())
        {
            let central = entry.central.bytes(&self.entries);
            let is_unchanged = matches!(&entry.central, PreparedCentral::Copy(_));
            write_prepared_central(&mut sink, central, patch, is_unchanged, accounting)?;
        }
        write_prepared_tail(&layout.tail, &mut sink, accounting)?;
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

    fn validate_output_layout(&self, prepared: &[PreparedEntry]) -> Result<OutputLayout, Error> {
        if prepared.len() < self.entries.len() {
            return Err(unsupported("prepared preservation plan length"));
        }
        let mut local_offsets = Vec::new();
        local_offsets
            .try_reserve_exact(prepared.len())
            .map_err(|source| allocation("preservation local offsets", source))?;
        local_offsets.resize(prepared.len(), 0_u64);
        let mut local_size = 0u64;
        for &index in &self.local_order {
            if prepared[index].omitted {
                continue;
            }
            validate_prepared_local(&prepared[index], self.archive_end_offset)?;
            local_offsets[index] = local_size;
            local_size = local_size
                .checked_add(prepared[index].local.len()?)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }
        for (index, entry) in prepared.iter().enumerate().skip(self.entries.len()) {
            if entry.omitted {
                continue;
            }
            validate_prepared_local(entry, self.archive_end_offset)?;
            local_offsets[index] = local_size;
            local_size = local_size
                .checked_add(entry.local.len()?)
                .ok_or_else(|| unsupported("output offset overflow"))?;
        }

        let entry_count = usize_to_u64(retained_entry_count(prepared), "preservation entry count")?;
        let central_size = prepared.iter().try_fold(0u64, |size, entry| {
            if entry.omitted {
                return Ok(size);
            }
            let central = entry.central.checked_bytes(&self.entries)?;
            let central_len = usize_to_u64(central.len(), "central-directory record length")?;
            size.checked_add(central_len)
                .ok_or_else(|| unsupported("output offset overflow"))
        })?;
        let central_start = local_size;
        let central_end = central_start
            .checked_add(central_size)
            .ok_or_else(|| unsupported("output offset overflow"))?;
        let archive_comment_len = usize_to_u64(self.archive_comment.len(), "archive comment")?;
        let tail_size = match self.zip64_tail.as_ref() {
            Some(tail) => tail.len()?,
            None => usize_to_u64(EndOfCentralDirectoryRecordFixed::SIZE, "fixed EOCD size")?,
        };
        let output_size = central_end
            .checked_add(tail_size)
            .and_then(|size| size.checked_add(archive_comment_len))
            .ok_or_else(|| unsupported("output offset overflow"))?;

        let mut central_patches = Vec::new();
        central_patches
            .try_reserve_exact(prepared.len())
            .map_err(|source| allocation("preservation central patches", source))?;
        for (index, entry) in prepared.iter().enumerate() {
            if entry.omitted {
                continue;
            }
            let central = entry.central.checked_bytes(&self.entries)?;
            let fixed32_offset = match &entry.central {
                PreparedCentral::Copy(source_index) => {
                    matches!(
                        &self.entries[*source_index].central_local_header_offset_patch,
                        CentralLocalHeaderOffsetPatch::Fixed32
                    )
                },
                PreparedCentral::Generated(_) | PreparedCentral::Shared { .. } => true,
            };
            if fixed32_offset && local_offsets[index] >= u64::from(u32::MAX) {
                return Err(unsupported("ZIP64 generated local-header offset"));
            }

            let patch = match &entry.central {
                PreparedCentral::Copy(source_index) => {
                    match &self.entries[*source_index].central_local_header_offset_patch {
                        CentralLocalHeaderOffsetPatch::Fixed32 => {
                            let offset = u32::try_from(local_offsets[index])
                                .map_err(|_| unsupported("ZIP64 output promotion"))?;
                            let mut bytes = [0u8; 8];
                            bytes[..4].copy_from_slice(&offset.to_le_bytes());
                            CentralOffsetPatch {
                                start: CENTRAL_LOCAL_HEADER_OFFSET.start,
                                end: CENTRAL_LOCAL_HEADER_OFFSET.end,
                                bytes,
                                len: 4,
                            }
                        },
                        CentralLocalHeaderOffsetPatch::Zip64(range) => {
                            let width = range
                                .end
                                .checked_sub(range.start)
                                .ok_or_else(|| unsupported("invalid central offset patch"))?;
                            if width != 8 || range.end > central.len() {
                                return Err(unsupported("invalid central offset patch"));
                            }
                            let mut bytes = [0u8; 8];
                            bytes.copy_from_slice(&local_offsets[index].to_le_bytes());
                            CentralOffsetPatch {
                                start: range.start,
                                end: range.end,
                                bytes,
                                len: 8,
                            }
                        },
                    }
                },
                PreparedCentral::Generated(_) | PreparedCentral::Shared { .. } => {
                    if CENTRAL_LOCAL_HEADER_OFFSET.end > central.len() {
                        return Err(unsupported("invalid central offset patch"));
                    }
                    let offset = u32::try_from(local_offsets[index])
                        .map_err(|_| unsupported("ZIP64 output promotion"))?;
                    let mut bytes = [0u8; 8];
                    bytes[..4].copy_from_slice(&offset.to_le_bytes());
                    CentralOffsetPatch {
                        start: CENTRAL_LOCAL_HEADER_OFFSET.start,
                        end: CENTRAL_LOCAL_HEADER_OFFSET.end,
                        bytes,
                        len: 4,
                    }
                },
            };
            central_patches.push(patch);
        }
        if (self.zip64_tail.is_none()
            && (local_size >= u64::from(u32::MAX)
                || central_size >= u64::from(u32::MAX)
                || output_size >= u64::from(u32::MAX)
                || retained_entry_count(prepared) >= usize::from(u16::MAX)))
            || self.archive_comment.len() > usize::from(u16::MAX)
        {
            return Err(unsupported("ZIP64 output promotion"));
        }

        let tail = if let Some(zip64_tail) = &self.zip64_tail {
            PreparedTail::Zip64(prepare_zip64_tail(
                zip64_tail,
                entry_count,
                central_size,
                central_start,
            )?)
        } else {
            PreparedTail::Zip32(prepare_zip32_tail(
                entry_count,
                central_size,
                central_start,
                archive_comment_len,
            )?)
        };
        Ok(OutputLayout {
            central_patches,
            tail,
        })
    }
}

#[derive(Debug)]
struct CentralOffsetPatch {
    start: usize,
    end: usize,
    bytes: [u8; 8],
    len: usize,
}

#[derive(Debug)]
struct PreparedZip64Tail {
    eocd: Vec<u8>,
    locator: [u8; ZIP64_LOCATOR_SIZE],
    classic_eocd: [u8; EndOfCentralDirectoryRecordFixed::SIZE],
}

#[derive(Debug)]
enum PreparedTail {
    Zip32([u8; EndOfCentralDirectoryRecordFixed::SIZE]),
    Zip64(PreparedZip64Tail),
}

#[derive(Debug)]
struct OutputLayout {
    central_patches: Vec<CentralOffsetPatch>,
    tail: PreparedTail,
}

fn write_prepared_central<W: Write>(
    sink: &mut W,
    bytes: &[u8],
    patch: &CentralOffsetPatch,
    count_source: bool,
    accounting: &mut ZipOperationAccounting,
) -> Result<(), Error> {
    debug_assert!(patch.start <= patch.end);
    debug_assert_eq!(patch.end.checked_sub(patch.start), Some(patch.len));
    debug_assert!(patch.end <= bytes.len());
    if count_source {
        write_all_counted(
            sink,
            &bytes[..patch.start],
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        )?;
    } else {
        sink.write_all(&bytes[..patch.start])?;
    }
    sink.write_all(&patch.bytes[..patch.len])?;
    if count_source {
        write_all_counted(
            sink,
            &bytes[patch.end..],
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        )?;
    } else {
        sink.write_all(&bytes[patch.end..])?;
    }
    Ok(())
}

fn write_prepared_tail<W: Write>(
    tail: &PreparedTail,
    sink: &mut W,
    accounting: &mut ZipOperationAccounting,
) -> Result<(), Error> {
    match tail {
        PreparedTail::Zip32(eocd) => sink.write_all(eocd)?,
        PreparedTail::Zip64(tail) => {
            write_all_counted(
                sink,
                &tail.eocd[..24],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;
            sink.write_all(&tail.eocd[24..56])?;
            write_all_counted(
                sink,
                &tail.eocd[56..],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;

            write_all_counted(
                sink,
                &tail.locator[..8],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;
            sink.write_all(&tail.locator[8..16])?;
            write_all_counted(
                sink,
                &tail.locator[16..],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;

            write_all_counted(
                sink,
                &tail.classic_eocd[..8],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;
            sink.write_all(&tail.classic_eocd[8..20])?;
            write_all_counted(
                sink,
                &tail.classic_eocd[20..],
                accounting,
                AccountingWriteKind::RawUnchangedSource,
            )?;
        },
    }
    Ok(())
}

fn prepare_zip32_tail(
    entry_count: u64,
    central_size: u64,
    central_offset: u64,
    comment_len: u64,
) -> Result<[u8; EndOfCentralDirectoryRecordFixed::SIZE], Error> {
    let entry_count =
        u16::try_from(entry_count).map_err(|_| unsupported("ZIP64 output promotion"))?;
    let central_size =
        u32::try_from(central_size).map_err(|_| unsupported("ZIP64 output promotion"))?;
    let central_offset =
        u32::try_from(central_offset).map_err(|_| unsupported("ZIP64 output promotion"))?;
    let comment_len =
        u16::try_from(comment_len).map_err(|_| unsupported("ZIP64 output promotion"))?;
    let mut eocd = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
    eocd[..4].copy_from_slice(&0x0605_4b50u32.to_le_bytes());
    eocd[8..10].copy_from_slice(&entry_count.to_le_bytes());
    eocd[10..12].copy_from_slice(&entry_count.to_le_bytes());
    eocd[12..16].copy_from_slice(&central_size.to_le_bytes());
    eocd[16..20].copy_from_slice(&central_offset.to_le_bytes());
    eocd[20..22].copy_from_slice(&comment_len.to_le_bytes());
    Ok(eocd)
}

fn prepare_zip64_tail(
    tail: &Zip64Tail,
    entry_count: u64,
    central_size: u64,
    central_offset: u64,
) -> Result<PreparedZip64Tail, Error> {
    if tail.eocd.len() < ZIP64_EOCD_FIXED_SIZE
        || tail.locator.len() != ZIP64_LOCATOR_SIZE
        || tail.classic_eocd.len() != EndOfCentralDirectoryRecordFixed::SIZE
    {
        return Err(unsupported("invalid retained ZIP64 tail"));
    }

    let size = u64::from_le_bytes(
        tail.eocd
            .get(4..12)
            .ok_or_else(|| unsupported("invalid retained ZIP64 tail"))?
            .try_into()
            .map_err(|_| unsupported("invalid retained ZIP64 tail"))?,
    );
    let expected_len = size
        .checked_add(12)
        .ok_or_else(|| unsupported("ZIP64 EOCD output length"))?;
    if usize_to_u64(tail.eocd.len(), "ZIP64 EOCD output length")? != expected_len {
        return Err(unsupported("invalid retained ZIP64 tail"));
    }
    let zip64_eocd_offset = central_offset
        .checked_add(central_size)
        .ok_or_else(|| unsupported("ZIP64 EOCD output offset"))?;

    let mut zip64_eocd = Vec::new();
    zip64_eocd
        .try_reserve_exact(tail.eocd.len())
        .map_err(|source| allocation("prepared ZIP64 EOCD", source))?;
    zip64_eocd.extend_from_slice(&tail.eocd);
    zip64_eocd[24..32].copy_from_slice(&entry_count.to_le_bytes());
    zip64_eocd[32..40].copy_from_slice(&entry_count.to_le_bytes());
    zip64_eocd[40..48].copy_from_slice(&central_size.to_le_bytes());
    zip64_eocd[48..56].copy_from_slice(&central_offset.to_le_bytes());
    let mut locator = [0u8; ZIP64_LOCATOR_SIZE];
    locator.copy_from_slice(&tail.locator);
    locator[8..16].copy_from_slice(&zip64_eocd_offset.to_le_bytes());

    let mut classic_eocd = [0u8; EndOfCentralDirectoryRecordFixed::SIZE];
    classic_eocd.copy_from_slice(&tail.classic_eocd);
    let original_num_entries = u16::from_le_bytes(
        tail.classic_eocd[8..10]
            .try_into()
            .map_err(|_| unsupported("invalid retained ZIP64 tail"))?,
    );
    let original_total_entries = u16::from_le_bytes(
        tail.classic_eocd[10..12]
            .try_into()
            .map_err(|_| unsupported("invalid retained ZIP64 tail"))?,
    );
    let original_central_size = u32::from_le_bytes(
        tail.classic_eocd[12..16]
            .try_into()
            .map_err(|_| unsupported("invalid retained ZIP64 tail"))?,
    );
    let original_central_offset = u32::from_le_bytes(
        tail.classic_eocd[16..20]
            .try_into()
            .map_err(|_| unsupported("invalid retained ZIP64 tail"))?,
    );
    classic_eocd[8..10].copy_from_slice(&zip32_u16_field(original_num_entries, entry_count)?);
    classic_eocd[10..12].copy_from_slice(&zip32_u16_field(original_total_entries, entry_count)?);
    classic_eocd[12..16].copy_from_slice(&zip32_u32_field(original_central_size, central_size)?);
    classic_eocd[16..20]
        .copy_from_slice(&zip32_u32_field(original_central_offset, central_offset)?);
    Ok(PreparedZip64Tail {
        eocd: zip64_eocd,
        locator,
        classic_eocd,
    })
}

fn zip32_u16_field(original: u16, value: u64) -> Result<[u8; 2], Error> {
    if original == u16::MAX || value >= u64::from(u16::MAX) {
        Ok(u16::MAX.to_le_bytes())
    } else {
        Ok(u16::try_from(value)
            .map_err(|_| unsupported("ZIP64 output promotion"))?
            .to_le_bytes())
    }
}

fn zip32_u32_field(original: u32, value: u64) -> Result<[u8; 4], Error> {
    if original == u32::MAX || value >= u64::from(u32::MAX) {
        Ok(u32::MAX.to_le_bytes())
    } else {
        Ok(u32::try_from(value)
            .map_err(|_| unsupported("ZIP64 output promotion"))?
            .to_le_bytes())
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
            Self::Shared { bytes, range } => &bytes[range.start..range.end],
        }
    }

    fn checked_bytes<'a>(&'a self, entries: &'a [PreservedEntry]) -> Result<&'a [u8], Error> {
        match self {
            Self::Copy(index) => entries
                .get(*index)
                .map(|entry| entry.central_bytes.as_slice())
                .ok_or_else(|| unsupported("invalid prepared central record")),
            Self::Generated(bytes) => Ok(bytes),
            Self::Shared { bytes, range } => bytes
                .get(range.clone())
                .ok_or_else(|| unsupported("invalid prepared central range")),
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
    fn len(&self) -> Result<u64, Error> {
        match self {
            Self::Copy(range) => range
                .end
                .checked_sub(range.start)
                .ok_or_else(|| unsupported("invalid prepared local range")),
            Self::Generated(bytes) => usize_to_u64(bytes.len(), "generated local bytes"),
            Self::Shared { bytes, range } => {
                if bytes.get(range.clone()).is_none() {
                    return Err(unsupported("invalid prepared local range"));
                }
                usize_to_u64(
                    range
                        .end
                        .checked_sub(range.start)
                        .ok_or_else(|| unsupported("invalid prepared local range"))?,
                    "prepared local range",
                )
            },
        }
    }
}

fn validate_prepared_local(entry: &PreparedEntry, source_end: u64) -> Result<(), Error> {
    match (&entry.local, &entry.generated_payload) {
        (PreparedLocal::Copy(range), None) => {
            if range.end < range.start || range.end > source_end {
                return Err(unsupported("invalid prepared source range"));
            }
            Ok(())
        },
        (PreparedLocal::Generated(_), None) => Ok(()),
        (
            PreparedLocal::Shared { bytes, range },
            Some(GeneratedPayload {
                range: payload_range,
                ..
            }),
        ) => {
            if bytes.get(range.clone()).is_none()
                || bytes.get(payload_range.clone()).is_none()
                || payload_range.start < range.start
                || payload_range.end > range.end
                || payload_range.start > payload_range.end
            {
                return Err(unsupported("generated payload outside local member"));
            }
            Ok(())
        },
        (PreparedLocal::Shared { bytes, range }, None) => {
            if bytes.get(range.clone()).is_none() {
                return Err(unsupported("invalid prepared local range"));
            }
            Ok(())
        },
        (PreparedLocal::Copy(_), Some(_)) => Err(unsupported("generated payload on copied member")),
        (PreparedLocal::Generated(_), Some(_)) => {
            Err(unsupported("generated payload metadata on owned local"))
        },
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
    archive: &ZipArchive<R>,
    entry: &PreservedEntry,
    local_end: u64,
) -> Result<bool, Error> {
    let (layout, local_central_name_mismatch) = archive
        .validate_preservation_entry_layout(entry.wayfinder, entry.raw_name_bytes())
        .map_err(map_local_layout_error)?;
    if layout.local_header_offset != entry.local_span.start {
        return Err(unsupported("local header offset mismatch"));
    }
    if layout.span_end > local_end {
        return Err(unsupported("truncated or overlapping local member"));
    }
    Ok(local_central_name_mismatch)
}

fn map_local_layout_error(error: Error) -> Error {
    match error.kind() {
        ErrorKind::Allocation { .. }
        | ErrorKind::IO(_)
        | ErrorKind::Io(_)
        | ErrorKind::LimitExceeded { .. } => error,
        _ => unsupported("invalid local member framing"),
    }
}

fn write_prepared_local<R, W>(
    local: &PreparedLocal,
    generated_payload: Option<&GeneratedPayload>,
    source: &R,
    sink: &mut W,
    buffer: &mut [u8],
    accounting: &mut ZipOperationAccounting,
) -> Result<(), Error>
where
    R: ReaderAt,
    W: Write,
{
    match (local, generated_payload) {
        (PreparedLocal::Copy(range), None) => {
            copy_range(source, range, sink, buffer, accounting)?;
            Ok(())
        },
        (PreparedLocal::Generated(bytes), None) => {
            sink.write_all(bytes)?;
            Ok(())
        },
        (PreparedLocal::Shared { bytes, range }, Some(payload)) => {
            let local_bytes = &bytes[range.start..range.end];
            let payload_start = payload
                .range
                .start
                .checked_sub(range.start)
                .expect("preflighted generated payload range is ordered");
            let payload_end = payload
                .range
                .end
                .checked_sub(range.start)
                .expect("preflighted generated payload range is ordered");
            sink.write_all(&local_bytes[..payload_start])?;
            write_all_counted(
                sink,
                &local_bytes[payload_start..payload_end],
                accounting,
                payload.kind,
            )?;
            sink.write_all(&local_bytes[payload_end..])?;
            Ok(())
        },
        (PreparedLocal::Shared { bytes, range }, None) => {
            let local_bytes = &bytes[range.start..range.end];
            sink.write_all(local_bytes)?;
            Ok(())
        },
        (PreparedLocal::Copy(_), Some(_)) | (PreparedLocal::Generated(_), Some(_)) => {
            unreachable!("invalid prepared local state was rejected during preflight")
        },
    }
}

fn copy_range<R, W>(
    source: &R,
    range: &Range<u64>,
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
        let remaining = range
            .end
            .checked_sub(offset)
            .expect("preflighted copy range is ordered");
        let buffer_len = u64::try_from(buffer.len()).expect("copy buffer length fits in u64");
        let len =
            usize::try_from(remaining.min(buffer_len)).expect("copy chunk length fits in usize");
        source.read_exact_at(&mut buffer[..len], offset)?;
        write_all_counted(
            sink,
            &buffer[..len],
            accounting,
            AccountingWriteKind::RawUnchangedSource,
        )?;
        offset = offset
            .checked_add(u64::try_from(len).expect("copy chunk length fits in u64"))
            .expect("preflighted copy range offset fits in u64");
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
    head_eocd_offset: u64,
) -> Result<(u64, usize, u64), Error> {
    let central_offset = record.central_directory_offset();
    let central_len = ZipFileHeaderFixed::SIZE
        .checked_add(
            usize::try_from(record.metadata_size_hint())
                .map_err(|_| unsupported("central-directory record length"))?,
        )
        .ok_or_else(|| unsupported("central-directory record length"))?;
    let central_len_u64 = usize_to_u64(central_len, "central-directory record length")?;
    let central_end = central_offset
        .checked_add(central_len_u64)
        .ok_or_else(|| unsupported("central-directory record range"))?;
    if central_end > head_eocd_offset {
        return Err(unsupported("truncated central-directory record"));
    }
    Ok((central_offset, central_len, central_end))
}

fn central_local_header_offset_patch(
    record: &crate::ZipFileHeaderRecord<'_>,
    central_fixed: &ZipFileHeaderFixed,
    central_bytes: &[u8],
    central_offset: u64,
    central_end: u64,
) -> Result<CentralLocalHeaderOffsetPatch, Error> {
    if central_fixed.local_header_offset != u32::MAX
        && u64::from(central_fixed.local_header_offset) != record.local_header_offset()
    {
        return Err(unsupported("local-header offset value mismatch"));
    }
    if let Some(range) = record.zip64_local_header_offset_range() {
        if central_fixed.local_header_offset != u32::MAX
            || range.end.checked_sub(range.start) != Some(8)
            || range.start < central_offset
            || range.end > central_end
        {
            return Err(unsupported("invalid ZIP64 local-header offset range"));
        }
        let start_offset = range
            .start
            .checked_sub(central_offset)
            .ok_or_else(|| unsupported("ZIP64 local-header offset range"))?;
        let end_offset = range
            .end
            .checked_sub(central_offset)
            .ok_or_else(|| unsupported("ZIP64 local-header offset range"))?;
        let start = usize::try_from(start_offset)
            .map_err(|_| unsupported("ZIP64 local-header offset range"))?;
        let end = usize::try_from(end_offset)
            .map_err(|_| unsupported("ZIP64 local-header offset range"))?;
        let expected = record.local_header_offset().to_le_bytes();
        if central_bytes.get(start..end) != Some(expected.as_slice()) {
            return Err(unsupported("ZIP64 local-header offset value mismatch"));
        }
        return Ok(CentralLocalHeaderOffsetPatch::Zip64(start..end));
    }

    if central_fixed.local_header_offset == u32::MAX {
        return Err(unsupported("missing ZIP64 local-header offset range"));
    }
    Ok(CentralLocalHeaderOffsetPatch::Fixed32)
}

fn valid_classic_disk_fields(eocd: &EndOfCentralDirectoryRecordFixed, zip64: bool) -> bool {
    if !zip64 {
        return eocd.disk_number == 0 && eocd.eocd_disk == 0;
    }
    (eocd.disk_number == 0 || eocd.disk_number == u16::MAX)
        && (eocd.eocd_disk == 0 || eocd.eocd_disk == u16::MAX)
}

fn valid_classic_zip64_fields(
    eocd: &EndOfCentralDirectoryRecordFixed,
    entries: u64,
    central_size: u64,
    central_offset: u64,
) -> bool {
    (eocd.num_entries == u16::MAX || u64::from(eocd.num_entries) == entries)
        && (eocd.total_entries == u16::MAX || u64::from(eocd.total_entries) == entries)
        && (eocd.central_dir_size == u32::MAX || u64::from(eocd.central_dir_size) == central_size)
        && (eocd.central_dir_offset == u32::MAX
            || u64::from(eocd.central_dir_offset) == central_offset)
}

fn read_zip64_tail<R: ReaderAt>(
    source: &R,
    head_eocd_offset: u64,
    eocd_offset: u64,
    entries: u64,
    central_size: u64,
    central_offset: u64,
    max_extensible_data: u64,
) -> Result<Zip64Tail, Error> {
    let locator_size = usize_to_u64(ZIP64_LOCATOR_SIZE, "ZIP64 locator size")?;
    let locator_offset = eocd_offset
        .checked_sub(locator_size)
        .ok_or_else(|| unsupported("ZIP64 locator offset"))?;
    if head_eocd_offset > locator_offset {
        return Err(unsupported("ZIP64 EOCD and locator bounds"));
    }

    let fixed = read_vec(source, head_eocd_offset, ZIP64_EOCD_FIXED_SIZE)?;
    let record = crate::Zip64EndOfCentralDirectoryRecord::parse(&fixed)?;
    if record.size < ZIP64_EOCD_FIXED_PAYLOAD_SIZE
        || record.disk_number != 0
        || record.cd_disk != 0
        || record.num_entries != entries
        || record.total_entries != entries
        || record.central_dir_size != central_size
        || record.central_dir_offset != central_offset
    {
        return Err(unsupported("invalid ZIP64 EOCD metadata"));
    }
    let extensible_data_len = record
        .size
        .checked_sub(ZIP64_EOCD_FIXED_PAYLOAD_SIZE)
        .ok_or_else(|| unsupported("ZIP64 EOCD extensible data"))?;
    if extensible_data_len > max_extensible_data {
        return Err(limit_error(
            LimitResource::MetadataBytes,
            extensible_data_len,
            max_extensible_data,
        ));
    }
    let record_len = 12usize
        .checked_add(
            usize::try_from(record.size).map_err(|_| unsupported("ZIP64 EOCD record length"))?,
        )
        .ok_or_else(|| unsupported("ZIP64 EOCD record length"))?;
    let record_len_u64 = usize_to_u64(record_len, "ZIP64 EOCD record length")?;
    let record_end = head_eocd_offset
        .checked_add(record_len_u64)
        .ok_or_else(|| unsupported("ZIP64 EOCD record range"))?;
    if record_end != locator_offset {
        return Err(unsupported("ZIP64 EOCD is not adjacent to its locator"));
    }
    let eocd = read_vec(source, head_eocd_offset, record_len)?;

    let locator = read_vec(source, locator_offset, ZIP64_LOCATOR_SIZE)?;
    let locator_signature = u32::from_le_bytes(
        locator
            .get(0..4)
            .ok_or_else(|| unsupported("invalid ZIP64 locator metadata"))?
            .try_into()
            .map_err(|_| unsupported("invalid ZIP64 locator metadata"))?,
    );
    let locator_disk = u32::from_le_bytes(
        locator
            .get(4..8)
            .ok_or_else(|| unsupported("invalid ZIP64 locator metadata"))?
            .try_into()
            .map_err(|_| unsupported("invalid ZIP64 locator metadata"))?,
    );
    let locator_eocd_offset = u64::from_le_bytes(
        locator
            .get(8..16)
            .ok_or_else(|| unsupported("invalid ZIP64 locator metadata"))?
            .try_into()
            .map_err(|_| unsupported("invalid ZIP64 locator metadata"))?,
    );
    let locator_disk_count = u32::from_le_bytes(
        locator
            .get(16..20)
            .ok_or_else(|| unsupported("invalid ZIP64 locator metadata"))?
            .try_into()
            .map_err(|_| unsupported("invalid ZIP64 locator metadata"))?,
    );
    if locator_signature != ZIP64_LOCATOR_SIGNATURE
        || locator_disk != 0
        || locator_eocd_offset != head_eocd_offset
        || locator_disk_count != 1
    {
        return Err(unsupported("invalid ZIP64 locator metadata"));
    }

    let classic_eocd = read_vec(source, eocd_offset, EndOfCentralDirectoryRecordFixed::SIZE)?;
    Ok(Zip64Tail {
        eocd,
        locator,
        classic_eocd,
    })
}

fn map_metadata_limit_error(error: Error, metadata_prefix: u64, limits: ArchiveLimits) -> Error {
    if let ErrorKind::LimitExceeded {
        resource: LimitResource::MetadataBytes,
        actual,
        ..
    } = error.kind()
    {
        let actual = metadata_prefix.checked_add(*actual).unwrap_or(u64::MAX);
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
    use crate::ZipLocalFileHeaderFixed;
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
        let mut records = archive.entries();
        records.next().unwrap().unwrap();
        let descriptor_record = records.next().unwrap().unwrap();
        let next_local = records.next().unwrap().unwrap().local_header_offset() as usize;
        let descriptor = usize::try_from(
            archive
                .get_entry(descriptor_record.wayfinder())
                .unwrap()
                .compressed_data_range()
                .1,
        )
        .unwrap();
        assert_eq!(next_local - descriptor, 16);
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
        include_bytes!("../assets/zip64.zip").to_vec()
    }

    fn prefixed_empty_archive() -> Vec<u8> {
        let mut data = b"empty archive prelude".to_vec();
        let empty = ZipArchiveWriter::new(Vec::new()).finish().unwrap();
        data.extend_from_slice(&empty);
        data
    }

    fn two_stored_archive() -> Vec<u8> {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        writer
            .write_stored_file("first.bin", b"first payload")
            .unwrap();
        writer
            .write_stored_file("second.bin", b"second payload")
            .unwrap();
        writer.finish().unwrap()
    }

    fn promote_to_zip64(mut source: Vec<u8>) -> Vec<u8> {
        let archive = ZipArchive::from_slice(&source).unwrap();
        assert!(!archive.is_zip64());
        let central_start = usize::try_from(archive.directory_offset()).unwrap();
        let old_eocd = usize::try_from(archive.eocd_offset()).unwrap();
        let original_classic_eocd = source[old_eocd..].to_vec();
        let mut central = Vec::new();
        let mut offset = central_start;
        while offset < old_eocd {
            let fixed = ZipFileHeaderFixed::parse(&source[offset..]).unwrap();
            let record_len = ZipFileHeaderFixed::SIZE + fixed.variable_length();
            let record = &source[offset..offset + record_len];
            let name_len = usize::from(fixed.file_name_len);
            let old_extra_len = usize::from(fixed.extra_field_len);
            let variable_start = ZipFileHeaderFixed::SIZE + name_len;
            let old_extra_end = variable_start + old_extra_len;
            let mut promoted = record[..variable_start].to_vec();
            promoted[20..24].copy_from_slice(&u32::MAX.to_le_bytes());
            promoted[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
            promoted[42..46].copy_from_slice(&u32::MAX.to_le_bytes());
            let mut zip64_extra = Vec::with_capacity(28);
            zip64_extra.extend_from_slice(&1u16.to_le_bytes());
            zip64_extra.extend_from_slice(&24u16.to_le_bytes());
            zip64_extra.extend_from_slice(&u64::from(fixed.uncompressed_size).to_le_bytes());
            zip64_extra.extend_from_slice(&u64::from(fixed.compressed_size).to_le_bytes());
            zip64_extra.extend_from_slice(&u64::from(fixed.local_header_offset).to_le_bytes());
            let preceding_extra = [0xaa, 0xaa, 0x03, 0x00, 0x70, 0x72, 0x65];
            promoted[30..32].copy_from_slice(
                &(u16::try_from(old_extra_len + preceding_extra.len() + zip64_extra.len())
                    .unwrap())
                .to_le_bytes(),
            );
            promoted.extend_from_slice(&preceding_extra);
            promoted.extend_from_slice(&zip64_extra);
            promoted.extend_from_slice(&record[old_extra_end..]);
            central.extend_from_slice(&promoted);
            offset += record_len;
        }

        source.truncate(central_start);
        source.extend_from_slice(&central);
        let zip64_offset = source.len();
        let mut zip64_eocd = Vec::new();
        zip64_eocd.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        zip64_eocd.extend_from_slice(&47u64.to_le_bytes());
        zip64_eocd.extend_from_slice(&45u16.to_le_bytes());
        zip64_eocd.extend_from_slice(&45u16.to_le_bytes());
        zip64_eocd.extend_from_slice(&0u32.to_le_bytes());
        zip64_eocd.extend_from_slice(&0u32.to_le_bytes());
        let entry_count = u64::from(u16::from_le_bytes(
            original_classic_eocd[8..10].try_into().unwrap(),
        ));
        zip64_eocd.extend_from_slice(&entry_count.to_le_bytes());
        zip64_eocd.extend_from_slice(&entry_count.to_le_bytes());
        zip64_eocd.extend_from_slice(&(central.len() as u64).to_le_bytes());
        zip64_eocd.extend_from_slice(&(central_start as u64).to_le_bytes());
        zip64_eocd.extend_from_slice(&[0xa1, 0xb2, 0xc3]);
        source.extend_from_slice(&zip64_eocd);
        source.extend_from_slice(&0x0706_4b50u32.to_le_bytes());
        source.extend_from_slice(&0u32.to_le_bytes());
        source.extend_from_slice(&(zip64_offset as u64).to_le_bytes());
        source.extend_from_slice(&1u32.to_le_bytes());
        let classic_offset = source.len();
        let mut classic = original_classic_eocd;
        classic[8..10].copy_from_slice(&u16::MAX.to_le_bytes());
        classic[10..12].copy_from_slice(&u16::MAX.to_le_bytes());
        classic[12..16].copy_from_slice(&u32::MAX.to_le_bytes());
        classic[16..20].copy_from_slice(&u32::MAX.to_le_bytes());
        source.extend_from_slice(&classic);
        assert_eq!(classic_offset + classic.len(), source.len());
        source
    }

    fn zip64_tail_extensible_data(data: &[u8]) -> Vec<u8> {
        let archive = ZipArchive::from_slice(data).unwrap();
        let start = usize::try_from(archive.head_eocd_offset()).unwrap() + ZIP64_EOCD_FIXED_SIZE;
        let end = usize::try_from(archive.eocd_offset()).unwrap() - ZIP64_LOCATOR_SIZE;
        data[start..end].to_vec()
    }

    fn ambiguous_zip64_tail() -> Vec<u8> {
        let mut data = vec![0u8; ZipFileHeaderFixed::SIZE];
        data[..4].copy_from_slice(&0x0201_4b50u32.to_le_bytes());
        let record_start = data.len();
        let mut extensible = vec![0xc7; 180];
        let mut false_record = Vec::new();
        false_record.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        false_record.extend_from_slice(&44u64.to_le_bytes());
        false_record.extend_from_slice(&45u16.to_le_bytes());
        false_record.extend_from_slice(&45u16.to_le_bytes());
        false_record.extend_from_slice(&[0; 8]);
        false_record.extend_from_slice(&1u64.to_le_bytes());
        false_record.extend_from_slice(&1u64.to_le_bytes());
        false_record.extend_from_slice(&(ZipFileHeaderFixed::SIZE as u64).to_le_bytes());
        false_record.extend_from_slice(&0u64.to_le_bytes());
        extensible[..false_record.len()].copy_from_slice(&false_record);
        let mut second_record = false_record.clone();
        second_record[4..12].copy_from_slice(&104u64.to_le_bytes());
        let second_start = 64usize;
        extensible[second_start..second_start + second_record.len()]
            .copy_from_slice(&second_record);
        data.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        data.extend_from_slice(&224u64.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&45u16.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&1u64.to_le_bytes());
        data.extend_from_slice(&1u64.to_le_bytes());
        data.extend_from_slice(&(ZipFileHeaderFixed::SIZE as u64).to_le_bytes());
        data.extend_from_slice(&0u64.to_le_bytes());
        data.extend_from_slice(&extensible);
        let false_offset = record_start + ZIP64_EOCD_FIXED_SIZE;
        data.extend_from_slice(&0x0706_4b50u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&(false_offset as u64).to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        assert_eq!(
            record_start + ZIP64_EOCD_FIXED_SIZE + second_start + 12 + 104,
            data.len() - 22 - ZIP64_LOCATOR_SIZE
        );
        data
    }

    fn masked_central_record(data: &[u8], entry: &PreservedEntry) -> Vec<u8> {
        let range = entry.central_record();
        let mut record = data[range.start as usize..range.end as usize].to_vec();
        record[CENTRAL_LOCAL_HEADER_OFFSET].fill(0);
        if let Some(range) = ZipArchive::from_slice(data)
            .unwrap()
            .entries()
            .find_map(|candidate| {
                let candidate = candidate.unwrap();
                (candidate.central_directory_offset() == entry.central_record.start)
                    .then(|| candidate.zip64_local_header_offset_range())
            })
            .flatten()
        {
            let start = usize::try_from(range.start - entry.central_record.start).unwrap();
            let end = usize::try_from(range.end - entry.central_record.start).unwrap();
            record[start..end].fill(0);
        }
        record
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

    fn stored_descriptor_archive(payload: &[u8]) -> Vec<u8> {
        let mut writer = ZipArchiveWriter::new(Vec::new());
        let (mut file, config) = writer
            .new_file("payload.bin")
            .compression_method(CompressionMethod::Store)
            .start()
            .unwrap();
        let mut data_writer = config.wrap(&mut file);
        data_writer.write_all(payload).unwrap();
        let (_, descriptor) = data_writer.finish().unwrap();
        file.finish(descriptor).unwrap();
        writer.finish().unwrap()
    }

    fn descriptor_offsets(data: &[u8]) -> (usize, usize, usize) {
        let archive = ZipArchive::from_slice(data).unwrap();
        let central = usize::try_from(archive.directory_offset()).unwrap();
        let eocd = usize::try_from(archive.eocd_offset()).unwrap();
        let record = archive
            .entries()
            .find(|record| {
                record
                    .as_ref()
                    .is_ok_and(|record| record.has_data_descriptor())
            })
            .unwrap()
            .unwrap();
        let payload_end = usize::try_from(
            archive
                .get_entry(record.wayfinder())
                .unwrap()
                .compressed_data_range()
                .1,
        )
        .unwrap();
        (payload_end, central, eocd)
    }

    fn unsigned_descriptor_archive(mut data: Vec<u8>) -> Vec<u8> {
        let (payload_end, central, eocd) = descriptor_offsets(&data);
        assert_eq!(central - payload_end, 16);
        data.drain(payload_end..payload_end + 4);
        let central = central - 4;
        let eocd = eocd - 4;
        data[eocd + 16..eocd + 20].copy_from_slice(&u32::try_from(central).unwrap().to_le_bytes());
        data
    }

    fn zip64_descriptor_archive(payload: &[u8], signed: bool) -> Vec<u8> {
        let mut data = stored_descriptor_archive(payload);
        let (payload_end, mut central, mut eocd) = descriptor_offsets(&data);
        if !signed {
            data.drain(payload_end..payload_end + 4);
            central -= 4;
            eocd -= 4;
            data[eocd + 16..eocd + 20]
                .copy_from_slice(&u32::try_from(central).unwrap().to_le_bytes());
        }

        let descriptor_len = if signed { 16 } else { 12 };
        let crc_offset = payload_end + if signed { 4 } else { 0 };
        let compressed_offset = crc_offset + 4;
        let uncompressed_offset = compressed_offset + 4;
        let crc = u32::from_le_bytes(data[crc_offset..crc_offset + 4].try_into().unwrap());
        let compressed = u32::from_le_bytes(
            data[compressed_offset..compressed_offset + 4]
                .try_into()
                .unwrap(),
        );
        let uncompressed = u32::from_le_bytes(
            data[uncompressed_offset..uncompressed_offset + 4]
                .try_into()
                .unwrap(),
        );
        let mut descriptor = Vec::new();
        if signed {
            descriptor.extend_from_slice(&crate::DataDescriptor::SIGNATURE.to_le_bytes());
        }
        descriptor.extend_from_slice(&crc.to_le_bytes());
        descriptor.extend_from_slice(&u64::from(compressed).to_le_bytes());
        descriptor.extend_from_slice(&u64::from(uncompressed).to_le_bytes());
        assert_eq!(descriptor.len(), descriptor_len + 8);
        data.splice(
            payload_end..payload_end + descriptor_len,
            descriptor.iter().copied(),
        );
        central += 8;
        eocd += 8;

        let fixed = ZipFileHeaderFixed::parse(&data[central..]).unwrap();
        let record_len = ZipFileHeaderFixed::SIZE + fixed.variable_length();
        let record = &data[central..central + record_len];
        let name_len = usize::from(fixed.file_name_len);
        let old_extra_len = usize::from(fixed.extra_field_len);
        let variable_start = ZipFileHeaderFixed::SIZE + name_len;
        let old_extra_end = variable_start + old_extra_len;
        let mut promoted = record[..variable_start].to_vec();
        promoted[20..24].copy_from_slice(&u32::MAX.to_le_bytes());
        promoted[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
        let mut zip64_extra = Vec::new();
        zip64_extra.extend_from_slice(&1u16.to_le_bytes());
        zip64_extra.extend_from_slice(&16u16.to_le_bytes());
        zip64_extra.extend_from_slice(&u64::from(uncompressed).to_le_bytes());
        zip64_extra.extend_from_slice(&u64::from(compressed).to_le_bytes());
        promoted[30..32].copy_from_slice(
            &u16::try_from(old_extra_len + zip64_extra.len())
                .unwrap()
                .to_le_bytes(),
        );
        promoted.extend_from_slice(&record[variable_start..old_extra_end]);
        promoted.extend_from_slice(&zip64_extra);
        promoted.extend_from_slice(&record[old_extra_end..]);
        data.splice(central..eocd, promoted.iter().copied());
        eocd += 20;
        data[eocd + 12..eocd + 16]
            .copy_from_slice(&u32::try_from(promoted.len()).unwrap().to_le_bytes());
        data[eocd + 16..eocd + 20].copy_from_slice(&u32::try_from(central).unwrap().to_le_bytes());
        data
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
    fn validates_signed_and_unsigned_32_and_64_descriptors_at_payload_end_with_decoy() {
        let payload = [
            crate::DataDescriptor::SIGNATURE.to_le_bytes().as_slice(),
            b"payload bytes".as_slice(),
        ]
        .concat();
        for (data, zip64) in [
            (stored_descriptor_archive(&payload), false),
            (
                unsigned_descriptor_archive(stored_descriptor_archive(&payload)),
                false,
            ),
            (zip64_descriptor_archive(&payload, true), true),
            (zip64_descriptor_archive(&payload, false), true),
        ] {
            let (archive, mut buffer) = indexed(&data);
            assert_eq!(archive.is_zip64(), false);
            let index = if zip64 {
                PreservationIndex::new_with_policy(
                    &archive,
                    &mut buffer,
                    PreservationPolicy::AllowZip64,
                )
                .unwrap()
            } else {
                PreservationIndex::new(&archive, &mut buffer).unwrap()
            };
            assert_eq!(
                index
                    .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                    .unwrap(),
                data
            );
        }
    }

    #[test]
    fn default_policy_refuses_projected_zip64_descriptor_before_publication() {
        let payload = [
            crate::DataDescriptor::SIGNATURE.to_le_bytes().as_slice(),
            b"payload bytes".as_slice(),
        ]
        .concat();
        for signed in [true, false] {
            let data = zip64_descriptor_archive(&payload, signed);
            let (archive, mut buffer) = indexed(&data);
            assert!(!archive.is_zip64());
            let sink = b"untouched".to_vec();
            let error = match PreservationIndex::new(&archive, &mut buffer) {
                Ok(_) => panic!("default policy must refuse projected ZIP64 entry"),
                Err(error) => error,
            };
            assert!(matches!(
                error.kind(),
                ErrorKind::UnsupportedPreservation { .. }
            ));
            assert_eq!(sink, b"untouched");
        }
    }

    #[test]
    fn rejects_malformed_descriptor_at_exact_payload_end_before_writing() {
        let mut data = stored_descriptor_archive(b"payload");
        let (payload_end, _, _) = descriptor_offsets(&data);
        data[payload_end + 8..payload_end + 12].copy_from_slice(&999u32.to_le_bytes());

        let (archive, mut buffer) = indexed(&data);
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("malformed data descriptor must be rejected"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
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
    fn fallible_plan_reservations_and_preparation_errors_leave_sink_untouched() {
        let data = ordinary_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();

        let mut plan = PreservationPlan::new();
        assert!(plan.try_reserve_exact(usize::MAX).is_err());
        assert!(plan.try_reserve_appended(usize::MAX).is_err());

        let mut sink = b"untouched".to_vec();
        let error = index
            .write_to(&plan, &mut sink)
            .expect_err("preparation must reject incomplete coverage before output");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(sink, b"untouched");
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
    fn rejects_fixed32_offsets_after_zip64_output_grows_past_their_range() {
        let source = promote_to_zip64(two_stored_archive());
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();

        // A Range is enough to model a layout past the ZIP32 limit; no large
        // allocation or write is needed.  The first source record retains an
        // existing ZIP64 offset field, while the second generated record
        // would still need a fixed-width central offset.
        let mut prepared = vec![
            PreparedEntry {
                local: PreparedLocal::Copy(0..u64::from(u32::MAX) + 1),
                central: PreparedCentral::Copy(0),
                generated_payload: None,
                omitted: false,
            },
            PreparedEntry {
                local: PreparedLocal::Generated(Vec::new()),
                central: PreparedCentral::Generated(vec![0; ZipFileHeaderFixed::SIZE]),
                generated_payload: None,
                omitted: false,
            },
        ];
        let error = index
            .validate_output_layout(&prepared)
            .expect_err("fixed-width offsets cannot represent the ZIP64 layout");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        prepared.clear();
    }

    #[test]
    fn rejects_reader_valid_zip64_local_and_disk_malformed_sources() {
        let source = promote_to_zip64(two_stored_archive());
        let source_archive = ZipArchive::from_slice(&source).unwrap();
        let first_local = source_archive
            .entries()
            .next_entry()
            .unwrap()
            .unwrap()
            .local_header_offset() as usize;
        let central_start = source_archive.directory_offset() as usize;
        let first_central = source_archive
            .entries()
            .next_entry()
            .unwrap()
            .unwrap()
            .central_directory_offset() as usize;

        let mut malformed_local = source.clone();
        malformed_local[first_local + 18..first_local + 22].copy_from_slice(&1000u32.to_le_bytes());
        let (archive, mut buffer) = indexed(&malformed_local);
        let error = match PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        ) {
            Ok(_) => panic!("mismatched local sizes must be rejected"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));

        let mut malformed_disk = source.clone();
        malformed_disk[first_central + 34..first_central + 36].copy_from_slice(&1u16.to_le_bytes());
        let (archive, mut buffer) = indexed(&malformed_disk);
        let error = match PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        ) {
            Ok(_) => panic!("nonzero disk-start metadata must be rejected"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(central_start, source_archive.directory_offset() as usize);
        assert_eq!(source, promote_to_zip64(two_stored_archive()));
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
    fn default_policy_refuses_zip64_while_internal_policy_preserves_it() {
        let data = zip64_archive();
        let (archive, mut buffer) = indexed(&data);
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("public preservation must refuse ZIP64"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));

        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        assert_eq!(
            index
                .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                .unwrap(),
            data
        );
    }

    #[test]
    fn empty_archive_prefix_is_refused_before_publication() {
        let data = prefixed_empty_archive();
        let (archive, mut buffer) = indexed(&data);
        assert_eq!(archive.entries_hint(), 0);
        let error = match PreservationIndex::new(&archive, &mut buffer) {
            Ok(_) => panic!("prefixed empty archive must be refused"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));

        let data = ZipArchiveWriter::new(Vec::new()).finish().unwrap();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        assert_eq!(
            index
                .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                .unwrap(),
            data
        );
    }

    #[test]
    fn zip64_metadata_limit_reports_aggregate_comment_and_extension() {
        let data = with_comment(promote_to_zip64(two_stored_archive()), b"abcd");
        let (archive, mut buffer) = indexed(&data);
        let mut limits = ArchiveLimits::UNBOUNDED;
        limits.max_metadata_bytes = 6;
        let error = match PreservationIndex::new_with_limits_and_policy(
            &archive,
            &mut buffer,
            limits,
            PreservationPolicy::AllowZip64,
        ) {
            Ok(_) => panic!("metadata limit must reject the ZIP64 tail"),
            Err(error) => error,
        };
        assert!(matches!(
            error.kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::MetadataBytes,
                actual: 7,
                maximum: 6,
            }
        ));
    }

    #[test]
    fn non_sentinel_classic_zip64_fields_must_match_resolved_values() {
        let source = zip64_archive();
        let source_archive = ZipArchive::from_slice(&source).unwrap();
        let eocd = usize::try_from(source_archive.eocd_offset()).unwrap();
        let cases: &[(usize, &[u8])] = &[
            (8, &2u16.to_le_bytes()),
            (10, &2u16.to_le_bytes()),
            (12, &73u32.to_le_bytes()),
            (16, &1u32.to_le_bytes()),
        ];

        for &(field_offset, replacement) in cases {
            let mut malformed = source.clone();
            malformed[eocd + field_offset..eocd + field_offset + replacement.len()]
                .copy_from_slice(replacement);
            let (archive, mut buffer) = indexed(&malformed);
            let error = match PreservationIndex::new_with_policy(
                &archive,
                &mut buffer,
                PreservationPolicy::AllowZip64,
            ) {
                Ok(_) => panic!("classic ZIP64 field mismatch must be rejected"),
                Err(error) => error,
            };
            assert!(matches!(
                error.kind(),
                ErrorKind::UnsupportedPreservation { .. }
            ));
        }
    }

    #[test]
    fn copies_zip64_sources_byte_for_byte_including_tail() {
        let data = zip64_archive();
        let (archive, mut buffer) = indexed(&data);
        assert!(archive.is_zip64());
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        let output = index
            .write_to(&PreservationPlan::copy_all(&index), Vec::new())
            .unwrap();
        assert_eq!(output, data);
    }

    #[test]
    fn copies_zip64_central_extras_and_tail_extensible_data_byte_for_byte() {
        for data in [
            zip64_archive(),
            include_bytes!("../assets/zip64-2.zip").to_vec(),
        ] {
            let source_tail_extension = zip64_tail_extensible_data(&data);
            let (archive, mut buffer) = indexed(&data);
            let index = PreservationIndex::new_with_policy(
                &archive,
                &mut buffer,
                PreservationPolicy::AllowZip64,
            )
            .unwrap();
            let output = index
                .write_to(&PreservationPlan::copy_all(&index), Vec::new())
                .unwrap();
            assert_eq!(output, data);
            assert_eq!(zip64_tail_extensible_data(&output), source_tail_extension);
        }
    }

    #[test]
    fn shifted_zip64_local_offsets_patch_only_existing_zip64_values() {
        let source = promote_to_zip64(two_stored_archive());
        let (archive, mut buffer) = indexed(&source);
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        let omitted_id = index.entries()[0].id();
        let retained = index.entries()[1].clone();
        let source_local = retained.local_span();
        let source_central = masked_central_record(&source, &retained);

        let mut plan = PreservationPlan::copy_all(&index);
        plan.actions[0] = PreservationAction::Omit(omitted_id);
        let output = index.write_to(&plan, Vec::new()).unwrap();
        let output_archive = ZipArchive::from_slice(&output).unwrap();
        assert!(output_archive.is_zip64());
        let output_entry = output_archive
            .entries()
            .map(|entry| entry.unwrap())
            .next()
            .unwrap();
        assert_eq!(output_entry.local_header_offset(), 0);
        assert_eq!(
            &output[..(source_local.end - source_local.start) as usize],
            &source[source_local.start as usize..source_local.end as usize]
        );

        let (output_archive, mut output_buffer) = indexed(&output);
        let output_index = PreservationIndex::new_with_policy(
            &output_archive,
            &mut output_buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        assert_eq!(
            masked_central_record(&output, &output_index.entries()[0]),
            source_central
        );
        assert_eq!(zip64_tail_extensible_data(&output), vec![0xa1, 0xb2, 0xc3]);
    }

    #[test]
    fn regenerating_a_zip64_member_keeps_zip64_framing() {
        let data = promote_to_zip64(two_stored_archive());
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        let id = index.entries()[0].id();
        let mut plan = PreservationPlan::copy_all(&index);
        plan.actions[0] = PreservationAction::Regenerate {
            id,
            entry: RegeneratedEntry::new("replacement.bin", b"replacement".to_vec()),
        };

        let output = index.write_to(&plan, Vec::new()).unwrap();
        let output_archive = ZipArchive::from_slice(&output).unwrap();
        assert!(output_archive.is_zip64());
        assert_eq!(zip64_tail_extensible_data(&output), vec![0xa1, 0xb2, 0xc3]);
        assert_eq!(
            output_archive.comment().as_bytes().len(),
            archive.comment().remaining() as usize
        );
        assert_eq!(
            crate::office::ArchiveReader::new(&output)
                .unwrap()
                .read("replacement.bin")
                .unwrap(),
            b"replacement"
        );
    }

    #[test]
    fn zip64_incomplete_plan_leaves_sink_untouched() {
        let data = zip64_archive();
        let (archive, mut buffer) = indexed(&data);
        let index = PreservationIndex::new_with_policy(
            &archive,
            &mut buffer,
            PreservationPolicy::AllowZip64,
        )
        .unwrap();
        let mut sink = b"untouched".to_vec();
        let error = index
            .write_to(&PreservationPlan::new(), &mut sink)
            .expect_err("incomplete ZIP64 plan must fail before output");
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedPreservation { .. }
        ));
        assert_eq!(sink, b"untouched");
    }

    #[test]
    fn rejects_ambiguous_malformed_multidisk_and_truncated_zip64_sources() {
        assert!(ZipArchive::from_slice(ambiguous_zip64_tail()).is_err());

        let source = zip64_archive();
        let archive = ZipArchive::from_slice(&source).unwrap();
        let head = usize::try_from(archive.head_eocd_offset()).unwrap();
        let locator = usize::try_from(archive.eocd_offset()).unwrap() - ZIP64_LOCATOR_SIZE;

        let mut malformed = source.clone();
        malformed[head + 4..head + 12].copy_from_slice(&43u64.to_le_bytes());
        assert!(ZipArchive::from_slice(&malformed).is_err());

        let mut multidisk = source.clone();
        multidisk[locator + 4..locator + 8].copy_from_slice(&1u32.to_le_bytes());
        assert!(ZipArchive::from_slice(&multidisk).is_err());

        let mut truncated = source.clone();
        truncated.pop();
        assert!(ZipArchive::from_slice(&truncated).is_err());
        assert_eq!(source, zip64_archive());
    }
}
