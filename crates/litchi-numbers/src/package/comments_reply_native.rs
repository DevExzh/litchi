//! Strict private native transitions for direct Numbers comment replies.
//!
//! The public comments facade deliberately does not expose any of the types in
//! this module.  A request is made from already-resolved package routes and a
//! result contains only private decompressed-member candidates.  In
//! particular, this module never selects a sheet/table or publishes a ZIP
//! candidate.  Metadata ownership, ZIP reassembly, previews, and publication
//! remain the responsibility of the package owner.
//!
//! The first reply slice is intentionally conservative.  It accepts one
//! complete current IWA member.  A caller that discovers a split graph must
//! provide a future multi-member implementation; silently editing only the
//! representative member would make a reply/list/BNC ownership edge stale.

use std::{
    collections::{HashMap, HashSet},
    fmt,
    mem::size_of,
};

use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, patch_length_delimited_field,
    patch_varint_field,
};
use litchi_iwa_core::archive::{
    ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, FieldObjectReferenceTransition, FieldType, ObjectReferenceTransition,
};
use litchi_iwa_core::{Archive, ArchiveObject, Limits, RawMessage};
use litchi_iwa_protos::{
    comment_storage_codec, numbers_table_cell_storage_codec as storage_codec, tst,
};
use litchi_numbers_wire::BncCell;

use super::{
    comments::CommentReplyPath,
    table_cell_pop_up_menu::{Path as BudgetPath, TransactionBudget},
    table_cell_pop_up_menu_native as popup_native,
};

const BUDGET_PATH: BudgetPath = BudgetPath::Package;

const TABLE_MODEL_TYPES: [u32; 2] = [6_000, 6_001];
const TILE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_TYPES: [u32; 2] = [6_005, 6_201];
const COMMENT_STORAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_TYPE: u32 = 212;
const COMMENT_LIST_TYPE: i32 = tst::table_data_list::ListType::CommentStorage as i32;
const AUTHOR_FIELD_PATH: &[u32] = &[3];
const REPLY_FIELD_PREFIX: u32 = 4;
const LIST_ENTRY_FIELD_PREFIX: u32 = 3;

/// One current, decompressed IWA member supplied by the package owner.
///
/// The first implementation requires exactly one member.  Keeping the member
/// identity in the request makes the eventual split-member adapter explicit
/// instead of accidentally treating a ZIP name as a component identity.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativeReplyMember<'source> {
    pub(super) archive: &'source Archive,
    pub(super) component_index: usize,
    pub(super) member_name: &'source str,
}

/// A fully resolved object/message route inside one supplied member.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NativeReplyObjectRoute {
    pub(super) member_index: usize,
    pub(super) identifier: u64,
    pub(super) message_index: usize,
}

/// Fresh identifier and payload UUID reserved by the metadata owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct NativeReplyStorageIdentity {
    pub(super) identifier: u64,
    pub(super) uuid_lower: u64,
    pub(super) uuid_upper: u64,
}

/// The supported direct-reply lifecycle operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NativeReplyOperation<'text> {
    /// Clone the root and append one direct leaf reply.
    Append { text: &'text str },
    /// Clone the root and replace one source-ordered direct leaf reply.
    Replace {
        ordinal: usize,
        expected_reply_identifier: u64,
        text: &'text str,
    },
    /// Clone the root and remove one source-ordered direct leaf reply.
    Remove {
        ordinal: usize,
        expected_reply_identifier: u64,
    },
}

impl<'text> NativeReplyOperation<'text> {
    const fn ordinal(self) -> Option<usize> {
        match self {
            Self::Append { .. } => None,
            Self::Replace { ordinal, .. } | Self::Remove { ordinal, .. } => Some(ordinal),
        }
    }

    const fn expected_reply_identifier(self) -> Option<u64> {
        match self {
            Self::Append { .. } => None,
            Self::Replace {
                expected_reply_identifier,
                ..
            }
            | Self::Remove {
                expected_reply_identifier,
                ..
            } => Some(expected_reply_identifier),
        }
    }

    const fn text(self) -> Option<&'text str> {
        match self {
            Self::Append { text } | Self::Replace { text, .. } => Some(text),
            Self::Remove { .. } => None,
        }
    }
}

/// A resolved direct-reply native graph transition.
///
/// `members` must contain every current physical archive that can contain the
/// selected model, tile, list, root, or reply objects.  This revision rejects
/// more than one member rather than mutating a representative and publishing
/// stale cross-member references.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativeReplyRequest<'source> {
    pub(super) members: &'source [NativeReplyMember<'source>],
    pub(super) model: NativeReplyObjectRoute,
    pub(super) tile: NativeReplyObjectRoute,
    pub(super) list: NativeReplyObjectRoute,
    pub(super) root: NativeReplyObjectRoute,
    pub(super) replies: &'source [NativeReplyObjectRoute],
    pub(super) tile_row: u32,
    pub(super) tile_column: u32,
    pub(super) comment_key: u32,
    pub(super) operation: NativeReplyOperation<'source>,
    pub(super) new_root: NativeReplyStorageIdentity,
    pub(super) new_reply: Option<NativeReplyStorageIdentity>,
    /// If present, this must equal the existing root author identifier.  A
    /// missing value means the owner elects to preserve an authorless root.
    pub(super) author_identifier: Option<u64>,
    pub(super) limits: Limits,
    pub(super) path: CommentReplyPath,
}

/// One private decompressed IWA member edit.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeReplyMemberEdit {
    pub(super) component_index: usize,
    pub(super) member_name: String,
    pub(super) member_bytes: Vec<u8>,
}

/// A metadata edge delta discovered by the native transition.
///
/// The package metadata owner may use this to construct its typed combined
/// edge batch.  The edge is intentionally an object-id fact only; component
/// selectors belong to the package catalog.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct NativeReplyReferenceEdge {
    pub(super) source_identifier: u64,
    pub(super) target_identifier: u64,
}

/// Compact aggregate native accounting returned to the package owner.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct NativeReplyReports {
    pub(super) source_bytes: usize,
    pub(super) candidate_bytes: usize,
    pub(super) fields: usize,
    pub(super) work_bytes: usize,
    pub(super) references: usize,
    pub(super) replies: usize,
    pub(super) allocations: usize,
    pub(super) scratch_bytes: usize,
    pub(super) retained_bytes: usize,
}

/// Private result of a direct-reply native transition.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeReplyOutput {
    pub(super) member_edits: Vec<NativeReplyMemberEdit>,
    pub(super) source_cell: Vec<u8>,
    pub(super) target_cell: Vec<u8>,
    pub(super) before_root_identifier: u64,
    pub(super) after_root_identifier: u64,
    pub(super) before_reply_identifiers: Vec<u64>,
    pub(super) after_reply_identifiers: Vec<u64>,
    pub(super) added_object_identifiers: Vec<u64>,
    pub(super) removed_object_identifiers: Vec<u64>,
    pub(super) added_edges: Vec<NativeReplyReferenceEdge>,
    pub(super) removed_edges: Vec<NativeReplyReferenceEdge>,
    pub(super) touched_components: Vec<usize>,
    pub(super) reports: NativeReplyReports,
}

/// Native graph failure.  The public facade maps this to its redacted error
/// vocabulary and never exposes payload bytes through this enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NativeReplyError {
    InvalidSource,
    UnsupportedDependency,
    Allocation,
    Limit,
    Codec,
    Archive,
}

impl fmt::Display for NativeReplyError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InvalidSource => "invalid comment-reply native source",
            Self::UnsupportedDependency => "unsupported comment-reply native dependency",
            Self::Allocation => "comment-reply native allocation failed",
            Self::Limit => "comment-reply native limit exceeded",
            Self::Codec => "comment-reply codec rejected the source",
            Self::Archive => "comment-reply archive rewrite failed",
        })
    }
}

impl std::error::Error for NativeReplyError {}

type Result<T> = std::result::Result<T, NativeReplyError>;

fn map_popup_native_error(error: popup_native::NativePopUpError) -> NativeReplyError {
    match error {
        popup_native::NativePopUpError::InvalidSource => NativeReplyError::InvalidSource,
        popup_native::NativePopUpError::UnsupportedDependency => {
            NativeReplyError::UnsupportedDependency
        },
        popup_native::NativePopUpError::Allocation => NativeReplyError::Allocation,
        popup_native::NativePopUpError::Limit => NativeReplyError::Limit,
        popup_native::NativePopUpError::Codec => NativeReplyError::Codec,
        popup_native::NativePopUpError::Archive => NativeReplyError::Archive,
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ListEntryFact {
    key: u32,
    ref_count: u32,
    storage_identifier: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentListFact {
    object_identifier: u64,
    message_index: usize,
    next_list_id: u32,
    entries: Vec<ListEntryFact>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StorageFact {
    object_identifier: u64,
    author_identifier: Option<u64>,
    reply_identifiers: Vec<u64>,
    uuid: (u64, u64),
}

#[derive(Debug, Default)]
struct ListProbe {
    entries: Vec<ListEntryFact>,
    segments: usize,
    invalid: bool,
}

impl storage_codec::StorageVisitor for ListProbe {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        let Some(reference) = snapshot.comment_storage() else {
            self.invalid = true;
            return Ok(());
        };
        if snapshot.string_value().is_some()
            || snapshot.reference().is_some()
            || snapshot.formula().is_some()
            || snapshot.format().is_some()
            || snapshot.custom_format().is_some()
            || snapshot.rich_text_payload().is_some()
            || snapshot.import_warning_set().is_some()
            || snapshot.cell_spec().is_some()
        {
            self.invalid = true;
            return Ok(());
        }
        if self.entries.try_reserve_exact(1).is_err() {
            self.invalid = true;
            return Ok(());
        }
        self.entries.push(ListEntryFact {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            storage_identifier: reference.identifier(),
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.segments = self.segments.saturating_add(1);
        Ok(())
    }
}

/// A fallible, direct identifier index used by the native reply validator.
///
/// The archive itself stores objects in source order, so this index is only a
/// lookup aid; it never becomes an ordering authority.  Reserving the map
/// before the first insertion makes malformed attacker-sized object/fact
/// collections fail as a typed allocation or budget error instead of growing
/// through an unaccounted sequence of rehashes.
#[derive(Debug)]
struct IdIndex {
    positions: HashMap<u64, usize>,
}

impl IdIndex {
    fn with_capacity(
        capacity: usize,
        budget: &mut TransactionBudget,
        path: BudgetPath,
    ) -> Result<Self> {
        let entry_size = size_of::<u64>()
            .checked_add(size_of::<usize>())
            .ok_or(NativeReplyError::InvalidSource)?;
        let scratch = capacity
            .checked_mul(entry_size)
            .ok_or(NativeReplyError::InvalidSource)?;
        budget
            .charge_allocations(1, path)
            .and_then(|_| budget.charge_scratch_bytes(scratch, path))
            .and_then(|_| budget.charge_transaction_work(capacity, path))
            .map_err(|_| NativeReplyError::Limit)?;
        Self::try_with_capacity(capacity)
    }

    fn try_with_capacity(capacity: usize) -> Result<Self> {
        let mut positions = HashMap::new();
        positions
            .try_reserve(capacity)
            .map_err(|_| NativeReplyError::Allocation)?;
        Ok(Self { positions })
    }

    fn insert(&mut self, identifier: u64, position: usize) -> Option<usize> {
        self.positions.insert(identifier, position)
    }

    fn get(&self, identifier: u64) -> Option<usize> {
        self.positions.get(&identifier).copied()
    }

    #[allow(dead_code)]
    fn len(&self) -> usize {
        self.positions.len()
    }
}

fn reserved_id_set(
    capacity: usize,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<HashSet<u64>> {
    let scratch = capacity
        .checked_mul(size_of::<u64>())
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, path)
        .and_then(|_| budget.charge_scratch_bytes(scratch, path))
        .and_then(|_| budget.charge_transaction_work(capacity, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    Ok(identifiers)
}

fn fallible_u64_vec(capacity: usize) -> Result<Vec<u64>> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    Ok(values)
}

#[derive(Debug)]
struct ReplyProbe {
    identifiers: Vec<u64>,
    identifier_set: HashSet<u64>,
    capacity: usize,
    unsupported_reference: bool,
    allocation_failed: bool,
}

impl ReplyProbe {
    fn with_capacity(capacity: usize) -> Result<Self> {
        let identifiers = fallible_u64_vec(capacity)?;
        let mut identifier_set = HashSet::new();
        identifier_set
            .try_reserve(capacity)
            .map_err(|_| NativeReplyError::Allocation)?;
        Ok(Self {
            identifiers,
            identifier_set,
            capacity,
            unsupported_reference: false,
            allocation_failed: false,
        })
    }

    fn record_identifier(
        &mut self,
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    ) {
        if self.allocation_failed {
            return;
        }
        if identifier == 0
            || deprecated_type.is_some_and(|value| value != 0)
            || deprecated_is_external == Some(true)
            || self.identifier_set.contains(&identifier)
        {
            self.unsupported_reference = true;
            return;
        }
        if self.identifiers.len() >= self.capacity {
            self.allocation_failed = true;
            return;
        }
        // Keep the callback infallible from the codec's perspective while
        // making both attacker-scaled collections explicitly fallible.  The
        // initial reservation is sized from the complete source envelope;
        // these one-item checks also cover any future change to that bound.
        if self.identifiers.try_reserve(1).is_err() || self.identifier_set.try_reserve(1).is_err() {
            self.allocation_failed = true;
            return;
        }
        if !self.identifier_set.insert(identifier) {
            self.unsupported_reference = true;
            return;
        }
        self.identifiers.push(identifier);
    }
}

impl comment_storage_codec::CommentStorageVisitor for ReplyProbe {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), comment_storage_codec::DecodeError> {
        let reference = reply.reference();
        self.record_identifier(
            reference.identifier(),
            reference.deprecated_type(),
            reference.deprecated_is_external(),
        );
        Ok(())
    }
}

#[derive(Debug, Default)]
struct TileCommentProbe {
    counts: Vec<(u32, u32)>,
    invalid: bool,
}

impl storage_codec::StorageVisitor for TileCommentProbe {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        if census_row_comments(row, &mut self.counts).is_err() {
            self.invalid = true;
        }
        Ok(())
    }
}

fn increment_count(counts: &mut Vec<(u32, u32)>, identifier: u32) -> Result<()> {
    if identifier == 0 {
        return Err(NativeReplyError::InvalidSource);
    }
    if let Some((_, count)) = counts.iter_mut().find(|(key, _)| *key == identifier) {
        *count = count
            .checked_add(1)
            .ok_or(NativeReplyError::InvalidSource)?;
        return Ok(());
    }
    counts
        .try_reserve_exact(1)
        .map_err(|_| NativeReplyError::Allocation)?;
    counts.push((identifier, 1));
    Ok(())
}

fn count_for(counts: &[(u32, u32)], identifier: u32) -> u32 {
    counts
        .iter()
        .find(|(key, _)| *key == identifier)
        .map_or(0, |(_, count)| *count)
}

fn census_row_comments(
    row: storage_codec::TileRowInfoSnapshot<'_>,
    counts: &mut Vec<(u32, u32)>,
) -> Result<()> {
    let (storage, offsets) = match (row.cell_storage_buffer(), row.cell_offsets()) {
        (Some(storage), Some(offsets)) => (storage, offsets),
        (None, None) => (
            row.cell_storage_buffer_pre_bnc(),
            row.cell_offsets_pre_bnc(),
        ),
        _ => return Err(NativeReplyError::InvalidSource),
    };
    let cell_count =
        usize::try_from(row.cell_count()).map_err(|_| NativeReplyError::InvalidSource)?;
    if offsets.is_empty() {
        if cell_count != 1 {
            return Err(NativeReplyError::InvalidSource);
        }
        census_one_cell(storage, counts)?;
        return Ok(());
    }
    let expected_offset_bytes = cell_count
        .checked_mul(2)
        .ok_or(NativeReplyError::InvalidSource)?;
    if offsets.len() != expected_offset_bytes {
        return Err(NativeReplyError::InvalidSource);
    }
    let unit = if row.has_wide_offsets().unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(cell_count)
        .map_err(|_| NativeReplyError::Allocation)?;
    let mut previous = None;
    for encoded in offsets.chunks_exact(2) {
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        let start = if raw == u16::MAX {
            None
        } else {
            let value = usize::from(raw)
                .checked_mul(unit)
                .ok_or(NativeReplyError::InvalidSource)?;
            if value > storage.len() || previous.is_some_and(|prior| prior > value) {
                return Err(NativeReplyError::InvalidSource);
            }
            previous = Some(value);
            Some(value)
        };
        starts.push(start);
    }
    let occupied = starts.iter().flatten().count();
    if occupied != cell_count {
        return Err(NativeReplyError::InvalidSource);
    }
    for (index, start) in starts.iter().flatten().enumerate() {
        let end = starts
            .iter()
            .skip(index + 1)
            .flatten()
            .next()
            .copied()
            .unwrap_or(storage.len());
        if *start > end || end > storage.len() {
            return Err(NativeReplyError::InvalidSource);
        }
        census_one_cell(&storage[*start..end], counts)?;
    }
    Ok(())
}

fn census_one_cell(source: &[u8], counts: &mut Vec<(u32, u32)>) -> Result<()> {
    let cell = BncCell::parse(source).map_err(|_| NativeReplyError::InvalidSource)?;
    if let Some(identifier) = cell.comment_identifier() {
        increment_count(counts, identifier)?;
    }
    Ok(())
}

fn storage_options(budget: &TransactionBudget, source: &[u8]) -> storage_codec::DecodeOptions {
    budget.residual_storage_options(source)
}

fn comment_options(source: &[u8], limits: Limits) -> comment_storage_codec::DecodeOptions {
    let bytes = source.len().max(1).min(limits.max_message_bytes().max(1));
    let fields = source
        .len()
        .saturating_mul(32)
        .max(1)
        .min(limits.max_header_fields().max(1));
    let work = source
        .len()
        .saturating_mul(128)
        .max(1)
        .min(limits.max_header_memory_bytes().max(1));
    let references = limits.max_metadata_items().max(1);
    comment_storage_codec::DecodeOptions::new(
        bytes,
        fields,
        work,
        limits.max_header_nesting().try_into().unwrap_or(u32::MAX),
        references,
        limits.max_message_bytes().max(1),
    )
}

fn charge_comment_report(
    budget: &mut TransactionBudget,
    report: comment_storage_codec::DecodeReport,
    path: BudgetPath,
) -> Result<()> {
    budget
        .charge_wire_bytes(report.source_bytes(), path)
        .and_then(|_| budget.charge_wire_fields(report.fields(), path))
        .and_then(|_| budget.charge_wire_work(report.work_bytes(), path))
        .and_then(|_| budget.charge_wire_nesting(report.max_depth(), path))
        .and_then(|_| budget.charge_payload_references(report.references(), path))
        .and_then(|_| budget.charge_wire_reference_bytes(report.reference_bytes(), path))
        .and_then(|_| budget.charge_wire_text_bytes(report.text_bytes(), path))
        .map_err(|_| NativeReplyError::Limit)
}

fn charge_storage_removal_report(
    budget: &mut TransactionBudget,
    report: storage_codec::TableDataListEntryRemovalReport,
    path: BudgetPath,
) -> Result<()> {
    budget
        .charge_storage_decode_report(report.source(), true, path)
        .and_then(|_| budget.charge_storage_decode_report(report.selection(), true, path))
        .and_then(|_| budget.charge_storage_decode_report(report.result(), true, path))
        .and_then(|_| budget.charge_storage_decode_report(report.verification(), true, path))
        .and_then(|_| budget.charge_output(report.output_bytes(), path))
        .and_then(|_| budget.charge_transaction_work(report.rewrite_work_bytes(), path))
        .map_err(|_| NativeReplyError::Limit)
}

fn decode_list(
    source: &[u8],
    budget: &mut TransactionBudget,
    limits: Limits,
    path: BudgetPath,
) -> Result<(storage_codec::TableDataListSnapshot, ListProbe)> {
    let mut probe = ListProbe::default();
    let (snapshot, report) = storage_codec::decode_table_data_list_with_visitor(
        source,
        storage_options(budget, source),
        &mut probe,
    )
    .map_err(|_| NativeReplyError::Codec)?;
    budget
        .charge_storage_decode_report(report, true, path)
        .map_err(|_| NativeReplyError::Limit)?;
    if probe.invalid {
        return Err(NativeReplyError::InvalidSource);
    }
    if snapshot.list_type() != COMMENT_LIST_TYPE || probe.segments != 0 {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    if probe.entries.is_empty()
        || probe
            .entries
            .iter()
            .any(|entry| entry.key == 0 || entry.ref_count == 0 || entry.storage_identifier == 0)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if probe.entries.iter().enumerate().any(|(index, entry)| {
        probe.entries[index + 1..].iter().any(|other| {
            other.key == entry.key || other.storage_identifier == entry.storage_identifier
        })
    }) {
        return Err(NativeReplyError::InvalidSource);
    }
    let _ = limits;
    Ok((snapshot, probe))
}

fn collect_comment_lists(
    archive: &Archive,
    budget: &mut TransactionBudget,
    limits: Limits,
    path: BudgetPath,
) -> Result<Vec<CommentListFact>> {
    let message_count = archive
        .objects
        .iter()
        .try_fold(0usize, |total, object| {
            total.checked_add(object.messages.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    let payload_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    // The list visitor owns a probe, one fact per comment list, and two
    // registry sets.  Reserve the complete envelope before the first codec
    // scan; the individual decode reports below then debit their exact wire
    // work from the same ledger.
    budget
        .charge_allocations(message_count.saturating_add(4), path)
        .and_then(|_| budget.charge_scratch_bytes(payload_bytes, path))
        .and_then(|_| budget.charge_transaction_work(payload_bytes, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut lists = Vec::new();
    for object in &archive.objects {
        let object_identifier = object
            .archive_info
            .identifier
            .ok_or(NativeReplyError::InvalidSource)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            if !TABLE_DATA_LIST_TYPES.contains(&message.type_) {
                continue;
            }
            let (list_type_snapshot, list_type_report) =
                storage_codec::decode_table_data_list_type_with_report(
                    &message.data,
                    storage_options(budget, &message.data),
                )
                .map_err(|_| NativeReplyError::Codec)?;
            budget
                .charge_storage_decode_report(list_type_report, true, path)
                .map_err(|_| NativeReplyError::Limit)?;
            let list_type = list_type_snapshot.list_type();
            if list_type != COMMENT_LIST_TYPE {
                continue;
            }
            let (snapshot, probe) = decode_list(&message.data, budget, limits, path)?;
            let info = object
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(NativeReplyError::InvalidSource)?;
            validate_comment_list_message_info(info, &probe.entries)?;
            lists
                .try_reserve_exact(1)
                .map_err(|_| NativeReplyError::Allocation)?;
            lists.push(CommentListFact {
                object_identifier,
                message_index,
                next_list_id: snapshot.next_list_id(),
                entries: probe.entries,
            });
        }
    }
    if lists.is_empty() {
        return Err(NativeReplyError::InvalidSource);
    }
    let capacity = lists
        .iter()
        .try_fold(0usize, |total, list| total.checked_add(list.entries.len()))
        .ok_or(NativeReplyError::InvalidSource)?;
    let mut keys = HashSet::new();
    keys.try_reserve(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    let mut storages = HashSet::new();
    storages
        .try_reserve(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    for list in &lists {
        if list.next_list_id == 0 {
            return Err(NativeReplyError::InvalidSource);
        }
        for entry in &list.entries {
            if !keys.insert(entry.key) || !storages.insert(entry.storage_identifier) {
                return Err(NativeReplyError::InvalidSource);
            }
        }
    }
    Ok(lists)
}

fn allocate_comment_key(lists: &[CommentListFact], next_list_id: u32) -> Result<u32> {
    let mut keys = HashSet::new();
    let capacity = lists
        .iter()
        .try_fold(0usize, |total, list| total.checked_add(list.entries.len()))
        .ok_or(NativeReplyError::InvalidSource)?;
    keys.try_reserve(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    for list in lists {
        for entry in &list.entries {
            keys.insert(entry.key);
        }
    }
    let mut candidate = next_list_id.max(1);
    loop {
        if !keys.contains(&candidate) {
            return Ok(candidate);
        }
        candidate = candidate
            .checked_add(1)
            .ok_or(NativeReplyError::InvalidSource)?;
    }
}

fn storage_id_set(
    lists: &[CommentListFact],
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<HashSet<u64>> {
    let capacity = lists
        .iter()
        .try_fold(0usize, |total, list| total.checked_add(list.entries.len()))
        .ok_or(NativeReplyError::InvalidSource)?;
    let mut identifiers = reserved_id_set(capacity, budget, path)?;
    for list in lists {
        for entry in &list.entries {
            if !identifiers.insert(entry.storage_identifier) {
                return Err(NativeReplyError::InvalidSource);
            }
        }
    }
    Ok(identifiers)
}

fn validate_comment_list_message_info(
    info: &litchi_iwa_core::MessageInfo,
    entries: &[ListEntryFact],
) -> Result<()> {
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(entries.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    expected.extend(entries.iter().map(|entry| entry.storage_identifier));
    if info.object_references != expected {
        return Err(NativeReplyError::InvalidSource);
    }
    if info.field_infos.is_empty() {
        return Ok(());
    }
    let mut seen = HashSet::new();
    seen.try_reserve(entries.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    for field in &info.field_infos {
        let path = field.path.as_slice();
        if path.len() == 2 && path[0] == LIST_ENTRY_FIELD_PREFIX {
            let key = path[1];
            let index = entries
                .iter()
                .position(|entry| u64::from(entry.key) == u64::from(key))
                .ok_or(NativeReplyError::InvalidSource)?;
            let expected_identifier = entries[index].storage_identifier;
            if !seen.insert(index)
                || field.object_references.as_slice() != [expected_identifier]
                || !field.data_references.is_empty()
                || field
                    .r#type
                    .is_some_and(|value| value != FieldType::ObjectReference)
            {
                return Err(NativeReplyError::InvalidSource);
            }
            continue;
        }
        if field
            .object_references
            .iter()
            .any(|identifier| !expected.contains(identifier))
        {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    if seen.len() != entries.len() {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok(())
}

fn decode_storage<'a>(
    payload: &'a [u8],
    budget: &mut TransactionBudget,
    limits: Limits,
) -> Result<(comment_storage_codec::CommentStorageSnapshot<'a>, Vec<u64>)> {
    // A length-delimited reply reference needs at least the outer tag and
    // length plus the nested identifier tag and value.  This is a safe upper
    // bound for both the ordered reply vector and its membership set, so the
    // visitor never has to grow either collection infallibly.
    let reply_capacity = (payload.len() / 4).min(limits.max_metadata_items().max(1));
    let reply_scratch = reply_capacity
        .checked_mul(2)
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(2, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(reply_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(reply_capacity, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut probe = ReplyProbe::with_capacity(reply_capacity)?;
    let (snapshot, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        payload,
        comment_options(payload, limits),
        &mut probe,
    )
    .map_err(|_| NativeReplyError::Codec)?;
    charge_comment_report(budget, report, BUDGET_PATH)?;
    if probe.allocation_failed {
        return Err(NativeReplyError::Allocation);
    }
    if probe.unsupported_reference {
        return Err(NativeReplyError::InvalidSource);
    }
    if report.replies() != probe.identifiers.len() {
        return Err(NativeReplyError::InvalidSource);
    }
    if snapshot.storage_uuid().is_none() {
        return Err(NativeReplyError::InvalidSource);
    }
    if snapshot
        .storage_uuid()
        .is_some_and(|uuid| uuid.lower() == 0 && uuid.upper() == 0)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if snapshot.author().is_some_and(|reference| {
        reference.identifier() == 0
            || reference.deprecated_type().is_some_and(|value| value != 0)
            || reference.deprecated_is_external() == Some(true)
    }) {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok((snapshot, probe.identifiers))
}

/// The exact source-ordered aggregate object-reference list for one storage
/// object, together with a constant-time membership index.  The ordered list
/// is retained for ArchiveInfo equality and transition authorization; the set
/// is only a validation aid and never changes source order.
#[derive(Debug, PartialEq, Eq)]
struct StorageReferenceAggregate {
    ordered: Vec<u64>,
    identifiers: HashSet<u64>,
}

impl StorageReferenceAggregate {
    fn new(
        author_identifier: Option<u64>,
        reply_identifiers: &[u64],
        budget: &mut TransactionBudget,
        path: BudgetPath,
    ) -> Result<Self> {
        let capacity = reply_identifiers
            .len()
            .checked_add(usize::from(author_identifier.is_some()))
            .ok_or(NativeReplyError::InvalidSource)?;
        let set_scratch = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(NativeReplyError::InvalidSource)?;
        budget
            .charge_allocations(1, path)
            .and_then(|_| budget.charge_scratch_bytes(set_scratch, path))
            .and_then(|_| budget.charge_transaction_work(capacity, path))
            .map_err(|_| NativeReplyError::Limit)?;
        let vector_scratch = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(NativeReplyError::InvalidSource)?;
        budget
            .charge_allocations(1, path)
            .and_then(|_| budget.charge_scratch_bytes(vector_scratch, path))
            .and_then(|_| budget.charge_transaction_work(capacity, path))
            .map_err(|_| NativeReplyError::Limit)?;
        Self::try_from_parts(author_identifier, reply_identifiers)
    }

    fn try_from_parts(author_identifier: Option<u64>, reply_identifiers: &[u64]) -> Result<Self> {
        let capacity = reply_identifiers
            .len()
            .checked_add(usize::from(author_identifier.is_some()))
            .ok_or(NativeReplyError::InvalidSource)?;
        let mut identifiers = HashSet::new();
        identifiers
            .try_reserve(capacity)
            .map_err(|_| NativeReplyError::Allocation)?;
        let mut ordered = fallible_u64_vec(capacity)?;
        if let Some(author) = author_identifier {
            if author == 0 || !identifiers.insert(author) {
                return Err(NativeReplyError::InvalidSource);
            }
            ordered.push(author);
        }
        for identifier in reply_identifiers {
            if *identifier == 0 || !identifiers.insert(*identifier) {
                return Err(NativeReplyError::InvalidSource);
            }
            ordered.push(*identifier);
        }
        Ok(Self {
            ordered,
            identifiers,
        })
    }

    fn as_slice(&self) -> &[u64] {
        &self.ordered
    }

    fn contains(&self, identifier: &u64) -> bool {
        self.identifiers.contains(identifier)
    }
}

fn collect_storage_facts(
    archive: &Archive,
    budget: &mut TransactionBudget,
    limits: Limits,
) -> Result<Vec<StorageFact>> {
    let storage_objects = archive
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == COMMENT_STORAGE_TYPE)
        })
        .count();
    let payload_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .filter(|message| message.type_ == COMMENT_STORAGE_TYPE)
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(
            storage_objects.saturating_mul(2).saturating_add(4),
            BUDGET_PATH,
        )
        .and_then(|_| budget.charge_scratch_bytes(payload_bytes, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(payload_bytes, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let uuid_scratch = storage_objects
        .checked_mul(2)
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(uuid_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(storage_objects, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut facts = Vec::new();
    let mut uuids = HashSet::new();
    uuids
        .try_reserve(storage_objects)
        .map_err(|_| NativeReplyError::Allocation)?;
    for object in &archive.objects {
        let Some(object_identifier) = object.archive_info.identifier else {
            return Err(NativeReplyError::InvalidSource);
        };
        let mut storage_index = None;
        let mut storage_count = 0usize;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == COMMENT_STORAGE_TYPE {
                storage_count = storage_count
                    .checked_add(1)
                    .ok_or(NativeReplyError::InvalidSource)?;
                storage_index = Some(index);
            }
        }
        if storage_count == 0 {
            continue;
        }
        if storage_count != 1 || object.messages.len() != 1 {
            return Err(NativeReplyError::InvalidSource);
        }
        let message_index = storage_index.ok_or(NativeReplyError::InvalidSource)?;
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(NativeReplyError::InvalidSource)?;
        let payload = &object.messages[message_index].data;
        if info.type_ != COMMENT_STORAGE_TYPE || info.length != payload.len() as u32 {
            return Err(NativeReplyError::InvalidSource);
        }
        let (snapshot, reply_identifiers) = decode_storage(payload, budget, limits)?;
        let uuid = snapshot
            .storage_uuid()
            .ok_or(NativeReplyError::InvalidSource)?;
        let uuid_pair = (uuid.lower(), uuid.upper());
        if !uuids.insert(uuid_pair) {
            return Err(NativeReplyError::UnsupportedDependency);
        }
        let author_identifier = snapshot.author().map(|reference| reference.identifier());
        let expected_aggregate = expected_storage_references(
            author_identifier,
            &reply_identifiers,
            budget,
            BUDGET_PATH,
        )?;
        validate_storage_message_info(
            info,
            author_identifier,
            &expected_aggregate,
            &reply_identifiers,
        )?;
        facts
            .try_reserve_exact(1)
            .map_err(|_| NativeReplyError::Allocation)?;
        facts.push(StorageFact {
            object_identifier,
            author_identifier,
            reply_identifiers,
            uuid: uuid_pair,
        });
    }
    if facts.is_empty() {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok(facts)
}

fn expected_storage_references(
    author_identifier: Option<u64>,
    reply_identifiers: &[u64],
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<StorageReferenceAggregate> {
    StorageReferenceAggregate::new(author_identifier, reply_identifiers, budget, path)
}

fn validate_storage_message_info(
    info: &litchi_iwa_core::MessageInfo,
    author_identifier: Option<u64>,
    expected_aggregate: &StorageReferenceAggregate,
    reply_identifiers: &[u64],
) -> Result<()> {
    if info.object_references != expected_aggregate.as_slice() {
        return Err(NativeReplyError::InvalidSource);
    }
    if info.field_infos.is_empty() {
        return Ok(());
    }
    let mut author_fields = 0usize;
    let mut reply_fields = Vec::new();
    reply_fields
        .try_reserve_exact(reply_identifiers.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    reply_fields.resize(reply_identifiers.len(), false);
    for field in &info.field_infos {
        let path = field.path.as_slice();
        if path == AUTHOR_FIELD_PATH {
            author_fields = author_fields
                .checked_add(1)
                .ok_or(NativeReplyError::InvalidSource)?;
            let Some(author) = author_identifier else {
                return Err(NativeReplyError::InvalidSource);
            };
            if field.object_references.as_slice() != [author]
                || !field.data_references.is_empty()
                || field
                    .r#type
                    .is_some_and(|value| value != FieldType::ObjectReference)
            {
                return Err(NativeReplyError::InvalidSource);
            }
            continue;
        }
        if path.len() == 2 && path[0] == REPLY_FIELD_PREFIX {
            let ordinal = usize::try_from(path[1]).map_err(|_| NativeReplyError::InvalidSource)?;
            let expected = reply_identifiers
                .get(ordinal)
                .ok_or(NativeReplyError::InvalidSource)?;
            let seen = reply_fields
                .get_mut(ordinal)
                .ok_or(NativeReplyError::InvalidSource)?;
            if *seen
                || field.object_references.as_slice() != [*expected]
                || !field.data_references.is_empty()
                || field
                    .r#type
                    .is_some_and(|value| value != FieldType::ObjectReference)
            {
                return Err(NativeReplyError::InvalidSource);
            }
            *seen = true;
            continue;
        }
        // Unknown FieldInfo paths are retained, but an object edge they carry
        // must still be accounted for by the exact aggregate transition.
        if field
            .object_references
            .iter()
            .any(|identifier| !expected_aggregate.contains(identifier))
        {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    if author_fields != usize::from(author_identifier.is_some())
        || reply_fields.iter().any(|seen| !seen)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok(())
}

fn validate_archive_headers(
    archive: &Archive,
    limits: Limits,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<IdIndex> {
    let mut identifiers = IdIndex::with_capacity(archive.objects.len(), budget, path)?;
    for (object_index, object) in archive.objects.iter().enumerate() {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(NativeReplyError::InvalidSource)?;
        if identifier == 0 || identifiers.insert(identifier, object_index).is_some() {
            return Err(NativeReplyError::UnsupportedDependency);
        }
        if object.messages.len() != object.archive_info.message_infos.len() {
            return Err(NativeReplyError::InvalidSource);
        }
        for (message, info) in object
            .messages
            .iter()
            .zip(&object.archive_info.message_infos)
        {
            if message.type_ != info.type_ || message.data.len() != info.length as usize {
                return Err(NativeReplyError::InvalidSource);
            }
        }
        object
            .validate_with_limits(limits)
            .map_err(|_| NativeReplyError::Archive)?;
    }
    Ok(identifiers)
}

fn storage_fact_index(
    facts: &[StorageFact],
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<IdIndex> {
    let mut identifiers = IdIndex::with_capacity(facts.len(), budget, path)?;
    for (fact_index, fact) in facts.iter().enumerate() {
        if fact.object_identifier == 0
            || identifiers
                .insert(fact.object_identifier, fact_index)
                .is_some()
        {
            return Err(NativeReplyError::UnsupportedDependency);
        }
    }
    Ok(identifiers)
}

/// Charge the parsed-archive envelope before any helper that allocates a
/// registry or walks ArchiveInfo.  `Archive::validate_with_limits` and the
/// ID index in `validate_archive_headers` do not expose a report, so their
/// object/message/payload envelope is conservatively reserved here while the
/// exact codec reports are charged by their callers.
fn preflight_archive(
    archive: &Archive,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<usize> {
    let message_count = archive
        .objects
        .iter()
        .try_fold(0usize, |total, object| {
            total.checked_add(object.messages.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    let decoded_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    let identifier_scratch = archive
        .objects
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(NativeReplyError::InvalidSource)?;
    let scratch = decoded_bytes
        .checked_add(identifier_scratch)
        .ok_or(NativeReplyError::InvalidSource)?;
    let allocations = archive
        .objects
        .len()
        .checked_add(message_count)
        .and_then(|count| count.checked_add(8))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_archive(archive, decoded_bytes, path)
        .and_then(|_| budget.charge_allocations(allocations, path))
        .and_then(|_| budget.charge_scratch_bytes(scratch, path))
        .and_then(|_| budget.charge_transaction_work(decoded_bytes, path))
        .map_err(|_| NativeReplyError::Limit)?;
    Ok(decoded_bytes)
}

fn collect_bnc_counts(
    archive: &Archive,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<Vec<(u32, u32)>> {
    let tile_count = archive
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == TILE_TYPE)
        })
        .count();
    let tile_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .filter(|message| message.type_ == TILE_TYPE)
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(tile_count.saturating_add(2), path)
        .and_then(|_| budget.charge_scratch_bytes(tile_bytes, path))
        .and_then(|_| budget.charge_transaction_work(tile_bytes, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut counts = Vec::new();
    for object in &archive.objects {
        let mut tile_indices = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == TILE_TYPE);
        let Some((_, tile)) = tile_indices.next() else {
            continue;
        };
        if tile_indices.next().is_some() {
            return Err(NativeReplyError::InvalidSource);
        }
        let mut visitor = TileCommentProbe::default();
        let (_, report) = storage_codec::decode_tile_with_visitor(
            &tile.data,
            storage_options(budget, &tile.data),
            &mut visitor,
        )
        .map_err(|_| NativeReplyError::Codec)?;
        budget
            .charge_storage_decode_report(report, true, path)
            .map_err(|_| NativeReplyError::Limit)?;
        if visitor.invalid {
            return Err(NativeReplyError::InvalidSource);
        }
        counts
            .try_reserve(visitor.counts.len())
            .map_err(|_| NativeReplyError::Allocation)?;
        for (identifier, count) in visitor.counts {
            for _ in 0..count {
                increment_count(&mut counts, identifier)?;
            }
        }
    }
    if counts.is_empty() {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok(counts)
}

fn validate_list_refcounts(
    lists: &[CommentListFact],
    counts: &[(u32, u32)],
    storage_ids: &HashSet<u64>,
    budget: &mut TransactionBudget,
) -> Result<()> {
    let capacity = lists
        .iter()
        .try_fold(0usize, |total, list| total.checked_add(list.entries.len()))
        .ok_or(NativeReplyError::InvalidSource)?;
    let seen_scratch = capacity
        .checked_mul(size_of::<u32>())
        .ok_or(NativeReplyError::InvalidSource)?;
    let count_scratch = counts
        .len()
        .checked_mul(size_of::<u32>() * 2)
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(2, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(seen_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_scratch_bytes(count_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(capacity, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(counts.len(), BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut seen_keys = HashSet::new();
    seen_keys
        .try_reserve(capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    let mut count_index = HashMap::new();
    count_index
        .try_reserve(counts.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    for (identifier, count) in counts {
        if count_index.insert(*identifier, *count).is_some() {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    for list in lists {
        for entry in &list.entries {
            if !storage_ids.contains(&entry.storage_identifier)
                || count_index.get(&entry.key).copied().unwrap_or(0) != entry.ref_count
                || !seen_keys.insert(entry.key)
            {
                return Err(NativeReplyError::InvalidSource);
            }
        }
    }
    if counts
        .iter()
        .any(|(identifier, _)| !seen_keys.contains(identifier))
    {
        return Err(NativeReplyError::InvalidSource);
    }
    Ok(())
}

fn selected_list<'a>(
    lists: &'a [CommentListFact],
    route: NativeReplyObjectRoute,
    archive: &Archive,
    object_index: &IdIndex,
) -> Result<&'a CommentListFact> {
    if route.member_index != 0 {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    let object_index = object_index
        .get(route.identifier)
        .ok_or(NativeReplyError::InvalidSource)?;
    let object = archive
        .objects
        .get(object_index)
        .ok_or(NativeReplyError::InvalidSource)?;
    let _ = object
        .messages
        .get(route.message_index)
        .ok_or(NativeReplyError::InvalidSource)?;
    let list = lists
        .iter()
        .find(|list| {
            list.object_identifier == route.identifier && list.message_index == route.message_index
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    Ok(list)
}

fn route_object<'archive>(
    archive: &'archive Archive,
    object_index: &IdIndex,
    route: NativeReplyObjectRoute,
    expected_type: Option<u32>,
) -> Result<&'archive ArchiveObject> {
    if route.member_index != 0 || route.identifier == 0 {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    let object = archive
        .objects
        .get(
            object_index
                .get(route.identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?;
    let message = object
        .messages
        .get(route.message_index)
        .ok_or(NativeReplyError::InvalidSource)?;
    if let Some(expected) = expected_type {
        if message.type_ != expected {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    Ok(object)
}

fn storage_route_payload<'archive>(
    archive: &'archive Archive,
    object_index: &IdIndex,
    route: NativeReplyObjectRoute,
) -> Result<&'archive [u8]> {
    let object = route_object(archive, object_index, route, Some(COMMENT_STORAGE_TYPE))?;
    Ok(&object
        .messages
        .get(route.message_index)
        .ok_or(NativeReplyError::InvalidSource)?
        .data)
}

fn validate_routes(
    request: &NativeReplyRequest<'_>,
    archive: &Archive,
    object_index: &IdIndex,
    budget: &mut TransactionBudget,
) -> Result<()> {
    route_object(archive, object_index, request.model, None).and_then(|object| {
        let message = object
            .messages
            .get(request.model.message_index)
            .ok_or(NativeReplyError::InvalidSource)?;
        if !TABLE_MODEL_TYPES.contains(&message.type_) {
            return Err(NativeReplyError::InvalidSource);
        }
        Ok(())
    })?;
    route_object(archive, object_index, request.tile, Some(TILE_TYPE))?;
    route_object(archive, object_index, request.list, None)?;
    route_object(
        archive,
        object_index,
        request.root,
        Some(COMMENT_STORAGE_TYPE),
    )?;
    let mut reply_ids = reserved_id_set(request.replies.len(), budget, BUDGET_PATH)?;
    for reply in request.replies {
        route_object(archive, object_index, *reply, Some(COMMENT_STORAGE_TYPE))?;
        if !reply_ids.insert(reply.identifier) {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    if request.comment_key == 0 || request.new_root.identifier == 0 {
        return Err(NativeReplyError::InvalidSource);
    }
    if request
        .new_reply
        .is_some_and(|identity| identity.identifier == 0)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    match request.operation {
        NativeReplyOperation::Append { .. } | NativeReplyOperation::Replace { .. }
            if request.new_reply.is_none() =>
        {
            return Err(NativeReplyError::InvalidSource);
        },
        NativeReplyOperation::Remove { .. } if request.new_reply.is_some() => {
            return Err(NativeReplyError::InvalidSource);
        },
        _ => {},
    }
    if request.new_root.identifier == request.root.identifier
        || request
            .new_reply
            .is_some_and(|identity| identity.identifier == request.root.identifier)
        || request
            .new_reply
            .is_some_and(|identity| identity.identifier == request.new_root.identifier)
    {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_reply_graph(
    archive: &Archive,
    object_index: &IdIndex,
    request: &NativeReplyRequest<'_>,
    root_fact: &StorageFact,
    facts: &[StorageFact],
    fact_index: &IdIndex,
    budget: &mut TransactionBudget,
) -> Result<()> {
    if root_fact.object_identifier != request.root.identifier
        || root_fact.reply_identifiers.len() != request.replies.len()
    {
        return Err(NativeReplyError::InvalidSource);
    }
    let mut reply_ids = reserved_id_set(request.replies.len(), budget, BUDGET_PATH)?;
    for route in request.replies {
        let id = route.identifier;
        if id == request.root.identifier || !reply_ids.insert(id) {
            return Err(NativeReplyError::InvalidSource);
        }
        let fact = fact_index
            .get(id)
            .and_then(|index| facts.get(index))
            .ok_or(NativeReplyError::InvalidSource)?;
        if !fact.reply_identifiers.is_empty() {
            return Err(NativeReplyError::UnsupportedDependency);
        }
        let _ = storage_route_payload(archive, object_index, *route)?;
    }
    if !root_fact
        .reply_identifiers
        .iter()
        .zip(request.replies)
        .all(|(expected, route)| *expected == route.identifier)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if request
        .author_identifier
        .is_some_and(|identifier| Some(identifier) != root_fact.author_identifier)
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if reply_ids.contains(&root_fact.object_identifier) {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    Ok(())
}

/// Validate every comment-storage object in the supplied member, not only the
/// selected root.  This prevents an orphan or nested storage object from
/// becoming an implicit cull/reuse target when the selected list changes.
fn validate_all_reply_graphs(
    facts: &[StorageFact],
    root_ids: &HashSet<u64>,
    fact_index: &IdIndex,
    budget: &mut TransactionBudget,
) -> Result<()> {
    let mut referenced_replies = reserved_id_set(facts.len(), budget, BUDGET_PATH)?;
    for fact in facts {
        let is_root = root_ids.contains(&fact.object_identifier);
        if !is_root && !fact.reply_identifiers.is_empty() {
            return Err(NativeReplyError::UnsupportedDependency);
        }
        if !is_root && fact.reply_identifiers.is_empty() {
            continue;
        }
        for reply_identifier in &fact.reply_identifiers {
            if *reply_identifier == fact.object_identifier
                || root_ids.contains(reply_identifier)
                || !referenced_replies.insert(*reply_identifier)
            {
                return Err(NativeReplyError::UnsupportedDependency);
            }
            let reply = fact_index
                .get(*reply_identifier)
                .and_then(|index| facts.get(index))
                .ok_or(NativeReplyError::InvalidSource)?;
            if !reply.reply_identifiers.is_empty() {
                return Err(NativeReplyError::UnsupportedDependency);
            }
        }
    }
    if root_ids
        .iter()
        .any(|identifier| fact_index.get(*identifier).is_none())
        || facts.len()
            != root_ids
                .len()
                .checked_add(referenced_replies.len())
                .ok_or(NativeReplyError::InvalidSource)?
        || facts.iter().any(|fact| {
            !root_ids.contains(&fact.object_identifier)
                && !referenced_replies.contains(&fact.object_identifier)
        })
    {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_uuid_inputs(
    request: &NativeReplyRequest<'_>,
    facts: &[StorageFact],
    object_index: &IdIndex,
    budget: &mut TransactionBudget,
) -> Result<()> {
    let valid =
        |identity: NativeReplyStorageIdentity| identity.uuid_lower != 0 || identity.uuid_upper != 0;
    if !valid(request.new_root) || request.new_reply.is_some_and(|identity| !valid(identity)) {
        return Err(NativeReplyError::InvalidSource);
    }
    let uuid_scratch = facts
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(uuid_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(facts.len(), BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut uuids = HashSet::new();
    uuids
        .try_reserve(facts.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    for fact in facts {
        if !uuids.insert(fact.uuid) {
            return Err(NativeReplyError::UnsupportedDependency);
        }
    }
    if uuids.contains(&(request.new_root.uuid_lower, request.new_root.uuid_upper))
        || request
            .new_reply
            .is_some_and(|identity| uuids.contains(&(identity.uuid_lower, identity.uuid_upper)))
        || request.new_reply.is_some_and(|identity| {
            (identity.uuid_lower, identity.uuid_upper)
                == (request.new_root.uuid_lower, request.new_root.uuid_upper)
        })
    {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    if object_index.get(request.new_root.identifier).is_some()
        || request
            .new_reply
            .is_some_and(|identity| object_index.get(identity.identifier).is_some())
    {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_author_object(
    archive: &Archive,
    object_index: &IdIndex,
    identifier: u64,
) -> Result<()> {
    let object = archive
        .objects
        .get(
            object_index
                .get(identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?;
    let mut matching = 0usize;
    for message in &object.messages {
        if message.type_ == ANNOTATION_AUTHOR_TYPE {
            matching = matching
                .checked_add(1)
                .ok_or(NativeReplyError::InvalidSource)?;
        }
    }
    if matching != 1 || object.messages.len() != 1 {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReplyFieldMutation {
    Append,
    Replace(usize),
    Remove(usize),
}

/// Build one transition authorization for every existing FieldInfo.  The
/// archive core requires a complete authorization vector; supplying identity
/// transitions for unknown/unselected fields both preserves them and proves
/// that no opaque object edge is silently dropped.
fn reply_field_transitions<'a>(
    info: &'a litchi_iwa_core::MessageInfo,
    author_identifier: Option<u64>,
    before: &'a [u64],
    after: &'a [u64],
    mutation: ReplyFieldMutation,
    budget: &mut TransactionBudget,
) -> Result<Vec<FieldObjectReferenceTransition<'a>>> {
    let before_aggregate =
        expected_storage_references(author_identifier, before, budget, BUDGET_PATH)?;
    validate_storage_message_info(info, author_identifier, &before_aggregate, before)?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(info.field_infos.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let field_after = if field.path.as_slice() == AUTHOR_FIELD_PATH {
            field.object_references.as_slice()
        } else if field.path.as_slice().len() == 2 && field.path.as_slice()[0] == REPLY_FIELD_PREFIX
        {
            let ordinal = usize::try_from(field.path.as_slice()[1])
                .map_err(|_| NativeReplyError::InvalidSource)?;
            let after_ordinal = match mutation {
                ReplyFieldMutation::Append => Some(ordinal),
                ReplyFieldMutation::Replace(_) => Some(ordinal),
                ReplyFieldMutation::Remove(removed) => {
                    if ordinal == removed {
                        None
                    } else if ordinal > removed {
                        Some(ordinal - 1)
                    } else {
                        Some(ordinal)
                    }
                },
            };
            after_ordinal.map_or(&[][..], |index| {
                after.get(index).map_or(&[][..], std::slice::from_ref)
            })
        } else {
            field.object_references.as_slice()
        };
        fields.push(FieldObjectReferenceTransition {
            field_info_index,
            expected_path: field.path.as_slice(),
            before: field.object_references.as_slice(),
            after: field_after,
        });
    }
    Ok(fields)
}

fn adjust_reply_field_infos(
    archive: &mut Archive,
    route: NativeReplyObjectRoute,
    mutation: ReplyFieldMutation,
    after: &[u64],
) -> Result<()> {
    let info = archive
        .object_mut(route.identifier)
        .ok_or(NativeReplyError::InvalidSource)?
        .archive_info
        .message_infos
        .get_mut(route.message_index)
        .ok_or(NativeReplyError::InvalidSource)?;
    if info.field_infos.is_empty() {
        return Ok(());
    }
    match mutation {
        ReplyFieldMutation::Append => {
            let identifier = *after.last().ok_or(NativeReplyError::InvalidSource)?;
            let ordinal = u32::try_from(
                after
                    .len()
                    .checked_sub(1)
                    .ok_or(NativeReplyError::InvalidSource)?,
            )
            .map_err(|_| NativeReplyError::InvalidSource)?;
            info.field_infos
                .try_reserve_exact(1)
                .map_err(|_| NativeReplyError::Allocation)?;
            let mut field = litchi_iwa_core::FieldInfo::new(vec![REPLY_FIELD_PREFIX, ordinal]);
            field.r#type = Some(FieldType::ObjectReference);
            field
                .object_references
                .try_reserve_exact(1)
                .map_err(|_| NativeReplyError::Allocation)?;
            field.object_references.push(identifier);
            info.field_infos.push(field);
        },
        ReplyFieldMutation::Replace(_) => {},
        ReplyFieldMutation::Remove(removed) => {
            let remove_index = info
                .field_infos
                .iter()
                .position(|field| {
                    field.path.as_slice().len() == 2
                        && field.path.as_slice()[0] == REPLY_FIELD_PREFIX
                        && field.path.as_slice()[1] == u32::try_from(removed).unwrap_or(u32::MAX)
                })
                .ok_or(NativeReplyError::InvalidSource)?;
            info.field_infos.remove(remove_index);
            for field in &mut info.field_infos {
                let path = field.path.as_slice();
                if path.len() == 2
                    && path[0] == REPLY_FIELD_PREFIX
                    && path[1] > u32::try_from(removed).unwrap_or(u32::MAX)
                {
                    field.path = vec![REPLY_FIELD_PREFIX, path[1] - 1].into();
                }
            }
        },
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ListFieldMutation {
    old_key: u32,
    old_ref_count: u32,
    new_key: u32,
}

fn list_field_transitions<'a>(
    info: &'a litchi_iwa_core::MessageInfo,
    entries: &[ListEntryFact],
    mutation: ListFieldMutation,
    before: &'a [u64],
    after: &'a [u64],
) -> Result<Vec<FieldObjectReferenceTransition<'a>>> {
    validate_comment_list_message_info(info, entries)?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(info.field_infos.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let field_after = if field.path.as_slice().len() == 2
            && field.path.as_slice()[0] == LIST_ENTRY_FIELD_PREFIX
            && field.path.as_slice()[1] == mutation.old_key
            && mutation.old_ref_count == 1
        {
            &[][..]
        } else {
            field.object_references.as_slice()
        };
        fields.push(FieldObjectReferenceTransition {
            field_info_index,
            expected_path: field.path.as_slice(),
            before: field.object_references.as_slice(),
            after: field_after,
        });
    }
    let _ = (before, after);
    Ok(fields)
}

fn adjust_list_field_infos(
    archive: &mut Archive,
    route: NativeReplyObjectRoute,
    mutation: ListFieldMutation,
    new_identifier: u64,
) -> Result<()> {
    let info = archive
        .object_mut(route.identifier)
        .ok_or(NativeReplyError::InvalidSource)?
        .archive_info
        .message_infos
        .get_mut(route.message_index)
        .ok_or(NativeReplyError::InvalidSource)?;
    if info.field_infos.is_empty() {
        return Ok(());
    }
    if mutation.old_ref_count == 1 {
        let old_path = [LIST_ENTRY_FIELD_PREFIX, mutation.old_key];
        let remove_index = info
            .field_infos
            .iter()
            .position(|field| field.path.as_slice() == old_path)
            .ok_or(NativeReplyError::InvalidSource)?;
        info.field_infos.remove(remove_index);
    }
    info.field_infos
        .try_reserve_exact(1)
        .map_err(|_| NativeReplyError::Allocation)?;
    let mut field =
        litchi_iwa_core::FieldInfo::new(vec![LIST_ENTRY_FIELD_PREFIX, mutation.new_key]);
    field.r#type = Some(FieldType::ObjectReference);
    field
        .object_references
        .try_reserve_exact(1)
        .map_err(|_| NativeReplyError::Allocation)?;
    field.object_references.push(new_identifier);
    info.field_infos.push(field);
    Ok(())
}

fn replace_with_reference_transition(
    archive: &mut Archive,
    route: NativeReplyObjectRoute,
    payload: Vec<u8>,
    aggregate_before: &[u64],
    aggregate_after: &[u64],
    fields: &[FieldObjectReferenceTransition<'_>],
    limits: Limits,
) -> Result<()> {
    let message_type = {
        let object = archive
            .object(route.identifier)
            .ok_or(NativeReplyError::InvalidSource)?;
        let info = object
            .archive_info
            .message_infos
            .get(route.message_index)
            .ok_or(NativeReplyError::InvalidSource)?;
        if info.object_references != aggregate_before {
            return Err(NativeReplyError::InvalidSource);
        }
        info.type_
    };
    let transition = ObjectReferenceTransition {
        aggregate_before,
        aggregate_after,
        fields,
    };
    archive
        .object_mut(route.identifier)
        .ok_or(NativeReplyError::InvalidSource)?
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            transition,
            limits,
        )
        .map_err(|_| NativeReplyError::Archive)?;
    Ok(())
}

fn uuid_payload(identity: NativeReplyStorageIdentity) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(2 * (1 + size_of::<u64>()))
        .map_err(|_| NativeReplyError::Allocation)?;
    append_varint_field(&mut output, 1, identity.uuid_lower)
        .map_err(|_| NativeReplyError::InvalidSource)?;
    append_varint_field(&mut output, 2, identity.uuid_upper)
        .map_err(|_| NativeReplyError::InvalidSource)?;
    Ok(output)
}

fn patch_storage_uuid(source: &[u8], identity: NativeReplyStorageIdentity) -> Result<Vec<u8>> {
    let uuid = uuid_payload(identity)?;
    patch_length_delimited_field(source, 5, true, Some(&uuid))
        .map_err(|_| NativeReplyError::InvalidSource)
}

fn patch_leaf_text_and_uuid(
    source: &[u8],
    expected_text: Option<&str>,
    text: &str,
    identity: NativeReplyStorageIdentity,
) -> Result<Vec<u8>> {
    let with_text =
        patch_length_delimited_field(source, 1, expected_text.is_some(), Some(text.as_bytes()))
            .map_err(|_| NativeReplyError::InvalidSource)?;
    patch_storage_uuid(&with_text, identity)
}

fn append_comment_entry(source: &[u8], key: u32, identifier: u64) -> Result<Vec<u8>> {
    if key == 0 || identifier == 0 {
        return Err(NativeReplyError::InvalidSource);
    }
    let mut entry = Vec::new();
    entry
        .try_reserve_exact(32)
        .map_err(|_| NativeReplyError::Allocation)?;
    append_varint_field(&mut entry, 1, u64::from(key))
        .map_err(|_| NativeReplyError::InvalidSource)?;
    append_varint_field(&mut entry, 2, 1).map_err(|_| NativeReplyError::InvalidSource)?;
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, identifier)
        .map_err(|_| NativeReplyError::InvalidSource)?;
    append_length_delimited_field(&mut entry, 10, &reference)
        .map_err(|_| NativeReplyError::InvalidSource)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len().saturating_add(entry.len()).saturating_add(8))
        .map_err(|_| NativeReplyError::Allocation)?;
    output.extend_from_slice(source);
    append_length_delimited_field(&mut output, 3, &entry)
        .map_err(|_| NativeReplyError::InvalidSource)?;
    Ok(output)
}

fn rewrite_comment_list(
    source: &[u8],
    old_key: u32,
    old_ref_count: u32,
    old_identifier: u64,
    new_key: u32,
    new_identifier: u64,
    budget: &mut TransactionBudget,
    limits: Limits,
    path: BudgetPath,
) -> Result<Vec<u8>> {
    let rewrite_scratch = source
        .len()
        .checked_add(128)
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(4, path)
        .and_then(|_| budget.charge_scratch_bytes(rewrite_scratch, path))
        .and_then(|_| budget.charge_transaction_work(rewrite_scratch, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let options = storage_options(budget, source);
    let reduced = if old_ref_count == 1 {
        let (reduced, report) = storage_codec::remove_table_data_list_entry_with_report(
            source,
            storage_codec::TableDataListEntryRemoval::new(old_key, 1, old_identifier),
            options,
        )
        .map_err(|_| NativeReplyError::Codec)?;
        charge_storage_removal_report(budget, report, path)?;
        reduced
    } else {
        let mutation = storage_codec::TableDataListEntryMutation::RefCount(
            storage_codec::TableDataListEntryRefCountEdit::new(
                old_key,
                old_ref_count,
                old_ref_count
                    .checked_sub(1)
                    .ok_or(NativeReplyError::InvalidSource)?,
            ),
        );
        let prepared =
            storage_codec::prepare_table_data_list_entry_rewrite(source, mutation, options)
                .map_err(|_| NativeReplyError::Codec)?;
        let requirements = prepared.requirements();
        budget
            .charge_storage_decode_report(requirements.source(), true, path)
            .and_then(|_| budget.charge_wire_fields(requirements.fields(), path))
            .and_then(|_| budget.charge_wire_work(requirements.work_bytes(), path))
            .and_then(|_| budget.charge_wire_nesting(requirements.max_depth(), path))
            .and_then(|_| budget.charge_payload_references(requirements.references(), path))
            .and_then(|_| budget.charge_allocations(requirements.allocations(), path))
            .and_then(|_| budget.charge_scratch_bytes(requirements.scratch_bytes(), path))
            .and_then(|_| budget.charge_retained_bytes(requirements.retained_bytes(), path))
            .and_then(|_| budget.charge_output(requirements.output_bytes(), path))
            .and_then(|_| budget.charge_transaction_work(requirements.output_bytes(), path))
            .map_err(|_| NativeReplyError::Limit)?;
        prepared
            .execute(requirements.exact_limits())
            .map_err(|_| NativeReplyError::Codec)?
            .0
    };
    let appended = append_comment_entry(&reduced, new_key, new_identifier)?;
    let next = new_key
        .checked_add(1)
        .ok_or(NativeReplyError::InvalidSource)?;
    let output = patch_varint_field(&appended, 2, true, Some(u64::from(next)))
        .map_err(|_| NativeReplyError::InvalidSource)?;
    if output.len() > limits.max_message_bytes() {
        return Err(NativeReplyError::Limit);
    }
    Ok(output)
}

fn archive_reference_presence(
    archive: &Archive,
    first_identifier: u64,
    second_identifier: Option<u64>,
    limits: Limits,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<(bool, bool)> {
    struct Search {
        first_identifier: u64,
        second_identifier: Option<u64>,
        first_found: bool,
        second_found: bool,
    }
    impl ArchiveReferenceVisitor for Search {
        fn visit_reference(
            &mut self,
            occurrence: ArchiveReferenceOccurrence,
        ) -> litchi_iwa_core::Result<()> {
            if occurrence.kind != ArchiveReferenceKind::Object {
                return Ok(());
            }
            if occurrence.referenced_identifier == self.first_identifier {
                self.first_found = true;
            }
            if self
                .second_identifier
                .is_some_and(|identifier| occurrence.referenced_identifier == identifier)
            {
                self.second_found = true;
            }
            Ok(())
        }
    }
    let scratch = archive
        .objects
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(NativeReplyError::InvalidSource)?;
    let work = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, path)
        .and_then(|_| budget.charge_scratch_bytes(scratch, path))
        .and_then(|_| budget.charge_transaction_work(work, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut search = Search {
        first_identifier,
        second_identifier,
        first_found: false,
        second_found: false,
    };
    for object in &archive.objects {
        object
            .inspect_references_with_policy_and_limits(
                &mut search,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                limits,
            )
            .map_err(|_| NativeReplyError::Archive)?;
    }
    Ok((search.first_found, search.second_found))
}

fn validate_known_archive_metadata(
    archive: &Archive,
    limits: Limits,
    budget: &mut TransactionBudget,
    path: BudgetPath,
) -> Result<()> {
    struct MetadataProbe;
    impl ArchiveReferenceVisitor for MetadataProbe {
        fn visit_reference(
            &mut self,
            _occurrence: ArchiveReferenceOccurrence,
        ) -> litchi_iwa_core::Result<()> {
            Ok(())
        }
    }
    let scratch = archive
        .objects
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(NativeReplyError::InvalidSource)?;
    let work = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, path)
        .and_then(|_| budget.charge_scratch_bytes(scratch, path))
        .and_then(|_| budget.charge_transaction_work(work, path))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut probe = MetadataProbe;
    for object in &archive.objects {
        object
            .inspect_references_with_policy_and_limits(
                &mut probe,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                limits,
            )
            .map_err(|_| NativeReplyError::Archive)?;
    }
    Ok(())
}

/// Verify reply-specific object locality.  Unlike the Pop-Up Menu helper,
/// this transition may add fresh COW root/reply objects and remove old
/// unshared objects. Every unchanged source object remains exact; candidate
/// additions and source removals must be explicitly authorized.
fn verify_reply_archive_object_locality(
    source: &Archive,
    source_index: &IdIndex,
    candidate: &Archive,
    candidate_index: &IdIndex,
    changed_identifiers: &[u64],
) -> Result<()> {
    for source_object in &source.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(NativeReplyError::InvalidSource)?;
        if changed_identifiers.contains(&identifier) {
            continue;
        }
        let candidate_object = candidate
            .objects
            .get(
                candidate_index
                    .get(identifier)
                    .ok_or(NativeReplyError::InvalidSource)?,
            )
            .ok_or(NativeReplyError::InvalidSource)?;
        if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    for candidate_object in &candidate.objects {
        let identifier = candidate_object
            .archive_info
            .identifier
            .ok_or(NativeReplyError::InvalidSource)?;
        if source_index.get(identifier).is_none() && !changed_identifiers.contains(&identifier) {
            return Err(NativeReplyError::InvalidSource);
        }
    }
    Ok(())
}

/// Execute one strict direct-reply native transition.
///
/// All graph validation and candidate mutation happen on a private archive
/// clone.  The source member is never modified.  A split graph is rejected
/// before any candidate clone so the package owner cannot accidentally publish
/// a stale component/token set.
pub(super) fn rewrite_native_comment_reply(
    request: NativeReplyRequest<'_>,
    budget: &mut TransactionBudget,
) -> Result<NativeReplyOutput> {
    if request.members.len() != 1 {
        return Err(NativeReplyError::UnsupportedDependency);
    }
    let member = request
        .members
        .first()
        .copied()
        .ok_or(NativeReplyError::InvalidSource)?;
    if member.member_name.is_empty() || member.component_index == usize::MAX {
        return Err(NativeReplyError::InvalidSource);
    }
    let source = member.archive;
    let decoded_bytes = preflight_archive(source, budget, BUDGET_PATH)?;
    let object_index = validate_archive_headers(source, request.limits, budget, BUDGET_PATH)?;
    validate_routes(&request, source, &object_index, budget)?;

    let lists = collect_comment_lists(source, budget, request.limits, BUDGET_PATH)?;
    let selected = selected_list(&lists, request.list, source, &object_index)?;
    let old_entry = selected
        .entries
        .iter()
        .find(|entry| entry.key == request.comment_key)
        .ok_or(NativeReplyError::InvalidSource)?;
    let root_identifier = request.root.identifier;
    if old_entry.storage_identifier != root_identifier {
        return Err(NativeReplyError::InvalidSource);
    }
    let storage_ids = storage_id_set(&lists, budget, BUDGET_PATH)?;
    let facts = collect_storage_facts(source, budget, request.limits)?;
    if facts
        .iter()
        .filter(|fact| storage_ids.contains(&fact.object_identifier))
        .count()
        != storage_ids.len()
    {
        return Err(NativeReplyError::InvalidSource);
    }
    let fact_index = storage_fact_index(&facts, budget, BUDGET_PATH)?;
    validate_all_reply_graphs(&facts, &storage_ids, &fact_index, budget)?;
    validate_uuid_inputs(&request, &facts, &object_index, budget)?;
    validate_known_archive_metadata(source, request.limits, budget, BUDGET_PATH)?;
    let counts = collect_bnc_counts(source, budget, BUDGET_PATH)?;
    validate_list_refcounts(&lists, &counts, &storage_ids, budget)?;
    if count_for(&counts, request.comment_key) != old_entry.ref_count {
        return Err(NativeReplyError::InvalidSource);
    }
    let root_fact = facts
        .get(
            fact_index
                .get(root_identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?;
    let mut validated_authors = reserved_id_set(facts.len(), budget, BUDGET_PATH)?;
    for fact in &facts {
        if let Some(author) = fact.author_identifier {
            if storage_ids.contains(&author) {
                return Err(NativeReplyError::UnsupportedDependency);
            }
            if validated_authors.insert(author) {
                validate_author_object(source, &object_index, author)?;
            }
        }
    }
    validate_reply_graph(
        source,
        &object_index,
        &request,
        root_fact,
        &facts,
        &fact_index,
        budget,
    )?;
    if request
        .operation
        .ordinal()
        .is_some_and(|ordinal| ordinal >= request.replies.len())
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if request
        .operation
        .expected_reply_identifier()
        .is_some_and(|expected| {
            request
                .operation
                .ordinal()
                .and_then(|ordinal| request.replies.get(ordinal))
                .is_none_or(|route| route.identifier != expected)
        })
    {
        return Err(NativeReplyError::InvalidSource);
    }
    if request
        .operation
        .text()
        .is_some_and(|text| text.len() > request.limits.max_message_bytes())
    {
        return Err(NativeReplyError::Limit);
    }

    // Prepare the root reply-reference rewrite before cloning any candidate.
    let reply_id_scratch = root_fact
        .reply_identifiers
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(2, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(reply_id_scratch, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut before_reply_ids = fallible_u64_vec(root_fact.reply_identifiers.len())?;
    before_reply_ids.extend_from_slice(&root_fact.reply_identifiers);
    let ordinal = request.operation.ordinal();
    let old_reply_identifier = ordinal.and_then(|index| before_reply_ids.get(index).copied());
    let after_capacity = before_reply_ids
        .len()
        .checked_add(usize::from(matches!(
            request.operation,
            NativeReplyOperation::Append { .. }
        )))
        .ok_or(NativeReplyError::InvalidSource)?;
    let mut after_reply_ids = fallible_u64_vec(after_capacity)?;
    after_reply_ids.extend_from_slice(&before_reply_ids);
    match request.operation {
        NativeReplyOperation::Append { .. } => {
            let identity = request.new_reply.ok_or(NativeReplyError::InvalidSource)?;
            after_reply_ids.push(identity.identifier);
        },
        NativeReplyOperation::Replace { ordinal, .. } => {
            let identity = request.new_reply.ok_or(NativeReplyError::InvalidSource)?;
            let slot = after_reply_ids
                .get_mut(ordinal)
                .ok_or(NativeReplyError::InvalidSource)?;
            *slot = identity.identifier;
        },
        NativeReplyOperation::Remove { ordinal, .. } => {
            after_reply_ids.remove(ordinal);
        },
    }
    let author_identifier = root_fact.author_identifier;
    let before_aggregate =
        expected_storage_references(author_identifier, &before_reply_ids, budget, BUDGET_PATH)?;
    let after_aggregate =
        expected_storage_references(author_identifier, &after_reply_ids, budget, BUDGET_PATH)?;
    let root_payload = storage_route_payload(source, &object_index, request.root)?;
    let root_rewrite = match request.operation {
        NativeReplyOperation::Append { .. } => {
            comment_storage_codec::CommentStorageReplyRewrite::append(
                request
                    .new_reply
                    .ok_or(NativeReplyError::InvalidSource)?
                    .identifier,
            )
        },
        NativeReplyOperation::Replace {
            ordinal,
            expected_reply_identifier,
            ..
        } => comment_storage_codec::CommentStorageReplyRewrite::replace(
            ordinal,
            expected_reply_identifier,
            request
                .new_reply
                .ok_or(NativeReplyError::InvalidSource)?
                .identifier,
        ),
        NativeReplyOperation::Remove {
            ordinal,
            expected_reply_identifier,
        } => comment_storage_codec::CommentStorageReplyRewrite::remove(
            ordinal,
            expected_reply_identifier,
        ),
    };
    let root_options = comment_options(root_payload, request.limits);
    // Preparation builds a parsed-field vector and retains the borrowed root
    // payload before it can return a report.  Reserve that bounded envelope
    // before entering the codec so a tight transaction budget fails before
    // private codec scratch is allocated.
    let root_scratch = root_payload
        .len()
        .checked_mul(2)
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(3, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(root_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_transaction_work(root_payload.len(), BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let prepared_root = comment_storage_codec::prepare_comment_storage_reply_rewrite(
        root_payload,
        root_rewrite,
        root_options,
    )
    .map_err(|_| NativeReplyError::Codec)?;
    let root_requirements = prepared_root.execution_requirements();
    budget
        .charge_allocations(root_requirements.allocations, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(root_requirements.scratch_bytes, BUDGET_PATH))
        .and_then(|_| budget.charge_retained_bytes(root_requirements.retained_bytes, BUDGET_PATH))
        .and_then(|_| budget.charge_wire_bytes(root_requirements.input_bytes, BUDGET_PATH))
        .and_then(|_| budget.charge_wire_fields(root_requirements.fields, BUDGET_PATH))
        .and_then(|_| budget.charge_wire_work(root_requirements.work_bytes, BUDGET_PATH))
        .and_then(|_| budget.charge_wire_nesting(root_requirements.max_depth, BUDGET_PATH))
        .and_then(|_| budget.charge_payload_references(root_requirements.references, BUDGET_PATH))
        .and_then(|_| {
            budget.charge_wire_reference_bytes(root_requirements.reference_bytes, BUDGET_PATH)
        })
        .and_then(|_| {
            budget.charge_wire_text_bytes(
                request.operation.text().map_or(0, |text| text.len()),
                BUDGET_PATH,
            )
        })
        .and_then(|_| budget.charge_output(root_requirements.output_bytes, BUDGET_PATH))
        .and_then(|_| {
            budget.charge_transaction_work(
                root_requirements
                    .input_bytes
                    .saturating_add(root_requirements.output_bytes),
                BUDGET_PATH,
            )
        })
        .map_err(|_| NativeReplyError::Limit)?;
    let root_rewritten = prepared_root
        .execute(root_requirements.exact_limits())
        .map_err(|_| NativeReplyError::Codec)?
        .into_bytes();
    let root_rewritten = patch_storage_uuid(&root_rewritten, request.new_root)?;
    let reply_field_mutation = match request.operation {
        NativeReplyOperation::Append { .. } => ReplyFieldMutation::Append,
        NativeReplyOperation::Replace { ordinal, .. } => ReplyFieldMutation::Replace(ordinal),
        NativeReplyOperation::Remove { ordinal, .. } => ReplyFieldMutation::Remove(ordinal),
    };
    let root_info = source
        .objects
        .get(
            object_index
                .get(root_identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?
        .archive_info
        .message_infos
        .get(request.root.message_index)
        .ok_or(NativeReplyError::InvalidSource)?
        .clone();
    let root_fields = reply_field_transitions(
        &root_info,
        author_identifier,
        &before_reply_ids,
        &after_reply_ids,
        reply_field_mutation,
        budget,
    )?;

    // Build the private candidate only after the source graph and prepared
    // codec ceilings have all passed.
    // Cloning the source is the first private candidate allocation.  Its
    // payloads are charged before the clone is requested; later archive
    // validation and codec reports debit their own exact traversals.
    let candidate_clone_scratch = decoded_bytes
        .checked_mul(2)
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(1, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(candidate_clone_scratch, BUDGET_PATH))
        .and_then(|_| budget.charge_retained_bytes(decoded_bytes, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut candidate = source.clone();
    let root_source_object = source
        .objects
        .get(
            object_index
                .get(root_identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(6, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(root_payload.len(), BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut root_archive = Archive::new();
    root_archive
        .objects
        .try_reserve_exact(1)
        .map_err(|_| NativeReplyError::Allocation)?;
    root_archive.objects.push(root_source_object.clone());
    root_archive.objects[0].archive_info.identifier = Some(request.new_root.identifier);
    replace_with_reference_transition(
        &mut root_archive,
        NativeReplyObjectRoute {
            member_index: 0,
            identifier: request.new_root.identifier,
            message_index: request.root.message_index,
        },
        root_rewritten,
        before_aggregate.as_slice(),
        after_aggregate.as_slice(),
        &root_fields,
        request.limits,
    )?;
    adjust_reply_field_infos(
        &mut root_archive,
        NativeReplyObjectRoute {
            member_index: 0,
            identifier: request.new_root.identifier,
            message_index: request.root.message_index,
        },
        reply_field_mutation,
        &after_reply_ids,
    )?;
    let new_root_object = root_archive
        .objects
        .pop()
        .ok_or(NativeReplyError::InvalidSource)?;
    candidate
        .objects
        .try_reserve_exact(2)
        .map_err(|_| NativeReplyError::Allocation)?;
    candidate.objects.push(new_root_object);

    // Clone/patch a direct leaf for append and replace.  The source root is a
    // valid leaf template only when no replies exist; otherwise use the first
    // validated direct leaf so a repeated field cannot leak into the new leaf.
    if let Some(new_reply) = request.new_reply {
        let template_route = match request.operation {
            NativeReplyOperation::Replace { ordinal, .. } => request
                .replies
                .get(ordinal)
                .copied()
                .ok_or(NativeReplyError::InvalidSource)?,
            NativeReplyOperation::Append { .. } => {
                request.replies.first().copied().unwrap_or(request.root)
            },
            NativeReplyOperation::Remove { .. } => {
                return Err(NativeReplyError::InvalidSource);
            },
        };
        let template_payload =
            storage_route_payload(source, &object_index, template_route)?.to_owned();
        let (template_snapshot, template_reply_ids) =
            decode_storage(&template_payload, budget, request.limits)?;
        if !template_reply_ids.is_empty() {
            return Err(NativeReplyError::UnsupportedDependency);
        }
        let text = request
            .operation
            .text()
            .ok_or(NativeReplyError::InvalidSource)?;
        let reply_payload =
            patch_leaf_text_and_uuid(&template_payload, template_snapshot.text(), text, new_reply)?;
        let mut reply_object = source
            .objects
            .get(
                object_index
                    .get(template_route.identifier)
                    .ok_or(NativeReplyError::InvalidSource)?,
            )
            .ok_or(NativeReplyError::InvalidSource)?
            .clone();
        reply_object.archive_info.identifier = Some(new_reply.identifier);
        let info = reply_object
            .archive_info
            .message_infos
            .get_mut(template_route.message_index)
            .ok_or(NativeReplyError::InvalidSource)?;
        info.length =
            u32::try_from(reply_payload.len()).map_err(|_| NativeReplyError::InvalidSource)?;
        let reply_author = template_snapshot
            .author()
            .map(|reference| reference.identifier());
        let reply_aggregate = expected_storage_references(reply_author, &[], budget, BUDGET_PATH)?;
        validate_storage_message_info(info, reply_author, &reply_aggregate, &[])?;
        info.object_references = reply_aggregate.ordered;
        reply_object.messages[template_route.message_index].data = reply_payload;
        candidate.objects.push(reply_object);
    }

    // Rewrite the selected root comment list.  New keys are always appended
    // after the root list's source next-list-id; the old key is decremented by
    // the already-validated global BNC census and removed only at refcount one.
    let new_key = allocate_comment_key(&lists, selected.next_list_id)?;
    let list_payload = &source
        .objects
        .get(
            object_index
                .get(request.list.identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?
        .messages
        .get(request.list.message_index)
        .ok_or(NativeReplyError::InvalidSource)?
        .data;
    let rewritten_list = rewrite_comment_list(
        list_payload,
        old_entry.key,
        old_entry.ref_count,
        old_entry.storage_identifier,
        new_key,
        request.new_root.identifier,
        budget,
        request.limits,
        BUDGET_PATH,
    )?;
    let list_ref_scratch = selected
        .entries
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(2, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(list_ref_scratch, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let before_list_refs = selected
        .entries
        .iter()
        .map(|entry| entry.storage_identifier)
        .collect::<Vec<_>>();
    let mut after_list_refs = before_list_refs.clone();
    if old_entry.ref_count == 1 {
        after_list_refs.retain(|identifier| *identifier != old_entry.storage_identifier);
    }
    after_list_refs.push(request.new_root.identifier);
    let list_info = source
        .objects
        .get(
            object_index
                .get(request.list.identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?
        .archive_info
        .message_infos
        .get(request.list.message_index)
        .ok_or(NativeReplyError::InvalidSource)?
        .clone();
    let list_mutation = ListFieldMutation {
        old_key: old_entry.key,
        old_ref_count: old_entry.ref_count,
        new_key,
    };
    let list_fields = list_field_transitions(
        &list_info,
        &selected.entries,
        list_mutation,
        &before_list_refs,
        &after_list_refs,
    )?;
    replace_with_reference_transition(
        &mut candidate,
        request.list,
        rewritten_list,
        &before_list_refs,
        &after_list_refs,
        &list_fields,
        request.limits,
    )?;
    adjust_list_field_infos(
        &mut candidate,
        request.list,
        list_mutation,
        request.new_root.identifier,
    )?;

    // The selected cell changes only its comment-list key.  The BNC helper
    // retains all other scalar/style/unknown bytes and handles offset-table
    // width changes in the enclosing row.
    let tile_payload = &source
        .objects
        .get(
            object_index
                .get(request.tile.identifier)
                .ok_or(NativeReplyError::InvalidSource)?,
        )
        .ok_or(NativeReplyError::InvalidSource)?
        .messages
        .get(request.tile.message_index)
        .ok_or(NativeReplyError::InvalidSource)?
        .data;
    let source_cell = popup_native::tile_cell(tile_payload, request.tile_row, request.tile_column)
        .map_err(|_| NativeReplyError::InvalidSource)?
        .to_owned();
    let mut cell = BncCell::parse(&source_cell).map_err(|_| NativeReplyError::InvalidSource)?;
    if cell.comment_identifier() != Some(request.comment_key) {
        return Err(NativeReplyError::InvalidSource);
    }
    cell.set_comment_identifier(Some(new_key));
    let target_cell = cell
        .try_encode_with_limit(request.limits.max_message_bytes())
        .map_err(|_| NativeReplyError::Limit)?;
    let rewritten_tile = popup_native::patch_tile_cell(
        tile_payload,
        request.tile_row,
        request.tile_column,
        &target_cell,
    )
    .map_err(|_| NativeReplyError::InvalidSource)?;
    popup_native::replace_message_preserving_header(
        &mut candidate,
        request.tile.identifier,
        request.tile.message_index,
        rewritten_tile,
        request.limits,
    )
    .map_err(map_popup_native_error)?;

    // Culling is performed only after an unknown-metadata global inbound
    // census.  A shared root stays alive while another cell still owns the old
    // key; a shared reply likewise stays alive when another root references it.
    let (root_referenced, old_reply_referenced) = archive_reference_presence(
        &candidate,
        root_identifier,
        old_reply_identifier,
        request.limits,
        budget,
        BUDGET_PATH,
    )?;
    let old_root_can_cull = old_entry.ref_count == 1 && !root_referenced;
    if old_root_can_cull {
        candidate
            .remove_object_checked_with_limits(root_identifier, request.limits)
            .map_err(|_| NativeReplyError::Archive)?;
    }
    if let Some(old_reply) = old_reply_identifier {
        if !old_reply_referenced {
            candidate
                .remove_object_checked_with_limits(old_reply, request.limits)
                .map_err(|_| NativeReplyError::Archive)?;
        }
    }

    // Verify that only the authorized graph objects changed.  This catches a
    // broad list/message replacement before bytes leave the private native
    // owner.
    // The locality walker compares every object and may retain temporary
    // identifier sets, so charge the candidate archive envelope before it is
    // entered as well as before its later semantic scans.
    let candidate_decoded_bytes = preflight_archive(&candidate, budget, BUDGET_PATH)?;
    let candidate_index =
        validate_archive_headers(&candidate, request.limits, budget, BUDGET_PATH)?;
    let changed_capacity = 6usize;
    let changed_scratch = changed_capacity
        .checked_mul(size_of::<u64>())
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(2, BUDGET_PATH)
        .and_then(|_| budget.charge_scratch_bytes(changed_scratch, BUDGET_PATH))
        .map_err(|_| NativeReplyError::Limit)?;
    let mut changed = vec![
        request.tile.identifier,
        request.list.identifier,
        request.new_root.identifier,
    ];
    if let Some(new_reply) = request.new_reply {
        changed.push(new_reply.identifier);
    }
    if let Some(old_reply) = old_reply_identifier {
        changed.push(old_reply);
    }
    if old_root_can_cull {
        changed.push(root_identifier);
    }
    verify_reply_archive_object_locality(
        source,
        &object_index,
        &candidate,
        &candidate_index,
        &changed,
    )?;
    let candidate_counts = collect_bnc_counts(&candidate, budget, BUDGET_PATH)?;
    let candidate_lists = collect_comment_lists(&candidate, budget, request.limits, BUDGET_PATH)?;
    let candidate_storage_ids = storage_id_set(&candidate_lists, budget, BUDGET_PATH)?;
    validate_list_refcounts(
        &candidate_lists,
        &candidate_counts,
        &candidate_storage_ids,
        budget,
    )?;

    // Archive serialization allocates its output buffer.  Compute both
    // source and candidate encoded bounds first, then charge the complete
    // output/scratch envelope before requesting either Vec.
    let candidate_encoded_len = candidate
        .encoded_len_with_limits(request.limits)
        .map_err(|_| NativeReplyError::Archive)?;
    let source_encoded_len = source
        .encoded_len_with_limits(request.limits)
        .map_err(|_| NativeReplyError::Archive)?;
    let serialization_scratch = candidate_encoded_len
        .checked_add(source_encoded_len)
        .ok_or(NativeReplyError::InvalidSource)?;
    budget
        .charge_allocations(3, BUDGET_PATH)
        .and_then(|_| budget.charge_output(candidate_encoded_len, BUDGET_PATH))
        .and_then(|_| budget.charge_retained_bytes(candidate_encoded_len, BUDGET_PATH))
        .and_then(|_| budget.charge_scratch_bytes(serialization_scratch, BUDGET_PATH))
        .and_then(|_| {
            budget.charge_transaction_work(
                candidate_encoded_len.saturating_add(source_encoded_len),
                BUDGET_PATH,
            )
        })
        .map_err(|_| NativeReplyError::Limit)?;
    let member_bytes = candidate
        .to_bytes_with_limits(request.limits)
        .map_err(|_| NativeReplyError::Archive)?;
    if member_bytes.len() > candidate_encoded_len {
        return Err(NativeReplyError::Archive);
    }
    let source_bytes = source
        .to_bytes_with_limits(request.limits)
        .map_err(|_| NativeReplyError::Archive)?;
    if source_bytes.len() > source_encoded_len {
        return Err(NativeReplyError::Archive);
    }
    let source_cell = popup_native::tile_cell(tile_payload, request.tile_row, request.tile_column)
        .map_err(|_| NativeReplyError::InvalidSource)?
        .to_owned();
    budget
        .charge_allocations(6, BUDGET_PATH)
        .map_err(|_| NativeReplyError::Limit)?;
    let mut member_edits = Vec::new();
    member_edits
        .try_reserve_exact(1)
        .map_err(|_| NativeReplyError::Allocation)?;
    if source_bytes == member_bytes {
        return Err(NativeReplyError::InvalidSource);
    }
    let candidate_byte_count = member_bytes.len();
    member_edits.push(NativeReplyMemberEdit {
        component_index: member.component_index,
        member_name: member.member_name.to_owned(),
        member_bytes,
    });
    let added_object_capacity = usize::from(request.new_reply.is_some()).saturating_add(1);
    let mut added_object_identifiers = Vec::new();
    added_object_identifiers
        .try_reserve_exact(added_object_capacity)
        .map_err(|_| NativeReplyError::Allocation)?;
    added_object_identifiers.push(request.new_root.identifier);
    let mut removed_object_identifiers = Vec::new();
    removed_object_identifiers
        .try_reserve_exact(2)
        .map_err(|_| NativeReplyError::Allocation)?;
    if let Some(new_reply) = request.new_reply {
        added_object_identifiers.push(new_reply.identifier);
    }
    if old_root_can_cull {
        removed_object_identifiers.push(root_identifier);
    }
    if let Some(old_reply) = old_reply_identifier {
        if candidate_index.get(old_reply).is_none() {
            removed_object_identifiers.push(old_reply);
        }
    }
    let mut added_edges = Vec::new();
    added_edges
        .try_reserve_exact(after_reply_ids.len())
        .map_err(|_| NativeReplyError::Allocation)?;
    let mut removed_edges = Vec::new();
    removed_edges
        .try_reserve_exact(if old_root_can_cull {
            before_reply_ids.len()
        } else {
            0
        })
        .map_err(|_| NativeReplyError::Allocation)?;
    for identifier in &after_reply_ids {
        added_edges.push(NativeReplyReferenceEdge {
            source_identifier: request.new_root.identifier,
            target_identifier: *identifier,
        });
    }
    if old_root_can_cull {
        for identifier in &before_reply_ids {
            removed_edges.push(NativeReplyReferenceEdge {
                source_identifier: root_identifier,
                target_identifier: *identifier,
            });
        }
    }
    let reply_count = before_reply_ids
        .len()
        .checked_add(after_reply_ids.len())
        .ok_or(NativeReplyError::InvalidSource)?;
    Ok(NativeReplyOutput {
        member_edits,
        source_cell,
        target_cell,
        before_root_identifier: root_identifier,
        after_root_identifier: request.new_root.identifier,
        before_reply_identifiers: before_reply_ids,
        after_reply_identifiers: after_reply_ids,
        added_object_identifiers,
        removed_object_identifiers,
        added_edges,
        removed_edges,
        touched_components: vec![member.component_index],
        reports: NativeReplyReports {
            source_bytes: source_bytes.len(),
            candidate_bytes: candidate_byte_count,
            fields: 0,
            work_bytes: decoded_bytes.saturating_add(candidate_decoded_bytes),
            references: before_aggregate
                .as_slice()
                .len()
                .saturating_add(after_aggregate.as_slice().len()),
            replies: reply_count,
            allocations: 8,
            scratch_bytes: decoded_bytes,
            retained_bytes: candidate_byte_count,
        },
    })
}

#[cfg(test)]
mod tests {
    use super::{IdIndex, NativeReplyError, ReplyProbe, StorageReferenceAggregate};

    #[test]
    fn reply_probe_handles_reverse_adversarial_order_and_duplicate() {
        const COUNT: usize = 16_384;
        let mut probe = ReplyProbe::with_capacity(COUNT)
            .expect("the bounded reply probe should reserve both ID collections");
        for identifier in (1..=u64::try_from(COUNT).expect("count fits in u64")).rev() {
            probe.record_identifier(identifier, None, None);
        }
        assert!(!probe.allocation_failed);
        assert!(!probe.unsupported_reference);
        assert_eq!(probe.identifiers.len(), COUNT);
        assert_eq!(
            probe.identifiers[0],
            u64::try_from(COUNT).expect("count fits in u64")
        );
        assert_eq!(probe.identifiers[COUNT - 1], 1);

        probe.record_identifier(
            u64::try_from(COUNT / 2).expect("count fits in u64"),
            None,
            None,
        );
        assert!(probe.unsupported_reference);
        assert!(!probe.allocation_failed);
        assert_eq!(probe.identifiers.len(), COUNT);
    }

    #[test]
    fn reply_probe_capacity_exhaustion_is_staged_as_allocation() {
        let mut probe = ReplyProbe::with_capacity(1).expect("one ID should reserve");
        probe.record_identifier(11, None, None);
        probe.record_identifier(12, None, None);
        assert!(probe.allocation_failed);
        assert_eq!(probe.identifiers, [11]);
        assert!(!probe.unsupported_reference);
    }

    #[test]
    fn id_index_is_direct_and_duplicate_typed() {
        const COUNT: usize = 16_384;
        let mut index = IdIndex::try_with_capacity(COUNT).expect("the ID index should reserve");
        for position in (0..COUNT).rev() {
            let identifier = u64::try_from(position + 1).expect("position fits in u64");
            assert_eq!(index.insert(identifier, position), None);
        }
        assert_eq!(index.len(), COUNT);
        for position in [0, COUNT / 2, COUNT - 1] {
            let identifier = u64::try_from(position + 1).expect("position fits in u64");
            assert_eq!(index.get(identifier), Some(position));
        }
        assert_eq!(index.insert(1, usize::MAX), Some(0));
    }

    #[test]
    fn storage_reference_aggregate_preserves_order_and_rejects_duplicates() {
        let aggregate = StorageReferenceAggregate::try_from_parts(Some(9), &[30, 10, 20])
            .expect("unique references should be accepted");
        assert_eq!(aggregate.as_slice(), [9, 30, 10, 20]);
        assert!(aggregate.contains(&10));
        assert!(!aggregate.contains(&99));
        assert_eq!(
            StorageReferenceAggregate::try_from_parts(Some(9), &[30, 9]),
            Err(NativeReplyError::InvalidSource)
        );
    }
}
