//! Private invalidation of stale Keynote slide-preview caches.

mod slide_number;

pub(super) use slide_number::{
    exact_slide_number_delta_with_allowance, set_slide_number_with_report,
};

use std::collections::{HashMap, HashSet};

use litchi_iwa_common::{
    Error as WireError, LimitKind as WireLimitKind, WireLimits, decode_varint_from_bytes,
    varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::archive::DataReferencePruning;
use litchi_iwa_core::{ArchiveObject, Limits as ArchiveLimits, RawMessage};
use thiserror::Error;

const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const DATABASE_THUMBNAIL_FIELD: u32 = 3;
const DATABASE_THUMBNAILS_FIELD: u32 = 9;
const THUMBNAIL_SIZES_FIELD: u32 = 10;
const THUMBNAILS_DIRTY_FIELD: u32 = 14;
const THUMBNAILS_FIELD: u32 = 16;
const THUMBNAIL_DIGESTS_FIELD: u32 = 25;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum InvalidationDirection {
    Forward,
    Inverse,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct InvalidationAllowance {
    work: usize,
    references: usize,
}

impl InvalidationAllowance {
    pub(super) const fn new(work: usize, references: usize) -> Self {
        Self { work, references }
    }

    const UNLIMITED: Self = Self::new(usize::MAX, usize::MAX);
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum InvalidationBudgetKind {
    References,
    Work,
}

#[derive(Debug, Error)]
pub(super) enum BudgetedInvalidationError {
    #[error(transparent)]
    Invalidation(#[from] InvalidationError),
    #[error(
        "Keynote slide-preview {kind:?} budget exceeded: observed {observed}, maximum {maximum}"
    )]
    BudgetExceeded {
        kind: InvalidationBudgetKind,
        observed: usize,
        maximum: usize,
    },
}

/// A private, content-redacted failure at the slide-preview boundary.
#[derive(Debug, Error)]
pub(super) enum InvalidationError {
    #[error("the selected Keynote slide node cannot be invalidated safely")]
    InvalidSource,
    #[error(transparent)]
    Wire(#[from] WireError),
    #[error(transparent)]
    Archive(#[from] litchi_iwa_core::Error),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct InvalidationReport {
    bytes: usize,
    fields: usize,
    references: usize,
    work: usize,
    allowance: InvalidationAllowance,
    exceeded: Option<(InvalidationBudgetKind, usize, usize)>,
}

impl Default for InvalidationReport {
    fn default() -> Self {
        Self::with_allowance(InvalidationAllowance::UNLIMITED)
    }
}

impl InvalidationReport {
    const fn with_allowance(allowance: InvalidationAllowance) -> Self {
        Self {
            bytes: 0,
            fields: 0,
            references: 0,
            work: 0,
            allowance,
            exceeded: None,
        }
    }

    pub(super) const fn references(self) -> usize {
        self.references
    }

    pub(super) const fn work(self) -> usize {
        self.work
    }

    fn charge_field(&mut self, limits: WireLimits) -> Result<(), InvalidationError> {
        self.fields = self.fields.saturating_add(1);
        if self.fields > limits.max_fields() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::Fields,
                observed: self.fields,
                limit: limits.max_fields(),
            }
            .into());
        }
        Ok(())
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), InvalidationError> {
        let observed = self
            .references
            .checked_add(amount)
            .ok_or(InvalidationError::InvalidSource)?;
        self.references = observed;
        if observed > self.allowance.references {
            self.exceeded.get_or_insert((
                InvalidationBudgetKind::References,
                observed,
                self.allowance.references,
            ));
            return Err(InvalidationError::InvalidSource);
        }
        Ok(())
    }

    fn charge_scan(&mut self, amount: usize, limits: WireLimits) -> Result<(), InvalidationError> {
        if amount > limits.max_input_bytes() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::InputBytes,
                observed: amount,
                limit: limits.max_input_bytes(),
            }
            .into());
        }
        self.bytes = self
            .bytes
            .checked_add(amount)
            .ok_or(InvalidationError::InvalidSource)?;
        self.charge_work(amount, limits)
    }

    fn charge_work(&mut self, amount: usize, limits: WireLimits) -> Result<(), InvalidationError> {
        self.work = self.work.saturating_add(amount);
        if self.work > self.allowance.work {
            self.exceeded.get_or_insert((
                InvalidationBudgetKind::Work,
                self.work,
                self.allowance.work,
            ));
            return Err(InvalidationError::InvalidSource);
        }
        if self.work > limits.max_rewrite_work() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::RewriteWork,
                observed: self.work,
                limit: limits.max_rewrite_work(),
            }
            .into());
        }
        Ok(())
    }

    fn budget_error(self) -> Option<BudgetedInvalidationError> {
        self.exceeded.map(
            |(kind, observed, maximum)| BudgetedInvalidationError::BudgetExceeded {
                kind,
                observed,
                maximum,
            },
        )
    }
}

#[derive(Clone, Copy)]
struct RawField<'source> {
    number: u32,
    wire_type: u8,
    raw: &'source [u8],
    key: &'source [u8],
    length_prefix: &'source [u8],
    payload: &'source [u8],
}

#[derive(Clone, Copy, Default)]
struct ExactPayloadState {
    child_references: usize,
    preview_object_references: usize,
    preview_data_references: usize,
    dirty: Option<bool>,
    has_preview_fields: bool,
}

#[derive(Clone, Copy, Default)]
struct PayloadReferenceOccurrences {
    child: usize,
    preview_data: usize,
    preview_object: usize,
    retained_object: usize,
}

#[derive(Default)]
struct ExactPayloadScan {
    state: ExactPayloadState,
    references: HashMap<u64, PayloadReferenceOccurrences>,
}

#[derive(Clone, Copy, Default)]
struct MetadataReferenceOccurrences {
    aggregate_object: usize,
    database_thumbnail_object: usize,
    other_field_object: usize,
    aggregate_data: usize,
    preview_field_data: usize,
    nonpreview_field_data: usize,
}

struct ScanBudget {
    limits: WireLimits,
    report: InvalidationReport,
}

impl ScanBudget {
    const fn new(limits: WireLimits) -> Self {
        Self::with_allowance(limits, InvalidationAllowance::UNLIMITED)
    }

    const fn with_allowance(limits: WireLimits, allowance: InvalidationAllowance) -> Self {
        Self {
            limits,
            report: InvalidationReport::with_allowance(allowance),
        }
    }

    fn parse<'source>(
        &mut self,
        source: &'source [u8],
    ) -> Result<WireView<'source>, InvalidationError> {
        // `WireView` deliberately rejects every deprecated protobuf group.
        // Changed transactions therefore fail closed before mutation instead
        // of trying to preserve an unbounded group subtree whose schema and
        // cache ownership cannot be proven here.
        self.report.charge_scan(source.len(), self.limits)?;
        let view = WireView::parse_with_limits(source, self.limits)?;
        let observed = self.report.fields.saturating_add(view.len());
        if observed > self.limits.max_fields() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::Fields,
                observed,
                limit: self.limits.max_fields(),
            }
            .into());
        }
        self.report.fields = observed;
        Ok(view)
    }

    fn add_work(&mut self, amount: usize) -> Result<(), InvalidationError> {
        self.report.charge_work(amount, self.limits)
    }

    fn add_references(&mut self, amount: usize) -> Result<(), InvalidationError> {
        self.report.charge_references(amount)
    }
}

struct PayloadRewrite {
    data: Vec<u8>,
    removed_object_references: Vec<u64>,
    removed_data_references: Vec<u64>,
}

struct PayloadScan<'source> {
    view: WireView<'source>,
    retained_bytes: usize,
    removed_object_references: Vec<u64>,
    removed_data_references: Vec<u64>,
    retained_child_object_references: Vec<u64>,
    dirty: Option<bool>,
    has_preview_fields: bool,
}

/// Remove every rendered preview owned by exactly one selected slide node.
///
/// The caller owns graph selection; this helper deliberately accepts neither
/// an object identifier nor a package component name. The mutation is limited
/// to the selected type-4 payload and its selected `MessageInfo` metadata.
pub(super) fn invalidate(
    object: &mut ArchiveObject,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
) -> Result<(), InvalidationError> {
    let mut budget = ScanBudget::new(wire_limits);
    invalidate_if_needed_inner(object, archive_limits, &mut budget)?;
    Ok(())
}

pub(super) fn invalidate_if_needed_with_report(
    object: &mut ArchiveObject,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    allowance: InvalidationAllowance,
) -> Result<(bool, InvalidationReport), BudgetedInvalidationError> {
    let mut budget = ScanBudget::with_allowance(wire_limits, allowance);
    let result = invalidate_if_needed_inner(object, archive_limits, &mut budget);
    if let Some(error) = budget.report.budget_error() {
        return Err(error);
    }
    Ok((result?, budget.report))
}

fn invalidate_if_needed_inner(
    object: &mut ArchiveObject,
    archive_limits: ArchiveLimits,
    budget: &mut ScanBudget,
) -> Result<bool, InvalidationError> {
    charge_scan_archive_structure(object, false, budget)?;
    object.validate_with_limits(archive_limits)?;
    charge_scan_message_selection(object, budget)?;
    let message_index = selected_message_index(object)?;
    budget.add_work(1)?;
    validate_metadata_topology(object, message_index)?;
    let original = object
        .messages
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    let scan = scan_payload(&original.data, budget)?;
    validate_reference_ownership(
        object,
        message_index,
        &scan.removed_object_references,
        &scan.retained_child_object_references,
        &scan.removed_data_references,
        budget,
    )?;
    if !scan.has_preview_fields && scan.dirty == Some(true) {
        return Ok(false);
    }
    let rewrite = rewrite_scanned_payload(scan, budget)?;
    validate_rewrite(&original.data, &rewrite.data, budget)?;
    let PayloadRewrite {
        data,
        removed_object_references,
        removed_data_references,
    } = rewrite;

    // The core replacement revalidates the enclosing object and walks its
    // metadata again before its atomic commit.
    charge_scan_archive_structure(object, true, budget)?;
    object.replace_message_pruning_references_preserving_header_with_limits(
        message_index,
        RawMessage {
            type_: SLIDE_NODE_MESSAGE_TYPE,
            data,
        },
        &removed_object_references,
        DataReferencePruning::Selected(&removed_data_references),
        archive_limits,
    )?;

    Ok(true)
}

/// Verify the raw preview state without decoding or cloning the slide node.
pub(super) fn is_invalidated(
    object: &ArchiveObject,
    wire_limits: WireLimits,
) -> Result<bool, InvalidationError> {
    let mut budget = ScanBudget::new(wire_limits);
    is_invalidated_inner(object, &mut budget)
}

fn is_invalidated_inner(
    object: &ArchiveObject,
    budget: &mut ScanBudget,
) -> Result<bool, InvalidationError> {
    charge_scan_message_selection(object, budget)?;
    let message_index = selected_message_index(object)?;
    budget.add_work(1)?;
    validate_metadata_topology(object, message_index)?;
    let payload = &object
        .messages
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?
        .data;
    let scan = scan_payload(payload, budget)?;
    validate_reference_ownership(
        object,
        message_index,
        &scan.removed_object_references,
        &scan.retained_child_object_references,
        &scan.removed_data_references,
        budget,
    )?;
    if scan.has_preview_fields {
        return Ok(false);
    }
    Ok(scan.dirty == Some(true))
}

#[allow(
    dead_code,
    reason = "kept for private callers that do not share a transaction budget"
)]
pub(super) fn exact_invalidation_delta(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    direction: InvalidationDirection,
    wire_limits: WireLimits,
) -> Result<(bool, InvalidationReport), InvalidationError> {
    let mut report = InvalidationReport::default();
    let matches =
        exact_invalidation_delta_inner(source, candidate, direction, wire_limits, &mut report)?;
    Ok((matches, report))
}

pub(super) fn exact_invalidation_delta_with_allowance(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    direction: InvalidationDirection,
    wire_limits: WireLimits,
    allowance: InvalidationAllowance,
) -> Result<(bool, InvalidationReport), BudgetedInvalidationError> {
    let mut report = InvalidationReport::with_allowance(allowance);
    let result =
        exact_invalidation_delta_inner(source, candidate, direction, wire_limits, &mut report);
    if let Some(error) = report.budget_error() {
        return Err(error);
    }
    Ok((result?, report))
}

fn exact_invalidation_delta_inner(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    direction: InvalidationDirection,
    wire_limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    match direction {
        InvalidationDirection::Forward => {
            exact_forward_delta(source, candidate, wire_limits, report)
        },
        InvalidationDirection::Inverse => {
            exact_forward_delta(candidate, source, wire_limits, report)
        },
    }
}

fn charge_scan_message_selection(
    object: &ArchiveObject,
    budget: &mut ScanBudget,
) -> Result<(), InvalidationError> {
    budget.add_work(1)?;
    for _message in &object.messages {
        budget.add_work(1)?;
    }
    for _info in &object.archive_info.message_infos {
        budget.add_work(1)?;
    }
    Ok(())
}

fn charge_scan_archive_structure(
    object: &ArchiveObject,
    compare_header: bool,
    budget: &mut ScanBudget,
) -> Result<(), InvalidationError> {
    charge_archive_contents(object, compare_header, |work, references| {
        budget.add_work(work)?;
        budget.add_references(references)
    })
}

fn charge_report_message_selection(
    object: &ArchiveObject,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<(), InvalidationError> {
    report.charge_work(1, limits)?;
    for _message in &object.messages {
        report.charge_work(1, limits)?;
    }
    for _info in &object.archive_info.message_infos {
        report.charge_work(1, limits)?;
    }
    Ok(())
}

fn charge_report_archive_structure(
    object: &ArchiveObject,
    compare_header: bool,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<(), InvalidationError> {
    charge_archive_contents(object, compare_header, |work, references| {
        report.charge_work(work, limits)?;
        report.charge_references(references)
    })
}

fn charge_exact_archive_structure(
    object: &ArchiveObject,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<(), InvalidationError> {
    charge_report_archive_structure(object, false, limits, report)
}

fn charge_exact_message_info_structure(
    info: &litchi_iwa_core::MessageInfo,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<(), InvalidationError> {
    charge_message_info_contents(info, &mut |work, references| {
        report.charge_work(work, limits)?;
        report.charge_references(references)
    })
}

fn charge_archive_contents(
    object: &ArchiveObject,
    compare_header: bool,
    mut charge: impl FnMut(usize, usize) -> Result<(), InvalidationError>,
) -> Result<(), InvalidationError> {
    charge(1, 0)?;
    if compare_header {
        charge(
            usize::try_from(object.header_length)
                .map_err(|_conversion| InvalidationError::InvalidSource)?,
            0,
        )?;
    }
    for message in &object.messages {
        charge(1, 0)?;
        charge(message.data.len(), 0)?;
    }
    for info in &object.archive_info.message_infos {
        charge_message_info_contents(info, &mut charge)?;
    }
    Ok(())
}

fn charge_message_info_contents(
    info: &litchi_iwa_core::MessageInfo,
    charge: &mut impl FnMut(usize, usize) -> Result<(), InvalidationError>,
) -> Result<(), InvalidationError> {
    charge(1, 0)?;
    charge(info.versions.len(), 0)?;
    charge(info.object_references.len(), info.object_references.len())?;
    charge(info.data_references.len(), info.data_references.len())?;
    charge(info.diff_merge_version.len(), 0)?;
    if let Some(path) = &info.diff_field_path {
        charge(1, 0)?;
        charge(path.as_slice().len(), 0)?;
    }
    for path in &info.fields_to_remove {
        charge(1, 0)?;
        charge(path.as_slice().len(), 0)?;
    }
    charge(info.diff_read_version.len(), 0)?;
    for field in &info.field_infos {
        charge(1, 0)?;
        charge(field.path.as_slice().len(), 0)?;
        charge(field.object_references.len(), field.object_references.len())?;
        charge(field.data_references.len(), field.data_references.len())?;
        charge(field.known_field_version.len(), 0)?;
        if let Some(identifier) = &field.known_field_feature_identifier {
            charge(identifier.len(), 0)?;
        }
    }
    Ok(())
}

fn walk_raw_fields(
    source: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
    mut visitor: impl FnMut(RawField<'_>, &mut InvalidationReport) -> Result<(), InvalidationError>,
) -> Result<(), InvalidationError> {
    report.charge_scan(source.len(), limits)?;
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_raw_field(source, offset)?;
        report.charge_field(limits)?;
        offset = offset
            .checked_add(field.raw.len())
            .ok_or(InvalidationError::InvalidSource)?;
        visitor(field, report)?;
    }
    Ok(())
}

fn parse_raw_field(source: &[u8], start: usize) -> Result<RawField<'_>, InvalidationError> {
    let remaining = source
        .get(start..)
        .ok_or(InvalidationError::InvalidSource)?;
    let (encoded_key, key_length) =
        decode_varint_from_bytes(remaining).map_err(|_error| InvalidationError::InvalidSource)?;
    let number =
        u32::try_from(encoded_key >> 3).map_err(|_conversion| InvalidationError::InvalidSource)?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(InvalidationError::InvalidSource);
    }
    let wire_type =
        u8::try_from(encoded_key & 7).map_err(|_conversion| InvalidationError::InvalidSource)?;
    let key_end = start
        .checked_add(key_length)
        .ok_or(InvalidationError::InvalidSource)?;
    let mut payload_start = key_end;
    let mut length_prefix_end = key_end;
    let end = match wire_type {
        0 => {
            let (_, width) = decode_varint_from_bytes(
                source
                    .get(key_end..)
                    .ok_or(InvalidationError::InvalidSource)?,
            )
            .map_err(|_error| InvalidationError::InvalidSource)?;
            key_end
                .checked_add(width)
                .ok_or(InvalidationError::InvalidSource)?
        },
        1 => key_end
            .checked_add(8)
            .ok_or(InvalidationError::InvalidSource)?,
        2 => {
            let (length, width) = decode_varint_from_bytes(
                source
                    .get(key_end..)
                    .ok_or(InvalidationError::InvalidSource)?,
            )
            .map_err(|_error| InvalidationError::InvalidSource)?;
            payload_start = key_end
                .checked_add(width)
                .ok_or(InvalidationError::InvalidSource)?;
            length_prefix_end = payload_start;
            payload_start
                .checked_add(
                    usize::try_from(length)
                        .map_err(|_conversion| InvalidationError::InvalidSource)?,
                )
                .ok_or(InvalidationError::InvalidSource)?
        },
        5 => key_end
            .checked_add(4)
            .ok_or(InvalidationError::InvalidSource)?,
        _ => return Err(InvalidationError::InvalidSource),
    };
    let raw = source
        .get(start..end)
        .ok_or(InvalidationError::InvalidSource)?;
    Ok(RawField {
        number,
        wire_type,
        raw,
        key: source
            .get(start..key_end)
            .ok_or(InvalidationError::InvalidSource)?,
        length_prefix: source
            .get(key_end..length_prefix_end)
            .ok_or(InvalidationError::InvalidSource)?,
        payload: source
            .get(payload_start..end)
            .ok_or(InvalidationError::InvalidSource)?,
    })
}

fn validate_exact_canonical_field(field: RawField<'_>) -> Result<(), InvalidationError> {
    let expected_key = (u64::from(field.number) << 3) | u64::from(field.wire_type);
    if field.key.len() != encoded_len(expected_key) {
        return Err(InvalidationError::InvalidSource);
    }
    if field.wire_type == 2 {
        let length = u64::try_from(field.payload.len())
            .map_err(|_conversion| InvalidationError::InvalidSource)?;
        if field.length_prefix.len() != encoded_len(length) {
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(())
}

fn exact_varint(field: RawField<'_>) -> Result<u64, InvalidationError> {
    if field.wire_type != 0 {
        return Err(InvalidationError::InvalidSource);
    }
    let (value, width) = decode_varint_from_bytes(field.payload)
        .map_err(|_error| InvalidationError::InvalidSource)?;
    if width != field.payload.len() || width != encoded_len(value) {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(value)
}

fn exact_reference_identifier(
    source: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<u64, InvalidationError> {
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    walk_raw_fields(source, limits, report, |field, _report| {
        match field.number {
            1 => {
                validate_exact_canonical_field(field)?;
                let value = exact_varint(field)?;
                if value == 0 || identifier.replace(value).is_some() {
                    return Err(InvalidationError::InvalidSource);
                }
            },
            2 => {
                validate_exact_canonical_field(field)?;
                let value = decode_canonical_int32(exact_varint(field)?)?;
                if deprecated_type.replace(value).is_some() {
                    return Err(InvalidationError::InvalidSource);
                }
            },
            3 => {
                validate_exact_canonical_field(field)?;
                let value = exact_varint(field)?;
                if value != 0 || external.replace(false).is_some() {
                    return Err(InvalidationError::InvalidSource);
                }
            },
            _ => {},
        }
        Ok(())
    })?;
    identifier.ok_or(InvalidationError::InvalidSource)
}

fn exact_data_reference_identifier(
    source: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<u64, InvalidationError> {
    let mut identifier = None;
    walk_raw_fields(source, limits, report, |field, _report| {
        if field.number == 1 {
            validate_exact_canonical_field(field)?;
            let value = exact_varint(field)?;
            if value == 0 || identifier.replace(value).is_some() {
                return Err(InvalidationError::InvalidSource);
            }
        }
        Ok(())
    })?;
    identifier.ok_or(InvalidationError::InvalidSource)
}

fn payload_reference_occurrences<'references>(
    references: &'references mut HashMap<u64, PayloadReferenceOccurrences>,
    identifier: u64,
    resource: &'static str,
) -> Result<&'references mut PayloadReferenceOccurrences, InvalidationError> {
    if references.len() == references.capacity() {
        references
            .try_reserve(1)
            .map_err(|_allocation| WireError::Allocation {
                resource,
                amount: references.len().saturating_add(1),
            })?;
    }
    Ok(references.entry(identifier).or_default())
}

fn increment_occurrence(occurrence: &mut usize) -> Result<(), InvalidationError> {
    *occurrence = occurrence
        .checked_add(1)
        .ok_or(InvalidationError::InvalidSource)?;
    Ok(())
}

fn scan_exact_payload(
    source: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<ExactPayloadScan, InvalidationError> {
    let mut scan = ExactPayloadScan::default();
    let mut database_thumbnail_seen = false;
    let mut slide_seen = false;
    walk_raw_fields(source, limits, report, |field, scan_report| {
        match field.number {
            1 => {
                validate_exact_canonical_field(field)?;
                if field.wire_type != 2 {
                    return Err(InvalidationError::InvalidSource);
                }
                let identifier = exact_reference_identifier(field.payload, limits, scan_report)?;
                scan_report.charge_references(1)?;
                scan_report.charge_work(1, limits)?;
                let occurrences = payload_reference_occurrences(
                    &mut scan.references,
                    identifier,
                    "Keynote exact payload reference index",
                )?;
                increment_occurrence(&mut occurrences.child)?;
                increment_occurrence(&mut occurrences.retained_object)?;
                scan.state.child_references = scan
                    .state
                    .child_references
                    .checked_add(1)
                    .ok_or(InvalidationError::InvalidSource)?;
            },
            2 => {
                validate_exact_canonical_field(field)?;
                if field.wire_type != 2 || std::mem::replace(&mut slide_seen, true) {
                    return Err(InvalidationError::InvalidSource);
                }
                let identifier = exact_reference_identifier(field.payload, limits, scan_report)?;
                scan_report.charge_references(1)?;
                scan_report.charge_work(1, limits)?;
                let occurrences = payload_reference_occurrences(
                    &mut scan.references,
                    identifier,
                    "Keynote exact payload reference index",
                )?;
                increment_occurrence(&mut occurrences.retained_object)?;
            },
            DATABASE_THUMBNAIL_FIELD | DATABASE_THUMBNAILS_FIELD => {
                validate_exact_canonical_field(field)?;
                if field.wire_type != 2
                    || (field.number == DATABASE_THUMBNAIL_FIELD
                        && std::mem::replace(&mut database_thumbnail_seen, true))
                {
                    return Err(InvalidationError::InvalidSource);
                }
                let identifier = exact_reference_identifier(field.payload, limits, scan_report)?;
                scan_report.charge_references(1)?;
                scan_report.charge_work(1, limits)?;
                let occurrences = payload_reference_occurrences(
                    &mut scan.references,
                    identifier,
                    "Keynote exact payload reference index",
                )?;
                increment_occurrence(&mut occurrences.preview_object)?;
                scan.state.preview_object_references = scan
                    .state
                    .preview_object_references
                    .checked_add(1)
                    .ok_or(InvalidationError::InvalidSource)?;
                scan.state.has_preview_fields = true;
            },
            THUMBNAIL_SIZES_FIELD | THUMBNAIL_DIGESTS_FIELD => {
                validate_exact_canonical_field(field)?;
                if field.wire_type != 2 {
                    return Err(InvalidationError::InvalidSource);
                }
                scan.state.has_preview_fields = true;
            },
            THUMBNAILS_DIRTY_FIELD => {
                validate_exact_canonical_field(field)?;
                let value = exact_varint(field)?;
                if value > 1 || scan.state.dirty.replace(value == 1).is_some() {
                    return Err(InvalidationError::InvalidSource);
                }
            },
            THUMBNAILS_FIELD => {
                validate_exact_canonical_field(field)?;
                if field.wire_type != 2 {
                    return Err(InvalidationError::InvalidSource);
                }
                let identifier =
                    exact_data_reference_identifier(field.payload, limits, scan_report)?;
                scan_report.charge_references(1)?;
                scan_report.charge_work(1, limits)?;
                let occurrences = payload_reference_occurrences(
                    &mut scan.references,
                    identifier,
                    "Keynote exact payload reference index",
                )?;
                increment_occurrence(&mut occurrences.preview_data)?;
                scan.state.preview_data_references = scan
                    .state
                    .preview_data_references
                    .checked_add(1)
                    .ok_or(InvalidationError::InvalidSource)?;
                scan.state.has_preview_fields = true;
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok(scan)
}

fn exact_forward_delta(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    charge_report_message_selection(source, limits, report)?;
    charge_report_message_selection(candidate, limits, report)?;
    let source_index = selected_message_index(source)?;
    let candidate_index = selected_message_index(candidate)?;
    report.charge_work(2, limits)?;
    validate_metadata_topology(source, source_index)?;
    validate_metadata_topology(candidate, candidate_index)?;
    if source_index != candidate_index
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
    {
        return Ok(false);
    }
    let source_payload = &source.messages[source_index].data;
    let candidate_payload = &candidate.messages[candidate_index].data;
    let source_scan = scan_exact_payload(source_payload, limits, report)?;
    let candidate_scan = scan_exact_payload(candidate_payload, limits, report)?;
    validate_exact_ownership(
        source,
        source_index,
        &source_scan.references,
        source_scan.state,
        limits,
        report,
    )?;
    validate_exact_ownership(
        candidate,
        candidate_index,
        &candidate_scan.references,
        candidate_scan.state,
        limits,
        report,
    )?;
    // Equality traverses every MessageInfo and FieldInfo, including empty
    // metadata that has no reference or path bytes to charge otherwise.
    charge_exact_archive_structure(source, limits, report)?;
    charge_exact_archive_structure(candidate, limits, report)?;
    let equal =
        source.archive_info == candidate.archive_info && source.messages == candidate.messages;
    if equal {
        report.charge_work(
            source_payload
                .len()
                .checked_add(candidate_payload.len())
                .ok_or(InvalidationError::InvalidSource)?,
            limits,
        )?;
        return Ok(!source_scan.state.has_preview_fields
            && source_scan.state.dirty == Some(true)
            && !candidate_scan.state.has_preview_fields
            && candidate_scan.state.dirty == Some(true));
    }
    if candidate_scan.state.has_preview_fields || candidate_scan.state.dirty != Some(true) {
        return Ok(false);
    }
    if !payload_delta_matches(source_payload, candidate_payload, limits, report)?
        || !object_delta_matches(
            source,
            candidate,
            source_index,
            &source_scan.references,
            limits,
            report,
        )?
    {
        return Ok(false);
    }
    Ok(true)
}

fn next_raw_field<'source>(
    source: &'source [u8],
    offset: &mut usize,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<Option<RawField<'source>>, InvalidationError> {
    if *offset == source.len() {
        return Ok(None);
    }
    let field = parse_raw_field(source, *offset)?;
    report.charge_field(limits)?;
    *offset = offset
        .checked_add(field.raw.len())
        .ok_or(InvalidationError::InvalidSource)?;
    Ok(Some(field))
}

fn payload_delta_matches(
    source: &[u8],
    candidate: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    report.charge_scan(source.len(), limits)?;
    report.charge_scan(candidate.len(), limits)?;
    let mut source_offset = 0usize;
    let mut candidate_offset = 0usize;
    let mut dirty_seen = false;
    while let Some(before) = next_raw_field(source, &mut source_offset, limits, report)? {
        if is_removed_preview_field(before.number) {
            continue;
        }
        let Some(after) = next_raw_field(candidate, &mut candidate_offset, limits, report)? else {
            return Ok(false);
        };
        if before.number == THUMBNAILS_DIRTY_FIELD {
            dirty_seen = true;
            report.charge_work(after.raw.len(), limits)?;
            if after.raw != [0x70, 0x01] {
                return Ok(false);
            }
        } else {
            report.charge_work(
                before
                    .raw
                    .len()
                    .checked_add(after.raw.len())
                    .ok_or(InvalidationError::InvalidSource)?,
                limits,
            )?;
            if before.raw != after.raw {
                return Ok(false);
            }
        }
    }
    if !dirty_seen {
        let Some(after) = next_raw_field(candidate, &mut candidate_offset, limits, report)? else {
            return Ok(false);
        };
        report.charge_work(after.raw.len(), limits)?;
        if after.raw != [0x70, 0x01] {
            return Ok(false);
        }
    }
    Ok(next_raw_field(candidate, &mut candidate_offset, limits, report)?.is_none())
}

const fn is_removed_preview_field(number: u32) -> bool {
    matches!(
        number,
        DATABASE_THUMBNAIL_FIELD
            | DATABASE_THUMBNAILS_FIELD
            | THUMBNAIL_SIZES_FIELD
            | THUMBNAILS_FIELD
            | THUMBNAIL_DIGESTS_FIELD
    )
}

fn metadata_reference_occurrences(
    references: &mut HashMap<u64, MetadataReferenceOccurrences>,
    identifier: u64,
) -> Result<&mut MetadataReferenceOccurrences, InvalidationError> {
    if identifier == 0 {
        return Err(InvalidationError::InvalidSource);
    }
    if references.len() == references.capacity() {
        references
            .try_reserve(1)
            .map_err(|_allocation| WireError::Allocation {
                resource: "Keynote exact metadata reference index",
                amount: references.len().saturating_add(1),
            })?;
    }
    Ok(references.entry(identifier).or_default())
}

fn index_exact_metadata(
    info: &litchi_iwa_core::MessageInfo,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<HashMap<u64, MetadataReferenceOccurrences>, InvalidationError> {
    let mut references = HashMap::new();
    report.charge_work(1, limits)?;
    for identifier in &info.object_references {
        report.charge_references(1)?;
        report.charge_work(1, limits)?;
        increment_occurrence(
            &mut metadata_reference_occurrences(&mut references, *identifier)?.aggregate_object,
        )?;
    }
    for identifier in &info.data_references {
        report.charge_references(1)?;
        report.charge_work(1, limits)?;
        increment_occurrence(
            &mut metadata_reference_occurrences(&mut references, *identifier)?.aggregate_data,
        )?;
    }
    for field in &info.field_infos {
        report.charge_work(1, limits)?;
        report.charge_work(field.path.as_slice().len(), limits)?;
        let database_thumbnail = is_database_thumbnail_path(field.path.as_slice());
        for identifier in &field.object_references {
            report.charge_references(1)?;
            report.charge_work(1, limits)?;
            let occurrences = metadata_reference_occurrences(&mut references, *identifier)?;
            if database_thumbnail {
                increment_occurrence(&mut occurrences.database_thumbnail_object)?;
            } else {
                increment_occurrence(&mut occurrences.other_field_object)?;
            }
        }
        let preview = is_preview_path(field.path.as_slice());
        for identifier in &field.data_references {
            report.charge_references(1)?;
            report.charge_work(1, limits)?;
            let occurrences = metadata_reference_occurrences(&mut references, *identifier)?;
            if preview {
                increment_occurrence(&mut occurrences.preview_field_data)?;
            } else {
                increment_occurrence(&mut occurrences.nonpreview_field_data)?;
            }
        }
    }
    Ok(references)
}

fn validate_exact_ownership(
    object: &ArchiveObject,
    message_index: usize,
    payload_references: &HashMap<u64, PayloadReferenceOccurrences>,
    state: ExactPayloadState,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<(), InvalidationError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    let metadata_references = index_exact_metadata(info, limits, report)?;
    report.charge_work(payload_references.len(), limits)?;
    let mut preview_objects = 0usize;
    let mut children = 0usize;
    let mut preview_data = 0usize;
    for (identifier, payload) in payload_references {
        let metadata = metadata_references
            .get(identifier)
            .copied()
            .unwrap_or_default();
        if payload.preview_object > 1
            || payload.child > 1
            || (payload.preview_object != 0 && payload.retained_object != 0)
            || (payload.preview_object != 0 && metadata.aggregate_object != 1)
            || (payload.child != 0 && metadata.aggregate_object != 1)
            || payload.preview_data > 1
            || (payload.preview_data != 0 && metadata.aggregate_data != 1)
        {
            return Err(InvalidationError::InvalidSource);
        }
        preview_objects = preview_objects
            .checked_add(payload.preview_object)
            .ok_or(InvalidationError::InvalidSource)?;
        children = children
            .checked_add(payload.child)
            .ok_or(InvalidationError::InvalidSource)?;
        preview_data = preview_data
            .checked_add(payload.preview_data)
            .ok_or(InvalidationError::InvalidSource)?;
    }
    if preview_objects != state.preview_object_references
        || children != state.child_references
        || preview_data != state.preview_data_references
    {
        return Err(InvalidationError::InvalidSource);
    }
    report.charge_work(metadata_references.len(), limits)?;
    for (identifier, metadata) in &metadata_references {
        let payload = payload_references
            .get(identifier)
            .copied()
            .unwrap_or_default();
        if metadata.aggregate_object > 1
            || metadata.aggregate_data > 1
            || (metadata.database_thumbnail_object != 0 && payload.preview_object != 1)
            || (metadata.other_field_object != 0 && payload.preview_object != 0)
            || (metadata.aggregate_data != 0
                && payload.preview_data == 0
                && metadata.nonpreview_field_data == 0)
            || (metadata.preview_field_data != 0 && payload.preview_data != 1)
            || (metadata.nonpreview_field_data != 0 && payload.preview_data != 0)
        {
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(())
}

fn object_delta_matches(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    selected_index: usize,
    source_references: &HashMap<u64, PayloadReferenceOccurrences>,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    for (index, (before, after)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        report.charge_work(2, limits)?;
        if before.type_ != after.type_ {
            return Ok(false);
        }
        if index != selected_index {
            report.charge_work(
                before
                    .data
                    .len()
                    .checked_add(after.data.len())
                    .ok_or(InvalidationError::InvalidSource)?,
                limits,
            )?;
            if before != after {
                return Ok(false);
            }
        }
    }
    if source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len() {
        return Ok(false);
    }
    for (index, (before, after)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        charge_exact_message_info_structure(before, limits, report)?;
        charge_exact_message_info_structure(after, limits, report)?;
        if index == selected_index {
            if !message_info_delta_matches(
                before,
                after,
                source_references,
                candidate.messages[index].data.len(),
                limits,
                report,
            )? {
                return Ok(false);
            }
        } else if before != after {
            return Ok(false);
        }
    }
    Ok(true)
}

fn message_info_delta_matches(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
    source_references: &HashMap<u64, PayloadReferenceOccurrences>,
    candidate_payload_length: usize,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    let Ok(expected_length) = u32::try_from(candidate_payload_length) else {
        return Ok(false);
    };
    if source.type_ != candidate.type_
        || source.versions != candidate.versions
        || candidate.length != expected_length
        || source.base_message_index != candidate.base_message_index
        || source.diff_merge_version != candidate.diff_merge_version
        || source.diff_field_path != candidate.diff_field_path
        || source.fields_to_remove != candidate.fields_to_remove
        || source.diff_read_version != candidate.diff_read_version
        || source.field_infos.len() != candidate.field_infos.len()
        || !object_reference_delta_matches(
            &source.object_references,
            &candidate.object_references,
            source_references,
            limits,
            report,
        )?
        || !data_reference_delta_matches(
            &source.data_references,
            &candidate.data_references,
            source_references,
            limits,
            report,
        )?
    {
        return Ok(false);
    }
    for (before, after) in source.field_infos.iter().zip(&candidate.field_infos) {
        if before.path != after.path
            || before.r#type != after.r#type
            || before.unknown_field_rule != after.unknown_field_rule
            || before.known_field_rule != after.known_field_rule
            || before.known_field_version != after.known_field_version
            || before.known_field_feature_identifier != after.known_field_feature_identifier
            || !object_reference_delta_matches(
                &before.object_references,
                &after.object_references,
                source_references,
                limits,
                report,
            )?
            || !data_reference_delta_matches(
                &before.data_references,
                &after.data_references,
                source_references,
                limits,
                report,
            )?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn object_reference_delta_matches(
    source: &[u64],
    candidate: &[u64],
    source_references: &HashMap<u64, PayloadReferenceOccurrences>,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    let compared = source
        .len()
        .checked_add(candidate.len())
        .ok_or(InvalidationError::InvalidSource)?;
    report.charge_references(compared)?;
    report.charge_work(compared, limits)?;
    let mut candidate_index = 0usize;
    for identifier in source {
        let removed = source_references
            .get(identifier)
            .is_some_and(|occurrences| occurrences.preview_object != 0);
        if !removed {
            if candidate.get(candidate_index) != Some(identifier) {
                return Ok(false);
            }
            candidate_index = candidate_index
                .checked_add(1)
                .ok_or(InvalidationError::InvalidSource)?;
        }
    }
    Ok(candidate_index == candidate.len())
}

fn data_reference_delta_matches(
    source: &[u64],
    candidate: &[u64],
    source_references: &HashMap<u64, PayloadReferenceOccurrences>,
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    let compared = source
        .len()
        .checked_add(candidate.len())
        .ok_or(InvalidationError::InvalidSource)?;
    report.charge_references(compared)?;
    report.charge_work(compared, limits)?;
    let mut candidate_index = 0usize;
    for identifier in source {
        let removed = source_references
            .get(identifier)
            .is_some_and(|occurrences| occurrences.preview_data != 0);
        if !removed {
            if candidate.get(candidate_index) != Some(identifier) {
                return Ok(false);
            }
            candidate_index = candidate_index
                .checked_add(1)
                .ok_or(InvalidationError::InvalidSource)?;
        }
    }
    Ok(candidate_index == candidate.len())
}

fn selected_message_index(object: &ArchiveObject) -> Result<usize, InvalidationError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(InvalidationError::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_ {
            return Err(InvalidationError::InvalidSource);
        }
        if message.type_ == SLIDE_NODE_MESSAGE_TYPE && selected.replace(index).is_some() {
            return Err(InvalidationError::InvalidSource);
        }
    }
    selected.ok_or(InvalidationError::InvalidSource)
}

fn validate_metadata_topology(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), InvalidationError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(())
}

fn rewrite_scanned_payload(
    scan: PayloadScan<'_>,
    budget: &mut ScanBudget,
) -> Result<PayloadRewrite, InvalidationError> {
    let PayloadScan {
        view,
        retained_bytes,
        removed_object_references,
        removed_data_references,
        retained_child_object_references: _,
        dirty,
        has_preview_fields: _,
    } = scan;
    if retained_bytes > budget.limits.max_output_bytes() {
        return Err(WireError::LimitExceeded {
            kind: WireLimitKind::OutputBytes,
            observed: retained_bytes,
            limit: budget.limits.max_output_bytes(),
        }
        .into());
    }
    budget.add_work(retained_bytes)?;
    let mut data = Vec::new();
    data.try_reserve_exact(retained_bytes)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote slide preview payload",
            amount: retained_bytes,
        })?;
    for field in view.fields() {
        match field.number() {
            DATABASE_THUMBNAIL_FIELD
            | DATABASE_THUMBNAILS_FIELD
            | THUMBNAIL_SIZES_FIELD
            | THUMBNAILS_FIELD
            | THUMBNAIL_DIGESTS_FIELD => {},
            THUMBNAILS_DIRTY_FIELD => {
                data.extend_from_slice(field.key());
                data.push(1);
            },
            _ => data.extend_from_slice(field.raw()),
        }
    }
    if dirty.is_none() {
        // Field 14, wire type 0, followed by boolean true.
        data.extend_from_slice(&[0x70, 0x01]);
    }
    if data.len() != retained_bytes {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(PayloadRewrite {
        data,
        removed_object_references,
        removed_data_references,
    })
}

fn scan_payload<'source>(
    source: &'source [u8],
    budget: &mut ScanBudget,
) -> Result<PayloadScan<'source>, InvalidationError> {
    let view = budget.parse(source)?;
    let mut database_thumbnail_seen = false;
    let mut slide_reference_seen = false;
    let mut dirty_seen = false;
    let mut dirty = None;
    let mut has_preview_fields = false;
    let mut retained_bytes = 0usize;
    let mut removed_object_references = Vec::new();
    let mut removed_data_references = Vec::new();
    let mut surviving_object_references = Vec::new();
    let mut retained_child_object_references = Vec::new();

    let removed_object_count = view
        .fields()
        .filter(|field| {
            matches!(
                field.number(),
                DATABASE_THUMBNAIL_FIELD | DATABASE_THUMBNAILS_FIELD
            )
        })
        .count();
    let removed_data_count = view
        .fields()
        .filter(|field| field.number() == THUMBNAILS_FIELD)
        .count();
    let surviving_object_count = view
        .fields()
        .filter(|field| matches!(field.number(), 1 | 2))
        .count();
    let child_object_count = view.fields().filter(|field| field.number() == 1).count();
    try_reserve_references(
        &mut removed_object_references,
        removed_object_count,
        "Keynote removed slide preview object references",
    )?;
    try_reserve_references(
        &mut removed_data_references,
        removed_data_count,
        "Keynote removed slide preview data references",
    )?;
    try_reserve_references(
        &mut surviving_object_references,
        surviving_object_count,
        "Keynote surviving slide-node object references",
    )?;
    try_reserve_references(
        &mut retained_child_object_references,
        child_object_count,
        "Keynote retained slide-node child references",
    )?;

    for field in view.fields() {
        match field.number() {
            1 => {
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
                let identifier = reference_identifier(field.payload(), budget)?;
                surviving_object_references.push(identifier);
                retained_child_object_references.push(identifier);
                retained_bytes = retained_bytes
                    .checked_add(field.raw().len())
                    .ok_or(InvalidationError::InvalidSource)?;
            },
            2 => {
                if std::mem::replace(&mut slide_reference_seen, true) {
                    return Err(InvalidationError::InvalidSource);
                }
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
                surviving_object_references.push(reference_identifier(field.payload(), budget)?);
                retained_bytes = retained_bytes
                    .checked_add(field.raw().len())
                    .ok_or(InvalidationError::InvalidSource)?;
            },
            DATABASE_THUMBNAIL_FIELD => {
                if std::mem::replace(&mut database_thumbnail_seen, true) {
                    return Err(InvalidationError::InvalidSource);
                }
                has_preview_fields = true;
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
                removed_object_references.push(reference_identifier(field.payload(), budget)?);
            },
            DATABASE_THUMBNAILS_FIELD => {
                has_preview_fields = true;
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
                removed_object_references.push(reference_identifier(field.payload(), budget)?);
            },
            THUMBNAILS_FIELD => {
                has_preview_fields = true;
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
                removed_data_references.push(data_reference_identifier(field.payload(), budget)?);
            },
            THUMBNAIL_SIZES_FIELD | THUMBNAIL_DIGESTS_FIELD => {
                has_preview_fields = true;
                validate_recognized_field(field)?;
                require_length_delimited(field.wire_type())?;
            },
            THUMBNAILS_DIRTY_FIELD => {
                if std::mem::replace(&mut dirty_seen, true) || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                validate_recognized_field(field)?;
                let value = strict_varint(field.payload())?;
                if value > 1 {
                    return Err(InvalidationError::InvalidSource);
                }
                dirty = Some(value == 1);
                retained_bytes = retained_bytes
                    .checked_add(field.key().len())
                    .and_then(|length| length.checked_add(1))
                    .ok_or(InvalidationError::InvalidSource)?;
            },
            _ => {
                retained_bytes = retained_bytes
                    .checked_add(field.raw().len())
                    .ok_or(InvalidationError::InvalidSource)?;
            },
        }
    }
    if !dirty_seen {
        retained_bytes = retained_bytes
            .checked_add(2)
            .ok_or(InvalidationError::InvalidSource)?;
    }
    let mut removed = HashSet::new();
    removed
        .try_reserve(removed_object_references.len())
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote slide preview reference set",
            amount: removed_object_references.len(),
        })?;
    for identifier in &removed_object_references {
        if !removed.insert(*identifier) {
            return Err(InvalidationError::InvalidSource);
        }
    }
    if surviving_object_references
        .iter()
        .any(|identifier| removed.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
    let mut data_identifiers = HashSet::new();
    data_identifiers
        .try_reserve(removed_data_references.len())
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote slide preview data-reference set",
            amount: removed_data_references.len(),
        })?;
    for identifier in &removed_data_references {
        if !data_identifiers.insert(*identifier) {
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(PayloadScan {
        view,
        retained_bytes,
        removed_object_references,
        removed_data_references,
        retained_child_object_references,
        dirty,
        has_preview_fields,
    })
}

fn reference_identifier(source: &[u8], budget: &mut ScanBudget) -> Result<u64, InvalidationError> {
    let view = budget.parse(source)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut deprecated_external_seen = false;
    for field in view.fields() {
        match field.number() {
            1 => {
                validate_recognized_field(field)?;
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                let value = strict_varint(field.payload())?;
                if value == 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 => {
                validate_recognized_field(field)?;
                if std::mem::replace(&mut deprecated_type_seen, true) || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                let _deprecated_type = decode_canonical_int32(strict_varint(field.payload())?)?;
            },
            3 => {
                validate_recognized_field(field)?;
                validate_optional_bool_reference_field(field, &mut deprecated_external_seen)?;
            },
            _ => {},
        }
    }
    let resolved = identifier.ok_or(InvalidationError::InvalidSource)?;
    budget.add_references(1)?;
    Ok(resolved)
}

fn data_reference_identifier(
    source: &[u8],
    budget: &mut ScanBudget,
) -> Result<u64, InvalidationError> {
    let view = budget.parse(source)?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() == 1 {
            validate_recognized_field(field)?;
            if identifier.is_some() || field.wire_type() != 0 {
                return Err(InvalidationError::InvalidSource);
            }
            let value = strict_varint(field.payload())?;
            if value == 0 {
                return Err(InvalidationError::InvalidSource);
            }
            identifier = Some(value);
        }
    }
    let resolved = identifier.ok_or(InvalidationError::InvalidSource)?;
    budget.add_references(1)?;
    Ok(resolved)
}

fn validate_reference_ownership(
    object: &ArchiveObject,
    message_index: usize,
    object_identifiers: &[u64],
    retained_child_identifiers: &[u64],
    payload_data_identifiers: &[u64],
    budget: &mut ScanBudget,
) -> Result<(), InvalidationError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    budget.add_work(1)?;
    budget.add_references(info.object_references.len())?;
    budget.add_work(info.object_references.len())?;
    budget.add_references(info.data_references.len())?;
    budget.add_work(info.data_references.len())?;
    let mut unrelated_capacity = 0usize;
    for field in &info.field_infos {
        budget.add_work(1)?;
        budget.add_work(field.path.as_slice().len())?;
        budget.add_references(field.object_references.len())?;
        budget.add_work(field.object_references.len())?;
        budget.add_references(field.data_references.len())?;
        budget.add_work(field.data_references.len())?;
        if !is_preview_path(field.path.as_slice()) {
            unrelated_capacity = unrelated_capacity
                .checked_add(field.data_references.len())
                .ok_or(InvalidationError::InvalidSource)?;
        }
    }
    let payload_reference_walk = object_identifiers
        .len()
        .checked_add(retained_child_identifiers.len())
        .and_then(|count| count.checked_add(payload_data_identifiers.len()))
        .ok_or(InvalidationError::InvalidSource)?;
    budget.add_references(payload_reference_walk)?;
    budget.add_work(payload_reference_walk)?;
    let object_removals =
        unique_reference_set(object_identifiers, "Keynote slide preview ownership set")?;
    let aggregate_objects = unique_reference_set(
        &info.object_references,
        "Keynote aggregate object-reference set",
    )?;
    let retained_children = unique_reference_set(
        retained_child_identifiers,
        "Keynote retained child object-reference set",
    )?;
    if object_removals
        .iter()
        .chain(&retained_children)
        .any(|identifier| !aggregate_objects.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
    let preview = unique_reference_set(
        payload_data_identifiers,
        "Keynote preview-owned data-reference set",
    )?;
    let aggregate_data = unique_reference_set(
        &info.data_references,
        "Keynote aggregate data-reference set",
    )?;
    if preview
        .iter()
        .any(|identifier| !aggregate_data.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
    let mut unrelated = HashSet::new();
    unrelated
        .try_reserve(unrelated_capacity)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote unrelated data-reference set",
            amount: unrelated_capacity,
        })?;
    for field in &info.field_infos {
        budget.add_work(1)?;
        budget.add_work(field.path.as_slice().len())?;
        budget.add_references(field.object_references.len())?;
        budget.add_work(field.object_references.len())?;
        budget.add_references(field.data_references.len())?;
        budget.add_work(field.data_references.len())?;
        let database_thumbnail_path = is_database_thumbnail_path(field.path.as_slice());
        if field
            .object_references
            .iter()
            .any(|identifier| object_removals.contains(identifier) != database_thumbnail_path)
        {
            return Err(InvalidationError::InvalidSource);
        }
        if is_preview_path(field.path.as_slice()) {
            if field
                .data_references
                .iter()
                .any(|identifier| !preview.contains(identifier))
            {
                return Err(InvalidationError::InvalidSource);
            }
        } else {
            for identifier in &field.data_references {
                if preview.contains(identifier) {
                    return Err(InvalidationError::InvalidSource);
                }
                unrelated.insert(*identifier);
            }
        }
    }
    budget.add_references(aggregate_data.len())?;
    budget.add_work(aggregate_data.len())?;
    for identifier in &aggregate_data {
        if !preview.contains(identifier) && !unrelated.contains(identifier) {
            // Aggregate metadata cannot prove whether this identifier belongs
            // to a removed preview field or to surviving opaque content.
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(())
}

fn unique_reference_set(
    identifiers: &[u64],
    resource: &'static str,
) -> Result<HashSet<u64>, InvalidationError> {
    let mut unique = HashSet::new();
    unique
        .try_reserve(identifiers.len())
        .map_err(|_allocation| WireError::Allocation {
            resource,
            amount: identifiers.len(),
        })?;
    for identifier in identifiers {
        if *identifier == 0 || !unique.insert(*identifier) {
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(unique)
}

fn try_reserve_references(
    identifiers: &mut Vec<u64>,
    amount: usize,
    resource: &'static str,
) -> Result<(), InvalidationError> {
    identifiers
        .try_reserve_exact(amount)
        .map_err(|_allocation| WireError::Allocation { resource, amount }.into())
}

fn is_database_thumbnail_path(path: &[u32]) -> bool {
    matches!(
        path.first(),
        Some(&DATABASE_THUMBNAIL_FIELD | &DATABASE_THUMBNAILS_FIELD)
    )
}

fn is_preview_path(path: &[u32]) -> bool {
    matches!(
        path.first(),
        Some(
            &DATABASE_THUMBNAIL_FIELD
                | &DATABASE_THUMBNAILS_FIELD
                | &THUMBNAIL_SIZES_FIELD
                | &THUMBNAILS_FIELD
                | &THUMBNAIL_DIGESTS_FIELD
        )
    )
}

fn validate_rewrite(
    source: &[u8],
    rewritten: &[u8],
    budget: &mut ScanBudget,
) -> Result<(), InvalidationError> {
    let source_view = budget.parse(source)?;
    let rewritten_view = budget.parse(rewritten)?;
    let source_untouched = source_view
        .fields()
        .filter(|field| !is_preview_field(field.number()))
        .map(litchi_iwa_common::wire::WireFieldView::raw);
    let rewritten_untouched = rewritten_view
        .fields()
        .filter(|field| !is_preview_field(field.number()))
        .map(litchi_iwa_common::wire::WireFieldView::raw);
    if !source_untouched.eq(rewritten_untouched) {
        return Err(InvalidationError::InvalidSource);
    }
    let mut dirty = None;
    for field in rewritten_view.fields() {
        match field.number() {
            DATABASE_THUMBNAIL_FIELD
            | DATABASE_THUMBNAILS_FIELD
            | THUMBNAIL_SIZES_FIELD
            | THUMBNAILS_FIELD
            | THUMBNAIL_DIGESTS_FIELD => return Err(InvalidationError::InvalidSource),
            THUMBNAILS_DIRTY_FIELD => {
                if dirty.is_some() || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                dirty = Some(strict_varint(field.payload())?);
            },
            _ => {},
        }
    }
    if dirty != Some(1) {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(())
}

const fn is_preview_field(number: u32) -> bool {
    matches!(
        number,
        DATABASE_THUMBNAIL_FIELD
            | DATABASE_THUMBNAILS_FIELD
            | THUMBNAIL_SIZES_FIELD
            | THUMBNAILS_DIRTY_FIELD
            | THUMBNAILS_FIELD
            | THUMBNAIL_DIGESTS_FIELD
    )
}

fn require_length_delimited(wire_type: u8) -> Result<(), InvalidationError> {
    if wire_type == 2 {
        Ok(())
    } else {
        Err(InvalidationError::InvalidSource)
    }
}

fn validate_recognized_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<(), InvalidationError> {
    field.validate_canonical_framing()?;
    Ok(())
}

fn validate_optional_bool_reference_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    seen: &mut bool,
) -> Result<(), InvalidationError> {
    if std::mem::replace(seen, true)
        || field.wire_type() != 0
        || strict_varint(field.payload())? != 0
    {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(())
}

fn decode_canonical_int32(value: u64) -> Result<i32, InvalidationError> {
    if value <= 2_147_483_647 {
        return i32::try_from(value).map_err(|_conversion| InvalidationError::InvalidSource);
    }
    if value < 0xffff_ffff_8000_0000 {
        return Err(InvalidationError::InvalidSource);
    }
    let bits = u32::try_from(value & u64::from(u32::MAX))
        .map_err(|_conversion| InvalidationError::InvalidSource)?;
    Ok(i32::from_ne_bytes(bits.to_ne_bytes()))
}

fn strict_varint(payload: &[u8]) -> Result<u64, InvalidationError> {
    let (value, encoded) =
        decode_varint_from_bytes(payload).map_err(|_error| InvalidationError::InvalidSource)?;
    if encoded == payload.len() && encoded == encoded_len(value) {
        Ok(value)
    } else {
        Err(InvalidationError::InvalidSource)
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{FieldInfo, FieldPath, MessageInfo};

    use super::*;

    const UNKNOWN_FIELD_BYTES: &[u8] = &[0xc2, 0x02, 0x03, 0xaa, 0xbb, 0xcc];

    #[test]
    fn invalidates_preview_bytes_and_reference_metadata() -> Result<(), InvalidationError> {
        let mut object = preview_object(false)?;
        invalidate(&mut object, ArchiveLimits::default(), WireLimits::default())?;

        assert!(is_invalidated(&object, WireLimits::default())?);
        let payload = &object.messages[0].data;
        assert!(
            payload
                .windows(UNKNOWN_FIELD_BYTES.len())
                .any(|bytes| bytes == UNKNOWN_FIELD_BYTES)
        );
        assert!(payload.windows(2).any(|bytes| bytes == [0x70, 0x01]));
        let info = &object.archive_info.message_infos[0];
        assert_eq!(info.object_references, [99]);
        assert_eq!(info.data_references, [56]);
        assert!(info.field_infos[0].object_references.is_empty());
        assert!(info.field_infos[1].object_references.is_empty());
        assert!(info.field_infos[0].data_references.is_empty());
        assert!(info.field_infos[2].data_references.is_empty());
        assert_eq!(info.field_infos[3].object_references, [99]);
        assert_eq!(info.field_infos[3].data_references, [56]);
        Ok(())
    }

    #[test]
    fn rejects_merge_topology_without_mutating() -> Result<(), InvalidationError> {
        let mut object = preview_object(false)?;
        object.archive_info.should_merge = Some(true);
        let original = object.clone();

        assert!(matches!(
            invalidate(&mut object, ArchiveLimits::default(), WireLimits::default()),
            Err(InvalidationError::InvalidSource)
        ));
        assert_eq!(object, original);
        Ok(())
    }

    #[test]
    fn rejects_reference_shared_with_surviving_slide_edge() -> Result<(), InvalidationError> {
        let mut object = preview_object(true)?;
        let original = object.clone();

        assert!(matches!(
            invalidate(&mut object, ArchiveLimits::default(), WireLimits::default()),
            Err(InvalidationError::InvalidSource)
        ));
        assert_eq!(object, original);
        Ok(())
    }

    #[test]
    fn rejects_aggregate_only_data_reference_without_mutating() -> Result<(), InvalidationError> {
        let mut object = preview_object(false)?;
        object.archive_info.message_infos[0]
            .data_references
            .push(57);
        let original = object.clone();

        assert!(matches!(
            invalidate(&mut object, ArchiveLimits::default(), WireLimits::default()),
            Err(InvalidationError::InvalidSource)
        ));
        assert_eq!(object, original);
        Ok(())
    }

    #[test]
    fn enforces_wire_field_limit_without_mutating() -> Result<(), InvalidationError> {
        let mut object = preview_object(false)?;
        let original = object.clone();
        let limits = WireLimits::default().with_fields(1)?;

        assert!(matches!(
            invalidate(&mut object, ArchiveLimits::default(), limits),
            Err(InvalidationError::Wire(WireError::LimitExceeded {
                kind: WireLimitKind::Fields,
                ..
            }))
        ));
        assert_eq!(object, original);
        Ok(())
    }

    #[test]
    fn rejects_noncanonical_recognized_framing_and_values_without_mutating()
    -> Result<(), InvalidationError> {
        let cases: &[(&[u8], &[u8])] = &[
            (&[0x70, 0x00], &[0xf0, 0x00, 0x00]),
            (&[0x70, 0x00], &[0x70, 0x80, 0x00]),
            (
                &[0x82, 0x01, 0x02, 0x08, 55],
                &[0x82, 0x01, 0x82, 0x00, 0x08, 55],
            ),
            (
                &[0x82, 0x01, 0x02, 0x08, 55],
                &[0x82, 0x01, 0x03, 0x08, 0xb7, 0x00],
            ),
            (&[0x1a, 0x02, 0x08, 77], &[0x1a, 0x03, 0x08, 0xcd, 0x00]),
            (&[0x1a, 0x02, 0x08, 77], &[0x1a, 0x03, 0x88, 0x00, 77]),
        ];
        for &(source, replacement) in cases {
            let mut object = preview_object(false)?;
            replace_once(&mut object.messages[0].data, source, replacement)?;
            assert_rejected_without_mutation(object);
        }
        Ok(())
    }

    #[test]
    fn requires_payload_cache_ids_exactly_once_in_aggregate_metadata()
    -> Result<(), InvalidationError> {
        let mut missing_object = preview_object(false)?;
        missing_object.archive_info.message_infos[0]
            .object_references
            .retain(|identifier| *identifier != 77);
        assert_rejected_without_mutation(missing_object);

        let mut duplicate_object = preview_object(false)?;
        duplicate_object.archive_info.message_infos[0]
            .object_references
            .push(77);
        assert_rejected_without_mutation(duplicate_object);

        let mut missing_data = preview_object(false)?;
        missing_data.archive_info.message_infos[0]
            .data_references
            .retain(|identifier| *identifier != 55);
        assert_rejected_without_mutation(missing_data);

        let mut duplicate_data = preview_object(false)?;
        duplicate_data.archive_info.message_infos[0]
            .data_references
            .push(55);
        assert_rejected_without_mutation(duplicate_data);
        Ok(())
    }

    #[test]
    fn rejects_unsafe_deprecated_reference_semantics_without_mutating()
    -> Result<(), InvalidationError> {
        let mut out_of_int32 = preview_object(false)?;
        replace_once(
            &mut out_of_int32.messages[0].data,
            &[0x4a, 0x02, 0x08, 78],
            &[0x4a, 0x08, 0x08, 78, 0x10, 0x80, 0x80, 0x80, 0x80, 0x08],
        )?;
        assert_rejected_without_mutation(out_of_int32);

        let mut external = preview_object(false)?;
        replace_once(
            &mut external.messages[0].data,
            &[0x4a, 0x02, 0x08, 78],
            &[0x4a, 0x04, 0x08, 78, 0x18, 0x01],
        )?;
        assert_rejected_without_mutation(external);

        let mut valid_boundary = preview_object(false)?;
        replace_once(
            &mut valid_boundary.messages[0].data,
            &[0x4a, 0x02, 0x08, 78],
            &[
                0x4a, 0x0a, 0x08, 78, 0x10, 0xff, 0xff, 0xff, 0xff, 0x07, 0x18, 0x00,
            ],
        )?;
        invalidate(
            &mut valid_boundary,
            ArchiveLimits::default(),
            WireLimits::default(),
        )?;
        assert!(is_invalidated(&valid_boundary, WireLimits::default())?);
        Ok(())
    }

    #[test]
    fn requires_each_retained_child_reference_in_aggregate_metadata()
    -> Result<(), InvalidationError> {
        let child = [0x0a, 0x02, 0x08, 88];
        let mut missing = preview_object(false)?;
        missing.messages[0].data.extend_from_slice(&child);
        assert_rejected_without_mutation(missing);

        let mut duplicate = preview_object(false)?;
        duplicate.messages[0].data.extend_from_slice(&child);
        duplicate.messages[0].data.extend_from_slice(&child);
        duplicate.archive_info.message_infos[0]
            .object_references
            .push(88);
        assert_rejected_without_mutation(duplicate);

        let mut valid = preview_object(false)?;
        valid.messages[0].data.extend_from_slice(&child);
        valid.archive_info.message_infos[0]
            .object_references
            .push(88);
        invalidate(&mut valid, ArchiveLimits::default(), WireLimits::default())?;
        assert!(is_invalidated(&valid, WireLimits::default())?);
        assert_eq!(
            valid.archive_info.message_infos[0].object_references,
            [99, 88]
        );
        Ok(())
    }

    #[test]
    fn exact_delta_proves_forward_inverse_and_invalidated_equality() -> Result<(), InvalidationError>
    {
        let source = preview_object(false)?;
        let mut invalidated = source.clone();
        invalidate(
            &mut invalidated,
            ArchiveLimits::default(),
            WireLimits::default(),
        )?;

        let (forward, forward_report) = exact_invalidation_delta(
            &source,
            &invalidated,
            InvalidationDirection::Forward,
            WireLimits::default(),
        )?;
        assert!(forward);
        assert!(forward_report.references() > 0);
        assert!(forward_report.work() > 0);

        let (inverse, inverse_report) = exact_invalidation_delta(
            &invalidated,
            &source,
            InvalidationDirection::Inverse,
            WireLimits::default(),
        )?;
        assert!(inverse);
        assert_eq!(inverse_report, forward_report);

        assert!(
            exact_invalidation_delta(
                &invalidated,
                &invalidated,
                InvalidationDirection::Forward,
                WireLimits::default(),
            )?
            .0
        );
        assert!(
            !exact_invalidation_delta(
                &source,
                &source,
                InvalidationDirection::Forward,
                WireLimits::default(),
            )?
            .0
        );
        Ok(())
    }

    #[test]
    fn exact_delta_rejects_partial_and_unrelated_candidate_changes() -> Result<(), InvalidationError>
    {
        let source = preview_object(false)?;
        let mut invalidated = source.clone();
        invalidate(
            &mut invalidated,
            ArchiveLimits::default(),
            WireLimits::default(),
        )?;

        let mut changed_unknown = invalidated.clone();
        let unknown = changed_unknown.messages[0]
            .data
            .iter_mut()
            .find(|byte| **byte == 0xaa)
            .ok_or(InvalidationError::InvalidSource)?;
        *unknown = 0xab;
        assert!(
            !exact_invalidation_delta(
                &source,
                &changed_unknown,
                InvalidationDirection::Forward,
                WireLimits::default(),
            )?
            .0
        );

        let mut changed_metadata = invalidated.clone();
        changed_metadata.archive_info.message_infos[0]
            .versions
            .push(99);
        assert!(
            !exact_invalidation_delta(
                &source,
                &changed_metadata,
                InvalidationDirection::Forward,
                WireLimits::default(),
            )?
            .0
        );

        let mut partial = source.clone();
        replace_once(&mut partial.messages[0].data, &[0x70, 0x00], &[0x70, 0x01])?;
        assert!(
            !exact_invalidation_delta(
                &source,
                &partial,
                InvalidationDirection::Forward,
                WireLimits::default(),
            )?
            .0
        );
        Ok(())
    }

    #[test]
    fn exact_delta_reference_index_scales_linearly_to_8192_distinct_ids()
    -> Result<(), InvalidationError> {
        let work_4096 = scaled_exact_delta_work(4_096)?;
        let work_8192 = scaled_exact_delta_work(8_192)?;
        assert!(
            work_8192.saturating_mul(10) <= work_4096.saturating_mul(23),
            "doubling distinct references grew work from {work_4096} to {work_8192}"
        );
        Ok(())
    }

    #[test]
    fn remaining_allowance_stops_exact_and_mutation_before_object_changes()
    -> Result<(), InvalidationError> {
        let source = preview_object(false)?;
        let mut candidate = source.clone();
        invalidate(
            &mut candidate,
            ArchiveLimits::default(),
            WireLimits::default(),
        )?;
        assert!(matches!(
            exact_invalidation_delta_with_allowance(
                &source,
                &candidate,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::new(0, usize::MAX),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                observed: 1..,
                maximum: 0,
            })
        ));

        let mut mutation = source.clone();
        let original = mutation.clone();
        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut mutation,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::new(0, usize::MAX),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                observed: 1..,
                maximum: 0,
            })
        ));
        assert_eq!(mutation, original);

        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut mutation,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::new(usize::MAX, 0),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::References,
                observed: 1..,
                maximum: 0,
            })
        ));
        assert_eq!(mutation, original);

        let (changed, report) = invalidate_if_needed_with_report(
            &mut mutation,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(|error| match error {
            BudgetedInvalidationError::Invalidation(inner) => inner,
            BudgetedInvalidationError::BudgetExceeded { .. } => InvalidationError::InvalidSource,
        })?;
        assert!(changed);
        assert!(report.work() > 0);
        assert!(report.references() > 0);
        let invalidated = mutation.clone();
        let (changed, noop_report) = invalidate_if_needed_with_report(
            &mut mutation,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(|error| match error {
            BudgetedInvalidationError::Invalidation(inner) => inner,
            BudgetedInvalidationError::BudgetExceeded { .. } => InvalidationError::InvalidSource,
        })?;
        assert!(!changed);
        assert!(noop_report.work() > 0);
        assert_eq!(mutation, invalidated);
        Ok(())
    }

    #[test]
    fn empty_metadata_structure_is_charged_before_mutation_and_exact_comparison()
    -> Result<(), InvalidationError> {
        const EMPTY_FIELDS: usize = 4_096;
        let source = empty_metadata_object(Vec::new(), EMPTY_FIELDS)?;
        let mut proven = source.clone();
        let (changed, report) = invalidate_if_needed_with_report(
            &mut proven,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;
        assert!(changed);
        assert!(report.work() >= EMPTY_FIELDS.saturating_mul(2));

        let mut zero_work = source.clone();
        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut zero_work,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::new(0, usize::MAX),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                observed: 1..,
                maximum: 0,
            })
        ));
        assert_eq!(zero_work, source);

        let invalidated = empty_metadata_object(vec![0x70, 0x01], EMPTY_FIELDS)?;
        let mut tiny_allowance = invalidated.clone();
        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut tiny_allowance,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::new(invalidated.messages[0].data.len(), usize::MAX),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                ..
            })
        ));
        assert_eq!(tiny_allowance, invalidated);

        let (matches, exact_report) = exact_invalidation_delta_with_allowance(
            &invalidated,
            &invalidated,
            InvalidationDirection::Forward,
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;
        assert!(matches);
        assert!(exact_report.work() >= EMPTY_FIELDS.saturating_mul(2));
        assert!(matches!(
            exact_invalidation_delta_with_allowance(
                &invalidated,
                &invalidated,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::new(
                    invalidated.messages[0].data.len().saturating_mul(2),
                    usize::MAX,
                ),
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn nonselected_payload_and_metadata_are_charged_before_broad_validation_or_equality()
    -> Result<(), InvalidationError> {
        let object = large_nonselected_object()?;
        let mut proven = object.clone();
        invalidate(&mut proven, ArchiveLimits::default(), WireLimits::default())?;
        assert_eq!(proven, object);

        let low_work = InvalidationAllowance::new(1_024, usize::MAX);
        assert!(matches!(
            exact_invalidation_delta_with_allowance(
                &object,
                &object,
                InvalidationDirection::Forward,
                WireLimits::default(),
                low_work,
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                observed: 1_025..,
                maximum: 1_024,
            })
        ));
        let mut mutation = object.clone();
        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut mutation,
                ArchiveLimits::default(),
                WireLimits::default(),
                low_work,
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::Work,
                observed: 1_025..,
                maximum: 1_024,
            })
        ));
        assert_eq!(mutation, object);

        let low_references = InvalidationAllowance::new(usize::MAX, 1_024);
        assert!(matches!(
            exact_invalidation_delta_with_allowance(
                &object,
                &object,
                InvalidationDirection::Forward,
                WireLimits::default(),
                low_references,
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::References,
                observed: 1_025..,
                maximum: 1_024,
            })
        ));
        let mut mutation = object.clone();
        assert!(matches!(
            invalidate_if_needed_with_report(
                &mut mutation,
                ArchiveLimits::default(),
                WireLimits::default(),
                low_references,
            ),
            Err(BudgetedInvalidationError::BudgetExceeded {
                kind: InvalidationBudgetKind::References,
                observed: 1_025..,
                maximum: 1_024,
            })
        ));
        assert_eq!(mutation, object);
        Ok(())
    }

    #[test]
    fn rejects_payload_and_metadata_reference_aliases_without_mutating()
    -> Result<(), InvalidationError> {
        let mut duplicate_payload_object = preview_object(false)?;
        duplicate_payload_object.messages[0]
            .data
            .extend_from_slice(&[0x4a, 0x02, 0x08, 77]);
        assert_rejected_without_mutation(duplicate_payload_object);

        let mut duplicate_payload_data = preview_object(false)?;
        duplicate_payload_data.messages[0]
            .data
            .extend_from_slice(&[0x82, 0x01, 0x02, 0x08, 55]);
        assert_rejected_without_mutation(duplicate_payload_data);

        let mut shared_metadata_object = preview_object(false)?;
        shared_metadata_object.archive_info.message_infos[0].field_infos[3]
            .object_references
            .push(77);
        assert_rejected_without_mutation(shared_metadata_object);

        let mut shared_metadata_data = preview_object(false)?;
        shared_metadata_data.archive_info.message_infos[0].field_infos[3]
            .data_references
            .push(55);
        assert_rejected_without_mutation(shared_metadata_data);

        let mut duplicate_unrelated_aggregate = preview_object(false)?;
        duplicate_unrelated_aggregate.archive_info.message_infos[0]
            .object_references
            .push(99);
        assert_rejected_without_mutation(duplicate_unrelated_aggregate);

        let mut duplicate_unrelated_data = preview_object(false)?;
        duplicate_unrelated_data.archive_info.message_infos[0]
            .data_references
            .push(56);
        assert_rejected_without_mutation(duplicate_unrelated_data);
        Ok(())
    }

    #[test]
    fn changed_invalidation_rejects_groups_without_mutating() -> Result<(), InvalidationError> {
        let mut recognized_group = preview_object(false)?;
        recognized_group.messages[0]
            .data
            .extend_from_slice(&[0x73, 0x74]);
        assert_rejected_without_mutation(recognized_group);

        let mut nested_group = preview_object(false)?;
        replace_once(
            &mut nested_group.messages[0].data,
            &[0x1a, 0x02, 0x08, 77],
            &[0x1a, 0x02, 0x0b, 0x0c],
        )?;
        assert_rejected_without_mutation(nested_group);

        let mut unknown_group = preview_object(false)?;
        unknown_group.messages[0]
            .data
            .extend_from_slice(&[0xd3, 0x01, 0xd4, 0x01]);
        assert_rejected_without_mutation(unknown_group);
        Ok(())
    }

    #[test]
    fn preserves_noncanonical_unknown_fields_and_read_only_noop_state()
    -> Result<(), InvalidationError> {
        let unknown = [0xf0, 0x81, 0x00, 0x80, 0x00];
        let mut object = preview_object(false)?;
        object.messages[0].data.extend_from_slice(&unknown);
        invalidate(&mut object, ArchiveLimits::default(), WireLimits::default())?;
        assert!(
            object.messages[0]
                .data
                .windows(unknown.len())
                .any(|field| field == unknown)
        );
        let invalidated = object.clone();
        assert!(is_invalidated(&object, WireLimits::default())?);
        assert_eq!(object, invalidated);
        Ok(())
    }

    fn assert_rejected_without_mutation(mut object: ArchiveObject) {
        let original = object.clone();
        assert!(is_invalidated(&object, WireLimits::default()).is_err());
        assert!(invalidate(&mut object, ArchiveLimits::default(), WireLimits::default()).is_err());
        assert_eq!(object, original);
    }

    pub(super) fn replace_once(
        data: &mut Vec<u8>,
        source: &[u8],
        replacement: &[u8],
    ) -> Result<(), InvalidationError> {
        let start = data
            .windows(source.len())
            .position(|bytes| bytes == source)
            .ok_or(InvalidationError::InvalidSource)?;
        data.splice(start..start + source.len(), replacement.iter().copied());
        Ok(())
    }

    pub(super) fn unbudgeted_test_error(error: BudgetedInvalidationError) -> InvalidationError {
        match error {
            BudgetedInvalidationError::Invalidation(inner) => inner,
            BudgetedInvalidationError::BudgetExceeded { .. } => InvalidationError::InvalidSource,
        }
    }

    fn empty_metadata_object(
        payload: Vec<u8>,
        empty_fields: usize,
    ) -> Result<ArchiveObject, InvalidationError> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data: payload,
            }],
        )?;
        let fields = &mut object.archive_info.message_infos[0].field_infos;
        fields
            .try_reserve_exact(empty_fields)
            .map_err(|_allocation| InvalidationError::InvalidSource)?;
        fields.extend((0..empty_fields).map(|_index| FieldInfo::default()));
        Ok(object)
    }

    fn large_nonselected_object() -> Result<ArchiveObject, InvalidationError> {
        const PAYLOAD_BYTES: usize = 256 * 1_024;
        const METADATA_ITEMS: usize = 2_048;
        let mut object = ArchiveObject::new(
            1,
            vec![
                RawMessage {
                    type_: SLIDE_NODE_MESSAGE_TYPE,
                    data: vec![0x70, 0x01],
                },
                RawMessage {
                    type_: 99,
                    data: vec![0xa5; PAYLOAD_BYTES],
                },
            ],
        )?;
        let info = &mut object.archive_info.message_infos[1];
        info.versions.extend(
            0..u32::try_from(METADATA_ITEMS)
                .map_err(|_conversion| InvalidationError::InvalidSource)?,
        );
        info.diff_merge_version.extend(
            0..u32::try_from(METADATA_ITEMS)
                .map_err(|_conversion| InvalidationError::InvalidSource)?,
        );
        info.diff_read_version.extend(
            0..u32::try_from(METADATA_ITEMS)
                .map_err(|_conversion| InvalidationError::InvalidSource)?,
        );
        for offset in 0..METADATA_ITEMS {
            let identifier = u64::try_from(offset)
                .ok()
                .and_then(|value| value.checked_add(10_000))
                .ok_or(InvalidationError::InvalidSource)?;
            info.object_references.push(identifier);
            info.data_references.push(identifier.saturating_add(10_000));
        }
        info.field_infos.push(FieldInfo {
            path: FieldPath::new(vec![27, 1, 2, 3]),
            object_references: info.object_references.clone(),
            data_references: info.data_references.clone(),
            known_field_version: (0..u32::try_from(METADATA_ITEMS)
                .map_err(|_conversion| InvalidationError::InvalidSource)?)
                .collect(),
            known_field_feature_identifier: Some("v".repeat(METADATA_ITEMS)),
            ..FieldInfo::default()
        });
        Ok(object)
    }

    fn scaled_exact_delta_work(distinct_children: usize) -> Result<usize, InvalidationError> {
        let mut source = preview_object(false)?;
        let info = &mut source.archive_info.message_infos[0];
        info.object_references
            .try_reserve_exact(distinct_children)
            .map_err(|_allocation| InvalidationError::InvalidSource)?;
        for offset in 0..distinct_children {
            let identifier = u64::try_from(offset)
                .ok()
                .and_then(|value| value.checked_add(1_000))
                .ok_or(InvalidationError::InvalidSource)?;
            let mut nested = vec![0x08];
            litchi_iwa_common::varint::encode_varint_into(&mut nested, identifier);
            source.messages[0].data.push(0x0a);
            litchi_iwa_common::varint::encode_varint_into(
                &mut source.messages[0].data,
                u64::try_from(nested.len())
                    .map_err(|_conversion| InvalidationError::InvalidSource)?,
            );
            source.messages[0].data.extend_from_slice(&nested);
            info.object_references.push(identifier);
        }
        let mut candidate = source.clone();
        invalidate(
            &mut candidate,
            ArchiveLimits::default(),
            WireLimits::default(),
        )?;
        let (matches, report) = exact_invalidation_delta(
            &source,
            &candidate,
            InvalidationDirection::Forward,
            WireLimits::default(),
        )?;
        assert!(matches);
        Ok(report.work())
    }

    pub(super) fn preview_object(
        shared_slide_reference: bool,
    ) -> Result<ArchiveObject, InvalidationError> {
        let mut payload = vec![0x20, 0x00, 0x30, 0x00, 0x38, 0x00];
        if shared_slide_reference {
            payload.extend_from_slice(&[0x12, 0x02, 0x08, 77]);
        }
        payload.extend_from_slice(&[0x1a, 0x02, 0x08, 77]);
        payload.extend_from_slice(UNKNOWN_FIELD_BYTES);
        payload.extend_from_slice(&[0x4a, 0x02, 0x08, 78]);
        payload.extend_from_slice(&[0x52, 0x00]);
        payload.extend_from_slice(&[0x70, 0x00]);
        payload.extend_from_slice(&[0x82, 0x01, 0x02, 0x08, 55]);
        payload.extend_from_slice(&[0xca, 0x01, 0x01, b'd']);
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data: payload,
            }],
        )?;
        object.archive_info.message_infos[0] = MessageInfo {
            object_references: vec![77, 78, 99],
            data_references: vec![55, 56],
            field_infos: vec![
                FieldInfo {
                    path: FieldPath::new(vec![DATABASE_THUMBNAIL_FIELD]),
                    object_references: vec![77],
                    ..FieldInfo::default()
                },
                FieldInfo {
                    path: FieldPath::new(vec![DATABASE_THUMBNAILS_FIELD]),
                    object_references: vec![78],
                    ..FieldInfo::default()
                },
                FieldInfo {
                    path: FieldPath::new(vec![THUMBNAILS_FIELD]),
                    data_references: vec![55],
                    ..FieldInfo::default()
                },
                FieldInfo {
                    path: FieldPath::new(vec![2]),
                    object_references: vec![99],
                    data_references: vec![56],
                    ..FieldInfo::default()
                },
            ],
            ..object.archive_info.message_infos[0].clone()
        };
        Ok(object)
    }
}
