//! Private invalidation of stale Keynote slide-preview caches.

use std::collections::HashSet;

use litchi_iwa_common::{
    Error as WireError, LimitKind as WireLimitKind, WireLimits, decode_varint_from_bytes,
    wire::WireView,
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

struct ScanBudget {
    limits: WireLimits,
    fields: usize,
    work: usize,
}

impl ScanBudget {
    const fn new(limits: WireLimits) -> Self {
        Self {
            limits,
            fields: 0,
            work: 0,
        }
    }

    fn parse<'source>(
        &mut self,
        source: &'source [u8],
    ) -> Result<WireView<'source>, InvalidationError> {
        self.add_work(source.len())?;
        let view = WireView::parse_with_limits(source, self.limits)?;
        let observed = self.fields.saturating_add(view.len());
        if observed > self.limits.max_fields() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::Fields,
                observed,
                limit: self.limits.max_fields(),
            }
            .into());
        }
        self.fields = observed;
        Ok(view)
    }

    fn add_work(&mut self, amount: usize) -> Result<(), InvalidationError> {
        let observed = self.work.saturating_add(amount);
        if observed > self.limits.max_rewrite_work() {
            return Err(WireError::LimitExceeded {
                kind: WireLimitKind::RewriteWork,
                observed,
                limit: self.limits.max_rewrite_work(),
            }
            .into());
        }
        self.work = observed;
        Ok(())
    }
}

struct PayloadRewrite {
    data: Vec<u8>,
    removed_object_references: Vec<u64>,
    removed_data_references: Vec<u64>,
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
    object.validate_with_limits(archive_limits)?;
    let message_index = selected_message_index(object)?;
    validate_metadata_topology(object, message_index)?;
    let original = object
        .messages
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    let mut budget = ScanBudget::new(wire_limits);
    let rewrite = rewrite_payload(&original.data, &mut budget)?;
    let removed_data_references = validate_reference_ownership(
        object,
        message_index,
        &rewrite.removed_object_references,
        &rewrite.removed_data_references,
    )?;
    validate_rewrite(&original.data, &rewrite.data, &mut budget)?;

    object.replace_message_pruning_references_preserving_header_with_limits(
        message_index,
        RawMessage {
            type_: SLIDE_NODE_MESSAGE_TYPE,
            data: rewrite.data,
        },
        &rewrite.removed_object_references,
        DataReferencePruning::Selected(&removed_data_references),
        archive_limits,
    )?;

    Ok(())
}

/// Verify the raw preview state without decoding or cloning the slide node.
pub(super) fn is_invalidated(
    object: &ArchiveObject,
    wire_limits: WireLimits,
) -> Result<bool, InvalidationError> {
    let message_index = selected_message_index(object)?;
    validate_metadata_topology(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    if info
        .field_infos
        .iter()
        .any(|field| is_preview_path(field.path.as_slice()) && !field.data_references.is_empty())
    {
        return Ok(false);
    }
    validate_retained_data_references(info)?;
    let payload = &object
        .messages
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?
        .data;
    let mut budget = ScanBudget::new(wire_limits);
    let view = budget.parse(payload)?;
    let mut dirty = None;
    for field in view.fields() {
        match field.number() {
            DATABASE_THUMBNAIL_FIELD
            | DATABASE_THUMBNAILS_FIELD
            | THUMBNAIL_SIZES_FIELD
            | THUMBNAILS_FIELD
            | THUMBNAIL_DIGESTS_FIELD => {
                require_length_delimited(field.wire_type())?;
                return Ok(false);
            },
            THUMBNAILS_DIRTY_FIELD => {
                if dirty.is_some() || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                dirty = Some(strict_varint(field.payload())? == 1);
            },
            _ => {},
        }
    }
    Ok(dirty == Some(true))
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

fn rewrite_payload(
    source: &[u8],
    budget: &mut ScanBudget,
) -> Result<PayloadRewrite, InvalidationError> {
    let view = budget.parse(source)?;
    let mut database_thumbnail_seen = false;
    let mut slide_reference_seen = false;
    let mut dirty_seen = false;
    let mut retained_bytes = 0usize;
    let mut removed_object_references = Vec::new();
    let mut removed_data_references = Vec::new();
    let mut surviving_object_references = Vec::new();

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

    for field in view.fields() {
        match field.number() {
            1 => {
                require_length_delimited(field.wire_type())?;
                surviving_object_references.push(reference_identifier(field.payload(), budget)?);
                retained_bytes = retained_bytes
                    .checked_add(field.raw().len())
                    .ok_or(InvalidationError::InvalidSource)?;
            },
            2 => {
                if std::mem::replace(&mut slide_reference_seen, true) {
                    return Err(InvalidationError::InvalidSource);
                }
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
                require_length_delimited(field.wire_type())?;
                removed_object_references.push(reference_identifier(field.payload(), budget)?);
            },
            DATABASE_THUMBNAILS_FIELD => {
                require_length_delimited(field.wire_type())?;
                removed_object_references.push(reference_identifier(field.payload(), budget)?);
            },
            THUMBNAILS_FIELD => {
                require_length_delimited(field.wire_type())?;
                removed_data_references.push(data_reference_identifier(field.payload(), budget)?);
            },
            THUMBNAIL_SIZES_FIELD | THUMBNAIL_DIGESTS_FIELD => {
                require_length_delimited(field.wire_type())?;
            },
            THUMBNAILS_DIRTY_FIELD => {
                if std::mem::replace(&mut dirty_seen, true) || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                let value = strict_varint(field.payload())?;
                if value > 1 {
                    return Err(InvalidationError::InvalidSource);
                }
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
    removed.extend(removed_object_references.iter().copied());
    if surviving_object_references
        .iter()
        .any(|identifier| removed.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
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
    if !dirty_seen {
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

fn reference_identifier(source: &[u8], budget: &mut ScanBudget) -> Result<u64, InvalidationError> {
    let view = budget.parse(source)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut deprecated_external_seen = false;
    for field in view.fields() {
        match field.number() {
            1 => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                identifier = Some(strict_varint(field.payload())?);
            },
            2 => {
                if std::mem::replace(&mut deprecated_type_seen, true) || field.wire_type() != 0 {
                    return Err(InvalidationError::InvalidSource);
                }
                let _value = strict_varint(field.payload())?;
            },
            3 => {
                validate_optional_bool_reference_field(field, &mut deprecated_external_seen)?;
            },
            _ => {},
        }
    }
    identifier.ok_or(InvalidationError::InvalidSource)
}

fn data_reference_identifier(
    source: &[u8],
    budget: &mut ScanBudget,
) -> Result<u64, InvalidationError> {
    let view = budget.parse(source)?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() == 1 {
            if identifier.is_some() || field.wire_type() != 0 {
                return Err(InvalidationError::InvalidSource);
            }
            identifier = Some(strict_varint(field.payload())?);
        }
    }
    identifier.ok_or(InvalidationError::InvalidSource)
}

fn validate_reference_ownership(
    object: &ArchiveObject,
    message_index: usize,
    object_identifiers: &[u64],
    payload_data_identifiers: &[u64],
) -> Result<Vec<u64>, InvalidationError> {
    let mut object_removals = HashSet::new();
    object_removals
        .try_reserve(object_identifiers.len())
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote slide preview ownership set",
            amount: object_identifiers.len(),
        })?;
    object_removals.extend(object_identifiers.iter().copied());
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?;
    for field in &info.field_infos {
        if field
            .object_references
            .iter()
            .any(|identifier| object_removals.contains(identifier))
            && !is_database_thumbnail_path(field.path.as_slice())
        {
            return Err(InvalidationError::InvalidSource);
        }
    }

    let preview_field_data_count = info
        .field_infos
        .iter()
        .filter(|field| is_preview_path(field.path.as_slice()))
        .try_fold(0usize, |count, field| {
            count.checked_add(field.data_references.len())
        })
        .ok_or(InvalidationError::InvalidSource)?;
    let preview_capacity = payload_data_identifiers
        .len()
        .checked_add(preview_field_data_count)
        .ok_or(InvalidationError::InvalidSource)?;
    let unrelated_capacity = info
        .field_infos
        .iter()
        .filter(|field| !is_preview_path(field.path.as_slice()))
        .try_fold(0usize, |count, field| {
            count.checked_add(field.data_references.len())
        })
        .ok_or(InvalidationError::InvalidSource)?;
    let mut preview = HashSet::new();
    preview
        .try_reserve(preview_capacity)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote preview-owned data-reference set",
            amount: preview_capacity,
        })?;
    let mut unrelated = HashSet::new();
    unrelated
        .try_reserve(unrelated_capacity)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote unrelated data-reference set",
            amount: unrelated_capacity,
        })?;
    let mut removals = Vec::new();
    try_reserve_references(
        &mut removals,
        preview_capacity,
        "Keynote preview data-reference removals",
    )?;
    for identifier in payload_data_identifiers.iter().copied().chain(
        info.field_infos
            .iter()
            .filter(|field| is_preview_path(field.path.as_slice()))
            .flat_map(|field| field.data_references.iter().copied()),
    ) {
        if preview.insert(identifier) {
            removals.push(identifier);
        }
    }
    for identifier in info
        .field_infos
        .iter()
        .filter(|field| !is_preview_path(field.path.as_slice()))
        .flat_map(|field| field.data_references.iter().copied())
    {
        unrelated.insert(identifier);
    }
    if preview
        .iter()
        .any(|identifier| unrelated.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
    for identifier in &info.data_references {
        if !preview.contains(identifier) && !unrelated.contains(identifier) {
            // Aggregate metadata cannot prove whether this identifier belongs
            // to a removed preview field or to surviving opaque content.
            return Err(InvalidationError::InvalidSource);
        }
    }
    Ok(removals)
}

fn validate_retained_data_references(
    info: &litchi_iwa_core::MessageInfo,
) -> Result<(), InvalidationError> {
    let capacity = info
        .field_infos
        .iter()
        .filter(|field| !is_preview_path(field.path.as_slice()))
        .try_fold(0usize, |count, field| {
            count.checked_add(field.data_references.len())
        })
        .ok_or(InvalidationError::InvalidSource)?;
    let mut supported = HashSet::new();
    supported
        .try_reserve(capacity)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote retained data-reference set",
            amount: capacity,
        })?;
    supported.extend(
        info.field_infos
            .iter()
            .filter(|field| !is_preview_path(field.path.as_slice()))
            .flat_map(|field| field.data_references.iter().copied()),
    );
    if info
        .data_references
        .iter()
        .any(|identifier| !supported.contains(identifier))
    {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(())
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

fn validate_optional_bool_reference_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    seen: &mut bool,
) -> Result<(), InvalidationError> {
    if std::mem::replace(seen, true)
        || field.wire_type() != 0
        || strict_varint(field.payload())? > 1
    {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(())
}

fn strict_varint(payload: &[u8]) -> Result<u64, InvalidationError> {
    let (value, encoded) =
        decode_varint_from_bytes(payload).map_err(|_error| InvalidationError::InvalidSource)?;
    if encoded == payload.len() {
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
        assert_eq!(info.field_infos[0].object_references, [99]);
        assert_eq!(info.field_infos[1].object_references, [99]);
        assert!(info.field_infos[0].data_references.is_empty());
        assert_eq!(info.field_infos[1].data_references, [56]);
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

    fn preview_object(shared_slide_reference: bool) -> Result<ArchiveObject, InvalidationError> {
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
                    object_references: vec![77, 99],
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
