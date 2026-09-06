//! Physical identity-clone and raw-header rewrite helpers for IWA objects.
//!
//! This module deliberately stays below format ownership. It only copies
//! source-authoritative archive metadata, rewrites known object references,
//! and leaves payload bytes/data references to the caller.

use super::*;

#[derive(Debug)]
pub(super) struct IdentityRemap {
    entries: Vec<(u64, u64)>,
}

impl IdentityRemap {
    fn get(&self, source: &u64) -> Option<&u64> {
        self.entries
            .binary_search_by_key(source, |(old, _)| *old)
            .ok()
            .map(|index| &self.entries[index].1)
    }

    pub(super) const fn len(&self) -> usize {
        self.entries.len()
    }
}

pub(super) fn validate_clone_arguments(
    source_identifier: u64,
    new_identifier: u64,
    object_remap: &[(u64, u64)],
    replacement_messages: &[RawMessage],
    source_messages: &[RawMessage],
    limits: Limits,
) -> Result<()> {
    if new_identifier == 0 {
        return Err(Error::invalid_archive(0, "clone object identifier is zero"));
    }
    if object_remap.len() > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            object_remap.len(),
            limits.max_metadata_items(),
        ));
    }
    if replacement_messages.len() != source_messages.len() {
        return Err(Error::invalid_archive(
            replacement_messages.len(),
            "clone replacement message count differs from source",
        ));
    }
    for (index, (source, replacement)) in
        source_messages.iter().zip(replacement_messages).enumerate()
    {
        if source.type_ != replacement.type_ {
            return Err(Error::invalid_archive(
                index,
                "clone replacement message type differs from source",
            ));
        }
        check_message_length(replacement.data.len(), limits)?;
        u32::try_from(replacement.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
    }
    if source_identifier == 0 {
        return Err(Error::invalid_archive(
            0,
            "source object identifier is zero",
        ));
    }
    Ok(())
}

pub(super) fn validate_clone_metadata_semantics(
    source: &ArchiveInfo,
    source_identifier: u64,
    new_identifier: u64,
) -> Result<()> {
    if source_identifier == new_identifier {
        return Ok(());
    }
    let has_diff_metadata = source.message_infos.iter().any(|message| {
        message.base_message_index.is_some()
            || !message.diff_merge_version.is_empty()
            || message.diff_field_path.is_some()
            || !message.fields_to_remove.is_empty()
            || !message.diff_read_version.is_empty()
    });
    if source.should_merge == Some(true) || has_diff_metadata {
        return Err(Error::invalid_archive(
            0,
            "cannot clone a changed-identity object carrying merge or diff metadata",
        ));
    }
    Ok(())
}

pub(super) fn validate_clone_self_reference_mapping(
    source: &ArchiveInfo,
    source_identifier: u64,
    new_identifier: u64,
    remap: &IdentityRemap,
) -> Result<()> {
    if source_identifier == new_identifier {
        return Ok(());
    }
    let has_self_reference = source.message_infos.iter().any(|message| {
        message.object_references.contains(&source_identifier)
            || message
                .field_infos
                .iter()
                .any(|field| field.object_references.contains(&source_identifier))
    });
    if has_self_reference && remap.get(&source_identifier).copied() != Some(new_identifier) {
        return Err(Error::invalid_archive(
            0,
            "clone self-reference requires an explicit source identity remap",
        ));
    }
    Ok(())
}

pub(super) fn validate_clone_input_scratch(
    remap_count: usize,
    replacement_count: usize,
    source_field_count: usize,
    limits: Limits,
) -> Result<()> {
    let remap_bytes = remap_count
        .checked_mul(size_of::<(u64, u64)>() * 2)
        .ok_or_else(|| Error::invalid_archive(0, "clone remap scratch size overflow"))?;
    let length_bytes = replacement_count
        .checked_mul(size_of::<u32>())
        .ok_or_else(|| Error::invalid_archive(0, "clone length scratch size overflow"))?;
    let wire_workspace = source_field_count
        .checked_mul(size_of::<WireField>() * 4)
        .ok_or_else(|| Error::invalid_archive(0, "clone wire workspace size overflow"))?;
    let rewrite_workspace = source_field_count
        .checked_mul(size_of::<HeaderFieldRewrite>() * 2)
        .ok_or_else(|| Error::invalid_archive(0, "clone rewrite workspace size overflow"))?;
    let total = remap_bytes
        .checked_add(length_bytes)
        .and_then(|value| value.checked_add(wire_workspace))
        .and_then(|value| value.checked_add(rewrite_workspace))
        .ok_or_else(|| Error::invalid_archive(0, "clone input scratch size overflow"))?;
    if total > limits.max_header_memory_bytes() {
        return Err(limit(
            LimitKind::HeaderMemoryBytes,
            total,
            limits.max_header_memory_bytes(),
        ));
    }
    Ok(())
}

pub(super) fn prepare_identity_remap(
    source_identifier: u64,
    new_identifier: u64,
    object_remap: &[(u64, u64)],
    limits: Limits,
) -> Result<IdentityRemap> {
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(object_remap.len())
        .map_err(|_| Error::allocation("IWA clone object-reference remap", object_remap.len()))?;
    for (old_identifier, replacement_identifier) in object_remap {
        if *old_identifier == 0 || *replacement_identifier == 0 {
            return Err(Error::invalid_archive(
                0,
                "clone object-reference remap contains a zero identifier",
            ));
        }
        if *old_identifier == source_identifier && *replacement_identifier != new_identifier {
            return Err(Error::invalid_archive(
                0,
                "clone source identity remap disagrees with new object identifier",
            ));
        }
        if *replacement_identifier == new_identifier && *old_identifier != source_identifier {
            return Err(Error::invalid_archive(
                0,
                "clone remap assigns the new object identifier to another source",
            ));
        }
        entries.push((*old_identifier, *replacement_identifier));
    }
    entries.sort_unstable_by_key(|(old, _)| *old);
    if entries.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::invalid_archive(
            0,
            "clone object-reference remap contains a duplicate source",
        ));
    }
    entries.sort_unstable_by_key(|(_, new)| *new);
    if entries.windows(2).any(|pair| pair[0].1 == pair[1].1) {
        return Err(Error::invalid_archive(
            0,
            "clone object-reference remap contains a duplicate target",
        ));
    }
    entries.sort_unstable_by_key(|(old, _)| *old);
    let remap = IdentityRemap { entries };
    if remap.len() > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            remap.len(),
            limits.max_metadata_items(),
        ));
    }
    Ok(remap)
}

pub(super) fn checked_replacement_lengths(
    messages: &[RawMessage],
    limits: Limits,
) -> Result<Vec<u32>> {
    let mut lengths = Vec::new();
    lengths
        .try_reserve_exact(messages.len())
        .map_err(|_| Error::allocation("IWA clone message lengths", messages.len()))?;
    for (index, message) in messages.iter().enumerate() {
        check_message_length(message.data.len(), limits)?;
        lengths.push(
            u32::try_from(message.data.len())
                .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?,
        );
    }
    Ok(lengths)
}

pub(super) fn clone_archive_info_with_identity_remap(
    source: &ArchiveInfo,
    new_identifier: u64,
    remap: &IdentityRemap,
    replacement_lengths: &[u32],
    limits: Limits,
) -> Result<ArchiveInfo> {
    if source.message_infos.len() != replacement_lengths.len() {
        return Err(Error::invalid_archive(
            replacement_lengths.len(),
            "clone replacement lengths differ from message metadata",
        ));
    }
    let mut message_infos = Vec::new();
    message_infos
        .try_reserve_exact(source.message_infos.len())
        .map_err(|_| Error::allocation("IWA clone message metadata", source.message_infos.len()))?;
    for (index, (message_info, replacement_length)) in source
        .message_infos
        .iter()
        .zip(replacement_lengths)
        .enumerate()
    {
        message_infos.push(clone_message_info_with_identity_remap(
            message_info,
            *replacement_length,
            remap,
            limits,
            index,
        )?);
    }
    Ok(ArchiveInfo {
        identifier: Some(new_identifier),
        message_infos,
        should_merge: source.should_merge,
    })
}

fn clone_message_info_with_identity_remap(
    source: &MessageInfo,
    replacement_length: u32,
    remap: &IdentityRemap,
    limits: Limits,
    message_index: usize,
) -> Result<MessageInfo> {
    let mut field_infos = Vec::new();
    field_infos
        .try_reserve_exact(source.field_infos.len())
        .map_err(|_| Error::allocation("IWA clone field metadata", source.field_infos.len()))?;
    for field in &source.field_infos {
        field_infos.push(clone_field_info_with_identity_remap(field, remap, limits)?);
    }

    let mut fields_to_remove = Vec::new();
    fields_to_remove
        .try_reserve_exact(source.fields_to_remove.len())
        .map_err(|_| {
            Error::allocation(
                "IWA clone removed field paths",
                source.fields_to_remove.len(),
            )
        })?;
    for path in &source.fields_to_remove {
        fields_to_remove.push(FieldPath {
            path: try_copy_slice(&path.path, "IWA clone removed field path")?,
        });
    }

    let message_info = MessageInfo {
        type_: source.type_,
        versions: try_copy_slice(&source.versions, "IWA clone message versions")?,
        length: replacement_length,
        field_infos,
        object_references: map_object_references(
            &source.object_references,
            remap,
            "IWA clone aggregate object references",
        )?,
        data_references: try_copy_slice(
            &source.data_references,
            "IWA clone aggregate data references",
        )?,
        base_message_index: source.base_message_index,
        diff_merge_version: try_copy_slice(
            &source.diff_merge_version,
            "IWA clone diff merge versions",
        )?,
        diff_field_path: source
            .diff_field_path
            .as_ref()
            .map(|path| -> Result<FieldPath> {
                Ok(FieldPath {
                    path: try_copy_slice(&path.path, "IWA clone diff field path")?,
                })
            })
            .transpose()?,
        fields_to_remove,
        diff_read_version: try_copy_slice(
            &source.diff_read_version,
            "IWA clone diff read versions",
        )?,
    };
    message_info
        .validate_with_limits(limits)
        .map_err(|error| match error {
            Error::InvalidArchive { .. } | Error::Limit { .. } | Error::Allocation { .. } => error,
            _ => Error::invalid_archive(message_index, "cloned MessageInfo is invalid"),
        })?;
    Ok(message_info)
}

fn clone_field_info_with_identity_remap(
    source: &FieldInfo,
    remap: &IdentityRemap,
    _limits: Limits,
) -> Result<FieldInfo> {
    Ok(FieldInfo {
        path: FieldPath {
            path: try_copy_slice(&source.path.path, "IWA clone field path")?,
        },
        r#type: source.r#type,
        unknown_field_rule: source.unknown_field_rule,
        object_references: map_object_references(
            &source.object_references,
            remap,
            "IWA clone field object references",
        )?,
        data_references: try_copy_slice(
            &source.data_references,
            "IWA clone field data references",
        )?,
        known_field_rule: source.known_field_rule,
        known_field_version: try_copy_slice(
            &source.known_field_version,
            "IWA clone known field versions",
        )?,
        known_field_feature_identifier: source
            .known_field_feature_identifier
            .as_deref()
            .map(|identifier| try_copy_string(identifier, "IWA clone field feature identifier"))
            .transpose()?,
    })
}

fn map_object_references(
    source: &[u64],
    remap: &IdentityRemap,
    resource: &'static str,
) -> Result<Vec<u64>> {
    let mut mapped = Vec::new();
    mapped
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    mapped.extend(
        source
            .iter()
            .map(|identifier| remap.get(identifier).copied().unwrap_or(*identifier)),
    );
    Ok(mapped)
}

pub(super) fn clone_raw_messages_with_limits(
    source: &[RawMessage],
    limits: Limits,
) -> Result<Vec<RawMessage>> {
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation("IWA clone messages", source.len()))?;
    for (index, message) in source.iter().enumerate() {
        check_message_length(message.data.len(), limits)?;
        let mut data = Vec::new();
        data.try_reserve_exact(message.data.len())
            .map_err(|_| Error::allocation("IWA clone message payload", message.data.len()))?;
        data.extend_from_slice(&message.data);
        if data.len() != message.data.len() {
            return Err(Error::invalid_archive(
                index,
                "cloned message payload length differs from source",
            ));
        }
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    Ok(messages)
}

pub(super) fn validate_clone_object_size(
    header_length: usize,
    replacement_messages: &[RawMessage],
    limits: Limits,
) -> Result<()> {
    check_header_length(header_length, limits)?;
    let mut object_length = varint_len(header_length)?
        .checked_add(header_length)
        .ok_or_else(|| Error::invalid_archive(0, "clone object prefix overflow"))?;
    for message in replacement_messages {
        check_message_length(message.data.len(), limits)?;
        object_length = object_length
            .checked_add(message.data.len())
            .ok_or_else(|| Error::invalid_archive(0, "clone object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_clone_header_scratch(
    canonical_before_length: usize,
    source_header_length: usize,
    rewritten_header_length: usize,
    source: &ArchiveInfo,
    remap_count: usize,
    replacement_count: usize,
    source_field_count: usize,
    limits: Limits,
) -> Result<()> {
    let metadata_bytes = clone_metadata_lower_bound(source)?;
    let mut scratch = 0usize;
    add_clone_scratch(&mut scratch, canonical_before_length, limits)?;
    add_clone_scratch(&mut scratch, source_header_length, limits)?;
    add_clone_scratch(
        &mut scratch,
        rewritten_header_length
            .checked_mul(3)
            .ok_or_else(|| Error::invalid_archive(0, "clone header scratch size overflow"))?,
        limits,
    )?;
    let wire_scratch = source_field_count
        .checked_mul(size_of::<WireField>() * 4)
        .and_then(|bytes| {
            source_field_count
                .checked_mul(size_of::<HeaderFieldRewrite>() * 2)
                .and_then(|rewrite_bytes| bytes.checked_add(rewrite_bytes))
        })
        .ok_or_else(|| Error::invalid_archive(0, "clone wire scratch size overflow"))?;
    add_clone_scratch(&mut scratch, wire_scratch, limits)?;
    add_clone_scratch(
        &mut scratch,
        metadata_bytes
            .checked_mul(2)
            .ok_or_else(|| Error::invalid_archive(0, "clone metadata scratch size overflow"))?,
        limits,
    )?;
    add_clone_scratch(
        &mut scratch,
        remap_count
            .checked_mul(size_of::<(u64, u64)>() * 2)
            .ok_or_else(|| Error::invalid_archive(0, "clone remap scratch size overflow"))?,
        limits,
    )?;
    add_clone_scratch(
        &mut scratch,
        replacement_count
            .checked_mul(size_of::<u32>())
            .ok_or_else(|| Error::invalid_archive(0, "clone length scratch size overflow"))?,
        limits,
    )?;
    Ok(())
}

fn add_clone_scratch(total: &mut usize, amount: usize, limits: Limits) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "clone header scratch size overflow"))?;
    if *total > limits.max_header_memory_bytes() {
        return Err(limit(
            LimitKind::HeaderMemoryBytes,
            *total,
            limits.max_header_memory_bytes(),
        ));
    }
    Ok(())
}

fn clone_metadata_lower_bound(source: &ArchiveInfo) -> Result<usize> {
    let mut total = size_of::<ArchiveInfo>();
    add_clone_metadata_bytes(
        &mut total,
        source.message_infos.len(),
        size_of::<MessageInfo>(),
    )?;
    for message in &source.message_infos {
        add_clone_metadata_bytes(&mut total, message.versions.len(), size_of::<u32>())?;
        add_clone_metadata_bytes(
            &mut total,
            message.field_infos.len(),
            size_of::<FieldInfo>(),
        )?;
        add_clone_metadata_bytes(
            &mut total,
            message.object_references.len(),
            size_of::<u64>(),
        )?;
        add_clone_metadata_bytes(&mut total, message.data_references.len(), size_of::<u64>())?;
        add_clone_metadata_bytes(
            &mut total,
            message.diff_merge_version.len(),
            size_of::<u32>(),
        )?;
        if let Some(path) = &message.diff_field_path {
            add_clone_metadata_bytes(&mut total, path.path.len(), size_of::<u32>())?;
        }
        add_clone_metadata_bytes(
            &mut total,
            message.fields_to_remove.len(),
            size_of::<FieldPath>(),
        )?;
        add_clone_metadata_bytes(
            &mut total,
            message.diff_read_version.len(),
            size_of::<u32>(),
        )?;
        for field in &message.field_infos {
            add_clone_metadata_bytes(&mut total, field.path.path.len(), size_of::<u32>())?;
            add_clone_metadata_bytes(&mut total, field.object_references.len(), size_of::<u64>())?;
            add_clone_metadata_bytes(&mut total, field.data_references.len(), size_of::<u64>())?;
            add_clone_metadata_bytes(
                &mut total,
                field.known_field_version.len(),
                size_of::<u32>(),
            )?;
            if let Some(identifier) = &field.known_field_feature_identifier {
                add_clone_metadata_bytes(&mut total, identifier.len(), size_of::<u8>())?;
            }
        }
        for path in &message.fields_to_remove {
            add_clone_metadata_bytes(&mut total, path.path.len(), size_of::<u32>())?;
        }
    }
    Ok(total)
}

fn add_clone_metadata_bytes(total: &mut usize, count: usize, size: usize) -> Result<()> {
    let bytes = count
        .checked_mul(size)
        .ok_or_else(|| Error::invalid_archive(0, "clone metadata size overflow"))?;
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| Error::invalid_archive(0, "clone metadata size overflow"))?;
    Ok(())
}

pub(super) fn retained_clone_source_header<'a>(
    object: &'a ArchiveObject,
    canonical: &'a [u8],
    limits: Limits,
) -> Result<&'a [u8]> {
    match (
        object.original_header.as_deref(),
        object.original_canonical_header.as_deref(),
    ) {
        (Some(original), Some(original_canonical)) if original_canonical == canonical => {
            preflight_header(original, HeaderKind::ArchiveInfo, limits)?;
            Ok(original)
        },
        (None, None) => {
            preflight_header(canonical, HeaderKind::ArchiveInfo, limits)?;
            Ok(canonical)
        },
        (Some(_), Some(_)) => Err(Error::invalid_archive(
            0,
            "retained ArchiveInfo header provenance is stale",
        )),
        _ => Err(Error::invalid_archive(
            0,
            "retained ArchiveInfo header provenance is incomplete",
        )),
    }
}

pub(super) fn preflight_clone_archive_info_header_length(
    source: &[u8],
    before: &ArchiveInfo,
    source_messages: &[RawMessage],
    replacement_messages: &[RawMessage],
    new_identifier: u64,
    replacement_lengths: &[u32],
    remap: &IdentityRemap,
    limits: Limits,
) -> Result<usize> {
    if before.message_infos.len() != replacement_lengths.len()
        || before.message_infos.len() != source_messages.len()
        || source_messages.len() != replacement_messages.len()
    {
        return Err(Error::invalid_archive(
            0,
            "clone ArchiveInfo and message counts differ",
        ));
    }
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let source_identifier = before
        .identifier
        .ok_or_else(|| Error::invalid_archive(0, "source object identifier is missing"))?;
    let identifier_field =
        effective_archive_identifier_field(source, &fields, source_identifier, 0)?;
    let mut output_length = source.len();
    if source_identifier != new_identifier {
        output_length = replace_varint_field_length(
            source,
            fields[identifier_field],
            new_identifier,
            0,
            output_length,
        )?;
    }

    let mut message_count = 0usize;
    for field in fields.iter().copied().filter(|field| field.number() == 2) {
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_count,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        let message_index = message_count;
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
        let before_message = before.message_infos.get(message_index).ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "raw MessageInfo count exceeds neutral metadata",
            )
        })?;
        let source_message = source_messages.get(message_index).ok_or_else(|| {
            Error::invalid_archive(message_index, "clone source message is missing")
        })?;
        let replacement_message = replacement_messages.get(message_index).ok_or_else(|| {
            Error::invalid_archive(message_index, "clone replacement message is missing")
        })?;
        let replacement_length =
            replacement_lengths
                .get(message_index)
                .copied()
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "clone replacement length is missing")
                })?;
        let message_source = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
        let message_length = preflight_clone_message_info_length(
            message_source,
            before_message,
            source_message,
            replacement_message,
            replacement_length,
            remap,
            wire_limits,
            message_index,
        )?;
        let original_length = field
            .raw(source)
            .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?
            .len();
        let rewritten_length =
            encoded_length_delimited_field_length(source, field, message_length, message_index)?;
        output_length = output_length
            .checked_sub(original_length)
            .and_then(|length| length.checked_add(rewritten_length))
            .ok_or_else(|| Error::invalid_archive(message_index, "clone header length overflow"))?;
    }
    if message_count != before.message_infos.len() {
        return Err(Error::invalid_archive(
            message_count,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    check_header_length(output_length, limits)?;
    Ok(output_length)
}

#[allow(
    clippy::too_many_arguments,
    reason = "The dry-run keeps source and target message metadata explicit for bounded sizing."
)]
fn preflight_clone_message_info_length(
    source: &[u8],
    before: &MessageInfo,
    source_message: &RawMessage,
    replacement_message: &RawMessage,
    replacement_length: u32,
    remap: &IdentityRemap,
    wire_limits: WireLimits,
    message_index: usize,
) -> Result<usize> {
    if source_message.type_ != replacement_message.type_ || source_message.type_ != before.type_ {
        return Err(Error::invalid_archive(
            message_index,
            "clone message type differs from neutral metadata",
        ));
    }
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let _type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(before.type_),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(before.length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;
    let mut output_length = source.len();
    if before.length != replacement_length {
        output_length = replace_varint_field_length(
            source,
            fields[length_field],
            u64::from(replacement_length),
            message_index,
            output_length,
        )?;
    }
    for field in fields.iter().copied() {
        match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let field_info = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                let rewritten_field_info_length = preflight_clone_field_info_length(
                    field_info,
                    remap,
                    wire_limits,
                    message_index,
                )?;
                output_length = replace_length_delimited_field_in_message_length(
                    source,
                    field,
                    rewritten_field_info_length,
                    message_index,
                    output_length,
                )?;
            },
            5 => {
                let rewritten_length =
                    mapped_reference_field_length(source, field, remap, message_index)?;
                output_length = replace_field_length(
                    source,
                    field,
                    rewritten_length,
                    message_index,
                    output_length,
                )?;
            },
            _ => {},
        }
    }
    Ok(output_length)
}

fn preflight_clone_field_info_length(
    source: &[u8],
    remap: &IdentityRemap,
    wire_limits: WireLimits,
    message_index: usize,
) -> Result<usize> {
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut output_length = source.len();
    for field in fields.iter().copied() {
        if field.number() != 4 {
            continue;
        }
        let rewritten_length = mapped_reference_field_length(source, field, remap, message_index)?;
        output_length = replace_field_length(
            source,
            field,
            rewritten_length,
            message_index,
            output_length,
        )?;
    }
    Ok(output_length)
}

fn replace_varint_field_length(
    source: &[u8],
    field: WireField,
    value: u64,
    message_index: usize,
    total: usize,
) -> Result<usize> {
    let key_length = field
        .key(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
        .len();
    let payload_length = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
        .len();
    let replacement = key_length
        .checked_add(encoded_varint_width(value, payload_length))
        .ok_or_else(|| Error::invalid_archive(message_index, "clone scalar length overflow"))?;
    replace_field_length_by_raw(source, field, replacement, message_index, total)
}

fn encoded_length_delimited_field_length(
    source: &[u8],
    field: WireField,
    payload_length: usize,
    message_index: usize,
) -> Result<usize> {
    let key_length = field
        .key(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
        .len();
    let prefix_width = field
        .payload_start()
        .checked_sub(field.key_end())
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "clone field prefix range is invalid")
        })?;
    let encoded_payload_length = u64::try_from(payload_length)
        .map_err(|_| Error::invalid_archive(message_index, "clone field payload exceeds u64"))?;
    key_length
        .checked_add(encoded_varint_width(encoded_payload_length, prefix_width))
        .and_then(|length| length.checked_add(payload_length))
        .ok_or_else(|| Error::invalid_archive(message_index, "clone field length overflow"))
}

fn replace_length_delimited_field_in_message_length(
    source: &[u8],
    field: WireField,
    payload_length: usize,
    message_index: usize,
    total: usize,
) -> Result<usize> {
    let replacement =
        encoded_length_delimited_field_length(source, field, payload_length, message_index)?;
    replace_field_length_by_raw(source, field, replacement, message_index, total)
}

fn replace_field_length(
    source: &[u8],
    field: WireField,
    replacement_length: usize,
    message_index: usize,
    total: usize,
) -> Result<usize> {
    replace_field_length_by_raw(source, field, replacement_length, message_index, total)
}

fn replace_field_length_by_raw(
    source: &[u8],
    field: WireField,
    replacement_length: usize,
    message_index: usize,
    total: usize,
) -> Result<usize> {
    let original_length = field
        .raw(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
        .len();
    total
        .checked_sub(original_length)
        .and_then(|length| length.checked_add(replacement_length))
        .ok_or_else(|| Error::invalid_archive(message_index, "clone header length overflow"))
}

fn mapped_reference_field_length(
    source: &[u8],
    field: WireField,
    remap: &IdentityRemap,
    message_index: usize,
) -> Result<usize> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| Error::invalid_archive(message_index, "malformed object reference"))?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "object reference has trailing bytes",
                ));
            }
            let replacement = remap.get(&value).copied().unwrap_or(value);
            if replacement == value {
                return Ok(field
                    .raw(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
                    .len());
            }
            let key_length = field
                .key(source)
                .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
                .len();
            key_length
                .checked_add(encoded_varint_width(replacement, encoded_length))
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "object reference length overflow")
                })
        },
        2 => mapped_packed_reference_field_length(source, payload, field, remap, message_index),
        _ => Err(Error::invalid_archive(
            message_index,
            "object-reference field has an ambiguous wire type",
        )),
    }
}

fn mapped_packed_reference_field_length(
    source: &[u8],
    payload: &[u8],
    field: WireField,
    remap: &IdentityRemap,
    message_index: usize,
) -> Result<usize> {
    let mut cursor = 0usize;
    let mut payload_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| {
                Error::invalid_archive(message_index, "malformed packed object reference")
            })?;
        let replacement = remap.get(&value).copied().unwrap_or(value);
        payload_length = payload_length
            .checked_add(encoded_varint_width(replacement, encoded_length))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference length overflow")
            })?;
        changed |= replacement != value;
        cursor = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range overflow")
        })?;
    }
    if !changed {
        return Ok(field
            .raw(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
            .len());
    }
    let key_length = field
        .key(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
        .len();
    let prefix_width = field
        .payload_start()
        .checked_sub(field.key_end())
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference prefix range is invalid")
        })?;
    key_length
        .checked_add(encoded_varint_width(
            u64::try_from(payload_length).map_err(|_| {
                Error::invalid_archive(message_index, "packed reference length exceeds u64")
            })?,
            prefix_width,
        ))
        .and_then(|length| length.checked_add(payload_length))
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference field length overflow")
        })
}

pub(super) fn rewrite_clone_archive_info_header(
    source: &[u8],
    before: &ArchiveInfo,
    source_messages: &[RawMessage],
    replacement_messages: &[RawMessage],
    new_identifier: u64,
    replacement_lengths: &[u32],
    remap: &IdentityRemap,
    limits: Limits,
) -> Result<Vec<u8>> {
    if before.message_infos.len() != replacement_lengths.len()
        || before.message_infos.len() != source_messages.len()
        || source_messages.len() != replacement_messages.len()
    {
        return Err(Error::invalid_archive(
            0,
            "clone ArchiveInfo and message counts differ",
        ));
    }
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let source_identifier = before
        .identifier
        .ok_or_else(|| Error::invalid_archive(0, "source object identifier is missing"))?;
    let identifier_field =
        effective_archive_identifier_field(source, &fields, source_identifier, 0)?;

    let mut rewrites = retained_field_rewrites(fields.len())?;
    if source_identifier != new_identifier {
        assign_header_field_rewrite(
            &mut rewrites,
            identifier_field,
            HeaderFieldRewrite::Varint(new_identifier),
            0,
        )?;
    }

    let mut message_count = 0usize;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_count,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        let message_index = message_count;
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
        let message_source = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
        let before_message = before.message_infos.get(message_index).ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "raw MessageInfo count exceeds neutral metadata",
            )
        })?;
        let replacement_length =
            replacement_lengths
                .get(message_index)
                .copied()
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "clone replacement length is missing")
                })?;
        let source_message = source_messages.get(message_index).ok_or_else(|| {
            Error::invalid_archive(message_index, "clone source message is missing")
        })?;
        let replacement_message = replacement_messages.get(message_index).ok_or_else(|| {
            Error::invalid_archive(message_index, "clone replacement message is missing")
        })?;
        let rewritten = rewrite_clone_message_info(
            message_source,
            before_message,
            source_message,
            replacement_message,
            replacement_length,
            remap,
            wire_limits,
            limits,
            message_index,
        )?;
        if let Some(rewritten) = rewritten {
            assign_header_field_rewrite(
                &mut rewrites,
                field_index,
                HeaderFieldRewrite::LengthDelimited(rewritten),
                message_index,
            )?;
        }
    }
    if message_count != before.message_infos.len() {
        return Err(Error::invalid_archive(
            message_count,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return try_copy_bytes(source, "IWA cloned ArchiveInfo header");
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        0,
        "IWA cloned ArchiveInfo header",
    )
}

fn effective_archive_identifier_field(
    source: &[u8],
    fields: &[WireField],
    expected: u64,
    message_index: usize,
) -> Result<usize> {
    let mut effective = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 1 {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous object identifier",
            ));
        }
        let payload = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
            .map_err(|_| Error::invalid_archive(message_index, "malformed object identifier"))?;
        if encoded_length != payload.len() {
            return Err(Error::invalid_archive(
                message_index,
                "object identifier has trailing bytes",
            ));
        }
        effective = Some((field_index, value));
    }
    let (field_index, value) = effective.ok_or_else(|| {
        Error::invalid_archive(message_index, "ArchiveInfo object identifier is missing")
    })?;
    if value != expected {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral object identifiers differ",
        ));
    }
    Ok(field_index)
}

#[allow(
    clippy::too_many_arguments,
    reason = "The physical clone rewrite keeps source and target metadata explicit for atomic verification."
)]
fn rewrite_clone_message_info(
    source: &[u8],
    before: &MessageInfo,
    source_message: &RawMessage,
    replacement_message: &RawMessage,
    replacement_length: u32,
    remap: &IdentityRemap,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<Vec<u8>>> {
    if source_message.type_ != replacement_message.type_ || source_message.type_ != before.type_ {
        return Err(Error::invalid_archive(
            message_index,
            "clone message type differs from neutral metadata",
        ));
    }
    let current_length = before.length;
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let _type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(before.type_),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    for (field_index, field) in fields.iter().copied().enumerate() {
        let rewrite = match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let field_info = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                rewrite_clone_field_info(field_info, remap, wire_limits, limits, message_index)?
            },
            5 => Some(rewrite_mapped_reference_field(
                source,
                field,
                remap,
                message_index,
            )?),
            _ => None,
        };
        if let Some(rewrite) = rewrite {
            assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
        }
    }
    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA cloned MessageInfo header",
    )
    .map(Some)
}

fn rewrite_clone_field_info(
    source: &[u8],
    remap: &IdentityRemap,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<HeaderFieldRewrite>> {
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 4 {
            continue;
        }
        let rewrite = rewrite_mapped_reference_field(source, field, remap, message_index)?;
        assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
    }
    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA cloned FieldInfo header",
    )
    .map(HeaderFieldRewrite::LengthDelimited)
    .map(Some)
}

fn rewrite_mapped_reference_field(
    source: &[u8],
    field: WireField,
    remap: &IdentityRemap,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| Error::invalid_archive(message_index, "malformed object reference"))?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "object reference has trailing bytes",
                ));
            }
            let replacement = remap.get(&value).copied().unwrap_or(value);
            Ok(if replacement == value {
                HeaderFieldRewrite::Retain
            } else {
                HeaderFieldRewrite::Varint(replacement)
            })
        },
        2 => rewrite_mapped_packed_references(payload, remap, message_index),
        _ => Err(Error::invalid_archive(
            message_index,
            "object-reference field has an ambiguous wire type",
        )),
    }
}

fn rewrite_mapped_packed_references(
    payload: &[u8],
    remap: &IdentityRemap,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let mut cursor = 0usize;
    let mut output_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| {
                Error::invalid_archive(message_index, "malformed packed object reference")
            })?;
        let replacement = remap.get(&value).copied().unwrap_or(value);
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range overflow")
        })?;
        output_length = output_length
            .checked_add(encoded_varint_width(replacement, encoded_length))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "packed object-reference length overflow")
            })?;
        changed |= replacement != value;
        cursor = end;
    }
    if !changed {
        return Ok(HeaderFieldRewrite::Retain);
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA cloned packed object references", output_length))?;
    cursor = 0;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| {
                Error::invalid_archive(message_index, "malformed packed object reference")
            })?;
        let replacement = remap.get(&value).copied().unwrap_or(value);
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed object-reference range overflow")
        })?;
        if replacement == value {
            output.extend_from_slice(payload.get(cursor..end).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed object-reference range is invalid")
            })?);
        } else {
            let mut encoded = [0u8; MAX_VARINT_BYTES];
            output.extend_from_slice(encode_varint_with_width(
                replacement,
                encoded_length,
                &mut encoded,
            ));
        }
        cursor = end;
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "packed object-reference rewrite length mismatch",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(output))
}
