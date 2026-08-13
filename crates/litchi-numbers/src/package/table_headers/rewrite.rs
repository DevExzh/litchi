use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireDescent,
        patch_nested_fields_batched_with_limits, preflight_wire_tree_with_limits,
    },
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};

use super::super::Package;
use super::error::{map_archive_error, map_candidate_read_error, map_core_error, map_wire_error};
use super::ownership::charge_work;
use super::resolve::{
    decode_settings, resolve_target, settings_from_snapshot, validate_message_metadata,
};
use super::{
    Error, FOOTER_ROWS_FIELD, HEADER_COLUMNS_FIELD, HEADER_COLUMNS_FROZEN_FIELD, HEADER_ROWS_FIELD,
    HEADER_ROWS_FROZEN_FIELD, Path, REPEATING_HEADER_COLUMNS_FIELD, REPEATING_HEADER_ROWS_FIELD,
    ROOT_PREVIEWS, Settings, Target,
};

fn rewritten_payload(source: &[u8], before: Settings, after: Settings) -> Result<Vec<u8>, Error> {
    if settings_from_snapshot(&decode_settings(source)?)? != before {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let paths = [
        [HEADER_ROWS_FIELD],
        [HEADER_COLUMNS_FIELD],
        [FOOTER_ROWS_FIELD],
        [HEADER_ROWS_FROZEN_FIELD],
        [HEADER_COLUMNS_FROZEN_FIELD],
        [REPEATING_HEADER_ROWS_FIELD],
        [REPEATING_HEADER_COLUMNS_FIELD],
    ];
    let before_values = [
        before
            .header_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before
            .header_columns
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before
            .footer_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before.header_rows_frozen.map(u64::from),
        before.header_columns_frozen.map(u64::from),
        before.repeating_header_rows_enabled.map(u64::from),
        before.repeating_header_columns_enabled.map(u64::from),
    ];
    let after_values = [
        after
            .header_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after
            .header_columns
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after
            .footer_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after.header_rows_frozen.map(u64::from),
        after.header_columns_frozen.map(u64::from),
        after.repeating_header_rows_enabled.map(u64::from),
        after.repeating_header_columns_enabled.map(u64::from),
    ];
    let edits = [
        NestedFieldEdit::new(
            &paths[0],
            before_values[0].is_some(),
            NestedFieldReplacement::Varint(after_values[0]),
        ),
        NestedFieldEdit::new(
            &paths[1],
            before_values[1].is_some(),
            NestedFieldReplacement::Varint(after_values[1]),
        ),
        NestedFieldEdit::new(
            &paths[2],
            before_values[2].is_some(),
            NestedFieldReplacement::Varint(after_values[2]),
        ),
        NestedFieldEdit::new(
            &paths[3],
            before_values[3].is_some(),
            NestedFieldReplacement::Varint(after_values[3]),
        ),
        NestedFieldEdit::new(
            &paths[4],
            before_values[4].is_some(),
            NestedFieldReplacement::Varint(after_values[4]),
        ),
        NestedFieldEdit::new(
            &paths[5],
            before_values[5].is_some(),
            NestedFieldReplacement::Varint(after_values[5]),
        ),
        NestedFieldEdit::new(
            &paths[6],
            before_values[6].is_some(),
            NestedFieldReplacement::Varint(after_values[6]),
        ),
    ];
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(map_wire_error)?
        .with_fields(source.len().max(1))
        .map_err(map_wire_error)?
        .with_nesting(WireLimits::MAX_NESTING)
        .map_err(map_wire_error)?
        .with_output_bytes(source.len().saturating_add(64).max(1))
        .map_err(map_wire_error)?
        .with_rewrite_work(source.len().saturating_mul(4).max(1))
        .map_err(map_wire_error)?;
    preflight_wire_tree_with_limits(source, limits, |_visit| Ok(WireDescent::Skip))
        .map_err(map_wire_error)?;
    let output =
        patch_nested_fields_batched_with_limits(source, &edits, limits).map_err(map_wire_error)?;
    if settings_from_snapshot(&decode_settings(&output)?)? != after {
        return Err(Error::Verification);
    }
    Ok(output)
}

pub(super) fn rewrite(
    source: &Package,
    target: Target,
    after: Settings,
    previews: &[&str],
) -> Result<(Package, Arc<[u8]>), Error> {
    let source_catalog = physical_source(source)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
    }
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    drop(stream);
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(object, target.message_index)?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let data = rewritten_payload(&message.data, target.settings, after)?;
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(data.len())
        .map_err(|_allocation| Error::Allocation {
            amount: data.len(),
            path: Path::Table {
                sheet: target.sheet_position,
                table: target.table_position,
            },
        })?;
    retained.extend_from_slice(&data);
    let retained_payload: Arc<[u8]> = retained.into();
    object
        .replace_message_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    drop(rewritten);
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    let selected = resolve_target(&candidate, target.sheet_position, target.table_position)?;
    if selected.settings != after {
        return Err(Error::Verification);
    }
    Ok((candidate, retained_payload))
}

pub(in crate::package) fn clone_selected_payload(
    source: &Package,
    target: Target,
) -> Result<Arc<[u8]>, Error> {
    let payload = selected_payload(source, target)?;
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(payload.len())
        .map_err(|_allocation| Error::Allocation {
            amount: payload.len(),
            path: Path::Table {
                sheet: target.sheet_position,
                table: target.table_position,
            },
        })?;
    retained.extend_from_slice(payload);
    Ok(retained.into())
}

pub(in crate::package) fn selected_payload(
    source: &Package,
    target: Target,
) -> Result<&[u8], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

pub(in crate::package) fn root_preview_deletions(
    source: &SourceCatalog,
) -> Result<Vec<&'static str>, Error> {
    let mut counts = [0usize; ROOT_PREVIEWS.len()];
    for entry in source.package().iter() {
        if let Some(index) = ROOT_PREVIEWS.iter().position(|name| *name == entry.name()) {
            counts[index] = counts[index].checked_add(1).ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        }
    }
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_allocation| Error::Allocation {
            amount: ROOT_PREVIEWS.len(),
            path: Path::Package,
        })?;
    for (name, count) in ROOT_PREVIEWS.into_iter().zip(counts) {
        match count {
            0 => {},
            1 => previews.push(name),
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        }
    }
    Ok(previews)
}

pub(in crate::package) fn verify_exact_locality(
    source: &Package,
    candidate: &Package,
    target: Target,
    source_previews: &[&str],
    expected_candidate_previews: usize,
    expected_payload: &[u8],
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let candidate_previews = root_preview_deletions(candidate_catalog)?;
    if candidate_previews.len() != expected_candidate_previews {
        return Err(Error::Verification);
    }
    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_previews.contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_previews.contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(left), Some(right)) if left.name() == right.name() => {
                let selected = source
                    .state
                    .components
                    .catalog()
                    .get_index(target.component_index)
                    .is_some_and(|component| component.name() == left.name());
                let preserved = if selected {
                    selected_package_member_preserved(left, right)
                } else {
                    package_member_preserved(left, right)
                };
                if !preserved {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    let before_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::Verification)?;
    let after_component = candidate
        .state
        .components
        .catalog()
        .get(before_component.name())
        .ok_or(Error::Verification)?;
    if before_component.archive().objects.len() != after_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (index, (left, right)) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
        .enumerate()
    {
        if index != target.object_index {
            if left.archive_info != right.archive_info || left.messages != right.messages {
                return Err(Error::Verification);
            }
            continue;
        }
        if left.messages.len() != right.messages.len()
            || left.archive_info.identifier != right.archive_info.identifier
            || left.archive_info.should_merge != right.archive_info.should_merge
            || left.archive_info.message_infos.len() != right.archive_info.message_infos.len()
        {
            return Err(Error::Verification);
        }
        for (message_index, (left_message, right_message)) in
            left.messages.iter().zip(&right.messages).enumerate()
        {
            if message_index == target.message_index {
                if left_message.type_ != right_message.type_
                    || expected_payload != right_message.data
                    || !message_info_preserved_except_length(
                        left.archive_info
                            .message_infos
                            .get(message_index)
                            .ok_or(Error::Verification)?,
                        right
                            .archive_info
                            .message_infos
                            .get(message_index)
                            .ok_or(Error::Verification)?,
                    )
                {
                    return Err(Error::Verification);
                }
            } else if left_message != right_message
                || left.archive_info.message_infos.get(message_index)
                    != right.archive_info.message_infos.get(message_index)
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

pub(in crate::package) fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

pub(in crate::package) fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_local_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let left = source.raw_record().local_record();
    let right = candidate.raw_record().local_record();
    let (Some(left_header), Some(right_header)) = (
        zip_local_header_length(left),
        zip_local_header_length(right),
    ) else {
        return false;
    };
    if left_header != right_header
        || left[..CRC_AND_SIZES.start] != right[..CRC_AND_SIZES.start]
        || left[CRC_AND_SIZES.end..left_header] != right[CRC_AND_SIZES.end..right_header]
    {
        return false;
    }
    let Some(left_end) = left_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= left.len())
    else {
        return false;
    };
    let Some(right_end) = right_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= right.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &left[left_end..],
        &right[right_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let left_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let right_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    left_prefix == right_prefix
        && source.len() == candidate.len()
        && source.len() >= left_prefix + 12
        && source[..left_prefix] == candidate[..right_prefix]
        && source[left_prefix + 12..] == candidate[right_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..OFFSET.start] == candidate[CRC_AND_SIZES.end..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

pub(in crate::package) fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    package
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

pub(in crate::package) fn preflight_transaction_work(
    source: &Package,
    retained_target: Option<&[u8]>,
) -> Result<usize, Error> {
    let maximum =
        usize::try_from(source.state.options.archive().max_total_bytes()).unwrap_or(usize::MAX);
    let source_bytes = source.source_bytes();
    let mut observed = 0;
    charge_work(&mut observed, source_bytes.len().saturating_mul(2), maximum)?;
    if let Some(target_bytes) = retained_target
        && target_bytes != source_bytes
    {
        charge_work(&mut observed, target_bytes.len().saturating_mul(2), maximum)?;
    }
    for object in source.state.components.iter_objects() {
        for message in &object.messages {
            charge_work(
                &mut observed,
                message.data.len().saturating_mul(16),
                maximum,
            )?;
        }
    }
    Ok(observed)
}
