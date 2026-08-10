use std::sync::Arc;

use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::wire::{
    NestedFieldEdit, NestedFieldReplacement, WireView, patch_nested_fields_batched_with_limits,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};

use super::media::next_raw_field;
use super::{
    Package, ReopenCost, SOUNDTRACK_MESSAGE_TYPE, Selection, TransactionBudget,
    charge_message_info, map_archive_error, map_core_error, map_read_error, map_wire_error,
    physical_catalog, remaining_wire_limits, select, selected_message,
};
use crate::package::rendering_invalidation;
use crate::soundtrack::{Commit, Diagnostics, Error, LimitKind, Patch, Settings};

pub(super) fn validate_component_framing(
    package: &Package,
    component_name: &str,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let catalog = physical_catalog(package)?;
    let component = catalog
        .components()
        .get(component_name)
        .ok_or(Error::InvalidSource)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let snappy_limits = transaction_snappy_limits(package, budget, entry.data().len(), 2)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|error| map_transaction_snappy_error(error, budget))?;
    budget.charge_work(
        entry
            .data()
            .len()
            .checked_add(stream.as_bytes().len().saturating_mul(2))
            .ok_or(Error::InvalidSource)?,
    )?;
    component
        .archive()
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)
}

fn transaction_snappy_limits(
    package: &Package,
    budget: &TransactionBudget,
    compressed_bytes: usize,
    decoded_passes: usize,
) -> Result<litchi_iwa_core::SnappyLimits, Error> {
    let base = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let remaining_work = budget.remaining_work();
    let decoded_allowance = remaining_work
        .checked_sub(compressed_bytes)
        .map_or(0, |available| available / decoded_passes);
    if decoded_allowance == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget
                .work
                .saturating_add(compressed_bytes)
                .saturating_add(decoded_passes) as u64,
            maximum: budget.max_work as u64,
        });
    }
    litchi_iwa_core::SnappyLimits::new(
        base.max_uncompressed_chunk().min(decoded_allowance),
        decoded_allowance,
    )
    .and_then(|limits| {
        limits.with_input_limits(
            base.max_compressed_chunk(),
            base.max_compressed_stream(),
            base.max_frames(),
        )
    })
    .map_err(map_core_error)
}

fn map_transaction_snappy_error(
    error: litchi_iwa_core::Error,
    budget: &TransactionBudget,
) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind:
                litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes,
            observed,
            ..
        } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.work.saturating_add(observed) as u64,
            maximum: budget.max_work as u64,
        },
        other @ (litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::Limit { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. }
        | litchi_iwa_core::Error::Allocation { .. }) => map_core_error(other),
    }
}

pub(super) fn rewrite_and_verify(
    source: &Package,
    source_bytes: Arc<[u8]>,
    selection: &Selection<'_>,
    before: Settings,
    after: Settings,
    budget: &mut TransactionBudget,
) -> Result<Commit, Error> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.soundtrack_component)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = transaction_snappy_limits(source, budget, entry.data().len(), 3)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|error| map_transaction_snappy_error(error, budget))?;
    let source_stream_len = stream.as_bytes().len();
    budget.charge_work(
        entry
            .data()
            .len()
            .checked_add(source_stream_len.saturating_mul(3))
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    drop(stream);
    let object = archive
        .object(selection.soundtrack_identifier)
        .ok_or(Error::InvalidSource)?;
    let (message_index, payload) = selected_message(object, SOUNDTRACK_MESSAGE_TYPE)?;
    if message_index != selection.soundtrack_message_index
        || payload != selection.soundtrack_payload
    {
        return Err(Error::InvalidSource);
    }
    let rewritten = rewrite_payload(source, payload, before, after, budget)?;
    let rewritten_len = rewritten.len();
    let payload_growth = rewritten_len.saturating_sub(payload.len());
    archive
        .object_mut(selection.soundtrack_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SOUNDTRACK_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let serialized_bound = source_stream_len
        .checked_add(payload_growth)
        .and_then(|amount| amount.checked_add(32))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(
        serialized_bound
            .checked_mul(3)
            .ok_or(Error::InvalidSource)?,
    )?;
    let serialized = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    drop(archive);
    if serialized.len() > serialized_bound {
        return Err(Error::Verification);
    }
    let serialized_len = serialized.len();
    let compressed = SnappyStream::compress(&serialized).map_err(map_core_error)?;
    let compressed_len = compressed.len();
    drop(serialized);
    let package_bound = source_bytes
        .len()
        .checked_add(compressed.len().saturating_mul(2))
        .and_then(|amount| amount.checked_add(1_024))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(package_bound.checked_mul(3).ok_or(Error::InvalidSource)?)?;
    let output = catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(selection.soundtrack_component, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let target_bytes: Arc<[u8]> = output.into();
    if target_bytes.len() > package_bound {
        return Err(Error::Verification);
    }
    let source_reopen = catalog_reopen_cost(catalog, source_bytes.len(), None)?;
    let target_reopen = catalog_reopen_cost(
        catalog,
        target_bytes.len(),
        Some((
            selection.soundtrack_component,
            compressed_len,
            serialized_len,
        )),
    )?;
    budget.charge_work(target_reopen.work)?;
    budget.charge_references(target_reopen.references)?;
    let candidate =
        Package::from_source_with_options(Arc::clone(&target_bytes), source.state.options)
            .map_err(map_read_error)?;
    let touched = verify_candidate(source, &candidate, selection, after, budget)?;
    budget.charge_work(
        source_bytes
            .len()
            .checked_add(target_bytes.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target_bytes),
        before,
        after,
        touched_components: touched,
        source_reopen_work: source_reopen.work,
        target_reopen_work: target_reopen.work,
        source_reopen_references: source_reopen.references,
        target_reopen_references: target_reopen.references,
    };
    Ok(Commit {
        package: candidate,
        patch,
        diagnostics: Diagnostics {
            changed: true,
            touched_components: touched,
            full_reparse_performed: true,
        },
    })
}

fn catalog_reopen_cost(
    catalog: &SourceCatalog,
    raw_package_bytes: usize,
    replacement: Option<(&str, usize, usize)>,
) -> Result<ReopenCost, Error> {
    let mut logical_bytes = 0usize;
    for entry in catalog.package().iter() {
        let bytes = replacement
            .filter(|(name, _, _)| *name == entry.name())
            .map_or(entry.data().len(), |(_, logical, _)| logical);
        logical_bytes = logical_bytes
            .checked_add(bytes)
            .ok_or(Error::InvalidSource)?;
    }
    let mut iwa_bytes = 0usize;
    for component in catalog.components().iter() {
        let bytes = replacement
            .filter(|(name, _, _)| *name == component.name())
            .map_or_else(
                || archive_stream_extent(component.archive()),
                |(_, _, decoded)| decoded,
            );
        iwa_bytes = iwa_bytes.checked_add(bytes).ok_or(Error::InvalidSource)?;
    }
    let physical_work = raw_package_bytes
        .checked_add(logical_bytes.saturating_mul(2))
        .and_then(|amount| amount.checked_add(iwa_bytes.saturating_mul(2)))
        .ok_or(Error::InvalidSource)?;
    let mut structural_work = 0usize;
    let mut references = 0usize;
    for component in catalog.components().iter() {
        let cost = archive_structure_cost(component.archive())?;
        structural_work = structural_work
            .checked_add(cost.work)
            .ok_or(Error::InvalidSource)?;
        references = references
            .checked_add(cost.references)
            .ok_or(Error::InvalidSource)?;
    }
    Ok(ReopenCost {
        work: physical_work
            .checked_add(structural_work)
            .ok_or(Error::InvalidSource)?,
        references,
    })
}

fn archive_structure_cost(archive: &Archive) -> Result<ReopenCost, Error> {
    let mut work = archive.objects.len();
    let mut references = 0usize;
    for object in &archive.objects {
        work = work
            .checked_add(object.messages.len())
            .and_then(|amount| amount.checked_add(object.archive_info.message_infos.len()))
            .ok_or(Error::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            let aggregate_references = info
                .object_references
                .len()
                .checked_add(info.data_references.len())
                .ok_or(Error::InvalidSource)?;
            references = references
                .checked_add(aggregate_references)
                .ok_or(Error::InvalidSource)?;
            work = work
                .checked_add(info.versions.len())
                .and_then(|amount| amount.checked_add(info.diff_merge_version.len()))
                .and_then(|amount| amount.checked_add(info.diff_read_version.len()))
                .and_then(|amount| amount.checked_add(aggregate_references))
                .and_then(|amount| {
                    amount.checked_add(
                        info.diff_field_path
                            .as_ref()
                            .map_or(0, |path| path.path.len()),
                    )
                })
                .ok_or(Error::InvalidSource)?;
            for path in &info.fields_to_remove {
                work = work
                    .checked_add(1)
                    .and_then(|amount| amount.checked_add(path.path.len()))
                    .ok_or(Error::InvalidSource)?;
            }
            for field in &info.field_infos {
                let field_references = field
                    .object_references
                    .len()
                    .checked_add(field.data_references.len())
                    .ok_or(Error::InvalidSource)?;
                references = references
                    .checked_add(field_references)
                    .ok_or(Error::InvalidSource)?;
                work = work
                    .checked_add(1)
                    .and_then(|amount| amount.checked_add(field.path.path.len()))
                    .and_then(|amount| amount.checked_add(field.known_field_version.len()))
                    .and_then(|amount| amount.checked_add(field_references))
                    .and_then(|amount| {
                        amount.checked_add(
                            field
                                .known_field_feature_identifier
                                .as_ref()
                                .map_or(0, String::len),
                        )
                    })
                    .ok_or(Error::InvalidSource)?;
            }
        }
    }
    Ok(ReopenCost { work, references })
}

fn archive_stream_extent(archive: &Archive) -> usize {
    archive
        .objects
        .iter()
        .filter_map(|object| {
            let payload = object.messages.iter().try_fold(0usize, |amount, message| {
                amount.checked_add(message.data.len())
            })?;
            usize::try_from(object.data_offset)
                .ok()?
                .checked_add(payload)
        })
        .max()
        .unwrap_or(0)
}

fn rewrite_payload(
    package: &Package,
    source: &[u8],
    before: Settings,
    after: Settings,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    drop(view);
    let edit_fields = [
        NestedFieldEdit::new(
            &[1],
            before.volume().is_some(),
            NestedFieldReplacement::Fixed64(after.volume().map(f64::to_bits)),
        ),
        NestedFieldEdit::new(
            &[2],
            before.mode().is_some(),
            NestedFieldReplacement::Varint(
                after
                    .mode()
                    .map(|mode| i64::from(mode.as_raw()).cast_unsigned()),
            ),
        ),
    ];
    let bounded = limits
        .with_fields(budget.remaining_fields())
        .and_then(|field_bounded| field_bounded.with_rewrite_work(budget.remaining_work()))
        .map_err(map_wire_error)?;
    let output = patch_nested_fields_batched_with_limits(source, &edit_fields, bounded)
        .map_err(map_wire_error)?;
    budget.charge_work(
        source
            .len()
            .checked_add(output.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    Ok(output)
}

pub(super) fn verify_candidate(
    source: &Package,
    candidate: &Package,
    source_selection: &Selection<'_>,
    expected: Settings,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    let target = select(candidate, budget)?.ok_or(Error::Verification)?;
    if target.show_identifier != source_selection.show_identifier
        || target.soundtrack_identifier != source_selection.soundtrack_identifier
        || target.soundtrack_component != source_selection.soundtrack_component
        || target.settings != expected
    {
        return Err(Error::Verification);
    }
    verify_package_members(
        source,
        candidate,
        source_selection.soundtrack_component,
        budget,
    )?;
    verify_selected_component(source, candidate, source_selection, &target, budget)?;
    Ok(1)
}

fn verify_package_members(
    source: &Package,
    candidate: &Package,
    selected_name: &str,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let mut before = source_catalog.package().iter();
    let mut after = candidate_catalog.package().iter();
    loop {
        match (before.next(), after.next()) {
            (Some(old), Some(new)) if old.name() == new.name() => {
                budget.charge_work(
                    old.raw_record()
                        .local_record()
                        .len()
                        .checked_add(new.raw_record().local_record().len())
                        .and_then(|amount| {
                            amount.checked_add(old.raw_record().central_directory_record().len())
                        })
                        .and_then(|amount| {
                            amount.checked_add(new.raw_record().central_directory_record().len())
                        })
                        .ok_or(Error::Verification)?,
                )?;
                let preserved = if old.name() == selected_name {
                    selected_package_member_preserved(old, new)
                } else {
                    package_member_preserved(old, new)
                };
                if !preserved {
                    return Err(Error::Verification);
                }
            },
            (None, None) => return Ok(()),
            _ => return Err(Error::Verification),
        }
    }
}

fn verify_selected_component(
    source: &Package,
    candidate: &Package,
    before: &Selection<'_>,
    after: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let old_component = physical_catalog(source)?
        .components()
        .get(before.soundtrack_component)
        .ok_or(Error::Verification)?;
    let new_component = physical_catalog(candidate)?
        .components()
        .get(after.soundtrack_component)
        .ok_or(Error::Verification)?;
    if old_component.archive().objects.len() != new_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (old, new) in old_component
        .archive()
        .objects
        .iter()
        .zip(&new_component.archive().objects)
    {
        charge_object_comparison(old, budget)?;
        charge_object_comparison(new, budget)?;
        if old.archive_info.identifier != new.archive_info.identifier {
            return Err(Error::Verification);
        }
        if old.archive_info.identifier == Some(before.soundtrack_identifier) {
            verify_soundtrack_object(old, new, before, after, budget)?;
        } else if old.archive_info != new.archive_info || old.messages != new.messages {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn charge_object_comparison(
    object: &ArchiveObject,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.charge_work(1)?;
    for (index, message) in object.messages.iter().enumerate() {
        budget.charge_work(message.data.len().saturating_add(1))?;
        charge_message_info(object, index, budget)?;
    }
    Ok(())
}

fn verify_soundtrack_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    before: &Selection<'_>,
    after: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (old_index, old_payload) = selected_message(source, SOUNDTRACK_MESSAGE_TYPE)?;
    let (new_index, new_payload) = selected_message(candidate, SOUNDTRACK_MESSAGE_TYPE)?;
    if old_index != new_index
        || old_index != before.soundtrack_message_index
        || new_index != after.soundtrack_message_index
        || source.messages.len() != candidate.messages.len()
    {
        return Err(Error::Verification);
    }
    for (index, (old, new)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if old.type_ != new.type_ || (index != old_index && old != new) {
            return Err(Error::Verification);
        }
    }
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(Error::Verification);
    }
    for (index, (old, new)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        if index == old_index {
            if !message_info_preserved_except_length(old, new) {
                return Err(Error::Verification);
            }
        } else if old != new {
            return Err(Error::Verification);
        }
    }
    if !payload_delta_matches(
        old_payload,
        new_payload,
        before.settings,
        after.settings,
        budget,
    )? {
        return Err(Error::Verification);
    }
    Ok(())
}

fn payload_delta_matches(
    source: &[u8],
    candidate: &[u8],
    before: Settings,
    after: Settings,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    budget.charge_work(source.len())?;
    budget.charge_work(candidate.len())?;
    let mut source_input = source;
    let mut candidate_input = candidate;
    while let Some(old) = next_raw_field(&mut source_input, budget)? {
        match old.number {
            1 => {
                if let Some(volume) = after.volume() {
                    let Some(new) = next_raw_field(&mut candidate_input, budget)? else {
                        return Ok(false);
                    };
                    if new.number != 1
                        || new.wire != 1
                        || new.bytes != Some(volume.to_bits().to_le_bytes().as_slice())
                    {
                        return Ok(false);
                    }
                }
            },
            2 => {
                if let Some(mode) = after.mode() {
                    let Some(new) = next_raw_field(&mut candidate_input, budget)? else {
                        return Ok(false);
                    };
                    if new.number != 2
                        || new.wire != 0
                        || new.varint != Some(i64::from(mode.as_raw()).cast_unsigned())
                    {
                        return Ok(false);
                    }
                }
            },
            _ => {
                if next_raw_field(&mut candidate_input, budget)?.map(|field| field.raw)
                    != Some(old.raw)
                {
                    return Ok(false);
                }
            },
        }
    }
    for (field, value) in [
        (1, before.volume().is_none() && after.volume().is_some()),
        (2, before.mode().is_none() && after.mode().is_some()),
    ] {
        if !value {
            continue;
        }
        let Some(appended) = next_raw_field(&mut candidate_input, budget)? else {
            return Ok(false);
        };
        if appended.number != field {
            return Ok(false);
        }
        if field == 1 {
            if appended.wire != 1
                || appended.bytes
                    != Some(
                        after
                            .volume()
                            .ok_or(Error::Verification)?
                            .to_bits()
                            .to_le_bytes()
                            .as_slice(),
                    )
            {
                return Ok(false);
            }
        } else if appended.wire != 0
            || appended.varint
                != Some(
                    i64::from(after.mode().ok_or(Error::Verification)?.as_raw()).cast_unsigned(),
                )
        {
            return Ok(false);
        }
    }
    Ok(next_raw_field(&mut candidate_input, budget)?.is_none())
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

fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && rendering_invalidation::central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
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
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    let Some(source_header) = zip_local_header_length(source_record) else {
        return false;
    };
    let Some(candidate_header) = zip_local_header_length(candidate_record) else {
        return false;
    };
    if source_header != candidate_header
        || source_record[..CRC_AND_SIZES.start] != candidate_record[..CRC_AND_SIZES.start]
        || source_record[CRC_AND_SIZES.end..source_header]
            != candidate_record[CRC_AND_SIZES.end..candidate_header]
    {
        return false;
    }
    let Some(source_end) = source_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= source_record.len())
    else {
        return false;
    };
    let Some(candidate_end) = candidate_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= candidate_record.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &source_record[source_end..],
        &candidate_record[candidate_end..],
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
    let source_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let candidate_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    source_prefix == candidate_prefix
        && source.len() == candidate.len()
        && source.len() >= source_prefix + 12
        && source[..source_prefix] == candidate[..candidate_prefix]
        && source[source_prefix + 12..] == candidate[candidate_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const LOCAL_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..LOCAL_OFFSET.start]
            == candidate[CRC_AND_SIZES.end..LOCAL_OFFSET.start]
        && source[LOCAL_OFFSET.end..] == candidate[LOCAL_OFFSET.end..]
}
