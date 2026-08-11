use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, wire::parse_wire_fields_with_limits};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::numbers_sheet_order_codec::ReferenceSnapshot;

use super::super::Package;
use super::error::{map_archive_error, map_candidate_read_error, map_core_error, map_wire_error};
use super::resolve::{NativeTarget, TransactionBudget, resolve_native_target};
use super::{Error, ROOT_PREVIEWS};

pub(super) struct Rewritten {
    pub(super) package: Package,
    pub(super) target: NativeTarget,
    pub(super) source_reopen: ReopenCost,
    pub(super) target_reopen: ReopenCost,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ReorderReport {
    work: usize,
    references: usize,
}

struct RawReorder {
    data: Vec<u8>,
    report: ReorderReport,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct ReopenCost {
    pub(super) work: usize,
    pub(super) references: usize,
}

pub(super) fn rewrite(
    source: &Package,
    target: &NativeTarget,
    budget: &mut TransactionBudget,
    source_position: usize,
    destination_position: usize,
    previews: &[&str],
) -> Result<Rewritten, Error> {
    if target.document.component_index != target.sidebar_root.component_index {
        return Err(Error::UnsupportedSource);
    }
    let source_catalog = physical_source(source)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.document.component_index)
        .ok_or(Error::InvalidSource)?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
    }
    let decoded_extent = archive_extent(component.archive())?;
    let decoded_allocation = archive_allocation_cost(component.archive())?;
    budget.charge_transaction_work(
        entry
            .data()
            .len()
            .checked_add(decoded_extent.checked_mul(3).ok_or(Error::InvalidSource)?)
            .and_then(|work| work.checked_add(decoded_allocation))
            .ok_or(Error::InvalidSource)?,
    )?;
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

    let document_payload = selected_payload(&archive, target.document)?;
    let expected_document_report = reorder_report(
        target.document_fields,
        target.document_snapshot.sheet_references().len(),
        document_payload.len(),
    )?;
    budget.charge_wire(target.document_fields, expected_document_report.work)?;
    let rewritten_document = reorder_raw_records(
        document_payload,
        1,
        target.document_snapshot.sheet_references(),
        source_position,
        destination_position,
    )?;
    if rewritten_document.report.references != expected_document_report.references
        || rewritten_document.report.work > expected_document_report.work
    {
        return Err(Error::Verification);
    }
    let reordered_sheet_identifiers = reordered_identifiers(
        target.document_snapshot.sheet_references(),
        source_position,
        destination_position,
    )?;
    rewrite_owner(
        &mut archive,
        target.document,
        rewritten_document.data,
        &reordered_sheet_identifiers,
        archive_limits,
        budget,
    )?;
    drop(reordered_sheet_identifiers);

    let sidebar_payload = selected_payload(&archive, target.sidebar_root)?;
    let expected_sidebar_report = reorder_report(
        target.sidebar_fields,
        target.sidebar_snapshot.child_references().len(),
        sidebar_payload.len(),
    )?;
    budget.charge_wire(target.sidebar_fields, expected_sidebar_report.work)?;
    let rewritten_sidebar = reorder_raw_records(
        sidebar_payload,
        2,
        target.sidebar_snapshot.child_references(),
        source_position,
        destination_position,
    )?;
    if rewritten_sidebar.report.references != expected_sidebar_report.references
        || rewritten_sidebar.report.work > expected_sidebar_report.work
    {
        return Err(Error::Verification);
    }
    let reordered_child_identifiers = reordered_identifiers(
        target.sidebar_snapshot.child_references(),
        source_position,
        destination_position,
    )?;
    rewrite_owner(
        &mut archive,
        target.sidebar_root,
        rewritten_sidebar.data,
        &reordered_child_identifiers,
        archive_limits,
        budget,
    )?;
    drop(reordered_child_identifiers);

    budget.charge_transaction_work(decoded_extent.saturating_mul(3))?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    let compressed_len = compressed.len();
    drop(rewritten);
    let package_bound = source
        .source_bytes()
        .len()
        .saturating_add(compressed.len().saturating_mul(2))
        .saturating_add(1_024);
    budget.charge_transaction_work(package_bound.saturating_mul(3))?;
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let source_reopen =
        catalog_reopen_cost(source_catalog, source.source_bytes().len(), None, &[])?;
    let target_reopen = catalog_reopen_cost(
        source_catalog,
        output.len(),
        Some((component_name, compressed_len, decoded_extent)),
        previews,
    )?;
    budget.charge_transaction_work(target_reopen.work)?;
    budget.charge_references(target_reopen.references)?;
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    let candidate_target = resolve_native_target(&candidate, budget)?;
    if candidate_target.prepared() != target.prepared()
        || !identifiers_equal(
            candidate_target.document_snapshot.sheet_references(),
            reordered_sheet_identifiers_from(
                target.document_snapshot.sheet_references(),
                source_position,
                destination_position,
            ),
        )
        || !identifiers_equal(
            candidate_target.sidebar_snapshot.child_references(),
            reordered_sheet_identifiers_from(
                target.sidebar_snapshot.child_references(),
                source_position,
                destination_position,
            ),
        )
    {
        return Err(Error::Verification);
    }
    Ok(Rewritten {
        package: candidate,
        target: candidate_target,
        source_reopen,
        target_reopen,
    })
}

fn rewrite_owner(
    archive: &mut Archive,
    target: super::resolve::MessageTarget,
    data: Vec<u8>,
    reordered_identifiers: &[u64],
    limits: litchi_iwa_core::Limits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let object = archive
        .objects
        .get_mut(target.object_index)
        .filter(|object| object.archive_info.identifier == Some(target.identifier))
        .ok_or(Error::InvalidSource)?;
    let report = metadata_reorder_report(
        object,
        target.message_index,
        reordered_identifiers.len(),
        data.len(),
    )?;
    if report.references < reordered_identifiers.len() {
        return Err(Error::InvalidSource);
    }
    budget.charge_transaction_work(report.work)?;
    object
        .replace_message_reordering_object_references_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data,
            },
            reordered_identifiers,
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn metadata_reorder_report(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    selected_references: usize,
    data_bytes: usize,
) -> Result<ReorderReport, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let work = object
        .messages
        .len()
        .saturating_add(object.archive_info.message_infos.len())
        .saturating_add(info.object_references.len().saturating_mul(4))
        .saturating_add(selected_references)
        .saturating_mul(32)
        .saturating_add(data_bytes.saturating_mul(10));
    Ok(ReorderReport {
        work,
        references: info
            .object_references
            .len()
            .checked_add(selected_references)
            .ok_or(Error::InvalidSource)?,
    })
}

fn selected_payload(
    archive: &Archive,
    target: super::resolve::MessageTarget,
) -> Result<&[u8], Error> {
    archive
        .objects
        .get(target.object_index)
        .filter(|object| object.archive_info.identifier == Some(target.identifier))
        .and_then(|object| object.messages.get(target.message_index))
        .filter(|message| message.type_ == target.message_type)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)
}

fn reorder_raw_records(
    source: &[u8],
    field_number: u32,
    references: &[ReferenceSnapshot],
    source_position: usize,
    destination_position: usize,
) -> Result<RawReorder, Error> {
    let limits = wire_limits(source.len())?;
    let fields = parse_wire_fields_with_limits(source, limits).map_err(map_wire_error)?;
    let selected_count = fields
        .iter()
        .filter(|field| field.number() == field_number)
        .count();
    if selected_count != references.len()
        || source_position >= selected_count
        || destination_position >= selected_count
    {
        return Err(Error::InvalidSource);
    }
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(selected_count)
        .map_err(|_allocation| Error::Allocation {
            amount: selected_count,
        })?;
    for field in &fields {
        if field.number() == field_number {
            if field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            field
                .validate_canonical_framing(source)
                .map_err(map_wire_error)?;
            selected.push(*field);
        }
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_allocation| Error::Allocation {
            amount: source.len(),
        })?;
    let mut selected_position = 0usize;
    for field in &fields {
        if field.number() == field_number {
            let source_slot =
                moved_source_slot(selected_position, source_position, destination_position);
            output.extend_from_slice(
                selected
                    .get(source_slot)
                    .ok_or(Error::InvalidSource)?
                    .raw(source)
                    .map_err(map_wire_error)?,
            );
            selected_position = selected_position
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        } else {
            output.extend_from_slice(field.raw(source).map_err(map_wire_error)?);
        }
    }
    if output.len() != source.len() {
        return Err(Error::Verification);
    }
    let report = reorder_report(fields.len(), selected_count, source.len())?;
    Ok(RawReorder {
        data: output,
        report,
    })
}

fn reorder_report(
    fields: usize,
    selected_references: usize,
    source_bytes: usize,
) -> Result<ReorderReport, Error> {
    let work = source_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(fields.checked_mul(2)?))
        .and_then(|value| value.checked_add(selected_references))
        .ok_or(Error::InvalidSource)?;
    Ok(ReorderReport {
        work,
        references: selected_references,
    })
}

fn moved_source_slot(slot: usize, source: usize, destination: usize) -> usize {
    match source.cmp(&destination) {
        std::cmp::Ordering::Less if slot < source => slot,
        std::cmp::Ordering::Less if slot < destination => slot + 1,
        std::cmp::Ordering::Greater if slot < destination => slot,
        std::cmp::Ordering::Less | std::cmp::Ordering::Greater if slot == destination => source,
        std::cmp::Ordering::Greater if slot <= source => slot - 1,
        std::cmp::Ordering::Less | std::cmp::Ordering::Equal | std::cmp::Ordering::Greater => slot,
    }
}

fn reordered_identifiers(
    references: &[ReferenceSnapshot],
    source: usize,
    destination: usize,
) -> Result<Vec<u64>, Error> {
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    for slot in 0..references.len() {
        identifiers.push(
            references
                .get(moved_source_slot(slot, source, destination))
                .copied()
                .ok_or(Error::InvalidSource)?
                .identifier(),
        );
    }
    Ok(identifiers)
}

fn reordered_sheet_identifiers_from(
    references: &[ReferenceSnapshot],
    source: usize,
    destination: usize,
) -> impl Iterator<Item = u64> + '_ {
    (0..references.len())
        .map(move |slot| references[moved_source_slot(slot, source, destination)].identifier())
}

fn identifiers_equal(actual: &[ReferenceSnapshot], expected: impl Iterator<Item = u64>) -> bool {
    actual
        .iter()
        .copied()
        .map(ReferenceSnapshot::identifier)
        .eq(expected)
}

fn archive_extent(archive: &Archive) -> Result<usize, Error> {
    archive.objects.iter().try_fold(0usize, |maximum, object| {
        let end = object
            .header_offset
            .checked_add(object.header_length)
            .and_then(|offset| offset.checked_add(object.data_length))
            .and_then(|value| usize::try_from(value).ok())
            .ok_or(Error::InvalidSource)?;
        Ok(maximum.max(end))
    })
}

fn archive_allocation_cost(archive: &Archive) -> Result<usize, Error> {
    use std::mem::size_of;

    fn add(cost: &mut usize, amount: usize) -> Result<(), Error> {
        *cost = cost.checked_add(amount).ok_or(Error::InvalidSource)?;
        Ok(())
    }

    let mut cost = archive
        .objects
        .len()
        .checked_mul(size_of::<litchi_iwa_core::ArchiveObject>())
        .ok_or(Error::InvalidSource)?;
    for object in &archive.objects {
        add(
            &mut cost,
            object
                .messages
                .len()
                .checked_mul(size_of::<RawMessage>())
                .ok_or(Error::InvalidSource)?,
        )?;
        add(
            &mut cost,
            object
                .archive_info
                .message_infos
                .len()
                .checked_mul(size_of::<litchi_iwa_core::MessageInfo>())
                .ok_or(Error::InvalidSource)?,
        )?;
        add(
            &mut cost,
            usize::try_from(object.header_length)
                .map_err(|_conversion| Error::InvalidSource)?
                .checked_mul(2)
                .ok_or(Error::InvalidSource)?,
        )?;
        for info in &object.archive_info.message_infos {
            for length in [
                info.versions.len(),
                info.diff_merge_version.len(),
                info.diff_read_version.len(),
            ] {
                add(
                    &mut cost,
                    length
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )?;
            }
            if let Some(path) = &info.diff_field_path {
                add(&mut cost, size_of::<litchi_iwa_core::FieldPath>())?;
                add(
                    &mut cost,
                    path.path
                        .len()
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )?;
            }
            add(
                &mut cost,
                info.fields_to_remove
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::FieldPath>())
                    .ok_or(Error::InvalidSource)?,
            )?;
            add(
                &mut cost,
                info.object_references
                    .len()
                    .checked_mul(size_of::<u64>())
                    .ok_or(Error::InvalidSource)?,
            )?;
            add(
                &mut cost,
                info.data_references
                    .len()
                    .checked_mul(size_of::<u64>())
                    .ok_or(Error::InvalidSource)?,
            )?;
            add(
                &mut cost,
                info.field_infos
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::FieldInfo>())
                    .ok_or(Error::InvalidSource)?,
            )?;
            for path in &info.fields_to_remove {
                add(
                    &mut cost,
                    path.path
                        .len()
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )?;
            }
            for field in &info.field_infos {
                let field_items = field
                    .path
                    .path
                    .len()
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or(Error::InvalidSource)?;
                add(
                    &mut cost,
                    field_items
                        .checked_mul(size_of::<u64>())
                        .ok_or(Error::InvalidSource)?,
                )?;
                add(
                    &mut cost,
                    field
                        .known_field_version
                        .len()
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )?;
                add(
                    &mut cost,
                    field
                        .known_field_feature_identifier
                        .as_ref()
                        .map_or(0, String::len),
                )?;
            }
        }
    }
    Ok(cost)
}

fn archive_reference_count(archive: &Archive) -> Result<usize, Error> {
    archive.objects.iter().try_fold(0usize, |total, object| {
        object
            .archive_info
            .message_infos
            .iter()
            .try_fold(total, |subtotal, info| {
                let aggregate = info
                    .object_references
                    .len()
                    .checked_add(info.data_references.len())
                    .ok_or(Error::InvalidSource)?;
                info.field_infos.iter().try_fold(
                    subtotal
                        .checked_add(aggregate)
                        .ok_or(Error::InvalidSource)?,
                    |field_total, field| {
                        field_total
                            .checked_add(field.object_references.len())
                            .and_then(|value| value.checked_add(field.data_references.len()))
                            .ok_or(Error::InvalidSource)
                    },
                )
            })
    })
}

fn catalog_reopen_cost(
    catalog: &SourceCatalog,
    raw_package_bytes: usize,
    replacement: Option<(&str, usize, usize)>,
    deletions: &[&str],
) -> Result<ReopenCost, Error> {
    let mut logical_bytes = 0usize;
    for entry in catalog.package().iter() {
        if deletions.contains(&entry.name()) {
            continue;
        }
        let bytes = replacement
            .filter(|(name, _, _)| *name == entry.name())
            .map_or(entry.data().len(), |(_, logical, _)| logical);
        logical_bytes = logical_bytes
            .checked_add(bytes)
            .ok_or(Error::InvalidSource)?;
    }
    let mut decoded_bytes = 0usize;
    let mut structure = 0usize;
    let mut references = 0usize;
    for component in catalog.components().iter() {
        let decoded = replacement
            .filter(|(name, _, _)| *name == component.name())
            .map_or(archive_extent(component.archive())?, |(_, _, extent)| {
                extent
            });
        decoded_bytes = decoded_bytes
            .checked_add(decoded)
            .ok_or(Error::InvalidSource)?;
        structure = structure
            .checked_add(archive_allocation_cost(component.archive())?)
            .ok_or(Error::InvalidSource)?;
        references = references
            .checked_add(archive_reference_count(component.archive())?)
            .ok_or(Error::InvalidSource)?;
    }
    let work = raw_package_bytes
        .checked_add(logical_bytes.checked_mul(2).ok_or(Error::InvalidSource)?)
        .and_then(|value| value.checked_add(decoded_bytes.checked_mul(2)?))
        .and_then(|value| value.checked_add(structure))
        .ok_or(Error::InvalidSource)?;
    Ok(ReopenCost { work, references })
}

fn wire_limits(bytes: usize) -> Result<WireLimits, Error> {
    WireLimits::default()
        .with_input_bytes(bytes.max(1))
        .and_then(|limits| limits.with_output_bytes(bytes.max(1)))
        .and_then(|limits| limits.with_fields(bytes.max(1)))
        .and_then(|limits| limits.with_rewrite_work(bytes.saturating_mul(2).max(1)))
        .map_err(map_wire_error)
}

pub(super) fn root_preview_deletions(source: &SourceCatalog) -> Result<Vec<&'static str>, Error> {
    let mut counts = [0usize; ROOT_PREVIEWS.len()];
    for entry in source.package().iter() {
        if let Some(index) = ROOT_PREVIEWS.iter().position(|name| *name == entry.name()) {
            counts[index] = counts[index].checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_allocation| Error::Allocation {
            amount: ROOT_PREVIEWS.len(),
        })?;
    for (name, count) in ROOT_PREVIEWS.into_iter().zip(counts) {
        match count {
            0 => {},
            1 => previews.push(name),
            _ => return Err(Error::InvalidSource),
        }
    }
    Ok(previews)
}

pub(super) fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    package
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

pub(super) fn preflight_transaction_work(
    source: &Package,
    retained_target: Option<&[u8]>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    if let Some(target) = retained_target {
        let work = source
            .source_bytes()
            .len()
            .checked_add(target.len().checked_mul(4).ok_or(Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        budget.charge_transaction_work(work)?;
    }
    Ok(())
}

pub(super) fn verify_exact_locality(
    source: &Package,
    candidate: &Package,
    source_target: &NativeTarget,
    candidate_target: &NativeTarget,
    budget: &mut TransactionBudget,
    source_position: usize,
    destination_position: usize,
    source_previews: &[&str],
    expected_candidate_previews: usize,
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let candidate_previews = root_preview_deletions(candidate_catalog)?;
    if candidate_previews.len() != expected_candidate_previews {
        return Err(Error::Verification);
    }
    budget.charge_transaction_work(
        source
            .source_bytes()
            .len()
            .saturating_add(candidate.source_bytes().len()),
    )?;
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
                    .get_index(source_target.document.component_index)
                    .is_some_and(|component| component.name() == left.name());
                if if selected {
                    !selected_package_member_preserved(left, right)
                } else {
                    !package_member_preserved(left, right)
                } {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    verify_selected_archive(
        source,
        candidate,
        source_target,
        candidate_target,
        source_position,
        destination_position,
        budget,
    )
}

fn verify_selected_archive(
    source: &Package,
    candidate: &Package,
    source_target: &NativeTarget,
    candidate_target: &NativeTarget,
    source_position: usize,
    destination_position: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let before_component = source
        .state
        .components
        .catalog()
        .get_index(source_target.document.component_index)
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
    for (object_index, (left, right)) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
        .enumerate()
    {
        let owner = if object_index == source_target.document.object_index {
            Some((
                source_target.document,
                candidate_target.document,
                source_target.document_snapshot.sheet_references(),
                candidate_target.document_snapshot.sheet_references(),
                1,
            ))
        } else if object_index == source_target.sidebar_root.object_index {
            Some((
                source_target.sidebar_root,
                candidate_target.sidebar_root,
                source_target.sidebar_snapshot.child_references(),
                candidate_target.sidebar_snapshot.child_references(),
                2,
            ))
        } else {
            None
        };
        let Some((before_target, after_target, before_order, after_order, field_number)) = owner
        else {
            if left.archive_info != right.archive_info || left.messages != right.messages {
                return Err(Error::Verification);
            }
            continue;
        };
        if before_target != after_target
            || left.messages.len() != right.messages.len()
            || left.archive_info.identifier != right.archive_info.identifier
            || left.archive_info.should_merge != right.archive_info.should_merge
            || left.archive_info.message_infos.len() != right.archive_info.message_infos.len()
        {
            return Err(Error::Verification);
        }
        for (message_index, (left_message, right_message)) in
            left.messages.iter().zip(&right.messages).enumerate()
        {
            if message_index != before_target.message_index {
                if left_message != right_message
                    || left.archive_info.message_infos.get(message_index)
                        != right.archive_info.message_infos.get(message_index)
                {
                    return Err(Error::Verification);
                }
                continue;
            }
            let fields = if field_number == 1 {
                source_target
                    .document_fields
                    .checked_add(candidate_target.document_fields)
                    .ok_or(Error::InvalidSource)?
            } else {
                source_target
                    .sidebar_fields
                    .checked_add(candidate_target.sidebar_fields)
                    .ok_or(Error::InvalidSource)?
            };
            budget.charge_wire(
                fields,
                left_message
                    .data
                    .len()
                    .checked_add(right_message.data.len())
                    .ok_or(Error::InvalidSource)?,
            )?;
            let payload_matches = raw_records_reordered_exact(
                &left_message.data,
                &right_message.data,
                field_number,
                before_order.len(),
                source_position,
                destination_position,
            )?;
            let left_info = left
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification)?;
            let right_info = right
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification)?;
            if !payload_matches
                || !message_info_preserved_except_length_and_order(
                    left_info,
                    right_info,
                    before_order,
                    after_order,
                )
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn raw_records_reordered_exact(
    source: &[u8],
    candidate: &[u8],
    field_number: u32,
    selected_count: usize,
    source_position: usize,
    destination_position: usize,
) -> Result<bool, Error> {
    if source.len() != candidate.len()
        || source_position >= selected_count
        || destination_position >= selected_count
    {
        return Ok(false);
    }
    let source_fields = parse_wire_fields_with_limits(source, wire_limits(source.len())?)
        .map_err(map_wire_error)?;
    let candidate_fields = parse_wire_fields_with_limits(candidate, wire_limits(candidate.len())?)
        .map_err(map_wire_error)?;
    if source_fields.len() != candidate_fields.len() {
        return Ok(false);
    }
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(selected_count)
        .map_err(|_allocation| Error::Allocation {
            amount: selected_count,
        })?;
    for field in &source_fields {
        if field.number() == field_number {
            if field.wire_type() != 2 {
                return Ok(false);
            }
            field
                .validate_canonical_framing(source)
                .map_err(map_wire_error)?;
            selected.push(*field);
        }
    }
    if selected.len() != selected_count {
        return Ok(false);
    }
    let mut selected_position = 0usize;
    for (left, right) in source_fields.iter().zip(&candidate_fields) {
        if left.number() == field_number {
            if right.number() != field_number || right.wire_type() != 2 {
                return Ok(false);
            }
            right
                .validate_canonical_framing(candidate)
                .map_err(map_wire_error)?;
            let source_slot =
                moved_source_slot(selected_position, source_position, destination_position);
            if selected
                .get(source_slot)
                .ok_or(Error::InvalidSource)?
                .raw(source)
                .map_err(map_wire_error)?
                != right.raw(candidate).map_err(map_wire_error)?
            {
                return Ok(false);
            }
            selected_position = selected_position
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        } else if right.number() == field_number
            || left.raw(source).map_err(map_wire_error)?
                != right.raw(candidate).map_err(map_wire_error)?
        {
            return Ok(false);
        }
    }
    Ok(selected_position == selected_count)
}

fn message_info_preserved_except_length_and_order(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
    source_order: &[ReferenceSnapshot],
    target_order: &[ReferenceSnapshot],
) -> bool {
    source.type_ == candidate.type_
        && source.length == candidate.length
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && selected_subsequence_reordered(
            &source.object_references,
            &candidate.object_references,
            source_order,
            target_order,
        )
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn selected_subsequence_reordered(
    source: &[u64],
    candidate: &[u64],
    source_order: &[ReferenceSnapshot],
    target_order: &[ReferenceSnapshot],
) -> bool {
    if source.len() != candidate.len() || source_order.len() != target_order.len() {
        return false;
    }
    let mut selected = 0usize;
    for (left, right) in source.iter().zip(candidate) {
        if source_order
            .get(selected)
            .is_some_and(|reference| reference.identifier() == *left)
        {
            if target_order
                .get(selected)
                .is_none_or(|reference| reference.identifier() != *right)
            {
                return false;
            }
            selected += 1;
        } else if left != right {
            return false;
        }
    }
    selected == source_order.len()
}

fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn package_member_preserved(
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

fn selected_package_member_preserved(
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

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::{Limits, package::to_bytes};
    use litchi_iwa_common::wire::append_length_delimited_field;
    use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
    use litchi_iwa_protos::{numbers_sheet_order_codec, tn, tsk, tsp};
    use prost::Message as _;

    use super::super::Error;
    use super::{metadata_reorder_report, reorder_raw_records, reordered_identifiers};
    use crate::Package;

    const DOCUMENT: u64 = 1;
    const SIDEBAR: u64 = 20;
    const SHEETS: [u64; 3] = [2, 3, 4];
    const CHILDREN: [u64; 3] = [30, 31, 32];

    type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
        Ok(ArchiveObject::new(
            identifier,
            vec![RawMessage { type_, data }],
        )?)
    }

    fn package_bytes(
        field_attribution: bool,
        mismatched_child: bool,
        preview_count: usize,
    ) -> TestResult<Vec<u8>> {
        let mut document = object(
            DOCUMENT,
            1,
            tn::DocumentArchive {
                sheets: SHEETS.into_iter().map(reference).collect(),
                sidebar_order: reference(SIDEBAR),
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        document.archive_info.message_infos[0].object_references =
            [SHEETS.as_slice(), &[SIDEBAR]].concat();
        if field_attribution {
            let mut field = FieldInfo::new(vec![1]);
            field.object_references = SHEETS.to_vec();
            document.archive_info.message_infos[0]
                .field_infos
                .push(field);
        }
        let mut sidebar = object(
            SIDEBAR,
            205,
            tsk::TreeNode {
                children: CHILDREN.into_iter().map(reference).collect(),
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        sidebar.archive_info.message_infos[0].object_references = CHILDREN.to_vec();
        let mut objects = vec![document, sidebar];
        for (position, identifier) in CHILDREN.into_iter().enumerate() {
            let sheet = if mismatched_child && position == 1 {
                SHEETS[0]
            } else {
                SHEETS[position]
            };
            let mut child = object(
                identifier,
                205,
                tsk::TreeNode {
                    object: Some(reference(sheet)),
                    ..Default::default()
                }
                .encode_to_vec(),
            )?;
            child.archive_info.message_infos[0].object_references = vec![sheet];
            objects.push(child);
        }
        for (position, identifier) in SHEETS.into_iter().enumerate() {
            objects.push(object(
                identifier,
                2,
                tn::SheetArchive {
                    name: format!("Sheet {}", position + 1),
                    ..Default::default()
                }
                .encode_to_vec(),
            )?);
        }
        let document_member = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
        let previews = [
            ("preview.jpg", b"preview".as_slice()),
            ("preview-micro.jpg", b"micro".as_slice()),
            ("preview-web.jpg", b"web".as_slice()),
        ];
        let mut entries: Vec<(&str, &[u8])> = previews.into_iter().take(preview_count).collect();
        entries.push(("Index/Document.iwa", document_member.as_slice()));
        Ok(to_bytes(entries, Limits::default())?)
    }

    fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }

    #[test]
    fn changed_move_apply_inverse_and_reverse_direction_are_exact() -> TestResult {
        let source = package_bytes(false, false, 3)?;
        let package = Package::from_bytes(&source)?;
        let commit = package.edit_sheet_order().move_sheet(0usize, 2)?.commit()?;
        assert_eq!(
            commit
                .package()
                .document()
                .sheets()
                .iter()
                .map(crate::sheet::Sheet::name)
                .collect::<Vec<_>>(),
            ["Sheet 2", "Sheet 3", "Sheet 1"]
        );
        assert_eq!(commit.diagnostics().deleted_previews(), 3);
        let target = exact_bytes(commit.package())?;
        assert_eq!(
            exact_bytes(package.apply_sheet_order(commit.patch())?.package())?,
            target
        );
        let reopened = Package::from_bytes(&target)?;
        assert_eq!(
            exact_bytes(
                reopened
                    .apply_sheet_order(&commit.patch().inverse())?
                    .package()
            )?,
            source
        );
        assert!(matches!(
            reopened.apply_sheet_order(commit.patch()),
            Err(Error::PatchConflict)
        ));
        let reverse = package.edit_sheet_order().move_sheet(2usize, 0)?.commit()?;
        assert_eq!(
            reverse
                .package()
                .document()
                .sheets()
                .iter()
                .map(crate::sheet::Sheet::name)
                .collect::<Vec<_>>(),
            ["Sheet 3", "Sheet 1", "Sheet 2"]
        );
        Ok(())
    }

    #[test]
    fn attributed_order_fields_and_mismatched_sidebar_fail_closed() -> TestResult {
        let attributed = Package::from_bytes(&package_bytes(true, false, 3)?)?;
        assert!(matches!(
            attributed.edit_sheet_order().move_sheet(0usize, 2),
            Err(Error::UnsupportedSource)
        ));
        let malformed = Package::from_bytes(&package_bytes(false, true, 3)?)?;
        assert!(matches!(
            malformed.edit_sheet_order().move_sheet(0usize, 2),
            Err(Error::InvalidSource)
        ));
        Ok(())
    }

    #[test]
    fn changed_move_requires_all_three_root_previews() -> TestResult {
        let source = package_bytes(false, false, 2)?;
        let package = Package::from_bytes(&source)?;
        let result = package.edit_sheet_order().move_sheet(0usize, 2)?.commit();
        assert!(matches!(result, Err(Error::UnsupportedSource)));
        assert_eq!(exact_bytes(&package)?, source);
        Ok(())
    }

    fn scaling_case(count: usize) -> TestResult<(usize, usize, usize)> {
        let identifiers: Vec<u64> = (0..count)
            .map(|index| u64::try_from(index).map(|value| value + 1_000))
            .collect::<Result<_, _>>()?;
        let mut payload = Vec::new();
        for identifier in &identifiers {
            append_length_delimited_field(
                &mut payload,
                1,
                &reference(*identifier).encode_to_vec(),
            )?;
        }
        append_length_delimited_field(&mut payload, 5, &reference(SIDEBAR).encode_to_vec())?;
        let options = numbers_sheet_order_codec::DecodeOptions::new(
            payload.len(),
            payload.len().saturating_mul(2),
            payload.len().saturating_mul(4),
            2,
            count + 1,
        );
        let (snapshot, report) =
            numbers_sheet_order_codec::decode_document_sheet_order_with_report(&payload, options)?;
        let rewritten =
            reorder_raw_records(&payload, 1, snapshot.sheet_references(), 0, count - 1)?;
        let raw_report = rewritten.report;
        let rewritten_data = rewritten.data;
        let reordered = reordered_identifiers(snapshot.sheet_references(), 0, count - 1)?;
        let mut owner = object(DOCUMENT, 1, payload.clone())?;
        owner.archive_info.message_infos[0].object_references =
            [identifiers.as_slice(), &[SIDEBAR]].concat();
        let metadata_report =
            metadata_reorder_report(&owner, 0, reordered.len(), rewritten_data.len())?;
        owner.replace_message_reordering_object_references_preserving_header(
            0,
            RawMessage {
                type_: 1,
                data: rewritten_data.clone(),
            },
            &reordered,
        )?;
        assert_eq!(owner.messages[0].data, rewritten_data);
        assert_eq!(
            owner.archive_info.message_infos[0].object_references[..count],
            reordered
        );
        assert_eq!(
            owner.archive_info.message_infos[0].object_references[count],
            SIDEBAR
        );
        assert_eq!(owner.messages[0].data.len(), payload.len());
        Ok((
            report
                .work_bytes()
                .saturating_add(raw_report.work)
                .saturating_add(metadata_report.work),
            report
                .references()
                .saturating_add(raw_report.references)
                .saturating_add(metadata_report.references),
            payload.len(),
        ))
    }

    #[test]
    fn production_reorder_reports_and_raw_sizes_scale_linearly() -> TestResult {
        let small = scaling_case(4_096)?;
        let large = scaling_case(8_192)?;
        for (small_value, large_value) in
            [(small.0, large.0), (small.1, large.1), (small.2, large.2)]
        {
            assert!(large_value <= small_value.saturating_mul(23) / 10 + 32);
        }
        Ok(())
    }
}
