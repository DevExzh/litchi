//! Exact candidate, locality, and directional cache verification.

use litchi_iwa_archive::package::Entry;
use litchi_iwa_core::ArchiveObject;

use super::{
    Error, Package, SLIDE_MESSAGE_TYPE, Selection, SlideSelector, State, TransactionBudget,
    map_rendering_invalidation_error, map_wire_error, physical_catalog,
    rewrite::visibility_payload_delta_matches, root_preview_count, select_with_budget,
    selected_message,
};

pub(super) fn verify_artifact_delta(
    source: &Package,
    candidate: &Package,
    selection: &Selection<'_>,
    expected: State,
    preview_count: usize,
    source_node_invalidated: bool,
    target_node_invalidated: bool,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    let target = select_with_budget(
        candidate,
        SlideSelector::Position(selection.position),
        selection.kind,
        true,
        budget,
    )?
    .ok_or(Error::Verification)?;
    if target.node_identifier != selection.node_identifier
        || target.slide_identifier != selection.slide_identifier
        || target.placeholder_identifier != selection.placeholder_identifier
        || target.state != expected
    {
        return Err(Error::Verification);
    }
    verify_cache_direction(candidate, preview_count)?;
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let source_previews =
        crate::package::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    let candidate_previews =
        crate::package::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    let mut before = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_previews.names().contains(&entry.name()));
    let mut after = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_previews.names().contains(&entry.name()));
    let selected_names = [selection.slide_component, selection.node_component];
    let mut touched = 0usize;
    loop {
        match (before.next(), after.next()) {
            (Some(old), Some(new)) if old.name() == new.name() => {
                for amount in [
                    old.raw_record().local_record().len(),
                    new.raw_record().local_record().len(),
                    old.raw_record().central_directory_record().len(),
                    new.raw_record().central_directory_record().len(),
                ] {
                    budget.charge_work(amount)?;
                }
                if selected_names.contains(&old.name()) {
                    if !selected_package_member_preserved(old, new) {
                        return Err(Error::Verification);
                    }
                    if old.raw_record().compressed_data() != new.raw_record().compressed_data() {
                        touched = touched.checked_add(1).ok_or(Error::Verification)?;
                    }
                } else if !package_member_preserved(old, new) {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    let mut verified = Vec::new();
    for name in selected_names {
        if verified.contains(&name) {
            continue;
        }
        verified.push(name);
        verify_selected_component(
            source,
            candidate,
            selection,
            &target,
            name,
            source_node_invalidated,
            target_node_invalidated,
            budget,
        )?;
    }
    if touched == 0 {
        return Err(Error::Verification);
    }
    Ok(touched)
}

fn verify_selected_component(
    source: &Package,
    candidate: &Package,
    selection: &Selection<'_>,
    target: &Selection<'_>,
    name: &str,
    source_node_invalidated: bool,
    target_node_invalidated: bool,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let before = physical_catalog(source)?
        .components()
        .get(name)
        .ok_or(Error::Verification)?;
    let after = physical_catalog(candidate)?
        .components()
        .get(name)
        .ok_or(Error::Verification)?;
    if before.archive().objects.len() != after.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (old, new) in before
        .archive()
        .objects
        .iter()
        .zip(&after.archive().objects)
    {
        if old.archive_info.identifier != new.archive_info.identifier {
            return Err(Error::Verification);
        }
        match old.archive_info.identifier {
            Some(identifier) if identifier == selection.slide_identifier => {
                verify_slide_object(source, old, new, selection, target, budget)?;
            },
            Some(identifier) if identifier == selection.node_identifier => {
                verify_node_object(
                    source,
                    old,
                    new,
                    selection.kind,
                    source_node_invalidated,
                    target_node_invalidated,
                    budget,
                )?;
            },
            _ if old.archive_info != new.archive_info || old.messages != new.messages => {
                return Err(Error::Verification);
            },
            _ => {},
        }
    }
    Ok(())
}

fn verify_slide_object(
    package: &Package,
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    before: &Selection<'_>,
    after: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    verify_selected_object(source, candidate, SLIDE_MESSAGE_TYPE, |old, new| {
        let forward = visibility_payload_delta_matches(
            old,
            new,
            before.placeholder_identifier,
            before.kind,
            before.state,
            after.state,
            limits,
            budget,
        )?;
        if forward {
            return Ok(true);
        }
        visibility_payload_delta_matches(
            new,
            old,
            before.placeholder_identifier,
            before.kind,
            after.state,
            before.state,
            limits,
            budget,
        )
    })
}

fn verify_node_object(
    package: &Package,
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    kind: super::Kind,
    source_invalidated: bool,
    target_invalidated: bool,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let direction = match (source_invalidated, target_invalidated) {
        (false | true, true) => crate::package::slide_preview::InvalidationDirection::Forward,
        (true, false) => crate::package::slide_preview::InvalidationDirection::Inverse,
        (false, false) => return Err(Error::Verification),
    };
    let (matches, report) = if kind == super::Kind::SlideNumber {
        crate::package::slide_preview::exact_slide_number_delta_with_allowance(
            source,
            candidate,
            source_invalidated,
            target_invalidated,
            direction,
            limits,
            budget.preview_allowance(),
        )
    } else {
        crate::package::slide_preview::exact_invalidation_delta_with_allowance(
            source,
            candidate,
            direction,
            limits,
            budget.preview_allowance(),
        )
    }
    .map_err(|error| budget.map_preview_budget_error(error))?;
    budget.charge_preview_report(report)?;
    if !matches {
        return Err(Error::Verification);
    }
    Ok(())
}

fn verify_selected_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    kind: u32,
    payload_matches: impl FnOnce(&[u8], &[u8]) -> Result<bool, Error>,
) -> Result<(), Error> {
    let (source_index, source_payload) = selected_message(source, kind)?;
    let (candidate_index, candidate_payload) = selected_message(candidate, kind)?;
    if source_index != candidate_index
        || source.messages.len() != candidate.messages.len()
        || !payload_matches(source_payload, candidate_payload)?
    {
        return Err(Error::Verification);
    }
    for (index, (old, new)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if old.type_ != new.type_ || (index != source_index && old != new) {
            return Err(Error::Verification);
        }
    }
    let mut expected = source.archive_info.clone();
    expected
        .message_infos
        .get_mut(source_index)
        .ok_or(Error::Verification)?
        .length = u32::try_from(candidate_payload.len()).map_err(|_error| Error::Verification)?;
    if expected != candidate.archive_info {
        return Err(Error::Verification);
    }
    Ok(())
}

fn package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && crate::package::rendering_invalidation::central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
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

fn selected_local_record_preserved(source: &Entry, candidate: &Entry) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let old = source.raw_record().local_record();
    let new = candidate.raw_record().local_record();
    let Some(old_header) = zip_local_header_length(old) else {
        return false;
    };
    let Some(new_header) = zip_local_header_length(new) else {
        return false;
    };
    if old_header != new_header
        || old[..CRC_AND_SIZES.start] != new[..CRC_AND_SIZES.start]
        || old[CRC_AND_SIZES.end..old_header] != new[CRC_AND_SIZES.end..new_header]
    {
        return false;
    }
    let Some(old_payload_end) = old_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= old.len())
    else {
        return false;
    };
    let Some(new_payload_end) = new_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= new.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &old[old_payload_end..],
        &new[new_payload_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize.checked_add(name)?.checked_add(extra)
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let source_descriptor = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let candidate_descriptor = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    source_descriptor == candidate_descriptor
        && source.len() == candidate.len()
        && source.len() >= source_descriptor + 12
        && source[..source_descriptor] == candidate[..candidate_descriptor]
        && source[source_descriptor + 12..] == candidate[candidate_descriptor + 12..]
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

fn verify_cache_direction(package: &Package, preview_count: usize) -> Result<(), Error> {
    if root_preview_count(package)? != preview_count {
        return Err(Error::Verification);
    }
    Ok(())
}
