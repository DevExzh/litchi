//! Private scalar editing and exact verification of Keynote `SlideNode` field 18.

use super::{
    ArchiveLimits, ArchiveObject, BudgetedInvalidationError, InvalidationAllowance,
    InvalidationDirection, InvalidationError, InvalidationReport, RawMessage,
    SLIDE_NODE_MESSAGE_TYPE, WireError, WireLimitKind, WireLimits, charge_exact_archive_structure,
    charge_report_archive_structure, charge_report_message_selection, exact_varint,
    selected_message_index, validate_exact_canonical_field, validate_metadata_topology,
    walk_raw_fields,
};

const SLIDE_NUMBER_FIELD: u32 = 18;
const CANONICAL_SLIDE_NUMBER_KEY: &[u8] = &[0x90, 0x01];

#[derive(Clone, Copy)]
struct SlideNumberScan {
    value: bool,
    range: Option<(usize, usize)>,
}

pub(in crate::package) fn set_slide_number_with_report(
    object: &mut ArchiveObject,
    target: bool,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    allowance: InvalidationAllowance,
) -> Result<(bool, InvalidationReport), BudgetedInvalidationError> {
    let mut report = InvalidationReport::with_allowance(allowance);
    let result = set_slide_number_inner(object, target, archive_limits, wire_limits, &mut report);
    if let Some(error) = report.budget_error() {
        return Err(error);
    }
    Ok((result?, report))
}

fn set_slide_number_inner(
    object: &mut ArchiveObject,
    target: bool,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<bool, InvalidationError> {
    charge_report_archive_structure(object, false, wire_limits, report)?;
    object.validate_with_limits(archive_limits)?;
    charge_report_message_selection(object, wire_limits, report)?;
    let message_index = selected_message_index(object)?;
    report.charge_work(1, wire_limits)?;
    validate_metadata_topology(object, message_index)?;
    let source = &object
        .messages
        .get(message_index)
        .ok_or(InvalidationError::InvalidSource)?
        .data;
    let scan = scan_slide_number(source, wire_limits, report)?;
    if scan.value == target {
        return Ok(false);
    }
    let output_length = if scan.range.is_some() {
        source.len()
    } else {
        source
            .len()
            .checked_add(CANONICAL_SLIDE_NUMBER_KEY.len() + 1)
            .ok_or(InvalidationError::InvalidSource)?
    };
    if output_length > wire_limits.max_output_bytes() {
        return Err(WireError::LimitExceeded {
            kind: WireLimitKind::OutputBytes,
            observed: output_length,
            limit: wire_limits.max_output_bytes(),
        }
        .into());
    }
    report.charge_work(output_length, wire_limits)?;
    charge_report_archive_structure(object, true, wire_limits, report)?;

    let mut data = Vec::new();
    data.try_reserve_exact(output_length)
        .map_err(|_allocation| WireError::Allocation {
            resource: "Keynote slide-number payload",
            amount: output_length,
        })?;
    if let Some((start, end)) = scan.range {
        data.extend_from_slice(
            source
                .get(..start)
                .ok_or(InvalidationError::InvalidSource)?,
        );
        data.extend_from_slice(CANONICAL_SLIDE_NUMBER_KEY);
        data.push(u8::from(target));
        data.extend_from_slice(source.get(end..).ok_or(InvalidationError::InvalidSource)?);
    } else {
        data.extend_from_slice(source);
        data.extend_from_slice(CANONICAL_SLIDE_NUMBER_KEY);
        data.push(u8::from(target));
    }
    if data.len() != output_length {
        return Err(InvalidationError::InvalidSource);
    }
    object.replace_message_preserving_header_with_limits(
        message_index,
        RawMessage {
            type_: SLIDE_NODE_MESSAGE_TYPE,
            data,
        },
        archive_limits,
    )?;
    Ok(true)
}

pub(in crate::package) fn exact_slide_number_delta_with_allowance(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    source_value: bool,
    target_value: bool,
    direction: InvalidationDirection,
    wire_limits: WireLimits,
    allowance: InvalidationAllowance,
) -> Result<(bool, InvalidationReport), BudgetedInvalidationError> {
    let mut report = InvalidationReport::with_allowance(allowance);
    let result = match direction {
        InvalidationDirection::Forward => exact_slide_number_forward_delta(
            source,
            candidate,
            source_value,
            target_value,
            wire_limits,
            &mut report,
        ),
        InvalidationDirection::Inverse => exact_slide_number_forward_delta(
            candidate,
            source,
            target_value,
            source_value,
            wire_limits,
            &mut report,
        ),
    };
    if let Some(error) = report.budget_error() {
        return Err(error);
    }
    Ok((result?, report))
}

fn scan_slide_number(
    source: &[u8],
    limits: WireLimits,
    report: &mut InvalidationReport,
) -> Result<SlideNumberScan, InvalidationError> {
    let mut cursor = 0usize;
    let mut value = None;
    let mut range = None;
    walk_raw_fields(source, limits, report, |field, _scan_report| {
        let start = cursor;
        cursor = cursor
            .checked_add(field.raw.len())
            .ok_or(InvalidationError::InvalidSource)?;
        if field.number == SLIDE_NUMBER_FIELD {
            validate_exact_canonical_field(field)?;
            let decoded = exact_varint(field)?;
            if decoded > 1 || value.replace(decoded == 1).is_some() {
                return Err(InvalidationError::InvalidSource);
            }
            range = Some((start, cursor));
        }
        Ok(())
    })?;
    if cursor != source.len() {
        return Err(InvalidationError::InvalidSource);
    }
    Ok(SlideNumberScan {
        value: value.unwrap_or(false),
        range,
    })
}

fn exact_slide_number_forward_delta(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    source_value: bool,
    target_value: bool,
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
    let source_scan = scan_slide_number(source_payload, limits, report)?;
    let candidate_scan = scan_slide_number(candidate_payload, limits, report)?;
    if source_scan.value != source_value || candidate_scan.value != target_value {
        return Ok(false);
    }
    charge_exact_archive_structure(source, limits, report)?;
    charge_exact_archive_structure(candidate, limits, report)?;
    if source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len() {
        return Ok(false);
    }
    for (index, (before, after)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if before.type_ != after.type_
            || (index == source_index
                && !slide_number_payload_delta_matches(
                    &before.data,
                    &after.data,
                    source_scan,
                    candidate_scan,
                    source_value,
                    target_value,
                ))
            || (index != source_index && before != after)
        {
            return Ok(false);
        }
    }
    for (index, (before, after)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        let Some(source_message) = source.messages.get(index) else {
            return Ok(false);
        };
        let Some(candidate_message) = candidate.messages.get(index) else {
            return Ok(false);
        };
        if index == source_index {
            if !slide_number_message_info_matches(
                before,
                after,
                source_message.data.len(),
                candidate_message.data.len(),
            ) {
                return Ok(false);
            }
        } else if before != after
            || usize::try_from(before.length).ok() != Some(source_message.data.len())
            || usize::try_from(after.length).ok() != Some(candidate_message.data.len())
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn slide_number_payload_delta_matches(
    source: &[u8],
    candidate: &[u8],
    source_scan: SlideNumberScan,
    candidate_scan: SlideNumberScan,
    source_value: bool,
    target_value: bool,
) -> bool {
    if source_value == target_value {
        return source == candidate;
    }
    let canonical = [
        CANONICAL_SLIDE_NUMBER_KEY[0],
        CANONICAL_SLIDE_NUMBER_KEY[1],
        u8::from(target_value),
    ];
    if let Some((source_start, source_end)) = source_scan.range {
        let Some((candidate_start, candidate_end)) = candidate_scan.range else {
            return false;
        };
        return source_start == candidate_start
            && source_end == candidate_end
            && candidate.get(candidate_start..candidate_end) == Some(canonical.as_slice())
            && source.get(..source_start) == candidate.get(..candidate_start)
            && source.get(source_end..) == candidate.get(candidate_end..);
    }
    let Some((candidate_start, candidate_end)) = candidate_scan.range else {
        return false;
    };
    !source_value
        && target_value
        && candidate_start == source.len()
        && candidate_end == candidate.len()
        && candidate.get(candidate_start..candidate_end) == Some(canonical.as_slice())
        && candidate.get(..candidate_start) == Some(source)
}

fn slide_number_message_info_matches(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
    source_payload_length: usize,
    candidate_payload_length: usize,
) -> bool {
    let Ok(source_length) = u32::try_from(source_payload_length) else {
        return false;
    };
    let Ok(candidate_length) = u32::try_from(candidate_payload_length) else {
        return false;
    };
    source.type_ == candidate.type_
        && source.length == source_length
        && candidate.length == candidate_length
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::slide_preview::{
        InvalidationBudgetKind,
        tests::{preview_object, replace_once, unbudgeted_test_error},
    };

    #[test]
    fn scalar_splice_preserves_all_preview_bytes_and_metadata() -> Result<(), InvalidationError> {
        let mut absent = preview_object(false)?;
        let absent_source = absent.clone();
        let (changed, report) = set_slide_number_with_report(
            &mut absent,
            false,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;
        assert!(!changed);
        assert!(report.work() > 0);
        assert_eq!(absent, absent_source);

        let mut enabled = absent_source.clone();
        let source_payload = enabled.messages[0].data.clone();
        let source_info = enabled.archive_info.message_infos[0].clone();
        let (changed, _) = set_slide_number_with_report(
            &mut enabled,
            true,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;
        assert!(changed);
        assert_eq!(
            enabled.messages[0].data.get(..source_payload.len()),
            Some(source_payload.as_slice())
        );
        assert_eq!(
            enabled.messages[0].data.get(source_payload.len()..),
            Some([0x90, 0x01, 0x01].as_slice())
        );
        let mut expected_info = source_info;
        expected_info.length = u32::try_from(enabled.messages[0].data.len())
            .map_err(|_conversion| InvalidationError::InvalidSource)?;
        assert_eq!(enabled.archive_info.message_infos[0], expected_info);
        assert!(
            exact_slide_number_delta_with_allowance(
                &absent_source,
                &enabled,
                false,
                true,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        assert!(
            exact_slide_number_delta_with_allowance(
                &absent_source,
                &absent_source,
                false,
                false,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );

        let mut explicit_false = absent_source.clone();
        explicit_false.messages[0]
            .data
            .extend_from_slice(&[0x90, 0x01, 0x00]);
        explicit_false.archive_info.message_infos[0].length =
            u32::try_from(explicit_false.messages[0].data.len())
                .map_err(|_conversion| InvalidationError::InvalidSource)?;
        let explicit_false_source = explicit_false.clone();
        assert!(
            !set_slide_number_with_report(
                &mut explicit_false,
                false,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        assert_eq!(explicit_false, explicit_false_source);
        assert!(
            set_slide_number_with_report(
                &mut explicit_false,
                true,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        assert_eq!(
            explicit_false.messages[0]
                .data
                .get(explicit_false.messages[0].data.len().saturating_sub(3)..),
            Some([0x90, 0x01, 0x01].as_slice())
        );
        assert!(
            set_slide_number_with_report(
                &mut explicit_false,
                false,
                ArchiveLimits::default(),
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        assert_eq!(
            explicit_false.messages[0]
                .data
                .get(explicit_false.messages[0].data.len().saturating_sub(3)..),
            Some([0x90, 0x01, 0x00].as_slice())
        );
        Ok(())
    }

    #[test]
    fn exact_delta_proves_forward_inverse_and_rejects_cache_changes()
    -> Result<(), InvalidationError> {
        let mut source = preview_object(false)?;
        source.messages[0]
            .data
            .extend_from_slice(&[0x90, 0x01, 0x00]);
        source.archive_info.message_infos[0].length = u32::try_from(source.messages[0].data.len())
            .map_err(|_conversion| InvalidationError::InvalidSource)?;
        let mut candidate = source.clone();
        set_slide_number_with_report(
            &mut candidate,
            true,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;

        assert!(
            exact_slide_number_delta_with_allowance(
                &source,
                &candidate,
                false,
                true,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        assert!(
            exact_slide_number_delta_with_allowance(
                &source,
                &candidate,
                false,
                true,
                InvalidationDirection::Inverse,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );

        let mut changed_cache = candidate.clone();
        replace_once(
            &mut changed_cache.messages[0].data,
            &[0x52, 0x00],
            &[0x52, 0x01, 0x00],
        )?;
        changed_cache.archive_info.message_infos[0].length =
            u32::try_from(changed_cache.messages[0].data.len())
                .map_err(|_conversion| InvalidationError::InvalidSource)?;
        assert!(
            !exact_slide_number_delta_with_allowance(
                &source,
                &changed_cache,
                false,
                true,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );

        let mut changed_metadata = candidate.clone();
        changed_metadata.archive_info.message_infos[0]
            .versions
            .push(99);
        assert!(
            !exact_slide_number_delta_with_allowance(
                &source,
                &changed_metadata,
                false,
                true,
                InvalidationDirection::Forward,
                WireLimits::default(),
                InvalidationAllowance::UNLIMITED,
            )
            .map_err(unbudgeted_test_error)?
            .0
        );
        Ok(())
    }

    #[test]
    fn scalar_rejects_noncanonical_or_ambiguous_f18_atomically() -> Result<(), InvalidationError> {
        for malformed in [
            &[0x90, 0x01, 0x00, 0x90, 0x01, 0x01][..],
            &[0x90, 0x81, 0x00, 0x00][..],
            &[0x90, 0x01, 0x80, 0x00][..],
            &[0x90, 0x01, 0x02][..],
            &[0x92, 0x01, 0x00][..],
        ] {
            let mut object = preview_object(false)?;
            object.messages[0].data.extend_from_slice(malformed);
            object.archive_info.message_infos[0].length =
                u32::try_from(object.messages[0].data.len())
                    .map_err(|_conversion| InvalidationError::InvalidSource)?;
            let original = object.clone();
            assert!(
                set_slide_number_with_report(
                    &mut object,
                    true,
                    ArchiveLimits::default(),
                    WireLimits::default(),
                    InvalidationAllowance::UNLIMITED,
                )
                .is_err()
            );
            assert_eq!(object, original);
        }
        Ok(())
    }

    #[test]
    fn scalar_honors_shared_allowance_before_mutation_or_comparison()
    -> Result<(), InvalidationError> {
        let source = preview_object(false)?;
        let mut mutation = source.clone();
        assert!(matches!(
            set_slide_number_with_report(
                &mut mutation,
                true,
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
        assert_eq!(mutation, source);

        let mut candidate = source.clone();
        set_slide_number_with_report(
            &mut candidate,
            true,
            ArchiveLimits::default(),
            WireLimits::default(),
            InvalidationAllowance::UNLIMITED,
        )
        .map_err(unbudgeted_test_error)?;
        assert!(matches!(
            exact_slide_number_delta_with_allowance(
                &source,
                &candidate,
                false,
                true,
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
        Ok(())
    }
}
