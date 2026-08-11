use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{
    WireLimits,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireDescent,
        patch_nested_fields_batched_with_limits, preflight_wire_tree_with_limits,
    },
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};

use super::{
    Error, Package, Path, ReopenCost, Settings, Target, TransactionBudget,
    decode_title_with_budget, map_header_error, map_wire_error,
    resolve_title_at_positions_with_budget,
};

const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_OUTLINED_FIELD: u32 = 37;

fn archive_extent(archive: &Archive) -> Result<usize, Error> {
    archive.objects.iter().try_fold(0usize, |maximum, object| {
        let end = object
            .header_offset
            .checked_add(object.header_length)
            .and_then(|offset| offset.checked_add(object.data_length))
            .and_then(|value| usize::try_from(value).ok())
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        Ok(maximum.max(end))
    })
}

pub(super) fn reassembly_cost(
    source_len: usize,
    compressed_len: usize,
) -> Result<(usize, usize), Error> {
    let package_bound = source_len
        .checked_add(compressed_len.saturating_mul(2))
        .and_then(|value| value.checked_add(1_024))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    Ok((package_bound, package_bound.saturating_mul(3)))
}

fn rewritten_payload(
    source: &[u8],
    before: Settings,
    after: Settings,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let decoded = decode_title_with_budget(source, budget)?;
    if decoded.settings != before {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let paths = [[TITLE_VISIBLE_FIELD], [TITLE_OUTLINED_FIELD]];
    let before_values = [
        before.visible().map(u64::from),
        before.outlined().map(u64::from),
    ];
    let after_values = [
        after.visible().map(u64::from),
        after.outlined().map(u64::from),
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
    ];
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(map_wire_error)?
        .with_fields(source.len().max(1))
        .map_err(map_wire_error)?
        .with_nesting(WireLimits::MAX_NESTING)
        .map_err(map_wire_error)?
        .with_output_bytes(source.len().saturating_add(32).max(1))
        .map_err(map_wire_error)?
        .with_rewrite_work(source.len().saturating_mul(4).max(1))
        .map_err(map_wire_error)?;
    budget.charge_transaction_work(source.len().saturating_mul(4))?;
    preflight_wire_tree_with_limits(source, limits, |_visit| Ok(WireDescent::Skip))
        .map_err(map_wire_error)?;
    let output =
        patch_nested_fields_batched_with_limits(source, &edits, limits).map_err(map_wire_error)?;
    let decoded = decode_title_with_budget(&output, budget)?;
    if decoded.settings != after {
        return Err(Error::Verification);
    }
    Ok(output)
}

pub(super) fn rewrite(
    source: &Package,
    target: Target,
    after: Settings,
    previews: &[&str],
    budget: &mut TransactionBudget,
    reopen_cost: ReopenCost,
) -> Result<Package, Error> {
    let source_catalog =
        super::super::table_headers::rewrite::physical_source(source).map_err(map_header_error)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.native.component_index)
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
    budget.charge_transaction_work(
        entry
            .data()
            .len()
            .saturating_mul(2)
            .saturating_add(archive_extent(component.archive())?.saturating_mul(3)),
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
    let object =
        archive
            .objects
            .get_mut(target.native.object_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    if object.archive_info.identifier != Some(target.native.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    super::super::table_headers::resolve::validate_message_metadata(
        object,
        target.native.message_index,
    )
    .map_err(map_header_error)?;
    let message = object
        .messages
        .get(target.native.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if message.type_ != target.native.message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let data = rewritten_payload(&message.data, target.settings, after, budget)?;
    object
        .replace_message_preserving_header_with_limits(
            target.native.message_index,
            RawMessage {
                type_: target.native.message_type,
                data,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    budget.charge_transaction_work(archive_extent(&archive)?.saturating_mul(3))?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    drop(archive);
    budget.charge_transaction_work(rewritten.len().saturating_mul(3))?;
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    drop(rewritten);
    let (package_bound, reassembly_work) =
        reassembly_cost(source.source_bytes().len(), compressed.len())?;
    budget.charge_transaction_work(reassembly_work)?;
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    if output.len() > package_bound {
        return Err(Error::Verification);
    }
    drop(compressed);
    budget.charge_references(reopen_cost.references)?;
    budget.charge_transaction_work(
        reopen_cost
            .work
            .saturating_add(output.len().saturating_mul(2))
            .saturating_add(128),
    )?;
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    let selected = resolve_title_at_positions_with_budget(
        &candidate,
        target.native.sheet_position,
        target.native.table_position,
        budget,
    )?;
    if selected.settings != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    use litchi_iwa_archive::LimitKind as ArchiveLimit;

    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                ArchiveLimit::InputBytes => super::LimitKind::InputBytes,
                ArchiveLimit::OutputBytes => super::LimitKind::OutputBytes,
                ArchiveLimit::Entries => super::LimitKind::Entries,
                ArchiveLimit::MemberNameBytes | ArchiveLimit::MetadataBytes => {
                    super::LimitKind::PackageBytes
                },
                ArchiveLimit::CompressedEntryBytes | ArchiveLimit::EntryBytes => {
                    super::LimitKind::EntryBytes
                },
                ArchiveLimit::TotalBytes => super::LimitKind::TotalEntryBytes,
                ArchiveLimit::IwaStreamBytes => super::LimitKind::PayloadBytes,
                ArchiveLimit::IwaTotalBytes => super::LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => super::LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    super::LimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => super::LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => super::LimitKind::PayloadItems,
                _ => super::LimitKind::PayloadBytes,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_candidate_read_error(error: super::super::Error) -> Error {
    match error {
        super::super::Error::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: super::LimitKind::OutputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Settings, TransactionBudget, rewritten_payload};
    use crate::package::table_title::LimitKind;

    fn payload(repetitions: usize) -> Vec<u8> {
        let mut source = Vec::with_capacity(repetitions.saturating_mul(6).saturating_add(3));
        for index in 0..repetitions {
            source.extend_from_slice(&[0xa5, 0x06]);
            source.extend_from_slice(&u32::try_from(index).unwrap_or(u32::MAX).to_le_bytes());
        }
        source.extend_from_slice(&[0xb0, 0x01, 0x01]);
        source
    }

    fn scaling_case(repetitions: usize) -> Result<(usize, usize, usize, usize), Error> {
        let source = payload(repetitions);
        let maximum = source.len().saturating_mul(32).max(1);
        let mut budget =
            TransactionBudget::test_with_limits(source.len().max(1), maximum, maximum, 2, maximum);
        let output = rewritten_payload(
            &source,
            Settings::new(Some(true), None),
            Settings::new(Some(false), None),
            &mut budget,
        )?;
        assert_eq!(output.len(), source.len());
        Ok(budget.test_usage())
    }

    #[test]
    fn production_title_rewrite_counters_scale_linearly_and_limit_before_output()
    -> Result<(), Error> {
        let small = scaling_case(682)?;
        let large = scaling_case(1_365)?;
        for (small_counter, large_counter) in
            [(small.0, large.0), (small.1, large.1), (small.3, large.3)]
        {
            assert!(large_counter.saturating_mul(10) <= small_counter.saturating_mul(23));
        }
        assert_eq!(small.2, 0);
        assert_eq!(large.2, 0);

        let source = payload(682);
        let fields = 683usize;
        let maximum = source.len().saturating_mul(32);
        let mut budget =
            TransactionBudget::test_with_limits(source.len(), fields - 1, maximum, 2, maximum);
        let attempted = rewritten_payload(
            &source,
            Settings::new(Some(true), None),
            Settings::new(Some(false), None),
            &mut budget,
        );
        let Err(error) = attempted else {
            return Err(Error::Verification);
        };
        assert!(matches!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::WireFields,
                ..
            }
        ));
        assert_eq!(budget.test_usage(), (0, 0, 0, 0));
        Ok(())
    }
}
