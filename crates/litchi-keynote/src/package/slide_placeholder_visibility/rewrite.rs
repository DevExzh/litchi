//! One-pass visibility splice and grouped physical publication.

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{WireLimits, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};

use super::{
    BODY_REFERENCE_FIELD, Error, LimitKind, OWNED_DRAWABLES_FIELD, Package, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE, Selection, State, TITLE_REFERENCE_FIELD, TransactionBudget,
    Z_ORDER_FIELD, canonical_varint_len, map_archive_error, map_core_error, map_read_error,
    map_rendering_invalidation_error, map_wire_error, physical_catalog, selected_message,
    strict_reference, verify_artifact_delta,
};

pub(super) fn rewrite_visibility(
    source: &Package,
    selection: &Selection<'_>,
    target: State,
    budget: &mut TransactionBudget,
) -> Result<(Package, usize, bool), Error> {
    let catalog = physical_catalog(source)?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let wire_limits = source.wire_limits().map_err(map_wire_error)?;
    let preview_plan =
        crate::package::rendering_invalidation::root_preview_deletions(catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(2)
        .map_err(|_allocation| Error::Allocation { amount: 2 })?;
    for name in [selection.slide_component, selection.node_component] {
        if !names.contains(&name) {
            names.push(name);
        }
    }
    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(names.len())
        .map_err(|_allocation| Error::Allocation {
            amount: names.len(),
        })?;
    let mut slide_changed = false;
    let mut node_seen = false;
    let mut source_node_invalidated = false;
    for name in &names {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == *name)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque() {
            return Err(Error::InvalidSource);
        }
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            physical_limits.snappy_limits().map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(map_core_error)?;
        drop(stream);
        if let Some(object) = archive.object(selection.slide_identifier) {
            let (index, payload) = selected_message(object, SLIDE_MESSAGE_TYPE)?;
            if payload != selection.slide_payload || std::mem::replace(&mut slide_changed, true) {
                return Err(Error::InvalidSource);
            }
            let changed = rewrite_visibility_payload(
                payload,
                selection.placeholder_identifier,
                selection.kind,
                selection.state,
                target,
                wire_limits,
                budget,
            )?;
            archive
                .object_mut(selection.slide_identifier)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header_with_limits(
                    index,
                    RawMessage {
                        type_: SLIDE_MESSAGE_TYPE,
                        data: changed,
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
        if let Some(node) = archive.object_mut(selection.node_identifier) {
            if std::mem::replace(&mut node_seen, true) {
                return Err(Error::InvalidSource);
            }
            let source_matches = {
                let (_index, payload) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
                payload == selection.node_payload
            };
            if !source_matches {
                return Err(Error::InvalidSource);
            }
            let (node_changed, report) = if selection.kind == super::Kind::SlideNumber {
                crate::package::slide_preview::set_slide_number_with_report(
                    node,
                    target == State::Visible,
                    archive_limits,
                    wire_limits,
                    budget.preview_allowance(),
                )
            } else {
                crate::package::slide_preview::invalidate_if_needed_with_report(
                    node,
                    archive_limits,
                    wire_limits,
                    budget.preview_allowance(),
                )
            }
            .map_err(|error| budget.map_preview_budget_error(error))?;
            budget.charge_preview_report(report)?;
            source_node_invalidated = if selection.kind == super::Kind::SlideNumber {
                selection.state == State::Visible
            } else {
                !node_changed
            };
        }
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        drop(archive);
        let replacement = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        drop(bytes);
        compressed.push((*name, replacement));
    }
    if !slide_changed || !node_seen {
        return Err(Error::InvalidSource);
    }
    let edits: Vec<_> = compressed
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect();
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, preview_plan.names(), physical_limits)
        .map_err(map_archive_error)?;
    drop(edits);
    drop(compressed);
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    let touched = verify_artifact_delta(
        source,
        &candidate,
        selection,
        target,
        0,
        source_node_invalidated,
        if selection.kind == super::Kind::SlideNumber {
            target == State::Visible
        } else {
            true
        },
        budget,
    )?;
    Ok((candidate, touched, source_node_invalidated))
}

pub(super) fn rewrite_visibility_payload(
    source: &[u8],
    identifier: u64,
    kind: super::Kind,
    before: State,
    after: State,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    if before == after {
        return Err(Error::Verification);
    }
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let role_number = match kind {
        super::Kind::Title => TITLE_REFERENCE_FIELD,
        super::Kind::Body => BODY_REFERENCE_FIELD,
        super::Kind::SlideNumber => super::SLIDE_NUMBER_PLACEHOLDER_FIELD,
    };
    let role = view
        .fields()
        .find(|field| field.number() == role_number)
        .ok_or(Error::InvalidSource)?;
    let mut retained = 0usize;
    let mut owned = 0usize;
    let mut z_order = 0usize;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let selected_list = matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD)
            && strict_reference(field, limits)? == identifier;
        if selected_list {
            if field.number() == OWNED_DRAWABLES_FIELD {
                owned = owned.checked_add(1).ok_or(Error::InvalidSource)?;
            } else {
                z_order = z_order.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
        if after == State::Visible || !selected_list {
            retained = retained
                .checked_add(field.raw().len())
                .ok_or(Error::InvalidSource)?;
        }
    }
    let expected = match before {
        State::Visible => 1,
        State::Hidden => 0,
    };
    if owned != expected || z_order != expected {
        return Err(Error::InvalidSource);
    }
    let append = if after == State::Visible {
        canonical_reference_record_len(OWNED_DRAWABLES_FIELD, role.payload().len())?
            .checked_add(canonical_reference_record_len(
                Z_ORDER_FIELD,
                role.payload().len(),
            )?)
            .ok_or(Error::InvalidSource)?
    } else {
        0
    };
    let output_len = retained.checked_add(append).ok_or(Error::InvalidSource)?;
    if output_len > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: output_len as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    let work = source
        .len()
        .checked_mul(6)
        .and_then(|value| value.checked_add(output_len))
        .and_then(|value| value.checked_add(view.len().saturating_mul(2)))
        .ok_or(Error::InvalidSource)?;
    if work > limits.max_rewrite_work() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: work as u64,
            maximum: limits.max_rewrite_work() as u64,
        });
    }
    // Charge the complete scan-and-emit plan before reserving output so a
    // transaction's remaining allowance cannot be exceeded post-allocation.
    budget.charge_work(work)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_allocation| Error::Allocation { amount: output_len })?;
    let mut remaining_owned_fields = view
        .fields()
        .filter(|field| field.number() == OWNED_DRAWABLES_FIELD)
        .count();
    let mut remaining_z_order_fields = view
        .fields()
        .filter(|field| field.number() == Z_ORDER_FIELD)
        .count();
    if kind == super::Kind::SlideNumber
        && after == State::Visible
        && (remaining_owned_fields == 0 || remaining_z_order_fields == 0)
    {
        return Err(Error::UnsupportedSource);
    }
    for field in view.fields() {
        let selected_list = matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD)
            && strict_reference(field, limits)? == identifier;
        if after == State::Visible || !selected_list {
            output.extend_from_slice(field.raw());
        }
        if kind == super::Kind::SlideNumber && after == State::Visible {
            if field.number() == OWNED_DRAWABLES_FIELD {
                remaining_owned_fields = remaining_owned_fields.saturating_sub(1);
                if remaining_owned_fields == 0 {
                    append_reference_record(&mut output, OWNED_DRAWABLES_FIELD, role.payload());
                }
            } else if field.number() == Z_ORDER_FIELD {
                remaining_z_order_fields = remaining_z_order_fields.saturating_sub(1);
                if remaining_z_order_fields == 0 {
                    append_reference_record(&mut output, Z_ORDER_FIELD, role.payload());
                }
            }
        }
    }
    if after == State::Visible && kind != super::Kind::SlideNumber {
        append_reference_record(&mut output, OWNED_DRAWABLES_FIELD, role.payload());
        append_reference_record(&mut output, Z_ORDER_FIELD, role.payload());
    }
    if output.len() != output_len {
        return Err(Error::Verification);
    }
    Ok(output)
}

/// Compare one directional visibility splice without allocating its output.
pub(super) fn visibility_payload_delta_matches(
    source: &[u8],
    target: &[u8],
    identifier: u64,
    kind: super::Kind,
    before: State,
    after: State,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    if before == after {
        return Ok(false);
    }
    let work = source
        .len()
        .checked_mul(5)
        .and_then(|amount| {
            target
                .len()
                .checked_mul(2)
                .and_then(|target_work| amount.checked_add(target_work))
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(work)?;
    let source_view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let target_view = WireView::parse_with_limits(target, limits).map_err(map_wire_error)?;
    let role_number = match kind {
        super::Kind::Title => TITLE_REFERENCE_FIELD,
        super::Kind::Body => BODY_REFERENCE_FIELD,
        super::Kind::SlideNumber => super::SLIDE_NUMBER_PLACEHOLDER_FIELD,
    };
    let role = source_view
        .fields()
        .find(|field| field.number() == role_number)
        .ok_or(Error::InvalidSource)?;
    let mut owned = 0usize;
    let mut z_order = 0usize;
    for field in source_view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD)
            && strict_reference(field, limits)? == identifier
        {
            if field.number() == OWNED_DRAWABLES_FIELD {
                owned = owned.checked_add(1).ok_or(Error::InvalidSource)?;
            } else {
                z_order = z_order.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
    }
    let expected = usize::from(before == State::Visible);
    if owned != expected || z_order != expected {
        return Err(Error::InvalidSource);
    }

    let mut target_fields = target_view.fields();
    let mut remaining_owned_fields = source_view
        .fields()
        .filter(|field| field.number() == OWNED_DRAWABLES_FIELD)
        .count();
    let mut remaining_z_order_fields = source_view
        .fields()
        .filter(|field| field.number() == Z_ORDER_FIELD)
        .count();
    if kind == super::Kind::SlideNumber
        && after == State::Visible
        && (remaining_owned_fields == 0 || remaining_z_order_fields == 0)
    {
        return Ok(false);
    }
    for field in source_view.fields() {
        let selected_list = matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD)
            && strict_reference(field, limits)? == identifier;
        if after == State::Visible || !selected_list {
            let Some(candidate) = target_fields.next() else {
                return Ok(false);
            };
            candidate
                .validate_canonical_framing()
                .map_err(map_wire_error)?;
            if candidate.raw() != field.raw() {
                return Ok(false);
            }
        }
        if kind == super::Kind::SlideNumber && after == State::Visible {
            let append_number = if field.number() == OWNED_DRAWABLES_FIELD {
                remaining_owned_fields = remaining_owned_fields.saturating_sub(1);
                (remaining_owned_fields == 0).then_some(OWNED_DRAWABLES_FIELD)
            } else if field.number() == Z_ORDER_FIELD {
                remaining_z_order_fields = remaining_z_order_fields.saturating_sub(1);
                (remaining_z_order_fields == 0).then_some(Z_ORDER_FIELD)
            } else {
                None
            };
            if let Some(number) = append_number {
                let Some(candidate) = target_fields.next() else {
                    return Ok(false);
                };
                candidate
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                if candidate.number() != number
                    || candidate.wire_type() != 2
                    || candidate.payload() != role.payload()
                {
                    return Ok(false);
                }
            }
        }
    }
    if after == State::Visible && kind != super::Kind::SlideNumber {
        for number in [OWNED_DRAWABLES_FIELD, Z_ORDER_FIELD] {
            let Some(candidate) = target_fields.next() else {
                return Ok(false);
            };
            candidate
                .validate_canonical_framing()
                .map_err(map_wire_error)?;
            if candidate.number() != number
                || candidate.wire_type() != 2
                || candidate.payload() != role.payload()
            {
                return Ok(false);
            }
        }
    }
    Ok(target_fields.next().is_none())
}

fn canonical_reference_record_len(field: u32, payload_len: usize) -> Result<usize, Error> {
    let key = (u64::from(field) << 3) | 2;
    canonical_varint_len(key)
        .checked_add(canonical_varint_len(payload_len as u64))
        .and_then(|value| value.checked_add(payload_len))
        .ok_or(Error::InvalidSource)
}

fn append_reference_record(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    encode_varint_into(output, (u64::from(field) << 3) | 2);
    encode_varint_into(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn payload_only_allowance_rejects_field_traversal_before_emission() {
        let reference = [0x08, 0x2a];
        let mut source = Vec::new();
        append_reference_record(&mut source, TITLE_REFERENCE_FIELD, &reference);
        append_reference_record(&mut source, OWNED_DRAWABLES_FIELD, &reference);
        append_reference_record(&mut source, Z_ORDER_FIELD, &reference);
        let mut budget = TransactionBudget {
            references: 0,
            fields: 0,
            work: 0,
            maximum_references: usize::MAX,
            maximum_fields: usize::MAX,
            maximum_work: source.len(),
        };
        assert!(matches!(
            rewrite_visibility_payload(
                &source,
                42,
                crate::slide::placeholder::Kind::Title,
                State::Visible,
                State::Hidden,
                WireLimits::default(),
                &mut budget,
            ),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed,
                maximum,
            }) if observed > maximum && maximum == source.len() as u64
        ));
    }
}
