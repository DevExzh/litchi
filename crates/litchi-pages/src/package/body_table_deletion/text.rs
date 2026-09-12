//! Source-preserving body-text anchor removal for Pages table deletion.
//!
//! The body storage is a `TSWP.StorageArchive` carried by the rooted body
//! object.  This phase only prepares its one-message replacement.  The graph
//! phase has already proved the selected table's ownership; this module adds
//! the final text witness that the selected UTF-16 unit is exactly the
//! replacement character belonging to the selected attachment.

use litchi_iwa_text_wire::{
    RewriteBehavior, RewriteError, RewriteLimits, StorageRewrite, StorageRewriteExecutionLimits,
    StorageRewriteExecutionReport, StorageRewritePrepareReport, StorageValidation,
    decode_storage_with_limits, prepare_storage_text_rewrite_with_behavior_and_limits,
    validate_storage_with_limits,
};
use std::num::NonZeroU64;

use super::{BodyTableDeletionError, GraphPlan, MessageEdit, Package, table_lock};

const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;
const TABLE_ATTACHMENT_STORAGE_FIELD: u32 = 9;

/// Prepare the body-storage message edit for one selected table.
///
/// The source catalog is already parsed and retained by [`Package`].  Every
/// source lookup below therefore borrows that catalog directly; no package
/// reopen or generated protobuf value is involved.  Text-wire performs the
/// strict schema walk and source-preserving rewrite, retaining unrelated text,
/// run tables, and unknown wire spans byte-for-byte.
pub(super) fn prepare(
    source: &Package,
    graph: &GraphPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<MessageEdit, BodyTableDeletionError> {
    let target = &graph.request.target;
    let source_catalog = &source.state.source;
    if !source_catalog.source_is_exact() {
        return Err(BodyTableDeletionError::UnsupportedSource);
    }

    let body_component = source_catalog
        .components()
        .get_index(target.body_component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let body_object = body_component
        .archive()
        .objects
        .get(target.body_object_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if body_object.archive_info.identifier != Some(target.body_identifier.get()) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    let body_message = body_object
        .messages
        .get(target.body_message_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if body_message.type_ != target.body_message_type {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    let body_message_info = body_object
        .archive_info
        .message_infos
        .get(target.body_message_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if body_message_info.type_ != body_message.type_ {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    let attachment_identifier = target.attachment_identifier.get();
    prove_attachment_header(body_message_info, attachment_identifier, budget)?;

    let payload = body_message.data.as_slice();
    // The target came from the same cached source catalog, but charge the
    // selected payload before any text projection can allocate. The strict
    // validation report below charges its finer-grained fields and work.
    budget
        .charge_payload_work(payload.len())
        .map_err(super::map_lock_error)?;

    let rewrite_limits = residual_rewrite_limits(source_catalog, budget)?;
    let validation =
        validate_storage_with_limits(payload, rewrite_limits).map_err(map_text_wire_error)?;
    charge_validation(validation, rewrite_limits, budget)?;

    // Decode through text-wire's archive-free projection only to prove the
    // selected UTF-16 unit.  The projection owns bounded semantic text and
    // never materializes a generated `StorageArchive` in Pages.
    let projection_limits = residual_rewrite_limits(source_catalog, budget)?;
    // Reserve the second strict walk and semantic projection before entering
    // text-wire's allocating decoder.  The decoder returns the same facts as
    // this already-proven validation; charging after it would leave the
    // temporary semantic text outside the cumulative transaction budget.
    charge_validation(validation, projection_limits, budget)?;
    let decoded =
        decode_storage_with_limits(payload, projection_limits).map_err(map_text_wire_error)?;
    let range_end = target
        .anchor_character_index
        .checked_add(1)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if range_end > validation.utf16_len() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    budget
        .charge_payload_work(range_end)
        .map_err(super::map_lock_error)?;
    prove_selected_anchor(decoded.storage().text(), target.anchor_character_index)?;

    // Validation and the bounded projection above have consumed part of the
    // shared wire budget. Recompute the profile before the output-free plan so
    // a large source cannot borrow the full codec allowance twice.
    let rewrite_limits = residual_rewrite_limits(source_catalog, budget)?;
    let prepared = prepare_storage_text_rewrite_with_behavior_and_limits(
        payload,
        target.anchor_character_index..range_end,
        "",
        RewriteBehavior::ReplaceSelection,
        rewrite_limits,
    )
    .map_err(map_text_wire_error)?;
    charge_prepare_report(prepared.prepare_report(), budget)?;
    let requirements = prepared.execution_requirements();
    charge_execution_requirements(requirements, budget)?;
    let rewritten = prepared
        .execute(StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(map_text_wire_error)?;
    verify_execution_report(&rewritten, requirements)?;
    verify_anchor_rewrite(
        rewritten,
        validation,
        target.attachment_identifier,
        target.body_component_index,
        target.body_object_index,
        target.body_message_index,
        target.body_message_type,
        target.body_identifier,
        budget,
    )
}

fn prove_selected_anchor(text: &str, anchor: usize) -> Result<(), BodyTableDeletionError> {
    text.encode_utf16()
        .enumerate()
        .find_map(|(index, unit)| (index == anchor).then_some(unit))
        .filter(|unit| *unit == OBJECT_REPLACEMENT_CHARACTER)
        .map(|_| ())
        .ok_or(BodyTableDeletionError::InvalidSource)
}

/// Prove the selected object's archive-header edge before decoding its body
/// storage. Object and data identifiers are separate metadata namespaces: a
/// data reference with the same numeric value as the selected object is
/// unrelated and must remain untouched by this object-reference edit.
fn prove_attachment_header(
    message: &litchi_iwa_core::MessageInfo,
    attachment: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut object_references = message.object_references.len();
    let mut data_references = message.data_references.len();
    let mut scan_work = object_references
        .checked_add(message.field_infos.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    for field in &message.field_infos {
        object_references = object_references
            .checked_add(field.object_references.len())
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        data_references = data_references
            .checked_add(field.data_references.len())
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        // Reserve the object membership/count and path comparison work before
        // any header scan starts. Data references are charged as inventory,
        // but deliberately never compared with the object identifier.
        scan_work = scan_work
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.path.path.len()))
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    budget
        .charge_payload_items(message.field_infos.len())
        .and_then(|_| budget.charge_payload_references(object_references))
        .and_then(|_| budget.charge_payload_references(data_references))
        .and_then(|_| budget.charge_payload_work(scan_work))
        .map_err(super::map_lock_error)?;

    if count_identifier(&message.object_references, attachment) != 1 {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    let mut attachment_field_occurrences = 0usize;
    for field in &message.field_infos {
        let occurrences = count_identifier(&field.object_references, attachment);
        if field.path.as_slice() == [TABLE_ATTACHMENT_STORAGE_FIELD] {
            // Native writers may aggregate several table attachments in one
            // field-9 declaration or repeat field-9 metadata for sibling
            // tables. Preserve those unrelated object edges; only the
            // selected edge must occur once when field metadata declares it.
            attachment_field_occurrences = attachment_field_occurrences
                .checked_add(occurrences)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
        } else if occurrences != 0 {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    // Some valid native sources retain the selected attachment only in the
    // aggregate MessageInfo.object_references inventory. If field metadata
    // declares the selected edge, however, it must identify it exactly once
    // among the field-9 declarations.
    if attachment_field_occurrences > 1 {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(())
}

fn verify_anchor_rewrite(
    rewritten: StorageRewrite,
    validation: StorageValidation,
    attachment: NonZeroU64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
    object_identifier: NonZeroU64,
    budget: &mut table_lock::WireBudget,
) -> Result<MessageEdit, BodyTableDeletionError> {
    let occurrence_before = rewritten.object_reference_occurrences_before().len();
    let occurrence_after = rewritten.object_reference_occurrences_after().len();
    let removed_by_field = rewritten.removed_object_references_by_field().len();
    let verification_work = rewritten
        .object_references_after()
        .len()
        .checked_add(
            occurrence_before
                .checked_mul(2)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(occurrence_after.checked_mul(2)?))
        .and_then(|value| value.checked_add(removed_by_field.checked_add(4)?))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(verification_work)
        .map_err(super::map_lock_error)?;
    let expected_before = validation.utf16_len();
    let expected_after = expected_before
        .checked_sub(1)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if !rewritten.changed()
        || rewritten.before_utf16_len() != expected_before
        || rewritten.after_utf16_len() != expected_after
        || (rewritten.bytes().is_empty() && expected_after != 0)
        || rewritten
            .object_references_after()
            .contains(&attachment.get())
        || rewritten.removed_object_references() != [attachment.get()]
    {
        return Err(BodyTableDeletionError::Verification);
    }

    let removed = rewritten.removed_object_references_by_field();
    if removed.len() != 1
        || removed[0].storage_field_number() != TABLE_ATTACHMENT_STORAGE_FIELD
        || removed[0].identifier() != attachment.get()
        || count_identifier(
            rewritten.object_reference_occurrences_before(),
            attachment.get(),
        ) != 1
        || count_identifier(
            rewritten.object_reference_occurrences_after(),
            attachment.get(),
        ) != 0
        || !occurrences_after_removing_selected(
            rewritten.object_reference_occurrences_before(),
            rewritten.object_reference_occurrences_after(),
            attachment.get(),
        )
    {
        return Err(BodyTableDeletionError::InvalidSource);
    }

    budget
        .charge_payload_items(1)
        .and_then(|_| budget.charge_payload_work(1))
        .map_err(super::map_lock_error)?;
    let mut remove_object_references = Vec::new();
    remove_object_references
        .try_reserve_exact(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    remove_object_references.push(attachment);
    let data = rewritten.into_bytes();
    Ok(MessageEdit {
        component_index,
        object_index,
        message_index,
        object_identifier,
        message_type,
        data,
        remove_object_references,
        remove_data_references: Vec::new(),
    })
}

fn count_identifier(values: &[u64], identifier: u64) -> usize {
    values
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
}

fn residual_rewrite_limits(
    source: &litchi_iwa_archive::SourceCatalog,
    budget: &table_lock::WireBudget,
) -> Result<RewriteLimits, BodyTableDeletionError> {
    let base =
        super::super::storage_rewrite_limits(source.limits()).map_err(map_storage_limits_error)?;
    let wire = budget.wire_limits();
    let fields = base.max_fields().min(wire.max_fields()).max(1);
    let fragments = base.max_fragments().min(fields).max(1);
    let table_entries = base.max_table_entries().min(fields).max(1);
    let object_references = base.max_object_references().min(wire.max_fields()).max(1);
    let rewrite_work = base.max_rewrite_work().min(wire.max_rewrite_work()).max(1);
    RewriteLimits::new(
        base.max_message_bytes().min(wire.max_input_bytes()).max(1),
        fields.min(budget.remaining_wire_fields().max(1)),
        base.max_nesting().min(wire.max_nesting()),
        fragments.min(budget.remaining_wire_fields().max(1)),
        base.max_text_bytes().min(wire.max_input_bytes()).max(1),
        table_entries.min(budget.remaining_wire_fields().max(1)),
        object_references.min(budget.remaining_payload_references().max(1)),
        base.max_output_bytes().min(wire.max_output_bytes()).max(1),
        rewrite_work.min(budget.remaining_wire_work().max(1)),
    )
    .map_err(map_text_wire_error)
}

fn occurrences_after_removing_selected(before: &[u64], after: &[u64], selected: u64) -> bool {
    let mut before_index = 0;
    let mut removed = false;
    for candidate in before {
        if !removed && *candidate == selected {
            removed = true;
            continue;
        }
        if after.get(before_index) != Some(candidate) {
            return false;
        }
        before_index += 1;
    }
    removed && before_index == after.len()
}

fn charge_validation(
    validation: StorageValidation,
    limits: RewriteLimits,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let max_depth = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    let table_items = validation
        .fragments()
        .checked_add(validation.table_entries())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let transaction_items = validation
        .utf8_len()
        .checked_add(validation.utf16_len())
        .and_then(|value| value.checked_add(validation.fragments()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_codec_report(
            validation.fields(),
            validation.validation_work(),
            max_depth,
            validation.reference_occurrences(),
        )
        .and_then(|_| budget.charge_payload_items(table_items))
        .and_then(|_| budget.charge_payload_work(transaction_items))
        .map_err(super::map_lock_error)
}

fn charge_prepare_report(
    report: StorageRewritePrepareReport,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let items = report
        .fragments()
        .checked_add(report.table_entries())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let transaction_work = report
        .text_bytes()
        .checked_add(report.text_units())
        .and_then(|value| value.checked_add(report.fragments()))
        .and_then(|value| value.checked_add(report.table_entries()))
        .and_then(|value| value.checked_add(report.max_nesting()))
        .and_then(|value| value.checked_add(usize::from(report.has_unknown_wire_fields())))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(report.input_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                report.fields(),
                report.work_bytes(),
                u32::try_from(report.max_nesting()).unwrap_or(u32::MAX),
                report.reference_occurrences(),
            )
        })
        .and_then(|_| budget.charge_payload_items(items))
        .and_then(|_| budget.charge_payload_work(transaction_work))
        .map_err(super::map_lock_error)
}

fn charge_execution_requirements(
    requirements: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.work()))
        .and_then(|_| budget.charge_payload_references(requirements.reference_occurrences()))
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.peak_scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(super::map_lock_error)
}

fn verify_execution_report(
    rewritten: &StorageRewrite,
    requirements: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
) -> Result<(), BodyTableDeletionError> {
    let report: StorageRewriteExecutionReport = rewritten.execution_report();
    if rewritten.bytes().len() != requirements.output_bytes()
        || report.retained_elements > requirements.retained_elements()
        || report.retained_bytes > requirements.retained_bytes()
        || report.peak_scratch_bytes > requirements.peak_scratch_bytes()
        || report.allocations > requirements.allocations()
        || report.work > requirements.work()
    {
        return Err(BodyTableDeletionError::Verification);
    }
    Ok(())
}

fn map_storage_limits_error(error: super::super::StorageWireLimitsError) -> BodyTableDeletionError {
    match error {
        super::super::StorageWireLimitsError::Physical(error) => super::map_archive_error(error),
        super::super::StorageWireLimitsError::Wire(error) => map_text_wire_error(error),
    }
}

fn map_text_wire_error(error: RewriteError) -> BodyTableDeletionError {
    match error {
        RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => BodyTableDeletionError::LimitExceeded {
            kind: match resource {
                "message bytes" => super::BodyTableDeletionLimitKind::WireBytes,
                "fields" => super::BodyTableDeletionLimitKind::WireFields,
                "nesting" => super::BodyTableDeletionLimitKind::WireNesting,
                "text fragments" | "table entries" => {
                    super::BodyTableDeletionLimitKind::PayloadItems
                },
                "text bytes" => super::BodyTableDeletionLimitKind::PayloadBytes,
                "object references" => super::BodyTableDeletionLimitKind::PayloadReferences,
                "output bytes" => super::BodyTableDeletionLimitKind::WireOutputBytes,
                _ => super::BodyTableDeletionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        RewriteError::Allocation { amount, .. } => BodyTableDeletionError::Allocation { amount },
        RewriteError::InvalidLimit { .. }
        | RewriteError::ReversedRange { .. }
        | RewriteError::RangeOutOfBounds { .. }
        | RewriteError::SurrogateSplit { .. }
        | RewriteError::ArithmeticOverflow { .. }
        | RewriteError::InvalidFormat(_)
        | RewriteError::Projection(_) => BodyTableDeletionError::InvalidSource,
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::{
        OBJECT_REPLACEMENT_CHARACTER, TABLE_ATTACHMENT_STORAGE_FIELD, prove_attachment_header,
        prove_selected_anchor,
    };
    use crate::package::table_lock;
    use litchi_iwa_common::varint::encode_varint_into;
    use litchi_iwa_core::{FieldInfo, FieldPath, MessageInfo};
    use litchi_iwa_text_wire::{
        RewriteBehavior, RewriteLimits, StorageRewriteExecutionLimits, decode_storage_with_limits,
        prepare_storage_text_rewrite_with_behavior_and_limits,
    };

    fn length_delimited(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        encode_varint_into(&mut output, u64::from(number) << 3 | 2);
        encode_varint_into(
            &mut output,
            u64::try_from(payload.len()).expect("test payload length fits in u64"),
        );
        output.extend_from_slice(payload);
        output
    }

    fn varint(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        encode_varint_into(&mut output, u64::from(number) << 3);
        encode_varint_into(&mut output, value);
        output
    }

    fn reference(identifier: u64) -> Vec<u8> {
        varint(1, identifier)
    }

    fn attachment_table(entries: &[(u32, u64)]) -> Vec<u8> {
        let mut table = Vec::new();
        for &(index, identifier) in entries {
            let entry = [
                varint(1, u64::from(index)),
                length_delimited(2, &reference(identifier)),
            ]
            .concat();
            table.extend(length_delimited(1, &entry));
        }
        table
    }

    fn storage(fragments: &[&str], entries: &[(u32, u64)]) -> Vec<u8> {
        let mut output = fragments
            .iter()
            .flat_map(|fragment| length_delimited(3, fragment.as_bytes()))
            .collect::<Vec<_>>();
        output.extend(length_delimited(
            TABLE_ATTACHMENT_STORAGE_FIELD,
            &attachment_table(entries),
        ));
        output
    }

    fn rewrite_anchor(
        source: &[u8],
        anchor: usize,
        attachment: u64,
    ) -> litchi_iwa_text_wire::StorageRewrite {
        let limits = RewriteLimits::default();
        let decoded = decode_storage_with_limits(source, limits).expect("valid storage");
        prove_selected_anchor(decoded.storage().text(), anchor).expect("selected anchor");
        let end = anchor.checked_add(1).expect("anchor end");
        let prepared = prepare_storage_text_rewrite_with_behavior_and_limits(
            source,
            anchor..end,
            "",
            RewriteBehavior::ReplaceSelection,
            limits,
        )
        .expect("prepared rewrite");
        let requirements = prepared.execution_requirements();
        let result = prepared
            .execute(StorageRewriteExecutionLimits {
                max_output_bytes: requirements.output_bytes(),
                max_retained_elements: requirements.retained_elements(),
                max_retained_bytes: requirements.retained_bytes(),
                max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
                max_allocations: requirements.allocations(),
                max_work: requirements.work(),
            })
            .expect("executed rewrite");
        assert_eq!(result.removed_object_references(), [attachment]);
        assert_eq!(
            result.removed_object_references_by_field()[0].storage_field_number(),
            TABLE_ATTACHMENT_STORAGE_FIELD
        );
        result
    }

    #[test]
    fn anchor_proof_rejects_non_replacement_character() {
        let source = storage(&["before after"], &[(7, 44)]);
        assert!(prove_selected_anchor("before after", 7).is_err());
        let limits = RewriteLimits::default();
        let decoded = decode_storage_with_limits(&source, limits).expect("valid storage");
        assert!(prove_selected_anchor(decoded.storage().text(), 7).is_err());
    }

    #[test]
    fn anchor_proof_rejects_emoji_surrogate_halves() {
        let text = "😀\u{fffc}";
        assert!(prove_selected_anchor(text, 0).is_err());
        assert!(prove_selected_anchor(text, 1).is_err());
        prove_selected_anchor(text, 2).expect("replacement character follows emoji");
    }

    #[test]
    fn attachment_header_proof_keeps_data_namespace_collision_unrelated() {
        let mut message = MessageInfo::new(2_001, 0);
        message.object_references = vec![44];
        message.data_references = vec![44];
        message.field_infos = vec![
            FieldInfo {
                path: FieldPath::new(vec![TABLE_ATTACHMENT_STORAGE_FIELD]),
                object_references: vec![44],
                ..FieldInfo::default()
            },
            FieldInfo {
                path: FieldPath::new(vec![17]),
                data_references: vec![44],
                ..FieldInfo::default()
            },
        ];
        let mut budget = table_lock::WireBudget::new(litchi_iwa_archive::Limits::default())
            .expect("default wire budget");

        prove_attachment_header(&message, 44, &mut budget)
            .expect("data and object identifiers use separate namespaces");
    }

    #[test]
    fn attachment_header_proof_allows_shared_field9_table_attachments() {
        let mut message = MessageInfo::new(2_001, 0);
        message.object_references = vec![44, 55];
        message.field_infos = vec![FieldInfo {
            path: FieldPath::new(vec![TABLE_ATTACHMENT_STORAGE_FIELD]),
            object_references: vec![44, 55],
            ..FieldInfo::default()
        }];
        let mut budget = table_lock::WireBudget::new(litchi_iwa_archive::Limits::default())
            .expect("default wire budget");

        prove_attachment_header(&message, 44, &mut budget)
            .expect("shared field-9 inventory should preserve the sibling edge");
    }

    #[test]
    fn attachment_header_proof_allows_repeated_field9_sibling_declarations() {
        let mut message = MessageInfo::new(2_001, 0);
        message.object_references = vec![44, 55];
        message.field_infos = vec![
            FieldInfo {
                path: FieldPath::new(vec![TABLE_ATTACHMENT_STORAGE_FIELD]),
                object_references: vec![44],
                ..FieldInfo::default()
            },
            FieldInfo {
                path: FieldPath::new(vec![TABLE_ATTACHMENT_STORAGE_FIELD]),
                object_references: vec![55],
                ..FieldInfo::default()
            },
        ];
        let mut budget = table_lock::WireBudget::new(litchi_iwa_archive::Limits::default())
            .expect("default wire budget");

        prove_attachment_header(&message, 44, &mut budget)
            .expect("sibling field-9 declaration should remain untouched");
    }

    #[test]
    fn attachment_header_proof_allows_aggregate_only_native_inventory() {
        let mut message = MessageInfo::new(2_001, 0);
        message.object_references = vec![44];
        let mut budget = table_lock::WireBudget::new(litchi_iwa_archive::Limits::default())
            .expect("default wire budget");

        prove_attachment_header(&message, 44, &mut budget)
            .expect("native aggregate-only attachment inventory is valid");
    }

    #[test]
    fn attachment_header_proof_rejects_selected_object_on_wrong_path() {
        let mut message = MessageInfo::new(2_001, 0);
        message.object_references = vec![44];
        message.field_infos = vec![FieldInfo {
            path: FieldPath::new(vec![17]),
            object_references: vec![44],
            ..FieldInfo::default()
        }];
        let mut budget = table_lock::WireBudget::new(litchi_iwa_archive::Limits::default())
            .expect("default wire budget");

        assert!(prove_attachment_header(&message, 44, &mut budget).is_err());
    }

    #[test]
    fn malformed_storage_is_rejected_before_anchor_rewrite() {
        let malformed = [0x1a, 0x01, 0xff];
        assert!(decode_storage_with_limits(&malformed, RewriteLimits::default()).is_err());
    }

    #[test]
    fn emoji_before_anchor_uses_utf16_and_removes_only_selected_table() {
        let source = storage(&["😀", "x\u{fffc}", "y\u{fffc}z"], &[(3, 44), (5, 55)]);
        let result = rewrite_anchor(&source, 3, 44);
        assert_eq!(result.before_utf16_len(), 7);
        assert_eq!(result.after_utf16_len(), 6);
        assert_eq!(result.removed_object_references(), [44]);
        assert_eq!(result.object_references_after(), [55]);
        let decoded = decode_storage_with_limits(result.bytes(), RewriteLimits::default())
            .expect("rewritten storage");
        assert_eq!(decoded.storage().text(), "😀xy\u{fffc}z");
    }

    #[test]
    fn unknown_root_span_survives_anchor_deletion() {
        let mut source = storage(&["a\u{fffc}b"], &[(1, 44)]);
        source.extend([0xa0, 0x06, 0x01]);
        let result = rewrite_anchor(&source, 1, 44);
        assert!(result.bytes().ends_with(&[0xa0, 0x06, 0x01]));
        assert_eq!(result.removed_object_references(), [44]);
    }

    #[test]
    fn replacement_unit_is_the_native_object_replacement_character() {
        assert_eq!(OBJECT_REPLACEMENT_CHARACTER, '\u{fffc}' as u16);
    }
}
