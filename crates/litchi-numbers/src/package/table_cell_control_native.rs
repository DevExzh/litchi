//! Private native seam for the unified cell-control owner.
//!
//! The selector/transaction facade lives in [`super::table_cell_control`].
//! This module intentionally contains no public archive or generated types;
//! scalar-control graph surgery will be added here as the source-preserving
//! list transition is shared with the audited Pop-Up Menu engine.

use std::mem::size_of;

use litchi_iwa_core::{
    Archive, ArchiveLimits, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, SnappyStream,
};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_number_format_codec as number_format_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
    package_metadata_codec::RewriteOptions as MetadataRewriteOptions,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};

use super::table_cell_pop_up_menu_native::NativePopUpError;
use super::{
    Package, table_cell_pop_up_menu as popup,
    table_cell_pop_up_menu::{CellTarget, Error, Path, TransactionBudget},
    table_cell_pop_up_menu_metadata as popup_metadata,
    table_cell_pop_up_menu_native as popup_native,
};
use crate::cell::data_format::CellControl;
use crate::cell::data_format::{
    control::DisplayFormat,
    number::{CurrencyStyle, DecimalPlaces, FractionAccuracy, NegativeStyle, ThousandsSeparator},
    numeral_system::{NegativeStyle as NumeralNegativeStyle, Places},
};

// BNC v5's fixed-layout format fields are private to the wire crate.  The
// package owner needs the masks here only to prove that a family-specific
// identifier is paired with the kind that gives it meaning; no native ID is
// exposed through this seam.
const BNC_FORMAT_FLAGS_START: usize = 8;
const BNC_FORMAT_FLAGS_END: usize = 12;
const BNC_CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_4000;
const BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_8000;
const BNC_DURATION_FORMAT_IDENTIFIER_FLAG: u32 = 0x0001_0000;
const BNC_TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG: u32 = 0x0004_0000;
const BNC_RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;

/// Validate the BNC format-field/kind relationship before using any format
/// identifier as a graph edge.
///
/// The low-level wire accessors are intentionally kind-directed: a Text
/// identifier is invisible when the kind field is absent, for example.  A
/// physical owner must not interpret that hidden field as harmless automatic
/// state, because clearing another visible edge could then leave the hidden
/// identifier dangling.  Shared generic identifiers are allowed only in the
/// native kinds where they are a primary/secondary edge; Text's marker and
/// generic-value rules are checked by its focused metadata validator.
pub(super) fn validate_bnc_format_metadata(
    source: &[u8],
    cell: &BncCell,
    path: Path,
) -> Result<(), Error> {
    let flags = source
        .get(BNC_FORMAT_FLAGS_START..BNC_FORMAT_FLAGS_END)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or(Error::InvalidSource { path })?;

    if flags & BNC_RESERVED_KNOWN_FIELD_FLAG != 0 {
        return Err(Error::UnsupportedDependency { path });
    }

    let kind = cell.cell_format_kind();
    let family_kind = |flag: u32, expected: u32| flags & flag != 0 && kind != Some(expected);
    if family_kind(
        BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG,
        litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND,
    ) || family_kind(
        BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG,
        litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND,
    ) || family_kind(
        BNC_DURATION_FORMAT_IDENTIFIER_FLAG,
        litchi_numbers_wire::DURATION_CELL_FORMAT_KIND,
    ) || family_kind(
        BNC_TEXT_FORMAT_IDENTIFIER_FLAG,
        litchi_numbers_wire::TEXT_CELL_FORMAT_KIND,
    ) || family_kind(
        BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG,
        litchi_numbers_wire::CHECKBOX_CELL_FORMAT_KIND,
    ) {
        return Err(Error::UnsupportedDependency { path });
    }

    // The shared identifier is a decimal primary, a Currency/Duration
    // secondary, or Text's retained Number edge.  It has no meaning without
    // a kind and is an orphan under Date/Time, Checkbox, or unknown kinds.
    if flags & BNC_CELL_FORMAT_IDENTIFIER_FLAG != 0
        && !matches!(
            kind,
            Some(
                litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND
                    | litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND
                    | litchi_numbers_wire::DURATION_CELL_FORMAT_KIND
                    | litchi_numbers_wire::TEXT_CELL_FORMAT_KIND
            )
        )
    {
        return Err(Error::UnsupportedDependency { path });
    }

    Ok(())
}

/// Marker used by the generic owner to keep native failures content-free.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
pub(super) enum NativeControlKind {
    Checkbox,
    StarRating,
    Slider,
    Stepper,
    PopUpMenu,
}

#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeControlMember {
    pub(super) member_name: String,
    pub(super) archive_bytes: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeControlOutput {
    pub(super) members: Vec<NativeControlMember>,
    pub(super) component_indices: Vec<usize>,
    pub(super) changed_objects: Vec<(usize, u64)>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct NativePreflight {
    pub(super) archive_bytes: usize,
    pub(super) compressed_bytes: usize,
}

/// Precharge the source archive and a conservative private scalar candidate
/// bound before archive cloning, list staging, or BNC rewrites allocate.
pub(super) fn preflight_copy_on_write_scalar_control(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativePreflight, Error> {
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let physical = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)?;
    let member = physical
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource { path })?;
    let archive = component.archive();
    // `member.data()` is the compressed IWA stream and was already charged by
    // the outer source-catalog preflight.  Archive accounting must use the
    // decompressed/native bound, otherwise a highly-compressible member can
    // bypass the decoded payload ceiling.
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_archive_bound = archive_serialized_bound(archive, archive_limits, path)?;
    budget.charge_archive(archive, source_archive_bound, path)?;

    // A split table graph can place the tile and the two DataList archives in
    // distinct members.  Reserve against the largest touched-member shape;
    // the native rewrite debits each actual member separately before it is
    // cloned or serialized.
    let mut graph_archive_bound = source_archive_bound;
    for candidate in source.state.components.catalog().iter() {
        graph_archive_bound = graph_archive_bound.max(archive_serialized_bound(
            candidate.archive(),
            archive_limits,
            path,
        )?);
    }
    let graph_member_bound = source
        .state
        .components
        .physical()
        .map(|physical| {
            physical
                .package()
                .iter()
                .map(|entry| entry.data().len())
                .max()
                .unwrap_or(member.data().len())
        })
        .unwrap_or_else(|| member.data().len());
    let archive_bytes = graph_archive_bound
        .saturating_add(graph_member_bound)
        .saturating_add(64 * 1024);
    let compressed_bytes = SnappyStream::maximum_compressed_len(archive_bytes)
        .map_err(|_| Error::InvalidSource { path })?;
    // These are operation-local private staging reservations.  The exact
    // prepared list requirements are charged by `apply_list_mutation`'s
    // resource-aware wrapper below; this allowance covers archive COW, tile
    // collection, and the BNC temporary buffers before those plans exist.
    budget.charge_allocations(8, path)?;
    budget.charge_transaction_work(
        archive_bytes
            .saturating_add(compressed_bytes)
            .saturating_add(archive.objects.len().saturating_mul(256)),
        path,
    )?;
    budget.charge_payload_objects(archive.objects.len().saturating_add(3), path)?;
    let messages = archive
        .objects
        .iter()
        .fold(0usize, |total, object| {
            total.saturating_add(object.messages.len())
        })
        .saturating_add(6);
    budget.charge_payload_messages(messages, path)?;
    budget.charge_payload_items(messages.saturating_add(archive.objects.len()), path)?;
    Ok(NativePreflight {
        archive_bytes,
        compressed_bytes,
    })
}

pub(super) fn archive_serialized_bound(
    archive: &Archive,
    limits: ArchiveLimits,
    path: Path,
) -> Result<usize, Error> {
    // `encoded_len_with_limits` is allocation-free and uses the same retained
    // raw-header selection as `to_bytes_with_limits`.  A structural estimate
    // is insufficient here: an ArchiveObject may preserve unknown/noncanonical
    // ArchiveInfo bytes in `original_header`, which can be larger than the
    // regenerated canonical header.  Resolve the exact bounded length before
    // charging any clone or serialization allocation.
    archive
        .encoded_len_with_limits(limits)
        .map_err(|_| Error::InvalidSource { path })
}

/// Rewrite a scalar control using the same strict list/storage primitives as
/// the Pop-Up Menu owner.  Popup-bearing transitions are deliberately left to
/// that owner so their model UUID/metadata lifecycle cannot be bypassed.
pub(super) fn rewrite_scalar_control(
    source: &Package,
    target: CellTarget,
    before: Option<&CellControl>,
    after: Option<&CellControl>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeControlOutput, Error> {
    if target.locked {
        return Err(Error::TableLocked { path });
    }
    if before.is_some_and(|value| matches!(value, CellControl::PopUpMenu(_)))
        || after.is_some_and(|value| matches!(value, CellControl::PopUpMenu(_)))
    {
        return Err(Error::UnsupportedDependency { path });
    }
    let model_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let model_archive = model_component.archive();
    let model_object = unique_object(model_archive, target.model_identifier, path)?;
    let model_index = unique_message_index(model_object, 6_001, path)?;
    let model_payload = &model_object.messages[model_index].data;
    let model_options = budget.residual_storage_options(model_payload);
    let (model, model_report) =
        storage_codec::decode_table_model_with_report(model_payload, model_options)
            .map_err(|error| map_storage_error(error, path))?;
    charge_storage_report(budget, model_report, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model.base_data_store(),
        budget.residual_storage_options(model.base_data_store()),
    )
    .map_err(|error| map_storage_error(error, path))?;
    charge_storage_report(budget, store_report, path)?;
    let tile_options = budget.residual_storage_options(store.tiles());
    let mut tile_visitor = TileReferenceCollector::new(budget, path);
    let decoded_tiles = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        tile_options,
        &mut tile_visitor,
    );
    let tile_references = tile_visitor.finish()?;
    let (tile_storage, tile_report) =
        decoded_tiles.map_err(|error| map_storage_error(error, path))?;
    charge_storage_report(budget, tile_report, path)?;
    let tile_id = target.position.row() / tile_storage.tile_size().unwrap_or(1).max(1);
    if tile_references.iter().any(|(_, reference)| *reference == 0)
        || tile_references
            .iter()
            .enumerate()
            .any(|(index, (id, reference))| {
                tile_references[index + 1..]
                    .iter()
                    .any(|(other_id, other_reference)| {
                        id == other_id || reference == other_reference
                    })
            })
    {
        return Err(Error::InvalidSource { path });
    }
    let tile_identifier = tile_references
        .iter()
        .find(|(id, _)| *id == tile_id)
        .map(|(_, reference)| *reference)
        .ok_or(Error::CellNotFound)?;
    if tile_identifier == target.model_identifier {
        return Err(Error::UnsupportedDependency { path });
    }
    let tile_component_index = resolved_component_index(source, tile_identifier, path)?;
    let tile_archive = source
        .state
        .components
        .catalog()
        .get_index(tile_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let tile_object = unique_object(tile_archive, tile_identifier, path)?;
    let tile_message_index = unique_message_index(tile_object, 6_002, path)?;
    let tile_payload =
        copy_payload_with_budget(&tile_object.messages[tile_message_index].data, budget, path)?;
    let cell_source = popup_native::tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
    let old_format = cell.format_identifier();
    let old_control = cell.control_cell_spec_identifier();
    if old_format.is_none() && old_control.is_some()
        || old_format.is_some() != old_control.is_some() && cell.cell_format_kind() == Some(5)
    {
        return Err(Error::InvalidSource { path });
    }
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let control_table_identifier = store
        .control_cell_spec_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let format_component_index = resolved_component_index(source, format_table_identifier, path)?;
    let format_archive = source
        .state
        .components
        .catalog()
        .get_index(format_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let control_component_index = resolved_component_index(source, control_table_identifier, path)?;
    let control_archive = source
        .state
        .components
        .catalog()
        .get_index(control_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let mut metadata_authority = None;
    popup::prove_cross_component_reference(
        source,
        &mut metadata_authority,
        target.component_index,
        tile_component_index,
        Some(tile_identifier),
        path,
    )?;
    popup::prove_cross_component_reference(
        source,
        &mut metadata_authority,
        target.component_index,
        format_component_index,
        Some(format_table_identifier),
        path,
    )?;
    popup::prove_cross_component_reference(
        source,
        &mut metadata_authority,
        target.component_index,
        control_component_index,
        Some(control_table_identifier),
        path,
    )?;
    let format_object = unique_object(format_archive, format_table_identifier, path)?;
    let control_object = unique_object(control_archive, control_table_identifier, path)?;
    let format_message = unique_list_message(format_object, 2, budget, path)?;
    let control_message = unique_list_message(control_object, 12, budget, path)?;
    let format_message_index = format_message.message_index;
    let control_message_index = control_message.message_index;
    let format_payload = copy_payload_with_budget(format_message.payload, budget, path)?;
    let control_payload = copy_payload_with_budget(control_message.payload, budget, path)?;
    let format_facts = list_facts(&format_payload, budget, path)?;
    let control_facts = list_facts(&control_payload, budget, path)?;
    validate_refcounts(
        source,
        &tile_references,
        format_facts.entries.as_slice(),
        control_facts.entries.as_slice(),
        budget,
        path,
    )?;
    let control_source_len = format_payload.len().max(control_payload.len());
    let owned_components = [
        target.component_index,
        tile_component_index,
        format_component_index,
        control_component_index,
    ];
    let mut popup_ownership = Vec::new();
    for entry in &control_facts.entries {
        let (spec, report) = control_codec::decode_any_cell_spec_with_report(
            &entry.payload,
            control_codec_options(entry.payload.len(), budget),
        )
        .map_err(|error| map_control_error(error, path))?;
        charge_control_decode_report(budget, report, path)?;
        let control_codec::CellSpecSnapshot::Popup(spec) = spec else {
            continue;
        };
        let identifier = spec.popup_model().identifier();
        let popup_component_index = resolved_component_index(source, identifier, path)?;
        popup::prove_cross_component_reference(
            source,
            &mut metadata_authority,
            control_component_index,
            popup_component_index,
            Some(identifier),
            path,
        )?;
        popup_ownership.push((popup_component_index, identifier));
        let popup_archive = source
            .state
            .components
            .catalog()
            .get_index(popup_component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let popup = unique_object(popup_archive, identifier, path)?;
        let popup_index = unique_message_index(popup, 6_206, path)?;
        let (_, report) = popup_codec::decode_popup_menu_model_with_report(
            &popup.messages[popup_index].data,
            budget.residual_popup_options(&popup.messages[popup_index].data),
        )
        .map_err(|_| Error::InvalidSource { path })?;
        budget.charge_wire_bytes(report.input_bytes(), path)?;
        budget.charge_wire_fields(report.fields(), path)?;
        budget.charge_wire_work(report.work_bytes(), path)?;
        budget.charge_payload_references(report.references(), path)?;
        budget.charge_payload_items(report.items(), path)?;
        budget.charge_transaction_work(
            report
                .output_bytes()
                .saturating_add(report.retained_bytes())
                .saturating_add(report.scratch_bytes()),
            path,
        )?;
        budget.charge_allocations(report.allocations(), path)?;
        let _ = identifier;
    }
    let metadata_facts = validate_metadata_ownership(source, &owned_components, budget, path)?;
    for (component_index, identifier) in popup_ownership {
        metadata_facts
            .require_current_uuid(component_index, identifier)
            .map_err(|_| Error::InvalidSource { path })?;
    }

    let (desired_kind, desired_format_payload, desired_spec_payload) = match after {
        Some(CellControl::Checkbox(_)) => (
            NativeControlKind::Checkbox,
            canonical_format(263, None, control_source_len, budget, path)?,
            canonical_spec(
                control_codec::CHECKBOX_INTERACTION_TYPE,
                None,
                control_source_len,
                budget,
                path,
            )?,
        ),
        Some(CellControl::StarRating(_)) => (
            NativeControlKind::StarRating,
            canonical_format(267, None, control_source_len, budget, path)?,
            canonical_spec(
                control_codec::STAR_RATING_INTERACTION_TYPE,
                Some((0.0, 5.0, 1.0)),
                control_source_len,
                budget,
                path,
            )?,
        ),
        Some(CellControl::Slider(value)) => (
            NativeControlKind::Slider,
            canonical_format(
                display_format_type(value.display_format(), path)?,
                Some(value.display_format()),
                control_source_len,
                budget,
                path,
            )?,
            canonical_spec(
                control_codec::SLIDER_INTERACTION_TYPE,
                Some((
                    value.range().minimum(),
                    value.range().maximum(),
                    value.range().increment(),
                )),
                control_source_len,
                budget,
                path,
            )?,
        ),
        Some(CellControl::Stepper(value)) => (
            NativeControlKind::Stepper,
            canonical_format(
                display_format_type(value.display_format(), path)?,
                Some(value.display_format()),
                control_source_len,
                budget,
                path,
            )?,
            canonical_spec(
                control_codec::STEPPER_INTERACTION_TYPE,
                Some((
                    value.range().minimum(),
                    value.range().maximum(),
                    value.range().increment(),
                )),
                control_source_len,
                budget,
                path,
            )?,
        ),
        Some(CellControl::PopUpMenu(_)) => return Err(Error::UnsupportedDependency { path }),
        None => (NativeControlKind::Checkbox, Vec::new(), Vec::new()),
    };
    let _ = desired_kind;
    let (new_format, new_format_key) = mutate_list(
        &format_payload,
        old_format,
        if after.is_some() {
            Some(desired_format_payload.as_slice())
        } else {
            None
        },
        true,
        budget,
        path,
    )?;
    let (new_control, new_control_key) = mutate_list(
        &control_payload,
        old_control,
        if after.is_some() {
            Some(desired_spec_payload.as_slice())
        } else {
            None
        },
        false,
        budget,
        path,
    )?;
    let replacement_cell = if let Some(after) = after {
        let (format_key, control_key) = (
            new_format_key.ok_or(Error::InvalidSource { path })?,
            new_control_key.ok_or(Error::InvalidSource { path })?,
        );
        let mut cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
        let (kind, _control) = native_kind(after, path)?;
        cell.set_data_format_identifier(format_key, kind, Some(control_key))
            .map_err(|_| Error::UnsupportedDependency { path })?;
        cell.encode()
    } else {
        let mut cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
        cell.clear_explicit_format();
        cell.encode()
    };
    let patched_tile = popup_native::patch_tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
        &replacement_cell,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let mut component_indices = vec![
        tile_component_index,
        format_component_index,
        control_component_index,
    ];
    component_indices.sort_unstable();
    component_indices.dedup();
    let mut archives = component_indices
        .iter()
        .map(|&component_index| {
            let component = source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?;
            budget.charge_archive(
                component.archive(),
                archive_serialized_bound(component.archive(), archive_limits, path)?,
                path,
            )?;
            Ok::<_, Error>((component_index, component.archive().clone()))
        })
        .collect::<Result<Vec<_>, Error>>()?;
    popup_native::validate_message_without_object_references(
        format_archive,
        format_table_identifier,
        format_message_index,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        archive_for_mut(&mut archives, format_component_index, path)?,
        format_table_identifier,
        format_message_index,
        new_format,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_control_message_with_transition(
        archive_for_mut(&mut archives, control_component_index, path)?,
        control_table_identifier,
        control_message_index,
        control_payload.as_slice(),
        new_control,
        archive_limits,
        true,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        archive_for_mut(&mut archives, tile_component_index, path)?,
        tile_identifier,
        tile_message_index,
        patched_tile,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;

    // `to_bytes_with_limits` performs the same length calculation internally,
    // but its output Vec is allocated before this transaction budget can see
    // the result.  Preflight every source/candidate pair while both archives
    // are still private, so the aggregate scratch/work/retained ceilings and
    // serialization reservations are charged before the first output buffer
    // is allocated.  Charging every pair is conservative: the scalar owner
    // currently publishes every candidate member, while a future exact-byte
    // filter may prove some candidates unchanged.
    let mut source_serialized_bytes = 0usize;
    let mut candidate_serialized_bytes = 0usize;
    for (component_index, candidate) in &archives {
        let source_archive = source
            .state
            .components
            .catalog()
            .get_index(*component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let source_length = source_archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let candidate_length = candidate
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        source_serialized_bytes = source_serialized_bytes
            .checked_add(source_length)
            .ok_or(Error::InvalidSource { path })?;
        candidate_serialized_bytes = candidate_serialized_bytes
            .checked_add(candidate_length)
            .ok_or(Error::InvalidSource { path })?;
    }
    let comparison_bytes = source_serialized_bytes
        .checked_add(candidate_serialized_bytes)
        .ok_or(Error::InvalidSource { path })?;
    let serialization_allocations = archives
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(serialization_allocations, path)?;
    budget.charge_scratch_bytes(comparison_bytes, path)?;
    budget.charge_retained_bytes(candidate_serialized_bytes, path)?;
    budget.charge_transaction_work(comparison_bytes, path)?;

    let mut members = Vec::new();
    members
        .try_reserve_exact(archives.len())
        .map_err(|_| Error::Allocation {
            amount: archives.len(),
            path,
        })?;
    for (component_index, archive) in archives {
        let changed_identifiers = [
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
            (control_component_index, control_table_identifier),
        ]
        .iter()
        .filter_map(|(owner, identifier)| (*owner == component_index).then_some(*identifier))
        .collect::<Vec<_>>();
        popup_native::verify_archive_object_locality(
            source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?
                .archive(),
            &archive,
            &changed_identifiers,
        )
        .map_err(|_| Error::Verification)?;
        // Recheck the candidate length immediately before serialization.  The
        // aggregate preflight above covers every member; this assertion
        // proves that no late archive-header change can make the output Vec
        // exceed the amount charged to the transaction.
        let expected_archive_bytes = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let archive_bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        if archive_bytes.len() != expected_archive_bytes {
            return Err(Error::Verification);
        }
        budget
            .charge_payload_bytes(archive_bytes.len(), path)
            .map_err(|_| Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::PayloadBytes,
                observed: archive_bytes.len() as u64,
                maximum: u64::MAX,
                path,
            })?;
        let member_name = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?
            .name()
            .to_owned();
        members.push(NativeControlMember {
            member_name,
            archive_bytes,
        });
    }
    Ok(NativeControlOutput {
        members,
        component_indices,
        changed_objects: vec![
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
            (control_component_index, control_table_identifier),
        ],
    })
}

/// Focused Number and Percentage display-format routes live in the private
/// family-parameterized native owner. Keep the Number spellings stable for
/// the existing transaction facade; the Percentage facade imports the shared
/// owner directly.
pub(super) use super::table_cell_display_format_native::{
    NumberFormatReadError, TextFormatReadError, read_number_format, read_number_format_with_budget,
    read_text_format, read_text_format_with_budget, rewrite_number_format, rewrite_text_format,
};

#[derive(Clone)]
pub(super) struct ListEntry {
    pub(super) key: u32,
    pub(super) ref_count: u32,
    pub(super) payload: Vec<u8>,
    pub(super) is_format: bool,
}

pub(super) struct ListFacts {
    pub(super) next_key: u32,
    pub(super) entries: Vec<ListEntry>,
}

/// Copy a borrowed native payload into transaction-owned staging memory.
///
/// The charge deliberately precedes `try_reserve_exact`: a hostile payload
/// must be rejected by the transaction ledger before the allocator is asked
/// to materialize it.  These bytes are retained until the surrounding native
/// operation has finished, rather than being counted as decoded wire input.
pub(super) fn copy_payload_with_budget(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    if source.is_empty() {
        return Ok(Vec::new());
    }
    budget.charge_allocations(1, path)?;
    budget.charge_retained_bytes(source.len(), path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
            amount: source.len(),
            path,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

/// Reserve more visitor slots while charging the expected growth footprint
/// before the allocator can grow the vector.  Callers only invoke this when
/// the current vector is full, so successful pushes do not over-debit the
/// transaction for existing capacity. Doubling the requested capacity keeps
/// hostile-but-valid lists from turning one fallible reservation per record
/// into quadratic copying work.
fn reserve_vec_slot<T>(
    values: &mut Vec<T>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    if values.len() < values.capacity() {
        return Ok(());
    }
    let additional = values.capacity().max(1);
    let retained = additional
        .checked_mul(size_of::<T>())
        .ok_or(Error::Allocation {
            amount: additional,
            path,
        })?;
    budget.charge_allocations(1, path)?;
    budget.charge_retained_bytes(retained, path)?;
    values
        .try_reserve_exact(additional)
        .map_err(|_| Error::Allocation {
            amount: additional,
            path,
        })
}

/// Stop a storage visitor after preserving the precise package-layer reason.
/// The storage codec's allocation error is only a control-flow sentinel; the
/// caller extracts `failure` before mapping any codec result.
fn visitor_failure(
    failure: &mut Option<Error>,
    error: Error,
) -> Result<(), storage_codec::DecodeError> {
    *failure = Some(error);
    Err(storage_codec::DecodeError::allocation(1))
}

pub(super) struct TileReferenceCollector<'budget> {
    pub(super) tiles: Vec<(u32, u64)>,
    budget: &'budget mut TransactionBudget,
    path: Path,
    failure: Option<Error>,
}

impl<'budget> TileReferenceCollector<'budget> {
    pub(super) fn new(budget: &'budget mut TransactionBudget, path: Path) -> Self {
        Self {
            tiles: Vec::new(),
            budget,
            path,
            failure: None,
        }
    }

    pub(super) fn finish(self) -> Result<Vec<(u32, u64)>, Error> {
        if let Some(error) = self.failure {
            Err(error)
        } else {
            Ok(self.tiles)
        }
    }
}

impl storage_codec::StorageVisitor for TileReferenceCollector<'_> {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Err(error) = reserve_vec_slot(&mut self.tiles, self.budget, self.path) {
            return visitor_failure(&mut self.failure, error);
        }
        self.tiles
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

struct BncReferenceCensus<'budget> {
    formats: Vec<(u32, usize)>,
    converted_text_generics: Vec<u32>,
    controls: Vec<(u32, usize)>,
    invalid: bool,
    budget: &'budget mut TransactionBudget,
    path: Path,
    failure: Option<Error>,
}

impl<'budget> BncReferenceCensus<'budget> {
    fn new(budget: &'budget mut TransactionBudget, path: Path) -> Self {
        Self {
            formats: Vec::new(),
            converted_text_generics: Vec::new(),
            controls: Vec::new(),
            invalid: false,
            budget,
            path,
            failure: None,
        }
    }

    fn residual_storage_options(&self, source: &[u8]) -> storage_codec::DecodeOptions {
        self.budget.residual_storage_options(source)
    }

    fn charge_storage_report(&mut self, report: storage_codec::DecodeReport) -> Result<(), Error> {
        charge_storage_report(self.budget, report, self.path)
    }

    fn increment(
        entries: &mut Vec<(u32, usize)>,
        identifier: u32,
        budget: &mut TransactionBudget,
        path: Path,
    ) -> Result<bool, Error> {
        if let Some((_, count)) = entries.iter_mut().find(|(key, _)| *key == identifier) {
            let Some(next) = count.checked_add(1) else {
                return Ok(false);
            };
            *count = next;
        } else {
            reserve_vec_slot(entries, budget, path)?;
            entries.push((identifier, 1));
        }
        Ok(true)
    }

    fn record_converted_text_generic(&mut self, identifier: u32) -> Result<(), Error> {
        if self.converted_text_generics.contains(&identifier) {
            return Ok(());
        }
        reserve_vec_slot(&mut self.converted_text_generics, self.budget, self.path)?;
        self.converted_text_generics.push(identifier);
        Ok(())
    }
}

impl storage_codec::StorageVisitor for BncReferenceCensus<'_> {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if self.invalid {
            return Ok(());
        }
        let storage = row
            .cell_storage_buffer()
            .unwrap_or_else(|| row.cell_storage_buffer_pre_bnc());
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc());
        if row.cell_count() == 0
            && (!storage.is_empty()
                || !row.cell_storage_buffer_pre_bnc().is_empty()
                || row
                    .cell_storage_buffer()
                    .is_some_and(|candidate| !candidate.is_empty()))
        {
            // A row that declares no materialized cells cannot carry a hidden
            // BNC payload, regardless of whether its offset table is empty or
            // merely padded with missing-column sentinels.
            self.invalid = true;
            return Ok(());
        }
        if offsets.is_empty() {
            match row.cell_count() {
                0 => {
                    return Ok(());
                },
                1 => {
                    if !self.count_cell(storage) {
                        if self.failure.is_some() {
                            return Err(storage_codec::DecodeError::allocation(1));
                        }
                        self.invalid = true;
                    }
                    return Ok(());
                },
                _ => {
                    self.invalid = true;
                    return Ok(());
                },
            }
        }
        if !offsets.len().is_multiple_of(2) {
            self.invalid = true;
            return Ok(());
        }
        let slot_count = offsets.len() / 2;
        let expected = usize::try_from(row.cell_count()).unwrap_or(usize::MAX);
        if expected > slot_count {
            self.invalid = true;
            return Ok(());
        }
        let unit = if row.has_wide_offsets().unwrap_or(false) {
            4usize
        } else {
            1usize
        };
        let mut occupied = 0usize;
        let mut previous = None;
        for encoded in offsets.chunks_exact(2) {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            if raw == u16::MAX {
                continue;
            }
            let Some(start) = usize::from(raw).checked_mul(unit) else {
                self.invalid = true;
                return Ok(());
            };
            // Automatic/minimal BNC cells have an empty payload. Adjacent
            // materialized columns may therefore share one offset, including
            // an offset at the end of the packed buffer. Only decreasing or
            // out-of-range starts are structurally invalid.
            if start > storage.len() || previous.is_some_and(|prior| prior > start) {
                self.invalid = true;
                return Ok(());
            }
            if let Some(prior) = previous {
                if !self.count_cell(&storage[prior..start]) {
                    if self.failure.is_some() {
                        return Err(storage_codec::DecodeError::allocation(1));
                    }
                    self.invalid = true;
                    return Ok(());
                }
            }
            let Some(next) = occupied.checked_add(1) else {
                self.invalid = true;
                return Ok(());
            };
            occupied = next;
            previous = Some(start);
        }
        if let Some(start) = previous {
            if !self.count_cell(&storage[start..]) {
                if self.failure.is_some() {
                    return Err(storage_codec::DecodeError::allocation(1));
                }
                self.invalid = true;
                return Ok(());
            }
        }
        if occupied != expected {
            self.invalid = true;
        }
        Ok(())
    }
}

impl BncReferenceCensus<'_> {
    fn count_cell(&mut self, source: &[u8]) -> bool {
        let Ok(cell) = BncCell::parse(source) else {
            return false;
        };
        if validate_bnc_format_metadata(source, &cell, self.path).is_err() {
            return false;
        }
        let format = cell.format_identifier();
        let control = cell.control_cell_spec_identifier();
        // Ordinary text/number cells legitimately own a format entry without
        // a control spec. A control key without a format key is the malformed
        // direction; interactive kind 5 is distinguished later by the strict
        // selected CellSpec, not merely by the presence of a format key.
        if format.is_none() && control.is_some() {
            return false;
        }
        if let Some(identifier) = format {
            match Self::increment(&mut self.formats, identifier, self.budget, self.path) {
                Ok(true) => {},
                Ok(false) => return false,
                Err(error) => {
                    self.failure = Some(error);
                    return false;
                },
            }
        }
        if let Some(identifier) = cell.secondary_format_identifier() {
            match Self::increment(&mut self.formats, identifier, self.budget, self.path) {
                Ok(true) => {},
                Ok(false) => return false,
                Err(error) => {
                    self.failure = Some(error);
                    return false;
                },
            }
        } else if control.is_none()
            && cell.cell_format_kind() == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)
        {
            // Older wire versions do not expose Text's retained generic
            // Number edge through `secondary_format_identifier`; recover it
            // only in this private census route. A later wire helper may
            // expose it directly, in which case the branch above wins and
            // prevents double-counting.
            let Ok(identifier) =
                super::table_cell_display_format_native::text_secondary_identifier_for_census(
                    source, &cell, self.path,
                )
            else {
                return false;
            };
            if let Some(identifier) = identifier {
                if let Err(error) = self.record_converted_text_generic(identifier) {
                    self.failure = Some(error);
                    return false;
                }
                match Self::increment(&mut self.formats, identifier, self.budget, self.path) {
                    Ok(true) => {},
                    Ok(false) => return false,
                    Err(error) => {
                        self.failure = Some(error);
                        return false;
                    },
                }
            }
        }
        if let Some(identifier) = control {
            match Self::increment(&mut self.controls, identifier, self.budget, self.path) {
                Ok(true) => {},
                Ok(false) => return false,
                Err(error) => {
                    self.failure = Some(error);
                    return false;
                },
            }
        }
        true
    }

    fn validate_converted_text_generic_targets(
        &mut self,
        formats: &[ListEntry],
    ) -> Result<(), Error> {
        // A converted Text cell carries a second edge in the shared generic
        // identifier slot. Refcount equality alone cannot prove that the
        // target entry is the required native Number format; validate every
        // distinct target after the complete tile walk, including targets
        // owned by nonselected cells.
        for index in 0..self.converted_text_generics.len() {
            let identifier = self.converted_text_generics[index];
            let entry = formats
                .iter()
                .find(|entry| entry.key == identifier)
                .ok_or(Error::InvalidSource { path: self.path })?;
            if !entry.is_format || entry.payload.is_empty() {
                return Err(Error::InvalidSource { path: self.path });
            }
            let (secondary, report) = number_format_codec::decode_number_format_with_report(
                &entry.payload,
                control_codec_options(entry.payload.len(), self.budget),
            )
            .map_err(|error| map_control_error(error, self.path))?;
            charge_control_decode_report(self.budget, report, self.path)?;
            if secondary.format_type() != number_format_codec::NATIVE_NUMBER_FORMAT_TYPE {
                return Err(Error::InvalidSource { path: self.path });
            }
        }
        Ok(())
    }
}

struct ListVisitor<'budget> {
    entries: Vec<ListEntry>,
    segments: usize,
    invalid_shape: bool,
    budget: &'budget mut TransactionBudget,
    path: Path,
    failure: Option<Error>,
}

impl<'budget> ListVisitor<'budget> {
    fn new(budget: &'budget mut TransactionBudget, path: Path) -> Self {
        Self {
            entries: Vec::new(),
            segments: 0,
            invalid_shape: false,
            budget,
            path,
            failure: None,
        }
    }

    fn finish(self) -> (Vec<ListEntry>, usize, bool, Option<Error>) {
        (
            self.entries,
            self.segments,
            self.invalid_shape,
            self.failure,
        )
    }
}

impl storage_codec::StorageVisitor for ListVisitor<'_> {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        let (payload_source, is_format) = match (snapshot.format(), snapshot.cell_spec()) {
            (Some(payload), None) => (Some(payload), true),
            (None, Some(payload)) => (Some(payload), false),
            _ => {
                self.invalid_shape = true;
                (None, false)
            },
        };
        if let Err(error) = reserve_vec_slot(&mut self.entries, self.budget, self.path) {
            return visitor_failure(&mut self.failure, error);
        }
        let payload = match payload_source {
            Some(payload) => match copy_payload_with_budget(payload, self.budget, self.path) {
                Ok(payload) => payload,
                Err(error) => return visitor_failure(&mut self.failure, error),
            },
            None => Vec::new(),
        };
        self.entries.push(ListEntry {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            payload,
            is_format,
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.segments = self.segments.saturating_add(1);
        Ok(())
    }
}

pub(super) fn list_facts(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListFacts, Error> {
    list_facts_with_input(source, budget, true, path)
}

pub(super) fn list_facts_without_input(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListFacts, Error> {
    list_facts_with_input(source, budget, false, path)
}

fn list_facts_with_input(
    source: &[u8],
    budget: &mut TransactionBudget,
    charge_input: bool,
    path: Path,
) -> Result<ListFacts, Error> {
    let options = budget.residual_storage_options(source);
    let mut visitor = ListVisitor::new(budget, path);
    let decoded = storage_codec::decode_table_data_list_with_visitor(source, options, &mut visitor);
    let (entries, segments, invalid_shape, failure) = visitor.finish();
    if let Some(error) = failure {
        return Err(error);
    }
    let (list, report) = decoded.map_err(|error| map_storage_error(error, path))?;
    if charge_input {
        charge_storage_report(budget, report, path)?;
    } else {
        charge_storage_report_without_input(budget, report, path)?;
    }
    budget.charge_payload_items(entries.len(), path)?;
    budget.charge_transaction_work(source.len(), path)?;
    let mut entries = entries;
    entries.sort_unstable_by_key(|entry| entry.key);
    if invalid_shape
        || segments != 0
        || entries.iter().any(|entry| entry.key == 0)
        || entries.iter().any(|entry| entry.ref_count == 0)
        || entries.windows(2).any(|pair| pair[0].key == pair[1].key)
    {
        return Err(Error::UnsupportedDependency { path });
    }
    // `next_list_id` is the source-authoritative allocation watermark.  A
    // stale/zero cursor would make deterministic COW allocation ambiguous and
    // could cause a new payload to collide with an existing key. Empty lists
    // still carry a positive cursor in valid native sources.
    let maximum_key = entries.last().map_or(0, |entry| entry.key);
    if list.next_list_id() == 0 || list.next_list_id() <= maximum_key {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(ListFacts {
        next_key: list.next_list_id(),
        entries,
    })
}

pub(super) fn mutate_list(
    source: &[u8],
    old_key: Option<u32>,
    desired_payload: Option<&[u8]>,
    format: bool,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(Vec<u8>, Option<u32>), Error> {
    let facts = list_facts_without_input(source, budget, path)?;
    let matches = desired_payload.and_then(|payload| {
        facts
            .entries
            .iter()
            .find(|entry| entry.is_format == format && entry.payload == payload)
            .map(|entry| entry.key)
    });
    let (desired_key, next_list_id_advance) = if desired_payload.is_some() {
        if let Some(existing) = matches {
            (Some(existing), None)
        } else {
            let mut candidate = 1u32;
            for entry in &facts.entries {
                if entry.key < candidate {
                    continue;
                }
                if entry.key == candidate {
                    candidate = candidate
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                } else {
                    break;
                }
            }
            // Prefer a free key below the source-authoritative cursor. A
            // dense registry appends exactly at the cursor and advances it in
            // the same prepared source-preserving rewrite.
            if facts.next_key == 0 || candidate > facts.next_key {
                return Err(Error::UnsupportedDependency { path });
            }
            let advance = if candidate == facts.next_key {
                Some((
                    facts.next_key,
                    facts
                        .next_key
                        .checked_add(1)
                        .ok_or(Error::UnsupportedDependency { path })?,
                ))
            } else {
                None
            };
            (Some(candidate), advance)
        }
    } else {
        (None, None)
    };
    let mut output = copy_payload_with_budget(source, budget, path)?;
    if old_key != desired_key {
        if let Some(old_key) = old_key {
            let entry = facts
                .entries
                .iter()
                .find(|entry| entry.key == old_key)
                .ok_or(Error::InvalidSource { path })?;
            if entry.ref_count == 0 {
                return Err(Error::InvalidSource { path });
            }
            let mutation = if entry.ref_count > 1 {
                storage_codec::TableDataListEntryMutation::RefCount(
                    storage_codec::TableDataListEntryRefCountEdit::new(
                        old_key,
                        entry.ref_count,
                        entry.ref_count - 1,
                    ),
                )
            } else if entry.is_format {
                storage_codec::TableDataListEntryMutation::Remove(
                    storage_codec::TableDataListEntryRemovalSpec::format(
                        old_key,
                        entry.ref_count,
                        &entry.payload,
                    ),
                )
            } else {
                storage_codec::TableDataListEntryMutation::Remove(
                    storage_codec::TableDataListEntryRemovalSpec::control_cell_spec(
                        old_key,
                        entry.ref_count,
                        &entry.payload,
                    ),
                )
            };
            output = apply_list_mutation_with_budget(&output, mutation, budget, path)?;
        }
        if let (Some(desired_key), Some(payload)) = (desired_key, desired_payload) {
            let after = list_facts(&output, budget, path)?;
            if let Some(entry) = after.entries.iter().find(|entry| entry.key == desired_key) {
                let mutation = storage_codec::TableDataListEntryMutation::RefCount(
                    storage_codec::TableDataListEntryRefCountEdit::new(
                        desired_key,
                        entry.ref_count,
                        entry
                            .ref_count
                            .checked_add(1)
                            .ok_or(Error::InvalidSource { path })?,
                    ),
                );
                output = apply_list_mutation_with_budget(&output, mutation, budget, path)?;
            } else {
                let mut append = if format {
                    storage_codec::TableDataListEntryAppend::format(desired_key, 1, payload)
                } else {
                    storage_codec::TableDataListEntryAppend::control_cell_spec(
                        desired_key,
                        1,
                        payload,
                    )
                };
                if let Some((expected, replacement)) = next_list_id_advance {
                    append = append.advance_next_list_id(expected, replacement);
                }
                output = apply_list_mutation_with_budget(
                    &output,
                    storage_codec::TableDataListEntryMutation::Append(append),
                    budget,
                    path,
                )?;
            }
        }
    }
    Ok((output, desired_key))
}

pub(super) fn apply_list_mutation_with_budget(
    source: &[u8],
    mutation: storage_codec::TableDataListEntryMutation<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    // The scalar native preflight has already reserved a conservative private
    // staging envelope.  The prepared requirements below are the sole exact
    // debit for this list plan; do not charge a second source-sized allowance
    // here, or a source traversal would be counted twice.
    let options = budget.residual_storage_rewrite_options(source);
    let prepared = storage_codec::prepare_table_data_list_entry_rewrite(source, mutation, options)
        .map_err(|error| map_storage_error(error, path))?;
    let preparation_report = prepared.prepare_report();
    // The prepared requirements below already aggregate source, payload,
    // result, and verification fields/work/references.  Charge only the
    // source-only reference/text byte classes from the preparation report;
    // charging its fields/work/references here would debit the same source
    // traversal twice.
    budget.charge_wire_bytes(preparation_report.source_bytes(), path)?;
    budget.charge_wire_reference_bytes(preparation_report.reference_bytes(), path)?;
    budget.charge_wire_text_bytes(preparation_report.text_bytes(), path)?;
    let requirements = prepared.requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_wire_nesting(requirements.max_depth(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_output(requirements.output_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)?;
    let exact_limits = requirements.exact_limits();
    let (output, report) = prepared
        .execute(exact_limits)
        .map_err(|error| map_storage_error(error, path))?;
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.references() != requirements.references()
        || report.scratch_bytes() != requirements.scratch_bytes()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
    {
        return Err(Error::Verification);
    }
    Ok(output)
}

pub(super) fn charge_storage_report(
    budget: &mut TransactionBudget,
    report: storage_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_bytes(report.source_bytes(), path)?;
    charge_storage_report_without_input(budget, report, path)
}

pub(super) fn charge_storage_report_without_input(
    budget: &mut TransactionBudget,
    report: storage_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_references(report.references(), path)?;
    budget.charge_wire_reference_bytes(report.reference_bytes(), path)?;
    budget.charge_wire_text_bytes(report.text_bytes(), path)?;
    Ok(())
}

pub(super) fn map_storage_error(error: storage_codec::DecodeError, path: Path) -> Error {
    match error.resource_limit() {
        Some(storage_codec::DecodeLimit::Bytes { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::References { observed, maximum }) => {
            Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::PayloadReferences,
                observed: observed as u64,
                maximum: maximum as u64,
                path,
            }
        },
        Some(storage_codec::DecodeLimit::Fields { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Work { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Nesting { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Allocation { requested }) => Error::Allocation {
            amount: requested,
            path,
        },
        Some(storage_codec::DecodeLimit::Retained { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Text { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireTextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        None | Some(_) => Error::InvalidSource { path },
    }
}

fn canonical_format(
    format_type: u32,
    display: Option<&DisplayFormat>,
    source_len: usize,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let options = control_codec_options(source_len, budget);
    match display {
        Some(DisplayFormat::Currency(value)) => {
            // CurrencyCode is a validated, three-byte semantic value. Keep a
            // local copy alive while the generated-free writer borrows its
            // string, so arbitrary valid codes do not require a leaked or
            // heap-owned string in the codec's borrowed write plan.
            let code = value.code();
            let write = decimal_format_write(
                control_codec::ControlFormatWrite::new(format_type)
                    .with_currency_code(code.as_str()),
                value.decimal_places(),
                value.negative_style(),
                value.thousands_separator(),
            )
            .with_use_accounting_style(matches!(value.style(), CurrencyStyle::Accounting));
            let prepared = popup_codec::prepare_control_format_write_fields(write, options)
                .map_err(|error| map_control_error(error, path))?;
            execute_control_format_plan(prepared, budget, path)
        },
        Some(display) => {
            let write = control_format_write(format_type, display, path)?;
            let prepared = popup_codec::prepare_control_format_write_fields(write, options)
                .map_err(|error| map_control_error(error, path))?;
            execute_control_format_plan(prepared, budget, path)
        },
        None => {
            let prepared = popup_codec::prepare_control_format_write_fields(
                control_codec::ControlFormatWrite::new(format_type),
                options,
            )
            .map_err(|error| map_control_error(error, path))?;
            execute_control_format_plan(prepared, budget, path)
        },
    }
}

fn execute_control_format_plan<'source>(
    prepared: popup_codec::PreparedControlFormatWrite<'source>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(control_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| map_control_error(error, path))?;
    verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn control_format_write(
    format_type: u32,
    display: &DisplayFormat,
    path: Path,
) -> Result<control_codec::ControlFormatWrite<'static>, Error> {
    if display_format_type(display, path)? != format_type {
        return Err(Error::InvalidSource { path });
    }
    let mut write = control_codec::ControlFormatWrite::new(format_type);
    match display {
        DisplayFormat::Number(value) => {
            write = decimal_format_write(
                write,
                value.decimal_places(),
                value.negative_style(),
                value.thousands_separator(),
            );
        },
        DisplayFormat::Percentage(value) => {
            write = decimal_format_write(
                write,
                value.decimal_places(),
                value.negative_style(),
                value.thousands_separator(),
            );
        },
        DisplayFormat::Currency(_) => {
            // Currency is handled in `canonical_format`, where the local
            // CurrencyCode copy can safely back the codec's borrowed string.
            return Err(Error::UnsupportedDependency { path });
        },
        DisplayFormat::Scientific(value) => {
            write = write.with_decimal_places(u32::from(value.decimal_places().value()));
        },
        DisplayFormat::Fraction(value) => {
            write = write.with_fraction_accuracy(match value.accuracy() {
                FractionAccuracy::UpToOneDigit => u32::MAX,
                FractionAccuracy::UpToTwoDigits => u32::MAX - 1,
                FractionAccuracy::UpToThreeDigits => u32::MAX - 2,
                FractionAccuracy::Halves => 2,
                FractionAccuracy::Quarters => 4,
                FractionAccuracy::Eighths => 8,
                FractionAccuracy::Sixteenths => 16,
                FractionAccuracy::Tenths => 10,
                FractionAccuracy::Hundredths => 100,
            });
        },
        DisplayFormat::NumeralSystem(value) => {
            write = write.with_base(u32::from(value.base().value()));
            write = write.with_base_places(match value.places() {
                Places::Minimum => 0,
                Places::Fixed(places) => u32::from(places.value()),
            });
            write = write.with_base_use_minus_sign(matches!(
                value.negative_style(),
                NumeralNegativeStyle::MinusSign
            ));
        },
    }
    Ok(write)
}

fn decimal_format_write<'source>(
    mut write: control_codec::ControlFormatWrite<'source>,
    decimal_places: DecimalPlaces,
    negative_style: NegativeStyle,
    thousands_separator: ThousandsSeparator,
) -> control_codec::ControlFormatWrite<'source> {
    write = write.with_decimal_places(match decimal_places {
        DecimalPlaces::Automatic => number_format_codec::NATIVE_AUTOMATIC_DECIMAL_PLACES,
        DecimalPlaces::Fixed(value) => u32::from(value.value()),
    });
    write
        .with_negative_style(match negative_style {
            NegativeStyle::MinusSign => 0,
            NegativeStyle::Red => 1,
            NegativeStyle::Parentheses => 2,
            NegativeStyle::RedParentheses => 3,
        })
        .with_show_thousands_separator(matches!(thousands_separator, ThousandsSeparator::Shown))
}

fn canonical_spec(
    interaction: u32,
    range: Option<(f64, f64, f64)>,
    source_len: usize,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let prepared = control_codec::prepare_control_cell_spec_write(
        interaction,
        range.map(|value| value.0),
        range.map(|value| value.1),
        range.map(|value| value.2),
        control_codec_options(source_len, budget),
    )
    .map_err(|error| map_control_error(error, path))?;
    charge_control_requirements(budget, prepared.execution_requirements(), path)?;
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(control_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| map_control_error(error, path))?;
    verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

pub(super) fn control_codec_options(
    source_len: usize,
    budget: &TransactionBudget,
) -> control_codec::DecodeOptions {
    budget.residual_control_options_for_len(source_len)
}

pub(super) fn charge_control_requirements(
    budget: &mut TransactionBudget,
    requirements: control_codec::RewriteExecutionRequirements,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_wire_nesting(requirements.max_depth(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_wire_text_bytes(requirements.text_bytes(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)
}

pub(super) fn charge_control_decode_report(
    budget: &mut TransactionBudget,
    report: control_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_bytes(report.input_bytes(), path)?;
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_references(report.references(), path)?;
    budget.charge_payload_items(report.items(), path)?;
    budget.charge_wire_text_bytes(report.text_bytes(), path)?;
    budget.charge_allocations(report.allocations(), path)?;
    budget.charge_scratch_bytes(report.scratch_bytes(), path)?;
    budget.charge_retained_bytes(report.retained_bytes(), path)?;
    budget.charge_transaction_work(report.output_bytes(), path)
}

/// Require the selected TableModel message to be the only package message
/// whose archive header owns an inbound edge to a mutable tile or format-list
/// object. A second table sharing either object would make a cell-local
/// rewrite affect data outside the selected semantic owner.
pub(super) fn require_exclusive_selected_model_reference(
    source: &Package,
    target: CellTarget,
    identifier: u64,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    struct Probe {
        identifier: u64,
        allowed: bool,
        expected_message: usize,
        matches: usize,
        unexpected: bool,
    }

    impl ArchiveReferenceVisitor for Probe {
        fn visit_reference(
            &mut self,
            occurrence: ArchiveReferenceOccurrence,
        ) -> litchi_iwa_core::Result<()> {
            if occurrence.kind == ArchiveReferenceKind::Object
                && occurrence.referenced_identifier == self.identifier
            {
                self.matches = self.matches.saturating_add(1);
                self.unexpected |=
                    !self.allowed || occurrence.message_index != self.expected_message;
            }
            Ok(())
        }
    }

    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let mut selected = 0usize;
    let mut inspected = 0usize;
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for object in &component.archive().objects {
            let object_identifier = object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource { path })?;
            let header_bytes =
                usize::try_from(object.header_length).map_err(|_| Error::InvalidSource { path })?;
            budget.charge_allocations(1, path)?;
            budget.charge_transaction_work(header_bytes.saturating_mul(4), path)?;
            let mut probe = Probe {
                identifier,
                allowed: component_index == target.component_index
                    && object_identifier == target.model_identifier,
                expected_message: target.message_index,
                matches: 0,
                unexpected: false,
            };
            let occurrences = object
                .inspect_references_with_policy_and_limits(
                    &mut probe,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::UnsupportedDependency { path })?;
            inspected = inspected
                .checked_add(occurrences)
                .ok_or(Error::InvalidSource { path })?;
            budget.charge_payload_references(occurrences, path)?;
            if probe.unexpected {
                return Err(Error::UnsupportedDependency { path });
            }
            selected = selected
                .checked_add(probe.matches)
                .ok_or(Error::InvalidSource { path })?;
        }
    }
    // The selected header contains one required aggregate edge and may mirror
    // it once in a typed FieldInfo. `validate_selected_model_reference`
    // already proves that exact shape; this package-wide census proves that
    // no other message or opaque metadata owns the mutable object.
    if selected == 0 || inspected == 0 {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(())
}

pub(super) fn validate_format_refcounts(
    source: &Package,
    tile_references: &[(u32, u64)],
    formats: &[ListEntry],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let mut census = BncReferenceCensus::new(budget, path);
    for (_, tile_identifier) in tile_references {
        let component_index = resolved_component_index(source, *tile_identifier, path)?;
        let archive = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let tile = unique_object(archive, *tile_identifier, path)?;
        let message_index = unique_message_index(tile, 6_002, path)?;
        let tile_payload = &tile.messages[message_index].data;
        let tile_options = census.residual_storage_options(tile_payload);
        let decoded =
            storage_codec::decode_tile_with_visitor(tile_payload, tile_options, &mut census);
        if let Some(error) = census.failure {
            return Err(error);
        }
        let (_, report) = decoded.map_err(|error| map_storage_error(error, path))?;
        census.charge_storage_report(report)?;
        if census.invalid {
            return Err(Error::InvalidSource { path });
        }
    }
    census.validate_converted_text_generic_targets(formats)?;
    for entry in formats {
        let observed = census
            .formats
            .iter()
            .find(|(identifier, _)| *identifier == entry.key)
            .map_or(0, |(_, count)| *count);
        if observed != usize::try_from(entry.ref_count).unwrap_or(usize::MAX) {
            return Err(Error::InvalidSource { path });
        }
    }
    if census
        .formats
        .iter()
        .any(|(identifier, _)| !formats.iter().any(|entry| entry.key == *identifier))
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

pub(super) fn verify_control_report(
    report: control_codec::DecodeReport,
    requirements: control_codec::RewriteExecutionRequirements,
    _path: Path,
) -> Result<(), Error> {
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.max_depth() != requirements.max_depth()
        || report.references() != requirements.references()
        || report.items() != requirements.items()
        || report.text_bytes() != requirements.text_bytes()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
        || report.scratch_bytes() != requirements.scratch_bytes()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

pub(super) fn map_control_error(error: control_codec::DecodeError, path: Path) -> Error {
    match error.resource_limit() {
        Some(control_codec::DecodeLimit::InputBytes { observed, maximum }) => {
            Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
                path,
            }
        },
        Some(control_codec::DecodeLimit::OutputBytes { observed, maximum }) => {
            Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::WireOutputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
                path,
            }
        },
        Some(control_codec::DecodeLimit::Fields { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Work { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Nesting { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::References { observed, maximum }) => {
            Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::PayloadReferences,
                observed: observed as u64,
                maximum: maximum as u64,
                path,
            }
        },
        Some(control_codec::DecodeLimit::Items { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::PayloadItems,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Text { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireTextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Allocation { requested }) => Error::Allocation {
            amount: requested,
            path,
        },
        Some(control_codec::DecodeLimit::Retained { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Scratch { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::ScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        None | Some(_) => Error::InvalidSource { path },
    }
}

fn display_format_type(format: &DisplayFormat, _path: Path) -> Result<u32, Error> {
    Ok(match format {
        DisplayFormat::Number(_) => 256,
        DisplayFormat::Currency(_) => 257,
        DisplayFormat::Percentage(_) => 258,
        DisplayFormat::Scientific(_) => 259,
        DisplayFormat::Fraction(_) => 262,
        DisplayFormat::NumeralSystem(_) => 269,
    })
}

fn native_kind(value: &CellControl, path: Path) -> Result<(CellDataFormatKind, u32), Error> {
    match value {
        CellControl::Checkbox(_) => Ok((CellDataFormatKind::Checkbox, 8)),
        CellControl::StarRating(_) => Ok((CellDataFormatKind::StarRating, 6)),
        CellControl::Slider(value) => Ok((
            match value.display_format() {
                DisplayFormat::Currency(_) => CellDataFormatKind::NumericControlCurrency,
                _ => CellDataFormatKind::NumericControlNumberOrPercentage,
            },
            5,
        )),
        CellControl::Stepper(value) => Ok((
            match value.display_format() {
                DisplayFormat::Currency(_) => CellDataFormatKind::NumericControlCurrency,
                _ => CellDataFormatKind::NumericControlNumberOrPercentage,
            },
            4,
        )),
        CellControl::PopUpMenu(_) => Err(Error::UnsupportedDependency { path }),
    }
}

pub(super) fn validate_metadata_ownership<'source>(
    source: &'source Package,
    component_indices: &[usize],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<popup_metadata::RegistryFacts<'source>, Error> {
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::InvalidSource { path })?;
    let bytes = catalog
        .package()
        .iter()
        .try_fold(0usize, |total, entry| total.checked_add(entry.data().len()))
        .ok_or(Error::InvalidSource { path })?
        .max(1);
    let options = MetadataRewriteOptions::new(
        bytes,
        bytes.saturating_mul(2),
        bytes.saturating_mul(16),
        bytes.saturating_mul(64),
        64,
        bytes,
        bytes.saturating_mul(2),
        1,
    );
    let facts =
        popup_metadata::inspect(source, options).map_err(|_| Error::InvalidSource { path })?;
    let report = facts.report();
    budget.charge_wire_bytes(report.input_bytes(), path)?;
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_payload_items(report.components_scanned(), path)?;
    budget.charge_payload_references(report.references_scanned(), path)?;
    budget.charge_allocations(report.allocations(), path)?;
    budget.charge_transaction_work(
        report
            .input_bytes()
            .saturating_add(report.output_bytes())
            .saturating_add(report.retained_bytes())
            .saturating_add(report.scratch_bytes()),
        path,
    )?;
    if facts.has_physical_alias() {
        return Err(Error::InvalidSource { path });
    }
    for (index, &component_index) in component_indices.iter().enumerate() {
        if component_indices[..index].contains(&component_index) {
            continue;
        }
        facts
            .selector(component_index)
            .map_err(|_| Error::InvalidSource { path })?;
    }
    Ok(facts)
}

pub(super) fn validate_cross_component_reference(
    facts: &popup_metadata::RegistryFacts<'_>,
    source_component_index: usize,
    target_component_index: usize,
    object_identifier: u64,
    path: Path,
) -> Result<(), Error> {
    if source_component_index == target_component_index {
        return Ok(());
    }
    facts
        .require_external_edge(
            source_component_index,
            target_component_index,
            Some(object_identifier),
            Some(false),
        )
        .map_err(|_| Error::UnsupportedDependency { path })
}

pub(super) fn unique_object(
    archive: &Archive,
    identifier: u64,
    path: Path,
) -> Result<&ArchiveObject, Error> {
    let mut matches = archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(identifier));
    let object = matches.next().ok_or(Error::InvalidSource { path })?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(object)
}

pub(super) fn unique_message_index(
    object: &ArchiveObject,
    type_: u32,
    path: Path,
) -> Result<usize, Error> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == type_);
    let index = matches
        .next()
        .map(|(index, _)| index)
        .ok_or(Error::InvalidSource { path })?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(index)
}

pub(super) fn archive_for_mut(
    archives: &mut [(usize, Archive)],
    component_index: usize,
    path: Path,
) -> Result<&mut Archive, Error> {
    archives
        .iter_mut()
        .find(|(index, _)| *index == component_index)
        .map(|(_, archive)| archive)
        .ok_or(Error::InvalidSource { path })
}

pub(super) fn resolved_component_index(
    source: &Package,
    identifier: u64,
    path: Path,
) -> Result<usize, Error> {
    source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .map(|resolved| resolved.component_index)
        .ok_or(Error::InvalidSource { path })
}

pub(super) struct ListMessage<'a> {
    pub(super) message_index: usize,
    pub(super) payload: &'a [u8],
}

pub(super) fn unique_list_message<'a>(
    object: &'a ArchiveObject,
    list_type: i32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListMessage<'a>, Error> {
    let mut selected = None;
    for (message_index, message) in object.messages.iter().enumerate() {
        if message.type_ != 6_005 {
            continue;
        }
        let (list, report) = storage_codec::decode_table_data_list_type_with_report(
            &message.data,
            budget.residual_storage_options(&message.data),
        )
        .map_err(|error| map_storage_error(error, path))?;
        charge_storage_report(budget, report, path)?;
        if list.list_type() != list_type {
            continue;
        }
        if selected.is_some() {
            return Err(Error::InvalidSource { path });
        }
        selected = Some(ListMessage {
            message_index,
            payload: &message.data,
        });
    }
    selected.ok_or(Error::InvalidSource { path })
}

fn validate_refcounts(
    source: &Package,
    tile_references: &[(u32, u64)],
    formats: &[ListEntry],
    controls: &[ListEntry],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let mut census = BncReferenceCensus::new(budget, path);
    for (_, tile_identifier) in tile_references {
        let component_index = resolved_component_index(source, *tile_identifier, path)?;
        let archive = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let tile = unique_object(archive, *tile_identifier, path)?;
        let message_index = unique_message_index(tile, 6_002, path)?;
        let tile_payload = &tile.messages[message_index].data;
        let tile_options = census.residual_storage_options(tile_payload);
        let decoded =
            storage_codec::decode_tile_with_visitor(tile_payload, tile_options, &mut census);
        if let Some(error) = census.failure {
            return Err(error);
        }
        let (_, report) = decoded.map_err(|error| map_storage_error(error, path))?;
        census.charge_storage_report(report)?;
        if census.invalid {
            return Err(Error::InvalidSource { path });
        }
    }
    census.validate_converted_text_generic_targets(formats)?;

    for entry in formats {
        let observed = census
            .formats
            .iter()
            .find(|(identifier, _)| *identifier == entry.key)
            .map_or(0, |(_, count)| *count);
        if observed != usize::try_from(entry.ref_count).unwrap_or(usize::MAX) {
            return Err(Error::InvalidSource { path });
        }
    }
    for entry in controls {
        let observed = census
            .controls
            .iter()
            .find(|(identifier, _)| *identifier == entry.key)
            .map_or(0, |(_, count)| *count);
        if observed != usize::try_from(entry.ref_count).unwrap_or(usize::MAX) {
            return Err(Error::InvalidSource { path });
        }
    }
    if census
        .formats
        .iter()
        .any(|(identifier, _)| !formats.iter().any(|entry| entry.key == *identifier))
        || census
            .controls
            .iter()
            .any(|(identifier, _)| !controls.iter().any(|entry| entry.key == *identifier))
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

/// Current implementation delegates Pop-Up Menu graph transitions to the
/// existing strict native owner.  Scalar controls are admitted only after
/// their corresponding source-preserving list transition is available.
#[allow(dead_code)]
pub(super) fn unsupported_scalar_control() -> NativePopUpError {
    NativePopUpError::UnsupportedDependency
}
