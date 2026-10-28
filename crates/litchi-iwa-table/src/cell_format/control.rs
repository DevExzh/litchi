//! Native seam of the cell-control owners.
//!
//! The selector/transaction facades live in the leaf crates (Numbers:
//! `package::table_cell_control`). This module contains no public generated
//! types; the scalar-control graph surgery shares its list transition with
//! the audited Pop-Up Menu engine.

use std::mem::size_of;
use std::sync::{Arc, OnceLock};

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, Limits as ArchiveLimits, SnappyStream,
};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_currency_format_codec as currency_format_codec,
    numbers_table_cell_duration_format_codec as duration_format_codec,
    numbers_table_cell_number_format_codec as number_format_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
    package_metadata_codec::RewriteOptions as MetadataRewriteOptions,
};
use litchi_numbers_wire::{BncCell, BncCellView, CellDataFormatKind};

use super::pop_up_menu_native::NativePopUpError;
use super::{
    Error, Path, Target, TransactionBudget, pop_up_menu as popup,
    pop_up_menu_metadata as popup_metadata, pop_up_menu_native as popup_native,
};
use crate::TableSource;
use litchi_iwa_common::table::cell::data_format::CellControl;
use litchi_iwa_common::table::cell::data_format::{
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
const BNC_CONTROL_CELL_SPEC_FLAG: u32 = 0x0000_0400;
const BNC_CELL_FORMAT_KIND_FLAG: u32 = 0x0000_1000;
const BNC_CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_4000;
const BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_8000;
const BNC_DURATION_FORMAT_IDENTIFIER_FLAG: u32 = 0x0001_0000;
const BNC_TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG: u32 = 0x0004_0000;
const BNC_RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;
const BNC_FORMAT_FLAGS: u32 = BNC_CONTROL_CELL_SPEC_FLAG
    | BNC_CELL_FORMAT_KIND_FLAG
    | BNC_CELL_FORMAT_IDENTIFIER_FLAG
    | BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG
    | BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG
    | BNC_DURATION_FORMAT_IDENTIFIER_FLAG
    | BNC_TEXT_FORMAT_IDENTIFIER_FLAG
    | BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG;
const BNC_TEXT_ALLOWED_FORMAT_FLAGS: u32 =
    BNC_CELL_FORMAT_KIND_FLAG | BNC_CELL_FORMAT_IDENTIFIER_FLAG | BNC_TEXT_FORMAT_IDENTIFIER_FLAG;
const BNC_FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0001, 16),
    (0x0000_0002, 8),
    (0x0000_0004, 8),
    (0x0000_0008, 4),
    (0x0000_0010, 4),
    (0x0000_0020, 4),
    (0x0000_0040, 4),
    (0x0000_0080, 4),
    (0x0000_0100, 4),
    (0x0000_0200, 4),
    (BNC_CONTROL_CELL_SPEC_FLAG, 4),
    (0x0000_0800, 4),
    (BNC_CELL_FORMAT_KIND_FLAG, 4),
    (BNC_CELL_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_DURATION_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_TEXT_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG, 4),
    (0x0008_0000, 4),
    (BNC_RESERVED_KNOWN_FIELD_FLAG, 4),
];

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
pub fn validate_bnc_format_metadata(
    source: &[u8],
    cell: &BncCell,
    path: Path,
) -> Result<(), Error> {
    validate_bnc_format_metadata_kind(source, cell.cell_format_kind(), path)
}

fn validate_bnc_format_metadata_kind(
    source: &[u8],
    kind: Option<u32>,
    path: Path,
) -> Result<(), Error> {
    let flags = source
        .get(BNC_FORMAT_FLAGS_START..BNC_FORMAT_FLAGS_END)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or(Error::InvalidSource { path })?;

    if flags & BNC_RESERVED_KNOWN_FIELD_FLAG != 0 {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }

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
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
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
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }

    Ok(())
}

/// Recover the generic Number edge retained by a converted Text cell without
/// materializing an owned [`BncCell`].  The package-wide census visits cells
/// that may carry arbitrary opaque tails; this fixed-width scan keeps that
/// path allocation-free while preserving the same strict Text metadata rules
/// as the focused Text owner.
fn text_secondary_identifier_for_census_view(
    source: &[u8],
    cell: &BncCellView<'_>,
    path: Path,
) -> Result<Option<u32>, Error> {
    if cell.cell_format_kind() != Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
        return Ok(None);
    }
    let flags = source
        .get(BNC_FORMAT_FLAGS_START..BNC_FORMAT_FLAGS_END)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or(Error::InvalidSource { path })?;
    if flags & BNC_RESERVED_KNOWN_FIELD_FLAG != 0
        || flags & BNC_FORMAT_FLAGS & !BNC_TEXT_ALLOWED_FORMAT_FLAGS != 0
    {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }

    let mut offset = BNC_FORMAT_FLAGS_END;
    let mut generic_identifier = None;
    let mut text_identifier = None;
    for &(flag, size) in BNC_FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        let end = offset
            .checked_add(size)
            .ok_or(Error::InvalidSource { path })?;
        let bytes = source
            .get(offset..end)
            .ok_or(Error::InvalidSource { path })?;
        if flag == BNC_CELL_FORMAT_IDENTIFIER_FLAG {
            generic_identifier = Some(u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_| Error::InvalidSource { path })?,
            ));
        } else if flag == BNC_TEXT_FORMAT_IDENTIFIER_FLAG {
            text_identifier = Some(u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_| Error::InvalidSource { path })?,
            ));
        }
        offset = end;
    }

    let valid_value = matches!(
        cell.stored_value(),
        litchi_numbers_wire::StoredValue::Empty | litchi_numbers_wire::StoredValue::Text(_)
    );
    let valid_marker = match cell.explicit_format_flags() {
        0 | litchi_numbers_wire::EXPLICIT_TEXT_FORMAT => generic_identifier.is_none(),
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT => {
            generic_identifier.is_some_and(|identifier| identifier != 0)
        },
        _ => false,
    };
    if !valid_value
        || !valid_marker
        || cell.control_cell_spec_identifier().is_some()
        || text_identifier.is_none_or(|identifier| identifier == 0)
    {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }
    if cell.explicit_format_flags() == litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT {
        Ok(generic_identifier)
    } else {
        Ok(None)
    }
}

/// Validate the complete metadata/value tuple for a native Currency or
/// Duration cell seen by the package-wide census.  The selected-cell routes
/// perform these checks before decoding their primary format; siblings need
/// the same admission so a later COW rewrite cannot strand an invalid tuple.
fn validate_native_cell_metadata(
    cell: &BncCellView<'_>,
    path: Path,
) -> Result<Option<(u32, NativePrimaryFormatKind)>, Error> {
    match cell.cell_format_kind() {
        Some(litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND) => {
            let primary = cell
                .format_identifier()
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?;
            let secondary = cell.secondary_format_identifier();
            let expected_marker =
                litchi_numbers_wire::explicit_currency_format_flags(secondary.is_some());
            // Currency marker zero represents native inherited/default state
            // only when no generic secondary edge is retained. Explicit
            // Currency tuples use the shape-dependent native marker.
            if (cell.explicit_format_flags() != 0
                && cell.explicit_format_flags() != expected_marker)
                || (cell.explicit_format_flags() == 0 && secondary.is_some())
                || secondary.is_some_and(|identifier| identifier == 0)
                || cell.control_cell_spec_identifier().is_some()
                || !currency_cell_type_matches_view(cell)
            {
                return Err(Error::Refused(
                    crate::cell_format::Refusal::UnsupportedDependency { path },
                ));
            }
            Ok(Some((primary, NativePrimaryFormatKind::Currency)))
        },
        Some(litchi_numbers_wire::DURATION_CELL_FORMAT_KIND) => {
            let primary = cell
                .format_identifier()
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?;
            let secondary = cell.secondary_format_identifier();
            let expected_marker =
                litchi_numbers_wire::explicit_duration_format_flags(secondary.is_some());
            if cell.explicit_format_flags() != expected_marker
                || secondary.is_some_and(|identifier| identifier == 0)
                || cell.control_cell_spec_identifier().is_some()
                || !cell.has_only_duration_format_metadata()
                || !cell.is_duration_format_compatible()
            {
                return Err(Error::Refused(
                    crate::cell_format::Refusal::UnsupportedDependency { path },
                ));
            }
            Ok(Some((primary, NativePrimaryFormatKind::Duration)))
        },
        _ => Ok(None),
    }
}

fn currency_cell_type_matches_view(cell: &BncCellView<'_>) -> bool {
    match cell.stored_value() {
        // Numbers can attach display metadata to an otherwise empty cell.
        // That shape has no numeric cell type and must remain empty through a
        // metadata-only set or clear operation.
        litchi_numbers_wire::StoredValue::Empty => cell.numeric_cell_type().is_none(),
        _ => {
            cell.numeric_cell_type() == Some(litchi_numbers_wire::NumericCellType::AlternateNumber)
        },
    }
}

/// Marker used by the generic owner to keep native failures content-free.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
pub enum NativeControlKind {
    Checkbox,
    StarRating,
    Slider,
    Stepper,
    PopUpMenu,
}

#[derive(Debug, PartialEq, Eq)]
pub struct NativeControlMember {
    pub member_name: String,
    pub archive_bytes: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
pub struct NativeControlOutput {
    pub members: Vec<NativeControlMember>,
    pub component_indices: Vec<usize>,
    pub changed_objects: Vec<(usize, u64)>,
}

#[derive(Debug, Clone, Copy)]
pub struct NativePreflight {
    pub archive_bytes: usize,
    pub compressed_bytes: usize,
}

/// Precharge the source archive and a conservative private scalar candidate
/// bound before archive cloning, list staging, or BNC rewrites allocate.
pub fn preflight_copy_on_write_scalar_control(
    source: &dyn TableSource,
    target: Target,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativePreflight, Error> {
    let component = source
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let physical = source.physical().ok_or(Error::Refused(
        crate::cell_format::Refusal::UnsupportedSource,
    ))?;
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
        .archive_limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_archive_bound = archive_serialized_bound(archive, archive_limits, path)?;
    budget.charge_archive(archive, source_archive_bound, path)?;

    // A split table graph can place the tile and the two DataList archives in
    // distinct members.  Reserve against the largest touched-member shape;
    // the native rewrite debits each actual member separately before it is
    // cloned or serialized.
    let mut graph_archive_bound = source_archive_bound;
    for candidate in source.catalog().iter() {
        graph_archive_bound = graph_archive_bound.max(archive_serialized_bound(
            candidate.archive(),
            archive_limits,
            path,
        )?);
    }
    let graph_member_bound = source
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
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let compressed_bytes = SnappyStream::maximum_compressed_len(archive_bytes, snappy_limits)
        .map_err(|error| super::map_core_error(error, path))?;
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

pub fn archive_serialized_bound(
    archive: &Archive,
    limits: ArchiveLimits,
    path: Path,
) -> Result<usize, Error> {
    // `encoded_len` is allocation-free and uses the same retained
    // raw-header selection as `to_bytes`.  A structural estimate
    // is insufficient here: an ArchiveObject may preserve unknown/noncanonical
    // ArchiveInfo bytes in `original_header`, which can be larger than the
    // regenerated canonical header.  Resolve the exact bounded length before
    // charging any clone or serialization allocation.
    archive
        .encoded_len(limits)
        .map_err(|_| Error::InvalidSource { path })
}

/// Rewrite a scalar control using the same strict list/storage primitives as
/// the Pop-Up Menu owner.  Popup-bearing transitions are deliberately left to
/// that owner so their model UUID/metadata lifecycle cannot be bypassed.
pub fn rewrite_scalar_control(
    source: &dyn TableSource,
    target: Target,
    before: Option<&CellControl>,
    after: Option<&CellControl>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeControlOutput, Error> {
    if target.locked {
        return Err(Error::Refused(crate::cell_format::Refusal::TableLocked {
            path,
        }));
    }
    if before.is_some_and(|value| matches!(value, CellControl::PopUpMenu(_)))
        || after.is_some_and(|value| matches!(value, CellControl::PopUpMenu(_)))
    {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }
    let model_component = source
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
        .ok_or(Error::Refused(crate::cell_format::Refusal::CellNotFound))?;
    if tile_identifier == target.model_identifier {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }
    let tile_component_index = resolved_component_index(source, tile_identifier, path)?;
    let tile_archive = source
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
    let cell = BncCell::parse(
        cell_source,
        &crate::cell_format::bnc_limits_for_source(cell_source),
    )
    .map_err(|_| Error::InvalidSource { path })?;
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
        .catalog()
        .get_index(format_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let control_component_index = resolved_component_index(source, control_table_identifier, path)?;
    let control_archive = source
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
        Some(CellControl::PopUpMenu(_)) => {
            return Err(Error::Refused(
                crate::cell_format::Refusal::UnsupportedDependency { path },
            ));
        },
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
        let limits = crate::cell_format::bnc_limits_for_source(cell_source);
        let mut cell =
            BncCell::parse(cell_source, &limits).map_err(|_| Error::InvalidSource { path })?;
        let (kind, _control) = native_kind(after, path)?;
        cell.set_data_format_identifier(format_key, kind, Some(control_key))
            .map_err(|_| {
                Error::Refused(crate::cell_format::Refusal::UnsupportedDependency { path })
            })?;
        cell.encode(&limits)
            .map_err(|_| Error::InvalidSource { path })?
    } else {
        let limits = crate::cell_format::bnc_limits_for_source(cell_source);
        let mut cell =
            BncCell::parse(cell_source, &limits).map_err(|_| Error::InvalidSource { path })?;
        cell.clear_explicit_format();
        cell.encode(&limits)
            .map_err(|_| Error::InvalidSource { path })?
    };
    let patched_tile = popup_native::patch_tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
        &replacement_cell,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let archive_limits = source
        .archive_limits()
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

    // `to_bytes` performs the same length calculation internally,
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
            .catalog()
            .get_index(*component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let source_length = source_archive
            .encoded_len(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let candidate_length = candidate
            .encoded_len(archive_limits)
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
            .encoded_len(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let archive_bytes = archive
            .to_bytes(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        if archive_bytes.len() != expected_archive_bytes {
            return Err(Error::Verification);
        }
        budget
            .charge_payload_bytes(archive_bytes.len(), path)
            .map_err(|_| Error::Limit {
                kind: crate::cell_format::LimitKind::PayloadBytes,
                observed: archive_bytes.len() as u64,
                maximum: u64::MAX,
                path,
            })?;
        let member_name = source
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
pub use super::display::{
    NumberFormatReadError, TextFormatReadError, read_number_format, read_number_format_with_budget,
    read_text_format, read_text_format_with_budget, rewrite_number_format, rewrite_text_format,
};

#[derive(Clone)]
pub struct ListEntry {
    pub key: u32,
    pub ref_count: u32,
    pub payload: Vec<u8>,
    pub is_format: bool,
}

pub struct ListFacts {
    pub next_key: u32,
    pub entries: Vec<ListEntry>,
}

/// Copy a borrowed native payload into transaction-owned staging memory.
///
/// The charge deliberately precedes `try_reserve_exact`: a hostile payload
/// must be rejected by the transaction ledger before the allocator is asked
/// to materialize it.  These bytes are retained until the surrounding native
/// operation has finished, rather than being counted as decoded wire input.
pub fn copy_payload_with_budget(
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

pub struct TileReferenceCollector<'budget> {
    pub tiles: Vec<(u32, u64)>,
    budget: &'budget mut TransactionBudget,
    path: Path,
    failure: Option<Error>,
}

impl<'budget> TileReferenceCollector<'budget> {
    pub fn new(budget: &'budget mut TransactionBudget, path: Path) -> Self {
        Self {
            tiles: Vec::new(),
            budget,
            path,
            failure: None,
        }
    }

    pub fn finish(self) -> Result<Vec<(u32, u64)>, Error> {
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
    number_secondaries: Vec<u32>,
    native_primary_formats: Vec<(u32, NativePrimaryFormatKind)>,
    controls: Vec<(u32, usize)>,
    invalid: bool,
    budget: &'budget mut TransactionBudget,
    path: Path,
    failure: Option<Error>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NativePrimaryFormatKind {
    Currency,
    Duration,
}

impl<'budget> BncReferenceCensus<'budget> {
    fn new(budget: &'budget mut TransactionBudget, path: Path) -> Self {
        Self {
            formats: Vec::new(),
            number_secondaries: Vec::new(),
            native_primary_formats: Vec::new(),
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

    fn record_number_secondary(&mut self, identifier: u32) -> Result<(), Error> {
        if self.number_secondaries.contains(&identifier) {
            return Ok(());
        }
        reserve_vec_slot(&mut self.number_secondaries, self.budget, self.path)?;
        self.number_secondaries.push(identifier);
        Ok(())
    }

    fn record_native_primary_format(
        &mut self,
        identifier: u32,
        kind: NativePrimaryFormatKind,
    ) -> Result<(), Error> {
        if self.native_primary_formats.contains(&(identifier, kind)) {
            return Ok(());
        }
        reserve_vec_slot(&mut self.native_primary_formats, self.budget, self.path)?;
        self.native_primary_formats.push((identifier, kind));
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
        let Ok(cell) =
            BncCellView::parse(source, &crate::cell_format::bnc_limits_for_source(source))
        else {
            return false;
        };
        if validate_bnc_format_metadata_kind(source, cell.cell_format_kind(), self.path).is_err() {
            return false;
        }
        let native_primary = match validate_native_cell_metadata(&cell, self.path) {
            Ok(value) => value,
            Err(_) => return false,
        };
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
            if let Err(error) = self.record_number_secondary(identifier) {
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
        } else if control.is_none()
            && cell.cell_format_kind() == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)
        {
            // Older wire versions do not expose Text's retained generic
            // Number edge through `secondary_format_identifier`; recover it
            // only in this private census route. A later wire helper may
            // expose it directly, in which case the branch above wins and
            // prevents double-counting.
            let Ok(identifier) =
                text_secondary_identifier_for_census_view(source, &cell, self.path)
            else {
                return false;
            };
            if let Some(identifier) = identifier {
                if let Err(error) = self.record_number_secondary(identifier) {
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
        if let Some((identifier, kind)) = native_primary {
            if let Err(error) = self.record_native_primary_format(identifier, kind) {
                self.failure = Some(error);
                return false;
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

    fn validate_number_format_target(
        &mut self,
        formats: &[ListEntry],
        identifier: u32,
    ) -> Result<(), Error> {
        let entry = formats
            .iter()
            .find(|entry| entry.key == identifier)
            .ok_or(Error::InvalidSource { path: self.path })?;
        // A secondary edge must point at a live format-list entry. `list_facts`
        // rejects zero refcounts globally, but keep this invariant local to
        // the edge validator as well so the target contract remains explicit.
        if !entry.is_format || entry.ref_count == 0 || entry.payload.is_empty() {
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
        Ok(())
    }

    fn validate_native_format_targets(&mut self, formats: &[ListEntry]) -> Result<(), Error> {
        // Converted Text cells and native Currency/Duration cells carry a
        // second edge in the shared generic identifier slot. Refcount
        // equality alone cannot prove that the target entry is the required
        // live native Number format; validate every distinct target after the
        // complete tile walk, including targets owned by nonselected cells.
        for index in 0..self.number_secondaries.len() {
            let identifier = self.number_secondaries[index];
            self.validate_number_format_target(formats, identifier)?;
        }
        // Currency and Duration primaries are family-specific strict payloads
        // as well. A refcount match must not make a Number/Fraction payload
        // usable as an unselected native Currency/Duration entry.
        for index in 0..self.native_primary_formats.len() {
            let (identifier, kind) = self.native_primary_formats[index];
            self.validate_native_primary_format_target(formats, identifier, kind)?;
        }
        Ok(())
    }

    fn validate_native_primary_format_target(
        &mut self,
        formats: &[ListEntry],
        identifier: u32,
        kind: NativePrimaryFormatKind,
    ) -> Result<(), Error> {
        let entry = formats
            .iter()
            .find(|entry| entry.key == identifier)
            .ok_or(Error::InvalidSource { path: self.path })?;
        if !entry.is_format || entry.ref_count == 0 || entry.payload.is_empty() {
            return Err(Error::InvalidSource { path: self.path });
        }
        let format_type = match kind {
            NativePrimaryFormatKind::Currency => {
                let (primary, report) = currency_format_codec::decode_currency_format_with_report(
                    &entry.payload,
                    control_codec_options(entry.payload.len(), self.budget),
                )
                .map_err(|error| map_control_error(error, self.path))?;
                charge_control_decode_report(self.budget, report, self.path)?;
                primary.format_type()
            },
            NativePrimaryFormatKind::Duration => {
                let (primary, report) = duration_format_codec::decode_duration_format_with_report(
                    &entry.payload,
                    control_codec_options(entry.payload.len(), self.budget),
                )
                .map_err(|error| map_control_error(error, self.path))?;
                charge_control_decode_report(self.budget, report, self.path)?;
                primary.format_type()
            },
        };
        let expected = match kind {
            NativePrimaryFormatKind::Currency => currency_format_codec::NATIVE_CURRENCY_FORMAT_TYPE,
            NativePrimaryFormatKind::Duration => duration_format_codec::NATIVE_DURATION_FORMAT_TYPE,
        };
        if format_type != expected {
            return Err(Error::InvalidSource { path: self.path });
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

pub fn list_facts(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListFacts, Error> {
    list_facts_with_input(source, budget, true, path)
}

pub fn list_facts_without_input(
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
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }
    // `next_list_id` is the source-authoritative allocation watermark.  A
    // stale/zero cursor would make deterministic COW allocation ambiguous and
    // could cause a new payload to collide with an existing key. Empty lists
    // still carry a positive cursor in valid native sources.
    let maximum_key = entries.last().map_or(0, |entry| entry.key);
    if list.next_list_id() == 0 || list.next_list_id() <= maximum_key {
        return Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        ));
    }
    Ok(ListFacts {
        next_key: list.next_list_id(),
        entries,
    })
}

pub fn mutate_list(
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
            let mut candidate = if budget.list_keys == super::ListKeys::Append {
                facts.next_key.max(1)
            } else {
                1u32
            };
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
                return Err(Error::Refused(
                    crate::cell_format::Refusal::UnsupportedDependency { path },
                ));
            }
            let advance = if candidate == facts.next_key {
                Some((
                    facts.next_key,
                    facts.next_key.checked_add(1).ok_or(Error::Refused(
                        crate::cell_format::Refusal::UnsupportedDependency { path },
                    ))?,
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
    let mut released = false;
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
            released = entry.ref_count == 1;
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
    if released && budget.list_keys == super::ListKeys::Append {
        if let Some(old_key) = old_key {
            output = hand_back_last_key(output, old_key, budget, path)?;
        }
    }
    Ok((output, desired_key))
}

/// Lower `next_list_id` to `key` when `key` was the last allocated key and
/// no entry holds it any more.
fn hand_back_last_key(
    list: Vec<u8>,
    key: u32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let facts = list_facts_without_input(&list, budget, path)?;
    if facts.next_key != key.saturating_add(1) || facts.entries.iter().any(|entry| entry.key == key)
    {
        return Ok(list);
    }
    budget.charge_scratch_bytes(list.len(), path)?;
    litchi_iwa_common::wire::patch_varint_field(&list, 2, true, Some(u64::from(key)))
        .map_err(|_| Error::InvalidSource { path })
}

pub fn apply_list_mutation_with_budget(
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

pub fn charge_storage_report(
    budget: &mut TransactionBudget,
    report: storage_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_bytes(report.source_bytes(), path)?;
    charge_storage_report_without_input(budget, report, path)
}

pub fn charge_storage_report_without_input(
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

pub fn map_storage_error(error: storage_codec::DecodeError, path: Path) -> Error {
    match error.resource_limit() {
        Some(storage_codec::DecodeLimit::Bytes { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::References { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::PayloadReferences,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Fields { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Work { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Nesting { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Allocation { requested }) => Error::Allocation {
            amount: requested,
            path,
        },
        Some(storage_codec::DecodeLimit::Retained { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Text { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireTextBytes,
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
            return Err(Error::Refused(
                crate::cell_format::Refusal::UnsupportedDependency { path },
            ));
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

pub fn control_codec_options(
    source_len: usize,
    budget: &TransactionBudget,
) -> control_codec::DecodeOptions {
    budget.residual_control_options_for_len(source_len)
}

pub fn charge_control_requirements(
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

pub fn charge_control_decode_report(
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
pub fn require_exclusive_selected_model_reference(
    source: &dyn TableSource,
    target: Target,
    identifier: u64,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    match source.proofs() {
        Some(proofs) => proofs
            .census(source, budget, path)?
            .prove_exclusive_after_preflight(target, identifier, path, budget),
        None => {
            // Keep the non-memoized Pages/Keynote route under the same
            // admission ordering as the Numbers proof cache. The census is
            // built only after all work/allocation charges have succeeded.
            preflight_reference_census(source, None, true, budget, path)?;
            ReferenceCensus::build(source, None)
                .prove_exclusive_after_preflight(target, identifier, path, budget)
        },
    }
}

/// The package-level proofs of the cell-format engines that a leaf memoizes
/// for one immutable package (record 0370k): the object-reference census
/// and the metadata registry inspections.
///
/// Each proof is a function of the package's components and limits, so a
/// memoized proof charges and refuses exactly as a fresh one. A package
/// reopened from an edited reassembly of its parent ([`Self::reopened`])
/// keeps the census of every component whose parsed archive it shares with
/// the parent and inspects only the others.
#[derive(Debug, Default)]
pub struct PackageProofs {
    census: OnceLock<Arc<ReferenceCensus>>,
    inherited: Option<InheritedCensus>,
    strict_registry: OnceLock<MemoizedRegistry>,
    cross_component_read_registry: OnceLock<MemoizedRegistry>,
}

impl PackageProofs {
    /// The proofs of `candidate`, a package reopened from an edited
    /// reassembly of `parent`, the package these proofs belong to. A
    /// candidate component whose parsed archive is one of the parent's own
    /// (a reopen shares byte-identical members) inherits that archive's
    /// census.
    #[must_use]
    pub fn reopened(&self, parent: &ComponentCatalog, candidate: &ComponentCatalog) -> Self {
        // Keep the parent census itself rather than allocating a pointer
        // index and a per-candidate Vec here. Candidate component matching is
        // performed after the transaction has charged the bounded census
        // build, so reopen handoff is allocation-free apart from the existing
        // Arc reference count.
        let _ = (parent, candidate);
        let inherited = self
            .census
            .get()
            .cloned()
            .map(|census| InheritedCensus { census });
        Self {
            inherited,
            ..Self::default()
        }
    }

    fn census(
        &self,
        source: &dyn TableSource,
        budget: &mut TransactionBudget,
        path: Path,
    ) -> Result<&ReferenceCensus, Error> {
        let build = self.census.get().is_none();
        preflight_reference_census(source, self.inherited.as_ref(), build, budget, path)?;
        Ok(self
            .census
            .get_or_init(|| Arc::new(ReferenceCensus::build(source, self.inherited.as_ref())))
            .as_ref())
    }

    /// [`popup_metadata::inspect_cross_component_read`], memoized for
    /// `options`.
    pub(crate) fn cross_component_read_registry<'source>(
        &self,
        source: &'source dyn TableSource,
        options: MetadataRewriteOptions,
    ) -> MetadataResult<'source> {
        if let Some(memoized) = self
            .cross_component_read_registry
            .get()
            .filter(|memoized| memoized.options == options)
        {
            return memoized.facts(source);
        }
        let facts = popup_metadata::inspect_cross_component_read(source, options)?;
        let _ = self
            .cross_component_read_registry
            .set(MemoizedRegistry::new(options, &facts));
        Ok(facts)
    }

    /// [`popup_metadata::inspect`], memoized for `options`. The memoized
    /// cross-component read registry stands in for a fresh inspection when
    /// it proves that inspection would succeed with the same facts and
    /// report ([`MemoizedRegistry::admits_strict`]).
    fn strict_registry<'source>(
        &self,
        source: &'source dyn TableSource,
        options: MetadataRewriteOptions,
    ) -> MetadataResult<'source> {
        if let Some(memoized) = self
            .strict_registry
            .get()
            .filter(|memoized| memoized.options == options)
        {
            return memoized.facts(source);
        }
        let facts = match self.cross_component_read_registry.get() {
            Some(read) => {
                let facts = read.facts(source)?;
                if read.admits_strict(options, facts.payload().len()) {
                    facts
                } else {
                    popup_metadata::inspect(source, options)?
                }
            },
            None => popup_metadata::inspect(source, options)?,
        };
        let _ = self
            .strict_registry
            .set(MemoizedRegistry::new(options, &facts));
        Ok(facts)
    }
}

type MetadataResult<'source> =
    Result<popup_metadata::RegistryFacts<'source>, popup_metadata::MetadataError>;

/// One memoized registry inspection and the options it ran under.
#[derive(Debug)]
struct MemoizedRegistry {
    options: MetadataRewriteOptions,
    registry: Arc<popup_metadata::Registry>,
}

impl MemoizedRegistry {
    fn new(options: MetadataRewriteOptions, facts: &popup_metadata::RegistryFacts<'_>) -> Self {
        Self {
            options,
            registry: Arc::clone(facts.shared()),
        }
    }

    fn facts<'source>(&self, source: &'source dyn TableSource) -> MetadataResult<'source> {
        Ok(popup_metadata::RegistryFacts::from_shared(
            popup_metadata::strict_source(source)?,
            Arc::clone(&self.registry),
        ))
    }

    /// Whether the strict inspection under `options` of the
    /// `payload_bytes`-byte metadata payload this cross-component read
    /// inspected succeeds with the same facts and report.
    ///
    /// Both inspections run the same two codec passes and visitor; the strict
    /// one also refuses a current component without a physical member. The
    /// codec gates an inspection only on its options' absolute ceiling, the
    /// payload size, its running counters (whose final values the report
    /// holds) and the generated decoder's message-size, unknown-field and
    /// recursion limits.
    fn admits_strict(&self, options: MetadataRewriteOptions, payload_bytes: usize) -> bool {
        let read = self.options;
        let report = self.registry.report();
        !self.registry.has_unmapped_current_component()
            // The read's options passed the absolute ceiling.
            && options.max_input_bytes() <= read.max_input_bytes()
            && options.max_output_bytes() <= read.max_output_bytes()
            && options.recursion_limit() == read.recursion_limit()
            && payload_bytes <= options.max_input_bytes()
            // Every unknown field takes at least one payload byte.
            && payload_bytes <= options.max_fields()
            && report.fields() <= options.max_fields()
            && report.work_bytes() <= options.max_work_bytes()
            && report.max_depth() <= options.recursion_limit()
            && report.components_scanned() <= options.max_components()
            && report.references_scanned() <= options.max_references()
    }
}

/// Every object's strict reference inspection
/// (`ArchiveReferencePolicy::RejectUnknownMetadata`) by catalog component,
/// replayed by each exclusivity proof in catalog order, so a proof charges
/// and refuses exactly as inspecting each object in turn does.
#[derive(Debug)]
struct ReferenceCensus {
    /// The effective archive limits of the inspections; `None` when the
    /// package's limits admit none.
    limits: Option<ArchiveLimits>,
    components: Vec<CensusComponent>,
}

/// The census of the components a reopened package shares with its parent,
/// by the reopened package's catalog position.
#[derive(Debug)]
struct InheritedCensus {
    census: Arc<ReferenceCensus>,
}

/// One component's stable parsed-archive identity and its memoized object
/// inspections. The identity is stored as an address-sized integer so the
/// proof cache remains `Send`/`Sync` without retaining a borrow into a
/// catalog owned by another package snapshot.
#[derive(Debug)]
struct CensusComponent {
    archive_key: usize,
    objects: Arc<[ObjectCensus]>,
}

/// One object's header size and strict reference inspection.
#[derive(Debug)]
struct ObjectCensus {
    identifier: Option<u64>,
    /// The occurrence count and the object references (referenced
    /// identifier, message index); `None` when the inspection refuses.
    inspection: Option<(usize, Vec<(u64, usize)>)>,
}

impl ReferenceCensus {
    /// Inspect every object of `source`, reusing the inherited components
    /// inspected under the same limits.
    fn build(source: &dyn TableSource, inherited: Option<&InheritedCensus>) -> Self {
        let Ok(limits) = source.archive_limits().effective_archive_limits() else {
            return Self {
                limits: None,
                components: Vec::new(),
            };
        };
        let inherited = inherited.filter(|inherited| inherited.census.limits == Some(limits));
        let components = source
            .catalog()
            .iter()
            .map(|component| {
                let archive_key = archive_key(component.archive());
                let objects = inherited
                    .and_then(|inherited| {
                        inherited
                            .census
                            .components
                            .iter()
                            .find(|component| component.archive_key == archive_key)
                            .map(|component| Arc::clone(&component.objects))
                    })
                    .unwrap_or_else(|| inspect_objects(component.archive(), limits));
                CensusComponent {
                    archive_key,
                    objects,
                }
            })
            .collect();
        Self {
            limits: Some(limits),
            components,
        }
    }

    /// The exclusivity proof after `preflight_reference_census` has admitted
    /// all census work and allocations. The selected model message is the
    /// only owner of `identifier`, and every object's header is admitted.
    fn prove_exclusive_after_preflight(
        &self,
        target: Target,
        identifier: u64,
        path: Path,
        budget: &mut TransactionBudget,
    ) -> Result<(), Error> {
        let refused =
            || Error::Refused(crate::cell_format::Refusal::UnsupportedDependency { path });
        if self.limits.is_none() {
            return Err(Error::InvalidSource { path });
        }
        let mut selected = 0usize;
        let mut inspected = 0usize;
        for (component_index, component) in self.components.iter().enumerate() {
            for object in component.objects.iter() {
                let object_identifier = object.identifier.ok_or(Error::InvalidSource { path })?;
                let allowed = component_index == target.component_index
                    && object_identifier == target.model_identifier;
                let (occurrences, references) = object.inspection.as_ref().ok_or_else(refused)?;
                let mut matches = 0usize;
                let mut unexpected = false;
                for &(referenced, message_index) in references {
                    if referenced == identifier {
                        matches = matches.saturating_add(1);
                        unexpected |= !allowed || message_index != target.message_index;
                    }
                }
                inspected = inspected
                    .checked_add(*occurrences)
                    .ok_or(Error::InvalidSource { path })?;
                budget.charge_payload_references(*occurrences, path)?;
                if unexpected {
                    return Err(refused());
                }
                selected = selected
                    .checked_add(matches)
                    .ok_or(Error::InvalidSource { path })?;
            }
        }
        // The selected header contains one required aggregate edge and may
        // mirror it once in a typed FieldInfo.
        // `validate_selected_model_reference` already proves that exact shape;
        // this package-wide census proves that no other message or opaque
        // metadata owns the mutable object.
        if selected == 0 || inspected == 0 {
            return Err(refused());
        }
        Ok(())
    }
}

/// Charge every operation-local resource that the reference census may use
/// before any archive header is inspected or census storage is allocated.
///
/// The previous memoized implementation built the complete census first and
/// charged the same work only while replaying it. That made a cold, tiny-
/// budget refusal perform an unadmitted package-wide inspection. This pass is
/// intentionally structural: it reads only already-parsed object metadata,
/// charges the complete cold-build bound up front, and leaves the strict
/// `RejectUnknownMetadata` inspection to the admitted build below.
fn preflight_reference_census(
    source: &dyn TableSource,
    inherited: Option<&InheritedCensus>,
    build: bool,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let limits = source
        .archive_limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let inherited = inherited.filter(|inherited| inherited.census.limits == Some(limits));
    let mut objects = 0usize;
    let mut work = 0usize;
    let mut changed_components = 0usize;
    let mut changed_reference_objects = 0usize;
    for component in source.catalog().iter() {
        let archive_key = archive_key(component.archive());
        let shared = inherited.and_then(|inherited| {
            inherited
                .census
                .components
                .iter()
                .find(|component| component.archive_key == archive_key)
        });
        if shared.is_none() {
            changed_components = changed_components.saturating_add(1);
        }
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource { path })?;
            let _ = identifier;
            let header_bytes = usize::try_from(object.header_length)
                .ok()
                .filter(|length| *length != 0)
                .ok_or(Error::InvalidSource { path })?;
            objects = objects.saturating_add(1);
            work = work.saturating_add(header_bytes.saturating_mul(4));
            if shared.is_none() && object_reference_capacity(object) != 0 {
                changed_reference_objects = changed_reference_objects.saturating_add(1);
            }
        }
    }

    // One logical allocation per replayed object is the existing proof
    // charge. A cold build additionally retains one object slice per changed
    // component, the catalog slice and its Arc wrapper; collector vectors are
    // reserved once for objects carrying references.
    let mut allocations = objects;
    if build {
        allocations = allocations
            .saturating_add(changed_components)
            .saturating_add(changed_reference_objects)
            .saturating_add(2);
    }
    budget.charge_allocations(allocations, path)?;
    budget.charge_transaction_work(work, path)
}

fn archive_key(archive: &Archive) -> usize {
    core::ptr::from_ref(archive) as usize
}

fn object_reference_capacity(object: &ArchiveObject) -> usize {
    object
        .archive_info
        .message_infos
        .iter()
        .map(|message| {
            message.object_references.len().saturating_add(
                message
                    .field_infos
                    .iter()
                    .map(|field| field.object_references.len())
                    .sum(),
            )
        })
        .sum()
}

/// Inspect every object of one component archive.
fn inspect_objects(archive: &Archive, limits: ArchiveLimits) -> Arc<[ObjectCensus]> {
    struct Collector {
        references: Vec<(u64, usize)>,
    }

    impl ArchiveReferenceVisitor for Collector {
        fn visit_reference(
            &mut self,
            occurrence: ArchiveReferenceOccurrence,
        ) -> litchi_iwa_core::Result<()> {
            if occurrence.kind == ArchiveReferenceKind::Object {
                self.references
                    .push((occurrence.referenced_identifier, occurrence.message_index));
            }
            Ok(())
        }
    }

    archive
        .objects
        .iter()
        .map(|object| {
            let mut collector = Collector {
                references: Vec::with_capacity(object_reference_capacity(object)),
            };
            let inspection = object
                .inspect_references_with_policy(
                    &mut collector,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .ok()
                .map(|occurrences| (occurrences, collector.references));
            ObjectCensus {
                identifier: object.archive_info.identifier,
                inspection,
            }
        })
        .collect()
}

pub fn validate_format_refcounts(
    source: &dyn TableSource,
    tile_references: &[(u32, u64)],
    formats: &[ListEntry],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let mut census = BncReferenceCensus::new(budget, path);
    for (_, tile_identifier) in tile_references {
        let component_index = resolved_component_index(source, *tile_identifier, path)?;
        let archive = source
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
    census.validate_native_format_targets(formats)?;
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

pub fn verify_control_report(
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

pub fn map_control_error(error: control_codec::DecodeError, path: Path) -> Error {
    match error.resource_limit() {
        Some(control_codec::DecodeLimit::InputBytes { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::OutputBytes { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireOutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Fields { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Work { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Nesting { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::References { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::PayloadReferences,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Items { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::PayloadItems,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Text { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::WireTextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Allocation { requested }) => Error::Allocation {
            amount: requested,
            path,
        },
        Some(control_codec::DecodeLimit::Retained { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Scratch { observed, maximum }) => Error::Limit {
            kind: crate::cell_format::LimitKind::ScratchBytes,
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
        CellControl::PopUpMenu(_) => Err(Error::Refused(
            crate::cell_format::Refusal::UnsupportedDependency { path },
        )),
    }
}

pub fn validate_metadata_ownership<'source>(
    source: &'source dyn TableSource,
    component_indices: &[usize],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<popup_metadata::RegistryFacts<'source>, Error> {
    let catalog = super::physical_source(source).map_err(|_| Error::InvalidSource { path })?;
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
    let facts = match source.proofs() {
        Some(proofs) => proofs.strict_registry(source, options),
        None => popup_metadata::inspect(source, options),
    }
    .map_err(|_| Error::InvalidSource { path })?;
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

pub fn validate_cross_component_reference(
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
        .map_err(|_| Error::Refused(crate::cell_format::Refusal::UnsupportedDependency { path }))
}

pub fn unique_object(
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

pub fn unique_message_index(
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

pub fn archive_for_mut(
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

pub fn resolved_component_index(
    source: &dyn TableSource,
    identifier: u64,
    path: Path,
) -> Result<usize, Error> {
    source
        .resolve_object(identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .map(|resolved| resolved.component_index)
        .ok_or(Error::InvalidSource { path })
}

pub struct ListMessage<'a> {
    pub message_index: usize,
    pub payload: &'a [u8],
}

pub fn unique_list_message<'a>(
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
    source: &dyn TableSource,
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
    census.validate_native_format_targets(formats)?;

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
pub fn unsupported_scalar_control() -> NativePopUpError {
    NativePopUpError::UnsupportedDependency
}

#[cfg(test)]
mod census_tests {
    use super::*;
    use litchi_iwa_archive::unstable::EntryEdit;
    use litchi_iwa_core::RawMessage;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    /// One IWA object whose header carries `extra` header bytes.
    fn object(identifier: u64, reference: Option<u64>, extra: &[u8]) -> Vec<u8> {
        let payload = [0x08, 0x01];
        let mut message_info = Vec::new();
        varint(&mut message_info, 1 << 3);
        varint(&mut message_info, 1);
        varint(&mut message_info, 3 << 3);
        varint(&mut message_info, payload.len() as u64);
        if let Some(reference) = reference {
            let mut packed = Vec::new();
            varint(&mut packed, reference);
            varint(&mut message_info, (5 << 3) | 2);
            varint(&mut message_info, packed.len() as u64);
            message_info.extend_from_slice(&packed);
        }
        let mut header = Vec::new();
        varint(&mut header, 1 << 3);
        varint(&mut header, identifier);
        varint(&mut header, (2 << 3) | 2);
        varint(&mut header, message_info.len() as u64);
        header.extend_from_slice(&message_info);
        header.extend_from_slice(extra);
        let mut output = Vec::new();
        varint(&mut output, header.len() as u64);
        output.extend_from_slice(&header);
        output.extend_from_slice(&payload);
        output
    }

    fn census_of(bytes: &[u8]) -> ReferenceCensus {
        let limits = ArchiveLimits::default();
        let archive = Archive::parse(bytes, limits).expect("archive");
        ReferenceCensus {
            limits: Some(limits),
            components: vec![CensusComponent {
                archive_key: archive_key(&archive),
                objects: inspect_objects(&archive, limits),
            }],
        }
    }

    fn target() -> Target {
        Target {
            scope: 0,
            table: 0,
            position: litchi_iwa_common::table::coordinate::At::new(0, 0),
            model_identifier: 10,
            component_index: 0,
            object_index: 0,
            message_index: 0,
            message_type: 6_001,
            locked: false,
        }
    }

    fn prove(bytes: &[u8]) -> Result<(), Error> {
        let source = Memoized::open(&fixture("numbers/basic.numbers"));
        let mut budget = TransactionBudget::for_cell_control(&source);
        census_of(bytes).prove_exclusive_after_preflight(target(), 20, Path::Package, &mut budget)
    }

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork")
                .join(name),
        )
        .expect("fixture bytes")
    }

    /// A package opened from bytes that memoizes its proofs, as a leaf does.
    struct Memoized {
        catalog: litchi_iwa_archive::SourceCatalog,
        proofs: PackageProofs,
    }

    impl Memoized {
        fn open(bytes: &[u8]) -> Self {
            Self {
                catalog: litchi_iwa_archive::SourceCatalog::from_bytes(
                    bytes,
                    litchi_iwa_archive::Limits::default(),
                )
                .expect("catalog"),
                proofs: PackageProofs::default(),
            }
        }

        /// `bytes` reopened as an edited reassembly of this package.
        fn reopen(&self, bytes: &[u8]) -> Self {
            let catalog = self.catalog.reopen(Arc::from(bytes)).expect("reopen");
            let proofs = self
                .proofs
                .reopened(self.catalog.components(), catalog.components());
            Self { catalog, proofs }
        }
    }

    impl TableSource for Memoized {
        fn catalog(&self) -> &ComponentCatalog {
            self.catalog.components()
        }

        fn physical(&self) -> Option<&litchi_iwa_archive::SourceCatalog> {
            Some(&self.catalog)
        }

        fn resolve_object(
            &self,
            _identifier: u64,
        ) -> Result<Option<crate::ResolvedObject<'_>>, crate::ResolveError> {
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            self.catalog.source_bytes()
        }

        fn archive_limits(&self) -> litchi_iwa_archive::Limits {
            litchi_iwa_archive::Limits::default()
        }

        fn max_objects(&self) -> usize {
            1_000_000
        }

        fn max_references(&self) -> usize {
            1_000_000
        }

        fn proofs(&self) -> Option<&PackageProofs> {
            Some(&self.proofs)
        }
    }

    #[test]
    fn the_cached_census_proves_an_exclusive_reference() {
        let mut bytes = object(10, Some(20), &[]);
        bytes.extend(object(20, None, &[]));
        bytes.extend(object(30, None, &[]));
        assert!(prove(&bytes).is_ok());
    }

    #[test]
    fn unknown_metadata_on_an_unrelated_object_still_refuses() {
        let mut bytes = object(10, Some(20), &[]);
        bytes.extend(object(20, None, &[]));
        // Field 9 of an ArchiveInfo is unknown header metadata.
        bytes.extend(object(30, None, &[9 << 3, 1]));
        assert!(matches!(
            prove(&bytes),
            Err(Error::Refused(
                crate::cell_format::Refusal::UnsupportedDependency { .. }
            ))
        ));
    }

    #[test]
    fn a_second_owner_still_refuses() {
        let mut bytes = object(10, Some(20), &[]);
        bytes.extend(object(20, None, &[]));
        bytes.extend(object(30, Some(20), &[]));
        assert!(prove(&bytes).is_err());
    }

    #[test]
    fn a_reopened_package_inherits_the_census_of_shared_archives_only() {
        let bytes = fixture("numbers/basic.numbers");
        let parent = Memoized::open(&bytes);
        let mut budget = TransactionBudget::for_cell_control(&parent);
        let census = parent
            .proofs
            .census(&parent, &mut budget, Path::Package)
            .expect("census");
        let fresh = ReferenceCensus::build(&parent, None);
        assert_eq!(format!("{census:?}"), format!("{fresh:?}"));

        // Reopening identical bytes shares every parsed archive.
        let reopened = parent.reopen(&bytes);
        let mut budget = TransactionBudget::for_cell_control(&reopened);
        let inherited = reopened
            .proofs
            .census(&reopened, &mut budget, Path::Package)
            .expect("inherited census");
        assert_eq!(inherited.components.len(), census.components.len());
        assert!(
            inherited
                .components
                .iter()
                .zip(&census.components)
                .all(|(left, right)| Arc::ptr_eq(&left.objects, &right.objects))
        );

        // A package parsed afresh shares none and is inspected in full.
        let mut unrelated = Memoized::open(&bytes);
        unrelated.proofs = parent
            .proofs
            .reopened(parent.catalog(), unrelated.catalog());
        let mut budget = TransactionBudget::for_cell_control(&unrelated);
        let rebuilt = unrelated
            .proofs
            .census(&unrelated, &mut budget, Path::Package)
            .expect("rebuilt census");
        assert!(
            rebuilt
                .components
                .iter()
                .zip(&census.components)
                .all(|(left, right)| !Arc::ptr_eq(&left.objects, &right.objects))
        );
        assert!(
            rebuilt
                .components
                .iter()
                .zip(&census.components)
                .all(
                    |(left, right)| format!("{:?}", left.objects) == format!("{:?}", right.objects)
                )
        );
    }

    #[test]
    fn a_low_budget_refuses_before_cold_or_cached_census_work() {
        let bytes = fixture("numbers/basic.numbers");
        let cold = Memoized::open(&bytes);
        let mut cold_budget = TransactionBudget::for_cell_control(&cold);
        cold_budget.remaining_allocations = 0;
        cold_budget.remaining_transaction_work = 0;
        assert!(matches!(
            require_exclusive_selected_model_reference(
                &cold,
                target(),
                20,
                Path::Package,
                &mut cold_budget,
            ),
            Err(Error::Limit { .. })
        ));
        assert!(cold.proofs.census.get().is_none());

        let cached = Memoized::open(&bytes);
        let mut warm_budget = TransactionBudget::for_cell_control(&cached);
        let _ = require_exclusive_selected_model_reference(
            &cached,
            target(),
            20,
            Path::Package,
            &mut warm_budget,
        );
        assert!(cached.proofs.census.get().is_some());

        let mut cached_budget = TransactionBudget::for_cell_control(&cached);
        cached_budget.remaining_allocations = 0;
        cached_budget.remaining_transaction_work = 0;
        assert!(matches!(
            require_exclusive_selected_model_reference(
                &cached,
                target(),
                20,
                Path::Package,
                &mut cached_budget,
            ),
            Err(Error::Limit { .. })
        ));
        assert!(cached.proofs.census.get().is_some());
    }

    #[test]
    fn a_changed_archive_is_reinspected_while_unchanged_archives_reuse_census() {
        let bytes = fixture("numbers/basic.numbers");
        let parent = Memoized::open(&bytes);
        let mut parent_budget = TransactionBudget::for_cell_control(&parent);
        let _ = require_exclusive_selected_model_reference(
            &parent,
            target(),
            20,
            Path::Package,
            &mut parent_budget,
        );
        let parent_census = Arc::clone(parent.proofs.census.get().expect("parent census"));

        let entry = parent
            .catalog
            .package()
            .iter()
            .find(|entry| entry.name() == "Index/Tables/DataList.iwa")
            .expect("DataList member");
        let archive_limits = parent
            .catalog
            .limits()
            .effective_archive_limits()
            .expect("archive limits");
        let snappy_limits = parent
            .catalog
            .limits()
            .snappy_limits()
            .expect("Snappy limits");
        let decompressed =
            SnappyStream::decompress(entry.data(), snappy_limits).expect("decompress DataList");
        let mut archive =
            Archive::parse(decompressed.as_bytes(), archive_limits).expect("DataList archive");
        let message = archive.objects[0].messages[0].clone();
        let mut data = message.data;
        data.push(0);
        archive.objects[0]
            .replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
                archive_limits,
            )
            .expect("mutate DataList payload");
        let changed = archive.to_bytes(archive_limits).expect("encode DataList");
        let changed = SnappyStream::compress(&changed, snappy_limits).expect("compress DataList");
        let candidate_bytes = parent
            .catalog
            .package()
            .reassemble_with_changes_to_bytes(
                &[],
                &[EntryEdit::new(entry.name(), &changed)],
                &[],
                parent.catalog.limits(),
            )
            .expect("reassemble candidate");
        let candidate = parent.reopen(&candidate_bytes);
        let mut candidate_budget = TransactionBudget::for_cell_control(&candidate);
        let candidate_census = candidate
            .proofs
            .census(&candidate, &mut candidate_budget, Path::Package)
            .expect("candidate census");

        let mut reused = 0usize;
        let mut rebuilt = 0usize;
        for component in &candidate_census.components {
            if let Some(parent_component) = parent_census
                .components
                .iter()
                .find(|parent_component| parent_component.archive_key == component.archive_key)
            {
                assert!(Arc::ptr_eq(&parent_component.objects, &component.objects));
                reused = reused.saturating_add(1);
            } else {
                rebuilt = rebuilt.saturating_add(1);
            }
        }
        assert!(
            reused > 0,
            "candidate should inherit unchanged component proofs"
        );
        assert_eq!(
            rebuilt, 1,
            "only the edited component should be re-inspected"
        );
    }

    /// The options of a cross-component read and of the strict census, as
    /// `pop_up_menu::prove_cross_component_reference` and
    /// `validate_metadata_ownership` size them.
    fn registry_options(package: &Memoized) -> (MetadataRewriteOptions, MetadataRewriteOptions) {
        let sized = |bytes: usize, components: usize| {
            MetadataRewriteOptions::new(
                bytes,
                bytes * 2,
                bytes * 16,
                bytes * 64,
                64,
                components,
                bytes * 2,
                1,
            )
        };
        let source = package.source_bytes().len();
        let members = package
            .catalog
            .package()
            .iter()
            .map(|entry| entry.data().len())
            .sum::<usize>();
        (sized(source, source * 2), sized(members, members))
    }

    #[test]
    fn a_memoized_cross_component_read_stands_in_for_the_strict_registry() {
        for name in [
            "numbers/basic.numbers",
            "numbers/custom-retirement-cleared.numbers",
        ] {
            let package = Memoized::open(&fixture(name));
            let (read_options, strict_options) = registry_options(&package);
            let read = package
                .proofs
                .cross_component_read_registry(&package, read_options)
                .expect("read registry");
            let strict = package
                .proofs
                .strict_registry(&package, strict_options)
                .expect("strict registry");
            assert!(Arc::ptr_eq(read.shared(), strict.shared()), "{name}");
            let fresh = popup_metadata::inspect(&package, strict_options).expect("fresh registry");
            assert_eq!(format!("{:?}", *strict), format!("{:?}", *fresh), "{name}");
            assert_eq!(strict.payload(), fresh.payload());
            let again = package
                .proofs
                .strict_registry(&package, strict_options)
                .expect("memoized registry");
            assert!(Arc::ptr_eq(again.shared(), strict.shared()));
        }
    }

    #[test]
    fn a_cross_component_read_outside_the_strict_limits_does_not_stand_in() {
        let package = Memoized::open(&fixture("numbers/basic.numbers"));
        let (read_options, _) = registry_options(&package);
        package
            .proofs
            .cross_component_read_registry(&package, read_options)
            .expect("read registry");
        let fields = 1;
        let strict_options = MetadataRewriteOptions::new(
            read_options.max_input_bytes(),
            read_options.max_output_bytes(),
            fields,
            read_options.max_work_bytes(),
            64,
            read_options.max_components(),
            read_options.max_references(),
            1,
        );
        assert!(popup_metadata::inspect(&package, strict_options).is_err());
        assert!(
            package
                .proofs
                .strict_registry(&package, strict_options)
                .is_err()
        );
        assert!(package.proofs.strict_registry.get().is_none());
    }
}
