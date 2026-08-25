//! Private native seam for the unified cell-control owner.
//!
//! The selector/transaction facade lives in [`super::table_cell_control`].
//! This module intentionally contains no public archive or generated types;
//! scalar-control graph surgery will be added here as the source-preserving
//! list transition is shared with the audited Pop-Up Menu engine.

use litchi_iwa_core::{Archive, ArchiveObject, SnappyStream};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
    package_metadata_codec::RewriteOptions as MetadataRewriteOptions,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};

use super::table_cell_pop_up_menu_native::NativePopUpError;
use super::{
    Package,
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

const NATIVE_AUTOMATIC_DECIMAL_PLACES: u32 = 253;

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
pub(super) struct NativeControlOutput {
    pub(super) member_name: String,
    pub(super) archive_bytes: Vec<u8>,
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
    budget.charge_archive(archive, member.data().len(), path)?;

    let archive_bytes = archive_serialized_bound(archive)
        .saturating_add(member.data().len())
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
        .map(|object| object.messages.len())
        .sum::<usize>()
        .saturating_add(6);
    budget.charge_payload_messages(messages, path)?;
    budget.charge_payload_items(messages.saturating_add(archive.objects.len()), path)?;
    Ok(NativePreflight {
        archive_bytes,
        compressed_bytes,
    })
}

fn archive_serialized_bound(archive: &Archive) -> usize {
    let mut bound = 64usize;
    for object in &archive.objects {
        bound = bound
            .saturating_add(64)
            .saturating_add(object.messages.len().saturating_mul(32));
        for message in &object.messages {
            bound = bound.saturating_add(message.data.len());
        }
        for info in &object.archive_info.message_infos {
            bound = bound
                .saturating_add(32)
                .saturating_add(info.object_references.len().saturating_mul(16));
            bound = bound.saturating_add(
                info.field_infos
                    .iter()
                    .map(|field| field.object_references.len().saturating_mul(16) + 32)
                    .sum::<usize>(),
            );
        }
    }
    bound
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
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let mut archive = component.archive().clone();
    let model_object = unique_object(&archive, target.model_identifier, path)?;
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
    let mut tile_references = TileReferenceCollector::default();
    let (tile_storage, tile_report) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        budget.residual_storage_options(store.tiles()),
        &mut tile_references,
    )
    .map_err(|error| map_storage_error(error, path))?;
    charge_storage_report(budget, tile_report, path)?;
    let tile_id = target.position.row() / tile_storage.tile_size().unwrap_or(1).max(1);
    if tile_references
        .tiles
        .iter()
        .any(|(_, reference)| *reference == 0)
        || tile_references
            .tiles
            .iter()
            .enumerate()
            .any(|(index, (id, reference))| {
                tile_references.tiles[index + 1..]
                    .iter()
                    .any(|(other_id, other_reference)| {
                        id == other_id || reference == other_reference
                    })
            })
    {
        return Err(Error::InvalidSource { path });
    }
    let tile_identifier = tile_references
        .tiles
        .iter()
        .find(|(id, _)| *id == tile_id)
        .map(|(_, reference)| *reference)
        .ok_or(Error::CellNotFound)?;
    if tile_identifier == target.model_identifier {
        return Err(Error::UnsupportedDependency { path });
    }
    let tile_object = unique_object(&archive, tile_identifier, path)?;
    let tile_message_index = unique_message_index(tile_object, 6_002, path)?;
    let tile_payload = tile_object.messages[tile_message_index].data.clone();
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
    let format_object = unique_object(&archive, format_table_identifier, path)?;
    let control_object = unique_object(&archive, control_table_identifier, path)?;
    let format_message = unique_list_message(format_object, 2, budget, path)?;
    let control_message = unique_list_message(control_object, 12, budget, path)?;
    let format_message_index = format_message.message_index;
    let control_message_index = control_message.message_index;
    let format_payload = format_message.payload.to_owned();
    let control_payload = control_message.payload.to_owned();
    let format_facts = list_facts(&format_payload, budget, path)?;
    let control_facts = list_facts(&control_payload, budget, path)?;
    validate_refcounts(
        &archive,
        &tile_references.tiles,
        format_facts.entries.as_slice(),
        control_facts.entries.as_slice(),
        budget,
        path,
    )?;
    let control_source_len = format_payload.len().max(control_payload.len());
    let mut owned_identifiers = vec![
        target.model_identifier,
        tile_identifier,
        format_table_identifier,
        control_table_identifier,
    ];
    for entry in &control_facts.entries {
        let (spec, report) = control_codec::decode_any_cell_spec_with_report(
            &entry.payload,
            control_codec_options(entry.payload.len()),
        )
        .map_err(|error| map_control_error(error, path))?;
        charge_control_decode_report(budget, report, path)?;
        let control_codec::CellSpecSnapshot::Popup(spec) = spec else {
            continue;
        };
        let identifier = spec.popup_model().identifier();
        let popup = unique_object(&archive, identifier, path)?;
        let popup_index = unique_message_index(popup, 6_206, path)?;
        let (_, report) = popup_codec::decode_popup_menu_model_with_report(
            &popup.messages[popup_index].data,
            popup_codec::DecodeOptions::for_source(&popup.messages[popup_index].data),
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
        owned_identifiers.push(identifier);
    }
    owned_identifiers.sort_unstable();
    owned_identifiers.dedup();
    validate_metadata_ownership(
        source,
        target.component_index,
        &owned_identifiers,
        budget,
        path,
    )?;

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
    popup_native::replace_message_preserving_header(
        &mut archive,
        format_table_identifier,
        format_message_index,
        new_format,
        source
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        &mut archive,
        control_table_identifier,
        control_message_index,
        new_control,
        source
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;

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
    popup_native::replace_message_preserving_header(
        &mut archive,
        tile_identifier,
        tile_message_index,
        patched_tile,
        source
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let archive_bytes = archive
        .to_bytes_with_limits(
            source
                .state
                .options
                .archive()
                .effective_archive_limits()
                .map_err(|_| Error::InvalidSource { path })?,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    budget
        .charge_payload_bytes(archive_bytes.len(), path)
        .map_err(|_| Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::PayloadBytes,
            observed: archive_bytes.len() as u64,
            maximum: u64::MAX,
            path,
        })?;
    Ok(NativeControlOutput {
        member_name: component.name().to_owned(),
        archive_bytes,
    })
}

#[derive(Clone)]
struct ListEntry {
    key: u32,
    ref_count: u32,
    payload: Vec<u8>,
    is_format: bool,
}

struct ListFacts {
    next_key: u32,
    entries: Vec<ListEntry>,
}

#[derive(Default)]
struct TileReferenceCollector {
    tiles: Vec<(u32, u64)>,
}

impl storage_codec::StorageVisitor for TileReferenceCollector {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.tiles
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

#[derive(Default)]
struct BncReferenceCensus {
    formats: Vec<(u32, usize)>,
    controls: Vec<(u32, usize)>,
    invalid: bool,
}

impl BncReferenceCensus {
    fn increment(entries: &mut Vec<(u32, usize)>, identifier: u32) -> bool {
        if let Some((_, count)) = entries.iter_mut().find(|(key, _)| *key == identifier) {
            let Some(next) = count.checked_add(1) else {
                return false;
            };
            *count = next;
        } else {
            entries.push((identifier, 1));
        }
        true
    }
}

impl storage_codec::StorageVisitor for BncReferenceCensus {
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
        if storage.is_empty() && row.cell_count() != 0 {
            self.invalid = true;
            return Ok(());
        }
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc());
        if offsets.is_empty() {
            match row.cell_count() {
                0 => return Ok(()),
                1 => {
                    if !self.count_cell(storage) {
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
            if start >= storage.len() || previous.is_some_and(|prior| prior >= start) {
                self.invalid = true;
                return Ok(());
            }
            if let Some(prior) = previous {
                if !self.count_cell(&storage[prior..start]) {
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

impl BncReferenceCensus {
    fn count_cell(&mut self, source: &[u8]) -> bool {
        let Ok(cell) = BncCell::parse(source) else {
            return false;
        };
        let format = cell.format_identifier();
        let control = cell.control_cell_spec_identifier();
        if (format.is_none() && control.is_some())
            || (format.is_some() != control.is_some() && cell.cell_format_kind() == Some(5))
        {
            return false;
        }
        if let Some(identifier) = format {
            if !Self::increment(&mut self.formats, identifier) {
                return false;
            }
        }
        if let Some(identifier) = control {
            if !Self::increment(&mut self.controls, identifier) {
                return false;
            }
        }
        true
    }
}

struct ListVisitor {
    entries: Vec<ListEntry>,
    segments: usize,
}

impl storage_codec::StorageVisitor for ListVisitor {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        let (payload, is_format) = if let Some(payload) = snapshot.format() {
            (payload.to_vec(), true)
        } else if let Some(payload) = snapshot.cell_spec() {
            (payload.to_vec(), false)
        } else {
            (Vec::new(), false)
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

fn list_facts(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListFacts, Error> {
    list_facts_with_input(source, budget, true, path)
}

fn list_facts_without_input(
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
    let mut visitor = ListVisitor {
        entries: Vec::new(),
        segments: 0,
    };
    let options = budget.residual_storage_options(source);
    let (list, report) =
        storage_codec::decode_table_data_list_with_visitor(source, options, &mut visitor)
            .map_err(|error| map_storage_error(error, path))?;
    if charge_input {
        charge_storage_report_without_input(budget, report, path)?;
    } else {
        charge_storage_report_without_input(budget, report, path)?;
    }
    budget.charge_payload_items(visitor.entries.len(), path)?;
    budget.charge_transaction_work(source.len(), path)?;
    if visitor.segments != 0
        || visitor.entries.iter().enumerate().any(|(index, entry)| {
            visitor.entries[index + 1..]
                .iter()
                .any(|other| other.key == entry.key)
        })
    {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(ListFacts {
        next_key: list.next_list_id(),
        entries: visitor.entries,
    })
}

fn mutate_list(
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
    let desired_key = if desired_payload.is_some() {
        if let Some(existing) = matches {
            Some(existing)
        } else {
            let maximum = facts
                .entries
                .iter()
                .map(|entry| entry.key)
                .max()
                .unwrap_or(0);
            let next = maximum
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            let candidate = facts.next_key.max(1).max(next);
            if facts.entries.iter().any(|entry| entry.key == candidate) {
                return Err(Error::InvalidSource { path });
            }
            Some(candidate)
        }
    } else {
        None
    };
    let mut output = source.to_owned();
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
                let append = if format {
                    storage_codec::TableDataListEntryAppend::format(desired_key, 1, payload)
                } else {
                    storage_codec::TableDataListEntryAppend::control_cell_spec(
                        desired_key,
                        1,
                        payload,
                    )
                };
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

fn apply_list_mutation_with_budget(
    source: &[u8],
    mutation: storage_codec::TableDataListEntryMutation<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    // Preparation owns fallible index/scratch vectors.  Reserve a bounded
    // private allowance before invoking it, then debit the exact prepared
    // requirements before its output buffer is allocated.
    budget.charge_allocations(1, path)?;
    budget.charge_transaction_work(source.len().saturating_mul(2), path)?;
    let options = budget.residual_storage_rewrite_options(source);
    let prepared = storage_codec::prepare_table_data_list_entry_rewrite(source, mutation, options)
        .map_err(|error| map_storage_error(error, path))?;
    let preparation_report = prepared.prepare_report();
    charge_storage_report_without_input(budget, preparation_report, path)?;
    let requirements = prepared.requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_transaction_work(
        requirements
            .output_bytes()
            .saturating_add(requirements.scratch_bytes())
            .saturating_add(requirements.retained_bytes()),
        path,
    )?;
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

fn charge_storage_report(
    budget: &mut TransactionBudget,
    report: storage_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_bytes(report.source_bytes(), path)?;
    charge_storage_report_without_input(budget, report, path)
}

fn charge_storage_report_without_input(
    budget: &mut TransactionBudget,
    report: storage_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_payload_references(report.references(), path)?;
    Ok(())
}

fn map_storage_error(error: storage_codec::DecodeError, path: Path) -> Error {
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
            kind: super::table_cell_pop_up_menu::LimitKind::TransactionWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(storage_codec::DecodeLimit::Text { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireWork,
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
    let options = control_codec_options(source_len);
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
        DecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
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
        control_codec_options(source_len),
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

fn control_codec_options(source_len: usize) -> control_codec::DecodeOptions {
    let bytes = source_len.max(256);
    control_codec::DecodeOptions::new(
        bytes,
        bytes.saturating_mul(2),
        bytes.saturating_mul(8),
        bytes.saturating_mul(16),
        64,
        bytes.saturating_mul(2),
        bytes.saturating_mul(2),
        bytes,
    )
}

fn charge_control_requirements(
    budget: &mut TransactionBudget,
    requirements: control_codec::RewriteExecutionRequirements,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_transaction_work(
        requirements
            .output_bytes()
            .saturating_add(requirements.retained_bytes())
            .saturating_add(requirements.scratch_bytes()),
        path,
    )
}

fn charge_control_decode_report(
    budget: &mut TransactionBudget,
    report: control_codec::DecodeReport,
    path: Path,
) -> Result<(), Error> {
    budget.charge_wire_bytes(report.input_bytes(), path)?;
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_payload_references(report.references(), path)?;
    budget.charge_payload_items(report.items(), path)?;
    budget.charge_allocations(report.allocations(), path)?;
    budget.charge_transaction_work(
        report
            .output_bytes()
            .saturating_add(report.retained_bytes())
            .saturating_add(report.scratch_bytes()),
        path,
    )
}

fn verify_control_report(
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

fn map_control_error(error: control_codec::DecodeError, path: Path) -> Error {
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
        Some(control_codec::DecodeLimit::Items { observed, maximum })
        | Some(control_codec::DecodeLimit::Text { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        Some(control_codec::DecodeLimit::Allocation { requested }) => Error::Allocation {
            amount: requested,
            path,
        },
        Some(control_codec::DecodeLimit::Retained { observed, maximum })
        | Some(control_codec::DecodeLimit::Scratch { observed, maximum }) => Error::LimitExceeded {
            kind: super::table_cell_pop_up_menu::LimitKind::TransactionWork,
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

fn validate_metadata_ownership(
    source: &Package,
    component_index: usize,
    identifiers: &[u64],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
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
    for &identifier in identifiers {
        facts
            .require_current_uuid(component_index, identifier)
            .map_err(|_| Error::InvalidSource { path })?;
    }
    Ok(())
}

fn unique_object(archive: &Archive, identifier: u64, path: Path) -> Result<&ArchiveObject, Error> {
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

fn unique_message_index(object: &ArchiveObject, type_: u32, path: Path) -> Result<usize, Error> {
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

struct ListMessage<'a> {
    message_index: usize,
    payload: &'a [u8],
}

fn unique_list_message<'a>(
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
    archive: &Archive,
    tile_references: &[(u32, u64)],
    formats: &[ListEntry],
    controls: &[ListEntry],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let mut census = BncReferenceCensus::default();
    for (_, tile_identifier) in tile_references {
        let tile = unique_object(archive, *tile_identifier, path)?;
        let message_index = unique_message_index(tile, 6_002, path)?;
        let (_, report) = storage_codec::decode_tile_with_visitor(
            &tile.messages[message_index].data,
            budget.residual_storage_options(&tile.messages[message_index].data),
            &mut census,
        )
        .map_err(|error| map_storage_error(error, path))?;
        charge_storage_report(budget, report, path)?;
        if census.invalid {
            return Err(Error::InvalidSource { path });
        }
    }

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
