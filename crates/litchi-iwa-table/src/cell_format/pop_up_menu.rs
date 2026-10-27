//! The Pop-Up Menu and control reads, native graph rewrite and metadata
//! transition of one resolved cell.

use litchi_iwa_common::table::cell::data_format::{
    CellControl, Checkbox, Currency, CurrencyStyle, Fraction, Number, NumeralSystem, Percentage,
    PopUpMenu, Scientific, Slider, StarRating, Stepper,
    control::{DisplayFormat, Range},
    number::{
        CurrencyCode, DecimalPlaces, FixedDecimalPlaces, FractionAccuracy,
        NegativeStyle as NumberNegativeStyle, ThousandsSeparator,
    },
    numeral_system::{Base, FixedPlaces, NegativeStyle as NumeralNegativeStyle, Places},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
    package_metadata_codec::{
        AdditionSaveTokenBatch, Batch as MetadataBatch, CombinedSaveTokenBatch, RemovalBatch,
        RemovalSaveTokenBatch, RewriteOptions as MetadataRewriteOptions, SaveTokenBatch,
    },
};
use litchi_numbers_wire::BncCell;

use super::{
    Error, LimitKind, Path, Refusal, Target, TransactionBudget, map_core_error,
    map_core_limit_kind, pop_up_menu_metadata as popup_metadata,
    pop_up_menu_native as popup_native,
};
use crate::TableSource;

const LIST_FORMAT: i32 = 2;
const LIST_CONTROL_CELL_SPEC: i32 = 12;

/// The native members and model lifecycle of one Pop-Up Menu rewrite.
pub struct NativeRewrite {
    /// Rewritten native members.
    pub member_edits: Vec<popup_native::NativeMemberEdit>,
    /// Component holding the control list.
    pub control_component_index: usize,
    /// A Pop-Up Menu model the rewrite stored.
    pub added_model: Option<u64>,
    /// Component of the added model.
    pub added_model_component: Option<usize>,
    /// Pop-Up Menu models the rewrite removed.
    pub removed_models: Vec<u64>,
    /// Components and identifiers of the removed models.
    pub removed_model_components: Vec<(usize, u64)>,
}

pub fn rewrite_native_entry(
    source: &dyn TableSource,
    target: Target,
    menu: Option<&PopUpMenu>,
    fresh_identifier: u64,
    path: Path,
    metadata_facts: &popup_metadata::RegistryFacts<'_>,
    budget: &mut TransactionBudget,
) -> Result<NativeRewrite, Error> {
    const POPUP_FORMAT_PAYLOAD: &[u8] = &[0x08, 0x84, 0x02];

    let component = source
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let archive = component.archive();
    let model = archive
        .objects
        .get(target.object_index)
        .and_then(|object| object.messages.get(target.message_index))
        .filter(|message| message.type_ == target.message_type && message.type_ == 6_001)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_wire_bytes(model.data.len(), path)?;
    let options = budget.residual_storage_options(&model.data);
    let (model_snapshot, model_report) =
        storage_codec::decode_table_model_with_report(&model.data, options)
            .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(model_report, false, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        budget.residual_storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(store_report, true, path)?;
    let control_table_identifier = store
        .control_cell_spec_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, tile_report) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        budget.residual_storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(tile_report, true, path)?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let mut tile_matches = tiles
        .tiles
        .iter()
        .filter(|tile| tile.0 == tile_id)
        .map(|tile| tile.1);
    let tile_identifier = tile_matches
        .next()
        .ok_or(Error::Refused(Refusal::CellNotFound))?;
    if tile_matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }

    let item_texts = menu.map(|value| {
        value
            .items()
            .iter()
            .map(|item| item.as_str())
            .collect::<Vec<_>>()
    });
    let desired =
        menu.zip(item_texts.as_deref())
            .map(|(value, items)| popup_native::NativePopUpValue {
                items,
                starts_with_first: value.initial_selection()
                    == litchi_iwa_common::table::cell::data_format::pop_up_menu::InitialSelection::FirstItem,
                first_item_string_identifier: None,
            });
    let physical = source
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let tile_component_index = resolved_component_index(source, tile_identifier, path)?;
    let control_component_index = resolved_component_index(source, control_table_identifier, path)?;
    let format_component_index = resolved_component_index(source, format_table_identifier, path)?;
    let string_table_identifier = store.string_table().identifier();
    let string_component_index = resolved_component_index(source, string_table_identifier, path)?;
    let mut members = Vec::new();
    members
        .try_reserve_exact(source.catalog().len())
        .map_err(|_| Error::Allocation {
            amount: source.catalog().len(),
            path,
        })?;
    for (component_index, member) in source.catalog().iter().enumerate() {
        members.push(popup_native::NativePopUpMember {
            archive: member.archive(),
            component_index,
            member_name: member.name(),
        });
    }
    let same_component = target.component_index == tile_component_index
        && target.component_index == control_component_index
        && target.component_index == format_component_index
        && target.component_index == string_component_index;
    let native_input = popup_native::NativePopUpInput {
        archive,
        component_index: target.component_index,
        model_identifier: target.model_identifier,
        tile_identifier,
        tile_row: target.position.row(),
        tile_column: target.position.column(),
        control_table_identifier,
        format_table_identifier,
        member_name: component.name(),
        format_payload: POPUP_FORMAT_PAYLOAD,
        new_popup_model_identifier: Some(fresh_identifier),
        desired,
        limits: archive_limits,
        snappy_limits,
        path,
    };
    let graph_input = popup_native::NativePopUpGraphInput {
        members: &members,
        model: popup_native::NativePopUpObjectRoute {
            member_index: target.component_index,
            identifier: target.model_identifier,
        },
        tile: popup_native::NativePopUpObjectRoute {
            member_index: tile_component_index,
            identifier: tile_identifier,
        },
        control_table_identifier,
        format_table_identifier,
        // Place a newly created popup model with the rooted control-list
        // member. This keeps its owner current and avoids inventing an
        // unproved external edge solely for creation.
        creation_member_index: control_component_index,
        tile_row: target.position.row(),
        tile_column: target.position.column(),
        format_payload: POPUP_FORMAT_PAYLOAD,
        new_popup_model_identifier: Some(fresh_identifier),
        desired,
        limits: archive_limits,
        snappy_limits,
        path,
    };
    let copy_on_write_candidates = if same_component {
        popup_native::existing_popup_model_identifiers(native_input, budget, path)
    } else {
        popup_native::existing_popup_model_identifiers_multi(graph_input, budget)
    }
    .map_err(|error| map_native_error(error, path))?;
    budget.charge_payload_items(copy_on_write_candidates.len().max(1), path)?;
    budget.charge_payload_references(copy_on_write_candidates.len(), path)?;
    budget.charge_transaction_work(
        copy_on_write_candidates
            .len()
            .saturating_add(1)
            .saturating_mul(archive.objects.len().max(1)),
        path,
    )?;
    // Every existing object whose archive message is mutated or whose
    // ownership is used to decide a cull must have exactly one current UUID
    // owner in this component.  Checking only popup models leaves tile/list
    // aliases and versioned list owners outside the transaction proof.
    let mut owners = vec![
        (target.component_index, target.model_identifier),
        (tile_component_index, tile_identifier),
        (control_component_index, control_table_identifier),
        (format_component_index, format_table_identifier),
    ];
    for identifier in &copy_on_write_candidates {
        let popup_component_index = resolved_component_index(source, *identifier, path)?;
        owners.push((popup_component_index, *identifier));
        if popup_component_index != control_component_index {
            metadata_facts
                .require_external_edge_exact(popup_metadata::ExternalEdge::new(
                    control_component_index,
                    popup_component_index,
                    *identifier,
                    Some(false),
                ))
                .map_err(|error| map_popup_metadata_error(error, path))?;
        }
        for (component_index, component) in source.catalog().iter().enumerate() {
            let found = popup_native::archive_has_popup_reference_strict(
                component.archive(),
                *identifier,
                archive_limits,
            )
            .map_err(|error| map_native_error(error, path))?;
            if component_index == control_component_index {
                // The rooted control list is the selected owner edge for
                // this popup model.  Other current members must not retain
                // an opaque inbound reference that would make cull/COW
                // ambiguous; the popup's own defining object is not an
                // inbound reference and is covered by native locality.
                continue;
            }
            if found {
                // A model referenced by another package member cannot be
                // safely removed or repointed by this same-component owner;
                // reject before the native candidate is allocated.
                return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
            }
        }
    }
    // A native sidecar object may co-locate the string, format, and control
    // list messages. Those semantic roles resolve to one physical ownership
    // route and must be proved once, not mistaken for a duplicate UUID owner.
    owners.sort_unstable();
    owners.dedup();
    if same_component {
        // The focused same-member transaction has always required strict UUID
        // ownership for every existing object it mutates or culls.  Do not
        // let the split-native compatibility policy weaken this established
        // graph boundary.
        for &(component_index, identifier) in &owners {
            metadata_facts
                .require_current_uuid(component_index, identifier)
                .map_err(|error| map_popup_metadata_error(error, path))?;
        }
    } else {
        // Native split Numbers producers may omit UUID records for rooted
        // table/list sidecars.  When a UUID binding exists it must still be a
        // unique current owner; only the genuinely absent binding is admitted
        // as this cross-component compatibility shape. Versioned, conflicting,
        // or alternate-registry ownership remains fail-closed.
        metadata_facts
            .current_uuids_if_registered(&owners)
            .map_err(|error| map_popup_metadata_error(error, path))?;
    }
    let output = if same_component {
        popup_native::rewrite_native_popup_menu(native_input, budget)
    } else {
        popup_native::rewrite_native_popup_menu_multi(graph_input, budget)
    }
    .map_err(|error| map_native_error(error, path))?;
    let mut member_edits = output.member_edits.edits;
    budget.charge_allocations(member_edits.len(), path)?;
    for edit in &mut member_edits {
        let maximum_compressed =
            SnappyStream::maximum_compressed_len(edit.member_bytes.len(), snappy_limits)
                .map_err(|error| map_core_error(error, path))?;
        budget.charge_compressed_bytes(maximum_compressed, path)?;
        budget.charge_transaction_work(
            edit.member_bytes.len().saturating_add(maximum_compressed),
            path,
        )?;
        edit.member_bytes = SnappyStream::compress(&edit.member_bytes, snappy_limits)
            .map_err(|error| map_core_error(error, path))?;
    }
    let added_model = match output.added_object_identifiers.as_slice() {
        [] => None,
        [identifier] => Some(*identifier),
        _ => return Err(Error::Refused(Refusal::UnsupportedDependency { path })),
    };
    let added_model_component = added_model.map(|_| control_component_index);
    let mut removed_model_components = Vec::new();
    removed_model_components
        .try_reserve_exact(output.removed_object_identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: output.removed_object_identifiers.len(),
            path,
        })?;
    for &identifier in &output.removed_object_identifiers {
        let component_index = resolved_component_index(source, identifier, path)?;
        if component_index != control_component_index {
            metadata_facts
                .require_external_edge_exact(popup_metadata::ExternalEdge::new(
                    control_component_index,
                    component_index,
                    identifier,
                    Some(false),
                ))
                .map_err(|error| map_popup_metadata_error(error, path))?;
        }
        removed_model_components.push((component_index, identifier));
    }
    Ok(NativeRewrite {
        member_edits,
        control_component_index,
        added_model,
        added_model_component,
        removed_models: output.removed_object_identifiers,
        removed_model_components,
    })
}

pub fn map_native_error(error: popup_native::NativePopUpError, path: Path) -> Error {
    match error {
        popup_native::NativePopUpError::UnsupportedDependency => {
            Error::Refused(Refusal::UnsupportedDependency { path })
        },
        popup_native::NativePopUpError::Allocation => Error::Allocation { amount: 1, path },
        popup_native::NativePopUpError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind,
            observed,
            maximum,
            path,
        },
        popup_native::NativePopUpError::CoreLimit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_core_limit_kind(kind),
            observed,
            maximum,
            path,
        },
        popup_native::NativePopUpError::InvalidSource
        | popup_native::NativePopUpError::Codec
        | popup_native::NativePopUpError::Archive => Error::InvalidSource { path },
    }
}

pub fn popup_metadata_options(
    payload_bytes: usize,
    additions: usize,
    budget: &TransactionBudget,
) -> MetadataRewriteOptions {
    MetadataRewriteOptions::new(
        payload_bytes.min(budget.remaining_transaction_work).max(1),
        payload_bytes
            .saturating_mul(2)
            .min(budget.remaining_transaction_work)
            .max(1),
        payload_bytes
            .saturating_mul(16)
            .min(budget.remaining_wire_fields)
            .max(1),
        payload_bytes
            .saturating_mul(64)
            .min(budget.remaining_wire_work)
            .max(1),
        64,
        budget.remaining_payload_items.max(1),
        budget.remaining_payload_references.max(1),
        additions.max(1).min(budget.remaining_payload_items.max(1)),
    )
}

pub fn rewrite_popup_metadata(
    _source: &dyn TableSource,
    facts: &popup_metadata::RegistryFacts<'_>,
    component_indices: &[usize],
    fresh: popup_metadata::FreshIdentifier,
    added_model: Option<u64>,
    added_model_component: Option<usize>,
    control_component_index: usize,
    removed_models: &[u64],
    removed_model_components: &[(usize, u64)],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let selectors = facts
        .touched_selectors(component_indices)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let save_tokens = SaveTokenBatch::new(&selectors);
    let registry_operations = usize::from(added_model.is_some())
        .saturating_add(removed_models.len())
        .saturating_add(usize::from(
            added_model_component.is_some_and(|component| component != control_component_index),
        ))
        .saturating_add(
            removed_model_components
                .iter()
                .filter(|(component, _)| *component != control_component_index)
                .count(),
        );
    let options = popup_metadata_options(facts.payload().len(), registry_operations, budget);

    if let Some(identifier) = added_model {
        if identifier != fresh.identifier {
            return Err(Error::InvalidSource { path });
        }
        let component_index = added_model_component
            .or_else(|| component_indices.first().copied())
            .ok_or(Error::InvalidSource { path })?;
        let additions = facts
            .uuid_additions(&[(component_index, fresh)])
            .map_err(|error| map_popup_metadata_error(error, path))?;
        let external_additions = if component_index != control_component_index {
            facts
                .external_additions(&[popup_metadata::ExternalEdge::new(
                    control_component_index,
                    component_index,
                    identifier,
                    Some(false),
                )])
                .map_err(|error| map_popup_metadata_error(error, path))?
        } else {
            Vec::new()
        };
        if !removed_models.is_empty() {
            if removed_model_components.len() != removed_models.len() {
                return Err(Error::InvalidSource { path });
            }
            let mut removals = Vec::new();
            removals
                .try_reserve_exact(removed_models.len())
                .map_err(|_| Error::Allocation {
                    amount: removed_models.len(),
                    path,
                })?;
            for &(removed_component_index, removed_identifier) in removed_model_components {
                match facts.uuid_removal(removed_component_index, removed_identifier) {
                    Ok(removal) => removals.push(removal),
                    // Native-compatible metadata may omit a UUID record for
                    // a physically unique popup model.  Its cull is still
                    // safe when the native/global inbound census proved the
                    // model unreachable; there is simply no UUID registry
                    // entry to remove.
                    Err(error) if error.kind == popup_metadata::FailureKind::MissingRoute => {},
                    Err(error) => return Err(map_popup_metadata_error(error, path)),
                }
            }
            let mut external_edges = Vec::new();
            external_edges
                .try_reserve_exact(removed_model_components.len())
                .map_err(|_| Error::Allocation {
                    amount: removed_model_components.len(),
                    path,
                })?;
            for &(removed_component_index, removed_identifier) in removed_model_components {
                if removed_component_index == control_component_index {
                    continue;
                }
                external_edges.push(popup_metadata::ExternalEdge::new(
                    control_component_index,
                    removed_component_index,
                    removed_identifier,
                    Some(false),
                ));
            }
            let external_removals = facts
                .external_removals(&external_edges)
                .map_err(|error| map_popup_metadata_error(error, path))?;
            let transition = popup_metadata::CombinedTransition::new(
                facts.last_object_identifier(),
                identifier,
                &additions,
                &external_additions,
                &removals,
                &external_removals,
                &[],
            );
            let prepared = popup_metadata::prepare_combined(
                facts,
                CombinedSaveTokenBatch::new(transition, save_tokens),
                options,
            )
            .map_err(|error| map_popup_metadata_error(error, path))?;
            let report = prepared.prepare_report();
            budget.charge_wire_fields(report.fields(), path)?;
            budget.charge_wire_work(report.work_bytes(), path)?;
            budget.charge_wire_nesting(report.max_depth(), path)?;
            budget.charge_payload_items(report.components_scanned(), path)?;
            budget.charge_payload_references(report.references_scanned(), path)?;
            budget.charge_transaction_work(
                report.input_bytes().saturating_add(report.output_bytes()),
                path,
            )?;
            let requirements = prepared.execution_requirements();
            budget.charge_wire_fields(requirements.fields(), path)?;
            budget.charge_wire_work(requirements.work_bytes(), path)?;
            budget.charge_payload_items(requirements.components(), path)?;
            budget.charge_payload_references(requirements.references(), path)?;
            budget.charge_allocations(requirements.allocations(), path)?;
            budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
            budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
            budget.charge_transaction_work(requirements.output_bytes(), path)?;
            return prepared
                .execute(requirements.exact_limits())
                .map(|output| output.into_bytes())
                .map_err(|error| {
                    map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
                });
        }
        let batch = MetadataBatch::new(
            facts.last_object_identifier(),
            identifier,
            &additions,
            &external_additions,
        );
        let prepared = popup_metadata::prepare_additions(
            facts,
            AdditionSaveTokenBatch::new(batch, save_tokens),
            options,
        )
        .map_err(|error| map_popup_metadata_error(error, path))?;
        let report = prepared.prepare_report();
        budget.charge_wire_fields(report.fields(), path)?;
        budget.charge_wire_work(report.work_bytes(), path)?;
        budget.charge_wire_nesting(report.max_depth(), path)?;
        budget.charge_payload_items(report.components_scanned(), path)?;
        budget.charge_payload_references(report.references_scanned(), path)?;
        budget.charge_transaction_work(
            report.input_bytes().saturating_add(report.output_bytes()),
            path,
        )?;
        let requirements = prepared.execution_requirements();
        budget.charge_wire_fields(requirements.fields(), path)?;
        budget.charge_wire_work(requirements.work_bytes(), path)?;
        budget.charge_payload_items(requirements.components(), path)?;
        budget.charge_payload_references(requirements.references(), path)?;
        budget.charge_allocations(requirements.allocations(), path)?;
        budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
        budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
        budget.charge_transaction_work(requirements.output_bytes(), path)?;
        let limits = requirements.exact_limits();
        return prepared
            .execute(limits)
            .map(|output| output.into_bytes())
            .map_err(|error| {
                map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
            });
    }

    if !removed_models.is_empty() {
        let mut removals = Vec::new();
        removals
            .try_reserve_exact(removed_models.len())
            .map_err(|_| Error::Allocation {
                amount: removed_models.len(),
                path,
            })?;
        if removed_model_components.len() != removed_models.len() {
            return Err(Error::InvalidSource { path });
        }
        for &(component_index, identifier) in removed_model_components {
            match facts.uuid_removal(component_index, identifier) {
                Ok(removal) => removals.push(removal),
                // See the combined transition above: absent UUID ownership
                // is admitted only for the strict physical/native route;
                // versioned or conflicting ownership remains an error.
                Err(error) if error.kind == popup_metadata::FailureKind::MissingRoute => {},
                Err(error) => return Err(map_popup_metadata_error(error, path)),
            }
        }
        let mut external_edges = Vec::new();
        external_edges
            .try_reserve_exact(removed_model_components.len())
            .map_err(|_| Error::Allocation {
                amount: removed_model_components.len(),
                path,
            })?;
        for &(component_index, identifier) in removed_model_components {
            if component_index == control_component_index {
                continue;
            }
            external_edges.push(popup_metadata::ExternalEdge::new(
                control_component_index,
                component_index,
                identifier,
                Some(false),
            ));
        }
        let external_removals = facts
            .external_removals(&external_edges)
            .map_err(|error| map_popup_metadata_error(error, path))?;
        if !removals.is_empty() || !external_removals.is_empty() {
            let batch = RemovalBatch::new(
                facts.last_object_identifier(),
                &removals,
                &external_removals,
                &[],
            );
            let prepared = popup_metadata::prepare_removals(
                facts,
                RemovalSaveTokenBatch::new(batch, save_tokens),
                options,
            )
            .map_err(|error| map_popup_metadata_error(error, path))?;
            let report = prepared.prepare_report();
            budget.charge_wire_fields(report.fields(), path)?;
            budget.charge_wire_work(report.work_bytes(), path)?;
            budget.charge_wire_nesting(report.max_depth(), path)?;
            budget.charge_payload_items(report.components_scanned(), path)?;
            budget.charge_payload_references(report.references_scanned(), path)?;
            budget.charge_transaction_work(
                report.input_bytes().saturating_add(report.output_bytes()),
                path,
            )?;
            let requirements = prepared.execution_requirements();
            budget.charge_wire_fields(requirements.fields(), path)?;
            budget.charge_wire_work(requirements.work_bytes(), path)?;
            budget.charge_payload_items(requirements.components(), path)?;
            budget.charge_payload_references(requirements.references(), path)?;
            budget.charge_allocations(requirements.allocations(), path)?;
            budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
            budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
            budget.charge_transaction_work(requirements.output_bytes(), path)?;
            let limits = requirements.exact_limits();
            return prepared
                .execute(limits)
                .map(|output| output.into_bytes())
                .map_err(|error| {
                    map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
                });
        }
    }

    let prepared = popup_metadata::prepare_save_tokens(facts, save_tokens, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let report = prepared.prepare_report();
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_items(report.components_scanned(), path)?;
    budget.charge_payload_references(report.references_scanned(), path)?;
    budget.charge_transaction_work(
        report.input_bytes().saturating_add(report.output_bytes()),
        path,
    )?;
    let requirements = prepared.execution_requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_items(requirements.components(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)?;
    let limits = requirements.exact_limits();
    prepared
        .execute(limits)
        .map(|output| output.into_bytes())
        .map_err(|error| map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path))
}

/// Rewrite only the root and save tokens for an already-validated set of
/// current native members.  Scalar cell controls do not allocate native
/// objects, but a split graph still mutates its tile and list members; their
/// current component tokens must advance atomically with the ZIP edits.
pub fn rewrite_component_save_tokens(
    source: &dyn TableSource,
    component_indices: &[usize],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let metadata_source = popup_metadata::strict_source(source)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    budget.charge_wire_bytes(metadata_source.payload.len(), path)?;
    let options = popup_metadata_options(metadata_source.payload.len(), 0, budget);
    let facts = popup_metadata::inspect(source, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let report = facts.report();
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_items(report.components_scanned(), path)?;
    budget.charge_payload_references(report.references_scanned(), path)?;
    budget.charge_allocations(report.allocations(), path)?;
    budget.charge_scratch_bytes(report.scratch_bytes(), path)?;
    budget.charge_retained_bytes(report.retained_bytes(), path)?;
    budget.charge_transaction_work(
        report.input_bytes().saturating_add(report.output_bytes()),
        path,
    )?;
    if facts.has_physical_alias() {
        return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
    }
    let selectors = facts
        .selectors_for_components(component_indices)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let save_tokens = SaveTokenBatch::new(&selectors);
    let prepared = popup_metadata::prepare_save_tokens(&facts, save_tokens, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let prepare_report = prepared.prepare_report();
    budget.charge_wire_fields(prepare_report.fields(), path)?;
    budget.charge_wire_work(prepare_report.work_bytes(), path)?;
    budget.charge_wire_nesting(prepare_report.max_depth(), path)?;
    budget.charge_payload_items(prepare_report.components_scanned(), path)?;
    budget.charge_payload_references(prepare_report.references_scanned(), path)?;
    budget.charge_transaction_work(
        prepare_report
            .input_bytes()
            .saturating_add(prepare_report.output_bytes()),
        path,
    )?;
    let requirements = prepared.execution_requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_items(requirements.components(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)?;
    let payload = prepared
        .execute(requirements.exact_limits())
        .map(|output| output.into_bytes())
        .map_err(|error| {
            map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
        })?;
    pack_metadata_payload(source, metadata_source.route, payload, path, budget)
}

pub fn pack_metadata_payload(
    source: &dyn TableSource,
    route: super::metadata::MessageRoute,
    payload: Vec<u8>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let component = source
        .catalog()
        .get_index(route.component_index)
        .ok_or(Error::InvalidSource { path })?;
    if component.name() != super::metadata::ENTRY_NAME {
        return Err(Error::InvalidSource { path });
    }
    let physical = source
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource { path })?;
    // `charge_package_source` already charged this borrowed Metadata.iwa
    // member.  The pack phase gets its own archive/payload/work charges
    // below, but must not debit the physical per-entry ceiling twice.
    budget.charge_transaction_work(entry.data().len().saturating_mul(4), path)?;
    budget.charge_allocations(2, path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let stream = SnappyStream::decompress(entry.data(), snappy_limits)
        .map_err(|error| map_core_error(error, path))?;
    let source_payload_bytes = stream.as_bytes().len();
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_archive(&archive, source_payload_bytes, path)?;
    let message = archive
        .objects
        .get_mut(route.object_index)
        .and_then(|object| object.messages.get_mut(route.message_index))
        .filter(|message| message.type_ == super::metadata::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource { path })?;
    let old_message_bytes = message.data.len();
    let archive_upper_bound = source_payload_bytes
        .saturating_sub(old_message_bytes)
        .saturating_add(payload.len());
    let maximum_compressed =
        SnappyStream::maximum_compressed_len(archive_upper_bound, snappy_limits)
            .map_err(|error| map_core_error(error, path))?;
    budget.charge_payload_bytes(archive_upper_bound, path)?;
    budget.charge_compressed_bytes(maximum_compressed, path)?;
    budget.charge_transaction_work(archive_upper_bound.saturating_add(maximum_compressed), path)?;
    message.data = payload;
    let archive_bytes = archive
        .to_bytes(archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    let compressed = SnappyStream::compress(&archive_bytes, snappy_limits)
        .map_err(|error| map_core_error(error, path))?;
    if archive_bytes.len() > archive_upper_bound || compressed.len() > maximum_compressed {
        return Err(Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(archive_bytes.len().max(compressed.len())).unwrap_or(u64::MAX),
            maximum: u64::try_from(archive_upper_bound.max(maximum_compressed)).unwrap_or(u64::MAX),
            path,
        });
    }
    Ok(compressed)
}

pub fn map_popup_metadata_error(error: popup_metadata::MetadataError, path: Path) -> Error {
    match error.kind {
        popup_metadata::FailureKind::Unsupported => {
            Error::Refused(Refusal::UnsupportedDependency { path })
        },
        popup_metadata::FailureKind::Limit => Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::try_from(error.observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(error.maximum).unwrap_or(u64::MAX),
            path,
        },
        popup_metadata::FailureKind::Allocation => Error::Allocation {
            amount: error.allocation,
            path,
        },
        popup_metadata::FailureKind::MissingRoute
        | popup_metadata::FailureKind::AmbiguousRoute
        | popup_metadata::FailureKind::InvalidSource
        | popup_metadata::FailureKind::Conflict
        | popup_metadata::FailureKind::VersionedOwnership => Error::InvalidSource { path },
    }
}

pub fn read_popup(source: &dyn TableSource, target: Target) -> Result<Option<PopUpMenu>, Error> {
    let mut authority = None;
    read_popup_with_policy(source, target, false, &mut authority)
}

/// Read the complete archive-free control sum for one rooted cell.
///
/// Pop-Up Menu remains on the audited popup path; the other four controls use
/// the neutral control codec for their `CellSpecArchive`/format payloads.  The
/// row projection deliberately resolves packed offsets, so a control in any
/// column of a multi-cell row is treated exactly like column zero.
pub fn read_cell_control(
    source: &dyn TableSource,
    target: Target,
) -> Result<Option<CellControl>, Error> {
    let mut authority = None;
    let scalar = read_non_popup_control(source, target, &mut authority)?;
    if let Some(control) = scalar {
        return Ok(Some(control));
    }
    Ok(read_popup_with_policy(source, target, false, &mut authority)?.map(CellControl::PopUpMenu))
}

fn read_non_popup_control<'source>(
    source: &'source dyn TableSource,
    target: Target,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
) -> Result<Option<CellControl>, Error> {
    let path = target.path();
    let model = source
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if model.type_ != target.message_type || model.type_ != 6_001 {
        return Err(Error::InvalidSource { path });
    }
    let options = storage_options(&model.data);
    let (model_snapshot, _) = storage_codec::decode_table_model_with_report(&model.data, options)
        .map_err(|_| Error::InvalidSource { path })?;
    let (store, _) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, _) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let tile_ref = tiles
        .tiles
        .iter()
        .find(|tile| tile.0 == tile_id)
        .map(|tile| tile.1)
        .ok_or(Error::Refused(Refusal::CellNotFound))?;
    let tile_component_index = resolved_component_index(source, tile_ref, path)?;
    validate_selected_model_reference(source, target, tile_ref, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        tile_component_index,
        Some(tile_ref),
        path,
    )?;
    let tile_message = resolve_typed_message_any(source, tile_ref, 6_002, path)?;
    let mut rows = RowCollector::default();
    storage_codec::decode_tile_with_visitor(
        &tile_message.data,
        storage_options(&tile_message.data),
        &mut rows,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let row = rows
        .rows
        .iter()
        .find(|row| row.index == target.position.row())
        .ok_or(Error::Refused(Refusal::CellNotFound))?;
    let cell_bytes = row
        .cell(target.position.column())
        .ok_or(Error::InvalidSource { path })?;
    let cell = BncCell::parse(
        cell_bytes,
        &crate::cell_format::bnc_limits_for_source(cell_bytes),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let format_identifier = cell.format_identifier();
    let control_identifier = cell.control_cell_spec_identifier();
    if format_identifier.is_none() && control_identifier.is_none() {
        return Ok(None);
    }
    if format_identifier.is_none() {
        return Err(Error::InvalidSource { path });
    }
    if cell.cell_format_kind() == Some(5) {
        return Ok(None);
    }
    if !matches!(cell.cell_format_kind(), Some(1 | 2 | 6)) {
        return Err(Error::InvalidSource { path });
    }
    let format_identifier = format_identifier.ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let format_resolved = source
        .resolve_object(format_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, format_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        format_resolved.component_index,
        Some(format_table_identifier),
        path,
    )?;
    let mut format_payload = None;
    for message in format_resolved
        .messages
        .iter()
        .filter(|message| message.type_ == 6_005)
    {
        let mut entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != LIST_FORMAT {
            continue;
        }
        if entries.segments != 0
            || duplicate_list_keys(&entries.entries)
            || format_payload.is_some()
        {
            return Err(Error::InvalidSource { path });
        }
        let entry = entries
            .entries
            .iter()
            .find(|entry| entry.key == format_identifier)
            .ok_or(Error::InvalidSource { path })?;
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        format_payload = entry.format.clone();
    }
    let format_payload = format_payload.ok_or(Error::InvalidSource { path })?;
    let (format, _) = control_codec::decode_control_format_with_report(
        &format_payload,
        control_codec::DecodeOptions::for_source(&format_payload),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let Some(control_identifier) = control_identifier else {
        if matches!(format.format_type(), 263 | 267) {
            return Err(Error::InvalidSource { path });
        }
        return Ok(None);
    };

    let control_table_identifier = store
        .control_cell_spec_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let control_resolved = source
        .resolve_object(control_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, control_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        control_resolved.component_index,
        Some(control_table_identifier),
        path,
    )?;
    let mut selected = None;
    for (message_index, message) in control_resolved
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == 6_005)
    {
        let mut entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != LIST_CONTROL_CELL_SPEC {
            continue;
        }
        if entries.segments != 0 || duplicate_list_keys(&entries.entries) || selected.is_some() {
            return Err(Error::InvalidSource { path });
        }
        let object = source
            .catalog()
            .get_index(control_resolved.component_index)
            .and_then(|component| {
                component
                    .archive()
                    .objects
                    .get(control_resolved.object_index)
            })
            .ok_or(Error::InvalidSource { path })?;
        validate_control_list_metadata(object, message_index, &entries.entries, false, path)?;
        for entry in &entries.entries {
            let spec = entry
                .cell_spec
                .as_deref()
                .ok_or(Error::InvalidSource { path })?;
            let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                spec,
                control_codec::DecodeOptions::for_source(spec),
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let control_codec::CellSpecSnapshot::Popup(spec) = spec else {
                continue;
            };
            let popup_identifier = spec.popup_model().identifier();
            let popup_component_index = resolved_component_index(source, popup_identifier, path)?;
            prove_cross_component_reference(
                source,
                authority,
                control_resolved.component_index,
                popup_component_index,
                Some(popup_identifier),
                path,
            )?;
            let popup = resolve_typed_message_any(source, popup_identifier, 6_206, path)?;
            popup_codec::decode_popup_menu_model_with_report(
                &popup.data,
                popup_codec::DecodeOptions::for_source(&popup.data),
            )
            .map_err(|_| Error::InvalidSource { path })?;
        }
        let entry = entries
            .entries
            .iter()
            .find(|entry| entry.key == control_identifier)
            .ok_or(Error::InvalidSource { path })?;
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        let spec = entry
            .cell_spec
            .as_deref()
            .ok_or(Error::InvalidSource { path })?;
        selected = Some(spec.to_owned());
    }
    let spec = selected.ok_or(Error::InvalidSource { path })?;
    let (spec, _) = control_codec::decode_control_cell_spec_with_report(
        &spec,
        control_codec::DecodeOptions::for_source(&spec),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    control_semantic_value(source, target.component_index, spec, format, path)
}

fn control_semantic_value(
    source: &dyn TableSource,
    component_index: usize,
    spec: control_codec::ControlCellSpecSnapshot<'_>,
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<Option<CellControl>, Error> {
    let interaction = spec.interaction_type();
    match interaction {
        control_codec::CHECKBOX_INTERACTION_TYPE => {
            if format.format_type() != 263 {
                return Err(Error::InvalidSource { path });
            }
            Ok(Some(CellControl::Checkbox(Checkbox)))
        },
        control_codec::STAR_RATING_INTERACTION_TYPE => {
            if format.format_type() != 267
                || spec.range_control_min() != Some(0.0)
                || spec.range_control_max() != Some(5.0)
                || spec.range_control_inc() != Some(1.0)
            {
                return Err(Error::InvalidSource { path });
            }
            Ok(Some(CellControl::StarRating(StarRating)))
        },
        control_codec::SLIDER_INTERACTION_TYPE | control_codec::STEPPER_INTERACTION_TYPE => {
            let range = Range::new(
                spec.range_control_min()
                    .ok_or(Error::InvalidSource { path })?,
                spec.range_control_max()
                    .ok_or(Error::InvalidSource { path })?,
                spec.range_control_inc()
                    .ok_or(Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let display = control_display_format(format, path)?;
            if interaction == control_codec::SLIDER_INTERACTION_TYPE {
                Ok(Some(CellControl::Slider(Slider::new(range, display))))
            } else {
                Ok(Some(CellControl::Stepper(Stepper::new(range, display))))
            }
        },
        7 => {
            // Use the strict popup route for the one graph-bearing variant;
            // resolving it from the current cell avoids accepting a stale
            // control spec that merely resembles a popup.
            let _ = (source, component_index);
            Err(Error::InvalidSource { path })
        },
        _ => Err(Error::InvalidSource { path }),
    }
}

fn control_display_format(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<DisplayFormat, Error> {
    let format_type = format.format_type();
    match format_type {
        256 | 258 => {
            ensure_decimal_format_fields(format, false, path)?;
            let decimal_places = decimal_places_from_native(format.decimal_places(), path)?;
            let negative_style = negative_style_from_native(format.negative_style(), path)?;
            let thousands_separator =
                thousands_separator_from_native(format.show_thousands_separator(), path)?;
            if format_type == 256 {
                Ok(DisplayFormat::Number(Number::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                )))
            } else {
                Ok(DisplayFormat::Percentage(Percentage::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                )))
            }
        },
        257 => {
            ensure_decimal_format_fields(format, true, path)?;
            let decimal_places = decimal_places_from_native(format.decimal_places(), path)?;
            let negative_style = negative_style_from_native(format.negative_style(), path)?;
            let thousands_separator =
                thousands_separator_from_native(format.show_thousands_separator(), path)?;
            let code = match format.currency_code() {
                Some(value) => {
                    CurrencyCode::new(value).map_err(|_| Error::InvalidSource { path })?
                },
                None => CurrencyCode::USD,
            };
            let style = match format.use_accounting_style() {
                Some(false) => CurrencyStyle::Standard,
                Some(true) => CurrencyStyle::Accounting,
                None => CurrencyStyle::Standard,
            };
            Ok(DisplayFormat::Currency(Currency::new(
                code,
                decimal_places,
                negative_style,
                thousands_separator,
                style,
            )))
        },
        259 => {
            ensure_decimal_format_fields(format, false, path)?;
            if format.negative_style().is_some_and(|value| value != 0)
                || format.show_thousands_separator().is_some_and(|value| value)
            {
                return Err(Error::InvalidSource { path });
            }
            let decimal_places = match format.decimal_places() {
                None => FixedDecimalPlaces::TWO,
                Some(value) => {
                    let DecimalPlaces::Fixed(value) =
                        decimal_places_from_native(Some(value), path)?
                    else {
                        return Err(Error::InvalidSource { path });
                    };
                    value
                },
            };
            Ok(DisplayFormat::Scientific(Scientific::new(decimal_places)))
        },
        262 => {
            ensure_fraction_format_fields(format, path)?;
            let accuracy = match format.fraction_accuracy().unwrap_or(u32::MAX - 2) as i32 {
                -1 => FractionAccuracy::UpToOneDigit,
                -2 => FractionAccuracy::UpToTwoDigits,
                -3 => FractionAccuracy::UpToThreeDigits,
                2 => FractionAccuracy::Halves,
                4 => FractionAccuracy::Quarters,
                8 => FractionAccuracy::Eighths,
                16 => FractionAccuracy::Sixteenths,
                10 => FractionAccuracy::Tenths,
                100 => FractionAccuracy::Hundredths,
                _ => return Err(Error::InvalidSource { path }),
            };
            Ok(DisplayFormat::Fraction(Fraction::new(accuracy)))
        },
        269 => {
            ensure_numeral_format_fields(format, path)?;
            let base = Base::new(
                u8::try_from(format.base().unwrap_or(10))
                    .map_err(|_| Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let places = match format.base_places().unwrap_or(0) {
                0 => Places::Minimum,
                value => Places::Fixed(
                    FixedPlaces::new(
                        u8::try_from(value).map_err(|_| Error::InvalidSource { path })?,
                    )
                    .map_err(|_| Error::InvalidSource { path })?,
                ),
            };
            let negative_style = match format.base_use_minus_sign().unwrap_or(true) {
                true => NumeralNegativeStyle::MinusSign,
                false => NumeralNegativeStyle::TwosComplement,
            };
            let value = NumeralSystem::new(base, places, negative_style)
                .map_err(|_| Error::InvalidSource { path })?;
            Ok(DisplayFormat::NumeralSystem(value))
        },
        _ => Err(Error::InvalidSource { path }),
    }
}

fn decimal_places_from_native(value: Option<u32>, path: Path) -> Result<DecimalPlaces, Error> {
    match value.unwrap_or(253) {
        253 => Ok(DecimalPlaces::Automatic),
        value => Ok(DecimalPlaces::Fixed(
            FixedDecimalPlaces::new(
                u8::try_from(value).map_err(|_| Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?,
        )),
    }
}

fn negative_style_from_native(
    value: Option<u32>,
    path: Path,
) -> Result<NumberNegativeStyle, Error> {
    match value {
        None | Some(0) => Ok(NumberNegativeStyle::MinusSign),
        Some(1) => Ok(NumberNegativeStyle::Red),
        Some(2) => Ok(NumberNegativeStyle::Parentheses),
        Some(3) => Ok(NumberNegativeStyle::RedParentheses),
        _ => Err(Error::InvalidSource { path }),
    }
}

fn thousands_separator_from_native(
    value: Option<bool>,
    _path: Path,
) -> Result<ThousandsSeparator, Error> {
    match value {
        Some(false) => Ok(ThousandsSeparator::Hidden),
        Some(true) => Ok(ThousandsSeparator::Shown),
        None => Ok(ThousandsSeparator::Hidden),
    }
}

fn ensure_decimal_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    currency: bool,
    path: Path,
) -> Result<(), Error> {
    if (!currency && (format.currency_code().is_some() || format.use_accounting_style().is_some()))
        || (currency
            && (format.currency_code().is_some() != format.use_accounting_style().is_some()))
        || format.duration_style().is_some()
        || format.base().is_some()
        || format.base_places().is_some()
        || format.base_use_minus_sign().is_some()
        || format.fraction_accuracy().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn ensure_fraction_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<(), Error> {
    if format.decimal_places().is_some()
        || format.currency_code().is_some()
        || format.negative_style().is_some()
        || format.show_thousands_separator().is_some()
        || format.use_accounting_style().is_some()
        || format.duration_style().is_some()
        || format.base().is_some()
        || format.base_places().is_some()
        || format.base_use_minus_sign().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn ensure_numeral_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<(), Error> {
    if format.decimal_places().is_some()
        || format.currency_code().is_some()
        || format.negative_style().is_some()
        || format.show_thousands_separator().is_some()
        || format.use_accounting_style().is_some()
        || format.duration_style().is_some()
        || format.fraction_accuracy().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

pub fn read_popup_with_policy<'source>(
    source: &'source dyn TableSource,
    target: Target,
    allow_missing_control_field_infos: bool,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
) -> Result<Option<PopUpMenu>, Error> {
    let path = target.path();
    let model = source
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if model.type_ != target.message_type || model.type_ != 6_001 {
        return Err(Error::InvalidSource { path });
    }
    let options = storage_options(model.data.as_slice());
    let (model_snapshot, _) = storage_codec::decode_table_model_with_report(&model.data, options)
        .map_err(|_| Error::InvalidSource { path })?;
    let (store, _) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, _) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let tile_ref = tiles
        .tiles
        .iter()
        .find(|tile| tile.0 == tile_id)
        .map(|tile| tile.1)
        .ok_or(Error::Refused(Refusal::CellNotFound))?;
    let tile_component_index = resolved_component_index(source, tile_ref, path)?;
    validate_selected_model_reference(source, target, tile_ref, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        tile_component_index,
        Some(tile_ref),
        path,
    )?;
    let tile_message = resolve_typed_message_any(source, tile_ref, 6_002, path)?;
    let mut rows = RowCollector::default();
    let (_, _) = storage_codec::decode_tile_with_visitor(
        &tile_message.data,
        storage_options(&tile_message.data),
        &mut rows,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let row = rows
        .rows
        .iter()
        .find(|row| row.index == target.position.row())
        .ok_or(Error::Refused(Refusal::CellNotFound))?;
    let cell_buffer = row
        .cell(target.position.column())
        .ok_or(Error::InvalidSource { path })?;
    let cell = BncCell::parse(
        cell_buffer,
        &crate::cell_format::bnc_limits_for_source(cell_buffer),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let format_id = cell.format_identifier();
    let control_id = cell.control_cell_spec_identifier();
    if format_id.is_none() && control_id.is_none() {
        return Ok(None);
    }
    if cell.cell_format_kind() != Some(5) || control_id.is_none() {
        return Ok(None);
    }
    if format_id.is_none() {
        return Err(Error::InvalidSource { path });
    }
    let format_id = format_id.ok_or(Error::InvalidSource { path })?;
    let control_key = control_id.ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let format_resolved = source
        .resolve_object(format_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, format_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        format_resolved.component_index,
        Some(format_table_identifier),
        path,
    )?;
    let mut format_entries = None;
    for message in format_resolved
        .messages
        .iter()
        .filter(|message| message.type_ == 6_005)
    {
        let mut candidate_entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut candidate_entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != 2 {
            continue;
        }
        if candidate_entries.segments != 0 || duplicate_list_keys(&candidate_entries.entries) {
            return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
        }
        if format_entries.replace(candidate_entries).is_some() {
            return Err(Error::InvalidSource { path });
        }
    }
    let format_entries = format_entries.ok_or(Error::InvalidSource { path })?;
    let format_entry = format_entries
        .entries
        .iter()
        .find(|entry| entry.key == format_id)
        .ok_or(Error::InvalidSource { path })?;
    if format_entry.ref_count == 0 || format_entry.format.is_none() {
        return Err(Error::InvalidSource { path });
    }
    let control_table_identifier = store
        .control_cell_spec_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let resolved = source
        .resolve_object(control_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, control_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        resolved.component_index,
        Some(control_table_identifier),
        path,
    )?;
    let mut selected_entries = None;
    for (message_index, message) in resolved.messages.iter().enumerate() {
        if message.type_ != 6_005 {
            continue;
        }
        let mut candidate_entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut candidate_entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() == 12 {
            if candidate_entries.segments != 0 {
                return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
            }
            let object = source
                .catalog()
                .get_index(resolved.component_index)
                .and_then(|component| component.archive().objects.get(resolved.object_index))
                .ok_or(Error::InvalidSource { path })?;
            // Native split producers can omit per-entry FieldInfo while
            // retaining the exact ArchiveInfo aggregate. Preserve strict
            // same-component reads and admit that producer shape only after
            // the control list resolves to a distinct current component.
            let allow_native_aggregate_only = resolved.component_index != target.component_index;
            validate_control_list_metadata(
                object,
                message_index,
                &candidate_entries.entries,
                allow_missing_control_field_infos || allow_native_aggregate_only,
                path,
            )?;
            if selected_entries.replace(candidate_entries).is_some() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    let entries = selected_entries.ok_or(Error::InvalidSource { path })?;
    for (index, entry) in entries.entries.iter().enumerate() {
        if entries.entries[..index]
            .iter()
            .any(|candidate| candidate.key == entry.key)
        {
            return Err(Error::InvalidSource { path });
        }
    }
    let entry = entries
        .entries
        .iter()
        .find(|entry| entry.key == control_key)
        .ok_or(Error::InvalidSource { path })?;
    if entry.ref_count == 0 {
        return Err(Error::InvalidSource { path });
    }
    let spec = entry
        .cell_spec
        .as_deref()
        .ok_or(Error::InvalidSource { path })?;
    let (spec, _) = popup_codec::decode_cell_spec_with_report(spec, popup_options(spec))
        .map_err(|_| Error::InvalidSource { path })?;
    if spec.interaction_type() != 7
        || spec.popup_model().deprecated_type().is_some()
        || spec.popup_model().deprecated_is_external().unwrap_or(false)
    {
        return Err(Error::InvalidSource { path });
    }
    let popup_identifier = spec.popup_model().identifier();
    let popup_component_index = resolved_component_index(source, popup_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        resolved.component_index,
        popup_component_index,
        Some(popup_identifier),
        path,
    )?;
    let popup_message = resolve_typed_message_any(source, popup_identifier, 6_206, path)?;
    let (popup, _) = popup_codec::decode_popup_menu_model_with_report(
        &popup_message.data,
        popup_options(&popup_message.data),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    if !popup.has_nil_sentinel() || popup.item_count() == 0 {
        return Err(Error::InvalidSource { path });
    }
    let items = popup.items().map(|item| item.value()).collect::<Vec<_>>();
    let menu = PopUpMenu::new(items).map_err(|_| Error::InvalidSource { path })?;
    Ok(Some(menu.with_initial_selection(
        if spec.starts_with_first() {
            litchi_iwa_common::table::cell::data_format::pop_up_menu::InitialSelection::FirstItem
        } else {
            litchi_iwa_common::table::cell::data_format::pop_up_menu::InitialSelection::Blank
        },
    )))
}

fn validate_control_list_metadata(
    object: &ArchiveObject,
    message_index: usize,
    entries: &[ListEntry],
    allow_missing_field_infos: bool,
    path: Path,
) -> Result<(), Error> {
    if duplicate_list_keys(entries) {
        return Err(Error::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    // An empty rooted control list is valid only when its control-specific
    // FieldInfo paths are empty too.  Leaving a stale [3,key] field behind is
    // precisely the metadata ghost that can keep a culled CellSpec alive.
    if entries.is_empty() {
        if info
            .field_infos
            .iter()
            .any(|field| field.path.as_slice().first() == Some(&3))
        {
            return Err(Error::InvalidSource { path });
        }
        return Ok(());
    }
    // A control list can legitimately mix the four scalar controls with
    // Pop-Up Menu entries.  Only the popup interaction owns a 6206 model;
    // scalar CellSpecs therefore contribute neither aggregate references nor
    // [3,key] FieldInfo edges.  Decode every entry through the neutral
    // control projection so scalar records are still validated strictly
    // rather than being treated as opaque bytes.
    let mut expected = Vec::with_capacity(entries.len());
    for entry in entries {
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        let payload = entry
            .cell_spec
            .as_deref()
            .ok_or(Error::InvalidSource { path })?;
        let (spec, _) = control_codec::decode_cell_spec_with_report(
            payload,
            control_codec::DecodeOptions::for_source(payload),
        )
        .map_err(|_| Error::InvalidSource { path })?;
        match spec {
            control_codec::CellSpecSnapshot::Control(spec) => {
                if !matches!(
                    spec.interaction_type(),
                    control_codec::CHECKBOX_INTERACTION_TYPE
                        | control_codec::STAR_RATING_INTERACTION_TYPE
                        | control_codec::SLIDER_INTERACTION_TYPE
                        | control_codec::STEPPER_INTERACTION_TYPE
                ) {
                    return Err(Error::InvalidSource { path });
                }
            },
            control_codec::CellSpecSnapshot::Popup(spec) => {
                if spec.interaction_type() != 7
                    || spec.popup_model().identifier() == 0
                    || spec.popup_model().deprecated_type().is_some()
                    || spec.popup_model().deprecated_is_external().unwrap_or(false)
                {
                    return Err(Error::InvalidSource { path });
                }
                expected.push((entry.key, spec.popup_model().identifier()));
            },
        }
    }
    let mut expected_aggregate = expected.iter().map(|(_, id)| *id).collect::<Vec<_>>();
    expected_aggregate.sort_unstable();
    expected_aggregate.dedup();
    let mut aggregate = info.object_references.clone();
    aggregate.sort_unstable();
    aggregate.dedup();
    if info.object_references.contains(&0)
        || info.object_references.len() != aggregate.len()
        || aggregate != expected_aggregate
    {
        return Err(Error::InvalidSource { path });
    }
    if allow_missing_field_infos && info.field_infos.is_empty() {
        return Ok(());
    }
    for (key, identifier) in &expected {
        let matches = info
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [3, *key]);
        let matching = matches.collect::<Vec<_>>();
        if matching.len() != 1
            || matching[0]
                .r#type
                .is_some_and(|kind| kind != litchi_iwa_core::FieldType::ObjectReference)
            || matching[0].object_references.as_slice() != [*identifier]
        {
            return Err(Error::InvalidSource { path });
        }
    }
    // Every control-list FieldInfo must be an exact [3,key] singleton above;
    // an aggregate-only [3] edge or an unrelated key is not a safe ownership
    // proof and is rejected rather than broadcast during a rewrite.
    if info.field_infos.iter().any(|field| {
        field.path.as_slice().first() == Some(&3)
            && !expected
                .iter()
                .any(|(key, _)| field.path.as_slice() == [3, *key])
    }) {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

/// Read one selected cell while charging the same transaction ledger used by
/// the changed route.  The strict read helper remains reusable by the public
/// no-op/read API; changed commits call this wrapper before native candidates
/// and again for the private candidate semantic readback.
pub fn read_popup_with_budget(
    source: &dyn TableSource,
    target: Target,
    budget: &mut TransactionBudget,
    allow_missing_control_field_infos: bool,
) -> Result<Option<PopUpMenu>, Error> {
    let path = target.path();
    let payload_len = source
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .map(|message| message.data.len())
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_wire_bytes(payload_len, path)?;
    budget.charge_wire_work(payload_len.saturating_mul(16), path)?;
    budget.charge_payload_items(1, path)?;
    budget.charge_transaction_work(payload_len.saturating_mul(16), path)?;
    let mut authority = None;
    read_popup_with_policy(
        source,
        target,
        allow_missing_control_field_infos,
        &mut authority,
    )
}

pub fn changed_member_count(
    source: &litchi_iwa_archive::SourceCatalog,
    target: &litchi_iwa_archive::SourceCatalog,
) -> usize {
    source
        .package()
        .iter()
        .filter(|entry| !entry.name().starts_with("preview"))
        .filter(|entry| {
            target
                .package()
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .count()
}

fn storage_options(source: &[u8]) -> storage_codec::DecodeOptions {
    let bytes = source.len().max(1);
    storage_codec::DecodeOptions::new(
        bytes,
        bytes.saturating_mul(8).max(1),
        bytes.saturating_mul(64).max(1),
        64,
        bytes.saturating_mul(2).max(1),
        bytes.saturating_mul(2).max(1),
    )
}

fn popup_options(source: &[u8]) -> popup_codec::DecodeOptions {
    let bytes = source.len().max(1);
    popup_codec::DecodeOptions::for_source(source)
        .with_max_output_bytes(bytes.saturating_mul(2).max(1))
        .with_max_items(bytes.max(1))
        .with_max_text_bytes(bytes.saturating_mul(2).max(1))
}

fn target_row_tile(tile_size: u32, row: u32) -> u32 {
    row / tile_size.max(1)
}

fn resolved_component_index(
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

/// Require the selected TableModel message header to declare each native
/// sidecar edge exactly once. Producer-omitted FieldInfo is accepted; when a
/// FieldInfo occurrence is present, it must be unique and object-reference
/// typed. This bounded slice does not claim exact field-path authority for
/// producer-omitted metadata.
pub fn validate_selected_model_reference(
    source: &dyn TableSource,
    target: Target,
    identifier: u64,
    path: Path,
) -> Result<(), Error> {
    let info = source
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.archive_info.message_infos.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource { path });
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if field
            .r#type
            .is_some_and(|kind| kind != litchi_iwa_core::FieldType::ObjectReference)
        {
            return Err(Error::InvalidSource { path });
        }
        field_occurrences = field_occurrences
            .checked_add(occurrences)
            .ok_or(Error::InvalidSource { path })?;
    }
    if field_occurrences > 1 {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

/// Resolve a generated-free message through the package index without
/// imposing the model component as an implicit owner.  The caller must prove
/// every cross-component edge with `prove_cross_component_reference`; this
/// split is what prevents a sidecar member from being mistaken for an
/// unowned/ambiguous object while still allowing the native graph to span
/// CalculationEngine, Tile, and DataList members.
fn resolve_typed_message_any(
    source: &dyn TableSource,
    identifier: u64,
    message_type: u32,
    path: Path,
) -> Result<&RawMessage, Error> {
    let resolved = source
        .resolve_object(identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let mut matches = resolved
        .messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let message = matches.next().ok_or(Error::InvalidSource { path })?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(message)
}

/// Validate the metadata ownership edge before a cross-component read is
/// exposed.  Sidecar components normally have no UUID object records, so an
/// exact current external component edge is the required authority.  Missing
/// Metadata, a missing/ambiguous effective locator, versioned edges, duplicate
/// edges, and weak/reference-shape conflicts all fail closed.
pub fn prove_cross_component_reference<'source>(
    source: &'source dyn TableSource,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
    source_component_index: usize,
    target_component_index: usize,
    object_identifier: Option<u64>,
    path: Path,
) -> Result<(), Error> {
    if source_component_index == target_component_index {
        return Ok(());
    }
    if authority.is_none() {
        let bytes = source.source_bytes().len().max(1);
        let options = MetadataRewriteOptions::new(
            bytes,
            bytes.saturating_mul(2),
            bytes.saturating_mul(16),
            bytes.saturating_mul(64),
            64,
            bytes.saturating_mul(2),
            bytes.saturating_mul(2),
            1,
        );
        *authority = Some(
            popup_metadata::inspect_cross_component_read(source, options)
                .map_err(|_| Error::Refused(Refusal::UnsupportedDependency { path }))?,
        );
    }
    let facts = authority
        .as_ref()
        .ok_or(Error::Refused(Refusal::UnsupportedDependency { path }))?;
    if facts.has_physical_alias() {
        return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
    }
    facts
        .require_external_edge(
            source_component_index,
            target_component_index,
            object_identifier,
            Some(false),
        )
        .map_err(|_| Error::Refused(Refusal::UnsupportedDependency { path }))
}

#[derive(Default)]
struct TileCollector {
    tiles: Vec<(u32, u64)>,
}

impl storage_codec::StorageVisitor for TileCollector {
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
struct RowCollector {
    rows: Vec<RowFact>,
}

struct RowFact {
    index: u32,
    cell_count: u32,
    storage: Vec<u8>,
    offsets: Vec<u8>,
    wide_offsets: bool,
}

impl RowFact {
    fn cell(&self, column: u32) -> Option<&[u8]> {
        let column = usize::try_from(column).ok()?;
        let count = usize::try_from(self.cell_count).ok()?;
        if count == 1 && self.offsets.is_empty() {
            return (column == 0).then_some(self.storage.as_slice());
        }
        if !self.offsets.len().is_multiple_of(2) {
            return None;
        }
        let slot_count = self.offsets.len() / 2;
        if column >= slot_count || slot_count < count {
            return None;
        }
        let unit = if self.wide_offsets { 4usize } else { 1usize };
        let mut starts = Vec::with_capacity(slot_count);
        let mut previous = None;
        for encoded in self.offsets.chunks_exact(2) {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            if raw == u16::MAX {
                starts.push(None);
                continue;
            }
            let start = usize::from(raw).checked_mul(unit)?;
            if previous.is_some_and(|prior| prior >= start) {
                return None;
            }
            starts.push(Some(start));
            previous = Some(start);
        }
        if starts.iter().flatten().count() != count {
            return None;
        }
        let start = starts.get(column).copied().flatten()?;
        let end = starts
            .iter()
            .skip(column + 1)
            .flatten()
            .next()
            .copied()
            .unwrap_or(self.storage.len());
        (start < end && end <= self.storage.len()).then_some(&self.storage[start..end])
    }
}

impl storage_codec::StorageVisitor for RowCollector {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let storage = row
            .cell_storage_buffer()
            .unwrap_or_else(|| row.cell_storage_buffer_pre_bnc())
            .to_vec();
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc())
            .to_vec();
        self.rows.push(RowFact {
            index: row.tile_row_index(),
            cell_count: row.cell_count(),
            storage,
            offsets,
            wide_offsets: row.has_wide_offsets().unwrap_or(false),
        });
        Ok(())
    }
}

#[derive(Default)]
struct ListCollector {
    entries: Vec<ListEntry>,
    segments: usize,
}

#[derive(Clone)]
struct ListEntry {
    key: u32,
    ref_count: u32,
    cell_spec: Option<Vec<u8>>,
    format: Option<Vec<u8>>,
}

fn duplicate_list_keys(entries: &[ListEntry]) -> bool {
    entries.iter().enumerate().any(|(index, entry)| {
        entries[..index]
            .iter()
            .any(|previous| previous.key == entry.key)
    })
}

impl storage_codec::StorageVisitor for ListCollector {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        self.entries.push(ListEntry {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            cell_spec: snapshot.cell_spec().map(ToOwned::to_owned),
            format: snapshot.format().map(ToOwned::to_owned),
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

#[cfg(test)]
mod tests {
    use super::{Error, LimitKind, Path, map_core_error, map_native_error};

    #[test]
    fn native_limit_payload_reaches_public_error() {
        let error = map_native_error(
            super::super::pop_up_menu_native::NativePopUpError::Limit {
                kind: LimitKind::WireWork,
                observed: 31,
                maximum: 13,
            },
            Path::Package,
        );
        assert_eq!(
            error,
            Error::Limit {
                kind: LimitKind::WireWork,
                observed: 31,
                maximum: 13,
                path: Path::Package,
            }
        );
    }

    #[test]
    fn bounded_snappy_failures_keep_typed_resource_kinds() {
        let limit = map_core_error(
            litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes,
                observed: 9,
                maximum: 8,
            },
            Path::Package,
        );
        assert_eq!(
            limit,
            Error::Limit {
                kind: LimitKind::CompressedBytes,
                observed: 9,
                maximum: 8,
                path: Path::Package,
            }
        );

        let allocation = map_core_error(
            litchi_iwa_core::Error::Allocation {
                resource: "test",
                requested: 12,
            },
            Path::Package,
        );
        assert_eq!(
            allocation,
            Error::Allocation {
                amount: 12,
                path: Path::Package,
            }
        );
    }
}
