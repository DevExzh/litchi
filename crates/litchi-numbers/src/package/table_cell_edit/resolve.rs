//! Changed-only rooted topology resolution for table-cell publication.
//!
//! Selection is deliberately finished before this module runs.  It admits one
//! changed batch only after the semantic table, native owner chain, storage
//! graph, and calculation-engine roots have all been proved.  The transaction
//! retains physical coordinates rather than publishing native object IDs.

use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::{
    numbers_table_cell_dependency_codec as dependency_codec,
    numbers_table_cell_storage_codec as storage_codec, tst::table_data_list::ListType,
};

use super::super::{Package, table_headers};
use crate::{
    package::table_cells::{DependencyKind, Error, LimitKind, Path},
    table::{CellPosition, Dimensions, lock::State as LockState},
};

const TILE_MESSAGE_TYPE: u32 = 6_002;
const DATA_LIST_MESSAGE_TYPES: [u32; 2] = [6_005, 6_201];
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const DATA_LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const REFERENCE_TRACKER_MESSAGE_TYPE: u32 = 4_004;
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;
const RANGE_PRECEDENTS_TILE_MESSAGE_TYPE: u32 = 4_010;
const HEADER_NAME_MANAGER_MESSAGE_TYPE: u32 = 6_366;
const HIDDEN_STATE_OWNER_MESSAGE_TYPE: u32 = 6_204;
const RICH_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
const RICH_STORAGE_MESSAGE_TYPE: u32 = 2_001;
const DEFAULT_TILE_SIZE: u32 = 256;

/// Exact physical coordinates for one already-validated payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub(super) struct MessageRoute {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
}

impl MessageRoute {
    const fn new(
        resolved: super::super::Resolved<'_>,
        message_index: usize,
        message_type: u32,
    ) -> Self {
        Self {
            component_index: resolved.component_index,
            object_index: resolved.object_index,
            message_index,
            message_type,
        }
    }
}

fn message_payload_at_route(
    source: &Package,
    route: MessageRoute,
    path: Path,
) -> Result<&[u8], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == route.message_type)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource { path })
}

/// One source tile keyed by its table-local row-strip number.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct TileRoute {
    pub(super) tile_id: u32,
    pub(super) message: MessageRoute,
}

/// One direct or segmented `TableDataList` owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ListRoute {
    pub(super) message: MessageRoute,
    pub(super) segments: Vec<MessageRoute>,
    pub(super) entries: usize,
}

/// Every model-owned list route which may participate in a scalar rewrite.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ListRoutes {
    pub(super) string: ListRoute,
    pub(super) style: ListRoute,
    pub(super) formula: ListRoute,
    pub(super) format_pre_bnc: ListRoute,
    pub(super) formula_error: Option<ListRoute>,
    pub(super) custom_format: Option<ListRoute>,
    pub(super) multiple_choice: Option<ListRoute>,
    pub(super) rich_text: Option<ListRoute>,
    pub(super) conditional_style: Option<ListRoute>,
    pub(super) comment: Option<ListRoute>,
    pub(super) import_warning: Option<ListRoute>,
    pub(super) control_cell_spec: Option<ListRoute>,
    pub(super) format: Option<ListRoute>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RichFieldRefs {
    pub(super) field_info_index: usize,
    pub(super) path: Vec<u32>,
    pub(super) object_references: Vec<u64>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RichObjectRoute {
    pub(super) message: MessageRoute,
    pub(super) object_id: u64,
    pub(super) message_type: u32,
    pub(super) object_references: Vec<u64>,
    pub(super) field_references: Vec<RichFieldRefs>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum RichEntryOwner {
    Root,
    Segment {
        message: MessageRoute,
        object_id: u64,
        owner_entries: u32,
        root_references: u32,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RichEntryRoute {
    pub(super) key: u32,
    pub(super) ref_count: u32,
    pub(super) owner: RichEntryOwner,
    pub(super) root: MessageRoute,
    pub(super) root_object_id: u64,
    pub(super) next_key: u32,
    pub(super) pair_index: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RichResolvedPairRoute {
    pub(super) payload: RichObjectRoute,
    pub(super) storage: RichObjectRoute,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RichRouteIndex {
    pub(super) entries: Vec<RichEntryRoute>,
    pub(super) pairs: Vec<RichResolvedPairRoute>,
    pub(super) local_object_ids: Vec<u64>,
    pub(super) payload_list_inbound: Vec<(u64, u32)>,
    pub(super) storage_payload_inbound: Vec<(u64, u32)>,
    pub(super) report: CodecUsage,
}

/// Strictly followed storage graph rooted in the selected `TableModel`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct StorageRoutes {
    pub(super) model: MessageRoute,
    pub(super) tile_size: u32,
    pub(super) tiles: Vec<TileRoute>,
    pub(super) row_headers: Vec<MessageRoute>,
    pub(super) column_headers: MessageRoute,
    pub(super) lists: ListRoutes,
    pub(super) hidden_state_owners: Vec<MessageRoute>,
    pub(super) merge_region_map: Option<MessageRoute>,
    pub(super) sort_rule_tracker: Option<MessageRoute>,
    pub(super) rich: Option<RichRouteIndex>,
}

/// Strictly followed calculation-engine graph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(super) struct DependencyRoutes {
    pub(super) engine: Option<MessageRoute>,
    /// Exact global formula-family count from the calculation-engine tracker.
    /// This includes opaque non-table formula families as well as table cells.
    pub(super) formula_count: u64,
    pub(super) engine_sidecars: Vec<MessageRoute>,
    pub(super) formula_owners: Vec<FormulaOwnerRoute>,
    pub(super) selected_formula_owner: Option<SelectedFormulaOwnerRoute>,
    pub(super) inert_marker_tiles: Vec<MessageRoute>,
    pub(super) cell_record_tiles: Vec<MessageRoute>,
    pub(super) range_precedent_tiles: Vec<MessageRoute>,
    pub(super) header_name_manager: Option<MessageRoute>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct FormulaOwnerRoute {
    pub(super) message: MessageRoute,
    /// Source-bound table/drawable object referenced by this owner, or `None`
    /// for non-table engine owners such as deferred marker owners.
    pub(super) formula_owner_object_id: Option<u64>,
    pub(super) internal_owner_id: u32,
    pub(super) uid_lower: u64,
    pub(super) uid_upper: u64,
    pub(super) cell_record_tiles: Vec<MessageRoute>,
    pub(super) range_precedent_tiles: Vec<RangePrecedentRoute>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct RangePrecedentRoute {
    pub(super) message: MessageRoute,
    pub(super) target_owner: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct SelectedFormulaOwnerRoute {
    pub(super) message: MessageRoute,
    pub(super) internal_owner_id: u32,
    pub(super) uid_lower: u64,
    pub(super) uid_upper: u64,
}

/// One exact native target admitted for a changed scalar-cell transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Target {
    pub(super) native: table_headers::Target,
    pub(super) storage: StorageRoutes,
    pub(super) dependencies: DependencyRoutes,
}

impl Target {
    pub(super) const fn path(&self) -> Path {
        selected_path(self.native)
    }
}

/// Aggregate strict-codec work already consumed by admission.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct CodecUsage {
    pub(super) source_bytes: usize,
    pub(super) fields: usize,
    pub(super) work_bytes: usize,
    pub(super) references: usize,
    pub(super) reference_bytes: usize,
    pub(super) text_bytes: usize,
    pub(super) max_depth: u32,
}

/// Exact work already consumed by changed-only target admission.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ResolveReport {
    pub(super) ownership: table_headers::ownership::OwnershipReport,
    pub(super) codecs: CodecUsage,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocation_events: usize,
}

/// Resolve and globally prove one changed edit target.
///
/// `positions` is the final, strictly sorted, duplicate-free set remaining
/// after semantic no-op elimination.  Empty batches bypass this resolver.
#[cfg(test)]
pub(super) fn resolve_changed_target(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    positions: &[CellPosition],
) -> Result<(Target, ResolveReport), Error> {
    let remaining = super::budget::TransactionBudget::new(source)?.remaining()?;
    resolve_changed_target_with_remaining(
        source,
        sheet_position,
        table_position,
        positions,
        remaining,
    )
}

#[cfg(test)]
pub(super) fn resolve_changed_target_with_remaining(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    positions: &[CellPosition],
    remaining: super::budget::Remaining,
) -> Result<(Target, ResolveReport), Error> {
    resolve_target_with_remaining(
        source,
        sheet_position,
        table_position,
        positions,
        remaining,
        false,
    )
}

pub(super) fn resolve_formula_read_target_with_remaining(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    positions: &[CellPosition],
    remaining: super::budget::Remaining,
) -> Result<(Target, ResolveReport), Error> {
    resolve_target_with_remaining(
        source,
        sheet_position,
        table_position,
        positions,
        remaining,
        true,
    )
}

fn resolve_target_with_remaining(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    positions: &[CellPosition],
    remaining: super::budget::Remaining,
    permit_locked_read: bool,
) -> Result<(Target, ResolveReport), Error> {
    let native = table_headers::resolve::resolve_target(source, sheet_position, table_position)
        .map_err(map_header_failure)?;
    let path = selected_path(native);
    validate_positions(positions, native.rows, native.columns, path)?;
    validate_header_counts(native, path)?;
    if native.locked == LockState::Locked && !permit_locked_read {
        return Err(Error::TableLocked { path });
    }

    let ownership =
        table_headers::ownership::validate_selected_ownership_with_report(source, native)
            .map_err(map_header_failure)?;
    table_headers::dependencies::validate_dependencies(
        source,
        native,
        native.settings,
        native.settings,
    )
    .map_err(map_header_failure)?;
    let payload =
        table_headers::rewrite::selected_payload(source, native).map_err(map_header_failure)?;
    let mut budget = ResolutionBudget::new(source, remaining, path)?;
    let (model, report) =
        storage_codec::decode_table_model_with_report(payload, budget.options(payload.len())?)
            .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    if model.number_of_rows() != native.rows || model.number_of_columns() != native.columns {
        return Err(Error::InvalidSource { path });
    }
    let owner_markers = validate_model_owners(source, native, model, payload, path)?;

    let data_payload = model.base_data_store();
    let mut storage_count = StorageCount::default();
    let (_store, report) = storage_codec::decode_data_store_with_visitor(
        data_payload,
        budget.options(data_payload.len())?,
        &mut storage_count,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    storage_count.validate(path)?;
    let mut storage_stage = StorageStage::with_capacity(storage_count, &mut budget, path)?;
    let (data_store, report) = storage_codec::decode_data_store_with_visitor(
        data_payload,
        budget.options(data_payload.len())?,
        &mut storage_stage,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    storage_stage.finish(storage_count, path)?;
    let (tile_storage, report) = storage_codec::decode_tile_storage_with_report(
        data_store.tiles(),
        budget.options(data_store.tiles().len())?,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;

    let model_object = object_at_native_model(source, native, path)?;
    let mut identities = RoleIdentities::new(native, &mut budget, path)?;
    let hidden_owner_count = usize::from(model.hidden_state_formula_owner_for_columns().is_some())
        .checked_add(usize::from(
            model.hidden_state_formula_owner_for_rows().is_some(),
        ))
        .ok_or(Error::InvalidSource { path })?;
    let mut hidden_state_owners =
        budget.reserve_retained(hidden_owner_count, LimitKind::Objects)?;
    if let Some(route) = validate_dormant_hidden_owner(
        source,
        model_object,
        native.message_index,
        model.hidden_state_formula_owner_for_columns(),
        34,
        &mut identities,
        &mut budget,
        path,
    )? {
        hidden_state_owners.push(route);
    }
    if let Some(route) = validate_dormant_hidden_owner(
        source,
        model_object,
        native.message_index,
        model.hidden_state_formula_owner_for_rows(),
        35,
        &mut identities,
        &mut budget,
        path,
    )? {
        hidden_state_owners.push(route);
    }
    let merge_region_map = resolve_merge_map(
        source,
        model_object,
        native.message_index,
        data_store.merge_region_map(),
        &mut identities,
        &mut budget,
        path,
    )?;
    let sort_rule_tracker = resolve_sort_rule_tracker(
        source,
        model_object,
        native.message_index,
        owner_markers.sort_rule_tracker,
        &mut identities,
        &mut budget,
        path,
    )?;

    let tile_size = tile_storage.tile_size().unwrap_or(DEFAULT_TILE_SIZE);
    if tile_size == 0 {
        return Err(Error::InvalidSource { path });
    }
    let tiles = resolve_tiles(
        source,
        model_object,
        native.message_index,
        &storage_stage.tiles,
        tile_size,
        native,
        &mut identities,
        &mut budget,
        path,
    )?;
    let row_headers = resolve_headers(
        source,
        model_object,
        native.message_index,
        &storage_stage.header_buckets,
        &[4, 1, 2],
        native.rows,
        HeaderAxis::Row,
        &mut identities,
        &mut budget,
        path,
    )?;
    let column_headers = resolve_one_header(
        source,
        model_object,
        native.message_index,
        data_store.column_headers(),
        &[4, 2],
        native.columns,
        HeaderAxis::Column,
        &mut identities,
        &mut budget,
        path,
    )?;
    budget.release_scratch(&storage_stage.tiles)?;
    budget.release_scratch(&storage_stage.header_buckets)?;
    drop(storage_stage);

    let lists = resolve_lists(
        source,
        model_object,
        native.message_index,
        data_store,
        native,
        &mut identities,
        &mut budget,
        path,
    )?;
    if owner_markers.conditional_style.is_some()
        && lists
            .conditional_style
            .as_ref()
            .is_some_and(|route| route.entries != 0)
    {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::ConditionalStyle,
        });
    }
    let rich = lists
        .rich_text
        .as_ref()
        .map(|route| resolve_rich_route_index(source, route, &mut identities, &mut budget, path))
        .transpose()?;
    let dependencies = resolve_dependencies(
        source,
        native,
        positions,
        owner_markers,
        &mut identities,
        &mut budget,
        path,
    )?;
    identities.finish(&mut budget, path)?;
    identities.release_scratch(&mut budget)?;
    if budget.current_scratch_bytes != 0 {
        return Err(Error::InvalidSource { path });
    }
    let model_route = MessageRoute {
        component_index: native.component_index,
        object_index: native.object_index,
        message_index: native.message_index,
        message_type: native.message_type,
    };
    Ok((
        Target {
            native,
            storage: StorageRoutes {
                model: model_route,
                tile_size,
                tiles,
                row_headers,
                column_headers,
                lists,
                hidden_state_owners,
                merge_region_map,
                sort_rule_tracker,
                rich,
            },
            dependencies,
        },
        ResolveReport {
            ownership,
            codecs: budget.usage,
            retained_elements: budget.retained_elements,
            retained_bytes: budget.retained_bytes,
            peak_scratch_bytes: budget.peak_scratch_bytes,
            allocation_events: budget.allocation_events,
        },
    ))
}

fn validate_positions(
    positions: &[CellPosition],
    rows: u32,
    columns: u32,
    path: Path,
) -> Result<(), Error> {
    let dimensions = Dimensions::new(rows, columns);
    let mut previous = None;
    for &position in positions {
        if previous.is_some_and(|prior| prior >= position) {
            return if previous == Some(position) {
                Err(Error::DuplicatePosition { position })
            } else {
                Err(Error::InvalidSource { path })
            };
        }
        if position.row() >= rows || position.column() >= columns {
            return Err(Error::OutOfBounds {
                position,
                dimensions,
            });
        }
        previous = Some(position);
    }
    Ok(())
}

fn validate_header_counts(native: table_headers::Target, path: Path) -> Result<(), Error> {
    let header_rows = u32::try_from(native.settings.header_row_count())
        .map_err(|_error| Error::InvalidSource { path })?;
    let footer_rows = u32::try_from(native.settings.footer_row_count())
        .map_err(|_error| Error::InvalidSource { path })?;
    let header_columns = u32::try_from(native.settings.header_column_count())
        .map_err(|_error| Error::InvalidSource { path })?;
    if header_rows
        .checked_add(footer_rows)
        .is_none_or(|count| count > native.rows)
        || header_columns > native.columns
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct ModelOwnerMarkers {
    spill: Option<UuidKey>,
    conditional_style: Option<UuidKey>,
    haunted: Option<UuidKey>,
    sort_rule_tracker: Option<u64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct UuidKey {
    lower: u64,
    upper: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct OwnerIdentity {
    uid: UuidKey,
    internal: u32,
}

fn validate_model_owners(
    source: &Package,
    native: table_headers::Target,
    model: storage_codec::TableModelSnapshot<'_>,
    payload: &[u8],
    path: Path,
) -> Result<ModelOwnerMarkers, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut category_reference = None;
    let mut filter_reference = None;
    let mut haunted = None;
    let mut haunted_seen = false;
    let mut sort_rule_tracker = None;
    for field in view.fields() {
        let kind = match field.number() {
            44 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if !inert_sort_order(field.payload()) {
                    return Err(Error::UnsupportedDependency {
                        path,
                        kind: DependencyKind::FormulaCache,
                    });
                }
                None
            },
            44 => return Err(Error::InvalidSource { path }),
            45 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                let id = parse_sort_rule_tracker(field.payload()).ok_or(
                    Error::UnsupportedDependency {
                        path,
                        kind: DependencyKind::FormulaCache,
                    },
                )?;
                if sort_rule_tracker.replace(id).is_some() {
                    return Err(Error::InvalidSource { path });
                }
                None
            },
            45 => return Err(Error::InvalidSource { path }),
            38 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if filter_reference.replace(field.payload()).is_some() {
                    return Err(Error::InvalidSource { path });
                }
                None
            },
            38 => return Err(Error::InvalidSource { path }),
            47 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if !inert_merge_owner(field.payload()) {
                    return Err(Error::UnsupportedDependency {
                        path,
                        kind: DependencyKind::Merge,
                    });
                }
                None
            },
            47 => return Err(Error::InvalidSource { path }),
            70 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                validate_inert_hidden_states_owner(source, field.payload(), path)?;
                None
            },
            70 => return Err(Error::InvalidSource { path }),
            81 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if !inactive_deprecated_category_owner(field.payload(), path)? {
                    return Err(Error::UnsupportedDependency {
                        path,
                        kind: DependencyKind::Category,
                    });
                }
                None
            },
            81 => return Err(Error::InvalidSource { path }),
            84 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if haunted_seen {
                    return Err(Error::InvalidSource { path });
                }
                haunted_seen = true;
                haunted = Some(require_owner_marker(
                    parse_single_uuid_owner(field.payload()),
                    DependencyKind::FormulaCache,
                    path,
                )?);
                None
            },
            84 => return Err(Error::InvalidSource { path }),
            86 if field.wire_type() == 2 => {
                field
                    .validate_canonical_framing()
                    .map_err(|_error| Error::InvalidSource { path })?;
                if category_reference.replace(field.payload()).is_some() {
                    return Err(Error::InvalidSource { path });
                }
                None
            },
            86 => return Err(Error::InvalidSource { path }),
            _ => None,
        };
        if let Some(kind) = kind {
            return Err(Error::UnsupportedDependency { path, kind });
        }
    }
    if model.pivot_owner().is_some() {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::Pivot,
        });
    }
    if let Some(reference) = filter_reference {
        validate_inert_filter_set(
            source,
            native,
            reference,
            &[38],
            DependencyKind::Pivot,
            path,
        )?;
    }
    if let Some(reference) = category_reference {
        let active =
            table_headers::dependencies::category_owner_reference_active(source, native, reference)
                .map_err(map_header_failure)?;
        if active {
            return Err(Error::UnsupportedDependency {
                path,
                kind: DependencyKind::Category,
            });
        }
    } else if model.category_owner().is_some() {
        return Err(Error::InvalidSource { path });
    }

    let spill = match model.spill_owner() {
        None => None,
        Some(owner) => Some(require_owner_marker(
            parse_spill_owner(owner),
            DependencyKind::Spill,
            path,
        )?),
    };
    let conditional_style = match model.conditional_style_formula_owner_id() {
        None => None,
        Some(owner) => Some(require_owner_marker(
            parse_cfuuid(owner),
            DependencyKind::ConditionalStyle,
            path,
        )?),
    };
    Ok(ModelOwnerMarkers {
        spill,
        conditional_style,
        haunted,
        sort_rule_tracker,
    })
}

fn inactive_deprecated_category_owner(source: &[u8], path: Path) -> Result<bool, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(source)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut owner = None;
    for field in view.fields() {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return Ok(false);
        }
        match field.number() {
            1 if owner.is_none() => owner = parse_uuid(field.payload()),
            2 => {},
            _ => return Ok(false),
        }
    }
    if owner.is_none() {
        return Ok(false);
    }
    let active = table_headers::dependencies::deprecated_category_grouping_active(source)
        .map_err(map_header_failure)?;
    Ok(!active)
}

fn parse_single_uuid_owner(source: &[u8]) -> Option<UuidKey> {
    let view = litchi_iwa_common::wire::WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let owner = fields.next()?;
    if fields.next().is_some()
        || owner.number() != 1
        || owner.wire_type() != 2
        || owner.validate_canonical_framing().is_err()
    {
        return None;
    }
    parse_uuid(owner.payload())
}

fn validate_inert_filter_set(
    source: &Package,
    native: table_headers::Target,
    reference: &[u8],
    declared_path: &[u32],
    kind: DependencyKind,
    path: Path,
) -> Result<(), Error> {
    let identifier = table_headers::resolve::local_reference_identifier(reference)
        .map_err(map_header_failure)?;
    let model_object = object_at_native_model(source, native, path)?;
    table_headers::resolve::require_declared_reference(
        model_object,
        native.message_index,
        identifier,
        declared_path,
    )
    .map_err(map_header_failure)?;
    let (_resolved, _index, payload) = resolve_exact_message(source, identifier, &[6_220], path)?;
    validate_inert_filter_payload(payload, kind, path)
}

fn validate_inert_filter_reference(
    source: &Package,
    reference: &[u8],
    path: Path,
) -> Result<(), Error> {
    let identifier = table_headers::resolve::local_reference_identifier(reference)
        .map_err(map_header_failure)?;
    let (_resolved, _index, payload) = resolve_exact_message(source, identifier, &[6_220], path)?;
    validate_inert_filter_payload(payload, DependencyKind::HiddenState, path)
}

fn validate_inert_filter_payload(
    payload: &[u8],
    kind: DependencyKind,
    path: Path,
) -> Result<(), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut seen = [false; 4];
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::InvalidSource { path })?;
        let slot = match field.number() {
            1 if field.wire_type() == 0 => 0,
            2 if field.wire_type() == 0 => 1,
            4 if field.wire_type() == 0 => 2,
            5 if field.wire_type() == 0 => 3,
            _ => {
                return Err(Error::UnsupportedDependency { path, kind });
            },
        };
        if seen[slot] {
            return Err(Error::InvalidSource { path });
        }
        seen[slot] = true;
        let (value, consumed) =
            litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                .map_err(|_error| Error::InvalidSource { path })?;
        if consumed != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != consumed
            || value != 0
        {
            return Err(Error::InvalidSource { path });
        }
    }
    if !seen[0] || !seen[1] || !seen[3] {
        return Err(Error::UnsupportedDependency { path, kind });
    }
    Ok(())
}

fn canonical_u64(source: &[u8], path: Path) -> Result<u64, Error> {
    let (value, consumed) = litchi_iwa_common::varint::decode_varint_from_bytes(source)
        .map_err(|_error| Error::InvalidSource { path })?;
    if consumed != source.len() || litchi_iwa_common::varint::encoded_len(value) != consumed {
        return Err(Error::InvalidSource { path });
    }
    Ok(value)
}

fn validate_inert_hidden_states_owner(
    source: &Package,
    payload: &[u8],
    path: Path,
) -> Result<(), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut owner_uid = false;
    let mut hidden_states = 0usize;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::InvalidSource { path })?;
        match field.number() {
            1 if field.wire_type() == 2 && !owner_uid => {
                owner_uid = parse_uuid(field.payload()).is_some();
                if !owner_uid {
                    return Err(Error::InvalidSource { path });
                }
            },
            2 if field.wire_type() == 2 => {
                hidden_states = hidden_states
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
                validate_inert_hidden_states(source, field.payload(), path)?;
            },
            _ => return Err(Error::InvalidSource { path }),
        }
    }
    if !owner_uid || hidden_states == 0 {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn validate_inert_hidden_states(source: &Package, payload: &[u8], path: Path) -> Result<(), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut uid = false;
    let mut column = false;
    let mut row = false;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::InvalidSource { path })?;
        match field.number() {
            1 if field.wire_type() == 2 && !uid => {
                uid = parse_uuid(field.payload()).is_some();
                if !uid {
                    return Err(Error::InvalidSource { path });
                }
            },
            2 if field.wire_type() == 2 && !column => {
                column = true;
                validate_inert_hidden_extent(source, field.payload(), 0, path)?;
            },
            3 if field.wire_type() == 2 && !row => {
                row = true;
                validate_inert_hidden_extent(source, field.payload(), 1, path)?;
            },
            _ => return Err(Error::InvalidSource { path }),
        }
    }
    if !uid || !column || !row {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn validate_inert_hidden_extent(
    source: &Package,
    payload: &[u8],
    expected_direction: u64,
    path: Path,
) -> Result<(), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut uid = false;
    let mut direction = false;
    let mut import_flag = false;
    let mut filter = false;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::InvalidSource { path })?;
        match field.number() {
            1 if field.wire_type() == 2 && !uid => {
                uid = parse_uuid(field.payload()).is_some();
                if !uid {
                    return Err(Error::InvalidSource { path });
                }
            },
            3 if field.wire_type() == 0 && !direction => {
                direction = true;
                if canonical_u64(field.payload(), path)? != expected_direction {
                    return Err(Error::InvalidSource { path });
                }
            },
            6 if field.wire_type() == 0 && !import_flag => {
                import_flag = true;
                if canonical_u64(field.payload(), path)? != 0 {
                    return Err(Error::UnsupportedDependency {
                        path,
                        kind: DependencyKind::HiddenState,
                    });
                }
            },
            8 if field.wire_type() == 2 && !filter => {
                filter = true;
                validate_inert_filter_reference(source, field.payload(), path)?;
            },
            _ => {
                return Err(Error::UnsupportedDependency {
                    path,
                    kind: DependencyKind::HiddenState,
                });
            },
        }
    }
    if !uid || !direction {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn inert_merge_owner(source: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(source) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(owner_uid) = fields.next() else {
        return false;
    };
    if owner_uid.number() != 1
        || owner_uid.wire_type() != 2
        || owner_uid.validate_canonical_framing().is_err()
        || parse_cfuuid(owner_uid.payload()).is_none()
    {
        return false;
    }
    let Some(store) = fields.next() else {
        return true;
    };
    if fields.next().is_some()
        || store.number() != 2
        || store.wire_type() != 2
        || store.validate_canonical_framing().is_err()
    {
        return false;
    }
    let Ok(store_view) = litchi_iwa_common::wire::WireView::parse(store.payload()) else {
        return false;
    };
    let mut store_fields = store_view.fields();
    let Some(next_index) = store_fields.next() else {
        return false;
    };
    next_index.number() == 2
        && next_index.wire_type() == 0
        && next_index.validate_canonical_key().is_ok()
        && litchi_iwa_common::varint::decode_varint_from_bytes(next_index.payload()).is_ok_and(
            |(value, consumed)| {
                consumed == next_index.payload().len()
                    && litchi_iwa_common::varint::encoded_len(value) == consumed
                    && value == 0
            },
        )
        && store_fields.next().is_none()
}

fn inert_single_uuid_owner(source: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(source) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(owner_uid) = fields.next() else {
        return false;
    };
    owner_uid.number() == 1
        && owner_uid.wire_type() == 2
        && owner_uid.validate_canonical_framing().is_ok()
        && parse_cfuuid(owner_uid.payload()).is_some()
        && fields.next().is_none()
}

fn inert_sort_order(source: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(source) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(sort_type) = fields.next() else {
        return false;
    };
    sort_type.number() == 1
        && sort_type.wire_type() == 0
        && sort_type.validate_canonical_key().is_ok()
        && litchi_iwa_common::varint::decode_varint_from_bytes(sort_type.payload()).is_ok_and(
            |(value, consumed)| {
                consumed == sort_type.payload().len()
                    && litchi_iwa_common::varint::encoded_len(value) == consumed
                    && value == 0
            },
        )
        && fields.next().is_none()
}

fn parse_sort_rule_tracker(source: &[u8]) -> Option<u64> {
    let view = litchi_iwa_common::wire::WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let reference = fields.next()?;
    if reference.number() != 1
        || reference.wire_type() != 2
        || reference.validate_canonical_framing().is_err()
        || fields.next().is_some()
    {
        return None;
    }
    table_headers::resolve::local_reference_identifier(reference.payload()).ok()
}

fn require_owner_marker(
    marker: Option<UuidKey>,
    kind: DependencyKind,
    path: Path,
) -> Result<UuidKey, Error> {
    marker.ok_or(Error::UnsupportedDependency { path, kind })
}

fn parse_spill_owner(source: &[u8]) -> Option<UuidKey> {
    let view = litchi_iwa_common::wire::WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let owner = fields.next()?;
    if fields.next().is_some()
        || owner.number() != 1
        || owner.wire_type() != 2
        || owner.validate_canonical_framing().is_err()
    {
        return None;
    }
    parse_uuid(owner.payload())
}

fn parse_uuid(source: &[u8]) -> Option<UuidKey> {
    let values = parse_varint_fields(source, [1, 2])?;
    Some(UuidKey {
        lower: values[0],
        upper: values[1],
    })
}

fn parse_cfuuid(source: &[u8]) -> Option<UuidKey> {
    let view = litchi_iwa_common::wire::WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let first = fields.next()?;
    if first.number() == 1 {
        if fields.next().is_some()
            || first.wire_type() != 2
            || first.validate_canonical_framing().is_err()
            || first.payload().len() != 16
        {
            return None;
        }
        let bytes: [u8; 16] = first.payload().try_into().ok()?;
        let value = u128::from_be_bytes(bytes);
        return Some(UuidKey {
            lower: value as u64,
            upper: (value >> 64) as u64,
        });
    }
    let values = parse_varint_field_iter(first, fields, [2, 3, 4, 5])?;
    let words = [
        u32::try_from(values[0]).ok()?,
        u32::try_from(values[1]).ok()?,
        u32::try_from(values[2]).ok()?,
        u32::try_from(values[3]).ok()?,
    ];
    Some(UuidKey {
        lower: u64::from(words[0]) | (u64::from(words[1]) << 32),
        upper: u64::from(words[2]) | (u64::from(words[3]) << 32),
    })
}

fn parse_varint_fields<const N: usize>(source: &[u8], expected: [u32; N]) -> Option<[u64; N]> {
    let view = litchi_iwa_common::wire::WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let first = fields.next()?;
    parse_varint_field_iter(first, fields, expected)
}

fn parse_varint_field_iter<'a, const N: usize>(
    first: litchi_iwa_common::wire::WireFieldView<'a>,
    fields: impl Iterator<Item = litchi_iwa_common::wire::WireFieldView<'a>>,
    expected: [u32; N],
) -> Option<[u64; N]> {
    let mut fields = core::iter::once(first).chain(fields);
    let mut values = [0u64; N];
    for (index, expected_number) in expected.into_iter().enumerate() {
        let field = fields.next()?;
        if field.number() != expected_number
            || field.wire_type() != 0
            || field.validate_canonical_key().is_err()
        {
            return None;
        }
        let (value, consumed) =
            litchi_iwa_common::varint::decode_varint_from_bytes(field.payload()).ok()?;
        if consumed != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != consumed
        {
            return None;
        }
        values[index] = value;
    }
    (fields.next().is_none()).then_some(values)
}

fn parse_owner_id_map(
    source: &[u8],
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(Vec<OwnerIdentity>, usize), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(source)
        .map_err(|_error| Error::InvalidSource { path })?;
    let count = view.fields().try_fold(0usize, |count, field| {
        if field.number() != 1
            || field.wire_type() != 2
            || field.validate_canonical_framing().is_err()
        {
            return Err(Error::InvalidSource { path });
        }
        count.checked_add(1).ok_or(Error::InvalidSource { path })
    })?;
    let mut identities = budget.reserve_scratch(count, LimitKind::References)?;
    let mut fields = count;
    for field in view.fields() {
        let entry = litchi_iwa_common::wire::WireView::parse(field.payload())
            .map_err(|_error| Error::InvalidSource { path })?;
        let mut entry_fields = entry.fields();
        let internal_field = entry_fields.next().ok_or(Error::InvalidSource { path })?;
        let uid_field = entry_fields.next().ok_or(Error::InvalidSource { path })?;
        if entry_fields.next().is_some()
            || internal_field.number() != 1
            || internal_field.wire_type() != 0
            || internal_field.validate_canonical_key().is_err()
            || uid_field.number() != 2
            || uid_field.wire_type() != 2
            || uid_field.validate_canonical_framing().is_err()
        {
            return Err(Error::InvalidSource { path });
        }
        let internal = u32::try_from(canonical_u64(internal_field.payload(), path)?)
            .ok()
            .filter(|value| *value != 0)
            .ok_or(Error::InvalidSource { path })?;
        let uid = parse_cfuuid(uid_field.payload()).ok_or(Error::InvalidSource { path })?;
        if identities
            .iter()
            .any(|identity: &OwnerIdentity| identity.internal == internal || identity.uid == uid)
        {
            return Err(Error::InvalidSource { path });
        }
        identities.push(OwnerIdentity { uid, internal });
        fields = fields.checked_add(2).ok_or(Error::InvalidSource { path })?;
    }
    Ok((identities, fields))
}

fn reject_unresolved_owner_markers(markers: ModelOwnerMarkers, path: Path) -> Result<(), Error> {
    for (present, kind) in [
        (
            markers.conditional_style.is_some(),
            DependencyKind::ConditionalStyle,
        ),
        (markers.spill.is_some(), DependencyKind::Spill),
        (markers.haunted.is_some(), DependencyKind::FormulaCache),
    ] {
        if present {
            return Err(Error::UnsupportedDependency { path, kind });
        }
    }
    Ok(())
}

fn spill_sizes_active(payload: &[u8], path: Path) -> Result<bool, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload).map_err(|_error| {
        Error::UnsupportedDependency {
            path,
            kind: DependencyKind::Spill,
        }
    })?;
    let mut active = false;
    for field in view.fields() {
        if field.number() != 1
            || field.wire_type() != 2
            || field.validate_canonical_framing().is_err()
        {
            return Err(Error::UnsupportedDependency {
                path,
                kind: DependencyKind::Spill,
            });
        }
        active = true;
    }
    Ok(active)
}

fn validate_empty_formula_owner_closure(
    owner: dependency_codec::FormulaOwnerDependenciesSnapshot<'_>,
    count: DependencyCount,
    expected_kind: u32,
    kind: DependencyKind,
    path: Path,
) -> Result<(), Error> {
    let expected_cell_tiles = usize::from(expected_kind == 35);
    let empty = owner.owner_kind() == Some(expected_kind)
        && owner.base_owner_uid().is_some()
        && owner.formula_owner().is_none()
        && count.cell_tiles == expected_cell_tiles
        && count.range_tiles == 0
        && owner.cell_dependencies().is_none_or(empty_message)
        && owner.range_dependencies().is_none_or(empty_message)
        && owner
            .volatile_dependencies()
            .is_none_or(empty_volatile_dependencies)
        && owner
            .spanning_column_dependencies()
            .is_none_or(empty_spanning_dependencies)
        && owner
            .spanning_row_dependencies()
            .is_none_or(empty_spanning_dependencies)
        && owner
            .whole_owner_dependencies()
            .is_none_or(empty_whole_owner_dependencies)
        && owner.cell_errors().is_none_or(empty_message)
        && (expected_kind == 35 || owner.tiled_cell_dependencies().is_none_or(empty_message))
        && owner.uuid_references().is_none_or(empty_message)
        && owner.tiled_range_dependencies().is_none_or(empty_message)
        && owner
            .spill_range_sizes()
            .is_none_or(|payload| spill_sizes_active(payload, path).is_ok_and(|active| !active));
    if !empty {
        return Err(Error::UnsupportedDependency { path, kind });
    }
    Ok(())
}

fn validate_selected_formula_owner_closure(
    owner: dependency_codec::FormulaOwnerDependenciesSnapshot<'_>,
    native: table_headers::Target,
    stage: &DependencyStage,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(), Error> {
    let inert_graph = stage.cell_tiles.is_empty()
        && stage.range_tiles.is_empty()
        && stage.cell_records.is_empty()
        && stage.edge_components.is_empty();
    let envelope = owner.owner_kind() == Some(1)
        && owner.base_owner_uid().is_none()
        && owner.formula_owner().is_some()
        // Inline range records are a supported part of the selected table's
        // formula graph. The strict dependency decoder has already validated
        // their wire/Buffa shape; cache planning later binds their target
        // owner and rectangle against the complete table registry.
        && owner
            .volatile_dependencies()
            .is_none_or(empty_volatile_dependencies)
        && owner.spanning_column_dependencies().is_none_or(|payload| {
            selected_spanning_dependencies(payload, native, path)
                || (inert_graph && empty_message(payload))
        })
        && owner.spanning_row_dependencies().is_none_or(|payload| {
            selected_spanning_dependencies(payload, native, path)
                || (inert_graph && empty_message(payload))
        })
        && owner
            .whole_owner_dependencies()
            .is_none_or(empty_whole_owner_dependencies)
        && owner.cell_errors().is_none_or(empty_message)
        && owner.uuid_references().is_none_or(empty_message)
        && owner
            .spill_range_sizes()
            .is_none_or(|payload| spill_sizes_active(payload, path).is_ok_and(|active| !active));
    if !envelope {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::FormulaCache,
        });
    }
    validate_selected_cell_dependencies(owner.cell_dependencies(), stage, native, path)?;
    budget.charge_auxiliary_work(
        stage
            .cell_records
            .len()
            .checked_add(stage.edge_components.len())
            .ok_or(Error::InvalidSource { path })?,
    )?;
    Ok(())
}

fn validate_selected_cell_dependencies(
    payload: Option<&[u8]>,
    stage: &DependencyStage,
    native: table_headers::Target,
    path: Path,
) -> Result<(), Error> {
    let record_count = match payload {
        None => 0,
        Some(payload) => {
            validate_selected_cell_dependency_wire(payload, &stage.cell_records, path)?
        },
    };
    if record_count != stage.cell_records.len() {
        return Err(Error::InvalidSource { path });
    }
    validate_selected_cell_graph(stage, native, path)
}

fn validate_selected_cell_graph(
    stage: &DependencyStage,
    native: table_headers::Target,
    path: Path,
) -> Result<(), Error> {
    let mut expected_start = 0usize;
    for record in &stage.cell_records {
        if record.component_start != expected_start
            || record.component_end < record.component_start
            || record.component_end > stage.edge_components.len()
            || record.column >= native.columns
            || record.row >= native.rows
            || record.dirty_self_plus_precedents_count.is_some()
            || record.is_in_a_cycle.is_some()
            || record.has_calculated_precedents.is_some()
        {
            return Err(Error::InvalidSource { path });
        }
        let components = &stage.edge_components[record.component_start..record.component_end];
        let Some(edges) = record.edges else {
            if !components.is_empty() {
                return Err(Error::InvalidSource { path });
            }
            expected_start = record.component_end;
            continue;
        };
        let local = edges.local();
        let external = edges.external();
        let expected_components = local
            .checked_mul(2)
            .and_then(|count| {
                external
                    .checked_mul(3)
                    .and_then(|external| count.checked_add(external))
            })
            .ok_or(Error::InvalidSource { path })?;
        if components.len() != expected_components {
            return Err(Error::InvalidSource { path });
        }
        let mut offset = 0usize;
        for (kind, count) in [
            (dependency_codec::ExpandedEdgeKind::LocalRow, local),
            (dependency_codec::ExpandedEdgeKind::LocalColumn, local),
            (dependency_codec::ExpandedEdgeKind::ExternalRow, external),
            (dependency_codec::ExpandedEdgeKind::ExternalColumn, external),
            (dependency_codec::ExpandedEdgeKind::InternalOwner, external),
        ] {
            for index in 0..count {
                let component = components
                    .get(offset)
                    .ok_or(Error::InvalidSource { path })?;
                if component.kind() != kind
                    || component.index() != index
                    || matches!(
                        kind,
                        dependency_codec::ExpandedEdgeKind::LocalRow
                            if component.value() >= native.rows
                    )
                    || matches!(
                        kind,
                        dependency_codec::ExpandedEdgeKind::LocalColumn
                            if component.value() >= native.columns
                    )
                {
                    return Err(Error::InvalidSource { path });
                }
                offset = offset.checked_add(1).ok_or(Error::InvalidSource { path })?;
            }
        }
        if offset != components.len() {
            return Err(Error::InvalidSource { path });
        }
        expected_start = record.component_end;
    }
    if expected_start != stage.edge_components.len() {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn validate_selected_cell_dependency_wire(
    payload: &[u8],
    records: &[SelectedCellRecord],
    path: Path,
) -> Result<usize, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut count = 0usize;
    for (field, record) in view.fields().zip(records.iter()) {
        if field.number() != 1
            || field.wire_type() != 2
            || field.validate_canonical_framing().is_err()
            || !selected_cell_record_wire(field.payload(), *record, path)?
        {
            return Err(Error::InvalidSource { path });
        }
        count = count.checked_add(1).ok_or(Error::InvalidSource { path })?;
    }
    if view.fields().count() != records.len() {
        return Err(Error::InvalidSource { path });
    }
    Ok(count)
}

fn selected_cell_record_wire(
    payload: &[u8],
    record: SelectedCellRecord,
    path: Path,
) -> Result<bool, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut column = false;
    let mut row = false;
    let mut edges = false;
    for field in view.fields() {
        match field.number() {
            1 if !column && field.wire_type() == 0 => {
                column = canonical_u64(field.payload(), path)? == u64::from(record.column);
            },
            2 if !row && field.wire_type() == 0 => {
                row = canonical_u64(field.payload(), path)? == u64::from(record.row);
            },
            6 if !edges && field.wire_type() == 2 && field.validate_canonical_framing().is_ok() => {
                edges = selected_expanded_edges_wire(field.payload());
            },
            _ => return Ok(false),
        }
    }
    Ok(column && row && edges == record.edges.is_some())
}

fn selected_expanded_edges_wire(payload: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(payload) else {
        return false;
    };
    view.fields().all(|field| {
        matches!(field.number(), 1..=5)
            && field.validate_canonical_key().is_ok()
            && match field.wire_type() {
                0 => canonical_varint_bytes(field.payload()),
                2 => {
                    field.validate_canonical_framing().is_ok()
                        && canonical_packed_varints(field.payload())
                },
                _ => false,
            }
    })
}

fn canonical_varint_bytes(payload: &[u8]) -> bool {
    let Ok((value, consumed)) = litchi_iwa_common::varint::decode_varint_from_bytes(payload) else {
        return false;
    };
    consumed == payload.len() && litchi_iwa_common::varint::encoded_len(value) == consumed
}

fn canonical_packed_varints(mut payload: &[u8]) -> bool {
    while !payload.is_empty() {
        let Ok((value, consumed)) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        else {
            return false;
        };
        if consumed == 0 || litchi_iwa_common::varint::encoded_len(value) != consumed {
            return false;
        }
        payload = &payload[consumed..];
    }
    true
}

fn selected_spanning_dependencies(
    payload: &[u8],
    native: table_headers::Target,
    path: Path,
) -> bool {
    let Some(last_column) = native.columns.checked_sub(1) else {
        return false;
    };
    let Some(last_row) = native.rows.checked_sub(1) else {
        return false;
    };
    let Ok(header_columns) = u32::try_from(native.settings.header_column_count()) else {
        return false;
    };
    let Ok(header_rows) = u32::try_from(native.settings.header_row_count()) else {
        return false;
    };
    let Ok(footer_rows) = u32::try_from(native.settings.footer_row_count()) else {
        return false;
    };
    let Some(body_last_row) = native
        .rows
        .checked_sub(footer_rows)
        .and_then(|rows| rows.checked_sub(1))
    else {
        return false;
    };
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(payload) else {
        return false;
    };
    let mut total = false;
    let mut body = false;
    for field in view.fields() {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return false;
        }
        let expected = match field.number() {
            2 if !total => {
                total = true;
                [0, 0, u64::from(last_column), u64::from(last_row)]
            },
            3 if !body => {
                body = true;
                [
                    u64::from(header_columns),
                    u64::from(header_rows),
                    u64::from(last_column),
                    u64::from(body_last_row),
                ]
            },
            _ => return false,
        };
        if parse_varint_fields(field.payload(), [1, 2, 3, 4]) != Some(expected) {
            return false;
        }
    }
    let _ = path;
    total && body
}

fn empty_message(payload: &[u8]) -> bool {
    litchi_iwa_common::wire::WireView::parse(payload)
        .is_ok_and(|view| view.fields().next().is_none())
}

fn empty_volatile_dependencies(payload: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(payload) else {
        return false;
    };
    let mut seen = [false; 7];
    for field in view.fields() {
        let Ok(index) = usize::try_from(field.number()) else {
            return false;
        };
        if !matches!(index, 1..=5 | 7)
            || seen[index - 1]
            || field.wire_type() != 2
            || field.validate_canonical_framing().is_err()
            || !empty_message(field.payload())
        {
            return false;
        }
        seen[index - 1] = true;
    }
    true
}

fn empty_whole_owner_dependencies(payload: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(payload) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(field) = fields.next() else {
        return true;
    };
    field.number() == 1
        && field.wire_type() == 2
        && field.validate_canonical_framing().is_ok()
        && empty_message(field.payload())
        && fields.next().is_none()
}

fn empty_spanning_dependencies(payload: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(payload) else {
        return false;
    };
    let mut total = false;
    let mut body = false;
    for field in view.fields() {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return false;
        }
        match field.number() {
            2 if !total => {
                total = true;
                if !sentinel_range(field.payload()) {
                    return false;
                }
            },
            3 if !body => {
                body = true;
                if !sentinel_range(field.payload()) {
                    return false;
                }
            },
            _ => return false,
        }
    }
    total && body
}

fn sentinel_range(payload: &[u8]) -> bool {
    parse_varint_fields(payload, [1, 2, 3, 4])
        == Some([32_767, 2_147_483_647, 32_767, 2_147_483_647])
}

#[allow(clippy::too_many_arguments, reason = "one dormant hidden-owner proof")]
fn validate_dormant_hidden_owner(
    source: &Package,
    model_object: &ArchiveObject,
    model_message: usize,
    reference: Option<storage_codec::ReferenceSnapshot>,
    field: u32,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Option<MessageRoute>, Error> {
    let Some(reference) = reference else {
        return Ok(None);
    };
    let id = checked_reference(reference, path)?;
    identities.claim(id, Role::HiddenStateOwner, true, budget, path)?;
    table_headers::resolve::require_declared_reference(model_object, model_message, id, &[field])
        .map_err(map_header_failure)?;
    let (resolved, index, payload) =
        resolve_exact_message(source, id, &[HIDDEN_STATE_OWNER_MESSAGE_TYPE], path)?;
    if !inert_hidden_state_owner(payload) {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::HiddenState,
        });
    }
    budget.charge_wire_scan(payload.len(), 6, 2)?;
    Ok(Some(MessageRoute::new(
        resolved,
        index,
        HIDDEN_STATE_OWNER_MESSAGE_TYPE,
    )))
}

#[allow(clippy::too_many_arguments, reason = "one dormant tracker proof")]
fn resolve_sort_rule_tracker(
    source: &Package,
    model_object: &ArchiveObject,
    model_message: usize,
    identifier: Option<u64>,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Option<MessageRoute>, Error> {
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    identities.claim(identifier, Role::SortRuleTracker, false, budget, path)?;
    table_headers::resolve::require_declared_reference(
        model_object,
        model_message,
        identifier,
        &[45, 1],
    )
    .map_err(map_header_failure)?;
    let (resolved, index, payload) =
        resolve_exact_message(source, identifier, &[REFERENCE_TRACKER_MESSAGE_TYPE], path)?;
    if !inert_single_uuid_owner(payload) {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::FormulaCache,
        });
    }
    budget.charge_wire_scan(payload.len(), 6, 2)?;
    Ok(Some(MessageRoute::new(
        resolved,
        index,
        REFERENCE_TRACKER_MESSAGE_TYPE,
    )))
}

fn inert_hidden_state_owner(source: &[u8]) -> bool {
    let Ok(view) = litchi_iwa_common::wire::WireView::parse(source) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(identifier) = fields.next() else {
        return false;
    };
    let Some(active) = fields.next() else {
        return false;
    };
    if fields.next().is_some()
        || identifier.number() != 1
        || identifier.wire_type() != 2
        || identifier.validate_canonical_framing().is_err()
        || parse_cfuuid(identifier.payload()).is_none()
        || active.number() != 3
        || active.wire_type() != 0
        || active.validate_canonical_key().is_err()
    {
        return false;
    }
    matches!(
        litchi_iwa_common::varint::decode_varint_from_bytes(active.payload()),
        Ok((0, 1))
    )
}

fn resolve_merge_map(
    source: &Package,
    model_object: &ArchiveObject,
    model_message: usize,
    reference: Option<storage_codec::ReferenceSnapshot>,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Option<MessageRoute>, Error> {
    let Some(reference) = reference else {
        return Ok(None);
    };
    let id = checked_reference(reference, path)?;
    identities.claim(id, Role::Merge, false, budget, path)?;
    // As with native formula-owner edges, Apple may omit FieldInfo while
    // retaining exactly one message-level object reference.
    require_declared(model_object, model_message, id, &[4, 13], path)?;
    let (resolved, index, payload) =
        resolve_exact_message(source, id, &DATA_LIST_MESSAGE_TYPES, path)?;
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut active = false;
    let mut field_count = 0usize;
    for field in view.fields() {
        field_count = field_count
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
        if field.number() != 1 || field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::InvalidSource { path })?;
        // Each field 1 is a CellRange. Scalar publication cannot safely
        // regenerate an active merge map or distinguish covered followers.
        active = true;
    }
    budget.charge_wire_scan(payload.len(), field_count, 1)?;
    if active {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::Merge,
        });
    }
    Ok(Some(MessageRoute::new(
        resolved,
        index,
        resolved.messages[index].type_,
    )))
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
struct StorageCount {
    tiles: usize,
    header_buckets: usize,
    overflowed: bool,
}

impl StorageCount {
    fn validate(&self, path: Path) -> Result<(), Error> {
        count_not_overflowed(self.overflowed, path)
    }
}

impl storage_codec::StorageVisitor for StorageCount {
    fn visit_tile_reference(
        &mut self,
        _record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.tiles.checked_add(1) {
            self.tiles = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }

    fn visit_header_bucket(
        &mut self,
        _record: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.header_buckets.checked_add(1) {
            self.header_buckets = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

struct StorageStage {
    tiles: Vec<(u32, storage_codec::ReferenceSnapshot)>,
    header_buckets: Vec<storage_codec::ReferenceSnapshot>,
}

impl StorageStage {
    fn with_capacity(
        count: StorageCount,
        budget: &mut ResolutionBudget,
        _path: Path,
    ) -> Result<Self, Error> {
        Ok(Self {
            tiles: budget.reserve_scratch(count.tiles, LimitKind::References)?,
            header_buckets: budget.reserve_scratch(count.header_buckets, LimitKind::References)?,
        })
    }

    fn finish(&self, count: StorageCount, path: Path) -> Result<(), Error> {
        if self.tiles.len() != count.tiles || self.header_buckets.len() != count.header_buckets {
            return Err(Error::InvalidSource { path });
        }
        Ok(())
    }
}

impl storage_codec::StorageVisitor for StorageStage {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.tiles.push((record.tile_id(), record.reference()));
        Ok(())
    }

    fn visit_header_bucket(
        &mut self,
        record: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.header_buckets.push(record.reference());
        Ok(())
    }
}

fn resolve_tiles(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    records: &[(u32, storage_codec::ReferenceSnapshot)],
    tile_size: u32,
    native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Vec<TileRoute>, Error> {
    let tile_count = native.rows.div_ceil(tile_size);
    let mut routes: Vec<TileRoute> = budget.reserve_retained(records.len(), LimitKind::Objects)?;
    for &(tile_id, reference) in records {
        if tile_id >= tile_count || routes.iter().any(|route| route.tile_id == tile_id) {
            return Err(Error::InvalidSource { path });
        }
        let id = checked_reference(reference, path)?;
        identities.claim(id, Role::Tile, false, budget, path)?;
        require_declared(owner, owner_message, id, &[4, 3, 1, 2], path)?;
        let (resolved, message_index, payload) =
            resolve_exact_message(source, id, &[TILE_MESSAGE_TYPE], path)?;
        let mut row_count = TileRowCount::default();
        let (tile, report) = storage_codec::decode_tile_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut row_count,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        row_count.validate(path)?;
        let mut row_stage = TileRowStage {
            rows: budget.reserve_scratch(row_count.rows, LimitKind::RetainedElements)?,
        };
        let (tile_again, report) = storage_codec::decode_tile_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut row_stage,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if tile != tile_again || row_stage.rows.len() != row_count.rows {
            return Err(Error::InvalidSource { path });
        }
        validate_tile(tile, &mut row_stage.rows, tile_id, tile_size, native, path)?;
        budget.release_scratch(&row_stage.rows)?;
        routes.push(TileRoute {
            tile_id,
            message: MessageRoute::new(resolved, message_index, TILE_MESSAGE_TYPE),
        });
    }
    Ok(routes)
}

fn validate_tile(
    tile: storage_codec::TileSnapshot,
    rows: &mut [TileRowProof],
    tile_id: u32,
    tile_size: u32,
    native: table_headers::Target,
    path: Path,
) -> Result<(), Error> {
    let origin = tile_id
        .checked_mul(tile_size)
        .ok_or(Error::InvalidSource { path })?;
    let remaining = native
        .rows
        .checked_sub(origin)
        .ok_or(Error::InvalidSource { path })?;
    rows.sort_unstable_by_key(|row| row.index);
    let duplicate_rows = rows.windows(2).any(|pair| pair[0].index == pair[1].index);
    let cell_count = rows.iter().try_fold(0u32, |total, row| {
        if row.index >= tile_size
            || origin
                .checked_add(row.index)
                .is_none_or(|value| value >= native.rows)
            || row.cells > native.columns
            || row.modern_storage != row.modern_offsets
        {
            return None;
        }
        total.checked_add(row.cells)
    });
    let modern = rows.iter().all(|row| row.modern_storage);
    if duplicate_rows
        || usize::try_from(tile.num_rows()).ok() != Some(rows.len())
        || (!modern && cell_count != Some(tile.num_cells()))
        || (modern && tile.num_cells() != 0)
        || tile.num_rows() > remaining.min(tile_size)
        || tile.num_cells()
            > native
                .columns
                .checked_mul(tile.num_rows())
                .ok_or(Error::InvalidSource { path })?
        || (!modern && tile.num_cells() != 0 && tile.max_column() >= native.columns)
        || (!modern && tile.num_rows() != 0 && tile.max_row() >= native.rows)
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

#[derive(Default)]
struct TileRowCount {
    rows: usize,
    overflowed: bool,
}

impl TileRowCount {
    fn validate(&self, path: Path) -> Result<(), Error> {
        count_not_overflowed(self.overflowed, path)
    }
}

impl storage_codec::StorageVisitor for TileRowCount {
    fn visit_tile_row(
        &mut self,
        _row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.rows.checked_add(1) {
            self.rows = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TileRowProof {
    index: u32,
    cells: u32,
    modern_storage: bool,
    modern_offsets: bool,
}

struct TileRowStage {
    rows: Vec<TileRowProof>,
}

impl storage_codec::StorageVisitor for TileRowStage {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.rows.push(TileRowProof {
            index: row.tile_row_index(),
            cells: row.cell_count(),
            modern_storage: row.cell_storage_buffer().is_some(),
            modern_offsets: row.cell_offsets().is_some(),
        });
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum HeaderAxis {
    Row,
    Column,
}

fn resolve_headers(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[storage_codec::ReferenceSnapshot],
    declared_path: &[u32],
    dimension: u32,
    axis: HeaderAxis,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Vec<MessageRoute>, Error> {
    let mut routes = budget.reserve_retained(references.len(), LimitKind::Objects)?;
    for &reference in references {
        routes.push(resolve_one_header(
            source,
            owner,
            owner_message,
            reference,
            declared_path,
            dimension,
            axis,
            identities,
            budget,
            path,
        )?);
    }
    Ok(routes)
}

fn resolve_one_header(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    reference: storage_codec::ReferenceSnapshot,
    declared_path: &[u32],
    dimension: u32,
    axis: HeaderAxis,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<MessageRoute, Error> {
    let id = checked_reference(reference, path)?;
    identities.claim(
        id,
        match axis {
            HeaderAxis::Row => Role::RowHeader,
            HeaderAxis::Column => Role::ColumnHeader,
        },
        false,
        budget,
        path,
    )?;
    require_declared(owner, owner_message, id, declared_path, path)?;
    let (resolved, message_index, payload) =
        resolve_exact_message(source, id, &[HEADER_BUCKET_MESSAGE_TYPE], path)?;
    let mut count = HeaderCount::default();
    let (_snapshot, report) = storage_codec::decode_header_storage_bucket_with_visitor(
        payload,
        budget.options(payload.len())?,
        &mut count,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    count.validate(path)?;
    let mut stage = HeaderStage {
        headers: budget.reserve_scratch(count.headers, LimitKind::RetainedElements)?,
    };
    let (_snapshot, report) = storage_codec::decode_header_storage_bucket_with_visitor(
        payload,
        budget.options(payload.len())?,
        &mut stage,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    if stage.headers.len() != count.headers {
        return Err(Error::InvalidSource { path });
    }
    stage.headers.sort_unstable();
    if stage.headers.windows(2).any(|pair| pair[0] == pair[1])
        || stage
            .headers
            .last()
            .is_some_and(|index| *index >= dimension)
    {
        return Err(Error::InvalidSource { path });
    }
    budget.release_scratch(&stage.headers)?;
    Ok(MessageRoute::new(
        resolved,
        message_index,
        HEADER_BUCKET_MESSAGE_TYPE,
    ))
}

#[derive(Default)]
struct HeaderCount {
    headers: usize,
    overflowed: bool,
}

impl HeaderCount {
    fn validate(&self, path: Path) -> Result<(), Error> {
        count_not_overflowed(self.overflowed, path)
    }
}

impl storage_codec::StorageVisitor for HeaderCount {
    fn visit_header(
        &mut self,
        _header: storage_codec::HeaderSnapshot,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.headers.checked_add(1) {
            self.headers = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

struct HeaderStage {
    headers: Vec<u32>,
}

impl storage_codec::StorageVisitor for HeaderStage {
    fn visit_header(
        &mut self,
        header: storage_codec::HeaderSnapshot,
    ) -> Result<(), storage_codec::DecodeError> {
        self.headers.push(header.index());
        Ok(())
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "all storage roles share one proof boundary"
)]
fn resolve_lists(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    store: storage_codec::DataStoreSnapshot<'_>,
    native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<ListRoutes, Error> {
    let required = |reference,
                    field,
                    kind,
                    role,
                    identities: &mut RoleIdentities,
                    budget: &mut ResolutionBudget| {
        resolve_list(
            source,
            owner,
            owner_message,
            reference,
            field,
            kind,
            role,
            native,
            identities,
            budget,
            path,
        )
    };
    let optional = |reference: Option<storage_codec::ReferenceSnapshot>,
                    field,
                    kind,
                    role,
                    identities: &mut RoleIdentities,
                    budget: &mut ResolutionBudget| {
        reference
            .map(|reference| {
                resolve_list(
                    source,
                    owner,
                    owner_message,
                    reference,
                    field,
                    kind,
                    role,
                    native,
                    identities,
                    budget,
                    path,
                )
            })
            .transpose()
    };
    Ok(ListRoutes {
        string: required(
            store.string_table(),
            4,
            ListType::String,
            Role::String,
            identities,
            budget,
        )?,
        style: required(
            store.style_table(),
            5,
            ListType::Style,
            Role::Style,
            identities,
            budget,
        )?,
        formula: required(
            store.formula_table(),
            6,
            ListType::Formula,
            Role::Formula,
            identities,
            budget,
        )?,
        format_pre_bnc: required(
            store.format_table_pre_bnc(),
            11,
            ListType::Format,
            Role::Format,
            identities,
            budget,
        )?,
        formula_error: optional(
            store.formula_error_table(),
            12,
            ListType::FormulaError,
            Role::FormulaError,
            identities,
            budget,
        )?,
        custom_format: optional(
            store.deprecated_custom_format_table(),
            15,
            ListType::CustomFormat,
            Role::CustomFormat,
            identities,
            budget,
        )?,
        multiple_choice: optional(
            store.multiple_choice_list_format_table(),
            16,
            ListType::MultipleChoiceListFormat,
            Role::MultipleChoice,
            identities,
            budget,
        )?,
        rich_text: optional(
            store.rich_text_table(),
            17,
            ListType::RichTextPayload,
            Role::RichText,
            identities,
            budget,
        )?,
        conditional_style: optional(
            store.conditional_style_table(),
            18,
            ListType::ConditionalStyle,
            Role::ConditionalStyle,
            identities,
            budget,
        )?,
        comment: optional(
            store.comment_storage_table(),
            19,
            ListType::CommentStorage,
            Role::Comment,
            identities,
            budget,
        )?,
        import_warning: optional(
            store.import_warning_set_table(),
            20,
            ListType::ImportWarning,
            Role::ImportWarning,
            identities,
            budget,
        )?,
        control_cell_spec: optional(
            store.control_cell_spec_table(),
            21,
            ListType::ControlCellSpec,
            Role::ControlCellSpec,
            identities,
            budget,
        )?,
        format: optional(
            store.format_table(),
            22,
            ListType::Format,
            Role::Format,
            identities,
            budget,
        )?,
    })
}

#[allow(clippy::too_many_arguments, reason = "one exact list-role proof")]
fn resolve_list(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    reference: storage_codec::ReferenceSnapshot,
    field: u32,
    expected_kind: ListType,
    role: Role,
    native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<ListRoute, Error> {
    let id = checked_reference(reference, path)?;
    identities.claim(id, role, role == Role::Format, budget, path)?;
    require_declared(owner, owner_message, id, &[4, field], path)?;
    let resolved = resolve_object(source, id, path)?;
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    let mut selected = None;
    for (index, message) in resolved.messages.iter().enumerate() {
        if !DATA_LIST_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        table_headers::resolve::validate_message_metadata(object, index)
            .map_err(map_header_failure)?;
        let (snapshot, report) = storage_codec::decode_table_data_list_with_report(
            &message.data,
            budget.options(message.data.len())?,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if snapshot.list_type() == expected_kind as i32 {
            if selected.replace((index, message.type_)).is_some() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    let (message_index, message_type) = selected.ok_or(Error::InvalidSource { path })?;
    let payload = &resolved.messages[message_index].data;
    let mut count = ListCount::default();
    let (_snapshot, report) = storage_codec::decode_table_data_list_with_visitor(
        payload,
        budget.options(payload.len())?,
        &mut count,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    count.validate(path)?;
    let mut stage = ListStage {
        segments: budget.reserve_scratch(count.segments, LimitKind::References)?,
        embedded_references: budget
            .reserve_scratch(count.embedded_references, LimitKind::References)?,
        entries: 0,
        overflowed: false,
    };
    let (_snapshot, report) = storage_codec::decode_table_data_list_with_visitor(
        payload,
        budget.options(payload.len())?,
        &mut stage,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    if stage.segments.len() != count.segments
        || stage.embedded_references.len() != count.embedded_references
        || stage.entries != count.entries
        || stage.overflowed
    {
        return Err(Error::InvalidSource { path });
    }
    validate_embedded_list_references(
        source,
        object,
        message_index,
        &stage.embedded_references,
        identities,
        budget,
        path,
    )?;
    let (segments, segment_entries) = resolve_list_segments(
        source,
        object,
        message_index,
        &stage.segments,
        expected_kind,
        native,
        identities,
        budget,
        path,
    )?;
    budget.release_scratch(&stage.segments)?;
    budget.release_scratch(&stage.embedded_references)?;
    Ok(ListRoute {
        message: MessageRoute::new(resolved, message_index, message_type),
        segments,
        entries: count
            .entries
            .checked_add(segment_entries)
            .ok_or(Error::InvalidSource { path })?,
    })
}

#[derive(Default)]
struct ListCount {
    segments: usize,
    embedded_references: usize,
    entries: usize,
    overflowed: bool,
}

impl ListCount {
    fn validate(&self, path: Path) -> Result<(), Error> {
        count_not_overflowed(self.overflowed, path)
    }
}

impl storage_codec::StorageVisitor for ListCount {
    fn visit_list_entry(
        &mut self,
        entry: storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.entries.checked_add(1) {
            self.entries = next;
        } else {
            self.overflowed = true;
        }
        for present in [
            entry.reference().is_some(),
            entry.rich_text_payload().is_some(),
            entry.comment_storage().is_some(),
        ] {
            if present {
                if let Some(next) = self.embedded_references.checked_add(1) {
                    self.embedded_references = next;
                } else {
                    self.overflowed = true;
                }
            }
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _record: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.segments.checked_add(1) {
            self.segments = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct EmbeddedReference {
    reference: storage_codec::ReferenceSnapshot,
    field: u32,
}

struct ListStage {
    segments: Vec<storage_codec::ReferenceSnapshot>,
    embedded_references: Vec<EmbeddedReference>,
    entries: usize,
    overflowed: bool,
}

impl storage_codec::StorageVisitor for ListStage {
    fn visit_list_entry(
        &mut self,
        entry: storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if let Some(next) = self.entries.checked_add(1) {
            self.entries = next;
        } else {
            self.overflowed = true;
        }
        for (reference, field) in [
            (entry.reference(), 4),
            (entry.rich_text_payload(), 9),
            (entry.comment_storage(), 10),
        ] {
            if let Some(reference) = reference {
                self.embedded_references
                    .push(EmbeddedReference { reference, field });
            }
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        record: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.segments.push(record.reference());
        Ok(())
    }
}

fn validate_embedded_list_references(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[EmbeddedReference],
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(), Error> {
    for &embedded in references {
        let id = checked_reference(embedded.reference, path)?;
        let role = if embedded.field == 9 {
            Role::RichPayload
        } else {
            Role::ListPayload
        };
        identities.claim(id, role, true, budget, path)?;
        require_declared(owner, owner_message, id, &[3, embedded.field], path)?;
        let _ = resolve_single_message(source, id, path)?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments, reason = "one segmented-list proof")]
fn resolve_list_segments(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[storage_codec::ReferenceSnapshot],
    expected_kind: ListType,
    _native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(Vec<MessageRoute>, usize), Error> {
    let mut routes = budget.reserve_retained(references.len(), LimitKind::Objects)?;
    let mut entries = 0usize;
    for &reference in references {
        let id = checked_reference(reference, path)?;
        identities.claim(id, Role::ListSegment, false, budget, path)?;
        require_declared(owner, owner_message, id, &[4], path)?;
        let (resolved, index, payload) =
            resolve_exact_message(source, id, &[DATA_LIST_SEGMENT_MESSAGE_TYPE], path)?;
        let mut count = ListCount::default();
        let (snapshot, report) = storage_codec::decode_table_data_list_segment_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut count,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        count.validate(path)?;
        if snapshot.list_type() != expected_kind as i32 || count.segments != 0 {
            return Err(Error::InvalidSource { path });
        }
        let mut stage = ListStage {
            segments: Vec::new(),
            embedded_references: budget
                .reserve_scratch(count.embedded_references, LimitKind::References)?,
            entries: 0,
            overflowed: false,
        };
        let (_snapshot, report) = storage_codec::decode_table_data_list_segment_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut stage,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if stage.entries != count.entries || stage.overflowed {
            return Err(Error::InvalidSource { path });
        }
        entries = entries
            .checked_add(count.entries)
            .ok_or(Error::InvalidSource { path })?;
        let segment_object = table_headers::resolve::resolved_object(source, resolved)
            .map_err(map_header_failure)?;
        validate_embedded_list_references(
            source,
            segment_object,
            index,
            &stage.embedded_references,
            identities,
            budget,
            path,
        )?;
        budget.release_scratch(&stage.embedded_references)?;
        routes.push(MessageRoute::new(
            resolved,
            index,
            DATA_LIST_SEGMENT_MESSAGE_TYPE,
        ));
    }
    Ok((routes, entries))
}

#[derive(Clone, Copy)]
struct RichEntrySeed {
    key: u32,
    ref_count: u32,
    payload: storage_codec::ReferenceSnapshot,
}

struct RichEntryCollector {
    entries: Vec<RichEntrySeed>,
    limit: usize,
    accepted: usize,
    invalid: bool,
}

impl RichEntryCollector {
    fn reserve_entry_slot(&mut self) -> bool {
        if self.accepted == self.limit {
            self.invalid = true;
            return false;
        }
        let Some(next) = self.accepted.checked_add(1) else {
            self.invalid = true;
            return false;
        };
        self.accepted = next;
        true
    }
}

impl storage_codec::StorageVisitor for RichEntryCollector {
    fn visit_list_entry(
        &mut self,
        entry: storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let Some(payload) = entry.rich_text_payload() else {
            self.invalid = true;
            return Ok(());
        };
        if entry.string_value().is_some()
            || entry.reference().is_some()
            || entry.formula().is_some()
            || entry.format().is_some()
            || entry.custom_format().is_some()
            || entry.comment_storage().is_some()
            || entry.import_warning_set().is_some()
            || entry.cell_spec().is_some()
            || entry.key() == 0
            || entry.ref_count() == 0
        {
            self.invalid = true;
            return Ok(());
        }
        if !self.reserve_entry_slot() {
            return Ok(());
        }
        self.entries.push(RichEntrySeed {
            key: entry.key(),
            ref_count: entry.ref_count(),
            payload,
        });
        Ok(())
    }
}

struct RichPendingEntry {
    seed: RichEntrySeed,
    owner: RichEntryOwner,
    payload_id: u64,
}

fn resolve_rich_route_index(
    source: &Package,
    route: &ListRoute,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<RichRouteIndex, Error> {
    let before = budget.usage;
    let (root_object, root_payload, root_object_id) = routed_message(source, route.message, path)?;
    let mut collector = RichEntryCollector {
        entries: budget.reserve_scratch(route.entries, LimitKind::RetainedElements)?,
        limit: route.entries,
        accepted: 0,
        invalid: false,
    };
    let (root_snapshot, report) = storage_codec::decode_table_data_list_with_visitor(
        root_payload,
        budget.options(root_payload.len())?,
        &mut collector,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    if root_snapshot.list_type() != ListType::RichTextPayload as i32 {
        return Err(Error::InvalidSource { path });
    }
    let root_entries = collector.entries.len();
    let mut owners = budget.reserve_scratch(route.entries, LimitKind::RetainedElements)?;
    owners.resize(root_entries, RichEntryOwner::Root);
    for &segment_route in &route.segments {
        let (segment_object, payload, segment_id) = routed_message(source, segment_route, path)?;
        let prior = collector.entries.len();
        let (snapshot, report) = storage_codec::decode_table_data_list_segment_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut collector,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if snapshot.list_type() != ListType::RichTextPayload as i32 {
            return Err(Error::InvalidSource { path });
        }
        let owner_entries = collector
            .entries
            .len()
            .checked_sub(prior)
            .and_then(|count| u32::try_from(count).ok())
            .ok_or(Error::InvalidSource { path })?;
        let root_references = root_object
            .archive_info
            .message_infos
            .get(route.message.message_index)
            .ok_or(Error::InvalidSource { path })?
            .object_references
            .iter()
            .filter(|candidate| **candidate == segment_id)
            .count();
        let root_references = u32::try_from(root_references)
            .ok()
            .filter(|count| *count == 1)
            .ok_or(Error::InvalidSource { path })?;
        let owner = RichEntryOwner::Segment {
            message: segment_route,
            object_id: segment_id,
            owner_entries,
            root_references,
        };
        owners.resize(collector.entries.len(), owner);
        let _ = segment_object;
    }
    if collector.invalid
        || collector.accepted != collector.entries.len()
        || collector.entries.len() != route.entries
        || owners.len() != route.entries
    {
        return Err(Error::InvalidSource { path });
    }

    let collector_capacity = collector.entries.capacity();
    let owners_capacity = owners.capacity();
    let mut pending = budget.reserve_scratch(route.entries, LimitKind::RetainedElements)?;
    for (seed, owner) in collector.entries.into_iter().zip(owners) {
        pending.push(RichPendingEntry {
            payload_id: checked_reference(seed.payload, path)?,
            seed,
            owner,
        });
    }
    budget.release_scratch_capacity::<RichEntrySeed>(collector_capacity)?;
    budget.release_scratch_capacity::<RichEntryOwner>(owners_capacity)?;
    budget.charge_auxiliary_work(sort_work_bound(pending.len(), path)?)?;
    pending.sort_unstable_by_key(|entry| (entry.payload_id, entry.seed.key));

    let mut pairs = budget.reserve_retained(route.entries, LimitKind::Objects)?;
    let mut entries = budget.reserve_retained(route.entries, LimitKind::RetainedElements)?;
    let pending_capacity = pending.capacity();
    let mut prior_payload = None;
    let mut pair_index = None;
    for pending_entry in pending {
        let current_pair = if prior_payload == Some(pending_entry.payload_id) {
            pair_index.ok_or(Error::InvalidSource { path })?
        } else {
            identities.claim(
                pending_entry.payload_id,
                Role::RichPayload,
                true,
                budget,
                path,
            )?;
            let (payload, storage_id) =
                resolve_rich_payload(source, pending_entry.payload_id, budget, path)?;
            identities.claim(storage_id, Role::RichStorage, true, budget, path)?;
            let storage = resolve_rich_storage(source, storage_id, identities, budget, path)?;
            let index = pairs.len();
            pairs.push(RichResolvedPairRoute { payload, storage });
            prior_payload = Some(pending_entry.payload_id);
            pair_index = Some(index);
            index
        };
        entries.push(RichEntryRoute {
            key: pending_entry.seed.key,
            ref_count: pending_entry.seed.ref_count,
            owner: pending_entry.owner,
            root: route.message,
            root_object_id,
            next_key: root_snapshot.next_list_id(),
            pair_index: current_pair,
        });
    }
    budget.release_scratch_capacity::<RichPendingEntry>(pending_capacity)?;
    budget.charge_auxiliary_work(sort_work_bound(entries.len(), path)?)?;
    entries.sort_unstable_by_key(|entry| entry.key);
    if entries.windows(2).any(|pair| pair[0].key >= pair[1].key) {
        return Err(Error::InvalidSource { path });
    }
    let next_key = entries
        .last()
        .map_or(Some(1), |entry| entry.key.checked_add(1))
        .map(|minimum| minimum.max(root_snapshot.next_list_id()).max(1))
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::RetainedElements,
            observed: u64::from(u32::MAX),
            maximum: u64::from(u32::MAX - 1),
            path,
        })?;
    if entries
        .binary_search_by_key(&next_key, |entry| entry.key)
        .is_ok()
    {
        return Err(Error::InvalidSource { path });
    }
    for entry in &mut entries {
        entry.next_key = next_key;
    }

    let local_object_ids = collect_local_object_ids(source, budget, path)?;
    let mut entry_payload_ids = budget.reserve_scratch(entries.len(), LimitKind::References)?;
    for entry in &entries {
        entry_payload_ids.push(
            pairs
                .get(entry.pair_index)
                .ok_or(Error::InvalidSource { path })?
                .payload
                .object_id,
        );
    }
    let payload_list_inbound = aggregate_rich_counts(
        pairs.iter().map(|pair| pair.payload.object_id),
        entry_payload_ids.iter().copied(),
        budget,
        path,
    )?;
    budget.release_scratch(&entry_payload_ids)?;
    drop(entry_payload_ids);
    let storage_payload_inbound = aggregate_rich_counts(
        pairs.iter().map(|pair| pair.storage.object_id),
        pairs.iter().map(|pair| pair.storage.object_id),
        budget,
        path,
    )?;
    validate_rich_inbound(
        source,
        &payload_list_inbound,
        &storage_payload_inbound,
        budget,
        path,
    )?;
    Ok(RichRouteIndex {
        entries,
        pairs,
        local_object_ids,
        payload_list_inbound,
        storage_payload_inbound,
        report: codec_usage_delta(before, budget.usage, path)?,
    })
}

fn routed_message(
    source: &Package,
    route: MessageRoute,
    path: Path,
) -> Result<(&ArchiveObject, &[u8], u64), Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .ok_or(Error::InvalidSource { path })?;
    let object_id = object
        .archive_info
        .identifier
        .ok_or(Error::InvalidSource { path })?;
    let message = object
        .messages
        .get(route.message_index)
        .filter(|message| message.type_ == route.message_type)
        .ok_or(Error::InvalidSource { path })?;
    table_headers::resolve::validate_message_metadata(object, route.message_index)
        .map_err(map_header_failure)?;
    Ok((object, &message.data, object_id))
}

fn resolve_rich_payload(
    source: &Package,
    identifier: u64,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(RichObjectRoute, u64), Error> {
    let (resolved, index, payload) =
        resolve_exact_message(source, identifier, &[RICH_PAYLOAD_MESSAGE_TYPE], path)?;
    if resolved.messages.len() != 1 {
        return Err(Error::InvalidSource { path });
    }
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut storage = None;
    let mut cell_id = false;
    for field in view.fields() {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return Err(Error::InvalidSource { path });
        }
        match field.number() {
            1 if storage.is_none() => {
                storage = Some(
                    table_headers::resolve::local_reference_identifier(field.payload())
                        .map_err(map_header_failure)?,
                );
            },
            2 => {},
            3 if !cell_id => cell_id = true,
            _ => return Err(Error::InvalidSource { path }),
        }
    }
    let storage = storage.ok_or(Error::InvalidSource { path })?;
    if !cell_id {
        return Err(Error::InvalidSource { path });
    }
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    let route = capture_rich_object(
        object,
        resolved,
        index,
        RICH_PAYLOAD_MESSAGE_TYPE,
        budget,
        path,
    )?;
    if route.object_references.as_slice() != [storage] {
        return Err(Error::InvalidSource { path });
    }
    require_rich_field_if_present(&route, storage, &[1], path)?;
    Ok((route, storage))
}

fn resolve_rich_storage(
    source: &Package,
    identifier: u64,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<RichObjectRoute, Error> {
    let (resolved, index, _payload) =
        resolve_exact_message(source, identifier, &[RICH_STORAGE_MESSAGE_TYPE], path)?;
    if resolved.messages.len() != 1 {
        return Err(Error::InvalidSource { path });
    }
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    let route = capture_rich_object(
        object,
        resolved,
        index,
        RICH_STORAGE_MESSAGE_TYPE,
        budget,
        path,
    )?;
    for &reference in &route.object_references {
        if reference == 0 {
            return Err(Error::InvalidSource { path });
        }
        identities.claim(reference, Role::RichStyle, true, budget, path)?;
        let _ = resolve_object(source, reference, path)?;
    }
    Ok(route)
}

fn capture_rich_object(
    object: &ArchiveObject,
    resolved: super::super::Resolved<'_>,
    message_index: usize,
    message_type: u32,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<RichObjectRoute, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    let mut object_references =
        budget.reserve_retained(info.object_references.len(), LimitKind::References)?;
    object_references.extend_from_slice(&info.object_references);
    let mut field_references =
        budget.reserve_retained(info.field_infos.len(), LimitKind::References)?;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let mut field_path =
            budget.reserve_retained(field.path.as_slice().len(), LimitKind::RetainedElements)?;
        field_path.extend_from_slice(field.path.as_slice());
        let mut references =
            budget.reserve_retained(field.object_references.len(), LimitKind::References)?;
        references.extend_from_slice(&field.object_references);
        for &reference in &references {
            if reference == 0
                || object_references
                    .iter()
                    .filter(|candidate| **candidate == reference)
                    .count()
                    != 1
            {
                return Err(Error::InvalidSource { path });
            }
        }
        field_references.push(RichFieldRefs {
            field_info_index,
            path: field_path,
            object_references: references,
        });
    }
    Ok(RichObjectRoute {
        message: MessageRoute::new(resolved, message_index, message_type),
        object_id: object
            .archive_info
            .identifier
            .ok_or(Error::InvalidSource { path })?,
        message_type,
        object_references,
        field_references,
    })
}

fn require_rich_field_if_present(
    route: &RichObjectRoute,
    identifier: u64,
    path: &[u32],
    error_path: Path,
) -> Result<(), Error> {
    let declarations = route
        .field_references
        .iter()
        .filter(|field| field.object_references.contains(&identifier))
        .count();
    if declarations == 0 {
        return Ok(());
    }
    if declarations != 1
        || !route.field_references.iter().any(|field| {
            field.path == path
                && field
                    .object_references
                    .iter()
                    .filter(|candidate| **candidate == identifier)
                    .count()
                    == 1
        })
    {
        return Err(Error::InvalidSource { path: error_path });
    }
    Ok(())
}

fn collect_local_object_ids(
    source: &Package,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Vec<u64>, Error> {
    let count = source
        .state
        .components
        .catalog()
        .iter()
        .try_fold(0usize, |count, component| {
            count.checked_add(component.archive().objects.len())
        })
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_auxiliary_work(
        count
            .checked_add(sort_work_bound(count, path)?)
            .ok_or(Error::InvalidSource { path })?,
    )?;
    let mut ids = budget.reserve_retained(count, LimitKind::Objects)?;
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            let id = object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource { path })?;
            if id == 0 {
                return Err(Error::InvalidSource { path });
            }
            ids.push(id);
        }
    }
    ids.sort_unstable();
    if ids.windows(2).any(|pair| pair[0] >= pair[1]) {
        return Err(Error::InvalidSource { path });
    }
    Ok(ids)
}

fn aggregate_rich_counts(
    unique: impl Iterator<Item = u64>,
    occurrences: impl Iterator<Item = u64>,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Vec<(u64, u32)>, Error> {
    let unique = unique;
    let (minimum, maximum) = unique.size_hint();
    if maximum != Some(minimum) {
        return Err(Error::InvalidSource { path });
    }
    let mut ids = budget.reserve_scratch(minimum, LimitKind::References)?;
    ids.extend(unique);
    let occurrences_hint = occurrences.size_hint();
    if occurrences_hint.1 != Some(occurrences_hint.0) {
        return Err(Error::InvalidSource { path });
    }
    let lookup_work = sort_work_bound(minimum, path)?
        .checked_add(
            occurrences_hint
                .0
                .checked_mul(binary_lookup_work(minimum, path)?)
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_auxiliary_work(lookup_work)?;
    ids.sort_unstable();
    ids.dedup();
    let mut counts = budget.reserve_retained(ids.len(), LimitKind::References)?;
    for &id in &ids {
        counts.push((id, 0u32));
    }
    for id in occurrences {
        let index = counts
            .binary_search_by_key(&id, |(identifier, _count)| *identifier)
            .map_err(|_error| Error::InvalidSource { path })?;
        counts[index].1 = counts[index]
            .1
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
    }
    budget.release_scratch(&ids)?;
    Ok(counts)
}

fn binary_lookup_work(length: usize, path: Path) -> Result<usize, Error> {
    if length < 2 {
        Ok(1)
    } else {
        usize::try_from(usize::BITS - (length - 1).leading_zeros())
            .map_err(|_error| Error::InvalidSource { path })
    }
}

fn sort_work_bound(length: usize, path: Path) -> Result<usize, Error> {
    if length < 2 {
        return Ok(length);
    }
    let passes = usize::BITS
        .checked_sub((length - 1).leading_zeros())
        .ok_or(Error::InvalidSource { path })?;
    length
        .checked_mul(usize::try_from(passes).map_err(|_error| Error::InvalidSource { path })?)
        .and_then(|work| work.checked_mul(2))
        .ok_or(Error::InvalidSource { path })
}

fn validate_rich_inbound(
    source: &Package,
    payloads: &[(u64, u32)],
    storages: &[(u64, u32)],
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(), Error> {
    let mut observed_payloads = budget.reserve_scratch(payloads.len(), LimitKind::References)?;
    observed_payloads.resize(payloads.len(), 0u32);
    let mut observed_storages = budget.reserve_scratch(storages.len(), LimitKind::References)?;
    observed_storages.resize(storages.len(), 0u32);
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            for info in &object.archive_info.message_infos {
                let work = 1usize
                    .checked_add(info.object_references.len())
                    .ok_or(Error::InvalidSource { path })?;
                budget.charge_auxiliary_work(work)?;
                for &reference in &info.object_references {
                    if let Ok(index) = payloads.binary_search_by_key(&reference, |item| item.0) {
                        observed_payloads[index] = observed_payloads[index]
                            .checked_add(1)
                            .ok_or(Error::InvalidSource { path })?;
                    }
                    if let Ok(index) = storages.binary_search_by_key(&reference, |item| item.0) {
                        observed_storages[index] = observed_storages[index]
                            .checked_add(1)
                            .ok_or(Error::InvalidSource { path })?;
                    }
                }
            }
        }
    }
    if payloads
        .iter()
        .map(|item| item.1)
        .ne(observed_payloads.iter().copied())
        || storages
            .iter()
            .map(|item| item.1)
            .ne(observed_storages.iter().copied())
    {
        return Err(Error::InvalidSource { path });
    }
    budget.release_scratch(&observed_payloads)?;
    budget.release_scratch(&observed_storages)?;
    Ok(())
}

fn codec_usage_delta(
    before: CodecUsage,
    after: CodecUsage,
    path: Path,
) -> Result<CodecUsage, Error> {
    Ok(CodecUsage {
        source_bytes: after
            .source_bytes
            .checked_sub(before.source_bytes)
            .ok_or(Error::InvalidSource { path })?,
        fields: after
            .fields
            .checked_sub(before.fields)
            .ok_or(Error::InvalidSource { path })?,
        work_bytes: after
            .work_bytes
            .checked_sub(before.work_bytes)
            .ok_or(Error::InvalidSource { path })?,
        references: after
            .references
            .checked_sub(before.references)
            .ok_or(Error::InvalidSource { path })?,
        reference_bytes: after
            .reference_bytes
            .checked_sub(before.reference_bytes)
            .ok_or(Error::InvalidSource { path })?,
        text_bytes: after
            .text_bytes
            .checked_sub(before.text_bytes)
            .ok_or(Error::InvalidSource { path })?,
        max_depth: after.max_depth,
    })
}

fn resolve_dependencies(
    source: &Package,
    native: table_headers::Target,
    positions: &[CellPosition],
    owner_markers: ModelOwnerMarkers,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<DependencyRoutes, Error> {
    let Some((engine_id, document, document_message, engine_path)) =
        rooted_engine_reference(source, path)?
    else {
        reject_unresolved_owner_markers(owner_markers, path)?;
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::FormulaCache,
        });
    };
    identities.claim(engine_id, Role::CalculationEngine, false, budget, path)?;
    require_declared(document, document_message, engine_id, engine_path, path)?;
    let (engine_resolved, engine_index, engine_payload) =
        resolve_exact_message(source, engine_id, &[CALCULATION_ENGINE_MESSAGE_TYPE], path)?;
    let mut count = DependencyCount::default();
    let (engine, report) = dependency_codec::decode_calculation_engine_with_visitor(
        engine_payload,
        budget.options(engine_payload.len())?,
        &mut count,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    count.validate(path)?;
    let mut stage = DependencyStage::with_capacity(count, budget, path)?;
    let (engine_again, report) = dependency_codec::decode_calculation_engine_with_visitor(
        engine_payload,
        budget.options(engine_payload.len())?,
        &mut stage,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    if engine != engine_again || !stage.matches(count) {
        return Err(Error::InvalidSource { path });
    }
    let tracker_payload = engine.dependency_tracker();
    let (tracker, report) = dependency_codec::decode_dependency_tracker_with_report(
        tracker_payload,
        budget.options(tracker_payload.len())?,
    )
    .map_err(|error| map_codec_failure(error, path))?;
    budget.charge(report)?;
    let formula_count = tracker
        .number_of_formulas()
        .ok_or(Error::InvalidSource { path })?;
    let owner_map_payload = tracker
        .owner_id_map()
        .ok_or(Error::InvalidSource { path })?;
    let (owner_identities, owner_map_fields) = parse_owner_id_map(owner_map_payload, budget, path)?;
    budget.charge_wire_scan(owner_map_payload.len(), owner_map_fields, 2)?;
    if owner_identities.len() != stage.formula_owners.len() {
        return Err(Error::InvalidSource { path });
    }
    let engine_object = table_headers::resolve::resolved_object(source, engine_resolved)
        .map_err(map_header_failure)?;
    let engine_sidecar_count = usize::from(engine.named_reference_manager().is_some())
        .checked_add(usize::from(engine.remote_data_store().is_some()))
        .and_then(|count| count.checked_add(usize::from(engine.refs_to_dirty().is_some())))
        .ok_or(Error::InvalidSource { path })?;
    let mut engine_sidecars = budget.reserve_retained(engine_sidecar_count, LimitKind::Objects)?;
    if let Some(route) = validate_optional_engine_reference(
        source,
        engine_object,
        engine_index,
        engine.named_reference_manager(),
        3,
        identities,
        budget,
        path,
    )? {
        engine_sidecars.push(route);
    }
    if let Some(route) = validate_optional_engine_reference(
        source,
        engine_object,
        engine_index,
        engine.remote_data_store(),
        12,
        identities,
        budget,
        path,
    )? {
        engine_sidecars.push(route);
    }
    if let Some(route) = validate_optional_engine_reference(
        source,
        engine_object,
        engine_index,
        engine.refs_to_dirty(),
        15,
        identities,
        budget,
        path,
    )? {
        engine_sidecars.push(route);
    }

    let header_name_manager = match engine.header_name_manager() {
        None => None,
        Some(reference) => {
            let id = checked_reference(reference, path)?;
            identities.claim(id, Role::HeaderNameManager, false, budget, path)?;
            require_declared(engine_object, engine_index, id, &[14], path)?;
            let (resolved, index, _payload) =
                resolve_exact_message(source, id, &[HEADER_NAME_MANAGER_MESSAGE_TYPE], path)?;
            Some(MessageRoute::new(
                resolved,
                index,
                HEADER_NAME_MANAGER_MESSAGE_TYPE,
            ))
        },
    };
    if header_name_manager.is_some() && touches_header_coordinate(native, positions, path)? {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::HeaderNameIndex,
        });
    }

    let mut owner_routes =
        budget.reserve_retained(stage.formula_owners.len(), LimitKind::Objects)?;
    let mut cell_routes = Vec::new();
    let mut range_routes = Vec::new();
    let mut inert_marker_tiles = budget.reserve_retained(
        usize::from(owner_markers.haunted.is_some()),
        LimitKind::Objects,
    )?;
    let mut selected_owner_count = 0usize;
    let mut conditional_owner_matches = 0usize;
    let mut spill_owner_matches = 0usize;
    let mut haunted_owner_matches = 0usize;
    let mut active_spill_owner = false;
    let mut selected_formula_owner = None;
    for &reference in &stage.formula_owners {
        let id = checked_reference(reference, path)?;
        identities.claim(id, Role::FormulaOwner, false, budget, path)?;
        require_declared(engine_object, engine_index, id, &[2, 6], path)?;
        let (resolved, index, payload) =
            resolve_exact_message(source, id, &[FORMULA_OWNER_MESSAGE_TYPE], path)?;
        let owner_object = table_headers::resolve::resolved_object(source, resolved)
            .map_err(map_header_failure)?;
        let mut count = DependencyCount::default();
        let (owner_snapshot, report) =
            dependency_codec::decode_formula_owner_dependencies_with_visitor(
                payload,
                budget.options(payload.len())?,
                &mut count,
            )
            .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        count.validate(path)?;
        let owner_key = UuidKey {
            lower: owner_snapshot.formula_owner_uid().lower(),
            upper: owner_snapshot.formula_owner_uid().upper(),
        };
        if owner_identities
            .iter()
            .filter(|identity| {
                identity.uid == owner_key
                    && identity.internal == owner_snapshot.internal_formula_owner_id()
            })
            .count()
            != 1
        {
            return Err(Error::InvalidSource { path });
        }
        if owner_markers.conditional_style == Some(owner_key) {
            conditional_owner_matches = conditional_owner_matches
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            validate_empty_formula_owner_closure(
                owner_snapshot,
                count,
                3,
                DependencyKind::ConditionalStyle,
                path,
            )?;
        }
        if owner_markers.spill == Some(owner_key) {
            spill_owner_matches = spill_owner_matches
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            validate_empty_formula_owner_closure(
                owner_snapshot,
                count,
                12,
                DependencyKind::Spill,
                path,
            )?;
            if let Some(payload) = owner_snapshot.spill_range_sizes() {
                active_spill_owner |= spill_sizes_active(payload, path)?;
            }
            if count.cell_tiles != 0 || count.range_tiles != 0 {
                active_spill_owner = true;
            }
        }
        let inert_marker_start = inert_marker_tiles.len();
        if owner_markers.haunted == Some(owner_key) {
            haunted_owner_matches = haunted_owner_matches
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            validate_empty_formula_owner_closure(
                owner_snapshot,
                count,
                35,
                DependencyKind::FormulaCache,
                path,
            )?;
        }
        let mut owner_stage = DependencyStage::with_capacity(count, budget, path)?;
        let (_again, report) = dependency_codec::decode_formula_owner_dependencies_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut owner_stage,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if !owner_stage.matches(count) {
            return Err(Error::InvalidSource { path });
        }
        if owner_markers.haunted == Some(owner_key) {
            resolve_empty_cell_dependency_tiles(
                source,
                owner_object,
                index,
                &owner_stage.cell_tiles,
                owner_snapshot.internal_formula_owner_id(),
                identities,
                budget,
                &mut inert_marker_tiles,
                path,
            )?;
        }
        let selected_owner = owner_snapshot
            .formula_owner()
            .is_some_and(|reference| reference.identifier() == native.drawable_identifier);
        if selected_owner {
            selected_owner_count = selected_owner_count
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            if selected_owner_count > 1 {
                return Err(Error::InvalidSource { path });
            }
            validate_selected_formula_owner_closure(
                owner_snapshot,
                native,
                &owner_stage,
                budget,
                path,
            )?;
        }
        validate_formula_owner_reference(
            source,
            owner_object,
            index,
            owner_snapshot.formula_owner(),
            native,
            identities,
            budget,
            path,
        )?;
        let mut owner_cell_routes = Vec::new();
        let mut owner_range_routes = Vec::new();
        if owner_markers.haunted == Some(owner_key) {
            budget.reserve_retained_growth(
                &mut owner_cell_routes,
                inert_marker_tiles.len() - inert_marker_start,
                LimitKind::Objects,
            )?;
            owner_cell_routes.extend_from_slice(&inert_marker_tiles[inert_marker_start..]);
        }
        if selected_owner {
            resolve_selected_cell_dependency_tiles(
                source,
                owner_object,
                index,
                &owner_stage.cell_tiles,
                owner_snapshot.internal_formula_owner_id(),
                native,
                identities,
                budget,
                &mut owner_cell_routes,
                path,
            )?;
            resolve_dependency_tiles(
                source,
                owner_object,
                index,
                &owner_stage.range_tiles,
                &[15, 1],
                RANGE_PRECEDENTS_TILE_MESSAGE_TYPE,
                Role::RangeDependencyTile,
                identities,
                budget,
                &mut owner_range_routes,
                path,
            )?;
        } else if owner_markers.conditional_style != Some(owner_key)
            && owner_markers.spill != Some(owner_key)
            && owner_markers.haunted != Some(owner_key)
        {
            resolve_dependency_tiles(
                source,
                owner_object,
                index,
                &owner_stage.cell_tiles,
                &[13, 1],
                CELL_RECORD_TILE_MESSAGE_TYPE,
                Role::CellDependencyTile,
                identities,
                budget,
                &mut owner_cell_routes,
                path,
            )?;
            resolve_dependency_tiles(
                source,
                owner_object,
                index,
                &owner_stage.range_tiles,
                &[15, 1],
                RANGE_PRECEDENTS_TILE_MESSAGE_TYPE,
                Role::RangeDependencyTile,
                identities,
                budget,
                &mut owner_range_routes,
                path,
            )?;
        }
        let mut owner_range_routes_with_targets =
            budget.reserve_retained(owner_range_routes.len(), LimitKind::Objects)?;
        for message in &owner_range_routes {
            let payload = message_payload_at_route(source, *message, path)?;
            let (snapshot, report) = dependency_codec::decode_range_precedents_tile_with_report(
                payload,
                budget.options(payload.len())?,
            )
            .map_err(|error| map_codec_failure(error, path))?;
            budget.charge(report)?;
            owner_range_routes_with_targets.push(RangePrecedentRoute {
                message: *message,
                target_owner: snapshot.to_owner_id(),
            });
        }
        if selected_owner {
            budget.reserve_retained_growth(
                &mut cell_routes,
                owner_cell_routes.len(),
                LimitKind::Objects,
            )?;
            cell_routes.extend_from_slice(&owner_cell_routes);
            budget.reserve_retained_growth(
                &mut range_routes,
                owner_range_routes.len(),
                LimitKind::Objects,
            )?;
            range_routes.extend_from_slice(&owner_range_routes);
        }
        let owner_route = MessageRoute::new(resolved, index, FORMULA_OWNER_MESSAGE_TYPE);
        if selected_owner {
            selected_formula_owner = Some(SelectedFormulaOwnerRoute {
                message: owner_route,
                internal_owner_id: owner_snapshot.internal_formula_owner_id(),
                uid_lower: owner_key.lower,
                uid_upper: owner_key.upper,
            });
        }
        owner_routes.push(FormulaOwnerRoute {
            message: owner_route,
            formula_owner_object_id: owner_snapshot
                .formula_owner()
                .map(|reference| reference.identifier()),
            internal_owner_id: owner_snapshot.internal_formula_owner_id(),
            uid_lower: owner_key.lower,
            uid_upper: owner_key.upper,
            cell_record_tiles: owner_cell_routes,
            range_precedent_tiles: owner_range_routes_with_targets,
        });
        owner_stage.release_scratch(budget)?;
    }
    validate_deferred_owner_closure(
        owner_markers,
        conditional_owner_matches,
        spill_owner_matches,
        active_spill_owner,
        haunted_owner_matches,
        path,
    )?;
    if selected_owner_count != 1 {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::FormulaCache,
        });
    }
    stage.release_scratch(budget)?;
    budget.release_scratch(&owner_identities)?;
    Ok(DependencyRoutes {
        engine: Some(MessageRoute::new(
            engine_resolved,
            engine_index,
            CALCULATION_ENGINE_MESSAGE_TYPE,
        )),
        formula_count,
        engine_sidecars,
        formula_owners: owner_routes,
        selected_formula_owner,
        inert_marker_tiles,
        cell_record_tiles: cell_routes,
        range_precedent_tiles: range_routes,
        header_name_manager,
    })
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
struct DependencyCount {
    formula_owners: usize,
    cell_tiles: usize,
    range_tiles: usize,
    cell_records: usize,
    edge_components: usize,
    overflowed: bool,
}

impl DependencyCount {
    fn validate(self, path: Path) -> Result<(), Error> {
        count_not_overflowed(self.overflowed, path)
    }
}

impl dependency_codec::DependencyVisitor for DependencyCount {
    fn visit_formula_owner_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.formula_owners.checked_add(1) {
            self.formula_owners = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.cell_tiles.checked_add(1) {
            self.cell_tiles = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }

    fn visit_tiled_range_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.range_tiles.checked_add(1) {
            self.range_tiles = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        _record: dependency_codec::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.cell_records.checked_add(1) {
            self.cell_records = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }

    fn visit_expanded_edge_component(
        &mut self,
        _component: dependency_codec::ExpandedEdgeComponent,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.edge_components.checked_add(1) {
            self.edge_components = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

struct DependencyStage {
    formula_owners: Vec<dependency_codec::ReferenceSnapshot>,
    cell_tiles: Vec<dependency_codec::ReferenceSnapshot>,
    range_tiles: Vec<dependency_codec::ReferenceSnapshot>,
    cell_records: Vec<SelectedCellRecord>,
    edge_components: Vec<dependency_codec::ExpandedEdgeComponent>,
}

#[derive(Debug, Clone, Copy)]
struct SelectedCellRecord {
    column: u32,
    row: u32,
    dirty_self_plus_precedents_count: Option<u64>,
    is_in_a_cycle: Option<bool>,
    has_calculated_precedents: Option<bool>,
    edges: Option<dependency_codec::ExpandedEdgesSnapshot>,
    component_start: usize,
    component_end: usize,
}

impl DependencyStage {
    fn with_capacity(
        count: DependencyCount,
        budget: &mut ResolutionBudget,
        _path: Path,
    ) -> Result<Self, Error> {
        Ok(Self {
            formula_owners: budget.reserve_scratch(count.formula_owners, LimitKind::References)?,
            cell_tiles: budget.reserve_scratch(count.cell_tiles, LimitKind::References)?,
            range_tiles: budget.reserve_scratch(count.range_tiles, LimitKind::References)?,
            cell_records: budget.reserve_scratch(count.cell_records, LimitKind::Objects)?,
            edge_components: budget
                .reserve_scratch(count.edge_components, LimitKind::References)?,
        })
    }

    fn matches(&self, count: DependencyCount) -> bool {
        self.formula_owners.len() == count.formula_owners
            && self.cell_tiles.len() == count.cell_tiles
            && self.range_tiles.len() == count.range_tiles
            && self.cell_records.len() == count.cell_records
            && self.edge_components.len() == count.edge_components
    }

    fn release_scratch(&self, budget: &mut ResolutionBudget) -> Result<(), Error> {
        budget.release_scratch(&self.formula_owners)?;
        budget.release_scratch(&self.cell_tiles)?;
        budget.release_scratch(&self.range_tiles)?;
        budget.release_scratch(&self.cell_records)?;
        budget.release_scratch(&self.edge_components)
    }
}

impl dependency_codec::DependencyVisitor for DependencyStage {
    fn visit_formula_owner_dependency(
        &mut self,
        reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        self.formula_owners.push(reference.reference());
        Ok(())
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        self.cell_tiles.push(reference.reference());
        Ok(())
    }

    fn visit_tiled_range_dependency(
        &mut self,
        reference: dependency_codec::ReferenceRecord<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        self.range_tiles.push(reference.reference());
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        record: dependency_codec::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        let component_start = self
            .cell_records
            .last()
            .map_or(0, |record| record.component_end);
        self.cell_records.push(SelectedCellRecord {
            column: record.column(),
            row: record.row(),
            dirty_self_plus_precedents_count: record.dirty_self_plus_precedents_count(),
            is_in_a_cycle: record.is_in_a_cycle(),
            has_calculated_precedents: record.has_calculated_precedents(),
            edges: record.expanded_edges_snapshot(),
            component_start,
            component_end: self.edge_components.len(),
        });
        Ok(())
    }

    fn visit_expanded_edge_component(
        &mut self,
        component: dependency_codec::ExpandedEdgeComponent,
    ) -> Result<(), dependency_codec::DecodeError> {
        self.edge_components.push(component);
        Ok(())
    }
}

struct CellRecordCount {
    records: usize,
    overflowed: bool,
}

impl CellRecordCount {
    const fn new() -> Self {
        Self {
            records: 0,
            overflowed: false,
        }
    }

    fn validate(self, path: Path) -> Result<usize, Error> {
        count_not_overflowed(self.overflowed, path)?;
        Ok(self.records)
    }
}

impl dependency_codec::DependencyVisitor for CellRecordCount {
    fn visit_cell_record(
        &mut self,
        _record: dependency_codec::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency_codec::DecodeError> {
        if let Some(next) = self.records.checked_add(1) {
            self.records = next;
        } else {
            self.overflowed = true;
        }
        Ok(())
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "one dormant dependency-tile proof"
)]
fn resolve_empty_cell_dependency_tiles(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[dependency_codec::ReferenceSnapshot],
    expected_internal_owner_id: u32,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    routes: &mut Vec<MessageRoute>,
    path: Path,
) -> Result<(), Error> {
    budget.reserve_retained_growth(routes, references.len(), LimitKind::Objects)?;
    for &reference in references {
        let id = checked_dependency_reference(reference, path)?;
        identities.claim(id, Role::CellDependencyTile, false, budget, path)?;
        require_declared(owner, owner_message, id, &[13, 1], path)?;
        let (resolved, index, payload) =
            resolve_exact_message(source, id, &[CELL_RECORD_TILE_MESSAGE_TYPE], path)?;
        let mut records = CellRecordCount::new();
        let (snapshot, report) = dependency_codec::decode_cell_record_tile_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut records,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if records.validate(path)? != 0
            || snapshot.internal_owner_id() != expected_internal_owner_id
        {
            return Err(Error::UnsupportedDependency {
                path,
                kind: DependencyKind::FormulaCache,
            });
        }
        routes.push(MessageRoute::new(
            resolved,
            index,
            CELL_RECORD_TILE_MESSAGE_TYPE,
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments, reason = "selected dependency-tile proof")]
fn resolve_selected_cell_dependency_tiles(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[dependency_codec::ReferenceSnapshot],
    expected_internal_owner_id: u32,
    native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    routes: &mut Vec<MessageRoute>,
    path: Path,
) -> Result<(), Error> {
    budget.reserve_retained_growth(routes, references.len(), LimitKind::Objects)?;
    for &reference in references {
        let id = checked_dependency_reference(reference, path)?;
        identities.claim(id, Role::CellDependencyTile, false, budget, path)?;
        require_declared(owner, owner_message, id, &[13, 1], path)?;
        let (resolved, index, payload) =
            resolve_exact_message(source, id, &[CELL_RECORD_TILE_MESSAGE_TYPE], path)?;
        let mut count = DependencyCount::default();
        let (snapshot, report) = dependency_codec::decode_cell_record_tile_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut count,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        count.validate(path)?;
        let mut stage = DependencyStage::with_capacity(count, budget, path)?;
        let (snapshot_again, report) = dependency_codec::decode_cell_record_tile_with_visitor(
            payload,
            budget.options(payload.len())?,
            &mut stage,
        )
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        if snapshot != snapshot_again
            || !stage.matches(count)
            || snapshot.internal_owner_id() != expected_internal_owner_id
            || snapshot.tile_column_begin() >= native.columns
            || snapshot.tile_row_begin() >= native.rows
        {
            return Err(Error::InvalidSource { path });
        }
        validate_selected_cell_tile_wire(payload, snapshot, &stage.cell_records, path)?;
        validate_selected_cell_graph(&stage, native, path)?;
        budget.charge_auxiliary_work(
            stage
                .cell_records
                .len()
                .checked_add(stage.edge_components.len())
                .ok_or(Error::InvalidSource { path })?,
        )?;
        routes.push(MessageRoute::new(
            resolved,
            index,
            CELL_RECORD_TILE_MESSAGE_TYPE,
        ));
        stage.release_scratch(budget)?;
    }
    Ok(())
}

fn validate_selected_cell_tile_wire(
    payload: &[u8],
    snapshot: dependency_codec::CellRecordTileSnapshot,
    records: &[SelectedCellRecord],
    path: Path,
) -> Result<(), Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload)
        .map_err(|_error| Error::InvalidSource { path })?;
    let mut owner = false;
    let mut column = false;
    let mut row = false;
    let mut record_index = 0usize;
    for field in view.fields() {
        match field.number() {
            1 if !owner && field.wire_type() == 0 => {
                owner = canonical_u64(field.payload(), path)?
                    == u64::from(snapshot.internal_owner_id());
            },
            2 if !column && field.wire_type() == 0 => {
                column = canonical_u64(field.payload(), path)?
                    == u64::from(snapshot.tile_column_begin());
            },
            3 if !row && field.wire_type() == 0 => {
                row = canonical_u64(field.payload(), path)? == u64::from(snapshot.tile_row_begin());
            },
            4 if field.wire_type() == 2
                && field.validate_canonical_framing().is_ok()
                && record_index < records.len()
                && selected_cell_record_wire(field.payload(), records[record_index], path)? =>
            {
                record_index = record_index
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
            },
            _ => return Err(Error::InvalidSource { path }),
        }
    }
    if !owner || !column || !row || record_index != records.len() {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

#[allow(clippy::too_many_arguments, reason = "one dependency-tile proof")]
fn resolve_dependency_tiles(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    references: &[dependency_codec::ReferenceSnapshot],
    declared_path: &[u32],
    message_type: u32,
    role: Role,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    routes: &mut Vec<MessageRoute>,
    path: Path,
) -> Result<(), Error> {
    budget.reserve_retained_growth(routes, references.len(), LimitKind::Objects)?;
    for &reference in references {
        let id = checked_reference(reference, path)?;
        identities.claim(id, role, false, budget, path)?;
        require_declared(owner, owner_message, id, declared_path, path)?;
        let (resolved, index, payload) = resolve_exact_message(source, id, &[message_type], path)?;
        let report = if message_type == CELL_RECORD_TILE_MESSAGE_TYPE {
            dependency_codec::decode_cell_record_tile_with_report(
                payload,
                budget.options(payload.len())?,
            )
            .map(|(_snapshot, report)| report)
        } else {
            dependency_codec::decode_range_precedents_tile_with_report(
                payload,
                budget.options(payload.len())?,
            )
            .map(|(_snapshot, report)| report)
        }
        .map_err(|error| map_codec_failure(error, path))?;
        budget.charge(report)?;
        routes.push(MessageRoute::new(resolved, index, message_type));
    }
    Ok(())
}

fn validate_formula_owner_reference(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    reference: Option<dependency_codec::ReferenceSnapshot>,
    native: table_headers::Target,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<(), Error> {
    if let Some(reference) = reference {
        let id = checked_dependency_reference(reference, path)?;
        if id == native.drawable_identifier {
            // Native formula-owner f11 is a weak semantic back-reference to
            // the already rooted TableInfo and is intentionally absent from
            // ArchiveInfo object-reference metadata.
            validate_selected_table_info_target(source, native, id, path)?;
        } else {
            require_declared(owner, owner_message, id, &[11], path)?;
            identities.claim(id, Role::EngineOwner, true, budget, path)?;
            let _ = resolve_exact_message(
                source,
                id,
                &[
                    super::super::TABLE_INFO_MESSAGE_TYPE,
                    super::super::LEGACY_TABLE_INFO_MESSAGE_TYPE,
                ],
                path,
            )?;
        }
    }
    Ok(())
}

fn validate_selected_table_info_target(
    source: &Package,
    native: table_headers::Target,
    id: u64,
    path: Path,
) -> Result<(), Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(native.info_component_index)
        .and_then(|component| component.archive().objects.get(native.info_object_index))
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.identifier != Some(id) {
        return Err(Error::InvalidSource { path });
    }
    let accepted = [
        super::super::TABLE_INFO_MESSAGE_TYPE,
        super::super::LEGACY_TABLE_INFO_MESSAGE_TYPE,
    ];
    let mut matching = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| accepted.contains(&message.type_));
    let Some((message_index, message)) = matching.next() else {
        return Err(Error::InvalidSource { path });
    };
    if matching.next().is_some()
        || message_index != native.info_message_index
        || message.type_ != native.info_message_type
    {
        return Err(Error::InvalidSource { path });
    }
    table_headers::resolve::validate_message_metadata(object, native.info_message_index)
        .map_err(map_header_failure)
}

fn validate_optional_engine_reference(
    source: &Package,
    owner: &ArchiveObject,
    owner_message: usize,
    reference: Option<dependency_codec::ReferenceSnapshot>,
    field: u32,
    identities: &mut RoleIdentities,
    budget: &mut ResolutionBudget,
    path: Path,
) -> Result<Option<MessageRoute>, Error> {
    if let Some(reference) = reference {
        let id = checked_dependency_reference(reference, path)?;
        identities.claim(id, Role::EngineSidecar, true, budget, path)?;
        require_declared(owner, owner_message, id, &[field], path)?;
        return resolve_single_message(source, id, path).map(Some);
    }
    Ok(None)
}

fn rooted_engine_reference(
    source: &Package,
    path: Path,
) -> Result<Option<(u64, &ArchiveObject, usize, &'static [u32])>, Error> {
    let document = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource { path })?;
    let (message_index, message) = table_headers::resolve::unique_message_index(
        &document.messages,
        super::super::DOCUMENT_MESSAGE_TYPE,
    )
    .map_err(map_header_failure)?
    .ok_or(Error::InvalidSource { path })?;
    table_headers::resolve::validate_message_metadata(document, message_index)
        .map_err(map_header_failure)?;
    let legacy = table_headers::resolve::repeated_length_payloads(&message.data, 3)
        .map_err(map_header_failure)?;
    let super_payloads = table_headers::resolve::repeated_length_payloads(&message.data, 8)
        .map_err(map_header_failure)?;
    let primary = match super_payloads.as_slice() {
        [] => Vec::new(),
        [super_payload] => table_headers::resolve::repeated_length_payloads(super_payload, 4)
            .map_err(map_header_failure)?,
        _ => return Err(Error::InvalidSource { path }),
    };
    let (payload, route): (&[u8], &'static [u32]) = match (primary.as_slice(), legacy.as_slice()) {
        ([], []) => return Ok(None),
        ([payload], []) => (*payload, &[8, 4]),
        ([], [payload]) => (*payload, &[3]),
        _ => return Err(Error::InvalidSource { path }),
    };
    let id =
        table_headers::resolve::local_reference_identifier(payload).map_err(map_header_failure)?;
    Ok(Some((id, document, message_index, route)))
}

fn touches_header_coordinate(
    native: table_headers::Target,
    positions: &[CellPosition],
    path: Path,
) -> Result<bool, Error> {
    let header_rows = u32::try_from(native.settings.header_row_count())
        .map_err(|_error| Error::InvalidSource { path })?;
    let header_columns = u32::try_from(native.settings.header_column_count())
        .map_err(|_error| Error::InvalidSource { path })?;
    Ok(positions
        .iter()
        .any(|position| position.row() < header_rows || position.column() < header_columns))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Role {
    Root,
    Sheet,
    Drawable,
    Model,
    Tile,
    RowHeader,
    ColumnHeader,
    String,
    Style,
    Formula,
    Format,
    FormulaError,
    CustomFormat,
    MultipleChoice,
    RichText,
    ConditionalStyle,
    Comment,
    ImportWarning,
    ControlCellSpec,
    ListSegment,
    ListPayload,
    CalculationEngine,
    EngineSidecar,
    HeaderNameManager,
    FormulaOwner,
    HiddenStateOwner,
    CellDependencyTile,
    RangeDependencyTile,
    Merge,
    EngineOwner,
    SortRuleTracker,
    RichPayload,
    RichStorage,
    RichStyle,
}

struct RoleIdentities {
    claimed: Vec<RoleClaim>,
    finalized: bool,
    usage: RoleIdentityWork,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RoleClaim {
    identifier: u64,
    role: Role,
    allow_same_role: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct RoleIdentityWork {
    movements: usize,
    work: usize,
}

impl RoleIdentityWork {
    fn total(self, path: Path) -> Result<usize, Error> {
        self.movements
            .checked_add(self.work)
            .ok_or(Error::InvalidSource { path })
    }

    fn checked_add(self, other: Self, path: Path) -> Result<Self, Error> {
        Ok(Self {
            movements: self
                .movements
                .checked_add(other.movements)
                .ok_or(Error::InvalidSource { path })?,
            work: self
                .work
                .checked_add(other.work)
                .ok_or(Error::InvalidSource { path })?,
        })
    }
}

impl RoleIdentities {
    fn new(
        native: table_headers::Target,
        budget: &mut ResolutionBudget,
        path: Path,
    ) -> Result<Self, Error> {
        let initial_usage = RoleIdentityWork {
            movements: 4,
            work: 4,
        };
        budget.charge_auxiliary_work(initial_usage.total(path)?)?;
        let mut claimed = budget.reserve_scratch(4, LimitKind::References)?;
        claimed.push(RoleClaim {
            identifier: 1,
            role: Role::Root,
            allow_same_role: false,
        });
        claimed.push(RoleClaim {
            identifier: native.sheet_identifier,
            role: Role::Sheet,
            allow_same_role: false,
        });
        claimed.push(RoleClaim {
            identifier: native.drawable_identifier,
            role: Role::Drawable,
            allow_same_role: false,
        });
        claimed.push(RoleClaim {
            identifier: native.model_identifier,
            role: Role::Model,
            allow_same_role: false,
        });
        if claimed.iter().any(|claim| claim.identifier == 0) {
            return Err(Error::InvalidSource { path });
        }
        claimed.sort_unstable_by_key(|claim| claim.identifier);
        if claimed
            .windows(2)
            .any(|pair| pair[0].identifier == pair[1].identifier)
        {
            return Err(Error::InvalidSource { path });
        }
        Ok(Self {
            claimed,
            finalized: false,
            usage: initial_usage,
        })
    }

    fn claim(
        &mut self,
        id: u64,
        role: Role,
        allow_same_role: bool,
        budget: &mut ResolutionBudget,
        path: Path,
    ) -> Result<(), Error> {
        if id == 0 || self.finalized {
            return Err(Error::InvalidSource { path });
        }
        let prior_capacity = self.claimed.capacity();
        let prior_length = self.claimed.len();
        let growth_movements = usize::from(prior_length == prior_capacity)
            .checked_mul(prior_length)
            .ok_or(Error::InvalidSource { path })?;
        let claim_usage = RoleIdentityWork {
            movements: growth_movements,
            work: 1,
        };
        budget.charge_auxiliary_work(claim_usage.total(path)?)?;
        budget.reserve_scratch_growth(&mut self.claimed, 1, LimitKind::References)?;
        if (self.claimed.capacity() != prior_capacity) != (growth_movements != 0) {
            return Err(Error::InvalidSource { path });
        }
        self.usage = self.usage.checked_add(claim_usage, path)?;
        self.claimed.push(RoleClaim {
            identifier: id,
            role,
            allow_same_role,
        });
        Ok(())
    }

    #[cfg(test)]
    fn is_known_owner(&self, id: u64) -> bool {
        self.finalized
            && self
                .claimed
                .binary_search_by_key(&id, |claim| claim.identifier)
                .is_ok_and(|index| {
                    matches!(
                        self.claimed[index].role,
                        Role::Drawable | Role::Model | Role::FormulaOwner
                    )
                })
    }

    fn finish(
        &mut self,
        budget: &mut ResolutionBudget,
        path: Path,
    ) -> Result<RoleIdentityWork, Error> {
        if self.finalized {
            return Err(Error::InvalidSource { path });
        }
        let length = self.claimed.len();
        let finish_usage = Self::finish_usage(length, path)?;
        budget.charge_auxiliary_work(finish_usage.total(path)?)?;
        let mut sorted = budget.reserve_scratch(length, LimitKind::References)?;
        sorted.extend_from_slice(&self.claimed);
        let mut scratch = budget.reserve_scratch(length, LimitKind::References)?;
        if let Some(&seed) = sorted.first() {
            scratch.resize(length, seed);
        }
        let mut usage = self.usage;
        for byte in 0..8usize {
            let mut counts = [0usize; 256];
            for claim in &sorted {
                let bucket = usize::from(claim.identifier.to_le_bytes()[byte]);
                counts[bucket] = counts[bucket]
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
            }
            let mut offsets = [0usize; 256];
            let mut next = 0usize;
            for (index, count) in counts.into_iter().enumerate() {
                offsets[index] = next;
                next = next
                    .checked_add(count)
                    .ok_or(Error::InvalidSource { path })?;
            }
            if next != length {
                return Err(Error::InvalidSource { path });
            }
            for &claim in &sorted {
                let bucket = usize::from(claim.identifier.to_le_bytes()[byte]);
                let destination = offsets[bucket];
                scratch[destination] = claim;
                offsets[bucket] = destination
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
            }
            core::mem::swap(&mut sorted, &mut scratch);
        }
        let mut read = 0usize;
        let mut unique = 0usize;
        scratch.clear();
        while read < sorted.len() {
            let first = sorted[read];
            let mut end = read.checked_add(1).ok_or(Error::InvalidSource { path })?;
            while end < sorted.len() && sorted[end].identifier == first.identifier {
                let repeated = sorted[end];
                if !repeated.allow_same_role || repeated.role != first.role {
                    return Err(Error::InvalidSource { path });
                }
                end = end.checked_add(1).ok_or(Error::InvalidSource { path })?;
            }
            scratch.push(first);
            unique = unique.checked_add(1).ok_or(Error::InvalidSource { path })?;
            read = end;
        }
        if let Some(&seed) = scratch.first() {
            scratch.resize(length, seed);
        }
        scratch.truncate(unique);
        core::mem::swap(&mut sorted, &mut scratch);
        usage = usage.checked_add(finish_usage, path)?;
        budget.release_scratch(&scratch)?;
        budget.release_scratch(&self.claimed)?;
        self.claimed = sorted;
        self.finalized = true;
        Ok(usage)
    }

    fn finish_usage(length: usize, path: Path) -> Result<RoleIdentityWork, Error> {
        // Two complete initializations (`sorted` and `scratch`), eight stable
        // radix movements, and one deterministic compaction movement per
        // claim. The work term is two visits per radix pass, 256 buckets per
        // pass, and one compaction visit per claim.
        let movements = length
            .checked_mul(11)
            .ok_or(Error::InvalidSource { path })?;
        let work = length
            .checked_mul(17)
            .and_then(|value| value.checked_add(8 * 256))
            .ok_or(Error::InvalidSource { path })?;
        Ok(RoleIdentityWork { movements, work })
    }

    fn release_scratch(&self, budget: &mut ResolutionBudget) -> Result<(), Error> {
        budget.release_scratch(&self.claimed)
    }

    #[cfg(test)]
    fn lookup_comparisons(&self, id: u64) -> usize {
        let mut left = 0usize;
        let mut right = self.claimed.len();
        let mut comparisons = 0usize;
        while left < right {
            comparisons += 1;
            let middle = left + (right - left) / 2;
            if self.claimed[middle].identifier < id {
                left = middle + 1;
            } else {
                right = middle;
            }
        }
        comparisons
    }
}

fn checked_reference(
    reference: storage_codec::ReferenceSnapshot,
    path: Path,
) -> Result<u64, Error> {
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(Error::InvalidSource { path });
    }
    Ok(reference.identifier())
}

fn checked_dependency_reference(
    reference: dependency_codec::ReferenceSnapshot,
    path: Path,
) -> Result<u64, Error> {
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(Error::InvalidSource { path });
    }
    Ok(reference.identifier())
}

fn require_declared(
    owner: &ArchiveObject,
    owner_message: usize,
    id: u64,
    accepted_path: &[u32],
    _path: Path,
) -> Result<(), Error> {
    table_headers::resolve::require_declared_reference(owner, owner_message, id, accepted_path)
        .map_err(map_header_failure)
}

fn resolve_object<'a>(
    source: &'a Package,
    id: u64,
    path: Path,
) -> Result<super::super::Resolved<'a>, Error> {
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, id)
        .map_err(|_error| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    if object.archive_info.identifier != Some(id) {
        return Err(Error::InvalidSource { path });
    }
    Ok(resolved)
}

fn resolve_exact_message<'a>(
    source: &'a Package,
    id: u64,
    accepted_types: &[u32],
    path: Path,
) -> Result<(super::super::Resolved<'a>, usize, &'a [u8]), Error> {
    let resolved = resolve_object(source, id, path)?;
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    let mut matches = resolved
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| accepted_types.contains(&message.type_));
    let (index, message) = matches.next().ok_or(Error::InvalidSource { path })?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    table_headers::resolve::validate_message_metadata(object, index).map_err(map_header_failure)?;
    Ok((resolved, index, &message.data))
}

fn resolve_single_message(source: &Package, id: u64, path: Path) -> Result<MessageRoute, Error> {
    let resolved = resolve_object(source, id, path)?;
    if resolved.messages.len() != 1 {
        return Err(Error::InvalidSource { path });
    }
    let message = resolved
        .messages
        .first()
        .ok_or(Error::InvalidSource { path })?;
    let object =
        table_headers::resolve::resolved_object(source, resolved).map_err(map_header_failure)?;
    table_headers::resolve::validate_message_metadata(object, 0).map_err(map_header_failure)?;
    Ok(MessageRoute::new(resolved, 0, message.type_))
}

fn object_at_native_model(
    source: &Package,
    native: table_headers::Target,
    path: Path,
) -> Result<&ArchiveObject, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(native.component_index)
        .and_then(|component| component.archive().objects.get(native.object_index))
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.identifier != Some(native.model_identifier) {
        return Err(Error::InvalidSource { path });
    }
    table_headers::resolve::validate_message_metadata(object, native.message_index)
        .map_err(map_header_failure)?;
    Ok(object)
}

struct ResolutionBudget {
    path: Path,
    maximum_bytes: usize,
    maximum_source_bytes: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_text: usize,
    usage: CodecUsage,
    maximum_retained_elements: usize,
    maximum_retained_bytes: usize,
    maximum_scratch_bytes: usize,
    maximum_allocation_events: usize,
    retained_elements: usize,
    retained_bytes: usize,
    current_scratch_bytes: usize,
    peak_scratch_bytes: usize,
    allocation_events: usize,
}

impl ResolutionBudget {
    fn new(
        source: &Package,
        remaining: super::budget::Remaining,
        path: Path,
    ) -> Result<Self, Error> {
        let maximum_bytes = source.state.options.archive().max_iwa_stream_bytes();
        if maximum_bytes == 0 {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: 0,
                maximum: 0,
                path,
            });
        }
        let maximum_work = maximum_bytes.checked_mul(32).ok_or(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_u64(maximum_bytes),
            path,
        })?;
        let remaining_wire_bytes =
            remaining_usize(remaining.wire_bytes, LimitKind::WireWork, path)?;
        let remaining_wire_fields =
            remaining_usize(remaining.wire_fields, LimitKind::WireFields, path)?;
        let remaining_wire_work = remaining_usize(remaining.wire_work, LimitKind::WireWork, path)?;
        let remaining_references =
            remaining_usize(remaining.references, LimitKind::References, path)?;
        Ok(Self {
            path,
            maximum_bytes,
            maximum_source_bytes: remaining_wire_bytes,
            maximum_fields: maximum_bytes.min(remaining_wire_fields),
            maximum_work: maximum_work.min(remaining_wire_work),
            maximum_references: source
                .state
                .options
                .semantic()
                .max_references()
                .min(remaining_references),
            maximum_text: source.state.options.semantic().max_output_text_bytes(),
            usage: CodecUsage::default(),
            maximum_retained_elements: remaining_usize(
                remaining.retained_elements,
                LimitKind::RetainedElements,
                path,
            )?,
            maximum_retained_bytes: remaining_usize(
                remaining.retained_bytes,
                LimitKind::RetainedBytes,
                path,
            )?,
            maximum_scratch_bytes: remaining_usize(
                remaining.peak_scratch_bytes,
                LimitKind::PeakScratchBytes,
                path,
            )?,
            maximum_allocation_events: remaining_usize(
                remaining.allocation_events,
                LimitKind::TransactionWork,
                path,
            )?,
            retained_elements: 0,
            retained_bytes: 0,
            current_scratch_bytes: 0,
            peak_scratch_bytes: 0,
            allocation_events: 0,
        })
    }

    fn options(&self, payload_bytes: usize) -> Result<storage_codec::DecodeOptions, Error> {
        if payload_bytes > self.maximum_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: usize_u64(payload_bytes),
                maximum: usize_u64(self.maximum_bytes),
                path: self.path,
            });
        }
        let remaining_source = remaining_budget(
            self.maximum_source_bytes,
            self.usage.source_bytes,
            LimitKind::WireWork,
            self.path,
        )?;
        if payload_bytes > remaining_source {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: usize_u64(
                    self.usage
                        .source_bytes
                        .checked_add(payload_bytes)
                        .ok_or(Error::InvalidSource { path: self.path })?,
                ),
                maximum: usize_u64(self.maximum_source_bytes),
                path: self.path,
            });
        }
        Ok(storage_codec::DecodeOptions::new(
            payload_bytes,
            remaining_budget(
                self.maximum_fields,
                self.usage.fields,
                LimitKind::WireFields,
                self.path,
            )?,
            remaining_budget(
                self.maximum_work,
                self.usage.work_bytes,
                LimitKind::WireWork,
                self.path,
            )?,
            64,
            remaining_budget(
                self.maximum_references,
                self.usage.references,
                LimitKind::References,
                self.path,
            )?,
            remaining_budget(
                self.maximum_text,
                self.usage.text_bytes,
                LimitKind::RetainedBytes,
                self.path,
            )?,
        ))
    }

    fn charge(&mut self, report: storage_codec::DecodeReport) -> Result<(), Error> {
        self.usage.source_bytes = checked_limit_add(
            self.usage.source_bytes,
            report.source_bytes(),
            self.maximum_source_bytes,
            LimitKind::WireWork,
            self.path,
        )?;
        self.usage.fields = checked_limit_add(
            self.usage.fields,
            report.fields(),
            self.maximum_fields,
            LimitKind::WireFields,
            self.path,
        )?;
        self.usage.work_bytes = checked_limit_add(
            self.usage.work_bytes,
            report.work_bytes(),
            self.maximum_work,
            LimitKind::WireWork,
            self.path,
        )?;
        self.usage.references = checked_limit_add(
            self.usage.references,
            report.references(),
            self.maximum_references,
            LimitKind::References,
            self.path,
        )?;
        self.usage.reference_bytes = checked_add(
            self.usage.reference_bytes,
            report.reference_bytes(),
            self.path,
        )?;
        self.usage.text_bytes = checked_limit_add(
            self.usage.text_bytes,
            report.text_bytes(),
            self.maximum_text,
            LimitKind::RetainedBytes,
            self.path,
        )?;
        self.usage.max_depth = self.usage.max_depth.max(report.max_depth());
        Ok(())
    }

    fn charge_wire_scan(
        &mut self,
        source_bytes: usize,
        fields: usize,
        max_depth: u32,
    ) -> Result<(), Error> {
        self.usage.source_bytes = checked_limit_add(
            self.usage.source_bytes,
            source_bytes,
            self.maximum_source_bytes,
            LimitKind::WireWork,
            self.path,
        )?;
        self.usage.fields = checked_limit_add(
            self.usage.fields,
            fields,
            self.maximum_fields,
            LimitKind::WireFields,
            self.path,
        )?;
        self.usage.work_bytes = checked_limit_add(
            self.usage.work_bytes,
            source_bytes,
            self.maximum_work,
            LimitKind::WireWork,
            self.path,
        )?;
        self.usage.max_depth = self.usage.max_depth.max(max_depth);
        Ok(())
    }

    fn charge_auxiliary_work(&mut self, work: usize) -> Result<(), Error> {
        self.usage.work_bytes = checked_limit_add(
            self.usage.work_bytes,
            work,
            self.maximum_work,
            LimitKind::WireWork,
            self.path,
        )?;
        Ok(())
    }

    fn reserve_retained<T>(
        &mut self,
        amount: usize,
        allocation_kind: LimitKind,
    ) -> Result<Vec<T>, Error> {
        let bytes = amount
            .checked_mul(size_of::<T>())
            .ok_or(Error::LimitExceeded {
                kind: LimitKind::RetainedBytes,
                observed: u64::MAX,
                maximum: usize_u64(self.maximum_retained_bytes),
                path: self.path,
            })?;
        let next_elements = checked_limit_add(
            self.retained_elements,
            amount,
            self.maximum_retained_elements,
            LimitKind::RetainedElements,
            self.path,
        )?;
        let next_bytes = checked_limit_add(
            self.retained_bytes,
            bytes,
            self.maximum_retained_bytes,
            LimitKind::RetainedBytes,
            self.path,
        )?;
        let event = usize::from(amount != 0 && size_of::<T>() != 0);
        let next_events = checked_limit_add(
            self.allocation_events,
            event,
            self.maximum_allocation_events,
            LimitKind::TransactionWork,
            self.path,
        )?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(amount)
            .map_err(|_allocation| Error::Allocation {
                kind: allocation_kind,
                amount,
            })?;
        if values.capacity() != amount && size_of::<T>() != 0 {
            return Err(Error::InvalidSource { path: self.path });
        }
        self.retained_elements = next_elements;
        self.retained_bytes = next_bytes;
        self.allocation_events = next_events;
        Ok(values)
    }

    fn reserve_scratch<T>(
        &mut self,
        amount: usize,
        allocation_kind: LimitKind,
    ) -> Result<Vec<T>, Error> {
        let bytes = amount
            .checked_mul(size_of::<T>())
            .ok_or(Error::LimitExceeded {
                kind: LimitKind::PeakScratchBytes,
                observed: u64::MAX,
                maximum: usize_u64(self.maximum_scratch_bytes),
                path: self.path,
            })?;
        let next_scratch = checked_limit_add(
            self.current_scratch_bytes,
            bytes,
            self.maximum_scratch_bytes,
            LimitKind::PeakScratchBytes,
            self.path,
        )?;
        let event = usize::from(amount != 0 && size_of::<T>() != 0);
        let next_events = checked_limit_add(
            self.allocation_events,
            event,
            self.maximum_allocation_events,
            LimitKind::TransactionWork,
            self.path,
        )?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(amount)
            .map_err(|_allocation| Error::Allocation {
                kind: allocation_kind,
                amount,
            })?;
        if values.capacity() != amount && size_of::<T>() != 0 {
            return Err(Error::InvalidSource { path: self.path });
        }
        self.current_scratch_bytes = next_scratch;
        self.peak_scratch_bytes = self.peak_scratch_bytes.max(next_scratch);
        self.allocation_events = next_events;
        Ok(values)
    }

    fn reserve_retained_growth<T>(
        &mut self,
        values: &mut Vec<T>,
        additional: usize,
        allocation_kind: LimitKind,
    ) -> Result<(), Error> {
        let spare = values
            .capacity()
            .checked_sub(values.len())
            .ok_or(Error::InvalidSource { path: self.path })?;
        if additional <= spare {
            return Ok(());
        }
        let minimum = values
            .len()
            .checked_add(additional)
            .ok_or(Error::InvalidSource { path: self.path })?;
        let doubled = values
            .capacity()
            .checked_mul(2)
            .ok_or(Error::InvalidSource { path: self.path })?;
        let target = minimum.max(doubled.max(4));
        let growth = target
            .checked_sub(values.capacity())
            .ok_or(Error::InvalidSource { path: self.path })?;
        let bytes = growth
            .checked_mul(size_of::<T>())
            .ok_or(Error::InvalidSource { path: self.path })?;
        let next_elements = checked_limit_add(
            self.retained_elements,
            growth,
            self.maximum_retained_elements,
            LimitKind::RetainedElements,
            self.path,
        )?;
        let next_bytes = checked_limit_add(
            self.retained_bytes,
            bytes,
            self.maximum_retained_bytes,
            LimitKind::RetainedBytes,
            self.path,
        )?;
        let event = usize::from(growth != 0 && size_of::<T>() != 0);
        let next_events = checked_limit_add(
            self.allocation_events,
            event,
            self.maximum_allocation_events,
            LimitKind::TransactionWork,
            self.path,
        )?;
        values
            .try_reserve_exact(
                target
                    .checked_sub(values.len())
                    .ok_or(Error::InvalidSource { path: self.path })?,
            )
            .map_err(|_allocation| Error::Allocation {
                kind: allocation_kind,
                amount: additional,
            })?;
        if values.capacity() != target && size_of::<T>() != 0 {
            return Err(Error::InvalidSource { path: self.path });
        }
        self.retained_elements = next_elements;
        self.retained_bytes = next_bytes;
        self.allocation_events = next_events;
        Ok(())
    }

    fn reserve_scratch_growth<T>(
        &mut self,
        values: &mut Vec<T>,
        additional: usize,
        allocation_kind: LimitKind,
    ) -> Result<(), Error> {
        let spare = values
            .capacity()
            .checked_sub(values.len())
            .ok_or(Error::InvalidSource { path: self.path })?;
        if additional <= spare {
            return Ok(());
        }
        let minimum = values
            .len()
            .checked_add(additional)
            .ok_or(Error::InvalidSource { path: self.path })?;
        let doubled = values
            .capacity()
            .checked_mul(2)
            .ok_or(Error::InvalidSource { path: self.path })?;
        let target = minimum.max(doubled.max(4));
        let growth = target
            .checked_sub(values.capacity())
            .ok_or(Error::InvalidSource { path: self.path })?;
        let bytes = growth
            .checked_mul(size_of::<T>())
            .ok_or(Error::InvalidSource { path: self.path })?;
        let next_scratch = checked_limit_add(
            self.current_scratch_bytes,
            bytes,
            self.maximum_scratch_bytes,
            LimitKind::PeakScratchBytes,
            self.path,
        )?;
        let event = usize::from(growth != 0 && size_of::<T>() != 0);
        let next_events = checked_limit_add(
            self.allocation_events,
            event,
            self.maximum_allocation_events,
            LimitKind::TransactionWork,
            self.path,
        )?;
        values
            .try_reserve_exact(
                target
                    .checked_sub(values.len())
                    .ok_or(Error::InvalidSource { path: self.path })?,
            )
            .map_err(|_allocation| Error::Allocation {
                kind: allocation_kind,
                amount: additional,
            })?;
        if values.capacity() != target && size_of::<T>() != 0 {
            return Err(Error::InvalidSource { path: self.path });
        }
        self.current_scratch_bytes = next_scratch;
        self.peak_scratch_bytes = self.peak_scratch_bytes.max(next_scratch);
        self.allocation_events = next_events;
        Ok(())
    }

    fn release_scratch<T>(&mut self, values: &Vec<T>) -> Result<(), Error> {
        self.release_scratch_capacity::<T>(values.capacity())
    }

    fn release_scratch_capacity<T>(&mut self, capacity: usize) -> Result<(), Error> {
        let bytes = capacity
            .checked_mul(size_of::<T>())
            .ok_or(Error::InvalidSource { path: self.path })?;
        self.current_scratch_bytes = self
            .current_scratch_bytes
            .checked_sub(bytes)
            .ok_or(Error::InvalidSource { path: self.path })?;
        Ok(())
    }
}

fn remaining_usize(value: u64, kind: LimitKind, path: Path) -> Result<usize, Error> {
    usize::try_from(value).map_err(|_error| Error::LimitExceeded {
        kind,
        observed: value,
        maximum: usize_u64(usize::MAX),
        path,
    })
}

fn remaining_budget(
    maximum: usize,
    used: usize,
    kind: LimitKind,
    path: Path,
) -> Result<usize, Error> {
    let remaining = maximum.checked_sub(used).ok_or(Error::LimitExceeded {
        kind,
        observed: usize_u64(used),
        maximum: usize_u64(maximum),
        path,
    })?;
    if remaining == 0 {
        return Err(Error::LimitExceeded {
            kind,
            observed: usize_u64(used),
            maximum: usize_u64(maximum),
            path,
        });
    }
    Ok(remaining)
}

fn checked_add(left: usize, right: usize, path: Path) -> Result<usize, Error> {
    left.checked_add(right).ok_or(Error::LimitExceeded {
        kind: LimitKind::TransactionWork,
        observed: u64::MAX,
        maximum: u64::MAX - 1,
        path,
    })
}

fn count_not_overflowed(overflowed: bool, path: Path) -> Result<(), Error> {
    if overflowed {
        return Err(Error::LimitExceeded {
            kind: LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        });
    }
    Ok(())
}

fn checked_limit_add(
    left: usize,
    right: usize,
    maximum: usize,
    kind: LimitKind,
    path: Path,
) -> Result<usize, Error> {
    let value = left.checked_add(right).ok_or(Error::LimitExceeded {
        kind,
        observed: u64::MAX,
        maximum: usize_u64(maximum),
        path,
    })?;
    if value > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: usize_u64(value),
            maximum: usize_u64(maximum),
            path,
        });
    }
    Ok(value)
}

const fn selected_path(native: table_headers::Target) -> Path {
    Path::Table {
        sheet: saturating_u32(native.sheet_position),
        table: saturating_u32(native.table_position),
    }
}

const fn saturating_u32(value: usize) -> u32 {
    if value > u32::MAX as usize {
        u32::MAX
    } else {
        value as u32
    }
}

const fn usize_u64(value: usize) -> u64 {
    value as u64
}

fn map_codec_failure(error: storage_codec::DecodeError, path: Path) -> Error {
    match error.resource_limit() {
        Some(limit) => {
            let (kind, observed, maximum) = match limit {
                storage_codec::DecodeLimit::Bytes { observed, maximum }
                | storage_codec::DecodeLimit::Fields { observed, maximum }
                | storage_codec::DecodeLimit::Work { observed, maximum }
                | storage_codec::DecodeLimit::References { observed, maximum }
                | storage_codec::DecodeLimit::Text { observed, maximum } => {
                    (LimitKind::WireWork, usize_u64(observed), usize_u64(maximum))
                },
                storage_codec::DecodeLimit::Nesting { observed, maximum } => {
                    (LimitKind::WireWork, u64::from(observed), u64::from(maximum))
                },
                _ => (LimitKind::WireWork, u64::MAX, u64::MAX - 1),
            };
            Error::LimitExceeded {
                kind,
                observed,
                maximum,
                path,
            }
        },
        None => Error::InvalidSource { path },
    }
}

fn map_header_failure(error: table_headers::Error) -> Error {
    use table_headers::Error as HeaderError;

    match error {
        HeaderError::SheetNotFound => Error::SheetNotFound,
        HeaderError::TableNotFound => Error::TableNotFound,
        HeaderError::TableLocked { path } => Error::TableLocked {
            path: map_header_path(path),
        },
        HeaderError::UnsupportedSource => Error::UnsupportedSource {
            path: Path::Package,
        },
        HeaderError::Allocation { amount, .. } => Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount,
        },
        HeaderError::LimitExceeded {
            observed,
            maximum,
            path,
            ..
        } => Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed,
            maximum,
            path: map_header_path(path),
        },
        HeaderError::InvalidSettings { path, .. }
        | HeaderError::UnsupportedDependency { path }
        | HeaderError::InvalidSource { path } => Error::InvalidSource {
            path: map_header_path(path),
        },
        HeaderError::Verification | HeaderError::PatchConflict => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

const fn map_header_path(path: table_headers::Path) -> Path {
    match path {
        table_headers::Path::Package => Path::Package,
        table_headers::Path::Table { sheet, table } => Path::Table {
            sheet: saturating_u32(sheet),
            table: saturating_u32(table),
        },
    }
}

fn validate_deferred_owner_closure(
    markers: ModelOwnerMarkers,
    conditional_owner_matches: usize,
    spill_owner_matches: usize,
    active_spill_owner: bool,
    haunted_owner_matches: usize,
    path: Path,
) -> Result<(), Error> {
    if markers.conditional_style.is_some() && conditional_owner_matches != 1 {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::ConditionalStyle,
        });
    }
    if markers.spill.is_some() && active_spill_owner {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::Spill,
        });
    }
    if markers.spill.is_some() && spill_owner_matches != 1 {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::Spill,
        });
    }
    if markers.haunted.is_some() && haunted_owner_matches != 1 {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::FormulaCache,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use super::*;

    #[test]
    fn native_basic_body_admits_and_header_name_coordinate_refuses() {
        let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers");
        let source = Package::open(fixture).expect("open native basic fixture");
        let body = CellPosition::from_a1("B3").expect("valid body coordinate");
        resolve_changed_target(&source, 0, 0, &[body]).expect("body admission");

        let header = CellPosition::from_a1("A1").expect("valid header coordinate");
        assert!(matches!(
            resolve_changed_target(&source, 0, 0, &[header]),
            Err(Error::UnsupportedDependency {
                kind: DependencyKind::HeaderNameIndex,
                ..
            })
        ));
    }

    #[test]
    #[ignore = "requires external Numbers 14.4 formula-rich oracle"]
    fn formula_rich_oracle_c2_admits() {
        let fixture = PathBuf::from(
            "/private/tmp/litchi-numbers-cell-batch-native.wuaiMp/oracle-preserved.numbers",
        );
        let source = Package::open(fixture).expect("open formula-rich oracle");
        let position = CellPosition::from_a1("C2").expect("valid formula-rich coordinate");
        resolve_changed_target(&source, 0, 0, &[position]).expect("formula-rich admission");
    }

    #[test]
    fn changed_positions_must_be_sorted_unique_and_in_bounds() {
        let path = Path::Table { sheet: 0, table: 0 };
        let a = CellPosition::new(0, 0);
        let b = CellPosition::new(1, 0);
        assert!(validate_positions(&[a, b], 2, 1, path).is_ok());
        assert!(matches!(
            validate_positions(&[a, a], 2, 1, path),
            Err(Error::DuplicatePosition { position }) if position == a
        ));
        assert!(matches!(
            validate_positions(&[b, a], 2, 1, path),
            Err(Error::InvalidSource { .. })
        ));
        assert!(matches!(
            validate_positions(&[CellPosition::new(2, 0)], 2, 1, path),
            Err(Error::OutOfBounds { .. })
        ));
    }

    #[test]
    fn role_aliases_fail_closed_except_same_format_role() {
        let native = table_headers::Target {
            sheet_position: 0,
            table_position: 0,
            model_identifier: 4,
            sheet_identifier: 2,
            drawable_identifier: 3,
            drawable_position: 0,
            sheet_component_index: 0,
            sheet_object_index: 0,
            sheet_message_index: 0,
            sheet_message_type: 2,
            info_component_index: 0,
            info_object_index: 0,
            info_message_index: 0,
            info_message_type: 6_000,
            component_index: 0,
            object_index: 0,
            message_index: 0,
            message_type: 6_001,
            settings: crate::table::headers::Settings::default(),
            rows: 1,
            columns: 1,
            locked: LockState::Unlocked,
        };
        let path = selected_path(native);
        let mut budget = test_resolution_budget(path);
        let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
        assert!(
            identities
                .claim(20, Role::Format, false, &mut budget, path)
                .is_ok()
        );
        assert!(
            identities
                .claim(20, Role::Format, true, &mut budget, path)
                .is_ok()
        );
        identities
            .finish(&mut budget, path)
            .expect("same-role alias is valid");

        let mut budget = test_resolution_budget(path);
        let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
        identities
            .claim(20, Role::Format, false, &mut budget, path)
            .unwrap();
        identities
            .claim(20, Role::String, false, &mut budget, path)
            .unwrap();
        assert!(matches!(
            identities.finish(&mut budget, path),
            Err(Error::InvalidSource { .. })
        ));

        let mut budget = test_resolution_budget(path);
        let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
        identities
            .claim(4, Role::Tile, false, &mut budget, path)
            .unwrap();
        assert!(matches!(
            identities.finish(&mut budget, path),
            Err(Error::InvalidSource { .. })
        ));
    }

    #[test]
    fn reverse_order_role_claims_scale_linearly_from_4k_to_8k() {
        let native = test_native_target();
        let path = selected_path(native);
        fn run(
            native: table_headers::Target,
            path: Path,
            count: u64,
        ) -> (RoleIdentityWork, RoleIdentities) {
            let mut budget = test_resolution_budget(path);
            let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
            for identifier in (10_000..10_000 + count).rev() {
                identities
                    .claim(identifier, Role::Tile, false, &mut budget, path)
                    .unwrap();
            }
            let work = identities.finish(&mut budget, path).unwrap();
            (work, identities)
        }
        let (work_4k, identities_4k) = run(native, path, 4_096);
        let (work_8k, identities_8k) = run(native, path, 8_192);
        assert!(work_8k.movements * 100 <= work_4k.movements * 220);
        assert!(work_8k.work * 100 <= work_4k.work * 220);
        assert!(identities_4k.lookup_comparisons(u64::MAX) <= 13);
        assert!(
            identities_8k.lookup_comparisons(u64::MAX)
                <= identities_4k.lookup_comparisons(u64::MAX) + 1
        );
        assert!(identities_8k.is_known_owner(native.model_identifier));
    }

    #[test]
    fn role_finish_max_minus_one_refuses_before_scratch_or_movement() {
        let native = test_native_target();
        let path = selected_path(native);
        let mut budget = test_resolution_budget(path);
        let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
        for identifier in (30_000..34_096).rev() {
            identities
                .claim(identifier, Role::Tile, false, &mut budget, path)
                .unwrap();
        }
        let required = RoleIdentities::finish_usage(identities.claimed.len(), path)
            .unwrap()
            .total(path)
            .unwrap();
        budget.maximum_work = budget.usage.work_bytes.checked_add(required - 1).unwrap();
        let allocation_events = budget.allocation_events;
        let current_scratch_bytes = budget.current_scratch_bytes;
        let peak_scratch_bytes = budget.peak_scratch_bytes;
        let identity_usage = identities.usage;
        let first = identities.claimed.first().copied();
        let last = identities.claimed.last().copied();

        assert!(matches!(
            identities.finish(&mut budget, path),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                ..
            })
        ));
        assert_eq!(budget.allocation_events, allocation_events);
        assert_eq!(budget.current_scratch_bytes, current_scratch_bytes);
        assert_eq!(budget.peak_scratch_bytes, peak_scratch_bytes);
        assert_eq!(identities.usage, identity_usage);
        assert_eq!(identities.claimed.first().copied(), first);
        assert_eq!(identities.claimed.last().copied(), last);
        assert!(!identities.finalized);
    }

    #[test]
    fn hostile_duplicate_identity_claims_do_not_mutate_the_sorted_index() {
        let native = test_native_target();
        let path = selected_path(native);
        let mut budget = test_resolution_budget(path);
        let mut identities = RoleIdentities::new(native, &mut budget, path).unwrap();
        for identifier in (20_000..28_192).rev() {
            identities
                .claim(identifier, Role::Tile, false, &mut budget, path)
                .unwrap();
        }
        for identifier in [20_000, 24_096, 28_191] {
            identities
                .claim(identifier, Role::String, true, &mut budget, path)
                .unwrap();
        }
        let before = identities.claimed.clone();
        assert!(matches!(
            identities.finish(&mut budget, path),
            Err(Error::InvalidSource { .. })
        ));
        assert_eq!(identities.claimed, before);
        for identifier in [20_000, 24_096, 28_191] {
            assert!(
                before
                    .iter()
                    .any(|claim| claim.identifier == identifier && claim.role == Role::String)
            );
        }
    }

    fn test_native_target() -> table_headers::Target {
        table_headers::Target {
            sheet_position: 0,
            table_position: 0,
            model_identifier: 4,
            sheet_identifier: 2,
            drawable_identifier: 3,
            drawable_position: 0,
            sheet_component_index: 0,
            sheet_object_index: 0,
            sheet_message_index: 0,
            sheet_message_type: 2,
            info_component_index: 0,
            info_object_index: 0,
            info_message_index: 0,
            info_message_type: 6_000,
            component_index: 0,
            object_index: 0,
            message_index: 0,
            message_type: 6_001,
            settings: crate::table::headers::Settings::default(),
            rows: 1,
            columns: 1,
            locked: LockState::Unlocked,
        }
    }

    fn test_resolution_budget(path: Path) -> ResolutionBudget {
        ResolutionBudget {
            path,
            maximum_bytes: usize::MAX,
            maximum_source_bytes: usize::MAX,
            maximum_fields: usize::MAX,
            maximum_work: usize::MAX,
            maximum_references: usize::MAX,
            maximum_text: usize::MAX,
            usage: CodecUsage::default(),
            maximum_retained_elements: usize::MAX,
            maximum_retained_bytes: usize::MAX,
            maximum_scratch_bytes: usize::MAX,
            maximum_allocation_events: usize::MAX,
            retained_elements: 0,
            retained_bytes: 0,
            current_scratch_bytes: 0,
            peak_scratch_bytes: 0,
            allocation_events: 0,
        }
    }

    #[test]
    fn codec_usage_enforces_one_shared_limit() {
        let path = Path::Package;
        assert_eq!(
            checked_limit_add(4, 5, 9, LimitKind::WireWork, path).unwrap(),
            9
        );
        assert!(matches!(
            checked_limit_add(4, 6, 9, LimitKind::WireWork, path),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: 10,
                maximum: 9,
                ..
            })
        ));
    }

    #[test]
    fn resolver_retained_and_scratch_max_minus_one_refuse_before_allocation() {
        let path = Path::Package;
        let bytes = size_of::<RichFieldRefs>();

        let mut retained = test_resolution_budget(path);
        retained.maximum_retained_bytes = bytes - 1;
        assert!(matches!(
            retained.reserve_retained::<RichFieldRefs>(1, LimitKind::RetainedElements),
            Err(Error::LimitExceeded {
                kind: LimitKind::RetainedBytes,
                ..
            })
        ));
        assert_eq!(retained.allocation_events, 0);
        assert_eq!(retained.retained_bytes, 0);

        let mut scratch = test_resolution_budget(path);
        scratch.maximum_scratch_bytes = bytes - 1;
        assert!(matches!(
            scratch.reserve_scratch::<RichFieldRefs>(1, LimitKind::RetainedElements),
            Err(Error::LimitExceeded {
                kind: LimitKind::PeakScratchBytes,
                ..
            })
        ));
        assert_eq!(scratch.allocation_events, 0);
        assert_eq!(scratch.current_scratch_bytes, 0);
        assert_eq!(scratch.peak_scratch_bytes, 0);
    }

    #[test]
    fn exhausted_decode_budget_refuses_before_visitor_entry() {
        #[derive(Default)]
        struct Probe(bool);
        impl storage_codec::StorageVisitor for Probe {
            fn visit_list_segment(
                &mut self,
                _record: storage_codec::ReferenceRecord<'_>,
            ) -> Result<(), storage_codec::DecodeError> {
                self.0 = true;
                Ok(())
            }
        }

        let path = Path::Package;
        let mut budget = test_resolution_budget(path);
        budget.maximum_fields = 1;
        budget.usage.fields = 1;
        let probe = Probe::default();
        assert!(matches!(
            budget.options(1),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireFields,
                ..
            })
        ));
        assert!(!probe.0);
    }

    #[test]
    fn rich_entry_collector_refuses_n_plus_one_before_push() {
        let mut collector = RichEntryCollector {
            entries: Vec::new(),
            limit: 4_096,
            accepted: 0,
            invalid: false,
        };
        for _ in 0..4_096 {
            assert!(collector.reserve_entry_slot());
        }
        assert_eq!(collector.accepted, 4_096);
        assert!(!collector.reserve_entry_slot());
        assert_eq!(collector.accepted, 4_096);
        assert!(collector.invalid);
        assert!(collector.entries.is_empty());
    }

    #[test]
    fn owner_markers_require_a_rooted_inert_closure() {
        let path = Path::Package;
        let marker = UuidKey { lower: 1, upper: 2 };
        for markers in [
            ModelOwnerMarkers {
                conditional_style: Some(marker),
                ..ModelOwnerMarkers::default()
            },
            ModelOwnerMarkers {
                spill: Some(marker),
                ..ModelOwnerMarkers::default()
            },
            ModelOwnerMarkers {
                haunted: Some(marker),
                ..ModelOwnerMarkers::default()
            },
        ] {
            assert!(matches!(
                reject_unresolved_owner_markers(markers, path),
                Err(Error::UnsupportedDependency { .. })
            ));
        }
    }

    #[test]
    fn spill_records_are_active_even_when_the_record_payload_is_empty() {
        let path = Path::Package;
        assert!(!spill_sizes_active(&[], path).expect("canonical empty archive"));
        assert!(spill_sizes_active(&[0x0a, 0x00], path).expect("one canonical empty record"));
        assert!(matches!(
            spill_sizes_active(&[0x12, 0x00], path),
            Err(Error::UnsupportedDependency {
                kind: DependencyKind::Spill,
                ..
            })
        ));
    }

    #[test]
    fn owner_uuid_varints_must_use_the_canonical_encoding() {
        assert_eq!(
            parse_uuid(&[0x08, 0x01, 0x10, 0x02]),
            Some(UuidKey { lower: 1, upper: 2 })
        );
        assert_eq!(parse_uuid(&[0x08, 0x81, 0x00, 0x10, 0x02]), None);
    }
}
