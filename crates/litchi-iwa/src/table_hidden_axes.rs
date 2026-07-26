//! Typed hidden-row and hidden-column state shared by native iWork tables.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::protobuf::{tsce, tsp, tst};
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const TABLE_INFO_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_INFO_HIDDEN_STATES_UUID_FIELD: u32 = 8;
const TABLE_MODEL_HIDDEN_ROWS_FIELD: u32 = 14;
const TABLE_MODEL_HIDDEN_COLUMNS_FIELD: u32 = 15;
const TABLE_MODEL_COLUMN_HIDDEN_FORMULA_OWNER_FIELD: u32 = 34;
const TABLE_MODEL_ROW_HIDDEN_FORMULA_OWNER_FIELD: u32 = 35;
const TABLE_MODEL_USER_HIDDEN_ROWS_FIELD: u32 = 41;
const TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD: u32 = 42;
const TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD: u32 = 70;
pub(crate) const HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE: u32 = 6_204;
pub(crate) const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const COLUMN_HIDDEN_EXTENT_UID_OFFSET: u64 = 7;

/// One zero-based row or column position in a native iWork table.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum TableAxisIndex {
    /// A zero-based row index.
    Row(usize),
    /// A zero-based column index.
    Column(usize),
}

impl TableAxisIndex {
    /// Address one row by zero-based index.
    pub const fn row(index: usize) -> Self {
        Self::Row(index)
    }

    /// Address one column by zero-based index.
    pub const fn column(index: usize) -> Self {
        Self::Column(index)
    }

    /// Return the zero-based index within this position's axis.
    pub const fn index(self) -> usize {
        match self {
            Self::Row(index) | Self::Column(index) => index,
        }
    }
}

/// Canonical, duplicate-free set of user-hidden table axes.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableHiddenAxes {
    axes: Vec<TableAxisIndex>,
}

impl TableHiddenAxes {
    /// No rows or columns are hidden.
    pub const fn empty() -> Self {
        Self { axes: Vec::new() }
    }

    /// Construct a sorted hidden-axis set, rejecting duplicate positions.
    pub fn new(axes: impl IntoIterator<Item = TableAxisIndex>) -> Result<Self> {
        let mut axes = axes.into_iter().collect::<Vec<_>>();
        axes.sort_unstable();
        if axes.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::ParseError(
                "hidden iWork table axes contain a duplicate position".to_owned(),
            ));
        }
        Ok(Self { axes })
    }

    /// Borrow the canonical row-then-column positions.
    pub fn as_slice(&self) -> &[TableAxisIndex] {
        &self.axes
    }

    /// Iterate over canonical row-then-column positions.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = TableAxisIndex> + '_ {
        self.axes.iter().copied()
    }

    /// Return whether every table axis is visible.
    pub fn is_empty(&self) -> bool {
        self.axes.is_empty()
    }

    /// Return whether one row or column is hidden.
    pub fn contains(&self, axis: TableAxisIndex) -> bool {
        self.axes.binary_search(&axis).is_ok()
    }
}

pub(crate) fn table_hidden_axes(
    package: &IWorkPackage,
    model_object_id: u64,
) -> Result<TableHiddenAxes> {
    let graph = table_hidden_graph(package, model_object_id)?;
    hidden_axes_from_graph(&graph)
}

pub(crate) fn set_table_hidden_axes(
    package: &mut IWorkPackage,
    model_object_id: u64,
    hidden: &TableHiddenAxes,
) -> Result<()> {
    let graph = table_hidden_graph(package, model_object_id)?;
    validate_axis_bounds(&graph.model, hidden)?;
    if hidden_axes_from_graph(&graph)? == *hidden {
        return Ok(());
    }
    if graph.model.hidden_states_owner.is_some()
        && (graph.model.hidden_state_formula_owner_for_columns.is_none()
            || graph.model.hidden_state_formula_owner_for_rows.is_none())
    {
        return Err(Error::InvalidFormat(
            "iWork table hidden-state owner is missing its formula owners".to_owned(),
        ));
    }

    let mut owner = graph.model.hidden_states_owner.clone();
    let mut info_uuid = graph.info.hidden_states_uuid;
    let mut new_object_ids = None;
    if owner.is_none() {
        if hidden.is_empty() {
            return Ok(());
        }
        let object_ids = HiddenStateObjectIds::allocate(package)?;
        let hidden_states_uid = graph.formula_owner_uid;
        let column_extent_uid = tsp::Uuid {
            lower: hidden_states_uid
                .lower
                .wrapping_add(COLUMN_HIDDEN_EXTENT_UID_OFFSET),
            upper: hidden_states_uid.upper,
        };
        info_uuid = Some(hidden_states_uid);
        owner = Some(tst::HiddenStatesOwnerArchive {
            owner_uid: hidden_states_uid,
            hidden_states: vec![tst::HiddenStatesArchive {
                hidden_states_uid,
                column_hidden_state_extent: empty_extent(
                    column_extent_uid,
                    tst::hidden_state_extent_archive::RowOrColumnDirection::ColumnDirection,
                    object_ids.filter_columns,
                ),
                row_hidden_state_extent: empty_extent(
                    hidden_states_uid,
                    tst::hidden_state_extent_archive::RowOrColumnDirection::RowDirection,
                    object_ids.filter_rows,
                ),
            }],
        });
        new_object_ids = Some(object_ids);
    }

    let owner = owner.as_mut().ok_or_else(|| {
        Error::InvalidFormat("iWork table hidden-state owner is absent".to_owned())
    })?;
    let active_uuid = info_uuid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table info {} has no active hidden-state UUID",
            graph.info_object_id
        ))
    })?;
    let active = unique_active_state_mut(owner, active_uuid)?;
    update_extent(
        &mut active.row_hidden_state_extent,
        &graph.row_uids,
        hidden.iter().filter_map(|axis| match axis {
            TableAxisIndex::Row(index) => Some(index),
            TableAxisIndex::Column(_) => None,
        }),
        "row",
    )?;
    update_extent(
        &mut active.column_hidden_state_extent,
        &graph.column_uids,
        hidden.iter().filter_map(|axis| match axis {
            TableAxisIndex::Column(index) => Some(index),
            TableAxisIndex::Row(_) => None,
        }),
        "column",
    )?;

    let mut staged = package.clone();
    if let Some(object_ids) = new_object_ids {
        insert_hidden_state_objects(&mut staged, &graph, owner, object_ids)?;
    }
    let active_uuid = info_uuid.as_ref().ok_or_else(|| {
        Error::InvalidFormat("iWork table hidden-state UUID is absent".to_owned())
    })?;
    patch_model(&mut staged, &graph, owner, active_uuid, new_object_ids)?;
    patch_info(&mut staged, &graph, active_uuid)?;
    if table_hidden_axes(&staged, model_object_id)? != *hidden {
        return Err(Error::InvalidFormat(
            "iWork table hidden axes failed round-trip validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

#[derive(Clone, Copy)]
struct HiddenStateObjectIds {
    filter_columns: u64,
    filter_rows: u64,
    formula_columns: u64,
    formula_rows: u64,
}

impl HiddenStateObjectIds {
    fn allocate(package: &IWorkPackage) -> Result<Self> {
        let filter_columns = next_object_identifier(package)?;
        let filter_rows = filter_columns.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("iWork hidden-state object identifier overflow".to_owned())
        })?;
        let formula_columns = filter_rows.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("iWork hidden-state object identifier overflow".to_owned())
        })?;
        let formula_rows = formula_columns.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("iWork hidden-state object identifier overflow".to_owned())
        })?;
        Ok(Self {
            filter_columns,
            filter_rows,
            formula_columns,
            formula_rows,
        })
    }
}

struct TableHiddenGraph {
    model_object_id: u64,
    model_archive: String,
    model_message_index: usize,
    model_message_type: u32,
    model: tst::TableModelArchive,
    info_archive: String,
    info_object_id: u64,
    info_message_index: usize,
    info_message_type: u32,
    info: tst::TableInfoArchive,
    formula_owner_uid: tsp::Uuid,
    row_uids: Vec<tsp::Uuid>,
    column_uids: Vec<tsp::Uuid>,
}

fn table_hidden_graph(package: &IWorkPackage, model_object_id: u64) -> Result<TableHiddenGraph> {
    let mut model_match = None;
    let mut info_match = None;
    let mut uid_maps = HashMap::new();
    let mut formula_owner_uids = HashMap::new();
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            let Some(object_id) = object.archive_info.identifier else {
                continue;
            };
            for (message_index, message) in object.messages.iter().enumerate() {
                if object_id == model_object_id
                    && TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_)
                    && let Ok(model) = tst::TableModelArchive::decode(message.data.as_slice())
                    && model_match
                        .replace((archive_name.to_owned(), message_index, message.type_, model))
                        .is_some()
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork table model {model_object_id} has multiple native payloads"
                    )));
                }
                if object_id != model_object_id
                    && TABLE_INFO_MESSAGE_TYPES.contains(&message.type_)
                    && let Ok(info) = tst::TableInfoArchive::decode(message.data.as_slice())
                    && info.table_model.identifier == model_object_id
                    && info_match
                        .replace((
                            archive_name.to_owned(),
                            object_id,
                            message_index,
                            message.type_,
                            info,
                        ))
                        .is_some()
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork table model {model_object_id} has multiple table-info owners"
                    )));
                }
                if let Ok(map) = tst::ColumnRowUidMapArchive::decode(message.data.as_slice()) {
                    uid_maps.insert(object_id, map);
                }
                if message.type_ == 4_008
                    && let Ok(owner) =
                        tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                    && let Some(formula_owner) = owner.formula_owner
                    && formula_owner.identifier != 0
                    && formula_owner_uids
                        .insert(formula_owner.identifier, owner.formula_owner_uid)
                        .is_some()
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork formula owner {} has multiple dependency records",
                        formula_owner.identifier
                    )));
                }
            }
        }
    }
    let (model_archive, model_message_index, model_message_type, model) =
        model_match.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table model {model_object_id} has no native payload"
            ))
        })?;
    let (info_archive, info_object_id, info_message_index, info_message_type, info) = info_match
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table model {model_object_id} has no table-info owner"
            ))
        })?;
    let formula_owner_uid = formula_owner_uids.remove(&info_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table info {info_object_id} has no formula-owner UUID"
        ))
    })?;
    let uid_map_id = info
        .view_column_row_uids
        .as_ref()
        .or(model.base_column_row_uids.as_ref())
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table model {model_object_id} has no stable axis UID map"
            ))
        })?;
    let uid_map = uid_maps.remove(&uid_map_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model {model_object_id} references missing axis UID map {uid_map_id}"
        ))
    })?;
    let row_uids = physical_uids(
        &uid_map.sorted_row_uids,
        &uid_map.row_index_for_uid,
        &uid_map.row_uid_for_index,
        usize::try_from(model.number_of_rows)
            .map_err(|_| Error::InvalidFormat("iWork row count exceeds usize".to_owned()))?,
        "row",
    )?;
    let column_uids = physical_uids(
        &uid_map.sorted_column_uids,
        &uid_map.column_index_for_uid,
        &uid_map.column_uid_for_index,
        usize::try_from(model.number_of_columns)
            .map_err(|_| Error::InvalidFormat("iWork column count exceeds usize".to_owned()))?,
        "column",
    )?;
    Ok(TableHiddenGraph {
        model_object_id,
        model_archive,
        model_message_index,
        model_message_type,
        model,
        info_archive,
        info_object_id,
        info_message_index,
        info_message_type,
        info,
        formula_owner_uid,
        row_uids,
        column_uids,
    })
}

fn physical_uids(
    sorted: &[tsp::Uuid],
    index_for_uid: &[u32],
    uid_for_index: &[u32],
    expected: usize,
    axis: &str,
) -> Result<Vec<tsp::Uuid>> {
    if sorted.len() != expected
        || index_for_uid.len() != expected
        || uid_for_index.len() != expected
    {
        return Err(Error::InvalidFormat(format!(
            "iWork {axis} UID map lengths do not match the table"
        )));
    }
    let mut seen = HashSet::with_capacity(expected);
    let mut physical = Vec::with_capacity(expected);
    for (index, &stable_index) in uid_for_index.iter().enumerate() {
        let stable_index = usize::try_from(stable_index)
            .map_err(|_| Error::InvalidFormat(format!("iWork {axis} UID exceeds usize")))?;
        let uid = sorted.get(stable_index).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork {axis} UID map references a missing UUID"))
        })?;
        let expected_index = u32::try_from(index)
            .map_err(|_| Error::InvalidFormat(format!("iWork {axis} index exceeds u32")))?;
        if index_for_uid.get(stable_index).copied() != Some(expected_index)
            || !seen.insert((uid.lower, uid.upper))
        {
            return Err(Error::InvalidFormat(format!(
                "iWork {axis} UID map is not a one-to-one permutation"
            )));
        }
        physical.push(*uid);
    }
    Ok(physical)
}

fn hidden_axes_from_graph(graph: &TableHiddenGraph) -> Result<TableHiddenAxes> {
    let Some(owner) = &graph.model.hidden_states_owner else {
        if graph.info.hidden_states_uuid.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork table info {} selects absent hidden states",
                graph.info_object_id
            )));
        }
        return Ok(TableHiddenAxes::empty());
    };
    if owner.owner_uid != graph.formula_owner_uid {
        return Err(Error::InvalidFormat(format!(
            "iWork table hidden-state owner does not match formula owner {}",
            graph.info_object_id
        )));
    }
    let active_uuid = graph.info.hidden_states_uuid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table info {} has no active hidden-state UUID",
            graph.info_object_id
        ))
    })?;
    let active = unique_active_state(owner, active_uuid)?;
    let mut axes = Vec::new();
    read_extent(
        &active.row_hidden_state_extent,
        &graph.row_uids,
        TableAxisIndex::Row,
        "row",
        &mut axes,
    )?;
    read_extent(
        &active.column_hidden_state_extent,
        &graph.column_uids,
        TableAxisIndex::Column,
        "column",
        &mut axes,
    )?;
    TableHiddenAxes::new(axes)
}

fn unique_active_state<'a>(
    owner: &'a tst::HiddenStatesOwnerArchive,
    active_uuid: &tsp::Uuid,
) -> Result<&'a tst::HiddenStatesArchive> {
    let mut states = owner
        .hidden_states
        .iter()
        .filter(|state| state.hidden_states_uid == *active_uuid);
    let state = states.next().ok_or_else(|| {
        Error::InvalidFormat("iWork table active hidden state is missing".to_owned())
    })?;
    if states.next().is_some() {
        return Err(Error::InvalidFormat(
            "iWork table active hidden state is duplicated".to_owned(),
        ));
    }
    Ok(state)
}

fn unique_active_state_mut<'a>(
    owner: &'a mut tst::HiddenStatesOwnerArchive,
    active_uuid: &tsp::Uuid,
) -> Result<&'a mut tst::HiddenStatesArchive> {
    let mut index = None;
    for (candidate, state) in owner.hidden_states.iter().enumerate() {
        if state.hidden_states_uid == *active_uuid && index.replace(candidate).is_some() {
            return Err(Error::InvalidFormat(
                "iWork table active hidden state is duplicated".to_owned(),
            ));
        }
    }
    let index = index.ok_or_else(|| {
        Error::InvalidFormat("iWork table active hidden state is missing".to_owned())
    })?;
    Ok(&mut owner.hidden_states[index])
}

fn read_extent(
    extent: &tst::HiddenStateExtentArchive,
    physical_uids: &[tsp::Uuid],
    axis: impl Fn(usize) -> TableAxisIndex,
    label: &str,
    output: &mut Vec<TableAxisIndex>,
) -> Result<()> {
    let expected_direction = match label {
        "row" => tst::hidden_state_extent_archive::RowOrColumnDirection::RowDirection,
        "column" => tst::hidden_state_extent_archive::RowOrColumnDirection::ColumnDirection,
        _ => unreachable!("validated hidden-axis label"),
    };
    if extent.row_or_column_direction != expected_direction as i32 {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} hidden-state extent has the wrong direction"
        )));
    }
    let indexes = physical_uids
        .iter()
        .enumerate()
        .map(|(index, uid)| ((uid.lower, uid.upper), index))
        .collect::<HashMap<_, _>>();
    let mut seen = HashSet::with_capacity(extent.base_hidden_states.len());
    for state in &extent.base_hidden_states {
        let key = (state.row_or_column_uid.lower, state.row_or_column_uid.upper);
        if !seen.insert(key) {
            return Err(Error::InvalidFormat(format!(
                "iWork table {label} hidden state contains duplicate UUIDs"
            )));
        }
        let index = indexes.get(&key).copied().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table {label} hidden state references an unknown UUID"
            ))
        })?;
        if state.user_hidden == Some(true) {
            output.push(axis(index));
        }
    }
    Ok(())
}

fn validate_axis_bounds(model: &tst::TableModelArchive, hidden: &TableHiddenAxes) -> Result<()> {
    let rows = usize::try_from(model.number_of_rows)
        .map_err(|_| Error::InvalidFormat("iWork row count exceeds usize".to_owned()))?;
    let columns = usize::try_from(model.number_of_columns)
        .map_err(|_| Error::InvalidFormat("iWork column count exceeds usize".to_owned()))?;
    for axis in hidden.iter() {
        let (index, length, label) = match axis {
            TableAxisIndex::Row(index) => (index, rows, "row"),
            TableAxisIndex::Column(index) => (index, columns, "column"),
        };
        if index >= length {
            return Err(Error::ParseError(format!(
                "Cannot hide iWork table {label} {index} in an axis of length {length}"
            )));
        }
    }
    Ok(())
}

fn update_extent(
    extent: &mut tst::HiddenStateExtentArchive,
    physical_uids: &[tsp::Uuid],
    hidden: impl Iterator<Item = usize>,
    axis: &str,
) -> Result<()> {
    let desired = hidden.collect::<HashSet<_>>();
    let indexes = physical_uids
        .iter()
        .enumerate()
        .map(|(index, uid)| ((uid.lower, uid.upper), index))
        .collect::<HashMap<_, _>>();
    let mut present = HashSet::with_capacity(extent.base_hidden_states.len());
    extent.base_hidden_states.retain_mut(|state| {
        let key = (state.row_or_column_uid.lower, state.row_or_column_uid.upper);
        let Some(index) = indexes.get(&key).copied() else {
            return true;
        };
        if desired.contains(&index) {
            state.user_hidden = Some(true);
            present.insert(index);
            true
        } else {
            state.user_hidden = None;
            state.filtered == Some(true) || state.pivot_hidden == Some(true)
        }
    });
    for &index in &desired {
        if present.contains(&index) {
            continue;
        }
        let uid = physical_uids.get(index).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork {axis} hidden index is out of bounds"))
        })?;
        extent
            .base_hidden_states
            .push(tst::hidden_state_extent_archive::RowOrColumnState {
                row_or_column_uid: *uid,
                user_hidden: Some(true),
                ..Default::default()
            });
    }
    extent.base_hidden_states.sort_by_key(|state| {
        indexes
            .get(&(state.row_or_column_uid.lower, state.row_or_column_uid.upper))
            .copied()
            .unwrap_or(usize::MAX)
    });
    Ok(())
}

fn empty_extent(
    uid: tsp::Uuid,
    direction: tst::hidden_state_extent_archive::RowOrColumnDirection,
    filter_set_id: u64,
) -> tst::HiddenStateExtentArchive {
    tst::HiddenStateExtentArchive {
        hidden_state_extent_uid: uid,
        row_or_column_direction: direction as i32,
        needs_to_update_filter_set_for_import: Some(false),
        filter_set: Some(tsp::Reference {
            identifier: filter_set_id,
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn insert_hidden_state_objects(
    package: &mut IWorkPackage,
    graph: &TableHiddenGraph,
    owner: &tst::HiddenStatesOwnerArchive,
    identifiers: HiddenStateObjectIds,
) -> Result<()> {
    let active = owner.hidden_states.first().ok_or_else(|| {
        Error::InvalidFormat("iWork table hidden-state owner has no state".to_owned())
    })?;
    package.update_archive(&graph.model_archive, |archive| {
        for identifier in [identifiers.filter_columns, identifiers.filter_rows] {
            let data = tst::FilterSetArchive {
                r#type: Some(
                    tst::filter_set_archive::FilterSetType::FilterSetArchiveTypeAll as i32,
                ),
                is_enabled: Some(false),
                needs_formula_rewrite_for_import: Some(false),
                filter_offsets: vec![0],
                ..Default::default()
            }
            .encode_to_vec();
            let mut object = ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: FILTER_SET_MESSAGE_TYPE,
                    data,
                }],
            )?;
            object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
            archive.insert_object(object)?;
        }
        for (identifier, uid) in [
            (
                identifiers.formula_columns,
                &active.column_hidden_state_extent.hidden_state_extent_uid,
            ),
            (
                identifiers.formula_rows,
                &active.row_hidden_state_extent.hidden_state_extent_uid,
            ),
        ] {
            let data = tst::HiddenStateFormulaOwnerArchive {
                owner_id: Some(uuid_as_cfuuid(uid)),
                needs_to_update_filter_set_for_import: Some(false),
                ..Default::default()
            }
            .encode_to_vec();
            let mut object = ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                    data,
                }],
            )?;
            object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
            archive.insert_object(object)?;
        }
        Ok(())
    })?;
    set_package_last_object_identifier(package, identifiers.formula_rows)
}

fn patch_model(
    package: &mut IWorkPackage,
    graph: &TableHiddenGraph,
    owner: &tst::HiddenStatesOwnerArchive,
    active_uuid: &tsp::Uuid,
    new_object_ids: Option<HiddenStateObjectIds>,
) -> Result<()> {
    package.update_archive(&graph.model_archive, |archive| {
        let object = archive.object_mut(graph.model_object_id).ok_or_else(|| {
            Error::InvalidFormat("iWork table model moved during update".to_owned())
        })?;
        let original = &object.messages[graph.model_message_index].data;
        validate_singular_field(
            original,
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
            graph.model.hidden_states_owner.is_some(),
            "hidden-state owner",
        )?;
        let mut data = patch_length_delimited_field(
            original,
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
            graph.model.hidden_states_owner.is_some(),
            Some(&owner.encode_to_vec()),
        )?;
        for (field, current, new_identifier) in [
            (
                TABLE_MODEL_COLUMN_HIDDEN_FORMULA_OWNER_FIELD,
                graph.model.hidden_state_formula_owner_for_columns.as_ref(),
                new_object_ids.map(|ids| ids.formula_columns),
            ),
            (
                TABLE_MODEL_ROW_HIDDEN_FORMULA_OWNER_FIELD,
                graph.model.hidden_state_formula_owner_for_rows.as_ref(),
                new_object_ids.map(|ids| ids.formula_rows),
            ),
        ] {
            validate_singular_field(
                original,
                field,
                current.is_some(),
                "hidden-state formula owner",
            )?;
            let replacement = new_identifier
                .or_else(|| current.map(|value| value.identifier))
                .map(reference)
                .map(|value| value.encode_to_vec());
            data = patch_length_delimited_field(
                &data,
                field,
                current.is_some(),
                replacement.as_deref(),
            )?;
        }
        let active = unique_active_state(owner, active_uuid)?;
        let (hidden_rows, user_hidden_rows) =
            extent_hidden_counts(&active.row_hidden_state_extent, "row")?;
        let (hidden_columns, user_hidden_columns) =
            extent_hidden_counts(&active.column_hidden_state_extent, "column")?;
        for (field, current, replacement) in [
            (
                TABLE_MODEL_HIDDEN_ROWS_FIELD,
                graph.model.number_of_hidden_rows,
                graph.model.number_of_hidden_rows.map(|_| hidden_rows),
            ),
            (
                TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
                graph.model.number_of_hidden_columns,
                graph.model.number_of_hidden_columns.map(|_| hidden_columns),
            ),
            (
                TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
                graph.model.number_of_user_hidden_rows,
                graph
                    .model
                    .number_of_user_hidden_rows
                    .map(|_| user_hidden_rows),
            ),
            (
                TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
                graph.model.number_of_user_hidden_columns,
                graph
                    .model
                    .number_of_user_hidden_columns
                    .map(|_| user_hidden_columns),
            ),
        ] {
            data = patch_varint_field(&data, field, current.is_some(), replacement.map(u64::from))?;
        }
        object.replace_message(
            graph.model_message_index,
            RawMessage {
                type_: graph.model_message_type,
                data,
            },
        )?;
        let references =
            &mut object.archive_info.message_infos[graph.model_message_index].object_references;
        for identifier in owner.hidden_states.iter().flat_map(|state| {
            [
                state.column_hidden_state_extent.filter_set.as_ref(),
                state.row_hidden_state_extent.filter_set.as_ref(),
            ]
            .into_iter()
            .flatten()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
        }) {
            if !references.contains(&identifier) {
                references.push(identifier);
            }
        }
        if let Some(identifiers) = new_object_ids {
            for identifier in [identifiers.formula_columns, identifiers.formula_rows] {
                if !references.contains(&identifier) {
                    references.push(identifier);
                }
            }
        }
        Ok(())
    })
}

fn extent_hidden_counts(extent: &tst::HiddenStateExtentArchive, axis: &str) -> Result<(u32, u32)> {
    let total = extent
        .base_hidden_states
        .iter()
        .filter(|state| {
            state.user_hidden == Some(true)
                || state.filtered == Some(true)
                || state.pivot_hidden == Some(true)
        })
        .count();
    let user = extent
        .base_hidden_states
        .iter()
        .filter(|state| state.user_hidden == Some(true))
        .count();
    Ok((
        u32::try_from(total)
            .map_err(|_| Error::ParseError(format!("hidden iWork {axis} count exceeds u32")))?,
        u32::try_from(user).map_err(|_| {
            Error::ParseError(format!("user-hidden iWork {axis} count exceeds u32"))
        })?,
    ))
}

fn patch_info(
    package: &mut IWorkPackage,
    graph: &TableHiddenGraph,
    active_uuid: &tsp::Uuid,
) -> Result<()> {
    package.update_archive(&graph.info_archive, |archive| {
        let object = archive.object_mut(graph.info_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table info {} moved during update",
                graph.info_object_id
            ))
        })?;
        let original = &object.messages[graph.info_message_index].data;
        validate_singular_field(
            original,
            TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
            graph.info.hidden_states_uuid.is_some(),
            "active hidden-state UUID",
        )?;
        let data = patch_length_delimited_field(
            original,
            TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
            graph.info.hidden_states_uuid.is_some(),
            Some(&active_uuid.encode_to_vec()),
        )?;
        object.replace_message(
            graph.info_message_index,
            RawMessage {
                type_: graph.info_message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn uuid_as_cfuuid(uuid: &tsp::Uuid) -> tsp::CfuuidArchive {
    tsp::CfuuidArchive {
        uuid_bytes: None,
        uuid_w0: Some(uuid.lower as u32),
        uuid_w1: Some((uuid.lower >> 32) as u32),
        uuid_w2: Some(uuid.upper as u32),
        uuid_w3: Some((uuid.upper >> 32) as u32),
    }
}

fn validate_singular_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    label: &str,
) -> Result<()> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != usize::from(expected_present)
        || matches.iter().any(|field| field.wire_type != 2)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} is missing, duplicated, or malformed"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hidden_axes_are_canonical_and_reject_duplicates() {
        let hidden = TableHiddenAxes::new([
            TableAxisIndex::column(2),
            TableAxisIndex::row(3),
            TableAxisIndex::row(1),
        ])
        .unwrap();

        assert_eq!(
            hidden.as_slice(),
            [
                TableAxisIndex::row(1),
                TableAxisIndex::row(3),
                TableAxisIndex::column(2),
            ]
        );
        assert!(TableHiddenAxes::new([TableAxisIndex::row(1), TableAxisIndex::row(1)]).is_err());
    }
}
