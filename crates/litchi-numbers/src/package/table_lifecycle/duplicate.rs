//! A populated copy of a table on its own sheet (record 0370h).
//!
//! The port of the host editor's `duplicate_table`
//! (`duplicate_attached_table_graph_in_package` and
//! `clone_table_formula_graph`), on the structural working set. The copy
//! takes:
//!
//! - fresh identifiers for its drawable, its model and every storage object
//!   the model owns, including the data-list segments those lists chain to;
//! - a table UUID derived from the new model's identifier, as the host does;
//! - a clone of its CalculationEngine owner family, with its dependency
//!   tiles, fresh internal owner identifiers and owner UUIDs translated from
//!   the new table UUID;
//! - its formula lists, with their host-table and self-table references
//!   rewritten to the new UUIDs.
//!
//! References to anything outside the table (styles, other tables' owners,
//! the parent sheet) are kept. Every field the clone does not rewrite is
//! copied verbatim, unknown fields included. The copy goes at the end of
//! the sheet's drawables, offset by ten points, with a sidebar node under
//! its sheet's node. Owner state that cannot be remapped without ambiguity
//! (volatile, spanning, whole-owner, error or spill dependency state) is
//! refused, as the host refuses it.

use std::collections::{BTreeMap, HashMap, HashSet};

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    append_length_delimited_field, patch_length_delimited_field, patch_nested_fixed32_field,
    patch_varint_field, repeated_fixed32_values, repeated_length_delimited_payloads,
    repeated_varint_values, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_numbers_wire::formula_graph;
use litchi_numbers_wire::formula_owner::Uuid as WireUuid;

use super::super::Package;
use super::super::structure::{
    self, DOCUMENT_SIDEBAR_FIELD, Fault, FaultResult, Published, TREE_CHILDREN_FIELD,
    TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    metadata::{self, Transition},
    references,
};
use super::{Error, locate, rewrite_list};

/// `TST.TableModelArchive` (and its pre-BNC form).
const TABLE_MODEL_MESSAGE_TYPES: [u32; 2] = [6_000, 6_001];
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `TSCE.CellRecordTileArchive`.
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;
/// `TSCE.RangePrecedentsTileArchive`.
const RANGE_PRECEDENTS_TILE_MESSAGE_TYPE: u32 = 4_010;
/// How far the copy is offset from its source (the host editor's
/// `DRAWABLE_DUPLICATE_OFFSET`).
const OFFSET: f32 = 10.0;

/// The model's owned-object reference paths.
const MODEL_REFERENCE_PATHS: [&[u32]; 19] = [
    &[4, 1, 2],
    &[4, 2],
    &[4, 3, 1, 2],
    &[4, 4],
    &[4, 5],
    &[4, 6],
    &[4, 11],
    &[4, 12],
    &[4, 13],
    &[4, 15],
    &[4, 16],
    &[4, 17],
    &[4, 18],
    &[4, 19],
    &[4, 20],
    &[4, 21],
    &[4, 22],
    &[46],
    &[49],
];
/// The table drawable's reference paths.
const INFO_REFERENCE_PATHS: [&[u32]; 10] = [
    &[1, 2],
    &[1, 6],
    &[1, 9],
    &[2],
    &[3],
    &[4],
    &[5],
    &[6],
    &[15],
    &[17],
];

/// A [`Fault`] carried through the wire and decode callbacks.
struct E(Fault);

impl From<Fault> for E {
    fn from(fault: Fault) -> Self {
        Self(fault)
    }
}

impl From<litchi_iwa_common::Error> for E {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(structure::map_wire(&error))
    }
}

impl From<formula_graph::Error> for E {
    fn from(error: formula_graph::Error) -> Self {
        Self(match error {
            formula_graph::Error::UnsupportedDependency => Fault::Dependency,
            formula_graph::Error::Allocation { requested } => {
                Fault::Allocation { amount: requested }
            },
            _ => Fault::Invalid,
        })
    }
}

type R<T> = Result<T, E>;

const fn invalid<T>() -> R<T> {
    Err(E(Fault::Invalid))
}

/// A `TSP.UUID`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct Uuid {
    lower: u64,
    upper: u64,
}

/// The last value of varint field `number`.
fn varint(data: &[u8], number: u32) -> R<Option<u64>> {
    Ok(repeated_varint_values(data, number)?.last().copied())
}

/// The last payload of length-delimited field `number`.
fn bytes(data: &[u8], number: u32) -> R<Option<&[u8]>> {
    Ok(repeated_length_delimited_payloads(data, number)?
        .last()
        .copied())
}

/// Every payload at `path`, through repeated messages at every level.
fn payloads_at<'d>(data: &'d [u8], path: &[u32]) -> R<Vec<&'d [u8]>> {
    let Some((first, rest)) = path.split_first() else {
        return invalid();
    };
    let payloads = repeated_length_delimited_payloads(data, *first)?;
    if rest.is_empty() {
        return Ok(payloads);
    }
    let mut output = Vec::new();
    for payload in payloads {
        output.extend(payloads_at(payload, rest)?);
    }
    Ok(output)
}

/// The identifiers of every `TSP.Reference` at `path`.
fn references_at(data: &[u8], path: &[u32]) -> R<Vec<u64>> {
    let mut output = Vec::new();
    for reference in payloads_at(data, path)? {
        output.push(varint(reference, 1)?.unwrap_or(0));
    }
    Ok(output)
}

/// Whether `data` holds field `number`.
fn has(data: &[u8], number: u32) -> R<bool> {
    let view = litchi_iwa_common::wire::parse_wire_view(data, WireLimits::DEFAULT)?;
    Ok(view.fields().any(|field| field.number() == number))
}

fn read_uuid(data: &[u8]) -> R<Uuid> {
    Ok(Uuid {
        lower: varint(data, 1)?.unwrap_or(0),
        upper: varint(data, 2)?.unwrap_or(0),
    })
}

/// The copy's name: `"<name> copy"`, then `"<name> copy 2"`, … until no
/// table of the sheet has it (the host editor's `duplicate_table_name`).
pub(super) fn copy_name(source: &str, existing: &[&str]) -> Option<String> {
    let base = format!("{source} copy");
    if !existing.contains(&base.as_str()) {
        return Some(base);
    }
    (2_u32..=u32::MAX)
        .map(|suffix| format!("{base} {suffix}"))
        .find(|candidate| !existing.contains(&candidate.as_str()))
}

/// Duplicate table `position` of sheet `sheet_position` as `name`.
pub(super) fn duplicate_table(
    source: &Package,
    sheet_position: usize,
    position: usize,
    name: &str,
) -> Result<Published, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let located = locate(&working, sheet_position, position)?;
    let mut transition = Transition::default();
    clone_table(
        &mut working,
        &mut transition,
        located.table,
        Placement {
            sheet: located.sheet,
            source_sheet: located.sheet,
            offset: OFFSET,
        },
        name,
    )?;
    Ok(working.publish(&transition)?)
}

/// Where a copy goes: the sheet whose drawables it joins (and under whose
/// sidebar node it is listed), the source table's sheet, and how far it is
/// offset from its source.
#[derive(Debug, Clone, Copy)]
pub(in crate::package) struct Placement {
    pub(in crate::package) sheet: u64,
    pub(in crate::package) source_sheet: u64,
    pub(in crate::package) offset: f32,
}

/// Copy the table drawable `info` as `name` at `placement`, recording its
/// registry consequences in `transition`; returns the copy's drawable.
pub(in crate::package) fn clone_table(
    working: &mut Working<'_>,
    transition: &mut Transition,
    info: u64,
    placement: Placement,
    name: &str,
) -> FaultResult<u64> {
    clone(working, transition, info, placement, name).map_err(|error| error.0)
}

/// One cloned object and the component it goes to.
struct Clone {
    component: String,
    object: ArchiveObject,
    source: u64,
}

/// The list-reference paths of a model's data store (the data lists).
const LIST_PATHS: [&[u32]; 14] = [
    &[4, 4],
    &[4, 5],
    &[4, 6],
    &[4, 11],
    &[4, 12],
    &[4, 13],
    &[4, 15],
    &[4, 16],
    &[4, 17],
    &[4, 18],
    &[4, 19],
    &[4, 20],
    &[4, 21],
    &[4, 22],
];
/// The other owned-object paths: header buckets, tiles, the UID map and
/// the stroke sidecar.
const STORAGE_PATHS: [&[u32]; 5] = [&[4, 1, 2], &[4, 2], &[4, 3, 1, 2], &[46], &[49]];
/// `TST.TableDataList.ListType.FORMULA`.
const FORMULA_LIST_TYPE: u64 = 3;

#[allow(
    clippy::too_many_lines,
    reason = "the port keeps the host editor's single clone sequence"
)]
fn clone(
    working: &mut Working<'_>,
    transition: &mut Transition,
    info_id: u64,
    placement: Placement,
    name: &str,
) -> R<u64> {
    let limits = working.archive_limits();
    let info_component = working.locate(info_id)?.ok_or(Fault::Invalid)?;
    let info_object = working.object(info_id)?.clone();
    let info_index = structure::table_info_message(&info_object)?;
    let info_data = info_object.messages[info_index].data.clone();
    let [model_id] = references_at(&info_data, &[2])?[..] else {
        return invalid();
    };
    let model_component = working.locate(model_id)?.ok_or(Fault::Invalid)?;
    let model_object = working.object(model_id)?.clone();
    let model_index = structure::unique_message(&model_object, &TABLE_MODEL_MESSAGE_TYPES)?;
    let model_data = model_object.messages[model_index].data.clone();
    let source_table_id = bytes(&model_data, 1)?.ok_or(Fault::Invalid)?;
    let source_table_id = core::str::from_utf8(source_table_id)
        .map_err(|_error| E(Fault::Invalid))?
        .to_owned();

    let graph = owned_graph(working, &model_data)?;
    if graph.contains_key(&info_id) || graph.contains_key(&model_id) {
        return invalid();
    }

    // Fresh identifiers: the drawable, the model, then the storage in
    // source-identifier order.
    let fresh = working.allocate(graph.len() + 2)?;
    let (new_info, new_model) = (fresh[0], fresh[1]);
    let mut remap: HashMap<u64, u64> = HashMap::new();
    remap.insert(info_id, new_info);
    remap.insert(model_id, new_model);
    if placement.sheet != placement.source_sheet {
        // The copy's drawable is parented to its new sheet.
        remap.insert(placement.source_sheet, placement.sheet);
    }
    for (source, replacement) in graph.keys().zip(&fresh[2..]) {
        remap.insert(*source, *replacement);
    }
    let table_uuid = allocate_table_uuid(new_model, &table_uuids(working)?);
    let source_uuid = parse_table_uuid(&source_table_id)?;
    let new_uuid = parse_table_uuid(&table_uuid)?;

    // The CalculationEngine owner family, when the workbook has an engine.
    let engine = engine(working)?;
    let family = match &engine {
        Some((component, _, _)) => Some(owner_family(working, component, info_id)?),
        None => None,
    };
    let owner_uuids = match &family {
        Some(family) => {
            if family.direct_uuid != owner_uuid_for_table(&source_uuid) {
                return invalid();
            }
            Some((family.direct_uuid, owner_uuid_for_table(&new_uuid)))
        },
        None => None,
    };

    // The model, the drawable and the storage.
    let new_model_data = {
        let data = patch_length_delimited_field(&model_data, 1, true, Some(table_uuid.as_bytes()))?;
        let data = patch_length_delimited_field(&data, 8, true, Some(name.as_bytes()))?;
        remap_paths(&data, &MODEL_REFERENCE_PATHS, &remap)?
    };
    let mut clones = vec![Clone {
        component: model_component,
        object: clone_metadata(
            &model_object,
            new_model,
            vec![RawMessage {
                type_: model_object.messages[model_index].type_,
                data: new_model_data,
            }],
            &remap,
        )?,
        source: model_id,
    }];
    let new_info_data = {
        let mut data = remap_paths(&info_data, &INFO_REFERENCE_PATHS, &remap)?;
        if let Some(position) = payloads_at(&info_data, &[1, 1, 1])?.last() {
            let x = repeated_fixed32_values(position, 1)?
                .last()
                .copied()
                .unwrap_or(0);
            let y = repeated_fixed32_values(position, 2)?
                .last()
                .copied()
                .unwrap_or(0);
            let (x, y) = (
                f32::from_bits(x) + placement.offset,
                f32::from_bits(y) + placement.offset,
            );
            if !x.is_finite() || !y.is_finite() {
                return invalid();
            }
            data = patch_nested_fixed32_field(&data, &[1, 1, 1, 1], true, Some(x.to_bits()))?;
            data = patch_nested_fixed32_field(&data, &[1, 1, 1, 2], true, Some(y.to_bits()))?;
        }
        data
    };
    clones.push(Clone {
        component: info_component.clone(),
        object: clone_metadata(
            &info_object,
            new_info,
            vec![RawMessage {
                type_: info_object.messages[info_index].type_,
                data: new_info_data,
            }],
            &remap,
        )?,
        source: info_id,
    });
    for (source, list) in &graph {
        let component = working.locate(*source)?.ok_or(Fault::Invalid)?;
        let object = working.object(*source)?;
        let mut messages = Vec::new();
        for (message, header) in object
            .messages
            .iter()
            .zip(&object.archive_info.message_infos)
        {
            let private: HashSet<u64> = header
                .object_references
                .iter()
                .copied()
                .filter(|identifier| remap.contains_key(identifier))
                .collect();
            let mut data = if private.is_empty() {
                message.data.clone()
            } else {
                remap_segments(&message.data, &private, &remap)?
            };
            if *list && varint(&data, 1)? == Some(FORMULA_LIST_TYPE) {
                data = remap_formula_list(&data, &source_uuid, &new_uuid)?;
                if let Some((source_owner, new_owner)) = &owner_uuids {
                    data = remap_formula_list(&data, source_owner, new_owner)?;
                }
            }
            messages.push(RawMessage {
                type_: message.type_,
                data,
            });
        }
        clones.push(Clone {
            component,
            object: clone_metadata(object, remap[source], messages, &remap)?,
            source: *source,
        });
    }
    for clone in &clones {
        working.insert(&clone.component, clone.object.clone())?;
    }

    // The CalculationEngine owner family.
    let mut engine_clones = Vec::new();
    if let (Some((component, engine_id, engine_index)), Some(family)) = (&engine, &family) {
        engine_clones = clone_owner_family(
            working,
            component,
            (*engine_id, *engine_index),
            family,
            (info_id, new_info),
            (&source_uuid, &new_uuid),
            limits,
        )?;
    }

    // Attach to the sheet's drawables.
    let sheet_object = working.object(placement.sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)?;
    let mut after = drawables.clone();
    after.push(new_info);
    rewrite_list(
        working,
        placement.sheet,
        sheet_index,
        sheet_path,
        &drawables,
        &after,
    )?;

    // The sheet's sidebar node lists its tables.
    let node = append_sidebar_node(working, placement.sheet, new_info)?;

    // Registry: UUIDs for the clones of registered objects, and every new
    // cross-component reference.
    let Some(registry) = metadata::registry(working)? else {
        return Ok(new_info);
    };
    let seed = metadata::identity_seed(working, new_info)?;
    let mut inserted: Vec<(u64, String, Option<u64>)> = Vec::new();
    for clone in &clones {
        let identifier = clone.object.archive_info.identifier.ok_or(Fault::Invalid)?;
        inserted.push((identifier, clone.component.clone(), Some(clone.source)));
    }
    for (identifier, source, component) in &engine_clones {
        inserted.push((*identifier, component.clone(), Some(*source)));
    }
    if let Some((node, component)) = &node {
        inserted.push((*node, component.clone(), None));
    }
    inserted.sort_by_key(|(identifier, _, _)| *identifier);
    for (identifier, component, source) in &inserted {
        let component_id = registry.component_for_entry(component)?.identifier;
        if let Some(source) = source
            && registry
                .uuids
                .iter()
                .any(|binding| binding.object == *source)
        {
            transition.register_uuid(&registry, component_id, *identifier, seed)?;
        }
        let referenced = working.object(*identifier)?.archive_info.message_infos[0]
            .object_references
            .clone();
        for target in referenced {
            if placement.sheet != placement.source_sheet && target == placement.sheet {
                // Registered last, below, where the host editor records it.
                continue;
            }
            let target_component = working.locate(target)?.ok_or(Fault::Invalid)?;
            let target_id = registry.component_for_entry(&target_component)?.identifier;
            transition.reference(&registry, component_id, target_id, target)?;
        }
    }
    let info_component_id = registry.component_for_entry(&info_component)?.identifier;
    let sheet_component = working.locate(placement.sheet)?.ok_or(Fault::Invalid)?;
    let sheet_component = registry.component_for_entry(&sheet_component)?.identifier;
    transition.reference(&registry, sheet_component, info_component_id, new_info)?;
    if let Some((component, _, _)) = &engine {
        let engine_component = registry.component_for_entry(component)?.identifier;
        transition.reference(&registry, engine_component, info_component_id, new_info)?;
    }
    if placement.sheet != placement.source_sheet {
        // A drawable outside the document component (a legacy table info)
        // addresses its new parent sheet across components.
        transition.reference(
            &registry,
            info_component_id,
            sheet_component,
            placement.sheet,
        )?;
    }
    Ok(new_info)
}

/// The storage objects a model owns: its tiles, header buckets, lists and
/// sidecars, and every data-list segment those lists chain to (the host
/// editor's `table_owned_graph`). The value says whether the object is a
/// data list.
fn owned_graph(working: &Working<'_>, model: &[u8]) -> R<BTreeMap<u64, bool>> {
    let mut graph: BTreeMap<u64, bool> = BTreeMap::new();
    let insert = |graph: &mut BTreeMap<u64, bool>, identifier: u64, list: bool| -> R<bool> {
        if identifier == 0 {
            return Ok(false);
        }
        match graph.insert(identifier, list) {
            Some(previous) if previous != list => invalid(),
            Some(_) => Ok(false),
            None => Ok(true),
        }
    };
    for path in STORAGE_PATHS {
        for identifier in references_at(model, path)? {
            insert(&mut graph, identifier, false)?;
        }
    }
    let mut pending = Vec::new();
    for path in LIST_PATHS {
        for identifier in references_at(model, path)? {
            if insert(&mut graph, identifier, true)? {
                pending.push(identifier);
            }
        }
    }
    let mut visited = HashSet::new();
    while let Some(identifier) = pending.pop() {
        if !visited.insert(identifier) {
            continue;
        }
        for message in &working.object(identifier)?.messages {
            if !varint(&message.data, 1)?.is_some_and(|kind| (1..=12).contains(&kind)) {
                continue;
            }
            for segment in references_at(&message.data, &[4])? {
                if insert(&mut graph, segment, true)? {
                    pending.push(segment);
                }
            }
        }
    }
    for identifier in graph.keys() {
        if working.locate(*identifier)?.is_none() {
            return invalid();
        }
    }
    Ok(graph)
}

/// Every table UUID of the workbook.
fn table_uuids(working: &Working<'_>) -> R<HashSet<String>> {
    let mut output = HashSet::new();
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            for message in &object.messages {
                if TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_)
                    && let Some(text) = bytes(&message.data, 1)?
                    && let Ok(text) = core::str::from_utf8(text)
                {
                    output.insert(text.to_owned());
                }
            }
        }
    }
    Ok(output)
}

/// The host editor's deterministic table UUID for a new model.
fn allocate_table_uuid(seed: u64, existing: &HashSet<String>) -> String {
    let mut suffix = seed & 0x0000_ffff_ffff_ffff;
    loop {
        let candidate = format!("00000000-0000-4000-8000-{suffix:012X}");
        if !existing.contains(&candidate) {
            return candidate;
        }
        suffix = suffix.wrapping_add(1) & 0x0000_ffff_ffff_ffff;
    }
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "splitting a 128-bit UUID into its halves"
)]
fn parse_table_uuid(value: &str) -> R<Uuid> {
    let compact: String = value
        .chars()
        .filter(|character| *character != '-')
        .collect();
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return invalid();
    }
    let raw = u128::from_str_radix(&compact, 16).map_err(|_error| E(Fault::Invalid))?;
    Ok(Uuid {
        lower: raw as u64,
        upper: (raw >> 64) as u64,
    })
}

/// A table's formula-owner UUID: its UUID with each half byte-reversed and
/// the halves swapped, as native Numbers derives it.
const fn owner_uuid_for_table(table: &Uuid) -> Uuid {
    Uuid {
        lower: table.upper.swap_bytes(),
        upper: table.lower.swap_bytes(),
    }
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "splitting a 128-bit UUID into its halves"
)]
const fn translate_owner_uuid(owner: &Uuid, source: &Uuid, replacement: &Uuid) -> Uuid {
    let owner = ((owner.upper as u128) << 64) | owner.lower as u128;
    let source = ((source.upper as u128) << 64) | source.lower as u128;
    let replacement = ((replacement.upper as u128) << 64) | replacement.lower as u128;
    let translated = replacement.wrapping_add(owner.wrapping_sub(source));
    Uuid {
        lower: translated as u64,
        upper: (translated >> 64) as u64,
    }
}

const fn uuid_key(uuid: &Uuid) -> (u64, u64) {
    (uuid.lower, uuid.upper)
}

fn uuid_bytes(uuid: &Uuid) -> [u8; 16] {
    ((u128::from(uuid.upper) << 64) | u128::from(uuid.lower)).to_be_bytes()
}

/// The `(lower, upper)` key of a `TSP.CFUUIDArchive` payload.
#[allow(
    clippy::cast_possible_truncation,
    reason = "splitting a 128-bit UUID into its halves"
)]
fn cfuuid_key(data: &[u8]) -> R<Option<(u64, u64)>> {
    let words = [
        varint(data, 2)?,
        varint(data, 3)?,
        varint(data, 4)?,
        varint(data, 5)?,
    ];
    if let [Some(w0), Some(w1), Some(w2), Some(w3)] = words {
        return Ok(Some((w0 | (w1 << 32), w2 | (w3 << 32))));
    }
    let Some(raw) = bytes(data, 1)? else {
        return Ok(None);
    };
    let Ok(raw) = <[u8; 16]>::try_from(raw) else {
        return Ok(None);
    };
    let value = u128::from_be_bytes(raw);
    Ok(Some((value as u64, (value >> 64) as u64)))
}

/// Replace every `TSP.Reference` at each of `paths` whose identifier is in
/// `remap`.
fn remap_paths(data: &[u8], paths: &[&[u32]], remap: &HashMap<u64, u64>) -> R<Vec<u8>> {
    let mut data = data.to_vec();
    for path in paths {
        data = transform_length_delimited_fields_at_path(&data, path, |reference| {
            let identifier = varint(reference, 1)?.unwrap_or(0);
            Ok::<_, E>(match remap.get(&identifier) {
                Some(replacement) => patch_varint_field(reference, 1, true, Some(*replacement))?,
                None => reference.to_vec(),
            })
        })?;
    }
    Ok(data)
}

/// A copy of `source`'s header for `identifier` carrying `messages`: the
/// merge flag and each message's header, with every reference in `remap`
/// replaced (the host editor's `clone_numbers_object_metadata`).
fn clone_metadata(
    source: &ArchiveObject,
    identifier: u64,
    messages: Vec<RawMessage>,
    remap: &HashMap<u64, u64>,
) -> R<ArchiveObject> {
    let mut cloned = ArchiveObject::new(identifier, messages, litchi_iwa_core::Limits::default())
        .map_err(|error| E(structure::map_core(&error)))?;
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    for ((target, header), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&source.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_error| E(Fault::Invalid))?;
        *target = header.clone();
        target.length = length;
        for reference in &mut target.object_references {
            if let Some(replacement) = remap.get(reference) {
                *reference = *replacement;
            }
        }
        for field in &mut target.field_infos {
            for reference in &mut field.object_references {
                if let Some(replacement) = remap.get(reference) {
                    *reference = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

/// A data list whose private references are its segments, with them
/// remapped.
fn remap_segments(data: &[u8], private: &HashSet<u64>, remap: &HashMap<u64, u64>) -> R<Vec<u8>> {
    let segments: HashSet<u64> = references_at(data, &[4])?.into_iter().collect();
    if !private.is_subset(&segments) {
        return Err(E(Fault::Unsupported));
    }
    remap_paths(data, &[&[4]], remap)
}

/// A formula list with every formula's host-table UUID and AST table
/// references equal to `source` replaced by `replacement`.
fn remap_formula_list(data: &[u8], source: &Uuid, replacement: &Uuid) -> R<Vec<u8>> {
    let remap = HashMap::from([(uuid_key(source), *replacement)]);
    transform_length_delimited_fields_at_path(data, &[3, 5], |formula| {
        let mut data = formula.to_vec();
        if let Some(host) = bytes(formula, 7)?
            && read_uuid(host)? == *source
        {
            data = transform_length_delimited_fields_at_path(&data, &[7], |uuid| {
                patch_uuid(uuid, replacement)
            })?;
        }
        transform_length_delimited_field(&data, 1, |ast| remap_ast(ast, &remap))
    })
}

fn remap_ast(data: &[u8], remap: &HashMap<(u64, u64), Uuid>) -> R<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, &[1], |node| {
        let node = transform_length_delimited_fields_at_path(node, &[16, 5], |uuid| {
            remap_cfuuid(uuid, remap)
        })?;
        let node = transform_length_delimited_fields_at_path(&node, &[28, 1], |uuid| {
            remap_cfuuid(uuid, remap)
        })?;
        transform_length_delimited_fields_at_path(&node, &[14], |nested| remap_ast(nested, remap))
    })
}

fn patch_uuid(data: &[u8], uuid: &Uuid) -> R<Vec<u8>> {
    let data = patch_varint_field(data, 1, true, Some(uuid.lower))?;
    Ok(patch_varint_field(&data, 2, true, Some(uuid.upper))?)
}

/// A `TSP.CFUUIDArchive` payload remapped through `remap`, in the forms it
/// already uses (bytes, words or both).
fn remap_cfuuid(data: &[u8], remap: &HashMap<(u64, u64), Uuid>) -> R<Vec<u8>> {
    let Some(replacement) = cfuuid_key(data)?.and_then(|key| remap.get(&key)) else {
        return Ok(data.to_vec());
    };
    let mut output = data.to_vec();
    if has(data, 1)? {
        output = patch_length_delimited_field(&output, 1, true, Some(&uuid_bytes(replacement)))?;
    }
    for (field, value) in [
        (2, replacement.lower & 0xffff_ffff),
        (3, replacement.lower >> 32),
        (4, replacement.upper & 0xffff_ffff),
        (5, replacement.upper >> 32),
    ] {
        if has(data, field)? {
            output = patch_varint_field(&output, field, true, Some(value))?;
        }
    }
    Ok(output)
}

/// The CalculationEngine root: its component, object and message.
fn engine(working: &Working<'_>) -> R<Option<(String, u64, usize)>> {
    let mut found = None;
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                    if found.replace((name.clone(), identifier, index)).is_some() {
                        return invalid();
                    }
                }
            }
        }
    }
    Ok(found)
}

/// One object of the owner family and the index of its payload.
struct Member {
    object: ArchiveObject,
    index: usize,
}

impl Member {
    fn data(&self) -> &[u8] {
        &self.object.messages[self.index].data
    }
}

/// The fields of one formula owner the clone reads.
struct Owner {
    uid: Uuid,
    internal: u32,
    cell_tiles: Vec<u64>,
    /// The cells its inline records hold.
    cells: Vec<(u64, u64)>,
}

/// A table's CalculationEngine owner family.
struct Family {
    owners: Vec<(Member, Owner)>,
    cell_tiles: Vec<Member>,
    range_tiles: Vec<Member>,
    direct_uuid: Uuid,
}

/// The single message of `object` of type `message_type`, if any.
fn single_message(object: &ArchiveObject, message_type: u32) -> R<Option<usize>> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ == message_type && found.replace(index).is_some() {
            return invalid();
        }
    }
    Ok(found)
}

fn internal_id(value: Option<u64>) -> R<u32> {
    u32::try_from(value.unwrap_or(0)).map_err(|_error| E(Fault::Invalid))
}

/// The `(row, column)` of every cell record at `path`.
fn cell_records(data: &[u8], path: &[u32]) -> R<Vec<(u64, u64)>> {
    let mut output = Vec::new();
    for record in payloads_at(data, path)? {
        output.push((
            varint(record, 2)?.unwrap_or(0),
            varint(record, 1)?.unwrap_or(0),
        ));
    }
    Ok(output)
}

/// The owners whose `formula_owner` is the table drawable or whose base
/// owner is its direct owner, with their dependency tiles (the host
/// editor's `formula_owner_family`, `cell_record_tiles` and
/// `range_precedent_tiles`).
fn owner_family(working: &Working<'_>, component: &str, info: u64) -> R<Family> {
    let mut all = Vec::new();
    let mut direct = None;
    for object in &working.archive(component)?.objects {
        let Some(index) = single_message(object, FORMULA_OWNER_MESSAGE_TYPE)? else {
            continue;
        };
        let data = &object.messages[index].data;
        let uid = read_uuid(bytes(data, 1)?.ok_or(Fault::Invalid)?)?;
        let internal = internal_id(varint(data, 2)?)?;
        let owned_by = references_at(data, &[11])?.first().copied();
        let base = match bytes(data, 12)? {
            Some(base) => Some(read_uuid(base)?),
            None => None,
        };
        if owned_by == Some(info) && direct.replace(uid).is_some() {
            return invalid();
        }
        all.push((object, index, uid, internal, owned_by, base));
    }
    // A table without its own owner is not a host-editable table.
    let direct_uuid = direct.ok_or(Fault::Unsupported)?;
    let mut owners = Vec::new();
    for (object, index, uid, internal, owned_by, base) in all {
        if owned_by != Some(info) && base != Some(direct_uuid) {
            continue;
        }
        let data = &object.messages[index].data;
        owners.push((
            Member {
                object: object.clone(),
                index,
            },
            Owner {
                uid,
                internal,
                cell_tiles: references_at(data, &[13, 1])?,
                cells: cell_records(data, &[4, 1])?,
            },
        ));
        // The range tiles are read below with the cell tiles.
    }
    owners.sort_by_key(|(member, _)| member.object.archive_info.identifier);
    let mut cell_ids = BTreeMap::new();
    let mut range_ids = BTreeMap::new();
    for (member, owner) in &owners {
        for identifier in &owner.cell_tiles {
            if cell_ids.insert(*identifier, ()).is_some() {
                return invalid();
            }
        }
        for identifier in references_at(member.data(), &[15, 1])? {
            if range_ids.insert(identifier, ()).is_some() {
                return invalid();
            }
        }
    }
    let archive = working.archive(component)?;
    let tiles = |identifiers: &BTreeMap<u64, ()>, message_type: u32| -> R<Vec<Member>> {
        let mut output = Vec::new();
        for identifier in identifiers.keys() {
            let object = archive.object(*identifier).ok_or(Fault::Invalid)?;
            let index = single_message(object, message_type)?.ok_or(Fault::Invalid)?;
            output.push(Member {
                object: object.clone(),
                index,
            });
        }
        Ok(output)
    };
    let cell_tiles = tiles(&cell_ids, CELL_RECORD_TILE_MESSAGE_TYPE)?;
    let range_tiles = tiles(&range_ids, RANGE_PRECEDENTS_TILE_MESSAGE_TYPE)?;
    Ok(Family {
        owners,
        cell_tiles,
        range_tiles,
        direct_uuid,
    })
}

/// Clone `family` for the new drawable on the shared formula-graph
/// rewrites (`litchi_numbers_wire::formula_graph`, record 0368); returns
/// each clone's identifier, source and component. The new internal owner
/// identifiers start above every owner and every owner-map entry (policy
/// 0368a: the host allocates above the owners only).
#[allow(
    clippy::too_many_lines,
    reason = "the port keeps the host editor's clone_table_formula_graph sequence"
)]
fn clone_owner_family(
    working: &mut Working<'_>,
    component: &str,
    (engine_id, engine_index): (u64, usize),
    family: &Family,
    (info, new_info): (u64, u64),
    (source_uuid, new_uuid): (&Uuid, &Uuid),
    limits: litchi_iwa_core::Limits,
) -> R<Vec<(u64, u64, String)>> {
    let graph_limits = formula_graph::Limits::DEFAULT;
    let count = family.owners.len() + family.cell_tiles.len() + family.range_tiles.len();
    let fresh = working.allocate(count)?;
    let mut object_ids: Vec<(u64, u64)> = vec![(info, new_info)];
    let sources = family
        .owners
        .iter()
        .map(|(member, _)| &member.object)
        .chain(family.cell_tiles.iter().map(|member| &member.object))
        .chain(family.range_tiles.iter().map(|member| &member.object));
    let mut order = Vec::new();
    for (object, replacement) in sources.zip(&fresh) {
        let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
        object_ids.push((identifier, *replacement));
        order.push((identifier, *replacement));
    }
    object_ids.sort_unstable();

    let engine_data = working.object(engine_id)?.messages[engine_index]
        .data
        .clone();
    let all_owners: Vec<&[u8]> = working
        .archive(component)?
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .collect();
    let mut next_internal = formula_graph::next_owner_id(&engine_data, &all_owners, &graph_limits)?;
    let mut owner_ids: Vec<(u32, u32)> = Vec::new();
    for (_, owner) in &family.owners {
        owner_ids.push((owner.internal, next_internal));
        next_internal = next_internal.checked_add(1).ok_or(Fault::Invalid)?;
    }
    let internal_of: HashMap<u32, u32> = owner_ids.iter().copied().collect();
    owner_ids.sort_unstable();

    let wire = |uuid: &Uuid| WireUuid::new(uuid.lower, uuid.upper);
    let new_owner_uuid = owner_uuid_for_table(new_uuid);
    let mut uuids: BTreeMap<WireUuid, WireUuid> = BTreeMap::new();
    uuids.insert(wire(source_uuid), wire(new_uuid));
    uuids.insert(wire(&family.direct_uuid), wire(&new_owner_uuid));
    for (_, owner) in &family.owners {
        let replacement = wire(&translate_owner_uuid(
            &owner.uid,
            &family.direct_uuid,
            &new_owner_uuid,
        ));
        match uuids.get(&wire(&owner.uid)) {
            Some(existing) if *existing != replacement => return invalid(),
            Some(_) => {},
            None => {
                uuids.insert(wire(&owner.uid), replacement);
            },
        }
    }
    let uuids: Vec<(WireUuid, WireUuid)> = uuids.into_iter().collect();
    let remap = formula_graph::Remap::new(&object_ids, &owner_ids, &uuids)?;

    let mut cloned = Vec::new();
    let mut new_owners = Vec::new();
    let mut map_entries = Vec::new();
    let mut formula_count = 0_u64;
    let target = |source: u64| -> R<u64> {
        order
            .iter()
            .find(|(identifier, _)| *identifier == source)
            .map(|(_, replacement)| *replacement)
            .ok_or(E(Fault::Invalid))
    };
    let object_remap: HashMap<u64, u64> = object_ids.iter().copied().collect();
    for (member, owner) in &family.owners {
        let source = member
            .object
            .archive_info
            .identifier
            .ok_or(Fault::Invalid)?;
        let new_id = target(source)?;
        formula_count = formula_count
            .checked_add(formula_cell_count(owner, &family.cell_tiles)?)
            .ok_or(Fault::Invalid)?;
        let mut messages = member.object.messages.clone();
        messages[member.index].data =
            formula_graph::remap_owner(member.data(), remap, &graph_limits)?;
        cloned.push(clone_metadata(
            &member.object,
            new_id,
            messages,
            &object_remap,
        )?);
        new_owners.push(new_id);
        let uid = uuids
            .iter()
            .find(|(source, _)| *source == wire(&owner.uid))
            .map_or(wire(&owner.uid), |(_, replacement)| *replacement);
        map_entries.push((internal_of[&owner.internal], uid));
    }
    for member in &family.cell_tiles {
        let source = member
            .object
            .archive_info
            .identifier
            .ok_or(Fault::Invalid)?;
        let mut messages = member.object.messages.clone();
        messages[member.index].data =
            formula_graph::remap_cell_record_tile(member.data(), remap, &graph_limits)?;
        cloned.push(clone_metadata(
            &member.object,
            target(source)?,
            messages,
            &object_remap,
        )?);
    }
    for member in &family.range_tiles {
        let source = member
            .object
            .archive_info
            .identifier
            .ok_or(Fault::Invalid)?;
        let mut messages = member.object.messages.clone();
        messages[member.index].data =
            formula_graph::remap_range_precedents_tile(member.data(), remap, &graph_limits)?;
        cloned.push(clone_metadata(
            &member.object,
            target(source)?,
            messages,
            &object_remap,
        )?);
    }

    // The engine lists the new owners, maps their internal identifiers and
    // counts their formulas.
    let engine_data = formula_graph::append_owners_to_engine(
        &engine_data,
        &new_owners,
        &map_entries,
        formula_count,
        &graph_limits,
    )?;
    let engine_object = working.object_mut(engine_id)?;
    references::replace(
        engine_object,
        engine_index,
        engine_data,
        &[],
        &new_owners,
        &[2, 6],
        limits,
    )?;
    let mut output = Vec::new();
    for (object, (source, identifier)) in cloned.into_iter().zip(order.iter().copied()) {
        working.insert(component, object)?;
        output.push((identifier, source, component.to_owned()));
    }
    Ok(output)
}

/// The number of distinct formula cells an owner records.
fn formula_cell_count(owner: &Owner, tiles: &[Member]) -> R<u64> {
    let mut coordinates: HashSet<(u64, u64)> = owner.cells.iter().copied().collect();
    for identifier in &owner.cell_tiles {
        let tile = tiles
            .iter()
            .find(|member| member.object.archive_info.identifier == Some(*identifier))
            .ok_or(Fault::Invalid)?;
        coordinates.extend(cell_records(tile.data(), &[4])?);
    }
    u64::try_from(coordinates.len()).map_err(|_error| E(Fault::Invalid))
}

/// Add a sidebar node for `target` under `sheet`'s node, when the sidebar
/// records the sheet; returns the node and its component.
fn append_sidebar_node(
    working: &mut Working<'_>,
    sheet: u64,
    target: u64,
) -> R<Option<(u64, String)>> {
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let roots = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SIDEBAR_FIELD],
    )?;
    let [root] = roots.as_slice() else {
        return Ok(None);
    };
    let root_object = working.object(*root)?;
    let root_index = structure::unique_message(root_object, &[TREE_NODE_MESSAGE_TYPE])?;
    let mut parent = None;
    for node in references::identifiers(
        &root_object.messages[root_index].data,
        &[TREE_CHILDREN_FIELD],
    )? {
        let node_object = working.object(node)?;
        let node_index = structure::unique_message(node_object, &[TREE_NODE_MESSAGE_TYPE])?;
        if references::identifiers(&node_object.messages[node_index].data, &[TREE_OBJECT_FIELD])?
            .as_slice()
            == [sheet]
        {
            parent = Some((node, node_index));
        }
    }
    let Some((parent, parent_index)) = parent else {
        return Ok(None);
    };
    let node = *working.allocate(1)?.first().ok_or(Fault::Invalid)?;
    let component = working.locate(parent)?.ok_or(Fault::Invalid)?;
    let mut data = Vec::new();
    append_length_delimited_field(
        &mut data,
        TREE_OBJECT_FIELD,
        &references::reference_payload(target)?,
        WireLimits::DEFAULT,
    )?;
    let template = working.object(parent)?.clone();
    let object = structure::new_object(
        node,
        Some(&template),
        TREE_NODE_MESSAGE_TYPE,
        data,
        vec![target],
        working.archive_limits(),
    )?;
    working.insert(&component, object)?;
    let children = references::identifiers(
        &working.object(parent)?.messages[parent_index].data,
        &[TREE_CHILDREN_FIELD],
    )?;
    let mut after = children.clone();
    after.push(node);
    rewrite_list(
        working,
        parent,
        parent_index,
        &[TREE_CHILDREN_FIELD],
        &children,
        &after,
    )?;
    Ok(Some((node, component)))
}

/// Whether every dependency edge of table drawable `info`'s owner family
/// stays inside the family (the host editor's
/// `table_formula_graph_is_self_contained`). A sheet copy refuses a table
/// whose formulas read another table: their edges would need a sheet-wide
/// remap.
pub(in crate::package) fn self_contained(working: &Working<'_>, info: u64) -> FaultResult<bool> {
    contained(working, info).map_err(|error| error.0)
}

fn contained(working: &Working<'_>, info: u64) -> R<bool> {
    let Some((component, _, _)) = engine(working)? else {
        return Ok(true);
    };
    let family = owner_family(working, &component, info)?;
    let internal: HashSet<u64> = family
        .owners
        .iter()
        .map(|(_, owner)| u64::from(owner.internal))
        .collect();
    let uuids: HashSet<(u64, u64)> = family
        .owners
        .iter()
        .map(|(_, owner)| uuid_key(&owner.uid))
        .collect();
    let edges_inside = |data: &[u8], path: &[u32]| -> R<bool> {
        for edges in payloads_at(data, path)? {
            if !repeated_varint_values(edges, 5)?
                .iter()
                .all(|identifier| internal.contains(identifier))
            {
                return Ok(false);
            }
        }
        Ok(true)
    };
    for (member, _) in &family.owners {
        let data = member.data();
        if !edges_inside(data, &[4, 1, 6])? {
            return Ok(false);
        }
        for dependency in payloads_at(data, &[5, 2])? {
            if bytes(dependency, 3)?.is_some() {
                return Ok(false);
            }
            if let Some(reference) = bytes(dependency, 4)?
                && !internal.contains(&varint(reference, 1)?.unwrap_or(0))
            {
                return Ok(false);
            }
        }
        for path in [&[14, 1, 1][..], &[14, 2, 1]] {
            for uuid in payloads_at(data, path)? {
                if !uuids.contains(&uuid_key(&read_uuid(uuid)?)) {
                    return Ok(false);
                }
            }
        }
    }
    for member in &family.cell_tiles {
        if !edges_inside(member.data(), &[4, 6])? {
            return Ok(false);
        }
    }
    for member in &family.range_tiles {
        if !internal.contains(&varint(member.data(), 1)?.unwrap_or(0)) {
            return Ok(false);
        }
    }
    Ok(true)
}
