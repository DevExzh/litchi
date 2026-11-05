//! A new, empty table graph on the structural working set (record 0420;
//! lifted from `litchi-numbers`' `table::lifecycle::add` so the Keynote
//! slide-table owner shares it).
//!
//! Two sources, as the host editor has them:
//!
//! - [`clone_template`], the host's `create_empty_table_graph`: the
//!   template table's drawable and model and one emptied copy of each
//!   storage object the model owns, with styles shared (the payloads of
//!   `litchi_numbers_wire::table_create`), every clone in its source's
//!   component;
//! - [`fresh`], the host's `bootstrap_empty_table_graph`: the shared
//!   from-scratch graph (`table_create::fresh`) styled by a caller-chosen
//!   stylesheet context, every object in the container's component.
//!
//! [`add_owner`] then gives the table its own CalculationEngine formula
//! owner when the package has an engine. The caller attaches the drawable
//! to its container (a sheet's drawables, a slide's owned drawables), adds
//! its own objects, and records the registry consequences of
//! [`Graph::inserted`] (UUIDs for clones of registered objects and every
//! fresh object, and every new cross-component reference), in the order
//! its host does. No clock or randomness is consulted: identities derive
//! from the registry payload and the allocated identifiers.

use std::collections::HashMap;

use litchi_iwa_common::wire::{parse_wire_view, transform_length_delimited_field};
use litchi_numbers_wire::table_create::{self, OwnerGeometry, Shape, TableUuid, fresh as graph};
use litchi_numbers_wire::table_write::WriteError;

use super::{
    Fault, FaultResult, TableLeaf, Working, map_wire, metadata, new_object, put_bytes, references,
    reserve, strip_fields, unique_message,
};

/// `TST.TableModelArchive`.
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// The header versions native Numbers writes on a formula owner.
const FORMULA_OWNER_VERSIONS: [u32; 3] = [3, 2, 10];
/// Table-info fields a clone drops, as the host editor clears them:
/// editing state, summary and category models, view UIDs and pivot state
/// (record 0370f; the group and hidden-state UUIDs, the pasteboard mapping
/// and the pivot flag are kept, as native tables carry them).
const DROPPED_INFO_FIELDS: [u32; 6] = [3, 4, 5, 6, 15, 17];
/// Drawable fields a clone drops: comment, pencil annotations, title and
/// caption.
const DROPPED_DRAWABLE_FIELDS: [u32; 4] = [6, 9, 10, 11];

/// The requested table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Request<'n> {
    /// The table name.
    pub name: &'n str,
    /// Rows, header rows included.
    pub rows: u32,
    /// Columns, header columns included.
    pub columns: u32,
}

/// A staged new table.
#[derive(Debug, Clone)]
pub struct Graph {
    /// The drawable.
    pub info: u64,
    /// The model.
    pub model: u64,
    /// The component holding the drawable.
    pub info_component: String,
    /// The identity seed (from the registry and the drawable's identifier).
    pub seed: u64,
    /// The table's identity.
    pub uuid: TableUuid,
    /// The geometry its formula owner records.
    pub geometry: OwnerGeometry,
    /// Every inserted object and its component, in insertion order.
    pub inserted: Vec<(u64, String)>,
    /// The object each inserted object was cloned from (`None`: built from
    /// scratch, registered in any case).
    pub origins: HashMap<u64, Option<u64>>,
}

/// A shared payload builder's error as a structural fault.
#[must_use]
pub fn write_fault(error: WriteError) -> Fault {
    match error {
        WriteError::Unsupported => Fault::Unsupported,
        WriteError::Allocation { requested } => Fault::Allocation { amount: requested },
        _ => Fault::Invalid,
    }
}

/// The table identity the host editor derives from `seed`.
#[must_use]
pub fn table_uuid(seed: u64) -> TableUuid {
    let high = metadata::mix(seed ^ 0x7461_626c_6555_5549);
    let low = metadata::mix(high ^ seed);
    TableUuid::from_halves(high, low)
}

/// A [`Fault`] carried through the common wire transform callbacks.
struct WireFault(Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(map_wire(&error))
    }
}

/// The single reference in field `number` of `data`.
fn single_reference(data: &[u8], number: u32) -> FaultResult<Option<u64>> {
    match references::identifiers(data, &[number])?.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Fault::Invalid),
    }
}

/// Move a `TSP.Point` by `offset` on both axes.
fn offset_point(point: &[u8], offset: f32) -> FaultResult<Vec<u8>> {
    let view = parse_wire_view(point, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| map_wire(&error))?;
    let mut output = Vec::new();
    reserve(&mut output, point.len())?;
    for field in view.fields() {
        let payload = field.payload();
        if (field.number() == 1 || field.number() == 2)
            && field.wire_type() == 5
            && payload.len() == 4
        {
            let value = f32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let moved = value + offset;
            if !moved.is_finite() {
                return Err(Fault::Invalid);
            }
            output.extend_from_slice(field.key());
            output.extend_from_slice(&moved.to_le_bytes());
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

/// The drawable of a clone: the template's without its dropped fields,
/// parented to `parent`, naming `model`, its position moved by `offset`
/// when given.
fn info_payload(
    template: &[u8],
    parent: u64,
    model: u64,
    offset: Option<f32>,
) -> FaultResult<Vec<u8>> {
    let mut data = strip_fields(template, &DROPPED_INFO_FIELDS)?;
    data = transform_length_delimited_field(&data, 1, |drawable| {
        let mut drawable = strip_fields(drawable, &DROPPED_DRAWABLE_FIELDS).map_err(WireFault)?;
        drawable = put_bytes(
            &drawable,
            2,
            &references::reference_payload(parent).map_err(WireFault)?,
        )
        .map_err(WireFault)?;
        if let Some(offset) = offset {
            drawable = transform_length_delimited_field(&drawable, 1, |geometry| {
                transform_length_delimited_field(geometry, 1, |position| {
                    offset_point(position, offset).map_err(WireFault)
                })
            })?;
        }
        Ok::<_, WireFault>(drawable)
    })
    .map_err(|error| error.0)?;
    put_bytes(&data, 2, &references::reference_payload(model)?)
}

/// The objects a model payload references, sorted, as the host editor's
/// `table_model_references` lists them (record 0370f): its style and owner
/// references, and its data store's headers, tiles and lists.
fn model_references(model: &[u8]) -> FaultResult<Vec<u64>> {
    const MODEL_FIELDS: [u32; 43] = [
        3, 24, 25, 26, 27, 18, 19, 20, 21, 30, 36, 48, 5, 34, 35, 38, 46, 49, 60, 61, 62, 63, 64,
        65, 66, 67, 68, 69, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 85, 86, 87, 88, 89,
    ];
    const STORE_FIELDS: [u32; 15] = [2, 4, 5, 6, 11, 12, 16, 13, 15, 17, 18, 19, 20, 21, 22];
    let mut output = Vec::new();
    for field in MODEL_FIELDS {
        output.extend(references::identifiers(model, &[field])?);
    }
    output.extend(references::identifiers(model, &[4, 1, 2])?);
    output.extend(references::identifiers(model, &[4, 3, 1, 2])?);
    for field in STORE_FIELDS {
        output.extend(references::identifiers(model, &[4, field])?);
    }
    output.retain(|identifier| *identifier != 0);
    output.sort_unstable();
    output.dedup();
    Ok(output)
}

/// Whether `data` holds a `TSP.Reference` to `identifier` anywhere in its
/// nested messages.
fn references_payload(data: &[u8], identifier: u64) -> bool {
    let Ok(view) = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT) else {
        return false;
    };
    view.fields().any(|field| {
        field.wire_type() == 2
            && (references::reference_identifier(field.payload()).ok() == Some(identifier)
                || references_payload(field.payload(), identifier))
    })
}

/// Clone the template table `template_info` (whose drawable names
/// `template_parent`, if any) as an empty table for `parent`; the clone
/// sits `offset` points from the template when both share a parent.
///
/// # Errors
///
/// [`Fault::Unsupported`] for a template of more than one message per
/// object, [`Fault::Invalid`] for a malformed one, and resource faults.
pub fn clone_template<P: TableLeaf>(
    working: &mut Working<'_, P>,
    request: Request<'_>,
    template_info: u64,
    template_parent: Option<u64>,
    parent: u64,
    offset: f32,
) -> FaultResult<Graph> {
    let info_object = working.object(template_info)?.clone();
    let info_component = working.locate(template_info)?.ok_or(Fault::Invalid)?;
    if info_object.messages.len() != 1 {
        return Err(Fault::Unsupported);
    }
    let model = single_reference(&info_object.messages[0].data, 2)?.ok_or(Fault::Invalid)?;
    let model_object = working.object(model)?.clone();
    let model_component = working.locate(model)?.ok_or(Fault::Invalid)?;
    if model_object.messages.len() != 1
        || model_object.messages[0].type_ != TABLE_MODEL_MESSAGE_TYPE
    {
        return Err(Fault::Unsupported);
    }
    let owned = table_create::owned_storage(&model_object.messages[0].data).map_err(write_fault)?;

    // The host editor's order: the drawable, the model and the owned
    // storage in source-identifier order (record 0370f).
    let fresh = working.allocate(owned.len() + 2)?;
    let (new_info, new_model) = (fresh[0], fresh[1]);
    let mut remap: HashMap<u64, u64> = HashMap::new();
    for ((identifier, _), replacement) in owned.iter().zip(&fresh[2..]) {
        remap.insert(*identifier, *replacement);
    }
    let seed = metadata::identity_seed(working, new_info)?;
    let uuid = table_uuid(seed);
    let limits = working.archive_limits();

    let mut inserted: Vec<(u64, String)> = Vec::new();
    let mut origins: HashMap<u64, Option<u64>> = HashMap::new();
    for (identifier, role) in &owned {
        let object = working.object(*identifier)?.clone();
        if object.messages.len() != 1 {
            return Err(Fault::Unsupported);
        }
        let component = working.locate(*identifier)?.ok_or(Fault::Invalid)?;
        let data = table_create::empty_storage(
            *role,
            &object.messages[0].data,
            request.rows,
            request.columns,
            new_model,
        )
        .map_err(write_fault)?;
        let clone = new_object(
            remap[identifier],
            Some(&object),
            object.messages[0].type_,
            data,
            Vec::new(),
            limits,
        )?;
        working.insert(&component, clone)?;
        inserted.push((remap[identifier], component));
        origins.insert(remap[identifier], Some(*identifier));
    }

    let model_data = table_create::model_payload(
        &model_object.messages[0].data,
        &|identifier| remap.get(&identifier).copied(),
        &uuid,
        &Shape {
            name: request.name,
            rows: request.rows,
            columns: request.columns,
            row_height: None,
            column_width: None,
        },
    )
    .map_err(write_fault)?;
    let geometry =
        OwnerGeometry::of(&model_data, request.rows, request.columns).map_err(write_fault)?;
    let declared = model_references(&model_data)?;
    let new_model_object = new_object(
        new_model,
        Some(&model_object),
        TABLE_MODEL_MESSAGE_TYPE,
        model_data,
        declared,
        limits,
    )?;
    working.insert(&model_component, new_model_object)?;
    inserted.push((new_model, model_component));
    origins.insert(new_model, Some(model));

    let offset = (template_parent == Some(parent)).then_some(offset);
    let info_data = info_payload(&info_object.messages[0].data, parent, new_model, offset)?;
    let mut info_references: Vec<u64> = info_object.archive_info.message_infos[0]
        .object_references
        .iter()
        .copied()
        .filter(|identifier| *identifier != model && Some(*identifier) != template_parent)
        .filter(|identifier| references_payload(&info_data, *identifier))
        .collect();
    info_references.push(parent);
    info_references.push(new_model);
    let new_info_object = new_object(
        new_info,
        Some(&info_object),
        info_object.messages[0].type_,
        info_data,
        info_references,
        limits,
    )?;
    working.insert(&info_component, new_info_object)?;
    inserted.push((new_info, info_component.clone()));
    origins.insert(new_info, Some(template_info));
    Ok(Graph {
        info: new_info,
        model: new_model,
        info_component,
        seed,
        uuid,
        geometry,
        inserted,
        origins,
    })
}

/// The table style presets of a theme: `TSS.ThemeArchive` (the theme's
/// field 1), its `TST.ThemePresetsArchive` extension (200),
/// `table_style_presets` (1).
const THEME_TABLE_PRESETS: [u32; 3] = [1, 200, 1];

/// The first table preset of the theme object `theme` and the styles its
/// network names (the host's `theme_table_style`), or `None` when the theme
/// has no table preset.
///
/// # Errors
///
/// [`Fault::Invalid`] for a malformed theme, preset or style network, or a
/// named style the package does not hold.
pub fn theme_styles<P: TableLeaf>(
    working: &Working<'_, P>,
    theme: u64,
) -> FaultResult<Option<graph::Styles>> {
    let payload = |identifier: u64, message_type: Option<u32>| -> FaultResult<Vec<u8>> {
        let object = working.object(identifier)?;
        let index = match message_type {
            Some(message_type) => unique_message(object, &[message_type])?,
            None if object.messages.len() == 1 => 0,
            None => return Err(Fault::Invalid),
        };
        Ok(object.messages[index].data.clone())
    };
    let theme = payload(theme, None)?;
    let presets = references::identifiers(&theme, &THEME_TABLE_PRESETS)?;
    let Some((position, preset)) = presets
        .iter()
        .copied()
        .enumerate()
        .find(|(_, identifier)| *identifier != 0)
    else {
        return Ok(None);
    };
    let preset_data = payload(preset, Some(graph::PRESET_MESSAGE_TYPE))?;
    let network = graph::preset_network(&preset_data).map_err(write_fault)?;
    let network_data = payload(network, Some(graph::STYLE_NETWORK_MESSAGE_TYPE))?;
    let position = u32::try_from(position).map_err(|_error| Fault::Invalid)?;
    let styles =
        graph::preset_styles(preset, position, &preset_data, &network_data).map_err(write_fault)?;
    for identifier in styles.objects() {
        working.locate(identifier)?.ok_or(Fault::Invalid)?;
    }
    Ok(Some(styles))
}

/// Build a table from scratch for `parent` in `component`, styled by
/// `styles` (the host editor's bootstrap for a package with no table left
/// to clone).
///
/// # Errors
///
/// [`Fault::Invalid`] for a size or style context the shared graph
/// refuses, and resource faults.
pub fn fresh<P: TableLeaf>(
    working: &mut Working<'_, P>,
    request: Request<'_>,
    parent: u64,
    component: &str,
    styles: &graph::Styles,
) -> FaultResult<Graph> {
    let count = graph::Ids::consecutive(0, graph::Profile::Bootstrap).count();
    let first = *working.allocate(count)?.first().ok_or(Fault::Invalid)?;
    let ids = graph::Ids::consecutive(first, graph::Profile::Bootstrap);
    let seed = metadata::identity_seed(working, ids.info)?;
    let uuid = table_uuid(seed);
    let objects = graph::table(
        &ids,
        graph::Placement {
            parent,
            frame: graph::DEFAULT_FRAME,
        },
        styles,
        &Shape {
            name: request.name,
            rows: request.rows,
            columns: request.columns,
            row_height: None,
            column_width: None,
        },
        &uuid.text,
        ids.model,
        graph::Profile::Bootstrap,
    )
    .map_err(write_fault)?;
    let model = objects.get(1).ok_or(Fault::Invalid)?;
    let geometry =
        OwnerGeometry::of(&model.payload, request.rows, request.columns).map_err(write_fault)?;
    let limits = working.archive_limits();
    let mut inserted = Vec::with_capacity(objects.len());
    let mut origins = HashMap::with_capacity(objects.len());
    for object in objects {
        let mut built = new_object(
            object.identifier,
            None,
            object.message_type,
            object.payload,
            object.references,
            limits,
        )?;
        built.archive_info.message_infos[0].versions = graph::OBJECT_VERSIONS.to_vec().into();
        working.insert(component, built)?;
        inserted.push((object.identifier, component.to_owned()));
        origins.insert(object.identifier, None);
    }
    Ok(Graph {
        info: ids.info,
        model: ids.model,
        info_component: component.to_owned(),
        seed,
        uuid,
        geometry,
        inserted,
        origins,
    })
}

/// The CalculationEngine root (component, object and message), when the
/// package has one.
fn engine<P: TableLeaf>(working: &Working<'_, P>) -> FaultResult<Option<(String, u64, usize)>> {
    let mut found = None;
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                    if found.replace((name.clone(), identifier, index)).is_some() {
                        return Err(Fault::Unsupported);
                    }
                }
            }
        }
    }
    Ok(found)
}

/// Give `table` its own empty formula owner in the package's calculation
/// engine, registered in the engine's dependency tracker; returns the
/// owner and the engine's component, or `None` without an engine. The
/// owner joins [`Graph::inserted`] as a fresh object.
///
/// # Errors
///
/// [`Fault::Unsupported`] for a package with two engines; malformed
/// engine payloads and resource faults.
pub fn add_owner<P: TableLeaf>(
    working: &mut Working<'_, P>,
    table: &mut Graph,
) -> FaultResult<Option<(u64, String)>> {
    let Some((component, engine_id, engine_index)) = engine(working)? else {
        return Ok(None);
    };
    let owner = *working.allocate(1)?.first().ok_or(Fault::Invalid)?;
    let limits = working.archive_limits();
    let engine_data = working.object(engine_id)?.messages[engine_index]
        .data
        .clone();
    let owners: Vec<&[u8]> = working
        .archive(&component)?
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .collect();
    let internal = table_create::next_internal_owner(&owners, &engine_data).map_err(write_fault)?;
    let uid = table.uuid.owner_uid();
    let owner_data = table_create::owner_payload(uid, internal, table.info, table.geometry)
        .map_err(write_fault)?;
    let mut owner_object = new_object(
        owner,
        None,
        FORMULA_OWNER_MESSAGE_TYPE,
        owner_data,
        Vec::new(),
        limits,
    )?;
    owner_object.archive_info.message_infos[0].versions = FORMULA_OWNER_VERSIONS.to_vec().into();
    working.insert(&component, owner_object)?;
    table.inserted.push((owner, component.clone()));
    table.origins.insert(owner, None);
    let engine_data =
        table_create::engine_payload(&engine_data, owner, internal, uid).map_err(write_fault)?;
    let engine_object = working.object_mut(engine_id)?;
    references::replace(
        engine_object,
        engine_index,
        engine_data,
        &[],
        &[owner],
        &[2, 6],
        limits,
    )?;
    Ok(Some((owner, component)))
}
