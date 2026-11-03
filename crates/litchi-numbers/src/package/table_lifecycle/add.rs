//! A new empty table, cloned from the structure of the workbook's first
//! table (the template).
//!
//! The template keeps the new table native: its drawable settings, styles,
//! default row and column sizes and storage layout are the workbook's own.
//! The clone takes the template's table drawable and model, and one fresh
//! copy of each storage object the model owns, emptied: tiles hold no rows,
//! header buckets no headers, data lists no entries, the merge map no
//! ranges, and the row/column UID map and stroke sidecar are rebuilt for the
//! requested size. Styles stay shared. Everything the template owns beyond
//! its storage (sort, filter, category, pivot and hidden-state owners,
//! title, caption, comment) is left out. The table gets a new identity (a
//! table UUID, object UUIDs in the registry), a CalculationEngine owner so
//! formulas can be authored in it, a place at the end of its sheet's
//! drawables and a sidebar node. No clock or randomness is consulted: fresh
//! identities derive from the registry payload and the allocated
//! identifiers, so the same edit of the same source is byte-identical.
//!
//! The payloads come from `litchi_numbers_wire::table_create`, shared with
//! the Pages and Keynote table owners. A workbook with no table left to
//! clone takes the shared from-scratch graph (`table_create::fresh`,
//! record 0420) styled by its theme's first table preset, as the host
//! editor's `bootstrap_empty_table_graph` builds it.

use std::collections::HashMap;

use litchi_iwa_common::wire::{
    append_length_delimited_field, parse_wire_view, transform_length_delimited_field,
};

use super::super::Package;
use super::super::structure::{
    self, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD, Fault, FaultResult, Published,
    TREE_CHILDREN_FIELD, TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    metadata::{self, Transition},
    references,
};
use super::{Error, rewrite_list};
use litchi_numbers_wire::table_create::{self, OwnerGeometry, Shape, TableUuid, fresh};
use litchi_numbers_wire::table_write::WriteError;

/// `TST.TableModelArchive`.
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// The header versions native Numbers writes on a formula owner.
const FORMULA_OWNER_VERSIONS: [u32; 3] = [3, 2, 10];
/// The most row and column identities one table may need (the native
/// editor's ceiling).
const MAX_TABLE_UIDS: u64 = 1_100_000;
/// How far a new table is offset from a template on the same sheet.
const POSITION_OFFSET: f32 = 40.0;
/// `TN.DocumentArchive.theme`.
const DOCUMENT_THEME_FIELD: u32 = 6;
/// `TN.ThemeArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
/// The theme's table style presets: `TSS.ThemeArchive` (field 1), its
/// `TST.ThemePresetsArchive` extension (200), `table_style_presets` (1).
const THEME_TABLE_PRESETS: [u32; 3] = [1, 200, 1];

/// Table-info fields the clone drops, as the host editor clears them:
/// editing state, summary and category models, view UIDs and pivot state
/// (record 0370f; the group and hidden-state UUIDs, the pasteboard mapping
/// and the pivot flag are kept, as native tables carry them).
const DROPPED_INFO_FIELDS: [u32; 6] = [3, 4, 5, 6, 15, 17];
/// Drawable fields the clone drops: comment, pencil annotations, title and
/// caption.
const DROPPED_DRAWABLE_FIELDS: [u32; 4] = [6, 9, 10, 11];

/// The requested shape of the new table.
pub(super) struct Request<'n> {
    pub(super) sheet: usize,
    pub(super) name: &'n str,
    pub(super) rows: u32,
    pub(super) columns: u32,
}

/// Validate the requested size.
pub(super) fn check_size(rows: u32, columns: u32) -> Result<(), Error> {
    if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::InvalidSize { rows, columns },
        ));
    }
    Ok(())
}

/// The first table drawable of the workbook, in sheet and drawable order.
fn template(working: &Working<'_>) -> FaultResult<Option<(u64, u64)>> {
    let document = working.object(1)?;
    let index = structure::document_message(document)?;
    for sheet in references::identifiers(&document.messages[index].data, &[DOCUMENT_SHEETS_FIELD])?
    {
        let object = working.object(sheet)?;
        let (sheet_index, path) = structure::sheet_message(object)?;
        for drawable in references::identifiers(&object.messages[sheet_index].data, path)? {
            if structure::table_info_message(working.object(drawable)?).is_ok() {
                return Ok(Some((sheet, drawable)));
            }
        }
    }
    Ok(None)
}

/// The single reference payload in field `number` of `data`.
fn single_reference(data: &[u8], number: u32) -> FaultResult<Option<u64>> {
    match references::identifiers(data, &[number])?.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Fault::Invalid),
    }
}

/// Encode `value` as a protobuf string payload of a canonical UUID.
fn table_uuid(seed: u64) -> TableUuid {
    let high = metadata::mix(seed ^ 0x7461_626c_6555_5549);
    let low = metadata::mix(high ^ seed);
    TableUuid::from_halves(high, low)
}

/// A [`Fault`] carried through the common wire transform callbacks.
struct WireFault(Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(structure::map_wire(&error))
    }
}

/// The drawable of the new table.
fn info_payload(
    template: &[u8],
    template_sheet: u64,
    sheet: u64,
    model: u64,
) -> FaultResult<Vec<u8>> {
    let mut data = structure::strip_fields(template, &DROPPED_INFO_FIELDS)?;
    data = transform_length_delimited_field(&data, 1, |drawable| {
        let mut drawable =
            structure::strip_fields(drawable, &DROPPED_DRAWABLE_FIELDS).map_err(WireFault)?;
        drawable = structure::put_bytes(
            &drawable,
            2,
            &references::reference_payload(sheet).map_err(WireFault)?,
        )
        .map_err(WireFault)?;
        if template_sheet == sheet {
            drawable = transform_length_delimited_field(&drawable, 1, |geometry| {
                transform_length_delimited_field(geometry, 1, |position| {
                    offset_point(position).map_err(WireFault)
                })
            })?;
        }
        Ok::<_, WireFault>(drawable)
    })
    .map_err(|error| error.0)?;
    structure::put_bytes(&data, 2, &references::reference_payload(model)?)
}

/// Move a `TSP.Point` by [`POSITION_OFFSET`] on both axes.
fn offset_point(point: &[u8]) -> FaultResult<Vec<u8>> {
    let view = parse_wire_view(point, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| structure::map_wire(&error))?;
    let mut output = Vec::new();
    structure::reserve(&mut output, point.len())?;
    for field in view.fields() {
        let payload = field.payload();
        if (field.number() == 1 || field.number() == 2)
            && field.wire_type() == 5
            && payload.len() == 4
        {
            let value = f32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let moved = value + POSITION_OFFSET;
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

/// The CalculationEngine root (component, object and message).
fn engine(working: &Working<'_>) -> FaultResult<Option<(String, u64, usize)>> {
    let mut found = None;
    for component in working.source().state.components.catalog().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                    if found
                        .replace((component.name().to_owned(), identifier, index))
                        .is_some()
                    {
                        return Err(Fault::Unsupported);
                    }
                }
            }
        }
    }
    Ok(found)
}

/// The graph a new table starts from: its drawable, model, the component
/// holding the drawable, its identity, its owner's geometry and every
/// object it inserted with the object it was cloned from (`None` for an
/// object built from scratch).
struct Graph {
    info: u64,
    info_component: String,
    seed: u64,
    uuid: TableUuid,
    geometry: OwnerGeometry,
    inserted: Vec<(u64, String)>,
    origins: HashMap<u64, Option<u64>>,
}

/// Clone the template table `template_info` of `template_sheet` for `sheet`.
fn clone_template(
    working: &mut Working<'_>,
    request: &Request<'_>,
    template_sheet: u64,
    template_info: u64,
    sheet: u64,
) -> Result<Graph, Error> {
    // The template's drawable, model and owned storage.
    let info_object = working.object(template_info)?.clone();
    let info_component = working.locate(template_info)?.ok_or(Error::InvalidSource)?;
    if info_object.messages.len() != 1 {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::UnsupportedSource,
        ));
    }
    let model = single_reference(&info_object.messages[0].data, 2)?.ok_or(Error::InvalidSource)?;
    let model_object = working.object(model)?.clone();
    let model_component = working.locate(model)?.ok_or(Error::InvalidSource)?;
    if model_object.messages.len() != 1
        || model_object.messages[0].type_ != TABLE_MODEL_MESSAGE_TYPE
    {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::UnsupportedSource,
        ));
    }
    let owned = table_create::owned_storage(&model_object.messages[0].data).map_err(write_fault)?;

    // Fresh identities.
    // The host editor's order: the drawable, the model and the owned
    // storage in source-identifier order; the formula owner and the sidebar
    // node follow (record 0370f).
    let fresh = working.allocate(owned.len() + 2)?;
    let (new_info, new_model) = (fresh[0], fresh[1]);
    let mut remap: HashMap<u64, u64> = HashMap::new();
    for ((identifier, _), replacement) in owned.iter().zip(&fresh[2..]) {
        remap.insert(*identifier, *replacement);
    }
    let seed = metadata::identity_seed(working, new_info)?;
    let uuid = table_uuid(seed);
    let limits = working.archive_limits();

    // The storage clones, in their sources' components.
    let mut inserted: Vec<(u64, String)> = Vec::new();
    let mut origins: HashMap<u64, Option<u64>> = HashMap::new();
    for (identifier, role) in &owned {
        let object = working.object(*identifier)?.clone();
        if object.messages.len() != 1 {
            return Err(Error::Refused(
                crate::package::table_lifecycle::Refusal::UnsupportedSource,
            ));
        }
        let component = working.locate(*identifier)?.ok_or(Error::InvalidSource)?;
        let data = table_create::empty_storage(
            *role,
            &object.messages[0].data,
            request.rows,
            request.columns,
            new_model,
        )
        .map_err(write_fault)?;
        let clone = structure::new_object(
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

    // The model, with its styles shared and its storage remapped.
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
    let model_references = model_references(&model_data)?;
    let new_model_object = structure::new_object(
        new_model,
        Some(&model_object),
        TABLE_MODEL_MESSAGE_TYPE,
        model_data,
        model_references.clone(),
        limits,
    )?;
    working.insert(&model_component, new_model_object)?;
    inserted.push((new_model, model_component.clone()));
    origins.insert(new_model, Some(model));

    // The drawable, parented to the target sheet.
    let info_data = info_payload(
        &info_object.messages[0].data,
        template_sheet,
        sheet,
        new_model,
    )?;
    let mut info_references: Vec<u64> = info_object.archive_info.message_infos[0]
        .object_references
        .iter()
        .copied()
        .filter(|identifier| *identifier != model && *identifier != template_sheet)
        .filter(|identifier| references_payload(&info_data, *identifier))
        .collect();
    info_references.push(sheet);
    info_references.push(new_model);
    let new_info_object = structure::new_object(
        new_info,
        Some(&info_object),
        info_object.messages[0].type_,
        info_data,
        info_references.clone(),
        limits,
    )?;
    working.insert(&info_component, new_info_object)?;
    inserted.push((new_info, info_component.clone()));
    origins.insert(new_info, Some(template_info));

    Ok(Graph {
        info: new_info,
        info_component,
        seed,
        uuid,
        geometry,
        inserted,
        origins,
    })
}

/// Build a table from scratch for a workbook with no table left to clone,
/// as the host editor's `bootstrap_empty_table_graph` does (record 0420):
/// the shared graph of `table_create::fresh` with the theme's first table
/// preset, every object in the sheet's component and registered.
fn bootstrap(working: &mut Working<'_>, request: &Request<'_>, sheet: u64) -> Result<Graph, Error> {
    let styles = theme_styles(working)?;
    let component = working.locate(sheet)?.ok_or(Error::InvalidSource)?;
    let first = *working
        .allocate(fresh::Ids::consecutive(0, fresh::Profile::Bootstrap).count())?
        .first()
        .ok_or(Error::InvalidSource)?;
    let ids = fresh::Ids::consecutive(first, fresh::Profile::Bootstrap);
    let seed = metadata::identity_seed(working, ids.info)?;
    let uuid = table_uuid(seed);
    let shape = Shape {
        name: request.name,
        rows: request.rows,
        columns: request.columns,
        row_height: None,
        column_width: None,
    };
    let objects = fresh::table(
        &ids,
        fresh::Placement {
            parent: sheet,
            frame: fresh::DEFAULT_FRAME,
        },
        &styles,
        &shape,
        &uuid.text,
        ids.model,
        fresh::Profile::Bootstrap,
    )
    .map_err(write_fault)?;
    let geometry = OwnerGeometry::of(&objects[1].payload, request.rows, request.columns)
        .map_err(write_fault)?;
    let limits = working.archive_limits();
    let mut inserted = Vec::with_capacity(objects.len());
    let mut origins = HashMap::with_capacity(objects.len());
    for object in objects {
        let mut built = structure::new_object(
            object.identifier,
            None,
            object.message_type,
            object.payload,
            object.references,
            limits,
        )?;
        built.archive_info.message_infos[0].versions = fresh::OBJECT_VERSIONS.to_vec().into();
        working.insert(&component, built)?;
        inserted.push((object.identifier, component.clone()));
        origins.insert(object.identifier, None);
    }
    Ok(Graph {
        info: ids.info,
        info_component: component,
        seed,
        uuid,
        geometry,
        inserted,
        origins,
    })
}

/// The theme's first table preset and the styles its network names (the
/// host's `theme_table_style`).
fn theme_styles(working: &Working<'_>) -> Result<fresh::Styles, Error> {
    let document = working.object(1)?;
    let index = structure::document_message(document)?;
    let theme = references::identifiers(&document.messages[index].data, &[DOCUMENT_THEME_FIELD])?;
    let [theme] = theme.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let payload = |identifier: u64, message_type: u32| -> Result<Vec<u8>, Error> {
        let object = working.object(identifier)?;
        let index = structure::unique_message(object, &[message_type])?;
        Ok(object.messages[index].data.clone())
    };
    let theme = payload(*theme, THEME_MESSAGE_TYPE)?;
    let presets = references::identifiers(&theme, &THEME_TABLE_PRESETS)?;
    let (position, preset) = presets
        .iter()
        .copied()
        .enumerate()
        .find(|(_, identifier)| *identifier != 0)
        .ok_or(Error::Refused(
            crate::package::table_lifecycle::Refusal::NoTemplate,
        ))?;
    let preset_data = payload(preset, fresh::PRESET_MESSAGE_TYPE)?;
    let network = fresh::preset_network(&preset_data).map_err(write_fault)?;
    let network_data = payload(network, fresh::STYLE_NETWORK_MESSAGE_TYPE)?;
    let position = u32::try_from(position).map_err(|_error| Error::InvalidSource)?;
    let styles =
        fresh::preset_styles(preset, position, &preset_data, &network_data).map_err(write_fault)?;
    for identifier in styles.objects() {
        working.locate(identifier)?.ok_or(Error::InvalidSource)?;
    }
    Ok(styles)
}

/// A shared payload builder's error as a structural fault.
fn write_fault(error: WriteError) -> Fault {
    match error {
        WriteError::Unsupported => Fault::Unsupported,
        WriteError::Allocation { requested } => Fault::Allocation { amount: requested },
        _ => Fault::Invalid,
    }
}

/// Add the new table and return the published candidate.
pub(super) fn add_table(source: &Package, request: &Request<'_>) -> Result<Published, Error> {
    check_size(request.rows, request.columns)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let sidebar_root = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SIDEBAR_FIELD],
    )?;
    let sheet = *sheets.get(request.sheet).ok_or(Error::Refused(
        crate::package::table_lifecycle::Refusal::SheetNotFound,
    ))?;
    let graph = match template(&working)? {
        Some((template_sheet, template_info)) => {
            clone_template(&mut working, request, template_sheet, template_info, sheet)?
        },
        None => bootstrap(&mut working, request, sheet)?,
    };
    let Graph {
        info: new_info,
        info_component,
        seed,
        uuid,
        geometry,
        mut inserted,
        origins,
    } = graph;
    let limits = working.archive_limits();

    // Attach to the sheet's drawables.
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)?;
    let mut after = drawables.clone();
    after.push(new_info);
    rewrite_list(
        &mut working,
        sheet,
        sheet_index,
        sheet_path,
        &drawables,
        &after,
    )?;

    // A formula owner for the new table, when the workbook has an engine.
    let mut owner_component = None;
    let mut new_owner = None;
    if let Some((component, engine_id, engine_index)) = engine(&working)? {
        let owner = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
        new_owner = Some(owner);
        let new_owner = owner;
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
        let internal =
            table_create::next_internal_owner(&owners, &engine_data).map_err(write_fault)?;
        let uid = uuid.owner_uid();
        let owner_data =
            table_create::owner_payload(uid, internal, new_info, geometry).map_err(write_fault)?;
        let mut owner_object = structure::new_object(
            new_owner,
            None,
            FORMULA_OWNER_MESSAGE_TYPE,
            owner_data,
            Vec::new(),
            limits,
        )?;
        owner_object.archive_info.message_infos[0].versions =
            FORMULA_OWNER_VERSIONS.to_vec().into();
        working.insert(&component, owner_object)?;
        inserted.push((new_owner, component.clone()));
        let engine_data = table_create::engine_payload(&engine_data, new_owner, internal, uid)
            .map_err(write_fault)?;
        let engine_object = working.object_mut(engine_id)?;
        references::replace(
            engine_object,
            engine_index,
            engine_data,
            &[],
            &[new_owner],
            &[2, 6],
            limits,
        )?;
        owner_component = Some(component);
    }

    // The sheet's sidebar node lists its tables.
    let mut node_parent = None;
    if let [root] = sidebar_root.as_slice() {
        let root_object = working.object(*root)?;
        let root_index = structure::unique_message(root_object, &[TREE_NODE_MESSAGE_TYPE])?;
        for node in references::identifiers(
            &root_object.messages[root_index].data,
            &[TREE_CHILDREN_FIELD],
        )? {
            let node_object = working.object(node)?;
            let node_index = structure::unique_message(node_object, &[TREE_NODE_MESSAGE_TYPE])?;
            if references::identifiers(
                &node_object.messages[node_index].data,
                &[TREE_OBJECT_FIELD],
            )?
            .as_slice()
                == [sheet]
            {
                node_parent = Some((node, node_index));
            }
        }
    }
    if let Some((parent, parent_index)) = node_parent {
        let new_node = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
        let parent_component = working.locate(parent)?.ok_or(Error::InvalidSource)?;
        let mut data = Vec::new();
        append_length_delimited_field(
            &mut data,
            TREE_OBJECT_FIELD,
            &references::reference_payload(new_info)?,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| structure::map_wire(&error))?;
        let template_node = working.object(parent)?.clone();
        let node_object = structure::new_object(
            new_node,
            Some(&template_node),
            TREE_NODE_MESSAGE_TYPE,
            data,
            vec![new_info],
            limits,
        )?;
        working.insert(&parent_component, node_object)?;
        inserted.push((new_node, parent_component));
        let children = references::identifiers(
            &working.object(parent)?.messages[parent_index].data,
            &[TREE_CHILDREN_FIELD],
        )?;
        let mut after = children.clone();
        after.push(new_node);
        rewrite_list(
            &mut working,
            parent,
            parent_index,
            &[TREE_CHILDREN_FIELD],
            &children,
            &after,
        )?;
    }

    // Registry: UUIDs for the clones of registered objects, and every new
    // cross-component reference.
    let mut transition = Transition::default();
    // In identifier order, as the host editor registers them.
    inserted.sort_by_key(|(identifier, _)| *identifier);
    if let Some(registry) = metadata::registry(&working)? {
        let registered = |identifier: u64| {
            registry
                .uuids
                .iter()
                .any(|binding| binding.object == identifier)
        };
        for (identifier, component) in &inserted {
            let source_identifier = origins.get(identifier).copied();
            let component_id = registry.component_for_entry(component)?.identifier;
            let needs_uuid = match source_identifier {
                Some(Some(source_identifier)) => registered(source_identifier),
                Some(None) => true,
                None => Some(*identifier) == new_owner,
            };
            if needs_uuid {
                transition.register_uuid(&registry, component_id, *identifier, seed)?;
            }
            let referenced: Vec<u64> = working.object(*identifier)?.archive_info.message_infos[0]
                .object_references
                .clone();
            for target in referenced {
                let target_component = working.locate(target)?.ok_or(Error::InvalidSource)?;
                let target_id = registry.component_for_entry(&target_component)?.identifier;
                transition.reference(&registry, component_id, target_id, target)?;
            }
        }
        // The sheet now addresses the drawable; the engine's owner addresses
        // the drawable through its `formula_owner`.
        let sheet_component = working.locate(sheet)?.ok_or(Error::InvalidSource)?;
        let sheet_component = registry.component_for_entry(&sheet_component)?.identifier;
        let info_component_id = registry.component_for_entry(&info_component)?.identifier;
        transition.reference(&registry, sheet_component, info_component_id, new_info)?;
        if let Some(component) = owner_component {
            let engine_component = registry.component_for_entry(&component)?.identifier;
            transition.reference(&registry, engine_component, info_component_id, new_info)?;
        }
    }
    Ok(working.publish(&transition)?)
}

/// The objects a model payload references, sorted, as the host editor's
/// `table_model_references` lists them (record 0370f): its style and owner
/// references, and its data store's headers, tiles and lists.
fn model_references(model: &[u8]) -> FaultResult<Vec<u64>> {
    const MODEL_FIELDS: [u32; 42] = [
        3, 24, 25, 26, 27, 18, 19, 20, 21, 30, 36, 48, 5, 34, 35, 38, 46, 49, 60, 61, 62, 63, 64,
        65, 66, 67, 68, 69, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 85, 86, 87, 88,
    ];
    const STORE_FIELDS: [u32; 15] = [2, 4, 5, 6, 11, 12, 16, 13, 15, 17, 18, 19, 20, 21, 22];
    let mut output = Vec::new();
    for field in MODEL_FIELDS.into_iter().chain([89]) {
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
/// nested messages. Only identifiers the header already declares are asked,
/// so a coincidental non-reference record is not a concern for correctness
/// of the declared set.
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
