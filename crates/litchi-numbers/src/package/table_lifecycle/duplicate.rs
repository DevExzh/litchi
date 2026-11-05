//! A populated copy of a table on a Numbers sheet (record 0370h).
//!
//! The copy itself — fresh identifiers for its drawable, model and owned
//! storage, the host editor's table UUID, its CalculationEngine owner
//! family, formula-list UUID rewrites and the clones' registry records — is
//! the shared engine `litchi_iwa_table::structure::duplicate`. This module
//! attaches the copy to its sheet: at the end of the sheet's drawables, with
//! a sidebar node under the sheet's node, and the sheet's, engine's and (for
//! a copy on another sheet) the drawable's references to each other.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_table::structure::duplicate::{self as engine, Drawable};

use super::super::Package;
use super::super::structure::{
    self, DOCUMENT_SIDEBAR_FIELD, Fault, FaultResult, Published, TREE_CHILDREN_FIELD,
    TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    metadata::{self, Transition},
    references,
};
use super::{Error, locate, rewrite_list};

pub(in crate::package) use engine::{copy_name, self_contained};

/// `TST.TableInfoArchive` and its legacy form.
const TABLE_INFO_MESSAGE_TYPES: [u32; 2] = [6_000, 6_003];
/// How far the copy is offset from its source (the host editor's
/// `DRAWABLE_DUPLICATE_OFFSET`).
const OFFSET: f32 = 10.0;

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
    let moved = placement.sheet != placement.source_sheet;
    let copy = engine::clone_graph(
        working,
        transition,
        info,
        Drawable {
            message_types: &TABLE_INFO_MESSAGE_TYPES,
            info_path: &[],
            offset: placement.offset,
            parent: moved.then_some((placement.source_sheet, placement.sheet)),
        },
        name,
    )?;
    let new_info = copy.info;

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

    // The node's, the sheet's, the engine's and (on another sheet) the
    // drawable's references across components.
    let Some(registry) = metadata::registry(working)? else {
        return Ok(new_info);
    };
    if let Some((node, component)) = &node {
        let component_id = registry.component_for_entry(component)?.identifier;
        let referenced = working.object(*node)?.archive_info.message_infos[0]
            .object_references
            .clone();
        for target in referenced {
            let target_component = working.locate(target)?.ok_or(Fault::Invalid)?;
            let target_id = registry.component_for_entry(&target_component)?.identifier;
            transition.reference(&registry, component_id, target_id, target)?;
        }
    }
    let info_component_id = registry
        .component_for_entry(&copy.info_component)?
        .identifier;
    let sheet_component = working.locate(placement.sheet)?.ok_or(Fault::Invalid)?;
    let sheet_component = registry.component_for_entry(&sheet_component)?.identifier;
    transition.reference(&registry, sheet_component, info_component_id, new_info)?;
    if let Some(component) = &copy.engine_component {
        let engine_component = registry.component_for_entry(component)?.identifier;
        transition.reference(&registry, engine_component, info_component_id, new_info)?;
    }
    if moved {
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

/// Add a sidebar node for `target` under `sheet`'s node, when the sidebar
/// records the sheet; returns the node and its component.
fn append_sidebar_node(
    working: &mut Working<'_>,
    sheet: u64,
    target: u64,
) -> FaultResult<Option<(u64, String)>> {
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
    )
    .map_err(|error| structure::map_wire(&error))?;
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
