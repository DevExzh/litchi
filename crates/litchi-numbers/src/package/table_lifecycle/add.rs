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

use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_table::structure::create::{self, Graph};
use litchi_numbers_wire::table_create::fresh;

use super::super::Package;
use super::super::structure::{
    self, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD, FaultResult, Published,
    TREE_CHILDREN_FIELD, TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    metadata::{self, Transition},
    references,
};
use super::{Error, rewrite_list};

/// The most row and column identities one table may need (the native
/// editor's ceiling).
const MAX_TABLE_UIDS: u64 = 1_100_000;
/// How far a new table is offset from a template on the same sheet.
const POSITION_OFFSET: f32 = 40.0;
/// `TN.DocumentArchive.theme`.
const DOCUMENT_THEME_FIELD: u32 = 6;

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

/// The theme's first table preset and the styles its network names (the
/// host's `theme_table_style`).
fn theme_styles(working: &Working<'_>) -> Result<fresh::Styles, Error> {
    let document = working.object(1)?;
    let index = structure::document_message(document)?;
    let theme = references::identifiers(&document.messages[index].data, &[DOCUMENT_THEME_FIELD])?;
    let [theme] = theme.as_slice() else {
        return Err(Error::InvalidSource);
    };
    create::theme_styles(working, *theme)?.ok_or(Error::Refused(
        crate::package::table_lifecycle::Refusal::NoTemplate,
    ))
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
    let shape = create::Request {
        name: request.name,
        rows: request.rows,
        columns: request.columns,
    };
    let mut graph = match template(&working)? {
        Some((template_sheet, template_info)) => create::clone_template(
            &mut working,
            shape,
            template_info,
            Some(template_sheet),
            sheet,
            POSITION_OFFSET,
        )?,
        None => {
            let styles = theme_styles(&working)?;
            let component = working.locate(sheet)?.ok_or(Error::InvalidSource)?;
            create::fresh(&mut working, shape, sheet, &component, &styles)?
        },
    };
    let limits = working.archive_limits();

    // Attach to the sheet's drawables.
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)?;
    let mut after = drawables.clone();
    after.push(graph.info);
    rewrite_list(
        &mut working,
        sheet,
        sheet_index,
        sheet_path,
        &drawables,
        &after,
    )?;

    // A formula owner for the new table, when the workbook has an engine.
    let owner = create::add_owner(&mut working, &mut graph)?;
    let Graph {
        info: new_info,
        info_component,
        seed,
        mut inserted,
        origins,
        ..
    } = graph;
    let new_owner = owner.as_ref().map(|(owner, _)| *owner);
    let owner_component = owner.map(|(_, component)| component);

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
