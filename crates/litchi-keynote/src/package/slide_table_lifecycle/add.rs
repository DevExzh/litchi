//! A new, empty slide table (the host's `add_slide_table`, record 0423).
//!
//! Builder decks carry a detached creation table: a `TST.TableInfoArchive`
//! whose drawable has no parent. As the host does, the new table is an
//! emptied copy of it (the shared `litchi_iwa_table::structure::create`
//! clone, lifted from Numbers' table add): fresh drawable, model and
//! storage, the scaffold's styles shared, a new table identity. The frame
//! is then set with flags 3 and angle 0, and every row and column sized
//! uniformly to fill it (`litchi_iwa_table::topology::{stage_frame,
//! stage_uniform_dimensions}`, record 0412g). The drawable joins the
//! slide's owned drawables and z-order (in front), the table gets its own
//! CalculationEngine formula owner, and the registry records the clones'
//! UUIDs and every new cross-component reference.
//!
//! A deck without the creation table (native decks) is refused, as the
//! host refuses it; `edit_slide_drawables().add_table` copies a slide table
//! there (record 0329). The shared from-scratch graph is not a fallback yet:
//! native Keynote themes archive their table style network in two
//! messages, which `table_create::fresh::preset_styles` does not read.

use litchi_iwa_table::structure::create::{self, Graph};
use litchi_iwa_table::structure::metadata::{self, Transition};
use litchi_iwa_table::structure::{Fault, FaultResult, Published, Working, references};
use litchi_iwa_table::topology::{self as engine, Frame, Target};

use super::{
    DRAWABLES_Z_ORDER, Error, OWNED_DRAWABLES, Refusal, SLIDE_MESSAGE_TYPE,
    TABLE_INFO_MESSAGE_TYPE, unique_message,
};
use crate::package::Package;
use crate::slide::drawable::Geometry;

/// The geometry flags the host gives a new slide table.
const TABLE_GEOMETRY_FLAGS: u32 = 3;
/// `TST.TableInfoArchive.super` (`TSD.DrawableArchive`) and its `parent`.
const DRAWABLE_PARENT: [u32; 2] = [1, 2];

/// A staged table addition.
#[derive(Debug, Clone)]
pub(super) struct Request {
    /// The slide's position.
    pub(super) slide_position: usize,
    /// The slide.
    pub(super) slide: u64,
    /// The table name.
    pub(super) name: String,
    /// Rows, header rows included.
    pub(super) rows: u32,
    /// Columns, header columns included.
    pub(super) columns: u32,
    /// The frame.
    pub(super) frame: Geometry,
}

/// The deck's one detached creation table (a table drawable with no
/// parent), as the host's `table_template_from_catalog` finds it; several
/// are ambiguous.
fn scaffold(working: &Working<'_, Package>) -> FaultResult<Option<u64>> {
    let mut found = None;
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            if !object
                .messages
                .iter()
                .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            {
                continue;
            }
            let [message] = object.messages.as_slice() else {
                return Err(Fault::Unsupported);
            };
            if !references::identifiers(&message.data, &DRAWABLE_PARENT)?.is_empty() {
                continue;
            }
            let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
            if found.replace(identifier).is_some() {
                return Err(Fault::Unsupported);
            }
        }
    }
    Ok(found)
}

fn layout_error(error: engine::Error) -> Error {
    match error {
        engine::Error::Refused(engine::Refusal::InvalidSize) => {
            Error::Refused(Refusal::InvalidGeometry)
        },
        engine::Error::Refused(_) => Error::Refused(Refusal::UnsupportedSource),
        engine::Error::Fault(fault) => Error::from(fault),
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// Add the requested table; returns the publication and the new drawable.
pub(super) fn add(source: &Package, request: &Request) -> Result<(Published<Package>, u64), Error> {
    let mut working = Working::new(source)?;
    let slide = request.slide;
    let shape = create::Request {
        name: &request.name,
        rows: request.rows,
        columns: request.columns,
    };
    let template = scaffold(&working)?.ok_or(Error::Refused(Refusal::NoTemplate))?;
    let mut graph = create::clone_template(&mut working, shape, template, None, slide, 0.0)?;

    let target = Target {
        info: graph.info,
        model: graph.model,
    };
    let frame = request.frame;
    engine::stage_frame(
        &mut working,
        &target,
        Frame {
            position: Some((frame.x, frame.y)),
            size: Some((frame.width, frame.height)),
            flags: Some(TABLE_GEOMETRY_FLAGS),
            angle: Some(frame.rotation.unwrap_or(0.0)),
        },
    )
    .map_err(layout_error)?;
    engine::stage_uniform_dimensions(&mut working, &target, frame.width, frame.height)
        .map_err(layout_error)?;

    // In front of the slide's drawables.
    let limits = working.archive_limits();
    let index = unique_message(working.object(slide)?, &[SLIDE_MESSAGE_TYPE])?;
    for path in [[OWNED_DRAWABLES], [DRAWABLES_Z_ORDER]] {
        let object = working.object_mut(slide)?;
        let mut after = references::identifiers(&object.messages[index].data, &path)?;
        after
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        after.push(graph.info);
        let data = references::rewrite(&object.messages[index].data, &path, &after)?;
        references::replace(object, index, data, &[], &[graph.info], &path, limits)?;
    }

    let owner = create::add_owner(&mut working, &mut graph)?;
    let transition = registry(&working, &graph, slide, owner.as_ref())?;
    Ok((working.publish(&transition)?, graph.info))
}

/// The registry records of the new graph, in identifier order as the host
/// registers them: UUIDs for the clones of registered objects and for every
/// object built from scratch, and every new cross-component reference.
fn registry(
    working: &Working<'_, Package>,
    graph: &Graph,
    slide: u64,
    owner: Option<&(u64, String)>,
) -> Result<Transition, Error> {
    let mut transition = Transition::default();
    let Some(registry) = metadata::registry(working)? else {
        return Ok(transition);
    };
    let registered = |identifier: u64| {
        registry
            .uuids
            .iter()
            .any(|binding| binding.object == identifier)
    };
    let mut inserted = graph.inserted.clone();
    inserted.sort_by_key(|(identifier, _)| *identifier);
    for (identifier, component) in &inserted {
        let component_id = registry.component_for_entry(component)?.identifier;
        let needs_uuid = match graph.origins.get(identifier) {
            Some(Some(origin)) => registered(*origin),
            Some(None) => true,
            None => false,
        };
        if needs_uuid {
            transition.register_uuid(&registry, component_id, *identifier, graph.seed)?;
        }
        let referenced = working.object(*identifier)?.archive_info.message_infos[0]
            .object_references
            .clone();
        for target in referenced {
            // The host registers no edge from the drawable to its slide.
            if *identifier == graph.info && target == slide {
                continue;
            }
            let target_component = working.locate(target)?.ok_or(Error::InvalidSource)?;
            let target_id = registry.component_for_entry(&target_component)?.identifier;
            transition.reference(&registry, component_id, target_id, target)?;
        }
    }
    let info_component = registry
        .component_for_entry(&graph.info_component)?
        .identifier;
    let slide_member = working.locate(slide)?.ok_or(Error::InvalidSource)?;
    let slide_component = registry.component_for_entry(&slide_member)?.identifier;
    transition.reference(&registry, slide_component, info_component, graph.info)?;
    if let Some((_, component)) = owner {
        let engine_component = registry.component_for_entry(component)?.identifier;
        transition.reference(&registry, engine_component, info_component, graph.info)?;
    }
    Ok(transition)
}
