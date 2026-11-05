//! Slide duplication: clone a slide's dedicated component and its node.

use super::graph::{self, MemberKind};
use super::{Error, Lift, Unsupported};
use crate::package::structure::workspace::Workspace;

/// Duplicate the slide of `node` into a new component; returns the new node.
/// The caller inserts the node into the show's slide tree.
/// The chart drawable, preset and style message types.
const CHART_FIRST_MESSAGE_TYPE: u32 = 5_020;
const CHART_LAST_MESSAGE_TYPE: u32 = 5_031;

pub(super) fn stage(workspace: &mut Workspace<'_>, node: u64) -> Result<u64, Error> {
    let slide = graph::node_slide(workspace, node)?;
    let node_member = workspace.component_of(node).ok_or(Error::InvalidSource)?;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    if member == node_member
        || graph::member_kind(&member) != MemberKind::Other
        || graph::other_slides(workspace, &member, slide)?
    {
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::SharedComponent,
            },
        ));
    }
    let source = graph::registered(workspace, &member)?.clone();
    graph::check_external_references(workspace, &source)?;

    let objects = workspace.member_objects(&member);
    // Charts register their styles in the document stylesheet and the
    // theme; duplicating a slide does not re-register them (record 0372:
    // the chart schema is admitted for the chart lifecycle family only).
    for object in &objects {
        if workspace
            .require(*object)
            .lift()?
            .messages
            .iter()
            .any(|message| {
                (CHART_FIRST_MESSAGE_TYPE..=CHART_LAST_MESSAGE_TYPE).contains(&message.type_)
            })
        {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedContent {
                    reason: Unsupported::UnmodelledContent,
                },
            ));
        }
    }
    let new_node = workspace.allocate().lift()?;
    let remap = graph::allocate_remap(workspace, &objects)?;
    let new_slide = *remap.get(&slide).ok_or(Error::InvalidSource)?;
    let locator = format!("Slide-{new_slide}");
    let new_member = format!("Index/{locator}.iwa");
    workspace.create_component(&new_member).lift()?;
    graph::clone_into(workspace, &new_member, &objects, &remap)?;
    graph::clone_node(workspace, node, new_node, slide, new_slide)?;

    let mut component = source
        .cloned(new_slide, &locator, None, &|object| {
            remap.get(&object).copied()
        })
        .lift()?;
    graph::derive_uuids(workspace, &mut component, source.object_uuids(), &remap)?;
    workspace.registry.add_component(component).lift()?;
    graph::register_node(workspace, node, new_node, new_slide)?;
    Ok(new_node)
}
