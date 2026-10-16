//! Slide duplication: clone a slide's dedicated component and its node.

use super::graph::{self, MemberKind};
use super::{Error, Lift, Unsupported};
use crate::package::structure::workspace::Workspace;

/// Duplicate the slide of `node` into a new component; returns the new node.
/// The caller inserts the node into the show's slide tree.
pub(super) fn stage(workspace: &mut Workspace<'_>, node: u64) -> Result<u64, Error> {
    let slide = graph::node_slide(workspace, node)?;
    let node_member = workspace.component_of(node).ok_or(Error::InvalidSource)?;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    if member == node_member
        || graph::member_kind(&member) != MemberKind::Other
        || graph::other_slides(workspace, &member, slide)?
    {
        return Err(Error::UnsupportedContent {
            reason: Unsupported::SharedComponent,
        });
    }
    let source = graph::registered(workspace, &member)?.clone();
    graph::check_external_references(workspace, &source)?;

    let new_node = workspace.allocate().lift()?;
    let objects = workspace.member_objects(&member);
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
