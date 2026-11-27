//! Graph helpers shared by slide duplication and layout materialization:
//! the show's slide tree, slide-node clones, header-synchronized payload
//! rewrites and the registry records every created slide needs.

use std::collections::HashMap;

use litchi_iwa_core::{ArchiveObject, MessageInfo, RawMessage};

use super::{Error, Lift, Unsupported};
use crate::package::structure::StructureError;
use crate::package::structure::metadata::{Component, ExternalReference};
use crate::package::structure::refs::{self, RefKind};
use crate::package::structure::wire::{self, Fields, LENGTH_DELIMITED};
use crate::package::structure::workspace::Workspace;
use crate::package::{SHOW_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE};

const SHOW_SLIDE_TREE: u32 = 3;
const SHOW_SLIDE_LIST: u32 = 19;
const TREE_ROOT_NODE: u32 = 1;
const TREE_SLIDES: u32 = 2;

const NODE_CHILDREN: u32 = 1;
const NODE_SLIDE: u32 = 2;
const NODE_DATABASE_THUMBNAIL: u32 = 3;
const NODE_DATABASE_THUMBNAILS: u32 = 9;
const NODE_THUMBNAIL_SIZES: u32 = 10;
const NODE_UNIQUE_IDENTIFIER: u32 = 11;
const NODE_COPY_FROM_IDENTIFIER: u32 = 12;
const NODE_HYPERLINK_COUNT: u32 = 13;
const NODE_THUMBNAILS_DIRTY: u32 = 14;
const NODE_THUMBNAILS: u32 = 16;
const NODE_HYPERLINK_MAP: u32 = 24;
const NODE_THUMBNAIL_DIGESTS: u32 = 25;

/// The index of the unique message of `message_type` in `object`.
pub(super) fn message_index(object: &ArchiveObject, message_type: u32) -> Result<usize, Error> {
    let mut found = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type)
        .map(|(index, _)| index);
    let index = found.next().ok_or(Error::InvalidSource)?;
    if found.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(index)
}

/// The object and data references a payload holds, by schema.
fn walk(
    message_type: u32,
    payload: &[u8],
    workspace: &Workspace<'_>,
) -> Result<(Vec<u64>, Vec<u64>), Error> {
    let mut objects = Vec::new();
    let mut data = Vec::new();
    refs::visit(
        message_type,
        payload,
        workspace.wire,
        &mut |kind, identifier| match kind {
            RefKind::Object => objects.push(identifier),
            RefKind::Data => data.push(identifier),
        },
    )
    .map_err(StructureError::from)
    .lift()?;
    Ok((objects, data))
}

/// The object references a payload holds, by schema.
pub(super) fn object_references(
    message_type: u32,
    payload: &[u8],
    workspace: &Workspace<'_>,
) -> Result<Vec<u64>, Error> {
    Ok(walk(message_type, payload, workspace)?.0)
}

/// Every object reference of an object: its walked payload references and
/// its header-listed references.
pub(super) fn all_object_references(
    object: &ArchiveObject,
    workspace: &Workspace<'_>,
) -> Result<Vec<u64>, Error> {
    let mut found = Vec::new();
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        if refs::admitted(message.type_) {
            found.extend(object_references(message.type_, &message.data, workspace)?);
        }
        found.extend_from_slice(&info.object_references);
        for field in &info.field_infos {
            found.extend_from_slice(&field.object_references);
        }
    }
    Ok(found)
}

/// The references walked under each top-level field of a payload that a
/// header attribution names.
fn references_by_field(
    message_type: u32,
    payload: &[u8],
    numbers: impl IntoIterator<Item = u32>,
    workspace: &Workspace<'_>,
) -> Result<HashMap<u32, (Vec<u64>, Vec<u64>)>, Error> {
    let fields = Fields::parse(payload, workspace.wire).lift()?;
    let mut found = HashMap::new();
    for number in numbers {
        if found.contains_key(&number) {
            continue;
        }
        let mut selected = Fields::parse(&[], workspace.wire).lift()?;
        let raw = fields
            .all(number)
            .map(|field| field.payload().to_vec())
            .collect::<Vec<_>>();
        let wire_type = fields
            .all(number)
            .next()
            .map_or(LENGTH_DELIMITED, |field| field.wire_type);
        selected.replace_all(number, wire_type, &raw);
        found.insert(number, walk(message_type, &selected.encode(), workspace)?);
    }
    Ok(found)
}

/// Keep a message header's references in step with a payload rewrite:
/// identifiers that disappeared from the payload leave the aggregate list,
/// identifiers that appeared join it, and every field attribution keeps
/// only the identifiers still found under its top-level field.
fn sync_header(
    info: &mut MessageInfo,
    before: &(Vec<u64>, Vec<u64>),
    after: &(Vec<u64>, Vec<u64>),
    by_field: &HashMap<u32, (Vec<u64>, Vec<u64>)>,
    attribute_like: Option<u64>,
) {
    let added_objects = after
        .0
        .iter()
        .filter(|identifier| !before.0.contains(identifier))
        .copied()
        .collect::<Vec<_>>();
    let added_data = after
        .1
        .iter()
        .filter(|identifier| !before.1.contains(identifier))
        .copied()
        .collect::<Vec<_>>();
    info.object_references
        .retain(|identifier| !before.0.contains(identifier) || after.0.contains(identifier));
    for identifier in &added_objects {
        if !info.object_references.contains(identifier) {
            info.object_references.push(*identifier);
        }
    }
    info.data_references
        .retain(|identifier| !before.1.contains(identifier) || after.1.contains(identifier));
    for identifier in &added_data {
        if !info.data_references.contains(identifier) {
            info.data_references.push(*identifier);
        }
    }
    for field in &mut info.field_infos {
        if let Some((objects, data)) = field
            .path
            .as_slice()
            .first()
            .and_then(|number| by_field.get(number))
        {
            field
                .object_references
                .retain(|identifier| objects.contains(identifier));
            field
                .data_references
                .retain(|identifier| data.contains(identifier));
        }
        if let Some(sibling) = attribute_like
            && field.object_references.contains(&sibling)
        {
            for identifier in &added_objects {
                if !field.object_references.contains(identifier) {
                    field.object_references.push(*identifier);
                }
            }
        }
    }
}

/// Replace the unique `message_type` payload of an owned object, keeping its
/// header references in step.
pub(super) fn rewrite_owned(
    object: &mut ArchiveObject,
    message_type: u32,
    payload: Vec<u8>,
    workspace: &Workspace<'_>,
    attribute_like: Option<u64>,
) -> Result<(), Error> {
    let index = message_index(object, message_type)?;
    let before = walk(message_type, &object.messages[index].data, workspace)?;
    let after = walk(message_type, &payload, workspace)?;
    let by_field = references_by_field(
        message_type,
        &payload,
        object.archive_info.message_infos[index]
            .field_infos
            .iter()
            .filter_map(|field| field.path.as_slice().first().copied()),
        workspace,
    )?;
    object
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            workspace.archive_limits(),
        )
        .map_err(StructureError::from)
        .lift()?;
    sync_header(
        &mut object.archive_info.message_infos[index],
        &before,
        &after,
        &by_field,
        attribute_like,
    );
    Ok(())
}

/// Replace the unique `message_type` payload of a staged object, keeping its
/// header references in step.
pub(super) fn rewrite(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    attribute_like: Option<u64>,
) -> Result<(), Error> {
    let mut object = workspace.require(identifier).lift()?.clone();
    rewrite_owned(
        &mut object,
        message_type,
        payload,
        workspace,
        attribute_like,
    )?;
    *workspace.object_mut(identifier).lift()? = object;
    Ok(())
}

/// The unique payload of `message_type` held by a staged object.
pub(super) fn payload<'w>(
    workspace: &'w Workspace<'_>,
    identifier: u64,
    message_type: u32,
) -> Result<&'w [u8], Error> {
    workspace.payload(identifier, message_type).lift()
}

/// The identifier of a singular reference field, if present.
pub(super) fn reference_field(fields: &Fields, number: u32) -> Result<Option<u64>, Error> {
    let Some(field) = fields.unique(number).lift()? else {
        return Ok(None);
    };
    let payload = field.message().lift()?;
    let identifier =
        wire::reference_identifier(payload, litchi_iwa_common::WireLimits::default()).lift()?;
    Ok(Some(identifier))
}

/// The identifiers of a repeated reference field, in order.
pub(super) fn repeated_references(fields: &Fields, number: u32) -> Result<Vec<u64>, Error> {
    fields
        .all(number)
        .map(|field| {
            wire::reference_identifier(
                field.message().lift()?,
                litchi_iwa_common::WireLimits::default(),
            )
            .lift()
        })
        .collect()
}

/// The show object and its ordered slide tree.
pub(super) struct ShowTree {
    identifier: u64,
    order: Vec<u64>,
}

impl ShowTree {
    /// Read the flat slide tree, refusing nested and list topologies.
    pub(super) fn read(workspace: &Workspace<'_>) -> Result<Self, Error> {
        let identifier = workspace
            .package()
            .root_show_identifier()
            .map_err(|_error| Error::InvalidSource)?;
        let show = Fields::parse(
            payload(workspace, identifier, SHOW_MESSAGE_TYPE)?,
            workspace.wire,
        )
        .lift()?;
        if show.unique(SHOW_SLIDE_LIST).lift()?.is_some() {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedTopology,
            ));
        }
        let tree = show
            .unique(SHOW_SLIDE_TREE)
            .lift()?
            .ok_or(Error::InvalidSource)?;
        let tree = Fields::parse(tree.message().lift()?, workspace.wire).lift()?;
        if tree.unique(TREE_ROOT_NODE).lift()?.is_some() {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedTopology,
            ));
        }
        let order = repeated_references(&tree, TREE_SLIDES)?;
        let mut sorted = order.clone();
        sorted.sort_unstable();
        if sorted.first() == Some(&0) || sorted.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedTopology,
            ));
        }
        Ok(Self { identifier, order })
    }

    /// The base slide-node order.
    pub(super) fn order(&self) -> &[u64] {
        &self.order
    }

    /// Write the final slide-node order into the show.
    pub(super) fn write(&self, workspace: &mut Workspace<'_>, order: &[u64]) -> Result<(), Error> {
        let mut show = Fields::parse(
            payload(workspace, self.identifier, SHOW_MESSAGE_TYPE)?,
            workspace.wire,
        )
        .lift()?;
        let tree = show
            .unique(SHOW_SLIDE_TREE)
            .lift()?
            .ok_or(Error::InvalidSource)?;
        let mut tree = Fields::parse(tree.message().lift()?, workspace.wire).lift()?;
        let references = order
            .iter()
            .map(|identifier| wire::reference(*identifier))
            .collect::<Vec<_>>();
        tree.replace_all(TREE_SLIDES, LENGTH_DELIMITED, &references);
        show.set_message(SHOW_SLIDE_TREE, &tree.encode());
        rewrite(
            workspace,
            self.identifier,
            SHOW_MESSAGE_TYPE,
            show.encode(),
            self.order.first().copied(),
        )
    }
}

/// The slide referenced by a slide node, refusing nodes with children.
pub(super) fn node_slide(workspace: &Workspace<'_>, node: u64) -> Result<u64, Error> {
    let fields = Fields::parse(
        payload(workspace, node, SLIDE_NODE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    if fields.all(NODE_CHILDREN).next().is_some() {
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedContent {
                reason: Unsupported::SharedComponent,
            },
        ));
    }
    reference_field(&fields, NODE_SLIDE)?.ok_or(Error::InvalidSource)
}

/// Clone a slide node for a new slide: drop children, thumbnails, hyperlink
/// maps and unique identifiers, mark thumbnails dirty, and point it at the
/// new slide. The clone joins the source node's component.
pub(super) fn clone_node(
    workspace: &mut Workspace<'_>,
    source_node: u64,
    new_node: u64,
    old_slide: u64,
    new_slide: u64,
) -> Result<(), Error> {
    let member = workspace
        .component_of(source_node)
        .ok_or(Error::InvalidSource)?;
    let source = workspace.require(source_node).lift()?.clone();
    if source.messages.len() != 1 {
        return Err(Error::InvalidSource);
    }
    let remap = HashMap::from([(old_slide, new_slide)]);
    let mut node = workspace.clone_object(&source, new_node, &remap).lift()?;
    let index = message_index(&node, SLIDE_NODE_MESSAGE_TYPE)?;
    let mut fields = Fields::parse(&node.messages[index].data, workspace.wire).lift()?;
    for number in [
        NODE_CHILDREN,
        NODE_DATABASE_THUMBNAIL,
        NODE_DATABASE_THUMBNAILS,
        NODE_THUMBNAIL_SIZES,
        NODE_UNIQUE_IDENTIFIER,
        NODE_COPY_FROM_IDENTIFIER,
        NODE_THUMBNAILS,
        NODE_HYPERLINK_MAP,
        NODE_THUMBNAIL_DIGESTS,
    ] {
        fields.remove(number);
    }
    if fields.unique(NODE_HYPERLINK_COUNT).lift()?.is_some() {
        fields.set_varint(NODE_HYPERLINK_COUNT, 0);
    }
    fields.set_varint(NODE_THUMBNAILS_DIRTY, 1);
    if reference_field(&fields, NODE_SLIDE)? != Some(new_slide) {
        return Err(Error::InvalidSource);
    }
    rewrite_owned(
        &mut node,
        SLIDE_NODE_MESSAGE_TYPE,
        fields.encode(),
        workspace,
        None,
    )?;
    workspace.insert(&member, node).lift()
}

/// Which registry component a package member is, for external-reference
/// admission.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum MemberKind {
    Document,
    Stylesheet,
    TemplateSlide,
    Other,
}

pub(super) fn member_kind(member: &str) -> MemberKind {
    match member {
        "Index/Document.iwa" => MemberKind::Document,
        "Index/DocumentStylesheet.iwa" => MemberKind::Stylesheet,
        _ if member.starts_with("Index/TemplateSlide") && member.ends_with(".iwa") => {
            MemberKind::TemplateSlide
        },
        _ => MemberKind::Other,
    }
}

/// The registry component of a package member.
pub(super) fn registered<'w>(
    workspace: &'w Workspace<'_>,
    member: &str,
) -> Result<&'w Component, Error> {
    workspace
        .registry
        .component_for_member(member)
        .ok_or(Error::InvalidSource)
}

/// Refuse a component whose external references reach components other than
/// the document, the stylesheet and template slides.
pub(super) fn check_external_references(
    workspace: &Workspace<'_>,
    component: &Component,
) -> Result<(), Error> {
    for reference in component.external_references() {
        let target = workspace
            .registry
            .component(reference.component)
            .ok_or(Error::InvalidSource)?;
        if member_kind(&target.member_name()) == MemberKind::Other {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedContent {
                    reason: Unsupported::ExternalContent,
                },
            ));
        }
    }
    Ok(())
}

/// Register a cloned component's object UUIDs, deriving one fresh UUID per
/// source mapping that `remap` carries over.
pub(super) fn derive_uuids(
    workspace: &mut Workspace<'_>,
    component: &mut Component,
    source: &[(u64, crate::package::structure::metadata::Uuid)],
    remap: &HashMap<u64, u64>,
) -> Result<(), Error> {
    for (object, uuid) in source {
        if let Some(target) = remap.get(object) {
            let fresh = workspace.clone_uuid(Some(*uuid), *target).lift()?;
            component.add_object_uuid(*target, fresh);
        }
    }
    Ok(())
}

/// Register a new slide node in the document component and link the
/// document to the new slide component.
pub(super) fn register_node(
    workspace: &mut Workspace<'_>,
    source_node: u64,
    new_node: u64,
    new_component: u64,
) -> Result<(), Error> {
    let member = workspace
        .component_of(new_node)
        .ok_or(Error::InvalidSource)?;
    let (document, source_uuid) = {
        let component = registered(workspace, &member)?;
        (component.identifier, component.uuid_of(source_node))
    };
    let uuid = workspace.clone_uuid(source_uuid, new_node).lift()?;
    let component = workspace
        .registry
        .component_mut(document)
        .ok_or(Error::InvalidSource)?;
    component.add_object_uuid(new_node, uuid);
    component.add_external_reference(ExternalReference {
        component: new_component,
        object: None,
        is_weak: None,
    });
    Ok(())
}

/// Allocate one identifier per source object, in order.
pub(super) fn allocate_remap(
    workspace: &mut Workspace<'_>,
    objects: &[u64],
) -> Result<HashMap<u64, u64>, Error> {
    let mut remap = HashMap::with_capacity(objects.len());
    for object in objects {
        let target = workspace.allocate().lift()?;
        if remap.insert(*object, target).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(remap)
}

/// Clone `objects` (all in staged components) into `member` under `remap`.
pub(super) fn clone_into(
    workspace: &mut Workspace<'_>,
    member: &str,
    objects: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<(), Error> {
    for object in objects {
        let source = workspace.require(*object).lift()?.clone();
        let target = *remap.get(object).ok_or(Error::InvalidSource)?;
        let clone = workspace.clone_object(&source, target, remap).lift()?;
        workspace.insert(member, clone).lift()?;
    }
    Ok(())
}

/// Every slide object in a member besides `slide`.
pub(super) fn other_slides(
    workspace: &Workspace<'_>,
    member: &str,
    slide: u64,
) -> Result<bool, Error> {
    for object in workspace.member_objects(member) {
        if object != slide
            && workspace
                .require(object)
                .lift()?
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        {
            return Ok(true);
        }
    }
    Ok(false)
}
