//! The document's sidebar order (`TN.DocumentArchive.sidebar_order`): one
//! `TSK.TreeNode` per sheet under the root, one per table under its sheet's
//! node, as the native app and this crate's builder record them.
//!
//! Host defect fixed with record 0370f: `add_empty_table` and
//! `add_empty_sheet` added no node, so a new sheet or table was missing from
//! the sidebar.

use super::*;
use crate::protobuf::tsk;

const TREE_NODE_MESSAGE_TYPE: u32 = 205;

/// Append a sidebar node addressing `target` under the node of `parent`
/// (a sheet), or under the root when `parent` is `None`. A document without
/// a sidebar, or a sheet without a node, takes none.
pub(super) fn append_node(
    package: &mut HostPackage,
    parent: Option<u64>,
    target: u64,
) -> Result<()> {
    insert_node(package, parent, target, None, None)
}

/// Add a sidebar node for `target` under `parent`'s node (the root when
/// `None`): right after the child addressing `after` when given, else
/// last. `existing` reuses a detached node instead of creating one.
pub(super) fn insert_node(
    package: &mut HostPackage,
    parent: Option<u64>,
    target: u64,
    after: Option<u64>,
    existing: Option<u64>,
) -> Result<()> {
    let document = numbers_document(package)?;
    let root = document.sidebar_order.identifier;
    if root == 0 {
        return Ok(());
    }
    let locations = object_locations(package)?;
    let node_of = |package: &HostPackage, identifier: u64| -> Result<Option<tsk::TreeNode>> {
        let Some(name) = locations.get(&identifier) else {
            return Ok(None);
        };
        Ok(package
            .archive(name)?
            .object(identifier)
            .and_then(|object| object.messages.first())
            .filter(|message| message.type_ == TREE_NODE_MESSAGE_TYPE)
            .and_then(|message| tsk::TreeNode::decode(message.data.as_slice()).ok()))
    };
    let Some(root_node) = node_of(package, root)? else {
        return Ok(());
    };
    let parent_id = match parent {
        None => root,
        Some(sheet) => {
            let mut found = None;
            for child in &root_node.children {
                if node_of(package, child.identifier)?
                    .and_then(|node| node.object)
                    .map(|object| object.identifier)
                    == Some(sheet)
                {
                    found = Some(child.identifier);
                }
            }
            let Some(found) = found else {
                return Ok(());
            };
            found
        },
    };
    let parent_archive = locations
        .get(&parent_id)
        .cloned()
        .ok_or_else(|| Error::InvalidFormat("Numbers sidebar node is missing".to_owned()))?;
    // The child to insert after, by the object it addresses.
    let after_node = match after {
        Some(after) => {
            let parent_node = node_of(package, parent_id)?.ok_or_else(|| {
                Error::InvalidFormat("Numbers sidebar node is missing".to_owned())
            })?;
            let mut found = None;
            for child in &parent_node.children {
                if node_of(package, child.identifier)?
                    .and_then(|node| node.object)
                    .map(|object| object.identifier)
                    == Some(after)
                {
                    found = Some(child.identifier);
                }
            }
            found
        },
        None => None,
    };
    let node_id = match existing {
        Some(node) => node,
        None => next_object_identifier(package)?,
    };
    let mut data = Vec::new();
    crate::wire::append_length_delimited_field(
        &mut data,
        3,
        &tsp::Reference {
            identifier: target,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    package.update_archive(&parent_archive, |archive| {
        let parent_object = archive
            .object(parent_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers sidebar node is missing".to_owned()))?;
        if existing.is_none() {
            let mut node = ArchiveObject::new(
                node_id,
                vec![RawMessage {
                    type_: TREE_NODE_MESSAGE_TYPE,
                    data,
                }],
                Default::default(),
            )?;
            node.archive_info.should_merge = parent_object.archive_info.should_merge;
            if let Some(info) = parent_object.archive_info.message_infos.first() {
                node.archive_info.message_infos[0].versions = info.versions.clone();
            }
            node.archive_info.message_infos[0].object_references = vec![target];
            archive.insert_object(node, Default::default())?;
        }
        let parent_object = archive
            .object_mut(parent_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers sidebar node is missing".to_owned()))?;
        let original = parent_object.messages[0].data.clone();
        let previous = tsk::TreeNode::decode(original.as_slice())?
            .children
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let mut current = previous.clone();
        match after_node.and_then(|after| current.iter().position(|child| *child == after)) {
            Some(at) => current.insert(at + 1, node_id),
            None => current.push(node_id),
        }
        let rewritten = model::rewrite_reference_list(&original, 2, &previous, &current)?;
        parent_object.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TREE_NODE_MESSAGE_TYPE,
                data: rewritten,
            },
            Default::default(),
        )?;
        let info = &mut parent_object.archive_info.message_infos[0];
        if !info.object_references.contains(&node_id) {
            info.object_references.push(node_id);
        }
        for field in &mut info.field_infos {
            if field.path.path.as_slice() == [2] && !field.object_references.contains(&node_id) {
                field.object_references.push(node_id);
            }
        }
        Ok(())
    })?;
    if existing.is_none() {
        set_package_last_object_identifier(package, node_id)?;
    }
    if let Some(target_archive) = locations.get(&target) {
        table_duplicate::register_numbers_component_reference(
            package,
            &parent_archive,
            target_archive,
            target,
        )?;
    }
    Ok(())
}

/// Remove the sidebar node addressing `target` (a table under its sheet's
/// node, or a sheet under the root); returns the removed node.
pub(super) fn remove_node(package: &mut HostPackage, target: u64) -> Result<Option<u64>> {
    let locations = object_locations(package)?;
    let node = detach_node(package, target)?;
    if let Some(node) = node {
        remove_object_or_empty_entry(package, &locations, node)?;
    }
    Ok(node)
}

/// Move the sidebar node addressing table `target` under `sheet`'s node, as
/// a table moved to another sheet is listed (record 0370j).
pub(super) fn move_node(package: &mut HostPackage, target: u64, sheet: u64) -> Result<()> {
    if let Some(node) = detach_node(package, target)? {
        insert_node(package, Some(sheet), target, None, Some(node))?;
    }
    Ok(())
}

/// Detach the sidebar node addressing `target` from its parent's children;
/// returns it.
fn detach_node(package: &mut HostPackage, target: u64) -> Result<Option<u64>> {
    let document = numbers_document(package)?;
    let root = document.sidebar_order.identifier;
    if root == 0 {
        return Ok(None);
    }
    let locations = object_locations(package)?;
    let node_of = |package: &HostPackage, identifier: u64| -> Result<Option<tsk::TreeNode>> {
        let Some(name) = locations.get(&identifier) else {
            return Ok(None);
        };
        Ok(package
            .archive(name)?
            .object(identifier)
            .and_then(|object| object.messages.first())
            .filter(|message| message.type_ == TREE_NODE_MESSAGE_TYPE)
            .and_then(|message| tsk::TreeNode::decode(message.data.as_slice()).ok()))
    };
    // (parent, node) of the node addressing `target`.
    let mut found = None;
    let mut pending = vec![root];
    while let Some(parent) = pending.pop() {
        let Some(node) = node_of(package, parent)? else {
            continue;
        };
        for child in &node.children {
            let Some(child_node) = node_of(package, child.identifier)? else {
                continue;
            };
            if child_node.object.as_ref().map(|object| object.identifier) == Some(target) {
                found = Some((parent, child.identifier));
            }
            pending.push(child.identifier);
        }
    }
    let Some((parent, node)) = found else {
        return Ok(None);
    };
    let parent_archive = locations
        .get(&parent)
        .cloned()
        .ok_or_else(|| Error::InvalidFormat("Numbers sidebar node is missing".to_owned()))?;
    package.update_archive(&parent_archive, |archive| {
        let parent_object = archive
            .object_mut(parent)
            .ok_or_else(|| Error::InvalidFormat("Numbers sidebar node is missing".to_owned()))?;
        let original = parent_object.messages[0].data.clone();
        let previous = tsk::TreeNode::decode(original.as_slice())?
            .children
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let current = previous
            .iter()
            .copied()
            .filter(|identifier| *identifier != node)
            .collect::<Vec<_>>();
        let rewritten = model::rewrite_reference_list(&original, 2, &previous, &current)?;
        parent_object.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TREE_NODE_MESSAGE_TYPE,
                data: rewritten,
            },
            Default::default(),
        )?;
        let info = &mut parent_object.archive_info.message_infos[0];
        info.object_references
            .retain(|identifier| *identifier != node);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| *identifier != node);
        }
        Ok(())
    })?;
    Ok(Some(node))
}
