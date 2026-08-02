//! Wire-preserving stylesheet style and parent-child registries.

use super::slide_style_graph::reference;
use super::*;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;

pub(super) fn patch_stylesheet(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    remove_style_id: Option<u64>,
    insertion: Option<(u64, u64, ArchiveObject)>,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote stylesheet {stylesheet_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote stylesheet {stylesheet_id} must have exactly one StylesheetArchive payload"
            )));
        }
        let index = indexes[0];
        let message = object.messages[index].clone();
        let insertion_ids = insertion
            .as_ref()
            .map(|(parent_style_id, new_style_id, _)| (*parent_style_id, *new_style_id));
        let data = rewrite_stylesheet_data(&message.data, remove_style_id, insertion_ids)?;
        tss::StylesheetArchive::decode(data.as_slice())?;
        object.replace_message(
            index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[index];
        if let Some(old) = remove_style_id {
            info.object_references
                .retain(|&identifier| identifier != old);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|&identifier| identifier != old);
            }
        }
        if let Some((_, new_style_id, _)) = insertion.as_ref()
            && !info.object_references.contains(new_style_id)
        {
            info.object_references.push(*new_style_id);
        }
        if let Some(old) = remove_style_id {
            archive.remove_object(old).ok_or_else(|| {
                Error::InvalidFormat(format!("disposable Keynote slide style {old} is missing"))
            })?;
        }
        if let Some((_, _, new_style)) = insertion {
            archive.insert_object(new_style)?;
        }
        Ok(())
    })
}

fn rewrite_stylesheet_data(
    data: &[u8],
    remove_style_id: Option<u64>,
    insertion: Option<(u64, u64)>,
) -> Result<Vec<u8>> {
    let mut style_references = repeated_length_delimited_payloads(data, 1)?
        .into_iter()
        .map(|payload| {
            Ok((
                tsp::Reference::decode(payload)?.identifier,
                payload.to_vec(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    if let Some((_, new_style_id)) = insertion
        && style_references.iter().any(|(id, _)| *id == new_style_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote stylesheet already contains style {new_style_id}"
        )));
    }
    if let Some(old) = remove_style_id {
        if style_references.iter().filter(|(id, _)| *id == old).count() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote stylesheet must contain disposable style {old} exactly once"
            )));
        }
        style_references.retain(|(id, _)| *id != old);
    }
    if let Some((_, new_style_id)) = insertion {
        style_references.push((new_style_id, reference(new_style_id).encode_to_vec()));
    }
    let replacements = style_references
        .into_iter()
        .map(|(_, payload)| payload)
        .collect::<Vec<_>>();
    let data = rewrite_repeated_length_delimited_fields(data, 1, &replacements)?;

    let mut child_entries = repeated_length_delimited_payloads(&data, 5)?
        .into_iter()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    let mut removed_count = 0usize;
    let mut parent_entry = None;
    for (index, payload) in child_entries.iter_mut().enumerate() {
        let entry = tss::stylesheet_archive::StyleChildrenEntry::decode(payload.as_slice())?;
        if let Some((_, new_style_id)) = insertion
            && entry
                .children
                .iter()
                .any(|child| child.identifier == new_style_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote stylesheet child map already contains style {new_style_id}"
            )));
        }
        let mut children = repeated_length_delimited_payloads(payload, 2)?
            .into_iter()
            .map(|raw| Ok((tsp::Reference::decode(raw)?.identifier, raw.to_vec())))
            .collect::<Result<Vec<_>>>()?;
        if let Some(old) = remove_style_id {
            if children.iter().any(|(id, _)| *id == old)
                && insertion
                    .is_some_and(|(parent_style_id, _)| entry.parent.identifier != parent_style_id)
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote stylesheet maps disposable style {old} under the wrong parent"
                )));
            }
            let before = children.len();
            children.retain(|(id, _)| *id != old);
            removed_count += before - children.len();
        }
        if let Some((parent_style_id, new_style_id)) = insertion
            && entry.parent.identifier == parent_style_id
        {
            if parent_entry.replace(index).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote stylesheet repeats child-map parent {parent_style_id}"
                )));
            }
            children.push((new_style_id, reference(new_style_id).encode_to_vec()));
        }
        let children = children.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
        *payload = rewrite_repeated_length_delimited_fields(payload, 2, &children)?;
    }
    if remove_style_id.is_some() && removed_count != 1 {
        return Err(Error::InvalidFormat(
            "Keynote stylesheet child map did not contain the disposable style exactly once"
                .to_owned(),
        ));
    }
    let mut nonempty_child_entries = Vec::with_capacity(child_entries.len());
    for payload in child_entries {
        if !repeated_length_delimited_payloads(&payload, 2)?.is_empty() {
            nonempty_child_entries.push(payload);
        }
    }
    let mut child_entries = nonempty_child_entries;
    if let Some((parent_style_id, new_style_id)) = insertion
        && parent_entry.is_none()
    {
        child_entries.push(
            tss::stylesheet_archive::StyleChildrenEntry {
                parent: reference(parent_style_id),
                children: vec![reference(new_style_id)],
            }
            .encode_to_vec(),
        );
    }
    let data = rewrite_repeated_length_delimited_fields(&data, 5, &child_entries)?;
    let verified = tss::StylesheetArchive::decode(data.as_slice())?;
    if let Some((parent_style_id, new_style_id)) = insertion
        && (verified
            .styles
            .iter()
            .filter(|style| style.identifier == new_style_id)
            .count()
            != 1
            || verified
                .parent_to_children_style_map
                .iter()
                .filter(|entry| entry.parent.identifier == parent_style_id)
                .flat_map(|entry| &entry.children)
                .filter(|child| child.identifier == new_style_id)
                .count()
                != 1)
    {
        return Err(Error::InvalidFormat(
            "Keynote stylesheet update failed validation".to_owned(),
        ));
    }
    if let Some(old) = remove_style_id
        && (verified.styles.iter().any(|style| style.identifier == old)
            || verified
                .parent_to_children_style_map
                .iter()
                .flat_map(|entry| &entry.children)
                .any(|child| child.identifier == old))
    {
        return Err(Error::InvalidFormat(
            "Keynote stylesheet retained a removed style".to_owned(),
        ));
    }
    Ok(data)
}
