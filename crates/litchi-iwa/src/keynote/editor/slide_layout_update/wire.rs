//! Lossless wire and metadata patches for Keynote layout reassignment.

use super::*;

pub(super) fn patch_placeholder_presentation(
    package: &mut IWorkPackage,
    graph: &ObjectGraph,
    current_id: u64,
    target_id: u64,
    label: &str,
) -> Result<()> {
    let current_raw = graph.message_data_type(
        current_id,
        PLACEHOLDER_MESSAGE_TYPE,
        "KN.PlaceholderArchive",
    )?;
    let target_raw =
        graph.message_data_type(target_id, PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")?;
    let current = kn::PlaceholderArchive::decode(current_raw)?;
    let target = kn::PlaceholderArchive::decode(target_raw)?;
    let old_style = current
        .super_
        .super_
        .style
        .as_ref()
        .map(|style| style.identifier);
    let new_style = target
        .super_
        .super_
        .style
        .as_ref()
        .map(|style| style.identifier);
    let geometry = nested_optional_payload(target_raw, PLACEHOLDER_GEOMETRY_PATH)?;
    let style = nested_optional_payload(target_raw, PLACEHOLDER_STYLE_PATH)?;
    let path_source = nested_optional_payload(target_raw, PLACEHOLDER_PATH_SOURCE_PATH)?;
    let archive_name = graph.archive_name(current_id)?.to_owned();
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(current_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote {label} placeholder {current_id} is missing"
            ))
        })?;
        let index = unique_message_index(object, PLACEHOLDER_MESSAGE_TYPE, "placeholder")?;
        let original = object.messages[index].data.as_slice();
        let mut data = patch_nested_length_delimited_field(
            original,
            PLACEHOLDER_GEOMETRY_PATH,
            current.super_.super_.super_.geometry.is_some(),
            geometry,
        )?;
        data = patch_nested_length_delimited_field(
            &data,
            PLACEHOLDER_STYLE_PATH,
            current.super_.super_.style.is_some(),
            style,
        )?;
        data = patch_nested_length_delimited_field(
            &data,
            PLACEHOLDER_PATH_SOURCE_PATH,
            current.super_.super_.pathsource.is_some(),
            path_source,
        )?;
        let mut expected = current.clone();
        expected.super_.super_.super_.geometry = target.super_.super_.super_.geometry;
        expected.super_.super_.style = target.super_.super_.style;
        expected.super_.super_.pathsource = target.super_.super_.pathsource;
        if kn::PlaceholderArchive::decode(data.as_slice())? != expected {
            return Err(Error::InvalidFormat(format!(
                "Keynote {label} placeholder presentation patch failed validation"
            )));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data,
            },
        )?;
        replace_metadata_reference(object, index, old_style, new_style, label)
    })
}

pub(super) fn patch_slide_relationship(
    package: &mut IWorkPackage,
    graph: &ObjectGraph,
    slide_id: u64,
    current: &kn::SlideArchive,
    target: &slide_create::layout::ResolvedLayout,
) -> Result<()> {
    let target_raw =
        graph.message_data_type(target.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    let target_style = required_length_delimited_payload(target_raw, SLIDE_STYLE_FIELD, "layout")?;
    let target_template = tsp::Reference {
        identifier: target.slide_id,
        ..Default::default()
    }
    .encode_to_vec();
    let old_style = current.style.identifier;
    let old_template = current.template_slide.as_ref().map(|item| item.identifier);
    let archive_name = graph.archive_name(slide_id)?.to_owned();
    package.update_archive(&archive_name, |archive| {
        let object = archive
            .object_mut(slide_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Keynote slide {slide_id} is missing")))?;
        let index = unique_message_index(object, SLIDE_MESSAGE_TYPE, "slide")?;
        let original = object.messages[index].data.as_slice();
        let mut data =
            patch_length_delimited_field(original, SLIDE_STYLE_FIELD, true, Some(target_style))?;
        data = patch_length_delimited_field(
            &data,
            SLIDE_TEMPLATE_FIELD,
            current.template_slide.is_some(),
            Some(&target_template),
        )?;
        let mut expected = current.clone();
        expected.style = target.slide.style;
        expected.template_slide = Some(tsp::Reference::decode(target_template.as_slice())?);
        if kn::SlideArchive::decode(data.as_slice())? != expected {
            return Err(Error::InvalidFormat(
                "Keynote slide layout relationship patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        replace_metadata_reference(
            object,
            index,
            Some(old_style),
            Some(target.slide.style.identifier),
            "slide style",
        )?;
        replace_metadata_reference(
            object,
            index,
            old_template,
            Some(target.slide_id),
            "slide template",
        )
    })
}

pub(super) fn patch_node_template_uuid(
    package: &mut IWorkPackage,
    graph: &ObjectGraph,
    node_id: u64,
    current: &kn::SlideNodeArchive,
    target: &kn::SlideNodeArchive,
) -> Result<()> {
    let target_uuid = target
        .template_slide_id
        .as_ref()
        .map(Message::encode_to_vec);
    let archive_name = graph.archive_name(node_id)?.to_owned();
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
        })?;
        let index = unique_message_index(object, SLIDE_NODE_MESSAGE_TYPE, "slide node")?;
        let original = object.messages[index].data.as_slice();
        let data = patch_length_delimited_field(
            original,
            SLIDE_NODE_TEMPLATE_UUID_FIELD,
            current.template_slide_id.is_some(),
            target_uuid.as_deref(),
        )?;
        let mut expected = current.clone();
        expected.template_slide_id = target.template_slide_id;
        if kn::SlideNodeArchive::decode(data.as_slice())? != expected {
            return Err(Error::InvalidFormat(
                "Keynote slide-node template UUID patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn nested_optional_payload<'a>(data: &'a [u8], path: &[u32]) -> Result<Option<&'a [u8]>> {
    let (&field, rest) = path.split_first().ok_or_else(|| {
        Error::InvalidFormat("protobuf nested field path cannot be empty".to_owned())
    })?;
    if rest.is_empty() {
        return optional_length_delimited_payload(data, field);
    }
    let parent = required_length_delimited_payload(data, field, "nested protobuf message")?;
    nested_optional_payload(parent, rest)
}

fn replace_metadata_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    old: Option<u64>,
    new: Option<u64>,
    label: &str,
) -> Result<()> {
    if old == new {
        return Ok(());
    }
    let info = &mut object.archive_info.message_infos[message_index];
    replace_reference_list(&mut info.object_references, old, new, label)?;
    for field in &mut info.field_infos {
        if old.is_some_and(|identifier| field.object_references.contains(&identifier)) {
            replace_reference_list(&mut field.object_references, old, new, label)?;
        }
    }
    Ok(())
}

fn replace_reference_list(
    references: &mut Vec<u64>,
    old: Option<u64>,
    new: Option<u64>,
    label: &str,
) -> Result<()> {
    let old_count = old.map_or(0, |identifier| {
        references
            .iter()
            .filter(|&&candidate| candidate == identifier)
            .count()
    });
    if old_count > 1 {
        return Err(Error::InvalidFormat(format!(
            "Keynote {label} metadata repeats object reference {}",
            old.unwrap_or_default()
        )));
    }
    match (old, new) {
        (Some(old), Some(new)) if old_count == 1 => {
            let index = references
                .iter()
                .position(|&identifier| identifier == old)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote {label} metadata reference disappeared during update"
                    ))
                })?;
            if references.contains(&new) {
                references.remove(index);
            } else {
                references[index] = new;
            }
        },
        (Some(old), None) => references.retain(|&identifier| identifier != old),
        (_, Some(new)) if !references.contains(&new) => references.push(new),
        _ => {},
    }
    Ok(())
}

fn unique_message_index(object: &ArchiveObject, message_type: u32, label: &str) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote {label} must contain exactly one message type {message_type} payload"
        )));
    };
    Ok(*index)
}
