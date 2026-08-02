//! Slide-style graph inspection and slide-reference mutation.

use super::*;

pub(super) const SLIDE_MESSAGE_TYPE: u32 = 5;
pub(super) const SLIDE_STYLE_MESSAGE_TYPE: u32 = 9;

pub(super) fn is_collapsible_background_variation(
    style: &kn::SlideStyleArchive,
    raw: &[u8],
) -> Result<bool> {
    let Some(properties) = style.slide_properties.as_ref() else {
        return Ok(false);
    };
    let semantic = style.super_.name.is_none()
        && style.super_.style_identifier.is_none()
        && style.super_.parent.is_some()
        && style.super_.is_variation == Some(true)
        && style.override_count == Some(1)
        && properties.fill.is_some()
        && properties.transition.is_none()
        && properties.transition_null.is_none()
        && properties.title_placeholder_visibility.is_none()
        && properties.body_placeholder_visibility.is_none()
        && properties.object_placeholder_visibility.is_none()
        && properties.slide_number_placeholder_visibility.is_none();
    if !semantic {
        return Ok(false);
    }
    let super_raw = required_length_delimited_payload(raw, 1, "Keynote slide style")?;
    let properties_raw = required_length_delimited_payload(raw, 11, "Keynote slide style")?;
    Ok(has_exact_fields(raw, &[1, 10, 11])?
        && has_exact_fields(super_raw, &[3, 4, 5])?
        && has_exact_fields(properties_raw, &[1])?)
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

pub(super) fn style_is_exclusive(graph: &ObjectGraph, style_id: u64) -> Result<bool> {
    let mut slide_count = 0usize;
    for messages in graph.objects.values() {
        for message in messages {
            match message.type_ {
                SLIDE_MESSAGE_TYPE => {
                    let slide = kn::SlideArchive::decode(message.data.as_slice())?;
                    if slide.style.identifier == style_id {
                        slide_count += 1;
                    }
                },
                SLIDE_STYLE_MESSAGE_TYPE => {
                    let style = kn::SlideStyleArchive::decode(message.data.as_slice())?;
                    if style.super_.parent.as_ref().map(|parent| parent.identifier)
                        == Some(style_id)
                    {
                        return Ok(false);
                    }
                },
                _ => {},
            }
        }
    }
    Ok(slide_count == 1)
}

pub(super) fn patch_slide_style_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive
            .object_mut(slide_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Keynote slide {slide_id} is missing")))?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        }
        let index = indexes[0];
        let message = object.messages[index].clone();
        let data = patch_length_delimited_field(
            &message.data,
            1,
            true,
            Some(&reference(new_style_id).encode_to_vec()),
        )?;
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        if verified.style.identifier != new_style_id {
            return Err(Error::InvalidFormat(
                "Keynote slide style-reference patch failed validation".to_owned(),
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
            &mut object.archive_info.message_infos[index].object_references,
            old_style_id,
            new_style_id,
        )?;
        for field in &mut object.archive_info.message_infos[index].field_infos {
            if field.object_references.contains(&old_style_id) {
                replace_metadata_reference(
                    &mut field.object_references,
                    old_style_id,
                    new_style_id,
                )?;
            }
        }
        Ok(())
    })
}

fn replace_metadata_reference(references: &mut Vec<u64>, old: u64, new: u64) -> Result<()> {
    let count = references
        .iter()
        .filter(|&&candidate| candidate == old)
        .count();
    if count > 1 || references.contains(&new) {
        return Err(Error::InvalidFormat(format!(
            "IWA metadata cannot replace reference {old} with {new} unambiguously"
        )));
    }
    if let Some(reference) = references.iter_mut().find(|candidate| **candidate == old) {
        *reference = new;
    } else {
        references.push(new);
    }
    Ok(())
}

pub(super) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
