//! Wire-preserving Keynote slide-style variation lifecycle.

use super::slide_background::{KeynoteRgbColorSpace, KeynoteSlideBackground};
use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_STYLE_MESSAGE_TYPE: u32 = 9;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;

pub(super) fn set_slide_background(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    background: KeynoteSlideBackground,
    inherited_fill_payload: &[u8],
) -> Result<()> {
    let slides = editor.slides()?;
    let slide_info = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let slide: kn::SlideArchive =
        graph.decode_type(slide_info.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    let old_style_id = slide.style.identifier;
    let old_style: kn::SlideStyleArchive = graph.decode_type(
        old_style_id,
        SLIDE_STYLE_MESSAGE_TYPE,
        "KN.SlideStyleArchive",
    )?;
    let stylesheet_id = old_style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide style {old_style_id} has no stylesheet reference"
            ))
        })?;
    let slide_archive = graph.archive_name(slide_info.slide_id)?.to_owned();
    let stylesheet_archive = graph.archive_name(stylesheet_id)?.to_owned();
    if graph.archive_name(old_style_id)? != stylesheet_archive {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide style {old_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let disposable = is_disposable_background_variation(&old_style)
        && count_slide_style_references(&graph, old_style_id)? == 1;
    let parent_style_id = if disposable {
        old_style.super_.parent.as_ref().unwrap().identifier
    } else {
        old_style_id
    };
    let new_style_id = next_object_identifier(editor.package())?;
    let fill_payload = encode_background_fill(&background, inherited_fill_payload)?;
    let new_style = new_style_object(new_style_id, parent_style_id, stylesheet_id, &fill_payload)?;

    let mut staged = editor.package().clone();
    patch_slide_style_reference(
        &mut staged,
        &slide_archive,
        slide_info.slide_id,
        old_style_id,
        new_style_id,
    )?;
    patch_stylesheet(
        &mut staged,
        &stylesheet_archive,
        stylesheet_id,
        disposable.then_some(old_style_id),
        parent_style_id,
        new_style_id,
        new_style,
    )?;
    update_package_metadata(
        &mut staged,
        &slide_archive,
        &stylesheet_archive,
        disposable.then_some(old_style_id),
        new_style_id,
    )?;
    set_package_last_object_identifier(&mut staged, new_style_id)?;

    let bytes = staged.to_bytes()?;
    let verified = KeynoteEditor::from_bytes(&bytes)?;
    if verified.slide_background(slide_index)? != background {
        return Err(Error::InvalidFormat(
            "Keynote slide-background update failed validation".to_owned(),
        ));
    }
    editor.text = IWorkTextEditor::from_package(staged);
    Ok(())
}

fn is_disposable_background_variation(style: &kn::SlideStyleArchive) -> bool {
    let Some(properties) = style.slide_properties.as_ref() else {
        return false;
    };
    style.super_.name.is_none()
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
        && properties.slide_number_placeholder_visibility.is_none()
}

fn count_slide_style_references(graph: &ObjectGraph, style_id: u64) -> Result<usize> {
    let mut count = 0usize;
    for messages in graph.objects.values() {
        for message in messages
            .iter()
            .filter(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        {
            let slide = kn::SlideArchive::decode(message.data.as_slice())?;
            if slide.style.identifier == style_id {
                count += 1;
            }
        }
    }
    Ok(count)
}

fn encode_background_fill(
    background: &KeynoteSlideBackground,
    inherited_fill_payload: &[u8],
) -> Result<Vec<u8>> {
    match background {
        KeynoteSlideBackground::None => Ok(tsd::FillArchive::default().encode_to_vec()),
        KeynoteSlideBackground::Opaque(payload) => {
            tsd::FillArchive::decode(payload.as_slice())?;
            Ok(payload.clone())
        },
        KeynoteSlideBackground::Solid(color) => {
            let existing = tsd::FillArchive::decode(inherited_fill_payload)?;
            let mut data = if existing.gradient.is_none() && existing.image.is_none() {
                if existing.color.is_some() {
                    inherited_fill_payload.to_vec()
                } else {
                    patch_length_delimited_field(
                        inherited_fill_payload,
                        1,
                        false,
                        Some(
                            &tsp::Color {
                                model: tsp::color::ColorModel::Rgb as i32,
                                ..Default::default()
                            }
                            .encode_to_vec(),
                        ),
                    )?
                }
            } else {
                tsd::FillArchive {
                    color: Some(tsp::Color {
                        model: tsp::color::ColorModel::Rgb as i32,
                        r: Some(color.red),
                        g: Some(color.green),
                        b: Some(color.blue),
                        rgbspace: Some(native_color_space(color.color_space)),
                        a: Some(color.alpha),
                        ..Default::default()
                    }),
                    ..Default::default()
                }
                .encode_to_vec()
            };
            data = patch_nested_varint_field(
                &data,
                &[1, 1],
                true,
                Some(tsp::color::ColorModel::Rgb as u64),
            )?;
            for (field, value) in [
                (3, color.red),
                (4, color.green),
                (5, color.blue),
                (6, color.alpha),
            ] {
                data = patch_nested_fixed32_field(&data, &[1, field], true, Some(value.to_bits()))?;
            }
            data = patch_nested_varint_field(
                &data,
                &[1, 12],
                true,
                Some(native_color_space(color.color_space) as u64),
            )?;
            let verified = tsd::FillArchive::decode(data.as_slice())?;
            if verified.gradient.is_some() || verified.image.is_some() {
                return Err(Error::InvalidFormat(
                    "Keynote solid background retained an incompatible fill".to_owned(),
                ));
            }
            Ok(data)
        },
    }
}

fn native_color_space(color_space: KeynoteRgbColorSpace) -> i32 {
    match color_space {
        KeynoteRgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
        KeynoteRgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
    }
}

fn new_style_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    fill_payload: &[u8],
) -> Result<ArchiveObject> {
    let style = kn::SlideStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        slide_properties: Some(kn::SlideStylePropertiesArchive {
            fill: Some(tsd::FillArchive::default()),
            ..Default::default()
        }),
    };
    let data = patch_nested_length_delimited_field(
        &style.encode_to_vec(),
        &[11, 1],
        true,
        Some(fill_payload),
    )?;
    kn::SlideStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SLIDE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    Ok(object)
}

fn patch_slide_style_reference(
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

fn patch_stylesheet(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    remove_style_id: Option<u64>,
    parent_style_id: u64,
    new_style_id: u64,
    new_style: ArchiveObject,
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
        let data = rewrite_stylesheet_data(
            &message.data,
            remove_style_id,
            parent_style_id,
            new_style_id,
        )?;
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
        if !info.object_references.contains(&new_style_id) {
            info.object_references.push(new_style_id);
        }
        if let Some(old) = remove_style_id {
            archive.remove_object(old).ok_or_else(|| {
                Error::InvalidFormat(format!("disposable Keynote slide style {old} is missing"))
            })?;
        }
        archive.insert_object(new_style)?;
        Ok(())
    })
}

fn rewrite_stylesheet_data(
    data: &[u8],
    remove_style_id: Option<u64>,
    parent_style_id: u64,
    new_style_id: u64,
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
    if style_references.iter().any(|(id, _)| *id == new_style_id) {
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
    style_references.push((new_style_id, reference(new_style_id).encode_to_vec()));
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
        if entry
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
                && entry.parent.identifier != parent_style_id
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote stylesheet maps disposable style {old} under the wrong parent"
                )));
            }
            let before = children.len();
            children.retain(|(id, _)| *id != old);
            removed_count += before - children.len();
        }
        if entry.parent.identifier == parent_style_id {
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
    child_entries.retain(|payload| {
        repeated_length_delimited_payloads(payload, 2)
            .map(|children| !children.is_empty())
            .unwrap_or(true)
    });
    if parent_entry.is_none() {
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
    if verified
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
            != 1
    {
        return Err(Error::InvalidFormat(
            "Keynote stylesheet update failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn update_package_metadata(
    package: &mut IWorkPackage,
    slide_archive: &str,
    stylesheet_archive: &str,
    remove_style_id: Option<u64>,
    new_style_id: u64,
) -> Result<()> {
    let slide_component = component_identifier_for_entry(package, slide_archive)?;
    let stylesheet_component = component_identifier_for_entry(package, stylesheet_archive)?;
    if let Some(component) = stylesheet_component {
        if let Some(old) = remove_style_id {
            remove_component_object_uuids(package, component, &[old])?;
            remove_component_external_references_to_object(package, component, old)?;
        }
        add_component_object_uuids(package, component, &[new_style_id])?;
    }
    if let (Some(source), Some(target)) = (slide_component, stylesheet_component)
        && source != target
    {
        add_component_external_reference(package, source, target, new_style_id)?;
    }
    Ok(())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
