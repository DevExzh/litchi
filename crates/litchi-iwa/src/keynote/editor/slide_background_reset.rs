//! Deletion of a direct Keynote slide-background style override.

use super::slide_style_graph::{
    SLIDE_MESSAGE_TYPE, SLIDE_STYLE_MESSAGE_TYPE, is_collapsible_background_variation,
    patch_slide_style_reference, style_is_exclusive,
};
use super::slide_style_metadata::{ensure_slide_style_external_reference, update_package_metadata};
use super::slide_style_registry::patch_stylesheet;
use super::*;

pub(super) fn reset_slide_background(editor: &mut KeynoteEditor, slide_index: usize) -> Result<()> {
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
    let style_id = slide.style.identifier;
    let style: kn::SlideStyleArchive =
        graph.decode_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
    if style.super_.is_variation != Some(true) {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide {slide_index} does not have a variation-style background override"
        )));
    }
    let parent_style_id = style
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide style {style_id} has no parent style"
            ))
        })?;
    let stylesheet_id = style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide style {style_id} has no stylesheet reference"
            ))
        })?;
    let style_raw =
        graph.message_data_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
    let properties_raw =
        required_length_delimited_payload(style_raw, 11, "Keynote slide-background variation")?;
    required_length_delimited_payload(properties_raw, 1, "Keynote slide-background variation")?;
    let slide_archive = graph.archive_name(slide_info.slide_id)?.to_owned();
    let stylesheet_archive = graph.archive_name(stylesheet_id)?.to_owned();
    if graph.archive_name(style_id)? != stylesheet_archive {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide style {style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let exclusive = style_is_exclusive(&graph, style_id)?;
    let collapsible = is_collapsible_background_variation(&style, style_raw)?;
    let mut staged = editor.package().clone();

    if collapsible {
        patch_slide_style_reference(
            &mut staged,
            &slide_archive,
            slide_info.slide_id,
            style_id,
            parent_style_id,
        )?;
        if exclusive {
            patch_stylesheet(
                &mut staged,
                &stylesheet_archive,
                stylesheet_id,
                Some(style_id),
                None,
            )?;
            update_package_metadata(
                &mut staged,
                &slide_archive,
                &stylesheet_archive,
                Some(style_id),
                None,
            )?;
            release_package_identifier_suffix(&mut staged, &[style_id])?;
        }
        ensure_slide_style_external_reference(
            &mut staged,
            &slide_archive,
            &stylesheet_archive,
            parent_style_id,
        )?;
    } else {
        let replacement_data = style_without_background(style_raw, &style)?;
        let protected_object_references = [parent_style_id, stylesheet_id];
        if exclusive {
            replace_style_data(
                &mut staged,
                &stylesheet_archive,
                style_id,
                replacement_data,
                &protected_object_references,
            )?;
        } else {
            let replacement_id = next_object_identifier(&staged)?;
            let replacement = clone_style_object(
                &staged,
                &stylesheet_archive,
                style_id,
                replacement_id,
                replacement_data,
                &protected_object_references,
            )?;
            patch_slide_style_reference(
                &mut staged,
                &slide_archive,
                slide_info.slide_id,
                style_id,
                replacement_id,
            )?;
            patch_stylesheet(
                &mut staged,
                &stylesheet_archive,
                stylesheet_id,
                None,
                Some((parent_style_id, replacement_id, replacement)),
            )?;
            update_package_metadata(
                &mut staged,
                &slide_archive,
                &stylesheet_archive,
                None,
                Some(replacement_id),
            )?;
            set_package_last_object_identifier(&mut staged, replacement_id)?;
        }
    }

    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_background_override(slide_index)?.is_some() {
        return Err(Error::InvalidFormat(
            "Keynote slide-background reset retained a direct override".to_owned(),
        ));
    }
    verified.slide_background(slide_index)?;
    editor.text = IWorkTextEditor::from_package(staged);
    Ok(())
}

fn style_without_background(original: &[u8], style: &kn::SlideStyleArchive) -> Result<Vec<u8>> {
    let override_count = style.override_count.ok_or_else(|| {
        Error::InvalidFormat("Keynote slide-background variation has no override count".to_owned())
    })?;
    let replacement_count = override_count.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat(
            "Keynote slide-background variation has a zero override count".to_owned(),
        )
    })?;
    let data = patch_nested_length_delimited_field(original, &[11, 1], true, None)?;
    let data = patch_varint_field(&data, 10, true, Some(u64::from(replacement_count)))?;
    let mut expected = style.clone();
    expected.override_count = Some(replacement_count);
    expected
        .slide_properties
        .as_mut()
        .ok_or_else(|| {
            Error::InvalidFormat(
                "Keynote slide-background variation has no slide properties".to_owned(),
            )
        })?
        .fill = None;
    if kn::SlideStyleArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote slide-background removal changed unrelated style properties".to_owned(),
        ));
    }
    Ok(data)
}

fn replace_style_data(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    data: Vec<u8>,
    protected_object_references: &[u64],
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide style {style_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_STYLE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide style {style_id} must have exactly one SlideStyleArchive payload"
            )));
        }
        let index = indexes[0];
        object.replace_message(
            index,
            RawMessage {
                type_: SLIDE_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        remove_background_field_metadata(
            &mut object.archive_info.message_infos[index],
            protected_object_references,
        );
        Ok(())
    })
}

fn clone_style_object(
    package: &IWorkPackage,
    archive_name: &str,
    source_id: u64,
    replacement_id: u64,
    data: Vec<u8>,
    protected_object_references: &[u64],
) -> Result<ArchiveObject> {
    let archive = package.archive(archive_name)?;
    let source = archive.object(source_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote slide style {source_id} is missing"))
    })?;
    if source.messages.len() != 1
        || source.messages[0].type_ != SLIDE_STYLE_MESSAGE_TYPE
        || source.archive_info.message_infos.len() != 1
    {
        return Err(Error::InvalidFormat(format!(
            "shared Keynote slide style {source_id} is not a single-message style object"
        )));
    }
    let length = u32::try_from(data.len())
        .map_err(|_| Error::InvalidFormat("Keynote slide style payload exceeds u32".to_owned()))?;
    let mut replacement = ArchiveObject::new(
        replacement_id,
        vec![RawMessage {
            type_: SLIDE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    replacement.archive_info.should_merge = source.archive_info.should_merge;
    replacement.archive_info.message_infos[0] = source.archive_info.message_infos[0].clone();
    replacement.archive_info.message_infos[0].length = length;
    replacement.archive_info.message_infos[0].type_ = SLIDE_STYLE_MESSAGE_TYPE;
    remove_background_field_metadata(
        &mut replacement.archive_info.message_infos[0],
        protected_object_references,
    );
    Ok(replacement)
}

fn remove_background_field_metadata(
    info: &mut crate::archive::MessageInfo,
    protected_object_references: &[u64],
) {
    const BACKGROUND_FIELD_PATH: &[u32] = &[11, 1];

    let mut removed_object_references = HashSet::new();
    let mut removed_data_references = HashSet::new();
    info.field_infos.retain(|field| {
        if field.path.path.starts_with(BACKGROUND_FIELD_PATH) {
            removed_object_references.extend(field.object_references.iter().copied());
            removed_data_references.extend(field.data_references.iter().copied());
            false
        } else {
            true
        }
    });
    let retained_object_references = info
        .field_infos
        .iter()
        .flat_map(|field| field.object_references.iter().copied())
        .collect::<HashSet<_>>();
    let retained_data_references = info
        .field_infos
        .iter()
        .flat_map(|field| field.data_references.iter().copied())
        .collect::<HashSet<_>>();
    info.object_references.retain(|identifier| {
        !removed_object_references.contains(identifier)
            || retained_object_references.contains(identifier)
            || protected_object_references.contains(identifier)
    });
    info.data_references.retain(|identifier| {
        !removed_data_references.contains(identifier)
            || retained_data_references.contains(identifier)
    });
}
