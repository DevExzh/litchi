//! Wire-preserving conversion of a theme layout graph into a live slide graph.

use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SHOW_MESSAGE_TYPE: u32 = 2;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const GUIDE_STORAGE_MESSAGE_TYPE: u32 = 3_047;
const SLIDE_NUMBER_ATTACHMENT_MESSAGE_TYPE: u32 = 2_043;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_DEPRECATED_BUILD_CHUNKS_FIELD: u32 = 3;
const SLIDE_NAME_FIELD: u32 = 10;
const SLIDE_TITLE_GEOMETRY_FIELD: u32 = 11;
const SLIDE_TITLE_SHAPE_STYLE_FIELD: u32 = 12;
const SLIDE_TITLE_TEXT_STYLE_FIELD: u32 = 13;
const SLIDE_BODY_GEOMETRY_FIELD: u32 = 14;
const SLIDE_BODY_SHAPE_STYLE_FIELD: u32 = 15;
const SLIDE_BODY_TEXT_STYLE_FIELD: u32 = 16;
const SLIDE_TEMPLATE_FIELD: u32 = 17;
const SLIDE_STATIC_GUIDES_FIELD: u32 = 18;
const SLIDE_IN_DOCUMENT_FIELD: u32 = 19;
const SLIDE_NUMBER_GEOMETRY_FIELD: u32 = 21;
const SLIDE_NUMBER_SHAPE_STYLE_FIELD: u32 = 22;
const SLIDE_NUMBER_TEXT_STYLE_FIELD: u32 = 23;
const SLIDE_TITLE_LAYOUT_FIELD: u32 = 24;
const SLIDE_BODY_LAYOUT_FIELD: u32 = 25;
const SLIDE_NUMBER_LAYOUT_FIELD: u32 = 26;
const SLIDE_NOTE_FIELD: u32 = 27;
const SLIDE_CLASSIC_STYLESHEET_FIELD: u32 = 29;
const SLIDE_OBJECT_PLACEHOLDER_FIELD: u32 = 30;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const SLIDE_BODY_PARAGRAPH_STYLES_FIELD: u32 = 31;
const SLIDE_DEPRECATED_OBJECT_VISIBLE_FIELD: u32 = 34;
const SLIDE_BODY_LIST_STYLES_FIELD: u32 = 35;
const SLIDE_TITLE_THUMBNAIL_TEXT_FIELD: u32 = 37;
const SLIDE_BODY_THUMBNAIL_TEXT_FIELD: u32 = 38;
const SLIDE_INFO_USING_OBJECT_GEOMETRY_FIELD: u32 = 39;
const SLIDE_INFO_OBJECT_GEOMETRY_MATCH_FIELD: u32 = 40;
const SLIDE_LAYER_WITH_TEMPLATE_FIELD: u32 = 41;
const SLIDE_BUILD_CHUNKS_FIELD: u32 = 43;
const SLIDE_INFOS_USING_OBJECT_GEOMETRY_FIELD: u32 = 44;
const SLIDE_INSTRUCTIONAL_TEXT_MAP_FIELD: u32 = 45;
const STORAGE_TEXT_FIELD: u32 = 3;
const STORAGE_ATTACHMENT_TABLE_FIELD: u32 = 9;
const GUIDE_STORAGE_GUIDES_FIELD: u32 = 1;
const OBJECT_REPLACEMENT_CHARACTER: &str = "\u{fffc}";
const IWORK_MESSAGE_VERSIONS: &[u32] = &[1, 0, 5];

pub(super) fn insert_slide_node(
    package: &mut IWorkPackage,
    archive_name: &str,
    show_id: u64,
    index: usize,
    node_id: u64,
    expected_count: usize,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(show_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHOW_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote show object {show_id} must contain exactly one show payload"
            )));
        };
        let message_index = *message_index;
        let show = kn::ShowArchive::decode(object.messages[message_index].data.as_slice())?;
        if show.slide_tree.slides.len() != expected_count {
            return Err(Error::InvalidFormat(
                "Keynote slide tree changed during slide creation".to_owned(),
            ));
        }
        let mut desired = show
            .slide_tree
            .slides
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        desired.insert(index, node_id);
        let data = rewrite_show_slide_references(
            object.messages[message_index].data.as_slice(),
            &show.slide_tree.slides,
            &desired,
        )?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[message_index];
        if !info.object_references.contains(&node_id) {
            info.object_references.push(node_id);
        }
        if let Some(source) = show.slide_tree.slides.first() {
            for field in &mut info.field_infos {
                if field.object_references.contains(&source.identifier)
                    && !field.object_references.contains(&node_id)
                {
                    field.object_references.push(node_id);
                }
            }
        }
        Ok(())
    })
}

#[allow(deprecated)]
pub(super) fn materialize_slide_object(
    object: &mut ArchiveObject,
    template_slide_id: u64,
    note_id: u64,
    template_only_drawables: &[u64],
) -> Result<()> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(
            "Keynote layout root must contain exactly one slide payload".to_owned(),
        ));
    };
    let message_index = *message_index;
    let original = object.messages[message_index].data.as_slice();
    let original_slide = kn::SlideArchive::decode(original)?;
    let template_only = template_only_drawables
        .iter()
        .copied()
        .collect::<HashSet<_>>();
    let mut removed_references = removed_layout_references(&original_slide);
    removed_references.extend(&template_only);
    let mut expected = original_slide.clone();

    expected.builds.clear();
    expected.build_chunk_archives.clear();
    expected.build_chunks.clear();
    expected.object_placeholder = None;
    expected.instructional_text_map = Some(kn::slide_archive::InstructionalTextMap::default());
    expected.name = None;
    expected.title_placeholder_geometry = None;
    expected.title_placeholder_shape_style_index = None;
    expected.title_placeholder_text_style_index = None;
    expected.title_layout_properties = None;
    expected.body_placeholder_geometry = None;
    expected.body_placeholder_shape_style_index = None;
    expected.body_placeholder_text_style_index = None;
    expected.body_layout_properties = None;
    expected.slide_number_placeholder_geometry = None;
    expected.slide_number_placeholder_shape_style_index = None;
    expected.slide_number_placeholder_text_style_index = None;
    expected.slide_number_layout_properties = None;
    expected.classic_stylesheet_record = None;
    expected.body_paragraph_styles.clear();
    expected.body_list_styles.clear();
    expected.thumbnail_text_for_title_placeholder = None;
    expected.thumbnail_text_for_body_placeholder = None;
    expected.slide_objects_layer_with_template = None;
    expected.template_slide = Some(tsp::Reference {
        identifier: template_slide_id,
        ..Default::default()
    });
    expected.static_guides.clear();
    expected.in_document = true;
    expected.note = Some(tsp::Reference {
        identifier: note_id,
        ..Default::default()
    });
    expected.infos_using_object_placeholder_geometry.clear();
    expected.deprecated_object_placeholder_visible_for_export = None;
    expected.info_using_object_placeholder_geometry = None;
    expected.info_using_object_placeholder_geometry_matches_object_placeholder_geometry = None;
    expected
        .owned_drawables
        .retain(|reference| !template_only.contains(&reference.identifier));
    expected
        .drawables_z_order
        .retain(|reference| !template_only.contains(&reference.identifier));

    let mut data = original.to_vec();
    for field in [
        SLIDE_BUILDS_FIELD,
        SLIDE_DEPRECATED_BUILD_CHUNKS_FIELD,
        SLIDE_STATIC_GUIDES_FIELD,
        SLIDE_BODY_PARAGRAPH_STYLES_FIELD,
        SLIDE_BODY_LIST_STYLES_FIELD,
        SLIDE_BUILD_CHUNKS_FIELD,
        SLIDE_INFOS_USING_OBJECT_GEOMETRY_FIELD,
    ] {
        data = rewrite_repeated_length_delimited_fields(&data, field, &[])?;
    }
    for (field, references) in [
        (SLIDE_OWNED_DRAWABLES_FIELD, &expected.owned_drawables),
        (SLIDE_DRAWABLES_Z_ORDER_FIELD, &expected.drawables_z_order),
    ] {
        let values = references
            .iter()
            .map(Message::encode_to_vec)
            .collect::<Vec<_>>();
        data = rewrite_repeated_length_delimited_fields(&data, field, &values)?;
    }
    for (field, was_present) in [
        (SLIDE_NAME_FIELD, original_slide.name.is_some()),
        (
            SLIDE_TITLE_GEOMETRY_FIELD,
            original_slide.title_placeholder_geometry.is_some(),
        ),
        (
            SLIDE_BODY_GEOMETRY_FIELD,
            original_slide.body_placeholder_geometry.is_some(),
        ),
        (
            SLIDE_NUMBER_GEOMETRY_FIELD,
            original_slide.slide_number_placeholder_geometry.is_some(),
        ),
        (
            SLIDE_TITLE_LAYOUT_FIELD,
            original_slide.title_layout_properties.is_some(),
        ),
        (
            SLIDE_BODY_LAYOUT_FIELD,
            original_slide.body_layout_properties.is_some(),
        ),
        (
            SLIDE_NUMBER_LAYOUT_FIELD,
            original_slide.slide_number_layout_properties.is_some(),
        ),
        (
            SLIDE_CLASSIC_STYLESHEET_FIELD,
            original_slide.classic_stylesheet_record.is_some(),
        ),
        (
            SLIDE_OBJECT_PLACEHOLDER_FIELD,
            original_slide.object_placeholder.is_some(),
        ),
        (
            SLIDE_TITLE_THUMBNAIL_TEXT_FIELD,
            original_slide
                .thumbnail_text_for_title_placeholder
                .is_some(),
        ),
        (
            SLIDE_BODY_THUMBNAIL_TEXT_FIELD,
            original_slide.thumbnail_text_for_body_placeholder.is_some(),
        ),
        (
            SLIDE_INFO_USING_OBJECT_GEOMETRY_FIELD,
            original_slide
                .info_using_object_placeholder_geometry
                .is_some(),
        ),
    ] {
        data = patch_length_delimited_field(&data, field, was_present, None)?;
    }
    for (field, was_present) in [
        (
            SLIDE_TITLE_SHAPE_STYLE_FIELD,
            original_slide.title_placeholder_shape_style_index.is_some(),
        ),
        (
            SLIDE_TITLE_TEXT_STYLE_FIELD,
            original_slide.title_placeholder_text_style_index.is_some(),
        ),
        (
            SLIDE_BODY_SHAPE_STYLE_FIELD,
            original_slide.body_placeholder_shape_style_index.is_some(),
        ),
        (
            SLIDE_BODY_TEXT_STYLE_FIELD,
            original_slide.body_placeholder_text_style_index.is_some(),
        ),
        (
            SLIDE_NUMBER_SHAPE_STYLE_FIELD,
            original_slide
                .slide_number_placeholder_shape_style_index
                .is_some(),
        ),
        (
            SLIDE_NUMBER_TEXT_STYLE_FIELD,
            original_slide
                .slide_number_placeholder_text_style_index
                .is_some(),
        ),
        (
            SLIDE_DEPRECATED_OBJECT_VISIBLE_FIELD,
            original_slide
                .deprecated_object_placeholder_visible_for_export
                .is_some(),
        ),
        (
            SLIDE_INFO_OBJECT_GEOMETRY_MATCH_FIELD,
            original_slide
                .info_using_object_placeholder_geometry_matches_object_placeholder_geometry
                .is_some(),
        ),
        (
            SLIDE_LAYER_WITH_TEMPLATE_FIELD,
            original_slide.slide_objects_layer_with_template.is_some(),
        ),
    ] {
        data = patch_varint_field(&data, field, was_present, None)?;
    }
    data = patch_length_delimited_field(
        &data,
        SLIDE_INSTRUCTIONAL_TEXT_MAP_FIELD,
        original_slide.instructional_text_map.is_some(),
        Some(&[] as &[u8]),
    )?;
    data = patch_length_delimited_field(
        &data,
        SLIDE_TEMPLATE_FIELD,
        original_slide.template_slide.is_some(),
        Some(
            &tsp::Reference {
                identifier: template_slide_id,
                ..Default::default()
            }
            .encode_to_vec(),
        ),
    )?;
    data = patch_varint_field(&data, SLIDE_IN_DOCUMENT_FIELD, true, Some(1))?;
    data = patch_length_delimited_field(
        &data,
        SLIDE_NOTE_FIELD,
        original_slide.note.is_some(),
        Some(
            &tsp::Reference {
                identifier: note_id,
                ..Default::default()
            }
            .encode_to_vec(),
        ),
    )?;
    if kn::SlideArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote layout materialization failed wire validation".to_owned(),
        ));
    }
    object.replace_message(
        message_index,
        RawMessage {
            type_: SLIDE_MESSAGE_TYPE,
            data,
        },
    )?;
    let info = &mut object.archive_info.message_infos[message_index];
    info.object_references
        .retain(|identifier| !removed_references.contains(identifier));
    for identifier in [template_slide_id, note_id] {
        if !info.object_references.contains(&identifier) {
            info.object_references.push(identifier);
        }
    }
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|identifier| !removed_references.contains(identifier));
    }
    Ok(())
}

pub(super) fn clear_user_guides(archive: &mut Archive, guide_storage_id: u64) -> Result<()> {
    let object = archive.object_mut(guide_storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote guide storage {guide_storage_id} is missing"
        ))
    })?;
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == GUIDE_STORAGE_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote guide storage {guide_storage_id} must contain exactly one guide payload"
        )));
    };
    let message_index = *message_index;
    let data = rewrite_repeated_length_delimited_fields(
        object.messages[message_index].data.as_slice(),
        GUIDE_STORAGE_GUIDES_FIELD,
        &[],
    )?;
    if !tsd::GuideStorageArchive::decode(data.as_slice())?
        .user_defined_guides
        .is_empty()
    {
        return Err(Error::InvalidFormat(
            "Keynote guide reset failed validation".to_owned(),
        ));
    }
    object.replace_message(
        message_index,
        RawMessage {
            type_: GUIDE_STORAGE_MESSAGE_TYPE,
            data,
        },
    )?;
    Ok(())
}

pub(super) fn prepare_slide_number(
    archive: &mut Archive,
    slide_id: u64,
    attachment_id: u64,
) -> Result<Option<ArchiveObject>> {
    let slide = archive.object(slide_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote created slide {slide_id} is missing"))
    })?;
    let slide = slide
        .messages
        .iter()
        .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .map(|message| kn::SlideArchive::decode(message.data.as_slice()))
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote created slide payload is missing".to_owned())
        })?;
    let Some(placeholder_id) = slide
        .slide_number_placeholder
        .map(|reference| reference.identifier)
    else {
        return Ok(None);
    };
    let placeholder = archive.object(placeholder_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote slide-number placeholder {placeholder_id} is missing"
        ))
    })?;
    let placeholder = placeholder
        .messages
        .iter()
        .find(|message| message.type_ == PLACEHOLDER_MESSAGE_TYPE)
        .map(|message| kn::PlaceholderArchive::decode(message.data.as_slice()))
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote slide-number placeholder payload is missing".to_owned())
        })?;
    let storage_id = placeholder
        .super_
        .owned_storage
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote slide-number placeholder has no storage".to_owned())
        })?
        .identifier;
    let table = tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: Some(tsp::Reference {
                identifier: attachment_id,
                ..Default::default()
            }),
        }],
    };
    let storage = archive.object_mut(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote slide-number storage {storage_id} is missing"
        ))
    })?;
    let indexes = storage
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide-number storage {storage_id} must contain exactly one storage payload"
        )));
    };
    let message_index = *message_index;
    let message_type = storage.messages[message_index].type_;
    let original = storage.messages[message_index].data.as_slice();
    let decoded = tswp::StorageArchive::decode(original)?;
    let data = rewrite_repeated_length_delimited_fields(
        original,
        STORAGE_TEXT_FIELD,
        &[OBJECT_REPLACEMENT_CHARACTER.as_bytes().to_vec()],
    )?;
    let data = patch_length_delimited_field(
        &data,
        STORAGE_ATTACHMENT_TABLE_FIELD,
        decoded.table_attachment.is_some(),
        Some(&table.encode_to_vec()),
    )?;
    let verified = tswp::StorageArchive::decode(data.as_slice())?;
    if verified.text != vec![OBJECT_REPLACEMENT_CHARACTER.to_owned()]
        || verified.table_attachment != Some(table)
    {
        return Err(Error::InvalidFormat(
            "Keynote slide-number storage initialization failed validation".to_owned(),
        ));
    }
    storage.replace_message(
        message_index,
        RawMessage {
            type_: message_type,
            data,
        },
    )?;
    let info = &mut storage.archive_info.message_infos[message_index];
    if !info.object_references.contains(&attachment_id) {
        info.object_references.push(attachment_id);
    }

    let mut attachment = ArchiveObject::new(
        attachment_id,
        vec![RawMessage {
            type_: SLIDE_NUMBER_ATTACHMENT_MESSAGE_TYPE,
            data: kn::SlideNumberAttachmentArchive {
                super_: tswp::TextualAttachmentArchive {
                    string_equivalent: Some(String::new()),
                    kind: Some(tswp::textual_attachment_archive::Kind::KKindPageNumber as i32),
                },
            }
            .encode_to_vec(),
        }],
    )?;
    attachment.archive_info.message_infos[0].versions = IWORK_MESSAGE_VERSIONS.to_vec();
    Ok(Some(attachment))
}

#[allow(deprecated)]
fn removed_layout_references(slide: &kn::SlideArchive) -> HashSet<u64> {
    slide
        .builds
        .iter()
        .chain(&slide.build_chunks)
        .chain(slide.object_placeholder.iter())
        .chain(slide.classic_stylesheet_record.iter())
        .chain(&slide.body_paragraph_styles)
        .chain(&slide.body_list_styles)
        .chain(slide.note.iter())
        .chain(&slide.infos_using_object_placeholder_geometry)
        .chain(slide.info_using_object_placeholder_geometry.iter())
        .map(|reference| reference.identifier)
        .chain(
            slide
                .build_chunk_archives
                .iter()
                .filter_map(|chunk| chunk.build.as_ref().map(|reference| reference.identifier)),
        )
        .collect()
}
