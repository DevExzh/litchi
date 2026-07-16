//! Typed construction support and strict discovery for Pages body shapes.

use super::*;
use crate::IWorkThemeArchive;
use crate::shapes::{shape_line_segment, shape_path_kind, shape_preset};

const THEME_MESSAGE_TYPE: u32 = 10_001;
const DRAWABLE_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
pub(super) const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;

pub(super) struct BodyShapeGraph {
    pub(super) archive_name: String,
    pub(super) info: PagesBodyShapeInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
}

pub(super) fn new_shape_geometry(
    position: DrawablePoint,
    size: DrawableSize,
) -> Result<DrawableGeometry> {
    if !position.x.is_finite()
        || !position.y.is_finite()
        || !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Pages shape position must be finite and size must be finite and positive".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_ROTATION_DEGREES),
    }
    .validate()
}

pub(super) fn shape_style_id(package: &IWorkPackage, root: &DocumentArchive) -> Result<u64> {
    let theme_id = root
        .theme
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".to_owned()))?
        .identifier;
    let archive_name = find_object_archive(package, theme_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Pages theme {theme_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages theme {theme_id} must have exactly one theme payload"
        )));
    };
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.shape_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no shape style preset".to_owned()))
}

pub(super) fn body_shape_infos(editor: &PagesEditor) -> Result<Vec<PagesBodyShapeInfo>> {
    let mut shapes = Vec::new();
    for text in editor.drawable_text_storages()? {
        let Some(shape) = shape_payload(editor.package(), text.drawable_object_id)? else {
            continue;
        };
        if shape.is_text_box == Some(false) {
            shapes.push(body_shape_graph_from_text(editor, text, shape)?.info);
        }
    }
    shapes.sort_by_key(|shape| shape.anchor_character_index);
    Ok(shapes)
}

pub(super) fn body_shape_graph(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<BodyShapeGraph> {
    let text = editor
        .drawable_text_storages()?
        .into_iter()
        .find(|text| text.drawable_object_id == drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a reachable text-bearing shape"
            ))
        })?;
    let shape = shape_payload(editor.package(), drawable_object_id)?.ok_or_else(|| {
        Error::ParseError(format!(
            "Pages drawable {drawable_object_id} is not an ordinary shape"
        ))
    })?;
    body_shape_graph_from_text(editor, text, shape)
}

#[allow(deprecated)]
fn body_shape_graph_from_text(
    editor: &PagesEditor,
    text: PagesDrawableTextInfo,
    shape: crate::protobuf::tswp::ShapeInfoArchive,
) -> Result<BodyShapeGraph> {
    let drawable_object_id = text.drawable_object_id;
    if shape.is_text_box != Some(false) {
        return Err(Error::ParseError(format!(
            "Pages drawable {drawable_object_id} is not an ordinary shape"
        )));
    }
    if shape
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(editor.body_storage_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} is not owned by the body storage"
        )));
    }
    if shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(text.storage.object_id)
        || shape
            .deprecated_storage
            .as_ref()
            .map(|reference| reference.identifier)
            != Some(text.storage.object_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} has inconsistent storage ownership"
        )));
    }
    if !shape.super_.super_.pencil_annotations.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} has unsupported pencil annotations"
        )));
    }
    if shape_storage_owner_count(editor.package(), text.storage.object_id)? != 1 {
        return Err(Error::InvalidFormat(format!(
            "Pages shape storage {} is not owned by exactly one shape",
            text.storage.object_id
        )));
    }

    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut attachments = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(reference) = entry.object else {
            continue;
        };
        let archive_name = find_object_archive(editor.package(), reference.identifier)?;
        let archive = editor.package().archive(&archive_name)?;
        let Some(object) = archive.object(reference.identifier) else {
            continue;
        };
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == DRAWABLE_ATTACHMENT_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages attachment {} repeats its payload",
                reference.identifier
            )));
        };
        let attachment = DrawableAttachmentArchive::decode(message.data.as_slice())?;
        if attachment
            .drawable
            .as_ref()
            .is_some_and(|drawable| drawable.identifier == drawable_object_id)
        {
            attachments.push((entry.character_index, reference.identifier));
        }
    }
    let [(anchor_character_index, attachment_id)] = attachments.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} has {} body attachments; expected one",
            attachments.len()
        )));
    };
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    if body_units.get(*anchor_character_index as usize) != Some(&0xfffc) {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} attachment is not backed by an object-replacement character"
        )));
    }

    let document = root_document(editor.package())?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let z_order: tp::DrawablesZOrderArchive = decode_typed_package_object(
        editor.package(),
        z_order_id.identifier,
        DRAWABLE_Z_ORDER_MESSAGE_TYPE,
        "TP.DrawablesZOrderArchive",
    )?;
    let z_order_count = z_order
        .drawables
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count();
    if z_order_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} occurs {z_order_count} times in drawable z-order"
        )));
    }

    let caption_id = required_reference(
        drawable_object_id,
        shape.super_.super_.caption.as_ref(),
        "caption stand-in",
    )?;
    let title_id = required_reference(
        drawable_object_id,
        shape.super_.super_.title.as_ref(),
        "title stand-in",
    )?;
    let object_ids = vec![
        drawable_object_id,
        caption_id,
        title_id,
        text.storage.object_id,
        *attachment_id,
    ];
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages shape {drawable_object_id} reuses private graph identifiers"
        )));
    }
    let archive_name = find_object_archive(editor.package(), drawable_object_id)?;
    for identifier in &object_ids {
        if find_object_archive(editor.package(), *identifier)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Pages shape {drawable_object_id} private graph spans multiple archives"
            )));
        }
    }
    for (identifier, message_types, label) in [
        (
            caption_id,
            &[STANDIN_CAPTION_MESSAGE_TYPE][..],
            "caption stand-in",
        ),
        (
            title_id,
            &[STANDIN_CAPTION_MESSAGE_TYPE][..],
            "title stand-in",
        ),
        (text.storage.object_id, STORAGE_MESSAGE_TYPES, "storage"),
        (
            *attachment_id,
            &[DRAWABLE_ATTACHMENT_MESSAGE_TYPE][..],
            "attachment",
        ),
    ] {
        let archive = editor.package().archive(&archive_name)?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages shape {label} {identifier} is missing"))
        })?;
        if object
            .messages
            .iter()
            .filter(|message| message_types.contains(&message.type_))
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Pages shape {label} {identifier} must have exactly one expected payload"
            )));
        }
    }

    let registered =
        component_uuid_identifiers(editor.package(), DOCUMENT_OBJECT_ID)?.unwrap_or_default();
    let expected_uuid_ids = [
        drawable_object_id,
        caption_id,
        title_id,
        text.storage.object_id,
    ];
    let uuid_object_ids = expected_uuid_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    if !registered.is_empty() && uuid_object_ids.len() != expected_uuid_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages document UUID map does not cover shape {drawable_object_id}"
        )));
    }
    let kind = shape_path_kind(&shape)?;
    let preset = shape_preset(&shape)?;
    let line_segment = shape_line_segment(&shape)?;
    let line_endpoints = line_segment
        .map(|_| shape_line_endpoints(editor.package(), &archive_name, drawable_object_id))
        .transpose()?;
    let geometry = shape_geometry(editor.package(), &archive_name, drawable_object_id)?;
    let properties = shape_properties(editor.package(), &archive_name, drawable_object_id)?;
    Ok(BodyShapeGraph {
        archive_name,
        info: PagesBodyShapeInfo {
            drawable_object_id,
            anchor_character_index: *anchor_character_index,
            kind,
            preset,
            line_segment,
            line_endpoints,
            storage: text.storage,
            geometry,
            properties,
        },
        object_ids,
        uuid_object_ids,
    })
}

fn required_reference(
    drawable_object_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages shape {drawable_object_id} has no {label} object"
            ))
        })
}

fn shape_payload(
    package: &IWorkPackage,
    drawable_object_id: u64,
) -> Result<Option<crate::protobuf::tswp::ShapeInfoArchive>> {
    let archive_name = find_object_archive(package, drawable_object_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages drawable {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.is_empty() {
        return Ok(None);
    }
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages drawable {drawable_object_id} repeats its shape payload"
        )));
    };
    Ok(Some(crate::protobuf::tswp::ShapeInfoArchive::decode(
        message.data.as_slice(),
    )?))
}

fn shape_storage_owner_count(package: &IWorkPackage, storage_id: u64) -> Result<usize> {
    let mut owners = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            {
                let shape =
                    crate::protobuf::tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                if shape
                    .owned_storage
                    .as_ref()
                    .is_some_and(|reference| reference.identifier == storage_id)
                {
                    owners += 1;
                }
            }
        }
    }
    Ok(owners)
}
