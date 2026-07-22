//! Typed construction and discovery of body-anchored Pages image graphs.

use super::*;
use crate::IWorkThemeArchive;
use crate::shapes::{geometry_from_drawable, patch_drawable_geometry};

const THEME_MESSAGE_TYPE: u32 = 10_001;
const DRAWABLE_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_IMAGE_FLAGS: u32 = 0;
const DEFAULT_IMAGE_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapType {
    Square = 4,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapDirection {
    BothSides = 2,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapFit {
    Text = 1,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum HorizontalAnchorBasis {
    BodyMargin = 0,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum VerticalAnchorBasis {
    Page = 1,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ImageObjectIds {
    pub(super) drawable: u64,
    title: u64,
    caption: u64,
    pub(super) attachment: u64,
}

impl ImageObjectIds {
    pub(super) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            drawable: identifier(0)?,
            title: identifier(1)?,
            caption: identifier(2)?,
            attachment: identifier(3)?,
        })
    }

    pub(super) const fn last(self) -> u64 {
        self.attachment
    }

    pub(super) const fn all(self) -> [u64; 4] {
        [self.drawable, self.title, self.caption, self.attachment]
    }

    pub(super) const fn uuid_objects(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(super) struct BodyImageGraph {
    pub(super) archive_name: String,
    pub(super) attachment_id: u64,
    pub(super) info: PagesImageInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn image_geometry(
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
            "Pages image position must be finite and size must be finite and positive".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_IMAGE_ROTATION_DEGREES),
    }
    .validate()
}

pub(super) fn image_style_id(package: &IWorkPackage, root: &DocumentArchive) -> Result<u64> {
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
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == THEME_MESSAGE_TYPE)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Pages theme {theme_id} has no theme payload"))
        })?;
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.image_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no image style preset".to_owned()))
}

pub(super) fn body_image_infos(editor: &PagesEditor) -> Result<Vec<PagesImageInfo>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut images = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        let package = editor.package();
        let archive_name = find_object_archive(package, drawable.identifier)?;
        let archive = package.archive(&archive_name)?;
        let object = archive.object(drawable.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages drawable {} is missing", drawable.identifier))
        })?;
        if object
            .messages
            .iter()
            .any(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        {
            images.push(image_info(
                editor.package(),
                drawable.identifier,
                entry.character_index,
            )?);
        }
    }
    images.sort_by_key(|image| image.anchor_character_index);
    Ok(images)
}

pub(super) fn body_image_graph(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<BodyImageGraph> {
    let info = body_image_infos(editor)?
        .into_iter()
        .find(|image| image.drawable_object_id == drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a body-anchored image"
            ))
        })?;
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut attachment_ids = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
        .filter(|entry| entry.character_index == info.anchor_character_index)
    {
        let Some(reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        if attachment
            .drawable
            .is_some_and(|drawable| drawable.identifier == drawable_object_id)
        {
            attachment_ids.push(reference.identifier);
        }
    }
    let [attachment_id] = attachment_ids.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages image {drawable_object_id} has {} body attachments; expected one",
            attachment_ids.len()
        )));
    };
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    if body_units.get(info.anchor_character_index as usize) != Some(&0xfffc) {
        return Err(Error::InvalidFormat(format!(
            "Pages image {drawable_object_id} attachment is not backed by an object-replacement character"
        )));
    }

    let image: tsd::ImageArchive = decode_typed_package_object(
        editor.package(),
        drawable_object_id,
        IMAGE_MESSAGE_TYPE,
        "TSD.ImageArchive",
    )?;
    if image.super_.parent.map(|parent| parent.identifier) != Some(editor.body_storage_id) {
        return Err(Error::InvalidFormat(format!(
            "Pages image {drawable_object_id} is not owned by the body storage"
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
            "Pages image {drawable_object_id} occurs {z_order_count} times in drawable z-order"
        )));
    }
    let mut object_ids = vec![drawable_object_id];
    for reference in [image.super_.title, image.super_.caption]
        .into_iter()
        .flatten()
    {
        decode_typed_package_object::<tsd::StandinCaptionArchive>(
            editor.package(),
            reference.identifier,
            STANDIN_CAPTION_MESSAGE_TYPE,
            "TSD.StandinCaptionArchive",
        )?;
        object_ids.push(reference.identifier);
    }
    object_ids.push(*attachment_id);
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages image {drawable_object_id} reuses private graph identifiers"
        )));
    }
    let archive_name = find_object_archive(editor.package(), drawable_object_id)?;
    for identifier in &object_ids {
        if find_object_archive(editor.package(), *identifier)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Pages image {drawable_object_id} private graph spans multiple archives"
            )));
        }
    }

    let registered =
        component_uuid_identifiers(editor.package(), DOCUMENT_OBJECT_ID)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    let archive = editor.package().archive(&archive_name)?;
    let mut data_references = Vec::new();
    for identifier in &object_ids {
        let object = archive.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages image object {identifier} is missing"))
        })?;
        data_references.extend(
            object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|message| {
                    message
                        .data_references
                        .iter()
                        .chain(
                            message
                                .field_infos
                                .iter()
                                .flat_map(|field| &field.data_references),
                        )
                        .map(|data| (*data, *identifier))
                }),
        );
    }
    Ok(BodyImageGraph {
        archive_name,
        attachment_id: *attachment_id,
        info,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

fn image_info(
    package: &IWorkPackage,
    identifier: u64,
    anchor_character_index: u32,
) -> Result<PagesImageInfo> {
    let image: tsd::ImageArchive =
        decode_typed_package_object(package, identifier, IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")?;
    let image_data_identifier = image
        .data
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages image {identifier} has no primary data reference"
            ))
        })?
        .identifier;
    Ok(PagesImageInfo {
        drawable_object_id: identifier,
        anchor_character_index,
        image_data_identifier,
        thumbnail_data_identifier: image.thumbnail_data.map(|reference| reference.identifier),
        geometry: geometry_from_drawable(&image.super_)?,
        original_size: image.original_size.map(drawable_size),
        natural_size: image.natural_size.map(drawable_size),
    })
}

pub(super) fn image_objects(
    ids: ImageObjectIds,
    body_storage_id: u64,
    style_id: u64,
    data_identifier: u64,
    geometry: DrawableGeometry,
    left_margin: f32,
) -> Result<[ArchiveObject; 4]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Pages image geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Pages image geometry has no size".to_owned())
    })?;
    let image = tsd::ImageArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point {
                    x: position.x,
                    y: position.y,
                }),
                size: Some(tsp::Size {
                    width: size.width,
                    height: size.height,
                }),
                flags: geometry.flags,
                angle: geometry.angle,
            }),
            parent: Some(reference(body_storage_id)),
            exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                r#type: Some(TextWrapType::Square as u32),
                direction: Some(TextWrapDirection::BothSides as u32),
                fit_type: Some(TextWrapFit::Text as u32),
                margin: Some(DEFAULT_TEXT_WRAP_MARGIN_POINTS),
                alpha_threshold: Some(DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD),
                is_html_wrap: Some(false),
            }),
            locked: Some(false),
            aspect_ratio_locked: Some(false),
            title: Some(reference(ids.title)),
            caption: Some(reference(ids.caption)),
            title_hidden: Some(false),
            caption_hidden: Some(false),
            ..Default::default()
        },
        data: Some(tsp::DataReference {
            identifier: data_identifier,
        }),
        style: Some(reference(style_id)),
        original_size: Some(tsp::Size {
            width: size.width,
            height: size.height,
        }),
        natural_size: Some(tsp::Size {
            width: size.width,
            height: size.height,
        }),
        flags: Some(DEFAULT_IMAGE_FLAGS),
        ..Default::default()
    };
    let attachment = DrawableAttachmentArchive {
        drawable: Some(reference(ids.drawable)),
        h_offset_type: Some(HorizontalAnchorBasis::BodyMargin as u32),
        h_offset: Some(position.x - left_margin),
        v_offset_type: Some(VerticalAnchorBasis::Page as u32),
        v_offset: Some(position.y),
    };
    Ok([
        pages_image_object(
            ids.drawable,
            IMAGE_MESSAGE_TYPE,
            image,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[data_identifier],
        )?,
        pages_image_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        pages_image_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        pages_image_object(
            ids.attachment,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            attachment,
            &STANDARD_MESSAGE_VERSION,
            &[ids.drawable],
            &[],
        )?,
    ])
}

pub(super) fn set_image_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    geometry.validate()?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(image_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages image object {image_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == IMAGE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages image {image_id} must have exactly one ImageArchive payload"
            )));
        };
        let data = transform_length_delimited_field(
            object.messages[*message_index].data.as_slice(),
            1,
            |drawable| patch_drawable_geometry(drawable, geometry),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: IMAGE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn pages_image_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = versions.to_vec();
    info.object_references = object_references.to_vec();
    info.data_references = data_references.to_vec();
    Ok(object)
}

fn drawable_size(size: tsp::Size) -> DrawableSize {
    DrawableSize {
        width: size.width,
        height: size.height,
    }
}

fn object_has_message_type(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
) -> Result<bool> {
    let archive_name = find_object_archive(package, identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages object {identifier} is missing from {archive_name}"
        ))
    })?;
    Ok(object
        .messages
        .iter()
        .any(|message| message.type_ == message_type))
}
