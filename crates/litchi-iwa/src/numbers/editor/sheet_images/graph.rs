//! Typed construction and discovery of sheet-owned Numbers image graphs.

use super::*;
use crate::DrawableTitleCaption;
use crate::IWorkThemeArchive;
use crate::image_adjustments::image_adjustments_from_archive;
use crate::image_caption::{
    CAPTION_INFO_MESSAGE_TYPE, CAPTION_PLACEMENT_MESSAGE_TYPE, CaptionObjectIds, CaptionThemeStyle,
    DrawableCaptionKind, SHAPE_STYLE_MESSAGE_TYPE, STORAGE_MESSAGE_TYPE, caption_objects,
    patch_drawable_caption_reference, replace_object_reference, standin_caption_object,
};
use crate::shapes::{
    drawable_properties, geometry_from_drawable, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};

const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ImageObjectIds {
    pub(super) drawable: u64,
    title: u64,
    caption: u64,
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
        })
    }

    pub(super) const fn last(self) -> u64 {
        self.caption
    }

    pub(super) const fn all(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(super) struct ImageCreationContext {
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) style_id: u64,
    pub(super) stylesheet_component_id: u64,
    pub(super) caption_theme: CaptionThemeStyle,
    pub(super) language: Option<String>,
}

#[derive(Debug, Clone)]
pub(super) struct ImageCaptionSlot {
    pub(super) reference_id: u64,
    pub(super) storage_id: Option<u64>,
    object_ids: Vec<u64>,
}

pub(super) struct SheetImageGraph {
    pub(super) sheet_id: u64,
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) info: NumbersSheetImageInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn image_creation_values(options: NumbersSheetImageOptions) -> Result<DrawableGeometry> {
    if !options.natural_size.width.is_finite()
        || !options.natural_size.height.is_finite()
        || options.natural_size.width <= 0.0
        || options.natural_size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Numbers image natural size must be finite and greater than zero".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(options.position),
        size: Some(options.size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_IMAGE_ROTATION_DEGREES),
    }
    .validate()
}

pub(super) fn image_creation_context(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<ImageCreationContext> {
    let (archive_name, _, _) = numbers_sheet(editor.package(), sheet_id)?;
    let document = numbers_document(editor.package())?;
    let style_id = image_style_id(editor.package(), document.theme.identifier)?;
    let caption_theme = image_caption_theme(
        editor.package(),
        document.theme.identifier,
        document.stylesheet.identifier,
    )?;
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let style_archive = object_locations(editor.package())?
        .get(&style_id)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers image style {style_id} is missing"))
        })?;
    let stylesheet_component_id = component_identifier_for_entry(editor.package(), &style_archive)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers stylesheet component {style_archive} is not registered"
            ))
        })?;
    Ok(ImageCreationContext {
        archive_name,
        component_id,
        style_id,
        stylesheet_component_id,
        caption_theme,
        language: document.super_.document_language,
    })
}

fn image_style_id(package: &IWorkPackage, theme_id: u64) -> Result<u64> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme object {theme_id} must have exactly one theme payload"
        )));
    };
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.image_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Numbers theme has no image style preset".to_owned()))
}

fn image_caption_theme(
    package: &IWorkPackage,
    theme_id: u64,
    stylesheet_id: u64,
) -> Result<CaptionThemeStyle> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme object {theme_id} must have exactly one theme payload"
        )));
    };
    let paragraph_style_id = IWorkThemeArchive::decode(&message.data)?
        .extensions
        .application
        .ok_or_else(|| Error::InvalidFormat("Numbers theme has no application presets".to_owned()))?
        .caption_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers theme has no caption style preset".to_owned())
        })?;
    for (identifier, label) in [
        (stylesheet_id, "stylesheet"),
        (paragraph_style_id, "caption paragraph style"),
    ] {
        if !locations.contains_key(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {label} object {identifier} is missing"
            )));
        }
    }
    Ok(CaptionThemeStyle {
        stylesheet_id,
        paragraph_style_id,
    })
}

pub(super) fn image_infos(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<Vec<NumbersSheetImageInfo>> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut images = Vec::new();
    for reference in sheet.drawable_infos {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let archive = editor.package().archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let image_messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == IMAGE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if image_messages.is_empty() {
            continue;
        }
        if image_messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {} has multiple image payloads",
                reference.identifier
            )));
        }
        images.push(image_info(
            editor.package(),
            sheet_id,
            reference.identifier,
        )?);
    }
    Ok(images)
}

pub(super) fn image_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetImageGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own image {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers image {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary image"
        )));
    };
    let image = tsd::ImageArchive::decode(message.data.as_slice())?;
    if image.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {drawable_object_id} is not owned by sheet {sheet_id}"
        )));
    }
    let title = image_caption_slot_from_reference(
        editor.package(),
        image.super_.title,
        DrawableCaptionKind::Title,
    )?;
    let caption = image_caption_slot_from_reference(
        editor.package(),
        image.super_.caption,
        DrawableCaptionKind::Caption,
    )?;
    let style_id = image
        .style
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers image {drawable_object_id} has no image style"
            ))
        })?
        .identifier;
    if !locations.contains_key(&style_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {drawable_object_id} style {style_id} is missing"
        )));
    }
    let mut object_ids = vec![drawable_object_id];
    object_ids.extend(title.object_ids.iter().copied());
    object_ids.extend(caption.object_ids.iter().copied());
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {drawable_object_id} reuses private graph identifiers"
        )));
    }
    for &identifier in &object_ids[1..] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers image {drawable_object_id} private graph spans multiple archives"
            )));
        }
    }
    let mut allowed_references = [sheet_id, title.reference_id, caption.reference_id, style_id]
        .into_iter()
        .collect::<HashSet<_>>();
    allowed_references.extend(image.super_.comment.map(|reference| reference.identifier));
    let unexpected_references = object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        })
        .copied()
        .filter(|identifier| !allowed_references.contains(identifier))
        .collect::<HashSet<_>>();
    if !unexpected_references.is_empty() {
        return Err(Error::ParseError(format!(
            "Numbers image {drawable_object_id} has unsupported private references {unexpected_references:?}"
        )));
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    let mut data_references = object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.data_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.data_references),
            )
        })
        .copied()
        .map(|data_identifier| (data_identifier, drawable_object_id))
        .collect::<Vec<_>>();
    data_references.sort_unstable();
    data_references.dedup();
    let info = image_info(editor.package(), sheet_id, drawable_object_id)?;
    for identifier in
        std::iter::once(info.image_data_identifier).chain(info.thumbnail_data_identifier)
    {
        if !data_references.contains(&(identifier, drawable_object_id)) {
            return Err(Error::InvalidFormat(format!(
                "Numbers image {drawable_object_id} data {identifier} is missing from archive metadata"
            )));
        }
    }
    Ok(SheetImageGraph {
        sheet_id,
        archive_name,
        component_id,
        info,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

pub(super) fn image_title_caption(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<DrawableTitleCaption> {
    image_graph(editor, sheet_id, drawable_object_id)?;
    let image = image_archive(editor.package(), drawable_object_id)?;
    let title = image_caption_slot_from_reference(
        editor.package(),
        image.super_.title,
        DrawableCaptionKind::Title,
    )?;
    let caption = image_caption_slot_from_reference(
        editor.package(),
        image.super_.caption,
        DrawableCaptionKind::Caption,
    )?;
    let text_editor = IWorkTextEditor::from_package(editor.package().clone());
    Ok(DrawableTitleCaption {
        title: title
            .storage_id
            .map(|storage_id| {
                text_editor
                    .storage(crate::text::native_storage_id(storage_id)?)
                    .map(|storage| storage.storage.into_text())
            })
            .transpose()?,
        caption: caption
            .storage_id
            .map(|storage_id| {
                text_editor
                    .storage(crate::text::native_storage_id(storage_id)?)
                    .map(|storage| storage.storage.into_text())
            })
            .transpose()?,
    })
}

pub(super) fn image_caption_slot(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    kind: DrawableCaptionKind,
) -> Result<ImageCaptionSlot> {
    image_graph(editor, sheet_id, drawable_object_id)?;
    let image = image_archive(editor.package(), drawable_object_id)?;
    let reference = match kind {
        DrawableCaptionKind::Caption => image.super_.caption,
        DrawableCaptionKind::Title => image.super_.title,
    };
    image_caption_slot_from_reference(editor.package(), reference, kind)
}

fn image_archive(package: &IWorkPackage, identifier: u64) -> Result<tsd::ImageArchive> {
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers image {identifier} is missing")))?;
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers image {identifier} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {identifier} must have exactly one image payload"
        )));
    };
    tsd::ImageArchive::decode(message.data.as_slice()).map_err(Into::into)
}

fn image_caption_slot_from_reference(
    package: &IWorkPackage,
    reference: Option<tsp::Reference>,
    kind: DrawableCaptionKind,
) -> Result<ImageCaptionSlot> {
    let reference_id = reference
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers image has no title/caption reference".to_owned())
        })?
        .identifier;
    let locations = object_locations(package)?;
    let archive_name = locations.get(&reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers title/caption object {reference_id} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers title/caption object {reference_id} is missing"
        ))
    })?;
    if object
        .messages
        .iter()
        .any(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
    {
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers title/caption object {reference_id} repeats its caption payload"
            )));
        };
        let info = crate::protobuf::tsa::CaptionInfoArchive::decode(message.data.as_slice())?;
        if info.child_info_kind != Some(kind.native_kind()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers title/caption object {reference_id} has the wrong native kind"
            )));
        }
        let storage_id = info
            .super_
            .owned_storage
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} has no text storage"
                ))
            })?
            .identifier;
        let style_id = info
            .super_
            .super_
            .style
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} has no shape style"
                ))
            })?
            .identifier;
        let placement_id = info
            .placement
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} has no placement"
                ))
            })?
            .identifier;
        for (identifier, message_type, type_name) in [
            (style_id, SHAPE_STYLE_MESSAGE_TYPE, "TSWP.ShapeStyleArchive"),
            (storage_id, STORAGE_MESSAGE_TYPE, "TSWP.StorageArchive"),
            (
                placement_id,
                CAPTION_PLACEMENT_MESSAGE_TYPE,
                "TSA.CaptionPlacementArchive",
            ),
        ] {
            let Some(object_archive_name) = locations.get(&identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} references missing {type_name} {identifier}"
                )));
            };
            let archive = package.archive(object_archive_name)?;
            let object = archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} references missing {type_name} {identifier}"
                ))
            })?;
            if object
                .messages
                .iter()
                .filter(|message| message.type_ == message_type)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers title/caption object {reference_id} references malformed {type_name} {identifier}"
                )));
            }
        }
        return Ok(ImageCaptionSlot {
            reference_id,
            storage_id: Some(storage_id),
            object_ids: vec![reference_id, style_id, storage_id, placement_id],
        });
    }
    if object
        .messages
        .iter()
        .filter(|message| message.type_ == STANDIN_CAPTION_MESSAGE_TYPE)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers title/caption stand-in {reference_id} is malformed"
        )));
    }
    Ok(ImageCaptionSlot {
        reference_id,
        storage_id: None,
        object_ids: vec![reference_id],
    })
}

pub(super) fn insert_image_caption(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    old_reference_id: u64,
    image_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
    ids: CaptionObjectIds,
) -> Result<()> {
    let objects = caption_objects(ids, image_id, image_width, text, kind, theme, language)?;
    package.update_archive(archive_name, |archive| {
        for object in objects {
            archive.insert_object(object)?;
        }
        replace_image_caption_reference(archive, image_id, old_reference_id, ids.info, kind)
    })
}

pub(super) fn insert_image_caption_standin(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    old_reference_id: u64,
    kind: DrawableCaptionKind,
    standin_id: u64,
) -> Result<()> {
    let standin = standin_caption_object(standin_id)?;
    package.update_archive(archive_name, |archive| {
        archive.insert_object(standin)?;
        replace_image_caption_reference(archive, image_id, old_reference_id, standin_id, kind)
    })
}

fn replace_image_caption_reference(
    archive: &mut crate::archive::Archive,
    image_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let object = archive.object_mut(image_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers image object {image_id} is missing"))
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
            "Numbers image {image_id} must have exactly one ImageArchive payload"
        )));
    };
    let original = object.messages[*message_index].data.as_slice();
    let current = tsd::ImageArchive::decode(original)?;
    let current_reference_id = match kind {
        DrawableCaptionKind::Caption => current.super_.caption,
        DrawableCaptionKind::Title => current.super_.title,
    }
    .map(|reference| reference.identifier);
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {image_id} title/caption reference changed unexpectedly"
        )));
    }
    let data = transform_length_delimited_field(original, 1, |drawable| {
        patch_drawable_caption_reference(drawable, kind, replacement_id)
    })?;
    object.replace_message(
        *message_index,
        RawMessage {
            type_: IMAGE_MESSAGE_TYPE,
            data,
        },
    )?;
    replace_object_reference(
        &mut object.archive_info.message_infos[*message_index].object_references,
        old_reference_id,
        replacement_id,
    );
    Ok(())
}

fn image_info(
    package: &IWorkPackage,
    sheet_id: u64,
    identifier: u64,
) -> Result<NumbersSheetImageInfo> {
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers image {identifier} is missing")))?;
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers image {identifier} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {identifier} must have exactly one image payload"
        )));
    };
    let image = tsd::ImageArchive::decode(message.data.as_slice())?;
    if image.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers image {identifier} is not owned by sheet {sheet_id}"
        )));
    }
    let image_data_identifier = image
        .data
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers image {identifier} has no primary data reference"
            ))
        })?
        .identifier;
    Ok(NumbersSheetImageInfo {
        sheet_id,
        drawable_object_id: identifier,
        image_data_identifier,
        thumbnail_data_identifier: image.thumbnail_data.map(|reference| reference.identifier),
        geometry: geometry_from_drawable(&image.super_)?,
        properties: drawable_properties(&image.super_),
        image_adjustments: image_adjustments_from_archive(&image)?,
        original_size: image.original_size.map(drawable_size),
        natural_size: image.natural_size.map(drawable_size),
    })
}

pub(super) fn image_objects(
    ids: ImageObjectIds,
    sheet_id: u64,
    style_id: u64,
    data_identifier: u64,
    geometry: DrawableGeometry,
    natural_size: DrawableSize,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers image geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers image geometry has no size".to_owned())
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
            parent: Some(reference(sheet_id)),
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
            width: natural_size.width,
            height: natural_size.height,
        }),
        natural_size: Some(tsp::Size {
            width: natural_size.width,
            height: natural_size.height,
        }),
        flags: Some(DEFAULT_IMAGE_FLAGS),
        ..Default::default()
    };
    Ok([
        numbers_image_object(
            ids.drawable,
            IMAGE_MESSAGE_TYPE,
            image,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[data_identifier],
        )?,
        numbers_image_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        numbers_image_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
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
            Error::InvalidFormat(format!("Numbers image object {image_id} is missing"))
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
                "Numbers image {image_id} must have exactly one ImageArchive payload"
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

pub(super) fn set_image_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    properties: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(image_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers image object {image_id} is missing"))
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
                "Numbers image {image_id} must have exactly one ImageArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::ImageArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::ImageArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Numbers image properties patch failed validation".to_owned(),
            ));
        }
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

fn numbers_image_object(
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

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
