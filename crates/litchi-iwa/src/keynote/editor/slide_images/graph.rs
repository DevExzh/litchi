//! Typed construction and discovery of slide-owned image graphs.

use super::*;
use crate::image_adjustments::image_adjustments_from_archive;
use crate::shapes::{
    drawable_properties, geometry_from_drawable, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MEDIA_STYLE_MESSAGE_TYPE: u32 = 3_016;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const LAYOUT_IMAGE_FLAG: u32 = 1;
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
    pub(super) slide_id: u64,
    pub(super) component_id: u64,
    pub(super) archive_name: String,
    pub(super) style_id: u64,
    pub(super) stylesheet_component_id: u64,
}

pub(super) struct SlideImageGraph {
    pub(super) slide_id: u64,
    pub(super) component_id: u64,
    pub(super) archive_name: String,
    pub(super) info: KeynoteSlideImageInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn image_geometry(
    position: DrawablePoint,
    size: DrawableSize,
) -> Result<DrawableGeometry> {
    if size.width <= 0.0 || size.height <= 0.0 {
        return Err(Error::ParseError(
            "Keynote image size must be greater than zero".to_owned(),
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

pub(super) fn image_creation_context(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<ImageCreationContext> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
    let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
    let stylesheet_id = show.stylesheet.identifier;
    let stylesheet: tss::StylesheetArchive = graph.decode_type(
        stylesheet_id,
        STYLESHEET_MESSAGE_TYPE,
        "TSS.StylesheetArchive",
    )?;
    let style_id = stylesheet
        .styles
        .iter()
        .map(|style| style.identifier)
        .find(|identifier| {
            graph.objects.get(identifier).is_some_and(|messages| {
                messages
                    .iter()
                    .any(|message| message.type_ == MEDIA_STYLE_MESSAGE_TYPE)
            })
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote stylesheet {stylesheet_id} has no media style"
            ))
        })?;
    let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {archive_name} is not registered"
            ))
        })?;
    let stylesheet_archive = graph.archive_name(stylesheet_id)?;
    let stylesheet_component_id =
        component_identifier_for_entry(editor.package(), stylesheet_archive)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote stylesheet component {stylesheet_archive} is not registered"
            ))
        })?;
    Ok(ImageCreationContext {
        slide_id: slide.slide_id,
        component_id,
        archive_name,
        style_id,
        stylesheet_component_id,
    })
}

pub(super) fn image_infos(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<Vec<KeynoteSlideImageInfo>> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    image_root_ids(&graph, slide.slide_id)?
        .into_iter()
        .map(|identifier| image_info(&graph, slide_index, identifier))
        .collect()
}

fn image_root_ids(graph: &ObjectGraph, slide_id: u64) -> Result<Vec<u64>> {
    let native: kn::SlideArchive =
        graph.decode_type(slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    Ok(native
        .owned_drawables
        .iter()
        .filter(|reference| {
            graph
                .objects
                .get(&reference.identifier)
                .is_some_and(|messages| {
                    messages
                        .iter()
                        .any(|message| message.type_ == IMAGE_MESSAGE_TYPE)
                })
        })
        .map(|reference| reference.identifier)
        .collect())
}

pub(super) fn require_file_image(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideImageGraph> {
    let graph = image_graph(editor, slide_index, drawable_object_id)?;
    if graph.info.kind != KeynoteSlideImageKind::File {
        return Err(Error::ParseError(format!(
            "Keynote image {drawable_object_id} is layout-owned, not an ordinary file-backed image"
        )));
    }
    Ok(graph)
}

pub(super) fn image_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideImageGraph> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    if !image_root_ids(&graph, slide.slide_id)?.contains(&drawable_object_id) {
        return Err(Error::ParseError(format!(
            "Keynote image {drawable_object_id} is not owned by slide {slide_index}"
        )));
    }
    let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
    if graph.archive_name(drawable_object_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Keynote image {drawable_object_id} is outside slide component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object_ids = slide_create::graph::private_clone_object_ids(
        &archive,
        [drawable_object_id],
        "slide image",
    )?;
    if object_ids.contains(&slide.slide_id) {
        return Err(Error::InvalidFormat(
            "Keynote image private graph reaches its owning slide".to_owned(),
        ));
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    let mut data_references = Vec::new();
    for identifier in &object_ids {
        let object = archive.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote image object {identifier} is missing"))
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
    Ok(SlideImageGraph {
        slide_id: slide.slide_id,
        component_id,
        archive_name,
        info: image_info(&graph, slide_index, drawable_object_id)?,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

fn image_info(
    graph: &ObjectGraph,
    slide_index: usize,
    identifier: u64,
) -> Result<KeynoteSlideImageInfo> {
    let image: tsd::ImageArchive =
        graph.decode_type(identifier, IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")?;
    let image_data_identifier = image
        .data
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote image {identifier} has no primary data reference"
            ))
        })?
        .identifier;
    Ok(KeynoteSlideImageInfo {
        slide_index,
        drawable_object_id: identifier,
        kind: if image
            .flags
            .is_some_and(|flags| flags & LAYOUT_IMAGE_FLAG != 0)
        {
            KeynoteSlideImageKind::Layout
        } else {
            KeynoteSlideImageKind::File
        },
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
    slide_id: u64,
    style_id: u64,
    data_identifier: u64,
    geometry: DrawableGeometry,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote image geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote image geometry has no size".to_owned())
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
            parent: Some(reference(slide_id)),
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
    Ok([
        keynote_object(
            ids.drawable,
            IMAGE_MESSAGE_TYPE,
            image,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[data_identifier],
        )?,
        keynote_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        keynote_object(
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
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(image_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote image object {image_id} is missing"))
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
                "Keynote image {image_id} must have exactly one ImageArchive payload"
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
            Error::InvalidFormat(format!("Keynote image object {image_id} is missing"))
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
                "Keynote image {image_id} must have exactly one ImageArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::ImageArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::ImageArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Keynote image properties patch failed validation".to_owned(),
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

fn keynote_object(
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
