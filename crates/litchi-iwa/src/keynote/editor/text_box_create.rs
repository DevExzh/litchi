//! Native Keynote text-box construction without a source drawable.

use super::*;
use crate::IWorkThemeArchive;
use crate::shapes::{DrawablePoint, DrawableSize, ShapePreset, shape_path_source};

const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_TEXT_BOX_ROTATION_DEGREES: f32 = 0.0;
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
pub(super) struct TextBoxObjectIds {
    pub(super) drawable: u64,
    pub(super) caption: u64,
    pub(super) title: u64,
    pub(super) storage: u64,
}

pub(super) struct TextBoxContext {
    pub(super) slide_id: u64,
    pub(super) theme_id: u64,
    pub(super) stylesheet_id: u64,
    pub(super) language: Option<String>,
    pub(super) slide: kn::SlideArchive,
}

pub(super) struct TextBoxThemeStyles {
    pub(super) shape: u64,
    stylesheet: tsp::Reference,
    paragraph: Option<tsp::Reference>,
    list: Option<tsp::Reference>,
}

impl TextBoxObjectIds {
    pub(super) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            drawable: identifier(0)?,
            caption: identifier(1)?,
            title: identifier(2)?,
            storage: identifier(3)?,
        })
    }

    pub(super) const fn last(self) -> u64 {
        self.storage
    }

    pub(super) const fn all(self) -> [u64; 4] {
        [self.drawable, self.caption, self.title, self.storage]
    }
}

impl KeynoteEditor {
    /// Add an independently editable text box to a slide.
    ///
    /// The shape, writable storage, title/caption stand-ins, drawable
    /// ownership, z-order, UUID metadata, and package high-water mark are
    /// constructed directly from typed values. No source drawable or package
    /// template is copied. Position and size are measured in slide points.
    pub fn add_slide_text_box(
        &mut self,
        slide_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<KeynoteSlideTextInfo> {
        let geometry = validate_text_box_geometry(position, size)?;
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_context(&graph, slide_index)?;
        let styles = text_box_theme_styles(&graph, context.theme_id, context.stylesheet_id)?;
        let base_storage = slide_text_storage_template(&graph, &context.slide)?;
        let storage = text_box_storage(
            text,
            base_storage.as_ref(),
            &styles,
            context.language.as_deref(),
        );
        let ids = TextBoxObjectIds::allocate(next_object_identifier(self.package())?)?;
        let archive_name = graph.archive_name(context.slide_id)?.to_owned();
        let objects = text_box_objects(
            ids,
            context.slide_id,
            styles.shape,
            geometry,
            storage,
            shape_path_source(ShapePreset::Rectangle, size)?,
            true,
        )?;

        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_slide_drawable_references(
            &mut staged,
            &archive_name,
            context.slide_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, context.slide_id, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .slide_text_storages(slide_index)?
            .into_iter()
            .find(|item| item.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote text-box creation failed validation".to_owned())
            })?;
        let created_graph = verified.text_box_graph(slide_index, ids.drawable)?;
        if created.role != KeynoteSlideTextRole::TextBox
            || created.storage.object_id != ids.storage
            || created.storage.text != text
            || created_graph.object_ids != ids.all()
            || verified.slide_text_box_geometry(slide_index, ids.drawable)? != geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote text-box creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }
}

pub(super) fn text_box_context(graph: &ObjectGraph, slide_index: usize) -> Result<TextBoxContext> {
    let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
    let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
    let node_reference = show.slide_tree.slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            show.slide_tree.slides.len()
        ))
    })?;
    let node: kn::SlideNodeArchive =
        graph.decode(node_reference.identifier, "KN.SlideNodeArchive")?;
    let slide_id = node
        .slide
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide node {} has no slide reference",
                node_reference.identifier
            ))
        })?
        .identifier;
    Ok(TextBoxContext {
        slide_id,
        theme_id: show.theme.identifier,
        stylesheet_id: show.stylesheet.identifier,
        language: document.super_.document_language,
        slide: graph.decode_type(slide_id, 5, "KN.SlideArchive")?,
    })
}

fn validate_text_box_geometry(
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
            "Keynote text-box position must be finite and size must be finite and positive"
                .to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_TEXT_BOX_ROTATION_DEGREES),
    }
    .validate()
}

pub(super) fn text_box_theme_styles(
    graph: &ObjectGraph,
    theme_id: u64,
    stylesheet_id: u64,
) -> Result<TextBoxThemeStyles> {
    let theme_data = graph.message_data_type(theme_id, 10, "KN.ThemeArchive")?;
    let theme = IWorkThemeArchive::decode(theme_data)?;
    let drawing = theme.extensions.drawing.ok_or_else(|| {
        Error::InvalidFormat("Keynote theme has no drawing style presets".to_owned())
    })?;
    let shape = drawing
        .textbox_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote theme has no text-box style preset".to_owned())
        })?;
    let text = theme.extensions.text.unwrap_or_default();
    Ok(TextBoxThemeStyles {
        shape,
        stylesheet: reference(stylesheet_id),
        paragraph: text.paragraph_style_presets.into_iter().next(),
        list: text.list_style_presets.into_iter().next(),
    })
}

pub(super) fn slide_text_storage_template(
    graph: &ObjectGraph,
    slide: &kn::SlideArchive,
) -> Result<Option<tswp::StorageArchive>> {
    for drawable in [
        slide.body_placeholder.as_ref(),
        slide.title_placeholder.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        if let Some(storage_id) = graph.drawable_storage(drawable.identifier)? {
            return graph.decode(storage_id, "TSWP.StorageArchive").map(Some);
        }
    }
    Ok(None)
}

pub(super) fn text_box_storage(
    text: &str,
    base: Option<&tswp::StorageArchive>,
    styles: &TextBoxThemeStyles,
    default_language: Option<&str>,
) -> tswp::StorageArchive {
    let paragraph_style = base
        .and_then(|storage| first_object_attribute(&storage.table_para_style))
        .or(styles.paragraph);
    let list_style = base
        .and_then(|storage| first_object_attribute(&storage.table_list_style))
        .or(styles.list);
    let drop_cap_style =
        base.and_then(|storage| first_object_attribute(&storage.table_drop_cap_style));
    let language = base
        .and_then(|storage| storage.table_language.as_ref())
        .and_then(|table| {
            table
                .entries
                .iter()
                .find_map(|entry| entry.object.as_ref().cloned())
        })
        .or_else(|| default_language.map(str::to_owned));
    tswp::StorageArchive {
        style_sheet: Some(
            base.and_then(|storage| storage.style_sheet)
                .unwrap_or(styles.stylesheet),
        ),
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(paragraph_style)),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_attribute_table(list_style)),
        table_para_starts: Some(zero_para_data()),
        table_language: language.map(|language| tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language),
            }],
        }),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_attribute_table(drop_cap_style)),
        ..Default::default()
    }
}

fn first_object_attribute(table: &Option<tswp::ObjectAttributeTable>) -> Option<tsp::Reference> {
    table
        .as_ref()
        .and_then(|table| table.entries.iter().find_map(|entry| entry.object))
}

fn object_attribute_table(object: Option<tsp::Reference>) -> tswp::ObjectAttributeTable {
    tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object,
        }],
    }
}

fn zero_para_data() -> tswp::ParaDataAttributeTable {
    tswp::ParaDataAttributeTable {
        entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        }],
    }
}

#[allow(deprecated)]
pub(super) fn text_box_objects(
    ids: TextBoxObjectIds,
    slide_id: u64,
    style_id: u64,
    geometry: DrawableGeometry,
    storage: tswp::StorageArchive,
    path_source: tsd::PathSourceArchive,
    is_text_box: bool,
) -> Result<[ArchiveObject; 4]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote text-box geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote text-box geometry has no size".to_owned())
    })?;
    let storage_references = storage_references(&storage);
    let shape = tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
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
            style: Some(reference(style_id)),
            pathsource: Some(path_source),
            stroke_pattern_offset_distance: Some(0.0),
            ..Default::default()
        },
        deprecated_storage: Some(reference(ids.storage)),
        owned_storage: Some(reference(ids.storage)),
        is_text_box: Some(is_text_box),
        ..Default::default()
    };
    Ok([
        keynote_object(
            ids.drawable,
            SHAPE_INFO_MESSAGE_TYPE,
            shape,
            &STANDARD_MESSAGE_VERSION,
            &[ids.caption, ids.title, style_id, ids.storage],
        )?,
        keynote_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        keynote_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        keynote_object(
            ids.storage,
            STORAGE_MESSAGE_TYPES[0],
            storage,
            &STANDARD_MESSAGE_VERSION,
            &storage_references,
        )?,
    ])
}

fn storage_references(storage: &tswp::StorageArchive) -> Vec<u64> {
    let mut references = Vec::with_capacity(4);
    references.extend(storage.style_sheet.map(|reference| reference.identifier));
    references
        .extend(first_object_attribute(&storage.table_para_style).map(|item| item.identifier));
    references
        .extend(first_object_attribute(&storage.table_list_style).map(|item| item.identifier));
    references
        .extend(first_object_attribute(&storage.table_drop_cap_style).map(|item| item.identifier));
    references.sort_unstable();
    references.dedup();
    references
}

fn keynote_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    versions: &[u32],
    references: &[u64],
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
    info.object_references = references.to_vec();
    Ok(object)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;

    const FIRST_POSITION: DrawablePoint = DrawablePoint { x: 144.0, y: 216.0 };
    const FIRST_SIZE: DrawableSize = DrawableSize {
        width: 640.0,
        height: 120.0,
    };

    #[test]
    fn scratch_presentation_supports_text_box_crud_without_a_source_drawable() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch presentation")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        assert!(
            editor
                .slide_text_storages(0)
                .unwrap()
                .iter()
                .all(|text| text.role != KeynoteSlideTextRole::TextBox)
        );
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_slide_text_box(0, "Built from typed objects", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        assert_eq!(created.role, KeynoteSlideTextRole::TextBox);
        assert_eq!(created.storage.text, "Built from typed objects");
        assert_eq!(
            editor
                .slide_text_box_geometry(0, created.drawable_object_id)
                .unwrap(),
            DrawableGeometry {
                position: Some(FIRST_POSITION),
                size: Some(FIRST_SIZE),
                flags: Some(DEFAULT_DRAWABLE_FLAGS),
                angle: Some(DEFAULT_TEXT_BOX_ROTATION_DEGREES),
            }
        );

        editor
            .set_slide_text_storage(0, created.drawable_object_id, "Updated independently")
            .unwrap();
        let duplicate = editor
            .duplicate_slide_text_box(0, created.drawable_object_id, "Independent copy")
            .unwrap();
        assert_ne!(duplicate.storage.object_id, created.storage.object_id);

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let text_boxes = reopened
            .slide_text_storages(0)
            .unwrap()
            .into_iter()
            .filter(|text| text.role == KeynoteSlideTextRole::TextBox)
            .collect::<Vec<_>>();
        assert_eq!(text_boxes.len(), 2);
        assert_eq!(text_boxes[0].storage.text, "Updated independently");
        assert_eq!(text_boxes[1].storage.text, "Independent copy");

        editor
            .remove_slide_text_box(0, duplicate.drawable_object_id)
            .unwrap();
        let removed = editor
            .remove_slide_text_box(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.text.storage.text, "Updated independently");
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_text_box_creation_is_transactional() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_slide_text_box(1, "Missing slide", FIRST_POSITION, FIRST_SIZE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_slide_text_box(
                    0,
                    "Invalid size",
                    FIRST_POSITION,
                    DrawableSize {
                        width: 0.0,
                        height: FIRST_SIZE.height,
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn slide_without_placeholders_uses_theme_text_styles() {
        let mut package = KeynoteDocumentBuilder::new()
            .build()
            .unwrap()
            .into_package();
        package
            .update_archive("Index/Slide-14.iwa", |archive| {
                let object = archive.object_mut(14).unwrap();
                let mut slide = kn::SlideArchive::decode(object.messages[0].data.as_slice())?;
                slide.title_placeholder = None;
                slide.body_placeholder = None;
                slide.owned_drawables.clear();
                slide.drawables_z_order.clear();
                object.replace_message(
                    0,
                    RawMessage {
                        type_: 5,
                        data: slide.encode_to_vec(),
                    },
                )?;
                let info = &mut object.archive_info.message_infos[0];
                info.object_references
                    .retain(|identifier| ![15, 16].contains(identifier));
                for field in &mut info.field_infos {
                    field
                        .object_references
                        .retain(|identifier| ![15, 16].contains(identifier));
                }
                Ok(())
            })
            .unwrap();
        let mut editor = KeynoteEditor::from_package(package).unwrap();

        let created = editor
            .add_slide_text_box(0, "On a truly blank slide", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        assert_eq!(created.storage.text, "On a truly blank slide");
        assert_eq!(created.role, KeynoteSlideTextRole::TextBox);
    }

    #[test]
    fn created_text_box_graph_has_private_objects_and_component_uuids() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let created = editor
            .add_slide_text_box(0, "Metadata", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        let graph = editor
            .text_box_graph(0, created.drawable_object_id)
            .unwrap();

        assert_eq!(graph.object_ids.len(), 4);
        assert_eq!(graph.uuid_object_ids, graph.object_ids);
        assert_eq!(graph.drawable_id, created.drawable_object_id);
        assert_eq!(graph.storage_id, created.storage.object_id);
    }
}
