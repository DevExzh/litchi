//! Native Pages text-box construction without a source drawable.

use super::*;
use crate::IWorkThemeArchive;
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

const DRAWABLE_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_TEXT_BOX_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const NORMALIZED_RECTANGLE_EXTENT: f32 = 100.0;
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

#[derive(Debug, Clone, Copy)]
struct TextBoxObjectIds {
    drawable: u64,
    caption: u64,
    title: u64,
    storage: u64,
    attachment: u64,
}

impl TextBoxObjectIds {
    fn allocate(first: u64) -> Result<Self> {
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
            attachment: identifier(4)?,
        })
    }

    fn last(self) -> u64 {
        self.attachment
    }

    fn uuid_objects(self) -> [u64; 4] {
        [self.drawable, self.caption, self.title, self.storage]
    }
}

impl PagesEditor {
    /// Add a body-anchored text box without requiring a source drawable.
    ///
    /// The text box, writable storage, title/caption stand-ins, body
    /// attachment, drawable z-order, and package UUID metadata are constructed
    /// from typed values. Position and size are measured in document points.
    pub fn add_text_box(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<PagesDrawableTextInfo> {
        let geometry = validate_text_box_geometry(position, size)?;
        let root = root_document(self.package())?;
        let body: StorageArchive = decode_typed_package_object(
            self.package(),
            self.body_storage_id,
            self.body_storage()?.message_type,
            "TSWP.StorageArchive",
        )?;
        let style_id = text_box_style_id(self.package(), &root)?;
        let storage = text_box_storage(text, &body);
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = TextBoxObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;

        let mut staged = self.package().clone();
        if creates_z_order {
            create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        let objects = text_box_objects(
            ids,
            self.body_storage_id,
            style_id,
            geometry,
            storage,
            root.left_margin.unwrap_or_default(),
        )?;
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;

        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            ids.attachment,
        )?;
        patch_pages_zorder(&mut staged, None, Some(ids.drawable))?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &ids.uuid_objects())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .drawable_text_storages()?
            .into_iter()
            .find(|item| item.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages text-box creation failed validation".to_owned())
            })?;
        let graph = verified.text_box_graph(ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.storage.object_id != ids.storage
            || created.storage.text != text
            || graph.anchor_character_index != expected_anchor
            || verified.text_box_geometry(ids.drawable)? != geometry
        {
            return Err(Error::InvalidFormat(
                "Pages text-box creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }
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
            "Pages text-box position must be finite and size must be finite and positive"
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

fn text_box_style_id(package: &IWorkPackage, root: &DocumentArchive) -> Result<u64> {
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
        .find(|message| message.type_ == 10_001)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Pages theme {theme_id} has no theme payload"))
        })?;
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.textbox_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no text-box style preset".to_owned()))
}

fn text_box_storage(text: &str, body: &StorageArchive) -> StorageArchive {
    let paragraph_style = first_object_attribute(&body.table_para_style);
    let list_style = first_object_attribute(&body.table_list_style);
    let language = body.table_language.as_ref().and_then(|table| {
        table
            .entries
            .iter()
            .find_map(|entry| entry.object.as_ref().cloned())
    });
    StorageArchive {
        style_sheet: body.style_sheet,
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(paragraph_style)),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_attribute_table(list_style)),
        table_para_starts: Some(zero_para_data()),
        table_language: language.map(|language| crate::protobuf::tswp::StringAttributeTable {
            entries: vec![
                crate::protobuf::tswp::string_attribute_table::StringAttribute {
                    character_index: 0,
                    object: Some(language),
                },
            ],
        }),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_attribute_table(None)),
        ..Default::default()
    }
}

fn first_object_attribute(
    table: &Option<crate::protobuf::tswp::ObjectAttributeTable>,
) -> Option<tsp::Reference> {
    table
        .as_ref()
        .and_then(|table| table.entries.iter().find_map(|entry| entry.object))
}

fn object_attribute_table(
    object: Option<tsp::Reference>,
) -> crate::protobuf::tswp::ObjectAttributeTable {
    crate::protobuf::tswp::ObjectAttributeTable {
        entries: vec![ObjectAttribute {
            character_index: 0,
            object,
        }],
    }
}

fn zero_para_data() -> crate::protobuf::tswp::ParaDataAttributeTable {
    crate::protobuf::tswp::ParaDataAttributeTable {
        entries: vec![
            crate::protobuf::tswp::para_data_attribute_table::ParaDataAttribute {
                character_index: 0,
                first: 0,
                second: 0,
            },
        ],
    }
}

#[allow(deprecated)]
fn text_box_objects(
    ids: TextBoxObjectIds,
    body_storage_id: u64,
    style_id: u64,
    geometry: DrawableGeometry,
    storage: StorageArchive,
    left_margin: f32,
) -> Result<[ArchiveObject; 5]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated text-box geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated text-box geometry has no size".to_owned())
    })?;
    let storage_references = storage_references(&storage);
    let shape = crate::protobuf::tswp::ShapeInfoArchive {
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
            style: Some(reference(style_id)),
            pathsource: Some(rectangle_path_source(size)),
            stroke_pattern_offset_distance: Some(0.0),
            ..Default::default()
        },
        deprecated_storage: Some(reference(ids.storage)),
        owned_storage: Some(reference(ids.storage)),
        is_text_box: Some(true),
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
        pages_object(
            ids.drawable,
            SHAPE_INFO_MESSAGE_TYPE,
            shape,
            &STANDARD_MESSAGE_VERSION,
            &[ids.caption, ids.title, style_id, ids.storage],
        )?,
        pages_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        pages_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        pages_object(
            ids.storage,
            STORAGE_MESSAGE_TYPES[0],
            storage,
            &STANDARD_MESSAGE_VERSION,
            &storage_references,
        )?,
        pages_object(
            ids.attachment,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            attachment,
            &STANDARD_MESSAGE_VERSION,
            &[ids.drawable],
        )?,
    ])
}

fn rectangle_path_source(size: DrawableSize) -> tsd::PathSourceArchive {
    use tsp::path::{Element, ElementType};

    let point = |x, y| tsp::Point { x, y };
    let element = |r#type: ElementType, points| Element {
        r#type: r#type as i32,
        points,
    };
    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size {
                width: size.width,
                height: size.height,
            }),
            path: Some(tsp::Path {
                elements: vec![
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                    element(
                        ElementType::LineTo,
                        vec![point(NORMALIZED_RECTANGLE_EXTENT, 0.0)],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(
                            NORMALIZED_RECTANGLE_EXTENT,
                            NORMALIZED_RECTANGLE_EXTENT,
                        )],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(0.0, NORMALIZED_RECTANGLE_EXTENT)],
                    ),
                    element(ElementType::CloseSubpath, Vec::new()),
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn storage_references(storage: &StorageArchive) -> Vec<u64> {
    let mut references = Vec::with_capacity(3);
    references.extend(storage.style_sheet.map(|reference| reference.identifier));
    references
        .extend(first_object_attribute(&storage.table_para_style).map(|item| item.identifier));
    references
        .extend(first_object_attribute(&storage.table_list_style).map(|item| item.identifier));
    references.sort_unstable();
    references.dedup();
    references
}

fn pages_object(
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

fn create_drawable_z_order(
    package: &mut IWorkPackage,
    archive_name: &str,
    z_order_id: u64,
) -> Result<()> {
    let z_order = pages_object(
        z_order_id,
        DRAWABLE_Z_ORDER_MESSAGE_TYPE,
        tp::DrawablesZOrderArchive::default(),
        &STANDARD_MESSAGE_VERSION,
        &[],
    )?;
    package.update_archive(archive_name, |archive| {
        archive.insert_object(z_order)?;
        let root = archive.object_mut(DOCUMENT_OBJECT_ID).ok_or_else(|| {
            Error::InvalidFormat("Pages document root object is missing".to_owned())
        })?;
        let indexes = root
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == DOCUMENT_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(
                "Pages document root must have exactly one document payload".to_owned(),
            ));
        }
        let message_index = indexes[0];
        let original = &root.messages[message_index];
        let previous = DocumentArchive::decode(original.data.as_slice())?;
        if previous.drawables_zorder.is_some() {
            return Err(Error::InvalidFormat(
                "Pages document acquired a drawable z-order during text-box creation".to_owned(),
            ));
        }
        let data = patch_length_delimited_field(
            &original.data,
            20,
            false,
            Some(&reference(z_order_id).encode_to_vec()),
        )?;
        if DocumentArchive::decode(data.as_slice())?
            .drawables_zorder
            .map(|reference| reference.identifier)
            != Some(z_order_id)
        {
            return Err(Error::InvalidFormat(
                "Pages drawable z-order link failed validation".to_owned(),
            ));
        }
        root.replace_message(
            message_index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data,
            },
        )?;
        update_reference_list(
            &mut root.archive_info.message_infos[message_index].object_references,
            None,
            Some(z_order_id),
        );
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIRST_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const FIRST_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 72.0,
    };

    #[test]
    fn scratch_document_supports_text_box_crud_without_a_source_drawable() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        assert!(editor.drawables().unwrap().is_empty());

        let created = editor
            .add_text_box(4, "Typed from scratch", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        let drawable_id = created.drawable_object_id;
        assert_eq!(created.storage.text, "Typed from scratch");
        assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}");
        assert_eq!(
            editor.text_box_geometry(drawable_id).unwrap(),
            DrawableGeometry {
                position: Some(FIRST_POSITION),
                size: Some(FIRST_SIZE),
                flags: Some(DEFAULT_DRAWABLE_FLAGS),
                angle: Some(DEFAULT_TEXT_BOX_ROTATION_DEGREES),
            }
        );

        editor
            .set_drawable_text(drawable_id, "Updated independently")
            .unwrap();
        let duplicate = editor
            .duplicate_text_box(drawable_id, 0, "Independent copy")
            .unwrap();
        assert_ne!(duplicate.storage.object_id, created.storage.object_id);
        assert_eq!(editor.body_text().unwrap(), "\u{fffc}Body\u{fffc}");

        let removed_copy = editor
            .remove_text_box(duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(removed_copy.text.storage.text, "Independent copy");
        let removed = editor.remove_text_box(drawable_id).unwrap();
        assert_eq!(removed.text.storage.text, "Updated independently");
        assert_eq!(editor.body_text().unwrap(), "Body");
        assert!(editor.drawables().unwrap().is_empty());
    }

    #[test]
    fn multiple_created_text_boxes_share_one_z_order_and_round_trip() {
        let mut editor = PagesEditor::create_with_text("Report").unwrap();
        let first = editor
            .add_text_box(6, "First", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        let second_position = DrawablePoint { x: 96.0, y: 240.0 };
        let second = editor
            .add_text_box(0, "Second", second_position, FIRST_SIZE)
            .unwrap();

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let root = root_document(reopened.package()).unwrap();
        let z_order_id = root.drawables_zorder.unwrap().identifier;
        let z_order: tp::DrawablesZOrderArchive = decode_typed_package_object(
            reopened.package(),
            z_order_id,
            DRAWABLE_Z_ORDER_MESSAGE_TYPE,
            "TP.DrawablesZOrderArchive",
        )
        .unwrap();
        assert_eq!(
            z_order
                .drawables
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>(),
            [first.drawable_object_id, second.drawable_object_id]
        );
        assert_eq!(
            reopened
                .text_box_geometry(second.drawable_object_id)
                .unwrap(),
            DrawableGeometry {
                position: Some(second_position),
                size: Some(FIRST_SIZE),
                flags: Some(DEFAULT_DRAWABLE_FLAGS),
                angle: Some(DEFAULT_TEXT_BOX_ROTATION_DEGREES),
            }
        );
    }

    #[test]
    fn invalid_text_box_creation_is_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_text_box(5, "Out of range", FIRST_POSITION, FIRST_SIZE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_text_box(
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
}
