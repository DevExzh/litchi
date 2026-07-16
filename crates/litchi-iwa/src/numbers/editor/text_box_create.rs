//! Native Numbers text-box construction without a source drawable.

use super::*;
use crate::IWorkThemeArchive;
use crate::shapes::{DrawablePoint, DrawableSize, ShapePreset, shape_path_source};

const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;
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

    pub(super) const fn all(self) -> [u64; 4] {
        [self.drawable, self.caption, self.title, self.storage]
    }

    pub(super) const fn last(self) -> u64 {
        self.storage
    }
}

pub(super) struct TextBoxThemeStyles {
    pub(super) shape: u64,
    stylesheet: tsp::Reference,
    paragraph: Option<tsp::Reference>,
    list: Option<tsp::Reference>,
}

impl NumbersEditor {
    /// Add an independently editable text box to a reachable sheet.
    ///
    /// The shape, writable storage, title/caption stand-ins, sheet ownership,
    /// UUID metadata, and package high-water mark are encoded from typed
    /// values. No source drawable or package template is copied. Position and
    /// size are measured in sheet points.
    pub fn add_sheet_text_box(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersTextBoxInfo> {
        let geometry = validate_text_box_geometry(position, size)?;
        let (archive_name, _, _) = numbers_sheet(&self.package, sheet_id)?;
        let document = numbers_document(&self.package)?;
        let styles = text_box_theme_styles(
            &self.package,
            document.theme.identifier,
            document.stylesheet.identifier,
        )?;
        let storage = text_box_storage(text, &styles);
        let ids = TextBoxObjectIds::allocate(next_object_identifier(&self.package)?)?;
        let objects = text_box_objects(
            ids,
            sheet_id,
            styles.shape,
            geometry,
            storage,
            shape_path_source(ShapePreset::Rectangle, size)?,
            true,
        )?;

        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &archive_name,
            sheet_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, DOCUMENT_COMPONENT_IDENTIFIER, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers text-box creation failed validation".to_owned())
            })?;
        let created_graph = numbers_text_box_graph(verified.package(), sheet_id, ids.drawable)?;
        if created.storage.object_id != ids.storage
            || created.storage.text != text
            || created_graph.object_ids != ids.all()
            || verified.sheet_text_box_geometry(sheet_id, ids.drawable)? != geometry
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box creation produced an inconsistent graph".to_owned(),
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
            "Numbers text-box position must be finite and size must be finite and positive"
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
    package: &IWorkPackage,
    theme_id: u64,
    stylesheet_id: u64,
) -> Result<TextBoxThemeStyles> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE);
    let message = messages.next().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers theme object {theme_id} has no theme payload"
        ))
    })?;
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme object {theme_id} repeats its theme payload"
        )));
    }
    let theme = IWorkThemeArchive::decode(&message.data)?;
    let drawing = theme.extensions.drawing.ok_or_else(|| {
        Error::InvalidFormat("Numbers theme has no drawing style presets".to_owned())
    })?;
    let shape = drawing
        .textbox_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers theme has no text-box style preset".to_owned())
        })?;
    let text = theme.extensions.text.unwrap_or_default();
    let styles = TextBoxThemeStyles {
        shape,
        stylesheet: reference(stylesheet_id),
        paragraph: text.paragraph_style_presets.into_iter().next(),
        list: text.list_style_presets.into_iter().next(),
    };
    for (identifier, label) in [
        (Some(styles.shape), "text-box style"),
        (Some(styles.stylesheet.identifier), "stylesheet"),
        (
            styles.paragraph.map(|reference| reference.identifier),
            "paragraph style",
        ),
        (
            styles.list.map(|reference| reference.identifier),
            "list style",
        ),
    ] {
        if let Some(identifier) = identifier
            && !locations.contains_key(&identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers theme {label} object {identifier} is missing"
            )));
        }
    }
    Ok(styles)
}

pub(super) fn text_box_storage(text: &str, styles: &TextBoxThemeStyles) -> tswp::StorageArchive {
    tswp::StorageArchive {
        style_sheet: Some(styles.stylesheet),
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(styles.paragraph)),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_attribute_table(styles.list)),
        table_para_starts: Some(zero_para_data()),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_attribute_table(None)),
        ..Default::default()
    }
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
    sheet_id: u64,
    style_id: u64,
    geometry: DrawableGeometry,
    storage: tswp::StorageArchive,
    path_source: tsd::PathSourceArchive,
    is_text_box: bool,
) -> Result<[ArchiveObject; 4]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers text-box geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers text-box geometry has no size".to_owned())
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
        numbers_object(
            ids.drawable,
            SHAPE_INFO_MESSAGE_TYPE,
            shape,
            &STANDARD_MESSAGE_VERSION,
            &[ids.caption, ids.title, style_id, ids.storage],
        )?,
        numbers_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        numbers_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
        )?,
        numbers_object(
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
    for table in [
        &storage.table_para_style,
        &storage.table_list_style,
        &storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        references.extend(
            table
                .entries
                .iter()
                .filter_map(|entry| entry.object.map(|reference| reference.identifier)),
        );
    }
    references.sort_unstable();
    references.dedup();
    references
}

fn numbers_object(
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

    const MISSING_SHEET_ID: u64 = 999;
    const FIRST_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 90.0 };
    const FIRST_SIZE: DrawableSize = DrawableSize {
        width: 260.0,
        height: 72.0,
    };

    #[test]
    fn scratch_spreadsheet_supports_text_box_crud_without_a_source_drawable() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        assert!(editor.sheet_text_boxes(sheet_id).unwrap().is_empty());
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_sheet_text_box(
                sheet_id,
                "Built from typed objects",
                FIRST_POSITION,
                FIRST_SIZE,
            )
            .unwrap();
        assert_eq!(created.storage.text, "Built from typed objects");
        assert_eq!(
            editor
                .sheet_text_box_geometry(sheet_id, created.drawable_object_id)
                .unwrap(),
            DrawableGeometry {
                position: Some(FIRST_POSITION),
                size: Some(FIRST_SIZE),
                flags: Some(DEFAULT_DRAWABLE_FLAGS),
                angle: Some(DEFAULT_TEXT_BOX_ROTATION_DEGREES),
            }
        );

        editor
            .set_sheet_text_box_text(
                sheet_id,
                created.drawable_object_id,
                "Updated independently",
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_text_box(sheet_id, created.drawable_object_id, "Independent copy")
            .unwrap();
        assert_ne!(duplicate.storage.object_id, created.storage.object_id);

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let text_boxes = reopened.sheet_text_boxes(sheet_id).unwrap();
        assert_eq!(text_boxes.len(), 2);
        assert_eq!(text_boxes[0].storage.text, "Updated independently");
        assert_eq!(text_boxes[1].storage.text, "Independent copy");

        editor
            .remove_sheet_text_box(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        let removed = editor
            .remove_sheet_text_box(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.text_box.storage.text, "Updated independently");
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_text_box_creation_is_transactional() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_sheet_text_box(
                    MISSING_SHEET_ID,
                    "Missing sheet",
                    FIRST_POSITION,
                    FIRST_SIZE,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_sheet_text_box(
                    sheet_id,
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
    fn created_text_box_graph_has_private_objects_and_document_uuids() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_text_box(sheet_id, "Metadata", FIRST_POSITION, FIRST_SIZE)
            .unwrap();
        let graph =
            numbers_text_box_graph(editor.package(), sheet_id, created.drawable_object_id).unwrap();

        assert_eq!(graph.object_ids.len(), 4);
        assert_eq!(graph.uuid_object_ids, graph.object_ids);
        assert_eq!(graph.drawable_id, created.drawable_object_id);
        assert_eq!(graph.storage_id, created.storage.object_id);

        let archive = editor.package().archive(&graph.archive_name).unwrap();
        let storage = tswp::StorageArchive::decode(
            archive.object(graph.storage_id).unwrap().messages[0]
                .data
                .as_slice(),
        )
        .unwrap();
        let drop_cap_entries = &storage.table_drop_cap_style.unwrap().entries;
        assert_eq!(drop_cap_entries.len(), 1);
        assert!(drop_cap_entries[0].object.is_none());
    }
}
