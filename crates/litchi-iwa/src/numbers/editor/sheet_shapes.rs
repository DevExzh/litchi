//! Ordinary text-bearing shape CRUD for Numbers sheets.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use super::*;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, ShapePathKind, ShapePreset,
    set_shape_preset, shape_path_kind, shape_preset,
};

use super::text_box_create::{
    TextBoxObjectIds, text_box_objects, text_box_storage, text_box_theme_styles,
};

const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;

/// Structural path family used by an ordinary Numbers shape.
pub type NumbersSheetShapeKind = ShapePathKind;

/// One ordinary, non-text-box shape owned directly by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetShapeInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub kind: NumbersSheetShapeKind,
    /// Source-buildable preset and its native controls, when recognized.
    pub preset: Option<ShapePreset>,
    pub storage: TextStorageInfo,
    pub geometry: DrawableGeometry,
    pub properties: DrawableProperties,
}

/// Result of removing an ordinary Numbers shape and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetShape {
    pub shape: NumbersSheetShapeInfo,
}

struct SheetShapeGraph {
    sheet_id: u64,
    archive_name: String,
    component_id: u64,
    info: NumbersSheetShapeInfo,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

impl NumbersEditor {
    /// List ordinary non-text-box shapes owned by one reachable sheet.
    pub fn sheet_shapes(&self, sheet_id: u64) -> Result<Vec<NumbersSheetShapeInfo>> {
        shape_infos(self, sheet_id)
    }

    /// Add a rectangular shape with independent writable text to one sheet.
    ///
    /// The shape, storage, title/caption stand-ins, sheet ownership, UUIDs,
    /// style relationship, and package high-water mark are built directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_rectangle(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersSheetShapeInfo> {
        self.add_sheet_shape(sheet_id, text, position, size, ShapePreset::Rectangle)
    }

    /// Add a typed preset shape with independent writable text to one sheet.
    ///
    /// The path, storage, title/caption stand-ins, sheet ownership, UUIDs,
    /// style relationship, and package high-water mark are built directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_shape(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: ShapePreset,
    ) -> Result<NumbersSheetShapeInfo> {
        let geometry = new_shape_geometry(position, size)?;
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
            preset,
            false,
        )?;
        let component_id = component_identifier_for_entry(&self.package, &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet component {archive_name} is not registered"
                ))
            })?;
        let style_archive = object_locations(&self.package)?
            .get(&styles.shape)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers shape style {} is missing", styles.shape))
            })?;
        let style_component_id = component_identifier_for_entry(&self.package, &style_archive)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers stylesheet component {style_archive} is not registered"
                ))
            })?;

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
        add_component_object_uuids(&mut staged, component_id, &ids.all())?;
        if component_id != style_component_id {
            add_component_external_reference(
                &mut staged,
                component_id,
                style_component_id,
                styles.shape,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_shapes(sheet_id)?
            .into_iter()
            .find(|shape| shape.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers shape creation failed validation".to_owned())
            })?;
        let created_graph = shape_graph(&verified, sheet_id, ids.drawable)?;
        if created.preset != Some(preset)
            || created.storage.object_id != ids.storage
            || created.storage.text != text
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
        {
            return Err(Error::InvalidFormat(
                "Numbers shape creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the recognized preset and native controls for one sheet shape.
    pub fn sheet_shape_preset(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<ShapePreset>> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?.info.preset)
    }

    /// Replace a sheet shape's preset path while retaining its text and style.
    pub fn set_sheet_shape_preset(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        preset: ShapePreset,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_preset(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            preset,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_preset(sheet_id, drawable_object_id)? != Some(preset) {
            return Err(Error::InvalidFormat(
                "Numbers shape preset update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read typed geometry for one ordinary sheet shape.
    pub fn sheet_shape_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update shape geometry while preserving unknown path fields.
    pub fn set_sheet_shape_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers shape geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary sheet shape.
    pub fn sheet_shape_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?
            .info
            .properties)
    }

    /// Update shared drawable properties while preserving unknown fields.
    pub fn set_sheet_shape_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers shape property update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a UTF-16 range in an ordinary sheet shape's text.
    pub fn replace_sheet_shape_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.replace_text(source.info.storage.object_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        shape_graph(&verified, sheet_id, drawable_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Replace all text in an ordinary sheet shape.
    pub fn set_sheet_shape_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text(source.info.storage.object_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = shape_graph(&verified, sheet_id, drawable_object_id)?;
        if updated.info.storage.text != replacement {
            return Err(Error::InvalidFormat(
                "Numbers shape text update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear an ordinary sheet shape without deleting it.
    pub fn clear_sheet_shape_text(&mut self, sheet_id: u64, drawable_object_id: u64) -> Result<()> {
        self.set_sheet_shape_text(sheet_id, drawable_object_id, "")
    }

    /// Remove an ordinary shape and its private text-bearing graph.
    pub fn remove_sheet_shape(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetShape> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            source.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers shape object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers shape object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .sheet_shapes(sheet_id)?
            .iter()
            .any(|shape| shape.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers shape deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetShape { shape: source.info })
    }
}

fn new_shape_geometry(position: DrawablePoint, size: DrawableSize) -> Result<DrawableGeometry> {
    if !position.x.is_finite()
        || !position.y.is_finite()
        || !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Numbers shape position must be finite and size must be finite and positive".to_owned(),
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

fn shape_infos(editor: &NumbersEditor, sheet_id: u64) -> Result<Vec<NumbersSheetShapeInfo>> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut shapes = Vec::new();
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
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {} has multiple shape payloads",
                reference.identifier
            )));
        };
        let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
        if shape.is_text_box == Some(false) {
            shapes.push(shape_info(
                editor,
                &locations,
                sheet_id,
                reference.identifier,
                &shape,
            )?);
        }
    }
    Ok(shapes)
}

#[allow(deprecated)]
fn shape_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetShapeGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own shape {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers shape {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary shape"
        )));
    };
    let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
    if shape.is_text_box != Some(false) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary shape"
        )));
    }
    if shape
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(sheet_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} does not name sheet {sheet_id} as its parent"
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
    let storage_id =
        required_reference(drawable_object_id, shape.owned_storage.as_ref(), "storage")?;
    if shape
        .deprecated_storage
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(storage_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has inconsistent storage ownership"
        )));
    }
    let object_ids = vec![drawable_object_id, caption_id, title_id, storage_id];
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} aliases private objects"
        )));
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
        (storage_id, STORAGE_MESSAGE_TYPES, "storage"),
    ] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers shape {label} {identifier} is outside {archive_name}"
            )));
        }
        let private_object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers shape {label} {identifier} is missing"))
        })?;
        if private_object
            .messages
            .iter()
            .filter(|message| message_types.contains(&message.type_))
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers shape {label} {identifier} must have exactly one expected payload"
            )));
        }
    }
    validate_shape_ownership(editor.package(), drawable_object_id, storage_id)?;
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
    if !registered.is_empty() && uuid_object_ids.len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers component {component_id} UUID map does not cover shape {drawable_object_id}"
        )));
    }
    Ok(SheetShapeGraph {
        sheet_id,
        archive_name,
        component_id,
        info: shape_info(editor, &locations, sheet_id, drawable_object_id, &shape)?,
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
                "Numbers shape {drawable_object_id} has no {label} object"
            ))
        })
}

fn validate_shape_ownership(
    package: &IWorkPackage,
    drawable_object_id: u64,
    storage_id: u64,
) -> Result<()> {
    let document = numbers_document(package)?;
    let locations = object_locations(package)?;
    let mut drawable_owners = 0usize;
    let mut storage_owners = 0usize;
    for sheet in document.sheets {
        let archive_name = locations.get(&sheet.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", sheet.identifier))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(sheet.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", sheet.identifier))
        })?;
        let (_, native) = decode_sheet(object)?;
        for drawable in native.drawable_infos {
            if drawable.identifier == drawable_object_id {
                drawable_owners += 1;
            }
            let Some(drawable_archive_name) = locations.get(&drawable.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} is missing",
                    drawable.identifier
                )));
            };
            let drawable_archive = package.archive(drawable_archive_name)?;
            let drawable_object =
                drawable_archive
                    .object(drawable.identifier)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers drawable {} is missing",
                            drawable.identifier
                        ))
                    })?;
            for message in drawable_object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            {
                let candidate = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                if candidate
                    .owned_storage
                    .as_ref()
                    .map(|reference| reference.identifier)
                    == Some(storage_id)
                {
                    storage_owners += 1;
                }
            }
        }
    }
    if drawable_owners != 1 || storage_owners != 1 {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has {drawable_owners} sheet owners and storage {storage_id} has {storage_owners} drawable owners"
        )));
    }
    Ok(())
}

fn shape_info(
    editor: &NumbersEditor,
    locations: &HashMap<u64, String>,
    sheet_id: u64,
    drawable_object_id: u64,
    shape: &tswp::ShapeInfoArchive,
) -> Result<NumbersSheetShapeInfo> {
    let storage_id = shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers shape {drawable_object_id} has no writable storage"
            ))
        })?;
    let archive_name = locations.get(&drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers shape {drawable_object_id} is missing"))
    })?;
    if locations.get(&storage_id) != Some(archive_name) {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has storage outside its component"
        )));
    }
    let text = IWorkTextEditor::from_package(editor.package.clone());
    Ok(NumbersSheetShapeInfo {
        sheet_id,
        drawable_object_id,
        kind: shape_path_kind(shape)?,
        preset: shape_preset(shape)?,
        storage: text.storage(storage_id)?,
        geometry: shape_geometry(editor.package(), archive_name, drawable_object_id)?,
        properties: shape_properties(editor.package(), archive_name, drawable_object_id)?,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::ShapeCornerRadius;

    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 300.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 300.0,
        height: 150.0,
    };

    #[test]
    fn scratch_spreadsheet_supports_rectangle_crud_without_a_source_drawable() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Scratch Shape")
            .table_name("Scratch Table")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        assert!(editor.sheet_shapes(sheet_id).unwrap().is_empty());
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_sheet_rectangle(sheet_id, "Built from typed objects", POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, NumbersSheetShapeKind::Rectangle);
        assert_eq!(created.preset, Some(ShapePreset::Rectangle));
        assert_eq!(created.storage.text, "Built from typed objects");
        editor
            .replace_sheet_shape_text(sheet_id, created.drawable_object_id, 0..5, "Made")
            .unwrap();
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 72.0, y: 180.0 }),
            size: Some(DrawableSize {
                width: 360.0,
                height: 180.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(12.0),
        };
        editor
            .set_sheet_shape_geometry(sheet_id, created.drawable_object_id, geometry)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/numbers-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Generated rectangle".to_owned()),
        };
        editor
            .set_sheet_shape_properties(sheet_id, created.drawable_object_id, properties.clone())
            .unwrap();

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let shape = &reopened.sheet_shapes(sheet_id).unwrap()[0];
        assert_eq!(shape.storage.text, "Made from typed objects");
        assert_eq!(shape.geometry, geometry);
        assert_eq!(shape.properties, properties);

        let removed = editor
            .remove_sheet_shape(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_typed_preset_shape_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Preset Shapes")
            .table_name("Source Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        let created = editor
            .add_sheet_shape(
                sheet_id,
                "Rounded",
                POSITION,
                SIZE,
                ShapePreset::ROUNDED_RECTANGLE,
            )
            .unwrap();
        assert_eq!(created.kind, NumbersSheetShapeKind::RoundedRectangle);
        assert_eq!(created.preset, Some(ShapePreset::ROUNDED_RECTANGLE));

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_preset(sheet_id, created.drawable_object_id)
                .unwrap(),
            Some(ShapePreset::ROUNDED_RECTANGLE)
        );

        for (preset, kind) in [
            (ShapePreset::Ellipse, NumbersSheetShapeKind::Ellipse),
            (ShapePreset::PENTAGON, NumbersSheetShapeKind::RegularPolygon),
            (ShapePreset::STAR, NumbersSheetShapeKind::Star),
        ] {
            editor
                .set_sheet_shape_preset(sheet_id, created.drawable_object_id, preset)
                .unwrap();
            let shape = &editor.sheet_shapes(sheet_id).unwrap()[0];
            assert_eq!(shape.kind, kind);
            assert_eq!(shape.preset, Some(preset));
            assert_eq!(shape.storage.text, "Rounded");
        }

        editor
            .remove_sheet_shape(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn invalid_rectangle_creation_and_cross_type_updates_are_transactional() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .add_sheet_rectangle(
                    sheet_id,
                    "invalid",
                    POSITION,
                    DrawableSize {
                        width: 0.0,
                        height: SIZE.height,
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_sheet_shape(
                    sheet_id,
                    "invalid radius",
                    POSITION,
                    SIZE,
                    ShapePreset::RoundedRectangle {
                        corner_radius: ShapeCornerRadius::new(SIZE.height).unwrap(),
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let text_box = editor
            .add_sheet_text_box(sheet_id, "not a shape", POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_shape_text(sheet_id, text_box.drawable_object_id, "wrong type")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_sheet_shape_preset(
                    sheet_id,
                    text_box.drawable_object_id,
                    ShapePreset::Ellipse,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
