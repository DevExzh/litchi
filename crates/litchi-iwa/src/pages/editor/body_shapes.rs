//! Body-anchored ordinary shape CRUD for Pages documents.

use std::ops::Range;

use super::*;
use crate::package_metadata::{
    add_component_external_reference, component_identifier_for_entry,
    remove_component_external_references_to_object,
};
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, ShapePathKind,
};

use super::text_box_create::{
    BodyTextShapeObjectIds, BodyTextShapeRole, body_text_shape_objects, body_text_storage,
};

mod graph;

use graph::*;

/// Structural path family used by an ordinary Pages body shape.
pub type PagesBodyShapeKind = ShapePathKind;

/// One ordinary, non-text-box shape anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesBodyShapeInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub kind: PagesBodyShapeKind,
    pub storage: TextStorageInfo,
    pub geometry: DrawableGeometry,
    pub properties: DrawableProperties,
}

/// Result of removing an ordinary body shape and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesBodyShape {
    pub shape: PagesBodyShapeInfo,
}

impl PagesEditor {
    /// List ordinary non-text-box shapes anchored to the document body.
    pub fn body_shapes(&self) -> Result<Vec<PagesBodyShapeInfo>> {
        body_shape_infos(self)
    }

    /// Add a text-bearing rectangle at one UTF-16 position in the body.
    ///
    /// The shape, path, writable storage, stand-ins, body attachment, z-order,
    /// UUIDs, and style relationship are constructed directly from typed
    /// values. No source drawable or package template is copied.
    pub fn add_body_rectangle(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<PagesBodyShapeInfo> {
        let geometry = rectangle_geometry(position, size)?;
        let root = root_document(self.package())?;
        let body: StorageArchive = decode_typed_package_object(
            self.package(),
            self.body_storage_id,
            self.body_storage()?.message_type,
            "TSWP.StorageArchive",
        )?;
        let style_id = shape_style_id(self.package(), &root)?;
        let storage = body_text_storage(text, &body);
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = BodyTextShapeObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;
        let objects = body_text_shape_objects(
            ids,
            self.body_storage_id,
            style_id,
            geometry,
            storage,
            root.left_margin.unwrap_or_default(),
            BodyTextShapeRole::Shape,
        )?;

        let mut staged = self.package().clone();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
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
        let style_archive = find_object_archive(&staged, style_id)?;
        if let Some(stylesheet_component) = component_identifier_for_entry(&staged, &style_archive)?
            && stylesheet_component != DOCUMENT_OBJECT_ID
        {
            add_component_external_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                stylesheet_component,
                style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .body_shapes()?
            .into_iter()
            .find(|shape| shape.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages rectangle creation failed validation".to_owned())
            })?;
        let created_graph = body_shape_graph(&verified, ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.anchor_character_index != expected_anchor
            || created.kind != PagesBodyShapeKind::Rectangle
            || created.storage.object_id != ids.storage
            || created.storage.text != text
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
        {
            return Err(Error::InvalidFormat(
                "Pages rectangle creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read typed geometry for one ordinary body shape.
    pub fn body_shape_geometry(&self, drawable_object_id: u64) -> Result<DrawableGeometry> {
        Ok(body_shape_graph(self, drawable_object_id)?.info.geometry)
    }

    /// Update body-shape geometry while preserving unknown shape fields.
    pub fn set_body_shape_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages shape geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary body shape.
    pub fn body_shape_properties(&self, drawable_object_id: u64) -> Result<DrawableProperties> {
        Ok(body_shape_graph(self, drawable_object_id)?.info.properties)
    }

    /// Update shared drawable properties while preserving unknown fields.
    pub fn set_body_shape_properties(
        &mut self,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_properties(drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Pages shape property update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a UTF-16 range in an ordinary body shape's text.
    pub fn replace_body_shape_text(
        &mut self,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package().clone());
        text.replace_text(source.info.storage.object_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        body_shape_graph(&verified, drawable_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Replace all text in an ordinary body shape.
    pub fn set_body_shape_text(
        &mut self,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package().clone());
        text.set_text(source.info.storage.object_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = body_shape_graph(&verified, drawable_object_id)?;
        if updated.info.storage.text != replacement {
            return Err(Error::InvalidFormat(
                "Pages shape text update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear an ordinary body shape without deleting it.
    pub fn clear_body_shape_text(&mut self, drawable_object_id: u64) -> Result<()> {
        self.set_body_shape_text(drawable_object_id, "")
    }

    /// Remove an ordinary body shape and its private text-bearing graph.
    pub fn remove_body_shape(&mut self, drawable_object_id: u64) -> Result<RemovedPagesBodyShape> {
        let graph = body_shape_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = graph.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        for identifier in &graph.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                *identifier,
            )?;
        }
        staged.update_archive(&graph.archive_name, |archive| {
            for identifier in &graph.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages shape object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        for identifier in &graph.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages shape object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &graph.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .body_shapes()?
            .iter()
            .any(|shape| shape.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages shape deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesBodyShape { shape: graph.info })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 300.0,
        height: 150.0,
    };

    #[test]
    fn scratch_document_supports_rectangle_crud_without_a_source_drawable() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline_body = editor.body_text().unwrap();
        assert!(editor.body_shapes().unwrap().is_empty());

        let created = editor
            .add_body_rectangle(4, "Built from typed objects", POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, PagesBodyShapeKind::Rectangle);
        assert_eq!(created.storage.text, "Built from typed objects");
        editor
            .replace_body_shape_text(created.drawable_object_id, 0..5, "Made")
            .unwrap();
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 96.0, y: 180.0 }),
            size: Some(DrawableSize {
                width: 360.0,
                height: 180.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(12.0),
        };
        editor
            .set_body_shape_geometry(created.drawable_object_id, geometry)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/pages-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Generated rectangle".to_owned()),
        };
        editor
            .set_body_shape_properties(created.drawable_object_id, properties.clone())
            .unwrap();

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let shape = &reopened.body_shapes().unwrap()[0];
        assert_eq!(shape.storage.text, "Made from typed objects");
        assert_eq!(shape.geometry, geometry);
        assert_eq!(shape.properties, properties);

        let mut package = editor.into_package();
        let root = root_document(&package).unwrap();
        let style_id = shape_style_id(&package, &root).unwrap();
        let style_archive = find_object_archive(&package, style_id).unwrap();
        let style_component = component_identifier_for_entry(&package, &style_archive)
            .unwrap()
            .unwrap();
        add_component_external_reference(
            &mut package,
            style_component,
            DOCUMENT_OBJECT_ID,
            created.drawable_object_id,
        )
        .unwrap();
        editor = PagesEditor::from_package(package).unwrap();

        let removed = editor
            .remove_body_shape(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.body_text().unwrap(), baseline_body);
        assert!(editor.body_shapes().unwrap().is_empty());
    }

    #[test]
    fn invalid_creation_and_cross_type_updates_are_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .add_body_rectangle(
                    4,
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

        let text_box = editor
            .add_text_box(4, "not a shape", POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_shape_text(text_box.drawable_object_id, "wrong type")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
