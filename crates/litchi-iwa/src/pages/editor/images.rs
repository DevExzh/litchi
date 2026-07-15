//! Body-anchored image CRUD for Pages documents.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::{add_component_external_reference, component_identifier_for_entry};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

mod graph;

use graph::*;

/// One image anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesImageInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub image_data_identifier: u64,
    pub thumbnail_data_identifier: Option<u64>,
    pub geometry: DrawableGeometry,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

/// Result of removing a body-anchored Pages image.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesImage {
    pub image: PagesImageInfo,
    /// Assets culled because the removed image held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl PagesEditor {
    /// List images anchored to the document body in text-flow order.
    pub fn body_images(&self) -> Result<Vec<PagesImageInfo>> {
        body_image_infos(self)
    }

    /// Add an independently editable image anchored at a UTF-16 body position.
    ///
    /// The image, title/caption stand-ins, body attachment, object-replacement
    /// character, z-order, UUIDs, component data reference, and `Data/*` asset
    /// are all created from typed values. No source object or package is copied.
    pub fn add_body_image(
        &mut self,
        anchor_character_index: usize,
        preferred_filename: &str,
        data: &[u8],
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<PagesImageInfo> {
        let geometry = image_geometry(position, size)?;
        let root = root_document(self.package())?;
        let style_id = image_style_id(self.package(), &root)?;
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = ImageObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Pages body images require image data, not {}",
                asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        let objects = image_objects(
            ids,
            self.body_storage_id,
            style_id,
            asset.data_identifier,
            geometry,
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
        add_component_data_reference(
            &mut staged,
            DOCUMENT_OBJECT_ID,
            asset.data_identifier,
            ids.drawable,
        )?;
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
            .body_images()?
            .into_iter()
            .find(|image| image.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages image creation failed validation".to_owned())
            })?;
        let created_graph = body_image_graph(&verified, ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.anchor_character_index != expected_anchor
            || created.image_data_identifier != asset.data_identifier
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Pages image creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read geometry for one body-anchored image.
    pub fn body_image_geometry(&self, drawable_object_id: u64) -> Result<DrawableGeometry> {
        Ok(body_image_graph(self, drawable_object_id)?.info.geometry)
    }

    /// Update body-image geometry while preserving unknown image fields.
    pub fn set_body_image_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = body_image_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_image_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages image geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the bytes referenced by one body-anchored image.
    pub fn replace_body_image_data(
        &mut self,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = body_image_graph(self, drawable_object_id)?;
        self.replace_media(source.info.image_data_identifier, replacement)
    }

    /// Remove a body-anchored image, its private graph, and unshared assets.
    pub fn remove_body_image(&mut self, drawable_object_id: u64) -> Result<RemovedPagesImage> {
        let source = body_image_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = source.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            let object_archive = find_object_archive(&staged, *identifier)?;
            staged.update_archive(&object_archive, |archive| {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages image object {identifier} is missing from {object_archive}"
                    ))
                })?;
                Ok(())
            })?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages image object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let mut media = IWorkMediaEditor::from_package(staged)?;
        let mut removed_data_identifiers = Vec::new();
        let data_identifiers = source
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<HashSet<_>>();
        for identifier in data_identifiers {
            if media
                .asset(identifier)
                .is_some_and(|asset| !asset.is_referenced())
            {
                media.remove_unreferenced(identifier)?;
                removed_data_identifiers.push(identifier);
            }
        }
        removed_data_identifiers.sort_unstable();
        let staged = media.into_package();
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let remaining_assets = verified.media_assets()?;
        if verified
            .body_images()?
            .iter()
            .any(|image| image.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Pages image deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesImage {
            image: source.info,
            removed_data_identifiers,
        })
    }
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::PathBuf;

    use super::*;

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 180.0,
    };
    const UPDATED_ANGLE_DEGREES: f32 = 7.5;

    #[test]
    fn scratch_document_supports_image_crud_without_a_source_package() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = PagesEditor::create_with_text("Quarterly report").unwrap();

        assert!(editor.body_images().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        let created = editor
            .add_body_image(
                "Quarterly report".encode_utf16().count(),
                "lena.png",
                &original,
                IMAGE_POSITION,
                IMAGE_SIZE,
            )
            .unwrap();
        assert_eq!(created.thumbnail_data_identifier, None);
        assert_eq!(created.original_size, Some(IMAGE_SIZE));
        assert_eq!(created.natural_size, Some(IMAGE_SIZE));
        assert_eq!(editor.body_text().unwrap(), "Quarterly report\u{fffc}");
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            original
        );

        let roundtripped = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.body_images().unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 72.0, y: 216.0 }),
            size: Some(DrawableSize {
                width: 192.0,
                height: 108.0,
            }),
            flags: Some(3),
            angle: Some(UPDATED_ANGLE_DEGREES),
        };
        editor
            .set_body_image_geometry(created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .body_image_geometry(created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );

        let previous = editor
            .replace_body_image_data(created.drawable_object_id, &replacement)
            .unwrap();
        assert_eq!(previous, original);
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            replacement
        );
        assert_eq!(
            editor.body_images().unwrap()[0].image_data_identifier,
            created.image_data_identifier
        );

        let removed = editor
            .remove_body_image(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.image.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.image_data_identifier]
        );
        assert_eq!(editor.body_text().unwrap(), "Quarterly report");
        assert!(editor.body_images().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_image_creation_and_geometry_are_transactional() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_body_image(
                    4,
                    "payload.bin",
                    b"not an image",
                    IMAGE_POSITION,
                    IMAGE_SIZE,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_body_image(5, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let created = editor
            .add_body_image(4, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        let before_geometry = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_image_geometry(
                    created.drawable_object_id,
                    DrawableGeometry {
                        size: Some(DrawableSize {
                            width: -1.0,
                            height: IMAGE_SIZE.height,
                        }),
                        ..created.geometry
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_geometry);
    }

    fn fixture(relative: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        fs::read(root.join(relative)).unwrap()
    }
}
