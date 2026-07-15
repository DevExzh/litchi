//! Standalone image-object CRUD for Keynote slides.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

mod graph;

use graph::*;

/// Semantic role of an image drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSlideImageKind {
    /// Ordinary file-backed image inserted by the user or this crate.
    File,
    /// Image materialized from the selected slide layout.
    Layout,
}

/// One image drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideImageInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub kind: KeynoteSlideImageKind,
    pub image_data_identifier: u64,
    pub thumbnail_data_identifier: Option<u64>,
    pub geometry: DrawableGeometry,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

/// Result of removing one slide-owned image and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideImage {
    pub image: KeynoteSlideImageInfo,
    /// Assets culled because the removed image held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl KeynoteEditor {
    /// List image drawables directly owned by one slide in drawable order.
    pub fn slide_images(&self, slide_index: usize) -> Result<Vec<KeynoteSlideImageInfo>> {
        image_infos(self, slide_index)
    }

    /// Add an independently editable image to a slide.
    ///
    /// The image object, title/caption stand-ins, drawable ownership, z-order,
    /// UUID metadata, component references, and embedded data record are built
    /// directly from typed values. No source drawable or package template is copied.
    pub fn add_slide_image(
        &mut self,
        slide_index: usize,
        preferred_filename: &str,
        data: &[u8],
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<KeynoteSlideImageInfo> {
        let geometry = image_geometry(position, size)?;
        let context = image_creation_context(self, slide_index)?;
        let ids = ImageObjectIds::allocate(next_object_identifier(self.package())?)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Keynote slide images require image data, not {}",
                asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        let objects = image_objects(
            ids,
            context.slide_id,
            context.style_id,
            asset.data_identifier,
            geometry,
        )?;
        staged.update_archive(&context.archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_slide_drawable_references(
            &mut staged,
            &context.archive_name,
            context.slide_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, context.component_id, &ids.all())?;
        add_component_data_reference(
            &mut staged,
            context.component_id,
            asset.data_identifier,
            ids.drawable,
        )?;
        add_component_external_reference(
            &mut staged,
            context.component_id,
            context.stylesheet_component_id,
            context.style_id,
        )?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slide_images(slide_index)?
            .into_iter()
            .find(|image| image.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote image creation failed validation".to_owned())
            })?;
        let created_graph = image_graph(&verified, slide_index, ids.drawable)?;
        if created.kind != KeynoteSlideImageKind::File
            || created.image_data_identifier != asset.data_identifier
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Keynote image creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read typed geometry for one ordinary slide image.
    pub fn slide_image_geometry(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(require_file_image(self, slide_index, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update image geometry while preserving unknown image fields.
    pub fn set_slide_image_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_image_geometry(slide_index, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Keynote image geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the bytes referenced by one ordinary slide image.
    ///
    /// Images sharing the data identifier observe the replacement, matching
    /// native Keynote shared-asset behavior.
    pub fn replace_slide_image_data(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        self.replace_media(source.info.image_data_identifier, replacement)
    }

    /// Remove an ordinary image, its private object graph, and unshared assets.
    pub fn remove_slide_image(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideImage> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            remove_object(&mut staged, &source.archive_name, *identifier)?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote image object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
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
            .slide_images(slide_index)?
            .iter()
            .any(|image| image.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Keynote image deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteSlideImage {
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
    use crate::keynote::KeynoteDocumentBuilder;

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 512.0,
        height: 512.0,
    };

    #[test]
    fn scratch_presentation_supports_image_crud_without_a_source_drawable() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch image")
            .subtitle("No embedded package")
            .build()
            .unwrap();

        assert!(editor.slide_images(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        let created = editor
            .add_slide_image(0, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        assert_eq!(created.kind, KeynoteSlideImageKind::File);
        assert_eq!(created.thumbnail_data_identifier, None);
        assert_eq!(created.original_size, Some(IMAGE_SIZE));
        assert_eq!(created.natural_size, Some(IMAGE_SIZE));
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            original
        );

        let roundtripped = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.slide_images(0).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 64.0, y: 96.0 }),
            size: Some(DrawableSize {
                width: 320.0,
                height: 180.0,
            }),
            flags: Some(3),
            angle: Some(12.5),
        };
        editor
            .set_slide_image_geometry(0, created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_geometry(0, created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );

        let previous = editor
            .replace_slide_image_data(0, created.drawable_object_id, &replacement)
            .unwrap();
        assert_eq!(previous, original);
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            replacement
        );
        assert_eq!(
            editor.slide_images(0).unwrap()[0].image_data_identifier,
            created.image_data_identifier
        );

        let removed = editor
            .remove_slide_image(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.image.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.image_data_identifier]
        );
        assert!(editor.slide_images(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_image_creation_and_geometry_are_transactional() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_slide_image(
                    0,
                    "payload.bin",
                    b"not an image",
                    IMAGE_POSITION,
                    IMAGE_SIZE
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let created = editor
            .add_slide_image(0, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        let before_geometry = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_image_geometry(
                    0,
                    created.drawable_object_id,
                    DrawableGeometry {
                        size: Some(DrawableSize {
                            width: -1.0,
                            height: 12.0,
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
