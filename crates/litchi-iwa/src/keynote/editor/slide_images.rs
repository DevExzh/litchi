//! Standalone image-object CRUD for Keynote slides.

use std::collections::HashMap;

use super::*;
use crate::ImageAdjustments;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::image_adjustments::replace_image_adjustments;
use crate::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize,
    flip_drawable_geometry, offset_drawable_geometry, restore_drawable_original_size,
};

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
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Exposure, saturation, and automatic-enhancement settings.
    pub image_adjustments: ImageAdjustments,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

/// Typed layout metadata for a newly created Keynote slide image.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteSlideImageOptions {
    /// Top-left position on the slide, in points.
    pub position: DrawablePoint,
    /// Displayed image size on the slide, in points.
    pub size: DrawableSize,
    /// Untransformed media dimensions reported to Keynote, in points.
    pub natural_size: DrawableSize,
}

impl KeynoteSlideImageOptions {
    /// Create options whose displayed and natural dimensions are identical.
    pub const fn new(position: DrawablePoint, size: DrawableSize) -> Self {
        Self {
            position,
            size,
            natural_size: size,
        }
    }

    /// Set media dimensions independently of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: DrawableSize) -> Self {
        self.natural_size = natural_size;
        self
    }
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
        options: KeynoteSlideImageOptions,
    ) -> Result<KeynoteSlideImageInfo> {
        let geometry = image_creation_values(options)?;
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
            options.natural_size,
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
            || created.original_size != Some(options.natural_size)
            || created.natural_size != Some(options.natural_size)
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

    /// Restore a file-backed slide image's displayed dimensions from its stored original size.
    ///
    /// This keeps the current position, rotation, reflection, media asset, adjustments,
    /// and properties unchanged. It returns an error when the image has no native
    /// original-size metadata.
    pub fn restore_slide_image_original_size(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let original_size = source.info.original_size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote image {drawable_object_id} has no original-size metadata"
            ))
        })?;
        let geometry = restore_drawable_original_size(source.info.geometry, original_size)?;
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
                "Keynote image original-size update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Apply one native Arrange Flip operation to an ordinary file-backed slide image.
    ///
    /// Returns the updated geometry after applying the same transform as the
    /// Keynote Flip Horizontally or Flip Vertically command.
    pub fn flip_slide_image(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let geometry = flip_drawable_geometry(source.info.geometry, axis)?;
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
                "Keynote image flip update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
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

    /// Read shared drawable properties for one ordinary file-backed slide image.
    pub fn slide_image_properties(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(require_file_image(self, slide_index, drawable_object_id)?
            .info
            .properties)
    }

    /// Update image accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native image fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_slide_image_properties(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_image_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_image_properties(slide_index, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Keynote image properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the basic controls in iWork's Image inspector for one ordinary slide image.
    pub fn slide_image_adjustments(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ImageAdjustments> {
        Ok(require_file_image(self, slide_index, drawable_object_id)?
            .info
            .image_adjustments)
    }

    /// Update image exposure, saturation, and automatic enhancement while preserving advanced
    /// and unknown native adjustment fields.
    pub fn set_slide_image_adjustments(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        adjustments: ImageAdjustments,
    ) -> Result<()> {
        let source = require_file_image(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        let expected = replace_image_adjustments(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            "Keynote image",
            adjustments,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_image_adjustments(slide_index, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Keynote image adjustment update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one ordinary file-backed slide image using Keynote's native
    /// shared-asset behavior.
    ///
    /// The image and its private title/caption graph receive fresh identifiers
    /// and UUIDs while retaining the source's style and unknown protobuf
    /// fields. The clone is offset on the same slide, but its embedded image
    /// asset remains shared, so replacing either image's data updates both.
    pub fn duplicate_slide_image(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
    ) -> Result<KeynoteSlideImageInfo> {
        let source = require_file_image(self, slide_index, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Keynote image graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Keynote image object {identifier} is missing"))
                })?;
                clone_slide_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Keynote image clone has no drawable identifier".to_owned())
        })?;
        let geometry = offset_drawable_geometry(source.info.geometry, DRAWABLE_DUPLICATE_OFFSET)?;
        set_image_geometry(&mut staged, &source.archive_name, new_drawable_id, geometry)?;
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Keynote image graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote image clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            let new_object_identifier =
                remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote image clone has no data-reference object for {object_identifier}"
                    ))
                })?;
            add_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                new_object_identifier,
            )?;
        }

        let verified = Self::from_package(staged)?;
        let created = verified
            .slide_images(slide_index)?
            .into_iter()
            .find(|image| image.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote image duplication failed validation".to_owned())
            })?;
        let created_graph = require_file_image(&verified, slide_index, new_drawable_id)?;
        let expected_data_references = source
            .data_references
            .iter()
            .map(|&(data_identifier, object_identifier)| {
                let new_object_identifier = remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote image clone has no validated data-reference object for {object_identifier}"
                    ))
                })?;
                Ok((data_identifier, new_object_identifier))
            })
            .collect::<Result<Vec<_>>>()?;
        if created.kind != KeynoteSlideImageKind::File
            || created.image_data_identifier != source.info.image_data_identifier
            || created.thumbnail_data_identifier != source.info.thumbnail_data_identifier
            || created.geometry != geometry
            || created.original_size != source.info.original_size
            || created.natural_size != source.info.natural_size
            || created_graph.object_ids.len() != source.object_ids.len()
            || created_graph.data_references != expected_data_references
        {
            return Err(Error::InvalidFormat(
                "Keynote image duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
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
    use crate::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 512.0,
        height: 512.0,
    };
    const NATURAL_IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 640.0,
        height: 480.0,
    };

    fn options() -> KeynoteSlideImageOptions {
        KeynoteSlideImageOptions::new(IMAGE_POSITION, IMAGE_SIZE)
            .with_natural_size(NATURAL_IMAGE_SIZE)
    }

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
            .add_slide_image(0, "lena.png", &original, options())
            .unwrap();
        assert_eq!(created.kind, KeynoteSlideImageKind::File);
        assert_eq!(created.thumbnail_data_identifier, None);
        assert_eq!(created.original_size, Some(NATURAL_IMAGE_SIZE));
        assert_eq!(created.natural_size, Some(NATURAL_IMAGE_SIZE));
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
        let restored_original_size = editor
            .restore_slide_image_original_size(0, created.drawable_object_id)
            .unwrap();
        let expected_original_size_geometry = DrawableGeometry {
            size: Some(NATURAL_IMAGE_SIZE),
            ..changed_geometry
        };
        assert_eq!(restored_original_size, expected_original_size_geometry);
        assert_eq!(
            editor
                .slide_image_geometry(0, created.drawable_object_id)
                .unwrap(),
            expected_original_size_geometry
        );
        let horizontally_flipped = editor
            .flip_slide_image(0, created.drawable_object_id, DrawableFlipAxis::Horizontal)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_geometry(0, created.drawable_object_id)
                .unwrap(),
            horizontally_flipped
        );
        assert_ne!(
            horizontally_flipped.flags,
            expected_original_size_geometry.flags
        );
        let vertically_flipped = editor
            .flip_slide_image(0, created.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_geometry(0, created.drawable_object_id)
                .unwrap(),
            vertically_flipped
        );
        assert_ne!(
            vertically_flipped.angle,
            expected_original_size_geometry.angle
        );

        let changed_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/keynote-image".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Quarterly-results portrait".to_owned()),
        };
        editor
            .set_slide_image_properties(0, created.drawable_object_id, changed_properties.clone())
            .unwrap();
        assert_eq!(
            editor
                .slide_image_properties(0, created.drawable_object_id)
                .unwrap(),
            changed_properties
        );
        editor
            .set_slide_image_properties(
                0,
                created.drawable_object_id,
                DrawableProperties::default(),
            )
            .unwrap();
        assert_eq!(
            editor
                .slide_image_properties(0, created.drawable_object_id)
                .unwrap(),
            DrawableProperties::default()
        );

        let changed_adjustments = ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        editor
            .set_slide_image_adjustments(0, created.drawable_object_id, changed_adjustments)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_adjustments(0, created.drawable_object_id)
                .unwrap(),
            changed_adjustments
        );
        editor
            .set_slide_image_adjustments(0, created.drawable_object_id, created.image_adjustments)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_adjustments(0, created.drawable_object_id)
                .unwrap(),
            created.image_adjustments
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
    fn scratch_presentation_supports_native_image_duplication() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Image clone")
            .subtitle("Shared native media")
            .build()
            .unwrap();
        let source = editor
            .add_slide_image(0, "lena.png", &original, options())
            .unwrap();
        let source_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/keynote-source".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Source portrait".to_owned()),
        };
        editor
            .set_slide_image_properties(0, source.drawable_object_id, source_properties.clone())
            .unwrap();
        let source_geometry = editor
            .flip_slide_image(0, source.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();

        let duplicate = editor
            .duplicate_slide_image(0, source.drawable_object_id)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = require_file_image(&editor, 0, source.drawable_object_id).unwrap();
        let duplicate_graph = require_file_image(&editor, 0, duplicate.drawable_object_id).unwrap();
        assert!(
            source_graph
                .object_ids
                .iter()
                .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
        );
        assert!(
            source_graph
                .uuid_object_ids
                .iter()
                .all(|identifier| !duplicate_graph.uuid_object_ids.contains(identifier))
        );
        assert_eq!(duplicate.kind, KeynoteSlideImageKind::File);
        assert_eq!(duplicate.slide_index, source.slide_index);
        assert_eq!(
            duplicate.image_data_identifier,
            source.image_data_identifier
        );
        assert_eq!(
            duplicate.thumbnail_data_identifier,
            source.thumbnail_data_identifier
        );
        assert_eq!(duplicate.original_size, source.original_size);
        assert_eq!(duplicate.natural_size, source.natural_size);
        assert_eq!(duplicate.properties, source_properties);
        assert_eq!(
            duplicate.geometry.position,
            source_geometry.position.map(|position| DrawablePoint {
                x: position.x + DRAWABLE_DUPLICATE_OFFSET,
                y: position.y + DRAWABLE_DUPLICATE_OFFSET,
            })
        );
        assert_eq!(duplicate.geometry.size, source_geometry.size);
        assert_eq!(duplicate.geometry.flags, source_geometry.flags);
        assert_eq!(duplicate.geometry.angle, source_geometry.angle);

        let moved_duplicate = DrawableGeometry {
            position: Some(DrawablePoint { x: 720.0, y: 420.0 }),
            ..duplicate.geometry
        };
        editor
            .set_slide_image_geometry(0, duplicate.drawable_object_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor
                .slide_image_geometry(0, source.drawable_object_id)
                .unwrap(),
            source_geometry
        );
        assert_eq!(
            editor
                .slide_image_geometry(0, duplicate.drawable_object_id)
                .unwrap(),
            moved_duplicate
        );

        let duplicate_properties = DrawableProperties {
            accessibility_description: Some("Independent portrait clone".to_owned()),
            ..source_properties.clone()
        };
        editor
            .set_slide_image_properties(
                0,
                duplicate.drawable_object_id,
                duplicate_properties.clone(),
            )
            .unwrap();
        assert_eq!(
            editor
                .slide_image_properties(0, source.drawable_object_id)
                .unwrap(),
            source_properties
        );
        assert_eq!(
            editor
                .slide_image_properties(0, duplicate.drawable_object_id)
                .unwrap(),
            duplicate_properties
        );

        assert_eq!(
            editor
                .replace_slide_image_data(0, duplicate.drawable_object_id, &replacement)
                .unwrap(),
            original
        );
        assert_eq!(
            editor.extract_media(source.image_data_identifier).unwrap(),
            replacement
        );
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.slide_images(0).unwrap().len(), 2);
        assert_eq!(
            reopened
                .slide_images(0)
                .unwrap()
                .into_iter()
                .find(|image| image.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .geometry,
            moved_duplicate
        );

        let removed_source = editor
            .remove_slide_image(0, source.drawable_object_id)
            .unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.slide_images(0).unwrap().len(), 1);
        let removed_duplicate = editor
            .remove_slide_image(0, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.image_data_identifier]
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

        assert!(editor.duplicate_slide_image(0, 999).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_slide_image(0, "payload.bin", b"not an image", options())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_slide_image(
                    0,
                    "lena.png",
                    &original,
                    options().with_natural_size(DrawableSize {
                        width: 0.0,
                        height: NATURAL_IMAGE_SIZE.height,
                    }),
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let created = editor
            .add_slide_image(0, "lena.png", &original, options())
            .unwrap();
        let before_restore = editor.to_bytes().unwrap();
        assert!(editor.restore_slide_image_original_size(0, 999).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before_restore);
        let before_flip = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_slide_image(0, 999, DrawableFlipAxis::Horizontal)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);
        let before_properties = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_image_properties(0, 999, DrawableProperties::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_properties);
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
