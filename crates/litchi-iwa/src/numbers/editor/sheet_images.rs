//! Standalone image-object CRUD for Numbers sheets.

use std::collections::HashMap;

use super::*;
use crate::ImageAdjustments;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::image_adjustments::replace_image_adjustments;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, offset_drawable_geometry,
};

mod graph;

use graph::*;

/// One ordinary image drawable owned directly by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetImageInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
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

/// Result of removing one sheet-owned image and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetImage {
    pub image: NumbersSheetImageInfo,
    /// Assets culled because the removed image held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl NumbersEditor {
    /// List ordinary image drawables owned directly by one reachable sheet.
    pub fn sheet_images(&self, sheet_id: u64) -> Result<Vec<NumbersSheetImageInfo>> {
        image_infos(self, sheet_id)
    }

    /// Add an independently editable image to a reachable sheet.
    ///
    /// The image, title/caption stand-ins, sheet ownership, style link, UUIDs,
    /// component data reference, and `Data/*` asset are constructed directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_image(
        &mut self,
        sheet_id: u64,
        preferred_filename: &str,
        data: &[u8],
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersSheetImageInfo> {
        let geometry = image_geometry(position, size)?;
        let context = image_creation_context(self, sheet_id)?;
        let ids = ImageObjectIds::allocate(next_object_identifier(&self.package)?)?;

        let mut media = IWorkMediaEditor::from_package(self.package.clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Numbers sheet images require image data, not {}",
                asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        let objects = image_objects(
            ids,
            sheet_id,
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
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &context.archive_name,
            sheet_id,
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
        if context.stylesheet_component_id != context.component_id {
            add_component_external_reference(
                &mut staged,
                context.component_id,
                context.stylesheet_component_id,
                context.style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_images(sheet_id)?
            .into_iter()
            .find(|image| image.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers image creation failed validation".to_owned())
            })?;
        let created_graph = image_graph(&verified, sheet_id, ids.drawable)?;
        if created.image_data_identifier != asset.data_identifier
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Numbers image creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read typed geometry for one ordinary sheet image.
    pub fn sheet_image_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(image_graph(self, sheet_id, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update image position, size, flags, and rotation while preserving
    /// unknown image fields.
    pub fn set_sheet_image_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = image_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_image_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers image geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary sheet image.
    pub fn sheet_image_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(image_graph(self, sheet_id, drawable_object_id)?
            .info
            .properties)
    }

    /// Update image accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native image fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_sheet_image_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = image_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_image_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_image_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers image properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the basic controls in iWork's Image inspector for one sheet image.
    pub fn sheet_image_adjustments(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ImageAdjustments> {
        Ok(image_graph(self, sheet_id, drawable_object_id)?
            .info
            .image_adjustments)
    }

    /// Update image exposure, saturation, and automatic enhancement while preserving advanced
    /// and unknown native adjustment fields.
    pub fn set_sheet_image_adjustments(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        adjustments: ImageAdjustments,
    ) -> Result<()> {
        let source = image_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        let expected = replace_image_adjustments(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            "Numbers image",
            adjustments,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_image_adjustments(sheet_id, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers image adjustment update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one ordinary sheet image using Numbers' native placement.
    ///
    /// The image and title/caption stand-ins receive fresh identifiers and
    /// UUIDs while retaining the source's style and unknown protobuf fields.
    /// The clone is added to the same sheet and reuses the embedded image
    /// asset, so replacing either image's data updates both images.
    pub fn duplicate_sheet_image(
        &mut self,
        sheet_id: u64,
        source_drawable_object_id: u64,
    ) -> Result<NumbersSheetImageInfo> {
        let source = image_graph(self, sheet_id, source_drawable_object_id)?;
        let mut staged = self.package.clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Numbers image graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers image object {identifier} is missing"))
                })?;
                clone_numbers_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers image clone has no drawable identifier".to_owned())
        })?;
        let geometry = offset_drawable_geometry(source.info.geometry, DRAWABLE_DUPLICATE_OFFSET)?;
        set_image_geometry(&mut staged, &source.archive_name, new_drawable_id, geometry)?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            source.sheet_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Numbers image graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers image clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            let new_object_identifier =
                remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers image clone has no data-reference object for {object_identifier}"
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
            .sheet_images(sheet_id)?
            .into_iter()
            .find(|image| image.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers image duplication failed validation".to_owned())
            })?;
        let created_graph = image_graph(&verified, sheet_id, new_drawable_id)?;
        let expected_data_references = source
            .data_references
            .iter()
            .map(|&(data_identifier, object_identifier)| {
                let new_object_identifier = remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers image clone has no validated data-reference object for {object_identifier}"
                    ))
                })?;
                Ok((data_identifier, new_object_identifier))
            })
            .collect::<Result<Vec<_>>>()?;
        if created.sheet_id != source.info.sheet_id
            || created.image_data_identifier != source.info.image_data_identifier
            || created.thumbnail_data_identifier != source.info.thumbnail_data_identifier
            || created.geometry != geometry
            || created.original_size != source.info.original_size
            || created.natural_size != source.info.natural_size
            || created_graph.object_ids.len() != source.object_ids.len()
            || created_graph.data_references != expected_data_references
        {
            return Err(Error::InvalidFormat(
                "Numbers image duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Replace the primary bytes referenced by one ordinary sheet image.
    ///
    /// Images sharing the data identifier observe the replacement, matching
    /// native Numbers shared-asset behavior.
    pub fn replace_sheet_image_data(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = image_graph(self, sheet_id, drawable_object_id)?;
        self.replace_media(source.info.image_data_identifier, replacement)
    }

    /// Remove an ordinary image, its private graph, and unshared assets.
    pub fn remove_sheet_image(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetImage> {
        let source = image_graph(self, sheet_id, drawable_object_id)?;
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
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                object_identifier,
            )?;
        }
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
                    Error::InvalidFormat(format!("Numbers image object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers image object {identifier} remains referenced after deletion"
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
            .sheet_images(sheet_id)?
            .iter()
            .any(|image| image.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Numbers image deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetImage {
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
    use crate::numbers::NumbersDocumentBuilder;
    use crate::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 240.0,
    };
    const UPDATED_ANGLE_DEGREES: f32 = 9.0;

    #[test]
    fn scratch_spreadsheet_supports_image_crud_without_a_source_package() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Media")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;

        assert!(editor.sheet_images(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        let created = editor
            .add_sheet_image(sheet_id, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        assert_eq!(created.sheet_id, sheet_id);
        assert_eq!(created.thumbnail_data_identifier, None);
        assert_eq!(created.original_size, Some(IMAGE_SIZE));
        assert_eq!(created.natural_size, Some(IMAGE_SIZE));
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            original
        );

        let roundtripped = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.sheet_images(sheet_id).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 64.0, y: 96.0 }),
            size: Some(DrawableSize {
                width: 192.0,
                height: 108.0,
            }),
            flags: Some(3),
            angle: Some(UPDATED_ANGLE_DEGREES),
        };
        editor
            .set_sheet_image_geometry(sheet_id, created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_geometry(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );

        let changed_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/numbers-image".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Quarterly-results portrait".to_owned()),
        };
        editor
            .set_sheet_image_properties(
                sheet_id,
                created.drawable_object_id,
                changed_properties.clone(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_properties(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_properties
        );
        editor
            .set_sheet_image_properties(
                sheet_id,
                created.drawable_object_id,
                DrawableProperties::default(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_properties(sheet_id, created.drawable_object_id)
                .unwrap(),
            DrawableProperties::default()
        );

        let changed_adjustments = ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        editor
            .set_sheet_image_adjustments(sheet_id, created.drawable_object_id, changed_adjustments)
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_adjustments(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_adjustments
        );
        editor
            .set_sheet_image_adjustments(
                sheet_id,
                created.drawable_object_id,
                created.image_adjustments,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_adjustments(sheet_id, created.drawable_object_id)
                .unwrap(),
            created.image_adjustments
        );

        let previous = editor
            .replace_sheet_image_data(sheet_id, created.drawable_object_id, &replacement)
            .unwrap();
        assert_eq!(previous, original);
        assert_eq!(
            editor.extract_media(created.image_data_identifier).unwrap(),
            replacement
        );
        assert_eq!(
            editor.sheet_images(sheet_id).unwrap()[0].image_data_identifier,
            created.image_data_identifier
        );
        editor
            .set_sheet_drawable_comment(
                sheet_id,
                created.drawable_object_id,
                "Remove this image after review",
            )
            .unwrap();

        let removed = editor
            .remove_sheet_image(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.image.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.image_data_identifier]
        );
        assert!(editor.sheet_images(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_image_duplication() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Media")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_image(sheet_id, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        let source_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/numbers-source".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Source portrait".to_owned()),
        };
        editor
            .set_sheet_image_properties(
                sheet_id,
                source.drawable_object_id,
                source_properties.clone(),
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet_image(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = image_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let duplicate_graph = image_graph(&editor, sheet_id, duplicate.drawable_object_id).unwrap();
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
        assert_eq!(duplicate.sheet_id, source.sheet_id);
        assert_eq!(
            duplicate.image_data_identifier,
            source.image_data_identifier
        );
        assert_eq!(
            duplicate.thumbnail_data_identifier,
            source.thumbnail_data_identifier
        );
        assert_eq!(duplicate.properties, source_properties);
        assert_eq!(
            duplicate.geometry.position,
            source.geometry.position.map(|position| DrawablePoint {
                x: position.x + DRAWABLE_DUPLICATE_OFFSET,
                y: position.y + DRAWABLE_DUPLICATE_OFFSET,
            })
        );
        assert_eq!(duplicate.geometry.size, source.geometry.size);
        assert_eq!(duplicate.geometry.flags, source.geometry.flags);
        assert_eq!(duplicate.geometry.angle, source.geometry.angle);

        let moved_duplicate = DrawableGeometry {
            position: Some(DrawablePoint { x: 640.0, y: 360.0 }),
            ..duplicate.geometry
        };
        editor
            .set_sheet_image_geometry(sheet_id, duplicate.drawable_object_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_geometry(sheet_id, source.drawable_object_id)
                .unwrap(),
            source.geometry
        );
        assert_eq!(
            editor
                .sheet_image_geometry(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            moved_duplicate
        );

        let duplicate_properties = DrawableProperties {
            accessibility_description: Some("Independent portrait clone".to_owned()),
            ..source_properties.clone()
        };
        editor
            .set_sheet_image_properties(
                sheet_id,
                duplicate.drawable_object_id,
                duplicate_properties.clone(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_properties(sheet_id, source.drawable_object_id)
                .unwrap(),
            source_properties
        );
        assert_eq!(
            editor
                .sheet_image_properties(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            duplicate_properties
        );

        assert_eq!(
            editor
                .replace_sheet_image_data(sheet_id, duplicate.drawable_object_id, &replacement)
                .unwrap(),
            original
        );
        assert_eq!(
            editor.extract_media(source.image_data_identifier).unwrap(),
            replacement
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheet_images(sheet_id).unwrap().len(), 2);
        assert_eq!(
            reopened
                .sheet_images(sheet_id)
                .unwrap()
                .into_iter()
                .find(|image| image.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .geometry,
            moved_duplicate
        );

        let removed_source = editor
            .remove_sheet_image(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.sheet_images(sheet_id).unwrap().len(), 1);
        let removed_duplicate = editor
            .remove_sheet_image(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.image_data_identifier]
        );
        assert!(editor.sheet_images(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_image_creation_and_geometry_are_transactional() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

        assert!(editor.duplicate_sheet_image(sheet_id, 999).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_sheet_image(
                    sheet_id,
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
                .add_sheet_image(999, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let created = editor
            .add_sheet_image(sheet_id, "lena.png", &original, IMAGE_POSITION, IMAGE_SIZE)
            .unwrap();
        let before_properties = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_image_properties(sheet_id, 999, DrawableProperties::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_properties);
        let before_geometry = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_image_geometry(
                    sheet_id,
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
