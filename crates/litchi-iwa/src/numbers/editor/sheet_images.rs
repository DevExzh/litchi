//! Standalone image-object CRUD for Numbers sheets.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

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
    fn invalid_image_creation_and_geometry_are_transactional() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

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
