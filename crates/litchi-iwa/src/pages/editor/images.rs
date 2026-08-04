//! Body-anchored image CRUD for Pages documents.

use std::collections::HashMap;

use super::*;
use crate::ImageAdjustments;
use crate::comments::DrawableObjectId;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::image_adjustments::replace_image_adjustments;
use crate::image_caption::{CaptionObjectIds, DrawableCaptionKind};
use crate::package_metadata::{add_component_external_reference, component_identifier_for_entry};
use crate::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize,
    flip_drawable_geometry, offset_drawable_geometry, restore_drawable_original_size,
};

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
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Exposure, saturation, and automatic-enhancement settings.
    pub image_adjustments: ImageAdjustments,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

/// Typed layout metadata for a newly created Pages image.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesImageOptions {
    /// Top-left position on the page, in points.
    pub position: DrawablePoint,
    /// Displayed image size, in points.
    pub size: DrawableSize,
    /// Untransformed media dimensions reported to Pages, in points.
    pub natural_size: DrawableSize,
}

impl PagesImageOptions {
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
        options: PagesImageOptions,
    ) -> Result<PagesImageInfo> {
        let geometry = image_creation_values(options)?;
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
            options.natural_size,
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
            || created.original_size != Some(options.natural_size)
            || created.natural_size != Some(options.natural_size)
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
    pub fn body_image_geometry(
        &self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<DrawableGeometry> {
        Ok(body_image_graph(self, drawable_object_id.object_id())?
            .info
            .geometry)
    }

    /// Restore a body image's displayed dimensions from its stored original size.
    ///
    /// This keeps the current position, rotation, reflection, media asset, adjustments,
    /// and properties unchanged. It returns an error when the image has no native
    /// original-size metadata.
    pub fn restore_body_image_original_size(
        &mut self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<DrawableGeometry> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let original_size = source.info.original_size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages image {drawable_object_id} has no original-size metadata"
            ))
        })?;
        let geometry = restore_drawable_original_size(source.info.geometry, original_size)?;
        let mut staged = self.package().clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_image_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages image original-size update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Apply one native Arrange Flip operation to a body-anchored image.
    ///
    /// Returns the updated geometry after applying the same transform as the
    /// Pages Flip Horizontally or Flip Vertically command.
    pub fn flip_body_image(
        &mut self,
        drawable_object_id: DrawableObjectId,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let geometry = flip_drawable_geometry(source.info.geometry, axis)?;
        let mut staged = self.package().clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_image_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages image flip update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Update body-image geometry while preserving unknown image fields.
    pub fn set_body_image_geometry(
        &mut self,
        drawable_object_id: DrawableObjectId,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let mut staged = self.package().clone();
        set_image_geometry(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
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

    /// Read shared drawable properties for one body-anchored image.
    pub fn body_image_properties(
        &self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<DrawableProperties> {
        Ok(body_image_graph(self, drawable_object_id.object_id())?
            .info
            .properties)
    }

    /// Update image accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native image fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_body_image_properties(
        &mut self,
        drawable_object_id: DrawableObjectId,
        properties: DrawableProperties,
    ) -> Result<()> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let mut staged = self.package().clone();
        set_image_properties(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_image_properties(drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Pages image properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the native title and caption attached to one body image.
    pub fn body_image_title_caption(
        &self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<crate::DrawableTitleCaption> {
        image_title_caption(self, drawable_object_id.object_id())
    }

    /// Create or replace one body image's native title.
    pub fn set_body_image_title(
        &mut self,
        drawable_object_id: DrawableObjectId,
        title: &str,
    ) -> Result<()> {
        set_body_image_caption(self, drawable_object_id, title, DrawableCaptionKind::Title)
    }

    /// Remove one body image's native title.
    ///
    /// Returns whether a title was present. Native iWork removal preserves the
    /// prior title graph for undo history and attaches a fresh empty stand-in.
    pub fn remove_body_image_title(
        &mut self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<bool> {
        remove_body_image_caption(self, drawable_object_id, DrawableCaptionKind::Title)
    }

    /// Create or replace one body image's native caption.
    pub fn set_body_image_caption(
        &mut self,
        drawable_object_id: DrawableObjectId,
        caption: &str,
    ) -> Result<()> {
        set_body_image_caption(
            self,
            drawable_object_id,
            caption,
            DrawableCaptionKind::Caption,
        )
    }

    /// Remove one body image's native caption.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_body_image_caption(
        &mut self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<bool> {
        remove_body_image_caption(self, drawable_object_id, DrawableCaptionKind::Caption)
    }

    /// Read the basic controls in iWork's Image inspector for one body image.
    pub fn body_image_adjustments(
        &self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<ImageAdjustments> {
        Ok(body_image_graph(self, drawable_object_id.object_id())?
            .info
            .image_adjustments)
    }

    /// Update image exposure, saturation, and automatic enhancement while preserving advanced
    /// and unknown native adjustment fields.
    pub fn set_body_image_adjustments(
        &mut self,
        drawable_object_id: DrawableObjectId,
        adjustments: ImageAdjustments,
    ) -> Result<()> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let mut staged = self.package().clone();
        let expected = replace_image_adjustments(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
            "Pages image",
            adjustments,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_image_adjustments(drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Pages image adjustment update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one body image at a UTF-16 body position.
    ///
    /// The image, title/caption stand-ins, and body attachment receive fresh
    /// identifiers and UUIDs while retaining the source's style and unknown
    /// protobuf fields. The clone is offset using Pages' native duplicate
    /// placement. Its embedded image asset remains shared with the source, so
    /// replacing either image's data updates both images.
    pub fn duplicate_body_image(
        &mut self,
        source_drawable_object_id: DrawableObjectId,
        anchor_character_index: usize,
    ) -> Result<PagesImageInfo> {
        let raw_source_drawable_object_id = source_drawable_object_id.object_id();
        let source = body_image_graph(self, raw_source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages image graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages image object {identifier} is missing"))
                })?;
                clone_pages_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&raw_source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Pages image clone has no drawable identifier".to_owned())
        })?;
        let new_attachment_id = *remap.get(&source.attachment_id).ok_or_else(|| {
            Error::InvalidFormat("Pages image clone has no attachment identifier".to_owned())
        })?;
        let geometry =
            offset_drawable_geometry(source.info.geometry, BODY_DRAWABLE_DUPLICATE_OFFSET)?;
        set_image_geometry(&mut staged, &source.archive_name, new_drawable_id, geometry)?;
        offset_pages_body_drawable_attachment_clone(
            &mut staged,
            new_attachment_id,
            BODY_DRAWABLE_DUPLICATE_OFFSET,
        )?;
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
            new_attachment_id,
        )?;
        patch_pages_zorder(&mut staged, None, Some(new_drawable_id))?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Pages image graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages image clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &new_uuid_object_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            let new_object_identifier =
                remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages image clone has no data-reference object for {object_identifier}"
                    ))
                })?;
            add_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                new_object_identifier,
            )?;
        }

        let verified = Self::from_package(staged)?;
        let created = verified
            .body_images()?
            .into_iter()
            .find(|image| image.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages image duplication failed validation".to_owned())
            })?;
        let created_graph = body_image_graph(&verified, new_drawable_id)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        let expected_data_references = source
            .data_references
            .iter()
            .map(|&(data_identifier, object_identifier)| {
                let new_object_identifier = remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages image clone has no validated data-reference object for {object_identifier}"
                    ))
                })?;
                Ok((data_identifier, new_object_identifier))
            })
            .collect::<Result<Vec<_>>>()?;
        if created.anchor_character_index != expected_anchor
            || created.image_data_identifier != source.info.image_data_identifier
            || created.thumbnail_data_identifier != source.info.thumbnail_data_identifier
            || created.geometry != geometry
            || created.original_size != source.info.original_size
            || created.natural_size != source.info.natural_size
            || created_graph.object_ids.len() != source.object_ids.len()
            || created_graph.data_references != expected_data_references
        {
            return Err(Error::InvalidFormat(
                "Pages image duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Replace the bytes referenced by one body-anchored image.
    ///
    /// Images duplicated with [`Self::duplicate_body_image`] share data and
    /// observe this replacement together, matching Pages' native behavior.
    pub fn replace_body_image_data(
        &mut self,
        drawable_object_id: DrawableObjectId,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = body_image_graph(self, drawable_object_id.object_id())?;
        self.replace_media(source.info.image_data_identifier, replacement)
    }

    /// Remove a body-anchored image, its private graph, and unshared assets.
    pub fn remove_body_image(
        &mut self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<RemovedPagesImage> {
        let raw_drawable_object_id = drawable_object_id.object_id();
        let source = body_image_graph(self, raw_drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(raw_drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = source.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(raw_drawable_object_id), None)?;
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
            .any(|image| image.drawable_object_id == raw_drawable_object_id)
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

fn set_body_image_caption(
    editor: &mut PagesEditor,
    drawable_object_id: DrawableObjectId,
    text: &str,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let raw_drawable_object_id = drawable_object_id.object_id();
    let source = body_image_graph(editor, raw_drawable_object_id)?;
    let slot = image_caption_slot(editor, raw_drawable_object_id, kind)?;
    let before = image_title_caption(editor, raw_drawable_object_id)?;
    let staged = if let Some(storage_id) = slot.storage_id {
        let mut text_editor = IWorkTextEditor::from_package(editor.package().clone());
        text_editor.set_text(storage_id, text)?;
        text_editor.into_package()
    } else {
        let root = root_document(editor.package())?;
        let (theme, language) = image_caption_theme(editor.package(), &root)?;
        let image_width = source
            .info
            .geometry
            .size
            .ok_or_else(|| Error::InvalidFormat("Pages image has no displayed size".to_owned()))?
            .width;
        let ids = CaptionObjectIds::allocate(next_object_identifier(editor.package())?)?;
        let mut staged = editor.package().clone();
        insert_image_caption(
            &mut staged,
            &source.archive_name,
            raw_drawable_object_id,
            slot.reference_id,
            image_width,
            text,
            kind,
            theme,
            language.as_deref(),
            ids,
        )?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;
        staged
    };
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    let actual = verified.body_image_title_caption(drawable_object_id)?;
    let expected = match kind {
        DrawableCaptionKind::Caption => crate::DrawableTitleCaption {
            title: before.title,
            caption: Some(text.to_owned()),
        },
        DrawableCaptionKind::Title => crate::DrawableTitleCaption {
            title: Some(text.to_owned()),
            caption: before.caption,
        },
    };
    if actual != expected {
        return Err(Error::InvalidFormat(
            "Pages image title/caption update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_body_image_caption(
    editor: &mut PagesEditor,
    drawable_object_id: DrawableObjectId,
    kind: DrawableCaptionKind,
) -> Result<bool> {
    let raw_drawable_object_id = drawable_object_id.object_id();
    let source = body_image_graph(editor, raw_drawable_object_id)?;
    let slot = image_caption_slot(editor, raw_drawable_object_id, kind)?;
    if slot.storage_id.is_none() {
        return Ok(false);
    }
    let before = image_title_caption(editor, raw_drawable_object_id)?;
    let standin_id = next_object_identifier(editor.package())?;
    let mut staged = editor.package().clone();
    insert_image_caption_standin(
        &mut staged,
        &source.archive_name,
        raw_drawable_object_id,
        slot.reference_id,
        kind,
        standin_id,
    )?;
    add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &[standin_id])?;
    set_package_last_object_identifier(&mut staged, standin_id)?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    let actual = verified.body_image_title_caption(drawable_object_id)?;
    let expected = match kind {
        DrawableCaptionKind::Caption => crate::DrawableTitleCaption {
            title: before.title,
            caption: None,
        },
        DrawableCaptionKind::Title => crate::DrawableTitleCaption {
            title: None,
            caption: before.caption,
        },
    };
    if actual != expected {
        return Err(Error::InvalidFormat(
            "Pages image title/caption removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::PathBuf;

    use super::*;
    use crate::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 180.0,
    };
    const NATURAL_IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 640.0,
        height: 480.0,
    };
    const UPDATED_ANGLE_DEGREES: f32 = 7.5;

    fn options() -> PagesImageOptions {
        PagesImageOptions::new(IMAGE_POSITION, IMAGE_SIZE).with_natural_size(NATURAL_IMAGE_SIZE)
    }

    fn selector(image: &PagesImageInfo) -> DrawableObjectId {
        DrawableObjectId::from_object_id(image.drawable_object_id).unwrap()
    }

    fn missing_selector() -> DrawableObjectId {
        DrawableObjectId::from_object_id(999).unwrap()
    }

    #[test]
    fn image_selectors_use_drawable_object_id() {
        let _: fn(&PagesEditor, DrawableObjectId) -> Result<DrawableGeometry> =
            PagesEditor::body_image_geometry;
        let _: fn(&mut PagesEditor, DrawableObjectId) -> Result<DrawableGeometry> =
            PagesEditor::restore_body_image_original_size;
        let _: fn(
            &mut PagesEditor,
            DrawableObjectId,
            DrawableFlipAxis,
        ) -> Result<DrawableGeometry> = PagesEditor::flip_body_image;
        let _: fn(&mut PagesEditor, DrawableObjectId, DrawableGeometry) -> Result<()> =
            PagesEditor::set_body_image_geometry;
        let _: fn(&PagesEditor, DrawableObjectId) -> Result<DrawableProperties> =
            PagesEditor::body_image_properties;
        let _: fn(&mut PagesEditor, DrawableObjectId, DrawableProperties) -> Result<()> =
            PagesEditor::set_body_image_properties;
        let _: fn(&PagesEditor, DrawableObjectId) -> Result<crate::DrawableTitleCaption> =
            PagesEditor::body_image_title_caption;
        let _: fn(&mut PagesEditor, DrawableObjectId, &str) -> Result<()> =
            PagesEditor::set_body_image_title;
        let _: fn(&mut PagesEditor, DrawableObjectId) -> Result<bool> =
            PagesEditor::remove_body_image_title;
        let _: fn(&mut PagesEditor, DrawableObjectId, &str) -> Result<()> =
            PagesEditor::set_body_image_caption;
        let _: fn(&mut PagesEditor, DrawableObjectId) -> Result<bool> =
            PagesEditor::remove_body_image_caption;
        let _: fn(&PagesEditor, DrawableObjectId) -> Result<ImageAdjustments> =
            PagesEditor::body_image_adjustments;
        let _: fn(&mut PagesEditor, DrawableObjectId, ImageAdjustments) -> Result<()> =
            PagesEditor::set_body_image_adjustments;
        let _: fn(&mut PagesEditor, DrawableObjectId, usize) -> Result<PagesImageInfo> =
            PagesEditor::duplicate_body_image;
        let _: fn(&mut PagesEditor, DrawableObjectId, &[u8]) -> Result<Vec<u8>> =
            PagesEditor::replace_body_image_data;
        let _: fn(&mut PagesEditor, DrawableObjectId) -> Result<RemovedPagesImage> =
            PagesEditor::remove_body_image;
    }

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
                options(),
            )
            .unwrap();
        let created_id = selector(&created);
        assert_eq!(created.thumbnail_data_identifier, None);
        assert_eq!(created.original_size, Some(NATURAL_IMAGE_SIZE));
        assert_eq!(created.natural_size, Some(NATURAL_IMAGE_SIZE));
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
            .set_body_image_geometry(created_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor.body_image_geometry(created_id).unwrap(),
            changed_geometry
        );
        let restored_original_size = editor.restore_body_image_original_size(created_id).unwrap();
        let expected_original_size_geometry = DrawableGeometry {
            size: Some(NATURAL_IMAGE_SIZE),
            ..changed_geometry
        };
        assert_eq!(restored_original_size, expected_original_size_geometry);
        assert_eq!(
            editor.body_image_geometry(created_id).unwrap(),
            expected_original_size_geometry
        );
        let horizontally_flipped = editor
            .flip_body_image(created_id, DrawableFlipAxis::Horizontal)
            .unwrap();
        assert_eq!(
            editor.body_image_geometry(created_id).unwrap(),
            horizontally_flipped
        );
        assert_ne!(
            horizontally_flipped.flags,
            expected_original_size_geometry.flags
        );
        let vertically_flipped = editor
            .flip_body_image(created_id, DrawableFlipAxis::Vertical)
            .unwrap();
        assert_eq!(
            editor.body_image_geometry(created_id).unwrap(),
            vertically_flipped
        );
        assert_ne!(
            vertically_flipped.angle,
            expected_original_size_geometry.angle
        );

        let changed_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/pages-image".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Quarterly-results portrait".to_owned()),
        };
        editor
            .set_body_image_properties(created_id, changed_properties.clone())
            .unwrap();
        assert_eq!(
            editor.body_image_properties(created_id).unwrap(),
            changed_properties
        );
        editor
            .set_body_image_properties(created_id, DrawableProperties::default())
            .unwrap();
        assert_eq!(
            editor.body_image_properties(created_id).unwrap(),
            DrawableProperties::default()
        );

        let changed_adjustments = ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        editor
            .set_body_image_adjustments(created_id, changed_adjustments)
            .unwrap();
        assert_eq!(
            editor.body_image_adjustments(created_id).unwrap(),
            changed_adjustments
        );
        editor
            .set_body_image_adjustments(created_id, created.image_adjustments)
            .unwrap();
        assert_eq!(
            editor.body_image_adjustments(created_id).unwrap(),
            created.image_adjustments
        );

        let previous = editor
            .replace_body_image_data(created_id, &replacement)
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

        let removed = editor.remove_body_image(created_id).unwrap();
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
    fn scratch_document_supports_native_image_title_caption_crud() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = PagesEditor::create_with_text("Image labels").unwrap();
        let image = editor
            .add_body_image(
                "Image labels".encode_utf16().count(),
                "lena.png",
                &original,
                options(),
            )
            .unwrap();
        let image_id = selector(&image);

        assert_eq!(
            editor.body_image_title_caption(image_id).unwrap(),
            crate::DrawableTitleCaption::default()
        );
        editor
            .set_body_image_title(image_id, "Quarterly portrait")
            .unwrap();
        editor
            .set_body_image_caption(image_id, "Revenue report")
            .unwrap();
        assert_eq!(
            editor.body_image_title_caption(image_id).unwrap(),
            crate::DrawableTitleCaption {
                title: Some("Quarterly portrait".to_owned()),
                caption: Some("Revenue report".to_owned()),
            }
        );

        editor
            .set_body_image_title(image_id, "Updated portrait")
            .unwrap();
        assert!(editor.remove_body_image_caption(image_id).unwrap());
        assert!(!editor.remove_body_image_caption(image_id).unwrap());
        assert!(editor.remove_body_image_title(image_id).unwrap());
        assert_eq!(
            editor.body_image_title_caption(image_id).unwrap(),
            crate::DrawableTitleCaption::default()
        );

        editor
            .set_body_image_caption(image_id, "Recreated caption")
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.body_image_title_caption(image_id).unwrap(),
            crate::DrawableTitleCaption {
                title: None,
                caption: Some("Recreated caption".to_owned()),
            }
        );
    }

    #[test]
    fn scratch_document_supports_native_image_duplication() {
        let original = fixture("test-data/images/png/lena.png");
        let replacement = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = PagesEditor::create_with_text("Quarterly report").unwrap();
        let source = editor
            .add_body_image(
                "Quarterly report".encode_utf16().count(),
                "lena.png",
                &original,
                options(),
            )
            .unwrap();
        let source_id = selector(&source);
        let source_properties = DrawableProperties {
            hyperlink_url: Some("https://example.test/pages-source".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Source portrait".to_owned()),
        };
        editor
            .set_body_image_properties(source_id, source_properties.clone())
            .unwrap();
        editor
            .set_body_image_title(source_id, "Source title")
            .unwrap();
        editor
            .set_body_image_caption(source_id, "Source caption")
            .unwrap();
        let source_geometry = editor
            .flip_body_image(source_id, DrawableFlipAxis::Vertical)
            .unwrap();
        let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();

        let duplicate = editor
            .duplicate_body_image(source_id, duplicate_anchor)
            .unwrap();
        let duplicate_id = selector(&duplicate);
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = body_image_graph(&editor, source.drawable_object_id).unwrap();
        let duplicate_graph = body_image_graph(&editor, duplicate.drawable_object_id).unwrap();
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
        assert_eq!(duplicate.anchor_character_index, duplicate_anchor as u32);
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
            editor.body_image_title_caption(duplicate_id).unwrap(),
            crate::DrawableTitleCaption {
                title: Some("Source title".to_owned()),
                caption: Some("Source caption".to_owned()),
            }
        );
        assert_eq!(
            duplicate.geometry.position,
            source_geometry.position.map(|position| DrawablePoint {
                x: position.x + BODY_DRAWABLE_DUPLICATE_OFFSET,
                y: position.y + BODY_DRAWABLE_DUPLICATE_OFFSET,
            })
        );
        assert_eq!(duplicate.geometry.size, source_geometry.size);
        assert_eq!(duplicate.geometry.flags, source_geometry.flags);
        assert_eq!(duplicate.geometry.angle, source_geometry.angle);

        let moved_duplicate = DrawableGeometry {
            position: Some(DrawablePoint { x: 312.0, y: 264.0 }),
            ..duplicate.geometry
        };
        editor
            .set_body_image_geometry(duplicate_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor.body_image_geometry(source_id).unwrap(),
            source_geometry
        );
        assert_eq!(
            editor.body_image_geometry(duplicate_id).unwrap(),
            moved_duplicate
        );

        let duplicate_properties = DrawableProperties {
            accessibility_description: Some("Independent portrait clone".to_owned()),
            ..source_properties.clone()
        };
        editor
            .set_body_image_properties(duplicate_id, duplicate_properties.clone())
            .unwrap();
        assert_eq!(
            editor.body_image_properties(source_id).unwrap(),
            source_properties
        );
        assert_eq!(
            editor.body_image_properties(duplicate_id).unwrap(),
            duplicate_properties
        );

        assert_eq!(
            editor
                .replace_body_image_data(duplicate_id, &replacement)
                .unwrap(),
            original
        );
        assert_eq!(
            editor.extract_media(source.image_data_identifier).unwrap(),
            replacement
        );
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_images().unwrap().len(), 2);
        assert_eq!(
            reopened
                .body_images()
                .unwrap()
                .into_iter()
                .find(|image| image.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .geometry,
            moved_duplicate
        );

        let removed_source = editor.remove_body_image(source_id).unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.body_images().unwrap().len(), 1);
        let removed_duplicate = editor.remove_body_image(duplicate_id).unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.image_data_identifier]
        );
        assert!(editor.body_images().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_image_creation_and_geometry_are_transactional() {
        let original = fixture("test-data/images/png/lena.png");
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();

        let missing_id = missing_selector();
        assert!(editor.duplicate_body_image(missing_id, 0).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_body_image(4, "payload.bin", b"not an image", options(),)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_body_image(5, "lena.png", &original, options())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_body_image(
                    4,
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
            .add_body_image(4, "lena.png", &original, options())
            .unwrap();
        let before_restore = editor.to_bytes().unwrap();
        assert!(editor.restore_body_image_original_size(missing_id).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before_restore);
        let before_flip = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_body_image(missing_id, DrawableFlipAxis::Horizontal)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);
        let before_properties = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_image_properties(missing_id, DrawableProperties::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_properties);
        let before_geometry = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_image_geometry(
                    selector(&created),
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
