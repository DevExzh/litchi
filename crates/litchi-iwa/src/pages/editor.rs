//! Semantic editing of reachable Pages body, header, footer, and drawable storages.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::comments::{
    DrawableCommentInfo, DrawableCommentReplyInfo, IWorkDrawableCommentEditor, IWorkDrawableInfo,
};
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    add_component_object_uuids, component_uuid_identifiers, next_object_identifier,
    release_package_identifier_suffix, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::protobuf::tp::{self, DocumentArchive, SectionArchive, SectionTemplateArchive};
use crate::protobuf::tsd;
use crate::protobuf::tsp;
use crate::protobuf::tswp::{
    DrawableAttachmentArchive, StorageArchive, object_attribute_table::ObjectAttribute,
};
use crate::shapes::{
    DrawableGeometry, DrawableProperties, set_shape_geometry, set_shape_properties, shape_geometry,
    shape_properties,
};
use crate::text::{IWorkTextEditor, TextStorageInfo};
use crate::wire::{
    append_repeated_length_delimited_field, patch_fixed32_field, patch_length_delimited_field,
    patch_nested_fixed32_field, patch_nested_varint_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};

const SECTION_MESSAGE_TYPE: u32 = 10011;
const SECTION_TEMPLATE_MESSAGE_TYPE: u32 = 10143;
const USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE: u32 = 10016;
const GUIDE_STORAGE_MESSAGE_TYPE: u32 = 3047;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2001, 2022];
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2011;
const DRAWABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2003;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3097;
const TEXT_BOX_DUPLICATE_OFFSET: f32 = 12.0;

/// Which page variant owns a Pages header/footer storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesTemplateKind {
    First,
    Even,
    Odd,
}

/// Whether a reachable Pages text region is a header or a footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesHeaderFooterKind {
    Header,
    Footer,
}

/// A reachable header/footer slot and its current writable text storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesHeaderFooterInfo {
    pub section_id: u64,
    pub section_name: Option<String>,
    /// UTF-16 position where the section begins in the body storage.
    pub section_character_index: u32,
    pub template_id: u64,
    pub template: PagesTemplateKind,
    pub kind: PagesHeaderFooterKind,
    /// Archive order within the header/footer list, normally left/center/right.
    pub slot: usize,
    pub storage: TextStorageInfo,
}

/// A writable text storage owned by a drawable reachable from a Pages document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesDrawableTextInfo {
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// Result of removing a body-anchored Pages text box.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedPagesTextBox {
    pub text: PagesDrawableTextInfo,
    /// UTF-16 body position formerly occupied by the object-replacement character.
    pub anchor_character_index: u32,
}

/// A section boundary reachable from the main Pages body storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesSectionInfo {
    pub object_id: u64,
    /// UTF-16 position where the section begins in the body storage.
    pub character_index: u32,
    pub name: Option<String>,
    pub first_template_id: Option<u64>,
    pub even_template_id: Option<u64>,
    pub odd_template_id: Option<u64>,
}

/// Writable settings stored directly on a Pages section.
///
/// Numeric kinds remain raw so newer iWork values can round-trip without an
/// artificial enum rejecting them. `background_fill_payload`, when present,
/// is the exact encoded `TSD.FillArchive` payload.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PagesSectionSettings {
    pub name: Option<String>,
    pub inherit_previous_header_footer: Option<bool>,
    pub first_page_different: Option<bool>,
    pub even_odd_pages_different: Option<bool>,
    pub start_kind: Option<u32>,
    pub page_number_kind: Option<u32>,
    pub page_number_start: Option<u32>,
    pub first_page_hides_header_footer: Option<bool>,
    pub background_fill_payload: Option<Vec<u8>>,
}

/// RGB color space used by a semantic Pages section background.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PagesRgbColorSpace {
    Srgb,
    DisplayP3,
}

/// Normalized RGB color components in the inclusive `0.0..=1.0` range.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesRgbaColor {
    pub red: f32,
    pub green: f32,
    pub blue: f32,
    pub alpha: f32,
    pub color_space: PagesRgbColorSpace,
}

/// Semantic Pages section background.
///
/// Gradient, image, extension, and future fills are exposed as `Opaque` so
/// callers can round-trip them losslessly through the same API.
#[derive(Debug, Clone, PartialEq)]
pub enum PagesSectionBackground {
    None,
    Solid(PagesRgbaColor),
    Opaque(Vec<u8>),
}

/// Writable page geometry stored on the Pages document root.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesPageLayout {
    pub page_width: Option<f32>,
    pub page_height: Option<f32>,
    pub left_margin: Option<f32>,
    pub right_margin: Option<f32>,
    pub top_margin: Option<f32>,
    pub bottom_margin: Option<f32>,
    pub header_margin: Option<f32>,
    pub footer_margin: Option<f32>,
    pub page_scale: Option<f32>,
    /// Raw Pages orientation value; `0` is the default used by portrait files.
    pub orientation: Option<u32>,
    pub lays_out_body_vertically: Option<bool>,
}

impl From<&DocumentArchive> for PagesPageLayout {
    fn from(document: &DocumentArchive) -> Self {
        Self {
            page_width: document.page_width,
            page_height: document.page_height,
            left_margin: document.left_margin,
            right_margin: document.right_margin,
            top_margin: document.top_margin,
            bottom_margin: document.bottom_margin,
            header_margin: document.header_margin,
            footer_margin: document.footer_margin,
            page_scale: document.page_scale,
            orientation: document.orientation,
            lays_out_body_vertically: document.lays_out_body_vertically,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct HeaderFooterLocation {
    section_id: u64,
    section_name: Option<String>,
    section_character_index: u32,
    template_id: u64,
    template: PagesTemplateKind,
    kind: PagesHeaderFooterKind,
    slot: usize,
    storage_id: u64,
}

/// Transactional editor for the body text of an existing Pages package.
#[derive(Debug, Clone)]
pub struct PagesEditor {
    text: IWorkTextEditor,
    body_storage_id: u64,
    sections: Vec<PagesSectionInfo>,
    header_footers: Vec<HeaderFooterLocation>,
}

impl PagesEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        let body_storage_id = body_storage_id(&package)?;
        let (sections, header_footers) = discover_structure(&package, body_storage_id)?;
        let text = IWorkTextEditor::from_package(package);
        if text.storage(body_storage_id).is_err() {
            return Err(Error::InvalidFormat(format!(
                "Pages body object {body_storage_id} has no writable text storage"
            )));
        }
        Ok(Self {
            text,
            body_storage_id,
            sections,
            header_footers,
        })
    }

    pub fn body_storage(&self) -> Result<TextStorageInfo> {
        self.text.storage(self.body_storage_id).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages body storage {} disappeared",
                self.body_storage_id
            ))
        })
    }

    pub fn body_text(&self) -> Result<String> {
        Ok(self.body_storage()?.text)
    }

    /// List supported direct-comment drawables reachable from the Pages root.
    pub fn drawables(&self) -> Result<Vec<IWorkDrawableInfo>> {
        let reachable = self.reachable_drawable_ids()?;
        let mut drawables = IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .drawables()?
            .into_iter()
            .filter(|drawable| reachable.contains(&drawable.object_id))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.object_id);
        Ok(drawables)
    }

    /// List writable text storages owned by reachable Pages drawables.
    ///
    /// Results are ordered by drawable object identifier. Shared storage
    /// ownership is rejected because a mutation could otherwise affect more
    /// than the caller-selected object.
    pub fn drawable_text_storages(&self) -> Result<Vec<PagesDrawableTextInfo>> {
        let mut storage_owners = HashMap::<u64, u64>::new();
        let mut result = Vec::new();
        for drawable in self.drawables()? {
            let Some(storage_id) = drawable_owned_text_storage(self.package(), &drawable)? else {
                continue;
            };
            if let Some(previous_drawable) = storage_owners.insert(storage_id, drawable.object_id) {
                return Err(Error::InvalidFormat(format!(
                    "Pages drawables {previous_drawable} and {} share owned text storage {storage_id}",
                    drawable.object_id
                )));
            }
            let storage = self.text.storage(storage_id).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Pages drawable {} owns invalid text storage {storage_id}: {error}",
                    drawable.object_id
                ))
            })?;
            result.push(PagesDrawableTextInfo {
                drawable_object_id: drawable.object_id,
                storage,
            });
        }
        Ok(result)
    }

    /// Replace a UTF-16 range in a reachable Pages text box or placeholder.
    pub fn replace_drawable_text(
        &mut self,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let storage_id = self.require_drawable_text_storage(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.replace_text(storage_id, range, replacement)?;
        Self::from_package(staged.package().clone())?;
        self.text = staged;
        Ok(())
    }

    /// Replace all text in a reachable Pages text box or placeholder.
    pub fn set_drawable_text(&mut self, drawable_object_id: u64, replacement: &str) -> Result<()> {
        let storage_id = self.require_drawable_text_storage(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text(storage_id, replacement)?;
        Self::from_package(staged.package().clone())?;
        self.text = staged;
        Ok(())
    }

    /// Clear a reachable Pages text box or placeholder without deleting it.
    pub fn clear_drawable_text(&mut self, drawable_object_id: u64) -> Result<()> {
        self.set_drawable_text(drawable_object_id, "")
    }

    /// Read the geometry of a reachable ordinary text box.
    pub fn text_box_geometry(&self, drawable_object_id: u64) -> Result<DrawableGeometry> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        shape_geometry(self.package(), &archive_name, drawable_object_id)
    }

    /// Update position, size, flags, and rotation on a reachable ordinary text box.
    pub fn set_text_box_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        self.text_box_graph(drawable_object_id)?;
        let mut staged = self.package().clone();
        let archive_name = find_object_archive(&staged, drawable_object_id)?;
        set_shape_geometry(&mut staged, &archive_name, drawable_object_id, geometry)?;
        let verified = Self::from_package(staged)?;
        if verified.text_box_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages text-box geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties from a reachable ordinary text box.
    pub fn text_box_properties(&self, drawable_object_id: u64) -> Result<DrawableProperties> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        shape_properties(self.package(), &archive_name, drawable_object_id)
    }

    /// Update shared drawable properties on a reachable ordinary text box.
    ///
    /// Pages may disable locking for anchored text boxes even though the
    /// corresponding property is present in the shared drawable archive.
    pub fn set_text_box_properties(
        &mut self,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        self.text_box_graph(drawable_object_id)?;
        let mut staged = self.package().clone();
        let archive_name = find_object_archive(&staged, drawable_object_id)?;
        set_shape_properties(&mut staged, &archive_name, drawable_object_id, &properties)?;
        let verified = Self::from_package(staged)?;
        if verified.text_box_properties(drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Pages text-box properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate a body-anchored Pages text box at a UTF-16 body position.
    ///
    /// The shape, writable storage, attachment, and stand-in caption objects
    /// are copied with fresh identifiers. Existing formatting and unknown
    /// protobuf fields on the source graph are preserved, while the new text
    /// storage is independent from its source.
    pub fn duplicate_text_box(
        &mut self,
        source_drawable_object_id: u64,
        anchor_character_index: usize,
        text: &str,
    ) -> Result<PagesDrawableTextInfo> {
        let source = self.text_box_graph(source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages text-box graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let archive_name = find_object_archive(&staged, *identifier)?;
            let cloned = {
                let archive = staged.archive(&archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages text-box object {identifier} is missing"))
                })?;
                clone_pages_text_box_object(source_object, &remap)?
            };
            staged.update_archive(&archive_name, |archive| archive.insert_object(cloned))?;
        }

        let new_drawable_id = remap[&source.drawable_id];
        let new_storage_id = remap[&source.storage_id];
        let new_attachment_id = remap[&source.attachment_id];
        offset_text_box_clone(
            &mut staged,
            new_drawable_id,
            new_attachment_id,
            TEXT_BOX_DUPLICATE_OFFSET,
        )?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.set_text(new_storage_id, text)?;
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
            Error::InvalidFormat("Pages text-box graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(&mut staged, 1, &new_uuid_object_ids)?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .drawable_text_storages()?
            .into_iter()
            .find(|item| item.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages text-box duplication failed validation".to_owned())
            })?;
        let created_graph = verified.text_box_graph(new_drawable_id)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created_graph.anchor_character_index != expected_anchor
            || created.storage.object_id != new_storage_id
            || created.storage.text != text
        {
            return Err(Error::InvalidFormat(
                "Pages text-box duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Remove a body-anchored Pages text box and its private object graph.
    pub fn remove_text_box(&mut self, drawable_object_id: u64) -> Result<RemovedPagesTextBox> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let text = self
            .drawable_text_storages()?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages text box {drawable_object_id} lost its writable storage"
                ))
            })?;

        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = graph.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;

        for identifier in &graph.object_ids {
            let archive_name = find_object_archive(&staged, *identifier)?;
            staged.update_archive(&archive_name, |archive| {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages text-box object {identifier} is missing from {archive_name}"
                    ))
                })?;
                Ok(())
            })?;
        }
        for identifier in &graph.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages text-box object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, 1, &graph.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_package(staged)?;
        if verified
            .drawables()?
            .iter()
            .any(|drawable| drawable.object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages text-box deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesTextBox {
            text,
            anchor_character_index: graph.anchor_character_index,
        })
    }

    /// Read a comment attached directly to a reachable Pages drawable.
    pub fn drawable_comment(&self, drawable_object_id: u64) -> Result<Option<DrawableCommentInfo>> {
        self.require_drawable(drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .comment(drawable_object_id)
    }

    /// Create or replace a direct comment on a reachable Pages drawable.
    pub fn set_drawable_comment(
        &mut self,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.set_comment(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Delete a direct comment from a reachable Pages drawable.
    pub fn clear_drawable_comment(&mut self, drawable_object_id: u64) -> Result<()> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Read the direct replies in a reachable Pages drawable comment thread.
    pub fn drawable_comment_replies(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<DrawableCommentReplyInfo>> {
        self.require_drawable(drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .replies(drawable_object_id)
    }

    /// Add a reply to a reachable Pages drawable comment.
    pub fn add_drawable_comment_reply(
        &mut self,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.add_reply(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Update a direct reply, returning its current storage identifier.
    pub fn set_drawable_comment_reply(
        &mut self,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.set_reply(drawable_object_id, reply_storage_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Remove a direct reply from a reachable Pages drawable comment.
    pub fn remove_drawable_comment_reply(
        &mut self,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.remove_reply(drawable_object_id, reply_storage_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Replace a UTF-16 range, matching the indexes in Pages text attributes.
    pub fn replace_body_text(&mut self, range: Range<usize>, replacement: &str) -> Result<()> {
        self.text
            .replace_text(self.body_storage_id, range, replacement)
    }

    pub fn set_body_text(&mut self, replacement: &str) -> Result<()> {
        self.text.set_text(self.body_storage_id, replacement)
    }

    pub fn clear_body(&mut self) -> Result<()> {
        self.set_body_text("")
    }

    /// List body section boundaries in UTF-16 document order.
    pub fn sections(&self) -> &[PagesSectionInfo] {
        &self.sections
    }

    /// Read the lossless settings payload of a reachable Pages section.
    pub fn section_settings(&self, section_id: u64) -> Result<PagesSectionSettings> {
        if !self
            .sections
            .iter()
            .any(|section| section.object_id == section_id)
        {
            return Err(Error::ParseError(format!(
                "Section {section_id} is not reachable from the Pages body"
            )));
        }
        let archive_name = find_section_archive(self.text.package(), section_id)?;
        let archive = self.text.package().archive(&archive_name)?;
        let object = archive.object(section_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == SECTION_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages section object {section_id} must have one section payload, found {}",
                messages.len()
            )));
        }
        decode_section_settings(messages[0].data.as_slice())
    }

    /// Replace the settings stored directly on a reachable Pages section.
    ///
    /// The update is transactional and patches only changed protobuf fields,
    /// preserving unknown fields, raw background-fill bytes, and field order.
    pub fn set_section_settings(
        &mut self,
        section_id: u64,
        settings: PagesSectionSettings,
    ) -> Result<()> {
        validate_section_settings(&settings)?;
        let current = self.section_settings(section_id)?;
        if current == settings {
            return Ok(());
        }

        let mut staged = self.text.package().clone();
        let archive_name = find_section_archive(&staged, section_id)?;
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(section_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
            })?;
            let message_indexes = object
                .messages
                .iter()
                .enumerate()
                .filter_map(|(index, message)| {
                    (message.type_ == SECTION_MESSAGE_TYPE).then_some(index)
                })
                .collect::<Vec<_>>();
            if message_indexes.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Pages section object {section_id} must have one section payload, found {}",
                    message_indexes.len()
                )));
            }
            let message_index = message_indexes[0];
            let original = object.messages[message_index].data.as_slice();
            let decoded_current = decode_section_settings(original)?;
            if decoded_current != current {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {section_id} changed during mutation"
                )));
            }

            let mut data = original.to_vec();
            for (field_number, before, after) in [
                (
                    17,
                    current.inherit_previous_header_footer,
                    settings.inherit_previous_header_footer,
                ),
                (
                    18,
                    current.first_page_different,
                    settings.first_page_different,
                ),
                (
                    19,
                    current.even_odd_pages_different,
                    settings.even_odd_pages_different,
                ),
                (
                    28,
                    current.first_page_hides_header_footer,
                    settings.first_page_hides_header_footer,
                ),
            ] {
                if before != after {
                    data = patch_varint_field(
                        &data,
                        field_number,
                        before.is_some(),
                        after.map(u64::from),
                    )?;
                }
            }
            for (field_number, before, after) in [
                (20, current.start_kind, settings.start_kind),
                (21, current.page_number_kind, settings.page_number_kind),
                (22, current.page_number_start, settings.page_number_start),
            ] {
                if before != after {
                    data = patch_varint_field(
                        &data,
                        field_number,
                        before.is_some(),
                        after.map(u64::from),
                    )?;
                }
            }
            if current.name != settings.name {
                data = patch_length_delimited_field(
                    &data,
                    26,
                    current.name.is_some(),
                    settings.name.as_deref().map(str::as_bytes),
                )?;
            }
            if current.background_fill_payload != settings.background_fill_payload {
                data = patch_length_delimited_field(
                    &data,
                    30,
                    current.background_fill_payload.is_some(),
                    settings.background_fill_payload.as_deref(),
                )?;
            }
            if decode_section_settings(&data)? != settings {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {section_id} settings patch failed validation"
                )));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: SECTION_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        *self = Self::from_package(staged)?;
        Ok(())
    }

    /// Read a section background as a semantic solid color when possible.
    pub fn section_background(&self, section_id: u64) -> Result<PagesSectionBackground> {
        let settings = self.section_settings(section_id)?;
        settings
            .background_fill_payload
            .as_deref()
            .map(decode_section_background)
            .transpose()
            .map(|background| background.unwrap_or(PagesSectionBackground::None))
    }

    /// Set, clear, or losslessly replace a section background fill.
    ///
    /// Editing an existing solid color patches only changed nested color
    /// scalars, preserving unknown protobuf fields in the fill and color.
    pub fn set_section_background(
        &mut self,
        section_id: u64,
        background: PagesSectionBackground,
    ) -> Result<()> {
        validate_section_background(&background)?;
        let mut settings = self.section_settings(section_id)?;
        let current = settings
            .background_fill_payload
            .as_deref()
            .map(decode_section_background)
            .transpose()?
            .unwrap_or(PagesSectionBackground::None);
        if current == background {
            return Ok(());
        }

        settings.background_fill_payload = match background {
            PagesSectionBackground::None => None,
            PagesSectionBackground::Opaque(payload) => Some(payload),
            PagesSectionBackground::Solid(color) => {
                let payload = match (current, settings.background_fill_payload.as_deref()) {
                    (PagesSectionBackground::Solid(current), Some(payload)) => {
                        patch_solid_background(payload, current, color)?
                    },
                    _ => encode_solid_background(color),
                };
                Some(payload)
            },
        };
        self.set_section_settings(section_id, settings)
    }

    /// Read the page geometry fields from the Pages document root.
    pub fn page_layout(&self) -> Result<PagesPageLayout> {
        Ok(PagesPageLayout::from(&root_document(self.text.package())?))
    }

    /// Replace the page geometry fields transactionally.
    pub fn set_page_layout(&mut self, layout: PagesPageLayout) -> Result<()> {
        validate_page_layout(&layout)?;
        let mut staged = self.text.package().clone();
        staged.update_archive("Index/Document.iwa", |archive| {
            let object = archive
                .object_mut(1)
                .ok_or_else(|| Error::InvalidFormat("Pages root object 1 is missing".to_owned()))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 10000)
                .ok_or_else(|| {
                    Error::InvalidFormat("Pages root has no TP.DocumentArchive payload".to_owned())
                })?;
            let original = &object.messages[message_index].data;
            let document = DocumentArchive::decode(original.as_slice())?;
            let mut data = original.clone();
            for (field_number, current, replacement) in [
                (30, document.page_width, layout.page_width),
                (31, document.page_height, layout.page_height),
                (32, document.left_margin, layout.left_margin),
                (33, document.right_margin, layout.right_margin),
                (34, document.top_margin, layout.top_margin),
                (35, document.bottom_margin, layout.bottom_margin),
                (36, document.header_margin, layout.header_margin),
                (37, document.footer_margin, layout.footer_margin),
                (38, document.page_scale, layout.page_scale),
            ] {
                data = patch_fixed32_field(
                    &data,
                    field_number,
                    current.is_some(),
                    replacement.map(f32::to_bits),
                )?;
            }
            data = patch_varint_field(
                &data,
                39,
                document.lays_out_body_vertically.is_some(),
                layout.lays_out_body_vertically.map(u64::from),
            )?;
            data = patch_varint_field(
                &data,
                42,
                document.orientation.is_some(),
                layout.orientation.map(u64::from),
            )?;
            let verified = DocumentArchive::decode(data.as_slice())?;
            if PagesPageLayout::from(&verified) != layout {
                return Err(Error::InvalidFormat(
                    "Pages page-layout wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                crate::archive::RawMessage { type_: 10000, data },
            )?;
            Ok(())
        })?;
        *self = Self::from_package(staged)?;
        Ok(())
    }

    /// Set or clear the display name of a reachable Pages section.
    pub fn set_section_name(&mut self, section_id: u64, name: Option<&str>) -> Result<()> {
        let mut settings = self.section_settings(section_id)?;
        settings.name = name.map(str::to_owned);
        self.set_section_settings(section_id, settings)
    }

    /// Insert a native section break at a UTF-16 body position.
    ///
    /// The new section copies the selected source section's page masters,
    /// headers, and footers into an independent private graph, matching Pages.
    /// The body gains the native U+0004 section-break marker at
    /// `character_index`, and the returned section starts immediately after it.
    pub fn insert_section(
        &mut self,
        source_section_id: u64,
        character_index: usize,
        name: &str,
    ) -> Result<PagesSectionInfo> {
        if name.contains('\0') {
            return Err(Error::ParseError(
                "Pages section names cannot contain NUL".to_owned(),
            ));
        }
        if character_index == 0 {
            return Err(Error::ParseError(
                "Pages section breaks must follow body content".to_owned(),
            ));
        }
        let body_units = self.body_text()?.encode_utf16().collect::<Vec<_>>();
        if character_index > body_units.len() {
            return Err(Error::ParseError(format!(
                "Pages section position {character_index} exceeds body length {}",
                body_units.len()
            )));
        }
        if body_units.get(character_index - 1) == Some(&0x0004) {
            return Err(Error::ParseError(
                "Pages body already has a section break at the requested position".to_owned(),
            ));
        }
        if self
            .sections
            .iter()
            .any(|section| usize::try_from(section.character_index).ok() == Some(character_index))
        {
            return Err(Error::ParseError(format!(
                "Pages already has a section boundary at UTF-16 index {character_index}"
            )));
        }

        let source = self
            .sections
            .iter()
            .find(|section| section.object_id == source_section_id)
            .cloned()
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Section {source_section_id} is not reachable from the Pages body"
                ))
            })?;
        let graph = pages_section_graph(self.package(), source_section_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(graph.clone_ids.len());
        for (offset, identifier) in graph.clone_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages section graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &graph.clone_ids {
            let cloned = {
                let archive = staged.archive(&graph.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages section-graph object {identifier} is missing"
                    ))
                })?;
                clone_pages_section_graph_object(source_object, &remap, name)?
            };
            staged.update_archive(&graph.archive_name, |archive| archive.insert_object(cloned))?;
        }
        let new_section_id = remap[&source_section_id];
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Pages section graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = graph
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(&mut staged, 1, &new_uuid_object_ids)?;

        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            character_index..character_index,
            "\u{4}",
        )?;
        staged = text_editor.into_package();
        let boundary = u32::try_from(character_index + 1).map_err(|_| {
            Error::ParseError("Pages section boundary exceeds the u32 UTF-16 range".to_owned())
        })?;
        let entry = ObjectAttribute {
            character_index: boundary,
            object: Some(reference(new_section_id)),
        };
        patch_body_section_table(
            &mut staged,
            self.body_storage_id,
            None,
            Some(new_section_id),
            |table| insert_section_table_entry(table, entry),
        )?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .sections
            .iter()
            .find(|section| section.object_id == new_section_id)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat("Pages section insertion failed validation".to_owned())
            })?;
        let verified_units = verified.body_text()?.encode_utf16().collect::<Vec<_>>();
        if created.character_index != boundary
            || created.name.as_deref() != Some(name)
            || verified_units.get(character_index) != Some(&0x0004)
            || !cloned_optional_identifier(source.first_template_id, created.first_template_id)
            || !cloned_optional_identifier(source.even_template_id, created.even_template_id)
            || !cloned_optional_identifier(source.odd_template_id, created.odd_template_id)
        {
            return Err(Error::InvalidFormat(
                "Pages inserted section did not create an independent native boundary".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Append an empty section at the current end of the body, inheriting a
    /// reachable source section's layout into independent page masters.
    pub fn append_section(
        &mut self,
        source_section_id: u64,
        name: &str,
    ) -> Result<PagesSectionInfo> {
        let character_index = self.body_text()?.encode_utf16().count();
        self.insert_section(source_section_id, character_index, name)
    }

    /// Remove a non-initial section boundary and its now-unreferenced section
    /// archive. Body text is retained and joins the preceding section.
    pub fn remove_section(&mut self, section_id: u64) -> Result<PagesSectionInfo> {
        if self.sections.len() <= 1 {
            return Err(Error::ParseError(
                "Cannot remove the final Pages section".to_owned(),
            ));
        }
        let removed = self
            .sections
            .iter()
            .find(|section| section.object_id == section_id)
            .cloned()
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Section {section_id} is not reachable from the Pages body"
                ))
            })?;
        if removed.character_index == 0 {
            return Err(Error::ParseError(
                "Cannot remove the initial Pages section boundary".to_owned(),
            ));
        }

        let graph = pages_section_graph(self.package(), section_id)?;
        let removed_index = usize::try_from(removed.character_index)
            .map_err(|_| Error::InvalidFormat("Pages section index overflow".to_owned()))?;
        let marker_index = removed_index.checked_sub(1).ok_or_else(|| {
            Error::InvalidFormat("Pages non-initial section has an invalid boundary".to_owned())
        })?;
        let body_units = self.body_text()?.encode_utf16().collect::<Vec<_>>();
        let remove_marker = body_units.get(marker_index) == Some(&0x0004);
        let mut staged = self.package().clone();
        patch_body_section_table(
            &mut staged,
            self.body_storage_id,
            Some(section_id),
            None,
            |table| {
                remove_repeated_length_delimited_field_where(table, 1, |entry| {
                    let entry = ObjectAttribute::decode(entry)?;
                    Ok(entry
                        .object
                        .is_some_and(|object| object.identifier == section_id))
                })
            },
        )?;
        let mut removed_object_ids = Vec::new();
        for identifier in graph.removal_order() {
            if package_references_object(&staged, identifier)? {
                if identifier == section_id {
                    return Err(Error::InvalidFormat(format!(
                        "Pages section object {section_id} remains referenced after removing its boundary"
                    )));
                }
                continue;
            }
            let archive_name = find_object_archive(&staged, identifier)?;
            staged.update_archive(&archive_name, |archive| {
                archive.remove_object(identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages section-graph object {identifier} is missing"
                    ))
                })?;
                Ok(())
            })?;
            removed_object_ids.push(identifier);
        }
        if remove_marker {
            let mut text_editor = IWorkTextEditor::from_package(staged);
            text_editor.replace_text(self.body_storage_id, marker_index..marker_index + 1, "")?;
            staged = text_editor.into_package();
        }
        let removed_uuid_ids = graph
            .uuid_object_ids
            .iter()
            .copied()
            .filter(|identifier| removed_object_ids.contains(identifier))
            .collect::<Vec<_>>();
        remove_component_object_uuids(&mut staged, 1, &removed_uuid_ids)?;
        release_package_identifier_suffix(&mut staged, &removed_object_ids)?;

        let verified = Self::from_package(staged)?;
        if verified
            .sections
            .iter()
            .any(|section| section.object_id == section_id)
            || verified.sections.len() + 1 != self.sections.len()
        {
            return Err(Error::InvalidFormat(
                "Pages section removal failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(removed)
    }

    /// List every header/footer slot reachable from the document's sections.
    ///
    /// A storage can appear in more than one slot when Pages intentionally
    /// shares content between page variants. Editing that storage updates all
    /// aliases, matching Pages' object graph semantics.
    pub fn header_footers(&self) -> Result<Vec<PagesHeaderFooterInfo>> {
        self.header_footers
            .iter()
            .map(|location| {
                Ok(PagesHeaderFooterInfo {
                    section_id: location.section_id,
                    section_name: location.section_name.clone(),
                    section_character_index: location.section_character_index,
                    template_id: location.template_id,
                    template: location.template,
                    kind: location.kind,
                    slot: location.slot,
                    storage: self.text.storage(location.storage_id)?,
                })
            })
            .collect()
    }

    /// Replace a UTF-16 range in a reachable header/footer storage.
    pub fn replace_header_footer_text(
        &mut self,
        storage_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        self.require_header_footer(storage_id)?;
        self.text.replace_text(storage_id, range, replacement)
    }

    /// Set the complete text of a reachable header/footer storage.
    pub fn set_header_footer_text(&mut self, storage_id: u64, replacement: &str) -> Result<()> {
        self.require_header_footer(storage_id)?;
        self.text.set_text(storage_id, replacement)
    }

    /// Clear a reachable header/footer storage without deleting its styled slot.
    pub fn clear_header_footer(&mut self, storage_id: u64) -> Result<()> {
        self.set_header_footer_text(storage_id, "")
    }

    pub fn package(&self) -> &IWorkPackage {
        self.text.package()
    }

    /// List metadata-backed media reachable from this Pages package.
    pub fn media_assets(&self) -> Result<Vec<EmbeddedMediaAsset>> {
        reachable_embedded_assets(self.package(), [1])
    }

    /// List media reachable from one Pages section object graph.
    pub fn section_media_assets(&self, section_id: u64) -> Result<Vec<EmbeddedMediaAsset>> {
        if !self
            .sections()
            .iter()
            .any(|section| section.object_id == section_id)
        {
            return Err(Error::ParseError(format!(
                "Pages section object {section_id} is not reachable"
            )));
        }
        reachable_embedded_assets(self.package(), [section_id])
    }

    pub fn extract_media(&self, data_identifier: u64) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Pages object graph"
            )));
        }
        IWorkMediaEditor::from_package(self.package().clone())?.extract(data_identifier)
    }

    /// Replace a referenced materialized asset without changing its data identifier.
    pub fn replace_media(&mut self, data_identifier: u64, replacement: &[u8]) -> Result<Vec<u8>> {
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Pages object graph"
            )));
        }
        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let old = media.replace(data_identifier, replacement)?;
        let staged = media.into_package();
        Self::from_package(staged.clone())?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(old)
    }

    pub fn into_package(self) -> IWorkPackage {
        self.text.into_package()
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.text.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.text.save(path)
    }

    fn require_header_footer(&self, storage_id: u64) -> Result<()> {
        if self
            .header_footers
            .iter()
            .any(|location| location.storage_id == storage_id)
        {
            Ok(())
        } else {
            Err(Error::ParseError(format!(
                "Text storage {storage_id} is not a reachable Pages header/footer"
            )))
        }
    }

    fn reachable_drawable_ids(&self) -> Result<HashSet<u64>> {
        let document = root_document(self.package())?;
        let mut reachable = metadata_reachable_objects(self.package(), [1, self.body_storage_id])?;

        if let Some(reference) = document.floating_drawables {
            reachable.insert(reference.identifier);
            let floating: tp::FloatingDrawablesArchive = decode_package_object(
                self.package(),
                reference.identifier,
                "TP.FloatingDrawablesArchive",
            )?;
            for group in floating.page_groups {
                for entry in group
                    .background_drawables
                    .into_iter()
                    .chain(group.foreground_drawables)
                    .chain(group.drawables)
                {
                    if let Some(drawable) = entry.drawable {
                        reachable.insert(drawable.identifier);
                    }
                }
            }
            if let Some(pairs) = floating.drawable_tag_pairs {
                reachable.extend(
                    pairs
                        .drawable_tag_pairs
                        .into_iter()
                        .map(|pair| pair.drawable.identifier),
                );
            }
        }

        if let Some(reference) = document.drawables_zorder {
            reachable.insert(reference.identifier);
            let zorder: tp::DrawablesZOrderArchive = decode_package_object(
                self.package(),
                reference.identifier,
                "TP.DrawablesZOrderArchive",
            )?;
            reachable.extend(
                zorder
                    .drawables
                    .into_iter()
                    .map(|drawable| drawable.identifier),
            );
        }

        for reference in document.page_templates {
            reachable.insert(reference.identifier);
            let template: tp::PageTemplateArchive = decode_package_object(
                self.package(),
                reference.identifier,
                "TP.PageTemplateArchive",
            )?;
            reachable.extend(
                template
                    .section_template_drawables
                    .into_iter()
                    .map(|drawable| drawable.identifier),
            );
            reachable.extend(
                template
                    .placeholder_drawables
                    .into_iter()
                    .map(|pair| pair.drawable.identifier),
            );
        }

        let template_ids = self.sections.iter().flat_map(|section| {
            [
                section.first_template_id,
                section.even_template_id,
                section.odd_template_id,
            ]
            .into_iter()
            .flatten()
        });
        for template_id in template_ids {
            reachable.insert(template_id);
            let template: SectionTemplateArchive =
                decode_package_object(self.package(), template_id, "TP.SectionTemplateArchive")?;
            reachable.extend(
                template
                    .section_template_drawables
                    .into_iter()
                    .map(|drawable| drawable.identifier),
            );
        }
        Ok(reachable)
    }

    fn require_drawable(&self, drawable_object_id: u64) -> Result<()> {
        if !self
            .drawables()?
            .iter()
            .any(|drawable| drawable.object_id == drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "drawable object {drawable_object_id} is not reachable from the Pages document"
            )));
        }
        Ok(())
    }

    fn require_drawable_text_storage(&self, drawable_object_id: u64) -> Result<u64> {
        self.drawable_text_storages()?
            .into_iter()
            .find(|text| text.drawable_object_id == drawable_object_id)
            .map(|text| text.storage.object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "drawable object {drawable_object_id} does not own writable text reachable from the Pages document"
                ))
            })
    }

    #[allow(deprecated)]
    fn text_box_graph(&self, drawable_object_id: u64) -> Result<PagesTextBoxGraph> {
        let text = self
            .drawable_text_storages()?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "drawable object {drawable_object_id} is not a reachable Pages text box"
                ))
            })?;
        let shape = decode_typed_package_object::<crate::protobuf::tswp::ShapeInfoArchive>(
            self.package(),
            drawable_object_id,
            SHAPE_INFO_MESSAGE_TYPE,
            "TSWP.ShapeInfoArchive",
        )?;
        if shape.is_text_box != Some(true) {
            return Err(Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not an ordinary text box"
            )));
        }
        if shape.owned_storage.map(|item| item.identifier) != Some(text.storage.object_id)
            || shape
                .deprecated_storage
                .is_some_and(|item| item.identifier != text.storage.object_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} has inconsistent storage ownership"
            )));
        }
        if !shape.super_.super_.pencil_annotations.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} has unsupported pencil annotations"
            )));
        }

        let body: StorageArchive = decode_typed_package_object(
            self.package(),
            self.body_storage_id,
            self.body_storage()?.message_type,
            "TSWP.StorageArchive",
        )?;
        let mut attachments = Vec::new();
        for entry in body
            .table_attachment
            .as_ref()
            .into_iter()
            .flat_map(|table| &table.entries)
        {
            let Some(reference) = entry.object else {
                continue;
            };
            let attachment: DrawableAttachmentArchive = decode_typed_package_object(
                self.package(),
                reference.identifier,
                DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                "TSWP.DrawableAttachmentArchive",
            )?;
            if attachment
                .drawable
                .is_some_and(|item| item.identifier == drawable_object_id)
            {
                attachments.push((entry.character_index, reference.identifier));
            }
        }
        if attachments.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} has {} body attachments; expected one",
                attachments.len()
            )));
        }
        let (anchor_character_index, attachment_id) = attachments[0];
        let body_units = self.body_text()?.encode_utf16().collect::<Vec<_>>();
        if body_units.get(anchor_character_index as usize) != Some(&0xfffc) {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} attachment is not backed by an object-replacement character"
            )));
        }

        let document = root_document(self.package())?;
        let zorder_id = document.drawables_zorder.ok_or_else(|| {
            Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
        })?;
        let zorder: tp::DrawablesZOrderArchive = decode_typed_package_object(
            self.package(),
            zorder_id.identifier,
            10015,
            "TP.DrawablesZOrderArchive",
        )?;
        let zorder_count = zorder
            .drawables
            .iter()
            .filter(|item| item.identifier == drawable_object_id)
            .count();
        if zorder_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} occurs {zorder_count} times in drawable z-order"
            )));
        }

        let mut object_ids = vec![drawable_object_id, text.storage.object_id, attachment_id];
        for reference in [shape.super_.super_.title, shape.super_.super_.caption]
            .into_iter()
            .flatten()
        {
            decode_typed_package_object::<crate::protobuf::tsd::StandinCaptionArchive>(
                self.package(),
                reference.identifier,
                STANDIN_CAPTION_MESSAGE_TYPE,
                "TSD.StandinCaptionArchive",
            )?;
            object_ids.push(reference.identifier);
        }
        let unique = object_ids.iter().copied().collect::<HashSet<_>>();
        if unique.len() != object_ids.len() {
            return Err(Error::InvalidFormat(format!(
                "Pages text box {drawable_object_id} reuses private graph identifiers"
            )));
        }
        let uuid_object_ids = component_uuid_identifiers(self.package(), 1)?
            .map(|identifiers| {
                object_ids
                    .iter()
                    .copied()
                    .filter(|identifier| identifiers.contains(identifier))
                    .collect()
            })
            .unwrap_or_default();
        Ok(PagesTextBoxGraph {
            drawable_id: drawable_object_id,
            storage_id: text.storage.object_id,
            attachment_id,
            anchor_character_index,
            object_ids,
            uuid_object_ids,
        })
    }
}

#[derive(Debug, Clone)]
struct PagesTextBoxGraph {
    drawable_id: u64,
    storage_id: u64,
    attachment_id: u64,
    anchor_character_index: u32,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

#[derive(Debug, Clone)]
struct PagesSectionGraph {
    archive_name: String,
    section_id: u64,
    template_ids: Vec<u64>,
    header_footer_ids: Vec<u64>,
    guide_map_id: Option<u64>,
    guide_storage_ids: Vec<u64>,
    clone_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

impl PagesSectionGraph {
    fn removal_order(&self) -> Vec<u64> {
        let mut result = vec![self.section_id];
        result.extend(self.template_ids.iter().copied());
        result.extend(self.guide_map_id);
        result.extend(self.header_footer_ids.iter().copied());
        result.extend(self.guide_storage_ids.iter().copied());
        result
    }
}

fn decode_typed_package_object<T: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
    type_name: &str,
) -> Result<T> {
    let mut decoded = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        if decoded.is_some() {
            return Err(Error::InvalidFormat(format!(
                "object {identifier} occurs in more than one Pages component"
            )));
        }
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == message_type)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "object {identifier} has no type-{message_type} {type_name} payload"
                ))
            })?;
        decoded = Some(T::decode(message.data.as_slice())?);
    }
    decoded.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "object {identifier} has no decodable {type_name} payload"
        ))
    })
}

fn remap_pages_reference_paths(
    data: &[u8],
    paths: &[&[u32]],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    paths.iter().try_fold(data.to_vec(), |data, path| {
        transform_length_delimited_fields_at_path(&data, path, |reference| {
            let decoded = crate::protobuf::tsp::Reference::decode(reference)?;
            let Some(identifier) = remap.get(&decoded.identifier).copied() else {
                return Ok(reference.to_vec());
            };
            let data = patch_varint_field(reference, 1, true, Some(identifier))?;
            if crate::protobuf::tsp::Reference::decode(data.as_slice())?.identifier != identifier {
                return Err(Error::InvalidFormat(
                    "Pages reference wire remap failed validation".to_owned(),
                ));
            }
            Ok(data)
        })
    })
}

#[allow(deprecated)]
fn remap_pages_shape_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 1, 2],
        &[1, 1, 6],
        &[1, 1, 9],
        &[1, 1, 10],
        &[1, 1, 11],
        &[1, 2],
        &[2],
        &[3],
        &[4],
    ];
    let mut expected = crate::protobuf::tswp::ShapeInfoArchive::decode(data)?;
    remap_pages_shape(&mut expected, remap);
    let data = remap_pages_reference_paths(data, REFERENCE_PATHS, remap)?;
    if crate::protobuf::tswp::ShapeInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages ShapeInfoArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_pages_storage_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const OBJECT_TABLE_FIELDS: &[u32] = &[5, 7, 8, 9, 11, 12, 15, 16, 17, 18, 21, 22, 23, 27, 28];
    let mut expected = StorageArchive::decode(data)?;
    remap_pages_storage(&mut expected, remap);
    let mut data = remap_pages_reference_paths(data, &[&[2]], remap)?;
    for field in OBJECT_TABLE_FIELDS {
        data = remap_pages_reference_paths(&data, &[&[*field, 1, 2]], remap)?;
    }
    for field in [25, 26] {
        data = remap_pages_reference_paths(&data, &[&[field, 1, 2]], remap)?;
    }
    if StorageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages StorageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_pages_attachment_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    let mut expected = DrawableAttachmentArchive::decode(data)?;
    remap_optional_pages_reference(&mut expected.drawable, remap);
    let data = remap_pages_reference_paths(data, &[&[1]], remap)?;
    if DrawableAttachmentArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages DrawableAttachmentArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn clone_pages_text_box_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let old_identifier = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Pages text-box object has no identifier".to_owned())
    })?;
    let new_identifier = *remap.get(&old_identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "No clone identifier allocated for Pages object {old_identifier}"
        ))
    })?;
    let mut messages = Vec::with_capacity(source.messages.len());
    for (message, info) in source
        .messages
        .iter()
        .zip(&source.archive_info.message_infos)
    {
        let data = match message.type_ {
            SHAPE_INFO_MESSAGE_TYPE => remap_pages_shape_wire(&message.data, remap)?,
            2001 | 2022 => remap_pages_storage_wire(&message.data, remap)?,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE => remap_pages_attachment_wire(&message.data, remap)?,
            STANDIN_CAPTION_MESSAGE_TYPE => message.data.clone(),
            _ => {
                if info
                    .object_references
                    .iter()
                    .any(|identifier| remap.contains_key(identifier))
                {
                    return Err(Error::InvalidFormat(format!(
                        "Cannot safely clone Pages message type {} with private text-box references",
                        message.type_
                    )));
                }
                message.data.clone()
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    clone_pages_object_metadata(source, new_identifier, messages, remap)
}

fn clone_pages_object_metadata(
    source: &ArchiveObject,
    new_identifier: u64,
    messages: Vec<RawMessage>,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let mut cloned = ArchiveObject::new(new_identifier, messages)?;
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    for ((target, source), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&source.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_owned())
        })?;
        *target = source.clone();
        target.length = length;
        target.object_references = source
            .object_references
            .iter()
            .map(|identifier| remap.get(identifier).copied().unwrap_or(*identifier))
            .collect();
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

fn offset_text_box_clone(
    package: &mut IWorkPackage,
    drawable_id: u64,
    attachment_id: u64,
    offset: f32,
) -> Result<()> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Pages text-box duplicate offset must be finite".to_owned(),
        ));
    }
    let shape: crate::protobuf::tswp::ShapeInfoArchive = decode_typed_package_object(
        package,
        drawable_id,
        SHAPE_INFO_MESSAGE_TYPE,
        "TSWP.ShapeInfoArchive",
    )?;
    let attachment: DrawableAttachmentArchive = decode_typed_package_object(
        package,
        attachment_id,
        DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        "TSWP.DrawableAttachmentArchive",
    )?;
    if shape
        .super_
        .super_
        .geometry
        .as_ref()
        .and_then(|geometry| geometry.position.as_ref())
        .is_none()
        || attachment.h_offset.is_none()
        || attachment.v_offset.is_none()
    {
        return Ok(());
    }
    let drawable_archive = find_object_archive(package, drawable_id)?;
    package.update_archive(&drawable_archive, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages text-box drawable {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages text-box drawable {drawable_id} must have exactly one shape payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let shape = crate::protobuf::tswp::ShapeInfoArchive::decode(original)?;
        let position = shape
            .super_
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.position.as_ref())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages text-box drawable {drawable_id} has no positioned geometry"
                ))
            })?;
        let x = position.x + offset;
        let y = position.y + offset;
        if !x.is_finite() || !y.is_finite() {
            return Err(Error::ParseError(
                "Pages text-box duplicate position overflow".to_owned(),
            ));
        }
        let data = transform_length_delimited_fields_at_path(original, &[1, 1, 1, 1], |point| {
            let point = patch_fixed32_field(point, 1, true, Some(x.to_bits()))?;
            patch_fixed32_field(&point, 2, true, Some(y.to_bits()))
        })?;
        let verified = crate::protobuf::tswp::ShapeInfoArchive::decode(data.as_slice())?;
        let verified_position = verified
            .super_
            .super_
            .geometry
            .and_then(|geometry| geometry.position)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages text-box geometry patch removed position".to_owned())
            })?;
        if verified_position.x != x || verified_position.y != y {
            return Err(Error::InvalidFormat(
                "Pages text-box geometry offset failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;

    let attachment_archive = find_object_archive(package, attachment_id)?;
    package.update_archive(&attachment_archive, |archive| {
        let object = archive.object_mut(attachment_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages text-box attachment {attachment_id} is missing"
            ))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == DRAWABLE_ATTACHMENT_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages text-box attachment {attachment_id} must have exactly one payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let attachment = DrawableAttachmentArchive::decode(original)?;
        let h_offset = attachment.h_offset.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages text-box attachment {attachment_id} has no horizontal offset"
            ))
        })? + offset;
        let v_offset = attachment.v_offset.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages text-box attachment {attachment_id} has no vertical offset"
            ))
        })? + offset;
        if !h_offset.is_finite() || !v_offset.is_finite() {
            return Err(Error::ParseError(
                "Pages text-box duplicate attachment offset overflow".to_owned(),
            ));
        }
        let data = patch_fixed32_field(original, 3, true, Some(h_offset.to_bits()))?;
        let data = patch_fixed32_field(&data, 5, true, Some(v_offset.to_bits()))?;
        let verified = DrawableAttachmentArchive::decode(data.as_slice())?;
        if verified.h_offset != Some(h_offset) || verified.v_offset != Some(v_offset) {
            return Err(Error::InvalidFormat(
                "Pages text-box attachment offset failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn remap_pages_reference(
    reference: &mut crate::protobuf::tsp::Reference,
    remap: &HashMap<u64, u64>,
) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

fn remap_optional_pages_reference(
    reference: &mut Option<crate::protobuf::tsp::Reference>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(reference) = reference {
        remap_pages_reference(reference, remap);
    }
}

#[allow(deprecated)]
fn remap_pages_shape(
    shape: &mut crate::protobuf::tswp::ShapeInfoArchive,
    remap: &HashMap<u64, u64>,
) {
    let drawable = &mut shape.super_.super_;
    remap_optional_pages_reference(&mut drawable.parent, remap);
    remap_optional_pages_reference(&mut drawable.comment, remap);
    for reference in &mut drawable.pencil_annotations {
        remap_pages_reference(reference, remap);
    }
    remap_optional_pages_reference(&mut drawable.title, remap);
    remap_optional_pages_reference(&mut drawable.caption, remap);
    remap_optional_pages_reference(&mut shape.super_.style, remap);
    remap_optional_pages_reference(&mut shape.deprecated_storage, remap);
    remap_optional_pages_reference(&mut shape.text_flow, remap);
    remap_optional_pages_reference(&mut shape.owned_storage, remap);
}

fn remap_pages_storage(storage: &mut StorageArchive, remap: &HashMap<u64, u64>) {
    remap_optional_pages_reference(&mut storage.style_sheet, remap);
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_attachment,
        &mut storage.table_smartfield,
        &mut storage.table_layout_style,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_section,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_highlight,
        &mut storage.table_tatechuyoko,
        &mut storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_optional_pages_reference(&mut entry.object, remap);
        }
    }
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_pages_reference(&mut entry.field, remap);
        }
    }
}

fn add_body_drawable_attachment(
    package: &mut IWorkPackage,
    body_storage_id: u64,
    anchor_character_index: usize,
    attachment_id: u64,
) -> Result<()> {
    let anchor_character_index = u32::try_from(anchor_character_index)
        .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
    let archive_name = find_object_archive(package, body_storage_id)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(body_storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages body storage {body_storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {body_storage_id} must have exactly one writable payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let storage = StorageArchive::decode(original)?;
        let existing_attachment_ids = storage
            .table_attachment
            .as_ref()
            .into_iter()
            .flat_map(|table| &table.entries)
            .filter_map(|entry| entry.object.map(|reference| reference.identifier))
            .collect::<HashSet<_>>();
        let entry = ObjectAttribute {
            character_index: anchor_character_index,
            object: Some(reference(attachment_id)),
        };
        let data = if storage.table_attachment.is_some() {
            transform_length_delimited_field(original, 9, |table| {
                let mut entries = repeated_length_delimited_payloads(table, 1)?
                    .into_iter()
                    .enumerate()
                    .map(|(order, data)| {
                        let decoded = ObjectAttribute::decode(data)?;
                        Ok((decoded.character_index, order, data.to_vec()))
                    })
                    .collect::<Result<Vec<_>>>()?;
                if entries
                    .iter()
                    .any(|(character_index, _, _)| *character_index == anchor_character_index)
                {
                    return Err(Error::InvalidFormat(format!(
                        "Pages body already has an attachment at UTF-16 index {anchor_character_index}"
                    )));
                }
                entries.push((
                    anchor_character_index,
                    entries.len(),
                    entry.encode_to_vec(),
                ));
                entries.sort_by_key(|(character_index, order, _)| (*character_index, *order));
                let entries = entries
                    .into_iter()
                    .map(|(_, _, data)| data)
                    .collect::<Vec<_>>();
                rewrite_repeated_length_delimited_fields(table, 1, &entries)
            })?
        } else {
            let table = crate::protobuf::tswp::ObjectAttributeTable {
                entries: vec![entry],
            };
            patch_length_delimited_field(original, 9, false, Some(&table.encode_to_vec()))?
        };
        let verified = StorageArchive::decode(data.as_slice())?;
        let matches = verified
            .table_attachment
            .as_ref()
            .into_iter()
            .flat_map(|table| &table.entries)
            .filter(|entry| {
                entry.character_index == anchor_character_index
                    && entry
                        .object
                        .is_some_and(|reference| reference.identifier == attachment_id)
            })
            .count();
        if matches != 1 {
            return Err(Error::InvalidFormat(
                "Pages body attachment insertion failed validation".to_owned(),
            ));
        }
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[message_index];
        update_reference_list(&mut info.object_references, None, Some(attachment_id));
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| existing_attachment_ids.contains(identifier))
            {
                update_reference_list(&mut field.object_references, None, Some(attachment_id));
            }
        }
        Ok(())
    })
}

fn patch_pages_zorder(
    package: &mut IWorkPackage,
    removed: Option<u64>,
    added: Option<u64>,
) -> Result<()> {
    let zorder_reference = root_document(package)?.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let archive_name = find_object_archive(package, zorder_reference.identifier)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive
            .object_mut(zorder_reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages drawable z-order object {} is missing",
                    zorder_reference.identifier
                ))
            })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 10015)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(
                "Pages drawable z-order must have exactly one type-10015 payload".to_owned(),
            ));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let previous = tp::DrawablesZOrderArchive::decode(original)?;
        let previous_ids = previous
            .drawables
            .iter()
            .map(|item| item.identifier)
            .collect::<HashSet<_>>();
        if removed.is_some_and(|identifier| {
            previous
                .drawables
                .iter()
                .filter(|item| item.identifier == identifier)
                .count()
                != 1
        }) || added.is_some_and(|identifier| {
            previous
                .drawables
                .iter()
                .any(|item| item.identifier == identifier)
        }) {
            return Err(Error::InvalidFormat(
                "Pages drawable z-order ownership is ambiguous".to_owned(),
            ));
        }
        let mut data = original.to_vec();
        if let Some(identifier) = removed {
            data = remove_repeated_length_delimited_field_where(&data, 1, |payload| {
                Ok(crate::protobuf::tsp::Reference::decode(payload)?.identifier == identifier)
            })?;
        }
        if let Some(identifier) = added {
            data = append_repeated_length_delimited_field(
                &data,
                1,
                &reference(identifier).encode_to_vec(),
            )?;
        }
        let verified = tp::DrawablesZOrderArchive::decode(data.as_slice())?;
        if removed.is_some_and(|identifier| {
            verified
                .drawables
                .iter()
                .any(|item| item.identifier == identifier)
        }) || added.is_some_and(|identifier| {
            verified
                .drawables
                .iter()
                .filter(|item| item.identifier == identifier)
                .count()
                != 1
        }) {
            return Err(Error::InvalidFormat(
                "Pages drawable z-order patch failed validation".to_owned(),
            ));
        }
        object.replace_message(message_index, RawMessage { type_: 10015, data })?;
        let info = &mut object.archive_info.message_infos[message_index];
        update_reference_list(&mut info.object_references, removed, added);
        for field in &mut info.field_infos {
            let contains_zorder = field
                .object_references
                .iter()
                .any(|identifier| previous_ids.contains(identifier));
            update_reference_list(&mut field.object_references, removed, None);
            if contains_zorder {
                update_reference_list(&mut field.object_references, None, added);
            }
        }
        Ok(())
    })
}

fn drawable_owned_text_storage(
    package: &IWorkPackage,
    drawable: &IWorkDrawableInfo,
) -> Result<Option<u64>> {
    if !matches!(
        drawable.message_type,
        PLACEHOLDER_MESSAGE_TYPE | SHAPE_INFO_MESSAGE_TYPE
    ) {
        return Ok(None);
    }

    let mut owned_storage = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(drawable.object_id) else {
            continue;
        };
        if owned_storage.is_some() {
            return Err(Error::InvalidFormat(format!(
                "drawable object {} occurs in more than one Pages component",
                drawable.object_id
            )));
        }
        let payload = &object
            .messages
            .iter()
            .find(|message| message.type_ == drawable.message_type)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages drawable {} has no type-{} payload",
                    drawable.object_id, drawable.message_type
                ))
            })?
            .data;
        owned_storage = Some(match drawable.message_type {
            PLACEHOLDER_MESSAGE_TYPE => {
                tp::PlaceholderArchive::decode(payload.as_slice())?
                    .super_
                    .owned_storage
            },
            SHAPE_INFO_MESSAGE_TYPE => {
                crate::protobuf::tswp::ShapeInfoArchive::decode(payload.as_slice())?.owned_storage
            },
            _ => unreachable!("message type was checked above"),
        });
    }
    owned_storage
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable object {} is missing",
                drawable.object_id
            ))
        })
        .map(|storage| storage.map(|reference| reference.identifier))
}

fn decode_package_object<T: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    type_name: &str,
) -> Result<T> {
    let mut decoded = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        if decoded.is_some() {
            return Err(Error::InvalidFormat(format!(
                "object {identifier} occurs in more than one Pages component"
            )));
        }
        decoded = object
            .messages
            .iter()
            .find_map(|message| T::decode(message.data.as_slice()).ok());
    }
    decoded.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "object {identifier} has no decodable {type_name} payload"
        ))
    })
}

fn metadata_reachable_objects(
    package: &IWorkPackage,
    roots: impl IntoIterator<Item = u64>,
) -> Result<HashSet<u64>> {
    let mut references = HashMap::<u64, Vec<u64>>::new();
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            let outgoing = references.entry(identifier).or_default();
            for info in object.archive_info.message_infos {
                outgoing.extend(info.object_references);
                for field in info.field_infos {
                    outgoing.extend(field.object_references);
                }
            }
            outgoing.sort_unstable();
            outgoing.dedup();
        }
    }
    let mut reachable = HashSet::new();
    let mut pending = roots.into_iter().collect::<Vec<_>>();
    while let Some(identifier) = pending.pop() {
        if !reachable.insert(identifier) {
            continue;
        }
        if let Some(outgoing) = references.get(&identifier) {
            pending.extend(outgoing.iter().copied());
        }
    }
    Ok(reachable)
}

fn reference(identifier: u64) -> crate::protobuf::tsp::Reference {
    crate::protobuf::tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn cloned_optional_identifier(source: Option<u64>, cloned: Option<u64>) -> bool {
    match source {
        Some(source) => cloned.is_some_and(|cloned| cloned != source),
        None => cloned.is_none(),
    }
}

fn pages_section_graph(package: &IWorkPackage, section_id: u64) -> Result<PagesSectionGraph> {
    let archive_name = find_section_archive(package, section_id)?;
    let section: SectionArchive = decode_typed_package_object(
        package,
        section_id,
        SECTION_MESSAGE_TYPE,
        "TP.SectionArchive",
    )?;
    let mut template_ids = [
        section.first_section_template_page,
        section.even_section_template_page,
        section.odd_section_template_page,
    ]
    .into_iter()
    .flatten()
    .map(|reference| reference.identifier)
    .collect::<Vec<_>>();
    template_ids.sort_unstable();
    template_ids.dedup();

    let mut header_footer_ids = section
        .obsolete_headers
        .into_iter()
        .chain(section.obsolete_footers)
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
    for template_id in &template_ids {
        let template: SectionTemplateArchive = decode_typed_package_object(
            package,
            *template_id,
            SECTION_TEMPLATE_MESSAGE_TYPE,
            "TP.SectionTemplateArchive",
        )?;
        header_footer_ids.extend(
            template
                .headers
                .into_iter()
                .chain(template.footers)
                .map(|reference| reference.identifier),
        );
    }
    header_footer_ids.sort_unstable();
    header_footer_ids.dedup();
    for identifier in &header_footer_ids {
        let object = package.archive(&find_object_archive(package, *identifier)?)?;
        let object = object.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages header/footer object {identifier} is missing"
            ))
        })?;
        if object.messages.len() != 1
            || !STORAGE_MESSAGE_TYPES.contains(&object.messages[0].type_)
            || StorageArchive::decode(object.messages[0].data.as_slice()).is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "Pages section header/footer {identifier} is not a single writable storage"
            )));
        }
    }

    let guide_map_id = section
        .user_defined_guide_storage
        .map(|reference| reference.identifier);
    let mut guide_storage_ids = Vec::new();
    if let Some(guide_map_id) = guide_map_id {
        let guide_map: tp::UserDefinedGuideMapArchive = decode_typed_package_object(
            package,
            guide_map_id,
            USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE,
            "TP.UserDefinedGuideMapArchive",
        )?;
        guide_storage_ids.extend(
            guide_map
                .user_defined_guide_storages
                .into_iter()
                .map(|entry| entry.guide_storage.identifier),
        );
        guide_storage_ids.sort_unstable();
        guide_storage_ids.dedup();
        for identifier in &guide_storage_ids {
            let _: crate::protobuf::tsd::GuideStorageArchive = decode_typed_package_object(
                package,
                *identifier,
                GUIDE_STORAGE_MESSAGE_TYPE,
                "TSD.GuideStorageArchive",
            )?;
        }
    }

    let mut clone_ids = vec![section_id];
    clone_ids.extend(template_ids.iter().copied());
    clone_ids.extend(header_footer_ids.iter().copied());
    clone_ids.extend(guide_map_id);
    clone_ids.extend(guide_storage_ids.first().copied());
    let unique = clone_ids.iter().copied().collect::<HashSet<_>>();
    if unique.len() != clone_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages section {section_id} reuses incompatible private graph identifiers"
        )));
    }
    for identifier in &clone_ids {
        let object_archive = find_object_archive(package, *identifier)?;
        if object_archive != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Pages section object {identifier} is split across {object_archive} instead of {archive_name}"
            )));
        }
    }
    let uuid_object_ids = component_uuid_identifiers(package, 1)?
        .map(|identifiers| {
            clone_ids
                .iter()
                .copied()
                .filter(|identifier| identifiers.contains(identifier))
                .collect()
        })
        .unwrap_or_default();
    Ok(PagesSectionGraph {
        archive_name,
        section_id,
        template_ids,
        header_footer_ids,
        guide_map_id,
        guide_storage_ids,
        clone_ids,
        uuid_object_ids,
    })
}

fn clone_pages_section_graph_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
    name: &str,
) -> Result<ArchiveObject> {
    let old_identifier = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Pages section-graph object has no identifier".to_owned())
    })?;
    let new_identifier = *remap.get(&old_identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "No clone identifier allocated for Pages section object {old_identifier}"
        ))
    })?;
    if source.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Cannot safely clone multi-payload Pages section object {old_identifier}"
        )));
    }
    let message = &source.messages[0];
    let data = match message.type_ {
        SECTION_MESSAGE_TYPE => {
            const REFERENCE_PATHS: &[&[u32]] = &[&[3], &[4], &[14], &[23], &[24], &[25], &[29]];
            let section = SectionArchive::decode(message.data.as_slice())?;
            let data = remap_pages_reference_paths(&message.data, REFERENCE_PATHS, remap)?;
            let data = patch_length_delimited_field(
                &data,
                26,
                section.name.is_some(),
                Some(name.as_bytes()),
            )?;
            if SectionArchive::decode(data.as_slice())?.name.as_deref() != Some(name) {
                return Err(Error::InvalidFormat(
                    "Pages section clone name patch failed validation".to_owned(),
                ));
            }
            data
        },
        SECTION_TEMPLATE_MESSAGE_TYPE => {
            remap_pages_reference_paths(&message.data, &[&[1], &[2], &[3]], remap)?
        },
        message_type if STORAGE_MESSAGE_TYPES.contains(&message_type) => {
            remap_pages_storage_wire(&message.data, remap)?
        },
        USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE => {
            let entries = repeated_length_delimited_payloads(&message.data, 1)?;
            let replacements = if let Some(entry) = entries.first() {
                let entry = patch_varint_field(entry, 1, true, Some(0))?;
                vec![remap_pages_reference_paths(&entry, &[&[2]], remap)?]
            } else {
                Vec::new()
            };
            let data = rewrite_repeated_length_delimited_fields(&message.data, 1, &replacements)?;
            let verified = tp::UserDefinedGuideMapArchive::decode(data.as_slice())?;
            if verified.user_defined_guide_storages.len() != replacements.len()
                || verified
                    .user_defined_guide_storages
                    .first()
                    .is_some_and(|entry| entry.page_index != 0)
            {
                return Err(Error::InvalidFormat(
                    "Pages guide-map clone failed validation".to_owned(),
                ));
            }
            data
        },
        GUIDE_STORAGE_MESSAGE_TYPE => {
            let data = rewrite_repeated_length_delimited_fields(&message.data, 1, &[])?;
            if !crate::protobuf::tsd::GuideStorageArchive::decode(data.as_slice())?
                .user_defined_guides
                .is_empty()
            {
                return Err(Error::InvalidFormat(
                    "Pages guide-storage clone failed validation".to_owned(),
                ));
            }
            data
        },
        _ => {
            return Err(Error::InvalidFormat(format!(
                "Cannot safely clone Pages section message type {}",
                message.type_
            )));
        },
    };
    let mut cloned = clone_pages_object_metadata(
        source,
        new_identifier,
        vec![RawMessage {
            type_: message.type_,
            data,
        }],
        remap,
    )?;
    if message.type_ == USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE {
        let guide_ids = tp::UserDefinedGuideMapArchive::decode(cloned.messages[0].data.as_slice())?
            .user_defined_guide_storages
            .into_iter()
            .map(|entry| entry.guide_storage.identifier)
            .collect::<Vec<_>>();
        let info = &mut cloned.archive_info.message_infos[0];
        info.object_references = guide_ids.clone();
        for field in &mut info.field_infos {
            if !field.object_references.is_empty() {
                field.object_references = guide_ids.clone();
            }
        }
    }
    Ok(cloned)
}

fn insert_section_table_entry(table: &[u8], entry: ObjectAttribute) -> Result<Vec<u8>> {
    let mut entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .enumerate()
        .map(|(order, data)| {
            let decoded = ObjectAttribute::decode(data)?;
            Ok((decoded.character_index, order, data.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    if entries
        .iter()
        .any(|(character_index, _, _)| *character_index == entry.character_index)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages already has a section at UTF-16 index {}",
            entry.character_index
        )));
    }
    entries.push((entry.character_index, entries.len(), entry.encode_to_vec()));
    entries.sort_by_key(|(character_index, order, _)| (*character_index, *order));
    let entries = entries
        .into_iter()
        .map(|(_, _, data)| data)
        .collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, 1, &entries)
}

fn patch_body_section_table<F>(
    package: &mut IWorkPackage,
    body_storage_id: u64,
    removed: Option<u64>,
    added: Option<u64>,
    update: F,
) -> Result<()>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    let archive_name = find_object_archive(package, body_storage_id)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(body_storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages body storage object {body_storage_id} is missing"
            ))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .filter(|(_, message)| StorageArchive::decode(message.data.as_slice()).is_ok())
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {body_storage_id} must have exactly one writable payload"
            )));
        }
        let message_index = indexes[0];
        let storage = StorageArchive::decode(object.messages[message_index].data.as_slice())?;
        let section_ids = storage
            .table_section
            .as_ref()
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages body storage {body_storage_id} has no section table"
                ))
            })?
            .entries
            .iter()
            .filter_map(|entry| entry.object.map(|object| object.identifier))
            .collect::<HashSet<_>>();
        let data = transform_length_delimited_field(
            object.messages[message_index].data.as_slice(),
            17,
            update,
        )?;
        let verified = StorageArchive::decode(data.as_slice())?;
        let verified_ids = verified
            .table_section
            .as_ref()
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages body storage {body_storage_id} lost its section table"
                ))
            })?
            .entries
            .iter()
            .filter_map(|entry| entry.object.map(|object| object.identifier))
            .collect::<Vec<_>>();
        if removed.is_some_and(|identifier| verified_ids.contains(&identifier))
            || added.is_some_and(|identifier| !verified_ids.contains(&identifier))
        {
            return Err(Error::InvalidFormat(
                "Pages body section-table patch failed validation".to_owned(),
            ));
        }
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[message_index];
        update_reference_list(&mut info.object_references, removed, added);
        for field in &mut info.field_infos {
            let contains_section = field
                .object_references
                .iter()
                .any(|identifier| section_ids.contains(identifier));
            update_reference_list(&mut field.object_references, removed, None);
            if contains_section {
                update_reference_list(&mut field.object_references, None, added);
            }
        }
        Ok(())
    })
}

fn update_reference_list(references: &mut Vec<u64>, old: Option<u64>, new: Option<u64>) {
    if let Some(old) = old {
        references.retain(|identifier| *identifier != old);
    }
    if let Some(new) = new
        && !references.contains(&new)
    {
        references.push(new);
    }
}

fn find_object_archive(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_none() {
            continue;
        }
        if found.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "Object {identifier} occurs in multiple IWA components"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))
}

fn package_references_object(package: &IWorkPackage, identifier: u64) -> Result<bool> {
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        if archive.objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|message| {
                message.object_references.contains(&identifier)
                    || message
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            })
        }) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn body_storage_id(package: &IWorkPackage) -> Result<u64> {
    root_document(package)?
        .body_storage
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Pages document has no body storage".to_owned()))
}

fn root_document(package: &IWorkPackage) -> Result<DocumentArchive> {
    let archive = package.archive("Index/Document.iwa")?;
    let object = archive
        .object(1)
        .ok_or_else(|| Error::InvalidFormat("Pages root object 1 is missing".to_owned()))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == 10000)
        .and_then(|message| DocumentArchive::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat("Pages root has no TP.DocumentArchive payload".to_owned())
        })
}

fn validate_page_layout(layout: &PagesPageLayout) -> Result<()> {
    for (name, value) in [
        ("page width", layout.page_width),
        ("page height", layout.page_height),
        ("page scale", layout.page_scale),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value <= 0.0) {
            return Err(Error::ParseError(format!(
                "Pages {name} must be finite and greater than zero"
            )));
        }
    }
    for (name, value) in [
        ("left margin", layout.left_margin),
        ("right margin", layout.right_margin),
        ("top margin", layout.top_margin),
        ("bottom margin", layout.bottom_margin),
        ("header margin", layout.header_margin),
        ("footer margin", layout.footer_margin),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value < 0.0) {
            return Err(Error::ParseError(format!(
                "Pages {name} must be finite and non-negative"
            )));
        }
    }
    if let (Some(width), Some(left), Some(right)) =
        (layout.page_width, layout.left_margin, layout.right_margin)
        && left + right >= width
    {
        return Err(Error::ParseError(
            "Pages horizontal margins must leave positive body width".to_owned(),
        ));
    }
    if let (Some(height), Some(top), Some(bottom)) =
        (layout.page_height, layout.top_margin, layout.bottom_margin)
        && top + bottom >= height
    {
        return Err(Error::ParseError(
            "Pages vertical margins must leave positive body height".to_owned(),
        ));
    }
    Ok(())
}

fn validate_section_settings(settings: &PagesSectionSettings) -> Result<()> {
    if settings
        .name
        .as_deref()
        .is_some_and(|name| name.contains('\0'))
    {
        return Err(Error::ParseError(
            "Pages section names cannot contain NUL".to_owned(),
        ));
    }
    if let Some(payload) = settings.background_fill_payload.as_deref() {
        tsd::FillArchive::decode(payload).map_err(|error| {
            Error::ParseError(format!(
                "Pages section background fill is not a TSD.FillArchive: {error}"
            ))
        })?;
    }
    Ok(())
}

fn decode_section_settings(data: &[u8]) -> Result<PagesSectionSettings> {
    let section = SectionArchive::decode(data)?;

    // Validate raw singularity and wire types even though prost accepts the
    // final occurrence of duplicate scalar fields.
    for (field_number, present, value) in [
        (
            17,
            section.inherit_previous_header_footer.is_some(),
            section.inherit_previous_header_footer.map(u64::from),
        ),
        (
            18,
            section.section_template_first_page_different.is_some(),
            section.section_template_first_page_different.map(u64::from),
        ),
        (
            19,
            section.section_template_even_odd_pages_different.is_some(),
            section
                .section_template_even_odd_pages_different
                .map(u64::from),
        ),
        (
            20,
            section.section_start_kind.is_some(),
            section.section_start_kind.map(u64::from),
        ),
        (
            21,
            section.section_page_number_kind.is_some(),
            section.section_page_number_kind.map(u64::from),
        ),
        (
            22,
            section.section_page_number_start.is_some(),
            section.section_page_number_start.map(u64::from),
        ),
        (
            28,
            section
                .section_template_first_page_hides_header_footer
                .is_some(),
            section
                .section_template_first_page_hides_header_footer
                .map(u64::from),
        ),
    ] {
        patch_varint_field(data, field_number, present, value)?;
    }
    patch_length_delimited_field(
        data,
        26,
        section.name.is_some(),
        section.name.as_deref().map(str::as_bytes),
    )?;

    let background_payloads = repeated_length_delimited_payloads(data, 30)?;
    if background_payloads.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field 30 occurs {} times",
            background_payloads.len()
        )));
    }
    if background_payloads.is_empty() != section.background_fill.is_none() {
        return Err(Error::InvalidFormat(
            "Pages section background-fill presence changed during decoding".to_owned(),
        ));
    }
    let background_fill_payload = background_payloads.first().map(|payload| payload.to_vec());
    if let Some(payload) = background_fill_payload.as_deref() {
        tsd::FillArchive::decode(payload)?;
    }
    patch_length_delimited_field(
        data,
        30,
        section.background_fill.is_some(),
        background_fill_payload.as_deref(),
    )?;

    Ok(PagesSectionSettings {
        name: section.name,
        inherit_previous_header_footer: section.inherit_previous_header_footer,
        first_page_different: section.section_template_first_page_different,
        even_odd_pages_different: section.section_template_even_odd_pages_different,
        start_kind: section.section_start_kind,
        page_number_kind: section.section_page_number_kind,
        page_number_start: section.section_page_number_start,
        first_page_hides_header_footer: section.section_template_first_page_hides_header_footer,
        background_fill_payload,
    })
}

fn validate_section_background(background: &PagesSectionBackground) -> Result<()> {
    match background {
        PagesSectionBackground::None => Ok(()),
        PagesSectionBackground::Solid(color) => validate_pages_color(*color),
        PagesSectionBackground::Opaque(payload) => {
            tsd::FillArchive::decode(payload.as_slice()).map_err(|error| {
                Error::ParseError(format!(
                    "Opaque Pages section background is not a TSD.FillArchive: {error}"
                ))
            })?;
            Ok(())
        },
    }
}

fn validate_pages_color(color: PagesRgbaColor) -> Result<()> {
    for (name, value) in [
        ("red", color.red),
        ("green", color.green),
        ("blue", color.blue),
        ("alpha", color.alpha),
    ] {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(Error::ParseError(format!(
                "Pages section background {name} must be finite and between zero and one"
            )));
        }
    }
    Ok(())
}

fn decode_section_background(payload: &[u8]) -> Result<PagesSectionBackground> {
    let fill = tsd::FillArchive::decode(payload)?;
    let Some(color) = fill.color.as_ref() else {
        return Ok(PagesSectionBackground::Opaque(payload.to_vec()));
    };
    if fill.gradient.is_some()
        || fill.image.is_some()
        || color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return Ok(PagesSectionBackground::Opaque(payload.to_vec()));
    }
    let Some((red, green, blue)) = color
        .r
        .zip(color.g)
        .zip(color.b)
        .map(|((r, g), b)| (r, g, b))
    else {
        return Ok(PagesSectionBackground::Opaque(payload.to_vec()));
    };
    let color_space = match color.rgbspace {
        None => PagesRgbColorSpace::Srgb,
        Some(value) if value == tsp::color::RgbColorSpace::Srgb as i32 => PagesRgbColorSpace::Srgb,
        Some(value) if value == tsp::color::RgbColorSpace::P3 as i32 => {
            PagesRgbColorSpace::DisplayP3
        },
        _ => return Ok(PagesSectionBackground::Opaque(payload.to_vec())),
    };
    let semantic = PagesRgbaColor {
        red,
        green,
        blue,
        alpha: color.a.unwrap_or(1.0),
        color_space,
    };
    if validate_pages_color(semantic).is_err() {
        return Ok(PagesSectionBackground::Opaque(payload.to_vec()));
    }
    Ok(PagesSectionBackground::Solid(semantic))
}

fn encode_solid_background(color: PagesRgbaColor) -> Vec<u8> {
    tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(color.red),
            g: Some(color.green),
            b: Some(color.blue),
            rgbspace: Some(match color.color_space {
                PagesRgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
                PagesRgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
            }),
            a: Some(color.alpha),
            ..Default::default()
        }),
        ..Default::default()
    }
    .encode_to_vec()
}

fn patch_solid_background(
    payload: &[u8],
    current: PagesRgbaColor,
    replacement: PagesRgbaColor,
) -> Result<Vec<u8>> {
    let fill = tsd::FillArchive::decode(payload)?;
    let color = fill.color.ok_or_else(|| {
        Error::InvalidFormat("Pages solid section background lost its color".to_owned())
    })?;
    let mut data = payload.to_vec();
    for (field_number, present, before, after) in [
        (3, color.r.is_some(), current.red, replacement.red),
        (4, color.g.is_some(), current.green, replacement.green),
        (5, color.b.is_some(), current.blue, replacement.blue),
        (6, color.a.is_some(), current.alpha, replacement.alpha),
    ] {
        if before != after {
            data = patch_nested_fixed32_field(
                &data,
                &[1, field_number],
                present,
                Some(after.to_bits()),
            )?;
        }
    }
    if current.color_space != replacement.color_space {
        let rgbspace = match replacement.color_space {
            PagesRgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as u64,
            PagesRgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as u64,
        };
        data =
            patch_nested_varint_field(&data, &[1, 12], color.rgbspace.is_some(), Some(rgbspace))?;
    }
    if decode_section_background(&data)? != PagesSectionBackground::Solid(replacement) {
        return Err(Error::InvalidFormat(
            "Pages solid section background patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn discover_structure(
    package: &IWorkPackage,
    body_storage_id: u64,
) -> Result<(Vec<PagesSectionInfo>, Vec<HeaderFooterLocation>)> {
    let document_archive = package.archive("Index/Document.iwa")?;
    let document = document_archive
        .object(1)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 10000)
        })
        .and_then(|message| DocumentArchive::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat("Pages root has no TP.DocumentArchive payload".to_owned())
        })?;

    let mut body = None;
    let mut sections = HashMap::new();
    let mut templates = HashMap::new();
    let mut writable_storages = HashSet::new();
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            for message in object.messages {
                match message.type_ {
                    message_type if STORAGE_MESSAGE_TYPES.contains(&message_type) => {
                        let Ok(storage) = StorageArchive::decode(message.data.as_slice()) else {
                            continue;
                        };
                        writable_storages.insert(identifier);
                        if identifier == body_storage_id {
                            body = Some(storage);
                        }
                    },
                    SECTION_MESSAGE_TYPE => {
                        if let Ok(section) = SectionArchive::decode(message.data.as_slice()) {
                            insert_unique(&mut sections, identifier, section, "section")?;
                        }
                    },
                    SECTION_TEMPLATE_MESSAGE_TYPE => {
                        if let Ok(template) =
                            SectionTemplateArchive::decode(message.data.as_slice())
                        {
                            insert_unique(
                                &mut templates,
                                identifier,
                                template,
                                "section template",
                            )?;
                        }
                    },
                    _ => {},
                }
            }
        }
    }

    if !writable_storages.contains(&body_storage_id) {
        return Err(Error::InvalidFormat(format!(
            "Pages body object {body_storage_id} has no writable text storage"
        )));
    }
    let body = body.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages body storage object {body_storage_id} could not be decoded"
        ))
    })?;

    let mut section_references = Vec::new();
    if let Some(reference) = document.section {
        section_references.push((0, reference.identifier));
    }
    if let Some(table) = body.table_section {
        section_references.extend(table.entries.into_iter().filter_map(|entry| {
            entry
                .object
                .map(|reference| (entry.character_index, reference.identifier))
        }));
    }
    section_references.sort_unstable();
    section_references.dedup();
    if section_references
        .first()
        .is_some_and(|(character_index, _)| *character_index != 0)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages initial section boundary starts at UTF-16 index {} instead of zero",
            section_references[0].0
        )));
    }
    let body_length = body
        .text
        .iter()
        .map(|text| text.encode_utf16().count())
        .sum::<usize>();
    let mut seen_sections = HashSet::new();
    for (index, (character_index, section_id)) in section_references.iter().enumerate() {
        let character_index_usize = usize::try_from(*character_index).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages section {section_id} boundary exceeds the platform index range"
            ))
        })?;
        if character_index_usize > body_length {
            return Err(Error::InvalidFormat(format!(
                "Pages section {section_id} boundary {character_index} exceeds body length {body_length}"
            )));
        }
        if index > 0 && section_references[index - 1].0 == *character_index {
            return Err(Error::InvalidFormat(format!(
                "Pages has multiple section boundaries at UTF-16 index {character_index}"
            )));
        }
        if !seen_sections.insert(*section_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages section object {section_id} is attached at multiple boundaries"
            )));
        }
    }

    let mut section_infos = Vec::with_capacity(section_references.len());
    let mut locations = Vec::new();
    for (character_index, section_id) in section_references {
        let section = sections.get(&section_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages section object {section_id} is missing or invalid"
            ))
        })?;
        section_infos.push(PagesSectionInfo {
            object_id: section_id,
            character_index,
            name: section.name.clone(),
            first_template_id: section
                .first_section_template_page
                .as_ref()
                .map(|reference| reference.identifier),
            even_template_id: section
                .even_section_template_page
                .as_ref()
                .map(|reference| reference.identifier),
            odd_template_id: section
                .odd_section_template_page
                .as_ref()
                .map(|reference| reference.identifier),
        });
        for (template_kind, reference) in [
            (
                PagesTemplateKind::First,
                section.first_section_template_page.as_ref(),
            ),
            (
                PagesTemplateKind::Even,
                section.even_section_template_page.as_ref(),
            ),
            (
                PagesTemplateKind::Odd,
                section.odd_section_template_page.as_ref(),
            ),
        ] {
            let Some(reference) = reference else {
                continue;
            };
            let template_id = reference.identifier;
            let template = templates.get(&template_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages section template object {template_id} is missing or invalid"
                ))
            })?;
            for (kind, references) in [
                (PagesHeaderFooterKind::Header, &template.headers),
                (PagesHeaderFooterKind::Footer, &template.footers),
            ] {
                for (slot, reference) in references.iter().enumerate() {
                    if !writable_storages.contains(&reference.identifier) {
                        return Err(Error::InvalidFormat(format!(
                            "Pages {kind:?} storage {} is missing or invalid",
                            reference.identifier
                        )));
                    }
                    locations.push(HeaderFooterLocation {
                        section_id,
                        section_name: section.name.clone(),
                        section_character_index: character_index,
                        template_id,
                        template: template_kind,
                        kind,
                        slot,
                        storage_id: reference.identifier,
                    });
                }
            }
        }
    }
    Ok((section_infos, locations))
}

fn find_section_archive(package: &IWorkPackage, section_id: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(section_id) else {
            continue;
        };
        if !object
            .messages
            .iter()
            .any(|message| message.type_ == SECTION_MESSAGE_TYPE)
        {
            continue;
        }
        if found.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "Pages section object {section_id} occurs in multiple IWA components"
            )));
        }
    }
    found.ok_or_else(|| {
        Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
    })
}

fn insert_unique<T>(
    values: &mut HashMap<u64, T>,
    identifier: u64,
    value: T,
    kind: &str,
) -> Result<()> {
    if values.insert(identifier, value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "Pages {kind} object {identifier} occurs more than once"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests;
