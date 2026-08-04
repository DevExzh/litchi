//! Semantic editing of reachable Pages body, header, footer, and drawable storages.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::comments::{
    CommentStorageId, DrawableCommentInfo, DrawableCommentReplyInfo, DrawableObjectId,
    IWorkDrawableCommentEditor, IWorkDrawableInfo,
};
use crate::media::reachable_embedded_assets;
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    component_uuid_identifiers, next_object_identifier, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
    remove_component_registration, set_package_last_object_identifier,
};
use crate::protobuf::tp::{self, DocumentArchive, SectionArchive, SectionTemplateArchive};
use crate::protobuf::tsd;
use crate::protobuf::tsp;
use crate::protobuf::tswp::{
    DrawableAttachmentArchive, StorageArchive, object_attribute_table::ObjectAttribute,
};
use crate::shapes::ShapeTextLayout;
use crate::shapes::{
    DrawableGeometry, DrawableProperties, RgbaColor, reset_shape_text_columns,
    reset_shape_text_layout, set_shape_geometry, set_shape_properties, set_shape_text_columns,
    set_shape_text_layout, shape_geometry, shape_properties, shape_text_columns, shape_text_layout,
};
use crate::text::{
    AppliedParagraphStyle, IWorkTextEditor, NamedParagraphStyle, ParagraphBackground,
    ParagraphBorders, ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphDropCap,
    ParagraphDropCapPlacement, ParagraphFlow, ParagraphFollowingStyle, ParagraphIndents,
    ParagraphLineSpacing, ParagraphList, ParagraphListBullet, ParagraphListBulletGeometry,
    ParagraphListIndentation, ParagraphListLabelColor, ParagraphListLevel,
    ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberScale,
    ParagraphListNumberTiering, ParagraphListNumbering, ParagraphSpacing, ParagraphStart,
    ParagraphTabStops, ParagraphWritingDirection, TextAlignment, TextBackground, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextColumns, TextComment, TextCommentBody,
    TextCommentId, TextCommentReply, TextCommentReplyBody, TextCommentReplyId, TextDecorations,
    TextFont, TextHighlight, TextHighlightId, TextHyperlink, TextHyperlinkId, TextHyperlinkTarget,
    TextLanguage, TextLanguageRun, TextLigatures, TextOutline, TextPosition, TextRange, TextScript,
    TextShadow, TextStorageInfo, TextStyle,
};
use crate::wire::{
    append_repeated_length_delimited_field, patch_fixed32_field, patch_length_delimited_field,
    patch_nested_fixed32_field, patch_nested_varint_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use crate::{EmbeddedMediaAsset, Error, IWorkMediaEditor, IWorkPackage, Result};

const DOCUMENT_ARCHIVE_NAME: &str = "Index/Document.iwa";
const DOCUMENT_OBJECT_ID: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 10000;
const SECTION_MESSAGE_TYPE: u32 = 10011;
const SECTION_TEMPLATE_MESSAGE_TYPE: u32 = 10143;
const USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE: u32 = 10016;
const GUIDE_STORAGE_MESSAGE_TYPE: u32 = 3047;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2001, 2022];
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2011;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const DRAWABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2003;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3097;
const BODY_DRAWABLE_DUPLICATE_OFFSET: f32 = 12.0;

pub use document_options::PagesDocumentOptions;
pub use types::{
    PagesDrawableTextInfo, PagesFootnote, PagesFootnoteFormat, PagesFootnoteGap, PagesFootnoteId,
    PagesFootnoteKind, PagesFootnoteNumbering, PagesFootnoteSettings, PagesHeaderFooterInfo,
    PagesHeaderFooterKind, PagesPageLayout, PagesPageNumber, PagesPageOrientation,
    PagesRgbColorSpace, PagesRgbaColor, PagesSectionBackground, PagesSectionInfo,
    PagesSectionPageNumbering, PagesSectionSettings, PagesSectionStart, PagesTemplateKind,
    RemovedPagesTextBox,
};

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

/// Transactional editor for a Pages package.
#[derive(Debug, Clone)]
pub struct PagesEditor {
    text: IWorkTextEditor,
    body_storage_id: u64,
    sections: Vec<PagesSectionInfo>,
    header_footers: Vec<HeaderFooterLocation>,
}

impl PagesEditor {
    /// Start configuring a new Pages package built from typed IWA objects.
    pub fn builder() -> crate::pages::PagesDocumentBuilder {
        crate::pages::PagesDocumentBuilder::new()
    }

    /// Create a blank, independent Pages document without a bundled template.
    pub fn create() -> Result<Self> {
        Self::builder().build()
    }

    /// Create an independent Pages document with initial body text.
    pub fn create_with_text(body_text: impl Into<String>) -> Result<Self> {
        Self::builder().body_text(body_text).build()
    }

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
            .filter(|drawable| reachable.contains(&drawable.object_id.object_id()))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.object_id.object_id());
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
            if let Some(previous_drawable) =
                storage_owners.insert(storage_id, drawable.object_id.object_id())
            {
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
                drawable_object_id: drawable.object_id.object_id(),
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

    /// Read vertical alignment, edge insets, and autosizing for a text box.
    pub fn text_box_text_layout(&self, drawable_object_id: u64) -> Result<ShapeTextLayout> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        shape_text_layout(self.package(), &archive_name, drawable_object_id)
    }

    /// Replace text-frame layout while preserving text, columns, and drawing style.
    pub fn set_text_box_text_layout(
        &mut self,
        drawable_object_id: u64,
        layout: ShapeTextLayout,
    ) -> Result<()> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package().clone(),
            &archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.text_box_text_layout(drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Pages text-box layout update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove crate-authored text-frame layout overrides.
    pub fn reset_text_box_text_layout(&mut self, drawable_object_id: u64) -> Result<bool> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        let (staged, changed) =
            reset_shape_text_layout(self.package().clone(), &archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the uniform column layout of a reachable ordinary text box.
    pub fn text_box_columns(&self, drawable_object_id: u64) -> Result<TextColumns> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        shape_text_columns(self.package(), &archive_name, drawable_object_id)
    }

    /// Replace the uniform column layout of a reachable ordinary text box.
    pub fn set_text_box_columns(
        &mut self,
        drawable_object_id: u64,
        columns: &TextColumns,
    ) -> Result<()> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        let staged = set_shape_text_columns(
            self.package().clone(),
            &archive_name,
            drawable_object_id,
            columns,
        )?;
        let verified = Self::from_package(staged)?;
        if &verified.text_box_columns(drawable_object_id)? != columns {
            return Err(Error::InvalidFormat(
                "Pages text-box column update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited column layout after a crate-authored override.
    pub fn reset_text_box_columns(&mut self, drawable_object_id: u64) -> Result<bool> {
        self.text_box_graph(drawable_object_id)?;
        let archive_name = find_object_archive(self.package(), drawable_object_id)?;
        let (staged, changed) =
            reset_shape_text_columns(self.package().clone(), &archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn text_box_text_style(&self, drawable_object_id: u64) -> Result<TextStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_style(graph.storage_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    pub fn set_text_box_text_style(
        &mut self,
        drawable_object_id: u64,
        style: TextStyle,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_style(graph.storage_id, style)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_style(drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Pages text-box character formatting update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited character formatting while preserving paragraph overrides.
    pub fn reset_text_box_text_style(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_style(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of an ordinary text box.
    pub fn text_box_text_font(&self, drawable_object_id: u64) -> Result<TextFont> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_font(graph.storage_id)
    }

    /// Atomically set a typed font identity across an ordinary text box.
    pub fn set_text_box_text_font(
        &mut self,
        drawable_object_id: u64,
        font: TextFont,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_font(graph.storage_id, font)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the inherited font while preserving sibling overrides.
    pub fn reset_text_box_text_font(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_font(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every explicit language boundary in an ordinary text box.
    pub fn text_box_text_languages(&self, drawable_object_id: u64) -> Result<Vec<TextLanguageRun>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_languages(graph.storage_id)
    }

    /// Read the effective language at one UTF-16 text boundary.
    pub fn text_box_text_language(
        &self,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<TextLanguage> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_language(graph.storage_id, position)
    }

    /// Atomically create or update one text-language boundary.
    pub fn set_text_box_text_language(
        &mut self,
        drawable_object_id: u64,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_language(graph.storage_id, position, language)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Delete one nonzero language boundary so it inherits the preceding run.
    pub fn remove_text_box_text_language_boundary(
        &mut self,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.remove_text_language_boundary(graph.storage_id, position)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Restore automatic language selection across an ordinary text box.
    pub fn reset_text_box_text_languages(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_languages(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every hyperlink in an ordinary text box.
    pub fn text_box_hyperlinks(&self, drawable_object_id: u64) -> Result<Vec<TextHyperlink>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_hyperlinks(graph.storage_id)
    }

    /// Create a hyperlink over a nonempty, unoccupied UTF-16 text range.
    pub fn add_text_box_hyperlink(
        &mut self,
        drawable_object_id: u64,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.add_text_hyperlink(graph.storage_id, range, target)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Update a text-box hyperlink's range and target without changing its ID.
    pub fn update_text_box_hyperlink(
        &mut self,
        drawable_object_id: u64,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.update_text_hyperlink(graph.storage_id, id, range, target)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Delete a text-box hyperlink and its owned smart-field object.
    pub fn remove_text_box_hyperlink(
        &mut self,
        drawable_object_id: u64,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let hyperlink = staged.remove_text_hyperlink(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(hyperlink)
    }

    /// Read every plain highlight in an ordinary text box.
    pub fn text_box_highlights(&self, drawable_object_id: u64) -> Result<Vec<TextHighlight>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_highlights(graph.storage_id)
    }

    /// Create a plain highlight over a nonempty UTF-16 text range.
    pub fn add_text_box_highlight(
        &mut self,
        drawable_object_id: u64,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.add_text_highlight(graph.storage_id, range)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Move a plain text-box highlight without changing its ID.
    pub fn update_text_box_highlight(
        &mut self,
        drawable_object_id: u64,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.update_text_highlight(graph.storage_id, id, range)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Delete a plain text-box highlight and its empty annotation graph.
    pub fn remove_text_box_highlight(
        &mut self,
        drawable_object_id: u64,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let highlight = staged.remove_text_highlight(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(highlight)
    }

    /// Read every ranged comment in an ordinary text box.
    pub fn text_box_comments(&self, drawable_object_id: u64) -> Result<Vec<TextComment>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_comments(graph.storage_id)
    }

    /// Create a ranged comment in an ordinary text box.
    pub fn add_text_box_comment(
        &mut self,
        drawable_object_id: u64,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.add_text_comment(graph.storage_id, range, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Update a text-box comment's range and body without changing its ID.
    pub fn update_text_box_comment(
        &mut self,
        drawable_object_id: u64,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.update_text_comment(graph.storage_id, id, range, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Delete a ranged text-box comment and its owned annotation graph.
    pub fn remove_text_box_comment(
        &mut self,
        drawable_object_id: u64,
        id: TextCommentId,
    ) -> Result<TextComment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let comment = staged.remove_text_comment(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(comment)
    }

    /// Read every direct reply to a text-box comment in stored order.
    pub fn text_box_comment_replies(
        &self,
        drawable_object_id: u64,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_comment_replies(graph.storage_id, comment_id)
    }

    /// Append a direct reply to a text-box comment.
    pub fn add_text_box_comment_reply(
        &mut self,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply = staged.add_text_comment_reply(graph.storage_id, comment_id, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Update a direct text-box comment reply without changing its ID.
    pub fn update_text_box_comment_reply(
        &mut self,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply =
            staged.update_text_comment_reply(graph.storage_id, comment_id, reply_id, body)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Delete one direct text-box comment reply and its storage.
    pub fn remove_text_box_comment_reply(
        &mut self,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let reply = staged.remove_text_comment_reply(graph.storage_id, comment_id, reply_id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(reply)
    }

    /// Read the canonical list preset of an ordinary text box.
    pub fn text_box_paragraph_list(&self, drawable_object_id: u64) -> Result<ParagraphList> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_list(graph.storage_id)
    }

    /// Atomically apply a canonical list preset to an ordinary text box.
    pub fn set_text_box_paragraph_list(
        &mut self,
        drawable_object_id: u64,
        list: ParagraphList,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list(graph.storage_id, list)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Remove list formatting from an ordinary text box.
    pub fn reset_text_box_paragraph_list(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read every list-level boundary in an ordinary text box.
    pub fn text_box_paragraph_list_levels(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_list_levels(graph.storage_id)
    }

    /// Read one paragraph's effective list nesting level.
    pub fn text_box_paragraph_list_level(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListLevel> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_list_level(graph.storage_id, paragraph)
    }

    /// Atomically set one paragraph's list nesting level.
    pub fn set_text_box_paragraph_list_level(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        level: ParagraphListLevel,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_level(graph.storage_id, paragraph, level)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_text_box_paragraph_list_level(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_level(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one text-box paragraph continues or restarts list numbering.
    pub fn text_box_paragraph_list_numbering(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListNumbering> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_numbering(graph.storage_id, paragraph)
    }

    /// Continue or restart numbered-list sequencing at one text-box paragraph.
    pub fn set_text_box_paragraph_list_numbering(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        numbering: ParagraphListNumbering,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_numbering(graph.storage_id, paragraph, numbering)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Read one numbered text-box paragraph's effective label format.
    pub fn text_box_paragraph_list_number_format(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListNumberFormat> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_number_format(graph.storage_id, paragraph)
    }

    /// Set one numbered text-box paragraph's locale-aware label format.
    pub fn set_text_box_paragraph_list_number_format(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        format: ParagraphListNumberFormat,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_format(graph.storage_id, paragraph, format)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_text_box_paragraph_list_number_format(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_format(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one numbered text-box paragraph displays hierarchical numbering.
    pub fn text_box_paragraph_list_number_tiering(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListNumberTiering> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_number_tiering(graph.storage_id, paragraph)
    }

    /// Choose flat or hierarchical numbering for one text-box list level.
    pub fn set_text_box_paragraph_list_number_tiering(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        tiering: ParagraphListNumberTiering,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_tiering(graph.storage_id, paragraph, tiering)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore flat numbering for one text-box list level.
    pub fn reset_text_box_paragraph_list_number_tiering(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_tiering(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one numbered text-box paragraph's number-label size.
    pub fn text_box_paragraph_list_number_scale(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListNumberScale> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_number_scale(graph.storage_id, paragraph)
    }

    /// Set one numbered text-box paragraph's number-label size.
    pub fn set_text_box_paragraph_list_number_scale(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        scale: ParagraphListNumberScale,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_number_scale(graph.storage_id, paragraph, scale)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_text_box_paragraph_list_number_scale(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_number_scale(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one text-box paragraph's effective text-bullet marker.
    pub fn text_box_paragraph_list_bullet(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListBullet> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_list_bullet(graph.storage_id, paragraph)
    }

    /// Set one text-box paragraph's text-bullet marker.
    pub fn set_text_box_paragraph_list_bullet(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        bullet: &ParagraphListBullet,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_bullet(graph.storage_id, paragraph, bullet)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one text-box paragraph.
    pub fn reset_text_box_paragraph_list_bullet(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_bullet(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one text-box paragraph's effective bullet size and baseline.
    pub fn text_box_paragraph_list_bullet_geometry(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListBulletGeometry> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_bullet_geometry(graph.storage_id, paragraph)
    }

    /// Set one text-box paragraph's bullet size and baseline.
    pub fn set_text_box_paragraph_list_bullet_geometry(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        geometry: ParagraphListBulletGeometry,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_bullet_geometry(graph.storage_id, paragraph, geometry)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_text_box_paragraph_list_bullet_geometry(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_bullet_geometry(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one text-box list paragraph's label and text-gap indentation.
    pub fn text_box_paragraph_list_indentation(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListIndentation> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_indentation(graph.storage_id, paragraph)
    }

    /// Set one text-box list paragraph's label and text-gap indentation.
    pub fn set_text_box_paragraph_list_indentation(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        indentation: ParagraphListIndentation,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_indentation(graph.storage_id, paragraph, indentation)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_text_box_paragraph_list_indentation(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_indentation(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read one text-box list paragraph's effective label color.
    pub fn text_box_paragraph_list_label_color(
        &self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListLabelColor> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_list_label_color(graph.storage_id, paragraph)
    }

    /// Set one text-box list paragraph's bullet or number color.
    pub fn set_text_box_paragraph_list_label_color(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
        color: ParagraphListLabelColor,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_list_label_color(graph.storage_id, paragraph, color)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_text_box_paragraph_list_label_color(
        &mut self,
        drawable_object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_list_label_color(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform underline and strikethrough formatting.
    pub fn text_box_text_decorations(&self, drawable_object_id: u64) -> Result<TextDecorations> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_decorations(graph.storage_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    pub fn set_text_box_text_decorations(
        &mut self,
        drawable_object_id: u64,
        decorations: TextDecorations,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_decorations(graph.storage_id, decorations)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_decorations(drawable_object_id)? != decorations {
            return Err(Error::InvalidFormat(
                "Pages text-box decoration update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited decorations while preserving sibling overrides.
    pub fn reset_text_box_text_decorations(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_decorations(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective uniform text color of a reachable ordinary text box.
    pub fn text_box_text_color(&self, drawable_object_id: u64) -> Result<RgbaColor> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_color(graph.storage_id)
    }

    /// Atomically set one text color across a reachable ordinary text box.
    pub fn set_text_box_text_color(
        &mut self,
        drawable_object_id: u64,
        color: RgbaColor,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_color(graph.storage_id, color)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_color(drawable_object_id)? != color {
            return Err(Error::InvalidFormat(
                "Pages text-box color update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited text color while preserving sibling overrides.
    pub fn reset_text_box_text_color(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_color(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform capitalization from a reachable ordinary text box.
    pub fn text_box_text_capitalization(
        &self,
        drawable_object_id: u64,
    ) -> Result<TextCapitalization> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_capitalization(graph.storage_id)
    }

    /// Atomically set one capitalization mode across a reachable ordinary text box.
    pub fn set_text_box_text_capitalization(
        &mut self,
        drawable_object_id: u64,
        capitalization: TextCapitalization,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_capitalization(graph.storage_id, capitalization)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_capitalization(drawable_object_id)? != capitalization {
            return Err(Error::InvalidFormat(
                "Pages text-box capitalization update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited capitalization while preserving sibling overrides.
    pub fn reset_text_box_text_capitalization(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_capitalization(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform baseline script from a reachable ordinary text box.
    pub fn text_box_text_script(&self, drawable_object_id: u64) -> Result<TextScript> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_script(graph.storage_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    pub fn set_text_box_text_script(
        &mut self,
        drawable_object_id: u64,
        script: TextScript,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_script(graph.storage_id, script)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_script(drawable_object_id)? != script {
            return Err(Error::InvalidFormat(
                "Pages text-box script update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited baseline script while preserving sibling overrides.
    pub fn reset_text_box_text_script(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_script(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of a reachable ordinary text box.
    pub fn text_box_text_baseline_shift(
        &self,
        drawable_object_id: u64,
    ) -> Result<TextBaselineShift> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_baseline_shift(graph.storage_id)
    }

    /// Atomically set a signed custom baseline displacement.
    pub fn set_text_box_text_baseline_shift(
        &mut self,
        drawable_object_id: u64,
        shift: TextBaselineShift,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_baseline_shift(graph.storage_id, shift)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_baseline_shift(drawable_object_id)? != shift {
            return Err(Error::InvalidFormat(
                "Pages text-box baseline-shift update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited baseline displacement while preserving sibling overrides.
    pub fn reset_text_box_text_baseline_shift(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_baseline_shift(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of a reachable ordinary text box.
    pub fn text_box_text_character_spacing(
        &self,
        drawable_object_id: u64,
    ) -> Result<TextCharacterSpacing> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_character_spacing(graph.storage_id)
    }

    /// Atomically set character spacing across a reachable ordinary text box.
    pub fn set_text_box_text_character_spacing(
        &mut self,
        drawable_object_id: u64,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_character_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_character_spacing(drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages text-box character-spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited character spacing while preserving sibling overrides.
    pub fn reset_text_box_text_character_spacing(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_character_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of a reachable ordinary text box.
    pub fn text_box_text_ligatures(&self, drawable_object_id: u64) -> Result<TextLigatures> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_ligatures(graph.storage_id)
    }

    /// Atomically set the ligature policy across a reachable ordinary text box.
    pub fn set_text_box_text_ligatures(
        &mut self,
        drawable_object_id: u64,
        ligatures: TextLigatures,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_ligatures(graph.storage_id, ligatures)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_ligatures(drawable_object_id)? != ligatures {
            return Err(Error::InvalidFormat(
                "Pages text-box ligature update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited ligatures while preserving sibling overrides.
    pub fn reset_text_box_text_ligatures(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_ligatures(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective outline of a reachable ordinary text box.
    pub fn text_box_text_outline(&self, drawable_object_id: u64) -> Result<TextOutline> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_outline(graph.storage_id)
    }

    /// Atomically set a typed outline across a reachable ordinary text box.
    pub fn set_text_box_text_outline(
        &mut self,
        drawable_object_id: u64,
        outline: TextOutline,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_outline(graph.storage_id, outline)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_outline(drawable_object_id)? != outline {
            return Err(Error::InvalidFormat(
                "Pages text-box outline update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited outline while preserving sibling overrides.
    pub fn reset_text_box_text_outline(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_outline(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective shadow of a reachable ordinary text box.
    pub fn text_box_text_shadow(&self, drawable_object_id: u64) -> Result<TextShadow> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_shadow(graph.storage_id)
    }

    /// Atomically set a typed drop shadow across a reachable ordinary text box.
    pub fn set_text_box_text_shadow(
        &mut self,
        drawable_object_id: u64,
        shadow: TextShadow,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_shadow(graph.storage_id, shadow)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_shadow(drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Pages text-box shadow update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited shadow while preserving sibling overrides.
    pub fn reset_text_box_text_shadow(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_shadow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective solid background of a reachable ordinary text box.
    pub fn text_box_text_background(&self, drawable_object_id: u64) -> Result<TextBackground> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_background(graph.storage_id)
    }

    /// Atomically set a solid background across a reachable ordinary text box.
    pub fn set_text_box_text_background(
        &mut self,
        drawable_object_id: u64,
        background: TextBackground,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_text_background(graph.storage_id, background)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_text_background(drawable_object_id)? != background {
            return Err(Error::InvalidFormat(
                "Pages text-box background update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited text background while preserving sibling overrides.
    pub fn reset_text_box_text_background(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_text_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph background.
    pub fn text_box_paragraph_background(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphBackground> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_background(graph.storage_id)
    }

    /// Atomically set the paragraph background across an ordinary text box.
    pub fn set_text_box_paragraph_background(
        &mut self,
        drawable_object_id: u64,
        background: ParagraphBackground,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_background(graph.storage_id, background)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_background(drawable_object_id)? != background {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph background update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph background.
    pub fn reset_text_box_paragraph_background(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph borders.
    pub fn text_box_paragraph_borders(&self, drawable_object_id: u64) -> Result<ParagraphBorders> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_borders(graph.storage_id)
    }

    /// Atomically set paragraph borders across an ordinary text box.
    pub fn set_text_box_paragraph_borders(
        &mut self,
        drawable_object_id: u64,
        borders: ParagraphBorders,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_borders(graph.storage_id, borders)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_borders(drawable_object_id)? != borders {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph border update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph borders.
    pub fn reset_text_box_paragraph_borders(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_borders(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph pagination and hyphenation controls.
    pub fn text_box_paragraph_flow(&self, drawable_object_id: u64) -> Result<ParagraphFlow> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_flow(graph.storage_id)
    }

    /// Atomically set paragraph pagination and hyphenation controls.
    pub fn set_text_box_paragraph_flow(
        &mut self,
        drawable_object_id: u64,
        flow: ParagraphFlow,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_flow(graph.storage_id, flow)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_flow(drawable_object_id)? != flow {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph flow update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited paragraph pagination and hyphenation controls.
    pub fn reset_text_box_paragraph_flow(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_flow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective base-writing direction of a reachable ordinary text box.
    pub fn text_box_paragraph_writing_direction(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphWritingDirection> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_writing_direction(graph.storage_id)
    }

    /// Set one base-writing direction across a reachable ordinary text box.
    pub fn set_text_box_paragraph_writing_direction(
        &mut self,
        drawable_object_id: u64,
        direction: ParagraphWritingDirection,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_writing_direction(graph.storage_id, direction)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_writing_direction(drawable_object_id)? != direction {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph writing-direction update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited base-writing direction.
    pub fn reset_text_box_paragraph_writing_direction(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_writing_direction(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// List the theme paragraph-style presets selectable for an ordinary text box.
    pub fn text_box_named_paragraph_styles(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<NamedParagraphStyle>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.named_paragraph_styles(graph.storage_id)
    }

    /// Read the named paragraph style selected for an ordinary text box.
    pub fn text_box_applied_named_paragraph_style(
        &self,
        drawable_object_id: u64,
    ) -> Result<AppliedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.applied_named_paragraph_style(graph.storage_id)
    }

    /// Redefine the selected named style from this text box's direct overrides.
    pub fn redefine_applied_text_box_named_paragraph_style(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let redefined = staged.redefine_applied_named_paragraph_style(graph.storage_id)?;
        let verified = Self::from_package(staged.package().clone())?;
        let selection = verified.text_box_applied_named_paragraph_style(drawable_object_id)?;
        if selection.style() != &redefined || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "Pages named paragraph-style redefinition failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(redefined)
    }

    /// Apply one named paragraph style and clear direct paragraph overrides.
    pub fn apply_text_box_named_paragraph_style(
        &mut self,
        drawable_object_id: u64,
        target: crate::text::ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let applied = staged.apply_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(staged.package().clone())?;
        let selection = verified.text_box_applied_named_paragraph_style(drawable_object_id)?;
        if selection.style() != &applied || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "Pages named paragraph-style application failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(applied)
    }

    /// Clone one selectable preset as a new named style in this text box's theme.
    pub fn create_text_box_named_paragraph_style(
        &mut self,
        drawable_object_id: u64,
        source: crate::text::ParagraphStyleId,
        name: crate::text::ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let created = staged.create_named_paragraph_style(graph.storage_id, source, name)?;
        let verified = Self::from_package(staged.package().clone())?;
        if !verified
            .text_box_named_paragraph_styles(drawable_object_id)?
            .contains(&created)
        {
            return Err(Error::InvalidFormat(
                "Pages named paragraph-style creation failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(created)
    }

    /// Rename one selectable paragraph style without changing its identifier.
    pub fn rename_text_box_named_paragraph_style(
        &mut self,
        drawable_object_id: u64,
        target: crate::text::ParagraphStyleId,
        name: crate::text::ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let renamed = staged.rename_named_paragraph_style(graph.storage_id, target, name)?;
        let verified = Self::from_package(staged.package().clone())?;
        if !verified
            .text_box_named_paragraph_styles(drawable_object_id)?
            .contains(&renamed)
        {
            return Err(Error::InvalidFormat(
                "Pages named paragraph-style rename failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(renamed)
    }

    /// Delete one unused selectable paragraph style.
    pub fn delete_text_box_named_paragraph_style(
        &mut self,
        drawable_object_id: u64,
        target: crate::text::ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let deleted = staged.delete_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified
            .text_box_named_paragraph_styles(drawable_object_id)?
            .iter()
            .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Pages named paragraph-style deletion failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(deleted)
    }

    /// Replace the style applied to this text box and delete it atomically.
    ///
    /// References from another object reject the deletion without changing the
    /// document, allowing callers to replace each remaining use explicitly.
    pub fn delete_applied_text_box_named_paragraph_style_with_replacement(
        &mut self,
        drawable_object_id: u64,
        target: crate::text::ParagraphStyleId,
        replacement: crate::text::ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let deleted = staged.delete_applied_named_paragraph_style_with_replacement(
            graph.storage_id,
            target,
            replacement,
        )?;
        let verified = Self::from_package(staged.package().clone())?;
        let selection = verified.text_box_applied_named_paragraph_style(drawable_object_id)?;
        if selection.style().id() != replacement
            || selection.has_overrides()
            || verified
                .text_box_named_paragraph_styles(drawable_object_id)?
                .iter()
                .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Pages paragraph-style replacement deletion failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(deleted)
    }

    /// Read the style selected for the paragraph created after this text box's current paragraph.
    pub fn text_box_paragraph_following_style(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphFollowingStyle> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_following_style(graph.storage_id)
    }

    /// Select the named or same style used for the following paragraph.
    pub fn set_text_box_paragraph_following_style(
        &mut self,
        drawable_object_id: u64,
        following_style: ParagraphFollowingStyle,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_following_style(graph.storage_id, following_style)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_following_style(drawable_object_id)? != following_style {
            return Err(Error::InvalidFormat(
                "Pages following paragraph-style update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited following paragraph style.
    pub fn reset_text_box_paragraph_following_style(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_following_style(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph alignment of a reachable ordinary text box.
    pub fn text_box_paragraph_alignment(&self, drawable_object_id: u64) -> Result<TextAlignment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_alignment(graph.storage_id)
    }

    /// Set one paragraph alignment across a reachable ordinary text box.
    pub fn set_text_box_paragraph_alignment(
        &mut self,
        drawable_object_id: u64,
        alignment: TextAlignment,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_alignment(graph.storage_id, alignment)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_alignment(drawable_object_id)? != alignment {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph-alignment update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited paragraph alignment after a private minimal override.
    pub fn reset_text_box_paragraph_alignment(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_alignment(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective line spacing of a reachable ordinary text box.
    pub fn text_box_paragraph_line_spacing(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphLineSpacing> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_line_spacing(graph.storage_id)
    }

    /// Set one typed line-spacing mode across a reachable ordinary text box.
    pub fn set_text_box_paragraph_line_spacing(
        &mut self,
        drawable_object_id: u64,
        spacing: ParagraphLineSpacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_line_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_line_spacing(drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph line-spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited line spacing while preserving sibling paragraph overrides.
    pub fn reset_text_box_paragraph_line_spacing(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_line_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing of an ordinary text box.
    pub fn text_box_paragraph_spacing(&self, drawable_object_id: u64) -> Result<ParagraphSpacing> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_spacing(graph.storage_id)
    }

    /// Atomically set before/after paragraph spacing across an ordinary text box.
    pub fn set_text_box_paragraph_spacing(
        &mut self,
        drawable_object_id: u64,
        spacing: ParagraphSpacing,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_spacing(drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph spacing update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited paragraph spacing while preserving sibling overrides.
    pub fn reset_text_box_paragraph_spacing(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right indentation of an ordinary text box.
    pub fn text_box_paragraph_indents(&self, drawable_object_id: u64) -> Result<ParagraphIndents> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_indents(graph.storage_id)
    }

    /// Atomically set paragraph indentation across an ordinary text box.
    pub fn set_text_box_paragraph_indents(
        &mut self,
        drawable_object_id: u64,
        indents: ParagraphIndents,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_indents(graph.storage_id, indents)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_indents(drawable_object_id)? != indents {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph indentation update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited indentation while preserving sibling paragraph overrides.
    pub fn reset_text_box_paragraph_indents(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_indents(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the character used to align decimal tab stops in an ordinary text box.
    pub fn text_box_paragraph_decimal_tab_character(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphDecimalTabCharacter> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_decimal_tab_character(graph.storage_id)
    }

    /// Atomically set the character used to align decimal tab stops.
    pub fn set_text_box_paragraph_decimal_tab_character(
        &mut self,
        drawable_object_id: u64,
        character: ParagraphDecimalTabCharacter,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_decimal_tab_character(graph.storage_id, character)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_decimal_tab_character(drawable_object_id)? != character {
            return Err(Error::InvalidFormat(
                "Pages text-box decimal-tab character update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited decimal-tab character.
    pub fn reset_text_box_paragraph_decimal_tab_character(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_decimal_tab_character(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the distance between implicit tab stops in an ordinary text box.
    pub fn text_box_paragraph_default_tab_interval(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphDefaultTabInterval> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_default_tab_interval(graph.storage_id)
    }

    /// Atomically set the distance between implicit tab stops.
    pub fn set_text_box_paragraph_default_tab_interval(
        &mut self,
        drawable_object_id: u64,
        interval: ParagraphDefaultTabInterval,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_default_tab_interval(graph.storage_id, interval)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_default_tab_interval(drawable_object_id)? != interval {
            return Err(Error::InvalidFormat(
                "Pages text-box default-tab interval update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore the inherited default-tab interval.
    pub fn reset_text_box_paragraph_default_tab_interval(
        &mut self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_default_tab_interval(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of an ordinary text box.
    pub fn text_box_paragraph_tab_stops(
        &self,
        drawable_object_id: u64,
    ) -> Result<ParagraphTabStops> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_tab_stops(graph.storage_id)
    }

    /// Atomically replace every explicit ruler tab stop of an ordinary text box.
    pub fn set_text_box_paragraph_tab_stops(
        &mut self,
        drawable_object_id: u64,
        stops: ParagraphTabStops,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_tab_stops(graph.storage_id, stops)?;
        let expected = staged.paragraph_tab_stops(graph.storage_id)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_tab_stops(drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Pages text-box paragraph tab-stop update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Restore inherited tab stops while preserving sibling paragraph overrides.
    pub fn reset_text_box_paragraph_tab_stops(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.reset_paragraph_tab_stops(graph.storage_id)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
    }

    /// List every Drop Cap in an ordinary text box.
    pub fn text_box_paragraph_drop_caps(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphDropCapPlacement>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.paragraph_drop_caps(graph.storage_id)
    }

    /// Read the Drop Cap attached to one text-box paragraph.
    pub fn text_box_paragraph_drop_cap(
        &self,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<Option<ParagraphDropCap>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text
            .paragraph_drop_cap(graph.storage_id, paragraph_start)
    }

    /// Atomically create or replace a text-box Drop Cap.
    pub fn set_text_box_paragraph_drop_cap(
        &mut self,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
        drop_cap: ParagraphDropCap,
    ) -> Result<()> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        staged.set_paragraph_drop_cap(graph.storage_id, paragraph_start, drop_cap)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified.text_box_paragraph_drop_cap(drawable_object_id, paragraph_start)?
            != Some(drop_cap)
        {
            return Err(Error::InvalidFormat(
                "Pages text-box Drop Cap update failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(())
    }

    /// Atomically remove a text-box Drop Cap.
    pub fn remove_text_box_paragraph_drop_cap(
        &mut self,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<bool> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let changed = staged.remove_paragraph_drop_cap(graph.storage_id, paragraph_start)?;
        if changed {
            *self = Self::from_package(staged.into_package())?;
        }
        Ok(changed)
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
                clone_pages_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&archive_name, |archive| Ok(archive.insert_object(cloned)?))?;
        }

        let new_drawable_id = remap[&source.drawable_id];
        let new_storage_id = remap[&source.storage_id];
        let new_attachment_id = remap[&source.attachment_id];
        offset_pages_body_drawable_clone(
            &mut staged,
            new_drawable_id,
            new_attachment_id,
            BODY_DRAWABLE_DUPLICATE_OFFSET,
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
        comments.clear_comment(DrawableObjectId::from_object_id(drawable_object_id)?)?;
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
            .any(|drawable| drawable.object_id.object_id() == drawable_object_id)
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
            .comment(DrawableObjectId::from_object_id(drawable_object_id)?)
    }

    /// Create or replace a direct comment on a reachable Pages drawable.
    pub fn set_drawable_comment(
        &mut self,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.set_comment(DrawableObjectId::from_object_id(drawable_object_id)?, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Delete a direct comment from a reachable Pages drawable.
    pub fn clear_drawable_comment(&mut self, drawable_object_id: u64) -> Result<()> {
        self.require_drawable(drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(DrawableObjectId::from_object_id(drawable_object_id)?)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Read the direct replies in a reachable Pages drawable comment thread.
    pub fn drawable_comment_replies(
        &self,
        drawable_object_id: DrawableObjectId,
    ) -> Result<Vec<DrawableCommentReplyInfo>> {
        self.require_drawable(drawable_object_id.object_id())?;
        IWorkDrawableCommentEditor::from_package(self.package().clone())?
            .replies(drawable_object_id)
    }

    /// Add a reply to a reachable Pages drawable comment.
    pub fn add_drawable_comment_reply(
        &mut self,
        drawable_object_id: DrawableObjectId,
        text: impl Into<String>,
    ) -> Result<CommentStorageId> {
        self.require_drawable(drawable_object_id.object_id())?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.add_reply(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Update a direct reply, returning its current storage identifier.
    pub fn set_drawable_comment_reply(
        &mut self,
        drawable_object_id: DrawableObjectId,
        reply_storage_object_id: CommentStorageId,
        text: impl Into<String>,
    ) -> Result<CommentStorageId> {
        self.require_drawable(drawable_object_id.object_id())?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        let reply_id = comments.set_reply(drawable_object_id, reply_storage_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id)
    }

    /// Remove a direct reply from a reachable Pages drawable comment.
    pub fn remove_drawable_comment_reply(
        &mut self,
        drawable_object_id: DrawableObjectId,
        reply_storage_object_id: CommentStorageId,
    ) -> Result<()> {
        self.require_drawable(drawable_object_id.object_id())?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.remove_reply(drawable_object_id, reply_storage_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// List body section boundaries in UTF-16 document order.
    pub fn sections(&self) -> &[PagesSectionInfo] {
        &self.sections
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
            staged.update_archive(&graph.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
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
            .any(|drawable| drawable.object_id.object_id() == drawable_object_id)
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
            let Some(attachment) = decode_optional_typed_package_object::<DrawableAttachmentArchive>(
                self.package(),
                reference.identifier,
                DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            )?
            else {
                continue;
            };
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

fn decode_optional_typed_package_object<T: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
) -> Result<Option<T>> {
    let mut found = false;
    let mut decoded = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        if found {
            return Err(Error::InvalidFormat(format!(
                "object {identifier} occurs in more than one Pages component"
            )));
        }
        found = true;
        decoded = object
            .messages
            .iter()
            .find(|message| message.type_ == message_type)
            .map(|message| T::decode(message.data.as_slice()))
            .transpose()?;
    }
    if !found {
        return Err(Error::InvalidFormat(format!(
            "Pages body attachment references missing object {identifier}"
        )));
    }
    Ok(decoded)
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

fn remap_pages_drawable_archive(drawable: &mut tsd::DrawableArchive, remap: &HashMap<u64, u64>) {
    remap_optional_pages_reference(&mut drawable.parent, remap);
    remap_optional_pages_reference(&mut drawable.comment, remap);
    for reference in &mut drawable.pencil_annotations {
        remap_pages_reference(reference, remap);
    }
    remap_optional_pages_reference(&mut drawable.title, remap);
    remap_optional_pages_reference(&mut drawable.caption, remap);
}

fn remap_pages_chart_wire(
    data: &[u8],
    recorded_references: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = crate::charts::IWorkChartArchive::decode(data)?;
    expected.remap_references(remap, recorded_references)?;
    let data = expected.encode()?;
    if crate::charts::IWorkChartArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_pages_image_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    let mut expected = tsd::ImageArchive::decode(data)?;
    remap_pages_drawable_archive(&mut expected.super_, remap);
    remap_optional_pages_reference(&mut expected.database_data, remap);
    remap_optional_pages_reference(&mut expected.style, remap);
    remap_optional_pages_reference(&mut expected.mask, remap);
    remap_optional_pages_reference(&mut expected.database_thumbnail_data, remap);
    remap_optional_pages_reference(&mut expected.database_original_data, remap);
    let data = remap_pages_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::ImageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages ImageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_pages_movie_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    let mut expected = tsd::MovieArchive::decode(data)?;
    remap_pages_drawable_archive(&mut expected.super_, remap);
    remap_optional_pages_reference(&mut expected.database_movie_data, remap);
    remap_optional_pages_reference(&mut expected.database_poster_image_data, remap);
    remap_optional_pages_reference(&mut expected.database_audio_only_image_data, remap);
    remap_optional_pages_reference(&mut expected.style, remap);
    let data = remap_pages_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::MovieArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages MovieArchive wire remap failed validation".to_owned(),
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

#[allow(deprecated)]
fn remap_pages_caption_info_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 1, 1, 2],
        &[1, 1, 1, 6],
        &[1, 1, 1, 9],
        &[1, 1, 1, 10],
        &[1, 1, 1, 11],
        &[1, 1, 2],
        &[1, 2],
        &[1, 3],
        &[1, 4],
        &[2],
    ];
    let mut expected = crate::protobuf::tsa::CaptionInfoArchive::decode(data)?;
    remap_pages_shape(&mut expected.super_, remap);
    remap_optional_pages_reference(&mut expected.placement, remap);
    let data = remap_pages_reference_paths(data, REFERENCE_PATHS, remap)?;
    if crate::protobuf::tsa::CaptionInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Pages CaptionInfoArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn clone_pages_drawable_graph_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let old_identifier = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Pages drawable object has no identifier".to_owned())
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
            crate::charts::source::CHART_MESSAGE_TYPE => {
                remap_pages_chart_wire(&message.data, &info.object_references, remap)?
            },
            crate::charts::source::CHART_PRESET_MESSAGE_TYPE => {
                crate::charts::source::remap_chart_preset_wire(
                    &message.data,
                    &info.object_references,
                    remap,
                )?
            },
            SHAPE_INFO_MESSAGE_TYPE => remap_pages_shape_wire(&message.data, remap)?,
            IMAGE_MESSAGE_TYPE => remap_pages_image_wire(&message.data, remap)?,
            MOVIE_MESSAGE_TYPE => remap_pages_movie_wire(&message.data, remap)?,
            2001 | 2022 => remap_pages_storage_wire(&message.data, remap)?,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE => remap_pages_attachment_wire(&message.data, remap)?,
            CAPTION_INFO_MESSAGE_TYPE => remap_pages_caption_info_wire(&message.data, remap)?,
            STANDIN_CAPTION_MESSAGE_TYPE => message.data.clone(),
            _ => {
                if info
                    .object_references
                    .iter()
                    .any(|identifier| remap.contains_key(identifier))
                {
                    return Err(Error::InvalidFormat(format!(
                        "Cannot safely clone Pages message type {} with private drawable-graph references",
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

fn offset_pages_body_drawable_clone(
    package: &mut IWorkPackage,
    drawable_id: u64,
    attachment_id: u64,
    offset: f32,
) -> Result<()> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Pages drawable duplicate offset must be finite".to_owned(),
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
            Error::InvalidFormat(format!("Pages drawable {drawable_id} is missing"))
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
                "Pages drawable {drawable_id} must have exactly one shape payload"
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
                    "Pages drawable {drawable_id} has no positioned geometry"
                ))
            })?;
        let x = position.x + offset;
        let y = position.y + offset;
        if !x.is_finite() || !y.is_finite() {
            return Err(Error::ParseError(
                "Pages drawable duplicate position overflow".to_owned(),
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
                Error::InvalidFormat("Pages drawable geometry patch removed position".to_owned())
            })?;
        if verified_position.x != x || verified_position.y != y {
            return Err(Error::InvalidFormat(
                "Pages drawable geometry offset failed validation".to_owned(),
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

    offset_pages_body_drawable_attachment_clone(package, attachment_id, offset)
}

pub(super) fn offset_pages_body_drawable_attachment_clone(
    package: &mut IWorkPackage,
    attachment_id: u64,
    offset: f32,
) -> Result<()> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Pages drawable duplicate offset must be finite".to_owned(),
        ));
    }
    let attachment_archive = find_object_archive(package, attachment_id)?;
    package.update_archive(&attachment_archive, |archive| {
        let object = archive.object_mut(attachment_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable attachment {attachment_id} is missing"
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
                "Pages drawable attachment {attachment_id} must have exactly one payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let attachment = DrawableAttachmentArchive::decode(original)?;
        let h_offset = attachment.h_offset.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable attachment {attachment_id} has no horizontal offset"
            ))
        })? + offset;
        let v_offset = attachment.v_offset.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable attachment {attachment_id} has no vertical offset"
            ))
        })? + offset;
        if !h_offset.is_finite() || !v_offset.is_finite() {
            return Err(Error::ParseError(
                "Pages drawable duplicate attachment offset overflow".to_owned(),
            ));
        }
        let data = patch_fixed32_field(original, 3, true, Some(h_offset.to_bits()))?;
        let data = patch_fixed32_field(&data, 5, true, Some(v_offset.to_bits()))?;
        let verified = DrawableAttachmentArchive::decode(data.as_slice())?;
        if verified.h_offset != Some(h_offset) || verified.v_offset != Some(v_offset) {
            return Err(Error::InvalidFormat(
                "Pages drawable attachment offset failed validation".to_owned(),
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
        let Some(object) = archive.object(drawable.object_id.object_id()) else {
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
    let archive = package.archive(DOCUMENT_ARCHIVE_NAME)?;
    let object = archive.object(DOCUMENT_OBJECT_ID).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages root object {DOCUMENT_OBJECT_ID} is missing"))
    })?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .and_then(|message| DocumentArchive::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat("Pages root has no TP.DocumentArchive payload".to_owned())
        })
}

fn discover_structure(
    package: &IWorkPackage,
    body_storage_id: u64,
) -> Result<(Vec<PagesSectionInfo>, Vec<HeaderFooterLocation>)> {
    let document_archive = package.archive(DOCUMENT_ARCHIVE_NAME)?;
    let document = document_archive
        .object(DOCUMENT_OBJECT_ID)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
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

mod audio;
mod body_shapes;
mod charts;
mod date_time_fields;
mod document_options;
mod drawable_order;
mod footnote_settings;
mod footnotes;
mod images;
mod movies;
mod number_attachments;
mod page_layout;
mod section_background;
mod section_content;
mod section_settings;
mod tables;
mod text_box_create;
mod types;

pub use audio::{PagesAudioInfo, PagesAudioOptions, RemovedPagesAudio};
pub use body_shapes::{PagesBodyShapeInfo, PagesBodyShapeKind, RemovedPagesBodyShape};
pub use charts::{PagesBodyChartInfo, RemovedPagesBodyChart};
pub use images::{PagesImageInfo, PagesImageOptions, RemovedPagesImage};
pub use movies::{PagesMovieInfo, PagesMovieOptions, RemovedPagesMovie};
pub use tables::{
    PagesCellValue, PagesTable, PagesTableAxisIndex, PagesTableCellCheckboxFormat,
    PagesTableCellComment, PagesTableCellCommentInfo, PagesTableCellCommentReplyInfo,
    PagesTableCellConditionalHighlightInfo, PagesTableCellCurrencyFormat, PagesTableCellDataFormat,
    PagesTableCellDateTimeFormat, PagesTableCellDecimalPlaces, PagesTableCellDurationFormat,
    PagesTableCellDurationStyle, PagesTableCellDurationUnit, PagesTableCellDurationUnitRange,
    PagesTableCellDurationUnits, PagesTableCellFixedDecimalPlaces, PagesTableCellFractionFormat,
    PagesTableCellInset, PagesTableCellInsets, PagesTableCellLayout,
    PagesTableCellNegativeNumberStyle, PagesTableCellNumberFormat,
    PagesTableCellNumeralSystemFormat, PagesTableCellParagraphIndents,
    PagesTableCellParagraphLineSpacing, PagesTableCellParagraphList,
    PagesTableCellParagraphListBullet, PagesTableCellParagraphListBulletGeometry,
    PagesTableCellParagraphListIndentation, PagesTableCellParagraphListLabelColor,
    PagesTableCellParagraphListLevel, PagesTableCellParagraphListLevelPlacement,
    PagesTableCellParagraphListNumberFormat, PagesTableCellParagraphListNumberScale,
    PagesTableCellParagraphListNumberTiering, PagesTableCellParagraphListNumbering,
    PagesTableCellParagraphListPlacement, PagesTableCellParagraphSpacing,
    PagesTableCellParagraphTabStops, PagesTableCellPercentageFormat, PagesTableCellPopUpMenuFormat,
    PagesTableCellPopUpMenuInitialSelection, PagesTableCellPopUpMenuItem, PagesTableCellRegion,
    PagesTableCellScientificFormat, PagesTableCellSliderDisplayFormat, PagesTableCellSliderFormat,
    PagesTableCellSliderRange, PagesTableCellStarRatingFormat, PagesTableCellStepperDisplayFormat,
    PagesTableCellStepperFormat, PagesTableCellStepperRange, PagesTableCellTextAlignment,
    PagesTableCellTextBackground, PagesTableCellTextBaselineShift,
    PagesTableCellTextCapitalization, PagesTableCellTextCharacterSpacing, PagesTableCellTextColor,
    PagesTableCellTextDecorations, PagesTableCellTextFont, PagesTableCellTextFormat,
    PagesTableCellTextLigatures, PagesTableCellTextOutline, PagesTableCellTextScript,
    PagesTableCellTextShadow, PagesTableCellTextStyle, PagesTableCellTextWrap,
    PagesTableCellThousandsSeparator, PagesTableCellUpdate, PagesTableCellVerticalAlignment,
    PagesTableColumnDeletion, PagesTableColumnInsertion, PagesTableDimension,
    PagesTableDimensionSize, PagesTableFormulaAxisReference, PagesTableFormulaBinaryOperator,
    PagesTableFormulaCachedValue, PagesTableFormulaCellReference, PagesTableFormulaExpression,
    PagesTableHeaderCount, PagesTableHeaderSettings, PagesTableHiddenAxes, PagesTableInfo,
    PagesTablePoints, PagesTableRowDeletion, PagesTableRowInsertion, PagesTableSortColumnIndex,
    PagesTableSortDirection, PagesTableSortOrder, PagesTableSortRowRange, PagesTableSortRule,
    PagesTableSortScope, PagesTableTitleSettings,
};

#[cfg(test)]
mod tests;

#[cfg(test)]
mod strict_selector_tests {
    use super::*;

    #[test]
    fn direct_comment_reply_api_uses_semantic_ids() {
        let drawable = DrawableObjectId::from_object_id(7).unwrap();
        let reply = CommentStorageId::from_object_id(9).unwrap();

        assert_eq!(drawable.object_id(), 7);
        assert_eq!(reply.object_id(), 9);
        assert_eq!(DrawableObjectId::new(0), None);
        assert_eq!(CommentStorageId::new(0), None);

        let _: fn(&PagesEditor, DrawableObjectId) -> Result<Vec<DrawableCommentReplyInfo>> =
            PagesEditor::drawable_comment_replies;
        let _: fn(&mut PagesEditor, DrawableObjectId, String) -> Result<CommentStorageId> =
            PagesEditor::add_drawable_comment_reply;
        let _: fn(
            &mut PagesEditor,
            DrawableObjectId,
            CommentStorageId,
            String,
        ) -> Result<CommentStorageId> = PagesEditor::set_drawable_comment_reply;
        let _: fn(&mut PagesEditor, DrawableObjectId, CommentStorageId) -> Result<()> =
            PagesEditor::remove_drawable_comment_reply;
    }
}
