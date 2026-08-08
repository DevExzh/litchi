//! Transactional editing of shared iWork text storage objects.

use std::collections::HashSet;
use std::ops::Range;
use std::path::Path;

use crate::archive::RawMessage;
use crate::protobuf::tswp::StorageArchive;
use crate::shapes::RgbaColor;
use crate::{Error, IWorkPackage, Result};

use super::TextStorageId;
use super::bookmark::{
    add_text_bookmark, remove_text_bookmark, remove_unreferenced_bookmark_objects, text_bookmarks,
    update_text_bookmark,
};
use super::bookmark_types::{TextBookmark, TextBookmarkId, TextBookmarkSettings};
use super::date_time::{
    add_text_date_time_field, insert_text_date_time_field, remove_text_date_time_field,
    remove_unreferenced_date_time_objects, text_date_time_fields, update_text_date_time_field,
};
use super::date_time_field::{TextDateTimeField, TextDateTimeFieldId};
use super::drop_cap::{
    paragraph_drop_cap, paragraph_drop_caps, remove_paragraph_drop_cap, set_paragraph_drop_cap,
};
use super::font::TextFont;
use super::highlight::{
    add_text_highlight, remove_text_highlight, remove_unreferenced_highlight_objects,
    text_highlights, update_text_highlight,
};
use super::hyperlink::{
    add_text_hyperlink, remove_text_hyperlink, remove_unreferenced_hyperlink_objects,
    text_hyperlinks, update_text_hyperlink,
};
use super::language::{
    remove_text_language_boundary, reset_text_languages, set_text_language, text_language,
    text_languages,
};
use super::number_attachment::{
    insert_text_number_attachment, remove_text_number_attachment,
    remove_unreferenced_number_attachment_objects, text_number_attachments,
    update_text_number_attachment,
};
use super::paragraph_alignment::{
    applied_named_paragraph_style, apply_named_paragraph_style, create_named_paragraph_style,
    delete_named_paragraph_style, named_paragraph_styles, paragraph_alignment,
    paragraph_background, paragraph_borders, paragraph_decimal_tab_character,
    paragraph_default_tab_interval, paragraph_flow, paragraph_following_style, paragraph_indents,
    paragraph_line_spacing, paragraph_spacing, paragraph_tab_stops, paragraph_writing_direction,
    redefine_applied_named_paragraph_style, rename_named_paragraph_style,
    reset_paragraph_alignment, reset_paragraph_background, reset_paragraph_borders,
    reset_paragraph_decimal_tab_character, reset_paragraph_default_tab_interval,
    reset_paragraph_flow, reset_paragraph_following_style, reset_paragraph_indents,
    reset_paragraph_line_spacing, reset_paragraph_spacing, reset_paragraph_tab_stops,
    reset_paragraph_writing_direction, reset_text_background, reset_text_baseline_shift,
    reset_text_capitalization, reset_text_character_spacing, reset_text_color,
    reset_text_decorations, reset_text_font, reset_text_ligatures, reset_text_outline,
    reset_text_script, reset_text_shadow, reset_text_style, set_paragraph_alignment,
    set_paragraph_background, set_paragraph_borders, set_paragraph_decimal_tab_character,
    set_paragraph_default_tab_interval, set_paragraph_flow, set_paragraph_following_style,
    set_paragraph_indents, set_paragraph_line_spacing, set_paragraph_spacing,
    set_paragraph_tab_stops, set_paragraph_writing_direction, set_text_background,
    set_text_baseline_shift, set_text_capitalization, set_text_character_spacing, set_text_color,
    set_text_decorations, set_text_font, set_text_ligatures, set_text_outline, set_text_script,
    set_text_shadow, set_text_style, text_background, text_baseline_shift, text_capitalization,
    text_character_spacing, text_color, text_decorations, text_font, text_ligatures, text_outline,
    text_script, text_shadow, text_style,
};
use super::paragraph_direction::ParagraphWritingDirection;
use super::paragraph_flow::ParagraphFlow;
use super::paragraph_list::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletGeometry, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLevel, ParagraphListLevelPlacement,
    ParagraphListNumberFormat, ParagraphListNumberScale, ParagraphListNumberTiering,
    ParagraphListNumbering, ParagraphListPlacement, paragraph_list, paragraph_list_bullet,
    paragraph_list_bullet_geometry, paragraph_list_indentation, paragraph_list_label_color,
    paragraph_list_level, paragraph_list_levels, paragraph_list_number_format,
    paragraph_list_number_scale, paragraph_list_number_tiering, paragraph_list_numbering,
    paragraph_lists, reset_paragraph_list, reset_paragraph_list_bullet,
    reset_paragraph_list_bullet_geometry, reset_paragraph_list_indentation,
    reset_paragraph_list_label_color, reset_paragraph_list_level,
    reset_paragraph_list_number_format, reset_paragraph_list_number_scale,
    reset_paragraph_list_number_tiering, set_paragraph_list, set_paragraph_list_bullet,
    set_paragraph_list_bullet_geometry, set_paragraph_list_indentation,
    set_paragraph_list_label_color, set_paragraph_list_level, set_paragraph_list_number_format,
    set_paragraph_list_number_scale, set_paragraph_list_number_tiering,
    set_paragraph_list_numbering, set_paragraph_lists,
};
use super::paragraph_style_apply::AppliedParagraphStyle;
use super::paragraph_tabs::{
    ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphTabStops,
};
use super::storage_wire::{
    locate_text_storage, locate_text_storage_with_archive, locate_text_storages,
    update_parsed_archive,
};
use super::text_comment::{
    add_text_comment, add_text_comment_reply, remove_text_comment, remove_text_comment_reply,
    text_comment_replies, text_comments, update_text_comment, update_text_comment_reply,
};
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::{
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextScript, TextStyle,
};
use litchi_iwa_text::comment::{
    TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
};
use litchi_iwa_text::date_time::{DisplayText, Settings};
use litchi_iwa_text::highlight::{TextHighlight, TextHighlightId};
use litchi_iwa_text::hyperlink::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
use litchi_iwa_text::number_attachment::{
    TextNumberAttachment, TextNumberAttachmentId, TextNumberAttachmentSettings,
};
use litchi_iwa_text::paragraph::drop_cap::{DropCap, Placement};
use litchi_iwa_text::paragraph::format::{Alignment, Borders, Indents, LineSpacing, Spacing};
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphFollowingStyle, ParagraphStyleId, ParagraphStyleName,
    raw::native_id,
};
use litchi_iwa_text::position::{TextPosition, TextRange};
use litchi_iwa_text::storage::Storage;
use litchi_iwa_text::{TextLanguage, TextLanguageRun};

/// A discoverable text storage within an iWork package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStorageInfo {
    /// Stable identity for this writable text storage.
    pub id: TextStorageId,
    /// Native message type retained inside the IWA adapter for wire writes.
    pub(crate) message_type: u32,
    /// Native storage kind retained inside the IWA adapter for format-specific
    /// structural decoding.
    pub(crate) kind: Option<i32>,
    pub storage: Storage,
}

/// Mutable editor for the TSWP text layer shared by Pages, Numbers, and Keynote.
#[derive(Debug, Clone)]
pub struct IWorkTextEditor {
    package: IWorkPackage,
}

impl IWorkTextEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Ok(Self {
            package: IWorkPackage::open(path)?,
        })
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            package: IWorkPackage::from_bytes(bytes)?,
        })
    }

    pub(crate) fn from_package(package: IWorkPackage) -> Self {
        Self { package }
    }

    pub fn storages(&self) -> Result<Vec<TextStorageInfo>> {
        let mut storages = locate_text_storages(&self.package)?
            .into_iter()
            .map(|location| {
                let id = TextStorageId::try_from(location.object_id)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                Ok(TextStorageInfo {
                    id,
                    message_type: location.message_type,
                    kind: location.storage.kind,
                    storage: Storage::from_text(location.storage.text.concat()),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        storages.sort_by_key(|storage| storage.id);
        Ok(storages)
    }

    pub fn storage(&self, id: TextStorageId) -> Result<TextStorageInfo> {
        let location = locate_text_storage(&self.package, id.get())?;
        Ok(TextStorageInfo {
            id,
            message_type: location.message_type,
            kind: location.storage.kind,
            storage: Storage::from_text(location.storage.text.concat()),
        })
    }

    /// Replace a UTF-16 range, matching the indexing used by iWork attributes.
    pub fn replace_text(
        &mut self,
        object_id: TextStorageId,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let object_id = object_id.get();
        if range.start > range.end {
            return Err(Error::ParseError(
                "Text replacement range starts after it ends".to_string(),
            ));
        }
        let mut staged = self.package.clone();
        replace_storage_text(&mut staged, object_id, range, replacement)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    pub fn set_text(&mut self, object_id: TextStorageId, replacement: &str) -> Result<()> {
        let storage = self.storage(object_id)?;
        self.replace_text(
            object_id,
            0..storage.storage.text().encode_utf16().count(),
            replacement,
        )
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn text_style(&self, object_id: TextStorageId) -> Result<TextStyle> {
        let object_id = object_id.get();
        text_style(&self.package, object_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_style(&mut self, object_id: TextStorageId, style: TextStyle) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_style(&mut staged, object_id, style)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_style(&verified, object_id)? != style {
            return Err(Error::InvalidFormat(
                "iWork text-style update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited character formatting while preserving paragraph overrides.
    pub fn reset_text_style(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_style(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of uniformly styled text.
    pub fn text_font(&self, object_id: TextStorageId) -> Result<TextFont> {
        let object_id = object_id.get();
        text_font(&self.package, object_id)
    }

    /// Atomically set a typed font identity across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_font(&mut self, object_id: TextStorageId, font: TextFont) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_font(&mut staged, object_id, font.clone())?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_font(&verified, object_id)? != font {
            return Err(Error::InvalidFormat(
                "iWork text font update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited font while preserving sibling overrides.
    pub fn reset_text_font(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_font(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read every explicit native language boundary in a text storage.
    pub fn text_languages(&self, object_id: TextStorageId) -> Result<Vec<TextLanguageRun>> {
        let object_id = object_id.get();
        text_languages(&self.package, object_id)
    }

    /// Read the effective language at one validated UTF-16 text boundary.
    pub fn text_language(
        &self,
        object_id: TextStorageId,
        position: TextPosition,
    ) -> Result<TextLanguage> {
        let object_id = object_id.get();
        text_language(&self.package, object_id, position)
    }

    /// Atomically create or update a language boundary.
    pub fn set_text_language(
        &mut self,
        object_id: TextStorageId,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_language(&mut staged, object_id, position, &language)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_language(&verified, object_id, position)? != language {
            return Err(Error::InvalidFormat(
                "iWork text-language update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Delete one nonzero language boundary so it inherits the preceding run.
    pub fn remove_text_language_boundary(
        &mut self,
        object_id: TextStorageId,
        position: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = remove_text_language_boundary(&mut staged, object_id, position)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if text_languages(&verified, object_id)?
                .iter()
                .any(|run| run.position == position)
            {
                return Err(Error::InvalidFormat(
                    "iWork text-language boundary removal failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Remove all explicit language boundaries and restore automatic selection.
    pub fn reset_text_languages(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_languages(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if !text_languages(&verified, object_id)?.is_empty() {
                return Err(Error::InvalidFormat(
                    "iWork text-language reset failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read every native hyperlink in a text storage.
    pub fn text_hyperlinks(&self, object_id: TextStorageId) -> Result<Vec<TextHyperlink>> {
        let object_id = object_id.get();
        text_hyperlinks(&self.package, object_id)
    }

    /// Create a native hyperlink over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_hyperlink(
        &mut self,
        object_id: TextStorageId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let object_id = object_id.get();
        add_text_hyperlink(&mut self.package, object_id, range, &target)
    }

    /// Atomically update a hyperlink's range and target while retaining its ID.
    pub fn update_text_hyperlink(
        &mut self,
        object_id: TextStorageId,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let object_id = object_id.get();
        update_text_hyperlink(&mut self.package, object_id, id, range, &target)
    }

    /// Delete one native hyperlink and reclaim its owned smart-field object.
    pub fn remove_text_hyperlink(
        &mut self,
        object_id: TextStorageId,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        let object_id = object_id.get();
        remove_text_hyperlink(&mut self.package, object_id, id)
    }

    /// Read every native ranged bookmark in a text storage.
    pub fn text_bookmarks(&self, object_id: TextStorageId) -> Result<Vec<TextBookmark>> {
        let object_id = object_id.get();
        text_bookmarks(&self.package, object_id)
    }

    /// Create a native bookmark over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_bookmark(
        &mut self,
        object_id: TextStorageId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        let object_id = object_id.get();
        add_text_bookmark(&mut self.package, object_id, range, &settings)
    }

    /// Atomically update a bookmark's range and settings while retaining its ID.
    pub fn update_text_bookmark(
        &mut self,
        object_id: TextStorageId,
        id: TextBookmarkId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        let object_id = object_id.get();
        update_text_bookmark(&mut self.package, object_id, id, range, &settings)
    }

    /// Delete one native bookmark and reclaim its owned bookmark-field object.
    pub fn remove_text_bookmark(
        &mut self,
        object_id: TextStorageId,
        id: TextBookmarkId,
    ) -> Result<TextBookmark> {
        let object_id = object_id.get();
        remove_text_bookmark(&mut self.package, object_id, id)
    }

    /// Read every native Date & Time smart field in a text storage.
    pub fn text_date_time_fields(
        &self,
        object_id: TextStorageId,
    ) -> Result<Vec<TextDateTimeField>> {
        let object_id = object_id.get();
        text_date_time_fields(&self.package, object_id)
    }

    /// Attach a Date & Time field to existing nonempty text.
    pub fn add_text_date_time_field(
        &mut self,
        object_id: TextStorageId,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let object_id = object_id.get();
        add_text_date_time_field(&mut self.package, object_id, range, &settings)
    }

    /// Atomically insert exact display text and attach a Date & Time field.
    pub fn insert_text_date_time_field(
        &mut self,
        object_id: TextStorageId,
        position: TextPosition,
        display_text: DisplayText,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let object_id = object_id.get();
        insert_text_date_time_field(
            &mut self.package,
            object_id,
            position,
            &display_text,
            &settings,
        )
    }

    /// Atomically update a Date & Time field's range and formatter payload.
    pub fn update_text_date_time_field(
        &mut self,
        object_id: TextStorageId,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let object_id = object_id.get();
        update_text_date_time_field(&mut self.package, object_id, id, range, &settings)
    }

    /// Delete one Date & Time field while retaining its visible text.
    pub fn remove_text_date_time_field(
        &mut self,
        object_id: TextStorageId,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        let object_id = object_id.get();
        remove_text_date_time_field(&mut self.package, object_id, id)
    }

    /// Read every native page/slide-number attachment in a text storage.
    pub fn text_number_attachments(
        &self,
        object_id: TextStorageId,
    ) -> Result<Vec<TextNumberAttachment>> {
        let object_id = object_id.get();
        text_number_attachments(&self.package, object_id)
    }

    /// Atomically insert U+FFFC and its native textual number attachment.
    ///
    /// Attachment evaluation is storage-context dependent. Prefer the Pages
    /// high-level wrappers when creating page numbers or page counts.
    pub fn insert_text_number_attachment(
        &mut self,
        object_id: TextStorageId,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        let object_id = object_id.get();
        insert_text_number_attachment(&mut self.package, object_id, position, &settings)
    }

    /// Atomically update the lossless payload of a number attachment.
    pub fn update_text_number_attachment(
        &mut self,
        object_id: TextStorageId,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        let object_id = object_id.get();
        update_text_number_attachment(&mut self.package, object_id, id, &settings)
    }

    /// Delete one number attachment together with its U+FFFC placeholder.
    pub fn remove_text_number_attachment(
        &mut self,
        object_id: TextStorageId,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment> {
        let object_id = object_id.get();
        remove_text_number_attachment(&mut self.package, object_id, id)
    }

    /// Read every native plain highlight in a text storage.
    pub fn text_highlights(&self, object_id: TextStorageId) -> Result<Vec<TextHighlight>> {
        let object_id = object_id.get();
        text_highlights(&self.package, object_id)
    }

    /// Create a native plain highlight over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_highlight(
        &mut self,
        object_id: TextStorageId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let object_id = object_id.get();
        add_text_highlight(&mut self.package, object_id, range)
    }

    /// Atomically move a plain highlight while retaining its native identity.
    pub fn update_text_highlight(
        &mut self,
        object_id: TextStorageId,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let object_id = object_id.get();
        update_text_highlight(&mut self.package, object_id, id, range)
    }

    /// Delete one plain highlight and its owned empty annotation graph.
    pub fn remove_text_highlight(
        &mut self,
        object_id: TextStorageId,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        let object_id = object_id.get();
        remove_text_highlight(&mut self.package, object_id, id)
    }

    /// Read every native ranged comment in a text storage.
    pub fn text_comments(&self, object_id: TextStorageId) -> Result<Vec<TextComment>> {
        let object_id = object_id.get();
        text_comments(&self.package, object_id)
    }

    /// Create a native comment over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_comment(
        &mut self,
        object_id: TextStorageId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let object_id = object_id.get();
        add_text_comment(&mut self.package, object_id, range, body)
    }

    /// Atomically update a comment's range and body while retaining its ID.
    pub fn update_text_comment(
        &mut self,
        object_id: TextStorageId,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let object_id = object_id.get();
        update_text_comment(&mut self.package, object_id, id, range, body)
    }

    /// Delete one ranged comment and its owned root/reply annotation graph.
    pub fn remove_text_comment(
        &mut self,
        object_id: TextStorageId,
        id: TextCommentId,
    ) -> Result<TextComment> {
        let object_id = object_id.get();
        remove_text_comment(&mut self.package, object_id, id)
    }

    /// Read every direct reply to one ranged comment in stored order.
    pub fn text_comment_replies(
        &self,
        object_id: TextStorageId,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        let object_id = object_id.get();
        text_comment_replies(&self.package, object_id, comment_id)
    }

    /// Append a direct reply to one ranged comment.
    pub fn add_text_comment_reply(
        &mut self,
        object_id: TextStorageId,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let object_id = object_id.get();
        add_text_comment_reply(&mut self.package, object_id, comment_id, body)
    }

    /// Update a direct reply while retaining its ID and native metadata.
    pub fn update_text_comment_reply(
        &mut self,
        object_id: TextStorageId,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let object_id = object_id.get();
        update_text_comment_reply(&mut self.package, object_id, comment_id, reply_id, body)
    }

    /// Delete one direct reply and its owned comment storage.
    pub fn remove_text_comment_reply(
        &mut self,
        object_id: TextStorageId,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        let object_id = object_id.get();
        remove_text_comment_reply(&mut self.package, object_id, comment_id, reply_id)
    }

    /// Read the canonical list preset applied uniformly to a text storage.
    pub fn paragraph_list(&self, object_id: TextStorageId) -> Result<ParagraphList> {
        let object_id = object_id.get();
        paragraph_list(&self.package, object_id)
    }

    /// Atomically apply a canonical nine-level list preset to a text storage.
    ///
    /// Storages with multiple list-style boundaries are rejected so the
    /// operation cannot flatten independently formatted paragraphs.
    pub fn set_paragraph_list(
        &mut self,
        object_id: TextStorageId,
        list: ParagraphList,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list(&mut staged, object_id, list)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list(&verified, object_id)? != list {
            return Err(Error::InvalidFormat(
                "iWork paragraph-list update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Remove uniform list formatting by applying the canonical None preset.
    pub fn reset_paragraph_list(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list(&verified, object_id)? != ParagraphList::None {
                return Err(Error::InvalidFormat(
                    "iWork paragraph-list reset failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read every canonical list-preset boundary in a text storage.
    pub fn paragraph_lists(&self, object_id: TextStorageId) -> Result<Vec<ParagraphListPlacement>> {
        let object_id = object_id.get();
        paragraph_lists(&self.package, object_id)
    }

    /// Atomically replace all list-preset boundaries in a text storage.
    ///
    /// Placements must be strictly increasing validated paragraph starts and
    /// begin at [`TextPosition::ZERO`]. Adjacent equal presets are coalesced.
    pub fn set_paragraph_lists(
        &mut self,
        object_id: TextStorageId,
        placements: &[ParagraphListPlacement],
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_lists(&mut staged, object_id, placements)?;
        let expected = paragraph_lists(&staged, object_id)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_lists(&verified, object_id)? != expected {
            return Err(Error::InvalidFormat(
                "iWork paragraph-list placements failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Read every effective list-level boundary in a text storage.
    pub fn paragraph_list_levels(
        &self,
        object_id: TextStorageId,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        let object_id = object_id.get();
        paragraph_list_levels(&self.package, object_id)
    }

    /// Read the effective list level at one validated paragraph start.
    pub fn paragraph_list_level(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel> {
        let object_id = object_id.get();
        paragraph_list_level(&self.package, object_id, paragraph)
    }

    /// Atomically set one paragraph's zero-based list nesting level.
    pub fn set_paragraph_list_level(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_level(&mut staged, object_id, paragraph, level)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_level(&verified, object_id, paragraph)? != level {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-level update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_paragraph_list_level(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_level(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_level(&verified, object_id, paragraph)? != ParagraphListLevel::ZERO {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-level reset failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read whether one paragraph continues or restarts numbered-list sequencing.
    pub fn paragraph_list_numbering(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering> {
        let object_id = object_id.get();
        paragraph_list_numbering(&self.package, object_id, paragraph)
    }

    /// Continue or restart numbered-list sequencing at one paragraph.
    pub fn set_paragraph_list_numbering(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_numbering(&mut staged, object_id, paragraph, numbering)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_numbering(&verified, object_id, paragraph)? != numbering {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-numbering update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Read one numbered paragraph's effective locale-aware label format.
    pub fn paragraph_list_number_format(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberFormat> {
        let object_id = object_id.get();
        paragraph_list_number_format(&self.package, object_id, paragraph)
    }

    /// Atomically set one numbered paragraph's locale-aware label format.
    pub fn set_paragraph_list_number_format(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        format: ParagraphListNumberFormat,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_number_format(&mut staged, object_id, paragraph, format)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_number_format(&verified, object_id, paragraph)? != format {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-number format update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore Apple's standard decimal-period format (`1.`, `2.`, `3.`).
    pub fn reset_paragraph_list_number_format(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_number_format(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_number_format(&verified, object_id, paragraph)?
                != ParagraphListNumberFormat::DECIMAL
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-number format reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read whether one numbered paragraph displays its full hierarchical number.
    pub fn paragraph_list_number_tiering(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberTiering> {
        let object_id = object_id.get();
        paragraph_list_number_tiering(&self.package, object_id, paragraph)
    }

    /// Atomically choose flat or hierarchical numbering for one list level.
    pub fn set_paragraph_list_number_tiering(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        tiering: ParagraphListNumberTiering,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_number_tiering(&mut staged, object_id, paragraph, tiering)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_number_tiering(&verified, object_id, paragraph)? != tiering {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-number tiering update failed round-trip validation"
                    .to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore flat numbering for one numbered-list level.
    pub fn reset_paragraph_list_number_tiering(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_number_tiering(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_number_tiering(&verified, object_id, paragraph)?
                != ParagraphListNumberTiering::Flat
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-number tiering reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read one numbered paragraph's effective number-label size.
    pub fn paragraph_list_number_scale(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberScale> {
        let object_id = object_id.get();
        paragraph_list_number_scale(&self.package, object_id, paragraph)
    }

    /// Atomically set one numbered paragraph's number-label size.
    pub fn set_paragraph_list_number_scale(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        scale: ParagraphListNumberScale,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_number_scale(&mut staged, object_id, paragraph, scale)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_number_scale(&verified, object_id, paragraph)? != scale {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-number scale update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_paragraph_list_number_scale(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_number_scale(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_number_scale(&verified, object_id, paragraph)?
                != ParagraphListNumberScale::ONE
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-number scale reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective text-bullet marker at one paragraph's list level.
    pub fn paragraph_list_bullet(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBullet> {
        let object_id = object_id.get();
        paragraph_list_bullet(&self.package, object_id, paragraph)
    }

    /// Atomically set one bullet-list paragraph's marker.
    pub fn set_paragraph_list_bullet(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        bullet: &ParagraphListBullet,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_bullet(&mut staged, object_id, paragraph, bullet)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_bullet(&verified, object_id, paragraph)? != *bullet {
            return Err(Error::InvalidFormat(
                "iWork paragraph text-bullet update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore Apple's standard `•` marker at one bullet-list paragraph.
    pub fn reset_paragraph_list_bullet(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_bullet(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_bullet(&verified, object_id, paragraph)?.as_str()
                != ParagraphListBullet::STANDARD
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph text-bullet reset failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read one bullet-list paragraph's effective marker size and baseline.
    pub fn paragraph_list_bullet_geometry(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBulletGeometry> {
        let object_id = object_id.get();
        paragraph_list_bullet_geometry(&self.package, object_id, paragraph)
    }

    /// Atomically set one bullet-list paragraph's marker size and baseline.
    pub fn set_paragraph_list_bullet_geometry(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        geometry: ParagraphListBulletGeometry,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_bullet_geometry(&mut staged, object_id, paragraph, geometry)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_bullet_geometry(&verified, object_id, paragraph)? != geometry {
            return Err(Error::InvalidFormat(
                "iWork paragraph text-bullet geometry update failed round-trip validation"
                    .to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore Apple's standard marker size and baseline for this nesting level.
    pub fn reset_paragraph_list_bullet_geometry(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_bullet_geometry(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            let level = paragraph_list_level(&verified, object_id, paragraph)?;
            if paragraph_list_bullet_geometry(&verified, object_id, paragraph)?
                != ParagraphListBulletGeometry::standard(level)
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph text-bullet geometry reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read one list paragraph's effective label and text-gap indentation.
    pub fn paragraph_list_indentation(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListIndentation> {
        let object_id = object_id.get();
        paragraph_list_indentation(&self.package, object_id, paragraph)
    }

    /// Atomically set one list paragraph's label and text-gap indentation.
    pub fn set_paragraph_list_indentation(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        indentation: ParagraphListIndentation,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_indentation(&mut staged, object_id, paragraph, indentation)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_indentation(&verified, object_id, paragraph)? != indentation {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-indentation update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_paragraph_list_indentation(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_indentation(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            let list = paragraph_lists(&verified, object_id)?
                .into_iter()
                .take_while(|placement| placement.paragraph <= paragraph)
                .last()
                .map(|placement| placement.list)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "iWork paragraph list-indentation reset lost its list preset".to_owned(),
                    )
                })?;
            let level = paragraph_list_level(&verified, object_id, paragraph)?;
            if paragraph_list_indentation(&verified, object_id, paragraph)?
                != ParagraphListIndentation::standard(list, level)?
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-indentation reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read one list paragraph's effective label color.
    pub fn paragraph_list_label_color(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLabelColor> {
        let object_id = object_id.get();
        paragraph_list_label_color(&self.package, object_id, paragraph)
    }

    /// Atomically set one list paragraph's bullet or number color.
    pub fn set_paragraph_list_label_color(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        color: ParagraphListLabelColor,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_list_label_color(&mut staged, object_id, paragraph, color)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_list_label_color(&verified, object_id, paragraph)? != color {
            return Err(Error::InvalidFormat(
                "iWork paragraph list-label color update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the label to the paragraph's automatic text color.
    pub fn reset_paragraph_list_label_color(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_list_label_color(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_list_label_color(&verified, object_id, paragraph)?
                != ParagraphListLabelColor::Automatic
            {
                return Err(Error::InvalidFormat(
                    "iWork paragraph list-label color reset failed round-trip validation"
                        .to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read effective uniform underline and strikethrough formatting.
    pub fn text_decorations(&self, object_id: TextStorageId) -> Result<TextDecorations> {
        let object_id = object_id.get();
        text_decorations(&self.package, object_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_decorations(
        &mut self,
        object_id: TextStorageId,
        decorations: TextDecorations,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_decorations(&mut staged, object_id, decorations)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_decorations(&verified, object_id)? != decorations {
            return Err(Error::InvalidFormat(
                "iWork text-decoration update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited text decorations while preserving sibling overrides.
    pub fn reset_text_decorations(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_decorations(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform text color.
    pub fn text_color(&self, object_id: TextStorageId) -> Result<RgbaColor> {
        let object_id = object_id.get();
        text_color(&self.package, object_id)
    }

    /// Atomically set one text color across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_color(&mut self, object_id: TextStorageId, color: RgbaColor) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_color(&mut staged, object_id, color)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_color(&verified, object_id)? != color {
            return Err(Error::InvalidFormat(
                "iWork text-color update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited text color while preserving sibling overrides.
    pub fn reset_text_color(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_color(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform capitalization mode.
    pub fn text_capitalization(&self, object_id: TextStorageId) -> Result<TextCapitalization> {
        let object_id = object_id.get();
        text_capitalization(&self.package, object_id)
    }

    /// Atomically set one capitalization mode across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_capitalization(
        &mut self,
        object_id: TextStorageId,
        capitalization: TextCapitalization,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_capitalization(&mut staged, object_id, capitalization)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_capitalization(&verified, object_id)? != capitalization {
            return Err(Error::InvalidFormat(
                "iWork text-capitalization update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited capitalization while preserving sibling overrides.
    pub fn reset_text_capitalization(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_capitalization(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform baseline script.
    pub fn text_script(&self, object_id: TextStorageId) -> Result<TextScript> {
        let object_id = object_id.get();
        text_script(&self.package, object_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_script(&mut self, object_id: TextStorageId, script: TextScript) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_script(&mut staged, object_id, script)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_script(&verified, object_id)? != script {
            return Err(Error::InvalidFormat(
                "iWork text-script update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited baseline script while preserving sibling overrides.
    pub fn reset_text_script(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_script(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform custom baseline displacement.
    pub fn text_baseline_shift(&self, object_id: TextStorageId) -> Result<TextBaselineShift> {
        let object_id = object_id.get();
        text_baseline_shift(&self.package, object_id)
    }

    /// Atomically set a signed custom baseline displacement.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_baseline_shift(
        &mut self,
        object_id: TextStorageId,
        shift: TextBaselineShift,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_baseline_shift(&mut staged, object_id, shift)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_baseline_shift(&verified, object_id)? != shift {
            return Err(Error::InvalidFormat(
                "iWork text baseline-shift update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited baseline displacement while preserving sibling overrides.
    pub fn reset_text_baseline_shift(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_baseline_shift(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform character spacing.
    pub fn text_character_spacing(&self, object_id: TextStorageId) -> Result<TextCharacterSpacing> {
        let object_id = object_id.get();
        text_character_spacing(&self.package, object_id)
    }

    /// Atomically set character spacing across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_character_spacing(
        &mut self,
        object_id: TextStorageId,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_character_spacing(&mut staged, object_id, spacing)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_character_spacing(&verified, object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "iWork text character-spacing update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited character spacing while preserving sibling overrides.
    pub fn reset_text_character_spacing(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_character_spacing(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform ligature policy.
    pub fn text_ligatures(&self, object_id: TextStorageId) -> Result<TextLigatures> {
        let object_id = object_id.get();
        text_ligatures(&self.package, object_id)
    }

    /// Atomically set the ligature policy across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_ligatures(
        &mut self,
        object_id: TextStorageId,
        ligatures: TextLigatures,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_ligatures(&mut staged, object_id, ligatures)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_ligatures(&verified, object_id)? != ligatures {
            return Err(Error::InvalidFormat(
                "iWork text ligature update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited ligatures while preserving sibling overrides.
    pub fn reset_text_ligatures(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_ligatures(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform text outline.
    pub fn text_outline(&self, object_id: TextStorageId) -> Result<Outline> {
        let object_id = object_id.get();
        text_outline(&self.package, object_id)
    }

    /// Atomically set a typed outline stroke across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_outline(&mut self, object_id: TextStorageId, outline: Outline) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_outline(&mut staged, object_id, outline)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_outline(&verified, object_id)? != outline {
            return Err(Error::InvalidFormat(
                "iWork text outline update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited outline while preserving sibling overrides.
    pub fn reset_text_outline(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_outline(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform text shadow.
    pub fn text_shadow(&self, object_id: TextStorageId) -> Result<Shadow> {
        let object_id = object_id.get();
        text_shadow(&self.package, object_id)
    }

    /// Atomically set a typed drop shadow across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_shadow(&mut self, object_id: TextStorageId, shadow: Shadow) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_shadow(&mut staged, object_id, shadow)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_shadow(&verified, object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "iWork text shadow update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited shadow while preserving sibling overrides.
    pub fn reset_text_shadow(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_shadow(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective solid background behind uniformly styled text.
    pub fn text_background(&self, object_id: TextStorageId) -> Result<Background> {
        let object_id = object_id.get();
        text_background(&self.package, object_id)
    }

    /// Atomically set a typed solid background across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_background(
        &mut self,
        object_id: TextStorageId,
        background: Background,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_text_background(&mut staged, object_id, background)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if text_background(&verified, object_id)? != background {
            return Err(Error::InvalidFormat(
                "iWork text background update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited text background while preserving sibling overrides.
    pub fn reset_text_background(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_text_background(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective solid fill across a uniform paragraph layout box.
    pub fn paragraph_background(&self, object_id: TextStorageId) -> Result<ParagraphBackground> {
        let object_id = object_id.get();
        paragraph_background(&self.package, object_id)
    }

    /// Atomically set the Text → Layout paragraph background.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_paragraph_background(
        &mut self,
        object_id: TextStorageId,
        background: ParagraphBackground,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_background(&mut staged, object_id, background)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_background(&verified, object_id)? != background {
            return Err(Error::InvalidFormat(
                "iWork paragraph background update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited paragraph background while preserving sibling overrides.
    pub fn reset_paragraph_background(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_background(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective paragraph borders across a uniform paragraph layout box.
    pub fn paragraph_borders(&self, object_id: TextStorageId) -> Result<Borders> {
        let object_id = object_id.get();
        paragraph_borders(&self.package, object_id)
    }

    /// Atomically set the Text → Layout paragraph borders.
    pub fn set_paragraph_borders(
        &mut self,
        object_id: TextStorageId,
        borders: Borders,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_borders(&mut staged, object_id, borders)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_borders(&verified, object_id)? != borders {
            return Err(Error::InvalidFormat(
                "iWork paragraph border update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited paragraph borders while preserving sibling overrides.
    pub fn reset_paragraph_borders(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_borders(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective paragraph pagination and hyphenation controls.
    pub fn paragraph_flow(&self, object_id: TextStorageId) -> Result<ParagraphFlow> {
        let object_id = object_id.get();
        paragraph_flow(&self.package, object_id)
    }

    /// Atomically set the paragraph pagination and hyphenation controls.
    pub fn set_paragraph_flow(
        &mut self,
        object_id: TextStorageId,
        flow: ParagraphFlow,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_flow(&mut staged, object_id, flow)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_flow(&verified, object_id)? != flow {
            return Err(Error::InvalidFormat(
                "iWork paragraph flow update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited paragraph flow while preserving sibling overrides.
    pub fn reset_paragraph_flow(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_flow(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective base-writing direction of a text storage.
    pub fn paragraph_writing_direction(
        &self,
        object_id: TextStorageId,
    ) -> Result<ParagraphWritingDirection> {
        let object_id = object_id.get();
        paragraph_writing_direction(&self.package, object_id)
    }

    /// Atomically set the base-writing direction for every paragraph.
    pub fn set_paragraph_writing_direction(
        &mut self,
        object_id: TextStorageId,
        direction: ParagraphWritingDirection,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_writing_direction(&mut staged, object_id, direction)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_writing_direction(&verified, object_id)? != direction {
            return Err(Error::InvalidFormat(
                "iWork paragraph writing-direction update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited base-writing direction while preserving sibling overrides.
    pub fn reset_paragraph_writing_direction(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_writing_direction(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// List the theme paragraph-style presets selectable for this text storage.
    pub fn named_paragraph_styles(
        &self,
        object_id: TextStorageId,
    ) -> Result<Vec<NamedParagraphStyle>> {
        let object_id = object_id.get();
        named_paragraph_styles(&self.package, object_id)
    }

    /// Read the named paragraph style selected for one uniform text storage.
    pub fn applied_named_paragraph_style(
        &self,
        object_id: TextStorageId,
    ) -> Result<AppliedParagraphStyle> {
        let object_id = object_id.get();
        applied_named_paragraph_style(&self.package, object_id)
    }

    /// Redefine the selected named style from this storage's direct overrides.
    ///
    /// The overrides become part of the shared named definition, update every
    /// use of that style, and are cleared from this storage.
    pub fn redefine_applied_named_paragraph_style(
        &mut self,
        object_id: TextStorageId,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let redefined = redefine_applied_named_paragraph_style(&mut staged, object_id)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        let selection = applied_named_paragraph_style(&verified, object_id)?;
        if selection.style() != &redefined || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "iWork named paragraph-style redefinition failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(redefined)
    }

    /// Apply one selectable paragraph style and clear direct paragraph overrides.
    pub fn apply_named_paragraph_style(
        &mut self,
        object_id: TextStorageId,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let applied = apply_named_paragraph_style(&mut staged, object_id, target)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        let selection = applied_named_paragraph_style(&verified, object_id)?;
        if selection.style() != &applied || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "iWork named paragraph-style application failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(applied)
    }

    /// Clone one theme preset as a new named paragraph style.
    pub fn create_named_paragraph_style(
        &mut self,
        object_id: TextStorageId,
        source: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let created = create_named_paragraph_style(&mut staged, object_id, source, name)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if !named_paragraph_styles(&verified, object_id)?.contains(&created) {
            return Err(Error::InvalidFormat(
                "iWork named paragraph-style creation failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(created)
    }

    /// Rename a selectable paragraph style without changing its stable identifier.
    pub fn rename_named_paragraph_style(
        &mut self,
        object_id: TextStorageId,
        target: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let renamed = rename_named_paragraph_style(&mut staged, object_id, target, name)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if !named_paragraph_styles(&verified, object_id)?.contains(&renamed) {
            return Err(Error::InvalidFormat(
                "iWork named paragraph-style rename failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(renamed)
    }

    /// Delete an unused selectable paragraph style.
    ///
    /// Styles referenced by text, inheritance, following-style rules, or an
    /// unknown object are rejected transactionally until the caller replaces
    /// those references.
    pub fn delete_named_paragraph_style(
        &mut self,
        object_id: TextStorageId,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let deleted = delete_named_paragraph_style(&mut staged, object_id, target)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if named_paragraph_styles(&verified, object_id)?
            .iter()
            .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "iWork named paragraph-style deletion failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(deleted)
    }

    /// Replace the current named style and delete it as one transaction.
    ///
    /// The operation clears current direct overrides, applies `replacement`,
    /// and removes `target`. References from any other object make the whole
    /// operation fail without changing the package.
    pub fn delete_applied_named_paragraph_style_with_replacement(
        &mut self,
        object_id: TextStorageId,
        target: ParagraphStyleId,
        replacement: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let object_id = object_id.get();
        if target == replacement {
            return Err(Error::InvalidFormat(
                "a deleted iWork paragraph style requires a different replacement".to_owned(),
            ));
        }
        let current = applied_named_paragraph_style(&self.package, object_id)?;
        if current.style().id() != target {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {object_id} does not currently use paragraph style {}",
                native_id(target)
            )));
        }

        let mut staged = self.package.clone();
        apply_named_paragraph_style(&mut staged, object_id, replacement)?;
        let deleted = delete_named_paragraph_style(&mut staged, object_id, target)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        let selection = applied_named_paragraph_style(&verified, object_id)?;
        if selection.style().id() != replacement
            || selection.has_overrides()
            || named_paragraph_styles(&verified, object_id)?
                .iter()
                .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "iWork paragraph-style replacement deletion failed round-trip validation"
                    .to_owned(),
            ));
        }
        self.package = staged;
        Ok(deleted)
    }

    /// Read the paragraph style selected for the paragraph created by pressing Return.
    pub fn paragraph_following_style(
        &self,
        object_id: TextStorageId,
    ) -> Result<ParagraphFollowingStyle> {
        let object_id = object_id.get();
        paragraph_following_style(&self.package, object_id)
    }

    /// Atomically select the paragraph style used after pressing Return.
    pub fn set_paragraph_following_style(
        &mut self,
        object_id: TextStorageId,
        following_style: ParagraphFollowingStyle,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_following_style(&mut staged, object_id, following_style)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_following_style(&verified, object_id)? != following_style {
            return Err(Error::InvalidFormat(
                "iWork following paragraph-style update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited following paragraph style while preserving sibling overrides.
    pub fn reset_paragraph_following_style(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_following_style(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform paragraph alignment of a text storage.
    pub fn paragraph_alignment(&self, object_id: TextStorageId) -> Result<Alignment> {
        let object_id = object_id.get();
        paragraph_alignment(&self.package, object_id)
    }

    /// Set one alignment for every paragraph in a uniformly styled text storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// this whole-storage operation can never flatten unrelated formatting.
    pub fn set_paragraph_alignment(
        &mut self,
        object_id: TextStorageId,
        alignment: Alignment,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_alignment(&mut staged, object_id, alignment)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_alignment(&verified, object_id)? != alignment {
            return Err(Error::InvalidFormat(
                "iWork paragraph-alignment update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Remove a private minimal alignment override and restore its parent style.
    pub fn reset_paragraph_alignment(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_alignment(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform line spacing of a text storage.
    pub fn paragraph_line_spacing(&self, object_id: TextStorageId) -> Result<LineSpacing> {
        let object_id = object_id.get();
        paragraph_line_spacing(&self.package, object_id)
    }

    /// Set one typed line-spacing mode across a uniformly styled text storage.
    pub fn set_paragraph_line_spacing(
        &mut self,
        object_id: TextStorageId,
        spacing: LineSpacing,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_line_spacing(&mut staged, object_id, spacing)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_line_spacing(&verified, object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "iWork paragraph line-spacing update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Remove a private line-spacing override while preserving sibling overrides.
    pub fn reset_paragraph_line_spacing(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_line_spacing(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective before/after spacing of a uniformly styled text storage.
    pub fn paragraph_spacing(&self, object_id: TextStorageId) -> Result<Spacing> {
        let object_id = object_id.get();
        paragraph_spacing(&self.package, object_id)
    }

    /// Atomically set before/after paragraph spacing across a uniform text storage.
    pub fn set_paragraph_spacing(
        &mut self,
        object_id: TextStorageId,
        spacing: Spacing,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_spacing(&mut staged, object_id, spacing)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_spacing(&verified, object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "iWork paragraph spacing update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited before/after spacing while preserving sibling overrides.
    pub fn reset_paragraph_spacing(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_spacing(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right indentation of a uniform storage.
    pub fn paragraph_indents(&self, object_id: TextStorageId) -> Result<Indents> {
        let object_id = object_id.get();
        paragraph_indents(&self.package, object_id)
    }

    /// Atomically set first-line, left, and right paragraph indentation.
    pub fn set_paragraph_indents(
        &mut self,
        object_id: TextStorageId,
        indents: Indents,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_indents(&mut staged, object_id, indents)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_indents(&verified, object_id)? != indents {
            return Err(Error::InvalidFormat(
                "iWork paragraph indentation update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited indentation while preserving sibling paragraph overrides.
    pub fn reset_paragraph_indents(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_indents(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective character used to align decimal tab stops.
    pub fn paragraph_decimal_tab_character(
        &self,
        object_id: TextStorageId,
    ) -> Result<ParagraphDecimalTabCharacter> {
        let object_id = object_id.get();
        paragraph_decimal_tab_character(&self.package, object_id)
    }

    /// Atomically set the character used to align decimal tab stops.
    pub fn set_paragraph_decimal_tab_character(
        &mut self,
        object_id: TextStorageId,
        character: ParagraphDecimalTabCharacter,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_decimal_tab_character(&mut staged, object_id, character)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_decimal_tab_character(&verified, object_id)? != character {
            return Err(Error::InvalidFormat(
                "iWork paragraph decimal-tab character update failed round-trip validation"
                    .to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited decimal-tab character.
    pub fn reset_paragraph_decimal_tab_character(
        &mut self,
        object_id: TextStorageId,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_decimal_tab_character(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective distance between implicit paragraph tab stops.
    pub fn paragraph_default_tab_interval(
        &self,
        object_id: TextStorageId,
    ) -> Result<ParagraphDefaultTabInterval> {
        let object_id = object_id.get();
        paragraph_default_tab_interval(&self.package, object_id)
    }

    /// Atomically set the distance between implicit paragraph tab stops.
    pub fn set_paragraph_default_tab_interval(
        &mut self,
        object_id: TextStorageId,
        interval: ParagraphDefaultTabInterval,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_default_tab_interval(&mut staged, object_id, interval)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_default_tab_interval(&verified, object_id)? != interval {
            return Err(Error::InvalidFormat(
                "iWork paragraph default-tab interval update failed round-trip validation"
                    .to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore the inherited default-tab interval.
    pub fn reset_paragraph_default_tab_interval(
        &mut self,
        object_id: TextStorageId,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_default_tab_interval(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of a uniform text storage.
    pub fn paragraph_tab_stops(&self, object_id: TextStorageId) -> Result<ParagraphTabStops> {
        let object_id = object_id.get();
        paragraph_tab_stops(&self.package, object_id)
    }

    /// Atomically replace every explicit ruler tab stop of a uniform text storage.
    pub fn set_paragraph_tab_stops(
        &mut self,
        object_id: TextStorageId,
        stops: ParagraphTabStops,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_tab_stops(&mut staged, object_id, &stops)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_tab_stops(&verified, object_id)? != stops {
            return Err(Error::InvalidFormat(
                "iWork paragraph tab-stop update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Restore inherited tab stops while preserving sibling paragraph overrides.
    pub fn reset_paragraph_tab_stops(&mut self, object_id: TextStorageId) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = reset_paragraph_tab_stops(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// List plain-text drop caps in paragraph order.
    pub fn paragraph_drop_caps(&self, object_id: TextStorageId) -> Result<Vec<Placement>> {
        let object_id = object_id.get();
        paragraph_drop_caps(&self.package, object_id)
    }

    /// Read the drop cap attached to one typed paragraph position.
    pub fn paragraph_drop_cap(
        &self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<Option<DropCap>> {
        let object_id = object_id.get();
        paragraph_drop_cap(&self.package, object_id, paragraph)
    }

    /// Atomically create or replace a plain-text Drop Cap.
    pub fn set_paragraph_drop_cap(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
        drop_cap: DropCap,
    ) -> Result<()> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        set_paragraph_drop_cap(&mut staged, object_id, paragraph, drop_cap)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_drop_cap(&verified, object_id, paragraph)? != Some(drop_cap) {
            return Err(Error::InvalidFormat(
                "iWork Drop Cap update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Atomically remove a drop cap while retaining its paragraph boundary.
    pub fn remove_paragraph_drop_cap(
        &mut self,
        object_id: TextStorageId,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let object_id = object_id.get();
        let mut staged = self.package.clone();
        let changed = remove_paragraph_drop_cap(&mut staged, object_id, paragraph)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_drop_cap(&verified, object_id, paragraph)?.is_some() {
                return Err(Error::InvalidFormat(
                    "iWork Drop Cap removal failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    pub(crate) fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub(crate) fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

pub(super) fn replace_storage_text(
    package: &mut IWorkPackage,
    object_id: u64,
    range: Range<usize>,
    replacement: &str,
) -> Result<()> {
    let located = locate_text_storage_with_archive(package, object_id)?;
    let location = located.location;
    let archive_name = location.archive_name;
    let message_index = location.message_index;
    let message_type = location.message_type;
    let object = located
        .archive
        .object(object_id)
        .ok_or_else(|| Error::ParseError(format!("Text storage object {object_id} not found")))?;
    let message_info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "text storage object {object_id} has no metadata for message {message_index}"
            ))
        })?;
    let original = object
        .messages
        .get(message_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "text storage object {object_id} has no payload at message {message_index}"
            ))
        })?
        .data
        .as_slice();
    let package_limits = package.limits();
    let retained_archive_limits = package_limits.archive_limits();
    let archive_limits = retained_archive_limits.with_archive_bytes(
        package_limits
            .max_iwa_stream_bytes()
            .min(retained_archive_limits.max_archive_bytes()),
    )?;
    let default_rewrite_limits = litchi_iwa_text_wire::RewriteLimits::default();
    let message_bytes = default_rewrite_limits
        .max_message_bytes()
        .min(archive_limits.max_message_bytes());
    let rewrite_limits = litchi_iwa_text_wire::RewriteLimits::new(
        message_bytes,
        default_rewrite_limits.max_fields(),
        default_rewrite_limits.max_nesting(),
        default_rewrite_limits.max_fragments(),
        default_rewrite_limits.max_text_bytes().min(message_bytes),
        default_rewrite_limits.max_table_entries(),
        default_rewrite_limits.max_object_references(),
        default_rewrite_limits.max_output_bytes().min(message_bytes),
        default_rewrite_limits.max_rewrite_work(),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!("TSWP text-storage rewrite limit failed: {error}"))
    })?;
    let rewrite = litchi_iwa_text_wire::rewrite_storage_text_with_behavior_and_limits(
        original,
        range,
        replacement,
        litchi_iwa_text_wire::RewriteBehavior::ReplaceSelection,
        rewrite_limits,
    )
    .map_err(|error| Error::InvalidFormat(format!("TSWP text-storage rewrite failed: {error}")))?;
    if !rewrite.changed() {
        return Ok(());
    }

    let prunable_references = proven_prunable_storage_references(
        message_info,
        rewrite.removed_object_references(),
        rewrite.removed_object_references_by_field(),
        rewrite.object_reference_occurrences_before(),
        rewrite.has_unknown_wire_fields(),
    );
    // Object cleanup is authorized only for references whose metadata ownership
    // was proven here and which remain unreferenced after the header rewrite.
    let cleanup_references = prunable_references.iter().copied().collect::<HashSet<_>>();
    let data = rewrite.into_bytes();
    update_parsed_archive(package, &archive_name, located.archive, |archive| {
        let object = archive.object_mut(object_id).ok_or_else(|| {
            Error::ParseError(format!("Text storage object {object_id} not found"))
        })?;
        object.replace_message_pruning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
            &prunable_references,
            archive_limits,
        )?;
        Ok(())
    })?;
    let cleanup_candidates = metadata_unreferenced_candidates(package, &cleanup_references)?;
    remove_unreferenced_hyperlink_objects(package, &archive_name, &cleanup_candidates)?;
    remove_unreferenced_date_time_objects(package, &archive_name, &cleanup_candidates)?;
    remove_unreferenced_number_attachment_objects(package, &archive_name, &cleanup_candidates)?;
    remove_unreferenced_bookmark_objects(package, &archive_name, &cleanup_candidates)?;
    remove_unreferenced_highlight_objects(package, &archive_name, &cleanup_candidates)
}

fn proven_prunable_storage_references(
    message_info: &crate::archive::MessageInfo,
    aggregate_removals: &[u64],
    field_removals: &[litchi_iwa_text_wire::RemovedObjectReference],
    reference_occurrences_before: &[u64],
    has_unknown_wire_fields: bool,
) -> Vec<u64> {
    // An opaque field at any traversed schema level can own the same identifier.
    if has_unknown_wire_fields {
        return Vec::new();
    }
    let mut indexed = message_info.object_references.clone();
    indexed.sort_unstable();
    let top_level_references_are_complete = indexed == reference_occurrences_before;
    // Field-local deltas are insufficient: only identifiers absent from every
    // known storage reference after the rewrite can be removed from metadata.
    let mut identifiers = aggregate_removals.to_vec();
    identifiers.sort_unstable();
    identifiers.dedup();
    identifiers.retain(|identifier| {
        let top_level_occurrence = message_info.object_references.contains(identifier);
        if top_level_occurrence && !top_level_references_are_complete {
            return false;
        }
        let mut found = false;
        for field in &message_info.field_infos {
            if !field.object_references.contains(identifier) {
                continue;
            }
            found = true;
            let Some(storage_field) = exact_storage_table_reference_path(&field.path.path) else {
                return false;
            };
            if field.effective_type() != crate::archive::FieldType::ObjectReference
                || !field_removals.iter().any(|removal| {
                    removal.identifier() == *identifier
                        && removal.storage_field_number() == storage_field
                })
            {
                return false;
            }
        }
        found || top_level_occurrence
    });
    identifiers
}

fn exact_storage_table_reference_path(path: &[u32]) -> Option<u32> {
    let [storage_field, 1, 2] = path else {
        return None;
    };
    matches!(
        storage_field,
        5 | 7 | 8 | 9 | 11 | 12 | 15 | 16 | 17 | 18 | 21 | 22 | 23 | 25 | 26 | 27 | 28
    )
    .then_some(*storage_field)
}

fn metadata_unreferenced_candidates(
    package: &IWorkPackage,
    candidates: &HashSet<u64>,
) -> Result<HashSet<u64>> {
    if candidates.is_empty() {
        return Ok(HashSet::new());
    }
    let mut referenced = HashSet::new();
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in archive.objects {
            for info in object.archive_info.message_infos {
                for identifier in info.object_references {
                    if candidates.contains(&identifier) {
                        referenced.insert(identifier);
                    }
                }
                for field in info.field_infos {
                    for identifier in field.object_references {
                        if candidates.contains(&identifier) {
                            referenced.insert(identifier);
                        }
                    }
                }
            }
        }
    }
    Ok(candidates
        .difference(&referenced)
        .copied()
        .collect::<HashSet<_>>())
}

pub(crate) fn storage_object_references(storage: &StorageArchive) -> Vec<u64> {
    let mut references = Vec::new();
    if let Some(reference) = &storage.style_sheet {
        references.push(reference.identifier);
    }
    for table in [
        &storage.table_para_style,
        &storage.table_list_style,
        &storage.table_char_style,
        &storage.table_attachment,
        &storage.table_smartfield,
        &storage.table_layout_style,
        &storage.table_bookmark,
        &storage.table_footnote,
        &storage.table_section,
        &storage.table_rubyfield,
        &storage.table_insertion,
        &storage.table_deletion,
        &storage.table_highlight,
        &storage.table_tatechuyoko,
        &storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        references.extend(
            table
                .entries
                .iter()
                .filter_map(|entry| entry.object.as_ref().map(|value| value.identifier)),
        );
    }
    for table in [
        &storage.table_overlapping_highlight,
        &storage.table_pencil_annotation,
    ]
    .into_iter()
    .flatten()
    {
        references.extend(table.entries.iter().map(|entry| entry.field.identifier));
    }
    references.sort_unstable();
    references.dedup();
    references
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, FieldInfo, FieldType};
    use crate::protobuf::tsp::{Range as TspRange, Reference};
    use crate::protobuf::tswp::object_attribute_table::ObjectAttribute;
    use crate::protobuf::tswp::overlapping_field_attribute_table::OverlappingFieldAttribute;
    use crate::protobuf::tswp::{ObjectAttributeTable, OverlappingFieldAttributeTable};
    use prost::Message;

    const UNKNOWN_ARCHIVE_HEADER_FIELD: &[u8] = &[0x92, 0x06, 0x04, 0xde, 0xad, 0xbe, 0xef];

    #[test]
    fn storage_discovery_rejects_malformed_recognized_payload() {
        let editor = IWorkTextEditor::from_package(test_package_with_messages(vec![RawMessage {
            type_: 2001,
            data: vec![0x80],
        }]));

        assert!(editor.storages().is_err());
    }

    #[test]
    fn storage_lookup_does_not_skip_malformed_first_payload() {
        let valid = StorageArchive {
            text: vec!["valid sibling".to_owned()],
            ..Default::default()
        };
        let editor = IWorkTextEditor::from_package(test_package_with_messages(vec![
            RawMessage {
                type_: 2001,
                data: vec![0x80],
            },
            RawMessage {
                type_: 2022,
                data: valid.encode_to_vec(),
            },
        ]));

        assert!(
            editor
                .storage(TextStorageId::new(42).expect("valid test storage ID"))
                .is_err()
        );
    }

    #[test]
    fn storage_lookup_rejects_duplicate_recognized_payloads() {
        let first = StorageArchive {
            text: vec!["first".to_owned()],
            ..Default::default()
        };
        let second = StorageArchive {
            text: vec!["second".to_owned()],
            ..Default::default()
        };
        let editor = IWorkTextEditor::from_package(test_package_with_messages(vec![
            RawMessage {
                type_: 2001,
                data: first.encode_to_vec(),
            },
            RawMessage {
                type_: 2022,
                data: second.encode_to_vec(),
            },
        ]));

        assert!(
            editor
                .storage(TextStorageId::new(42).expect("valid test storage ID"))
                .is_err()
        );
        assert!(editor.storages().is_err());
    }

    #[test]
    fn storage_lookup_rejects_duplicate_archives() {
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2001,
                data: StorageArchive {
                    text: vec!["duplicated".to_owned()],
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
        )
        .unwrap();
        let archive = Archive {
            objects: vec![object],
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/Document.iwa", &archive)
            .unwrap();
        package
            .replace_archive("Index/Other.iwa", &archive)
            .unwrap();

        let editor = IWorkTextEditor::from_package(package);
        assert!(
            editor
                .storage(TextStorageId::new(42).expect("valid test storage ID"))
                .is_err()
        );
    }

    #[test]
    fn storage_replacement_is_atomic_when_resolution_fails() {
        let mut editor =
            IWorkTextEditor::from_package(test_package_with_messages(vec![RawMessage {
                type_: 2001,
                data: vec![0x80],
            }]));
        let before = editor.to_bytes().unwrap();

        assert!(
            editor
                .replace_text(
                    TextStorageId::new(42).expect("valid test storage ID"),
                    0..0,
                    "replacement",
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn unknown_message_types_are_ignored_and_preserved() {
        let unknown = RawMessage {
            type_: 9_999,
            data: vec![0xde, 0xad, 0xbe, 0xef],
        };
        let storage = StorageArchive {
            text: vec!["Source".to_owned()],
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_messages(vec![
            unknown.clone(),
            RawMessage {
                type_: 2001,
                data: storage.encode_to_vec(),
            },
        ]));

        assert_eq!(
            editor
                .storage(TextStorageId::new(42).expect("valid test storage ID"))
                .unwrap()
                .storage
                .text(),
            "Source"
        );
        editor
            .set_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                "Updated",
            )
            .unwrap();
        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let object = archive.object(42).unwrap();
        assert_eq!(object.messages[0], unknown);
        assert_eq!(object.messages[1].type_, 2001);
    }

    #[test]
    fn whole_text_replacement_preserves_drop_cap_sentinel_exactly() {
        let storage = StorageArchive {
            text: vec!["Source".to_owned()],
            table_drop_cap_style: Some(ObjectAttributeTable {
                entries: vec![ObjectAttribute {
                    character_index: 0,
                    object: None,
                }],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                "Changed text",
            )
            .unwrap();
        editor
            .set_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                "Source",
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn replacement_uses_utf16_and_shifts_style_boundaries() {
        let storage = StorageArchive {
            text: vec!["A🚀BC".to_string()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![
                    attribute(0, 10),
                    attribute(1, 11),
                    attribute(3, 12),
                    attribute(5, 13),
                ],
            }),
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 99)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..4,
                "東京",
            )
            .unwrap();

        let package = editor.into_package();
        let archive = package.archive("Index/Document.iwa").unwrap();
        let message = &archive.object(42).unwrap().messages[0];
        let storage = StorageArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(storage.text, ["A東京C"]);
        let indexes = storage.table_char_style.unwrap().entries;
        assert_eq!(
            indexes
                .iter()
                .map(|entry| entry.character_index)
                .collect::<Vec<_>>(),
            [0, 1, 4]
        );
        assert!(storage.table_attachment.is_none());
    }

    #[test]
    fn exclusive_table_reference_is_pruned_without_normalizing_unknown_header_bytes() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 91)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_reference_metadata(
            storage,
            Vec::new(),
            vec![reference_field(vec![9, 1, 2], 91)],
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "longer",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let raw = archive.to_bytes().unwrap();
        assert!(archive_header(&raw).ends_with(UNKNOWN_ARCHIVE_HEADER_FIELD));
        let object = archive.object(42).unwrap();
        assert!(
            object.archive_info.message_infos[0].field_infos[0]
                .object_references
                .is_empty()
        );
        let storage = StorageArchive::decode(object.messages[0].data.as_slice()).unwrap();
        assert_eq!(storage.text, ["Alonger"]);
        assert!(storage.table_attachment.is_none());
    }

    #[test]
    fn reference_shared_with_opaque_field_info_is_retained_conservatively() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 91)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_reference_metadata(
            storage,
            Vec::new(),
            vec![
                reference_field(vec![9, 1, 2], 91),
                reference_field(vec![99], 91),
            ],
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "longer",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let raw = archive.to_bytes().unwrap();
        assert!(archive_header(&raw).ends_with(UNKNOWN_ARCHIVE_HEADER_FIELD));
        let object = archive.object(42).unwrap();
        let field_infos = &object.archive_info.message_infos[0].field_infos;
        assert_eq!(field_infos[0].object_references, [91]);
        assert_eq!(field_infos[1].object_references, [91]);
        let storage = StorageArchive::decode(object.messages[0].data.as_slice()).unwrap();
        assert!(storage.table_attachment.is_none());
    }

    #[test]
    fn complete_known_top_level_reference_multiset_is_safe_to_prune() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 91)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_reference_metadata(
            storage,
            vec![91],
            Vec::new(),
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "longer",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let object = archive.object(42).unwrap();
        assert!(
            object.archive_info.message_infos[0]
                .object_references
                .is_empty()
        );
        assert!(
            archive_header(&archive.to_bytes().unwrap()).ends_with(UNKNOWN_ARCHIVE_HEADER_FIELD)
        );
    }

    #[test]
    fn reference_retained_by_another_storage_field_is_not_pruned() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 91)],
            }),
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 91)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_reference_metadata(
            storage,
            vec![91, 91],
            Vec::new(),
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "longer",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let object = archive.object(42).unwrap();
        assert_eq!(
            object.archive_info.message_infos[0].object_references,
            [91, 91]
        );
        let storage = StorageArchive::decode(object.messages[0].data.as_slice()).unwrap();
        assert!(storage.table_attachment.is_none());
        assert_eq!(
            storage.table_char_style.unwrap().entries[0]
                .object
                .unwrap()
                .identifier,
            91
        );
    }

    #[test]
    fn same_text_replacement_still_removes_intersecting_object_attributes() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_smartfield: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 91)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package_with_reference_metadata(
            storage,
            vec![91],
            Vec::new(),
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                0..1,
                "A",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let object = archive.object(42).unwrap();
        let storage = StorageArchive::decode(object.messages[0].data.as_slice()).unwrap();
        assert!(storage.table_smartfield.is_none());
        assert!(
            object.archive_info.message_infos[0]
                .object_references
                .is_empty()
        );
    }

    #[test]
    fn unknown_storage_field_blocks_all_reference_pruning() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 91)],
            }),
            ..Default::default()
        };
        let mut data = storage.encode_to_vec();
        append_unknown_varint(&mut data, 99, 91);
        let mut editor = IWorkTextEditor::from_package(test_package_with_raw_reference_metadata(
            data,
            vec![91],
            vec![reference_field(vec![9, 1, 2], 91)],
        ));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "longer",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let message_info = &archive.object(42).unwrap().archive_info.message_infos[0];
        assert_eq!(message_info.object_references, [91]);
        assert_eq!(message_info.field_infos[0].object_references, [91]);
    }

    #[test]
    fn invalid_surrogate_boundary_is_transactional() {
        let storage = StorageArchive {
            text: vec!["🚀".to_string()],
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .replace_text(
                    TextStorageId::new(42).expect("valid test storage ID"),
                    1..2,
                    "x",
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn generic_editor_preserves_structural_control_insertion_compatibility() {
        let storage = StorageArchive {
            text: vec!["AB".to_owned()],
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));

        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..1,
                "\u{4}\u{e}",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let storage =
            StorageArchive::decode(archive.object(42).unwrap().messages[0].data.as_slice())
                .unwrap();
        assert_eq!(storage.text, ["A\u{4}\u{e}B"]);
    }

    #[test]
    fn insertion_retains_section_boundary_at_replacement_start() {
        let storage = StorageArchive {
            text: Vec::new(),
            table_section: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 77)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                0..0,
                "Body",
            )
            .unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let storage =
            StorageArchive::decode(archive.object(42).unwrap().messages[0].data.as_slice())
                .unwrap();
        assert_eq!(storage.text, ["Body"]);
        assert_eq!(storage.table_section.unwrap().entries[0].character_index, 0);
    }

    #[test]
    fn text_replace_restore_preserves_deep_unknown_fields_exactly() {
        let storage = StorageArchive {
            text: vec!["A🚀BC".to_owned()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 10), attribute(5, 11)],
            }),
            table_overlapping_highlight: Some(OverlappingFieldAttributeTable {
                entries: vec![OverlappingFieldAttribute {
                    range: TspRange {
                        location: 4,
                        length: 1,
                    },
                    field: Reference {
                        identifier: 77,
                        ..Default::default()
                    },
                }],
            }),
            ..Default::default()
        };
        let mut package = test_package(storage);
        package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive.object_mut(42).unwrap();
                let message_type = object.messages[0].type_;
                let mut data = crate::wire::transform_length_delimited_fields_at_path(
                    object.messages[0].data.as_slice(),
                    &[8],
                    |table| {
                        let mut table = crate::wire::transform_length_delimited_fields_at_path(
                            table,
                            &[1],
                            |entry| {
                                let mut entry =
                                    crate::wire::transform_length_delimited_fields_at_path(
                                        entry,
                                        &[2],
                                        |reference| {
                                            let mut reference = reference.to_vec();
                                            append_unknown_varint(&mut reference, 96, 960);
                                            Ok(reference)
                                        },
                                    )?;
                                append_unknown_varint(&mut entry, 97, 970);
                                Ok(entry)
                            },
                        )?;
                        append_unknown_varint(&mut table, 98, 980);
                        Ok(table)
                    },
                )?;
                data = crate::wire::transform_length_delimited_fields_at_path(
                    &data,
                    &[25],
                    |table| {
                        let mut table = crate::wire::transform_length_delimited_fields_at_path(
                            table,
                            &[1],
                            |entry| {
                                let mut entry =
                                    crate::wire::transform_length_delimited_fields_at_path(
                                        entry,
                                        &[1],
                                        |range| {
                                            let mut range = range.to_vec();
                                            append_unknown_varint(&mut range, 93, 930);
                                            Ok(range)
                                        },
                                    )?;
                                entry = crate::wire::transform_length_delimited_fields_at_path(
                                    &entry,
                                    &[2],
                                    |reference| {
                                        let mut reference = reference.to_vec();
                                        append_unknown_varint(&mut reference, 92, 920);
                                        Ok(reference)
                                    },
                                )?;
                                append_unknown_varint(&mut entry, 94, 940);
                                Ok(entry)
                            },
                        )?;
                        append_unknown_varint(&mut table, 95, 950);
                        Ok(table)
                    },
                )?;
                append_unknown_varint(&mut data, 99, 990);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                object.archive_info.message_infos[0].object_references = vec![10, 11, 77];
                Ok(())
            })
            .unwrap();
        let before = package
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap();
        let mut editor = IWorkTextEditor::from_package(package);
        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..3,
                "X",
            )
            .unwrap();
        editor
            .replace_text(
                TextStorageId::new(42).expect("valid test storage ID"),
                1..2,
                "🚀",
            )
            .unwrap();
        let after = editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap();
        assert_eq!(after, before);
    }

    #[test]
    fn duplicate_attribute_indexes_fail_transactionally() {
        let storage = StorageArchive {
            text: vec!["Body".to_owned()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 10)],
            }),
            ..Default::default()
        };
        let mut package = test_package(storage);
        package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive.object_mut(42).unwrap();
                let message_type = object.messages[0].type_;
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    object.messages[0].data.as_slice(),
                    &[8, 1],
                    |entry| {
                        let mut entry = entry.to_vec();
                        entry.extend(litchi_iwa_common::varint::encode_varint(8));
                        entry.extend(litchi_iwa_common::varint::encode_varint(0));
                        Ok(entry)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut editor = IWorkTextEditor::from_package(package);
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .replace_text(
                    TextStorageId::new(42).expect("valid test storage ID"),
                    0..0,
                    "X",
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    fn attribute(character_index: u32, identifier: u64) -> ObjectAttribute {
        ObjectAttribute {
            character_index,
            object: Some(Reference {
                identifier,
                ..Default::default()
            }),
        }
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }

    fn reference_field(path: Vec<u32>, identifier: u64) -> FieldInfo {
        let mut field = FieldInfo::new(path);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references = vec![identifier];
        field
    }

    fn archive_header(data: &[u8]) -> &[u8] {
        let (header_length, prefix_length) =
            litchi_iwa_common::varint::decode_varint_from_bytes(data).unwrap();
        let header_length = usize::try_from(header_length).unwrap();
        &data[prefix_length..prefix_length + header_length]
    }

    fn append_unknown_archive_header(data: &[u8]) -> Vec<u8> {
        let (header_length, prefix_length) =
            litchi_iwa_common::varint::decode_varint_from_bytes(data).unwrap();
        let header_length = usize::try_from(header_length).unwrap();
        let payload_start = prefix_length + header_length;
        let mut header = data[prefix_length..payload_start].to_vec();
        header.extend_from_slice(UNKNOWN_ARCHIVE_HEADER_FIELD);
        let mut rewritten = litchi_iwa_common::varint::encode_varint(
            u64::try_from(header.len()).expect("test header length fits u64"),
        );
        rewritten.extend_from_slice(&header);
        rewritten.extend_from_slice(&data[payload_start..]);
        rewritten
    }

    fn test_package(storage: StorageArchive) -> IWorkPackage {
        test_package_with_messages(vec![RawMessage {
            type_: 2001,
            data: storage.encode_to_vec(),
        }])
    }

    fn test_package_with_messages(messages: Vec<RawMessage>) -> IWorkPackage {
        let object = ArchiveObject::new(42, messages).unwrap();
        test_package_with_object(object)
    }

    fn test_package_with_reference_metadata(
        storage: StorageArchive,
        object_references: Vec<u64>,
        field_infos: Vec<FieldInfo>,
    ) -> IWorkPackage {
        test_package_with_raw_reference_metadata(
            storage.encode_to_vec(),
            object_references,
            field_infos,
        )
    }

    fn test_package_with_raw_reference_metadata(
        data: Vec<u8>,
        object_references: Vec<u64>,
        field_infos: Vec<FieldInfo>,
    ) -> IWorkPackage {
        let mut object = ArchiveObject::new(42, vec![RawMessage { type_: 2001, data }]).unwrap();
        object.archive_info.message_infos[0].object_references = object_references;
        object.archive_info.message_infos[0].field_infos = field_infos;
        let source = Archive {
            objects: vec![object],
        }
        .to_bytes()
        .unwrap();
        let source = append_unknown_archive_header(&source);
        let archive = Archive::parse(&source).unwrap();
        test_package_with_archive(archive)
    }

    fn test_package_with_object(object: ArchiveObject) -> IWorkPackage {
        test_package_with_archive(Archive {
            objects: vec![object],
        })
    }

    fn test_package_with_archive(archive: Archive) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/Document.iwa", &archive)
            .unwrap();
        package
    }
}
