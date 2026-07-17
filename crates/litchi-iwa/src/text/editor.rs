//! Transactional editing of shared iWork text storage objects.

use std::ops::Range;
use std::path::Path;

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::{
    ObjectAttributeTable, OverlappingFieldAttributeTable, ParaDataAttributeTable, StorageArchive,
    StringAttributeTable,
};
use crate::shapes::RgbaColor;
use crate::wire::{
    parse_wire_fields, patch_length_delimited_field, patch_nested_varint_field, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::bookmark::{
    add_text_bookmark, remove_text_bookmark, remove_unreferenced_bookmark_objects, text_bookmarks,
    update_text_bookmark,
};
use super::bookmark_types::{TextBookmark, TextBookmarkId, TextBookmarkSettings};
use super::date_time::{
    add_text_date_time_field, insert_text_date_time_field, remove_text_date_time_field,
    remove_unreferenced_date_time_objects, text_date_time_fields, update_text_date_time_field,
};
use super::date_time_types::{
    TextDateTimeDisplayText, TextDateTimeField, TextDateTimeFieldId, TextDateTimeFieldSettings,
};
use super::drop_cap::{
    ParagraphDropCap, ParagraphDropCapPlacement, ParagraphStart, paragraph_drop_cap,
    paragraph_drop_caps, remove_paragraph_drop_cap, set_paragraph_drop_cap,
};
use super::font::TextFont;
use super::highlight::{
    add_text_highlight, remove_text_highlight, remove_unreferenced_highlight_objects,
    text_highlights, update_text_highlight,
};
use super::highlight_types::{TextHighlight, TextHighlightId};
use super::hyperlink::{
    add_text_hyperlink, remove_text_hyperlink, remove_unreferenced_hyperlink_objects,
    text_hyperlinks, update_text_hyperlink,
};
use super::hyperlink_types::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
use super::language::{
    remove_text_language_boundary, reset_text_languages, set_text_language, text_language,
    text_languages,
};
use super::language_types::{TextLanguage, TextLanguageRun};
use super::number_attachment::{
    insert_text_number_attachment, remove_text_number_attachment,
    remove_unreferenced_number_attachment_objects, text_number_attachments,
    update_text_number_attachment,
};
use super::number_attachment_types::{
    TextNumberAttachment, TextNumberAttachmentId, TextNumberAttachmentSettings,
};
use super::paragraph_alignment::{
    paragraph_alignment, paragraph_indents, paragraph_line_spacing, paragraph_spacing,
    paragraph_tab_stops, reset_paragraph_alignment, reset_paragraph_indents,
    reset_paragraph_line_spacing, reset_paragraph_spacing, reset_paragraph_tab_stops,
    reset_text_background, reset_text_baseline_shift, reset_text_capitalization,
    reset_text_character_spacing, reset_text_color, reset_text_decorations, reset_text_font,
    reset_text_ligatures, reset_text_outline, reset_text_script, reset_text_shadow,
    reset_text_style, set_paragraph_alignment, set_paragraph_indents, set_paragraph_line_spacing,
    set_paragraph_spacing, set_paragraph_tab_stops, set_text_background, set_text_baseline_shift,
    set_text_capitalization, set_text_character_spacing, set_text_color, set_text_decorations,
    set_text_font, set_text_ligatures, set_text_outline, set_text_script, set_text_shadow,
    set_text_style, text_background, text_baseline_shift, text_capitalization,
    text_character_spacing, text_color, text_decorations, text_font, text_ligatures, text_outline,
    text_script, text_shadow, text_style,
};
use super::paragraph_list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, paragraph_list,
    paragraph_list_level, paragraph_list_levels, reset_paragraph_list, reset_paragraph_list_level,
    set_paragraph_list, set_paragraph_list_level,
};
use super::paragraph_tabs::ParagraphTabStops;
use super::position::{TextPosition, TextRange};
use super::style::{
    ParagraphIndents, ParagraphLineSpacing, ParagraphSpacing, TextAlignment, TextBackground,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextOutline, TextScript, TextShadow, TextStyle,
};
use super::text_comment::{
    add_text_comment, add_text_comment_reply, remove_text_comment, remove_text_comment_reply,
    text_comment_replies, text_comments, update_text_comment, update_text_comment_reply,
};
use super::text_comment_types::{
    TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2001, 2022];

/// A discoverable text storage within an iWork package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStorageInfo {
    pub object_id: u64,
    pub message_type: u32,
    pub kind: Option<i32>,
    pub text: String,
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

    pub fn from_package(package: IWorkPackage) -> Self {
        Self { package }
    }

    pub fn storages(&self) -> Result<Vec<TextStorageInfo>> {
        let mut storages = Vec::new();
        for name in self.package.iwa_entry_names() {
            let archive = self.package.archive(name)?;
            for object in archive.objects {
                let Some(object_id) = object.archive_info.identifier else {
                    continue;
                };
                for message in object.messages {
                    if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                        continue;
                    }
                    let Ok(storage) = StorageArchive::decode(message.data.as_slice()) else {
                        continue;
                    };
                    storages.push(TextStorageInfo {
                        object_id,
                        message_type: message.type_,
                        kind: storage.kind,
                        text: storage.text.concat(),
                    });
                }
            }
        }
        storages.sort_by_key(|storage| storage.object_id);
        Ok(storages)
    }

    pub fn storage(&self, object_id: u64) -> Result<TextStorageInfo> {
        let archive_name = find_storage_archive(&self.package, object_id)?;
        let archive = self.package.archive(&archive_name)?;
        let object = archive.object(object_id).ok_or_else(|| {
            Error::ParseError(format!("Text storage object {object_id} not found"))
        })?;
        object
            .messages
            .iter()
            .find_map(|message| {
                if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                    return None;
                }
                StorageArchive::decode(message.data.as_slice())
                    .ok()
                    .map(|storage| TextStorageInfo {
                        object_id,
                        message_type: message.type_,
                        kind: storage.kind,
                        text: storage.text.concat(),
                    })
            })
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Object {object_id} has no writable TSWP storage payload"
                ))
            })
    }

    /// Replace a UTF-16 range, matching the indexing used by iWork attributes.
    pub fn replace_text(
        &mut self,
        object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
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

    pub fn set_text(&mut self, object_id: u64, replacement: &str) -> Result<()> {
        let storage = self.storage(object_id)?;
        self.replace_text(
            object_id,
            0..storage.text.encode_utf16().count(),
            replacement,
        )
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn text_style(&self, object_id: u64) -> Result<TextStyle> {
        text_style(&self.package, object_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_style(&mut self, object_id: u64, style: TextStyle) -> Result<()> {
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
    pub fn reset_text_style(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_font(&self, object_id: u64) -> Result<TextFont> {
        text_font(&self.package, object_id)
    }

    /// Atomically set a typed font identity across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_font(&mut self, object_id: u64, font: TextFont) -> Result<()> {
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
    pub fn reset_text_font(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_languages(&self, object_id: u64) -> Result<Vec<TextLanguageRun>> {
        text_languages(&self.package, object_id)
    }

    /// Read the effective language at one validated UTF-16 text boundary.
    pub fn text_language(&self, object_id: u64, position: TextPosition) -> Result<TextLanguage> {
        text_language(&self.package, object_id, position)
    }

    /// Atomically create or update a language boundary.
    pub fn set_text_language(
        &mut self,
        object_id: u64,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
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
        object_id: u64,
        position: TextPosition,
    ) -> Result<bool> {
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
    pub fn reset_text_languages(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_hyperlinks(&self, object_id: u64) -> Result<Vec<TextHyperlink>> {
        text_hyperlinks(&self.package, object_id)
    }

    /// Create a native hyperlink over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_hyperlink(
        &mut self,
        object_id: u64,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        add_text_hyperlink(&mut self.package, object_id, range, &target)
    }

    /// Atomically update a hyperlink's range and target while retaining its ID.
    pub fn update_text_hyperlink(
        &mut self,
        object_id: u64,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        update_text_hyperlink(&mut self.package, object_id, id, range, &target)
    }

    /// Delete one native hyperlink and reclaim its owned smart-field object.
    pub fn remove_text_hyperlink(
        &mut self,
        object_id: u64,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        remove_text_hyperlink(&mut self.package, object_id, id)
    }

    /// Read every native ranged bookmark in a text storage.
    pub fn text_bookmarks(&self, object_id: u64) -> Result<Vec<TextBookmark>> {
        text_bookmarks(&self.package, object_id)
    }

    /// Create a native bookmark over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_bookmark(
        &mut self,
        object_id: u64,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        add_text_bookmark(&mut self.package, object_id, range, &settings)
    }

    /// Atomically update a bookmark's range and settings while retaining its ID.
    pub fn update_text_bookmark(
        &mut self,
        object_id: u64,
        id: TextBookmarkId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        update_text_bookmark(&mut self.package, object_id, id, range, &settings)
    }

    /// Delete one native bookmark and reclaim its owned bookmark-field object.
    pub fn remove_text_bookmark(
        &mut self,
        object_id: u64,
        id: TextBookmarkId,
    ) -> Result<TextBookmark> {
        remove_text_bookmark(&mut self.package, object_id, id)
    }

    /// Read every native Date & Time smart field in a text storage.
    pub fn text_date_time_fields(&self, object_id: u64) -> Result<Vec<TextDateTimeField>> {
        text_date_time_fields(&self.package, object_id)
    }

    /// Attach a Date & Time field to existing nonempty text.
    pub fn add_text_date_time_field(
        &mut self,
        object_id: u64,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        add_text_date_time_field(&mut self.package, object_id, range, &settings)
    }

    /// Atomically insert exact display text and attach a Date & Time field.
    pub fn insert_text_date_time_field(
        &mut self,
        object_id: u64,
        position: TextPosition,
        display_text: TextDateTimeDisplayText,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
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
        object_id: u64,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        update_text_date_time_field(&mut self.package, object_id, id, range, &settings)
    }

    /// Delete one Date & Time field while retaining its visible text.
    pub fn remove_text_date_time_field(
        &mut self,
        object_id: u64,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        remove_text_date_time_field(&mut self.package, object_id, id)
    }

    /// Read every native page/slide-number attachment in a text storage.
    pub fn text_number_attachments(&self, object_id: u64) -> Result<Vec<TextNumberAttachment>> {
        text_number_attachments(&self.package, object_id)
    }

    /// Atomically insert U+FFFC and its native textual number attachment.
    ///
    /// Attachment evaluation is storage-context dependent. Prefer the Pages
    /// high-level wrappers when creating page numbers or page counts.
    pub fn insert_text_number_attachment(
        &mut self,
        object_id: u64,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        insert_text_number_attachment(&mut self.package, object_id, position, &settings)
    }

    /// Atomically update the lossless payload of a number attachment.
    pub fn update_text_number_attachment(
        &mut self,
        object_id: u64,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        update_text_number_attachment(&mut self.package, object_id, id, &settings)
    }

    /// Delete one number attachment together with its U+FFFC placeholder.
    pub fn remove_text_number_attachment(
        &mut self,
        object_id: u64,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment> {
        remove_text_number_attachment(&mut self.package, object_id, id)
    }

    /// Read every native plain highlight in a text storage.
    pub fn text_highlights(&self, object_id: u64) -> Result<Vec<TextHighlight>> {
        text_highlights(&self.package, object_id)
    }

    /// Create a native plain highlight over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_highlight(
        &mut self,
        object_id: u64,
        range: TextRange,
    ) -> Result<TextHighlight> {
        add_text_highlight(&mut self.package, object_id, range)
    }

    /// Atomically move a plain highlight while retaining its native identity.
    pub fn update_text_highlight(
        &mut self,
        object_id: u64,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        update_text_highlight(&mut self.package, object_id, id, range)
    }

    /// Delete one plain highlight and its owned empty annotation graph.
    pub fn remove_text_highlight(
        &mut self,
        object_id: u64,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        remove_text_highlight(&mut self.package, object_id, id)
    }

    /// Read every native ranged comment in a text storage.
    pub fn text_comments(&self, object_id: u64) -> Result<Vec<TextComment>> {
        text_comments(&self.package, object_id)
    }

    /// Create a native comment over a nonempty, unoccupied UTF-16 range.
    pub fn add_text_comment(
        &mut self,
        object_id: u64,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        add_text_comment(&mut self.package, object_id, range, body)
    }

    /// Atomically update a comment's range and body while retaining its ID.
    pub fn update_text_comment(
        &mut self,
        object_id: u64,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        update_text_comment(&mut self.package, object_id, id, range, body)
    }

    /// Delete one ranged comment and its owned root/reply annotation graph.
    pub fn remove_text_comment(
        &mut self,
        object_id: u64,
        id: TextCommentId,
    ) -> Result<TextComment> {
        remove_text_comment(&mut self.package, object_id, id)
    }

    /// Read every direct reply to one ranged comment in stored order.
    pub fn text_comment_replies(
        &self,
        object_id: u64,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        text_comment_replies(&self.package, object_id, comment_id)
    }

    /// Append a direct reply to one ranged comment.
    pub fn add_text_comment_reply(
        &mut self,
        object_id: u64,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        add_text_comment_reply(&mut self.package, object_id, comment_id, body)
    }

    /// Update a direct reply while retaining its ID and native metadata.
    pub fn update_text_comment_reply(
        &mut self,
        object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        update_text_comment_reply(&mut self.package, object_id, comment_id, reply_id, body)
    }

    /// Delete one direct reply and its owned comment storage.
    pub fn remove_text_comment_reply(
        &mut self,
        object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        remove_text_comment_reply(&mut self.package, object_id, comment_id, reply_id)
    }

    /// Read the canonical list preset applied uniformly to a text storage.
    pub fn paragraph_list(&self, object_id: u64) -> Result<ParagraphList> {
        paragraph_list(&self.package, object_id)
    }

    /// Atomically apply a canonical nine-level list preset to a text storage.
    ///
    /// Storages with multiple list-style boundaries are rejected so the
    /// operation cannot flatten independently formatted paragraphs.
    pub fn set_paragraph_list(&mut self, object_id: u64, list: ParagraphList) -> Result<()> {
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
    pub fn reset_paragraph_list(&mut self, object_id: u64) -> Result<bool> {
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

    /// Read every effective list-level boundary in a text storage.
    pub fn paragraph_list_levels(
        &self,
        object_id: u64,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        paragraph_list_levels(&self.package, object_id)
    }

    /// Read the effective list level at one validated paragraph start.
    pub fn paragraph_list_level(
        &self,
        object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<ParagraphListLevel> {
        paragraph_list_level(&self.package, object_id, paragraph)
    }

    /// Atomically set one paragraph's zero-based list nesting level.
    pub fn set_paragraph_list_level(
        &mut self,
        object_id: u64,
        paragraph: ParagraphStart,
        level: ParagraphListLevel,
    ) -> Result<()> {
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
        object_id: u64,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
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

    /// Read effective uniform underline and strikethrough formatting.
    pub fn text_decorations(&self, object_id: u64) -> Result<TextDecorations> {
        text_decorations(&self.package, object_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_decorations(
        &mut self,
        object_id: u64,
        decorations: TextDecorations,
    ) -> Result<()> {
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
    pub fn reset_text_decorations(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_color(&self, object_id: u64) -> Result<RgbaColor> {
        text_color(&self.package, object_id)
    }

    /// Atomically set one text color across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_color(&mut self, object_id: u64, color: RgbaColor) -> Result<()> {
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
    pub fn reset_text_color(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_capitalization(&self, object_id: u64) -> Result<TextCapitalization> {
        text_capitalization(&self.package, object_id)
    }

    /// Atomically set one capitalization mode across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_capitalization(
        &mut self,
        object_id: u64,
        capitalization: TextCapitalization,
    ) -> Result<()> {
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
    pub fn reset_text_capitalization(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_script(&self, object_id: u64) -> Result<TextScript> {
        text_script(&self.package, object_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_script(&mut self, object_id: u64, script: TextScript) -> Result<()> {
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
    pub fn reset_text_script(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_baseline_shift(&self, object_id: u64) -> Result<TextBaselineShift> {
        text_baseline_shift(&self.package, object_id)
    }

    /// Atomically set a signed custom baseline displacement.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_baseline_shift(
        &mut self,
        object_id: u64,
        shift: TextBaselineShift,
    ) -> Result<()> {
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
    pub fn reset_text_baseline_shift(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_character_spacing(&self, object_id: u64) -> Result<TextCharacterSpacing> {
        text_character_spacing(&self.package, object_id)
    }

    /// Atomically set character spacing across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_character_spacing(
        &mut self,
        object_id: u64,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
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
    pub fn reset_text_character_spacing(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_ligatures(&self, object_id: u64) -> Result<TextLigatures> {
        text_ligatures(&self.package, object_id)
    }

    /// Atomically set the ligature policy across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_ligatures(&mut self, object_id: u64, ligatures: TextLigatures) -> Result<()> {
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
    pub fn reset_text_ligatures(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_outline(&self, object_id: u64) -> Result<TextOutline> {
        text_outline(&self.package, object_id)
    }

    /// Atomically set a typed outline stroke across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_outline(&mut self, object_id: u64, outline: TextOutline) -> Result<()> {
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
    pub fn reset_text_outline(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_shadow(&self, object_id: u64) -> Result<TextShadow> {
        text_shadow(&self.package, object_id)
    }

    /// Atomically set a typed drop shadow across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_shadow(&mut self, object_id: u64, shadow: TextShadow) -> Result<()> {
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
    pub fn reset_text_shadow(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn text_background(&self, object_id: u64) -> Result<TextBackground> {
        text_background(&self.package, object_id)
    }

    /// Atomically set a typed solid background across a uniformly styled storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// the operation cannot flatten independently formatted paragraphs.
    pub fn set_text_background(
        &mut self,
        object_id: u64,
        background: TextBackground,
    ) -> Result<()> {
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
    pub fn reset_text_background(&mut self, object_id: u64) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = reset_text_background(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective uniform paragraph alignment of a text storage.
    pub fn paragraph_alignment(&self, object_id: u64) -> Result<TextAlignment> {
        paragraph_alignment(&self.package, object_id)
    }

    /// Set one alignment for every paragraph in a uniformly styled text storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// this whole-storage operation can never flatten unrelated formatting.
    pub fn set_paragraph_alignment(
        &mut self,
        object_id: u64,
        alignment: TextAlignment,
    ) -> Result<()> {
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
    pub fn reset_paragraph_alignment(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn paragraph_line_spacing(&self, object_id: u64) -> Result<ParagraphLineSpacing> {
        paragraph_line_spacing(&self.package, object_id)
    }

    /// Set one typed line-spacing mode across a uniformly styled text storage.
    pub fn set_paragraph_line_spacing(
        &mut self,
        object_id: u64,
        spacing: ParagraphLineSpacing,
    ) -> Result<()> {
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
    pub fn reset_paragraph_line_spacing(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn paragraph_spacing(&self, object_id: u64) -> Result<ParagraphSpacing> {
        paragraph_spacing(&self.package, object_id)
    }

    /// Atomically set before/after paragraph spacing across a uniform text storage.
    pub fn set_paragraph_spacing(
        &mut self,
        object_id: u64,
        spacing: ParagraphSpacing,
    ) -> Result<()> {
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
    pub fn reset_paragraph_spacing(&mut self, object_id: u64) -> Result<bool> {
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
    pub fn paragraph_indents(&self, object_id: u64) -> Result<ParagraphIndents> {
        paragraph_indents(&self.package, object_id)
    }

    /// Atomically set first-line, left, and right paragraph indentation.
    pub fn set_paragraph_indents(
        &mut self,
        object_id: u64,
        indents: ParagraphIndents,
    ) -> Result<()> {
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
    pub fn reset_paragraph_indents(&mut self, object_id: u64) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = reset_paragraph_indents(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of a uniform text storage.
    pub fn paragraph_tab_stops(&self, object_id: u64) -> Result<ParagraphTabStops> {
        paragraph_tab_stops(&self.package, object_id)
    }

    /// Atomically replace every explicit ruler tab stop of a uniform text storage.
    pub fn set_paragraph_tab_stops(
        &mut self,
        object_id: u64,
        stops: ParagraphTabStops,
    ) -> Result<()> {
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
    pub fn reset_paragraph_tab_stops(&mut self, object_id: u64) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = reset_paragraph_tab_stops(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    /// List plain-text Drop Caps in paragraph-start order.
    pub fn paragraph_drop_caps(&self, object_id: u64) -> Result<Vec<ParagraphDropCapPlacement>> {
        paragraph_drop_caps(&self.package, object_id)
    }

    /// Read the Drop Cap attached to one typed paragraph start.
    pub fn paragraph_drop_cap(
        &self,
        object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<Option<ParagraphDropCap>> {
        paragraph_drop_cap(&self.package, object_id, paragraph_start)
    }

    /// Atomically create or replace a plain-text Drop Cap.
    pub fn set_paragraph_drop_cap(
        &mut self,
        object_id: u64,
        paragraph_start: ParagraphStart,
        drop_cap: ParagraphDropCap,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_paragraph_drop_cap(&mut staged, object_id, paragraph_start, drop_cap)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_drop_cap(&verified, object_id, paragraph_start)? != Some(drop_cap) {
            return Err(Error::InvalidFormat(
                "iWork Drop Cap update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Atomically remove a Drop Cap while retaining its paragraph boundary.
    pub fn remove_paragraph_drop_cap(
        &mut self,
        object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = remove_paragraph_drop_cap(&mut staged, object_id, paragraph_start)?;
        if changed {
            let bytes = staged.to_bytes()?;
            let verified = IWorkPackage::from_bytes(&bytes)?;
            if paragraph_drop_cap(&verified, object_id, paragraph_start)?.is_some() {
                return Err(Error::InvalidFormat(
                    "iWork Drop Cap removal failed round-trip validation".to_owned(),
                ));
            }
            self.package = staged;
        }
        Ok(changed)
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub fn into_package(self) -> IWorkPackage {
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
    let archive_name = find_storage_archive(package, object_id)?;
    let mut removed_references = std::collections::HashSet::new();
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(object_id).ok_or_else(|| {
            Error::ParseError(format!("Text storage object {object_id} not found"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Object {object_id} has no writable TSWP storage payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let mut storage = StorageArchive::decode(original)?;
        let previous_references = storage_object_references(&storage);
        let current = storage.text.concat();
        let start = utf16_to_byte_index(&current, range.start)?;
        let end = utf16_to_byte_index(&current, range.end)?;
        let mut updated = String::with_capacity(
            current
                .len()
                .saturating_sub(end - start)
                .saturating_add(replacement.len()),
        );
        updated.push_str(&current[..start]);
        updated.push_str(replacement);
        updated.push_str(&current[end..]);

        let replacement_units = replacement.encode_utf16().count();
        adjust_storage_attributes(&mut storage, range.clone(), replacement_units)?;
        storage.text = if updated.is_empty() {
            if current.is_empty() {
                storage.text.clone()
            } else {
                Vec::new()
            }
        } else {
            vec![updated]
        };
        let data = patch_storage_text_wire(original, &range, replacement_units, &storage)?;
        if StorageArchive::decode(data.as_slice())? != storage {
            return Err(Error::InvalidFormat(
                "TSWP text-storage wire patch failed validation".to_owned(),
            ));
        }
        let current_references = storage_object_references(&storage);
        removed_references = previous_references
            .into_iter()
            .filter(|identifier| !current_references.contains(identifier))
            .collect::<std::collections::HashSet<_>>();
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .retain(|identifier| !removed_references.contains(identifier));
        for field in &mut object.archive_info.message_infos[message_index].field_infos {
            field
                .object_references
                .retain(|identifier| !removed_references.contains(identifier));
        }
        Ok(())
    })?;
    remove_unreferenced_hyperlink_objects(package, &archive_name, &removed_references)?;
    remove_unreferenced_date_time_objects(package, &archive_name, &removed_references)?;
    remove_unreferenced_number_attachment_objects(package, &archive_name, &removed_references)?;
    remove_unreferenced_bookmark_objects(package, &archive_name, &removed_references)?;
    remove_unreferenced_highlight_objects(package, &archive_name, &removed_references)
}

fn find_storage_archive(package: &IWorkPackage, object_id: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(object_id) else {
            continue;
        };
        if !object.messages.iter().any(|message| {
            STORAGE_MESSAGE_TYPES.contains(&message.type_)
                && StorageArchive::decode(message.data.as_slice()).is_ok()
        }) {
            continue;
        }
        if found.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "Text storage object {object_id} occurs in multiple IWA components"
            )));
        }
    }
    found.ok_or_else(|| Error::ParseError(format!("Text storage object {object_id} not found")))
}

fn utf16_to_byte_index(text: &str, target: usize) -> Result<usize> {
    if target == 0 {
        return Ok(0);
    }
    let mut units = 0usize;
    for (byte_index, character) in text.char_indices() {
        if units == target {
            return Ok(byte_index);
        }
        units += character.len_utf16();
        if units > target {
            return Err(Error::ParseError(format!(
                "UTF-16 index {target} splits a surrogate pair"
            )));
        }
    }
    if units == target {
        Ok(text.len())
    } else {
        Err(Error::ParseError(format!(
            "UTF-16 index {target} exceeds text length {units}"
        )))
    }
}

fn adjust_storage_attributes(
    storage: &mut StorageArchive,
    range: Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_layout_style,
    ] {
        adjust_object_table(table, &range, replacement_units, true)?;
    }
    for table in [
        &mut storage.table_para_data,
        &mut storage.table_para_starts,
        &mut storage.table_para_bidi,
    ] {
        adjust_para_table(table, &range, replacement_units)?;
    }
    for table in [&mut storage.table_language, &mut storage.table_dictation] {
        adjust_string_table(table, &range, replacement_units)?;
    }
    for table in [
        &mut storage.table_attachment,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_tatechuyoko,
    ] {
        adjust_object_table(table, &range, replacement_units, false)?;
    }
    if storage
        .table_attachment
        .as_ref()
        .is_some_and(|table| table.entries.is_empty())
    {
        storage.table_attachment = None;
    }
    normalize_ranged_object_table(&mut storage.table_bookmark);
    adjust_object_table(
        &mut storage.table_smartfield,
        &range,
        replacement_units,
        false,
    )?;
    normalize_ranged_object_table(&mut storage.table_smartfield);
    adjust_object_table(
        &mut storage.table_highlight,
        &range,
        replacement_units,
        false,
    )?;
    normalize_ranged_object_table(&mut storage.table_highlight);
    // A drop-cap record is a paragraph-start boundary. Numbers commonly emits
    // an explicit index-zero entry with no style reference as a sentinel, and
    // replacing the paragraph text must not erase that structural marker.
    adjust_object_table(
        &mut storage.table_drop_cap_style,
        &range,
        replacement_units,
        true,
    )?;
    // A section boundary identifies the section owning text inserted exactly
    // at that boundary. In particular, the mandatory first boundary must stay
    // at UTF-16 index zero when text is inserted into an empty Pages body.
    adjust_object_table(&mut storage.table_section, &range, replacement_units, true)?;
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ] {
        adjust_overlapping_table(table, &range, replacement_units)?;
    }
    Ok(())
}

fn patch_storage_text_wire(
    original: &[u8],
    range: &Range<usize>,
    replacement_units: usize,
    storage: &StorageArchive,
) -> Result<Vec<u8>> {
    let text = storage
        .text
        .iter()
        .map(|value| value.as_bytes().to_vec())
        .collect::<Vec<_>>();
    let mut data = rewrite_repeated_length_delimited_fields(original, 3, &text)?;
    for field in [5, 7, 8, 12, 17, 28] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, true)
        })?;
    }
    let attachment_tables = repeated_length_delimited_payloads(&data, 9)?;
    if attachment_tables.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular TSWP storage table field 9 occurs {} times",
            attachment_tables.len()
        )));
    }
    if let Some(table) = attachment_tables.first() {
        let adjusted = adjust_index_table_wire(table, range, replacement_units, false)?;
        let replacement =
            (!repeated_length_delimited_payloads(&adjusted, 1)?.is_empty()).then_some(adjusted);
        data = patch_length_delimited_field(&data, 9, true, replacement.as_deref())?;
    }
    for field in [16, 18, 21, 22, 27] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, false)
        })?;
    }
    for field in [11, 15, 23] {
        let tables = repeated_length_delimited_payloads(&data, field)?;
        if tables.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "singular TSWP storage table field {field} occurs {} times",
                tables.len()
            )));
        }
        if let Some(table) = tables.first() {
            let adjusted = adjust_index_table_wire(table, range, replacement_units, false)?;
            let normalized = normalize_ranged_object_table_wire(&adjusted)?;
            data = patch_length_delimited_field(&data, field, true, normalized.as_deref())?;
        }
    }
    for field in [6, 14, 19, 20, 24] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, true)
        })?;
    }
    for field in [25, 26] {
        data = transform_optional_table(&data, field, |table| {
            adjust_overlapping_table_wire(table, range, replacement_units)
        })?;
    }
    Ok(data)
}

fn normalize_ranged_object_table(table: &mut Option<ObjectAttributeTable>) {
    let Some(entries) = table.as_mut().map(|table| &mut table.entries) else {
        return;
    };
    let Some(first_object) = entries.iter().position(|entry| entry.object.is_some()) else {
        *table = None;
        return;
    };
    if first_object > 1 {
        entries.drain(..first_object - 1);
    }
    if entries[0].object.is_none() {
        entries[0].character_index = 0;
    } else if entries[0].character_index != 0 {
        entries.insert(
            0,
            crate::protobuf::tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: None,
            },
        );
    }
}

fn normalize_ranged_object_table_wire(table: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .map(|raw| {
            let entry =
                crate::protobuf::tswp::object_attribute_table::ObjectAttribute::decode(raw)?;
            Ok((entry, raw.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    let Some(first_object) = entries.iter().position(|(entry, _)| entry.object.is_some()) else {
        return Ok(None);
    };
    if first_object > 1 {
        entries.drain(..first_object - 1);
    }
    if entries[0].0.object.is_none() {
        entries[0].0.character_index = 0;
        entries[0].1 = patch_varint_field(&entries[0].1, 1, true, Some(0))?;
    } else if entries[0].0.character_index != 0 {
        let sentinel = crate::protobuf::tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: None,
        };
        let raw = sentinel.encode_to_vec();
        entries.insert(0, (sentinel, raw));
    }
    let entries = entries.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, 1, &entries).map(Some)
}

fn transform_optional_table<F>(data: &[u8], field_number: u32, transform: F) -> Result<Vec<u8>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    match repeated_length_delimited_payloads(data, field_number)?.len() {
        0 => Ok(data.to_vec()),
        1 => transform_length_delimited_field(data, field_number, transform),
        count => Err(Error::InvalidFormat(format!(
            "singular TSWP storage table field {field_number} occurs {count} times"
        ))),
    }
}

fn adjust_index_table_wire(
    table: &[u8],
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<Vec<u8>> {
    let mut entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .enumerate()
        .filter_map(|(order, entry)| {
            let index = required_u32_varint(entry, 1);
            match index.and_then(|index| {
                adjust_index(index, range, replacement_units, retain_start_boundary)
            }) {
                Ok(Some(index)) => Some(
                    patch_varint_field(entry, 1, true, Some(u64::from(index)))
                        .map(|entry| (index, order, entry)),
                ),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    entries.sort_by_key(|(index, order, _)| (*index, *order));
    entries.dedup_by_key(|(index, _, _)| *index);
    let entries = entries
        .into_iter()
        .map(|(_, _, entry)| entry)
        .collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, 1, &entries)
}

fn adjust_overlapping_table_wire(
    table: &[u8],
    replacement: &Range<usize>,
    replacement_units: usize,
) -> Result<Vec<u8>> {
    let entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .filter_map(|entry| {
            let adjusted = (|| {
                let ranges = repeated_length_delimited_payloads(entry, 1)?;
                if ranges.len() != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "TSWP overlapping attribute range occurs {} times",
                        ranges.len()
                    )));
                }
                let start = usize::try_from(required_u32_varint(ranges[0], 1)?)
                    .map_err(|_| Error::ParseError("Text attribute index overflow".to_owned()))?;
                let length = usize::try_from(required_u32_varint(ranges[0], 2)?)
                    .map_err(|_| Error::ParseError("Text attribute length overflow".to_owned()))?;
                let end = start
                    .checked_add(length)
                    .ok_or_else(|| Error::ParseError("Text attribute range overflow".to_owned()))?;
                if end <= replacement.start {
                    return Ok(Some(entry.to_vec()));
                }
                if start >= replacement.end {
                    let shifted = shift_index(start, replacement, replacement_units)?;
                    return patch_nested_varint_field(
                        entry,
                        &[1, 1],
                        true,
                        Some(u64::from(shifted)),
                    )
                    .map(Some);
                }
                Ok(None)
            })();
            match adjusted {
                Ok(Some(entry)) => Some(Ok(entry)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(table, 1, &entries)
}

fn required_u32_varint(data: &[u8], field_number: u32) -> Result<u32> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != 1 || matches[0].wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "required protobuf varint field {field_number} occurs {} times or has the wrong wire type",
            matches.len()
        )));
    }
    let field = matches[0];
    let (value, length) = crate::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
        .map_err(|error| Error::InvalidFormat(format!("invalid protobuf varint: {error}")))?;
    if field.key_end + length != field.end {
        return Err(Error::InvalidFormat(
            "protobuf varint field has trailing bytes".to_owned(),
        ));
    }
    u32::try_from(value).map_err(|_| Error::InvalidFormat("protobuf varint exceeds u32".to_owned()))
}

fn adjust_object_table(
    table: &mut Option<ObjectAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(
                entry.character_index,
                range,
                replacement_units,
                retain_start_boundary,
            )
            .transpose()
            .map(|result| {
                result.map(|index| {
                    entry.character_index = index;
                    entry
                })
            })
        })
        .collect::<Result<Vec<_>>>()?;
    deduplicate_object_entries(&mut table.entries);
    Ok(())
}

fn adjust_para_table(
    table: &mut Option<ParaDataAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(entry.character_index, range, replacement_units, true)
                .transpose()
                .map(|result| {
                    result.map(|index| {
                        entry.character_index = index;
                        entry
                    })
                })
        })
        .collect::<Result<Vec<_>>>()?;
    table.entries.sort_by_key(|entry| entry.character_index);
    table.entries.dedup_by_key(|entry| entry.character_index);
    Ok(())
}

fn adjust_string_table(
    table: &mut Option<StringAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(entry.character_index, range, replacement_units, true)
                .transpose()
                .map(|result| {
                    result.map(|index| {
                        entry.character_index = index;
                        entry
                    })
                })
        })
        .collect::<Result<Vec<_>>>()?;
    table.entries.sort_by_key(|entry| entry.character_index);
    table.entries.dedup_by_key(|entry| entry.character_index);
    Ok(())
}

fn adjust_overlapping_table(
    table: &mut Option<OverlappingFieldAttributeTable>,
    replacement: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    let mut adjusted = Vec::new();
    for mut entry in table.entries.drain(..) {
        let start = entry.range.location as usize;
        let end = start
            .checked_add(entry.range.length as usize)
            .ok_or_else(|| Error::ParseError("Text attribute range overflow".to_string()))?;
        if end <= replacement.start {
            adjusted.push(entry);
        } else if start >= replacement.end {
            entry.range.location = shift_index(start, replacement, replacement_units)?;
            adjusted.push(entry);
        }
        // An annotation intersecting replaced text is intentionally removed.
    }
    table.entries = adjusted;
    Ok(())
}

fn adjust_index(
    index: u32,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<Option<u32>> {
    let index_usize = index as usize;
    if index_usize < range.start || (retain_start_boundary && index_usize == range.start) {
        return Ok(Some(index));
    }
    if index_usize < range.end {
        return Ok(None);
    }
    Ok(Some(shift_index(index_usize, range, replacement_units)?))
}

fn shift_index(index: usize, range: &Range<usize>, replacement_units: usize) -> Result<u32> {
    let removed = range.end - range.start;
    let shifted = if replacement_units >= removed {
        index.checked_add(replacement_units - removed)
    } else {
        index.checked_sub(removed - replacement_units)
    }
    .ok_or_else(|| Error::ParseError("Text attribute index overflow".to_string()))?;
    u32::try_from(shifted)
        .map_err(|_| Error::ParseError("Text attribute index exceeds u32".to_string()))
}

fn deduplicate_object_entries(
    entries: &mut Vec<crate::protobuf::tswp::object_attribute_table::ObjectAttribute>,
) {
    entries.sort_by_key(|entry| entry.character_index);
    entries.dedup_by_key(|entry| entry.character_index);
}

fn storage_object_references(storage: &StorageArchive) -> Vec<u64> {
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
    use crate::archive::{Archive, ArchiveObject};
    use crate::protobuf::tsp::{Range as TspRange, Reference};
    use crate::protobuf::tswp::object_attribute_table::ObjectAttribute;
    use crate::protobuf::tswp::overlapping_field_attribute_table::OverlappingFieldAttribute;

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
        editor.set_text(42, "Changed text").unwrap();
        editor.set_text(42, "Source").unwrap();
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
        editor.replace_text(42, 1..4, "東京").unwrap();

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
    fn invalid_surrogate_boundary_is_transactional() {
        let storage = StorageArchive {
            text: vec!["🚀".to_string()],
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        let before = editor.to_bytes().unwrap();
        assert!(editor.replace_text(42, 1..2, "x").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
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
        editor.replace_text(42, 0..0, "Body").unwrap();

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
        editor.replace_text(42, 1..3, "X").unwrap();
        editor.replace_text(42, 1..2, "🚀").unwrap();
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
                        entry.extend(crate::varint::encode_varint(8));
                        entry.extend(crate::varint::encode_varint(0));
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
        assert!(editor.replace_text(42, 0..0, "X").is_err());
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
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }

    fn test_package(storage: StorageArchive) -> IWorkPackage {
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2001,
                data: storage.encode_to_vec(),
            }],
        )
        .unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();
        package
    }
}
