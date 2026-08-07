//! Text Extraction Utilities for iWork Documents
//!
//! This module provides shared text extraction functionality used across
//! Pages, Numbers, and Keynote documents.

mod annotation;
mod annotation_reply;
mod bookmark;
mod bookmark_object;
mod bookmark_types;
mod character;
pub(crate) mod columns;
mod date_time;
mod date_time_field;
mod date_time_object;
mod drop_cap;
pub mod editor;
pub(crate) mod extractor;
mod font;
mod highlight;
mod highlight_object;
mod highlight_storage;
mod hyperlink;
mod hyperlink_object;
mod hyperlink_storage;
mod language;
mod number_attachment;
mod number_attachment_object;
mod number_attachment_storage;
pub(crate) mod paragraph_alignment;
mod paragraph_direction;
mod paragraph_flow;
mod paragraph_list;
mod paragraph_style_apply;
mod paragraph_style_catalog;
mod paragraph_style_delete;
mod paragraph_style_redefine;
mod paragraph_style_rename;
mod paragraph_tabs;
mod smart_field_object;
mod storage_id;
mod storage_wire;
pub(crate) mod style_registry;
mod text_comment;

#[cfg(test)]
mod bookmark_tests;
#[cfg(test)]
mod date_time_tests;
#[cfg(test)]
mod highlight_tests;
#[cfg(test)]
mod hyperlink_tests;
#[cfg(test)]
mod language_tests;
#[cfg(test)]
mod number_attachment_tests;
#[cfg(test)]
mod text_comment_tests;

pub use bookmark_types::{
    TextBookmark, TextBookmarkId, TextBookmarkName, TextBookmarkSettings, TextBookmarkVisibility,
};
pub use date_time_field::{TextDateTimeField, TextDateTimeFieldId};
pub use editor::{IWorkTextEditor, TextStorageInfo};
pub use storage_id::{TextStorageId, TextStorageIdError};

/// Validate a native storage reference at the IWA adapter boundary.
pub(crate) fn native_storage_id(raw: u64) -> crate::Result<TextStorageId> {
    TextStorageId::try_from(raw).map_err(|error| crate::Error::InvalidFormat(error.to_string()))
}
pub(crate) use extractor::TextExtractor;
pub use font::{Font, Name, NameError, TextFont, TextFontName};
pub use litchi_iwa_common::text::layout;
pub use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
pub use litchi_iwa_text::character::{
    Error as CharacterError, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextPointSize, TextScript, TextStrikethrough, TextStyle,
    TextUnderline,
};
pub use litchi_iwa_text::comment::{
    TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
};
pub use litchi_iwa_text::highlight::{TextHighlight, TextHighlightId};
pub use litchi_iwa_text::hyperlink::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
pub use litchi_iwa_text::language::{
    Error as TextLanguageError, TextLanguage, TextLanguageRun, TextLanguageTag,
};
pub use litchi_iwa_text::number_attachment::{
    TextNumberAttachment, TextNumberAttachmentFormat, TextNumberAttachmentId,
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText,
};
pub use litchi_iwa_text::paragraph;
pub use litchi_iwa_text::paragraph::format::{
    Alignment, Border, Borders, Format, IndentPoints, Indents, LineSpacing, LineSpacingMultiple,
    LineSpacingPoints, Spacing, SpacingPoints,
};
pub use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphFollowingStyle, ParagraphStyleId, ParagraphStyleName,
};
pub use litchi_iwa_text::position::{Error as TextPositionError, TextPosition, TextRange};
pub use litchi_iwa_text::storage::{Error as StorageError, Fragment, Run, Storage};
pub(crate) use paragraph_alignment::{
    applied_named_paragraph_style as applied_named_paragraph_style_in_storage,
    named_paragraph_styles as named_paragraph_styles_in_storage,
};
pub use paragraph_direction::ParagraphWritingDirection;
pub use paragraph_flow::{ParagraphFlow, ParagraphHyphenation};
pub use paragraph_list::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
    ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
    ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberPunctuation,
    ParagraphListNumberScale, ParagraphListNumberSequence, ParagraphListNumberTiering,
    ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
};
pub(crate) use paragraph_list::{
    paragraph_list as paragraph_list_in_storage,
    paragraph_list_bullet as paragraph_list_bullet_in_storage,
    paragraph_list_bullet_geometry as paragraph_list_bullet_geometry_in_storage,
    paragraph_list_indentation as paragraph_list_indentation_in_storage,
    paragraph_list_label_color as paragraph_list_label_color_in_storage,
    paragraph_list_level as paragraph_list_level_in_storage,
    paragraph_list_levels as paragraph_list_levels_in_storage,
    paragraph_list_number_format as paragraph_list_number_format_in_storage,
    paragraph_list_number_scale as paragraph_list_number_scale_in_storage,
    paragraph_list_number_tiering as paragraph_list_number_tiering_in_storage,
    paragraph_list_numbering as paragraph_list_numbering_in_storage,
    paragraph_lists as paragraph_lists_in_storage, preset_style_id, preset_style_object,
};
pub use paragraph_style_apply::AppliedParagraphStyle;
pub use paragraph_tabs::{
    ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops,
};

impl From<litchi_iwa_text::appearance::Error> for crate::Error {
    fn from(error: litchi_iwa_text::appearance::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_text::paragraph::format::Error> for crate::Error {
    fn from(error: litchi_iwa_text::paragraph::format::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_text::NameError> for crate::Error {
    fn from(error: litchi_iwa_text::NameError) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_text::date_time::Error> for crate::Error {
    fn from(error: litchi_iwa_text::date_time::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_text::language::Error> for crate::Error {
    fn from(error: litchi_iwa_text::language::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

impl From<litchi_iwa_text::position::Error> for crate::Error {
    fn from(error: litchi_iwa_text::position::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

impl From<litchi_iwa_text::paragraph::list::Error> for crate::Error {
    fn from(error: litchi_iwa_text::paragraph::list::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_text::paragraph::drop_cap::Error> for crate::Error {
    fn from(error: litchi_iwa_text::paragraph::drop_cap::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}
