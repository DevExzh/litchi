//! Text Extraction Utilities for iWork Documents
//!
//! This module provides shared text extraction functionality used across
//! Pages, Numbers, and Keynote documents.

mod annotation;
mod annotation_reply;
mod bookmark;
mod bookmark_object;
mod bookmark_types;
pub(crate) mod columns;
mod date_time;
mod date_time_object;
mod date_time_types;
mod drop_cap;
pub mod editor;
pub mod extractor;
mod font;
mod highlight;
mod highlight_object;
mod highlight_storage;
mod highlight_types;
mod hyperlink;
mod hyperlink_object;
mod hyperlink_storage;
mod hyperlink_types;
mod language;
mod language_types;
mod number_attachment;
mod number_attachment_object;
mod number_attachment_storage;
mod number_attachment_types;
pub(crate) mod paragraph_alignment;
mod paragraph_list;
mod paragraph_tabs;
mod position;
mod smart_field_object;
pub mod storage;
mod storage_wire;
pub mod style;
pub(crate) mod style_registry;
mod text_comment;
mod text_comment_types;

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
pub use columns::{
    EqualTextColumns, FollowingTextColumn, TextColumnCount, TextColumnGap, TextColumnWidth,
    TextColumns, VariableTextColumns,
};
pub use date_time_types::{
    TextDateTimeDisplayText, TextDateTimeField, TextDateTimeFieldId, TextDateTimeFieldSettings,
    TextDateTimeFormat, TextDateTimeFormatterStyle, TextDateTimeInstant,
    TextDateTimeLocaleIdentifier, TextDateTimeUpdatePlan,
};
pub use drop_cap::{
    DropCapCharacterCount, DropCapCharacterScale, DropCapCornerRadius, DropCapLineCount,
    DropCapOutdent, DropCapPadding, DropCapRaisedLines, DropCapWrap, ParagraphDropCap,
    ParagraphDropCapPlacement, ParagraphStart,
};
pub use editor::{IWorkTextEditor, TextStorageInfo};
pub use font::{TextFont, TextFontName};
pub use highlight_types::{TextHighlight, TextHighlightId};
pub use hyperlink_types::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
pub use language_types::{TextLanguage, TextLanguageRun, TextLanguageTag};
pub use number_attachment_types::{
    TextNumberAttachment, TextNumberAttachmentFormat, TextNumberAttachmentId,
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText,
};
pub use paragraph_list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListPlacement,
};
pub(crate) use paragraph_list::{
    paragraph_list as paragraph_list_in_storage,
    paragraph_list_level as paragraph_list_level_in_storage,
    paragraph_list_levels as paragraph_list_levels_in_storage,
    paragraph_lists as paragraph_lists_in_storage, preset_style_id, preset_style_object,
};
pub use paragraph_tabs::{
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops,
};
pub use position::{TextPosition, TextRange};
pub use text_comment_types::{
    TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
};

pub use extractor::TextExtractor;
pub use storage::{TextFragment, TextRun, TextStorage};
pub use style::{
    ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStyle,
    TextAlignment, TextBackground, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextOutline, TextPointSize, TextScript, TextShadow,
    TextStrikethrough, TextStyle, TextUnderline,
};
