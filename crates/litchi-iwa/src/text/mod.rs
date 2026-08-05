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
mod paragraph_direction;
mod paragraph_flow;
mod paragraph_following_style;
mod paragraph_list;
mod paragraph_style_apply;
mod paragraph_style_catalog;
mod paragraph_style_delete;
mod paragraph_style_redefine;
mod paragraph_style_rename;
mod paragraph_tabs;
mod position;
mod smart_field_object;
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
pub use font::{Font, Name, NameError, TextFont, TextFontName};
pub use highlight_types::{TextHighlight, TextHighlightId};
pub use hyperlink_types::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
pub use language_types::{TextLanguage, TextLanguageRun, TextLanguageTag};
pub use number_attachment_types::{
    TextNumberAttachment, TextNumberAttachmentFormat, TextNumberAttachmentId,
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText,
};
pub(crate) use paragraph_alignment::{
    applied_named_paragraph_style as applied_named_paragraph_style_in_storage,
    named_paragraph_styles as named_paragraph_styles_in_storage,
};
pub use paragraph_direction::ParagraphWritingDirection;
pub use paragraph_flow::{ParagraphFlow, ParagraphHyphenation};
pub use paragraph_following_style::{
    NamedParagraphStyle, ParagraphFollowingStyle, ParagraphStyleId, ParagraphStyleName,
};
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
pub use position::{TextPosition, TextRange};
pub use text_comment_types::{
    TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
};

pub use extractor::TextExtractor;
pub use litchi_iwa_text::{
    TextFragment, TextFragmentIter, TextRun, TextStorage, extract_text_from_storages,
    parse_storage_archive,
};
pub use style::{
    ParagraphBackground, ParagraphBorder, ParagraphBorderOffset, ParagraphBorderSides,
    ParagraphBorders, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
    ParagraphLineSpacingMultiple, ParagraphLineSpacingPoints, ParagraphSpacing,
    ParagraphSpacingPoints, ParagraphStyle, TextAlignment, TextBackground, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures, TextOutline,
    TextPointSize, TextScript, TextShadow, TextStrikethrough, TextStyle, TextUnderline,
};

impl From<litchi_iwa_text::NameError> for crate::Error {
    fn from(error: litchi_iwa_text::NameError) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}
