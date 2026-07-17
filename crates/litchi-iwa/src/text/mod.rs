//! Text Extraction Utilities for iWork Documents
//!
//! This module provides shared text extraction functionality used across
//! Pages, Numbers, and Keynote documents.

pub(crate) mod columns;
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
mod paragraph_alignment;
mod paragraph_list;
mod paragraph_tabs;
mod position;
pub mod storage;
mod storage_wire;
pub mod style;
mod style_registry;

#[cfg(test)]
mod highlight_tests;
#[cfg(test)]
mod hyperlink_tests;
#[cfg(test)]
mod language_tests;

pub use columns::{
    EqualTextColumns, FollowingTextColumn, TextColumnCount, TextColumnGap, TextColumnWidth,
    TextColumns, VariableTextColumns,
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
pub use paragraph_list::{ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement};
pub use paragraph_tabs::{
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops,
};
pub use position::{TextPosition, TextRange};

pub use extractor::TextExtractor;
pub use storage::{TextFragment, TextRun, TextStorage};
pub use style::{
    ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStyle,
    TextAlignment, TextBackground, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextOutline, TextPointSize, TextScript, TextShadow,
    TextStrikethrough, TextStyle, TextUnderline,
};
