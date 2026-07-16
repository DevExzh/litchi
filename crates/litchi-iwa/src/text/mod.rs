//! Text Extraction Utilities for iWork Documents
//!
//! This module provides shared text extraction functionality used across
//! Pages, Numbers, and Keynote documents.

pub(crate) mod columns;
mod drop_cap;
pub mod editor;
pub mod extractor;
mod paragraph_alignment;
mod paragraph_tabs;
pub mod storage;
pub mod style;
mod style_registry;

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
pub use paragraph_tabs::{
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops,
};

pub use extractor::TextExtractor;
pub use storage::{TextFragment, TextRun, TextStorage};
pub use style::{
    ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStyle,
    TextAlignment, TextCapitalization, TextDecorations, TextPointSize, TextStrikethrough,
    TextStyle, TextUnderline,
};
