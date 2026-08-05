//! Public semantic section-layout model for Word 97+ documents.
//!
//! The section facade keeps the typed model separate from the focused page
//! border and column owners. Binary SEPX parsing and writing remains in the
//! document package and writer layers, while callers continue to use
//! `crate::section` paths.

pub mod borders;
pub mod columns;

mod model;

pub use model::{
    Behavior, BreakKind, ChapterNumberSeparator, DocSection, FootnotePosition, LineNumberRestart,
    LineNumbering, Margins, NoteNumberRestart, NoteSettings, PageGrid, PageGridMode, PageLayout,
    PageNumbering, PageOrientation, PaperSettings, Protection, TextFlow, VerticalJustification,
    VerticalMargin,
};
