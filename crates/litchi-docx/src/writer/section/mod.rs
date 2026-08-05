//! Contextual WordprocessingML section authoring facade.
//!
//! The public section vocabulary remains at this module boundary while the
//! implementation is layered into semantic models, validation, and codecs.

mod model;
pub use model::{
    Art, ChapterSep, Color, Display, EndnotePos, Endnotes, FootnotePos, Footnotes, GridType,
    LineNumberRestart, NoteNumberRestart, OffsetFrom, PageNumberFormat, PageOrientation,
    SectionColumn, SectionColumns, SectionDocumentGrid, SectionHeaderFooterPart,
    SectionHeaderFooterReference, SectionLineNumbering, SectionPageNumbering, SectionPaperSource,
    SectionProperties, SectionTextDirection, SectionVerticalAlignment, Style, ZOrder,
};

pub mod borders;
mod codec;
mod validation;

#[cfg(test)]
mod tests;
