#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Contextual `WordprocessingML` section authoring facade.
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
