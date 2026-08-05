//! TextRun parsing for PowerPoint presentations.
//!
//! Based on Apache POI's HSLF TextRun and related classes, this module
//! provides proper text extraction with formatting from PPT files.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

#[cfg(test)]
use crate::consts::RecordType;
#[cfg(test)]
use crate::records::Record;
#[cfg(test)]
use codec::{
    decode_color_index_struct, formatting_from_style, paragraph_formatting_from_style, utf16_prefix,
};

pub use model::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphRun, ParagraphRunFormatting,
    ParagraphTabAlignment, ParagraphTabStop, ParagraphTextDirection, TextRun, TextRunFormatting,
};
pub use package::TextRunExtractor;
