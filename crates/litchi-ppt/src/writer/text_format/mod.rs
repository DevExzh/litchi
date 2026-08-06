//! Hierarchical text-formatting writer for legacy PowerPoint files.
//!
//! The facade keeps the established ergonomic exports while organizing the
//! owner into semantic models, binary codecs, and wire-level validation.
//!
//! Reference: [MS-PPT] Section 2.9 - Text Formatting

pub mod codec;
pub mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::TextPropsBuilder;
pub use semantic::{
    FontEntity, FontStyle, Paragraph, StyleTextPropHeader, TabAlign, TabStop, TextAlign, TextColor,
    TextDirection, TextFontAlign, TextHeaderType, TextRun, char_mask, para_mask,
};
