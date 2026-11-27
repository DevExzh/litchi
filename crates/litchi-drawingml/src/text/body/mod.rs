//! Shared `DrawingML` text-body vocabulary.
//!
//! `SpreadsheetDrawing` and `PresentationML` both embed the neutral
//! `a:CT_TextBody` story. Host crates retain their wrapper and relationship
//! semantics while this module owns the reusable body model, its reader
//! ([`reader`]) and its writer ([`writer`]).

mod model;
pub mod raw;
pub mod reader;
pub mod writer;

#[cfg(test)]
mod tests;

pub use model::{
    Body, Bullet, BulletColor, BulletFont, BulletSize, Field, Font, Hyperlink, Insets, Paragraph,
    ParagraphProperties, Picture, Properties, Run, RunKind, RunProperties, Spacing, TabStop,
    TextFill,
};
