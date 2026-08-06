//! Layered web-settings codec facade.
//!
//! XML, semantic validation, relationship validation, and OPC part reads are
//! intentionally kept behind this small public surface.

mod package;
mod relationship;
mod rewrite;
mod semantic;
mod xml;

pub use package::read;
pub use xml::{parse, write};

pub(super) use rewrite::rewrite;

pub(super) use semantic::{
    div_position, parse_i64, validate_border_style, validate_divs, validate_encoding,
    validate_pixels_per_inch, validate_relationship_id, validate_text, validate_word_color,
};
