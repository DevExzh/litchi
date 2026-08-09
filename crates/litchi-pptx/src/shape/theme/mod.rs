//! Layered `PresentationML` theme semantics.
//!
//! `DrawingML` color/font schemes and their XML codecs are owned by
//! `litchi-drawingml`; this module retains only `PresentationML` package
//! relationships and the format-specific summary projection.

pub mod package;
pub mod part;

pub use litchi_drawingml::theme::{Color, Face, FontSet, Override, Palette, Script, Slot, System};
pub use package::{
    Authored, add, attach, load, load_override, next_part_uri, put_colors, put_fonts, put_override,
    remove_override, validate,
};
pub use part::{
    Color as ParsedColor, Font as ParsedFont, Part as ThemePart, Summary as ThemeSummary,
};
