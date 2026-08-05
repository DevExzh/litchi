//! Layered PresentationML theme semantics.
//!
//! `model` contains contextual values, `codec` owns bounded XML parsing and
//! serialization, and `package` owns only the OPC relationship graph.

pub mod codec;
pub mod model;
pub mod package;
pub mod part;

pub use model::{Authored, Color, Face, FontSet, Override, Palette, Part, Script, Slot, System};
pub use package::{
    add, attach, load, load_override, next_part_uri, put_colors, put_fonts, put_override,
    remove_override, validate,
};
pub use part::{
    Color as ParsedColor, Font as ParsedFont, Part as ThemePart, Summary as ThemeSummary,
};
