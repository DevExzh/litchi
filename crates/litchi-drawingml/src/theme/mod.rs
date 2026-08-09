//! Shared `DrawingML` theme model and codecs.
//!
//! This module owns the format-neutral `a:theme` and `a:themeOverride`
//! vocabulary. DOCX, PPTX, XLSX, and XLSB retain their package relationships,
//! resource paths, and format-specific projections; the color/font schemes
//! themselves belong here.
//!
//! The ownership split follows `[MS-ODRAWXML]` theme structures, while
//! `[MS-OI29500]` package/content-type/relationship rules stay with each
//! format facade. `[MS-PPTX]` and `[MS-XLSX]` therefore own their theme-part
//! placement and links; `[MS-XLSB]` retains BIFF12 host indices and binary
//! records around any linked `DrawingML` theme resources.

pub mod codec;
pub mod model;

pub use model::{Color, Face, FontSet, Override, Palette, Script, Slot, System, Theme};
