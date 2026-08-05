//! Shared DrawingML theme model and codecs.
//!
//! This module owns the format-neutral `a:theme` and `a:themeOverride`
//! vocabulary. DOCX, PPTX, XLSX, and XLSB retain their package relationships,
//! resource paths, and format-specific projections; the color/font schemes
//! themselves belong here.

pub mod codec;
pub mod model;

pub use model::{Color, Face, FontSet, Override, Palette, Script, Slot, System, Theme};
