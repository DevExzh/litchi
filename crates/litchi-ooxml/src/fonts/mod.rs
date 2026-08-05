//! Temporary migration boundary for the legacy OOXML package writer.
//!
//! Font discovery, cmap mapping, subsetting, standalone-face validation, EOT
//! publication, and obfuscation are owned by `litchi-fonts`. This host module
//! only adapts the existing PPTX package integration until that consumer is
//! migrated; it is intentionally not a public compatibility surface.

#[cfg(feature = "fonts")]
use crate::error::{OoxmlError, Result};
#[cfg(feature = "fonts")]
use litchi_fonts::GlyphMap;

#[cfg(feature = "fonts")]
pub(crate) trait EmbedFonts {
    /// Embed fonts into the given OPC package based on used glyphs and save options.
    fn embed_fonts(&mut self) -> Result<()>;
}

#[cfg(feature = "fonts")]
pub(crate) type PreparedFont = litchi_fonts::embedding::Prepared;

#[cfg(feature = "fonts")]
pub(crate) fn prepare_fonts(
    used_glyphs: GlyphMap,
    subset_requested: bool,
) -> Result<Vec<PreparedFont>> {
    litchi_fonts::embedding::prepare(used_glyphs, subset_requested)
        .map_err(|error| OoxmlError::Other(error.to_string()))
}

#[cfg(feature = "fonts")]
pub(crate) fn powerpoint_data(font: &mut PreparedFont) -> Result<Vec<u8>> {
    litchi_fonts::embedding::powerpoint::data(font)
        .map_err(|error| OoxmlError::Other(error.to_string()))
}
