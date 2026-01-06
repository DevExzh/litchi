use std::collections::HashMap;

#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;

#[cfg(feature = "fonts")]
pub mod loader;
#[cfg(feature = "fonts")]
pub mod subsetter;

#[cfg(feature = "fonts")]
pub use loader::*;
#[cfg(feature = "fonts")]
pub use subsetter::*;

/// Trait for document types that can collect all glyphs (characters) used in the document.
///
/// This is used to determine which fonts need to be embedded and which glyphs
/// should be included in font subsets.
///
/// Uses `RoaringBitmap` instead of `HashSet<char>` for better cache locality and memory efficiency.
/// The bitmap stores Unicode code points (u32 values from chars).
#[cfg(feature = "fonts")]
pub trait CollectGlyphs {
    /// Returns a map of font names to the set of character code points used with that font.
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap>;
}

#[cfg(feature = "fonts")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontData {
    pub name: String,
    pub data: Vec<u8>,
    pub index: u32,
    pub properties: Option<FontProperties>,
}

/// Font properties needed for Office font embedding
#[cfg(feature = "fonts")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontProperties {
    pub panose: Option<String>,
    pub charset: Option<String>,
    pub family: Option<String>,
    pub pitch: Option<String>,
    /// Unicode signature (usb0, usb1, usb2, usb3, csb0, csb1)
    pub sig: Option<(String, String, String, String, String, String)>,
}

#[cfg(feature = "fonts")]
pub trait FontSubsetter {
    fn subset(&self, font: &FontData, glyph_ids: &[u16]) -> Result<Vec<u8>, FontError>;
}

#[cfg(feature = "fonts")]
#[derive(Debug, thiserror::Error)]
pub enum FontError {
    #[error("Font not found: {0}")]
    NotFound(String),
    #[error("Invalid font data")]
    InvalidData,
    #[error("Subsetting failed: {0}")]
    SubsettingFailed(String),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
}
