//! Typed font collection models.

/// One embedded OpenType font facet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFont {
    /// Facet index: plain, bold, italic, or bold-italic (`0..=3`).
    pub style: u8,
    /// Embedded font bytes in the format specified by MS-PPT.
    pub data: Vec<u8>,
}

/// Font attributes from a `FontEntityAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    /// Zero-based font index from the atom's record instance.
    pub index: u16,
    /// Null-terminated UTF-16 typeface name.
    pub name: String,
    /// Windows character-set identifier.
    pub charset: u8,
    /// Raw font flags byte.
    pub font_flags: u8,
    /// Whether only a subset of the font is embedded.
    pub embedded_subset: bool,
    /// Raw four-bit font type flags.
    pub font_type_flags: u8,
    /// Whether this is a raster font.
    pub raster: bool,
    /// Whether this is a device font.
    pub device: bool,
    /// Whether this is a TrueType font.
    pub truetype: bool,
    /// Whether font substitution is disabled.
    pub no_substitution: bool,
    /// Windows pitch and family byte.
    pub pitch_and_family: u8,
    /// Optional embedded font facets in record order.
    pub embedded_fonts: Vec<EmbeddedFont>,
}

/// Parsed base or PowerPoint 10 font collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontCollection {
    /// Whether this is the international `FontCollection10Container`.
    pub international: bool,
    /// Fonts in collection order.
    pub fonts: Vec<Font>,
}

impl FontCollection {
    /// Resolve a zero-based font reference.
    pub fn get(&self, index: u16) -> Option<&Font> {
        self.fonts.iter().find(|font| font.index == index)
    }
}

/// PowerPoint 10 document-wide font embedding settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FontEmbeddingFlags {
    /// Raw flags word. Only bits 0 and 1 are defined by MS-PPT.
    pub raw: u32,
    /// Whether embedded fonts contain only the characters used by the presentation.
    pub subset: bool,
    /// Whether the user confirmed the subset choice in the user interface.
    pub subset_option_confirmed: bool,
}

/// Base and international font collections resolved from a PPT record tree.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FontCollections {
    /// Base font collection from `DocumentTextInfoContainer`.
    pub base: Option<FontCollection>,
    /// PowerPoint 10 international font collection from `___PPT10`.
    pub international: Option<FontCollection>,
    /// PowerPoint 10 document-wide font embedding settings from `___PPT10`.
    pub embedding_flags: Option<FontEmbeddingFlags>,
}

impl FontCollections {
    /// Resolve a base-font reference.
    pub fn get_base(&self, index: u16) -> Option<&Font> {
        self.base.as_ref()?.get(index)
    }

    /// Resolve a PowerPoint 10 international-font reference.
    pub fn get_international(&self, index: u16) -> Option<&Font> {
        self.international.as_ref()?.get(index)
    }
}
