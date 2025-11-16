//! SPRM (Single Property Modifier) generation for DOC files
//!
//! SPRMs are instructions that modify document properties. They are used in
//! character and paragraph formatting (CHP and PAP structures).
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's SprmOperation.

/// SPRM operation codes for character properties
pub mod chp {
    /// Bold
    pub const BOLD: u16 = 0x0835;
    /// Italic
    pub const ITALIC: u16 = 0x0836;
    /// Underline
    pub const UNDERLINE: u16 = 0x2A3E;
    /// Font size (half-points)
    pub const FONT_SIZE: u16 = 0x4A43;
    /// Font family
    pub const FONT_FAMILY: u16 = 0x4A4F;
    /// Text color
    pub const COLOR: u16 = 0x2A42;
    /// Strike through
    pub const STRIKE: u16 = 0x0838;
}

/// SPRM operation codes for paragraph properties
pub mod pap {
    /// Justification (alignment)
    pub const JUSTIFICATION: u16 = 0x2403;
    /// Left indent
    pub const LEFT_INDENT: u16 = 0x840F;
    /// Right indent
    pub const RIGHT_INDENT: u16 = 0x840E;
    /// First line indent
    pub const FIRST_LINE_INDENT: u16 = 0x8411;
    /// Space before
    pub const SPACE_BEFORE: u16 = 0xA413;
    /// Space after
    pub const SPACE_AFTER: u16 = 0xA414;
    /// Line spacing
    pub const LINE_SPACING: u16 = 0x6412;
}

/// SPRM builder for generating property modification sequences
#[derive(Debug, Default)]
pub struct SprmBuilder {
    sprms: Vec<u8>,
}

impl SprmBuilder {
    /// Create a new SPRM builder
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a boolean SPRM (0 or 1 value)
    pub fn add_bool(&mut self, code: u16, value: bool) {
        self.sprms.extend_from_slice(&code.to_le_bytes());
        self.sprms.push(if value { 1 } else { 0 });
    }

    /// Add a byte SPRM
    pub fn add_byte(&mut self, code: u16, value: u8) {
        self.sprms.extend_from_slice(&code.to_le_bytes());
        self.sprms.push(value);
    }

    /// Add a word (u16) SPRM
    pub fn add_word(&mut self, code: u16, value: u16) {
        self.sprms.extend_from_slice(&code.to_le_bytes());
        self.sprms.extend_from_slice(&value.to_le_bytes());
    }

    /// Add a dword (u32) SPRM
    pub fn add_dword(&mut self, code: u16, value: u32) {
        self.sprms.extend_from_slice(&code.to_le_bytes());
        self.sprms.extend_from_slice(&value.to_le_bytes());
    }

    /// Add a signed word SPRM
    pub fn add_signed_word(&mut self, code: u16, value: i16) {
        self.sprms.extend_from_slice(&code.to_le_bytes());
        self.sprms.extend_from_slice(&value.to_le_bytes());
    }

    /// Get the SPRM sequence as bytes
    pub fn build(&self) -> Vec<u8> {
        self.sprms.clone()
    }

    /// Clear all SPRMs
    pub fn clear(&mut self) {
        self.sprms.clear();
    }
}

/// Helper to create character property SPRMs
pub fn build_chp_sprms(bold: bool, italic: bool, font_size: Option<u16>) -> Vec<u8> {
    let mut builder = SprmBuilder::new();

    if bold {
        builder.add_bool(chp::BOLD, true);
    }
    if italic {
        builder.add_bool(chp::ITALIC, true);
    }
    if let Some(size) = font_size {
        builder.add_word(chp::FONT_SIZE, size * 2); // Convert to half-points
    }

    builder.build()
}

/// Helper to create paragraph property SPRMs
pub fn build_pap_sprms(
    alignment: Option<u8>,
    left_indent: Option<i16>,
    space_before: Option<u16>,
) -> Vec<u8> {
    let mut builder = SprmBuilder::new();

    if let Some(align) = alignment {
        builder.add_byte(pap::JUSTIFICATION, align);
    }
    if let Some(indent) = left_indent {
        builder.add_signed_word(pap::LEFT_INDENT, indent);
    }
    if let Some(space) = space_before {
        builder.add_word(pap::SPACE_BEFORE, space);
    }

    builder.build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sprm_builder() {
        let mut builder = SprmBuilder::new();
        builder.add_bool(chp::BOLD, true);
        builder.add_word(chp::FONT_SIZE, 24); // 12pt

        let sprms = builder.build();
        assert!(!sprms.is_empty());
        assert_eq!(sprms.len(), 3 + 4); // bool SPRM (3 bytes) + word SPRM (4 bytes)
    }

    #[test]
    fn test_build_chp_sprms() {
        let sprms = build_chp_sprms(true, true, Some(12));
        assert!(!sprms.is_empty());
    }
}
